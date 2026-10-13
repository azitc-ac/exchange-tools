#Requires -Version 5.1
<#
.SYNOPSIS
    Prüft, ob eine Postfachdatenbank leer ist - und nennt für jeden Rest den Befehl, der ihn entfernt.

.DESCRIPTION
    Remove-MailboxDatabase und die Deinstallation des letzten Exchange-Servers scheitern, solange
    noch ein Postfach in der Datenbank liegt. Gemeint ist nicht nur das Benutzerpostfach: auch
    Archive, Public-Folder-, Arbitration- und AuditLog-Postfächer blockieren. Das Skript sucht alle
    Arten zusammen und gibt je Fund aus:

      - Typ und Name
      - Blockiert: ja | nein | unklar
      - den passenden Befehl zum Entfernen bzw. Deaktivieren

    Getrennte Postfächer (soft-deleted, disabled) blockieren nicht, werden aber angezeigt: Ihr
    Inhalt verschwindet mit der Datenbank.

    Das Skript ändert nichts. Die Befehle sind Vorschläge zum Kopieren.

.PARAMETER Database
    Eine oder mehrere Datenbanken. Ohne Angabe erscheint eine Auswahl (Mehrfachauswahl möglich).

.PARAMETER PassThru
    Gibt die Funde zusätzlich als Objekte zurück, etwa für Export-Csv.

.EXAMPLE
    .\Test-MailboxDatabaseEmpty.ps1
    .\Test-MailboxDatabaseEmpty.ps1 -Database DB01
    .\Test-MailboxDatabaseEmpty.ps1 -Database DB01, DB02 -PassThru | Export-Csv rest.csv -NoTypeInformation -Encoding UTF8

.NOTES
    Exchange 2013/2016/2019/SE, Windows PowerShell 5.1, Exchange-Verwaltungsshell.
    Exitcode: 0 = alle gewählten Datenbanken ohne blockierende Postfächer,
              1 = mindestens ein blockierendes oder unklares Postfach,
              2 = nichts geprüft (keine Auswahl, Datenbank unbekannt).
#>
[CmdletBinding()]
param(
    [string[]]$Database,
    [switch]$PassThru
)

$script:Version = '1.0.0'
$ErrorActionPreference = 'Stop'

# Voraussetzungen. Der Block zwischen den Markern stammt aus build/Prolog.OnPrem.ps1
# und wird von build/Sync-Prolog.ps1 gepflegt - nicht von Hand aendern.
$script:RequiredCmdlets = @('Get-Mailbox', 'Get-MailboxDatabase', 'Get-MailboxStatistics', 'Set-ADServerSettings')
$script:ToolIsGui       = $false
# <prolog:onprem v1 - Quelle: build/Prolog.OnPrem.ps1, eingefügt von build/Sync-Prolog.ps1.
#                     NICHT von Hand ändern - build/Test-Prolog.ps1 meldet jede Abweichung.>
# Erwartet davor gesetzt:
#   $script:RequiredCmdlets = @('Get-Queue')      (Cmdlets, die das Werkzeug wirklich braucht)
#   $script:ToolIsGui       = $true|$false        (bei $true kommen Fehler als MessageBox)
function Initialize-OnPremPrerequisite {
    [CmdletBinding()]
    param(
        [string[]]$Cmdlets = $script:RequiredCmdlets,
        [bool]$Gui         = [bool]$script:ToolIsGui
    )

    function Stop-WithReason([string]$Text) {
        # Eine mit -noConsole gebaute EXE hat kein Fenster für Write-Host oder throw;
        # ohne MessageBox bliebe der Start wortlos erfolglos.
        if ($Gui) {
            try {
                Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
                [void][System.Windows.Forms.MessageBox]::Show($Text, 'Voraussetzung fehlt', 'OK', 'Warning')
            } catch { Write-Warning $Text }
        }
        throw $Text
    }

    # 1. Das Exchange-Snap-in ist .NET-Framework-Code und lädt nur in Windows PowerShell 5.1.
    #    In PowerShell 7 scheitert Add-PSSnapin mit einer Meldung, die das nicht verrät.
    if ($PSVersionTable.PSEdition -eq 'Core') {
        Stop-WithReason ("Dieses Werkzeug braucht die Exchange-Verwaltungsshell und läuft deshalb nur in " +
                         "Windows PowerShell 5.1, nicht in PowerShell $($PSVersionTable.PSVersion).`n`n" +
                         'Über "Exchange Management Shell" starten oder powershell.exe statt pwsh.exe verwenden.')
    }

    # 2. Snap-in laden, falls die Sitzung es noch nicht hat (etwa beim Start per Doppelklick
    #    oder aus einer EXE heraus - dort ist die EMS-Umgebung nicht vorhanden).
    if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
        if (-not (Get-PSSnapin -Registered -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
            Stop-WithReason ("Auf diesem Rechner ist keine Exchange-Verwaltungsshell installiert.`n`n" +
                             'Dieses Werkzeug gehört auf einen Exchange-Server oder einen Rechner mit den Exchange-Verwaltungswerkzeugen.')
        }
        try { Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop }
        catch { Stop-WithReason "Das Exchange-Snap-in ließ sich nicht laden: $($_.Exception.Message)" }
    }

    # 3. Snap-in geladen heißt noch nicht berechtigt: auch on-premises blendet RBAC
    #    Cmdlets aus, für die die Rolle nicht reicht.
    $fehlt = @($Cmdlets | Where-Object { -not (Get-Command $_ -ErrorAction SilentlyContinue) })
    if ($fehlt) {
        Stop-WithReason ("Die Exchange-Verwaltungsshell ist geladen, aber diese Cmdlets fehlen:`n" +
                         ("    " + ($fehlt -join "`n    ")) +
                         "`n`nDas ist eine Frage der RBAC-Rolle, nicht der Installation.")
    }

    # 4. Exchange gibt Datumsangaben und Zahlen in der Sprache des Servers zurück.
    #    Ohne feste Kultur brechen Vergleiche und Parser auf deutschen Systemen.
    [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::InvariantCulture
}
Initialize-OnPremPrerequisite
# </prolog:onprem>

function ConvertTo-PsLiteral {
    # Identitäten können Apostrophe enthalten (O'Brien) - im Vorschlag sicher quoten.
    param([string]$Text)
    "'" + ($Text -replace "'", "''") + "'"
}

function New-Occupant {
    param($Database, $Typ, $Detail, $Name, $Identity, $Blockiert, $Befehl)
    [pscustomobject]@{
        Database  = $Database
        Typ       = $Typ
        Detail    = $Detail
        Name      = $Name
        Identity  = $Identity
        Blockiert = $Blockiert
        Befehl    = $Befehl
    }
}

function Get-DatabaseOccupant {
    <#
        Sammelt alles, was in einer Datenbank liegt. Exchange-Aufrufe nur über Get-Mailbox und
        Get-MailboxStatistics, damit der Test sie durch Attrappen ersetzen kann.
        $AllMailboxes: einmal organisationsweit gelesen, für Archive, deren primäres Postfach in
        einer anderen Datenbank liegt (ArchiveDatabase zeigt hierher).
    #>
    param(
        [Parameter(Mandatory)][string]$Database,
        [object[]]$AllMailboxes = @(),
        [bool]$Mounted = $true
    )
    $funde = New-Object System.Collections.Generic.List[object]

    # Postfächer ohne Schalter. Discovery-Postfächer kommen hier mit und werden nur einmal
    # gezählt - unter ihrem eigenen Typ.
    foreach ($m in @(Get-Mailbox -Database $Database -ResultSize Unlimited)) {
        $id = ConvertTo-PsLiteral "$($m.Identity)"
        if ("$($m.RecipientTypeDetails)" -eq 'DiscoveryMailbox') {
            $funde.Add((New-Occupant $Database 'Discovery' "$($m.RecipientTypeDetails)" $m.DisplayName "$($m.Identity)" 'ja' "Remove-Mailbox -Identity $id"))
        } else {
            $funde.Add((New-Occupant $Database 'Postfach' "$($m.RecipientTypeDetails)" $m.DisplayName "$($m.Identity)" 'ja' "nach Exchange Online migrieren; wenn verzichtbar: Disable-Mailbox -Identity $id"))
        }
    }

    $sonder = @(
        @{ Schalter = 'Arbitration';  Typ = 'Arbitration';  Blockiert = 'ja';     Befehl = 'Remove-Mailbox -Identity {0} -Arbitration -RemoveArbitrationMailboxWithOABsAllowed -RemoveLastArbitrationMailboxAllowed' }
        @{ Schalter = 'AuditLog';     Typ = 'AuditLog';     Blockiert = 'ja';     Befehl = 'Remove-Mailbox -Identity {0} -AuditLog' }
        @{ Schalter = 'PublicFolder'; Typ = 'PublicFolder'; Blockiert = 'ja';     Befehl = 'Remove-Mailbox -Identity {0} -PublicFolder   (erst nach der Public-Folder-Migration)' }
        # Health-Postfächer stehen nicht in Microsofts Liste der Blocker; ob sie stören, ist offen.
        @{ Schalter = 'Monitoring';   Typ = 'Monitoring';   Blockiert = 'unklar'; Befehl = 'Disable-Mailbox -Identity {0}   (Health-Postfach)' }
    )
    foreach ($s in $sonder) {
        $p = @{ Database = $Database; ResultSize = 'Unlimited'; $s.Schalter = $true }
        foreach ($m in @(Get-Mailbox @p)) {
            $funde.Add((New-Occupant $Database $s.Typ "$($m.RecipientTypeDetails)" $m.DisplayName "$($m.Identity)" $s.Blockiert ($s.Befehl -f (ConvertTo-PsLiteral "$($m.Identity)"))))
        }
    }

    # Archive: beide Wege wie in Microsofts Anleitung, doppelte Treffer nur einmal.
    $archive = @(Get-Mailbox -Archive -Database $Database -ResultSize Unlimited)
    $archive += @($AllMailboxes | Where-Object { "$($_.ArchiveDatabase)" -eq $Database })
    $gesehen = @{}
    foreach ($m in $archive) {
        $key = "$($m.Identity)"
        if ($gesehen.ContainsKey($key)) { continue }
        $gesehen[$key] = $true
        $funde.Add((New-Occupant $Database 'Archiv' "$($m.RecipientTypeDetails)" $m.DisplayName $key 'ja' "Disable-Mailbox -Identity $(ConvertTo-PsLiteral $key) -Archive"))
    }

    # Getrennte Postfächer liest nur Get-MailboxStatistics - und das nur bei eingehängter DB.
    if ($Mounted) {
        foreach ($st in @(Get-MailboxStatistics -Database $Database | Where-Object { "$($_.DisconnectReason)" -in 'SoftDeleted', 'Disabled' })) {
            $funde.Add((New-Occupant $Database "$($st.DisconnectReason)" '' $st.DisplayName "$($st.MailboxGuid)" 'nein' 'blockiert nicht; Inhalt geht mit der Datenbank verloren (bei Bedarf vorher New-MailboxRestoreRequest)'))
        }
    }

    # Einzeln in die Pipeline - der Aufrufer sammelt mit @(...).
    $funde.ToArray()
}

# ------------------------------------------------------------------------------------------
# Ablauf
# ------------------------------------------------------------------------------------------

# Bei mehreren Domänen sonst nur die eigene - Postfächer anderer Domänen fehlten in der Liste.
Set-ADServerSettings -ViewEntireForest $true

$alleDbs = @(Get-MailboxDatabase -Status)
if ($Database) {
    $unbekannt = @($Database | Where-Object { $_ -notin @($alleDbs.Name) })
    if ($unbekannt) {
        Write-Warning ("Unbekannte Datenbank: {0}. Vorhanden: {1}" -f ($unbekannt -join ', '), (@($alleDbs.Name) -join ', '))
        exit 2
    }
    $wahl = @($alleDbs | Where-Object { $_.Name -in $Database })
} else {
    if (-not (Get-Command Out-GridView -ErrorAction SilentlyContinue)) {
        Write-Warning 'Out-GridView ist hier nicht verfügbar (etwa Server Core). Datenbank mit -Database angeben.'
        exit 2
    }
    $namen = @($alleDbs | Select-Object Name, Server, Mounted |
               Out-GridView -PassThru -Title 'Datenbank(en) zum Prüfen wählen (Mehrfachauswahl mit Strg)')
    if (-not $namen) {
        Write-Warning 'Keine Datenbank gewählt - nichts geprüft.'
        exit 2
    }
    $wahl = @($alleDbs | Where-Object { $_.Name -in @($namen.Name) })
}

# Einmal für alle Datenbanken: Archive, deren Hauptpostfach woanders liegt.
$alleMbx = @(Get-Mailbox -ResultSize Unlimited | Where-Object { $_.ArchiveDatabase })

$ergebnis = @()
foreach ($db in $wahl) {
    Write-Host ''
    Write-Host "Prüfe Datenbank: $($db.Name)" -ForegroundColor Cyan
    $mounted = [bool]$db.Mounted
    if (-not $mounted) {
        Write-Warning "$($db.Name) ist nicht eingehängt: getrennte Postfächer (soft-deleted, disabled) lassen sich nicht prüfen."
    }

    $funde = @(Get-DatabaseOccupant -Database $db.Name -AllMailboxes $alleMbx -Mounted $mounted)
    $ergebnis += $funde
    $block = @($funde | Where-Object { $_.Blockiert -ne 'nein' })

    if (-not $funde) {
        Write-Host '  Leer. Die Datenbank kann entfernt werden.' -ForegroundColor Green
        continue
    }
    if (-not $block) {
        Write-Host '  Keine blockierenden Postfächer. Nur getrennte Postfächer, deren Inhalt mit der Datenbank verloren geht:' -ForegroundColor Green
    } else {
        Write-Host ("  Noch {0} blockierende(s) Postfach/Postfächer:" -f $block.Count) -ForegroundColor Yellow
    }
    $funde | Sort-Object Blockiert, Typ, Name | Format-Table Typ, Detail, Name, Blockiert -AutoSize | Out-Host
    if ($block) {
        Write-Host '  Vorschläge (nichts wurde ausgeführt):'
        $block | ForEach-Object { Write-Host "    $($_.Befehl)" }
    }
}

if ($PassThru) { $ergebnis }
if (@($ergebnis | Where-Object { $_.Blockiert -ne 'nein' }).Count -gt 0) { exit 1 } else { exit 0 }
