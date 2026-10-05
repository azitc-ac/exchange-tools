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
