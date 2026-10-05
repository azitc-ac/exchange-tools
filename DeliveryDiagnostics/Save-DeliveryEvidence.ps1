<#
.SYNOPSIS
    Sichert im Störungsfall den LAUFZEITZUSTAND der Exchange-Zustellung - bevor jemand neu startet.

.DESCRIPTION
    Gedacht für den Fall "Zustellung ins Postfach hängt, Konfiguration sieht aber sauber aus".
    Ein Neustart räumt genau die Spuren weg, die man braucht. Dieses Skript sammelt sie in
    wenigen Sekunden ein - rein lesend, es ändert und startet nichts.

    Gesammelt wird:
      1. Warteschlangen samt letztem Fehler
      2. Transportprozesse: Handles, Threads, Speicher, Laufzeit (Lecks und Neustarts erkennen)
      3. TCP-Zustand: Verbindungen nach Status, Auslastung des dynamischen Portbereichs,
         AUSGESCHLOSSENE Portbereiche (Hyper-V/WinNAT reserviert gern Bereiche und bricht damit
         ausgehende Verbindungen - verschwindet nach Reboot, kommt wieder)
      4. Verbindungen auf Port 475 und ausgehende Verbindungen der Transportprozesse
      5. Connectivity- und Protokollprotokolle der Zustellung (zeigen den echten Abbruch)
      6. Ereignisprotokoll
      7. Dienste und deren Startzeit

    WICHTIG für die Eingrenzung: Vor einem Server-Neustart erst die DIENSTE neu starten
    (MSExchangeTransport, MSExchangeDelivery). Hilft das schon, liegt es im Dienst (Leck,
    hängende Worker). Hilft nur der Reboot, liegt es im Netzwerkstack oder Systemzustand.

.EXAMPLE
    .\Save-DeliveryEvidence.ps1
    .\Save-DeliveryEvidence.ps1 -Server SRVA -OutFolder D:\Temp -EventHours 12
#>
[CmdletBinding()]
param(
    [string]$Server = $env:COMPUTERNAME,
    [string]$OutFolder = (Join-Path $env:TEMP 'DeliveryEvidence'),
    [int]$EventHours = 6,
    [int]$LogTailLines = 60
)
$ErrorActionPreference = 'Continue'

if (-not (Test-Path $OutFolder)) { New-Item -ItemType Directory -Path $OutFolder -Force | Out-Null }
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outFile = Join-Path $OutFolder ("DeliveryEvidence_{0}_{1}.txt" -f $Server, $stamp)
$script:Buffer = New-Object Text.StringBuilder

function Out-Both {
    param([string]$Text, [string]$Color = 'Gray')
    Write-Host $Text -ForegroundColor $Color
    [void]$script:Buffer.AppendLine($Text)
}
function Section([string]$T) { Out-Both ''; Out-Both ("=== $T ===") 'Cyan' }
function Dump($Object) {
    if ($null -eq $Object) { Out-Both '  (nichts)'; return }
    $text = ($Object | Format-Table -AutoSize -Wrap | Out-String).TrimEnd()
    if (-not $text) { $text = ($Object | Out-String).TrimEnd() }
    foreach ($l in ($text -split "`r?`n")) { Out-Both "  $l" }
}

$local = ($Server -eq $env:COMPUTERNAME -or $Server -like "$env:COMPUTERNAME.*")
function Remote([scriptblock]$Script, [object[]]$ArgumentList = @()) {
    # ArgumentList statt Closure: bei Invoke-Command kaemen lokale Variablen sonst nicht mit.
    if ($local) { return (& $Script @ArgumentList) }
    try { return (Invoke-Command -ComputerName $Server -ScriptBlock $Script -ArgumentList $ArgumentList -ErrorAction Stop) }
    catch { return "nicht abfragbar: $($_.Exception.Message)" }
}

Out-Both ("Beweissicherung Zustellung - {0} - {1}" -f $Server, (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) 'White'
Out-Both ("Erhoben von {0} auf {1}" -f $env:USERNAME, $env:COMPUTERNAME)

#--------------------------------------------------------------- 1. Warteschlangen
Section '1. Warteschlangen'
# Voraussetzungen. Der Block zwischen den Markern stammt aus build/Prolog.OnPrem.ps1
# und wird von build/Sync-Prolog.ps1 gepflegt - nicht von Hand aendern.
$script:RequiredCmdlets = @('Get-Queue', 'Get-TransportService', 'Get-MailboxTransportService')
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
try {
    $q = @(Get-Queue -Server $Server -ErrorAction Stop)
    Dump ($q | Select-Object Identity, Status, MessageCount, NextHopDomain, DeliveryType, LastRetryTime)
    foreach ($item in ($q | Where-Object { $_.LastError })) {
        Out-Both ("  LETZTER FEHLER [{0}]: {1}" -f $item.Identity, $item.LastError) 'Red'
    }
    Out-Both ("  Nachrichten gesamt in Warteschlangen: {0}" -f (($q | Measure-Object MessageCount -Sum).Sum))
} catch { Out-Both "  Get-Queue fehlgeschlagen: $($_.Exception.Message)" 'Yellow' }

#--------------------------------------------------------------- 2. Prozesse
Section '2. Transportprozesse (Lecks und Laufzeit)'
$proc = Remote {
    Get-Process MSExchangeTransport, MSExchangeDelivery, MSExchangeSubmission, MSExchangeFrontendTransport, EdgeTransport -ErrorAction SilentlyContinue |
        Select-Object Name, Id, Handles,
            @{n='Threads';e={$_.Threads.Count}},
            @{n='PrivateMB';e={[math]::Round($_.PrivateMemorySize64/1MB)}},
            @{n='WorkingMB';e={[math]::Round($_.WorkingSet64/1MB)}},
            @{n='CPUsek';e={[math]::Round($_.CPU)}},
            StartTime,
            @{n='LaeuftSeitStd';e={[math]::Round(((Get-Date) - $_.StartTime).TotalHours,1)}}
}
Dump $proc
Out-Both '  Richtwerte: Handles pro Prozess dauerhaft > 20000 oder PrivateMB > 4000 sind verdaechtig;'
Out-Both '  eine sehr kurze Laufzeit verraet einen unbemerkten Dienstneustart (Absturz).'

#--------------------------------------------------------------- 3. TCP-Zustand
Section '3. TCP-Zustand und Portbereiche'
$tcp = Remote {
    $all = Get-NetTCPConnection -ErrorAction SilentlyContinue
    [pscustomobject]@{
        ByState  = ($all | Group-Object State | Select-Object Name, Count | Sort-Object Count -Descending)
        Total    = $all.Count
        TimeWait = @($all | Where-Object State -eq 'TimeWait').Count
    }
}
if ($tcp -is [string]) { Out-Both "  $tcp" 'Yellow' }
else {
    Out-Both ("  Verbindungen gesamt: {0}, davon TimeWait: {1}" -f $tcp.Total, $tcp.TimeWait) `
             $(if ($tcp.TimeWait -gt 10000) { 'Red' } elseif ($tcp.TimeWait -gt 4000) { 'Yellow' } else { 'Gray' })
    Dump $tcp.ByState
}

$ports = Remote {
    [pscustomobject]@{
        Dynamic   = (netsh int ipv4 show dynamicport tcp | Out-String).Trim()
        Excluded  = (netsh int ipv4 show excludedportrange tcp | Out-String).Trim()
        Dynamic6  = (netsh int ipv6 show dynamicport tcp | Out-String).Trim()
        Excluded6 = (netsh int ipv6 show excludedportrange tcp | Out-String).Trim()
    }
}
if ($ports -is [string]) { Out-Both "  $ports" 'Yellow' }
else {
    Out-Both '  --- dynamischer Portbereich (IPv4) ---'
    foreach ($l in ($ports.Dynamic -split "`r?`n")) { Out-Both "  $l" }
    Out-Both '  --- AUSGESCHLOSSENE Portbereiche (IPv4) ---'
    foreach ($l in ($ports.Excluded -split "`r?`n")) { Out-Both "  $l" }
    Out-Both '  --- dynamischer Portbereich (IPv6) ---'
    foreach ($l in ($ports.Dynamic6 -split "`r?`n")) { Out-Both "  $l" }
    Out-Both '  --- AUSGESCHLOSSENE Portbereiche (IPv6) ---'
    foreach ($l in ($ports.Excluded6 -split "`r?`n")) { Out-Both "  $l" }
    Out-Both '  Hinweis: Ueberschneiden sich ausgeschlossene Bereiche mit dem dynamischen Bereich,'
    Out-Both '  fehlen dem Transportdienst Quellports - typischerweise nach einem Hyper-V/WinNAT-Start.'
}

#--------------------------------------------------------------- 4. Port 475
Section '4. Verbindungen der Zustellung (Port 475)'
$conn = Remote {
    $l = @(Get-NetTCPConnection -LocalPort 475 -ErrorAction SilentlyContinue |
           Group-Object State | Select-Object Name, Count)
    $procIds = @(Get-Process MSExchangeTransport, MSExchangeDelivery -ErrorAction SilentlyContinue | Select-Object -Expand Id)
    $out = @(Get-NetTCPConnection -ErrorAction SilentlyContinue |
             Where-Object { $procIds -contains $_.OwningProcess } |
             Group-Object State | Select-Object Name, Count)
    [pscustomobject]@{ Port475 = $l; ByTransport = $out; TransportPids = ($procIds -join ', ') }
}
if ($conn -is [string]) { Out-Both "  $conn" 'Yellow' }
else {
    Out-Both '  Verbindungen auf Port 475 nach Status:'
    Dump $conn.Port475
    Out-Both ("  Verbindungen der Transportprozesse (PID {0}) nach Status:" -f $conn.TransportPids)
    Dump $conn.ByTransport
}

#--------------------------------------------------------------- 5. Protokolle
Section '5. Protokollprotokolle der Zustellung (die eigentliche Fehlerquelle)'
try {
    $mts = Get-MailboxTransportService $Server -ErrorAction Stop
    $ts  = Get-TransportService $Server -ErrorAction SilentlyContinue
    $paths = @(
        @{ Name = 'Delivery Connectivity'; Path = $mts.ConnectivityLogPath }
        @{ Name = 'Delivery SendProtocol'; Path = $mts.SendProtocolLogPath }
        @{ Name = 'Delivery ReceiveProtocol'; Path = $mts.ReceiveProtocolLogPath }
        @{ Name = 'Transport Connectivity'; Path = $ts.ConnectivityLogPath }
    )
    foreach ($p in $paths) {
        if (-not $p.Path) { continue }
        $unc = "$($p.Path)" -replace '^([A-Za-z]):', "\\$Server\`$1`$"
        Out-Both ("  --- {0}: {1}" -f $p.Name, $p.Path)
        try {
            $newest = Get-ChildItem $unc -Filter *.log -ErrorAction Stop | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if (-not $newest) { Out-Both '      (keine Protokolldatei)'; continue }
            Out-Both ("      neueste Datei: {0} ({1:yyyy-MM-dd HH:mm})" -f $newest.Name, $newest.LastWriteTime)
            $tail = Get-Content $newest.FullName -Tail $LogTailLines -ErrorAction Stop
            foreach ($l in $tail) { Out-Both "      $l" }
        } catch { Out-Both "      nicht lesbar: $($_.Exception.Message)" 'Yellow' }
    }
} catch { Out-Both "  Protokollpfade nicht ermittelbar: $($_.Exception.Message)" 'Yellow' }

#--------------------------------------------------------------- 6. Ereignisse
Section "6. Ereignisse der letzten $EventHours Stunden"
$ev = Remote {
    param($h)
    Get-WinEvent -FilterHashtable @{ LogName = 'Application', 'System'; StartTime = (Get-Date).AddHours(-$h); Level = 1, 2, 3 } -ErrorAction SilentlyContinue |
        Where-Object { $_.ProviderName -match 'MSExchange|Schannel|Kerberos|LsaSrv|Tcpip|NETLOGON' } |
        Select-Object TimeCreated, ProviderName, Id, LevelDisplayName, @{n='Msg';e={ ($_.Message -split "`n")[0].Trim() }}
} @($EventHours)
if ($ev -is [string]) { Out-Both "  $ev" 'Yellow' }
elseif (-not $ev) { Out-Both '  keine passenden Ereignisse' }
else {
    $seen = @{}
    foreach ($e in $ev) {
        $k = "$($e.ProviderName)/$($e.Id)"
        if ($seen.ContainsKey($k)) { $seen[$k]++; continue }
        $seen[$k] = 1
        Out-Both ("  {0:yyyy-MM-dd HH:mm} {1} ({2}) {3}: {4}" -f $e.TimeCreated, $e.ProviderName, $e.Id, $e.LevelDisplayName, $e.Msg) `
                 $(if ($e.LevelDisplayName -match 'Fehler|Error') { 'Red' } else { 'Gray' })
    }
    foreach ($k in ($seen.Keys | Where-Object { $seen[$_] -gt 1 } | Sort-Object)) {
        Out-Both ("     ({0} weitere von {1})" -f ($seen[$k] - 1), $k)
    }
}

#--------------------------------------------------------------- 7. Dienste
Section '7. Dienste'
Dump (Remote {
    Get-CimInstance Win32_Service -Filter "Name LIKE 'MSExchange%'" |
        Where-Object { $_.Name -match 'Transport|Delivery|Submission|IS$' } |
        Select-Object Name, State, StartMode, ProcessId
})

#--------------------------------------------------------------- Ende
Section 'Nächste Schritte'
Out-Both '  1. NICHT sofort neu starten - erst diese Datei sichern.'
Out-Both '  2. Dann NUR die Dienste neu starten:'
Out-Both '       Restart-Service MSExchangeTransport, MSExchangeDelivery'
Out-Both '     Hilft das -> Ursache im Dienst (Leck, haengende Worker).'
Out-Both '     Hilft nur ein Reboot -> Ursache im Netzwerkstack/Systemzustand (Portbereiche, Treiber).'
Out-Both '  3. Diese Datei mit der Aufnahme eines fehlerfreien Zeitpunkts vergleichen -'
Out-Both '     die Differenz zeigt, was sich aufgebaut hat.'

[IO.File]::WriteAllText($outFile, $script:Buffer.ToString(), (New-Object Text.UTF8Encoding $true))
Write-Host ''
Write-Host "Gesichert in: $outFile" -ForegroundColor Green
