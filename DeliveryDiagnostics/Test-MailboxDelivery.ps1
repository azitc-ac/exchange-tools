<#
.SYNOPSIS
    Diagnose für "SMTP delivery to mailbox schlägt fehl / socket error", wenn es nur mit
    Datenbanken auf einem bestimmten Server auftritt. Vergleicht den kaputten Server mit dem
    funktionierenden - rein lesend, ändert nichts.

.DESCRIPTION
    Zustellweg in Exchange 2013+: Transportdienst -> SMTP an Port 475 (Mailbox Transport
    Delivery) DES SERVERS MIT DER AKTIVEN DATENBANKKOPIE. Diese Verbindung ist per
    X-ANONYMOUSTLS verschlüsselt und benutzt das INTERNE TRANSPORTZERTIFIKAT des Zielservers.
    Ist dieser Pfad gestört, bleibt die Queue stehen - aber nur, solange Datenbanken auf dem
    betroffenen Server aktiv sind.

    Geprüft wird je Server:
      1. Queues und deren letzter Fehler
      2. Dienste (Transport, Delivery, Submission)
      3. Port 475: lauscht er, ist er erreichbar, kommt ein SMTP-Dialog zustande,
         gelingt der TLS-Handshake (X-ANONYMOUSTLS)
      4. Internes Transportzertifikat: vorhanden, gültig, privater Schlüssel lesbar
      5. Back Pressure (Ressourcendruck), freier Plattenplatz
      6. Schannel-/TLS-Einstellungen und .NET-Strong-Crypto
      7. Netzwerkkarten-Offload (häufige Ursache für abbrechende Verbindungen)
      8. Ereignisprotokoll der letzten Stunden
    Zum Schluss werden die Unterschiede zwischen den Servern hervorgehoben.

.EXAMPLE
    .\Test-MailboxDelivery.ps1 -BadServer SRVA -GoodServer SRVB
    .\Test-MailboxDelivery.ps1 -BadServer SRVA -GoodServer SRVB -EventHours 24
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string]$BadServer,
    [string]$GoodServer,
    [int]$EventHours = 6,
    [switch]$SkipRemote          # nur lokal prüfbare Dinge (wenn kein WinRM zum Zielserver)
)
$ErrorActionPreference = 'Continue'
$script:Findings = New-Object Collections.ArrayList

function Section([string]$T) { Write-Host ''; Write-Host "=== $T ===" -ForegroundColor Cyan }
function Line([string]$T, [string]$Level = 'INFO') {
    $c = switch ($Level) { 'OK' { 'Green' } 'WARN' { 'Yellow' } 'BAD' { 'Red' } default { 'Gray' } }
    Write-Host "  $T" -ForegroundColor $c
    if ($Level -in 'WARN', 'BAD') { [void]$script:Findings.Add("[$Level] $T") }
}

if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
}
$servers = @($BadServer) + @($GoodServer | Where-Object { $_ })

#--------------------------------------------------------------------- Hilfsfunktionen
function Invoke-OnServer {
    # ArgumentList ist Pflicht: GetNewClosure() wirkt NUR lokal - bei Invoke-Command kämen die
    # Variablen des Aufrufers nicht mit, der Scriptblock liefe mit leeren Parametern.
    param([string]$Name, [scriptblock]$Script, [object[]]$ArgumentList = @())
    $local = ($Name -eq $env:COMPUTERNAME -or $Name -like "$env:COMPUTERNAME.*")
    if ($local) { return (& $Script @ArgumentList) }
    if ($SkipRemote) { return 'übersprungen (-SkipRemote)' }
    try { return (Invoke-Command -ComputerName $Name -ScriptBlock $Script -ArgumentList $ArgumentList -ErrorAction Stop) }
    catch { return "nicht abfragbar: $($_.Exception.Message)" }
}

function Test-DeliveryPort {
    <#
      Spricht Port 475 an wie der Transportdienst: Begrüßung lesen, EHLO, und prüfen ob
      X-ANONYMOUSTLS angeboten wird. Danach den TLS-Handshake tatsächlich versuchen.
    #>
    param([string]$Fqdn, [int]$Port = 475, [int]$TimeoutMs = 10000)
    $res = [ordered]@{ Reachable = $false; Banner = ''; AnonymousTls = $false; TlsHandshake = ''; Error = '' }
    $tcp = New-Object Net.Sockets.TcpClient
    try {
        if (-not $tcp.ConnectAsync($Fqdn, $Port).Wait($TimeoutMs)) { $res.Error = 'Zeitüberschreitung beim Verbinden'; return [pscustomobject]$res }
        $res.Reachable = $true
        $stream = $tcp.GetStream(); $stream.ReadTimeout = $TimeoutMs
        $reader = New-Object IO.StreamReader($stream)
        $writer = New-Object IO.StreamWriter($stream); $writer.AutoFlush = $true
        $res.Banner = $reader.ReadLine()
        $writer.WriteLine("EHLO $env:COMPUTERNAME")
        $caps = @()
        while ($null -ne ($l = $reader.ReadLine())) {
            $caps += $l
            if ($l -match '^\d{3} ') { break }
        }
        $res.AnonymousTls = [bool]($caps -match 'X-ANONYMOUSTLS')
        if ($res.AnonymousTls) {
            $writer.WriteLine('X-ANONYMOUSTLS')
            $r = $reader.ReadLine()
            if ($r -match '^220') {
                try {
                    $ssl = New-Object Net.Security.SslStream($stream, $false, { $true })
                    $ssl.AuthenticateAsClient($Fqdn)
                    $cert = New-Object Security.Cryptography.X509Certificates.X509Certificate2($ssl.RemoteCertificate)
                    $res.TlsHandshake = "OK - Zertifikat $($cert.Thumbprint) ($($cert.Subject)), gültig bis $($cert.NotAfter.ToString('yyyy-MM-dd'))"
                } catch { $res.TlsHandshake = "FEHLGESCHLAGEN: $($_.Exception.Message)" }
            } else { $res.TlsHandshake = "X-ANONYMOUSTLS abgelehnt: $r" }
        }
    } catch { $res.Error = $_.Exception.Message }
    finally { $tcp.Close() }
    return [pscustomobject]$res
}

#--------------------------------------------------------------------- 1. Queues
Section '1. Warteschlangen und letzter Fehler'
foreach ($s in $servers) {
    Write-Host "$s" -ForegroundColor White
    $q = @(Get-Queue -Server $s -ErrorAction SilentlyContinue | Where-Object { $_.MessageCount -gt 0 -or $_.Status -ne 'Ready' })
    if (-not $q.Count) { Line 'keine auffälligen Warteschlangen' 'OK' }
    foreach ($item in $q) {
        $lvl = if ($item.LastError) { 'BAD' } else { 'INFO' }
        Line ("{0} | Status={1} | Nachrichten={2} | NextHop={3}" -f $item.Identity, $item.Status, $item.MessageCount, $item.NextHopDomain) $lvl
        if ($item.LastError) { Line ("   letzter Fehler: {0}" -f $item.LastError) 'BAD' }
    }
}

#--------------------------------------------------------------------- 2. Dienste
Section '2. Transportdienste'
foreach ($s in $servers) {
    Write-Host "$s" -ForegroundColor White
    $svc = Invoke-OnServer $s { Get-Service MSExchangeTransport, MSExchangeDelivery, MSExchangeSubmission, MSExchangeIS -ErrorAction SilentlyContinue |
                                 Select-Object Name, Status, StartType }
    if ($svc -is [string]) { Line $svc 'WARN' }
    else { foreach ($x in $svc) { Line ("{0,-24} {1}" -f $x.Name, $x.Status) $(if ($x.Status -ne 'Running') { 'BAD' } else { 'OK' }) } }
}

#--------------------------------------------------------------------- 3. Port 475
Section '3. Zustellport 475 (Mailbox Transport Delivery)'
foreach ($s in $servers) {
    $fqdn = (Get-ExchangeServer $s).Fqdn
    Write-Host "$s ($fqdn)" -ForegroundColor White
    $listen = Invoke-OnServer $s {
        @(Get-NetTCPConnection -LocalPort 475 -State Listen -ErrorAction SilentlyContinue |
          Select-Object LocalAddress, @{n='Process';e={(Get-Process -Id $_.OwningProcess -ErrorAction SilentlyContinue).Name}})
    }
    if ($listen -is [string]) { Line $listen 'WARN' }
    elseif (-not $listen) { Line 'NIEMAND lauscht auf Port 475 - MSExchangeDelivery läuft nicht oder ist blockiert' 'BAD' }
    else { foreach ($x in $listen) { Line ("lauscht auf {0} (Prozess {1})" -f $x.LocalAddress, $x.Process) 'OK' } }

    $probe = Test-DeliveryPort -Fqdn $fqdn
    if (-not $probe.Reachable) { Line ("Port 475 von hier NICHT erreichbar: {0}" -f $probe.Error) 'BAD' }
    else {
        Line ("Verbindung steht, Begrüßung: {0}" -f $probe.Banner) 'OK'
        if (-not $probe.AnonymousTls) { Line 'X-ANONYMOUSTLS wird NICHT angeboten - genau das braucht die interne Zustellung' 'BAD' }
        else { Line ("TLS-Handshake: {0}" -f $probe.TlsHandshake) $(if ($probe.TlsHandshake -like 'OK*') { 'OK' } else { 'BAD' }) }
    }
}

#--------------------------------------------------------------------- 4. Internes Transportzertifikat
Section '4. Internes Transportzertifikat (wird für X-ANONYMOUSTLS benutzt)'
foreach ($s in $servers) {
    Write-Host "$s" -ForegroundColor White
    $srv = Get-ExchangeServer $s
    $thumb = $null
    try {
        $de = New-Object DirectoryServices.DirectoryEntry("LDAP://$($srv.DistinguishedName)")
        $raw = $de.Properties['msExchServerInternalTLSCert'].Value
        if ($raw) {
            $coll = New-Object Security.Cryptography.X509Certificates.X509Certificate2Collection
            $coll.Import([byte[]]$raw)
            if ($coll.Count) {
                $c = @($coll)[0]; $thumb = $c.Thumbprint
                $days = [math]::Floor(($c.NotAfter - (Get-Date)).TotalDays)
                $lvl = if ($days -lt 0) { 'BAD' } elseif ($days -lt 30) { 'WARN' } else { 'OK' }
                Line ("laut AD: {0} | {1} | gültig bis {2} ({3} Tage)" -f $c.Thumbprint, $c.Subject, $c.NotAfter.ToString('yyyy-MM-dd'), $days) $lvl
            }
        } else { Line 'AD-Attribut msExchServerInternalTLSCert ist leer' 'BAD' }
    } catch { Line "AD-Abfrage fehlgeschlagen: $($_.Exception.Message)" 'WARN' }

    if ($thumb) {
        $store = Get-ExchangeCertificate -Server $s -Thumbprint $thumb -ErrorAction SilentlyContinue
        if (-not $store) { Line 'Dieses Zertifikat liegt NICHT im Speicher des Servers - die interne TLS-Aushandlung kann so nicht gelingen' 'BAD' }
        else {
            Line ("im Speicher: Status={0} | Dienste={1} | PrivateKey={2}" -f $store.Status, $store.Services, $store.HasPrivateKey) `
                 $(if ($store.Status -ne 'Valid' -or -not $store.HasPrivateKey) { 'BAD' } else { 'OK' })
            if ($store.Services -notmatch 'SMTP') { Line 'Der SMTP-Dienst ist für dieses Zertifikat NICHT aktiviert' 'BAD' }
            # Zugriff auf den privaten Schlüssel (ACL) prüfen - reiner Lesetest
            $keyCheck = Invoke-OnServer $s {
                param($tp)
                try {
                    $c = Get-Item "Cert:\LocalMachine\My\$tp" -ErrorAction Stop
                    if (-not $c.HasPrivateKey) { return 'kein privater Schlüssel am Zertifikat' }
                    $null = $c.PrivateKey
                    return 'privater Schlüssel lesbar'
                } catch { return "privater Schlüssel NICHT lesbar: $($_.Exception.Message)" }
            } -ArgumentList @($thumb)
            if ($keyCheck -is [string]) { Line $keyCheck $(if ($keyCheck -like 'privater Schlüssel lesbar*') { 'OK' } else { 'BAD' }) }
        }
    }
}

#--------------------------------------------------------------------- 5. Ressourcendruck
Section '5. Ressourcendruck (Back Pressure) und Plattenplatz'
foreach ($s in $servers) {
    Write-Host "$s" -ForegroundColor White
    $bp = Invoke-OnServer $s {
        $ev = Get-EventLog -LogName Application -Source MSExchangeTransport -Newest 200 -ErrorAction SilentlyContinue |
              Where-Object { $_.EventID -in 15004, 15005, 15006, 15007 } | Select-Object -First 3
        $disk = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" |
                Select-Object DeviceID, @{n='FreeGB';e={[math]::Round($_.FreeSpace/1GB,1)}}, @{n='PctFree';e={[math]::Round(100*$_.FreeSpace/$_.Size,1)}}
        [pscustomobject]@{ Events = $ev; Disks = $disk }
    }
    if ($bp -is [string]) { Line $bp 'WARN'; continue }
    foreach ($d in $bp.Disks) {
        $lvl = if ($d.PctFree -lt 10) { 'BAD' } elseif ($d.PctFree -lt 20) { 'WARN' } else { 'OK' }
        Line ("{0} frei: {1} GB ({2} %)" -f $d.DeviceID, $d.FreeGB, $d.PctFree) $lvl
    }
    if ($bp.Events) { foreach ($e in $bp.Events) { Line ("Back-Pressure-Ereignis {0}: {1}" -f $e.EventID, ($e.Message -split "`n")[0]) 'BAD' } }
    else { Line 'keine Back-Pressure-Ereignisse' 'OK' }
}

#--------------------------------------------------------------------- 6. TLS-Einstellungen
Section '6. Schannel/TLS und .NET-Strong-Crypto (Unterschiede sind hier besonders verdächtig)'
$tlsData = @{}
foreach ($s in $servers) {
    Write-Host "$s" -ForegroundColor White
    $tls = Invoke-OnServer $s {
        $out = [ordered]@{}
        $base = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols'
        foreach ($p in 'TLS 1.0', 'TLS 1.1', 'TLS 1.2', 'TLS 1.3') {
            foreach ($r in 'Client', 'Server') {
                $k = Join-Path $base "$p\$r"
                if (Test-Path $k) {
                    $v = Get-ItemProperty $k -ErrorAction SilentlyContinue
                    $out["$p/$r"] = "Enabled=$($v.Enabled) DisabledByDefault=$($v.DisabledByDefault)"
                } else { $out["$p/$r"] = 'kein Schlüssel (Standard)' }
            }
        }
        foreach ($k in 'HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319') {
            $v = Get-ItemProperty $k -ErrorAction SilentlyContinue
            $out[$k] = "SchUseStrongCrypto=$($v.SchUseStrongCrypto) SystemDefaultTlsVersions=$($v.SystemDefaultTlsVersions)"
        }
        $out
    }
    if ($tls -is [string]) { Line $tls 'WARN'; continue }
    $tlsData[$s] = $tls
    foreach ($k in $tls.Keys) { Line ("{0,-28} {1}" -f $k, $tls[$k]) }
}

#--------------------------------------------------------------------- 7. Netzwerkkarte
Section '7. Netzwerkkarten-Offload (häufige Ursache abbrechender Verbindungen)'
foreach ($s in $servers) {
    Write-Host "$s" -ForegroundColor White
    $nic = Invoke-OnServer $s {
        Get-NetAdapter -Physical | Where-Object Status -eq 'Up' | ForEach-Object {
            $a = $_
            $adv = Get-NetAdapterAdvancedProperty -Name $a.Name -ErrorAction SilentlyContinue |
                   Where-Object { $_.DisplayName -match 'Offload|RSS|Jumbo|Large Send|Checksum|Flow Control|Receive Side' }
            [pscustomobject]@{
                Name = $a.Name; Speed = $a.LinkSpeed; MTU = $a.MtuSize
                Props = ($adv | ForEach-Object { "$($_.DisplayName)=$($_.DisplayValue)" })
            }
        }
    }
    if ($nic -is [string]) { Line $nic 'WARN'; continue }
    foreach ($n in $nic) {
        Line ("{0} | {1} | MTU {2}" -f $n.Name, $n.Speed, $n.MTU)
        foreach ($p in $n.Props) { Line ("   $p") }
    }
}

#--------------------------------------------------------------------- 8. Ereignisse
Section "8. Ereignisprotokoll der letzten $EventHours Stunden (Transport/Delivery/Schannel)"
foreach ($s in $servers) {
    Write-Host "$s" -ForegroundColor White
    $ev = Invoke-OnServer $s {
        param($h)
        Get-WinEvent -FilterHashtable @{ LogName = 'Application', 'System'; StartTime = (Get-Date).AddHours(-$h); Level = 1, 2, 3 } -ErrorAction SilentlyContinue |
            Where-Object { $_.ProviderName -match 'MSExchangeTransport|MSExchangeDelivery|MSExchangeSubmission|Schannel|MSExchange Common' } |
            Select-Object -First 15 TimeCreated, ProviderName, Id, LevelDisplayName, @{n='Msg';e={ ($_.Message -split "`n")[0].Trim() }}
    } -ArgumentList @($EventHours)
    if ($ev -is [string]) { Line $ev 'WARN'; continue }
    if (-not $ev) { Line 'keine passenden Ereignisse' 'OK'; continue }
    # Gleiche Ereignis-ID nur einmal melden, sonst erschlaegt Routine-Rauschen die echten Funde.
    $seen = @{}
    foreach ($e in $ev) {
        $key = "$($e.ProviderName)/$($e.Id)"
        if ($seen.ContainsKey($key)) { $seen[$key]++; continue }
        $seen[$key] = 1
        $lvl = if ($e.LevelDisplayName -match 'Fehler|Error') { 'BAD' } else { 'WARN' }
        # "Zertifikat laeuft bald ab" ist ein Hinweis, kein Stoerungsgrund
        if ($e.Id -eq 12017) { $lvl = 'INFO' }
        Line ("{0:HH:mm} {1} ({2}): {3}" -f $e.TimeCreated, $e.ProviderName, $e.Id, $e.Msg) $lvl
    }
    foreach ($k in ($seen.Keys | Where-Object { $seen[$_] -gt 1 })) { Line ("   ({0} weitere Ereignisse von {1})" -f ($seen[$k] - 1), $k) }
}

#--------------------------------------------------------------------- Unterschiede
if ($GoodServer -and $tlsData.Count -eq 2) {
    Section 'Unterschiede der TLS-Einstellungen zwischen beiden Servern'
    $a = $tlsData[$BadServer]; $b = $tlsData[$GoodServer]
    $diff = $false
    foreach ($k in $a.Keys) {
        if ("$($a[$k])" -ne "$($b[$k])") {
            Line ("{0}:`n     {1}: {2}`n     {3}: {4}" -f $k, $BadServer, $a[$k], $GoodServer, $b[$k]) 'WARN'
            $diff = $true
        }
    }
    if (-not $diff) { Line 'keine Unterschiede in den TLS-Einstellungen' 'OK' }
}

Section 'Zusammenfassung der Auffälligkeiten'
if (-not $script:Findings.Count) { Write-Host '  nichts Auffälliges gefunden' -ForegroundColor Green }
else { foreach ($f in $script:Findings) { Write-Host "  $f" -ForegroundColor $(if ($f -like '`[BAD`]*') { 'Red' } else { 'Yellow' }) } }
Write-Host ''
Write-Host 'Nächster Schritt, falls Port 475 erreichbar ist und der TLS-Handshake scheitert:' -ForegroundColor Cyan
Write-Host '  Das interne Transportzertifikat des betroffenen Servers neu erzeugen und zuweisen'
Write-Host '  (Abschnitt 4 zeigt, ob es fehlt, abgelaufen ist oder der private Schlüssel klemmt).'
Write-Host 'Falls die Verbindung gar nicht zustande kommt: Virenschutz/EDR und lokale Firewall auf dem'
Write-Host '  betroffenen Server prüfen - Port 475 wird gern von Mail-Scannern abgefangen.'
