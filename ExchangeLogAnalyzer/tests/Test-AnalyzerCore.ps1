<#
.SYNOPSIS
    Prueft den Parsing-Kern und den Zwischenspeicher des Log Analyzers.

.DESCRIPTION
    Legt kleine Protokolldateien mit von Hand nachgerechneten Erwartungswerten an und laesst
    den echten Code darauf laufen (geladen ohne Fenster). Geprueft werden:

      * Message Tracking - alle vier Sichten aus EINEM Lauf, samt der beiden Tuecken des
        Formats: ein Betreff mit Komma verschiebt jedes Feld dahinter, und ein Eintrag kann
        mehrere Empfaenger mit ';' tragen
      * SMTP Receive     - EHLO/HELO je Gegenstelle, Connector-Filter, IP aus remote-endpoint
      * IIS              - Treffer je Adresse und je Konto mit Statusklassen
      * Zeitraumfilter
      * Zwischenspeicher - zweite Abfrage ohne neues Lesen; nach Aenderung einer Datei wird
        wieder gelesen

.EXAMPLE
    powershell.exe -STA -ExecutionPolicy Bypass -File .\tests\Test-AnalyzerCore.ps1
#>
$ErrorActionPreference = 'Stop'

$root   = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path $MyInvocation.MyCommand.Path -Parent }
$script = Join-Path (Split-Path $root -Parent) 'Exchange Log Analyzer.ps1'
if (-not (Test-Path $script)) { throw "Skript nicht gefunden: $script" }

# ohne Fenster laden
$text = [IO.File]::ReadAllText($script).Replace('[void]$form.ShowDialog()', '# im Test kein Fenster')
$tmp  = Join-Path $env:TEMP ('analyzer-test-' + [Guid]::NewGuid().ToString('N') + '.ps1')
[IO.File]::WriteAllText($tmp, $text, (New-Object Text.UTF8Encoding $true))
. $tmp

$dir = Join-Path $env:TEMP ('analyzer-test-' + [Guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $dir -Force

$script:Fehler = 0
function Test-Wert {
    param([string]$Name, $Erwartet, $Erhalten)
    $ok = ("$Erwartet" -eq "$Erhalten")
    '{0} {1,-46} erwartet {2,-12} erhalten {3}' -f $(if ($ok) { '[OK]  ' } else { '[FEHL]' }), $Name, $Erwartet, $Erhalten
    if (-not $ok) { $script:Fehler++ }
}

# ── Testdaten: Message Tracking ──────────────────────────────────────────────────
#  Feldnummern: 0 date-time, 4 server-hostname, 8 event-id, 12 recipient-address,
#               14 total-bytes, 18 message-subject, 19 sender-address, 30 Felder gesamt
function New-TrackLine {
    param($Date, $Server, $Event, $Rcpt, $Bytes, $Subject, $Sender)
    @(
        $Date, '192.0.2.1', 'client.corp.example.com', '192.0.2.5', $Server, 'SMTP', "Default $Server",
        'SMTP', $Event, '1', '<m@corp.example.com>', 'net1', $Rcpt, '250 2.1.5', $Bytes, '1', '', '',
        $Subject, $Sender, $Sender, '', 'Originating', '', '192.0.2.9', '192.0.2.5', '', 'Email', '1', '15'
    ) -join ','
}
$trackFile = Join-Path $dir 'MSGTRK20260901-1.LOG'
$trackLines = @(
    '#Software: Microsoft Exchange Server'
    '#Fields: date-time,client-ip,client-hostname,server-ip,server-hostname,source-context,connector-id,source,event-id,internal-message-id,message-id,network-message-id,recipient-address,recipient-status,total-bytes,recipient-count,related-recipient-address,reference,message-subject,sender-address,return-path,message-info,directionality,tenant-id,original-client-ip,original-server-ip,custom-data,transport-traffic-type,log-id,schema-version'
    # Betreff MIT Komma -> alle Felder dahinter verschieben sich um eins
    (New-TrackLine '2026-09-01T08:00:00.000Z' 'EX01' 'RECEIVE' '"a@x.test;b@x.test"' 1000 '"Hallo, Welt"'          's1@x.test')
    (New-TrackLine '2026-09-01T09:00:00.000Z' 'EX01' 'SEND'    'c@x.test'             2000 'Ohne Komma'             's1@x.test')
    (New-TrackLine '2026-09-02T08:00:00.000Z' 'EX02' 'DELIVER' 'd@x.test'              500 'Auch ohne'              's2@x.test')
    (New-TrackLine '2026-09-02T09:00:00.000Z' 'EX02' 'RECEIVE' 'a@x.test'             3000 '"Re: Angebot, dringend"' 's2@x.test')
)
[IO.File]::WriteAllLines($trackFile, $trackLines, (New-Object Text.UTF8Encoding $false))

# ── Testdaten: SMTP Receive ──────────────────────────────────────────────────────
$smtpFile = Join-Path $dir 'RECV20260901-1.LOG'
$smtpLines = @(
    '#Software: Microsoft Exchange Server'
    '#Fields: date-time,connector-id,session-id,sequence-number,local-endpoint,remote-endpoint,event,data,context'
    '2026-09-01T08:00:00.000Z,EX01\Default Frontend,08DD,1,192.0.2.5:25,198.51.100.10:51000,>,"EHLO mail.partner.test",'
    '2026-09-01T08:00:01.000Z,EX01\Default Frontend,08DD,2,192.0.2.5:25,198.51.100.10:51001,>,"EHLO mail.partner.test",'
    '2026-09-01T08:00:02.000Z,EX01\Relay Connector,08DE,1,192.0.2.5:25,198.51.100.20:52000,>,"HELO altsystem",'
    '2026-09-02T08:00:03.000Z,EX01\Relay Connector,08DF,1,192.0.2.5:25,[2001:db8::5]:53000,>,"EHLO v6host",'
    '2026-09-01T08:00:04.000Z,EX01\Default Frontend,08E0,3,192.0.2.5:25,198.51.100.10:51002,>,"MAIL FROM:<a@x.test>",'
)
[IO.File]::WriteAllLines($smtpFile, $smtpLines, (New-Object Text.UTF8Encoding $false))

# ── Testdaten: IIS ───────────────────────────────────────────────────────────────
$iisFile = Join-Path $dir 'u_ex260901.log'
$iisLines = @(
    '#Software: Microsoft Internet Information Services 10.0'
    '#Fields: date time s-ip cs-method cs-uri-stem cs-uri-query s-port cs-username c-ip cs(User-Agent) cs(Referer) sc-status sc-substatus sc-win32-status time-taken'
    '2026-09-01 08:00:00 192.0.2.5 GET /owa - 443 EXAMPLE\anna 198.51.100.10 Mozilla/5.0 - 200 0 0 15'
    '2026-09-01 08:00:01 192.0.2.5 GET /owa - 443 EXAMPLE\anna 198.51.100.10 Mozilla/5.0 - 302 0 0 5'
    '2026-09-01 08:00:02 192.0.2.5 GET /ews - 443 EXAMPLE\bernd 198.51.100.11 Mozilla/5.0 - 401 0 0 2'
    '2026-09-02 08:00:03 192.0.2.5 GET /ews - 443 EXAMPLE\bernd 198.51.100.11 Mozilla/5.0 - 500 0 0 9'
)
[IO.File]::WriteAllLines($iisFile, $iisLines, (New-Object Text.UTF8Encoding $false))

# ══ Message Tracking ═════════════════════════════════════════════════════════════
'--- Message Tracking: alle vier Sichten aus einem Lauf ---'
$srv = @(Invoke-MessageTrackingQuery -Files @($trackFile) -UI $ui -Mode 'Server' -DateFrom '' -DateTo '')
$EX01 = $srv | Where-Object { $_.Servername -eq 'EX01' }
$EX02 = $srv | Where-Object { $_.Servername -eq 'EX02' }
Test-Wert 'Server EX01: gesendet'                   1    $EX01.SendCount
Test-Wert 'Server EX01: empfangen'                  1    $EX01.RecvCount
Test-Wert 'Server EX02: empfangen (DELIVER zaehlt nicht)' 1 $EX02.RecvCount
Test-Wert 'Server EX02: gesendet'                   0    $EX02.SendCount

$day = @(Invoke-MessageTrackingQuery -Files @($trackFile) -UI $ui -Mode 'Day' -DateFrom '' -DateTo '')
Test-Wert 'aus dem Zwischenspeicher (2. Abfrage)'   $true $script:LastFromCache
$d1 = $day | Where-Object { $_.Date -eq '2026-09-01' }
$d2 = $day | Where-Object { $_.Date -eq '2026-09-02' }
Test-Wert 'Tag 2026-09-01: gesendet/empfangen'      '1/1' "$($d1.SendCount)/$($d1.RecvCount)"
Test-Wert 'Tag 2026-09-02: gesendet/empfangen'      '0/1' "$($d2.SendCount)/$($d2.RecvCount)"

$rcp = @(Invoke-MessageTrackingQuery -Files @($trackFile) -UI $ui -Mode 'Recipient' -DateFrom '' -DateTo '')
$a = $rcp | Where-Object { $_.Recipient -eq 'a@x.test' }
$b = $rcp | Where-Object { $_.Recipient -eq 'b@x.test' }
$c = $rcp | Where-Object { $_.Recipient -eq 'c@x.test' }
$d = $rcp | Where-Object { $_.Recipient -eq 'd@x.test' }
Test-Wert 'Empfaenger a@x.test (RECEIVE 2x)'        2 $a.Count
Test-Wert 'Empfaenger b@x.test (aus Liste mit ;)'   1 $b.Count
Test-Wert 'Empfaenger c@x.test (SEND zaehlt nicht)' 0 ([int]$c.Count)
Test-Wert 'Empfaenger d@x.test (DELIVER zaehlt)'    1 $d.Count

# Der Absender steht HINTER dem Betreff - stimmt der Versatz nicht, faellt das hier auf
$snd = @(Invoke-MessageTrackingQuery -Files @($trackFile) -UI $ui -Mode 'Sender' -DateFrom '' -DateTo '')
$s1 = $snd | Where-Object { $_.Sender -eq 's1@x.test' }
$s2 = $snd | Where-Object { $_.Sender -eq 's2@x.test' }
Test-Wert 'Absender s1@x.test (Betreff mit Komma)'  1 $s1.Count
Test-Wert 'Absender s2@x.test (Betreff mit Komma)'  1 $s2.Count
Test-Wert 'keine unbekannten Absender'              2 $snd.Count

# ── Zeitraumfilter: anderer Schluessel, also frisch lesen ────────────────────────
''
'--- Zeitraumfilter ---'
$day2 = @(Invoke-MessageTrackingQuery -Files @($trackFile) -UI $ui -Mode 'Day' -DateFrom '2026-09-02' -DateTo '2026-09-02')
Test-Wert 'gefiltert: neu gelesen, nicht zwischengespeichert' $false $script:LastFromCache
Test-Wert 'gefiltert: nur ein Tag'                  1 $day2.Count
Test-Wert 'gefiltert: Tag ist 2026-09-02'           '2026-09-02' $day2[0].Date

# ── Zwischenspeicher faellt um, wenn die Datei waechst ──────────────────────────
''
'--- Zwischenspeicher nach Dateiaenderung ---'
$null = @(Invoke-MessageTrackingQuery -Files @($trackFile) -UI $ui -Mode 'Day' -DateFrom '' -DateTo '')
Test-Wert 'vor der Aenderung: aus dem Zwischenspeicher' $true $script:LastFromCache
Add-Content -LiteralPath $trackFile -Value (New-TrackLine '2026-09-03T08:00:00.000Z' 'EX01' 'RECEIVE' 'e@x.test' 4000 'Neu' 's3@x.test')
$day3 = @(Invoke-MessageTrackingQuery -Files @($trackFile) -UI $ui -Mode 'Day' -DateFrom '' -DateTo '')
Test-Wert 'nach der Aenderung: neu gelesen'         $false $script:LastFromCache
Test-Wert 'nach der Aenderung: drei Tage'           3 $day3.Count

# ── Fortschritt nach Bytes ──────────────────────────────────────────────────────
# Der Balken folgt den gelesenen Bytes. Am Ende eines Laufs muessen sich die gemeldeten
# Bytes genau auf die Summe der Dateigroessen addieren - sonst bleibt der Balken stehen
# oder schiesst ueber 100 hinaus.
''
'--- Fortschritt ---'
$soll = (Get-Item $trackFile).Length
Test-Wert 'TotalBytes = Groesse der Datei'        $soll ([AnalyzerCore]::TotalBytes)
Test-Wert 'DoneBytes am Ende = TotalBytes'        ([AnalyzerCore]::TotalBytes) ([AnalyzerCore]::DoneBytes)

# ══ SMTP Receive ═════════════════════════════════════════════════════════════════
''
'--- SMTP Receive ---'
$smtp = @(Invoke-SmtpReceiveHits -Files @($smtpFile) -UI $ui -ConnectorFilter '' -DateFrom '' -DateTo '' -ResolveDns $false)
$ip10 = $smtp | Where-Object { $_.IP -eq '198.51.100.10' }
$ip20 = $smtp | Where-Object { $_.IP -eq '198.51.100.20' }
$ipv6 = $smtp | Where-Object { $_.IP -eq '2001:db8::5' }
Test-Wert '198.51.100.10: EHLO (MAIL FROM zaehlt nicht)' 2 $ip10.EHLO
Test-Wert '198.51.100.20: HELO'                      1 $ip20.HELO
Test-Wert '198.51.100.20: kein EHLO'                 0 $ip20.EHLO
Test-Wert 'IPv6 in Klammern richtig ausgelesen'     1 $ipv6.EHLO
Test-Wert 'drei Gegenstellen'                       3 $smtp.Count

$smtpF = @(Invoke-SmtpReceiveHits -Files @($smtpFile) -UI $ui -ConnectorFilter 'Relay' -DateFrom '' -DateTo '' -ResolveDns $false)
Test-Wert 'Connector-Filter "Relay": zwei Gegenstellen' 2 $smtpF.Count

# ══ IIS ══════════════════════════════════════════════════════════════════════════
''
'--- IIS ---'
$iisIp = @(Invoke-IisLogQuery -Files @($iisFile) -UI $ui -Mode 'IisIp' -DateFrom '' -DateTo '')
$i10 = $iisIp | Where-Object { $_.IP -eq '198.51.100.10' }
Test-Wert 'IIS 198.51.100.10: Treffer'               2 $i10.Hits
Test-Wert 'IIS 198.51.100.10: 2xx / 3xx'             '1/1' "$($i10.S2xx)/$($i10.S3xx)"

$iisUser = @(Invoke-IisLogQuery -Files @($iisFile) -UI $ui -Mode 'IisUser' -DateFrom '' -DateTo '')
Test-Wert 'IIS nach Konto: aus dem Zwischenspeicher' $true $script:LastFromCache
$bernd = $iisUser | Where-Object { $_.User -eq 'EXAMPLE\bernd' }
Test-Wert 'IIS EXAMPLE\bernd: Treffer'                 2 $bernd.Hits
Test-Wert 'IIS EXAMPLE\bernd: 4xx / 5xx'               '1/1' "$($bernd.S4xx)/$($bernd.S5xx)"

# ── Aufraeumen ──────────────────────────────────────────────────────────────────
Remove-Item $dir -Recurse -Force
Remove-Item $tmp -Force
''
if ($script:Fehler -eq 0) { 'ERGEBNIS: alle Pruefungen bestanden'; exit 0 }
else { "ERGEBNIS: $script:Fehler Pruefung(en) fehlgeschlagen"; exit 1 }
