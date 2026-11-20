# Prüft DmarcReportAnalyzer ohne Postfach und ohne Netz: künstliche Reports (Beispieldomains,
# Dokumentations-IP-Adressen nach RFC 5737) als .xml, .gz und .zip. Keine echten Report-Daten.
# Die Funktionen werden per AST herausgelöst; am Ende läuft das Skript einmal ganz gegen einen Ordner.
$ErrorActionPreference = 'Stop'
$hier = if ($PSScriptRoot) { $PSScriptRoot }
        elseif ($MyInvocation.MyCommand.Path) { Split-Path $MyInvocation.MyCommand.Path -Parent }
        else { (Get-Location).Path }
$skript = Join-Path (Split-Path $hier -Parent) 'Invoke-DmarcReportAnalyzer.ps1'
Add-Type -AssemblyName System.IO.Compression

$ast = [System.Management.Automation.Language.Parser]::ParseFile($skript, [ref]$null, [ref]$null)
foreach ($fn in 'ConvertFrom-DmarcAttachment', 'ConvertFrom-DmarcXml', 'Get-DmarcRows', 'Get-SenderService',
                'Test-DomainRelated', 'Get-DmarcRowFinding', 'Get-FindingLabel', 'Format-DayList',
                'Get-DmarcSources', 'Get-DmarcDomainSummary', 'Get-DmarcRecommendation', 'New-DmarcHtmlReport',
                'Get-GraphPages') {
    $def = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $args[0].Name -eq $fn }, $false)
    if (-not $def) { throw "Funktion nicht gefunden: $fn" }
    Invoke-Expression $def[0].Extent.Text
}
$script:Version = 'test'

$pass = 0; $fail = 0
function T($name, $ist, $soll) {
    if ("$ist" -eq "$soll") { $script:pass++; Write-Output "  OK   $name" }
    else { $script:fail++; Write-Output "  FEHL $name : ist '$ist', soll '$soll'" }
}
function Werfen([scriptblock]$sb) { try { & $sb | Out-Null; $false } catch { $true } }

# --- Bausteine ---------------------------------------------------------------------------
# $auth: $null = eine Signatur der Absenderdomain mit Ergebnis $dkim; sonst Liste 'domain=ergebnis' (leer = keine Signatur).
function Rec($ip, $count, $dkim, $spf, $disp = 'none', $from = 'contoso.com', $auth = $null) {
    $dk = if ($null -eq $auth) { "<dkim><domain>$from</domain><result>$dkim</result></dkim>" }
          else { (@($auth) | ForEach-Object { $d, $r = $_ -split '='; "<dkim><domain>$d</domain><result>$r</result></dkim>" }) -join '' }
@"
  <record><row><source_ip>$ip</source_ip><count>$count</count>
    <policy_evaluated><disposition>$disp</disposition><dkim>$dkim</dkim><spf>$spf</spf></policy_evaluated></row>
    <identifiers><header_from>$from</header_from><envelope_from>$from</envelope_from></identifiers>
    <auth_results>$dk<spf><domain>$from</domain><result>$spf</result></spf></auth_results>
  </record>
"@
}
function Rep($org, $id, $begin, $end, $records, $p = 'none', $pct = '100', $ns = '') {
    $b = [DateTimeOffset]::new([datetime]$begin, [TimeSpan]::Zero).ToUnixTimeSeconds()
    $e = [DateTimeOffset]::new([datetime]$end, [TimeSpan]::Zero).ToUnixTimeSeconds()
@"
<?xml version="1.0" encoding="UTF-8"?>
<feedback$ns>
  <report_metadata><org_name>$org</org_name><email>noreply@example.net</email><report_id>$id</report_id>
    <date_range><begin>$b</begin><end>$e</end></date_range></report_metadata>
  <policy_published><domain>contoso.com</domain><adkim>r</adkim><aspf>r</aspf><p>$p</p><sp>$p</sp><pct>$pct</pct></policy_published>
$($records -join "`n")
</feedback>
"@
}
function GzBytes([string]$s) {
    $ms = New-Object IO.MemoryStream
    $gz = New-Object IO.Compression.GZipStream($ms, [IO.Compression.CompressionMode]::Compress)
    $b = [Text.Encoding]::UTF8.GetBytes($s); $gz.Write($b, 0, $b.Length); $gz.Dispose(); $ms.ToArray()
}
function ZipBytes([hashtable]$files) {
    $ms = New-Object IO.MemoryStream
    $z = New-Object IO.Compression.ZipArchive($ms, [IO.Compression.ZipArchiveMode]::Create, $true)
    foreach ($k in $files.Keys) { $e = $z.CreateEntry($k); $s = $e.Open(); $b = [Text.Encoding]::UTF8.GetBytes($files[$k]); $s.Write($b, 0, $b.Length); $s.Dispose() }
    $z.Dispose(); $ms.ToArray()
}

$r1 = Rep 'Example Mail' 'r-1' '2026-11-01' '2026-11-02' @(
    (Rec '192.0.2.10' 100 'pass' 'pass'),
    (Rec '198.51.100.7' 5 'fail' 'fail' 'quarantine'),
    (Rec '203.0.113.5' 20 'fail' 'pass'))

Write-Output "`n######## Einlesen ########"
$x = ConvertFrom-DmarcXml -Xml $r1
T 'Reporter'                    $x.Reporter 'Example Mail'
T 'Beginn aus Unixzeit'         $x.Begin.ToString('yyyy-MM-dd HH:mm') '2026-11-01 00:00'
T 'Policy'                      $x.Policy 'none'
T 'drei Datensätze'             $x.Records.Count 3
T 'DMARC-Pass nur mit Alignment' (($x.Records | Where-Object DmarcPass).Count) 2
T 'SPF-only gilt als Pass'      ($x.Records | Where-Object SourceIp -eq '203.0.113.5').DmarcPass $true
T 'Auth-Ergebnis DKIM'          ($x.Records[0].DkimAuth) 'contoso.com=pass'

$ns = Rep 'NS Org' 'r-ns' '2026-11-01' '2026-11-02' @((Rec '192.0.2.10' 3 'pass' 'pass')) 'none' '100' ' xmlns="urn:ietf:params:xml:ns:dmarc-2.0"'
T 'mit Namespace lesbar'        (ConvertFrom-DmarcXml -Xml $ns).Records[0].Count 3
T 'kein feedback -> wirft'      (Werfen { ConvertFrom-DmarcXml -Xml '<?xml version="1.0"?><other/>' }) $true
$xxe = '<?xml version="1.0"?><!DOCTYPE feedback [<!ENTITY x SYSTEM "file:///c:/windows/win.ini">]><feedback><report_metadata><org_name>&x;</org_name></report_metadata></feedback>'
T 'DTD/externe Entität -> wirft' (Werfen { ConvertFrom-DmarcXml -Xml $xxe }) $true

Write-Output "`n######## Anhänge erkennen ########"
T 'xml roh'                     @(ConvertFrom-DmarcAttachment -Bytes ([Text.Encoding]::UTF8.GetBytes($r1))).Count 1
T 'gzip'                        @(ConvertFrom-DmarcAttachment -Bytes (GzBytes $r1)).Count 1
T 'zip mit zwei Reports'        @(ConvertFrom-DmarcAttachment -Bytes (ZipBytes @{ 'a.xml' = $r1; 'b.xml' = $ns })).Count 2
T 'xml mit BOM'                 @(ConvertFrom-DmarcAttachment -Bytes ([byte[]](0xEF,0xBB,0xBF) + [Text.Encoding]::UTF8.GetBytes($r1))).Count 1
T 'PNG ist kein Report'         @(ConvertFrom-DmarcAttachment -Bytes ([byte[]](0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A))).Count 0
T 'fremdes XML ist kein Report' @(ConvertFrom-DmarcAttachment -Bytes ([Text.Encoding]::UTF8.GetBytes('<?xml version="1.0"?><note>x</note>'))).Count 0

Write-Output "`n######## Zusammenfassen ########"
$rows = @(Get-DmarcRows -Reports @($x, $x))
T 'doppelter Report zählt einmal' $rows.Count 3
$src = @(Get-DmarcSources -Rows $rows)
T 'drei Quellen'                $src.Count 3
T 'größte Quelle zuerst'        $src[0].SourceIp '192.0.2.10'
T 'Quarantäne gezählt'          ($src | Where-Object SourceIp -eq '198.51.100.7').Quarantined 5
$dom = @(Get-DmarcDomainSummary -Rows $rows)
T 'Pass-Anteil'                 $dom[0].PassPct 96
T 'Mails gesamt'                $dom[0].Count 125

Write-Output "`n######## Dienste ########"
T 'Microsoft 365'               (Get-SenderService 'mail-am6eur05on2070.outbound.protection.outlook.com') 'Microsoft 365'
T 'Google'                      (Get-SenderService 'mail-wr1-f41.google.com') 'Google'
T 'SendGrid'                    (Get-SenderService 'o1.ptr1234.sendgrid.net') 'SendGrid'
T 'Google-Weiterleitung'        (Get-SenderService 'x08-01.v6.unverified-forwarding.1e100.net') 'Weiterleitung (Google)'
T 'GMX-Weiterleitung'           (Get-SenderService 'mout-bounce.gmx.net') 'Weiterleitung (GMX/WEB.DE)'
T 'GMX-Versand ist keine Weiterleitung' (Get-SenderService 'mout.gmx.net') ''
T 'unbekannt bleibt leer'      (Get-SenderService 'mx.example.org') ''
T 'Endung muss passen'          (Get-SenderService 'google.com.attacker.example') ''

Write-Output "`n######## Befunde je Zeile und Quelle ########"
$bf = Rep 'Example Mail' 'r-bf' '2026-11-16' '2026-11-17' @(
    (Rec '192.0.2.20' 4 'fail' 'pass' 'none' 'contoso.com' @()),
    (Rec '192.0.2.21' 2 'fail' 'pass' 'none' 'contoso.com' @('contoso.com=fail')),
    (Rec '192.0.2.22' 3 'pass' 'pass' 'none' 'contoso.com' @('contoso.com=pass', 'contoso.com=fail')),
    (Rec '192.0.2.23' 5 'pass' 'fail' 'none' 'contoso.com' @('sub.contoso.com=fail', 'contoso.com=pass')),
    (Rec '192.0.2.24' 7 'pass' 'pass'),
    (Rec '198.51.100.9' 1 'fail' 'fail'))
$bfr = @(Get-DmarcRows -Reports @(ConvertFrom-DmarcXml -Xml $bf))
$fd = @{}; foreach ($r in $bfr) { $fd[$r.SourceIp] = Get-DmarcRowFinding $r }
T 'keine Signatur -> NurSpf'     $fd['192.0.2.20'] 'NurSpf'
T 'gebrochen -> DkimGebrochen'   $fd['192.0.2.21'] 'DkimGebrochen'
T 'pass+fail -> ZweitSignatur'   $fd['192.0.2.22'] 'ZweitSignatur'
T 'Subdomain-Signatur zählt mit' $fd['192.0.2.23'] 'ZweitSignatur'
T 'sauber -> leer'               $fd['192.0.2.24'] ''
T 'nicht bestanden -> Fail'      $fd['198.51.100.9'] 'Fail'
T 'fremde Signatur bricht nichts' (Get-DmarcRowFinding ([pscustomobject]@{ DmarcPass = $true; DkimAligned = $true; HeaderFrom = 'contoso.com'; DkimAuth = 'contoso.com=pass; mailer.example=fail' })) ''
T 'Weiterleiter vor Fail'        (Get-DmarcRowFinding ([pscustomobject]@{ Service = 'Weiterleitung (GMX/WEB.DE)'; DmarcPass = $false })) 'Weiterleitung'
T 'contoso.com.evil ist nicht verwandt' (Test-DomainRelated 'contoso.com.evil.example' 'contoso.com') $false
$bfs = @(Get-DmarcSources -Rows $bfr)
$q = $bfs | Where-Object SourceIp -eq '192.0.2.20'
T 'Quelle zählt Befund'          $q.Findings['NurSpf'] 4
T 'Quelle kennt Befund-Tag'      (@($q.FindingDays['NurSpf']) -join ',') '2026-11-16'
T 'Quelle hat Zeitraum'          $q.FirstDay.ToString('yyyy-MM-dd') '2026-11-16'
T 'rohes DKIM-Ergebnis'          ($bfs | Where-Object SourceIp -eq '192.0.2.22').DkimResults 'contoso.com=pass; contoso.com=fail'
T 'keine Signatur sichtbar'      $q.DkimResults '(keine Signatur)'
T 'jede Quelle hat eine Id'      (@($bfs | Where-Object { $_.Id -match '^src-\d+$' }).Count) $bfs.Count
T 'Tage-Liste kurz'              (Format-DayList @('2026-11-17', '2026-11-16', '2026-11-16')) '16.11., 17.11.'

Write-Output "`n######## Empfehlungen ########"
function D($pctPass, $policy, $tage, $pct = '100') { [pscustomobject]@{ Domain = 'contoso.com'; PassPct = $pctPass; Policy = $policy; Pct = $pct; From = [datetime]'2026-11-01'; To = ([datetime]'2026-11-01').AddDays($tage) } }
function S($count, $pass, $dkim, $svc = '', $ptr = '', $f = @{}, $days = '2026-11-16') {
    $fdays = @{}; foreach ($k in @($f.Keys)) { $fdays[$k] = @($days) }
    [pscustomobject]@{ Id = 'src-7'; Domain = 'contoso.com'; SourceIp = '198.51.100.7'; Ptr = $ptr; Service = $svc; Count = $count; DmarcPass = $pass; DkimAligned = $dkim; Findings = $f; FindingDays = $fdays } }
function Empf($d, $s) { @(Get-DmarcRecommendation -Domains @($d) -Sources @($s)) }
$r = Empf (D 99 'reject' 30) (S 50 0 0)
T 'unbekannte Fehlquelle: Hoch'  ((($r | Where-Object Stufe -eq 'Hoch').Text) -match 'keinem bekannten Dienst') $true
$r = Empf (D 99 'reject' 30) (S 50 0 0 'SendGrid')
T 'bekannter Dienst: DKIM-Hinweis' ((($r | Where-Object Stufe -eq 'Hoch').Text) -match 'DKIM mit der eigenen Domain') $true
$r = Empf (D 99 'reject' 30) (S 50 50 10 '' '' @{ NurSpf = 40 })
T 'nur über SPF: Mittel'         @($r | Where-Object Stufe -eq 'Mittel').Count 1
T '... mit Tag'                  ($r[0].Text -match 'am 16\.11\.') $true
T '... mit Sprungziel'           $r[0].Anchor 'src-7'
$r = Empf (D 99 'reject' 30) (S 50 50 10 'Microsoft 365' '' @{ DkimGebrochen = 40 })
T 'gebrochen: Veränderung statt einrichten' ((($r | Where-Object Stufe -eq 'Mittel').Text) -match 'nach dem Signieren verändert') $true
$r = Empf (D 99 'reject' 30) (S 6 6 6 'Weiterleitung (Google)' '' @{ Weiterleitung = 6 })
T 'Weiterleitung mit Pass: still' $r.Count 0
$r = Empf (D 99 'reject' 30) (S 6 0 0 'Weiterleitung (Google)' '' @{ Weiterleitung = 6 })
T 'Weiterleitung mit Fail: nur Info' (($r | ForEach-Object Stufe) -join ',') 'Info'
$r = Empf (D 99 'reject' 30) (S 50 50 50)
T 'alles sauber: keine Empfehlung' $r.Count 0
T 'none + 99 % + 20 Tage -> quarantine' ((Empf (D 99 'none' 20) $null).Text -match 'p=quarantine') $true
T 'none + 80 % -> noch nicht'    ((Empf (D 80 'none' 20) $null).Text -match 'mehr als 98') $true
T 'none + 99 % + 5 Tage -> noch nicht' ((Empf (D 99 'none' 5) $null).Text -match 'mehr als 98') $true
T 'quarantine + 99,5 % -> reject' ((Empf (D 99.5 'quarantine' 20) $null).Text -match 'p=reject') $true
$r = Empf (D 99.5 'quarantine' 20) (S 50 50 10 '' '' @{ NurSpf = 40 } '2026-11-17')
$p = ($r | Where-Object { $_.Text -match 'Policy ist' }).Text
T 'SPF-only hält reject zurück'  ($p -match 'Nächster Schritt') $false
T '... und sagt warum, mit Datum' ($p -match 'Vor dem nächsten Schritt.*17\.11\.2026') $true
$r = Empf (D 99.5 'none' 20) (S 50 50 10 '' '' @{ DkimGebrochen = 40 })
T 'gilt auch vor quarantine'     ((($r | Where-Object { $_.Text -match 'Policy ist' }).Text) -match 'Nächster Schritt') $false
T 'pct=50 -> Hinweis'           (@((Empf (D 99.5 'quarantine' 20 '50') $null) | Where-Object { $_.Text -match 'pct=50' }).Count) 1

Write-Output "`n######## HTML ########"
# Im XML maskiert, damit beim Leser wirklich der TEXT "<script>..." ankommt (roh wäre es ein Element).
$evil = Rep '&lt;script&gt;alert(1)&lt;/script&gt;' 'r-x' '2026-11-01' '2026-11-02' @((Rec '192.0.2.99' 1 'fail' 'fail'))
$er = @(Get-DmarcRows -Reports @(ConvertFrom-DmarcXml -Xml $evil))
T 'Testeingabe: Script-Text kam an' $er[0].Reporter '<script>alert(1)</script>'
$html = New-DmarcHtmlReport -Domains @(Get-DmarcDomainSummary -Rows $er) -Sources @(Get-DmarcSources -Rows $er) -Recommendations @() -Rows $er -Title 'Test' -Origin 'Test'
T 'Fremddaten werden kodiert'    ($html -match '<script>alert') $false
T '... und stehen kodiert drin'  ($html -match '&lt;script&gt;alert') $true
T 'Domain-Karte vorhanden'       ($html -match 'contoso\.com') $true
$bd = @(Get-DmarcDomainSummary -Rows $bfr)
$bh = New-DmarcHtmlReport -Domains $bd -Sources $bfs -Recommendations @(Get-DmarcRecommendation -Domains $bd -Sources $bfs) -Rows $bfr -Title 'T' -Origin 'T'
$zid = ($bfs | Where-Object SourceIp -eq '192.0.2.20').Id
T 'auffällige Zeile hervorgehoben' ($bh -match "<tr id='$zid' class='hl'>") $true
T 'saubere Zeile nicht'          ($bh -match "<tr id='$(($bfs | Where-Object SourceIp -eq '192.0.2.24').Id)' class='hl'>") $false
T 'Empfehlung verlinkt Zeile'    ($bh -match "href='#$zid'") $true
T 'pass; fail im HTML sichtbar'  ($bh -match 'contoso\.com=pass; contoso\.com=fail') $true
T 'Befund mit Anzahl im HTML'    ($bh -match 'kein DKIM, nur SPF: 4') $true
T 'Lesehilfe vorhanden'          ($bh -match 'So liest du den Bericht') $true
T 'keine externen Quellen'      ($html -match '(src|href)\s*=\s*["'']https?://') $false

Write-Output "`n######## Graph-Seiten: unsortiert, ältere nur überspringen ########"
# Attrappe für Invoke-Graph: zwei Seiten, die Daten NICHT absteigend sortiert (wie bei
# Gruppen-Unterhaltungen am 20.11.2026 beobachtet).
$script:Seiten = @{
    'p1' = [pscustomobject]@{ value = @([pscustomobject]@{ id = 'a'; d = '2026-11-19T10:00:00Z' }, [pscustomobject]@{ id = 'alt'; d = '2026-09-01T10:00:00Z' }); '@odata.nextLink' = 'p2' }
    'p2' = [pscustomobject]@{ value = @([pscustomobject]@{ id = 'b'; d = '2026-11-18T10:00:00Z' }) }
}
function Invoke-Graph { param($Uri, $Token) $script:Seiten[$Uri] }
$ids = @(Get-GraphPages -Uri 'p1' -Token 'x' -SkipBefore ([datetime]'2026-10-20') -DateField 'd' | ForEach-Object id)
T 'nach älterem Eintrag weiterlesen' ($ids -join ',') 'a,b'

Write-Output "`n######## Ganzer Lauf gegen einen Ordner ########"
$tmp = Join-Path $env:TEMP ('dmarc-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory $tmp | Out-Null
try {
    $heute = (Get-Date).ToUniversalTime().Date
    $neu = Rep 'Example Mail' 'r-neu' $heute.AddDays(-2).ToString('s') $heute.AddDays(-1).ToString('s') @((Rec '192.0.2.10' 10 'pass' 'pass'), (Rec '198.51.100.7' 2 'fail' 'fail'))
    $alt = Rep 'Example Mail' 'r-alt' $heute.AddDays(-90).ToString('s') $heute.AddDays(-89).ToString('s') @((Rec '192.0.2.10' 1000 'pass' 'pass'))
    [IO.File]::WriteAllBytes((Join-Path $tmp 'neu.xml.gz'), (GzBytes $neu))
    [IO.File]::WriteAllText((Join-Path $tmp 'alt.xml'), $alt)
    [IO.File]::WriteAllBytes((Join-Path $tmp 'kaputt.zip'), [byte[]](0x50, 0x4B, 0x03, 0x04, 0x00))
    $out = Join-Path $tmp 'neu\unter\bericht.html'   # Zielordner gibt es noch nicht
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $skript -Path $tmp -Days 30 -OutputPath $out -NoReverseDns *> $null
    T 'Exitcode 0'                 $LASTEXITCODE 0
    T 'HTML geschrieben'           (Test-Path $out) $true
    $csv = @(Import-Csv ([IO.Path]::ChangeExtension($out, '.csv')) -Delimiter ';')
    T 'nur Reports im Zeitraum (2 Zeilen, nicht 3)' $csv.Count 2
    T 'CSV hat Befund je Zeile'    (($csv | Where-Object SourceIp -eq '198.51.100.7').Auffaellig) 'DMARC fail'
    T 'kaputte Datei bricht nichts ab' ((Get-Content $out -Raw) -match 'Pass') $true

    $leer = Join-Path $tmp 'leer'; New-Item -ItemType Directory $leer | Out-Null
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $skript -Path $leer -OutputPath (Join-Path $tmp 'x.html') -NoReverseDns *> $null
    T 'keine Daten -> Exitcode 2'  $LASTEXITCODE 2
}
finally { Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue }

Write-Output "`n================ Bestanden: $pass   Fehlgeschlagen: $fail ================"
if ($fail -gt 0) { exit 1 }
