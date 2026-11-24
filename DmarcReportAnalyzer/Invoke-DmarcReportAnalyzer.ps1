<#
.SYNOPSIS
    Wertet DMARC-Aggregate-Reports (RUA) aus und erzeugt einen HTML-Bericht samt CSV.

.DESCRIPTION
    Liest Reports als .xml, .xml.gz oder .zip - aus einem Ordner, aus dem Gruppenpostfach einer
    Microsoft-365-Gruppe oder aus einem persönlichen bzw. freigegebenen Postfach (Microsoft Graph,
    delegiert). Welche Datei ein Report ist, entscheidet der Inhalt, nicht Betreff oder Dateiname.

    Der Bericht zeigt je Domain Zeitraum, Menge und DMARC-Pass-Anteil, je sendender Quelle die
    Ergebnisse von SPF, DKIM und Alignment, ordnet Quellen bekannten Diensten zu und gibt
    Empfehlungen - auch zur nächsten Policy-Stufe. Er ist eine einzelne HTML-Datei ohne externe
    Skripte oder Schriften und lässt sich offline öffnen.

    Report-Inhalte sind Fremddaten: Alles, was in den Bericht geht, wird HTML-kodiert.

.PARAMETER Path
    Ordner mit Reports (.xml, .gz, .zip), Unterordner eingeschlossen.

.PARAMETER GroupMailbox
    Adresse einer Microsoft-365-Gruppe, in deren Postfach die Reports eingehen.
    Benötigt delegiert Group.Read.All. Dafür muss im Tenant einmal ein Administrator der App
    "Microsoft Graph Command Line Tools" zugestimmt haben; danach genügt die Anmeldung des
    Benutzers, der Zugriff auf die Gruppe hat.

.PARAMETER Mailbox
    Adresse eines persönlichen oder freigegebenen Postfachs (Vollzugriff nötig).
    Benötigt delegiert Mail.Read.Shared (keine Administratorzustimmung nötig).

.PARAMETER Days
    Nur Reports, deren Zeitraum in den letzten n Tagen endet. Standard 30.

.PARAMETER OutputPath
    Pfad der HTML-Datei. Die CSV entsteht daneben mit gleichem Namen. Standard:
    .\DMARC-Bericht-<Datum>.html im aktuellen Ordner.

.PARAMETER AccessToken
    Vorhandenes Graph-Token statt einer Anmeldung.

.PARAMETER Tenant
    Tenant für die Anmeldung (Domain oder ID). Standard 'organizations'.

.PARAMETER UseDeviceCode
    Anmeldung per Gerätecode statt im Browser - für Rechner ohne Browser oder Sitzungen, in denen
    sich keiner öffnen lässt. Standard ist die Anmeldung im Browser: Das Skript öffnet ihn selbst
    und nimmt die Antwort auf http://localhost an einem freien Port entgegen (PKCE).

.PARAMETER NoReverseDns
    Keine PTR-Abfragen - schneller, aber ohne Zuordnung zu Diensten über den Hostnamen.

.PARAMETER NoMicrosoftRanges
    Die Adressbereiche von Exchange Online nicht laden. Standard ist, die öffentliche Liste von
    Microsoft (endpoints.office.com) einmal abzurufen und Quellen darin als "Microsoft 365" zu
    erkennen, auch ohne passenden Rückwärtsnamen. Gesendet wird dabei nur eine Zufallskennung.

.PARAMETER MessageTrace
    Zu jeder auffälligen Quelle die Mails aus dem Message Trace von Exchange Online suchen:
    Absender @Domain an den Tagen des Befunds, passende Empfänger (Empfängerdomain aus dem
    Report, sonst die Domains des meldenden Anbieters) zuerst. Braucht das Modul
    ExchangeOnlineManagement und eine Anmeldung mit Leserecht auf den Message Trace; eine
    bestehende Verbindung wird genutzt. Findet nur Mails, die durch den eigenen Tenant liefen.

.EXAMPLE
    .\Invoke-DmarcReportAnalyzer.ps1 -Path C:\Reports
.EXAMPLE
    .\Invoke-DmarcReportAnalyzer.ps1 -GroupMailbox dmarc@contoso.onmicrosoft.com -Days 14
.EXAMPLE
    .\Invoke-DmarcReportAnalyzer.ps1 -Mailbox dmarc-reports@contoso.com -OutputPath C:\Temp\dmarc.html
.EXAMPLE
    .\Invoke-DmarcReportAnalyzer.ps1 -GroupMailbox dmarc@contoso.onmicrosoft.com -UseDeviceCode
.EXAMPLE
    .\Invoke-DmarcReportAnalyzer.ps1 -Mailbox dmarc-reports@contoso.com -MessageTrace
#>
[CmdletBinding(DefaultParameterSetName = 'Folder')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Folder')][string]$Path,
    [Parameter(Mandatory, ParameterSetName = 'Group')][string]$GroupMailbox,
    [Parameter(Mandatory, ParameterSetName = 'Mailbox')][string]$Mailbox,
    [ValidateRange(1, 400)][int]$Days = 30,
    [string]$OutputPath,
    [Parameter(ParameterSetName = 'Group')][Parameter(ParameterSetName = 'Mailbox')][string]$AccessToken,
    [Parameter(ParameterSetName = 'Group')][Parameter(ParameterSetName = 'Mailbox')][string]$Tenant = 'organizations',
    [Parameter(ParameterSetName = 'Group')][Parameter(ParameterSetName = 'Mailbox')][switch]$UseDeviceCode,
    [switch]$NoReverseDns,
    [switch]$NoMicrosoftRanges,
    [switch]$MessageTrace
)

# Einzige Stelle fuer die Versionsnummer; Build und Release-Tag lesen sie hier aus.
$script:Version = '1.2.4'

#region ------------------------------------------------------------ Fortschritt

$script:ProgressTimer = [Diagnostics.Stopwatch]::StartNew()
function Show-Progress {
    <#
        Fortschritt für lange Phasen, höchstens alle 300 ms (Write-Progress ist unter 5.1 teuer).
        -Force zeigt sofort (Phasenbeginn). Ohne Konsole oder mit $ProgressPreference =
        'SilentlyContinue' passiert nichts; ein Host, der Write-Progress ablehnt, bricht den Lauf nicht ab.
    #>
    param([string]$Activity, [string]$Status, [int]$Current = 0, [int]$Total = 0, [switch]$Completed, [switch]$Force)
    try {
        if ($Completed) { Write-Progress -Activity $Activity -Completed; return }
        if (-not $Force -and $script:ProgressTimer.ElapsedMilliseconds -lt 300) { return }
        $script:ProgressTimer.Restart()
        $p = @{ Activity = $Activity; Status = $(if ($Status) { $Status } else { '...' }) }
        if ($Total -gt 0) { $p.PercentComplete = [math]::Min(100, [int](100 * $Current / $Total)) }
        Write-Progress @p
    } catch { }
}

#endregion

#region ------------------------------------------------------------ Einlesen

function ConvertFrom-DmarcAttachment {
    <#
        Ein Anhang oder eine Datei -> XML-Texte. Erkennung über die ersten Bytes:
        1F 8B = gzip, 50 4B = zip (jede .xml darin), sonst Text mit '<'.
        Liefert nichts, wenn es kein Report ist (z. B. ein Logo in der Mail).
    #>
    param([Parameter(Mandatory)][byte[]]$Bytes, [string]$Name = '')
    if ($Bytes.Length -lt 4) { return }
    if ($Bytes[0] -eq 0x1F -and $Bytes[1] -eq 0x8B) {
        $in = New-Object System.IO.MemoryStream(, $Bytes)
        $gz = New-Object System.IO.Compression.GZipStream($in, [System.IO.Compression.CompressionMode]::Decompress)
        $out = New-Object System.IO.MemoryStream
        try { $gz.CopyTo($out) } finally { $gz.Dispose() }
        return ConvertFrom-DmarcAttachment -Bytes $out.ToArray() -Name ($Name -replace '\.gz$', '')
    }
    if ($Bytes[0] -eq 0x50 -and $Bytes[1] -eq 0x4B) {
        Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
        $zip = New-Object System.IO.Compression.ZipArchive((New-Object System.IO.MemoryStream(, $Bytes)), [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            foreach ($e in $zip.Entries) {
                if ($e.Length -eq 0) { continue }
                $ms = New-Object System.IO.MemoryStream
                $s = $e.Open(); try { $s.CopyTo($ms) } finally { $s.Dispose() }
                ConvertFrom-DmarcAttachment -Bytes $ms.ToArray() -Name $e.FullName
            }
        } finally { $zip.Dispose() }
        return
    }
    # Text: BOM überspringen, Kodierung aus dem XML-Kopf übernimmt der XmlReader.
    $text = [System.Text.Encoding]::UTF8.GetString($Bytes).TrimStart([char]0xFEFF, ' ', "`r", "`n", "`t")
    if ($text.StartsWith('<') -and $text -match '<feedback[\s>]') { return $text }
}

function ConvertFrom-DmarcXml {
    <#
        Ein Report (XML-Text) -> ein Objekt mit Kopf und Zeilen. Wirft bei XML, das kein
        DMARC-Report ist. DTD-Verarbeitung ist aus (Fremddaten).
    #>
    param([Parameter(Mandatory)][string]$Xml)
    $settings = New-Object System.Xml.XmlReaderSettings
    $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
    $settings.XmlResolver = $null
    $reader = [System.Xml.XmlReader]::Create((New-Object System.IO.StringReader($Xml)), $settings)
    $doc = New-Object System.Xml.XmlDocument
    try { $doc.Load($reader) } finally { $reader.Dispose() }

    # Namespace (DMARC 2.0 / RFC 9990-Entwurf) ignorieren: über local-name() lesen.
    $fb = $doc.SelectSingleNode("/*[local-name()='feedback']")
    if (-not $fb) { throw 'Kein DMARC-Report (Wurzelelement <feedback> fehlt).' }
    function V($node, [string]$path) {
        $parts = $path -split '/'
        $x = '.'
        foreach ($p in $parts) { $x += "/*[local-name()='$p']" }
        $n = $node.SelectSingleNode($x)
        if ($n) { return $n.InnerText.Trim() } else { return '' }
    }
    function T([string]$unix) {
        $u = 0L
        if ([long]::TryParse($unix, [ref]$u)) { return [DateTimeOffset]::FromUnixTimeSeconds($u).UtcDateTime }
        return $null
    }

    $report = [pscustomobject]@{
        Reporter  = V $fb 'report_metadata/org_name'
        Email     = V $fb 'report_metadata/email'
        ReportId  = V $fb 'report_metadata/report_id'
        Begin     = T (V $fb 'report_metadata/date_range/begin')
        End       = T (V $fb 'report_metadata/date_range/end')
        Domain    = (V $fb 'policy_published/domain').ToLower()
        Adkim     = V $fb 'policy_published/adkim'
        Aspf      = V $fb 'policy_published/aspf'
        Policy    = V $fb 'policy_published/p'
        SubPolicy = V $fb 'policy_published/sp'
        Pct       = V $fb 'policy_published/pct'
        Records   = @()
    }

    $rows = @()
    foreach ($rec in $fb.SelectNodes("./*[local-name()='record']")) {
        $count = 0; [void][int]::TryParse((V $rec 'row/count'), [ref]$count)
        $dkimAuth = @(foreach ($d in $rec.SelectNodes("./*[local-name()='auth_results']/*[local-name()='dkim']")) {
            '{0}={1}' -f (V $d 'domain'), (V $d 'result')
        })
        $dkimSel = @(foreach ($d in $rec.SelectNodes("./*[local-name()='auth_results']/*[local-name()='dkim']")) {
            $s = V $d 'selector'; if ($s) { '{0}:{1}' -f (V $d 'domain'), $s }
        })
        # Freiwillig (RFC 7489 Anhang C): Grund für eine abweichende Behandlung, z. B. forwarded.
        $reason = @(foreach ($g in $rec.SelectNodes("./*[local-name()='row']/*[local-name()='policy_evaluated']/*[local-name()='reason']")) {
            $ty = V $g 'type'; $co = V $g 'comment'
            if ($co) { '{0}: {1}' -f $ty, $co } elseif ($ty) { $ty }
        })
        $spfAuth = @(foreach ($s in $rec.SelectNodes("./*[local-name()='auth_results']/*[local-name()='spf']")) {
            '{0}={1}' -f (V $s 'domain'), (V $s 'result')
        })
        $dkimAligned = (V $rec 'row/policy_evaluated/dkim') -eq 'pass'
        $spfAligned  = (V $rec 'row/policy_evaluated/spf') -eq 'pass'
        $rows += [pscustomobject]@{
            SourceIp     = V $rec 'row/source_ip'
            Count        = $count
            Disposition  = V $rec 'row/policy_evaluated/disposition'
            DkimAligned  = $dkimAligned
            SpfAligned   = $spfAligned
            DmarcPass    = ($dkimAligned -or $spfAligned)
            HeaderFrom   = (V $rec 'identifiers/header_from').ToLower()
            EnvelopeFrom = (V $rec 'identifiers/envelope_from').ToLower()
            EnvelopeTo   = (V $rec 'identifiers/envelope_to').ToLower()
            Reason       = ($reason -join '; ')
            DkimSelector = ($dkimSel -join '; ')
            DkimAuth     = ($dkimAuth -join '; ')
            SpfAuth      = ($spfAuth -join '; ')
        }
    }
    $report.Records = $rows
    return $report
}

function Get-DmarcRows {
    <#
        Reports -> flache Zeilen (eine je Record) mit Kopfdaten. Doppelte Reports (gleicher
        Reporter + Report-ID, etwa zweimal zugestellt) zählen einmal.
    #>
    param([object[]]$Reports)
    $seen = @{}
    foreach ($r in $Reports) {
        $key = "$($r.Reporter)|$($r.ReportId)"
        if ($r.ReportId -and $seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        $begin = if ($r.Begin) { $r.Begin } else { $r.End }   # fehlender Beginn: Ende nehmen, nicht abstürzen
        foreach ($rec in $r.Records) {
            [pscustomobject]@{
                Reporter = $r.Reporter; ReportId = $r.ReportId; Begin = $begin; End = $r.End
                Domain = $r.Domain; Policy = $r.Policy; Pct = $r.Pct
                SourceIp = $rec.SourceIp; Count = $rec.Count; Disposition = $rec.Disposition
                DkimAligned = $rec.DkimAligned; SpfAligned = $rec.SpfAligned; DmarcPass = $rec.DmarcPass
                HeaderFrom = $rec.HeaderFrom; EnvelopeFrom = $rec.EnvelopeFrom
                EnvelopeTo = $rec.EnvelopeTo; Reason = $rec.Reason; DkimSelector = $rec.DkimSelector
                DkimAuth = $rec.DkimAuth; SpfAuth = $rec.SpfAuth
                Ptr = ''; Service = ''
            }
        }
    }
}

#endregion

#region ------------------------------------------------------------ Auswerten

function Get-SenderService {
    <# Bekannte Dienste nach Hostnamen-Endung (PTR). Bewusst kurz: lieber "unbekannt" als falsch. #>
    param([string]$Ptr)
    if (-not $Ptr) { return '' }
    # Weiterleitungen zuerst: Sie senden fremde Mails weiter, sind also keine Absender der Domain.
    # Name beginnt mit 'Weiterleitung' - daran erkennt Get-DmarcRowFinding sie.
    $known = @(
    @{ Pattern = '\.unverified-forwarding\.1e100\.net$';             Name = 'Weiterleitung (Google)' }
    @{ Pattern = '^mout-bounce\.(gmx|web)\.(net|de|com)$';           Name = 'Weiterleitung (GMX/WEB.DE)' }
    @{ Pattern = '\.(outlook|protection\.outlook)\.com$';            Name = 'Microsoft 365' }
    @{ Pattern = '\.1e100\.net$';                                     Name = 'Google' }
    @{ Pattern = '\.(google|googlemail)\.com$';                       Name = 'Google' }
    @{ Pattern = '\.sendgrid\.net$';                                  Name = 'SendGrid' }
    @{ Pattern = '\.(mcsv\.net|mcdlv\.net|rsgsv\.net|mandrillapp\.com)$'; Name = 'Mailchimp/Mandrill' }
    @{ Pattern = '\.amazonses\.com$';                                 Name = 'Amazon SES' }
    @{ Pattern = '\.mailgun\.(net|org)$';                             Name = 'Mailgun' }
    @{ Pattern = '\.sparkpostmail\.com$';                             Name = 'SparkPost' }
    @{ Pattern = '\.(sendinblue|brevo)\.com$';                        Name = 'Brevo' }
    @{ Pattern = '\.hubspotemail\.net$';                              Name = 'HubSpot' }
    @{ Pattern = '\.(exacttarget|salesforce)\.com$';                  Name = 'Salesforce' }
    @{ Pattern = '\.zendesk\.com$';                                   Name = 'Zendesk' }
    @{ Pattern = '\.(yahoo|yahoodns)\.(com|net)$';                    Name = 'Yahoo' }
    )
    foreach ($k in $known) { if ($Ptr -match $k.Pattern) { return $k.Name } }
    return ''
}

function Resolve-Ptr {
    param([string]$Ip)
    if (-not $script:PtrCache) { $script:PtrCache = @{} }
    if ($script:PtrCache.ContainsKey($Ip)) { return $script:PtrCache[$Ip] }
    $name = ''
    try {
        $r = Resolve-DnsName -Name $Ip -Type PTR -QuickTimeout -ErrorAction Stop | Where-Object { $_.Type -eq 'PTR' } | Select-Object -First 1
        if ($r) { $name = "$($r.NameHost)".TrimEnd('.').ToLower() }
    } catch { }
    $script:PtrCache[$Ip] = $name
    return $name
}

function Resolve-PtrBatch {
    <#
        Rückwärtsnamen parallel (Runspaces, auch unter 5.1). Nacheinander wartet jede Adresse ohne PTR
        ihren Timeout ab - gemessen 24 s für 19 Quellen. Ergebnis landet in $script:PtrCache wie bei
        Resolve-Ptr. -Resolver nur für Tests: Skriptblock mit param($ip), gibt den Namen oder '' zurück.
    #>
    param([string[]]$Ips, [int]$Threads = 16, [scriptblock]$Resolver)
    if (-not $script:PtrCache) { $script:PtrCache = @{} }
    $todo = @($Ips | Where-Object { $_ -and -not $script:PtrCache.ContainsKey($_) } | Sort-Object -Unique)
    if (-not $todo) { return }
    if (-not $Resolver) {
        $Resolver = { param($ip)
            try {
                $r = Resolve-DnsName -Name $ip -Type PTR -QuickTimeout -ErrorAction Stop | Where-Object { $_.Type -eq 'PTR' } | Select-Object -First 1
                if ($r) { "$($r.NameHost)".TrimEnd('.').ToLower() } else { '' }
            } catch { '' } }
    }
    $pool = [runspacefactory]::CreateRunspacePool(1, [math]::Max(1, [math]::Min($Threads, $todo.Count)))
    $pool.Open()
    $jobs = @(); $i = 0
    try {
        foreach ($ip in $todo) {
            $ps = [powershell]::Create().AddScript($Resolver.ToString()).AddArgument($ip)
            $ps.RunspacePool = $pool
            $jobs += [pscustomobject]@{ Ip = $ip; Ps = $ps; H = $ps.BeginInvoke() }
        }
        foreach ($j in $jobs) {
            $out = @(try { $j.Ps.EndInvoke($j.H) } catch { })
            $script:PtrCache[$j.Ip] = if ($out.Count) { "$($out[0])" } else { '' }
            $i++
            Show-Progress -Activity 'Rückwärtsnamen (PTR)' -Status ("IP {0} von {1}" -f $i, $todo.Count) -Current $i -Total $todo.Count
        }
    } finally {
        foreach ($j in $jobs) { $j.Ps.Dispose() }
        $pool.Close(); $pool.Dispose()
        Show-Progress -Activity 'Rückwärtsnamen (PTR)' -Completed
    }
}

function Test-DomainRelated {
    <# Entspannte Zuordnung: gleiche Domain oder eine ist Subdomain der anderen. #>
    param([string]$A, [string]$B)
    $a = $A.Trim().TrimEnd('.').ToLower(); $b = $B.Trim().TrimEnd('.').ToLower()
    if (-not $a -or -not $b) { return $false }
    return ($a -eq $b -or $a.EndsWith(".$b") -or $b.EndsWith(".$a"))
}

function Get-DmarcRowFinding {
    <#
        Eine Zeile -> Befund-Schlüssel ('' = unauffällig). Der eine gemeinsame Weg für
        Quellen-Tabelle, Empfehlungen und CSV - damit nichts nur an einer Stelle sichtbar ist.
          Weiterleitung  Quelle ist ein bekannter Weiterleiter
          Fail           DMARC nicht bestanden
          DkimGebrochen  Signatur der eigenen Domain vorhanden, aber ungültig; Pass nur über SPF
          NurSpf         keine Signatur der eigenen Domain; Pass nur über SPF
          ZweitSignatur  eine eigene Signatur gültig, eine weitere gebrochen (harmlos)
    #>
    param($Row)
    if ("$($Row.Service)" -like 'Weiterleitung*') { return 'Weiterleitung' }
    # Der Empfänger selbst meldet die Weiterleitung (policy_evaluated/reason).
    if ("$($Row.Reason)" -match '(^|; )(forwarded|mailing_list|trusted_forwarder)\b') { return 'Weiterleitung' }
    if (-not $Row.DmarcPass) { return 'Fail' }
    $hf = if ($Row.HeaderFrom) { $Row.HeaderFrom } else { $Row.Domain }
    $broken = $false
    foreach ($p in ("$($Row.DkimAuth)" -split ';')) {
        $kv = $p.Trim() -split '=', 2
        if ($kv.Count -eq 2 -and $kv[1] -in 'fail', 'permerror', 'temperror' -and (Test-DomainRelated $kv[0] $hf)) { $broken = $true }
    }
    if (-not $Row.DkimAligned) { if ($broken) { return 'DkimGebrochen' } else { return 'NurSpf' } }
    if ($broken) { return 'ZweitSignatur' }
    return ''
}

function Get-FindingLabel {
    param([string]$Key)
    switch ($Key) {
        'Weiterleitung' { 'Weiterleitung' }
        'Fail'          { 'DMARC fail' }
        'DkimGebrochen' { 'DKIM gebrochen, nur SPF' }
        'NurSpf'        { 'kein DKIM, nur SPF' }
        'ZweitSignatur' { 'zweite Signatur gebrochen (harmlos)' }
        default         { $Key }
    }
}

function Format-DayList {
    <# ['2026-11-16','2026-11-17'] -> '16.11., 17.11.' (höchstens 5, dann '…'). #>
    param([string[]]$Days)
    $d = @($Days | Sort-Object -Unique)
    $s = (@($d | Select-Object -First 5 | ForEach-Object { ([datetime]$_).ToString('dd.MM.') }) -join ', ')
    if ($d.Count -gt 5) { $s += " … (zuletzt $(([datetime]$d[-1]).ToString('dd.MM.')))" }
    return $s
}

function Get-DmarcSources {
    <#
        Zeilen -> je Domain und Quell-IP zusammengefasst, mit Zeitraum, rohen DKIM-Ergebnissen
        und Befunden (Anzahl Mails und Tage je Befund). Id dient als Sprungziel im HTML.
    #>
    param([object[]]$Rows)
    $i = 0
    $Rows | Group-Object Domain, SourceIp | ForEach-Object {
        $g = $_.Group
        $total = ($g | Measure-Object Count -Sum).Sum
        $sum = { param($cond) (($g | Where-Object $cond | Measure-Object Count -Sum).Sum) + 0 }
        $findings = @{}; $findingDays = @{}
        foreach ($r in $g) {
            $k = Get-DmarcRowFinding $r
            if (-not $k) { continue }
            $findings[$k] = [int]$findings[$k] + $r.Count
            $findingDays[$k] = @($findingDays[$k]) + $r.Begin.ToString('yyyy-MM-dd') | Where-Object { $_ } | Sort-Object -Unique
        }
        $days = @($g | ForEach-Object { $_.Begin } | Sort-Object)
        [pscustomobject]@{
            FirstDay     = $days[0]
            LastDay      = $days[-1]
            DkimResults  = (@($g | ForEach-Object { if ($_.DkimAuth) { $_.DkimAuth } else { '(keine Signatur)' } } | Sort-Object -Unique) -join ' | ')
            Findings     = $findings
            FindingDays  = $findingDays
            Domain       = $g[0].Domain
            SourceIp     = $g[0].SourceIp
            Ptr          = $g[0].Ptr
            Service      = $g[0].Service
            Count        = $total
            DmarcPass    = & $sum { $_.DmarcPass }
            DkimAligned  = & $sum { $_.DkimAligned }
            SpfAligned   = & $sum { $_.SpfAligned }
            Rejected     = & $sum { $_.Disposition -eq 'reject' }
            Quarantined  = & $sum { $_.Disposition -eq 'quarantine' }
            Reporters    = (@($g.Reporter | Sort-Object -Unique) -join ', ')
            HeaderFrom   = (@($g.HeaderFrom | Sort-Object -Unique) -join ', ')
        }
    } | Sort-Object Domain, @{ Expression = 'Count'; Descending = $true } |
        ForEach-Object { $i++; $_ | Add-Member -NotePropertyName Id -NotePropertyValue "src-$i" -PassThru }
}

function Get-DmarcDomainSummary {
    param([object[]]$Rows)
    $Rows | Group-Object Domain | ForEach-Object {
        $g = $_.Group
        $total = ($g | Measure-Object Count -Sum).Sum
        $pass  = (($g | Where-Object DmarcPass | Measure-Object Count -Sum).Sum) + 0
        $latest = $g | Sort-Object End -Descending | Select-Object -First 1
        [pscustomobject]@{
            Domain   = $_.Name
            From     = ($g | Sort-Object Begin | Select-Object -First 1).Begin
            To       = $latest.End
            Count    = $total
            Pass     = $pass
            PassPct  = $(if ($total) { [math]::Round(100.0 * $pass / $total, 1) } else { 0 })
            Policy   = $latest.Policy
            Pct      = $latest.Pct
            Quarantined = (($g | Where-Object Disposition -eq 'quarantine' | Measure-Object Count -Sum).Sum) + 0
            Rejected    = (($g | Where-Object Disposition -eq 'reject' | Measure-Object Count -Sum).Sum) + 0
            Reports  = @($g | Select-Object Reporter, ReportId -Unique).Count
            Sources  = @($g.SourceIp | Sort-Object -Unique).Count
            Reporters = (@($g.Reporter | Sort-Object -Unique) -join ', ')
        }
    }
}

function Get-DmarcRecommendation {
    <#
        Reine Regeln, testbar. Eingabe: Domain-Zusammenfassungen und Quellen.
        Ausgabe: Objekte mit Domain, Stufe (Hoch/Mittel/Info) und Text.
    #>
    param([object[]]$Domains, [object[]]$Sources)
    $Sources = @($Sources | Where-Object { $_ })
    foreach ($s in $Sources) {
        $f = if ($s.Findings) { $s.Findings } else { @{} }
        $am = { param($k) if ($s.FindingDays -and $s.FindingDays[$k]) { " (am $(Format-DayList $s.FindingDays[$k]))" } else { '' } }
        $failCount = $s.Count - $s.DmarcPass
        $failShare = if ($s.Count) { $failCount / $s.Count } else { 0 }
        $label = if ($s.Service) { "$($s.Service) ($($s.SourceIp))" } elseif ($s.Ptr) { "$($s.Ptr) ($($s.SourceIp))" } else { $s.SourceIp }
        $new = { param($stufe, $text) [pscustomobject]@{ Domain = $s.Domain; Stufe = $stufe; Text = $text; Anchor = $s.Id } }
        if ("$($s.Service)" -like 'Weiterleitung*') {
            # Weiterleitungen sind vom Absender aus nicht zu beheben - nur melden, wenn dabei etwas durchfiel.
            if ($failCount -gt 0) { & $new 'Info' "$label hat $failCount Mails ohne DMARC-Pass weitergeleitet$(& $am 'Weiterleitung'). Die Weiterleitung hat dabei die DKIM-Signatur verändert; vom Absender aus nicht zu beheben, nur zur Kenntnis." }
            continue
        }
        if ($failShare -ge 0.5) {
            if ($s.Service) {
                & $new 'Hoch' "$label sendet $($s.Count) Mails, davon $failCount ohne DMARC-Pass$(& $am 'Fail'). Ist der Dienst gewollt: DKIM mit der eigenen Domain einrichten (bevorzugt) oder den Dienst in den SPF-Eintrag aufnehmen."
            } else {
                & $new 'Hoch' "$label sendet $($s.Count) Mails ohne DMARC-Pass$(& $am 'Fail') und ist keinem bekannten Dienst zuzuordnen. Prüfen: vergessener eigener Dienst oder Missbrauch der Domain."
            }
        }
        if ($f['DkimGebrochen']) {
            & $new 'Mittel' "$label`: $($f['DkimGebrochen']) Mails mit gebrochener DKIM-Signatur der eigenen Domain$(& $am 'DkimGebrochen'), DMARC bestand nur über SPF. DKIM ist also eingerichtet, aber die Mail wurde nach dem Signieren verändert - etwa durch ein Signatur- oder Disclaimer-Gateway, einen Virenscanner oder einen Verteiler. Diesen Schritt vor das Signieren legen."
        }
        if ($f['NurSpf']) {
            & $new 'Mittel' "$label`: $($f['NurSpf']) Mails ohne DKIM-Signatur der eigenen Domain$(& $am 'NurSpf'), DMARC bestand nur über SPF. Bei Weiterleitungen geht SPF verloren - für diesen Versandweg DKIM mit der eigenen Domain einrichten oder ihn über einen signierenden Dienst leiten."
        }
    }
    foreach ($d in $Domains) {
        $tage = if ($d.From -and $d.To) { ($d.To - $d.From).TotalDays } else { 0 }
        # Strengere Stufe erst, wenn keine Mail mehr allein an SPF hängt: Unter quarantine/reject
        # fiele so eine Mail bei jeder Weiterleitung durch.
        $spfOnly = 0; $spfDays = @()
        foreach ($s in ($Sources | Where-Object Domain -eq $d.Domain)) {
            foreach ($k in 'DkimGebrochen', 'NurSpf') {
                if ($s.Findings -and $s.Findings[$k]) { $spfOnly += $s.Findings[$k]; $spfDays += @($s.FindingDays[$k]) }
            }
        }
        $vorher = if ($spfOnly) { "Vor dem nächsten Schritt erst die Mittel-Punkte klären: $spfOnly Mails bestanden nur über SPF (zuletzt am $(([datetime](@($spfDays | Sort-Object)[-1])).ToString('dd.MM.yyyy'))). Unter einer strengeren Policy würden solche Mails nach einer Weiterleitung abgewiesen." }
        switch ($d.Policy) {
            'none' {
                if ($d.PassPct -ge 98 -and $tage -ge 14) {
                    if ($spfOnly) { [pscustomobject]@{ Domain = $d.Domain; Stufe = 'Info'; Text = "Policy ist p=none, $($d.PassPct) % bestehen DMARC über $([int]$tage) Tage. $vorher Danach: p=quarantine, zunächst mit pct=25."; Anchor = '' } }
                    else { [pscustomobject]@{ Domain = $d.Domain; Stufe = 'Info'; Text = "Policy ist p=none, $($d.PassPct) % bestehen DMARC über $([int]$tage) Tage. Nächster Schritt: p=quarantine, zunächst mit pct=25."; Anchor = '' } }
                } else {
                    [pscustomobject]@{ Domain = $d.Domain; Stufe = 'Info'; Text = "Policy ist p=none (nur beobachten). Für eine strengere Stufe sollten über mindestens 14 Tage mehr als 98 % bestehen - aktuell $($d.PassPct) % über $([int]$tage) Tage."; Anchor = '' }
                }
            }
            'quarantine' {
                if ($d.PassPct -ge 99 -and $tage -ge 14) {
                    if ($spfOnly) { [pscustomobject]@{ Domain = $d.Domain; Stufe = 'Info'; Text = "Policy ist p=quarantine bei $($d.PassPct) % Pass über $([int]$tage) Tage. $vorher Danach: p=reject."; Anchor = '' } }
                    else { [pscustomobject]@{ Domain = $d.Domain; Stufe = 'Info'; Text = "Policy ist p=quarantine bei $($d.PassPct) % Pass über $([int]$tage) Tage. Nächster Schritt: p=reject."; Anchor = '' } }
                }
            }
        }
        if ($d.Pct -and $d.Pct -ne '100' -and $d.Policy -ne 'none') {
            [pscustomobject]@{ Domain = $d.Domain; Stufe = 'Info'; Text = "pct=$($d.Pct): die Policy gilt nur für einen Teil der Mails. Auf 100 erhöhen, sobald keine Fehlquellen mehr auftauchen."; Anchor = '' }
        }
    }
}

function Test-IpInCidr {
    <# Liegt die Adresse im Netz? IPv4 und IPv6, ohne .NET-8-Typen (läuft auch in Windows PowerShell 5.1). #>
    param([string]$Ip, [string]$Cidr)
    $teil = $Cidr -split '/', 2
    $a = $null; $n = $null
    if (-not [System.Net.IPAddress]::TryParse($Ip, [ref]$a) -or -not [System.Net.IPAddress]::TryParse($teil[0], [ref]$n)) { return $false }
    if ($a.AddressFamily -ne $n.AddressFamily) { return $false }
    $ab = $a.GetAddressBytes(); $nb = $n.GetAddressBytes()
    $bits = if ($teil.Count -eq 2) { [int]$teil[1] } else { $ab.Length * 8 }
    for ($i = 0; $i -lt $ab.Length -and $bits -gt 0; $i++) {
        $m = if ($bits -ge 8) { 0xFF } else { (0xFF -shl (8 - $bits)) -band 0xFF }
        if (($ab[$i] -band $m) -ne ($nb[$i] -band $m)) { return $false }
        $bits -= 8
    }
    return $true
}

function Get-Microsoft365Ranges {
    <#
        Adressbereiche von Exchange Online aus der öffentlichen Liste von Microsoft. Ohne Netz oder
        bei Fehlern leer - dann bleibt es bei der Erkennung über den Rückwärtsnamen.
    #>
    try {
        $url = 'https://endpoints.office.com/endpoints/worldwide?ServiceAreas=Exchange&NoIPv6=false&clientrequestid=' + [guid]::NewGuid()
        $liste = Invoke-RestMethod -Uri $url -TimeoutSec 20 -ErrorAction Stop
        return @($liste | Where-Object { $_.ips } | ForEach-Object { $_.ips } | Sort-Object -Unique)
    } catch {
        Write-Warning "Adressbereiche von Microsoft nicht abrufbar ($($_.Exception.Message)) - Microsoft 365 nur über den Rückwärtsnamen erkannt."
        return @()
    }
}

function Test-Microsoft365Ip {
    param([string]$Ip, [string[]]$Ranges)
    foreach ($c in $Ranges) { if (Test-IpInCidr $Ip $c) { return $true } }
    return $false
}

function Get-ReporterTargets {
    <#
        Empfängerdomains, die zu einem Melder gehören. $null = nicht eingrenzbar, etwa
        "Enterprise Outlook" (Microsoft 365): Das kann jede Domain eines Tenants sein.
    #>
    param([string]$Reporter)
    $o = "$Reporter".Trim().ToLower()
    if ($o -match 'google') { return @('gmail.com', 'googlemail.com') }
    if ($o -match 'outlook\.com') { return @('outlook.com', 'outlook.de', 'hotmail.com', 'hotmail.de', 'live.com', 'live.de', 'msn.com') }
    if ($o -match 'yahoo') { return @('yahoo.com', 'yahoo.de', 'ymail.com', 'aol.com', 'aol.de') }
    if ($o -match '^[a-z0-9-]+(\.[a-z0-9-]+)+$') { return @($o) }
    return $null
}

function Select-TraceCandidates {
    <#
        Message-Trace-Zeilen -> Kandidaten für eine Quelle. Passend = zugestellt und Empfängerdomain
        gleich oder Subdomain eines Ziels. Passende zuerst, dann neueste. Rein, testbar.
    #>
    param([object[]]$Trace, [string[]]$Targets, [int]$Max = 20)
    $out = foreach ($m in @($Trace | Where-Object { $_ })) {
        $rd = ("$($m.RecipientAddress)" -split '@')[-1].ToLower()
        $passt = $false
        foreach ($z in @($Targets | Where-Object { $_ })) { if ($rd -eq $z -or $rd.EndsWith(".$z")) { $passt = $true } }
        # Nie zugestellt = nie beim Melder angekommen, kann also nicht im Report stehen.
        if ("$($m.Status)" -eq 'Failed') { $passt = $false }
        [pscustomobject]@{ Received = $m.Received; Sender = $m.SenderAddress; Recipient = $m.RecipientAddress
                           Subject = $m.Subject; Status = $m.Status; Passt = $passt }
    }
    @($out | Sort-Object @{ Expression = 'Passt'; Descending = $true }, @{ Expression = 'Received'; Descending = $true } | Select-Object -First $Max)
}

function Get-TraceCandidates {
    <#
        Je auffälliger Quelle (DMARC fail, nur SPF, DKIM gebrochen) die Mails aus dem Message Trace:
        Absender *@Domain an den Tagen des Befunds (je Tag eine Abfrage, je Domain und Tag nur
        einmal). Ergebnis: Hashtable Quellen-Id -> @{ Anzahl; Passend; Eingrenzbar; Mails }.
    #>
    param([object[]]$Sources, [object[]]$Rows, [int]$MaxAbfragen = 40)
    $ergebnis = @{}; $cache = @{}; $abfragen = 0; $i = 0
    foreach ($s in $Sources) {
        $i++
        Show-Progress -Activity 'Message Trace' -Status ("Quelle {0} von {1}, Abfragen {2}" -f $i, @($Sources).Count, $abfragen) -Current $i -Total @($Sources).Count
        if ("$($s.Service)" -like 'Weiterleitung*' -or -not $s.Findings) { continue }
        $tage = @(); foreach ($k in 'Fail', 'DkimGebrochen', 'NurSpf') { if ($s.Findings[$k]) { $tage += @($s.FindingDays[$k]) } }
        $tage = @($tage | Where-Object { $_ } | Sort-Object -Unique)
        if (-not $tage) { continue }
        $eigene = @($Rows | Where-Object { $_.Domain -eq $s.Domain -and $_.SourceIp -eq $s.SourceIp -and (Get-DmarcRowFinding $_) -in 'Fail', 'DkimGebrochen', 'NurSpf' })
        $ziele = @(); $offen = $false
        foreach ($r in $eigene) {
            if ($r.EnvelopeTo) { $ziele += $r.EnvelopeTo; continue }
            $z = Get-ReporterTargets $r.Reporter
            if ($null -eq $z) { $offen = $true } else { $ziele += $z }
        }
        $trace = @()
        foreach ($tag in $tage) {
            $key = "$($s.Domain)|$tag"
            if (-not $cache.ContainsKey($key)) {
                if ($abfragen -ge $MaxAbfragen) { Write-Warning "Message Trace: höchstens $MaxAbfragen Abfragen je Lauf, weitere Tage ausgelassen."; $cache[$key] = @(); continue }
                $abfragen++
                $von = [datetime]::SpecifyKind([datetime]$tag, [DateTimeKind]::Utc)
                try {
                    $cache[$key] = @(Get-MessageTraceV2 -SenderAddress "*@$($s.Domain)" -StartDate $von -EndDate $von.AddDays(1) -ResultSize 5000 -ErrorAction Stop)
                } catch {
                    Write-Warning "Message Trace für $($s.Domain) am $tag gescheitert: $($_.Exception.Message)"; $cache[$key] = @()
                }
            }
            $trace += $cache[$key]
        }
        $ziele = @($ziele | Sort-Object -Unique)
        $kand = @(Select-TraceCandidates -Trace $trace -Targets $ziele)
        $ergebnis[$s.Id] = [pscustomobject]@{ Anzahl = $trace.Count; Passend = @($kand | Where-Object Passt).Count
                                               Eingrenzbar = ($ziele.Count -gt 0 -and -not $offen); Mails = $kand }
    }
    Show-Progress -Activity 'Message Trace' -Completed
    return $ergebnis
}

function Connect-MessageTrace {
    <# Bestehende Verbindung nutzen, sonst ExchangeOnlineManagement laden und anmelden. $false = nicht möglich. #>
    if (Get-Command Get-MessageTraceV2 -ErrorAction SilentlyContinue) {
        try { if (Get-ConnectionInformation -ErrorAction Stop | Where-Object State -eq 'Connected') { return $true } } catch { }
    }
    if (-not (Get-Module -ListAvailable ExchangeOnlineManagement)) {
        Write-Warning 'Für -MessageTrace fehlt das Modul ExchangeOnlineManagement (Install-Module ExchangeOnlineManagement).'
        return $false
    }
    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        return [bool](Get-Command Get-MessageTraceV2 -ErrorAction SilentlyContinue)
    } catch {
        Write-Warning "Anmeldung an Exchange Online gescheitert: $($_.Exception.Message)"
        return $false
    }
}

#endregion

#region ------------------------------------------------------------ Darstellung

function New-DmarcHtmlReport {
    param([object[]]$Domains, [object[]]$Sources, [object[]]$Recommendations, [object[]]$Rows, [string]$Title, [string]$Origin,
          [hashtable]$Trace = @{})
    $e = { param($v) [System.Net.WebUtility]::HtmlEncode("$v") }
    $brk = { param($v) ([System.Net.WebUtility]::HtmlEncode("$v")) -replace '([.:])', '$1<wbr>' }
    $sb = New-Object System.Text.StringBuilder
    $fmtDate = { param($d) if ($d) { $d.ToString('dd.MM.yyyy') } else { '' } }

    [void]$sb.Append(@"
<!doctype html><html lang="de"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>$(& $e $Title)</title><style>
:root{--bg:#fff;--fg:#1f2328;--mut:#59636e;--line:#d1d9e0;--ok:#1a7f37;--bad:#cf222e;--warn:#9a6700;--card:#f6f8fa;--hl:#fff8c5}
@media (prefers-color-scheme:dark){:root{--bg:#0d1117;--fg:#e6edf3;--mut:#9198a1;--line:#3d444d;--ok:#3fb950;--bad:#f85149;--warn:#d29922;--card:#151b23;--hl:#2e2a14}}
tr.hl td{background:var(--hl)}tr.det td{border-bottom:1px solid var(--line);padding-top:0}tr.det details{margin:2px 0;box-sizing:border-box;max-width:min(100%,calc(100vw - 56px));position:sticky;left:0}tr.det li,tr.det p{white-space:normal;overflow-wrap:anywhere}tr:target td{outline:2px solid var(--warn)}a{color:inherit}
details{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:8px 14px;margin:14px 0}summary{cursor:pointer;font-weight:600}
details li{margin:3px 0}.small{font-size:12px}
body{background:var(--bg);color:var(--fg);font:14px/1.45 system-ui,Segoe UI,sans-serif;margin:0;padding:16px;max-width:1200px;margin:auto}
h1{font-size:22px;margin:0 0 4px}h2{font-size:17px;margin:28px 0 8px;border-bottom:1px solid var(--line);padding-bottom:4px}
.mut{color:var(--mut)}table{border-collapse:collapse;width:100%;margin:6px 0}th,td{text-align:left;padding:5px 6px;border-bottom:1px solid var(--line);vertical-align:top}
th{font-weight:600;color:var(--mut)}td.n{text-align:right;font-variant-numeric:tabular-nums}.wrap{overflow-x:auto}
.ok{color:var(--ok)}.bad{color:var(--bad)}.warn{color:var(--warn)}
.cards{display:flex;flex-wrap:wrap;gap:10px}.card{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:10px 14px;min-width:150px}
.card b{display:block;font-size:20px}.tag{font-size:12px;padding:1px 6px;border-radius:10px;border:1px solid var(--line)}
svg text{fill:var(--mut);font-size:10px}
td.dkim{overflow-wrap:anywhere;min-width:8em}td.auff{min-width:7em}
@media (max-width:1100px){table{font-size:13px}th,td{padding:4px 5px}}@media (max-width:700px){body{padding:10px}}
</style></head><body>
<h1>$(& $e $Title)</h1><div class="mut">Quelle: $(& $e $Origin) · erstellt $(Get-Date -Format 'dd.MM.yyyy HH:mm') · DmarcReportAnalyzer $script:Version</div>
"@)

    [void]$sb.Append(@'
<details><summary>So liest du den Bericht</summary><ul>
<li><b>Woher die Zahlen kommen:</b> Empfänger wie Microsoft oder Google schicken täglich eine Zusammenfassung: welche IP-Adresse wie viele Mails mit deiner Domain im Absender geschickt hat und ob sie die Prüfungen bestanden. Inhalte und Empfänger stehen nicht drin. Nicht jeder Empfänger schickt Reports - die Zahlen sind eine Stichprobe, keine Gesamtmenge.</li>
<li><b>DMARC-Pass</b> heißt: SPF <i>oder</i> DKIM hat bestanden <i>und</i> passt zur Absenderdomain (&bdquo;aligned&ldquo;). Damit ist bestätigt, dass die Mail wirklich von dir kommt.</li>
<li><b>DKIM ist das Wichtigere:</b> Die Signatur übersteht Weiterleitungen, SPF nicht. Ziel ist, dass jede eigene Quelle mit DKIM aligned ist. Eine Mail, die nur über SPF besteht, fällt nach einer Weiterleitung durch.</li>
<li><b>Spalte &bdquo;Auffällig&ldquo;</b> (gelb hinterlegte Zeilen): <i>kein DKIM, nur SPF</i> - der Versandweg signiert nicht; <i>DKIM gebrochen, nur SPF</i> - signiert, aber die Mail wurde danach verändert; <i>zweite Signatur gebrochen</i> - eine gültige Signatur blieb übrig, harmlos; <i>Weiterleitung</i> - jemand hat deine Mail weitergeleitet, kein fremder Absender; <i>DMARC fail</i> - nicht bestanden: fremder Dienst, vergessener eigener Dienst oder Missbrauch.</li>
<li><b>Spalte &bdquo;DKIM-Ergebnis&ldquo;:</b> je Zeile eine Kombination, wie sie in den Datensätzen vorkam (Signaturdomain=Ergebnis). Trug eine Mail mehrere Signaturen, sind sie mit &bdquo;+&ldquo; verbunden.</li>
<li><b>Übersicht je Domain:</b> &bdquo;ausgewertet&ldquo; ist die Zahl der gemeldeten Mails; wie viele davon tatsächlich in Quarantäne kamen oder abgewiesen wurden, steht darunter. Die Policy ist nur die veröffentlichte Regel.</li>
<li><b>Einzelheiten</b> unter einer auffälligen Quelle: je Datensatz Tag, Melder, Anzahl, Empfängerdomain (falls gemeldet), SPF, DKIM mit Selektor, Return-Path und Grund (etwa <i>forwarded</i>). Mit <code>-MessageTrace</code> stehen darunter die passenden Mails aus dem Message Trace des eigenen Tenants.</li>
<li><b>Empfehlungen</b> nennen die Tage und verlinken die betroffene Zeile. Alle Einzeldatensätze (je Report und Tag) stehen in der CSV daneben, mit derselben Spalte &bdquo;Auffaellig&ldquo;.</li>
<li><b>Policy-Stufen:</b> p=none beobachtet nur, p=quarantine schickt Fehlschläge in den Junk-Ordner, p=reject weist sie ab. Eine Stufe weiter erst, wenn über mindestens 14 Tage nahezu alles besteht und keine Mail mehr allein an SPF hängt.</li>
<li><b>Wie oft:</b> wöchentlich, solange du an der Policy drehst; sonst monatlich und nach jeder Änderung am Mailfluss (Connector, Gateway, neuer Versanddienst).</li>
</ul></details>
'@)

    [void]$sb.Append('<h2>Übersicht je Domain</h2><div class="cards">')
    foreach ($d in $Domains) {
        $cls = if ($d.PassPct -ge 98) { 'ok' } elseif ($d.PassPct -ge 90) { 'warn' } else { 'bad' }
        # Menge und Policy getrennt: "40 Mails · p=quarantine" las sich wie "40 Mails in Quarantäne".
        $hart = [int]$d.Quarantined + [int]$d.Rejected
        $hartCls = if ($hart -gt 0) { 'bad' } else { 'mut' }
        [void]$sb.Append("<div class='card'><span class='mut'>$(& $e $d.Domain)</span><b class='$cls'>$($d.PassPct) % Pass</b>$($d.Count) Mails ausgewertet<br><span class='$hartCls'>davon $hart in Quarantäne/abgewiesen</span><br>Policy: p=$(& $e $d.Policy)$(if ($d.Pct -and $d.Pct -ne '100') { ", pct=$(& $e $d.Pct)" })<br><span class='mut'>$(& $fmtDate $d.From) – $(& $fmtDate $d.To) · $($d.Reports) Reports · $($d.Sources) Quellen</span><br><span class='mut'>Berichtet von: $(& $e $d.Reporters)</span></div>")
    }
    [void]$sb.Append('</div>')

    [void]$sb.Append('<h2>Empfehlungen</h2>')
    if (-not $Recommendations) { [void]$sb.Append('<p class="ok">Keine Auffälligkeiten.</p>') }
    else {
        [void]$sb.Append('<div class="wrap"><table><tr><th>Stufe</th><th>Domain</th><th>Hinweis</th></tr>')
        foreach ($r in ($Recommendations | Sort-Object @{ Expression = { @{ Hoch = 0; Mittel = 1; Info = 2 }[$_.Stufe] } }, Domain)) {
            $cls = @{ Hoch = 'bad'; Mittel = 'warn'; Info = 'mut' }[$r.Stufe]
            $link = if ($r.Anchor) { " <a class='small' href='#$(& $e $r.Anchor)'>&rarr; Zeile</a>" } else { '' }
            [void]$sb.Append("<tr><td class='$cls'>$(& $e $r.Stufe)</td><td>$(& $e $r.Domain)</td><td>$(& $e $r.Text)$link</td></tr>")
        }
        [void]$sb.Append('</table></div>')
    }

    # Tagesverlauf: Mails je Tag (Report-Beginn), Pass grün, Fail rot. Inline-SVG, keine Bibliothek.
    $daily = @($Rows | Group-Object { $_.Begin.ToString('yyyy-MM-dd') } | Sort-Object Name | ForEach-Object {
        [pscustomobject]@{ Day = $_.Name
            Pass = (($_.Group | Where-Object DmarcPass | Measure-Object Count -Sum).Sum) + 0
            Fail = (($_.Group | Where-Object { -not $_.DmarcPass } | Measure-Object Count -Sum).Sum) + 0 }
    })
    if ($daily.Count -gt 0) {
        # Durchgehende Zeitachse: Tage ohne Report bleiben als Lücke sichtbar, statt zu verschwinden.
        $byDay = @{}; foreach ($d in $daily) { $byDay[$d.Day] = $d }
        $first = [datetime]::ParseExact($daily[0].Day, 'yyyy-MM-dd', $null); $last = [datetime]::ParseExact($daily[-1].Day, 'yyyy-MM-dd', $null)
        $span = @(for ($t = $first; $t -le $last; $t = $t.AddDays(1)) { $t })
        $max = ($daily | ForEach-Object { $_.Pass + $_.Fail } | Measure-Object -Maximum).Maximum
        if (-not $max) { $max = 1 }
        $pitch = 22; $left = 34; $h = 140; $top = 8
        # Beschriftung höchstens alle 44 px - "dd.MM." braucht etwa 34 px.
        $step = [math]::Max(1, [math]::Ceiling(44 / $pitch))
        $w = [math]::Max(300, $left + $span.Count * $pitch + 4)
        [void]$sb.Append("<h2>Verlauf je Tag</h2><div class='wrap'><svg width='$w' height='$($top + $h + 24)' role='img' aria-label='Mails je Tag'>")
        $yb = $top + $h
        [void]$sb.Append("<line x1='$left' y1='$top' x2='$($w - 2)' y2='$top' stroke='var(--line)' stroke-dasharray='3 3'/><line x1='$left' y1='$yb' x2='$($w - 2)' y2='$yb' stroke='var(--line)'/>")
        [void]$sb.Append("<text x='$($left - 4)' y='$($top + 4)' text-anchor='end'>$max</text><text x='$($left - 4)' y='$($yb + 4)' text-anchor='end'>0</text>")
        $i = 0
        foreach ($t in $span) {
            $key = $t.ToString('yyyy-MM-dd'); $label = $t.ToString('dd.MM.')
            $x = $left + $i * $pitch + 3
            $day = $byDay[$key]
            if ($day) {
                $hp = [math]::Round($h * $day.Pass / $max); $hf = [math]::Round($h * $day.Fail / $max)
                [void]$sb.Append("<rect x='$x' y='$($yb - $hp - $hf)' width='16' height='$hf' fill='var(--bad)'><title>$label $($day.Fail) Fail</title></rect>")
                [void]$sb.Append("<rect x='$x' y='$($yb - $hp)' width='16' height='$hp' fill='var(--ok)'><title>$label $($day.Pass) Pass</title></rect>")
            } else {
                [void]$sb.Append("<rect class='leer' x='$x' y='$($yb - 3)' width='16' height='3' fill='var(--line)'><title>$label kein Report</title></rect>")
            }
            if ($i % $step -eq 0) { [void]$sb.Append("<text x='$($x + 8)' y='$($yb + 16)' text-anchor='middle'>$label</text>") }
            $i++
        }
        [void]$sb.Append("</svg></div><div class='mut'><span class='ok'>■</span> DMARC-Pass &nbsp; <span class='bad'>■</span> DMARC-Fail &nbsp; <span style='color:var(--line)'>▬</span> Tag ohne Report</div>")
        [void]$sb.Append("<div class='mut small'>Höhe: Mails je Tag nach Beginn des Report-Zeitraums; die gestrichelte Linie ist der höchste Tag ($max Mails).</div>")
    }

    [void]$sb.Append('<h2>Quellen</h2><div class="wrap"><table><tr><th>Domain</th><th>Quelle</th><th>Dienst</th><th>Auffällig</th><th>Zeitraum</th><th class="n">Mails</th><th class="n">DMARC-Pass</th><th class="n">DKIM aligned</th><th class="n">SPF aligned</th><th class="n">Quarantäne/<wbr>Reject</th><th>DKIM-Ergebnis</th><th>Berichtet von</th></tr>')
    foreach ($s in $Sources) {
        $pct = if ($s.Count) { [math]::Round(100.0 * $s.DmarcPass / $s.Count) } else { 0 }
        $cls = if ($pct -ge 98) { 'ok' } elseif ($pct -ge 50) { 'warn' } else { 'bad' }
        # Hostnamen und IPv6 dürfen nach Punkt/Doppelpunkt umbrechen - sonst schiebt ein langer
        # Rückwärtsname (mail-...protection.outlook.com) die Tabelle über den Rand.
        $src = if ($s.Ptr) { "$(& $brk $s.SourceIp)<br><span class='mut'>$(& $brk $s.Ptr)</span>" } else { & $brk $s.SourceIp }
        $zeit = if ($s.FirstDay) { $a = $s.FirstDay.ToString('dd.MM.'); $b = $s.LastDay.ToString('dd.MM.'); if ($a -eq $b) { $a } else { "$a – $b" } } else { '' }
        $auff = @(); $hl = $false
        foreach ($k in 'Fail', 'DkimGebrochen', 'NurSpf', 'Weiterleitung', 'ZweitSignatur') {
            if (-not ($s.Findings -and $s.Findings[$k])) { continue }
            $tcls = if ($k -eq 'Fail') { 'bad' } elseif ($k -in 'DkimGebrochen', 'NurSpf') { 'warn' } else { 'mut' }
            if ($k -in 'Fail', 'DkimGebrochen', 'NurSpf') { $hl = $true }
            $auff += "<span class='$tcls'>$(& $e (Get-FindingLabel $k)): $($s.Findings[$k])</span><br><span class='mut small'>$(& $e (Format-DayList $s.FindingDays[$k]))</span>"
        }
        $trc = if ($hl) { " class='hl'" } else { '' }
        $dkimCls = if ($s.DkimAligned -lt $s.Count -and "$($s.Service)" -notlike 'Weiterleitung*') { 'n warn' } else { 'n' }
        # DKIM: eine Zeile je vorgekommener Kombination, mehrere Signaturen derselben Mail mit "+".
        $dkimHtml = (@("$($s.DkimResults)" -split ' \| ' | ForEach-Object { & $e (($_ -split ';\s*') -join ' + ') }) -join '<br>')
        [void]$sb.Append("<tr id='$(& $e $s.Id)'$trc><td>$(& $brk $s.Domain)</td><td>$src</td><td>$(& $e $s.Service)</td><td class='auff'>$($auff -join '<br>')</td><td>$zeit</td><td class='n'>$($s.Count)</td><td class='n $cls'>$pct %</td><td class='$dkimCls'>$($s.DkimAligned)</td><td class='n'>$($s.SpfAligned)</td><td class='n'>$($s.Quarantined + $s.Rejected)</td><td class='small dkim'>$dkimHtml</td><td>$(& $brk $s.Reporters)</td></tr>")
        # Einzelheiten: die auffälligen Datensätze dieser Quelle - wann, wer meldet, wohin, wie.
        $det = @($Rows | Where-Object { $_.Domain -eq $s.Domain -and $_.SourceIp -eq $s.SourceIp -and (Get-DmarcRowFinding $_) } | Sort-Object Begin -Descending)
        if ($det.Count) {
            [void]$sb.Append("<tr class='det'><td colspan='12'><details><summary>Einzelheiten ($($det.Count) Datensätze)</summary><ul class='small'>")
            foreach ($r in ($det | Select-Object -First 50)) {
                $teile = @($r.Begin.ToString('dd.MM.yyyy'), $r.Reporter, "$($r.Count) Mails", (Get-FindingLabel (Get-DmarcRowFinding $r)))
                $teile += $(if ($r.EnvelopeTo) { "an $($r.EnvelopeTo)" } else { 'Empfängerdomain nicht gemeldet' })
                $teile += "SPF $($r.SpfAuth)"
                # Mehrere Signaturen wie in der Tabelle mit "+" (in der CSV bleibt "; ").
                $teile += $(if ($r.DkimAuth) { "DKIM $(($r.DkimAuth -split '\s*;\s*') -join ' + ')" } else { 'DKIM: keine Signatur' })
                if ($r.DkimSelector) { $teile += "Selektor $(($r.DkimSelector -split '\s*;\s*') -join ' + ')" }
                if ($r.EnvelopeFrom) { $teile += "Return-Path $($r.EnvelopeFrom)" }
                if ($r.Disposition -and $r.Disposition -ne 'none') { $teile += $r.Disposition }
                if ($r.Reason) { $teile += "Grund $($r.Reason)" }
                [void]$sb.Append("<li>$(& $e ($teile -join ' · '))</li>")
            }
            [void]$sb.Append('</ul>')
            $tr = $Trace[$s.Id]
            if ($tr) {
                $wie = if ($tr.Eingrenzbar) { "davon $($tr.Passend) an Empfänger, die zum Melder passen (oben)" } else { 'der Melder lässt sich keiner Empfängerdomain zuordnen' }
                [void]$sb.Append("<p class='mut small'>Message Trace an diesen Tagen: $($tr.Anzahl) Mails von @$(& $e $s.Domain), $(& $e $wie). Kandidaten, kein Nachweis - der Report nennt keine einzelne Mail.</p>")
                if ($tr.Mails) {
                    [void]$sb.Append('<ul class="small">')
                    foreach ($m in $tr.Mails) {
                        $mc = if ($m.Passt) { '' } else { " class='mut'" }
                        $wann = if ($m.Received) { ([datetime]$m.Received).ToString('dd.MM.yyyy HH:mm') + ' UTC' } else { '' }
                        [void]$sb.Append("<li$mc>$(& $e "$wann · $($m.Sender) → $($m.Recipient) · Betreff: $($m.Subject) · $($m.Status)")</li>")
                    }
                    [void]$sb.Append('</ul>')
                }
            }
            [void]$sb.Append('</details></td></tr>')
        }
    }
    [void]$sb.Append('</table></div>')
    [void]$sb.Append("<p class='mut'>&bdquo;Aligned&ldquo; heißt: Die Domain aus DKIM bzw. SPF passt zur Absenderdomain im From. DMARC besteht, wenn mindestens eins davon aligned ist. Alle Einzeldatensätze stehen in der CSV daneben.</p></body></html>")
    return $sb.ToString()
}

#endregion

#region ------------------------------------------------------------ Graph

$script:GraphClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'   # Microsoft Graph Command Line Tools

function Get-GraphAccessToken {
    <# Einziger Weg zum Token: vorhandenes Token, sonst Browser (Standard), sonst Gerätecode. #>
    param([string]$Scope, [string]$Tenant = 'organizations', [string]$AccessToken, [switch]$UseDeviceCode)
    if ($AccessToken) { return $AccessToken }
    if ($UseDeviceCode) { return Get-GraphTokenDeviceCode -Scope $Scope -Tenant $Tenant }
    Get-GraphTokenBrowser -Scope $Scope -Tenant $Tenant
}

function ConvertTo-Base64Url {
    param([byte[]]$Bytes)
    [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

function New-PkceVerifier {
    <# 32 Zufallsbytes -> 43 Zeichen base64url (RFC 7636, 4.1). #>
    $b = New-Object byte[] 32
    $rng = [Security.Cryptography.RandomNumberGenerator]::Create()
    try { $rng.GetBytes($b) } finally { $rng.Dispose() }
    ConvertTo-Base64Url $b
}

function Get-PkceChallenge {
    <# S256: base64url(SHA256(ASCII(verifier))) (RFC 7636, 4.2). #>
    param([Parameter(Mandatory)][string]$Verifier)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { ConvertTo-Base64Url $sha.ComputeHash([Text.Encoding]::ASCII.GetBytes($Verifier)) } finally { $sha.Dispose() }
}

function ConvertFrom-AuthRedirect {
    <#
        Wertet die erste Zeile der HTTP-Anfrage aus, mit der der Browser nach der Anmeldung auf den
        lokalen Port zurückkommt. Liefert den Code; $null für fremde Anfragen (favicon, leere
        Vorab-Verbindungen), die einfach übergangen werden. Fehler und falscher state werfen.
    #>
    param([string]$RequestLine, [Parameter(Mandatory)][string]$ExpectedState)
    if ($RequestLine -notmatch '^GET\s+/?\?(\S*)\s') { return $null }
    $q = @{}
    foreach ($pair in ($Matches[1] -split '&')) {
        if (-not $pair) { continue }
        $k, $v = $pair -split '=', 2
        $q[[uri]::UnescapeDataString($k.Replace('+', ' '))] = [uri]::UnescapeDataString("$v".Replace('+', ' '))
    }
    if (-not ($q.ContainsKey('code') -or $q.ContainsKey('error'))) { return $null }
    if ($q['state'] -ne $ExpectedState) { throw 'Anmeldung verworfen: Die Antwort an den lokalen Port gehört nicht zu dieser Anmeldung (state passt nicht).' }
    if ($q['error']) {
        $txt = if ($q['error'] -eq 'access_denied') { 'Anmeldung abgelehnt oder abgebrochen' } else { 'Anmeldung fehlgeschlagen' }
        throw "$txt ($($q['error'])): $($q['error_description'])"
    }
    $q['code']
}

function Get-GraphTokenBrowser {
    <#
        Anmeldung im Browser: Autorisierungscode mit PKCE (S256), Rückkehr auf http://localhost:<freier Port>.
        Der Port wird per TcpListener nur auf der Loopback-Adresse geöffnet - ohne URL-ACL,
        ohne Adminrechte, ohne Firewall-Freigabe. Bei Microsoft ist für diese App http://localhost
        registriert; der Port ist bei localhost frei wählbar.
    #>
    param([string]$Scope, [string]$Tenant = 'organizations', [int]$TimeoutSeconds = 300)
    $base = "https://login.microsoftonline.com/$Tenant/oauth2/v2.0"
    $verifier = New-PkceVerifier
    $state = New-PkceVerifier
    $listeners = @()
    $v4 = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback, 0)
    $v4.Start(); $listeners += $v4
    $port = ([Net.IPEndPoint]$v4.LocalEndpoint).Port
    # localhost kann im Browser auch auf ::1 auflösen - denselben Port dort mit öffnen, wenn möglich.
    try { $v6 = New-Object Net.Sockets.TcpListener([Net.IPAddress]::IPv6Loopback, $port); $v6.Start(); $listeners += $v6 } catch { }
    $redirect = "http://localhost:$port/"
    $scopes = "https://graph.microsoft.com/$Scope offline_access"
    $url = "$base/authorize?client_id=$($script:GraphClientId)&response_type=code&response_mode=query" +
           "&redirect_uri=$([uri]::EscapeDataString($redirect))&scope=$([uri]::EscapeDataString($scopes))" +
           "&state=$state&code_challenge=$(Get-PkceChallenge $verifier)&code_challenge_method=S256&prompt=select_account"
    $code = $null
    try {
        Write-Host 'Anmeldung im Browser - falls sich keiner öffnet, diese Adresse aufrufen:' -ForegroundColor Yellow
        Write-Host $url -ForegroundColor Yellow
        try { Start-Process $url } catch { Write-Warning "Browser ließ sich nicht starten: $($_.Exception.Message)" }
        $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
        while (-not $code) {
            if ((Get-Date) -ge $deadline) {
                throw "Keine Anmeldung im Browser innerhalb von $([math]::Ceiling($TimeoutSeconds / 60)) Minuten. Skript neu starten - oder mit -UseDeviceCode per Gerätecode anmelden."
            }
            $l = $listeners | Where-Object { $_.Pending() } | Select-Object -First 1
            if (-not $l) { Start-Sleep -Milliseconds 200; continue }
            $client = $l.AcceptTcpClient()
            try {
                $stream = $client.GetStream(); $stream.ReadTimeout = 5000
                $line = $null
                try { $line = (New-Object IO.StreamReader($stream, [Text.Encoding]::ASCII)).ReadLine() } catch { }
                $err = $null
                try { $code = ConvertFrom-AuthRedirect -RequestLine $line -ExpectedState $state } catch { $err = $_ }
                $msg = if ($err) { 'Anmeldung fehlgeschlagen - Details im PowerShell-Fenster.' }
                       elseif ($code) { 'Anmeldung abgeschlossen. Dieses Fenster kann geschlossen werden.' } else { $null }
                $status = if ($msg) { '200 OK' } else { '404 Not Found' }
                $body = [Text.Encoding]::UTF8.GetBytes("<!doctype html><meta charset=`"utf-8`"><title>DmarcReportAnalyzer</title><p style=`"font-family:sans-serif`">$msg</p>")
                $head = [Text.Encoding]::ASCII.GetBytes("HTTP/1.1 $status`r`nContent-Type: text/html; charset=utf-8`r`nContent-Length: $($body.Length)`r`nConnection: close`r`n`r`n")
                try { $stream.Write($head, 0, $head.Length); $stream.Write($body, 0, $body.Length); $stream.Flush() } catch { }
                if ($err) { throw $err }
            } finally { $client.Close() }
        }
    } finally { foreach ($x in $listeners) { $x.Stop() } }
    $t = Invoke-RestMethod -Method POST -Uri "$base/token" -Body @{
        grant_type = 'authorization_code'; client_id = $script:GraphClientId; code = $code
        redirect_uri = $redirect; code_verifier = $verifier; scope = $scopes }
    $t.access_token
}

function Get-GraphTokenDeviceCode {
    <#
        Anmeldung per Gerätecode über REST mit der Microsoft-Graph-Kommandozeilen-App (-UseDeviceCode).
        Für Rechner ohne Browser. Bewusst ohne Graph-Modul: Connect-MgGraph scheitert in Hosts ohne
        Konsole am Listener-Fehler (gemessen 20.11.2026).
    #>
    param([string]$Scope, [string]$Tenant = 'organizations')
    $client = $script:GraphClientId
    $base = "https://login.microsoftonline.com/$Tenant/oauth2/v2.0"
    $dc = Invoke-RestMethod -Method POST -Uri "$base/devicecode" -Body @{ client_id = $client; scope = "https://graph.microsoft.com/$Scope" }
    Write-Host $dc.message -ForegroundColor Yellow
    $deadline = (Get-Date).AddSeconds([int]$dc.expires_in)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds ([int]$dc.interval)
        try {
            $t = Invoke-RestMethod -Method POST -Uri "$base/token" -Body @{ grant_type = 'urn:ietf:params:oauth:grant-type:device_code'; client_id = $client; device_code = $dc.device_code }
            return $t.access_token
        } catch {
            if ("$_" -match 'expired_token|code_expired') { break }
            if ("$_" -match 'authorization_declined') { throw 'Anmeldung abgelehnt.' }
            if ("$_" -notmatch 'authorization_pending|slow_down') { throw }
        }
    }
    throw "Keine Anmeldung innerhalb der Frist ($([int]([int]$dc.expires_in / 60)) Minuten) - der Code ist abgelaufen. Skript neu starten und mit dem neuen Code anmelden."
}

function Invoke-Graph {
    param([string]$Uri, [string]$Token)
    Invoke-RestMethod -Method GET -Uri $Uri -Headers @{ Authorization = "Bearer $Token" }
}

function Get-GraphPages {
    <#
        Alle Seiten einer Graph-Liste. Mit -DateField werden ältere Einträge ÜBERSPRUNGEN, nicht
        als Abbruchsignal genommen: Die Unterhaltungen einer Gruppe kommen nicht verlässlich nach
        Datum sortiert (gemessen 20.11.2026 - ein Abbruch beim ersten älteren Thread verlor
        die Mehrzahl der Reports).
    #>
    param([string]$Uri, [string]$Token, [datetime]$SkipBefore, [string]$DateField)
    while ($Uri) {
        $r = Invoke-Graph -Uri $Uri -Token $Token
        foreach ($v in $r.value) {
            if ($DateField -and $v.$DateField -and ([datetime]$v.$DateField) -lt $SkipBefore) { continue }
            $v
        }
        $Uri = $r.'@odata.nextLink'
    }
}

function Get-GroupReportAttachments {
    <# Gruppenpostfach: Unterhaltungen (neueste zuerst) -> Beiträge -> Anhänge mit Inhalt. #>
    param([string]$Group, [string]$Token, [datetime]$Since)
    $g = Invoke-Graph -Uri ("https://graph.microsoft.com/v1.0/groups?`$filter=mail eq '{0}'&`$select=id,displayName" -f $Group.Replace("'", "''")) -Token $Token
    if (-not $g.value) { throw "Keine Microsoft-365-Gruppe mit der Adresse '$Group' gefunden (oder kein Zugriff)." }
    $gid = $g.value[0].id
    Write-Host "Gruppe: $($g.value[0].displayName)" -ForegroundColor Gray
    Write-Host 'Lese Unterhaltungen und Anhänge ... (bei vielen Reports einige Minuten)' -ForegroundColor Gray
    $akt = 'Gruppenpostfach lesen'
    Show-Progress -Activity $akt -Status 'Unterhaltungen auflisten' -Force
    # Erst beim Server nach Datum filtern: große Gruppen haben Hunderte alter Unterhaltungen (gemessen
    # 741 Threads, 55 s nur fürs Auflisten). /threads nahm $filter nicht an (Lauf 24.11.2026);
    # /conversations unterstützt laut Learn $select und $filter. Abgelehnt oder leer -> wie bisher
    # alle Threads holen. Die Datumsprüfung hier unten gilt in beiden Fällen.
    $base = "https://graph.microsoft.com/v1.0/groups/$gid/threads?`$select=id,topic,lastDeliveredDateTime,hasAttachments&`$top=50"
    $sinceIso = $Since.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $all = @(); $serverFilter = $false; $sw = [Diagnostics.Stopwatch]::StartNew()
    try {
        $conv = @(Get-GraphPages -Uri "https://graph.microsoft.com/v1.0/groups/$gid/conversations?`$filter=lastDeliveredDateTime ge $sinceIso&`$select=id,lastDeliveredDateTime" -Token $Token)
        Write-Verbose ("Datumsfilter auf conversations: {0} Treffer seit {1} ({2:N1} s)" -f $conv.Count, $sinceIso, $sw.Elapsed.TotalSeconds)
        if ($conv.Count) {
            $all = @(foreach ($c in $conv) { Get-GraphPages -Uri "https://graph.microsoft.com/v1.0/groups/$gid/conversations/$($c.id)/threads?`$select=id,topic,lastDeliveredDateTime,hasAttachments" -Token $Token })
            $serverFilter = $true
        }
    } catch {
        # Unter 5.1 steht die Antwort von Graph (Fehlercode, Text) in ErrorDetails, nicht in der Exception.
        Write-Verbose ("Datumsfilter auf conversations abgelehnt: {0} {1}" -f $_.Exception.Message, $_.ErrorDetails.Message)
    }
    if (-not $serverFilter) {
        $sw.Restart(); $all = @(Get-GraphPages -Uri $base -Token $Token)
        Write-Verbose ("Alle Threads gelistet: {0} ({1:N1} s)" -f $all.Count, $sw.Elapsed.TotalSeconds)
    }
    $inRange = @($all | Where-Object { -not $_.lastDeliveredDateTime -or ([datetime]$_.lastDeliveredDateTime) -ge $Since })
    $n = @{ Posts = 0; PostsMitAnhang = 0; Anhaenge = 0 }
    $i = 0
    foreach ($t in $inRange) {
        $i++
        Show-Progress -Activity $akt -Status ("Unterhaltung {0} von {1}, Anhänge bisher {2}" -f $i, $inRange.Count, $n.Anhaenge) -Current $i -Total $inRange.Count
        # hasAttachments der Unterhaltung nicht als Filter nehmen - nur zählen, die Beiträge entscheiden.
        foreach ($p in (Get-GraphPages -Uri "https://graph.microsoft.com/v1.0/groups/$gid/threads/$($t.id)/posts?`$select=id,receivedDateTime,hasAttachments" -Token $Token)) {
            $n.Posts++
            if (-not $p.hasAttachments) { continue }
            $n.PostsMitAnhang++
            foreach ($a in (Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/groups/$gid/threads/$($t.id)/posts/$($p.id)/attachments" -Token $Token).value) {
                if ($a.contentBytes) { $n.Anhaenge++; [pscustomobject]@{ Name = $a.name; Bytes = [Convert]::FromBase64String($a.contentBytes) } }
            }
        }
    }
    Show-Progress -Activity $akt -Completed
    $gesamt = if ($serverFilter) { 'Filter beim Server' } else { "$($all.Count) gesamt" }
    Write-Host ("Unterhaltungen: {0}, {1} seit {2:dd.MM.yyyy}; Beiträge: {3}, davon mit Anhang {4}; Anhänge: {5}" -f `
        $gesamt, $inRange.Count, $Since, $n.Posts, $n.PostsMitAnhang, $n.Anhaenge) -ForegroundColor Gray
}

function Get-MailboxReportAttachments {
    <# Persönliches oder freigegebenes Postfach: Nachrichten mit Anhang seit $Since. #>
    param([string]$Mailbox, [string]$Token, [datetime]$Since)
    $since = $Since.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $uri = "https://graph.microsoft.com/v1.0/users/$([uri]::EscapeDataString($Mailbox))/messages?`$filter=receivedDateTime ge $since and hasAttachments eq true&`$select=id,receivedDateTime&`$top=50"
    Write-Host 'Lese Nachrichten mit Anhang ... (bei vielen Reports einige Minuten)' -ForegroundColor Gray
    $akt = 'Postfach lesen'
    Show-Progress -Activity $akt -Status 'Nachrichten auflisten' -Force
    $msgs = @(Get-GraphPages -Uri $uri -Token $Token)
    $i = 0
    foreach ($m in $msgs) {
        $i++
        Show-Progress -Activity $akt -Status ("Nachricht {0} von {1}" -f $i, $msgs.Count) -Current $i -Total $msgs.Count
        foreach ($a in (Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/users/$([uri]::EscapeDataString($Mailbox))/messages/$($m.id)/attachments" -Token $Token).value) {
            if ($a.contentBytes) { [pscustomobject]@{ Name = $a.name; Bytes = [Convert]::FromBase64String($a.contentBytes) } }
        }
    }
    Show-Progress -Activity $akt -Completed
}

#endregion

#region ------------------------------------------------------------ Ablauf

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue

$since = (Get-Date).ToUniversalTime().AddDays(-$Days)
# Reports kommen mit 1-3 Tagen Verzug - im Postfach etwas weiter zurück lesen.
$mailSince = $since.AddDays(-3)

# Dauer je Phase in den Statuszeilen, damit sichtbar wird, wo ein langer Lauf die Zeit verbringt.
$phase = [Diagnostics.Stopwatch]::StartNew()
function Get-PhaseTime { $s = '({0:N0} s)' -f $phase.Elapsed.TotalSeconds; $phase.Restart(); $s }

$files = @()
switch ($PSCmdlet.ParameterSetName) {
    'Folder' {
        if (-not (Test-Path -LiteralPath $Path -PathType Container)) { throw "Ordner nicht gefunden: $Path" }
        $origin = "Ordner $Path"
        $files = @(Get-ChildItem -LiteralPath $Path -Recurse -File | Where-Object { $_.Extension -match '^\.(xml|gz|zip)$' } |
                   ForEach-Object { [pscustomobject]@{ Name = $_.Name; Bytes = [IO.File]::ReadAllBytes($_.FullName) } })
    }
    'Group' {
        $origin = "Gruppenpostfach $GroupMailbox"
        $tok = Get-GraphAccessToken -Scope 'Group.Read.All' -Tenant $Tenant -AccessToken $AccessToken -UseDeviceCode:$UseDeviceCode
        $files = @(Get-GroupReportAttachments -Group $GroupMailbox -Token $tok -Since $mailSince)
    }
    'Mailbox' {
        $origin = "Postfach $Mailbox"
        $tok = Get-GraphAccessToken -Scope 'Mail.Read.Shared' -Tenant $Tenant -AccessToken $AccessToken -UseDeviceCode:$UseDeviceCode
        $files = @(Get-MailboxReportAttachments -Mailbox $Mailbox -Token $tok -Since $mailSince)
    }
}
Write-Host "Anhänge/Dateien gelesen: $($files.Count) $(Get-PhaseTime)" -ForegroundColor Gray

$reports = @(); $broken = 0; $i = 0
foreach ($f in $files) {
    $i++
    Show-Progress -Activity 'Reports entpacken und lesen' -Status ("Datei {0} von {1}" -f $i, $files.Count) -Current $i -Total $files.Count
    try { foreach ($x in (ConvertFrom-DmarcAttachment -Bytes $f.Bytes -Name $f.Name)) { $reports += ConvertFrom-DmarcXml -Xml $x } }
    catch { $broken++; Write-Warning "Nicht lesbar: $($f.Name) - $($_.Exception.Message)" }
}
Show-Progress -Activity 'Reports entpacken und lesen' -Completed
$reports = @($reports | Where-Object { $_.End -and $_.End -ge $since })
$rows = @(Get-DmarcRows -Reports $reports)
Write-Host ("Reports im Zeitraum: {0}, Datensätze: {1}, nicht lesbar: {2} {3}" -f $reports.Count, $rows.Count, $broken, (Get-PhaseTime)) -ForegroundColor Gray
if ($rows.Count -eq 0) { Write-Warning "Keine DMARC-Daten in den letzten $Days Tagen gefunden."; exit 2 }

if (-not $NoReverseDns) {
    $ips = @($rows.SourceIp | Sort-Object -Unique)
    Write-Host "Reverse-DNS für $($ips.Count) Quellen ... (abschalten mit -NoReverseDns)" -ForegroundColor Gray
    Resolve-PtrBatch -Ips $ips
    Write-Host "Reverse-DNS fertig $(Get-PhaseTime)" -ForegroundColor Gray
    foreach ($r in $rows) { $r.Ptr = $script:PtrCache[$r.SourceIp]; $r.Service = Get-SenderService $r.Ptr }
}
if (-not $NoMicrosoftRanges) {
    # Viele Ausgänge von Exchange Online haben keinen sprechenden Rückwärtsnamen (vor allem IPv6).
    $msRanges = @(Get-Microsoft365Ranges)
    if ($msRanges.Count) {
        $msIp = @{}
        foreach ($ip in @($rows.SourceIp | Sort-Object -Unique)) { $msIp[$ip] = Test-Microsoft365Ip $ip $msRanges }
        foreach ($r in $rows) { if (-not $r.Service -and $msIp[$r.SourceIp]) { $r.Service = 'Microsoft 365' } }
    }
    Write-Host "Adressbereiche Microsoft 365 geprüft $(Get-PhaseTime)" -ForegroundColor Gray
}

$domains = @(Get-DmarcDomainSummary -Rows $rows)
$sources = @(Get-DmarcSources -Rows $rows)
$recs    = @(Get-DmarcRecommendation -Domains $domains -Sources $sources)
$trace   = @{}
if ($MessageTrace) {
    if (Connect-MessageTrace) {
        Write-Host 'Message Trace für auffällige Quellen ...' -ForegroundColor Gray
        $trace = Get-TraceCandidates -Sources $sources -Rows $rows
    }
}

if (-not $OutputPath) { $OutputPath = Join-Path (Get-Location) ("DMARC-Bericht-{0}.html" -f (Get-Date -Format 'yyyy-MM-dd')) }
# Relativ zum PowerShell-Ort auflösen, nicht zum Prozess-Arbeitsverzeichnis von .NET.
$OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
$outDir = Split-Path $OutputPath -Parent
if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
$html = New-DmarcHtmlReport -Domains $domains -Sources $sources -Recommendations $recs -Rows $rows -Title "DMARC-Bericht, letzte $Days Tage" -Origin $origin -Trace $trace
[IO.File]::WriteAllText($OutputPath, $html, (New-Object Text.UTF8Encoding($true)))
$csv = [IO.Path]::ChangeExtension($OutputPath, '.csv')
$rows | Select-Object Domain, Reporter, ReportId, @{ n = 'Begin'; e = { $_.Begin.ToString('yyyy-MM-dd HH:mm') } }, @{ n = 'End'; e = { $_.End.ToString('yyyy-MM-dd HH:mm') } },
    SourceIp, Ptr, Service, Count, DmarcPass, DkimAligned, SpfAligned, @{ n = 'Auffaellig'; e = { Get-FindingLabel (Get-DmarcRowFinding $_) } }, Disposition, HeaderFrom, EnvelopeFrom, EnvelopeTo, Reason, DkimAuth, DkimSelector, SpfAuth, Policy, Pct |
    Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8 -Delimiter ';'
if ($trace.Count) {
    $traceCsv = [IO.Path]::ChangeExtension($OutputPath, '.trace.csv')
    @(foreach ($s in $sources) { if ($trace[$s.Id]) { foreach ($m in $trace[$s.Id].Mails) {
        [pscustomobject]@{ Domain = $s.Domain; SourceIp = $s.SourceIp; Received = $m.Received; Sender = $m.Sender
                           Recipient = $m.Recipient; Subject = $m.Subject; Status = $m.Status; PasstZumMelder = $m.Passt } } } }) |
        Export-Csv -LiteralPath $traceCsv -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-Host "Trace:   $traceCsv" -ForegroundColor Green
}

Write-Host "Bericht: $OutputPath" -ForegroundColor Green
Write-Host "CSV:     $csv" -ForegroundColor Green
foreach ($r in ($recs | Where-Object Stufe -eq 'Hoch')) { Write-Warning "$($r.Domain): $($r.Text)" }
exit 0

#endregion
