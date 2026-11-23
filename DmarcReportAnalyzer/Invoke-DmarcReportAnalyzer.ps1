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

.EXAMPLE
    .\Invoke-DmarcReportAnalyzer.ps1 -Path C:\Reports
.EXAMPLE
    .\Invoke-DmarcReportAnalyzer.ps1 -GroupMailbox dmarc@contoso.onmicrosoft.com -Days 14
.EXAMPLE
    .\Invoke-DmarcReportAnalyzer.ps1 -Mailbox dmarc-reports@contoso.com -OutputPath C:\Temp\dmarc.html
.EXAMPLE
    .\Invoke-DmarcReportAnalyzer.ps1 -GroupMailbox dmarc@contoso.onmicrosoft.com -UseDeviceCode
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
    [switch]$NoReverseDns
)

# Einzige Stelle fuer die Versionsnummer; Build und Release-Tag lesen sie hier aus.
$script:Version = '1.1.0'

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

#endregion

#region ------------------------------------------------------------ Darstellung

function New-DmarcHtmlReport {
    param([object[]]$Domains, [object[]]$Sources, [object[]]$Recommendations, [object[]]$Rows, [string]$Title, [string]$Origin)
    $e = { param($v) [System.Net.WebUtility]::HtmlEncode("$v") }
    $sb = New-Object System.Text.StringBuilder
    $fmtDate = { param($d) if ($d) { $d.ToString('dd.MM.yyyy') } else { '' } }

    [void]$sb.Append(@"
<!doctype html><html lang="de"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>$(& $e $Title)</title><style>
:root{--bg:#fff;--fg:#1f2328;--mut:#59636e;--line:#d1d9e0;--ok:#1a7f37;--bad:#cf222e;--warn:#9a6700;--card:#f6f8fa;--hl:#fff8c5}
@media (prefers-color-scheme:dark){:root{--bg:#0d1117;--fg:#e6edf3;--mut:#9198a1;--line:#3d444d;--ok:#3fb950;--bad:#f85149;--warn:#d29922;--card:#151b23;--hl:#2e2a14}}
tr.hl td{background:var(--hl)}tr:target td{outline:2px solid var(--warn)}a{color:inherit}
details{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:8px 14px;margin:14px 0}summary{cursor:pointer;font-weight:600}
details li{margin:3px 0}.small{font-size:12px}
body{background:var(--bg);color:var(--fg);font:14px/1.45 system-ui,Segoe UI,sans-serif;margin:0;padding:16px;max-width:1200px;margin:auto}
h1{font-size:22px;margin:0 0 4px}h2{font-size:17px;margin:28px 0 8px;border-bottom:1px solid var(--line);padding-bottom:4px}
.mut{color:var(--mut)}table{border-collapse:collapse;width:100%;margin:6px 0}th,td{text-align:left;padding:5px 8px;border-bottom:1px solid var(--line);vertical-align:top}
th{font-weight:600;color:var(--mut)}td.n{text-align:right;font-variant-numeric:tabular-nums}.wrap{overflow-x:auto}
.ok{color:var(--ok)}.bad{color:var(--bad)}.warn{color:var(--warn)}
.cards{display:flex;flex-wrap:wrap;gap:10px}.card{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:10px 14px;min-width:150px}
.card b{display:block;font-size:20px}.tag{font-size:12px;padding:1px 6px;border-radius:10px;border:1px solid var(--line)}
svg text{fill:var(--mut);font-size:10px}
</style></head><body>
<h1>$(& $e $Title)</h1><div class="mut">Quelle: $(& $e $Origin) · erstellt $(Get-Date -Format 'dd.MM.yyyy HH:mm') · DmarcReportAnalyzer $script:Version</div>
"@)

    [void]$sb.Append(@'
<details><summary>So liest du den Bericht</summary><ul>
<li><b>Woher die Zahlen kommen:</b> Empfänger wie Microsoft oder Google schicken täglich eine Zusammenfassung: welche IP-Adresse wie viele Mails mit deiner Domain im Absender geschickt hat und ob sie die Prüfungen bestanden. Inhalte und Empfänger stehen nicht drin. Nicht jeder Empfänger schickt Reports - die Zahlen sind eine Stichprobe, keine Gesamtmenge.</li>
<li><b>DMARC-Pass</b> heißt: SPF <i>oder</i> DKIM hat bestanden <i>und</i> passt zur Absenderdomain (&bdquo;aligned&ldquo;). Damit ist bestätigt, dass die Mail wirklich von dir kommt.</li>
<li><b>DKIM ist das Wichtigere:</b> Die Signatur übersteht Weiterleitungen, SPF nicht. Ziel ist, dass jede eigene Quelle mit DKIM aligned ist. Eine Mail, die nur über SPF besteht, fällt nach einer Weiterleitung durch.</li>
<li><b>Spalte &bdquo;Auffällig&ldquo;</b> (gelb hinterlegte Zeilen): <i>kein DKIM, nur SPF</i> - der Versandweg signiert nicht; <i>DKIM gebrochen, nur SPF</i> - signiert, aber die Mail wurde danach verändert; <i>zweite Signatur gebrochen</i> - eine gültige Signatur blieb übrig, harmlos; <i>Weiterleitung</i> - jemand hat deine Mail weitergeleitet, kein fremder Absender; <i>DMARC fail</i> - nicht bestanden: fremder Dienst, vergessener eigener Dienst oder Missbrauch.</li>
<li><b>Empfehlungen</b> nennen die Tage und verlinken die betroffene Zeile. Alle Einzeldatensätze (je Report und Tag) stehen in der CSV daneben, mit derselben Spalte &bdquo;Auffaellig&ldquo;.</li>
<li><b>Policy-Stufen:</b> p=none beobachtet nur, p=quarantine schickt Fehlschläge in den Junk-Ordner, p=reject weist sie ab. Eine Stufe weiter erst, wenn über mindestens 14 Tage nahezu alles besteht und keine Mail mehr allein an SPF hängt.</li>
<li><b>Wie oft:</b> wöchentlich, solange du an der Policy drehst; sonst monatlich und nach jeder Änderung am Mailfluss (Connector, Gateway, neuer Versanddienst).</li>
</ul></details>
'@)

    [void]$sb.Append('<h2>Übersicht je Domain</h2><div class="cards">')
    foreach ($d in $Domains) {
        $cls = if ($d.PassPct -ge 98) { 'ok' } elseif ($d.PassPct -ge 90) { 'warn' } else { 'bad' }
        [void]$sb.Append("<div class='card'><span class='mut'>$(& $e $d.Domain)</span><b class='$cls'>$($d.PassPct) % Pass</b>$($d.Count) Mails · p=$(& $e $d.Policy)$(if ($d.Pct -and $d.Pct -ne '100') { " · pct=$(& $e $d.Pct)" })<br><span class='mut'>$(& $fmtDate $d.From) – $(& $fmtDate $d.To) · $($d.Reports) Reports · $($d.Sources) Quellen</span><br><span class='mut'>Berichtet von: $(& $e $d.Reporters)</span></div>")
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
        $max = ($daily | ForEach-Object { $_.Pass + $_.Fail } | Measure-Object -Maximum).Maximum
        if (-not $max) { $max = 1 }
        $w = [math]::Max(300, $daily.Count * 22); $h = 140
        [void]$sb.Append("<h2>Verlauf je Tag</h2><div class='wrap'><svg width='$w' height='$($h + 30)' role='img' aria-label='Mails je Tag'>")
        $i = 0
        foreach ($day in $daily) {
            $x = $i * 22 + 4
            $hp = [math]::Round($h * $day.Pass / $max); $hf = [math]::Round($h * $day.Fail / $max)
            [void]$sb.Append("<rect x='$x' y='$($h - $hp - $hf)' width='16' height='$hf' fill='var(--bad)'><title>$($day.Day): $($day.Fail) Fail</title></rect>")
            [void]$sb.Append("<rect x='$x' y='$($h - $hp)' width='16' height='$hp' fill='var(--ok)'><title>$($day.Day): $($day.Pass) Pass</title></rect>")
            if ($i % [math]::Max(1, [math]::Ceiling($daily.Count / 10)) -eq 0) { [void]$sb.Append("<text x='$x' y='$($h + 14)'>$($day.Day.Substring(5))</text>") }
            $i++
        }
        [void]$sb.Append("</svg></div><div class='mut'><span class='ok'>■</span> DMARC-Pass &nbsp; <span class='bad'>■</span> Fail · Maximum $max Mails/Tag</div>")
    }

    [void]$sb.Append('<h2>Quellen</h2><div class="wrap"><table><tr><th>Domain</th><th>Quelle</th><th>Dienst</th><th>Zeitraum</th><th class="n">Mails</th><th class="n">DMARC-Pass</th><th class="n">DKIM aligned</th><th class="n">SPF aligned</th><th class="n">Quarantäne/Reject</th><th>DKIM-Ergebnis</th><th>Auffällig</th><th>Berichtet von</th></tr>')
    foreach ($s in $Sources) {
        $pct = if ($s.Count) { [math]::Round(100.0 * $s.DmarcPass / $s.Count) } else { 0 }
        $cls = if ($pct -ge 98) { 'ok' } elseif ($pct -ge 50) { 'warn' } else { 'bad' }
        $src = if ($s.Ptr) { "$(& $e $s.SourceIp)<br><span class='mut'>$(& $e $s.Ptr)</span>" } else { & $e $s.SourceIp }
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
        [void]$sb.Append("<tr id='$(& $e $s.Id)'$trc><td>$(& $e $s.Domain)</td><td>$src</td><td>$(& $e $s.Service)</td><td>$zeit</td><td class='n'>$($s.Count)</td><td class='n $cls'>$pct %</td><td class='$dkimCls'>$($s.DkimAligned)</td><td class='n'>$($s.SpfAligned)</td><td class='n'>$($s.Quarantined + $s.Rejected)</td><td class='small'>$(& $e $s.DkimResults)</td><td>$($auff -join '<br>')</td><td>$(& $e $s.Reporters)</td></tr>")
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
    # Alle Unterhaltungen holen und hier filtern, damit die Zählung zeigt, was wo wegfällt.
    $all = @(Get-GraphPages -Uri "https://graph.microsoft.com/v1.0/groups/$gid/threads?`$select=id,topic,lastDeliveredDateTime,hasAttachments&`$top=50" -Token $Token)
    $inRange = @($all | Where-Object { -not $_.lastDeliveredDateTime -or ([datetime]$_.lastDeliveredDateTime) -ge $Since })
    $n = @{ Posts = 0; PostsMitAnhang = 0; Anhaenge = 0 }
    foreach ($t in $inRange) {
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
    Write-Host ("Unterhaltungen: {0} gesamt, {1} seit {2:dd.MM.yyyy}; Beiträge: {3}, davon mit Anhang {4}; Anhänge: {5}" -f `
        $all.Count, $inRange.Count, $Since, $n.Posts, $n.PostsMitAnhang, $n.Anhaenge) -ForegroundColor Gray
}

function Get-MailboxReportAttachments {
    <# Persönliches oder freigegebenes Postfach: Nachrichten mit Anhang seit $Since. #>
    param([string]$Mailbox, [string]$Token, [datetime]$Since)
    $since = $Since.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $uri = "https://graph.microsoft.com/v1.0/users/$([uri]::EscapeDataString($Mailbox))/messages?`$filter=receivedDateTime ge $since and hasAttachments eq true&`$select=id,receivedDateTime&`$top=50"
    foreach ($m in (Get-GraphPages -Uri $uri -Token $Token)) {
        foreach ($a in (Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/users/$([uri]::EscapeDataString($Mailbox))/messages/$($m.id)/attachments" -Token $Token).value) {
            if ($a.contentBytes) { [pscustomobject]@{ Name = $a.name; Bytes = [Convert]::FromBase64String($a.contentBytes) } }
        }
    }
}

#endregion

#region ------------------------------------------------------------ Ablauf

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue

$since = (Get-Date).ToUniversalTime().AddDays(-$Days)
# Reports kommen mit 1-3 Tagen Verzug - im Postfach etwas weiter zurück lesen.
$mailSince = $since.AddDays(-3)

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
Write-Host "Anhänge/Dateien gelesen: $($files.Count)" -ForegroundColor Gray

$reports = @(); $broken = 0
foreach ($f in $files) {
    try { foreach ($x in (ConvertFrom-DmarcAttachment -Bytes $f.Bytes -Name $f.Name)) { $reports += ConvertFrom-DmarcXml -Xml $x } }
    catch { $broken++; Write-Warning "Nicht lesbar: $($f.Name) - $($_.Exception.Message)" }
}
$reports = @($reports | Where-Object { $_.End -and $_.End -ge $since })
$rows = @(Get-DmarcRows -Reports $reports)
Write-Host ("Reports im Zeitraum: {0}, Datensätze: {1}, nicht lesbar: {2}" -f $reports.Count, $rows.Count, $broken) -ForegroundColor Gray
if ($rows.Count -eq 0) { Write-Warning "Keine DMARC-Daten in den letzten $Days Tagen gefunden."; exit 2 }

if (-not $NoReverseDns) {
    $ips = @($rows.SourceIp | Sort-Object -Unique)
    Write-Host "Reverse-DNS für $($ips.Count) Quellen ..." -ForegroundColor Gray
    foreach ($ip in $ips) { [void](Resolve-Ptr $ip) }
    foreach ($r in $rows) { $r.Ptr = $script:PtrCache[$r.SourceIp]; $r.Service = Get-SenderService $r.Ptr }
}

$domains = @(Get-DmarcDomainSummary -Rows $rows)
$sources = @(Get-DmarcSources -Rows $rows)
$recs    = @(Get-DmarcRecommendation -Domains $domains -Sources $sources)

if (-not $OutputPath) { $OutputPath = Join-Path (Get-Location) ("DMARC-Bericht-{0}.html" -f (Get-Date -Format 'yyyy-MM-dd')) }
# Relativ zum PowerShell-Ort auflösen, nicht zum Prozess-Arbeitsverzeichnis von .NET.
$OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
$outDir = Split-Path $OutputPath -Parent
if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
$html = New-DmarcHtmlReport -Domains $domains -Sources $sources -Recommendations $recs -Rows $rows -Title "DMARC-Bericht, letzte $Days Tage" -Origin $origin
[IO.File]::WriteAllText($OutputPath, $html, (New-Object Text.UTF8Encoding($true)))
$csv = [IO.Path]::ChangeExtension($OutputPath, '.csv')
$rows | Select-Object Domain, Reporter, ReportId, @{ n = 'Begin'; e = { $_.Begin.ToString('yyyy-MM-dd HH:mm') } }, @{ n = 'End'; e = { $_.End.ToString('yyyy-MM-dd HH:mm') } },
    SourceIp, Ptr, Service, Count, DmarcPass, DkimAligned, SpfAligned, @{ n = 'Auffaellig'; e = { Get-FindingLabel (Get-DmarcRowFinding $_) } }, Disposition, HeaderFrom, EnvelopeFrom, DkimAuth, SpfAuth, Policy, Pct |
    Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8 -Delimiter ';'

Write-Host "Bericht: $OutputPath" -ForegroundColor Green
Write-Host "CSV:     $csv" -ForegroundColor Green
foreach ($r in ($recs | Where-Object Stufe -eq 'Hoch')) { Write-Warning "$($r.Domain): $($r.Text)" }
exit 0

#endregion
