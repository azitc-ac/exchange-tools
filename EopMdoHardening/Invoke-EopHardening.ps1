#Requires -Version 5.1
<#
.SYNOPSIS
    Setzt die Baseline des EOP-/MDO-Hardening-Standards.

.DESCRIPTION
    WICHTIG: Das Skript laeuft standardmaessig im Vorschaumodus. Ohne den Schalter
    -Execute wird NICHTS am Tenant geaendert - es wird lediglich angezeigt, welche
    Befehle ausgefuehrt wuerden.

    Empfohlener Ablauf:
      1. Invoke-EopAudit.ps1 -ExportJson    (Ist-Zustand sichern)
      2. Invoke-EopHardening.ps1            (Vorschau, zeigt jede geplante Aenderung)
      3. Mit dem Kunden durchgehen
      4. Invoke-EopHardening.ps1 -Execute   (Umsetzung)

    Die IDs in der Ausgabe entsprechen der Assessment-Checkliste und dem
    Best Practice Guide.

    Es gibt in EOP/MDO keine Rollback-Funktion. Der JSON-Export aus dem Audit-Skript
    ist die einzige Grundlage, um Werte zurueckzusetzen.

.PARAMETER Execute
    Fuehrt die Aenderungen tatsaechlich aus. Ohne diesen Schalter passiert nichts.

.PARAMETER BaselineProfile
    Baseline    Der in diesem Standard dokumentierte Zielzustand (Standard-Preset-Niveau
                mit den begruendeten Abweichungen). Empfohlen.
    Strict      Zusaetzlich Spam in Quarantaene, Bulk in Quarantaene, Phishing-Schwellwert 4.
    Legacy2023  Die Werte des Basisdokuments von 2023 inkl. Admin-only-Quarantaene.

.PARAMETER ContentFilterPolicy
    Name der Anti-Spam-Policy. Standard: "Default".

.PARAMETER AntiPhishPolicy
    Name der Anti-Phishing-Policy. Standard: "Office365 AntiPhish Default".

.PARAMETER MalwareFilterPolicy
    Name der Anti-Malware-Policy. Standard: "Default".

.PARAMETER OutboundSpamPolicy
    Name der Outbound-Spam-Policy. Standard: "Default".

.PARAMETER SafeLinksPolicy
    Name der Safe-Links-Policy. Ohne Angabe wird der Safe-Links-Teil uebersprungen.

.PARAMETER SafeAttachmentPolicy
    Name der Safe-Attachments-Policy. Ohne Angabe wird der Teil uebersprungen.

.PARAMETER ProtectedUsers
    Zu schuetzende Benutzer im Format "Anzeigename;E-Mail-Adresse".
    Beispiel: @("Max Mustermann;max@kunde.de","Erika Beispiel;erika@kunde.de")

.PARAMETER ProtectedDomains
    Zu schuetzende Fremd-Domaenen, z. B. @("hausbank.de","steuerberater.de")

.PARAMETER SkipSections
    Abschnitte ueberspringen. Gueltig: AntiSpam, ASF, ConnectionFilter, Outbound,
    AntiPhish, Impersonation, AntiMalware, SafeLinks, SafeAttachments, Quarantine,
    ExternalTag, Teams

.PARAMETER SkipConnect
    Baut unter keinen Umstaenden eine Verbindung auf, auch dann nicht, wenn keine
    besteht. Im Normalfall nicht noetig: das Skript erkennt eine bestehende Verbindung
    von selbst und verwendet sie weiter.

.PARAMETER ForceNewConnection
    Trennt eine bestehende Verbindung und meldet sich neu an. Sinnvoll beim Wechsel
    zwischen Kundentenants im selben PowerShell-Fenster.

.PARAMETER UserPrincipalName
    UPN fuer die Anmeldung.

.PARAMETER LogFile
    Pfad fuer das Protokoll. Standard: EOP-Hardening_<Datum>.log im aktuellen Ordner.

.EXAMPLE
    .\Invoke-EopHardening.ps1
    Vorschau. Es wird nichts geaendert.

.EXAMPLE
    .\Invoke-EopHardening.ps1 -Execute -ProtectedUsers @("Max Mustermann;max.mustermann@kunde.de")

.EXAMPLE
    .\Invoke-EopHardening.ps1 -Execute -BaselineProfile Strict -SkipSections ASF,ExternalTag

.NOTES
    Benoetigte Rollen: Security Administrator (EOP/MDO-Policies) und
    Exchange Administrator (Set-ExternalInOutlook).

    Alle Cmdlets und Parameter sind gegen die Microsoft-Learn-Referenz
    (Stand August 2026) verifiziert. Sie wurden NICHT gegen einen Produktivtenant
    getestet - der erste Lauf gehoert in die Vorschau.
#>

[CmdletBinding()]
param(
    [switch]$Execute,

    [ValidateSet('Baseline','Strict','Legacy2023')]
    [string]$BaselineProfile = 'Baseline',

    [string]$ContentFilterPolicy  = 'Default',
    [string]$AntiPhishPolicy      = 'Office365 AntiPhish Default',
    [string]$MalwareFilterPolicy  = 'Default',
    [string]$OutboundSpamPolicy   = 'Default',
    [string]$ConnectionFilterPolicy = 'Default',
    [string]$SafeLinksPolicy      = '',
    [string]$SafeAttachmentPolicy = '',

    [string[]]$ProtectedUsers   = @(),
    [string[]]$ProtectedDomains = @(),

    [ValidateSet('AntiSpam','ASF','ConnectionFilter','Outbound','AntiPhish','Impersonation',
                 'AntiMalware','SafeLinks','SafeAttachments','Quarantine','ExternalTag','Teams')]
    [string[]]$SkipSections = @(),

    [switch]$SkipConnect,
    [switch]$ForceNewConnection,
    [string]$UserPrincipalName,
    [string]$LogFile
)

$ErrorActionPreference = 'Continue'

# Mindestversion des Moduls ExchangeOnlineManagement. Gleicher Wert wie in
# Invoke-EopAudit.ps1; build/Test-Prolog.ps1 wacht darueber, dass beide ihn erzwingen.
$script:MinimumExoModuleVersion = '3.6.0'

# ===================================================================================
#  Baseline-Werte je Profil
# ===================================================================================

$B = @{}

# --- gemeinsam ---------------------------------------------------------------------
$B['BulkThreshold']                    = 6
$B['SpamAction']                       = 'MoveToJmf'
$B['HighConfidenceSpamAction']         = 'Quarantine'
$B['PhishSpamAction']                  = 'Quarantine'
$B['HighConfidencePhishAction']        = 'Quarantine'
$B['BulkSpamAction']                   = 'MoveToJmf'
$B['SpamQuarantineTag']                = 'DefaultFullAccessWithNotificationPolicy'
$B['HighConfidenceSpamQuarantineTag']  = 'DefaultFullAccessWithNotificationPolicy'
$B['PhishQuarantineTag']               = 'DefaultFullAccessWithNotificationPolicy'
$B['HighConfidencePhishQuarantineTag'] = 'AdminOnlyAccessPolicy'
$B['BulkQuarantineTag']                = 'DefaultFullAccessWithNotificationPolicy'
$B['QuarantineRetentionPeriod']        = 30
$B['PhishThresholdLevel']              = 3
$B['AuthenticationFailAction']         = 'Quarantine'
$B['SpoofQuarantineTag']               = 'DefaultFullAccessWithNotificationPolicy'
$B['DmarcRejectAction']                = 'Reject'
$B['ImpersonationAction']              = 'Quarantine'
$B['MailboxIntelligenceAction']        = 'MoveToJmf'
$B['RecipientLimitExternalPerHour']    = 500
$B['RecipientLimitInternalPerHour']    = 1000
$B['RecipientLimitPerDay']             = 1000
$B['NotificationFrequency']            = '04:00:00'

switch ($BaselineProfile) {
    'Strict' {
        $B['BulkThreshold']                 = 5
        $B['SpamAction']                    = 'Quarantine'
        $B['BulkSpamAction']                = 'Quarantine'
        $B['PhishThresholdLevel']           = 4
        $B['MailboxIntelligenceAction']     = 'Quarantine'
        $B['RecipientLimitExternalPerHour'] = 400
        $B['RecipientLimitInternalPerHour'] = 800
        $B['RecipientLimitPerDay']          = 800
    }
    'Legacy2023' {
        # Werte des Basisdokuments von 2023, inklusive Admin-only-Quarantaene.
        $B['HighConfidenceSpamQuarantineTag'] = 'AdminOnlyAccessPolicy'
        $B['PhishQuarantineTag']              = 'AdminOnlyAccessPolicy'
        $B['SpoofQuarantineTag']              = 'AdminOnlyAccessPolicy'
        $B['PhishThresholdLevel']             = 2
        $B['RecipientLimitPerDay']            = 200
        $B['RecipientLimitExternalPerHour']   = 0
        $B['RecipientLimitInternalPerHour']   = 0
    }
}

# ===================================================================================
#  Infrastruktur
# ===================================================================================

if ([string]::IsNullOrWhiteSpace($LogFile)) {
    $LogFile = Join-Path (Get-Location).Path ('EOP-Hardening_' + (Get-Date -Format 'yyyyMMdd-HHmm') + '.log')
}

$script:Planned  = New-Object System.Collections.ArrayList
$script:Applied  = 0
$script:Failed   = 0
$script:Skipped  = 0

function Write-Log {
    param([string]$Text, [string]$Color = 'Gray')
    Write-Host $Text -ForegroundColor $Color
    try { Add-Content -Path $LogFile -Value ((Get-Date -Format 'HH:mm:ss') + '  ' + $Text) -Encoding UTF8 } catch { }
}

function Write-Section {
    param([string]$Text)
    Write-Host ''
    Write-Log ('== ' + $Text) 'Cyan'
}

function Test-SectionSkipped {
    param([string]$Name)
    if ($SkipSections -contains $Name) {
        Write-Log ('   Abschnitt ' + $Name + ' uebersprungen (-SkipSections).') 'DarkGray'
        return $true
    }
    return $false
}

function Invoke-Change {
    <#
        Fuehrt eine Aenderung aus oder zeigt sie nur an.
        Der Aufrufer uebergibt eine Beschreibung und einen Scriptblock.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$Id,
        [Parameter(Mandatory=$true)][string]$Description,
        [Parameter(Mandatory=$true)][string]$CommandText,
        [Parameter(Mandatory=$true)][scriptblock]$Action
    )

    [void]$script:Planned.Add([pscustomobject]@{
        ID          = $Id
        Beschreibung= $Description
        Befehl      = $CommandText
    })

    if (-not $Execute) {
        Write-Log ('[VORSCHAU] ' + $Id + '  ' + $Description) 'Yellow'
        Write-Log ('           ' + $CommandText) 'DarkGray'
        $script:Skipped++
        return
    }

    Write-Log ('[SETZE]    ' + $Id + '  ' + $Description) 'White'
    try {
        & $Action | Out-Null
        Write-Log ('           OK') 'Green'
        $script:Applied++
    }
    catch {
        Write-Log ('           FEHLER: ' + $_.Exception.Message) 'Red'
        $script:Failed++
    }
}


function Get-SafeProperty {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p) { return $null }
    return $p.Value
}

function Get-ExoConnection {
    <#
        Liefert eine bestehende, nutzbare Verbindung zurueck - oder $null.

        Grundlage ist Get-ConnectionInformation (Modul 3.0.0 oder neuer). Microsoft
        woertlich: "This cmdlet is required because the Get-PSSession cmdlet in Windows
        PowerShell doesn't return information for REST API connections."

        IsEopSession unterscheidet: $false = Exchange Online, $true = Security & Compliance.
        Fuer State und TokenStatus dokumentiert Microsoft keine Werteliste, daher wird
        defensiv geprueft.
    #>
    param([switch]$Compliance)

    if ($null -eq (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
        return $null
    }
    $all = $null
    try { $all = @(Get-ConnectionInformation -ErrorAction Stop) } catch { return $null }
    if ($null -eq $all -or $all.Count -eq 0) { return $null }

    $result = $null
    foreach ($c in $all) {
        $isEop = Get-SafeProperty $c 'IsEopSession'
        if ($Compliance.IsPresent) {
            if ($isEop -ne $true) { continue }
        } else {
            if ($isEop -eq $true) { continue }
        }
        $state = [string](Get-SafeProperty $c 'State')
        if (-not [string]::IsNullOrWhiteSpace($state) -and $state -ne 'Connected') { continue }
        $tok = [string](Get-SafeProperty $c 'TokenStatus')
        if (-not [string]::IsNullOrWhiteSpace($tok) -and $tok -ne 'Active') { continue }
        $exp = Get-SafeProperty $c 'TokenExpiryTimeUTC'
        if ($null -ne $exp) {
            $expUtc = $null
            try { $expUtc = ([datetimeoffset]$exp).UtcDateTime } catch {
                try { $expUtc = ([datetime]$exp).ToUniversalTime() } catch { }
            }
            if ($null -ne $expUtc -and $expUtc -lt (Get-Date).ToUniversalTime()) { continue }
        }
        $result = $c
        break
    }
    return $result
}

function Write-ConnectionSummary {
    <# Zeigt an, in welchem Tenant gearbeitet wird - schuetzt vor Tenant-Verwechslung. #>
    param($Connection, [string]$Label)
    if ($null -eq $Connection) { return }
    $org = [string](Get-SafeProperty $Connection 'Organization')
    if ([string]::IsNullOrWhiteSpace($org)) { $org = [string](Get-SafeProperty $Connection 'DelegatedOrganization') }
    $upn = [string](Get-SafeProperty $Connection 'UserPrincipalName')
    $tid = [string](Get-SafeProperty $Connection 'TenantID')
    Write-Log ('   ' + $Label + ': ' + $org) 'White'
    if (-not [string]::IsNullOrWhiteSpace($upn)) { Write-Log ('   Angemeldet als : ' + $upn) 'Gray' }
    if (-not [string]::IsNullOrWhiteSpace($tid)) { Write-Log ('   Tenant-ID      : ' + $tid) 'Gray' }
}

function Test-CommandAvailable {
    param([string]$Name)
    $c = Get-Command -Name $Name -ErrorAction SilentlyContinue
    return ($null -ne $c)
}

# ===================================================================================
#  Start
# ===================================================================================

Write-Host ''
Write-Host '=============================================================' -ForegroundColor White
Write-Host ' EOP / Microsoft Defender for Office 365 - Hardening' -ForegroundColor White
Write-Host '=============================================================' -ForegroundColor White
Write-Log (' Profil     : ' + $BaselineProfile) 'White'
Write-Log (' Modus      : ' + $(if ($Execute) { 'UMSETZUNG - es wird geaendert' } else { 'VORSCHAU - es wird NICHTS geaendert' })) $(if ($Execute) { 'Red' } else { 'Yellow' })
Write-Log (' Protokoll  : ' + $LogFile) 'White'
Write-Log (' PowerShell : ' + $PSVersionTable.PSVersion.ToString()) 'White'

if ($Execute) {
    Write-Host ''
    Write-Host ' Es werden jetzt Aenderungen am Tenant vorgenommen.' -ForegroundColor Red
    Write-Host ' Wurde vorher Invoke-EopAudit.ps1 -ExportJson ausgefuehrt?' -ForegroundColor Red
    $answer = Read-Host ' Fortfahren? (ja/nein)'
    if ($answer -ne 'ja') {
        Write-Log ' Abgebrochen durch Benutzer.' 'Yellow'
        return
    }
}

Write-Section 'Verbindung'

if ($ForceNewConnection -and -not $SkipConnect) {
    Write-Log '   ForceNewConnection: bestehende Verbindungen werden getrennt.' 'Yellow'
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { }
}

$existing = Get-ExoConnection

if ($null -ne $existing) {
    Write-Log '   Bestehende Verbindung gefunden - sie wird weiterverwendet.' 'Gray'
    Write-ConnectionSummary -Connection $existing -Label 'Organisation  '

    # Warnung bei abweichendem Konto: der teuerste Irrtum beim Arbeiten mit
    # mehreren Kundentenants im selben Fenster.
    if (-not [string]::IsNullOrWhiteSpace($UserPrincipalName)) {
        $upn = [string](Get-SafeProperty $existing 'UserPrincipalName')
        if (-not [string]::IsNullOrWhiteSpace($upn) -and $upn -ne $UserPrincipalName) {
            Write-Host ''
            Write-Warning ('Die bestehende Verbindung laeuft unter ' + $upn + ', angegeben war ' + $UserPrincipalName + '.')
            Write-Warning 'Pruefen Sie den Tenant. Mit -ForceNewConnection wird neu angemeldet.'
            if ($Execute) {
                $go = Read-Host ' Trotzdem fortfahren? (ja/nein)'
                if ($go -ne 'ja') { Write-Log ' Abgebrochen durch Benutzer.' 'Yellow'; return }
            }
        }
    }
}
elseif ($SkipConnect) {
    Write-Log '   Keine bestehende Verbindung, -SkipConnect verhindert den Aufbau.' 'Yellow'
}
else {
    $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement |
              Sort-Object Version -Descending | Select-Object -First 1
    if ($null -eq $module) {
        throw ('Modul ExchangeOnlineManagement nicht gefunden. Install-Module ExchangeOnlineManagement ' +
               '-Scope CurrentUser -MinimumVersion ' + $script:MinimumExoModuleVersion)
    }
    if ($module.Version -lt [Version]$script:MinimumExoModuleVersion) {
        throw ('ExchangeOnlineManagement ' + $module.Version.ToString() + ' ist zu alt, mindestens ' +
               $script:MinimumExoModuleVersion + ' wird gebraucht. Aktualisieren mit: Update-Module ExchangeOnlineManagement')
    }
    Write-Log ('   Modulversion   : ' + $module.Version.ToString() + ' (mindestens ' + $script:MinimumExoModuleVersion + ')') 'Gray'
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    } catch { }
    Import-Module ExchangeOnlineManagement -MinimumVersion $script:MinimumExoModuleVersion -ErrorAction Stop
    Write-Log '   Keine bestehende Verbindung - es wird eine neue aufgebaut.' 'Gray'
    if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    } else {
        Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowBanner:$false -ErrorAction Stop
    }
    Write-ConnectionSummary -Connection (Get-ExoConnection) -Label 'Verbunden mit '
}

# ===================================================================================
#  ASI  Anti-Spam eingehend
# ===================================================================================

if (-not (Test-SectionSkipped 'AntiSpam')) {
    Write-Section 'ASI  Anti-Spam eingehend'

    $p = $ContentFilterPolicy

    Invoke-Change -Id 'ASI-01' -Description ('Bulk-Schwellwert auf ' + $B['BulkThreshold']) `
        -CommandText ('Set-HostedContentFilterPolicy -Identity "' + $p + '" -BulkThreshold ' + $B['BulkThreshold'] + ' -MarkAsSpamBulkMail On') `
        -Action { Set-HostedContentFilterPolicy -Identity $p -BulkThreshold $B['BulkThreshold'] -MarkAsSpamBulkMail On }

    Invoke-Change -Id 'ASI-02..06' -Description 'Aktionen und Quarantaene-Policies je Verdict' `
        -CommandText ('Set-HostedContentFilterPolicy -Identity "' + $p + '" ' +
                      '-SpamAction ' + $B['SpamAction'] + ' -SpamQuarantineTag ' + $B['SpamQuarantineTag'] + ' ' +
                      '-HighConfidenceSpamAction ' + $B['HighConfidenceSpamAction'] + ' -HighConfidenceSpamQuarantineTag ' + $B['HighConfidenceSpamQuarantineTag'] + ' ' +
                      '-PhishSpamAction ' + $B['PhishSpamAction'] + ' -PhishQuarantineTag ' + $B['PhishQuarantineTag'] + ' ' +
                      '-HighConfidencePhishAction ' + $B['HighConfidencePhishAction'] + ' -HighConfidencePhishQuarantineTag ' + $B['HighConfidencePhishQuarantineTag'] + ' ' +
                      '-BulkSpamAction ' + $B['BulkSpamAction'] + ' -BulkQuarantineTag ' + $B['BulkQuarantineTag']) `
        -Action {
            Set-HostedContentFilterPolicy -Identity $p `
                -SpamAction $B['SpamAction'] -SpamQuarantineTag $B['SpamQuarantineTag'] `
                -HighConfidenceSpamAction $B['HighConfidenceSpamAction'] -HighConfidenceSpamQuarantineTag $B['HighConfidenceSpamQuarantineTag'] `
                -PhishSpamAction $B['PhishSpamAction'] -PhishQuarantineTag $B['PhishQuarantineTag'] `
                -HighConfidencePhishAction $B['HighConfidencePhishAction'] -HighConfidencePhishQuarantineTag $B['HighConfidencePhishQuarantineTag'] `
                -BulkSpamAction $B['BulkSpamAction'] -BulkQuarantineTag $B['BulkQuarantineTag']
        }

    Invoke-Change -Id 'ASI-07/08/09' -Description 'Safety Tips, ZAP und Quarantaene-Aufbewahrung' `
        -CommandText ('Set-HostedContentFilterPolicy -Identity "' + $p + '" -InlineSafetyTipsEnabled $true ' +
                      '-SpamZapEnabled $true -PhishZapEnabled $true -QuarantineRetentionPeriod ' + $B['QuarantineRetentionPeriod']) `
        -Action {
            Set-HostedContentFilterPolicy -Identity $p `
                -InlineSafetyTipsEnabled $true -SpamZapEnabled $true -PhishZapEnabled $true `
                -QuarantineRetentionPeriod $B['QuarantineRetentionPeriod']
        }

    Write-Log '   Hinweis ASI-10: Allow-Listen werden bewusst NICHT automatisch geleert.' 'DarkYellow'
    Write-Log '   Bestehende Eintraege koennen produktiv benoetigt werden. Manuell pruefen:' 'DarkYellow'
    Write-Log ('   Get-HostedContentFilterPolicy -Identity "' + $p + '" | Format-List AllowedSenders,AllowedSenderDomains') 'DarkGray'
}

# ===================================================================================
#  ASF
# ===================================================================================

if (-not (Test-SectionSkipped 'ASF')) {
    Write-Section 'ASF  Advanced Spam Filter abschalten'

    $p = $ContentFilterPolicy

    Invoke-Change -Id 'ASF-01' -Description 'Score-Erhoeher auf Off' `
        -CommandText ('Set-HostedContentFilterPolicy -Identity "' + $p + '" -IncreaseScoreWithImageLinks Off ' +
                      '-IncreaseScoreWithNumericIps Off -IncreaseScoreWithRedirectToOtherPort Off -IncreaseScoreWithBizOrInfoUrls Off') `
        -Action {
            Set-HostedContentFilterPolicy -Identity $p `
                -IncreaseScoreWithImageLinks Off -IncreaseScoreWithNumericIps Off `
                -IncreaseScoreWithRedirectToOtherPort Off -IncreaseScoreWithBizOrInfoUrls Off
        }

    Invoke-Change -Id 'ASF-02' -Description 'Klassifizierungs-Features auf Off' `
        -CommandText ('Set-HostedContentFilterPolicy -Identity "' + $p + '" -MarkAsSpamEmptyMessages Off ' +
                      '-MarkAsSpamEmbedTagsInHtml Off -MarkAsSpamJavaScriptInHtml Off -MarkAsSpamFormTagsInHtml Off ' +
                      '-MarkAsSpamFramesInHtml Off -MarkAsSpamWebBugsInHtml Off -MarkAsSpamObjectTagsInHtml Off ' +
                      '-MarkAsSpamSensitiveWordList Off') `
        -Action {
            Set-HostedContentFilterPolicy -Identity $p `
                -MarkAsSpamEmptyMessages Off -MarkAsSpamEmbedTagsInHtml Off -MarkAsSpamJavaScriptInHtml Off `
                -MarkAsSpamFormTagsInHtml Off -MarkAsSpamFramesInHtml Off -MarkAsSpamWebBugsInHtml Off `
                -MarkAsSpamObjectTagsInHtml Off -MarkAsSpamSensitiveWordList Off
        }

    Invoke-Change -Id 'ASF-03/04/05' -Description 'SPF-Hardfail, Sender ID, Backscatter und Testmodus auf Off' `
        -CommandText ('Set-HostedContentFilterPolicy -Identity "' + $p + '" -MarkAsSpamSpfRecordHardFail Off ' +
                      '-MarkAsSpamFromAddressAuthFail Off -MarkAsSpamNdrBackscatter Off -TestModeAction None') `
        -Action {
            Set-HostedContentFilterPolicy -Identity $p `
                -MarkAsSpamSpfRecordHardFail Off -MarkAsSpamFromAddressAuthFail Off `
                -MarkAsSpamNdrBackscatter Off -TestModeAction None
        }
}

# ===================================================================================
#  CF  Connection Filter
# ===================================================================================

if (-not (Test-SectionSkipped 'ConnectionFilter')) {
    Write-Section 'CF  Connection Filter'

    Write-Log '   Hinweis CF-03: EnableSafeList wird NICHT automatisch gesetzt.' 'DarkYellow'
    Write-Log '   Die Cmdlet-Referenz kennzeichnet den Parameter als "reserved for internal Microsoft use",' 'DarkYellow'
    Write-Log '   der Konfigurationsartikel beschreibt ihn als regulaeres Feature. Ist-Zustand pruefen:' 'DarkYellow'
    Write-Log ('   Get-HostedConnectionFilterPolicy -Identity ' + $ConnectionFilterPolicy + ' | Format-List EnableSafeList') 'DarkGray'

    Write-Log '   Hinweis CF-01: Die IP Allow List wird NICHT automatisch geleert.' 'DarkYellow'
    Write-Log '   Jeden Eintrag einzeln bewerten - ein Loeschen kann den Mailfluss brechen.' 'DarkYellow'
}

# ===================================================================================
#  ASO  Anti-Spam ausgehend
# ===================================================================================

if (-not (Test-SectionSkipped 'Outbound')) {
    Write-Section 'ASO  Anti-Spam ausgehend'

    $p = $OutboundSpamPolicy

    Invoke-Change -Id 'ASO-02/03' -Description 'Empfaenger-Limits und Aktion bei Ueberschreitung' `
        -CommandText ('Set-HostedOutboundSpamFilterPolicy -Identity "' + $p + '" ' +
                      '-RecipientLimitExternalPerHour ' + $B['RecipientLimitExternalPerHour'] + ' ' +
                      '-RecipientLimitInternalPerHour ' + $B['RecipientLimitInternalPerHour'] + ' ' +
                      '-RecipientLimitPerDay ' + $B['RecipientLimitPerDay'] + ' ' +
                      '-ActionWhenThresholdReached BlockUser') `
        -Action {
            Set-HostedOutboundSpamFilterPolicy -Identity $p `
                -RecipientLimitExternalPerHour $B['RecipientLimitExternalPerHour'] `
                -RecipientLimitInternalPerHour $B['RecipientLimitInternalPerHour'] `
                -RecipientLimitPerDay $B['RecipientLimitPerDay'] `
                -ActionWhenThresholdReached BlockUser
        }

    Invoke-Change -Id 'ASO-04' -Description 'Automatische externe Weiterleitungen unterbinden' `
        -CommandText ('Set-HostedOutboundSpamFilterPolicy -Identity "' + $p + '" -AutoForwardingMode Automatic') `
        -Action { Set-HostedOutboundSpamFilterPolicy -Identity $p -AutoForwardingMode Automatic }

    Invoke-Change -Id 'ASO-07' -Description 'Remote Domain: automatische Weiterleitung unterbinden' `
        -CommandText 'Set-RemoteDomain -Identity Default -AutoForwardEnabled $false' `
        -Action { Set-RemoteDomain -Identity Default -AutoForwardEnabled $false }

    Write-Log '   Hinweis ASO-07: Nur die DEFAULT-Remote-Domain umstellen.' 'DarkYellow'
    Write-Log '   In Hybrid braucht die Remote Domain fuer *.mail.onmicrosoft.com weiterhin $true.' 'DarkYellow'
}

# ===================================================================================
#  TEAMS  Microsoft Teams Protection
# ===================================================================================

if (-not (Test-SectionSkipped 'Teams')) {
    Write-Section 'TEAMS  Microsoft Teams Protection'

    if (-not (Test-CommandAvailable 'Set-TeamsProtectionPolicy')) {
        Write-Log '   Uebersprungen: Set-TeamsProtectionPolicy nicht verfuegbar (keine Defender-Lizenz).' 'DarkGray'
    }
    else {
        Invoke-Change -Id 'TEAMS-01' -Description 'ZAP fuer Teams-Nachrichten und Admin-Quarantaene' `
            -CommandText ('Set-TeamsProtectionPolicy -Identity "Teams Protection Policy" -ZapEnabled $true ' +
                          '-HighConfidencePhishQuarantineTag AdminOnlyAccessPolicy -MalwareQuarantineTag AdminOnlyAccessPolicy') `
            -Action {
                Set-TeamsProtectionPolicy -Identity 'Teams Protection Policy' -ZapEnabled $true `
                    -HighConfidencePhishQuarantineTag AdminOnlyAccessPolicy `
                    -MalwareQuarantineTag AdminOnlyAccessPolicy
            }
    }
}

# ===================================================================================
#  APH  Anti-Phishing (EOP-Teil)
# ===================================================================================

if (-not (Test-SectionSkipped 'AntiPhish')) {
    Write-Section 'APH  Anti-Phishing (Spoof und DMARC)'

    $p = $AntiPhishPolicy

    Invoke-Change -Id 'APH-01/02' -Description 'Spoof Intelligence und Aktion bei Spoof-Erkennung' `
        -CommandText ('Set-AntiPhishPolicy -Identity "' + $p + '" -EnableSpoofIntelligence $true ' +
                      '-AuthenticationFailAction ' + $B['AuthenticationFailAction'] + ' -SpoofQuarantineTag ' + $B['SpoofQuarantineTag']) `
        -Action {
            Set-AntiPhishPolicy -Identity $p -EnableSpoofIntelligence $true `
                -AuthenticationFailAction $B['AuthenticationFailAction'] -SpoofQuarantineTag $B['SpoofQuarantineTag']
        }

    Invoke-Change -Id 'APH-03/04/05' -Description 'DMARC-Policy des Absenders durchsetzen' `
        -CommandText ('Set-AntiPhishPolicy -Identity "' + $p + '" -HonorDmarcPolicy $true ' +
                      '-DmarcQuarantineAction Quarantine -DmarcRejectAction ' + $B['DmarcRejectAction']) `
        -Action {
            Set-AntiPhishPolicy -Identity $p -HonorDmarcPolicy $true `
                -DmarcQuarantineAction Quarantine -DmarcRejectAction $B['DmarcRejectAction']
        }

    Invoke-Change -Id 'APH-06/07' -Description 'Absender-Indikatoren und First Contact Safety Tip' `
        -CommandText ('Set-AntiPhishPolicy -Identity "' + $p + '" -EnableUnauthenticatedSender $true ' +
                      '-EnableViaTag $true -EnableFirstContactSafetyTips $true') `
        -Action {
            Set-AntiPhishPolicy -Identity $p -EnableUnauthenticatedSender $true `
                -EnableViaTag $true -EnableFirstContactSafetyTips $true
        }

    Write-Log '   Hinweis APH-03: Steht ein Gateway vor EOP, wirkt HonorDmarcPolicy nur mit' 'DarkYellow'
    Write-Log '   aktivem Enhanced Filtering am Inbound-Connector (EF-01).' 'DarkYellow'
}

# ===================================================================================
#  IMP  Impersonation (Defender)
# ===================================================================================

if (-not (Test-SectionSkipped 'Impersonation')) {
    Write-Section 'IMP  Impersonation (Defender P1+)'

    $p = $AntiPhishPolicy
    $policyObj = $null
    try { $policyObj = Get-AntiPhishPolicy -Identity $p -ErrorAction Stop } catch { }

    $hasDefender = $false
    if ($null -ne $policyObj) {
        if ($null -ne $policyObj.PSObject.Properties['PhishThresholdLevel']) { $hasDefender = $true }
    }

    if (-not $hasDefender) {
        Write-Log '   Keine Defender-Eigenschaften in der Policy gefunden - Abschnitt uebersprungen.' 'DarkGray'
        Write-Log '   Impersonation-Schutz und Phishing-Schwellwert erfordern Defender for Office 365 P1 oder P2.' 'DarkGray'
    }
    else {
        if ($p -eq 'Office365 AntiPhish Default') {
            Write-Log '   WARNUNG: In der Default-Policy sind Impersonation-Einstellungen und der' 'DarkYellow'
            Write-Log '   Phishing-Schwellwert nicht wirksam. Fuer diesen Abschnitt eine Custom-Policy' 'DarkYellow'
            Write-Log '   anlegen und ueber -AntiPhishPolicy angeben.' 'DarkYellow'
        }

        Invoke-Change -Id 'APH-08' -Description ('Phishing-Schwellwert auf ' + $B['PhishThresholdLevel']) `
            -CommandText ('Set-AntiPhishPolicy -Identity "' + $p + '" -PhishThresholdLevel ' + $B['PhishThresholdLevel']) `
            -Action { Set-AntiPhishPolicy -Identity $p -PhishThresholdLevel $B['PhishThresholdLevel'] }

        Invoke-Change -Id 'IMP-01' -Description 'Eigene Domaenen gegen Impersonation schuetzen' `
            -CommandText ('Set-AntiPhishPolicy -Identity "' + $p + '" -EnableOrganizationDomainsProtection $true ' +
                          '-TargetedDomainProtectionAction ' + $B['ImpersonationAction'] + ' -TargetedDomainQuarantineTag ' + $B['SpamQuarantineTag']) `
            -Action {
                Set-AntiPhishPolicy -Identity $p -EnableOrganizationDomainsProtection $true `
                    -TargetedDomainProtectionAction $B['ImpersonationAction'] `
                    -TargetedDomainQuarantineTag $B['SpamQuarantineTag']
            }

        if ($ProtectedDomains.Count -gt 0) {
            Invoke-Change -Id 'IMP-02' -Description ('Fremd-Domaenen schuetzen (' + $ProtectedDomains.Count + ')') `
                -CommandText ('Set-AntiPhishPolicy -Identity "' + $p + '" -EnableTargetedDomainsProtection $true ' +
                              '-TargetedDomainsToProtect @{Add="' + ($ProtectedDomains -join '","') + '"}') `
                -Action {
                    Set-AntiPhishPolicy -Identity $p -EnableTargetedDomainsProtection $true `
                        -TargetedDomainsToProtect @{Add=$ProtectedDomains}
                }
        } else {
            Write-Log '   IMP-02 uebersprungen: keine -ProtectedDomains angegeben.' 'DarkGray'
        }

        if ($ProtectedUsers.Count -gt 0) {
            Invoke-Change -Id 'IMP-03' -Description ('Benutzer schuetzen (' + $ProtectedUsers.Count + ')') `
                -CommandText ('Set-AntiPhishPolicy -Identity "' + $p + '" -EnableTargetedUserProtection $true ' +
                              '-TargetedUsersToProtect @{Add="' + ($ProtectedUsers -join '","') + '"} ' +
                              '-TargetedUserProtectionAction ' + $B['ImpersonationAction']) `
                -Action {
                    Set-AntiPhishPolicy -Identity $p -EnableTargetedUserProtection $true `
                        -TargetedUsersToProtect @{Add=$ProtectedUsers} `
                        -TargetedUserProtectionAction $B['ImpersonationAction'] `
                        -TargetedUserQuarantineTag $B['SpamQuarantineTag']
                }
        } else {
            Write-Log '   IMP-03 uebersprungen: keine -ProtectedUsers angegeben.' 'DarkYellow'
            Write-Log '   Format: -ProtectedUsers @("Max Mustermann;max@kunde.de")' 'DarkGray'
        }

        Invoke-Change -Id 'IMP-05' -Description 'Mailbox Intelligence inklusive eigener Aktion' `
            -CommandText ('Set-AntiPhishPolicy -Identity "' + $p + '" -EnableMailboxIntelligence $true ' +
                          '-EnableMailboxIntelligenceProtection $true -MailboxIntelligenceProtectionAction ' + $B['MailboxIntelligenceAction']) `
            -Action {
                Set-AntiPhishPolicy -Identity $p -EnableMailboxIntelligence $true `
                    -EnableMailboxIntelligenceProtection $true `
                    -MailboxIntelligenceProtectionAction $B['MailboxIntelligenceAction']
            }

        Invoke-Change -Id 'IMP-06' -Description 'Impersonation Safety Tips' `
            -CommandText ('Set-AntiPhishPolicy -Identity "' + $p + '" -EnableSimilarUsersSafetyTips $true ' +
                          '-EnableSimilarDomainsSafetyTips $true -EnableUnusualCharactersSafetyTips $true') `
            -Action {
                Set-AntiPhishPolicy -Identity $p -EnableSimilarUsersSafetyTips $true `
                    -EnableSimilarDomainsSafetyTips $true -EnableUnusualCharactersSafetyTips $true
            }

        if (Test-CommandAvailable 'Set-EmailTenantSettings') {
            Invoke-Change -Id 'IMP-08' -Description 'Priority Account Protection (Defender P2)' `
                -CommandText 'Set-EmailTenantSettings -EnablePriorityAccountProtection $true' `
                -Action { Set-EmailTenantSettings -EnablePriorityAccountProtection $true }
        }
    }
}

# ===================================================================================
#  AMW  Anti-Malware
# ===================================================================================

if (-not (Test-SectionSkipped 'AntiMalware')) {
    Write-Section 'AMW  Anti-Malware'

    $p = $MalwareFilterPolicy

    Invoke-Change -Id 'AMW-01..04' -Description 'Dateityp-Filter, ZAP und Quarantaene-Policy' `
        -CommandText ('Set-MalwareFilterPolicy -Identity "' + $p + '" -EnableFileFilter $true ' +
                      '-FileTypeAction Reject -ZapEnabled $true -QuarantineTag AdminOnlyAccessPolicy') `
        -Action {
            Set-MalwareFilterPolicy -Identity $p -EnableFileFilter $true `
                -FileTypeAction Reject -ZapEnabled $true -QuarantineTag AdminOnlyAccessPolicy
        }

    Write-Log '   Hinweis AMW-01: Die Dateityp-Liste wird NICHT veraendert.' 'DarkYellow'
    Write-Log '   Microsoft dokumentiert keine empfohlene Erweiterung - jede Ergaenzung ist eine' 'DarkYellow'
    Write-Log '   eigene Entscheidung und sollte vorher im Threat Explorer gemessen werden.' 'DarkYellow'
}

# ===================================================================================
#  SL  Safe Links
# ===================================================================================

if (-not (Test-SectionSkipped 'SafeLinks')) {
    Write-Section 'SL  Safe Links (Defender P1+)'

    if ([string]::IsNullOrWhiteSpace($SafeLinksPolicy)) {
        Write-Log '   Uebersprungen: kein -SafeLinksPolicy angegeben.' 'DarkGray'
        Write-Log '   Vorhandene Policies anzeigen: Get-SafeLinksPolicy | Format-Table Name' 'DarkGray'
    }
    elseif (-not (Test-CommandAvailable 'Set-SafeLinksPolicy')) {
        Write-Log '   Uebersprungen: Set-SafeLinksPolicy nicht verfuegbar (keine Defender-Lizenz).' 'DarkGray'
    }
    else {
        $p = $SafeLinksPolicy
        Invoke-Change -Id 'SL-01..05' -Description 'Safe-Links-Kernschalter' `
            -CommandText ('Set-SafeLinksPolicy -Identity "' + $p + '" -EnableSafeLinksForEmail $true ' +
                          '-EnableSafeLinksForTeams $true -EnableSafeLinksForOffice $true -EnableForInternalSenders $true ' +
                          '-ScanUrls $true -DeliverMessageAfterScan $true -TrackClicks $true -AllowClickThrough $false ' +
                          '-DisableUrlRewrite $false') `
            -Action {
                Set-SafeLinksPolicy -Identity $p `
                    -EnableSafeLinksForEmail $true -EnableSafeLinksForTeams $true -EnableSafeLinksForOffice $true `
                    -EnableForInternalSenders $true -ScanUrls $true -DeliverMessageAfterScan $true `
                    -TrackClicks $true -AllowClickThrough $false -DisableUrlRewrite $false
            }

        Invoke-Change -Id 'SL-07' -Description 'Organisationsbranding auf den Warnseiten' `
            -CommandText ('Set-SafeLinksPolicy -Identity "' + $p + '" -EnableOrganizationBranding $true') `
            -Action { Set-SafeLinksPolicy -Identity $p -EnableOrganizationBranding $true }
    }
}

# ===================================================================================
#  SA  Safe Attachments
# ===================================================================================

if (-not (Test-SectionSkipped 'SafeAttachments')) {
    Write-Section 'SA  Safe Attachments (Defender P1+)'

    if (-not (Test-CommandAvailable 'Set-SafeAttachmentPolicy')) {
        Write-Log '   Uebersprungen: Set-SafeAttachmentPolicy nicht verfuegbar (keine Defender-Lizenz).' 'DarkGray'
    }
    else {
        if (-not [string]::IsNullOrWhiteSpace($SafeAttachmentPolicy)) {
            $p = $SafeAttachmentPolicy
            Invoke-Change -Id 'SA-01' -Description 'Safe Attachments aktivieren, Aktion Block' `
                -CommandText ('Set-SafeAttachmentPolicy -Identity "' + $p + '" -Enable $true -Action Block ' +
                              '-QuarantineTag AdminOnlyAccessPolicy -Redirect $false') `
                -Action {
                    Set-SafeAttachmentPolicy -Identity $p -Enable $true -Action Block `
                        -QuarantineTag AdminOnlyAccessPolicy -Redirect $false
                }
        } else {
            Write-Log '   SA-01 uebersprungen: kein -SafeAttachmentPolicy angegeben.' 'DarkGray'
        }

        if (Test-CommandAvailable 'Set-AtpPolicyForO365') {
            Invoke-Change -Id 'SA-03' -Description 'Safe Attachments fuer SharePoint, OneDrive und Teams' `
                -CommandText 'Set-AtpPolicyForO365 -Identity Default -EnableATPForSPOTeamsODB $true' `
                -Action { Set-AtpPolicyForO365 -Identity Default -EnableATPForSPOTeamsODB $true }

            Write-Log '   Hinweis SA-03: Zusaetzlich noetig, damit erkannte Dateien nicht heruntergeladen' 'DarkYellow'
            Write-Log '   werden koennen (SharePoint Online Management Shell):' 'DarkYellow'
            Write-Log '   Set-SPOTenant -DisallowInfectedFileDownload $true' 'DarkGray'
            Write-Log '   Hinweis SA-04: Safe Documents braucht M365 E5 bzw. E5 Security - hier nicht gesetzt.' 'DarkYellow'
        }
    }
}

# ===================================================================================
#  QUA  Quarantaene
# ===================================================================================

if (-not (Test-SectionSkipped 'Quarantine')) {
    Write-Section 'QUA  Quarantaene-Benachrichtigungen'

    Invoke-Change -Id 'QUA-03' -Description ('Benachrichtigungsfrequenz auf ' + $B['NotificationFrequency']) `
        -CommandText ('Get-QuarantinePolicy -QuarantinePolicyType GlobalQuarantinePolicy | ' +
                      'Set-QuarantinePolicy -EndUserSpamNotificationFrequency ' + $B['NotificationFrequency']) `
        -Action {
            Get-QuarantinePolicy -QuarantinePolicyType GlobalQuarantinePolicy |
                Set-QuarantinePolicy -EndUserSpamNotificationFrequency $B['NotificationFrequency']
        }

    Write-Log '   Hinweis QUA-04: Branding und eigene Absenderadresse werden nicht automatisch gesetzt,' 'DarkYellow'
    Write-Log '   weil sie ein konfiguriertes M365-Organisationsdesign und eine gueltige interne' 'DarkYellow'
    Write-Log '   Absenderadresse voraussetzen. Manuell:' 'DarkYellow'
    Write-Log '   Get-QuarantinePolicy -QuarantinePolicyType GlobalQuarantinePolicy | Set-QuarantinePolicy -OrganizationBrandingEnabled $true -EndUserSpamNotificationCustomFromAddress "quarantaene@kunde.de"' 'DarkGray'
}

# ===================================================================================
#  EXT  Externe Kennzeichnung
# ===================================================================================

if (-not (Test-SectionSkipped 'ExternalTag')) {
    Write-Section 'EXT  Kennzeichnung externer E-Mails'

    Invoke-Change -Id 'EXT-01' -Description 'Natives External-Tag aktivieren' `
        -CommandText 'Set-ExternalInOutlook -Enabled $true' `
        -Action { Set-ExternalInOutlook -Enabled $true }

    Write-Log '   Hinweis: Wirksam nach 24 bis 48 Stunden, nur fuer neu eingehende Nachrichten.' 'DarkGray'
}

# ===================================================================================
#  Abschluss
# ===================================================================================

Write-Host ''
Write-Host '-------------------------------------------------------------' -ForegroundColor White
if ($Execute) {
    Write-Log (' Umgesetzt : ' + $script:Applied) 'Green'
    Write-Log (' Fehler    : ' + $script:Failed)  $(if ($script:Failed -gt 0) { 'Red' } else { 'Green' })
} else {
    Write-Log (' Geplante Aenderungen: ' + $script:Planned.Count) 'Yellow'
    Write-Log ' Es wurde NICHTS geaendert. Zur Umsetzung mit -Execute erneut ausfuehren.' 'Yellow'
}
Write-Log (' Protokoll : ' + $LogFile) 'White'
Write-Host '-------------------------------------------------------------' -ForegroundColor White

Write-Host ''
Write-Host ' Nicht automatisiert - bewusst manuell zu entscheiden:' -ForegroundColor White
Write-Host '   ASI-10  Allow-Listen der Anti-Spam-Policy leeren'          -ForegroundColor Gray
Write-Host '   CF-01   IP Allow List bereinigen'                          -ForegroundColor Gray
Write-Host '   TABL    Bestehende Allow- und Spoof-Eintraege durchsehen'  -ForegroundColor Gray
Write-Host '   ADV-03  Transportregeln mit Filter-Bypass abbauen'         -ForegroundColor Gray
Write-Host '   EF-01   Enhanced Filtering, wenn ein Gateway vor EOP steht'-ForegroundColor Gray
Write-Host '   AUTH    SPF, DKIM und DMARC im DNS'                        -ForegroundColor Gray
Write-Host '   ASO-08  Versand aus der onmicrosoft.com-Domaene abloesen'      -ForegroundColor Gray
Write-Host '   HYB-10  Accepted-Domain-Typ pruefen (DBEB)'                  -ForegroundColor Gray
Write-Host '   VER     Wirksamkeit nachweisen (EICAR, Spoof-Test, Header)'   -ForegroundColor Gray
Write-Host '   HYB     On-Premises- und Firewall-Themen'                  -ForegroundColor Gray
Write-Host ''

$script:Planned
