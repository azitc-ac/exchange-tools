#Requires -Version 5.1
<#
.SYNOPSIS
    Liest den Ist-Zustand der EOP-/Microsoft-Defender-for-Office-365-Konfiguration aus
    und vergleicht ihn gegen die Baseline des Hardening-Standards.

.DESCRIPTION
    Das Skript ist ausschliesslich lesend. Es fuehrt keine schreibenden Cmdlets aus und
    veraendert nichts am Tenant.

    Erzeugt werden:
      - eine CSV-Datei mit einer Zeile je Pruefpunkt (ID, Ist, Soll, Bewertung)
      - ein HTML-Report zur Uebergabe an den Kunden
      - optional ein vollstaendiger JSON-Export aller Policy-Objekte als Rollback-Grundlage

    Policy-Inventar und Wirksamkeit
    Das Skript liest zu jeder Policy auch die zugehoerige Regel und ermittelt daraus,
    ob die Policy ueberhaupt auf Empfaenger angewendet wird. Der Abschnitt
    "Policy-Inventar" (IDs INV-*) listet je Typ alle gefundenen Policies mit
    Regelzuordnung, Prioritaet und Empfaengerbereich auf.

    Vier Zustaende werden unterschieden:
      aktiv - Prio n                        wird angewendet, wird bewertet
      Default-Policy                        wird angewendet, wird bewertet
      Regel deaktiviert                     wirkt auf niemanden, wird NICHT bemaengelt
      ohne Regel                            im Portal unsichtbar, wird NICHT bemaengelt
      Preset - nicht editierbar             von Microsoft verwaltet, wird NICHT bemaengelt

    Befunde zu unwirksamen Policies erscheinen als "Info" mit Statusvermerk in der
    Spalte Objekt. Eine Policy, die auf niemanden angewendet wird, kann keinen
    Mangel darstellen.

    Die IDs entsprechen der Assessment-Checkliste (EOP-MDO_Assessment-Checkliste.xlsx)
    und dem Best Practice Guide.

.PARAMETER OutputFolder
    Zielordner fuer die Ausgabedateien. Standard: aktuelles Verzeichnis.

.PARAMETER CustomerName
    Name des Kunden. Erscheint im HTML-Report und in den Dateinamen.

.PARAMETER ExportJson
    Exportiert zusaetzlich alle gelesenen Policy-Objekte als JSON. Dringend empfohlen
    vor der ersten Aenderung - in EOP/MDO gibt es keine Rollback-Funktion.

.PARAMETER SkipConnect
    Baut unter keinen Umstaenden eine Verbindung auf, auch dann nicht, wenn keine
    besteht. Im Normalfall nicht noetig: das Skript erkennt eine bestehende Verbindung
    von selbst und verwendet sie weiter.

.PARAMETER ForceNewConnection
    Trennt eine bestehende Verbindung und meldet sich neu an. Sinnvoll beim Wechsel
    zwischen Kundentenants im selben PowerShell-Fenster.

.PARAMETER IncludeProtectionAlerts
    Prueft zusaetzlich die Alert Policies. Erfordert eine zweite Verbindung
    (Connect-IPPSSession) und damit eine weitere Anmeldung.

.PARAMETER UserPrincipalName
    UPN fuer die Anmeldung an Exchange Online.

.EXAMPLE
    .\Invoke-EopAudit.ps1 -CustomerName "Kunde" -ExportJson

.EXAMPLE
    .\Invoke-EopAudit.ps1 -CustomerName "Kunde" -OutputFolder C:\Temp -IncludeProtectionAlerts

.NOTES
    Benoetigte Rollen (nur lesend): Global Reader und Security Reader.
    Getestet gegen Windows PowerShell 5.1 mit dem Modul ExchangeOnlineManagement (V3).

    Alle verwendeten Cmdlets und Parameter sind gegen die Microsoft-Learn-Referenz
    (Stand August 2026) verifiziert.
#>

[CmdletBinding()]
param(
    [string]$OutputFolder = (Get-Location).Path,
    [string]$CustomerName = 'Kunde',
    [switch]$ExportJson,
    [switch]$SkipConnect,
    [switch]$ForceNewConnection,
    [switch]$IncludeProtectionAlerts,
    [string]$UserPrincipalName
)

# Einzige Stelle fuer die Versionsnummer; Build und Release-Tag lesen sie hier aus.
$script:Version = '1.0.0'


$ErrorActionPreference = 'Continue'

# ===================================================================================
#  Hilfsfunktionen
# ===================================================================================

$script:Findings = New-Object System.Collections.ArrayList
$script:RawData  = @{}

# Wirksamkeit der gerade geprueften Policy. Steht sie auf $false, werden alle
# Befunde dieser Policy auf 'Info' herabgestuft - eine Policy ohne wirksame Regel
# kann nicht maengelbehaftet sein, weil sie auf niemanden angewendet wird.
$script:CurrentPolicyActive = $true
$script:CurrentPolicyStatus = ''

function Write-Step {
    param([string]$Text)
    Write-Host ''
    Write-Host ('== ' + $Text) -ForegroundColor Cyan
}

function Write-Info {
    param([string]$Text)
    Write-Host ('   ' + $Text) -ForegroundColor Gray
}

function ConvertTo-DisplayValue {
    param($Value)
    if ($null -eq $Value) { return '<nicht gesetzt>' }
    if ($Value -is [System.Array]) {
        if ($Value.Count -eq 0) { return '<leer>' }
        return ($Value -join ', ')
    }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        $items = @()
        foreach ($i in $Value) { $items += [string]$i }
        if ($items.Count -eq 0) { return '<leer>' }
        return ($items -join ', ')
    }
    $s = [string]$Value
    if ([string]::IsNullOrWhiteSpace($s)) { return '<leer>' }
    return $s
}

# ===================================================================================
#  Fundorte  (generiert aus links.py - nicht von Hand bearbeiten)
#
#  Je Pruefpunkt: Klickpfad im Portal, Deep-Link und der Microsoft-Learn-Abschnitt,
#  der die Einstellung beschreibt. Alle Learn-Anker sind gegen die Quelldateien
#  von MicrosoftDocs verifiziert.
# ===================================================================================

$script:Links = @{
    'ADV-01' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Erweiterte Zustellung (Advanced delivery) > Reiter "SecOps-Postfach" (nur echte Postfaecher, keine Verteilergruppen)'; Portal = 'https://security.microsoft.com/advanceddelivery'; Learn = 'https://learn.microsoft.com/defender-office-365/advanced-delivery-policy-configure#use-the-microsoft-defender-portal-to-configure-secops-mailboxes-in-the-advanced-delivery-policy' }
    'ADV-02' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Erweiterte Zustellung > Reiter "Phishingsimulation" (Domaene + Sende-IP; IPv6 nur per PowerShell)'; Portal = 'https://security.microsoft.com/advanceddelivery'; Learn = 'https://learn.microsoft.com/defender-office-365/advanced-delivery-policy-configure#use-the-microsoft-defender-portal-to-configure-non-microsoft-phishing-simulations-in-the-advanced-delivery-policy' }
    'ADV-03' = @{ Nav = 'Exchange Admin Center > Nachrichtenfluss > Regeln: jede Regel auf "Spamfilterung umgehen" (SCL -1) pruefen. Solche Bypaesse gehoeren in die Erweiterte Zustellung oder die TABL'; Portal = 'https://admin.exchange.microsoft.com/#/transportrules'; Learn = 'https://learn.microsoft.com/defender-office-365/create-safe-sender-lists-in-office-365#use-mail-flow-rules' }
    'AMW-01' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antischadsoftware > <Richtlinie> > "Schutzeinstellungen" > "Aktivieren Sie den allgemeinen Anlagenfilter" > "Dateitypen auswaehlen"'; Portal = 'https://security.microsoft.com/antimalwarev2'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-malware-protection-about#common-attachments-filter-in-anti-malware-policies' }
    'AMW-02' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antischadsoftware > <Richtlinie> > "Schutzeinstellungen" > "Automatische Null-Stunden-Bereinigung fuer Schadsoftware aktivieren"'; Portal = 'https://security.microsoft.com/antimalwarev2'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-malware-protection-about#zero-hour-auto-purge-zap-in-anti-malware-policies' }
    'AMW-03' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antischadsoftware > <Richtlinie> > "Schutzeinstellungen" > "Wenn diese Dateitypen gefunden werden" (NDR ablehnen vs. Quarantaene)'; Portal = 'https://security.microsoft.com/antimalwarev2'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-malware-protection-about#common-attachments-filter-in-anti-malware-policies' }
    'AMW-04' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antischadsoftware > <Richtlinie> > "Schutzeinstellungen" > "Quarantaenerichtlinie" (Default AdminOnlyAccessPolicy)'; Portal = 'https://security.microsoft.com/antimalwarev2'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-malware-protection-about#quarantine-policies-in-anti-malware-policies' }
    'AMW-05' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antischadsoftware > <Richtlinie> > "Schutzeinstellungen" > Benachrichtigungen > Administratorbenachrichtigungen'; Portal = 'https://security.microsoft.com/antimalwarev2'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-malware-protection-about#admin-notifications-in-anti-malware-policies' }
    'APH-01' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antiphishing > <Richtlinie> > "Phishingschwellenwert & Schutz" > Abschnitt "Spoofing" > "Spoofintelligenz aktivieren"'; Portal = 'https://security.microsoft.com/antiphishing'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about#spoof-settings' }
    'APH-02' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antiphishing > <Richtlinie> > "Aktionen" > "Wenn die Nachricht durch Spoofintelligenz als Spoofing erkannt wird"'; Portal = 'https://security.microsoft.com/antiphishing'; Learn = 'https://learn.microsoft.com/defender-office-365/recommended-settings-for-eop-and-office365#anti-phishing-policy-settings-for-all-cloud-mailboxes' }
    'APH-03' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antiphishing > <Richtlinie> > "Phishingschwellenwert & Schutz" > "DMARC-Eintragsrichtlinie beruecksichtigen, wenn die Nachricht als Spoofing erkannt wird"'; Portal = 'https://security.microsoft.com/antiphishing'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about#spoof-protection-and-sender-dmarc-policies' }
    'APH-04' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antiphishing > <Richtlinie> > "Aktionen" > "... und die DMARC-Richtlinie p=quarantine lautet"'; Portal = 'https://security.microsoft.com/antiphishing'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about#spoof-protection-and-sender-dmarc-policies' }
    'APH-05' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antiphishing > <Richtlinie> > "Aktionen" > "... und die DMARC-Richtlinie p=reject lautet"'; Portal = 'https://security.microsoft.com/antiphishing'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about#spoof-protection-and-sender-dmarc-policies' }
    'APH-06' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antiphishing > <Richtlinie> > "Aktionen" > Sicherheitstipps & Indikatoren > "(?) fuer nicht authentifizierte Absender anzeigen" und "Tag via anzeigen"'; Portal = 'https://security.microsoft.com/antiphishing'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about#unauthenticated-sender-indicators' }
    'APH-07' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antiphishing > <Richtlinie> > "Aktionen" > Sicherheitstipps & Indikatoren > "Sicherheitstipp fuer ersten Kontakt anzeigen"'; Portal = 'https://security.microsoft.com/antiphishing'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about#first-contact-safety-tip' }
    'APH-08' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antiphishing > <Richtlinie> > "Phishingschwellenwert & Schutz" > "Schwellenwert fuer Phishing-E-Mails" (1-4, nur mit Defender-Lizenz sichtbar)'; Portal = 'https://security.microsoft.com/antiphishing'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about#phishing-email-thresholds-in-anti-phishing-policies-in-microsoft-defender-for-office-365' }
    'ASF-01' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antispam > <Richtlinie> > "Massen-E-Mail-Schwellenwert & Spameigenschaften" > Spameigenschaften > Gruppe "Erhoehen der Spambewertung"'; Portal = 'https://security.microsoft.com/antispam'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-spam-policies-asf-settings-about#increase-spam-score-settings' }
    'ASF-02' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antispam > <Richtlinie> > "Massen-E-Mail-Schwellenwert & Spameigenschaften" > Spameigenschaften > Gruppe "Als Spam markieren"'; Portal = 'https://security.microsoft.com/antispam'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-spam-policies-asf-settings-about#mark-as-spam-settings' }
    'ASF-03' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antispam > <Richtlinie> > "Massen-E-Mail-Schwellenwert & Spameigenschaften" > Spameigenschaften > "SPF-Eintrag: Hard Fail"'; Portal = 'https://security.microsoft.com/antispam'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-spam-policies-asf-settings-about#mark-as-spam-settings' }
    'ASF-04' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antispam > <Richtlinie> > "Massen-E-Mail-Schwellenwert & Spameigenschaften" > Spameigenschaften > "Absender-ID-Filterung: Hard Fail" und "Backscatter"'; Portal = 'https://security.microsoft.com/antispam'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-spam-policies-asf-settings-about#mark-as-spam-settings' }
    'ASF-05' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antispam > <Richtlinie> > "Massen-E-Mail-Schwellenwert & Spameigenschaften" > Spameigenschaften > "Testmodus" (gilt global fuer alle auf Test gesetzten ASF-Optionen)'; Portal = 'https://security.microsoft.com/antispam'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-spam-policies-asf-settings-about#enable-disable-or-test-asf-settings' }
    'ASI-01' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antispam > <Richtlinie> > "Massen-E-Mail-Schwellenwert & Spameigenschaften" > Schieberegler "Massen-E-Mail-Schwellenwert"'; Portal = 'https://security.microsoft.com/antispam'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-spam-protection-about#bulk-complaint-threshold-bcl-in-anti-spam-policies' }
    'ASI-02' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antispam > <Richtlinie> > "Aktionen" > Nachrichtenaktionen > Zeile "Spam"'; Portal = 'https://security.microsoft.com/antispam'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-spam-protection-about#actions-in-anti-spam-policies' }
    'ASI-03' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antispam > <Richtlinie> > "Aktionen" > Nachrichtenaktionen > Zeile "Spam mit hoher Sicherheit"'; Portal = 'https://security.microsoft.com/antispam'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-spam-protection-about#actions-in-anti-spam-policies' }
    'ASI-04' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antispam > <Richtlinie> > "Aktionen" > Nachrichtenaktionen > Zeile "Phishing"'; Portal = 'https://security.microsoft.com/antispam'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-spam-protection-about#actions-in-anti-spam-policies' }
    'ASI-05' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antispam > <Richtlinie> > "Aktionen" > Nachrichtenaktionen > Zeile "Phishing mit hoher Sicherheit"'; Portal = 'https://security.microsoft.com/antispam'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-spam-protection-about#actions-in-anti-spam-policies' }
    'ASI-06' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antispam > <Richtlinie> > "Aktionen" > Nachrichtenaktionen > Zeile "Massenkonforme Ebene (BCL) erreicht oder ueberschritten"'; Portal = 'https://security.microsoft.com/antispam'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-spam-protection-about#actions-in-anti-spam-policies' }
    'ASI-07' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antispam > <Richtlinie> > "Aktionen" > Abschnitt "Sicherheitstipps" > Haken "Sicherheitstipps aktivieren"'; Portal = 'https://security.microsoft.com/antispam'; Learn = 'https://learn.microsoft.com/defender-office-365/recommended-settings-for-eop-and-office365#anti-spam-policy-settings' }
    'ASI-08' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antispam > <Richtlinie> > "Aktionen" > "Automatische Bereinigung (ZAP) aktivieren" mit den zwei Unterhaken fuer Phishing und Spam'; Portal = 'https://security.microsoft.com/antispam'; Learn = 'https://learn.microsoft.com/defender-office-365/zero-hour-auto-purge#zero-hour-auto-purge-zap-for-spam' }
    'ASI-09' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antispam > <Richtlinie> > "Aktionen" > Feld "Spamnachrichten so viele Tage lang in Quarantaene aufbewahren" (1-30, Default 15)'; Portal = 'https://security.microsoft.com/antispam'; Learn = 'https://learn.microsoft.com/defender-office-365/quarantine-about#quarantine-retention' }
    'ASI-10' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antispam > <Richtlinie> > "Zulassen- und Sperrliste" > "Absender verwalten" / "Domaenen zulassen" (im Portal max. 30 Eintraege)'; Portal = 'https://security.microsoft.com/antispam'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-spam-protection-about#allow-and-block-lists-in-anti-spam-policies' }
    'ASI-11' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antispam > <Richtlinie> > "Massen-E-Mail-Schwellenwert & Spameigenschaften" > Spameigenschaften > "Enthaelt bestimmte Sprachen" / "Aus diesen Laendern" (gehoert NICHT zum ASF)'; Portal = 'https://security.microsoft.com/antispam'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-spam-protection-about#spam-properties-in-anti-spam-policies' }
    'ASI-12' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antispam > <Richtlinie> > "Aktionen" > "Organisationsinterne Nachrichten, fuer die Massnahmen ergriffen werden sollen"'; Portal = 'https://security.microsoft.com/antispam'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-spam-protection-about#actions-in-anti-spam-policies' }
    'ASO-01' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antispam > Liste auf "Ausgehende Antispamrichtlinie (Standard)" und eigene Richtlinien pruefen; neu ueber "+ Richtlinie erstellen > Ausgehend"'; Portal = 'https://security.microsoft.com/antispam'; Learn = 'https://learn.microsoft.com/defender-office-365/outbound-spam-policies-configure#use-the-microsoft-defender-portal-to-create-outbound-spam-policies' }
    'ASO-02' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antispam > <ausgehende Richtlinie> > "Schutzeinstellungen" > Nachrichtengrenzwerte (extern / intern / taeglich)'; Portal = 'https://security.microsoft.com/antispam'; Learn = 'https://learn.microsoft.com/defender-office-365/outbound-spam-policies-configure#use-the-microsoft-defender-portal-to-create-outbound-spam-policies' }
    'ASO-03' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antispam > <ausgehende Richtlinie> > "Schutzeinstellungen" > "Einschraenkung fuer Benutzer, die das Nachrichtenlimit erreichen"'; Portal = 'https://security.microsoft.com/antispam'; Learn = 'https://learn.microsoft.com/defender-office-365/recommended-settings-for-eop-and-office365#outbound-spam-policy-settings' }
    'ASO-04' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antispam > <ausgehende Richtlinie> > "Schutzeinstellungen" > Weiterleitungsregeln > "Automatische Weiterleitungsregeln"'; Portal = 'https://security.microsoft.com/antispam'; Learn = 'https://learn.microsoft.com/defender-office-365/outbound-spam-policies-external-email-forwarding' }
    'ASO-05' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antispam > "Ausgehende Antispamrichtlinie (Standard)" > "Schutzeinstellungen" > Benachrichtigungen (BCC-Kopie); gesperrte Absender unter Defender-Portal > Ueberpruefen > Eingeschraenkte Benutzer'; Portal = 'https://security.microsoft.com/restrictedusers'; Learn = 'https://learn.microsoft.com/defender-office-365/outbound-spam-policies-configure#use-the-microsoft-defender-portal-to-create-outbound-spam-policies' }
    'ASO-06' = @{ Nav = 'Nicht konfigurierbar (mandantenweites Limit). Auswertung: Exchange Admin Center > Berichte > Nachrichtenfluss > "Tenant Outbound External Recipients"'; Portal = 'https://admin.exchange.microsoft.com/#/reports/mailflowreportsmain'; Learn = 'https://learn.microsoft.com/defender-office-365/outbound-spam-sending-limits-troubleshoot#tenant-external-recipient-rate-limit' }
    'ASO-07' = @{ Nav = 'NICHT im Defender-Portal: Exchange Admin Center > Nachrichtenfluss > Remotedomaenen > "Default" > Abschnitt "Automatische Antworten" (dort sitzt die Weiterleitungsoption)'; Portal = 'https://admin.exchange.microsoft.com/#/remotedomains'; Learn = 'https://learn.microsoft.com/exchange/mail-flow-best-practices/remote-domains/remote-domains#reducing-or-increasing-information-flow-to-another-company' }
    'ASO-08' = @{ Nav = 'Keine Einstellung, mandantenweite Drosselung. Pruefung: Exchange Admin Center > Nachrichtenfluss > Akzeptierte Domaenen sowie Absenderdomaenen in Connectors und Anwendungen'; Portal = 'https://admin.exchange.microsoft.com/#/accepteddomains'; Learn = 'https://learn.microsoft.com/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits#sending-limits' }
    'AUTH-01' = @{ Nav = 'Oeffentliches DNS beim Domain-Hoster: TXT auf <domaene> mit v=spf1 include:spf.protection.outlook.com -all, je akzeptierter Domaene genau einer'; Portal = ''; Learn = 'https://learn.microsoft.com/defender-office-365/email-authentication-spf-configure#spf-txt-records-for-custom-domains-in-microsoft-365' }
    'AUTH-02' = @{ Nav = 'Oeffentliches DNS: TXT auf <domaene>, Endqualifier -all statt ~all'; Portal = ''; Learn = 'https://learn.microsoft.com/defender-office-365/email-authentication-spf-configure#syntax-for-spf-txt-records' }
    'AUTH-03' = @{ Nav = 'Oeffentliches DNS: fuer geparkte Domaenen TXT v=spf1 -all auf <domaene> plus TXT v=DMARC1; p=reject; auf _dmarc.<domaene>. Gilt auch fuer die onmicrosoft.com-Domaene'; Portal = ''; Learn = 'https://learn.microsoft.com/defender-office-365/email-authentication-spf-configure#scenario-parked-domains' }
    'AUTH-04' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > E-Mail-Authentifizierungseinstellungen > DKIM > Domaene waehlen, Umschalter aktivieren. Vorher die beiden CNAMEs selector1/2._domainkey.<domaene> ins oeffentliche DNS'; Portal = 'https://security.microsoft.com/authentication?viewid=DKIM'; Learn = 'https://learn.microsoft.com/defender-office-365/email-authentication-dkim-configure#use-the-defender-portal-to-enable-dkim-signing-of-outbound-messages-using-a-custom-domain' }
    'AUTH-05' = @{ Nav = 'Die Bitlaenge steuert nur PowerShell: Rotate-DkimSigningConfig -KeySize 2048. Im Portal rotiert man ueber E-Mail-Authentifizierungseinstellungen > DKIM > Domaenenzeile ANKLICKEN > Flyout > "DKIM-Schluessel rotieren"'; Portal = 'https://security.microsoft.com/authentication?viewid=DKIM'; Learn = 'https://learn.microsoft.com/defender-office-365/email-authentication-dkim-configure#use-exchange-online-powershell-to-rotate-the-dkim-keys-for-a-domain-and-change-the-bit-depth' }
    'AUTH-06' = @{ Nav = 'Oeffentliches DNS: TXT auf _dmarc.<domaene> mit v=DMARC1; p=...; rua=mailto:... Fuer die onmicrosoft.com-Domaene ueber Microsoft 365 Admin Center > Einstellungen > Domaenen'; Portal = 'https://admin.microsoft.com/Adminportal/Home#/Domains'; Learn = 'https://learn.microsoft.com/defender-office-365/email-authentication-dmarc-configure#set-up-dmarc-for-active-custom-domains-in-microsoft-365' }
    'AUTH-07' = @{ Nav = 'Oeffentliches DNS: rua=mailto: im TXT auf _dmarc.<domaene>, Ziel ein dediziertes Shared Mailbox, kein Benutzerpostfach'; Portal = ''; Learn = 'https://learn.microsoft.com/defender-office-365/email-authentication-dmarc-configure#best-practices-for-dmarc-reports' }
    'AUTH-08' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > E-Mail-Authentifizierungseinstellungen > ARC; alternativ Set-ArcConfig -ArcTrustedSealers'; Portal = 'https://security.microsoft.com/authentication'; Learn = 'https://learn.microsoft.com/defender-office-365/email-authentication-arc-configure#use-the-microsoft-defender-portal-to-add-trusted-arc-sealers' }
    'AUTH-09' = @{ Nav = 'Oeffentliches DNS: TXT auf _mta-sts.<domaene> PLUS extern gehostete Policy-Datei unter https://mta-sts.<domaene>/.well-known/mta-sts.txt. Exchange Online hostet die Datei nicht. Ausgehend ist MTA-STS immer aktiv'; Portal = ''; Learn = 'https://learn.microsoft.com/exchange/security-and-compliance/enhance-mail-flow-using-strict-transport-security#adopt-mta-sts-for-your-domain' }
    'AUTH-10' = @{ Nav = 'Ausgehend standardmaessig an. Eingehend per Exchange Online PowerShell: Enable-DnssecForVerifiedDomain, dann Enable-SmtpDaneInbound, danach den ausgegebenen MX-Wert beim Hoster setzen und die Delegation DNSSEC-signieren'; Portal = ''; Learn = 'https://learn.microsoft.com/exchange/security-and-compliance/how-dane-secures-email#inbound-smtp-dane-with-dnssec' }
    'AUTH-11' = @{ Nav = 'Kein Portal: SPF, DKIM und DMARC ausgerichtet auf die 5322.From-Domaene plus auffindbarer Abmeldemechanismus. Greift ab 5.000 Nachrichten/Tag an Microsoft-Consumer-Dienste, sonst NDR 550 5.7.515'; Portal = ''; Learn = 'https://learn.microsoft.com/defender-office-365/external-senders-policies-practices-guidelines' }
    'AUTH-12' = @{ Nav = 'Oeffentliches DNS: TXT auf default._bimi.<domaene>. Voraussetzung ist DMARC p=quarantine oder p=reject. Exchange Online wertet BIMI derzeit nicht aus'; Portal = ''; Learn = 'https://learn.microsoft.com/dynamics365/customer-insights/journeys/bimi-support' }
    'CF-01' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antispam > in der Liste die Zeile "Verbindungsfilterrichtlinie (Standard)" ANKLICKEN > Flyout > "Verbindungsfilterrichtlinie bearbeiten" > "Nachrichten aus den folgenden IP-Adressen immer zulassen"'; Portal = 'https://security.microsoft.com/antispam'; Learn = 'https://learn.microsoft.com/defender-office-365/connection-filter-policies-configure#use-the-microsoft-defender-portal-to-modify-the-default-connection-filter-policy' }
    'CF-02' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antispam > Zeile "Verbindungsfilterrichtlinie (Standard)" > "Verbindungsfilterrichtlinie bearbeiten" > "Nachrichten aus den folgenden IP-Adressen immer blockieren"'; Portal = 'https://security.microsoft.com/antispam'; Learn = 'https://learn.microsoft.com/defender-office-365/connection-filter-policies-configure#use-the-microsoft-defender-portal-to-modify-the-default-connection-filter-policy' }
    'CF-03' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antispam > Zeile "Verbindungsfilterrichtlinie (Standard)" > "Verbindungsfilterrichtlinie bearbeiten" > Haken "Sichere Liste aktivieren"'; Portal = 'https://security.microsoft.com/antispam'; Learn = 'https://learn.microsoft.com/defender-office-365/connection-filter-policies-configure#use-the-microsoft-defender-portal-to-modify-the-default-connection-filter-policy' }
    'EF-01' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Erweiterte Filterung (Enhanced filtering) > Eintrag des Inbound-Connectors > zu ueberspringende IPs bzw. "Letzte IP ueberspringen"'; Portal = 'https://security.microsoft.com/skiplisting'; Learn = 'https://learn.microsoft.com/exchange/mail-flow-best-practices/use-connectors-to-configure-mail-flow/enhanced-filtering-for-connectors#use-the-microsoft-defender-portal-to-configure-enhanced-filtering-for-connectors-on-an-inbound-connector' }
    'EF-02' = @{ Nav = 'Exchange Admin Center > Nachrichtenfluss > Regeln: SCL-(-1)-Regeln fuer Nachrichten ueber diesen Connector abschalten'; Portal = 'https://admin.exchange.microsoft.com/#/transportrules'; Learn = 'https://learn.microsoft.com/exchange/mail-flow-best-practices/use-connectors-to-configure-mail-flow/enhanced-filtering-for-connectors#what-do-you-need-to-know-before-you-begin' }
    'EXT-01' = @{ Nav = 'Kein Portal, nur Exchange Online PowerShell: Set-ExternalInOutlook -Enabled $true, Ausnahmen ueber -AllowList (max. 200). Wirkung erst nach 24-48 Stunden'; Portal = ''; Learn = 'https://learn.microsoft.com/powershell/module/exchangepowershell/set-externalinoutlook' }
    'EXT-02' = @{ Nav = 'Entscheidung zwischen nativem Tag (PowerShell), First-Contact-Tipp (Antiphishing-Richtlinie) und Banner per Transportregel. Nicht mehrere gleichzeitig'; Portal = 'https://security.microsoft.com/antiphishing'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about#first-contact-safety-tip' }
    'EXT-03' = @{ Nav = 'Exchange Admin Center > Nachrichtenfluss > Regeln: Regel mit "Absender ausserhalb der Organisation" und Aktion "Haftungsausschluss voranstellen", Fallback-Aktion und Ausnahmen setzen'; Portal = 'https://admin.exchange.microsoft.com/#/transportrules'; Learn = 'https://learn.microsoft.com/exchange/security-and-compliance/mail-flow-rules/disclaimers-signatures-footers-or-headers#use-the-eac-to-add-a-disclaimer-or-other-email-header-or-footer' }
    'GOV-01' = @{ Nav = 'Kein Microsoft-Portalpunkt: organisatorisch - Datenschutz-Folgenabschaetzung, Verfahrensverzeichnis und Mitbestimmung fuer Quarantaene-Einsicht und Explorer-Preview'; Portal = ''; Learn = 'https://learn.microsoft.com/compliance/regulatory/gdpr#data-protection-impact-assessment' }
    'HYB-01' = @{ Nav = 'Nur On-Premises: Exchange Management Shell (Get-ExchangeServer, ExSetup /Version) gegen die Supportability-Matrix. Exchange 2016 und 2019 sind seit 14.10.2025 out of support'; Portal = ''; Learn = 'https://learn.microsoft.com/exchange/plan-and-deploy/supportability-matrix#supported-versions-and-builds' }
    'HYB-02' = @{ Nav = 'On-Premises: Skript ConfigureExchangeHybridApplication.ps1 bzw. HCW. Kontrolle der Dienstprinzipal-Anmeldungen in Microsoft Entra ID > Ueberwachung > Anmeldeprotokolle'; Portal = 'https://entra.microsoft.com'; Learn = 'https://learn.microsoft.com/exchange/hybrid-deployment/deploy-dedicated-hybrid-app#configure-the-dedicated-exchange-hybrid-application' }
    'HYB-03' = @{ Nav = 'Nur On-Premises: "Default Frontend <Server>" nicht fuer anonymes Relay oeffnen, stattdessen dedizierter Receive-Connector mit RemoteIpRanges'; Portal = ''; Learn = 'https://learn.microsoft.com/exchange/mail-flow/connectors/allow-anonymous-relay#step-1-create-a-dedicated-receive-connector-for-anonymous-relay' }
    'HYB-04' = @{ Nav = 'Perimeter- und Host-Firewall: ausgehend SMTP 25/587 nur von den Transport-Servern, POP3 110/995 und IMAP4 143/993 nach aussen sperren'; Portal = ''; Learn = 'https://learn.microsoft.com/exchange/plan-and-deploy/deployment-ref/network-ports#network-ports-required-for-mail-flow' }
    'HYB-05' = @{ Nav = 'Exchange Admin Center > Einstellungen > E-Mail-Fluss > "SMTP AUTH-Protokoll fuer Ihre Organisation deaktivieren". Ausnahmen je Postfach ueber Set-CASMailbox'; Portal = 'https://admin.exchange.microsoft.com/#/settings'; Learn = 'https://learn.microsoft.com/exchange/clients-and-mobile-in-exchange-online/authenticated-client-smtp-submission#disable-smtp-auth-in-your-organization' }
    'HYB-06' = @{ Nav = 'Exchange Admin Center > Empfaenger > Postfaecher > <Postfach> > "E-Mail-App-Einstellungen verwalten"; organisationsweit ueber Set-CASMailbox bzw. Set-CASMailboxPlan fuer neue Postfaecher'; Portal = 'https://admin.exchange.microsoft.com/#/mailboxes'; Learn = 'https://learn.microsoft.com/exchange/recipients-in-exchange-online/manage-user-mailboxes/managing-email-apps-for-user-mailboxes#use-exchange-online-powershell-to-enable-or-disable-email-apps' }
    'HYB-07' = @{ Nav = 'Exchange Admin Center > Empfaenger > Postfaecher > <Postfach> > "Nachrichtengroessenbeschraenkung verwalten"; die Dienstgrenzwerte sind harte Limits'; Portal = 'https://admin.exchange.microsoft.com/#/mailboxes'; Learn = 'https://learn.microsoft.com/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits#message-limits' }
    'HYB-08' = @{ Nav = 'Exchange Admin Center > Nachrichtenfluss > Connectors > ausgehender Connector: "Immer eine TLS-gesicherte Verbindung verwenden" plus Zertifikatspruefung'; Portal = 'https://admin.exchange.microsoft.com/#/connectors'; Learn = 'https://learn.microsoft.com/exchange/mail-flow-best-practices/use-connectors-to-configure-mail-flow/set-up-connectors-to-secure-mail-sent-to-partner-organization#for-new-eac' }
    'HYB-09' = @{ Nav = 'Exchange Admin Center > Nachrichtenfluss > Connectors: alle ein- und ausgehenden Connectors inventarisieren (HCW vs. manuell, IP-Bereiche, Zertifikate)'; Portal = 'https://admin.exchange.microsoft.com/#/connectors'; Learn = 'https://learn.microsoft.com/exchange/mail-flow-best-practices/use-connectors-to-configure-mail-flow/use-connectors-to-configure-mail-flow#when-do-i-need-a-connector' }
    'HYB-10' = @{ Nav = 'Exchange Admin Center > Nachrichtenfluss > Akzeptierte Domaenen: Typ "Autorisierend" schaltet DBEB ein, "Internes Relay" schaltet es ab'; Portal = 'https://admin.exchange.microsoft.com/#/accepteddomains'; Learn = 'https://learn.microsoft.com/exchange/mail-flow-best-practices/use-directory-based-edge-blocking#configure-dbeb' }
    'IMP-01' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antiphishing > <Richtlinie> > "Phishingschwellenwert & Schutz" > Identitaetswechsel > "Zu schuetzende Domaenen aktivieren" > "Domaenen einschliessen, die ich besitze"'; Portal = 'https://security.microsoft.com/antiphishing'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about#domain-impersonation-protection' }
    'IMP-02' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antiphishing > <Richtlinie> > "Phishingschwellenwert & Schutz" > Identitaetswechsel > "Benutzerdefinierte Domaenen einschliessen" (max. 50)'; Portal = 'https://security.microsoft.com/antiphishing'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about#domain-impersonation-protection' }
    'IMP-03' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antiphishing > <Richtlinie> > "Phishingschwellenwert & Schutz" > Identitaetswechsel > "Benutzern das Schuetzen ermoeglichen" > "Geschuetzte Benutzer verwalten" (max. 350)'; Portal = 'https://security.microsoft.com/antiphishing'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about#user-impersonation-protection' }
    'IMP-04' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antiphishing > <Richtlinie> > "Aktionen" > die drei Zeilen zu Benutzer-, Domaenenidentitaetswechsel und Postfachintelligenz (nicht "Keine Aktion anwenden")'; Portal = 'https://security.microsoft.com/antiphishing'; Learn = 'https://learn.microsoft.com/defender-office-365/recommended-settings-for-eop-and-office365#impersonation-settings-in-anti-phishing-policies-in-microsoft-defender-for-office-365' }
    'IMP-05' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antiphishing > <Richtlinie> > "Phishingschwellenwert & Schutz" > Identitaetswechsel > "Postfachintelligenz aktivieren" und "Intelligenz fuer Identitaetswechselschutz aktivieren"'; Portal = 'https://security.microsoft.com/antiphishing'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about#mailbox-intelligence-impersonation-protection' }
    'IMP-06' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antiphishing > <Richtlinie> > "Aktionen" > Sicherheitstipps & Indikatoren > die drei Identitaetswechsel-Tipps'; Portal = 'https://security.microsoft.com/antiphishing'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about#impersonation-safety-tips' }
    'IMP-07' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antiphishing > <Richtlinie> > "Phishingschwellenwert & Schutz" > Identitaetswechsel > "Vertrauenswuerdige Absender und Domaenen hinzufuegen" (max. 1024)'; Portal = 'https://security.microsoft.com/antiphishing'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about#trusted-senders-and-domains' }
    'IMP-08' = @{ Nav = 'Defender-Portal > Einstellungen > E-Mail & Zusammenarbeit > Schutz von Prioritaetskonten (Priority account protection); die Konten selbst werden im Microsoft 365 Admin Center markiert'; Portal = 'https://security.microsoft.com/securitysettings/priorityAccountProtection'; Learn = 'https://learn.microsoft.com/defender-office-365/priority-accounts-turn-on-priority-account-protection#review-or-turn-on-priority-account-protection-in-the-microsoft-defender-portal' }
    'OPS-01' = @{ Nav = 'Defender-Portal > Berechtigungen: Rollengruppen "E-Mail & Zusammenarbeit" bzw. Defender XDR Unified RBAC; dazu die Entra-Rollen'; Portal = 'https://security.microsoft.com/emailandcollabpermissions'; Learn = 'https://learn.microsoft.com/defender-office-365/scc-permissions#role-groups-in-microsoft-defender-for-office-365-and-microsoft-purview' }
    'OPS-02' = @{ Nav = 'Kein Portal: Admin-Arbeitsplatz. Modul ExchangeOnlineManagement, Windows PowerShell 5.1 mit .NET 4.7.2+ oder PowerShell 7'; Portal = ''; Learn = 'https://learn.microsoft.com/powershell/exchange/exchange-online-powershell-v2' }
    'OPS-03' = @{ Nav = 'Baseline per PowerShell exportieren; ergaenzend Konfigurationsanalyse > Reiter "Configuration drift analysis and history" (setzt Unified Auditing voraus)'; Portal = 'https://security.microsoft.com/configurationAnalyzer'; Learn = 'https://learn.microsoft.com/defender-office-365/configuration-analyzer-for-security-policies#configuration-drift-analysis-and-history-tab-in-the-configuration-analyzer' }
    'OPS-04' = @{ Nav = 'Nachweis ueber Microsoft Purview > Audit (Unified Audit Log); Nachkontrolle als wiederkehrenden Termin setzen'; Portal = 'https://purview.microsoft.com'; Learn = 'https://learn.microsoft.com/purview/audit-log-enable-disable' }
    'P2-01' = @{ Nav = 'Defender-Portal > Untersuchungen (AIR, nur MDO Plan 2; setzt aktiviertes Audit-Logging voraus)'; Portal = 'https://security.microsoft.com/airinvestigation'; Learn = 'https://learn.microsoft.com/defender-office-365/air-about#the-overall-flow-of-air' }
    'PRE-01' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Voreingestellte Sicherheitsrichtlinien (Preset security policies)'; Portal = 'https://security.microsoft.com/presetSecurityPolicies'; Learn = 'https://learn.microsoft.com/defender-office-365/mdo-deployment-guide#determine-your-threat-policy-strategy' }
    'PRE-02' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Voreingestellte Sicherheitsrichtlinien > Integrierter Schutz (Built-in protection) > "Ausschluesse verwalten"'; Portal = 'https://security.microsoft.com/presetSecurityPolicies'; Learn = 'https://learn.microsoft.com/defender-office-365/preset-security-policies#use-the-microsoft-defender-portal-to-add-exclusions-to-the-built-in-protection-preset-security-policy' }
    'PRE-03' = @{ Nav = 'Kein eigener Portal-Screen. Praezedenz ablesen an der Spalte "Prioritaet" je Richtlinientyp (Antispam, Antiphishing, ...) plus der Preset-Seite'; Portal = 'https://security.microsoft.com/presetSecurityPolicies'; Learn = 'https://learn.microsoft.com/defender-office-365/preset-security-policies#order-of-precedence-for-preset-security-policies-and-other-threat-policies' }
    'PRE-04' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Konfigurationsanalyse (Configuration analyzer) > Reiter "Standardempfehlungen" / "Strenge Empfehlungen"'; Portal = 'https://security.microsoft.com/configurationAnalyzer'; Learn = 'https://learn.microsoft.com/defender-office-365/configuration-analyzer-for-security-policies#standard-recommendations-and-strict-recommendations-tabs-in-the-configuration-analyzer' }
    'PRE-05' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Antispam (Spalten Status / Prioritaet / Typ), analog Antiphishing, Antischadsoftware, Sichere Links, Sichere Anlagen'; Portal = 'https://security.microsoft.com/antispam'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-spam-policies-configure#use-the-microsoft-defender-portal-to-enable-or-disable-anti-spam-policies' }
    'QUA-01' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Quarantaenerichtlinie (definieren); zugewiesen wird sie je Verdict in Antispam > <Richtlinie> > "Aktionen", in Antiphishing > "Aktionen" sowie in Antischadsoftware und Sichere Anlagen'; Portal = 'https://security.microsoft.com/quarantinePolicies'; Learn = 'https://learn.microsoft.com/defender-office-365/quarantine-policies#assign-quarantine-policies-in-supported-policies-in-the-microsoft-defender-portal' }
    'QUA-02' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Quarantaenerichtlinie > "Benutzerdefinierte Richtlinie hinzufuegen" > "Zugriff auf Empfaengernachrichten" > "Spezifischen Zugriff festlegen (Erweitert)" > Quarantaenebenachrichtigung aktivieren'; Portal = 'https://security.microsoft.com/quarantinePolicies'; Learn = 'https://learn.microsoft.com/defender-office-365/quarantine-policies#step-1-create-quarantine-policies-in-the-microsoft-defender-portal' }
    'QUA-03' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Quarantaenerichtlinie > Zahnrad "Globale Einstellungen" > "Spambenachrichtigung fuer Endbenutzer senden alle" (4 Stunden / taeglich / woechentlich)'; Portal = 'https://security.microsoft.com/quarantinePolicies'; Learn = 'https://learn.microsoft.com/defender-office-365/quarantine-policies#customize-all-quarantine-notifications' }
    'QUA-04' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Quarantaenerichtlinie > Zahnrad "Globale Einstellungen" > Absenderadresse, Anzeigename, Betreff, Haftungsausschluss und Firmenlogo'; Portal = 'https://security.microsoft.com/quarantinePolicies'; Learn = 'https://learn.microsoft.com/defender-office-365/quarantine-policies#customize-all-quarantine-notifications' }
    'REP-01' = @{ Nav = 'Defender-Portal > Berichte > E-Mail & Zusammenarbeit; Zeitplan je Bericht ueber "Create schedule"'; Portal = 'https://security.microsoft.com/emailandcollabreport'; Learn = 'https://learn.microsoft.com/defender-office-365/reports-email-security' }
    'REP-02' = @{ Nav = 'Defender-Portal > Richtlinien & Regeln > Warnungsrichtlinie; Empfaenger je Richtlinie im Feld "Email recipients"'; Portal = 'https://security.microsoft.com/alertpolicies'; Learn = 'https://learn.microsoft.com/defender-xdr/alert-policies#alert-policy-settings' }
    'REP-03' = @{ Nav = 'Defender-Portal > Einstellungen > E-Mail & Zusammenarbeit > Benutzerdefinierte Meldungen (User reported settings); alternativ *-ReportSubmissionPolicy in PowerShell'; Portal = 'https://security.microsoft.com/securitysettings/userSubmission'; Learn = 'https://learn.microsoft.com/defender-office-365/submissions-user-reported-messages-custom-mailbox#use-the-microsoft-defender-portal-to-configure-user-reported-settings' }
    'REP-04' = @{ Nav = 'Kein Portalpunkt, sondern Skripte: auf Get-MessageTraceV2 / Get-MessageTraceDetailV2 umstellen. Interaktiv fuehrt der Defender-Portal-Eintrag nur ins EAC'; Portal = 'https://admin.exchange.microsoft.com/#/messagetrace'; Learn = 'https://learn.microsoft.com/powershell/module/exchangepowershell/get-messagetracev2' }
    'SA-01' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Sichere Anlagen > <Richtlinie> > "Einstellungen" > "Safe Attachments-Antwort bei unbekannter Schadsoftware" = Blockieren'; Portal = 'https://security.microsoft.com/safeattachmentv2'; Learn = 'https://learn.microsoft.com/defender-office-365/safe-attachments-about#safe-attachments-policy-settings' }
    'SA-02' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Sichere Anlagen > <Richtlinie> > "Einstellungen" > "Umleiten von Nachrichten mit erkannten Anlagen" (wirkt laut Doku nur bei Aktion "Ueberwachen")'; Portal = 'https://security.microsoft.com/safeattachmentv2'; Learn = 'https://learn.microsoft.com/defender-office-365/safe-attachments-about#safe-attachments-policy-settings' }
    'SA-03' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Sichere Anlagen > Zahnrad "Globale Einstellungen" > "Defender for Office 365 fuer SharePoint, OneDrive und Microsoft Teams aktivieren". Download-Sperre zusaetzlich nur per SharePoint-PowerShell'; Portal = 'https://security.microsoft.com/safeattachmentv2'; Learn = 'https://learn.microsoft.com/defender-office-365/safe-attachments-for-spo-odfb-teams-configure#step-1-use-the-microsoft-defender-portal-to-turn-on-safe-attachments-for-sharepoint-onedrive-and-microsoft-teams' }
    'SA-04' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Sichere Anlagen > Zahnrad "Globale Einstellungen" > "Safe Documents fuer Office-Clients aktivieren" plus Durchklicken verbieten'; Portal = 'https://security.microsoft.com/safeattachmentv2'; Learn = 'https://learn.microsoft.com/defender-office-365/safe-documents-in-e5-plus-security-about#use-the-microsoft-defender-portal-to-configure-safe-documents' }
    'SA-05' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Sichere Anlagen > <Richtlinie> > "Einstellungen" > "Safe Attachments-Antwort" = "Dynamische Uebermittlung (Vorschau von Nachrichten)"'; Portal = 'https://security.microsoft.com/safeattachmentv2'; Learn = 'https://learn.microsoft.com/defender-office-365/safe-attachments-about#dynamic-delivery-in-safe-attachments-policies' }
    'SL-01' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Sichere Links > <Richtlinie> > "URL- & Klickschutzeinstellungen" > Abschnitte "E-Mail", "Teams" und "Office 365-Apps" je auf Ein'; Portal = 'https://security.microsoft.com/safelinksv2'; Learn = 'https://learn.microsoft.com/defender-office-365/safe-links-about#safe-links-settings-for-email-messages' }
    'SL-02' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Sichere Links > <Richtlinie> > "URL- & Klickschutzeinstellungen" > E-Mail > "Sichere Links auf E-Mail-Nachrichten anwenden, die innerhalb der Organisation gesendet werden"'; Portal = 'https://security.microsoft.com/safelinksv2'; Learn = 'https://learn.microsoft.com/defender-office-365/safe-links-about#safe-links-settings-for-email-messages' }
    'SL-03' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Sichere Links > <Richtlinie> > "URL- & Klickschutzeinstellungen" > "URL-Ueberpruefung in Echtzeit ..." mit der Unteroption "Warten, bis die URL-Ueberpruefung abgeschlossen ist"'; Portal = 'https://security.microsoft.com/safelinksv2'; Learn = 'https://learn.microsoft.com/defender-office-365/safe-links-about#safe-links-settings-for-email-messages' }
    'SL-04' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Sichere Links > <Richtlinie> > "URL- & Klickschutzeinstellungen" > Klickschutzeinstellungen > "Benutzern das Durchklicken zur urspruenglichen URL erlauben" ausschalten'; Portal = 'https://security.microsoft.com/safelinksv2'; Learn = 'https://learn.microsoft.com/defender-office-365/safe-links-about#click-protection-settings-in-safe-links-policies' }
    'SL-05' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Sichere Links > <Richtlinie> > "URL- & Klickschutzeinstellungen" > E-Mail > "URLs nicht umschreiben, Ueberpruefungen nur ueber die SafeLinks-API" muss AUS sein'; Portal = 'https://security.microsoft.com/safelinksv2'; Learn = 'https://learn.microsoft.com/defender-office-365/safe-links-about#safe-links-settings-for-email-messages' }
    'SL-06' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Sichere Links > <Richtlinie> > "URL- & Klickschutzeinstellungen" > "Die folgenden URLs in E-Mails nicht umschreiben" > "Nicht umzuschreibende URLs verwalten" (pro Richtlinie, nicht global)'; Portal = 'https://security.microsoft.com/safelinksv2'; Learn = 'https://learn.microsoft.com/defender-office-365/safe-links-about#entry-syntax-for-the-do-not-rewrite-the-following-urls-list' }
    'SL-07' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Sichere Links > <Richtlinie> > "URL- & Klickschutzeinstellungen" > "Organisationsbranding anzeigen" (Logo aus dem M365-Organisationsdesign) sowie > "Benachrichtigung" > eigener Text (max. 200 Zeichen)'; Portal = 'https://security.microsoft.com/safelinksv2'; Learn = 'https://learn.microsoft.com/defender-office-365/safe-links-about#click-protection-settings-in-safe-links-policies' }
    'TABL-01' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Mandanten-Zulassungs-/Sperrlisten. Zulassungen entstehen bevorzugt ueber Uebermittlungen (Submissions), nicht von Hand'; Portal = 'https://security.microsoft.com/tenantAllowBlockList'; Learn = 'https://learn.microsoft.com/defender-office-365/tenant-allow-block-list-about#allow-entries-in-the-tenant-allowblock-list' }
    'TABL-02' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Mandanten-Zulassungs-/Sperrlisten > Reiter "Domaenen und E-Mail-Adressen" > Spalten "Aktion" und "Laeuft ab am"'; Portal = 'https://security.microsoft.com/tenantAllowBlockList'; Learn = 'https://learn.microsoft.com/defender-office-365/tenant-allow-block-list-email-spoof-configure#domains-and-email-addresses-in-the-tenant-allowblock-list' }
    'TABL-03' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Mandanten-Zulassungs-/Sperrlisten > Reiter "Gefaelschte Absender"; ergaenzend die Spoofintelligenz-Erkenntnisse'; Portal = 'https://security.microsoft.com/tenantAllowBlockList?viewid=SpoofItem'; Learn = 'https://learn.microsoft.com/defender-office-365/tenant-allow-block-list-email-spoof-configure#spoofed-senders-in-the-tenant-allowblock-list' }
    'TABL-04' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Mandanten-Zulassungs-/Sperrlisten > Reiter "URLs"'; Portal = 'https://security.microsoft.com/tenantAllowBlockList'; Learn = 'https://learn.microsoft.com/defender-office-365/tenant-allow-block-list-urls-configure#url-syntax-for-the-tenant-allowblock-list' }
    'TABL-05' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Mandanten-Zulassungs-/Sperrlisten > Reiter "Dateien" (nur SHA256, Zulassungen nur ueber Uebermittlungen)'; Portal = 'https://security.microsoft.com/tenantAllowBlockList'; Learn = 'https://learn.microsoft.com/defender-office-365/tenant-allow-block-list-files-configure#create-block-entries-for-files' }
    'TABL-06' = @{ Nav = 'Defender-Portal > Bedrohungsrichtlinien > Mandanten-Zulassungs-/Sperrlisten > Reiter "IP-Adressen" (ausschliesslich IPv6)'; Portal = 'https://security.microsoft.com/tenantAllowBlockList'; Learn = 'https://learn.microsoft.com/defender-office-365/tenant-allow-block-list-ip-addresses-configure#create-block-entries-for-ipv6-addresses' }
    'TEAMS-01' = @{ Nav = 'Defender-Portal > Einstellungen > E-Mail & Zusammenarbeit > Microsoft Teams-Schutz (NICHT mehr unter Bedrohungsrichtlinien) > ZAP fuer Teams, Quarantaenerichtlinie, Ausnahmen'; Portal = 'https://security.microsoft.com/securitysettings/teamsProtectionPolicy'; Learn = 'https://learn.microsoft.com/defender-office-365/mdo-support-teams-about#configure-zap-for-teams-protection-in-defender-for-office-365' }
    'VER-01' = @{ Nav = 'Defender-Portal > Ueberpruefen > Quarantaene, Filter Richtlinientyp = Antischadsoftware-Richtlinie'; Portal = 'https://security.microsoft.com/quarantine'; Learn = 'https://learn.microsoft.com/defender-office-365/anti-malware-protection-about#common-attachments-filter-in-anti-malware-policies' }
    'VER-02' = @{ Nav = 'Kein Portal: Outlook > Datei > Eigenschaften > Internetkopfzeilen (neues Outlook/OWA: Nachricht > ... > Details anzeigen). Ausgewertet werden X-Forefront-Antispam-Report, X-Microsoft-Antispam und Authentication-Results'; Portal = ''; Learn = 'https://learn.microsoft.com/defender-office-365/message-headers-eop-mdo#x-forefront-antispam-report-message-header-fields' }
    'VER-03' = @{ Nav = 'Defender-Portal > E-Mail & Zusammenarbeit > Explorer (P2) bzw. Echtzeiterkennungen (P1), Ansichten Phish und URL-Klicks'; Portal = 'https://security.microsoft.com/threatexplorerv3'; Learn = 'https://learn.microsoft.com/defender-office-365/threat-explorer-real-time-detections-about' }
}

$script:LinkAliases = @{
    'AUTH-DNS'         = 'AUTH-01'
    'INV-AMW'          = 'AMW-01'
    'INV-APH'          = 'APH-01'
    'INV-ASI'          = 'PRE-05'
    'INV-ASO'          = 'ASO-01'
    'INV-SA'           = 'SA-01'
    'INV-SL'           = 'SL-01'
    'TABL-FileHash'    = 'TABL-05'
    'TABL-IP'          = 'TABL-06'
    'TABL-Sender'      = 'TABL-02'
    'TABL-Url'         = 'TABL-04'
}

$script:Prefixes = @(
    @{ Key = 'PRE'; Area = 'Preset Security Policies und Governance'; Desc = 'Voreingestellte Sicherheitsrichtlinien, Praezedenz, Configuration Analyzer' }
    @{ Key = 'ASI'; Area = 'Anti-Spam inbound'; Desc = 'Anti-Spam-Richtlinie, eingehend' }
    @{ Key = 'ASF'; Area = 'Advanced Spam Filter'; Desc = 'Die ASF-Schalter innerhalb der Anti-Spam-Richtlinie' }
    @{ Key = 'CF'; Area = 'Connection Filter'; Desc = 'Verbindungsfilterrichtlinie - liegt als eigene Zeile in der Liste der Anti-Spam-Richtlinien' }
    @{ Key = 'ASO'; Area = 'Anti-Spam outbound'; Desc = 'Ausgehende Anti-Spam-Richtlinie plus Weiterleitungswege' }
    @{ Key = 'APH'; Area = 'Anti-Phishing (EOP-Teil)'; Desc = 'Spoof-Intelligence, DMARC-Behandlung, Sicherheitstipps - ohne Defender-Lizenz verfuegbar' }
    @{ Key = 'IMP'; Area = 'Impersonation (Defender-Teil)'; Desc = 'Identitaetswechselschutz - sitzt in derselben Anti-Phishing-Richtlinie, ist aber lizenzpflichtig' }
    @{ Key = 'TABL'; Area = 'Tenant Allow/Block List'; Desc = 'Mandanten-Zulassungs-/Sperrliste' }
    @{ Key = 'AMW'; Area = 'Anti-Malware'; Desc = 'Anti-Malware-Richtlinie und Anlagenfilter' }
    @{ Key = 'SA'; Area = 'Safe Attachments'; Desc = 'Sichere Anlagen (Defender)' }
    @{ Key = 'SL'; Area = 'Safe Links'; Desc = 'Sichere Links (Defender)' }
    @{ Key = 'QUA'; Area = 'Quarantaene-Richtlinien'; Desc = 'Zugriffsrechte und Benachrichtigungen fuer die Quarantaene' }
    @{ Key = 'ADV'; Area = 'Advanced Delivery'; Desc = 'Erweiterte Zustellung: SecOps-Postfach und Phishing-Simulation' }
    @{ Key = 'AUTH'; Area = 'E-Mail-Authentifizierung'; Desc = 'SPF, DKIM, DMARC, ARC, MTA-STS, DANE - ueberwiegend im oeffentlichen DNS' }
    @{ Key = 'EF'; Area = 'Enhanced Filtering'; Desc = 'Erweiterte Filterung am Inbound-Connector (Skip Listing)' }
    @{ Key = 'EXT'; Area = 'Externe Kennzeichnung'; Desc = 'External-Tag, First-Contact-Tipp, Banner per Transportregel' }
    @{ Key = 'REP'; Area = 'Reporting und Alerting'; Desc = 'Berichte, Warnungsrichtlinien, Meldeweg fuer Nutzer' }
    @{ Key = 'HYB'; Area = 'Hybrid und On-Premises'; Desc = 'Exchange Server, Connectors, akzeptierte Domaenen, Legacy-Protokolle' }
    @{ Key = 'OPS'; Area = 'Betrieb und Voraussetzungen'; Desc = 'Rollen, PowerShell-Arbeitsplatz, Baseline, Nachkontrolle' }
    @{ Key = 'TEAMS'; Area = 'Microsoft Teams Protection'; Desc = 'ZAP fuer Teams-Nachrichten' }
    @{ Key = 'VER'; Area = 'Wirksamkeitsnachweis'; Desc = 'Belegen, dass der Schutz greift - Quarantaene, Header, Explorer' }
    @{ Key = 'GOV'; Area = 'Governance und Datenschutz'; Desc = 'DSGVO und Mitbestimmung rund um Quarantaene-Einsicht' }
    @{ Key = 'P2'; Area = 'Defender Plan 2'; Desc = 'Nur mit MDO P2: automatisierte Untersuchung (AIR)' }
)

# Deep-Links, die Microsoft nicht selbst dokumentiert (aus der Praxis uebernommen).
$script:UndocumentedLinks = @('ASO-07', 'ASO-08', 'HYB-06', 'HYB-07', 'HYB-10')

function Get-LinkInfo {
    <# Fundort zu einer Pruefpunkt-ID. Auch fuer abgeleitete IDs wie "TEAMS-01b" oder "AUTH-DNS". #>
    param([string]$Id)
    if ([string]::IsNullOrWhiteSpace($Id)) { return $null }
    if ($script:Links.ContainsKey($Id)) { return $script:Links[$Id] }
    # Sammel- und Inventar-IDs auf ihren Pruefpunkt abbilden
    if ($script:LinkAliases.ContainsKey($Id)) {
        return $script:Links[$script:LinkAliases[$Id]]
    }
    # Suffixe wie 'b' abschneiden: TEAMS-01b -> TEAMS-01
    $m = [regex]::Match($Id, '^([A-Z0-9]+-[0-9]+)')
    if ($m.Success -and $script:Links.ContainsKey($m.Groups[1].Value)) {
        return $script:Links[$m.Groups[1].Value]
    }
    return $null
}

function ConvertTo-HtmlText {
    <# Ohne System.Web - die Assembly ist nicht garantiert geladen. #>
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    $t = $Text -replace '&', '&amp;'
    $t = $t -replace '<', '&lt;'
    $t = $t -replace '>', '&gt;'
    $t = $t -replace '"', '&quot;'
    return $t
}


function Add-Finding {
    <#
        Bewertung:
          OK                     Ist entspricht Soll
          Verbesserung empfohlen Abweichung ohne akutes Risiko
          Kritisch               Schutzwirkung fehlt oder ist wirkungslos
          n.a.                   Feature nicht verfuegbar / nicht lizenziert
          Info                   reine Bestandsaufnahme, keine Bewertung
    #>
    param(
        [Parameter(Mandatory=$true)][string]$Id,
        [Parameter(Mandatory=$true)][string]$Category,
        [Parameter(Mandatory=$true)][string]$Check,
        [string]$Scope = '',
        $Actual,
        [string]$Expected = '',
        [ValidateSet('OK','Verbesserung empfohlen','Kritisch','n.a.','Info')]
        [string]$Rating = 'Info',
        [string]$Note = ''
    )

    # Policies ohne wirksame Regel werden nicht bemaengelt, nur berichtet.
    if (-not $script:CurrentPolicyActive -and $Rating -ne 'n.a.') {
        $Rating = 'Info'
        if (-not [string]::IsNullOrWhiteSpace($script:CurrentPolicyStatus)) {
            $Note = '[' + $script:CurrentPolicyStatus + '] ' + $Note
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($script:CurrentPolicyStatus)) {
        $Scope = $Scope + '  (' + $script:CurrentPolicyStatus + ')'
    }

    $obj = [pscustomobject]@{
        ID        = $Id
        Kategorie = $Category
        Pruefpunkt= $Check
        Objekt    = $Scope
        IstWert   = (ConvertTo-DisplayValue $Actual)
        SollWert  = $Expected
        Bewertung = $Rating
        Hinweis   = $Note
    }
    [void]$script:Findings.Add($obj)

    $color = 'Gray'
    switch ($Rating) {
        'OK'                     { $color = 'Green' }
        'Verbesserung empfohlen' { $color = 'Yellow' }
        'Kritisch'               { $color = 'Red' }
        'n.a.'                   { $color = 'DarkGray' }
    }
    $line = '{0,-9} {1,-24} {2}' -f $Id, $Rating, $Check
    Write-Host $line -ForegroundColor $color
}

function Test-Value {
    <#
        Vergleicht einen Ist-Wert gegen einen oder mehrere zulaessige Soll-Werte
        und legt daraus die Bewertung fest.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$Id,
        [Parameter(Mandatory=$true)][string]$Category,
        [Parameter(Mandatory=$true)][string]$Check,
        [string]$Scope = '',
        $Actual,
        [Parameter(Mandatory=$true)][string[]]$Accept,
        [string]$ExpectedText = '',
        [ValidateSet('Verbesserung empfohlen','Kritisch')]
        [string]$FailRating = 'Verbesserung empfohlen',
        [string]$Note = ''
    )

    if ([string]::IsNullOrEmpty($ExpectedText)) { $ExpectedText = ($Accept -join ' oder ') }
    $actualText = ConvertTo-DisplayValue $Actual

    $match = $false
    foreach ($a in $Accept) {
        if ($actualText -eq $a) { $match = $true; break }
    }

    $rating = $FailRating
    if ($match) { $rating = 'OK' }

    Add-Finding -Id $Id -Category $Category -Check $Check -Scope $Scope `
                -Actual $actualText -Expected $ExpectedText -Rating $rating -Note $Note
}

function Get-SafeProperty {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p) { return $null }
    return $p.Value
}

function Invoke-SafeCommand {
    <#
        Fuehrt ein lesendes Cmdlet aus und faengt Fehler ab, damit fehlende Lizenzen
        oder Berechtigungen den gesamten Lauf nicht abbrechen.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$Name,
        [Parameter(Mandatory=$true)][scriptblock]$Script
    )
    try {
        return & $Script
    }
    catch {
        Write-Info ("{0}: nicht verfuegbar ({1})" -f $Name, $_.Exception.Message)
        return $null
    }
}

function Test-IsPresetPolicy {
    <# Erkennt die von Microsoft verwalteten Preset-Policies an ihrem Namen. #>
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    if ($Name -match '^(Standard|Strict) Preset Security Policy') { return $true }
    if ($Name -match 'Built-In Protection Policy')                { return $true }
    if ($Name -match 'Evaluation Policy')                         { return $true }
    return $false
}

function Get-RuleForPolicy {
    <#
        Sucht die Regel, die einer Policy zugeordnet ist. Der Name des
        zustaendigen Policy-Feldes unterscheidet sich je Policy-Typ.
    #>
    param($Rules, [string]$PolicyName, [string[]]$PolicyProperties)
    if ($null -eq $Rules -or [string]::IsNullOrWhiteSpace($PolicyName)) { return $null }
    foreach ($r in @($Rules)) {
        foreach ($prop in $PolicyProperties) {
            $v = Get-SafeProperty $r $prop
            if ($null -ne $v -and ([string]$v) -eq $PolicyName) { return $r }
        }
    }
    return $null
}

function Set-PolicyContext {
    <#
        Ermittelt, ob eine Policy tatsaechlich auf Empfaenger angewendet wird,
        und setzt den Kontext fuer die nachfolgenden Pruefungen.

        Rueckgabe: Statustext. Nebenwirkung: setzt $script:CurrentPolicyActive
        und $script:CurrentPolicyStatus.
    #>
    param(
        $Policy,
        $Rules,
        [string[]]$PolicyProperties,
        [switch]$IsDefaultPolicy
    )

    $name = [string](Get-SafeProperty $Policy 'Name')
    if ([string]::IsNullOrWhiteSpace($name)) { $name = [string](Get-SafeProperty $Policy 'Identity') }

    # Von Microsoft verwaltete Preset-Policies: nicht editierbar, per Definition konform
    if (Test-IsPresetPolicy $name) {
        $script:CurrentPolicyActive = $false
        $script:CurrentPolicyStatus = 'Preset - nicht editierbar'
        return $script:CurrentPolicyStatus
    }

    # Default-Policy: immer wirksam, hat keine Regel
    $isDefault = $IsDefaultPolicy.IsPresent
    if (-not $isDefault) {
        $d = Get-SafeProperty $Policy 'IsDefault'
        if ($d -eq $true) { $isDefault = $true }
        if ($name -eq 'Default' -or $name -match '\(Default\)$') { $isDefault = $true }
    }
    if ($isDefault) {
        $script:CurrentPolicyActive = $true
        $script:CurrentPolicyStatus = 'Default-Policy'
        return $script:CurrentPolicyStatus
    }

    $rule = Get-RuleForPolicy -Rules $Rules -PolicyName $name -PolicyProperties $PolicyProperties

    if ($null -eq $rule) {
        $script:CurrentPolicyActive = $false
        $script:CurrentPolicyStatus = 'ohne Regel - wirkt auf niemanden'
        return $script:CurrentPolicyStatus
    }

    # Regelzustand: je nach Cmdlet heisst das Feld State oder Enabled
    $state   = [string](Get-SafeProperty $rule 'State')
    $enabled = Get-SafeProperty $rule 'Enabled'
    $isOff = $false
    if ($state -eq 'Disabled') { $isOff = $true }
    if ($enabled -eq $false)   { $isOff = $true }

    if ($isOff) {
        $script:CurrentPolicyActive = $false
        $script:CurrentPolicyStatus = 'Regel deaktiviert - wirkt auf niemanden'
        return $script:CurrentPolicyStatus
    }

    $prio = Get-SafeProperty $rule 'Priority'
    $script:CurrentPolicyActive = $true
    $script:CurrentPolicyStatus = 'aktiv - Prio ' + $prio
    return $script:CurrentPolicyStatus
}

function Reset-PolicyContext {
    $script:CurrentPolicyActive = $true
    $script:CurrentPolicyStatus = ''
}

function Get-RuleScopeText {
    <# Fasst den Empfaengerbereich einer Regel lesbar zusammen. #>
    param($Rule)
    if ($null -eq $Rule) { return 'keine Regel' }
    $parts = @()
    foreach ($f in @('SentTo','SentToMemberOf','RecipientDomainIs')) {
        $v = @(Get-SafeProperty $Rule $f | Where-Object { $null -ne $_ -and ([string]$_).Trim() -ne '' })
        if ($v.Count -gt 0) { $parts += ($f + '=' + $v.Count) }
    }
    foreach ($f in @('ExceptIfSentTo','ExceptIfSentToMemberOf','ExceptIfRecipientDomainIs')) {
        $v = @(Get-SafeProperty $Rule $f | Where-Object { $null -ne $_ -and ([string]$_).Trim() -ne '' })
        if ($v.Count -gt 0) { $parts += ('Ausnahme ' + $f.Replace('ExceptIf','') + '=' + $v.Count) }
    }
    if ($parts.Count -eq 0) { return 'alle Empfaenger' }
    return ($parts -join ' + ')
}

function Get-ExoConnection {
    <#
        Liefert eine bestehende, nutzbare Verbindung zurueck - oder $null.

        Grundlage ist Get-ConnectionInformation (Modul 3.0.0 oder neuer). Microsoft
        woertlich: "This cmdlet is required because the Get-PSSession cmdlet in Windows
        PowerShell doesn't return information for REST API connections." Get-PSSession
        ist fuer REST-Verbindungen also wertlos und wird hier nicht verwendet.

        Unterschieden wird ueber IsEopSession: $false = Exchange Online,
        $true = Security & Compliance.

        Fuer State und TokenStatus dokumentiert Microsoft keine Werteliste (nur
        "For example, Connected." bzw. "For example, Active."). Deshalb wird defensiv
        geprueft: ein Objekt gilt als nutzbar, solange es sich nicht nachweislich als
        unbrauchbar zu erkennen gibt.
    #>
    param([switch]$Compliance)

    if ($null -eq (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
        return $null   # Modul aelter als 3.0.0 - Erkennung nicht moeglich
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

        # Abgelaufenes Token aussortieren, falls der Zeitstempel lesbar ist
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
    $org  = [string](Get-SafeProperty $Connection 'Organization')
    if ([string]::IsNullOrWhiteSpace($org)) { $org = [string](Get-SafeProperty $Connection 'DelegatedOrganization') }
    $upn  = [string](Get-SafeProperty $Connection 'UserPrincipalName')
    $tid  = [string](Get-SafeProperty $Connection 'TenantID')
    Write-Info ($Label + ': ' + $org)
    if (-not [string]::IsNullOrWhiteSpace($upn)) { Write-Info ('  Angemeldet als : ' + $upn) }
    if (-not [string]::IsNullOrWhiteSpace($tid)) { Write-Info ('  Tenant-ID      : ' + $tid) }
}

function Confirm-ExpectedAccount {
    <#
        Warnt, wenn eine bestehende Verbindung unter einem anderen Konto laeuft als
        ueber -UserPrincipalName angegeben. Beim Arbeiten mit mehreren Kundentenants
        ist das der haeufigste und teuerste Irrtum.
    #>
    param($Connection, [string]$Expected)
    if ($null -eq $Connection -or [string]::IsNullOrWhiteSpace($Expected)) { return $true }
    $upn = [string](Get-SafeProperty $Connection 'UserPrincipalName')
    if ([string]::IsNullOrWhiteSpace($upn)) { return $true }
    if ($upn -eq $Expected) { return $true }
    Write-Host ''
    Write-Warning ('Die bestehende Verbindung laeuft unter ' + $upn + ', angegeben war ' + $Expected + '.')
    Write-Warning 'Pruefen Sie, ob Sie im richtigen Tenant arbeiten. Mit -ForceNewConnection wird neu angemeldet.'
    Write-Host ''
    return $false
}

$script:MinimumExoModuleVersion = '3.6.0'

function Initialize-ExoModule {
    <# Modulpruefung samt Mindestversion, TLS 1.2 fuer PowerShell 5.1 und Import. #>
    $min = [Version]$script:MinimumExoModuleVersion
    $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement |
              Sort-Object Version -Descending | Select-Object -First 1
    if ($null -eq $module) {
        throw ('Das Modul ExchangeOnlineManagement wurde nicht gefunden. Installation: ' +
               'Install-Module ExchangeOnlineManagement -Scope CurrentUser -MinimumVersion ' + $script:MinimumExoModuleVersion)
    }
    if ($module.Version -lt $min) {
        throw ('ExchangeOnlineManagement ' + $module.Version.ToString() + ' ist zu alt, mindestens ' +
               $script:MinimumExoModuleVersion + ' wird gebraucht. Aktualisieren mit: Update-Module ExchangeOnlineManagement')
    }
    Write-Info ('Modulversion   : ' + $module.Version.ToString() + ' (mindestens ' + $script:MinimumExoModuleVersion + ')')
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    } catch { }
    Import-Module ExchangeOnlineManagement -MinimumVersion $script:MinimumExoModuleVersion -ErrorAction Stop
    return $module
}

# ===================================================================================
#  Verbindung
# ===================================================================================

Write-Host ''
Write-Host '=============================================================' -ForegroundColor White
Write-Host ' EOP / Microsoft Defender for Office 365 - Audit (nur lesend)' -ForegroundColor White
Write-Host '=============================================================' -ForegroundColor White
Write-Host (' Kunde  : ' + $CustomerName)
Write-Host (' Datum  : ' + (Get-Date -Format 'dd.MM.yyyy HH:mm'))
Write-Host (' PS     : ' + $PSVersionTable.PSVersion.ToString())

Write-Step 'Verbindung zu Exchange Online'

$script:ConnectionWasReused = $false

if ($ForceNewConnection -and -not $SkipConnect) {
    Write-Info 'ForceNewConnection: bestehende Verbindungen werden getrennt.'
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { }
}

$existing = Get-ExoConnection

if ($null -ne $existing) {
    Write-Info 'Bestehende Verbindung gefunden - sie wird weiterverwendet.'
    Write-ConnectionSummary -Connection $existing -Label 'Organisation  '
    $null = Confirm-ExpectedAccount -Connection $existing -Expected $UserPrincipalName
    $script:ConnectionWasReused = $true
}
elseif ($SkipConnect) {
    Write-Warning 'Keine bestehende Verbindung gefunden, und -SkipConnect verhindert den Verbindungsaufbau.'
    Write-Warning 'Die Pruefungen werden weitgehend leer ausfallen. Skript ohne -SkipConnect erneut ausfuehren.'
}
else {
    $null = Initialize-ExoModule
    Write-Info 'Keine bestehende Verbindung - es wird eine neue aufgebaut.'
    if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    } else {
        Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowBanner:$false -ErrorAction Stop
    }
    Write-ConnectionSummary -Connection (Get-ExoConnection) -Label 'Verbunden mit '
}

if ($IncludeProtectionAlerts) {
    $existingIpps = Get-ExoConnection -Compliance
    if ($null -ne $existingIpps) {
        Write-Info 'Bestehende Security-&-Compliance-Verbindung gefunden - sie wird weiterverwendet.'
    }
    elseif ($SkipConnect) {
        Write-Warning 'Keine Security-&-Compliance-Verbindung vorhanden - Alert Policies werden uebersprungen.'
        $IncludeProtectionAlerts = $false
    }
    else {
        Write-Info 'Zweite Verbindung fuer Alert Policies (Security & Compliance) ...'
        try {
            if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
                Connect-IPPSSession -ErrorAction Stop
            } else {
                Connect-IPPSSession -UserPrincipalName $UserPrincipalName -ErrorAction Stop
            }
        }
        catch {
            Write-Warning ('Connect-IPPSSession fehlgeschlagen: ' + $_.Exception.Message)
            $IncludeProtectionAlerts = $false
        }
    }
}

# ===================================================================================
#  Tenant-Basisdaten
# ===================================================================================

Write-Step 'Tenant-Basisdaten'

$acceptedDomains = Invoke-SafeCommand 'Get-AcceptedDomain' { Get-AcceptedDomain }
if ($null -ne $acceptedDomains) {
    $script:RawData['AcceptedDomains'] = $acceptedDomains
    $domList = @()
    foreach ($d in $acceptedDomains) { $domList += $d.DomainName }
    Write-Info ('Akzeptierte Domaenen: ' + ($domList -join ', '))
    Add-Finding -Id 'INFO-01' -Category 'Tenant' -Check 'Akzeptierte Domaenen' `
        -Actual $domList -Rating 'Info'
}

# ===================================================================================
#  Policy-Inventar: welche Policies existieren, welche Regel gehoert dazu,
#  und wirkt die Policy ueberhaupt auf jemanden?
# ===================================================================================

Write-Step 'Policy-Inventar'

$script:Rules = @{}
$script:Rules['HostedContentFilter']     = Invoke-SafeCommand 'Get-HostedContentFilterRule'     { Get-HostedContentFilterRule }
$script:Rules['AntiPhish']               = Invoke-SafeCommand 'Get-AntiPhishRule'               { Get-AntiPhishRule }
$script:Rules['MalwareFilter']           = Invoke-SafeCommand 'Get-MalwareFilterRule'           { Get-MalwareFilterRule }
$script:Rules['SafeLinks']               = Invoke-SafeCommand 'Get-SafeLinksRule'               { Get-SafeLinksRule }
$script:Rules['SafeAttachment']          = Invoke-SafeCommand 'Get-SafeAttachmentRule'          { Get-SafeAttachmentRule }
$script:Rules['HostedOutboundSpamFilter']= Invoke-SafeCommand 'Get-HostedOutboundSpamFilterRule'{ Get-HostedOutboundSpamFilterRule }
foreach ($k in @($script:Rules.Keys)) { $script:RawData[$k + 'Rule'] = $script:Rules[$k] }

# Policy-Eigenschaft, ueber die eine Regel auf ihre Policy zeigt (je Cmdlet anders benannt)
$script:RuleProps = @{
    'HostedContentFilter'      = @('HostedContentFilterPolicy')
    'AntiPhish'                = @('AntiPhishPolicy')
    'MalwareFilter'            = @('MalwareFilterPolicy')
    'SafeLinks'                = @('SafeLinksPolicy')
    'SafeAttachment'           = @('SafeAttachmentPolicy')
    'HostedOutboundSpamFilter' = @('HostedOutboundSpamFilterPolicy')
}

function Write-PolicyInventory {
    <# Listet alle Policies eines Typs mit Regelzuordnung und Wirksamkeit auf. #>
    param(
        [string]$Id,
        [string]$Label,
        $Policies,
        [string]$RuleKey,
        [switch]$DefaultByName
    )

    if ($null -eq $Policies) {
        Add-Finding -Id $Id -Category 'Policy-Inventar' -Check ($Label + ': vorhandene Policies') `
            -Actual 'nicht verfuegbar' -Expected 'mindestens die Default-Policy' -Rating 'n.a.'
        return
    }

    $rules = $null
    if ($script:Rules.ContainsKey($RuleKey)) { $rules = $script:Rules[$RuleKey] }
    $props = @('Policy')
    if ($script:RuleProps.ContainsKey($RuleKey)) { $props = $script:RuleProps[$RuleKey] }

    $lines = @()
    $aktiv = 0; $inaktiv = 0; $preset = 0
    foreach ($p in @($Policies)) {
        $name = [string](Get-SafeProperty $p 'Name')
        if ([string]::IsNullOrWhiteSpace($name)) { $name = [string](Get-SafeProperty $p 'Identity') }

        $isDef = $false
        if ($DefaultByName.IsPresent -and $name -eq 'Default') { $isDef = $true }

        $status = Set-PolicyContext -Policy $p -Rules $rules -PolicyProperties $props -IsDefaultPolicy:$isDef
        $rule   = Get-RuleForPolicy -Rules $rules -PolicyName $name -PolicyProperties $props
        $scope  = Get-RuleScopeText $rule

        if     ($status -like 'Preset*')   { $preset++ }
        elseif ($script:CurrentPolicyActive) { $aktiv++ }
        else                                 { $inaktiv++ }

        $lines += ('{0}  [{1}]  Scope: {2}' -f $name, $status, $scope)
    }
    Reset-PolicyContext

    $summary = ('{0} Policies: {1} wirksam, {2} ohne Wirkung, {3} Preset' -f @($Policies).Count, $aktiv, $inaktiv, $preset)
    Write-Info ($Label + ': ' + $summary)

    $rate = 'Info'
    if ($inaktiv -gt 0) { $rate = 'Verbesserung empfohlen' }

    Add-Finding -Id $Id -Category 'Policy-Inventar' -Check ($Label + ': vorhandene Policies') `
        -Actual $lines -Expected 'jede Policy hat eine aktive Regel oder ist bewusst stillgelegt' -Rating $rate `
        -Note ($summary + '. Policies ohne Regel sind im Defender-Portal unsichtbar und wirken auf niemanden - sie werden im Audit nur berichtet, nicht bemaengelt.')
}

# ===================================================================================
#  PRE  Preset Security Policies
# ===================================================================================

Write-Step 'PRE  Preset Security Policies'

$eopRules = Invoke-SafeCommand 'Get-EOPProtectionPolicyRule' { Get-EOPProtectionPolicyRule }
$atpRules = Invoke-SafeCommand 'Get-ATPProtectionPolicyRule' { Get-ATPProtectionPolicyRule }
$script:RawData['EOPProtectionPolicyRule'] = $eopRules
$script:RawData['ATPProtectionPolicyRule'] = $atpRules

$eopStates = @()
if ($null -ne $eopRules) {
    foreach ($r in $eopRules) { $eopStates += ('{0} = {1}' -f $r.Name, $r.State) }
}
if ($eopStates.Count -eq 0) { $eopStates = @('keine Preset-Regeln vorhanden') }

$atpStates = @()
if ($null -ne $atpRules) {
    foreach ($r in $atpRules) { $atpStates += ('{0} = {1}' -f $r.Name, $r.State) }
}
if ($atpStates.Count -eq 0) { $atpStates = @('keine Preset-Regeln vorhanden (oder kein Defender)') }

Add-Finding -Id 'PRE-01' -Category 'Preset Security Policies' `
    -Check 'Preset-Status EOP-Teil (Standard / Strict)' -Actual $eopStates `
    -Expected 'Bewusste Entscheidung: Preset ODER Custom-Policies' -Rating 'Info' `
    -Note 'EOP-Teil und MDO-Teil werden getrennt aktiviert. Ein Preset kann halb aktiv sein.'

Add-Finding -Id 'PRE-01b' -Category 'Preset Security Policies' `
    -Check 'Preset-Status MDO-Teil (Standard / Strict)' -Actual $atpStates `
    -Expected 'Bewusste Entscheidung' -Rating 'Info'

$builtIn = Invoke-SafeCommand 'Get-ATPBuiltInProtectionRule' { Get-ATPBuiltInProtectionRule }
$script:RawData['ATPBuiltInProtectionRule'] = $builtIn
if ($null -ne $builtIn) {
    $ex = @()
    $exDom  = Get-SafeProperty $builtIn 'ExceptIfRecipientDomainIs'
    $exTo   = Get-SafeProperty $builtIn 'ExceptIfSentTo'
    $exGrp  = Get-SafeProperty $builtIn 'ExceptIfSentToMemberOf'
    if ($null -ne $exDom -and @($exDom).Count -gt 0) { $ex += ('Domaenen: ' + (@($exDom) -join ', ')) }
    if ($null -ne $exTo  -and @($exTo).Count  -gt 0) { $ex += ('Empfaenger: ' + @($exTo).Count) }
    if ($null -ne $exGrp -and @($exGrp).Count -gt 0) { $ex += ('Gruppen: ' + @($exGrp).Count) }

    if ($ex.Count -eq 0) {
        Add-Finding -Id 'PRE-02' -Category 'Preset Security Policies' `
            -Check 'Built-in protection ohne Ausnahmen' -Actual 'keine Ausnahmen' `
            -Expected 'keine Ausnahmen' -Rating 'OK'
    } else {
        $rate = 'Verbesserung empfohlen'
        if ($null -ne $acceptedDomains -and $null -ne $exDom) {
            if (@($exDom).Count -ge @($acceptedDomains).Count) { $rate = 'Kritisch' }
        }
        Add-Finding -Id 'PRE-02' -Category 'Preset Security Policies' `
            -Check 'Built-in protection ohne Ausnahmen' -Actual $ex `
            -Expected 'keine Ausnahmen' -Rating $rate `
            -Note 'Sind alle Accepted Domains ausgenommen, ist Safe Links / Safe Attachments per Built-in abgeschaltet.'
    }
}

# ===================================================================================
#  ASI / ASF  Anti-Spam inbound
# ===================================================================================

Write-Step 'ASI / ASF  Anti-Spam eingehend'

$contentPolicies = Invoke-SafeCommand 'Get-HostedContentFilterPolicy' { Get-HostedContentFilterPolicy }
$script:RawData['HostedContentFilterPolicy'] = $contentPolicies

Write-PolicyInventory -Id 'INV-ASI' -Label 'Anti-Spam inbound' -Policies $contentPolicies -RuleKey 'HostedContentFilter' -DefaultByName

if ($null -ne $contentPolicies) {
    foreach ($p in $contentPolicies) {
        $n = $p.Name
        $null = Set-PolicyContext -Policy $p -Rules $script:Rules['HostedContentFilter'] `
                    -PolicyProperties $script:RuleProps['HostedContentFilter'] -IsDefaultPolicy:($n -eq 'Default')

        Test-Value -Id 'ASI-01' -Category 'Anti-Spam inbound' -Check 'Bulk-Schwellwert (BulkThreshold)' `
            -Scope $n -Actual (Get-SafeProperty $p 'BulkThreshold') -Accept @('6','5') `
            -ExpectedText '6 (MS Standard) oder 5 (MS Strict)' `
            -Note 'Default ist 7. Empfehlung von 2023 6 deckt sich mit dem MS-Standard-Preset.'

        Test-Value -Id 'ASI-02' -Category 'Anti-Spam inbound' -Check 'Aktion bei Spam' `
            -Scope $n -Actual (Get-SafeProperty $p 'SpamAction') -Accept @('MoveToJmf','Quarantine') `
            -ExpectedText 'MoveToJmf oder Quarantine'

        Test-Value -Id 'ASI-03' -Category 'Anti-Spam inbound' -Check 'Aktion bei High Confidence Spam' `
            -Scope $n -Actual (Get-SafeProperty $p 'HighConfidenceSpamAction') -Accept @('Quarantine') `
            -ExpectedText 'Quarantine' -FailRating 'Verbesserung empfohlen' `
            -Note 'Default ist MoveToJmf. Beide Presets setzen Quarantine.'

        Test-Value -Id 'ASI-04' -Category 'Anti-Spam inbound' -Check 'Aktion bei Phishing' `
            -Scope $n -Actual (Get-SafeProperty $p 'PhishSpamAction') -Accept @('Quarantine') `
            -ExpectedText 'Quarantine' -FailRating 'Kritisch' `
            -Note 'Steht dies auf MoveToJmf, landet erkanntes Phishing im Junk-Ordner des Nutzers und ist dort anklickbar.'

        Test-Value -Id 'ASI-05' -Category 'Anti-Spam inbound' -Check 'Aktion bei High Confidence Phishing' `
            -Scope $n -Actual (Get-SafeProperty $p 'HighConfidencePhishAction') -Accept @('Quarantine') `
            -ExpectedText 'Quarantine' -FailRating 'Kritisch'

        Test-Value -Id 'ASI-05b' -Category 'Anti-Spam inbound' -Check 'Quarantaene-Policy fuer High Confidence Phishing' `
            -Scope $n -Actual (Get-SafeProperty $p 'HighConfidencePhishQuarantineTag') -Accept @('AdminOnlyAccessPolicy') `
            -ExpectedText 'AdminOnlyAccessPolicy' -FailRating 'Kritisch'

        Test-Value -Id 'ASI-06' -Category 'Anti-Spam inbound' -Check 'Aktion bei Bulk' `
            -Scope $n -Actual (Get-SafeProperty $p 'BulkSpamAction') -Accept @('MoveToJmf','Quarantine') `
            -ExpectedText 'MoveToJmf oder Quarantine'

        Test-Value -Id 'ASI-06b' -Category 'Anti-Spam inbound' -Check 'Bulk-Filterung aktiv (MarkAsSpamBulkMail)' `
            -Scope $n -Actual (Get-SafeProperty $p 'MarkAsSpamBulkMail') -Accept @('On') `
            -ExpectedText 'On' -FailRating 'Kritisch' `
            -Note 'Steht dies auf Off, ist die BCL-Schwelle wirkungslos.'

        Test-Value -Id 'ASI-07' -Category 'Anti-Spam inbound' -Check 'Spam Safety Tips' `
            -Scope $n -Actual (Get-SafeProperty $p 'InlineSafetyTipsEnabled') -Accept @('True') `
            -ExpectedText 'True'

        Test-Value -Id 'ASI-08a' -Category 'Anti-Spam inbound' -Check 'ZAP fuer Spam' `
            -Scope $n -Actual (Get-SafeProperty $p 'SpamZapEnabled') -Accept @('True') `
            -ExpectedText 'True' -FailRating 'Kritisch'

        Test-Value -Id 'ASI-08b' -Category 'Anti-Spam inbound' -Check 'ZAP fuer Phishing' `
            -Scope $n -Actual (Get-SafeProperty $p 'PhishZapEnabled') -Accept @('True') `
            -ExpectedText 'True' -FailRating 'Kritisch'

        Test-Value -Id 'ASI-09' -Category 'Anti-Spam inbound' -Check 'Quarantaene-Aufbewahrung (Tage)' `
            -Scope $n -Actual (Get-SafeProperty $p 'QuarantineRetentionPeriod') -Accept @('30') `
            -ExpectedText '30 (Maximum)' `
            -Note 'Default ist 15. Bei 15 Tagen findet ein Nutzer nach dem Urlaub nichts mehr vor.'

        # Allow-Listen
        $allowSenders = @(Get-SafeProperty $p 'AllowedSenders')
        $allowDomains = @(Get-SafeProperty $p 'AllowedSenderDomains')
        $cntA = 0
        if ($null -ne $allowSenders) { $cntA += $allowSenders.Count }
        $cntD = 0
        if ($null -ne $allowDomains) { $cntD += $allowDomains.Count }

        $rate = 'OK'
        if ($cntD -gt 0) { $rate = 'Kritisch' }
        elseif ($cntA -gt 0) { $rate = 'Verbesserung empfohlen' }

        Add-Finding -Id 'ASI-10' -Category 'Anti-Spam inbound' `
            -Check 'Allow-Listen in der Anti-Spam-Policy' -Scope $n `
            -Actual ('Absender: ' + $cntA + ', Domaenen: ' + $cntD) -Expected 'beide leer' -Rating $rate `
            -Note 'AllowedSenderDomains setzt SCL -1 und ueberspringt die komplette Spam- und Phishing-Filterung. Ausnahmen gehoeren in die Tenant Allow/Block List.'

        Test-Value -Id 'ASI-11a' -Category 'Anti-Spam inbound' -Check 'Sprachfilter' `
            -Scope $n -Actual (Get-SafeProperty $p 'EnableLanguageBlockList') -Accept @('False') `
            -ExpectedText 'False (kein Preset aktiviert ihn)'

        Test-Value -Id 'ASI-11b' -Category 'Anti-Spam inbound' -Check 'Laenderfilter' `
            -Scope $n -Actual (Get-SafeProperty $p 'EnableRegionBlockList') -Accept @('False') `
            -ExpectedText 'False (kein Preset aktiviert ihn)'

        Add-Finding -Id 'ASI-12' -Category 'Anti-Spam inbound' -Check 'Filterung interner Nachrichten' `
            -Scope $n -Actual (Get-SafeProperty $p 'IntraOrgFilterState') `
            -Expected 'Default (= High Confidence Phishing)' -Rating 'Info'

        # ---- ASF -------------------------------------------------------------
        $asfIncrease = @(
            'IncreaseScoreWithImageLinks','IncreaseScoreWithNumericIps',
            'IncreaseScoreWithRedirectToOtherPort','IncreaseScoreWithBizOrInfoUrls'
        )
        $asfMark = @(
            'MarkAsSpamEmptyMessages','MarkAsSpamEmbedTagsInHtml','MarkAsSpamJavaScriptInHtml',
            'MarkAsSpamFormTagsInHtml','MarkAsSpamFramesInHtml','MarkAsSpamWebBugsInHtml',
            'MarkAsSpamObjectTagsInHtml','MarkAsSpamSensitiveWordList'
        )

        $onIncrease = @()
        foreach ($a in $asfIncrease) {
            $v = Get-SafeProperty $p $a
            if ($null -ne $v -and [string]$v -ne 'Off') { $onIncrease += ($a + ' = ' + $v) }
        }
        $rate = 'OK'
        if ($onIncrease.Count -gt 0) { $rate = 'Verbesserung empfohlen' }
        $val = 'alle Off'
        if ($onIncrease.Count -gt 0) { $val = $onIncrease }
        Add-Finding -Id 'ASF-01' -Category 'Advanced Spam Filter (ASF)' `
            -Check 'ASF Score-Erhoeher' -Scope $n -Actual $val -Expected 'alle Off' -Rating $rate `
            -Note 'ASF-Treffer koennen nicht als False Positive an Microsoft gemeldet werden.'

        $onMark = @()
        foreach ($a in $asfMark) {
            $v = Get-SafeProperty $p $a
            if ($null -ne $v -and [string]$v -ne 'Off') { $onMark += ($a + ' = ' + $v) }
        }
        $rate = 'OK'
        if ($onMark.Count -gt 0) { $rate = 'Verbesserung empfohlen' }
        $val = 'alle Off'
        if ($onMark.Count -gt 0) { $val = $onMark }
        Add-Finding -Id 'ASF-02' -Category 'Advanced Spam Filter (ASF)' `
            -Check 'ASF Klassifizierung als Spam' -Scope $n -Actual $val -Expected 'alle Off' -Rating $rate `
            -Note 'Diese Treffer setzen SCL 9. Newsletter mit Tracking-Pixel oder Formular sind damit High Confidence Spam.'

        Test-Value -Id 'ASF-03' -Category 'Advanced Spam Filter (ASF)' -Check 'SPF Hard Fail als Spam-Kriterium' `
            -Scope $n -Actual (Get-SafeProperty $p 'MarkAsSpamSpfRecordHardFail') -Accept @('Off') `
            -ExpectedText 'Off' `
            -Note 'Microsoft: "No. This ASF setting is no longer required."'

        Test-Value -Id 'ASF-04a' -Category 'Advanced Spam Filter (ASF)' -Check 'Sender ID Hard Fail' `
            -Scope $n -Actual (Get-SafeProperty $p 'MarkAsSpamFromAddressAuthFail') -Accept @('Off') `
            -ExpectedText 'Off'

        Test-Value -Id 'ASF-04b' -Category 'Advanced Spam Filter (ASF)' -Check 'NDR Backscatter' `
            -Scope $n -Actual (Get-SafeProperty $p 'MarkAsSpamNdrBackscatter') -Accept @('Off') `
            -ExpectedText 'Off'

        Test-Value -Id 'ASF-05' -Category 'Advanced Spam Filter (ASF)' -Check 'ASF Test-Modus' `
            -Scope $n -Actual (Get-SafeProperty $p 'TestModeAction') -Accept @('None') `
            -ExpectedText 'None'
    }
    Reset-PolicyContext
}

# ===================================================================================
#  CF  Connection Filter
# ===================================================================================

Write-Step 'CF  Connection Filter'

$connFilter = Invoke-SafeCommand 'Get-HostedConnectionFilterPolicy' { Get-HostedConnectionFilterPolicy -Identity Default }
$script:RawData['HostedConnectionFilterPolicy'] = $connFilter

if ($null -ne $connFilter) {
    $ipAllow = @(Get-SafeProperty $connFilter 'IPAllowList')
    $ipBlock = @(Get-SafeProperty $connFilter 'IPBlockList')

    $cnt = 0
    if ($null -ne $ipAllow) { $cnt = $ipAllow.Count }
    $rate = 'OK'
    if ($cnt -gt 0) { $rate = 'Verbesserung empfohlen' }
    $val = '<leer>'
    if ($cnt -gt 0) { $val = $ipAllow }
    Add-Finding -Id 'CF-01' -Category 'Connection Filter' -Check 'IP Allow List' `
        -Scope 'Default' -Actual $val -Expected 'leer' -Rating $rate `
        -Note 'Setzt SCL -1: Spam-, Phishing-, Bulk- und Spoof-Pruefung werden uebersprungen. Jeden Eintrag einzeln begruenden.'

    $cntB = 0
    if ($null -ne $ipBlock) { $cntB = $ipBlock.Count }
    Add-Finding -Id 'CF-02' -Category 'Connection Filter' -Check 'IP Block List' `
        -Scope 'Default' -Actual ($cntB.ToString() + ' Eintraege') -Expected 'nach Bedarf' -Rating 'Info' `
        -Note 'Blockierte Nachrichten erscheinen NICHT im Message Trace.'

    Test-Value -Id 'CF-03' -Category 'Connection Filter' -Check 'Safe List' `
        -Scope 'Default' -Actual (Get-SafeProperty $connFilter 'EnableSafeList') -Accept @('False') `
        -ExpectedText 'False'
}

# ===================================================================================
#  ASO  Anti-Spam outbound
# ===================================================================================

Write-Step 'ASO  Anti-Spam ausgehend'

$outPolicies = Invoke-SafeCommand 'Get-HostedOutboundSpamFilterPolicy' { Get-HostedOutboundSpamFilterPolicy }
$script:RawData['HostedOutboundSpamFilterPolicy'] = $outPolicies

Write-PolicyInventory -Id 'INV-ASO' -Label 'Anti-Spam outbound' -Policies $outPolicies -RuleKey 'HostedOutboundSpamFilter' -DefaultByName

if ($null -ne $outPolicies) {
    foreach ($p in $outPolicies) {
        $n = $p.Name
        $null = Set-PolicyContext -Policy $p -Rules $script:Rules['HostedOutboundSpamFilter'] `
                    -PolicyProperties $script:RuleProps['HostedOutboundSpamFilter'] -IsDefaultPolicy:($n -eq 'Default')

        Add-Finding -Id 'ASO-02' -Category 'Anti-Spam outbound' -Check 'Empfaenger-Limits' -Scope $n `
            -Actual ('extern/h: ' + (Get-SafeProperty $p 'RecipientLimitExternalPerHour') +
                     ', intern/h: ' + (Get-SafeProperty $p 'RecipientLimitInternalPerHour') +
                     ', pro Tag: '  + (Get-SafeProperty $p 'RecipientLimitPerDay')) `
            -Expected '500 / 1000 / 1000 (MS Standard)' -Rating 'Info' `
            -Note '0 bedeutet: Service-Defaults. Vor dem Absenken das tatsaechliche Spitzenvolumen messen.'

        Test-Value -Id 'ASO-03' -Category 'Anti-Spam outbound' -Check 'Aktion bei Limitueberschreitung' `
            -Scope $n -Actual (Get-SafeProperty $p 'ActionWhenThresholdReached') -Accept @('BlockUser') `
            -ExpectedText 'BlockUser' `
            -Note 'Beide Presets setzen BlockUser. Alert allein stoppt ein kompromittiertes Konto nicht.'

        Test-Value -Id 'ASO-04' -Category 'Anti-Spam outbound' -Check 'Automatische externe Weiterleitungen' `
            -Scope $n -Actual (Get-SafeProperty $p 'AutoForwardingMode') -Accept @('Automatic','Off') `
            -ExpectedText 'Automatic (systemgesteuert = aus) oder Off' -FailRating 'Kritisch' `
            -Note 'Steht dies auf On, ist der haeufigste Exfiltrationsweg nach einer Kontokompromittierung offen.'

        Add-Finding -Id 'ASO-05' -Category 'Anti-Spam outbound' -Check 'Benachrichtigung / BCC bei verdaechtigem Ausgangsverkehr' `
            -Scope $n -Actual ('NotifyOutboundSpam: ' + (Get-SafeProperty $p 'NotifyOutboundSpam') +
                               ', Bcc: ' + (Get-SafeProperty $p 'BccSuspiciousOutboundMail')) `
            -Expected 'bewusste Entscheidung (Preset: beide False, Secure Score verlangt True)' -Rating 'Info' `
            -Note 'NotifyOutboundSpam ist laut Microsoft in Abkuendigung. Stattdessen Alert Policies nutzen.'
    }
    Reset-PolicyContext
}

# ===================================================================================
#  APH / IMP  Anti-Phishing
# ===================================================================================

Write-Step 'APH / IMP  Anti-Phishing und Impersonation'

$phishPolicies = Invoke-SafeCommand 'Get-AntiPhishPolicy' { Get-AntiPhishPolicy }
$script:RawData['AntiPhishPolicy'] = $phishPolicies

Write-PolicyInventory -Id 'INV-APH' -Label 'Anti-Phishing' -Policies $phishPolicies -RuleKey 'AntiPhish'

if ($null -ne $phishPolicies) {
    foreach ($p in $phishPolicies) {
        $n = $p.Name
        $null = Set-PolicyContext -Policy $p -Rules $script:Rules['AntiPhish'] `
                    -PolicyProperties $script:RuleProps['AntiPhish'] `
                    -IsDefaultPolicy:($n -eq 'Office365 AntiPhish Default')

        Test-Value -Id 'APH-01' -Category 'Anti-Phishing / Spoof' -Check 'Spoof Intelligence' `
            -Scope $n -Actual (Get-SafeProperty $p 'EnableSpoofIntelligence') -Accept @('True') `
            -ExpectedText 'True' -FailRating 'Kritisch'

        Test-Value -Id 'APH-02' -Category 'Anti-Phishing / Spoof' -Check 'Aktion bei Spoof-Erkennung' `
            -Scope $n -Actual (Get-SafeProperty $p 'AuthenticationFailAction') -Accept @('Quarantine') `
            -ExpectedText 'Quarantine (MS Strict)' `
            -Note 'Default und MS Standard setzen MoveToJmf. Die Empfehlung von 2023 liegt auf Strict-Niveau.'

        Test-Value -Id 'APH-03' -Category 'Anti-Phishing / Spoof' -Check 'DMARC-Policy des Absenders beruecksichtigen' `
            -Scope $n -Actual (Get-SafeProperty $p 'HonorDmarcPolicy') -Accept @('True') `
            -ExpectedText 'True' -FailRating 'Kritisch' `
            -Note 'Wirkt nur, wenn der MX direkt auf EXO zeigt ODER Enhanced Filtering am Inbound-Connector aktiv ist.'

        Test-Value -Id 'APH-04' -Category 'Anti-Phishing / Spoof' -Check 'Aktion bei DMARC p=quarantine' `
            -Scope $n -Actual (Get-SafeProperty $p 'DmarcQuarantineAction') -Accept @('Quarantine') `
            -ExpectedText 'Quarantine'

        Test-Value -Id 'APH-05' -Category 'Anti-Phishing / Spoof' -Check 'Aktion bei DMARC p=reject' `
            -Scope $n -Actual (Get-SafeProperty $p 'DmarcRejectAction') -Accept @('Reject','Quarantine') `
            -ExpectedText 'Reject (Zielzustand) oder Quarantine (Einfuehrungsphase)'

        Test-Value -Id 'APH-06a' -Category 'Anti-Phishing / Spoof' -Check 'Unauthentifizierter Absender (Fragezeichen)' `
            -Scope $n -Actual (Get-SafeProperty $p 'EnableUnauthenticatedSender') -Accept @('True') `
            -ExpectedText 'True'

        Test-Value -Id 'APH-06b' -Category 'Anti-Phishing / Spoof' -Check 'via-Tag' `
            -Scope $n -Actual (Get-SafeProperty $p 'EnableViaTag') -Accept @('True') `
            -ExpectedText 'True'

        Test-Value -Id 'APH-07' -Category 'Anti-Phishing / Spoof' -Check 'First Contact Safety Tip' `
            -Scope $n -Actual (Get-SafeProperty $p 'EnableFirstContactSafetyTips') -Accept @('True') `
            -ExpectedText 'True' `
            -Note 'Default der Default-Policy ist False. Beide Presets setzen True, Microsoft empfiehlt es ausdruecklich.'

        # --- Defender-Teil: nur bewerten, wenn die Eigenschaften vorhanden sind
        $thr = Get-SafeProperty $p 'PhishThresholdLevel'
        if ($null -ne $thr) {
            Test-Value -Id 'APH-08' -Category 'Impersonation (MDO)' -Check 'Phishing-Schwellwert' `
                -Scope $n -Actual $thr -Accept @('3','4') `
                -ExpectedText '3 (MS Standard) oder 4 (MS Strict)' `
                -Note 'Die Empfehlung von 2023 2 liegt unter dem heutigen MS-Standard. In der Default-Policy wirkt der Wert nicht.'
        }

        $orgDom = Get-SafeProperty $p 'EnableOrganizationDomainsProtection'
        if ($null -ne $orgDom) {
            Test-Value -Id 'IMP-01' -Category 'Impersonation (MDO)' -Check 'Eigene Domaenen geschuetzt' `
                -Scope $n -Actual $orgDom -Accept @('True') -ExpectedText 'True'

            $targetUsers = @(Get-SafeProperty $p 'TargetedUsersToProtect')
            $cntU = 0
            if ($null -ne $targetUsers) { $cntU = $targetUsers.Count }
            $rate = 'Verbesserung empfohlen'
            if ((Get-SafeProperty $p 'EnableTargetedUserProtection') -eq $true -and $cntU -gt 0) { $rate = 'OK' }
            Add-Finding -Id 'IMP-03' -Category 'Impersonation (MDO)' -Check 'Geschuetzte Benutzer' -Scope $n `
                -Actual ('aktiv: ' + (Get-SafeProperty $p 'EnableTargetedUserProtection') + ', Anzahl: ' + $cntU) `
                -Expected 'aktiv, mit Geschaeftsfuehrung / Buchhaltung / Einkauf' -Rating $rate `
                -Note 'Verteidigung gegen CEO-Fraud. Format: "Anzeigename;E-Mail-Adresse". Max. 350 je Policy.'

            $targetDoms = @(Get-SafeProperty $p 'TargetedDomainsToProtect')
            $cntD2 = 0
            if ($null -ne $targetDoms) { $cntD2 = $targetDoms.Count }
            Add-Finding -Id 'IMP-02' -Category 'Impersonation (MDO)' -Check 'Geschuetzte Fremd-Domaenen' -Scope $n `
                -Actual ('aktiv: ' + (Get-SafeProperty $p 'EnableTargetedDomainsProtection') + ', Anzahl: ' + $cntD2) `
                -Expected 'Hausbank, Steuerberater, Hauptlieferanten' -Rating 'Info' -Note 'Max. 50 je Policy.'

            $actU = [string](Get-SafeProperty $p 'TargetedUserProtectionAction')
            $actD = [string](Get-SafeProperty $p 'TargetedDomainProtectionAction')
            $actM = [string](Get-SafeProperty $p 'MailboxIntelligenceProtectionAction')
            $noAction = @()
            if ($actU -eq 'NoAction') { $noAction += 'User' }
            if ($actD -eq 'NoAction') { $noAction += 'Domain' }
            if ($actM -eq 'NoAction') { $noAction += 'MailboxIntelligence' }
            $rate = 'OK'
            if ($noAction.Count -gt 0) { $rate = 'Kritisch' }
            Add-Finding -Id 'IMP-04' -Category 'Impersonation (MDO)' -Check 'Impersonation-Aktionen gesetzt' -Scope $n `
                -Actual ('User: ' + $actU + ', Domain: ' + $actD + ', MbxInt: ' + $actM) `
                -Expected 'Quarantine / Quarantine / MoveToJmf' -Rating $rate `
                -Note 'NoAction bedeutet: Erkennung laeuft, aber es passiert nichts. Haeufigster Befund.'

            Test-Value -Id 'IMP-05a' -Category 'Impersonation (MDO)' -Check 'Mailbox Intelligence' `
                -Scope $n -Actual (Get-SafeProperty $p 'EnableMailboxIntelligence') -Accept @('True') -ExpectedText 'True'

            Test-Value -Id 'IMP-05b' -Category 'Impersonation (MDO)' -Check 'Mailbox Intelligence Schutz (eigene Aktion)' `
                -Scope $n -Actual (Get-SafeProperty $p 'EnableMailboxIntelligenceProtection') -Accept @('True') `
                -ExpectedText 'True' -Note 'Zweiter Schalter, ab Werk aus.'

            $tips = @()
            foreach ($t in @('EnableSimilarUsersSafetyTips','EnableSimilarDomainsSafetyTips','EnableUnusualCharactersSafetyTips')) {
                $tips += ($t + '=' + (Get-SafeProperty $p $t))
            }
            $allOn = ((Get-SafeProperty $p 'EnableSimilarUsersSafetyTips') -eq $true -and
                      (Get-SafeProperty $p 'EnableSimilarDomainsSafetyTips') -eq $true -and
                      (Get-SafeProperty $p 'EnableUnusualCharactersSafetyTips') -eq $true)
            $rate = 'Verbesserung empfohlen'
            if ($allOn) { $rate = 'OK' }
            Add-Finding -Id 'IMP-06' -Category 'Impersonation (MDO)' -Check 'Impersonation Safety Tips' -Scope $n `
                -Actual $tips -Expected 'alle True' -Rating $rate

            $exS = @(Get-SafeProperty $p 'ExcludedSenders')
            $exD = @(Get-SafeProperty $p 'ExcludedDomains')
            $cS = 0
            if ($null -ne $exS) { $cS = $exS.Count }
            $cD = 0
            if ($null -ne $exD) { $cD = $exD.Count }
            Add-Finding -Id 'IMP-07' -Category 'Impersonation (MDO)' -Check 'Impersonation-Ausnahmen' -Scope $n `
                -Actual ('Absender: ' + $cS + ', Domaenen: ' + $cD) -Expected 'minimal und begruendet' -Rating 'Info' `
                -Note 'Waechst im Betrieb automatisch durch False-Positive-Meldungen. Regelmaessig durchsehen.'
        }
        else {
            Add-Finding -Id 'IMP-00' -Category 'Impersonation (MDO)' `
                -Check 'Impersonation-Einstellungen vorhanden' -Scope $n `
                -Actual 'nicht vorhanden' -Expected 'Defender P1 oder P2' -Rating 'n.a.' `
                -Note 'Impersonation-Schutz und Phishing-Schwellwert erfordern Defender for Office 365.'
        }
    }
    Reset-PolicyContext
}

$emailTenant = Invoke-SafeCommand 'Get-EmailTenantSettings' { Get-EmailTenantSettings }
if ($null -ne $emailTenant) {
    $script:RawData['EmailTenantSettings'] = $emailTenant
    Test-Value -Id 'IMP-08' -Category 'Impersonation (MDO)' -Check 'Priority Account Protection' `
        -Actual (Get-SafeProperty $emailTenant 'EnablePriorityAccountProtection') -Accept @('True') `
        -ExpectedText 'True (Defender P2)' `
        -Note 'VIPs zusaetzlich im M365 Admin Center als Priority Account taggen - das ersetzt TargetedUsersToProtect nicht.'
}

# ===================================================================================
#  TABL
# ===================================================================================

Write-Step 'TABL  Tenant Allow/Block List'

foreach ($lt in @('Sender','Url','FileHash','IP')) {
    $allow = Invoke-SafeCommand ('Get-TenantAllowBlockListItems ' + $lt) { Get-TenantAllowBlockListItems -ListType $lt -Allow }
    $block = Invoke-SafeCommand ('Get-TenantAllowBlockListItems ' + $lt) { Get-TenantAllowBlockListItems -ListType $lt -Block }
    $script:RawData['TABL_' + $lt + '_Allow'] = $allow
    $script:RawData['TABL_' + $lt + '_Block'] = $block

    $ca = 0
    if ($null -ne $allow) { $ca = @($allow).Count }
    $cb = 0
    if ($null -ne $block) { $cb = @($block).Count }

    Add-Finding -Id ('TABL-' + $lt) -Category 'Tenant Allow/Block List' `
        -Check ('Eintraege vom Typ ' + $lt) -Actual ('Allow: ' + $ca + ', Block: ' + $cb) `
        -Expected 'jeder Allow-Eintrag mit Begruendung und Ablaufdatum' -Rating 'Info'
}

$spoofItems = Invoke-SafeCommand 'Get-TenantAllowBlockListSpoofItems' { Get-TenantAllowBlockListSpoofItems }
$script:RawData['TABL_SpoofItems'] = $spoofItems
if ($null -ne $spoofItems) {
    $spoofAllow = @($spoofItems | Where-Object { $_.Action -eq 'Allow' })
    $cnt = 0
    if ($null -ne $spoofAllow) { $cnt = $spoofAllow.Count }
    $rate = 'Info'
    if ($cnt -gt 0) { $rate = 'Verbesserung empfohlen' }
    Add-Finding -Id 'TABL-03' -Category 'Tenant Allow/Block List' -Check 'Freigegebene Spoof-Absender' `
        -Actual ($cnt.ToString() + ' Allow-Eintraege') -Expected 'nur wenn nicht anders loesbar' -Rating $rate `
        -Note 'Spoof-Allow-Eintraege laufen NIE ab. Vorrang hat die Korrektur von SPF/DKIM beim Absender.'
}

# ===================================================================================
#  AMW  Anti-Malware
# ===================================================================================

Write-Step 'AMW  Anti-Malware'

$malwarePolicies = Invoke-SafeCommand 'Get-MalwareFilterPolicy' { Get-MalwareFilterPolicy }
$script:RawData['MalwareFilterPolicy'] = $malwarePolicies

Write-PolicyInventory -Id 'INV-AMW' -Label 'Anti-Malware' -Policies $malwarePolicies -RuleKey 'MalwareFilter' -DefaultByName

if ($null -ne $malwarePolicies) {
    foreach ($p in $malwarePolicies) {
        $n = $p.Name
        $null = Set-PolicyContext -Policy $p -Rules $script:Rules['MalwareFilter'] `
                    -PolicyProperties $script:RuleProps['MalwareFilter'] -IsDefaultPolicy:($n -eq 'Default')

        Test-Value -Id 'AMW-01' -Category 'Anti-Malware' -Check 'Common Attachment Filter' `
            -Scope $n -Actual (Get-SafeProperty $p 'EnableFileFilter') -Accept @('True') `
            -ExpectedText 'True' -FailRating 'Kritisch'

        $ft = @(Get-SafeProperty $p 'FileTypes')
        $cnt = 0
        if ($null -ne $ft) { $cnt = $ft.Count }
        Add-Finding -Id 'AMW-01b' -Category 'Anti-Malware' -Check 'Anzahl gefilterter Dateitypen' `
            -Scope $n -Actual ($cnt.ToString() + ' Typen') -Expected 'mindestens die 54 Default-Typen' -Rating 'Info'

        Test-Value -Id 'AMW-02' -Category 'Anti-Malware' -Check 'ZAP fuer Malware' `
            -Scope $n -Actual (Get-SafeProperty $p 'ZapEnabled') -Accept @('True') `
            -ExpectedText 'True' -FailRating 'Kritisch'

        Test-Value -Id 'AMW-03' -Category 'Anti-Malware' -Check 'Aktion des Dateityp-Filters' `
            -Scope $n -Actual (Get-SafeProperty $p 'FileTypeAction') -Accept @('Reject','Quarantine') `
            -ExpectedText 'Reject (MS-Preset) oder Quarantine'

        Test-Value -Id 'AMW-04' -Category 'Anti-Malware' -Check 'Quarantaene-Policy fuer Malware' `
            -Scope $n -Actual (Get-SafeProperty $p 'QuarantineTag') -Accept @('AdminOnlyAccessPolicy') `
            -ExpectedText 'AdminOnlyAccessPolicy'

        Add-Finding -Id 'AMW-05' -Category 'Anti-Malware' -Check 'Admin-Benachrichtigungen' -Scope $n `
            -Actual ('intern: ' + (Get-SafeProperty $p 'EnableInternalSenderAdminNotifications') +
                     ', extern: ' + (Get-SafeProperty $p 'EnableExternalSenderAdminNotifications')) `
            -Expected 'optional - MS-Presets nutzen stattdessen Alert Policies' -Rating 'Info'
    }
    Reset-PolicyContext
}

# ===================================================================================
#  SA / SL  Safe Attachments und Safe Links
# ===================================================================================

Write-Step 'SA / SL  Safe Attachments und Safe Links (Defender)'

$safeAtt = Invoke-SafeCommand 'Get-SafeAttachmentPolicy' { Get-SafeAttachmentPolicy }
$script:RawData['SafeAttachmentPolicy'] = $safeAtt

if ($null -eq $safeAtt) {
    Add-Finding -Id 'SA-00' -Category 'Safe Attachments (MDO)' -Check 'Safe Attachments verfuegbar' `
        -Actual 'nicht verfuegbar' -Expected 'Defender P1 oder P2' -Rating 'n.a.'
}
else {
    Write-PolicyInventory -Id 'INV-SA' -Label 'Safe Attachments' -Policies $safeAtt -RuleKey 'SafeAttachment'
    foreach ($p in $safeAtt) {
        $n = $p.Name
        $null = Set-PolicyContext -Policy $p -Rules $script:Rules['SafeAttachment'] `
                    -PolicyProperties $script:RuleProps['SafeAttachment']
        Test-Value -Id 'SA-01a' -Category 'Safe Attachments (MDO)' -Check 'Policy aktiv' `
            -Scope $n -Actual (Get-SafeProperty $p 'Enable') -Accept @('True') -ExpectedText 'True' -FailRating 'Kritisch'

        Test-Value -Id 'SA-01b' -Category 'Safe Attachments (MDO)' -Check 'Aktion' `
            -Scope $n -Actual (Get-SafeProperty $p 'Action') -Accept @('Block') `
            -ExpectedText 'Block' -Note '"Replace" existiert nicht mehr. Gueltig sind Allow, Block, DynamicDelivery.'

        Test-Value -Id 'SA-01c' -Category 'Safe Attachments (MDO)' -Check 'Quarantaene-Policy' `
            -Scope $n -Actual (Get-SafeProperty $p 'QuarantineTag') -Accept @('AdminOnlyAccessPolicy') `
            -ExpectedText 'AdminOnlyAccessPolicy'

        $act = [string](Get-SafeProperty $p 'Action')
        $red = (Get-SafeProperty $p 'Redirect')
        if ($red -eq $true -and $act -ne 'Allow') {
            Add-Finding -Id 'SA-02' -Category 'Safe Attachments (MDO)' -Check 'Redirect konsistent zur Aktion' `
                -Scope $n -Actual ('Redirect=True bei Action=' + $act) -Expected 'Redirect nur bei Action=Allow' `
                -Rating 'Verbesserung empfohlen' -Note 'Redirect ist bei Action=Block wirkungslos, wird aber ohne Fehler akzeptiert.'
        } else {
            Add-Finding -Id 'SA-02' -Category 'Safe Attachments (MDO)' -Check 'Redirect konsistent zur Aktion' `
                -Scope $n -Actual ('Redirect=' + $red) -Expected 'konsistent' -Rating 'OK'
        }
    }
    Reset-PolicyContext
}

$atpO365 = Invoke-SafeCommand 'Get-AtpPolicyForO365' { Get-AtpPolicyForO365 }
$script:RawData['AtpPolicyForO365'] = $atpO365
if ($null -ne $atpO365) {
    Test-Value -Id 'SA-03' -Category 'Safe Attachments (MDO)' -Check 'Safe Attachments fuer SPO / OneDrive / Teams' `
        -Actual (Get-SafeProperty $atpO365 'EnableATPForSPOTeamsODB') -Accept @('True') -ExpectedText 'True' `
        -Note 'Zusaetzlich noetig: Set-SPOTenant -DisallowInfectedFileDownload $true (SharePoint Online Management Shell).'

    Add-Finding -Id 'SA-04' -Category 'Safe Attachments (MDO)' -Check 'Safe Documents' `
        -Actual ('EnableSafeDocs: ' + (Get-SafeProperty $atpO365 'EnableSafeDocs') +
                 ', AllowSafeDocsOpen: ' + (Get-SafeProperty $atpO365 'AllowSafeDocsOpen')) `
        -Expected 'True / False - nur mit M365 E5 bzw. E5 Security' -Rating 'Info' `
        -Note 'Safe Documents ist NICHT in MDO P1 oder P2 standalone enthalten.'
}

$safeLinks = Invoke-SafeCommand 'Get-SafeLinksPolicy' { Get-SafeLinksPolicy }
$script:RawData['SafeLinksPolicy'] = $safeLinks

if ($null -eq $safeLinks) {
    Add-Finding -Id 'SL-00' -Category 'Safe Links (MDO)' -Check 'Safe Links verfuegbar' `
        -Actual 'nicht verfuegbar' -Expected 'Defender P1 oder P2' -Rating 'n.a.'
}
else {
    Write-PolicyInventory -Id 'INV-SL' -Label 'Safe Links' -Policies $safeLinks -RuleKey 'SafeLinks'
    foreach ($p in $safeLinks) {
        $n = $p.Name
        $null = Set-PolicyContext -Policy $p -Rules $script:Rules['SafeLinks'] `
                    -PolicyProperties $script:RuleProps['SafeLinks']
        Test-Value -Id 'SL-01a' -Category 'Safe Links (MDO)' -Check 'Safe Links fuer E-Mail' `
            -Scope $n -Actual (Get-SafeProperty $p 'EnableSafeLinksForEmail') -Accept @('True') -ExpectedText 'True' -FailRating 'Kritisch'
        Test-Value -Id 'SL-01b' -Category 'Safe Links (MDO)' -Check 'Safe Links fuer Teams' `
            -Scope $n -Actual (Get-SafeProperty $p 'EnableSafeLinksForTeams') -Accept @('True') -ExpectedText 'True'
        Test-Value -Id 'SL-01c' -Category 'Safe Links (MDO)' -Check 'Safe Links fuer Office-Apps' `
            -Scope $n -Actual (Get-SafeProperty $p 'EnableSafeLinksForOffice') -Accept @('True') -ExpectedText 'True'
        Test-Value -Id 'SL-02' -Category 'Safe Links (MDO)' -Check 'Interne Absender einbeziehen' `
            -Scope $n -Actual (Get-SafeProperty $p 'EnableForInternalSenders') -Accept @('True') -ExpectedText 'True' `
            -Note 'Wichtig gegen laterale Ausbreitung. Built-in protection setzt hier False.'
        Test-Value -Id 'SL-03a' -Category 'Safe Links (MDO)' -Check 'Echtzeit-URL-Pruefung' `
            -Scope $n -Actual (Get-SafeProperty $p 'ScanUrls') -Accept @('True') -ExpectedText 'True'
        Test-Value -Id 'SL-03b' -Category 'Safe Links (MDO)' -Check 'Auf Scan-Abschluss warten' `
            -Scope $n -Actual (Get-SafeProperty $p 'DeliverMessageAfterScan') -Accept @('True') -ExpectedText 'True'
        Test-Value -Id 'SL-04a' -Category 'Safe Links (MDO)' -Check 'Klicks protokollieren' `
            -Scope $n -Actual (Get-SafeProperty $p 'TrackClicks') -Accept @('True') -ExpectedText 'True'
        Test-Value -Id 'SL-04b' -Category 'Safe Links (MDO)' -Check 'Durchklicken zur Original-URL' `
            -Scope $n -Actual (Get-SafeProperty $p 'AllowClickThrough') -Accept @('False') -ExpectedText 'False' -FailRating 'Kritisch'
        Test-Value -Id 'SL-05' -Category 'Safe Links (MDO)' -Check 'URL-Rewriting aktiv (nicht API-only)' `
            -Scope $n -Actual (Get-SafeProperty $p 'DisableUrlRewrite') -Accept @('False') -ExpectedText 'False'

        $dnr = @(Get-SafeProperty $p 'DoNotRewriteUrls')
        $cnt = 0
        if ($null -ne $dnr) { $cnt = $dnr.Count }
        $rate = 'OK'
        if ($cnt -gt 0) { $rate = 'Info' }
        Add-Finding -Id 'SL-06' -Category 'Safe Links (MDO)' -Check '"Do not rewrite"-Liste' -Scope $n `
            -Actual ($cnt.ToString() + ' Eintraege') -Expected 'leer oder minimal' -Rating $rate `
            -Note 'Diese Eintraege verhindern nur das Wrapping, NICHT die Einstufung als schaedlich. Dafuer TABL-URL-Allow.'

        Add-Finding -Id 'SL-07' -Category 'Safe Links (MDO)' -Check 'Warnseiten-Text und Branding' -Scope $n `
            -Actual ('Branding: ' + (Get-SafeProperty $p 'EnableOrganizationBranding') +
                     ', Text: ' + (ConvertTo-DisplayValue (Get-SafeProperty $p 'CustomNotificationText'))) `
            -Expected 'produktiver Text, Branding aktiv' -Rating 'Info' `
            -Note 'Pruefen, ob dort ein Testeintrag steht.'
    }
    Reset-PolicyContext
}

# ===================================================================================
#  QUA  Quarantaene
# ===================================================================================

Write-Step 'QUA  Quarantaene-Policies'

$quarPolicies = Invoke-SafeCommand 'Get-QuarantinePolicy' { Get-QuarantinePolicy }
$script:RawData['QuarantinePolicy'] = $quarPolicies
if ($null -ne $quarPolicies) {
    $names = @()
    foreach ($q in $quarPolicies) { $names += $q.Name }
    Add-Finding -Id 'QUA-01' -Category 'Quarantaene-Policies' -Check 'Vorhandene Quarantaene-Policies' `
        -Actual $names -Expected 'AdminOnlyAccessPolicy, DefaultFullAccessPolicy, DefaultFullAccessWithNotificationPolicy' `
        -Rating 'Info' -Note 'Eine vorgefertigte "AdminOnlyAccessWithNotifyPolicy" existiert bei Microsoft nicht.'
}

$globalQuar = Invoke-SafeCommand 'Get-QuarantinePolicy (Global)' { Get-QuarantinePolicy -QuarantinePolicyType GlobalQuarantinePolicy }
$script:RawData['GlobalQuarantinePolicy'] = $globalQuar
if ($null -ne $globalQuar) {
    $freq = Get-SafeProperty $globalQuar 'EndUserSpamNotificationFrequency'
    Test-Value -Id 'QUA-03' -Category 'Quarantaene-Policies' -Check 'Benachrichtigungsfrequenz' `
        -Actual $freq -Accept @('04:00:00') -ExpectedText '04:00:00 (alle 4 Stunden)' `
        -Note 'Gueltig sind nur 04:00:00, 1.00:00:00 und 7.00:00:00.'

    Add-Finding -Id 'QUA-04' -Category 'Quarantaene-Policies' -Check 'Branding und Absenderadresse' `
        -Actual ('Branding: ' + (Get-SafeProperty $globalQuar 'OrganizationBrandingEnabled') +
                 ', Absender: ' + (ConvertTo-DisplayValue (Get-SafeProperty $globalQuar 'EndUserSpamNotificationCustomFromAddress'))) `
        -Expected 'Branding aktiv, eigene Absenderadresse' -Rating 'Info' `
        -Note 'Gefaelschte Quarantaene-Benachrichtigungen sind eine der verbreitetsten Phishing-Kampagnen.'
}

# ===================================================================================
#  ADV  Advanced Delivery
# ===================================================================================

Write-Step 'ADV  Advanced Delivery'

$secOps = Invoke-SafeCommand 'Get-SecOpsOverridePolicy' { Get-SecOpsOverridePolicy }
$script:RawData['SecOpsOverridePolicy'] = $secOps
if ($null -eq $secOps) {
    Add-Finding -Id 'ADV-01' -Category 'Advanced Delivery' -Check 'SecOps-Postfach' `
        -Actual 'nicht konfiguriert' -Expected 'nur mit dediziertem, isoliertem Postfach' -Rating 'OK'
} else {
    Add-Finding -Id 'ADV-01' -Category 'Advanced Delivery' -Check 'SecOps-Postfach' `
        -Actual (Get-SafeProperty $secOps 'SentTo') -Expected 'dediziertes Security-Postfach' -Rating 'Info' `
        -Note 'ACHTUNG: SecOps-Postfaecher erhalten auch Malware ungefiltert - Malware-Filter und Malware-ZAP werden uebersprungen.'
}

$phishSim = Invoke-SafeCommand 'Get-PhishSimOverridePolicy' { Get-PhishSimOverridePolicy }
$script:RawData['PhishSimOverridePolicy'] = $phishSim
if ($null -eq $phishSim) {
    Add-Finding -Id 'ADV-02' -Category 'Advanced Delivery' -Check 'Phishing-Simulation (Drittanbieter)' `
        -Actual 'nicht konfiguriert' -Expected 'nur bei laufenden Simulationen' -Rating 'Info' `
        -Note 'Verfuegbar ab Defender P1 - KEIN E5-Feature.'
} else {
    Add-Finding -Id 'ADV-02' -Category 'Advanced Delivery' -Check 'Phishing-Simulation (Drittanbieter)' `
        -Actual ('konfiguriert, aktiv: ' + (Get-SafeProperty $phishSim 'Enabled')) -Expected 'ueber Advanced Delivery' -Rating 'OK'
}

# ===================================================================================
#  ADV-03 / EF  Transportregeln und Connectors
# ===================================================================================

Write-Step 'ADV-03 / EF  Transportregeln und Connectors'

$rules = Invoke-SafeCommand 'Get-TransportRule' { Get-TransportRule }
$script:RawData['TransportRule'] = $rules

if ($null -ne $rules) {
    $bypass = @($rules | Where-Object { [string]$_.SetSCL -eq '-1' })
    $cnt = 0
    if ($null -ne $bypass) { $cnt = $bypass.Count }
    $rate = 'OK'
    if ($cnt -gt 0) { $rate = 'Kritisch' }
    $val = 'keine'
    if ($cnt -gt 0) {
        $names = @()
        foreach ($r in $bypass) { $names += ($r.Name + ' [' + $r.State + ']') }
        $val = $names
    }
    Add-Finding -Id 'ADV-03' -Category 'Advanced Delivery' -Check 'Transportregeln mit Filter-Bypass (SCL -1)' `
        -Actual $val -Expected 'keine, bzw. dokumentiert und befristet' -Rating $rate `
        -Note 'Haeufigster Grund, warum ein gehaerteter Tenant trotzdem Phishing durchlaesst.'

    $banner = @($rules | Where-Object { $null -ne $_.ApplyHtmlDisclaimerText -and $_.ApplyHtmlDisclaimerText -ne '' })
    foreach ($r in $banner) {
        $fb = [string]$r.ApplyHtmlDisclaimerFallbackAction
        $rate = 'OK'
        if ($fb -eq 'Wrap') { $rate = 'Kritisch' }
        Add-Finding -Id 'EXT-03' -Category 'Kennzeichnung externer E-Mails' `
            -Check 'Banner-Transportregel: Fallback-Aktion' -Scope $r.Name `
            -Actual ('FallbackAction=' + $fb + ', Mode=' + $r.Mode + ', State=' + $r.State) `
            -Expected 'Ignore (niemals Wrap)' -Rating $rate `
            -Note 'Microsoft: "Wrap" stoert die Safe-Attachments-Pruefung bei Nachrichten externer Absender. Wrap ist der Default.'
    }
}

$inbound = Invoke-SafeCommand 'Get-InboundConnector' { Get-InboundConnector }
$outbound = Invoke-SafeCommand 'Get-OutboundConnector' { Get-OutboundConnector }
$script:RawData['InboundConnector']  = $inbound
$script:RawData['OutboundConnector'] = $outbound

if ($null -ne $inbound) {
    foreach ($c in $inbound) {
        $ef = Get-SafeProperty $c 'EFSkipLastIP'
        $efIps = @(Get-SafeProperty $c 'EFSkipIPs')
        $efCnt = 0
        if ($null -ne $efIps) { $efCnt = $efIps.Count }
        $efActive = ($ef -eq $true -or $efCnt -gt 0)
        Add-Finding -Id 'EF-01' -Category 'Enhanced Filtering' -Check 'Enhanced Filtering am Inbound-Connector' `
            -Scope $c.Name -Actual ('EFSkipLastIP=' + $ef + ', EFSkipIPs=' + $efCnt + ', aktiv=' + $efActive) `
            -Expected 'aktiv, WENN ein MTA vor EOP steht - sonst aus' -Rating 'Info' `
            -Note 'Pruefkriterium: Zeigt der MX-Record direkt auf Exchange Online? Wenn nein, ist Enhanced Filtering zwingend.'

        $tmi = Get-SafeProperty $c 'TreatMessagesAsInternal'
        if ($tmi -eq $true) {
            Add-Finding -Id 'HYB-09' -Category 'Hybrid & On-Premises' -Check 'TreatMessagesAsInternal' `
                -Scope $c.Name -Actual 'True' -Expected 'False, sofern nicht begruendet' -Rating 'Kritisch' `
                -Note 'Hebt die Unterscheidung intern/extern auf. External-Tag, First Contact Safety Tip und TABL-Blocks greifen dann nicht.'
        }
    }
}

if ($null -ne $outbound) {
    foreach ($c in $outbound) {
        Add-Finding -Id 'HYB-08' -Category 'Hybrid & On-Premises' -Check 'Outbound-Connector TLS' `
            -Scope $c.Name -Actual ('TlsSettings=' + (ConvertTo-DisplayValue (Get-SafeProperty $c 'TlsSettings')) +
                                    ', TlsDomain=' + (ConvertTo-DisplayValue (Get-SafeProperty $c 'TlsDomain')) +
                                    ', SmtpDaneMode=' + (ConvertTo-DisplayValue (Get-SafeProperty $c 'SmtpDaneMode'))) `
            -Expected 'DomainValidation fuer definierte Partner' -Rating 'Info'
    }
}

# ===================================================================================
#  AUTH  E-Mail-Authentifizierung
# ===================================================================================

Write-Step 'AUTH  E-Mail-Authentifizierung'

$dkim = Invoke-SafeCommand 'Get-DkimSigningConfig' { Get-DkimSigningConfig }
$script:RawData['DkimSigningConfig'] = $dkim

if ($null -ne $dkim) {
    foreach ($d in $dkim) {
        $enabled = Get-SafeProperty $d 'Enabled'
        $status  = Get-SafeProperty $d 'Status'
        $isOnMs  = ([string]$d.Domain -like '*.onmicrosoft.com')

        $rate = 'Kritisch'
        if ($enabled -eq $true) { $rate = 'OK' }
        if ($isOnMs) { $rate = 'Info' }

        Add-Finding -Id 'AUTH-04' -Category 'E-Mail-Authentifizierung' -Check 'DKIM aktiv' `
            -Scope ([string]$d.Domain) -Actual ('Enabled=' + $enabled + ', Status=' + $status) `
            -Expected 'Enabled=True, Status=Valid' -Rating $rate `
            -Note 'Ohne Custom-Domain-DKIM haengt DMARC allein an SPF - und SPF bricht bei jeder Weiterleitung.'

        $k1 = Get-SafeProperty $d 'Selector1KeySize'
        $k2 = Get-SafeProperty $d 'Selector2KeySize'
        if ($null -ne $k1) {
            $rate = 'Verbesserung empfohlen'
            if ([string]$k1 -eq '2048' -and [string]$k2 -eq '2048') { $rate = 'OK' }
            Add-Finding -Id 'AUTH-05' -Category 'E-Mail-Authentifizierung' -Check 'DKIM-Schluessellaenge' `
                -Scope ([string]$d.Domain) -Actual ('Selector1: ' + $k1 + ', Selector2: ' + $k2) `
                -Expected '2048 (beide Selektoren)' -Rating $rate `
                -Note 'Eigene Empfehlung, kein MS-Zitat. Set-DkimSigningConfig hat KEIN -KeySize; Umstellung nur ueber Rotate-DkimSigningConfig, und zwar zweimal.'
        }
    }
}

$arc = Invoke-SafeCommand 'Get-ArcConfig' { Get-ArcConfig }
$script:RawData['ArcConfig'] = $arc
$arcSealers = @()
if ($null -ne $arc) { $arcSealers = @(Get-SafeProperty $arc 'ArcTrustedSealers') }
$cnt = 0
if ($null -ne $arcSealers) { $cnt = $arcSealers.Count }
$val = 'keine'
if ($cnt -gt 0) { $val = $arcSealers }
Add-Finding -Id 'AUTH-08' -Category 'E-Mail-Authentifizierung' -Check 'ARC Trusted Sealers' `
    -Actual $val -Expected 'nur tatsaechlich genutzte Dienste' -Rating 'Info' `
    -Note 'Jeder Trusted Sealer darf beliebige Authentifizierungsergebnisse bezeugen. Set-ArcConfig ersetzt die komplette Liste.'

# DNS-Pruefungen, sofern Resolve-DnsName verfuegbar ist (Windows)
$hasResolve = $null -ne (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue)
if ($hasResolve -and $null -ne $acceptedDomains) {
    foreach ($d in $acceptedDomains) {
        $dn = [string]$d.DomainName
        if ($dn -like '*.onmicrosoft.com') { continue }

        # SPF
        $spf = $null
        try {
            $txt = Resolve-DnsName -Name $dn -Type TXT -ErrorAction Stop
            foreach ($t in $txt) {
                $strings = Get-SafeProperty $t 'Strings'
                if ($null -ne $strings) {
                    $joined = ($strings -join '')
                    if ($joined -like 'v=spf1*') { $spf = $joined }
                }
            }
        } catch { }

        if ($null -eq $spf) {
            Add-Finding -Id 'AUTH-01' -Category 'E-Mail-Authentifizierung' -Check 'SPF-Record' -Scope $dn `
                -Actual 'kein SPF-Record gefunden' -Expected 'v=spf1 ... -all' -Rating 'Kritisch'
        } else {
            $rate = 'Verbesserung empfohlen'
            if ($spf -like '*-all*') { $rate = 'OK' }
            Add-Finding -Id 'AUTH-01' -Category 'E-Mail-Authentifizierung' -Check 'SPF-Record' -Scope $dn `
                -Actual $spf -Expected 'endet auf -all' -Rating $rate `
                -Note 'Vor der Umstellung auf -all pruefen, ob alle legitimen Versandquellen enthalten sind.'
        }

        # DMARC
        $dmarc = $null
        try {
            $txt = Resolve-DnsName -Name ('_dmarc.' + $dn) -Type TXT -ErrorAction Stop
            foreach ($t in $txt) {
                $strings = Get-SafeProperty $t 'Strings'
                if ($null -ne $strings) {
                    $joined = ($strings -join '')
                    if ($joined -like 'v=DMARC1*') { $dmarc = $joined }
                }
            }
        } catch { }

        if ($null -eq $dmarc) {
            Add-Finding -Id 'AUTH-06' -Category 'E-Mail-Authentifizierung' -Check 'DMARC-Record' -Scope $dn `
                -Actual 'kein DMARC-Record gefunden' -Expected 'mindestens p=none, Zielzustand p=reject' -Rating 'Kritisch' `
                -Note 'Seit Mai 2025 verlangt Microsoft DMARC von Absendern mit mehr als 5.000 Nachrichten/Tag an Consumer-Adressen.'
        } else {
            $rate = 'Verbesserung empfohlen'
            if ($dmarc -like '*p=reject*') { $rate = 'OK' }
            Add-Finding -Id 'AUTH-06' -Category 'E-Mail-Authentifizierung' -Check 'DMARC-Record' -Scope $dn `
                -Actual $dmarc -Expected 'Zielzustand p=reject' -Rating $rate
        }

        # MX
        try {
            $mx = Resolve-DnsName -Name $dn -Type MX -ErrorAction Stop |
                  Where-Object { $_.QueryType -eq 'MX' } |
                  Sort-Object Preference
            $mxNames = @()
            foreach ($m in $mx) { $mxNames += ([string]$m.NameExchange) }
            $direct = $false
            foreach ($m in $mxNames) {
                if ($m -like '*.mail.protection.outlook.com' -or $m -like '*.mx.microsoft') { $direct = $true }
            }
            Add-Finding -Id 'EF-00' -Category 'Enhanced Filtering' -Check 'MX-Ziel' -Scope $dn `
                -Actual ($mxNames -join ', ') `
                -Expected 'direkt auf Exchange Online, sonst Enhanced Filtering zwingend' -Rating 'Info' `
                -Note ('Direkt auf EXO: ' + $direct + '. Wenn False, muss Enhanced Filtering am Inbound-Connector aktiv sein, sonst sind SPF/DMARC/Spoof-Ergebnisse wertlos.')
        } catch { }
    }
}
elseif (-not $hasResolve) {
    Add-Finding -Id 'AUTH-DNS' -Category 'E-Mail-Authentifizierung' -Check 'DNS-Pruefung' `
        -Actual 'Resolve-DnsName nicht verfuegbar' -Expected 'SPF/DMARC/MX manuell pruefen' -Rating 'Info'
}

# ===================================================================================
#  EXT / HYB / OPS
# ===================================================================================

Write-Step 'EXT / HYB  Kennzeichnung und Legacy-Protokolle'

$extOutlook = Invoke-SafeCommand 'Get-ExternalInOutlook' { Get-ExternalInOutlook }
$script:RawData['ExternalInOutlook'] = $extOutlook
if ($null -ne $extOutlook) {
    Test-Value -Id 'EXT-01' -Category 'Kennzeichnung externer E-Mails' -Check 'Natives External-Tag' `
        -Actual (Get-SafeProperty $extOutlook 'Enabled') -Accept @('True') -ExpectedText 'True' `
        -Note 'Eigenes Feature, nicht identisch mit dem First Contact Safety Tip. Wirksam nach 24-48 Stunden.'
}

$transportConfig = Invoke-SafeCommand 'Get-TransportConfig' { Get-TransportConfig }
$script:RawData['TransportConfig'] = $transportConfig
if ($null -ne $transportConfig) {
    $smtpDisabled = Get-SafeProperty $transportConfig 'SmtpClientAuthenticationDisabled'
    Test-Value -Id 'HYB-05' -Category 'Hybrid & On-Premises' -Check 'SMTP AUTH tenantweit deaktiviert' `
        -Actual $smtpDisabled -Accept @('True') -ExpectedText 'True' `
        -Note 'Ausnahmen einzeln per Set-CASMailbox, nicht tenantweit. Besser: interner Relay-Connector.'
}

$casPlans = Invoke-SafeCommand 'Get-CASMailboxPlan' { Get-CASMailboxPlan }
$script:RawData['CASMailboxPlan'] = $casPlans
if ($null -ne $casPlans) {
    $popOn = @($casPlans | Where-Object { $_.PopEnabled -eq $true })
    $imapOn = @($casPlans | Where-Object { $_.ImapEnabled -eq $true })
    $cp = 0
    if ($null -ne $popOn) { $cp = $popOn.Count }
    $ci = 0
    if ($null -ne $imapOn) { $ci = $imapOn.Count }
    $rate = 'OK'
    if ($cp -gt 0 -or $ci -gt 0) { $rate = 'Verbesserung empfohlen' }
    Add-Finding -Id 'HYB-06' -Category 'Hybrid & On-Premises' -Check 'POP3 / IMAP4 in den Mailbox-Plaenen' `
        -Actual ('Plaene mit POP: ' + $cp + ', mit IMAP: ' + $ci) -Expected 'beide deaktiviert' -Rating $rate `
        -Note 'Der Mailbox-Plan wirkt nur auf NEUE Postfaecher. Bestandspostfaecher separat pruefen.'
}

if ($IncludeProtectionAlerts) {
    Write-Step 'REP  Alert Policies'
    $alerts = Invoke-SafeCommand 'Get-ProtectionAlert' { Get-ProtectionAlert }
    $script:RawData['ProtectionAlert'] = $alerts
    if ($null -ne $alerts) {
        $relevant = @('User restricted from sending email','Suspicious email sending patterns detected','Suspicious connector activity')
        foreach ($name in $relevant) {
            $a = $alerts | Where-Object { $_.Name -eq $name } | Select-Object -First 1
            if ($null -eq $a) {
                Add-Finding -Id 'REP-02' -Category 'Reporting & Alerting' -Check ('Alert Policy: ' + $name) `
                    -Actual 'nicht gefunden' -Expected 'vorhanden und aktiv' -Rating 'Info'
            } else {
                $disabled = Get-SafeProperty $a 'Disabled'
                $rate = 'OK'
                if ($disabled -eq $true) { $rate = 'Kritisch' }
                Add-Finding -Id 'REP-02' -Category 'Reporting & Alerting' -Check ('Alert Policy: ' + $name) `
                    -Actual ('Disabled=' + $disabled + ', Empfaenger: ' + (ConvertTo-DisplayValue (Get-SafeProperty $a 'NotifyUser'))) `
                    -Expected 'aktiv, Empfaenger auf dem Security-Verteiler' -Rating $rate
            }
        }
    }
}

# --- Remote Domain (zweiter Weg fuer automatische Weiterleitungen) ---
$remoteDomains = Invoke-SafeCommand 'Get-RemoteDomain' { Get-RemoteDomain }
$script:RawData['RemoteDomain'] = $remoteDomains
if ($null -ne $remoteDomains) {
    $def = $remoteDomains | Where-Object { $_.Name -eq 'Default' } | Select-Object -First 1
    if ($null -ne $def) {
        Test-Value -Id 'ASO-07' -Category 'Anti-Spam outbound' -Check 'Remote Domain: automatische Weiterleitung' `
            -Scope 'Default' -Actual (Get-SafeProperty $def 'AutoForwardEnabled') -Accept @('False') `
            -ExpectedText 'False' -FailRating 'Kritisch' `
            -Note 'Zweiter, unabhaengiger Weg neben AutoForwardingMode (ASO-04). Default in Exchange Online ist True. In Hybrid die Remote Domain fuer *.mail.onmicrosoft.com nicht umstellen.'
    }
}

# --- Accepted-Domain-Typ / DBEB ---
if ($null -ne $acceptedDomains) {
    foreach ($d in $acceptedDomains) {
        $dt = [string](Get-SafeProperty $d 'DomainType')
        $rate = 'OK'
        if ($dt -eq 'InternalRelay') { $rate = 'Verbesserung empfohlen' }
        Add-Finding -Id 'HYB-10' -Category 'Hybrid & On-Premises' -Check 'Accepted-Domain-Typ (DBEB)' `
            -Scope ([string]$d.DomainName) -Actual $dt -Expected 'Authoritative, sofern alle Empfaenger in EXO liegen' `
            -Rating $rate `
            -Note 'InternalRelay nimmt Post an unbekannte Empfaenger an und verschenkt damit Directory Based Edge Blocking.'
    }
}

# --- onmicrosoft.com als Primaeradresse ---
$onmsMailboxes = Invoke-SafeCommand 'Get-Mailbox (onmicrosoft)' {
    Get-Mailbox -ResultSize Unlimited | Where-Object { [string]$_.PrimarySmtpAddress -like '*.onmicrosoft.com' }
}
if ($null -ne $onmsMailboxes) {
    $cnt = @($onmsMailboxes).Count
    $rate = 'OK'
    if ($cnt -gt 0) { $rate = 'Verbesserung empfohlen' }
    Add-Finding -Id 'ASO-08' -Category 'Anti-Spam outbound' -Check 'Postfaecher mit onmicrosoft.com als Primaeradresse' `
        -Actual ($cnt.ToString() + ' Postfaecher') -Expected '0 - alle auf eine Custom-Domaene umstellen' -Rating $rate `
        -Note 'Microsoft drosselt den Versand aus der Default-onmicrosoft.com-Domaene. Symptom: NDR 550 5.7.236.'
}

# --- Teams Protection ---
$teamsPolicy = Invoke-SafeCommand 'Get-TeamsProtectionPolicy' { Get-TeamsProtectionPolicy }
$script:RawData['TeamsProtectionPolicy'] = $teamsPolicy
if ($null -eq $teamsPolicy) {
    Add-Finding -Id 'TEAMS-01' -Category 'Microsoft Teams Protection' -Check 'Teams Protection verfuegbar' `
        -Actual 'nicht verfuegbar' -Expected 'Defender for Office 365' -Rating 'n.a.'
} else {
    Test-Value -Id 'TEAMS-01' -Category 'Microsoft Teams Protection' -Check 'ZAP fuer Teams-Nachrichten' `
        -Scope ([string]$teamsPolicy.Identity) -Actual (Get-SafeProperty $teamsPolicy 'ZapEnabled') -Accept @('True') `
        -ExpectedText 'True' `
        -Note 'Default laut Cmdlet-Referenz ist False. Nicht zu verwechseln mit Safe Links fuer Teams (SL-01).'

    Add-Finding -Id 'TEAMS-01b' -Category 'Microsoft Teams Protection' -Check 'Quarantaene-Policies fuer Teams' `
        -Scope ([string]$teamsPolicy.Identity) `
        -Actual ('HC-Phish: ' + (Get-SafeProperty $teamsPolicy 'HighConfidencePhishQuarantineTag') +
                 ', Malware: ' + (Get-SafeProperty $teamsPolicy 'MalwareQuarantineTag')) `
        -Expected 'beide AdminOnlyAccessPolicy' -Rating 'Info'
}

$reportSub = Invoke-SafeCommand 'Get-ReportSubmissionPolicy' { Get-ReportSubmissionPolicy }
$script:RawData['ReportSubmissionPolicy'] = $reportSub
if ($null -ne $reportSub) {
    Add-Finding -Id 'REP-03' -Category 'Reporting & Alerting' -Check 'Meldeweg fuer Nutzer' `
        -Actual ('An Microsoft: ' + (Get-SafeProperty $reportSub 'EnableReportToMicrosoft') +
                 ', eigene Adresse: ' + (Get-SafeProperty $reportSub 'ReportJunkToCustomizedAddress')) `
        -Expected 'mindestens einer der beiden Wege aktiv' -Rating 'Info'
}

# ===================================================================================
#  Ausgabe
# ===================================================================================

Write-Step 'Ausgabe'

if (-not (Test-Path -LiteralPath $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
}

$safeCustomer = ($CustomerName -replace '[^\w\-]', '_')
$stamp = Get-Date -Format 'yyyyMMdd-HHmm'
$baseName = 'EOP-Audit_' + $safeCustomer + '_' + $stamp

$csvPath  = Join-Path $OutputFolder ($baseName + '.csv')
$htmlPath = Join-Path $OutputFolder ($baseName + '.html')
$jsonPath = Join-Path $OutputFolder ($baseName + '_RawConfig.json')

$csvRows = foreach ($f in $script:Findings) {
    $li = Get-LinkInfo $f.ID
    $f | Select-Object ID, Kategorie, Pruefpunkt, Objekt, IstWert, SollWert, Bewertung, Hinweis,
        @{ n = 'Fundort';    e = { if ($li) { $li.Nav }    else { '' } } },
        @{ n = 'PortalLink'; e = { if ($li) { $li.Portal } else { '' } } },
        @{ n = 'LearnLink';  e = { if ($li) { $li.Learn }  else { '' } } }
}
$csvRows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
Write-Info ('CSV : ' + $csvPath)

# Kennzahlen
$total    = $script:Findings.Count
$critical = @($script:Findings | Where-Object { $_.Bewertung -eq 'Kritisch' }).Count
$improve  = @($script:Findings | Where-Object { $_.Bewertung -eq 'Verbesserung empfohlen' }).Count
$ok       = @($script:Findings | Where-Object { $_.Bewertung -eq 'OK' }).Count
$na       = @($script:Findings | Where-Object { $_.Bewertung -eq 'n.a.' }).Count
$info     = @($script:Findings | Where-Object { $_.Bewertung -eq 'Info' }).Count

$style = @'
<style>
 body   { font-family: Arial, Helvetica, sans-serif; font-size: 13px; color:#222; margin:24px; }
 h1     { color:#1F3864; font-size:22px; margin-bottom:2px; }
 h2     { color:#1F3864; font-size:16px; margin-top:26px; border-bottom:1px solid #D0D0D0; padding-bottom:4px; }
 .sub   { color:#666; font-size:12px; margin-bottom:18px; }
 table  { border-collapse: collapse; width:100%; margin-top:8px; }
 table  { table-layout: auto; }
 th     { background:#1F3864; color:#fff; text-align:left; padding:6px 8px; font-size:12px; }
 td     { border-bottom:1px solid #E0E0E0; padding:5px 8px; vertical-align:top; font-size:12px; }
 tr:nth-child(even) td { background:#FAFAFA; }
 .kpi   { display:inline-block; padding:10px 16px; margin:4px 8px 4px 0; border-radius:4px; color:#fff; font-weight:bold; }
 .k-crit{ background:#C00000; } .k-imp { background:#BF8F00; } .k-ok { background:#548235; }
 .k-na  { background:#808080; } .k-inf { background:#2E5496; }
 .b-Kritisch { color:#C00000; font-weight:bold; }
 .b-Verbesserungempfohlen { color:#BF8F00; font-weight:bold; }
 .b-OK { color:#548235; }
 .b-na { color:#808080; }
 .b-Info { color:#2E5496; }
 .note  { color:#555; font-size:11px; }
 .find  { color:#444; font-size:11px; margin-top:5px; padding:4px 6px; background:#F1F4F9;
          border-left:3px solid #2E5496; }
 .find b { color:#1F3864; }
 .find a { color:#0563C1; text-decoration:none; font-weight:bold; white-space:nowrap; }
 .find a:hover { text-decoration:underline; }
 .legend td { font-size:11px; }
 .legend th { font-size:11px; }
 code   { font-family: Consolas, monospace; font-size:11px; }
 .foot  { margin-top:30px; color:#777; font-size:11px; border-top:1px solid #DDD; padding-top:10px; }
</style>
'@

$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine('<!DOCTYPE html><html lang="de"><head><meta charset="utf-8">')
[void]$sb.AppendLine('<title>EOP/MDO Audit - ' + $CustomerName + '</title>')
[void]$sb.AppendLine($style)
[void]$sb.AppendLine('</head><body>')
[void]$sb.AppendLine('<h1>EOP / Microsoft Defender for Office 365 &ndash; Bestandsaufnahme</h1>')
[void]$sb.AppendLine('<div class="sub">Kunde: <b>' + $CustomerName + '</b> &middot; erstellt am ' +
                     (Get-Date -Format 'dd.MM.yyyy HH:mm') + ' &middot; ' + $total + ' Pr&uuml;fpunkte</div>')

[void]$sb.AppendLine('<div>')
[void]$sb.AppendLine('<span class="kpi k-crit">Kritisch: ' + $critical + '</span>')
[void]$sb.AppendLine('<span class="kpi k-imp">Verbesserung: ' + $improve + '</span>')
[void]$sb.AppendLine('<span class="kpi k-ok">OK: ' + $ok + '</span>')
[void]$sb.AppendLine('<span class="kpi k-na">n.a.: ' + $na + '</span>')
[void]$sb.AppendLine('<span class="kpi k-inf">Info: ' + $info + '</span>')
[void]$sb.AppendLine('</div>')

function Add-HtmlTable {
    param($Builder, $Rows, [string]$Heading)
    if ($null -eq $Rows -or @($Rows).Count -eq 0) { return }
    [void]$Builder.AppendLine('<h2>' + $Heading + '</h2>')
    [void]$Builder.AppendLine('<table><tr><th style="width:85px">ID</th><th style="width:340px">Pr&uuml;fpunkt und Fundort</th>' +
                              '<th style="width:130px">Objekt</th><th>Ist-Wert</th><th style="width:170px">Soll</th>' +
                              '<th style="width:110px">Bewertung</th></tr>')
    foreach ($r in $Rows) {
        $cls = 'b-' + ($r.Bewertung -replace '[^A-Za-z]','')
        $note = ''
        if (-not [string]::IsNullOrWhiteSpace($r.Hinweis)) {
            $note = '<div class="note">' + (ConvertTo-HtmlText $r.Hinweis) + '</div>'
        }

        # Fundort: wo im Portal die Einstellung sitzt, plus Direktlinks.
        $find = ''
        $li = Get-LinkInfo $r.ID
        if ($null -ne $li) {
            $find = '<div class="find"><b>Fundort:</b> ' + (ConvertTo-HtmlText $li.Nav)
            if (-not [string]::IsNullOrWhiteSpace($li.Portal)) {
                $find += ' &nbsp;<a href="' + $li.Portal + '" target="_blank">Portal &#8599;</a>'
            }
            if (-not [string]::IsNullOrWhiteSpace($li.Learn)) {
                $find += ' &nbsp;<a href="' + $li.Learn + '" target="_blank">Learn &#8599;</a>'
            }
            $find += '</div>'
        }

        [void]$Builder.AppendLine('<tr><td style="white-space:nowrap"><b>' + (ConvertTo-HtmlText $r.ID) + '</b></td>' +
            '<td>' + (ConvertTo-HtmlText $r.Pruefpunkt) + $note + $find + '</td>' +
            '<td><code>' + (ConvertTo-HtmlText $r.Objekt) + '</code></td>' +
            '<td><code>' + (ConvertTo-HtmlText $r.IstWert) + '</code></td>' +
            '<td>' + (ConvertTo-HtmlText $r.SollWert) + '</td>' +
            '<td class="' + $cls + '">' + $r.Bewertung + '</td></tr>')
    }
    [void]$Builder.AppendLine('</table>')
}

$critRows = @($script:Findings | Where-Object { $_.Bewertung -eq 'Kritisch' })
Add-HtmlTable -Builder $sb -Rows $critRows -Heading ('Sofortmassnahmen (' + $critical + ')')

$impRows = @($script:Findings | Where-Object { $_.Bewertung -eq 'Verbesserung empfohlen' })
Add-HtmlTable -Builder $sb -Rows $impRows -Heading ('Verbesserung empfohlen (' + $improve + ')')

foreach ($cat in ($script:Findings | Select-Object -ExpandProperty Kategorie -Unique)) {
    $rows = @($script:Findings | Where-Object { $_.Kategorie -eq $cat })
    Add-HtmlTable -Builder $sb -Rows $rows -Heading ('Vollstaendig: ' + $cat)
}

[void]$sb.AppendLine('<h2>Was die K&uuml;rzel bedeuten</h2>')
[void]$sb.AppendLine('<p style="font-size:12px;color:#444;margin:4px 0 0 0">Das Pr&auml;fix sagt, in welchem Bereich ' +
    'ein Pr&uuml;fpunkt sitzt &ndash; nicht zwingend, in welcher Portal-Richtlinie. APH und IMP stehen zum Beispiel ' +
    'beide in der Anti-Phishing-Richtlinie, sind aber lizenzrechtlich getrennt.</p>')
[void]$sb.AppendLine('<table class="legend"><tr><th style="width:70px">K&uuml;rzel</th>' +
    '<th style="width:260px">Bereich</th><th>Was dahintersteckt</th></tr>')
foreach ($p in $script:Prefixes) {
    [void]$sb.AppendLine('<tr><td><b>' + $p.Key + '</b></td><td>' + (ConvertTo-HtmlText $p.Area) + '</td>' +
        '<td>' + (ConvertTo-HtmlText $p.Desc) + '</td></tr>')
}
[void]$sb.AppendLine('</table>')
[void]$sb.AppendLine('<p style="font-size:11px;color:#666;margin-top:10px">Der Pfad &bdquo;Bedrohungsrichtlinien&ldquo; ' +
    'liegt im Defender-Portal unter <i>E-Mail &amp; Zusammenarbeit &gt; Richtlinien &amp; Regeln</i>. ' +
    'S&auml;mtliche Defender-Deep-Links stammen aus der Microsoft-Dokumentation. Nicht von Microsoft dokumentiert ' +
    'und aus der Praxis &uuml;bernommen sind die Exchange-Admin-Center-Links von ' +
    ($script:UndocumentedLinks -join ', ') + ' &ndash; f&uuml;hrt einer ins Leere, hilft der Klickpfad daneben.</p>')

[void]$sb.AppendLine('<div class="foot">Erzeugt von Invoke-EopAudit.ps1. Rein lesende Bestandsaufnahme &ndash; ' +
    'am Tenant wurde nichts ver&auml;ndert. Die IDs entsprechen der Assessment-Checkliste und dem Best Practice Guide. ' +
    'Alle Soll-Werte sind gegen die Microsoft-Learn-Referenz (Stand August 2026) verifiziert.</div>')
[void]$sb.AppendLine('</body></html>')

Set-Content -Path $htmlPath -Value $sb.ToString() -Encoding UTF8
Write-Info ('HTML: ' + $htmlPath)

if ($ExportJson) {
    try {
        $script:RawData | ConvertTo-Json -Depth 6 -Compress:$false | Set-Content -Path $jsonPath -Encoding UTF8
        Write-Info ('JSON: ' + $jsonPath)
    }
    catch {
        Write-Warning ('JSON-Export fehlgeschlagen: ' + $_.Exception.Message)
    }
}

Write-Host ''
Write-Host '-------------------------------------------------------------' -ForegroundColor White
Write-Host (' Kritisch               : ' + $critical) -ForegroundColor Red
Write-Host (' Verbesserung empfohlen : ' + $improve)  -ForegroundColor Yellow
Write-Host (' OK                     : ' + $ok)       -ForegroundColor Green
Write-Host (' n.a. / Info            : ' + ($na + $info)) -ForegroundColor Gray
Write-Host '-------------------------------------------------------------' -ForegroundColor White
Write-Host ''
Write-Host ' Naechster Schritt: CSV in die Assessment-Checkliste uebernehmen,' -ForegroundColor White
Write-Host ' danach Invoke-EopHardening.ps1 OHNE -Execute zur Vorschau ausfuehren.' -ForegroundColor White
Write-Host ''

# Ergebnis auch in die Pipeline geben
$script:Findings
