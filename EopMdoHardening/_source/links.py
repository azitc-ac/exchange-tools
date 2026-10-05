# -*- coding: utf-8 -*-
"""
Fundort je Pruefpunkt: Klickpfad, Portal-Deep-Link, Microsoft-Learn-Artikel.

Alle Angaben wurden gegen learn.microsoft.com und die Quell-Repositories
MicrosoftDocs/defender-docs bzw. office-docs-powershell verifiziert.
Deep-Links, die Microsoft nicht selbst dokumentiert, sind unten in
UNDOCUMENTED aufgefuehrt und werden im LIESMICH offengelegt.
"""

# --- wiederkehrende Ziele -----------------------------------------------------
P_SPAM   = 'https://security.microsoft.com/antispam'
P_PHISH  = 'https://security.microsoft.com/antiphishing'
P_MAL    = 'https://security.microsoft.com/antimalwarev2'
P_SA     = 'https://security.microsoft.com/safeattachmentv2'
P_SL     = 'https://security.microsoft.com/safelinksv2'
P_QUA    = 'https://security.microsoft.com/quarantinePolicies'
P_TABL   = 'https://security.microsoft.com/tenantAllowBlockList'
P_ADV    = 'https://security.microsoft.com/advanceddelivery'
P_PRESET = 'https://security.microsoft.com/presetSecurityPolicies'
P_CFGAN  = 'https://security.microsoft.com/configurationAnalyzer'
P_AUTHN  = 'https://security.microsoft.com/authentication'
P_EAC    = 'https://admin.exchange.microsoft.com'
P_RULES  = P_EAC + '/#/transportrules'
P_CONN   = P_EAC + '/#/connectors'

L_SPAMABOUT = 'https://learn.microsoft.com/defender-office-365/anti-spam-protection-about'
L_ASF       = 'https://learn.microsoft.com/defender-office-365/anti-spam-policies-asf-settings-about'
L_PHISHABT  = 'https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about'
L_MALABOUT  = 'https://learn.microsoft.com/defender-office-365/anti-malware-protection-about'
L_SAABOUT   = 'https://learn.microsoft.com/defender-office-365/safe-attachments-about'
L_SLABOUT   = 'https://learn.microsoft.com/defender-office-365/safe-links-about'
L_QUAPOL    = 'https://learn.microsoft.com/defender-office-365/quarantine-policies'
L_RECO      = 'https://learn.microsoft.com/defender-office-365/recommended-settings-for-eop-and-office365'

# Klickpfad-Praefixe. "Bedrohungsrichtlinien" liegt im Defender-Portal unter
# E-Mail & Zusammenarbeit > Richtlinien & Regeln.
TP  = 'Defender-Portal > Bedrohungsrichtlinien > '
SET = 'Defender-Portal > Einstellungen > E-Mail & Zusammenarbeit > '

# --- Fundorte -----------------------------------------------------------------
# id: (nav, portal, learn)
LINKS = {

# ============================================================ Preset / Governance
'PRE-01': (TP + 'Voreingestellte Sicherheitsrichtlinien (Preset security policies)',
    P_PRESET,
    'https://learn.microsoft.com/defender-office-365/mdo-deployment-guide#determine-your-threat-policy-strategy'),
'PRE-02': (TP + 'Voreingestellte Sicherheitsrichtlinien > Integrierter Schutz (Built-in protection) > "Ausschluesse verwalten"',
    P_PRESET,
    'https://learn.microsoft.com/defender-office-365/preset-security-policies#use-the-microsoft-defender-portal-to-add-exclusions-to-the-built-in-protection-preset-security-policy'),
'PRE-03': ('Kein eigener Portal-Screen. Praezedenz ablesen an der Spalte "Prioritaet" je Richtlinientyp (Antispam, Antiphishing, ...) plus der Preset-Seite',
    P_PRESET,
    'https://learn.microsoft.com/defender-office-365/preset-security-policies#order-of-precedence-for-preset-security-policies-and-other-threat-policies'),
'PRE-04': (TP + 'Konfigurationsanalyse (Configuration analyzer) > Reiter "Standardempfehlungen" / "Strenge Empfehlungen"',
    P_CFGAN,
    'https://learn.microsoft.com/defender-office-365/configuration-analyzer-for-security-policies#standard-recommendations-and-strict-recommendations-tabs-in-the-configuration-analyzer'),
'PRE-05': (TP + 'Antispam (Spalten Status / Prioritaet / Typ), analog Antiphishing, Antischadsoftware, Sichere Links, Sichere Anlagen',
    P_SPAM,
    'https://learn.microsoft.com/defender-office-365/anti-spam-policies-configure#use-the-microsoft-defender-portal-to-enable-or-disable-anti-spam-policies'),

# ============================================================ Anti-Spam inbound
'ASI-01': (TP + 'Antispam > <Richtlinie> > "Massen-E-Mail-Schwellenwert & Spameigenschaften" > Schieberegler "Massen-E-Mail-Schwellenwert"',
    P_SPAM, L_SPAMABOUT + '#bulk-complaint-threshold-bcl-in-anti-spam-policies'),
'ASI-02': (TP + 'Antispam > <Richtlinie> > "Aktionen" > Nachrichtenaktionen > Zeile "Spam"',
    P_SPAM, L_SPAMABOUT + '#actions-in-anti-spam-policies'),
'ASI-03': (TP + 'Antispam > <Richtlinie> > "Aktionen" > Nachrichtenaktionen > Zeile "Spam mit hoher Sicherheit"',
    P_SPAM, L_SPAMABOUT + '#actions-in-anti-spam-policies'),
'ASI-04': (TP + 'Antispam > <Richtlinie> > "Aktionen" > Nachrichtenaktionen > Zeile "Phishing"',
    P_SPAM, L_SPAMABOUT + '#actions-in-anti-spam-policies'),
'ASI-05': (TP + 'Antispam > <Richtlinie> > "Aktionen" > Nachrichtenaktionen > Zeile "Phishing mit hoher Sicherheit"',
    P_SPAM, L_SPAMABOUT + '#actions-in-anti-spam-policies'),
'ASI-06': (TP + 'Antispam > <Richtlinie> > "Aktionen" > Nachrichtenaktionen > Zeile "Massenkonforme Ebene (BCL) erreicht oder ueberschritten"',
    P_SPAM, L_SPAMABOUT + '#actions-in-anti-spam-policies'),
'ASI-07': (TP + 'Antispam > <Richtlinie> > "Aktionen" > Abschnitt "Sicherheitstipps" > Haken "Sicherheitstipps aktivieren"',
    P_SPAM, L_RECO + '#anti-spam-policy-settings'),
'ASI-08': (TP + 'Antispam > <Richtlinie> > "Aktionen" > "Automatische Bereinigung (ZAP) aktivieren" mit den zwei Unterhaken fuer Phishing und Spam',
    P_SPAM, 'https://learn.microsoft.com/defender-office-365/zero-hour-auto-purge#zero-hour-auto-purge-zap-for-spam'),
'ASI-09': (TP + 'Antispam > <Richtlinie> > "Aktionen" > Feld "Spamnachrichten so viele Tage lang in Quarantaene aufbewahren" (1-30, Default 15)',
    P_SPAM, 'https://learn.microsoft.com/defender-office-365/quarantine-about#quarantine-retention'),
'ASI-10': (TP + 'Antispam > <Richtlinie> > "Zulassen- und Sperrliste" > "Absender verwalten" / "Domaenen zulassen" (im Portal max. 30 Eintraege)',
    P_SPAM, L_SPAMABOUT + '#allow-and-block-lists-in-anti-spam-policies'),
'ASI-11': (TP + 'Antispam > <Richtlinie> > "Massen-E-Mail-Schwellenwert & Spameigenschaften" > Spameigenschaften > "Enthaelt bestimmte Sprachen" / "Aus diesen Laendern" (gehoert NICHT zum ASF)',
    P_SPAM, L_SPAMABOUT + '#spam-properties-in-anti-spam-policies'),
'ASI-12': (TP + 'Antispam > <Richtlinie> > "Aktionen" > "Organisationsinterne Nachrichten, fuer die Massnahmen ergriffen werden sollen"',
    P_SPAM, L_SPAMABOUT + '#actions-in-anti-spam-policies'),

# ============================================================ ASF
'ASF-01': (TP + 'Antispam > <Richtlinie> > "Massen-E-Mail-Schwellenwert & Spameigenschaften" > Spameigenschaften > Gruppe "Erhoehen der Spambewertung"',
    P_SPAM, L_ASF + '#increase-spam-score-settings'),
'ASF-02': (TP + 'Antispam > <Richtlinie> > "Massen-E-Mail-Schwellenwert & Spameigenschaften" > Spameigenschaften > Gruppe "Als Spam markieren"',
    P_SPAM, L_ASF + '#mark-as-spam-settings'),
'ASF-03': (TP + 'Antispam > <Richtlinie> > "Massen-E-Mail-Schwellenwert & Spameigenschaften" > Spameigenschaften > "SPF-Eintrag: Hard Fail"',
    P_SPAM, L_ASF + '#mark-as-spam-settings'),
'ASF-04': (TP + 'Antispam > <Richtlinie> > "Massen-E-Mail-Schwellenwert & Spameigenschaften" > Spameigenschaften > "Absender-ID-Filterung: Hard Fail" und "Backscatter"',
    P_SPAM, L_ASF + '#mark-as-spam-settings'),
'ASF-05': (TP + 'Antispam > <Richtlinie> > "Massen-E-Mail-Schwellenwert & Spameigenschaften" > Spameigenschaften > "Testmodus" (gilt global fuer alle auf Test gesetzten ASF-Optionen)',
    P_SPAM, L_ASF + '#enable-disable-or-test-asf-settings'),

# ============================================================ Connection Filter
'CF-01': (TP + 'Antispam > in der Liste die Zeile "Verbindungsfilterrichtlinie (Standard)" ANKLICKEN > Flyout > "Verbindungsfilterrichtlinie bearbeiten" > "Nachrichten aus den folgenden IP-Adressen immer zulassen"',
    P_SPAM,
    'https://learn.microsoft.com/defender-office-365/connection-filter-policies-configure#use-the-microsoft-defender-portal-to-modify-the-default-connection-filter-policy'),
'CF-02': (TP + 'Antispam > Zeile "Verbindungsfilterrichtlinie (Standard)" > "Verbindungsfilterrichtlinie bearbeiten" > "Nachrichten aus den folgenden IP-Adressen immer blockieren"',
    P_SPAM,
    'https://learn.microsoft.com/defender-office-365/connection-filter-policies-configure#use-the-microsoft-defender-portal-to-modify-the-default-connection-filter-policy'),
'CF-03': (TP + 'Antispam > Zeile "Verbindungsfilterrichtlinie (Standard)" > "Verbindungsfilterrichtlinie bearbeiten" > Haken "Sichere Liste aktivieren"',
    P_SPAM,
    'https://learn.microsoft.com/defender-office-365/connection-filter-policies-configure#use-the-microsoft-defender-portal-to-modify-the-default-connection-filter-policy'),

# ============================================================ Anti-Spam outbound
'ASO-01': (TP + 'Antispam > Liste auf "Ausgehende Antispamrichtlinie (Standard)" und eigene Richtlinien pruefen; neu ueber "+ Richtlinie erstellen > Ausgehend"',
    P_SPAM,
    'https://learn.microsoft.com/defender-office-365/outbound-spam-policies-configure#use-the-microsoft-defender-portal-to-create-outbound-spam-policies'),
'ASO-02': (TP + 'Antispam > <ausgehende Richtlinie> > "Schutzeinstellungen" > Nachrichtengrenzwerte (extern / intern / taeglich)',
    P_SPAM,
    'https://learn.microsoft.com/defender-office-365/outbound-spam-policies-configure#use-the-microsoft-defender-portal-to-create-outbound-spam-policies'),
'ASO-03': (TP + 'Antispam > <ausgehende Richtlinie> > "Schutzeinstellungen" > "Einschraenkung fuer Benutzer, die das Nachrichtenlimit erreichen"',
    P_SPAM, L_RECO + '#outbound-spam-policy-settings'),
'ASO-04': (TP + 'Antispam > <ausgehende Richtlinie> > "Schutzeinstellungen" > Weiterleitungsregeln > "Automatische Weiterleitungsregeln"',
    P_SPAM,
    'https://learn.microsoft.com/defender-office-365/outbound-spam-policies-external-email-forwarding'),
'ASO-05': (TP + 'Antispam > "Ausgehende Antispamrichtlinie (Standard)" > "Schutzeinstellungen" > Benachrichtigungen (BCC-Kopie); gesperrte Absender unter Defender-Portal > Ueberpruefen > Eingeschraenkte Benutzer',
    'https://security.microsoft.com/restrictedusers',
    'https://learn.microsoft.com/defender-office-365/outbound-spam-policies-configure#use-the-microsoft-defender-portal-to-create-outbound-spam-policies'),
'ASO-06': ('Nicht konfigurierbar (mandantenweites Limit). Auswertung: Exchange Admin Center > Berichte > Nachrichtenfluss > "Tenant Outbound External Recipients"',
    P_EAC + '/#/reports/mailflowreportsmain',
    'https://learn.microsoft.com/defender-office-365/outbound-spam-sending-limits-troubleshoot#tenant-external-recipient-rate-limit'),
'ASO-07': ('NICHT im Defender-Portal: Exchange Admin Center > Nachrichtenfluss > Remotedomaenen > "Default" > Abschnitt "Automatische Antworten" (dort sitzt die Weiterleitungsoption)',
    P_EAC + '/#/remotedomains',
    'https://learn.microsoft.com/exchange/mail-flow-best-practices/remote-domains/remote-domains#reducing-or-increasing-information-flow-to-another-company'),
'ASO-08': ('Keine Einstellung, mandantenweite Drosselung. Pruefung: Exchange Admin Center > Nachrichtenfluss > Akzeptierte Domaenen sowie Absenderdomaenen in Connectors und Anwendungen',
    P_EAC + '/#/accepteddomains',
    'https://learn.microsoft.com/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits#sending-limits'),

# ============================================================ Anti-Phishing
'APH-01': (TP + 'Antiphishing > <Richtlinie> > "Phishingschwellenwert & Schutz" > Abschnitt "Spoofing" > "Spoofintelligenz aktivieren"',
    P_PHISH, L_PHISHABT + '#spoof-settings'),
'APH-02': (TP + 'Antiphishing > <Richtlinie> > "Aktionen" > "Wenn die Nachricht durch Spoofintelligenz als Spoofing erkannt wird"',
    P_PHISH, L_RECO + '#anti-phishing-policy-settings-for-all-cloud-mailboxes'),
'APH-03': (TP + 'Antiphishing > <Richtlinie> > "Phishingschwellenwert & Schutz" > "DMARC-Eintragsrichtlinie beruecksichtigen, wenn die Nachricht als Spoofing erkannt wird"',
    P_PHISH, L_PHISHABT + '#spoof-protection-and-sender-dmarc-policies'),
'APH-04': (TP + 'Antiphishing > <Richtlinie> > "Aktionen" > "... und die DMARC-Richtlinie p=quarantine lautet"',
    P_PHISH, L_PHISHABT + '#spoof-protection-and-sender-dmarc-policies'),
'APH-05': (TP + 'Antiphishing > <Richtlinie> > "Aktionen" > "... und die DMARC-Richtlinie p=reject lautet"',
    P_PHISH, L_PHISHABT + '#spoof-protection-and-sender-dmarc-policies'),
'APH-06': (TP + 'Antiphishing > <Richtlinie> > "Aktionen" > Sicherheitstipps & Indikatoren > "(?) fuer nicht authentifizierte Absender anzeigen" und "Tag via anzeigen"',
    P_PHISH, L_PHISHABT + '#unauthenticated-sender-indicators'),
'APH-07': (TP + 'Antiphishing > <Richtlinie> > "Aktionen" > Sicherheitstipps & Indikatoren > "Sicherheitstipp fuer ersten Kontakt anzeigen"',
    P_PHISH, L_PHISHABT + '#first-contact-safety-tip'),
'APH-08': (TP + 'Antiphishing > <Richtlinie> > "Phishingschwellenwert & Schutz" > "Schwellenwert fuer Phishing-E-Mails" (1-4, nur mit Defender-Lizenz sichtbar)',
    P_PHISH, L_PHISHABT + '#phishing-email-thresholds-in-anti-phishing-policies-in-microsoft-defender-for-office-365'),

# ============================================================ Impersonation (in der Antiphishing-Richtlinie)
'IMP-01': (TP + 'Antiphishing > <Richtlinie> > "Phishingschwellenwert & Schutz" > Identitaetswechsel > "Zu schuetzende Domaenen aktivieren" > "Domaenen einschliessen, die ich besitze"',
    P_PHISH, L_PHISHABT + '#domain-impersonation-protection'),
'IMP-02': (TP + 'Antiphishing > <Richtlinie> > "Phishingschwellenwert & Schutz" > Identitaetswechsel > "Benutzerdefinierte Domaenen einschliessen" (max. 50)',
    P_PHISH, L_PHISHABT + '#domain-impersonation-protection'),
'IMP-03': (TP + 'Antiphishing > <Richtlinie> > "Phishingschwellenwert & Schutz" > Identitaetswechsel > "Benutzern das Schuetzen ermoeglichen" > "Geschuetzte Benutzer verwalten" (max. 350)',
    P_PHISH, L_PHISHABT + '#user-impersonation-protection'),
'IMP-04': (TP + 'Antiphishing > <Richtlinie> > "Aktionen" > die drei Zeilen zu Benutzer-, Domaenenidentitaetswechsel und Postfachintelligenz (nicht "Keine Aktion anwenden")',
    P_PHISH, L_RECO + '#impersonation-settings-in-anti-phishing-policies-in-microsoft-defender-for-office-365'),
'IMP-05': (TP + 'Antiphishing > <Richtlinie> > "Phishingschwellenwert & Schutz" > Identitaetswechsel > "Postfachintelligenz aktivieren" und "Intelligenz fuer Identitaetswechselschutz aktivieren"',
    P_PHISH, L_PHISHABT + '#mailbox-intelligence-impersonation-protection'),
'IMP-06': (TP + 'Antiphishing > <Richtlinie> > "Aktionen" > Sicherheitstipps & Indikatoren > die drei Identitaetswechsel-Tipps',
    P_PHISH, L_PHISHABT + '#impersonation-safety-tips'),
'IMP-07': (TP + 'Antiphishing > <Richtlinie> > "Phishingschwellenwert & Schutz" > Identitaetswechsel > "Vertrauenswuerdige Absender und Domaenen hinzufuegen" (max. 1024)',
    P_PHISH, L_PHISHABT + '#trusted-senders-and-domains'),
'IMP-08': (SET + 'Schutz von Prioritaetskonten (Priority account protection); die Konten selbst werden im Microsoft 365 Admin Center markiert',
    'https://security.microsoft.com/securitysettings/priorityAccountProtection',
    'https://learn.microsoft.com/defender-office-365/priority-accounts-turn-on-priority-account-protection#review-or-turn-on-priority-account-protection-in-the-microsoft-defender-portal'),

# ============================================================ Tenant Allow/Block List
'TABL-01': (TP + 'Mandanten-Zulassungs-/Sperrlisten. Zulassungen entstehen bevorzugt ueber Uebermittlungen (Submissions), nicht von Hand',
    P_TABL,
    'https://learn.microsoft.com/defender-office-365/tenant-allow-block-list-about#allow-entries-in-the-tenant-allowblock-list'),
'TABL-02': (TP + 'Mandanten-Zulassungs-/Sperrlisten > Reiter "Domaenen und E-Mail-Adressen" > Spalten "Aktion" und "Laeuft ab am"',
    P_TABL,
    'https://learn.microsoft.com/defender-office-365/tenant-allow-block-list-email-spoof-configure#domains-and-email-addresses-in-the-tenant-allowblock-list'),
'TABL-03': (TP + 'Mandanten-Zulassungs-/Sperrlisten > Reiter "Gefaelschte Absender"; ergaenzend die Spoofintelligenz-Erkenntnisse',
    P_TABL + '?viewid=SpoofItem',
    'https://learn.microsoft.com/defender-office-365/tenant-allow-block-list-email-spoof-configure#spoofed-senders-in-the-tenant-allowblock-list'),
'TABL-04': (TP + 'Mandanten-Zulassungs-/Sperrlisten > Reiter "URLs"',
    P_TABL,
    'https://learn.microsoft.com/defender-office-365/tenant-allow-block-list-urls-configure#url-syntax-for-the-tenant-allowblock-list'),
'TABL-05': (TP + 'Mandanten-Zulassungs-/Sperrlisten > Reiter "Dateien" (nur SHA256, Zulassungen nur ueber Uebermittlungen)',
    P_TABL,
    'https://learn.microsoft.com/defender-office-365/tenant-allow-block-list-files-configure#create-block-entries-for-files'),
'TABL-06': (TP + 'Mandanten-Zulassungs-/Sperrlisten > Reiter "IP-Adressen" (ausschliesslich IPv6)',
    P_TABL,
    'https://learn.microsoft.com/defender-office-365/tenant-allow-block-list-ip-addresses-configure#create-block-entries-for-ipv6-addresses'),

# ============================================================ Anti-Malware
'AMW-01': (TP + 'Antischadsoftware > <Richtlinie> > "Schutzeinstellungen" > "Aktivieren Sie den allgemeinen Anlagenfilter" > "Dateitypen auswaehlen"',
    P_MAL, L_MALABOUT + '#common-attachments-filter-in-anti-malware-policies'),
'AMW-02': (TP + 'Antischadsoftware > <Richtlinie> > "Schutzeinstellungen" > "Automatische Null-Stunden-Bereinigung fuer Schadsoftware aktivieren"',
    P_MAL, L_MALABOUT + '#zero-hour-auto-purge-zap-in-anti-malware-policies'),
'AMW-03': (TP + 'Antischadsoftware > <Richtlinie> > "Schutzeinstellungen" > "Wenn diese Dateitypen gefunden werden" (NDR ablehnen vs. Quarantaene)',
    P_MAL, L_MALABOUT + '#common-attachments-filter-in-anti-malware-policies'),
'AMW-04': (TP + 'Antischadsoftware > <Richtlinie> > "Schutzeinstellungen" > "Quarantaenerichtlinie" (Default AdminOnlyAccessPolicy)',
    P_MAL, L_MALABOUT + '#quarantine-policies-in-anti-malware-policies'),
'AMW-05': (TP + 'Antischadsoftware > <Richtlinie> > "Schutzeinstellungen" > Benachrichtigungen > Administratorbenachrichtigungen',
    P_MAL, L_MALABOUT + '#admin-notifications-in-anti-malware-policies'),

# ============================================================ Safe Attachments
'SA-01': (TP + 'Sichere Anlagen > <Richtlinie> > "Einstellungen" > "Safe Attachments-Antwort bei unbekannter Schadsoftware" = Blockieren',
    P_SA, L_SAABOUT + '#safe-attachments-policy-settings'),
'SA-02': (TP + 'Sichere Anlagen > <Richtlinie> > "Einstellungen" > "Umleiten von Nachrichten mit erkannten Anlagen" (wirkt laut Doku nur bei Aktion "Ueberwachen")',
    P_SA, L_SAABOUT + '#safe-attachments-policy-settings'),
'SA-03': (TP + 'Sichere Anlagen > Zahnrad "Globale Einstellungen" > "Defender for Office 365 fuer SharePoint, OneDrive und Microsoft Teams aktivieren". Download-Sperre zusaetzlich nur per SharePoint-PowerShell',
    P_SA,
    'https://learn.microsoft.com/defender-office-365/safe-attachments-for-spo-odfb-teams-configure#step-1-use-the-microsoft-defender-portal-to-turn-on-safe-attachments-for-sharepoint-onedrive-and-microsoft-teams'),
'SA-04': (TP + 'Sichere Anlagen > Zahnrad "Globale Einstellungen" > "Safe Documents fuer Office-Clients aktivieren" plus Durchklicken verbieten',
    P_SA,
    'https://learn.microsoft.com/defender-office-365/safe-documents-in-e5-plus-security-about#use-the-microsoft-defender-portal-to-configure-safe-documents'),
'SA-05': (TP + 'Sichere Anlagen > <Richtlinie> > "Einstellungen" > "Safe Attachments-Antwort" = "Dynamische Uebermittlung (Vorschau von Nachrichten)"',
    P_SA, L_SAABOUT + '#dynamic-delivery-in-safe-attachments-policies'),

# ============================================================ Safe Links
'SL-01': (TP + 'Sichere Links > <Richtlinie> > "URL- & Klickschutzeinstellungen" > Abschnitte "E-Mail", "Teams" und "Office 365-Apps" je auf Ein',
    P_SL, L_SLABOUT + '#safe-links-settings-for-email-messages'),
'SL-02': (TP + 'Sichere Links > <Richtlinie> > "URL- & Klickschutzeinstellungen" > E-Mail > "Sichere Links auf E-Mail-Nachrichten anwenden, die innerhalb der Organisation gesendet werden"',
    P_SL, L_SLABOUT + '#safe-links-settings-for-email-messages'),
'SL-03': (TP + 'Sichere Links > <Richtlinie> > "URL- & Klickschutzeinstellungen" > "URL-Ueberpruefung in Echtzeit ..." mit der Unteroption "Warten, bis die URL-Ueberpruefung abgeschlossen ist"',
    P_SL, L_SLABOUT + '#safe-links-settings-for-email-messages'),
'SL-04': (TP + 'Sichere Links > <Richtlinie> > "URL- & Klickschutzeinstellungen" > Klickschutzeinstellungen > "Benutzern das Durchklicken zur urspruenglichen URL erlauben" ausschalten',
    P_SL, L_SLABOUT + '#click-protection-settings-in-safe-links-policies'),
'SL-05': (TP + 'Sichere Links > <Richtlinie> > "URL- & Klickschutzeinstellungen" > E-Mail > "URLs nicht umschreiben, Ueberpruefungen nur ueber die SafeLinks-API" muss AUS sein',
    P_SL, L_SLABOUT + '#safe-links-settings-for-email-messages'),
'SL-06': (TP + 'Sichere Links > <Richtlinie> > "URL- & Klickschutzeinstellungen" > "Die folgenden URLs in E-Mails nicht umschreiben" > "Nicht umzuschreibende URLs verwalten" (pro Richtlinie, nicht global)',
    P_SL, L_SLABOUT + '#entry-syntax-for-the-do-not-rewrite-the-following-urls-list'),
'SL-07': (TP + 'Sichere Links > <Richtlinie> > "URL- & Klickschutzeinstellungen" > "Organisationsbranding anzeigen" (Logo aus dem M365-Organisationsdesign) sowie > "Benachrichtigung" > eigener Text (max. 200 Zeichen)',
    P_SL, L_SLABOUT + '#click-protection-settings-in-safe-links-policies'),

# ============================================================ Quarantaene
'QUA-01': (TP + 'Quarantaenerichtlinie (definieren); zugewiesen wird sie je Verdict in Antispam > <Richtlinie> > "Aktionen", in Antiphishing > "Aktionen" sowie in Antischadsoftware und Sichere Anlagen',
    P_QUA, L_QUAPOL + '#assign-quarantine-policies-in-supported-policies-in-the-microsoft-defender-portal'),
'QUA-02': (TP + 'Quarantaenerichtlinie > "Benutzerdefinierte Richtlinie hinzufuegen" > "Zugriff auf Empfaengernachrichten" > "Spezifischen Zugriff festlegen (Erweitert)" > Quarantaenebenachrichtigung aktivieren',
    P_QUA, L_QUAPOL + '#step-1-create-quarantine-policies-in-the-microsoft-defender-portal'),
'QUA-03': (TP + 'Quarantaenerichtlinie > Zahnrad "Globale Einstellungen" > "Spambenachrichtigung fuer Endbenutzer senden alle" (4 Stunden / taeglich / woechentlich)',
    P_QUA, L_QUAPOL + '#customize-all-quarantine-notifications'),
'QUA-04': (TP + 'Quarantaenerichtlinie > Zahnrad "Globale Einstellungen" > Absenderadresse, Anzeigename, Betreff, Haftungsausschluss und Firmenlogo',
    P_QUA, L_QUAPOL + '#customize-all-quarantine-notifications'),

# ============================================================ Advanced Delivery
'ADV-01': (TP + 'Erweiterte Zustellung (Advanced delivery) > Reiter "SecOps-Postfach" (nur echte Postfaecher, keine Verteilergruppen)',
    P_ADV,
    'https://learn.microsoft.com/defender-office-365/advanced-delivery-policy-configure#use-the-microsoft-defender-portal-to-configure-secops-mailboxes-in-the-advanced-delivery-policy'),
'ADV-02': (TP + 'Erweiterte Zustellung > Reiter "Phishingsimulation" (Domaene + Sende-IP; IPv6 nur per PowerShell)',
    P_ADV,
    'https://learn.microsoft.com/defender-office-365/advanced-delivery-policy-configure#use-the-microsoft-defender-portal-to-configure-non-microsoft-phishing-simulations-in-the-advanced-delivery-policy'),
'ADV-03': ('Exchange Admin Center > Nachrichtenfluss > Regeln: jede Regel auf "Spamfilterung umgehen" (SCL -1) pruefen. Solche Bypaesse gehoeren in die Erweiterte Zustellung oder die TABL',
    P_RULES,
    'https://learn.microsoft.com/defender-office-365/create-safe-sender-lists-in-office-365#use-mail-flow-rules'),

# ============================================================ E-Mail-Authentifizierung
'AUTH-01': ('Oeffentliches DNS beim Domain-Hoster: TXT auf <domaene> mit v=spf1 include:spf.protection.outlook.com -all, je akzeptierter Domaene genau einer',
    '', 'https://learn.microsoft.com/defender-office-365/email-authentication-spf-configure#spf-txt-records-for-custom-domains-in-microsoft-365'),
'AUTH-02': ('Oeffentliches DNS: TXT auf <domaene>, Endqualifier -all statt ~all',
    '', 'https://learn.microsoft.com/defender-office-365/email-authentication-spf-configure#syntax-for-spf-txt-records'),
'AUTH-03': ('Oeffentliches DNS: fuer geparkte Domaenen TXT v=spf1 -all auf <domaene> plus TXT v=DMARC1; p=reject; auf _dmarc.<domaene>. Gilt auch fuer die onmicrosoft.com-Domaene',
    '', 'https://learn.microsoft.com/defender-office-365/email-authentication-spf-configure#scenario-parked-domains'),
'AUTH-04': (TP + 'E-Mail-Authentifizierungseinstellungen > DKIM > Domaene waehlen, Umschalter aktivieren. Vorher die beiden CNAMEs selector1/2._domainkey.<domaene> ins oeffentliche DNS',
    P_AUTHN + '?viewid=DKIM',
    'https://learn.microsoft.com/defender-office-365/email-authentication-dkim-configure#use-the-defender-portal-to-enable-dkim-signing-of-outbound-messages-using-a-custom-domain'),
'AUTH-05': ('Die Bitlaenge steuert nur PowerShell: Rotate-DkimSigningConfig -KeySize 2048. Im Portal rotiert man ueber E-Mail-Authentifizierungseinstellungen > DKIM > Domaenenzeile ANKLICKEN > Flyout > "DKIM-Schluessel rotieren"',
    P_AUTHN + '?viewid=DKIM',
    'https://learn.microsoft.com/defender-office-365/email-authentication-dkim-configure#use-exchange-online-powershell-to-rotate-the-dkim-keys-for-a-domain-and-change-the-bit-depth'),
'AUTH-06': ('Oeffentliches DNS: TXT auf _dmarc.<domaene> mit v=DMARC1; p=...; rua=mailto:... Fuer die onmicrosoft.com-Domaene ueber Microsoft 365 Admin Center > Einstellungen > Domaenen',
    'https://admin.microsoft.com/Adminportal/Home#/Domains',
    'https://learn.microsoft.com/defender-office-365/email-authentication-dmarc-configure#set-up-dmarc-for-active-custom-domains-in-microsoft-365'),
'AUTH-07': ('Oeffentliches DNS: rua=mailto: im TXT auf _dmarc.<domaene>, Ziel ein dediziertes Shared Mailbox, kein Benutzerpostfach',
    '', 'https://learn.microsoft.com/defender-office-365/email-authentication-dmarc-configure#best-practices-for-dmarc-reports'),
'AUTH-08': (TP + 'E-Mail-Authentifizierungseinstellungen > ARC; alternativ Set-ArcConfig -ArcTrustedSealers',
    P_AUTHN,
    'https://learn.microsoft.com/defender-office-365/email-authentication-arc-configure#use-the-microsoft-defender-portal-to-add-trusted-arc-sealers'),
'AUTH-09': ('Oeffentliches DNS: TXT auf _mta-sts.<domaene> PLUS extern gehostete Policy-Datei unter https://mta-sts.<domaene>/.well-known/mta-sts.txt. Exchange Online hostet die Datei nicht. Ausgehend ist MTA-STS immer aktiv',
    '', 'https://learn.microsoft.com/exchange/security-and-compliance/enhance-mail-flow-using-strict-transport-security#adopt-mta-sts-for-your-domain'),
'AUTH-10': ('Ausgehend standardmaessig an. Eingehend per Exchange Online PowerShell: Enable-DnssecForVerifiedDomain, dann Enable-SmtpDaneInbound, danach den ausgegebenen MX-Wert beim Hoster setzen und die Delegation DNSSEC-signieren',
    '', 'https://learn.microsoft.com/exchange/security-and-compliance/how-dane-secures-email#inbound-smtp-dane-with-dnssec'),
'AUTH-11': ('Kein Portal: SPF, DKIM und DMARC ausgerichtet auf die 5322.From-Domaene plus auffindbarer Abmeldemechanismus. Greift ab 5.000 Nachrichten/Tag an Microsoft-Consumer-Dienste, sonst NDR 550 5.7.515',
    '', 'https://learn.microsoft.com/defender-office-365/external-senders-policies-practices-guidelines'),
'AUTH-12': ('Oeffentliches DNS: TXT auf default._bimi.<domaene>. Voraussetzung ist DMARC p=quarantine oder p=reject. Exchange Online wertet BIMI derzeit nicht aus',
    '', 'https://learn.microsoft.com/dynamics365/customer-insights/journeys/bimi-support'),

# ============================================================ Enhanced Filtering
'EF-01': (TP + 'Erweiterte Filterung (Enhanced filtering) > Eintrag des Inbound-Connectors > zu ueberspringende IPs bzw. "Letzte IP ueberspringen"',
    'https://security.microsoft.com/skiplisting',
    'https://learn.microsoft.com/exchange/mail-flow-best-practices/use-connectors-to-configure-mail-flow/enhanced-filtering-for-connectors#use-the-microsoft-defender-portal-to-configure-enhanced-filtering-for-connectors-on-an-inbound-connector'),
'EF-02': ('Exchange Admin Center > Nachrichtenfluss > Regeln: SCL-(-1)-Regeln fuer Nachrichten ueber diesen Connector abschalten',
    P_RULES,
    'https://learn.microsoft.com/exchange/mail-flow-best-practices/use-connectors-to-configure-mail-flow/enhanced-filtering-for-connectors#what-do-you-need-to-know-before-you-begin'),

# ============================================================ Externe Kennzeichnung
'EXT-01': ('Kein Portal, nur Exchange Online PowerShell: Set-ExternalInOutlook -Enabled $true, Ausnahmen ueber -AllowList (max. 200). Wirkung erst nach 24-48 Stunden',
    '', 'https://learn.microsoft.com/powershell/module/exchangepowershell/set-externalinoutlook'),
'EXT-02': ('Entscheidung zwischen nativem Tag (PowerShell), First-Contact-Tipp (Antiphishing-Richtlinie) und Banner per Transportregel. Nicht mehrere gleichzeitig',
    P_PHISH, L_PHISHABT + '#first-contact-safety-tip'),
'EXT-03': ('Exchange Admin Center > Nachrichtenfluss > Regeln: Regel mit "Absender ausserhalb der Organisation" und Aktion "Haftungsausschluss voranstellen", Fallback-Aktion und Ausnahmen setzen',
    P_RULES,
    'https://learn.microsoft.com/exchange/security-and-compliance/mail-flow-rules/disclaimers-signatures-footers-or-headers#use-the-eac-to-add-a-disclaimer-or-other-email-header-or-footer'),

# ============================================================ Reporting
'REP-01': ('Defender-Portal > Berichte > E-Mail & Zusammenarbeit; Zeitplan je Bericht ueber "Create schedule"',
    'https://security.microsoft.com/emailandcollabreport',
    'https://learn.microsoft.com/defender-office-365/reports-email-security'),
'REP-02': ('Defender-Portal > Richtlinien & Regeln > Warnungsrichtlinie; Empfaenger je Richtlinie im Feld "Email recipients"',
    'https://security.microsoft.com/alertpolicies',
    'https://learn.microsoft.com/defender-xdr/alert-policies#alert-policy-settings'),
'REP-03': (SET + 'Benutzerdefinierte Meldungen (User reported settings); alternativ *-ReportSubmissionPolicy in PowerShell',
    'https://security.microsoft.com/securitysettings/userSubmission',
    'https://learn.microsoft.com/defender-office-365/submissions-user-reported-messages-custom-mailbox#use-the-microsoft-defender-portal-to-configure-user-reported-settings'),
'REP-04': ('Kein Portalpunkt, sondern Skripte: auf Get-MessageTraceV2 / Get-MessageTraceDetailV2 umstellen. Interaktiv fuehrt der Defender-Portal-Eintrag nur ins EAC',
    P_EAC + '/#/messagetrace',
    'https://learn.microsoft.com/powershell/module/exchangepowershell/get-messagetracev2'),

# ============================================================ Hybrid / On-Premises
'HYB-01': ('Nur On-Premises: Exchange Management Shell (Get-ExchangeServer, ExSetup /Version) gegen die Supportability-Matrix. Exchange 2016 und 2019 sind seit 14.10.2025 out of support',
    '', 'https://learn.microsoft.com/exchange/plan-and-deploy/supportability-matrix#supported-versions-and-builds'),
'HYB-02': ('On-Premises: Skript ConfigureExchangeHybridApplication.ps1 bzw. HCW. Kontrolle der Dienstprinzipal-Anmeldungen in Microsoft Entra ID > Ueberwachung > Anmeldeprotokolle',
    'https://entra.microsoft.com',
    'https://learn.microsoft.com/exchange/hybrid-deployment/deploy-dedicated-hybrid-app#configure-the-dedicated-exchange-hybrid-application'),
'HYB-03': ('Nur On-Premises: "Default Frontend <Server>" nicht fuer anonymes Relay oeffnen, stattdessen dedizierter Receive-Connector mit RemoteIpRanges',
    '', 'https://learn.microsoft.com/exchange/mail-flow/connectors/allow-anonymous-relay#step-1-create-a-dedicated-receive-connector-for-anonymous-relay'),
'HYB-04': ('Perimeter- und Host-Firewall: ausgehend SMTP 25/587 nur von den Transport-Servern, POP3 110/995 und IMAP4 143/993 nach aussen sperren',
    '', 'https://learn.microsoft.com/exchange/plan-and-deploy/deployment-ref/network-ports#network-ports-required-for-mail-flow'),
'HYB-05': ('Exchange Admin Center > Einstellungen > E-Mail-Fluss > "SMTP AUTH-Protokoll fuer Ihre Organisation deaktivieren". Ausnahmen je Postfach ueber Set-CASMailbox',
    P_EAC + '/#/settings',
    'https://learn.microsoft.com/exchange/clients-and-mobile-in-exchange-online/authenticated-client-smtp-submission#disable-smtp-auth-in-your-organization'),
'HYB-06': ('Exchange Admin Center > Empfaenger > Postfaecher > <Postfach> > "E-Mail-App-Einstellungen verwalten"; organisationsweit ueber Set-CASMailbox bzw. Set-CASMailboxPlan fuer neue Postfaecher',
    P_EAC + '/#/mailboxes',
    'https://learn.microsoft.com/exchange/recipients-in-exchange-online/manage-user-mailboxes/managing-email-apps-for-user-mailboxes#use-exchange-online-powershell-to-enable-or-disable-email-apps'),
'HYB-07': ('Exchange Admin Center > Empfaenger > Postfaecher > <Postfach> > "Nachrichtengroessenbeschraenkung verwalten"; die Dienstgrenzwerte sind harte Limits',
    P_EAC + '/#/mailboxes',
    'https://learn.microsoft.com/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits#message-limits'),
'HYB-08': ('Exchange Admin Center > Nachrichtenfluss > Connectors > ausgehender Connector: "Immer eine TLS-gesicherte Verbindung verwenden" plus Zertifikatspruefung',
    P_CONN,
    'https://learn.microsoft.com/exchange/mail-flow-best-practices/use-connectors-to-configure-mail-flow/set-up-connectors-to-secure-mail-sent-to-partner-organization#for-new-eac'),
'HYB-09': ('Exchange Admin Center > Nachrichtenfluss > Connectors: alle ein- und ausgehenden Connectors inventarisieren (HCW vs. manuell, IP-Bereiche, Zertifikate)',
    P_CONN,
    'https://learn.microsoft.com/exchange/mail-flow-best-practices/use-connectors-to-configure-mail-flow/use-connectors-to-configure-mail-flow#when-do-i-need-a-connector'),
'HYB-10': ('Exchange Admin Center > Nachrichtenfluss > Akzeptierte Domaenen: Typ "Autorisierend" schaltet DBEB ein, "Internes Relay" schaltet es ab',
    P_EAC + '/#/accepteddomains',
    'https://learn.microsoft.com/exchange/mail-flow-best-practices/use-directory-based-edge-blocking#configure-dbeb'),

# ============================================================ Betrieb
'OPS-01': ('Defender-Portal > Berechtigungen: Rollengruppen "E-Mail & Zusammenarbeit" bzw. Defender XDR Unified RBAC; dazu die Entra-Rollen',
    'https://security.microsoft.com/emailandcollabpermissions',
    'https://learn.microsoft.com/defender-office-365/scc-permissions#role-groups-in-microsoft-defender-for-office-365-and-microsoft-purview'),
'OPS-02': ('Kein Portal: Admin-Arbeitsplatz. Modul ExchangeOnlineManagement, Windows PowerShell 5.1 mit .NET 4.7.2+ oder PowerShell 7',
    '', 'https://learn.microsoft.com/powershell/exchange/exchange-online-powershell-v2'),
'OPS-03': ('Baseline per PowerShell exportieren; ergaenzend Konfigurationsanalyse > Reiter "Configuration drift analysis and history" (setzt Unified Auditing voraus)',
    P_CFGAN,
    'https://learn.microsoft.com/defender-office-365/configuration-analyzer-for-security-policies#configuration-drift-analysis-and-history-tab-in-the-configuration-analyzer'),
'OPS-04': ('Nachweis ueber Microsoft Purview > Audit (Unified Audit Log); Nachkontrolle als wiederkehrenden Termin setzen',
    'https://purview.microsoft.com',
    'https://learn.microsoft.com/purview/audit-log-enable-disable'),

# ============================================================ Teams
'TEAMS-01': (SET + 'Microsoft Teams-Schutz (NICHT mehr unter Bedrohungsrichtlinien) > ZAP fuer Teams, Quarantaenerichtlinie, Ausnahmen',
    'https://security.microsoft.com/securitysettings/teamsProtectionPolicy',
    'https://learn.microsoft.com/defender-office-365/mdo-support-teams-about#configure-zap-for-teams-protection-in-defender-for-office-365'),

# ============================================================ Wirksamkeitsnachweis
'VER-01': ('Defender-Portal > Ueberpruefen > Quarantaene, Filter Richtlinientyp = Antischadsoftware-Richtlinie',
    'https://security.microsoft.com/quarantine',
    L_MALABOUT + '#common-attachments-filter-in-anti-malware-policies'),
'VER-02': ('Kein Portal: Outlook > Datei > Eigenschaften > Internetkopfzeilen (neues Outlook/OWA: Nachricht > ... > Details anzeigen). Ausgewertet werden X-Forefront-Antispam-Report, X-Microsoft-Antispam und Authentication-Results',
    '', 'https://learn.microsoft.com/defender-office-365/message-headers-eop-mdo#x-forefront-antispam-report-message-header-fields'),
'VER-03': ('Defender-Portal > E-Mail & Zusammenarbeit > Explorer (P2) bzw. Echtzeiterkennungen (P1), Ansichten Phish und URL-Klicks',
    'https://security.microsoft.com/threatexplorerv3',
    'https://learn.microsoft.com/defender-office-365/threat-explorer-real-time-detections-about'),

# ============================================================ Governance / P2
'GOV-01': ('Kein Microsoft-Portalpunkt: organisatorisch - Datenschutz-Folgenabschaetzung, Verfahrensverzeichnis und Mitbestimmung fuer Quarantaene-Einsicht und Explorer-Preview',
    '', 'https://learn.microsoft.com/compliance/regulatory/gdpr#data-protection-impact-assessment'),
'P2-01': ('Defender-Portal > Untersuchungen (AIR, nur MDO Plan 2; setzt aktiviertes Audit-Logging voraus)',
    'https://security.microsoft.com/airinvestigation',
    'https://learn.microsoft.com/defender-office-365/air-about#the-overall-flow-of-air'),
}

# Deep-Links, die Microsoft nicht selbst dokumentiert (aus der Praxis uebernommen).
UNDOCUMENTED = ['ASO-07', 'ASO-08', 'HYB-06', 'HYB-07', 'HYB-10']

# Pruefpunkte ohne passgenauen Learn-Artikel - der Link fuehrt auf das Naechstliegende.
WEAK_LEARN = {
    'AUTH-11': 'Learn deckt die High-Volume-Sender-Anforderungen nicht ab; belastbar sind nur der Support-Artikel zu NDR 550 5.7.515 und der MDO-Blogpost vom 30.04.2025.',
    'AUTH-12': 'Fuer BIMI gibt es keinen EOP/MDO-Artikel. Die verlinkte Seite gehoert zu Dynamics 365 und stellt fest, dass Exchange Online BIMI nicht auswertet.',
    'HYB-04': 'Kein Firewall-Artikel bei Microsoft; verlinkt ist die Exchange-Portreferenz.',
    'GOV-01': 'Mitbestimmung nach BetrVG ist bei Microsoft nicht dokumentiert; verlinkt ist der DSGVO-Artikel.',
}

# Kuerzel-Legende
PREFIXES = [
    ('PRE',   'Preset Security Policies und Governance', 'Voreingestellte Sicherheitsrichtlinien, Praezedenz, Configuration Analyzer'),
    ('ASI',   'Anti-Spam inbound',                       'Anti-Spam-Richtlinie, eingehend'),
    ('ASF',   'Advanced Spam Filter',                    'Die ASF-Schalter innerhalb der Anti-Spam-Richtlinie'),
    ('CF',    'Connection Filter',                       'Verbindungsfilterrichtlinie - liegt als eigene Zeile in der Liste der Anti-Spam-Richtlinien'),
    ('ASO',   'Anti-Spam outbound',                      'Ausgehende Anti-Spam-Richtlinie plus Weiterleitungswege'),
    ('APH',   'Anti-Phishing (EOP-Teil)',                'Spoof-Intelligence, DMARC-Behandlung, Sicherheitstipps - ohne Defender-Lizenz verfuegbar'),
    ('IMP',   'Impersonation (Defender-Teil)',           'Identitaetswechselschutz - sitzt in derselben Anti-Phishing-Richtlinie, ist aber lizenzpflichtig'),
    ('TABL',  'Tenant Allow/Block List',                 'Mandanten-Zulassungs-/Sperrliste'),
    ('AMW',   'Anti-Malware',                            'Anti-Malware-Richtlinie und Anlagenfilter'),
    ('SA',    'Safe Attachments',                        'Sichere Anlagen (Defender)'),
    ('SL',    'Safe Links',                              'Sichere Links (Defender)'),
    ('QUA',   'Quarantaene-Richtlinien',                 'Zugriffsrechte und Benachrichtigungen fuer die Quarantaene'),
    ('ADV',   'Advanced Delivery',                       'Erweiterte Zustellung: SecOps-Postfach und Phishing-Simulation'),
    ('AUTH',  'E-Mail-Authentifizierung',                'SPF, DKIM, DMARC, ARC, MTA-STS, DANE - ueberwiegend im oeffentlichen DNS'),
    ('EF',    'Enhanced Filtering',                      'Erweiterte Filterung am Inbound-Connector (Skip Listing)'),
    ('EXT',   'Externe Kennzeichnung',                   'External-Tag, First-Contact-Tipp, Banner per Transportregel'),
    ('REP',   'Reporting und Alerting',                  'Berichte, Warnungsrichtlinien, Meldeweg fuer Nutzer'),
    ('HYB',   'Hybrid und On-Premises',                  'Exchange Server, Connectors, akzeptierte Domaenen, Legacy-Protokolle'),
    ('OPS',   'Betrieb und Voraussetzungen',             'Rollen, PowerShell-Arbeitsplatz, Baseline, Nachkontrolle'),
    ('TEAMS', 'Microsoft Teams Protection',              'ZAP fuer Teams-Nachrichten'),
    ('VER',   'Wirksamkeitsnachweis',                    'Belegen, dass der Schutz greift - Quarantaene, Header, Explorer'),
    ('GOV',   'Governance und Datenschutz',              'DSGVO und Mitbestimmung rund um Quarantaene-Einsicht'),
    ('P2',    'Defender Plan 2',                         'Nur mit MDO P2: automatisierte Untersuchung (AIR)'),
]

# Zusatz-IDs, die die Skripte selbst erzeugen (Inventar, Sammelbefunde),
# und der Pruefpunkt, dessen Fundort fuer sie gilt.
ALIASES = {
    'INV-ASI':        'PRE-05',
    'INV-ASO':        'ASO-01',
    'INV-APH':        'APH-01',
    'INV-AMW':        'AMW-01',
    'INV-SA':         'SA-01',
    'INV-SL':         'SL-01',
    'AUTH-DNS':       'AUTH-01',
    'TABL-Sender':    'TABL-02',
    'TABL-Url':       'TABL-04',
    'TABL-FileHash':  'TABL-05',
    'TABL-IP':        'TABL-06',
}
