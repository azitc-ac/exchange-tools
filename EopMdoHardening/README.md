# EOP / Microsoft Defender for Office 365 — Hardening-Standard

Wiederverwendbare Methodik für EOP- und MDO-Assessments: 120 Prüfpunkte mit Soll-Werten, Begründung, PowerShell-Befehl und — der Grund für diese README — **dem Fundort im Portal**.

Alle Soll-Werte, Cmdlets, Parameter und Default-/Preset-Werte sind gegen die Microsoft-Learn-Referenz geprüft (Recherchestand August 2026). Nicht belegbare Aussagen sind im Guide als solche gekennzeichnet.

## Inhalt

| Datei | Zweck |
|---|---|
| `Invoke-EopAudit.ps1` | Liest den Ist-Zustand aus. **Rein lesend.** Liefert CSV, HTML-Report und optional einen JSON-Export als Rollback-Grundlage. Das Arbeitstier. |
| `Invoke-EopHardening.ps1` | Setzt die Baseline. Vorschau ist der Standardmodus — ohne `-Execute` wird nichts geändert. |
| `EOP-MDO_Assessment-Checkliste.xlsx` | 120 Prüfpunkte zum Ausfüllen. Ist-Wert, Bewertung, Notizen; die Auswertung rechnet mit. Blatt „Kürzel" enthält die Legende. |
| `EOP-MDO_Best-Practice-Guide.docx` / `.pdf` | Begründung zu jedem Prüfpunkt: Funktion, Nutzen, Microsoft-Stand 2026, Lizenz, Fallstricke. Anhang E ist die Fundort-Tabelle. |

Checkliste, Guide und beide Skripte verwenden **dieselben IDs**. Steht im Audit-Report `ASI-04`, findet sich im Guide unter derselben Nummer die Begründung und in der Checkliste die Zeile zum Ausfüllen.

## Audit beim Kunden

```powershell
# Verbindung steht schon? Dann einfach:
.\Invoke-EopAudit.ps1 -CustomerName "Kunde" -ExportJson

# Noch nicht verbunden — das Skript meldet sich selbst an:
.\Invoke-EopAudit.ps1 -CustomerName "Kunde" -UserPrincipalName admin@kunde.onmicrosoft.com -ExportJson

# Tenantwechsel im selben Fenster:
.\Invoke-EopAudit.ps1 -CustomerName "Kunde2" -ForceNewConnection
```

Rollen: **Global Reader** plus **Security Reader** genügen fürs Audit, für die Umsetzung braucht es **Security Administrator** und **Exchange Administrator**.

Voraussetzungen: Windows PowerShell 5.1 (unterstützt) oder PowerShell 7, Modul `ExchangeOnlineManagement` V3, TLS 1.2, .NET Framework 4.7.2+. Beide Skripte sind UTF-8 **mit BOM** gespeichert und PS-5.1-kompatibel.

Blockt Windows die Datei als „aus dem Internet": `Unblock-File .\Invoke-EopAudit.ps1`.

### Verbindung

Beide Skripte erkennen eine bestehende Exchange-Online-Verbindung und verwenden sie weiter — einmal anmelden, dann Audit und Hardening nacheinander laufen lassen. Grundlage ist `Get-ConnectionInformation`; `Get-PSSession` liefert für die REST-Verbindungen seit Modul V3 nichts mehr.

| Situation | Verhalten |
|---|---|
| Verbindung besteht | wird weiterverwendet |
| keine Verbindung | es wird eine neue aufgebaut |
| nur Security-&-Compliance-Verbindung da | Exchange Online wird zusätzlich verbunden (`IsEopSession` unterscheidet) |
| Token abgelaufen | wird verworfen, neue Anmeldung |
| `-ForceNewConnection` | bestehende wird getrennt, neue Anmeldung |
| `-SkipConnect` | es wird unter keinen Umständen verbunden |

Beim Start zeigen beide Skripte **Organisation, angemeldetes Konto und Tenant-ID** — die Absicherung gegen den teuersten Irrtum beim Arbeiten mit mehreren Kundentenants. Weicht das Konto von `-UserPrincipalName` ab, kommt eine Warnung; das Hardening-Skript fragt im `-Execute`-Modus zusätzlich nach. Getrennt wird nie automatisch.

### Was der Report ausgibt

- **CSV** — für die Übernahme in die Checkliste, mit den Spalten `Fundort`, `PortalLink`, `LearnLink`
- **HTML** — Befunde nach Dringlichkeit, unter jedem Befund der Klickpfad plus Direktlinks ins Portal und zu Learn. Das ist die Ansicht für die Besprechung mit dem Kunden.
- **JSON** (`-ExportJson`) — der Ist-Zustand als Rollback-Grundlage, bevor etwas geändert wird

## Die Kürzel

Jede ID besteht aus Präfix und laufender Nummer. Das Präfix benennt den **fachlichen Bereich** — nicht zwingend die Portal-Richtlinie, in der die Einstellung sitzt. Das fällt an drei Stellen auseinander, und zwar absichtlich:

- **APH und IMP** stehen beide in derselben Anti-Phishing-Richtlinie. Getrennt sind sie, weil APH mit jeder Lizenz funktioniert und IMP eine Defender-Lizenz braucht. Bei einem EOP-only-Kunden lassen sich damit alle IMP-Punkte in einem Zug auf *n.a.* setzen — das ist der praktische Nutzen.
- **ASF** sitzt innerhalb der Anti-Spam-Richtlinie, ist aber ein eigenes, historisch gewachsenes Regelwerk mit eigener Bewertungslogik.
- **CF** erscheint im Portal als Zeile *in der Liste der Anti-Spam-Richtlinien*, ist technisch aber eine eigene Richtlinie mit eigenen Cmdlets und greift vor allen anderen.

| Kürzel | Steht für | Punkte | Wo es im Portal liegt |
|---|---|---|---|
| **PRE** | Preset Security Policies und Governance | 5 | Voreingestellte Sicherheitsrichtlinien, Präzedenz, Configuration Analyzer |
| **ASI** | Anti-Spam inbound | 12 | Anti-Spam-Richtlinie, eingehend |
| **ASF** | Advanced Spam Filter | 5 | Die ASF-Schalter innerhalb der Anti-Spam-Richtlinie |
| **CF** | Connection Filter | 3 | Verbindungsfilterrichtlinie — liegt als eigene Zeile in der Liste der Anti-Spam-Richtlinien |
| **ASO** | Anti-Spam outbound | 8 | Ausgehende Anti-Spam-Richtlinie plus Weiterleitungswege |
| **APH** | Anti-Phishing (EOP-Teil) | 8 | Spoof-Intelligence, DMARC-Behandlung, Sicherheitstipps — ohne Defender-Lizenz verfügbar |
| **IMP** | Impersonation (Defender-Teil) | 8 | Identitätswechselschutz — sitzt in derselben Anti-Phishing-Richtlinie, ist aber lizenzpflichtig |
| **TABL** | Tenant Allow/Block List | 6 | Mandanten-Zulassungs-/Sperrliste |
| **AMW** | Anti-Malware | 5 | Anti-Malware-Richtlinie und Anlagenfilter |
| **SA** | Safe Attachments | 5 | Sichere Anlagen (Defender) |
| **SL** | Safe Links | 7 | Sichere Links (Defender) |
| **QUA** | Quarantäne-Richtlinien | 4 | Zugriffsrechte und Benachrichtigungen für die Quarantäne |
| **ADV** | Advanced Delivery | 3 | Erweiterte Zustellung: SecOps-Postfach und Phishing-Simulation |
| **AUTH** | E-Mail-Authentifizierung | 12 | SPF, DKIM, DMARC, ARC, MTA-STS, DANE — überwiegend im öffentlichen DNS |
| **EF** | Enhanced Filtering | 2 | Erweiterte Filterung am Inbound-Connector (Skip Listing) |
| **EXT** | Externe Kennzeichnung | 3 | External-Tag, First-Contact-Tipp, Banner per Transportregel |
| **REP** | Reporting und Alerting | 4 | Berichte, Warnungsrichtlinien, Meldeweg für Nutzer |
| **HYB** | Hybrid und On-Premises | 10 | Exchange Server, Connectors, akzeptierte Domänen, Legacy-Protokolle |
| **OPS** | Betrieb und Voraussetzungen | 4 | Rollen, PowerShell-Arbeitsplatz, Baseline, Nachkontrolle |
| **TEAMS** | Microsoft Teams Protection | 1 | ZAP für Teams-Nachrichten |
| **VER** | Wirksamkeitsnachweis | 3 | Belegen, dass der Schutz greift — Quarantäne, Header, Explorer |
| **GOV** | Governance und Datenschutz | 1 | DSGVO und Mitbestimmung rund um Quarantäne-Einsicht |
| **P2** | Defender Plan 2 | 1 | Nur mit MDO P2: automatisierte Untersuchung (AIR) |

Dazu kommen IDs, die die Skripte selbst erzeugen und die auf den Fundort ihres Prüfpunkts verweisen:

| Erzeugte ID | Bedeutung | Fundort von |
|---|---|---|
| `AUTH-DNS` | Sammelbefund der DNS-Prüfung (SPF, DKIM, DMARC) | AUTH-01 |
| `INV-AMW` | Inventar der Anti-Malware-Richtlinien | AMW-01 |
| `INV-APH` | Inventar der Anti-Phishing-Richtlinien | APH-01 |
| `INV-ASI` | Inventar der Anti-Spam-Richtlinien (eingehend) | PRE-05 |
| `INV-ASO` | Inventar der Anti-Spam-Richtlinien (ausgehend) | ASO-01 |
| `INV-SA` | Inventar der Safe-Attachments-Richtlinien | SA-01 |
| `INV-SL` | Inventar der Safe-Links-Richtlinien | SL-01 |
| `TABL-FileHash` | Einträge für Dateien (SHA256) | TABL-05 |
| `TABL-IP` | Einträge für IP-Adressen | TABL-06 |
| `TABL-Sender` | Einträge für Absender und Domänen | TABL-02 |
| `TABL-Url` | Einträge für URLs | TABL-04 |

`INFO-01` ist die Tenant-Information im Kopf des Reports und hat keinen Fundort.

## Fundorte: wo die Einstellung im Portal sitzt

Der Pfad **Bedrohungsrichtlinien** liegt im Defender-Portal unter *E-Mail & Zusammenarbeit > Richtlinien & Regeln*. Wo kein Portal-Link steht, ist die Einstellung nicht über die Oberfläche erreichbar — das betrifft vor allem das öffentliche DNS und einige PowerShell-only-Schalter.

Zwei Fundorte, an denen man erfahrungsgemäß vorbeiscrollt:

- Die **Verbindungsfilterrichtlinie** (`CF-*`) hat keine eigene Seite. Sie steht als Zeile *in der Liste der Anti-Spam-Richtlinien* — die **Zeile** anklicken, nicht die Checkbox, dann im Flyout auf „Verbindungsfilterrichtlinie bearbeiten".
- Der **Teams-Schutz** (`TEAMS-01`) liegt nicht mehr unter Bedrohungsrichtlinien, sondern unter *Einstellungen > E-Mail & Zusammenarbeit > Microsoft Teams-Schutz*.

### Voreingestellte Richtlinien und Governance

| ID | Klickpfad | Direkt |
|:------|:------------------------------------------------------------------------------------------|:------|
| `PRE-01` | Defender-Portal > Bedrohungsrichtlinien > Voreingestellte Sicherheitsrichtlinien (Preset security policies) | [Portal](https://security.microsoft.com/presetSecurityPolicies) · [Learn](https://learn.microsoft.com/defender-office-365/mdo-deployment-guide#determine-your-threat-policy-strategy) |
| `PRE-02` | Defender-Portal > Bedrohungsrichtlinien > Voreingestellte Sicherheitsrichtlinien > Integrierter Schutz (Built-in protection) > "Ausschlüsse verwalten" | [Portal](https://security.microsoft.com/presetSecurityPolicies) · [Learn](https://learn.microsoft.com/defender-office-365/preset-security-policies#use-the-microsoft-defender-portal-to-add-exclusions-to-the-built-in-protection-preset-security-policy) |
| `PRE-03` | Kein eigener Portal-Screen. Präzedenz ablesen an der Spalte "Priorität" je Richtlinientyp (Antispam, Antiphishing, ...) plus der Preset-Seite | [Portal](https://security.microsoft.com/presetSecurityPolicies) · [Learn](https://learn.microsoft.com/defender-office-365/preset-security-policies#order-of-precedence-for-preset-security-policies-and-other-threat-policies) |
| `PRE-04` | Defender-Portal > Bedrohungsrichtlinien > Konfigurationsanalyse (Configuration analyzer) > Reiter "Standardempfehlungen" / "Strenge Empfehlungen" | [Portal](https://security.microsoft.com/configurationAnalyzer) · [Learn](https://learn.microsoft.com/defender-office-365/configuration-analyzer-for-security-policies#standard-recommendations-and-strict-recommendations-tabs-in-the-configuration-analyzer) |
| `PRE-05` | Defender-Portal > Bedrohungsrichtlinien > Antispam (Spalten Status / Priorität / Typ), analog Antiphishing, Antischadsoftware, Sichere Links, Sichere Anlagen | [Portal](https://security.microsoft.com/antispam) · [Learn](https://learn.microsoft.com/defender-office-365/anti-spam-policies-configure#use-the-microsoft-defender-portal-to-enable-or-disable-anti-spam-policies) |

### Anti-Spam eingehend

| ID | Klickpfad | Direkt |
|:------|:------------------------------------------------------------------------------------------|:------|
| `ASF-01` | Defender-Portal > Bedrohungsrichtlinien > Antispam > `<Richtlinie>` > "Massen-E-Mail-Schwellenwert & Spameigenschaften" > Spameigenschaften > Gruppe "Erhöhen der Spambewertung" | [Portal](https://security.microsoft.com/antispam) · [Learn](https://learn.microsoft.com/defender-office-365/anti-spam-policies-asf-settings-about#increase-spam-score-settings) |
| `ASF-02` | Defender-Portal > Bedrohungsrichtlinien > Antispam > `<Richtlinie>` > "Massen-E-Mail-Schwellenwert & Spameigenschaften" > Spameigenschaften > Gruppe "Als Spam markieren" | [Portal](https://security.microsoft.com/antispam) · [Learn](https://learn.microsoft.com/defender-office-365/anti-spam-policies-asf-settings-about#mark-as-spam-settings) |
| `ASF-03` | Defender-Portal > Bedrohungsrichtlinien > Antispam > `<Richtlinie>` > "Massen-E-Mail-Schwellenwert & Spameigenschaften" > Spameigenschaften > "SPF-Eintrag: Hard Fail" | [Portal](https://security.microsoft.com/antispam) · [Learn](https://learn.microsoft.com/defender-office-365/anti-spam-policies-asf-settings-about#mark-as-spam-settings) |
| `ASF-04` | Defender-Portal > Bedrohungsrichtlinien > Antispam > `<Richtlinie>` > "Massen-E-Mail-Schwellenwert & Spameigenschaften" > Spameigenschaften > "Absender-ID-Filterung: Hard Fail" und "Backscatter" | [Portal](https://security.microsoft.com/antispam) · [Learn](https://learn.microsoft.com/defender-office-365/anti-spam-policies-asf-settings-about#mark-as-spam-settings) |
| `ASF-05` | Defender-Portal > Bedrohungsrichtlinien > Antispam > `<Richtlinie>` > "Massen-E-Mail-Schwellenwert & Spameigenschaften" > Spameigenschaften > "Testmodus" (gilt global für alle auf Test gesetzten ASF-Optionen) | [Portal](https://security.microsoft.com/antispam) · [Learn](https://learn.microsoft.com/defender-office-365/anti-spam-policies-asf-settings-about#enable-disable-or-test-asf-settings) |
| `ASI-01` | Defender-Portal > Bedrohungsrichtlinien > Antispam > `<Richtlinie>` > "Massen-E-Mail-Schwellenwert & Spameigenschaften" > Schieberegler "Massen-E-Mail-Schwellenwert" | [Portal](https://security.microsoft.com/antispam) · [Learn](https://learn.microsoft.com/defender-office-365/anti-spam-protection-about#bulk-complaint-threshold-bcl-in-anti-spam-policies) |
| `ASI-02` | Defender-Portal > Bedrohungsrichtlinien > Antispam > `<Richtlinie>` > "Aktionen" > Nachrichtenaktionen > Zeile "Spam" | [Portal](https://security.microsoft.com/antispam) · [Learn](https://learn.microsoft.com/defender-office-365/anti-spam-protection-about#actions-in-anti-spam-policies) |
| `ASI-03` | Defender-Portal > Bedrohungsrichtlinien > Antispam > `<Richtlinie>` > "Aktionen" > Nachrichtenaktionen > Zeile "Spam mit hoher Sicherheit" | [Portal](https://security.microsoft.com/antispam) · [Learn](https://learn.microsoft.com/defender-office-365/anti-spam-protection-about#actions-in-anti-spam-policies) |
| `ASI-04` | Defender-Portal > Bedrohungsrichtlinien > Antispam > `<Richtlinie>` > "Aktionen" > Nachrichtenaktionen > Zeile "Phishing" | [Portal](https://security.microsoft.com/antispam) · [Learn](https://learn.microsoft.com/defender-office-365/anti-spam-protection-about#actions-in-anti-spam-policies) |
| `ASI-05` | Defender-Portal > Bedrohungsrichtlinien > Antispam > `<Richtlinie>` > "Aktionen" > Nachrichtenaktionen > Zeile "Phishing mit hoher Sicherheit" | [Portal](https://security.microsoft.com/antispam) · [Learn](https://learn.microsoft.com/defender-office-365/anti-spam-protection-about#actions-in-anti-spam-policies) |
| `ASI-06` | Defender-Portal > Bedrohungsrichtlinien > Antispam > `<Richtlinie>` > "Aktionen" > Nachrichtenaktionen > Zeile "Massenkonforme Ebene (BCL) erreicht oder überschritten" | [Portal](https://security.microsoft.com/antispam) · [Learn](https://learn.microsoft.com/defender-office-365/anti-spam-protection-about#actions-in-anti-spam-policies) |
| `ASI-07` | Defender-Portal > Bedrohungsrichtlinien > Antispam > `<Richtlinie>` > "Aktionen" > Abschnitt "Sicherheitstipps" > Haken "Sicherheitstipps aktivieren" | [Portal](https://security.microsoft.com/antispam) · [Learn](https://learn.microsoft.com/defender-office-365/recommended-settings-for-eop-and-office365#anti-spam-policy-settings) |
| `ASI-08` | Defender-Portal > Bedrohungsrichtlinien > Antispam > `<Richtlinie>` > "Aktionen" > "Automatische Bereinigung (ZAP) aktivieren" mit den zwei Unterhaken für Phishing und Spam | [Portal](https://security.microsoft.com/antispam) · [Learn](https://learn.microsoft.com/defender-office-365/zero-hour-auto-purge#zero-hour-auto-purge-zap-for-spam) |
| `ASI-09` | Defender-Portal > Bedrohungsrichtlinien > Antispam > `<Richtlinie>` > "Aktionen" > Feld "Spamnachrichten so viele Tage lang in Quarantäne aufbewahren" (1-30, Default 15) | [Portal](https://security.microsoft.com/antispam) · [Learn](https://learn.microsoft.com/defender-office-365/quarantine-about#quarantine-retention) |
| `ASI-10` | Defender-Portal > Bedrohungsrichtlinien > Antispam > `<Richtlinie>` > "Zulassen- und Sperrliste" > "Absender verwalten" / "Domänen zulassen" (im Portal max. 30 Einträge) | [Portal](https://security.microsoft.com/antispam) · [Learn](https://learn.microsoft.com/defender-office-365/anti-spam-protection-about#allow-and-block-lists-in-anti-spam-policies) |
| `ASI-11` | Defender-Portal > Bedrohungsrichtlinien > Antispam > `<Richtlinie>` > "Massen-E-Mail-Schwellenwert & Spameigenschaften" > Spameigenschaften > "Enthält bestimmte Sprachen" / "Aus diesen Ländern" (gehört NICHT zum ASF) | [Portal](https://security.microsoft.com/antispam) · [Learn](https://learn.microsoft.com/defender-office-365/anti-spam-protection-about#spam-properties-in-anti-spam-policies) |
| `ASI-12` | Defender-Portal > Bedrohungsrichtlinien > Antispam > `<Richtlinie>` > "Aktionen" > "Organisationsinterne Nachrichten, für die Maßnahmen ergriffen werden sollen" | [Portal](https://security.microsoft.com/antispam) · [Learn](https://learn.microsoft.com/defender-office-365/anti-spam-protection-about#actions-in-anti-spam-policies) |

### Verbindungsfilter

| ID | Klickpfad | Direkt |
|:------|:------------------------------------------------------------------------------------------|:------|
| `CF-01` | Defender-Portal > Bedrohungsrichtlinien > Antispam > in der Liste die Zeile "Verbindungsfilterrichtlinie (Standard)" ANKLICKEN > Flyout > "Verbindungsfilterrichtlinie bearbeiten" > "Nachrichten aus den folgenden IP-Adressen immer zulassen" | [Portal](https://security.microsoft.com/antispam) · [Learn](https://learn.microsoft.com/defender-office-365/connection-filter-policies-configure#use-the-microsoft-defender-portal-to-modify-the-default-connection-filter-policy) |
| `CF-02` | Defender-Portal > Bedrohungsrichtlinien > Antispam > Zeile "Verbindungsfilterrichtlinie (Standard)" > "Verbindungsfilterrichtlinie bearbeiten" > "Nachrichten aus den folgenden IP-Adressen immer blockieren" | [Portal](https://security.microsoft.com/antispam) · [Learn](https://learn.microsoft.com/defender-office-365/connection-filter-policies-configure#use-the-microsoft-defender-portal-to-modify-the-default-connection-filter-policy) |
| `CF-03` | Defender-Portal > Bedrohungsrichtlinien > Antispam > Zeile "Verbindungsfilterrichtlinie (Standard)" > "Verbindungsfilterrichtlinie bearbeiten" > Haken "Sichere Liste aktivieren" | [Portal](https://security.microsoft.com/antispam) · [Learn](https://learn.microsoft.com/defender-office-365/connection-filter-policies-configure#use-the-microsoft-defender-portal-to-modify-the-default-connection-filter-policy) |

### Anti-Spam ausgehend

| ID | Klickpfad | Direkt |
|:------|:------------------------------------------------------------------------------------------|:------|
| `ASO-01` | Defender-Portal > Bedrohungsrichtlinien > Antispam > Liste auf "Ausgehende Antispamrichtlinie (Standard)" und eigene Richtlinien prüfen; neu über "+ Richtlinie erstellen > Ausgehend" | [Portal](https://security.microsoft.com/antispam) · [Learn](https://learn.microsoft.com/defender-office-365/outbound-spam-policies-configure#use-the-microsoft-defender-portal-to-create-outbound-spam-policies) |
| `ASO-02` | Defender-Portal > Bedrohungsrichtlinien > Antispam > `<ausgehende Richtlinie>` > "Schutzeinstellungen" > Nachrichtengrenzwerte (extern / intern / täglich) | [Portal](https://security.microsoft.com/antispam) · [Learn](https://learn.microsoft.com/defender-office-365/outbound-spam-policies-configure#use-the-microsoft-defender-portal-to-create-outbound-spam-policies) |
| `ASO-03` | Defender-Portal > Bedrohungsrichtlinien > Antispam > `<ausgehende Richtlinie>` > "Schutzeinstellungen" > "Einschränkung für Benutzer, die das Nachrichtenlimit erreichen" | [Portal](https://security.microsoft.com/antispam) · [Learn](https://learn.microsoft.com/defender-office-365/recommended-settings-for-eop-and-office365#outbound-spam-policy-settings) |
| `ASO-04` | Defender-Portal > Bedrohungsrichtlinien > Antispam > `<ausgehende Richtlinie>` > "Schutzeinstellungen" > Weiterleitungsregeln > "Automatische Weiterleitungsregeln" | [Portal](https://security.microsoft.com/antispam) · [Learn](https://learn.microsoft.com/defender-office-365/outbound-spam-policies-external-email-forwarding) |
| `ASO-05` | Defender-Portal > Bedrohungsrichtlinien > Antispam > "Ausgehende Antispamrichtlinie (Standard)" > "Schutzeinstellungen" > Benachrichtigungen (BCC-Kopie); gesperrte Absender unter Defender-Portal > Überprüfen > Eingeschränkte Benutzer | [Portal](https://security.microsoft.com/restrictedusers) · [Learn](https://learn.microsoft.com/defender-office-365/outbound-spam-policies-configure#use-the-microsoft-defender-portal-to-create-outbound-spam-policies) |
| `ASO-06` | Nicht konfigurierbar (mandantenweites Limit). Auswertung: Exchange Admin Center > Berichte > Nachrichtenfluss > "Tenant Outbound External Recipients" | [Portal](https://admin.exchange.microsoft.com/#/reports/mailflowreportsmain) · [Learn](https://learn.microsoft.com/defender-office-365/outbound-spam-sending-limits-troubleshoot#tenant-external-recipient-rate-limit) |
| `ASO-07` | NICHT im Defender-Portal: Exchange Admin Center > Nachrichtenfluss > Remotedomänen > "Default" > Abschnitt "Automatische Antworten" (dort sitzt die Weiterleitungsoption) | [Portal](https://admin.exchange.microsoft.com/#/remotedomains) · [Learn](https://learn.microsoft.com/exchange/mail-flow-best-practices/remote-domains/remote-domains#reducing-or-increasing-information-flow-to-another-company) |
| `ASO-08` | Keine Einstellung, mandantenweite Drosselung. Prüfung: Exchange Admin Center > Nachrichtenfluss > Akzeptierte Domänen sowie Absenderdomänen in Connectors und Anwendungen | [Portal](https://admin.exchange.microsoft.com/#/accepteddomains) · [Learn](https://learn.microsoft.com/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits#sending-limits) |

### Anti-Phishing und Impersonation

| ID | Klickpfad | Direkt |
|:------|:------------------------------------------------------------------------------------------|:------|
| `APH-01` | Defender-Portal > Bedrohungsrichtlinien > Antiphishing > `<Richtlinie>` > "Phishingschwellenwert & Schutz" > Abschnitt "Spoofing" > "Spoofintelligenz aktivieren" | [Portal](https://security.microsoft.com/antiphishing) · [Learn](https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about#spoof-settings) |
| `APH-02` | Defender-Portal > Bedrohungsrichtlinien > Antiphishing > `<Richtlinie>` > "Aktionen" > "Wenn die Nachricht durch Spoofintelligenz als Spoofing erkannt wird" | [Portal](https://security.microsoft.com/antiphishing) · [Learn](https://learn.microsoft.com/defender-office-365/recommended-settings-for-eop-and-office365#anti-phishing-policy-settings-for-all-cloud-mailboxes) |
| `APH-03` | Defender-Portal > Bedrohungsrichtlinien > Antiphishing > `<Richtlinie>` > "Phishingschwellenwert & Schutz" > "DMARC-Eintragsrichtlinie berücksichtigen, wenn die Nachricht als Spoofing erkannt wird" | [Portal](https://security.microsoft.com/antiphishing) · [Learn](https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about#spoof-protection-and-sender-dmarc-policies) |
| `APH-04` | Defender-Portal > Bedrohungsrichtlinien > Antiphishing > `<Richtlinie>` > "Aktionen" > "... und die DMARC-Richtlinie p=quarantine lautet" | [Portal](https://security.microsoft.com/antiphishing) · [Learn](https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about#spoof-protection-and-sender-dmarc-policies) |
| `APH-05` | Defender-Portal > Bedrohungsrichtlinien > Antiphishing > `<Richtlinie>` > "Aktionen" > "... und die DMARC-Richtlinie p=reject lautet" | [Portal](https://security.microsoft.com/antiphishing) · [Learn](https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about#spoof-protection-and-sender-dmarc-policies) |
| `APH-06` | Defender-Portal > Bedrohungsrichtlinien > Antiphishing > `<Richtlinie>` > "Aktionen" > Sicherheitstipps & Indikatoren > "(?) für nicht authentifizierte Absender anzeigen" und "Tag via anzeigen" | [Portal](https://security.microsoft.com/antiphishing) · [Learn](https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about#unauthenticated-sender-indicators) |
| `APH-07` | Defender-Portal > Bedrohungsrichtlinien > Antiphishing > `<Richtlinie>` > "Aktionen" > Sicherheitstipps & Indikatoren > "Sicherheitstipp für ersten Kontakt anzeigen" | [Portal](https://security.microsoft.com/antiphishing) · [Learn](https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about#first-contact-safety-tip) |
| `APH-08` | Defender-Portal > Bedrohungsrichtlinien > Antiphishing > `<Richtlinie>` > "Phishingschwellenwert & Schutz" > "Schwellenwert für Phishing-E-Mails" (1-4, nur mit Defender-Lizenz sichtbar) | [Portal](https://security.microsoft.com/antiphishing) · [Learn](https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about#phishing-email-thresholds-in-anti-phishing-policies-in-microsoft-defender-for-office-365) |
| `IMP-01` | Defender-Portal > Bedrohungsrichtlinien > Antiphishing > `<Richtlinie>` > "Phishingschwellenwert & Schutz" > Identitätswechsel > "Zu schützende Domänen aktivieren" > "Domänen einschließen, die ich besitze" | [Portal](https://security.microsoft.com/antiphishing) · [Learn](https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about#domain-impersonation-protection) |
| `IMP-02` | Defender-Portal > Bedrohungsrichtlinien > Antiphishing > `<Richtlinie>` > "Phishingschwellenwert & Schutz" > Identitätswechsel > "Benutzerdefinierte Domänen einschließen" (max. 50) | [Portal](https://security.microsoft.com/antiphishing) · [Learn](https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about#domain-impersonation-protection) |
| `IMP-03` | Defender-Portal > Bedrohungsrichtlinien > Antiphishing > `<Richtlinie>` > "Phishingschwellenwert & Schutz" > Identitätswechsel > "Benutzern das Schützen ermöglichen" > "Geschützte Benutzer verwalten" (max. 350) | [Portal](https://security.microsoft.com/antiphishing) · [Learn](https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about#user-impersonation-protection) |
| `IMP-04` | Defender-Portal > Bedrohungsrichtlinien > Antiphishing > `<Richtlinie>` > "Aktionen" > die drei Zeilen zu Benutzer-, Domänenidentitätswechsel und Postfachintelligenz (nicht "Keine Aktion anwenden") | [Portal](https://security.microsoft.com/antiphishing) · [Learn](https://learn.microsoft.com/defender-office-365/recommended-settings-for-eop-and-office365#impersonation-settings-in-anti-phishing-policies-in-microsoft-defender-for-office-365) |
| `IMP-05` | Defender-Portal > Bedrohungsrichtlinien > Antiphishing > `<Richtlinie>` > "Phishingschwellenwert & Schutz" > Identitätswechsel > "Postfachintelligenz aktivieren" und "Intelligenz für Identitätswechselschutz aktivieren" | [Portal](https://security.microsoft.com/antiphishing) · [Learn](https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about#mailbox-intelligence-impersonation-protection) |
| `IMP-06` | Defender-Portal > Bedrohungsrichtlinien > Antiphishing > `<Richtlinie>` > "Aktionen" > Sicherheitstipps & Indikatoren > die drei Identitätswechsel-Tipps | [Portal](https://security.microsoft.com/antiphishing) · [Learn](https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about#impersonation-safety-tips) |
| `IMP-07` | Defender-Portal > Bedrohungsrichtlinien > Antiphishing > `<Richtlinie>` > "Phishingschwellenwert & Schutz" > Identitätswechsel > "Vertrauenswürdige Absender und Domänen hinzufügen" (max. 1024) | [Portal](https://security.microsoft.com/antiphishing) · [Learn](https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about#trusted-senders-and-domains) |
| `IMP-08` | Defender-Portal > Einstellungen > E-Mail & Zusammenarbeit > Schutz von Prioritätskonten (Priority account protection); die Konten selbst werden im Microsoft 365 Admin Center markiert | [Portal](https://security.microsoft.com/securitysettings/priorityAccountProtection) · [Learn](https://learn.microsoft.com/defender-office-365/priority-accounts-turn-on-priority-account-protection#review-or-turn-on-priority-account-protection-in-the-microsoft-defender-portal) |

### Tenant Allow/Block List

| ID | Klickpfad | Direkt |
|:------|:------------------------------------------------------------------------------------------|:------|
| `TABL-01` | Defender-Portal > Bedrohungsrichtlinien > Mandanten-Zulassungs-/Sperrlisten. Zulassungen entstehen bevorzugt über Übermittlungen (Submissions), nicht von Hand | [Portal](https://security.microsoft.com/tenantAllowBlockList) · [Learn](https://learn.microsoft.com/defender-office-365/tenant-allow-block-list-about#allow-entries-in-the-tenant-allowblock-list) |
| `TABL-02` | Defender-Portal > Bedrohungsrichtlinien > Mandanten-Zulassungs-/Sperrlisten > Reiter "Domänen und E-Mail-Adressen" > Spalten "Aktion" und "Läuft ab am" | [Portal](https://security.microsoft.com/tenantAllowBlockList) · [Learn](https://learn.microsoft.com/defender-office-365/tenant-allow-block-list-email-spoof-configure#domains-and-email-addresses-in-the-tenant-allowblock-list) |
| `TABL-03` | Defender-Portal > Bedrohungsrichtlinien > Mandanten-Zulassungs-/Sperrlisten > Reiter "Gefälschte Absender"; ergänzend die Spoofintelligenz-Erkenntnisse | [Portal](https://security.microsoft.com/tenantAllowBlockList?viewid=SpoofItem) · [Learn](https://learn.microsoft.com/defender-office-365/tenant-allow-block-list-email-spoof-configure#spoofed-senders-in-the-tenant-allowblock-list) |
| `TABL-04` | Defender-Portal > Bedrohungsrichtlinien > Mandanten-Zulassungs-/Sperrlisten > Reiter "URLs" | [Portal](https://security.microsoft.com/tenantAllowBlockList) · [Learn](https://learn.microsoft.com/defender-office-365/tenant-allow-block-list-urls-configure#url-syntax-for-the-tenant-allowblock-list) |
| `TABL-05` | Defender-Portal > Bedrohungsrichtlinien > Mandanten-Zulassungs-/Sperrlisten > Reiter "Dateien" (nur SHA256, Zulassungen nur über Übermittlungen) | [Portal](https://security.microsoft.com/tenantAllowBlockList) · [Learn](https://learn.microsoft.com/defender-office-365/tenant-allow-block-list-files-configure#create-block-entries-for-files) |
| `TABL-06` | Defender-Portal > Bedrohungsrichtlinien > Mandanten-Zulassungs-/Sperrlisten > Reiter "IP-Adressen" (ausschließlich IPv6) | [Portal](https://security.microsoft.com/tenantAllowBlockList) · [Learn](https://learn.microsoft.com/defender-office-365/tenant-allow-block-list-ip-addresses-configure#create-block-entries-for-ipv6-addresses) |

### Anti-Malware, Safe Attachments, Safe Links

| ID | Klickpfad | Direkt |
|:------|:------------------------------------------------------------------------------------------|:------|
| `AMW-01` | Defender-Portal > Bedrohungsrichtlinien > Antischadsoftware > `<Richtlinie>` > "Schutzeinstellungen" > "Aktivieren Sie den allgemeinen Anlagenfilter" > "Dateitypen auswählen" | [Portal](https://security.microsoft.com/antimalwarev2) · [Learn](https://learn.microsoft.com/defender-office-365/anti-malware-protection-about#common-attachments-filter-in-anti-malware-policies) |
| `AMW-02` | Defender-Portal > Bedrohungsrichtlinien > Antischadsoftware > `<Richtlinie>` > "Schutzeinstellungen" > "Automatische Null-Stunden-Bereinigung für Schadsoftware aktivieren" | [Portal](https://security.microsoft.com/antimalwarev2) · [Learn](https://learn.microsoft.com/defender-office-365/anti-malware-protection-about#zero-hour-auto-purge-zap-in-anti-malware-policies) |
| `AMW-03` | Defender-Portal > Bedrohungsrichtlinien > Antischadsoftware > `<Richtlinie>` > "Schutzeinstellungen" > "Wenn diese Dateitypen gefunden werden" (NDR ablehnen vs. Quarantäne) | [Portal](https://security.microsoft.com/antimalwarev2) · [Learn](https://learn.microsoft.com/defender-office-365/anti-malware-protection-about#common-attachments-filter-in-anti-malware-policies) |
| `AMW-04` | Defender-Portal > Bedrohungsrichtlinien > Antischadsoftware > `<Richtlinie>` > "Schutzeinstellungen" > "Quarantänerichtlinie" (Default AdminOnlyAccessPolicy) | [Portal](https://security.microsoft.com/antimalwarev2) · [Learn](https://learn.microsoft.com/defender-office-365/anti-malware-protection-about#quarantine-policies-in-anti-malware-policies) |
| `AMW-05` | Defender-Portal > Bedrohungsrichtlinien > Antischadsoftware > `<Richtlinie>` > "Schutzeinstellungen" > Benachrichtigungen > Administratorbenachrichtigungen | [Portal](https://security.microsoft.com/antimalwarev2) · [Learn](https://learn.microsoft.com/defender-office-365/anti-malware-protection-about#admin-notifications-in-anti-malware-policies) |
| `SA-01` | Defender-Portal > Bedrohungsrichtlinien > Sichere Anlagen > `<Richtlinie>` > "Einstellungen" > "Safe Attachments-Antwort bei unbekannter Schadsoftware" = Blockieren | [Portal](https://security.microsoft.com/safeattachmentv2) · [Learn](https://learn.microsoft.com/defender-office-365/safe-attachments-about#safe-attachments-policy-settings) |
| `SA-02` | Defender-Portal > Bedrohungsrichtlinien > Sichere Anlagen > `<Richtlinie>` > "Einstellungen" > "Umleiten von Nachrichten mit erkannten Anlagen" (wirkt laut Doku nur bei Aktion "Überwachen") | [Portal](https://security.microsoft.com/safeattachmentv2) · [Learn](https://learn.microsoft.com/defender-office-365/safe-attachments-about#safe-attachments-policy-settings) |
| `SA-03` | Defender-Portal > Bedrohungsrichtlinien > Sichere Anlagen > Zahnrad "Globale Einstellungen" > "Defender for Office 365 für SharePoint, OneDrive und Microsoft Teams aktivieren". Download-Sperre zusätzlich nur per SharePoint-PowerShell | [Portal](https://security.microsoft.com/safeattachmentv2) · [Learn](https://learn.microsoft.com/defender-office-365/safe-attachments-for-spo-odfb-teams-configure#step-1-use-the-microsoft-defender-portal-to-turn-on-safe-attachments-for-sharepoint-onedrive-and-microsoft-teams) |
| `SA-04` | Defender-Portal > Bedrohungsrichtlinien > Sichere Anlagen > Zahnrad "Globale Einstellungen" > "Safe Documents für Office-Clients aktivieren" plus Durchklicken verbieten | [Portal](https://security.microsoft.com/safeattachmentv2) · [Learn](https://learn.microsoft.com/defender-office-365/safe-documents-in-e5-plus-security-about#use-the-microsoft-defender-portal-to-configure-safe-documents) |
| `SA-05` | Defender-Portal > Bedrohungsrichtlinien > Sichere Anlagen > `<Richtlinie>` > "Einstellungen" > "Safe Attachments-Antwort" = "Dynamische Übermittlung (Vorschau von Nachrichten)" | [Portal](https://security.microsoft.com/safeattachmentv2) · [Learn](https://learn.microsoft.com/defender-office-365/safe-attachments-about#dynamic-delivery-in-safe-attachments-policies) |
| `SL-01` | Defender-Portal > Bedrohungsrichtlinien > Sichere Links > `<Richtlinie>` > "URL- & Klickschutzeinstellungen" > Abschnitte "E-Mail", "Teams" und "Office 365-Apps" je auf Ein | [Portal](https://security.microsoft.com/safelinksv2) · [Learn](https://learn.microsoft.com/defender-office-365/safe-links-about#safe-links-settings-for-email-messages) |
| `SL-02` | Defender-Portal > Bedrohungsrichtlinien > Sichere Links > `<Richtlinie>` > "URL- & Klickschutzeinstellungen" > E-Mail > "Sichere Links auf E-Mail-Nachrichten anwenden, die innerhalb der Organisation gesendet werden" | [Portal](https://security.microsoft.com/safelinksv2) · [Learn](https://learn.microsoft.com/defender-office-365/safe-links-about#safe-links-settings-for-email-messages) |
| `SL-03` | Defender-Portal > Bedrohungsrichtlinien > Sichere Links > `<Richtlinie>` > "URL- & Klickschutzeinstellungen" > "URL-Überprüfung in Echtzeit ..." mit der Unteroption "Warten, bis die URL-Überprüfung abgeschlossen ist" | [Portal](https://security.microsoft.com/safelinksv2) · [Learn](https://learn.microsoft.com/defender-office-365/safe-links-about#safe-links-settings-for-email-messages) |
| `SL-04` | Defender-Portal > Bedrohungsrichtlinien > Sichere Links > `<Richtlinie>` > "URL- & Klickschutzeinstellungen" > Klickschutzeinstellungen > "Benutzern das Durchklicken zur ursprünglichen URL erlauben" ausschalten | [Portal](https://security.microsoft.com/safelinksv2) · [Learn](https://learn.microsoft.com/defender-office-365/safe-links-about#click-protection-settings-in-safe-links-policies) |
| `SL-05` | Defender-Portal > Bedrohungsrichtlinien > Sichere Links > `<Richtlinie>` > "URL- & Klickschutzeinstellungen" > E-Mail > "URLs nicht umschreiben, Überprüfungen nur über die SafeLinks-API" muss AUS sein | [Portal](https://security.microsoft.com/safelinksv2) · [Learn](https://learn.microsoft.com/defender-office-365/safe-links-about#safe-links-settings-for-email-messages) |
| `SL-06` | Defender-Portal > Bedrohungsrichtlinien > Sichere Links > `<Richtlinie>` > "URL- & Klickschutzeinstellungen" > "Die folgenden URLs in E-Mails nicht umschreiben" > "Nicht umzuschreibende URLs verwalten" (pro Richtlinie, nicht global) | [Portal](https://security.microsoft.com/safelinksv2) · [Learn](https://learn.microsoft.com/defender-office-365/safe-links-about#entry-syntax-for-the-do-not-rewrite-the-following-urls-list) |
| `SL-07` | Defender-Portal > Bedrohungsrichtlinien > Sichere Links > `<Richtlinie>` > "URL- & Klickschutzeinstellungen" > "Organisationsbranding anzeigen" (Logo aus dem M365-Organisationsdesign) sowie > "Benachrichtigung" > eigener Text (max. 200 Zeichen) | [Portal](https://security.microsoft.com/safelinksv2) · [Learn](https://learn.microsoft.com/defender-office-365/safe-links-about#click-protection-settings-in-safe-links-policies) |

### Quarantäne und Advanced Delivery

| ID | Klickpfad | Direkt |
|:------|:------------------------------------------------------------------------------------------|:------|
| `ADV-01` | Defender-Portal > Bedrohungsrichtlinien > Erweiterte Zustellung (Advanced delivery) > Reiter "SecOps-Postfach" (nur echte Postfächer, keine Verteilergruppen) | [Portal](https://security.microsoft.com/advanceddelivery) · [Learn](https://learn.microsoft.com/defender-office-365/advanced-delivery-policy-configure#use-the-microsoft-defender-portal-to-configure-secops-mailboxes-in-the-advanced-delivery-policy) |
| `ADV-02` | Defender-Portal > Bedrohungsrichtlinien > Erweiterte Zustellung > Reiter "Phishingsimulation" (Domäne + Sende-IP; IPv6 nur per PowerShell) | [Portal](https://security.microsoft.com/advanceddelivery) · [Learn](https://learn.microsoft.com/defender-office-365/advanced-delivery-policy-configure#use-the-microsoft-defender-portal-to-configure-non-microsoft-phishing-simulations-in-the-advanced-delivery-policy) |
| `ADV-03` | Exchange Admin Center > Nachrichtenfluss > Regeln: jede Regel auf "Spamfilterung umgehen" (SCL -1) prüfen. Solche Bypässe gehören in die Erweiterte Zustellung oder die TABL | [Portal](https://admin.exchange.microsoft.com/#/transportrules) · [Learn](https://learn.microsoft.com/defender-office-365/create-safe-sender-lists-in-office-365#use-mail-flow-rules) |
| `QUA-01` | Defender-Portal > Bedrohungsrichtlinien > Quarantänerichtlinie (definieren); zugewiesen wird sie je Verdict in Antispam > `<Richtlinie>` > "Aktionen", in Antiphishing > "Aktionen" sowie in Antischadsoftware und Sichere Anlagen | [Portal](https://security.microsoft.com/quarantinePolicies) · [Learn](https://learn.microsoft.com/defender-office-365/quarantine-policies#assign-quarantine-policies-in-supported-policies-in-the-microsoft-defender-portal) |
| `QUA-02` | Defender-Portal > Bedrohungsrichtlinien > Quarantänerichtlinie > "Benutzerdefinierte Richtlinie hinzufügen" > "Zugriff auf Empfängernachrichten" > "Spezifischen Zugriff festlegen (Erweitert)" > Quarantänebenachrichtigung aktivieren | [Portal](https://security.microsoft.com/quarantinePolicies) · [Learn](https://learn.microsoft.com/defender-office-365/quarantine-policies#step-1-create-quarantine-policies-in-the-microsoft-defender-portal) |
| `QUA-03` | Defender-Portal > Bedrohungsrichtlinien > Quarantänerichtlinie > Zahnrad "Globale Einstellungen" > "Spambenachrichtigung für Endbenutzer senden alle" (4 Stunden / täglich / wöchentlich) | [Portal](https://security.microsoft.com/quarantinePolicies) · [Learn](https://learn.microsoft.com/defender-office-365/quarantine-policies#customize-all-quarantine-notifications) |
| `QUA-04` | Defender-Portal > Bedrohungsrichtlinien > Quarantänerichtlinie > Zahnrad "Globale Einstellungen" > Absenderadresse, Anzeigename, Betreff, Haftungsausschluss und Firmenlogo | [Portal](https://security.microsoft.com/quarantinePolicies) · [Learn](https://learn.microsoft.com/defender-office-365/quarantine-policies#customize-all-quarantine-notifications) |

### E-Mail-Authentifizierung

| ID | Klickpfad | Direkt |
|:------|:------------------------------------------------------------------------------------------|:------|
| `AUTH-01` | Öffentliches DNS beim Domain-Hoster: TXT auf `<domäne>` mit v=spf1 include:spf.protection.outlook.com -all, je akzeptierter Domäne genau einer | [Learn](https://learn.microsoft.com/defender-office-365/email-authentication-spf-configure#spf-txt-records-for-custom-domains-in-microsoft-365) |
| `AUTH-02` | Öffentliches DNS: TXT auf `<domäne>`, Endqualifier -all statt ~all | [Learn](https://learn.microsoft.com/defender-office-365/email-authentication-spf-configure#syntax-for-spf-txt-records) |
| `AUTH-03` | Öffentliches DNS: für geparkte Domänen TXT v=spf1 -all auf `<domäne>` plus TXT v=DMARC1; p=reject; auf _dmarc.`<domäne>`. Gilt auch für die onmicrosoft.com-Domäne | [Learn](https://learn.microsoft.com/defender-office-365/email-authentication-spf-configure#scenario-parked-domains) |
| `AUTH-04` | Defender-Portal > Bedrohungsrichtlinien > E-Mail-Authentifizierungseinstellungen > DKIM > Domäne wählen, Umschalter aktivieren. Vorher die beiden CNAMEs selector1/2._domainkey.`<domäne>` ins öffentliche DNS | [Portal](https://security.microsoft.com/authentication?viewid=DKIM) · [Learn](https://learn.microsoft.com/defender-office-365/email-authentication-dkim-configure#use-the-defender-portal-to-enable-dkim-signing-of-outbound-messages-using-a-custom-domain) |
| `AUTH-05` | Die Bitlänge steuert nur PowerShell: Rotate-DkimSigningConfig -KeySize 2048. Im Portal rotiert man über E-Mail-Authentifizierungseinstellungen > DKIM > Domänenzeile ANKLICKEN > Flyout > "DKIM-Schlüssel rotieren" | [Portal](https://security.microsoft.com/authentication?viewid=DKIM) · [Learn](https://learn.microsoft.com/defender-office-365/email-authentication-dkim-configure#use-exchange-online-powershell-to-rotate-the-dkim-keys-for-a-domain-and-change-the-bit-depth) |
| `AUTH-06` | Öffentliches DNS: TXT auf _dmarc.`<domäne>` mit v=DMARC1; p=...; rua=mailto:... Für die onmicrosoft.com-Domäne über Microsoft 365 Admin Center > Einstellungen > Domänen | [Portal](https://admin.microsoft.com/Adminportal/Home#/Domains) · [Learn](https://learn.microsoft.com/defender-office-365/email-authentication-dmarc-configure#set-up-dmarc-for-active-custom-domains-in-microsoft-365) |
| `AUTH-07` | Öffentliches DNS: rua=mailto: im TXT auf _dmarc.`<domäne>`, Ziel ein dediziertes Shared Mailbox, kein Benutzerpostfach | [Learn](https://learn.microsoft.com/defender-office-365/email-authentication-dmarc-configure#best-practices-for-dmarc-reports) |
| `AUTH-08` | Defender-Portal > Bedrohungsrichtlinien > E-Mail-Authentifizierungseinstellungen > ARC; alternativ Set-ArcConfig -ArcTrustedSealers | [Portal](https://security.microsoft.com/authentication) · [Learn](https://learn.microsoft.com/defender-office-365/email-authentication-arc-configure#use-the-microsoft-defender-portal-to-add-trusted-arc-sealers) |
| `AUTH-09` | Öffentliches DNS: TXT auf _mta-sts.`<domäne>` PLUS extern gehostete Policy-Datei unter https://mta-sts.`<domäne>`/.well-known/mta-sts.txt. Exchange Online hostet die Datei nicht. Ausgehend ist MTA-STS immer aktiv | [Learn](https://learn.microsoft.com/exchange/security-and-compliance/enhance-mail-flow-using-strict-transport-security#adopt-mta-sts-for-your-domain) |
| `AUTH-10` | Ausgehend standardmäßig an. Eingehend per Exchange Online PowerShell: Enable-DnssecForVerifiedDomain, dann Enable-SmtpDaneInbound, danach den ausgegebenen MX-Wert beim Hoster setzen und die Delegation DNSSEC-signieren | [Learn](https://learn.microsoft.com/exchange/security-and-compliance/how-dane-secures-email#inbound-smtp-dane-with-dnssec) |
| `AUTH-11` | Kein Portal: SPF, DKIM und DMARC ausgerichtet auf die 5322.From-Domäne plus auffindbarer Abmeldemechanismus. Greift ab 5.000 Nachrichten/Tag an Microsoft-Consumer-Dienste, sonst NDR 550 5.7.515 | [Learn](https://learn.microsoft.com/defender-office-365/external-senders-policies-practices-guidelines) |
| `AUTH-12` | Öffentliches DNS: TXT auf default._bimi.`<domäne>`. Voraussetzung ist DMARC p=quarantine oder p=reject. Exchange Online wertet BIMI derzeit nicht aus | [Learn](https://learn.microsoft.com/dynamics365/customer-insights/journeys/bimi-support) |

### Enhanced Filtering und externe Kennzeichnung

| ID | Klickpfad | Direkt |
|:------|:------------------------------------------------------------------------------------------|:------|
| `EF-01` | Defender-Portal > Bedrohungsrichtlinien > Erweiterte Filterung (Enhanced filtering) > Eintrag des Inbound-Connectors > zu überspringende IPs bzw. "Letzte IP überspringen" | [Portal](https://security.microsoft.com/skiplisting) · [Learn](https://learn.microsoft.com/exchange/mail-flow-best-practices/use-connectors-to-configure-mail-flow/enhanced-filtering-for-connectors#use-the-microsoft-defender-portal-to-configure-enhanced-filtering-for-connectors-on-an-inbound-connector) |
| `EF-02` | Exchange Admin Center > Nachrichtenfluss > Regeln: SCL-(-1)-Regeln für Nachrichten über diesen Connector abschalten | [Portal](https://admin.exchange.microsoft.com/#/transportrules) · [Learn](https://learn.microsoft.com/exchange/mail-flow-best-practices/use-connectors-to-configure-mail-flow/enhanced-filtering-for-connectors#what-do-you-need-to-know-before-you-begin) |
| `EXT-01` | Kein Portal, nur Exchange Online PowerShell: Set-ExternalInOutlook -Enabled $true, Ausnahmen über -AllowList (max. 200). Wirkung erst nach 24-48 Stunden | [Learn](https://learn.microsoft.com/powershell/module/exchangepowershell/set-externalinoutlook) |
| `EXT-02` | Entscheidung zwischen nativem Tag (PowerShell), First-Contact-Tipp (Antiphishing-Richtlinie) und Banner per Transportregel. Nicht mehrere gleichzeitig | [Portal](https://security.microsoft.com/antiphishing) · [Learn](https://learn.microsoft.com/defender-office-365/anti-phishing-policies-about#first-contact-safety-tip) |
| `EXT-03` | Exchange Admin Center > Nachrichtenfluss > Regeln: Regel mit "Absender ausserhalb der Organisation" und Aktion "Haftungsausschluss voranstellen", Fallback-Aktion und Ausnahmen setzen | [Portal](https://admin.exchange.microsoft.com/#/transportrules) · [Learn](https://learn.microsoft.com/exchange/security-and-compliance/mail-flow-rules/disclaimers-signatures-footers-or-headers#use-the-eac-to-add-a-disclaimer-or-other-email-header-or-footer) |

### Reporting, Hybrid, Betrieb

| ID | Klickpfad | Direkt |
|:------|:------------------------------------------------------------------------------------------|:------|
| `HYB-01` | Nur On-Premises: Exchange Management Shell (Get-ExchangeServer, ExSetup /Version) gegen die Supportability-Matrix. Exchange 2016 und 2019 sind seit 14.10.2025 out of support | [Learn](https://learn.microsoft.com/exchange/plan-and-deploy/supportability-matrix#supported-versions-and-builds) |
| `HYB-02` | On-Premises: Skript ConfigureExchangeHybridApplication.ps1 bzw. HCW. Kontrolle der Dienstprinzipal-Anmeldungen in Microsoft Entra ID > Überwachung > Anmeldeprotokolle | [Portal](https://entra.microsoft.com) · [Learn](https://learn.microsoft.com/exchange/hybrid-deployment/deploy-dedicated-hybrid-app#configure-the-dedicated-exchange-hybrid-application) |
| `HYB-03` | Nur On-Premises: "Default Frontend `<Server>`" nicht für anonymes Relay öffnen, stattdessen dedizierter Receive-Connector mit RemoteIpRanges | [Learn](https://learn.microsoft.com/exchange/mail-flow/connectors/allow-anonymous-relay#step-1-create-a-dedicated-receive-connector-for-anonymous-relay) |
| `HYB-04` | Perimeter- und Host-Firewall: ausgehend SMTP 25/587 nur von den Transport-Servern, POP3 110/995 und IMAP4 143/993 nach aussen sperren | [Learn](https://learn.microsoft.com/exchange/plan-and-deploy/deployment-ref/network-ports#network-ports-required-for-mail-flow) |
| `HYB-05` | Exchange Admin Center > Einstellungen > E-Mail-Fluss > "SMTP AUTH-Protokoll für Ihre Organisation deaktivieren". Ausnahmen je Postfach über Set-CASMailbox | [Portal](https://admin.exchange.microsoft.com/#/settings) · [Learn](https://learn.microsoft.com/exchange/clients-and-mobile-in-exchange-online/authenticated-client-smtp-submission#disable-smtp-auth-in-your-organization) |
| `HYB-06` | Exchange Admin Center > Empfänger > Postfächer > `<Postfach>` > "E-Mail-App-Einstellungen verwalten"; organisationsweit über Set-CASMailbox bzw. Set-CASMailboxPlan für neue Postfächer | [Portal](https://admin.exchange.microsoft.com/#/mailboxes) · [Learn](https://learn.microsoft.com/exchange/recipients-in-exchange-online/manage-user-mailboxes/managing-email-apps-for-user-mailboxes#use-exchange-online-powershell-to-enable-or-disable-email-apps) |
| `HYB-07` | Exchange Admin Center > Empfänger > Postfächer > `<Postfach>` > "Nachrichtengrößenbeschränkung verwalten"; die Dienstgrenzwerte sind harte Limits | [Portal](https://admin.exchange.microsoft.com/#/mailboxes) · [Learn](https://learn.microsoft.com/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits#message-limits) |
| `HYB-08` | Exchange Admin Center > Nachrichtenfluss > Connectors > ausgehender Connector: "Immer eine TLS-gesicherte Verbindung verwenden" plus Zertifikatsprüfung | [Portal](https://admin.exchange.microsoft.com/#/connectors) · [Learn](https://learn.microsoft.com/exchange/mail-flow-best-practices/use-connectors-to-configure-mail-flow/set-up-connectors-to-secure-mail-sent-to-partner-organization#for-new-eac) |
| `HYB-09` | Exchange Admin Center > Nachrichtenfluss > Connectors: alle ein- und ausgehenden Connectors inventarisieren (HCW vs. manuell, IP-Bereiche, Zertifikate) | [Portal](https://admin.exchange.microsoft.com/#/connectors) · [Learn](https://learn.microsoft.com/exchange/mail-flow-best-practices/use-connectors-to-configure-mail-flow/use-connectors-to-configure-mail-flow#when-do-i-need-a-connector) |
| `HYB-10` | Exchange Admin Center > Nachrichtenfluss > Akzeptierte Domänen: Typ "Autorisierend" schaltet DBEB ein, "Internes Relay" schaltet es ab | [Portal](https://admin.exchange.microsoft.com/#/accepteddomains) · [Learn](https://learn.microsoft.com/exchange/mail-flow-best-practices/use-directory-based-edge-blocking#configure-dbeb) |
| `OPS-01` | Defender-Portal > Berechtigungen: Rollengruppen "E-Mail & Zusammenarbeit" bzw. Defender XDR Unified RBAC; dazu die Entra-Rollen | [Portal](https://security.microsoft.com/emailandcollabpermissions) · [Learn](https://learn.microsoft.com/defender-office-365/scc-permissions#role-groups-in-microsoft-defender-for-office-365-and-microsoft-purview) |
| `OPS-02` | Kein Portal: Admin-Arbeitsplatz. Modul ExchangeOnlineManagement, Windows PowerShell 5.1 mit .NET 4.7.2+ oder PowerShell 7 | [Learn](https://learn.microsoft.com/powershell/exchange/exchange-online-powershell-v2) |
| `OPS-03` | Baseline per PowerShell exportieren; ergänzend Konfigurationsanalyse > Reiter "Configuration drift analysis and history" (setzt Unified Auditing voraus) | [Portal](https://security.microsoft.com/configurationAnalyzer) · [Learn](https://learn.microsoft.com/defender-office-365/configuration-analyzer-for-security-policies#configuration-drift-analysis-and-history-tab-in-the-configuration-analyzer) |
| `OPS-04` | Nachweis über Microsoft Purview > Audit (Unified Audit Log); Nachkontrolle als wiederkehrenden Termin setzen | [Portal](https://purview.microsoft.com) · [Learn](https://learn.microsoft.com/purview/audit-log-enable-disable) |
| `REP-01` | Defender-Portal > Berichte > E-Mail & Zusammenarbeit; Zeitplan je Bericht über "Create schedule" | [Portal](https://security.microsoft.com/emailandcollabreport) · [Learn](https://learn.microsoft.com/defender-office-365/reports-email-security) |
| `REP-02` | Defender-Portal > Richtlinien & Regeln > Warnungsrichtlinie; Empfänger je Richtlinie im Feld "Email recipients" | [Portal](https://security.microsoft.com/alertpolicies) · [Learn](https://learn.microsoft.com/defender-xdr/alert-policies#alert-policy-settings) |
| `REP-03` | Defender-Portal > Einstellungen > E-Mail & Zusammenarbeit > Benutzerdefinierte Meldungen (User reported settings); alternativ *-ReportSubmissionPolicy in PowerShell | [Portal](https://security.microsoft.com/securitysettings/userSubmission) · [Learn](https://learn.microsoft.com/defender-office-365/submissions-user-reported-messages-custom-mailbox#use-the-microsoft-defender-portal-to-configure-user-reported-settings) |
| `REP-04` | Kein Portalpunkt, sondern Skripte: auf Get-MessageTraceV2 / Get-MessageTraceDetailV2 umstellen. Interaktiv führt der Defender-Portal-Eintrag nur ins EAC | [Portal](https://admin.exchange.microsoft.com/#/messagetrace) · [Learn](https://learn.microsoft.com/powershell/module/exchangepowershell/get-messagetracev2) |

### Teams, Wirksamkeitsnachweis, Governance, P2

| ID | Klickpfad | Direkt |
|:------|:------------------------------------------------------------------------------------------|:------|
| `GOV-01` | Kein Microsoft-Portalpunkt: organisatorisch - Datenschutz-Folgenabschätzung, Verfahrensverzeichnis und Mitbestimmung für Quarantäne-Einsicht und Explorer-Preview | [Learn](https://learn.microsoft.com/compliance/regulatory/gdpr#data-protection-impact-assessment) |
| `P2-01` | Defender-Portal > Untersuchungen (AIR, nur MDO Plan 2; setzt aktiviertes Audit-Logging voraus) | [Portal](https://security.microsoft.com/airinvestigation) · [Learn](https://learn.microsoft.com/defender-office-365/air-about#the-overall-flow-of-air) |
| `TEAMS-01` | Defender-Portal > Einstellungen > E-Mail & Zusammenarbeit > Microsoft Teams-Schutz (NICHT mehr unter Bedrohungsrichtlinien) > ZAP für Teams, Quarantänerichtlinie, Ausnahmen | [Portal](https://security.microsoft.com/securitysettings/teamsProtectionPolicy) · [Learn](https://learn.microsoft.com/defender-office-365/mdo-support-teams-about#configure-zap-for-teams-protection-in-defender-for-office-365) |
| `VER-01` | Defender-Portal > Überprüfen > Quarantäne, Filter Richtlinientyp = Antischadsoftware-Richtlinie | [Portal](https://security.microsoft.com/quarantine) · [Learn](https://learn.microsoft.com/defender-office-365/anti-malware-protection-about#common-attachments-filter-in-anti-malware-policies) |
| `VER-02` | Kein Portal: Outlook > Datei > Eigenschaften > Internetkopfzeilen (neues Outlook/OWA: Nachricht > ... > Details anzeigen). Ausgewertet werden X-Forefront-Antispam-Report, X-Microsoft-Antispam und Authentication-Results | [Learn](https://learn.microsoft.com/defender-office-365/message-headers-eop-mdo#x-forefront-antispam-report-message-header-fields) |
| `VER-03` | Defender-Portal > E-Mail & Zusammenarbeit > Explorer (P2) bzw. Echtzeiterkennungen (P1), Ansichten Phish und URL-Klicks | [Portal](https://security.microsoft.com/threatexplorerv3) · [Learn](https://learn.microsoft.com/defender-office-365/threat-explorer-real-time-detections-about) |

## Was an diesen Links unsicher ist

Sämtliche Deep-Links ins Defender-Portal stammen aus der Microsoft-Dokumentation — Microsoft nennt sie in den Learn-Artikeln ausdrücklich. Alle Learn-Sprungmarken sind gegen die Quelldateien der MicrosoftDocs-Repositories geprüft, also nicht geraten. Drei Einschränkungen gehören dazu:

- Für das **Exchange Admin Center** dokumentiert Microsoft nur wenige Direktlinks. Nicht belegt und aus der Praxis übernommen sind die von `ASO-07`, `ASO-08`, `HYB-06`, `HYB-07`, `HYB-10`. Führt einer ins Leere, steht der Klickpfad daneben. Der belegte Einstieg ist [admin.exchange.microsoft.com](https://admin.exchange.microsoft.com); der EAC-Übersichtsartikel nennt inzwischen zusätzlich `admin.cloud.microsoft/exchange`.
- Vier Prüfpunkte haben **keinen passgenauen Learn-Artikel**; der Link führt auf das Nächstliegende:

    - `AUTH-11` — Learn deckt die High-Volume-Sender-Anforderungen nicht ab; belastbar sind nur der Support-Artikel zu NDR 550 5.7.515 und der MDO-Blogpost vom 30.04.2025.
    - `AUTH-12` — Für BIMI gibt es keinen EOP/MDO-Artikel. Die verlinkte Seite gehört zu Dynamics 365 und stellt fest, dass Exchange Online BIMI nicht auswertet.
    - `GOV-01` — Mitbestimmung nach BetrVG ist bei Microsoft nicht dokumentiert; verlinkt ist der DSGVO-Artikel.
    - `HYB-04` — Kein Firewall-Artikel bei Microsoft; verlinkt ist die Exchange-Portreferenz.

- Die Skripte sind syntaktisch validiert, auf PS-5.1-Kompatibilität geprüft und mit Mock-Cmdlets end-to-end durchgelaufen. **Gegen einen Produktivtenant getestet wurden sie nicht** — der erste Lauf beim Kunden gehört in den Lesemodus.

## Policy-Inventar: warum manche Befunde nur „Info" sind

Eine Schutzrichtlinie in EOP besteht immer aus **zwei** Objekten: der *Policy* mit den Werten und der *Rule*, die festlegt, für wen sie gilt. Das Portal legt beides gemeinsam an, PowerShell nicht — daraus entstehen zwei Zustände, die wie funktionierende Richtlinien aussehen:

| Status | Bedeutung | Wird bewertet? |
|---|---|---|
| `aktiv - Prio n` | wird auf Empfänger angewendet | ja |
| `Default-Policy` | greift immer, hat keine Regel | ja |
| `Regel deaktiviert` | Policy sichtbar, wirkt aber auf niemanden | nein, nur Info |
| `ohne Regel` | im Defender-Portal unsichtbar, wirkt auf niemanden | nein, nur Info |
| `Preset - nicht editierbar` | von Microsoft verwaltet | nein, nur Info |

Eine Richtlinie, die auf niemanden angewendet wird, kann keinen Mangel darstellen — ihre Werte gegen die Baseline zu prüfen erzeugt nur Rauschen. Das Audit weist solche Objekte im Abschnitt `INV-*` mit ihrem Status aus und bewertet sie als *Info*. Prüfpunkt `PRE-05` greift das auf: verwaiste und stillgelegte Policies gehören reaktiviert, gelöscht oder als bewusst stillgelegt dokumentiert.

## Was die Skripte bewusst nicht anfassen

Diese Punkte gehören in die manuelle Bewertung, weil automatisches Setzen den Mailfluss brechen kann:

- Allow-Listen der Anti-Spam-Richtlinie (`ASI-10`) und die IP Allow List (`CF-01`)
- Bestehende Einträge der Tenant Allow/Block List (`TABL-01` bis `TABL-06`)
- Transportregeln mit Filter-Bypass (`ADV-03`)
- Enhanced Filtering (`EF-01`) — die Entscheidung hängt am MX-Ziel
- Alles im DNS: SPF, DKIM-Rotation, DMARC, MTA-STS, DANE (`AUTH-01` bis `AUTH-12`)
- On-Premises und Firewall (`HYB-01` bis `HYB-10`)

