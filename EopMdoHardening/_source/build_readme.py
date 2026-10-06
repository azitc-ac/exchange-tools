# -*- coding: utf-8 -*-
"""README fuer den Repo-Unterordner: Kuerzel-Legende und vollstaendige Fundort-Tabelle.

Aufruf: python3 build_readme.py  (im Ordner _source)

links.py ist die einzige Stelle, an der die Fundorte gepflegt werden. Die README wird
daraus erzeugt und nicht von Hand bearbeitet - sonst geht die Aenderung beim naechsten
Lauf verloren. Dieselben Daten speisen die Spalten Fundort / Portal oeffnen / Microsoft
Learn der Checkliste, Anhang E des Guides und die Tabelle $script:Links in
Invoke-EopAudit.ps1. Test-Links.ps1 meldet, wenn README und Audit-Skript auseinanderlaufen.
"""
import io, re, sys, collections
sys.path.insert(0, '.')
from links import LINKS, PREFIXES, UNDOCUMENTED, WEAK_LEARN, ALIASES

# Transliteration zurueck in echtes Deutsch (links.py ist ASCII, damit es in
# die PowerShell-Skripte passt; in Markdown wollen wir Umlaute).
UML = [
    ('Ausschluesse','Ausschlüsse'), ('Praezedenz','Präzedenz'), ('Prioritaet','Priorität'),
    ('Enthaelt','Enthält'), ('Laendern','Ländern'), ('gehoert','gehört'), ('Erhoehen','Erhöhen'),
    ('Ueberpruefungen','Überprüfungen'), ('Ueberpruefung','Überprüfung'), ('Ueberpruefen','Überprüfen'),
    ('urspruenglichen','ursprünglichen'), ('Quarantaene','Quarantäne'), ('quarantaene','quarantäne'),
    ('Uebermittlungen','Übermittlungen'), ('Uebermittlung','Übermittlung'),
    ('Gefaelschte','Gefälschte'), ('Domaenen','Domänen'), ('domaenen','domänen'),
    ('Domaene','Domäne'), ('domaene','domäne'), ('Identitaetswechsel','Identitätswechsel'),
    ('Vertrauenswuerdige','Vertrauenswürdige'), ('Geschuetzte','Geschützte'),
    ('geschuetzte','geschützte'), ('schuetzende','schützende'), ('Schuetzen','Schützen'),
    ('ermoeglichen','ermöglichen'), ('einschliessen','einschließen'),
    ('Prioritaetskonten','Prioritätskonten'), ('Schluessel','Schlüssel'), ('Bitlaenge','Bitlänge'),
    ('Oeffentliches','Öffentliches'), ('oeffentlichen','öffentlichen'), ('oeffentliche','öffentliche'),
    ('taeglich','täglich'), ('Nachrichtengroessenbeschraenkung','Nachrichtengrößenbeschränkung'),
    ('Empfaenger','Empfänger'), ('Postfaecher','Postfächer'), ('Remotedomaenen','Remotedomänen'),
    ('naechsten','nächsten'), ('ausschliesslich','ausschließlich'), ('Massnahmen','Maßnahmen'),
    ('beruecksichtigen','berücksichtigen'), ('Einschraenkung','Einschränkung'),
    ('Eingeschraenkte','Eingeschränkte'), ('ueberspringende','überspringende'),
    ('ueberspringen','überspringen'), ('Zertifikatspruefung','Zertifikatsprüfung'),
    ('waehlen','wählen'), ('pruefen','prüfen'), ('Pruefung','Prüfung'), ('Pruefpunkt','Prüfpunkt'),
    ('Haken','Haken'), ('Eintraege','Einträge'), ('Datenschutz-Folgenabschaetzung','Datenschutz-Folgenabschätzung'),
    ('Abmeldemechanismus','Abmeldemechanismus'), ('aufbewahren','aufbewahren'),
    ('Spameigenschaften','Spameigenschaften'), ('Schwellenwert','Schwellenwert'),
    ('Massen-E-Mail','Massen-E-Mail'), ('Antischadsoftware','Antischadsoftware'),
    ('Anlagenfilter','Anlagenfilter'), ('Null-Stunden','Null-Stunden'),
    ('verfuegbar','verfügbar'), ('ueberwiegend','überwiegend'), ('Zugriffsrechte','Zugriffsrechte'),
    ('fuer','für'), ('Fuer','Für'), ('ueberschritten','überschritten'),
    ('Domaenenidentitaetswechsel','Domänenidentitätswechsel'), ('hinzufuegen','hinzufügen'),
    ('Laeuft','Läuft'), ('ergaenzend','ergänzend'), ('zusaetzlich','zusätzlich'),
    ('ueber','über'), ('Ueber','Über'), ('Benutzeridentitaetswechsel','Benutzeridentitätswechsel'),
    ('aehnlich','ähnlich'), ('naemlich','nämlich'), ('moeglich','möglich'),
    ('Groesse','Größe'), ('groesser','größer'), ('loeschen','löschen'),
    ('Haeufig','Häufig'), ('haeufig','häufig'), ('Aenderung','Änderung'),
    ('zugehoerige','zugehörige'), ('Uebersicht','Übersicht'), ('Loesung','Lösung'),
    ('Bypaesse','Bypässe'), ('fuehrt','führt'), ('gehoeren','gehören'), ('oeffnen','öffnen'),
    ('standardmaessig','standardmäßig'), ('woechentlich','wöchentlich'),
    ('Domaenenidentitaetswechsel','Domänenidentitätswechsel'),
    ('identitaetswechsel','identitätswechsel'),
]
def de(t):
    for a, b in UML:
        t = t.replace(a, b)
    return t

def cell(t):
    t = de(t).replace('|', r'\|')
    # <Richtlinie> als Code, sonst schluckt es der Markdown-Renderer
    return re.sub(r'<([^<>]{1,40})>', r'`<\1>`', t)

GROUPS = [
    ('Voreingestellte Richtlinien und Governance', ['PRE']),
    ('Anti-Spam eingehend',                        ['ASI', 'ASF']),
    ('Verbindungsfilter',                          ['CF']),
    ('Anti-Spam ausgehend',                        ['ASO']),
    ('Anti-Phishing und Impersonation',            ['APH', 'IMP']),
    ('Tenant Allow/Block List',                    ['TABL']),
    ('Anti-Malware, Safe Attachments, Safe Links', ['AMW', 'SA', 'SL']),
    ('Quarantäne und Advanced Delivery',           ['QUA', 'ADV']),
    ('E-Mail-Authentifizierung',                   ['AUTH']),
    ('Enhanced Filtering und externe Kennzeichnung', ['EF', 'EXT']),
    ('Reporting, Hybrid, Betrieb',                 ['REP', 'HYB', 'OPS']),
    ('Teams, Wirksamkeitsnachweis, Governance, P2',['TEAMS', 'VER', 'GOV', 'P2']),
]

def pid_sort(pid):
    m = re.match(r'^([A-Z0-9]+)-(\d+)$', pid)
    return (m.group(1), int(m.group(2))) if m else (pid, 0)

counts = collections.Counter(p.split('-')[0] for p in LINKS)
O = []
w = O.append

w('# EOP / Microsoft Defender for Office 365 — Hardening-Standard')
w('')
w('Wiederverwendbare Methodik für EOP- und MDO-Assessments: 120 Prüfpunkte mit Soll-Werten, '
  'Begründung, PowerShell-Befehl und — der Grund für diese README — **dem Fundort im Portal**.')
w('')
w('Alle Soll-Werte, Cmdlets, Parameter und Default-/Preset-Werte sind gegen die Microsoft-Learn-Referenz '
  'geprüft (Recherchestand August 2026). Nicht belegbare Aussagen sind im Guide als solche gekennzeichnet.')
w('')
w('## Inhalt')
w('')
w('| Datei | Zweck |')
w('|---|---|')
w('| `Invoke-EopAudit.ps1` | Liest den Ist-Zustand aus. **Rein lesend.** Liefert CSV, HTML-Report und optional einen JSON-Export als Rollback-Grundlage. Das Arbeitstier. |')
w('| `Invoke-EopHardening.ps1` | Setzt die Baseline. Vorschau ist der Standardmodus — ohne `-Execute` wird nichts geändert. |')
w('| `EOP-MDO_Assessment-Checkliste.xlsx` | 120 Prüfpunkte zum Ausfüllen. Ist-Wert, Bewertung, Notizen; die Auswertung rechnet mit. Blatt „Kürzel" enthält die Legende. |')
w('| `EOP-MDO_Best-Practice-Guide.docx` / `.pdf` | Begründung zu jedem Prüfpunkt: Funktion, Nutzen, Microsoft-Stand 2026, Lizenz, Fallstricke. Anhang E ist die Fundort-Tabelle. |')
w('')
w('Checkliste, Guide und beide Skripte verwenden **dieselben IDs**. Steht im Audit-Report `ASI-04`, '
  'findet sich im Guide unter derselben Nummer die Begründung und in der Checkliste die Zeile zum Ausfüllen.')
w('')
w('## Audit beim Kunden')
w('')
w('```powershell')
w('# Verbindung steht schon? Dann einfach:')
w('.\\Invoke-EopAudit.ps1 -CustomerName "Kunde" -ExportJson')
w('')
w('# Noch nicht verbunden — das Skript meldet sich selbst an:')
w('.\\Invoke-EopAudit.ps1 -CustomerName "Kunde" -UserPrincipalName admin@kunde.onmicrosoft.com -ExportJson')
w('')
w('# Tenantwechsel im selben Fenster:')
w('.\\Invoke-EopAudit.ps1 -CustomerName "Kunde2" -ForceNewConnection')
w('```')
w('')
w('Rollen: **Global Reader** plus **Security Reader** genügen fürs Audit, für die Umsetzung braucht es '
  '**Security Administrator** und **Exchange Administrator**.')
w('')
w('Voraussetzungen: Windows PowerShell 5.1 (unterstützt) oder PowerShell 7, Modul `ExchangeOnlineManagement` V3, '
  'TLS 1.2, .NET Framework 4.7.2+. Beide Skripte sind UTF-8 **mit BOM** gespeichert und PS-5.1-kompatibel.')
w('')
w('Blockt Windows die Datei als „aus dem Internet": `Unblock-File .\\Invoke-EopAudit.ps1`.')
w('')
w('### Verbindung')
w('')
w('Beide Skripte erkennen eine bestehende Exchange-Online-Verbindung und verwenden sie weiter — einmal anmelden, '
  'dann Audit und Hardening nacheinander laufen lassen. Grundlage ist `Get-ConnectionInformation`; '
  '`Get-PSSession` liefert für die REST-Verbindungen seit Modul V3 nichts mehr.')
w('')
w('| Situation | Verhalten |')
w('|---|---|')
w('| Verbindung besteht | wird weiterverwendet |')
w('| keine Verbindung | es wird eine neue aufgebaut |')
w('| nur Security-&-Compliance-Verbindung da | Exchange Online wird zusätzlich verbunden (`IsEopSession` unterscheidet) |')
w('| Token abgelaufen | wird verworfen, neue Anmeldung |')
w('| `-ForceNewConnection` | bestehende wird getrennt, neue Anmeldung |')
w('| `-SkipConnect` | es wird unter keinen Umständen verbunden |')
w('')
w('Beim Start zeigen beide Skripte **Organisation, angemeldetes Konto und Tenant-ID** — die Absicherung gegen '
  'den teuersten Irrtum beim Arbeiten mit mehreren Kundentenants. Weicht das Konto von `-UserPrincipalName` ab, '
  'kommt eine Warnung; das Hardening-Skript fragt im `-Execute`-Modus zusätzlich nach. '
  'Getrennt wird nie automatisch.')
w('')
w('### Was der Report ausgibt')
w('')
w('- **CSV** — für die Übernahme in die Checkliste, mit den Spalten `Fundort`, `PortalLink`, `LearnLink`')
w('- **HTML** — Befunde nach Dringlichkeit, unter jedem Befund der Klickpfad plus Direktlinks ins Portal und zu Learn. '
  'Das ist die Ansicht für die Besprechung mit dem Kunden.')
w('- **JSON** (`-ExportJson`) — der Ist-Zustand als Rollback-Grundlage, bevor etwas geändert wird')
w('')
w('## Die Kürzel')
w('')
w('Jede ID besteht aus Präfix und laufender Nummer. Das Präfix benennt den **fachlichen Bereich** — nicht '
  'zwingend die Portal-Richtlinie, in der die Einstellung sitzt. Das fällt an drei Stellen auseinander, und zwar absichtlich:')
w('')
w('- **APH und IMP** stehen beide in derselben Anti-Phishing-Richtlinie. Getrennt sind sie, weil APH mit jeder '
  'Lizenz funktioniert und IMP eine Defender-Lizenz braucht. Bei einem EOP-only-Kunden lassen sich damit alle '
  'IMP-Punkte in einem Zug auf *n.a.* setzen — das ist der praktische Nutzen.')
w('- **ASF** sitzt innerhalb der Anti-Spam-Richtlinie, ist aber ein eigenes, historisch gewachsenes Regelwerk '
  'mit eigener Bewertungslogik.')
w('- **CF** erscheint im Portal als Zeile *in der Liste der Anti-Spam-Richtlinien*, ist technisch aber eine '
  'eigene Richtlinie mit eigenen Cmdlets und greift vor allen anderen.')
w('')
w('| Kürzel | Steht für | Punkte | Wo es im Portal liegt |')
w('|---|---|---|---|')
for pref, bereich, was in PREFIXES:
    w('| **%s** | %s | %d | %s |' % (pref, de(bereich), counts.get(pref, 0),
                                     de(was).replace(' - ', ' — ')))
w('')
w('Dazu kommen IDs, die die Skripte selbst erzeugen und die auf den Fundort ihres Prüfpunkts verweisen:')
w('')
w('| Erzeugte ID | Bedeutung | Fundort von |')
w('|---|---|---|')
ALIAS_DESC = {
    'INV-ASI': 'Inventar der Anti-Spam-Richtlinien (eingehend)',
    'INV-ASO': 'Inventar der Anti-Spam-Richtlinien (ausgehend)',
    'INV-APH': 'Inventar der Anti-Phishing-Richtlinien',
    'INV-AMW': 'Inventar der Anti-Malware-Richtlinien',
    'INV-SA':  'Inventar der Safe-Attachments-Richtlinien',
    'INV-SL':  'Inventar der Safe-Links-Richtlinien',
    'AUTH-DNS': 'Sammelbefund der DNS-Prüfung (SPF, DKIM, DMARC)',
    'TABL-Sender': 'Einträge für Absender und Domänen',
    'TABL-Url': 'Einträge für URLs',
    'TABL-FileHash': 'Einträge für Dateien (SHA256)',
    'TABL-IP': 'Einträge für IP-Adressen',
}
for k in sorted(ALIASES):
    w('| `%s` | %s | %s |' % (k, ALIAS_DESC.get(k, ''), ALIASES[k]))
w('')
w('`INFO-01` ist die Tenant-Information im Kopf des Reports und hat keinen Fundort.')
w('')
w('## Fundorte: wo die Einstellung im Portal sitzt')
w('')
w('Der Pfad **Bedrohungsrichtlinien** liegt im Defender-Portal unter *E-Mail & Zusammenarbeit > '
  'Richtlinien & Regeln*. Wo kein Portal-Link steht, ist die Einstellung nicht über die Oberfläche '
  'erreichbar — das betrifft vor allem das öffentliche DNS und einige PowerShell-only-Schalter.')
w('')
w('Zwei Fundorte, an denen man erfahrungsgemäß vorbeiscrollt:')
w('')
w('- Die **Verbindungsfilterrichtlinie** (`CF-*`) hat keine eigene Seite. Sie steht als Zeile *in der Liste der '
  'Anti-Spam-Richtlinien* — die **Zeile** anklicken, nicht die Checkbox, dann im Flyout auf '
  '„Verbindungsfilterrichtlinie bearbeiten".')
w('- Der **Teams-Schutz** (`TEAMS-01`) liegt nicht mehr unter Bedrohungsrichtlinien, sondern unter '
  '*Einstellungen > E-Mail & Zusammenarbeit > Microsoft Teams-Schutz*.')
w('')
done = set()
for titel, prefs in GROUPS:
    ids = sorted([p for p in LINKS if p.split('-')[0] in prefs], key=pid_sort)
    if not ids:
        continue
    w('### ' + titel)
    w('')
    w('| ID | Klickpfad | Direkt |')
    w('|:------|:' + '-' * 90 + '|:------|')
    for pid in ids:
        nav, portal, learn = LINKS[pid]
        done.add(pid)
        d = []
        if portal:
            d.append('[Portal](%s)' % portal)
        if learn:
            d.append('[Learn](%s)' % learn)
        w('| `%s` | %s | %s |' % (pid, cell(nav), ' · '.join(d) if d else '—'))
    w('')
missing = sorted(set(LINKS) - done, key=pid_sort)
assert not missing, 'nicht einsortiert: %s' % missing

w('## Was an diesen Links unsicher ist')
w('')
w('Sämtliche Deep-Links ins Defender-Portal stammen aus der Microsoft-Dokumentation — Microsoft nennt sie in '
  'den Learn-Artikeln ausdrücklich. Alle Learn-Sprungmarken sind gegen die Quelldateien der '
  'MicrosoftDocs-Repositories geprüft, also nicht geraten. Drei Einschränkungen gehören dazu:')
w('')
w('- Für das **Exchange Admin Center** dokumentiert Microsoft nur wenige Direktlinks. Nicht belegt und aus der '
  'Praxis übernommen sind die von ' + ', '.join('`%s`' % x for x in UNDOCUMENTED) +
  '. Führt einer ins Leere, steht der Klickpfad daneben. Der belegte Einstieg ist '
  '[admin.exchange.microsoft.com](https://admin.exchange.microsoft.com); der EAC-Übersichtsartikel nennt '
  'inzwischen zusätzlich `admin.cloud.microsoft/exchange`.')
w('- Vier Prüfpunkte haben **keinen passgenauen Learn-Artikel**; der Link führt auf das Nächstliegende:')
w('')
for k in sorted(WEAK_LEARN):
    w('    - `%s` — %s' % (k, de(WEAK_LEARN[k])))
w('')
w('- Die Skripte sind syntaktisch validiert, auf PS-5.1-Kompatibilität geprüft und mit Mock-Cmdlets '
  'end-to-end durchgelaufen. **Gegen einen Produktivtenant getestet wurden sie nicht** — der erste Lauf '
  'beim Kunden gehört in den Lesemodus.')
w('')
w('## Policy-Inventar: warum manche Befunde nur „Info" sind')
w('')
w('Eine Schutzrichtlinie in EOP besteht immer aus **zwei** Objekten: der *Policy* mit den Werten und der '
  '*Rule*, die festlegt, für wen sie gilt. Das Portal legt beides gemeinsam an, PowerShell nicht — '
  'daraus entstehen zwei Zustände, die wie funktionierende Richtlinien aussehen:')
w('')
w('| Status | Bedeutung | Wird bewertet? |')
w('|---|---|---|')
w('| `aktiv - Prio n` | wird auf Empfänger angewendet | ja |')
w('| `Default-Policy` | greift immer, hat keine Regel | ja |')
w('| `Regel deaktiviert` | Policy sichtbar, wirkt aber auf niemanden | nein, nur Info |')
w('| `ohne Regel` | im Defender-Portal unsichtbar, wirkt auf niemanden | nein, nur Info |')
w('| `Preset - nicht editierbar` | von Microsoft verwaltet | nein, nur Info |')
w('')
w('Eine Richtlinie, die auf niemanden angewendet wird, kann keinen Mangel darstellen — ihre Werte gegen die '
  'Baseline zu prüfen erzeugt nur Rauschen. Das Audit weist solche Objekte im Abschnitt `INV-*` mit ihrem '
  'Status aus und bewertet sie als *Info*. Prüfpunkt `PRE-05` greift das auf: verwaiste und stillgelegte '
  'Policies gehören reaktiviert, gelöscht oder als bewusst stillgelegt dokumentiert.')
w('')
w('## Was die Skripte bewusst nicht anfassen')
w('')
w('Diese Punkte gehören in die manuelle Bewertung, weil automatisches Setzen den Mailfluss brechen kann:')
w('')
w('- Allow-Listen der Anti-Spam-Richtlinie (`ASI-10`) und die IP Allow List (`CF-01`)')
w('- Bestehende Einträge der Tenant Allow/Block List (`TABL-01` bis `TABL-06`)')
w('- Transportregeln mit Filter-Bypass (`ADV-03`)')
w('- Enhanced Filtering (`EF-01`) — die Entscheidung hängt am MX-Ziel')
w('- Alles im DNS: SPF, DKIM-Rotation, DMARC, MTA-STS, DANE (`AUTH-01` bis `AUTH-12`)')
w('- On-Premises und Firewall (`HYB-01` bis `HYB-10`)')
w('')

io.open('README.md', 'w', encoding='utf8').write('\n'.join(O) + '\n')
print('README.md: %d Zeilen, %d Fundorte, %d Kuerzel' % (len(O), len(LINKS), len(PREFIXES)))
