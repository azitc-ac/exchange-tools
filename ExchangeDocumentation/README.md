# ExchangeDocumentation

Nimmt eine Exchange-Organisation vollständig auf und schreibt sie als CSV- und Textdateien
heraus – einmal für on-premises, einmal für Exchange Online. Gedacht für Übernahmen, vor
Migrationen und als regelmäßige Bestandsaufnahme.

| Skript | Ziel |
|---|---|
| `Get-ExchangeDocumentation.ps1` | Exchange on-premises, läuft in der Verwaltungsshell auf einem Server |
| `Get-ExchangeOnlineDocumentation.ps1` | Exchange Online, verbindet sich bei Bedarf selbst |

## Aufrufen

```powershell
# on-premises
.\Get-ExchangeDocumentation.ps1
.\Get-ExchangeDocumentation.ps1 -includeMailboxFolderPermissions -includeInboxRules -includeMobileDevices

# Exchange Online
.\Get-ExchangeOnlineDocumentation.ps1
.\Get-ExchangeOnlineDocumentation.ps1 -IncludeMailboxFolderPermissions $true -IncludeRecipients $false
```

Die drei Schalter der on-prem-Fassung sind bewusst nicht voreingestellt: Ordnerrechte,
Posteingangsregeln und Mobilgeräte werden je Postfach einzeln abgefragt und dauern bei großen
Organisationen entsprechend lange. Die Online-Fassung nimmt dieselbe Entscheidung als `[bool]`
entgegen – Postfachrechte und Empfänger sind dort an, Ordnerrechte aus.

Die Online-Fassung stellt selbst eine Verbindung her (`Connect-ExchangeOnline -ShowBanner:$false`),
wenn noch keine Sitzung besteht, und legt ihre Ausgabe unter `.\EXO-Documentation\<Tenant>\` ab.

## Was herauskommt

Rund 45 Dateien je Lauf, thematisch gruppiert:

| Gruppe | Inhalt |
|---|---|
| Infrastructure | Server, Datenbanken, DAG, Zertifikate |
| Org | Akzeptierte Domänen, Adresslisten, Adressbuchrichtlinien, OAB, Organisationskonfiguration |
| Mailboxes | Postfächer, Statistiken, Archive, Vollzugriff, SendAs, SendOnBehalf, Weiterleitungen |
| Groups | Verteilergruppen samt Mitgliedern, dynamische Gruppen |
| Transport | Send- und Receive-Connectors, Transportregeln, Journalregeln |
| Policies | Aufbewahrung samt Tags, OWA, Mobilgeräte, IRM |
| Hybrid | Hybridkonfiguration, Verbundvertrauensstellung, Organisationsbeziehungen, AuthServer, IntraOrg-Connector |
| Recipients | Kontakte, MailUser, Ressourcen |

Mehrwertige Felder stehen in einer Zelle, getrennt durch `§`. Das Zeichen wird im Skript aus
seinem Zeichencode gebaut, damit die Datei selbst reines ASCII bleibt und eine falsche Kodierung
beim Kopieren nichts kaputt macht.

## Herkunft

Aus einer langen Reihe gewachsen (v2.1 bis v3.2 bzw. v2.2 bis v3.1, verteilt über mehrere Server
und OneDrive). Hier liegen die jeweils jüngsten Stände: on-prem v3.2 vom 28.05.2026, online v3.1
vom 10.09.2026. Alle älteren Fassungen sind damit abgelöst.

## Voraussetzungen

Windows PowerShell 5.1. On-prem: Exchange-Verwaltungsshell auf dem Server. Online: Modul
`ExchangeOnlineManagement` und ein Konto mit Leserechten in der Organisation.
