# MailboxDatabaseCheck

Ist die Postfachdatenbank wirklich leer? `Remove-MailboxDatabase` und die Deinstallation des
letzten Exchange-Servers scheitern, solange noch irgendein Postfach darin liegt – und „irgendein“
heißt mehr als Benutzerpostfächer. Das Skript sucht alle Arten zusammen und schlägt für jeden Rest
den Befehl vor, der ihn entfernt. Es ändert selbst nichts.

```powershell
.\Test-MailboxDatabaseEmpty.ps1                          # Auswahl im Dialog, Mehrfachauswahl mit Strg
.\Test-MailboxDatabaseEmpty.ps1 -Database DB01, DB02     # ohne Dialog, auch auf Server Core
.\Test-MailboxDatabaseEmpty.ps1 -Database DB01 -PassThru | Export-Csv rest.csv -NoTypeInformation -Encoding UTF8
```

## Was geprüft wird

| Typ | Abfrage | Blockiert | Vorschlag |
|---|---|---|---|
| Postfach | `Get-Mailbox -Database` | ja | nach Exchange Online migrieren, sonst `Disable-Mailbox` |
| Discovery | dieselbe Abfrage, `RecipientTypeDetails = DiscoveryMailbox` – einmal gezählt | ja | `Remove-Mailbox` |
| Archiv | `Get-Mailbox -Archive -Database` **und** jedes Postfach, dessen `ArchiveDatabase` hierher zeigt | ja | `Disable-Mailbox -Archive` |
| PublicFolder | `Get-Mailbox -PublicFolder` | ja | `Remove-Mailbox -PublicFolder` (nach der PF-Migration) |
| Arbitration | `Get-Mailbox -Arbitration` | ja | `Remove-Mailbox -Arbitration -RemoveArbitrationMailboxWithOABsAllowed -RemoveLastArbitrationMailboxAllowed` |
| AuditLog | `Get-Mailbox -AuditLog` | ja | `Remove-Mailbox -AuditLog` |
| Monitoring | `Get-Mailbox -Monitoring` | unklar | `Disable-Mailbox` |
| SoftDeleted, Disabled | `Get-MailboxStatistics`, nur bei eingehängter Datenbank | nein | Inhalt geht mit der Datenbank verloren |

Grundlage für „blockiert“ ist Microsofts Liste in
[Error when you try to remove a mailbox database](https://learn.microsoft.com/en-us/troubleshoot/exchange/administration/error-remove-a-mailbox-database):
User-, Archive-, Public-Folder-, Arbitration- und AuditLog-Postfächer. Monitoring-Postfächer stehen
dort nicht – deshalb „unklar“ statt einer Behauptung.

Archive fragt Microsofts Anleitung auf zwei Wegen ab: `Get-Mailbox -Archive -Database` und alle
Postfächer, deren `ArchiveDatabase` auf die Datenbank zeigt. Das Skript übernimmt beide und zählt
Doppelte einmal. Für den zweiten Weg wird die Organisation einmal gelesen, nicht je Datenbank.

## Ausgabe und Exitcode

Je Datenbank eine Tabelle (Typ, Detail, Name, Blockiert) und darunter die Befehle zum Kopieren.
Mit `-PassThru` kommen die Funde zusätzlich als Objekte mit `Database`, `Typ`, `Detail`, `Name`,
`Identity`, `Blockiert` und `Befehl`.

| Exitcode | Bedeutung |
|---|---|
| 0 | keine blockierenden Postfächer in den gewählten Datenbanken |
| 1 | mindestens ein Postfach blockiert oder ist unklar |
| 2 | nichts geprüft: keine Auswahl getroffen oder Datenbank unbekannt |

Die Sitzung wird auf den ganzen Forest gestellt (`Set-ADServerSettings -ViewEntireForest $true`),
sonst fehlen bei mehreren Domänen Postfächer aus den anderen.

## Voraussetzungen

Windows PowerShell 5.1 mit der Exchange-Verwaltungsshell, Exchange 2013 bis SE. Der gemeinsame
Prolog prüft Snap-in und Cmdlets vorab.

## Tests

`tests\Test-OccupantLogik.ps1` prüft die Sammellogik mit Attrappen statt Exchange: Discovery nur
einmal, Archive über beide Wege ohne Doppelte, Befehle je Typ, getrennte Postfächer nicht
blockierend, keine Statistikabfrage bei ausgehängter Datenbank, Quoting von Apostrophen.
