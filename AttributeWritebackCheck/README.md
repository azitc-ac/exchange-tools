# AttributeWritebackCheck

Passen die Exchange-Attribute im lokalen AD zu Exchange Online? Das Skript vergleicht für jedes
synchronisierte Postfach den Empfängertyp, `mail` und die primäre Adresse in `proxyAddresses`,
erklärt jede Abweichung in einem Satz und gibt am Ende **ein** Urteil aus: GRÜN, GELB, ROT oder
UNVOLLSTÄNDIG. Es ändert nichts, weder im AD noch in Exchange Online.

Gedacht für die Cloud-Verwaltung der Exchange-Attribute (`IsExchangeCloudManaged`) und den
Cloud-Sync-Job „EXO to AD attribute sync“: **vorher** als Vorschau, was der Writeback im AD ändern
würde, **nachher** als Kontrolle, ob alles passt.

```powershell
.\Test-ExchangeAttributeWriteback.ps1 -BeforeWriteback   # vor dem Einrichten: Vorschau
.\Test-ExchangeAttributeWriteback.ps1                    # nach dem Einrichten: Befund
.\Test-ExchangeAttributeWriteback.ps1 -ExtraServer dc01.fabrikam.com -Credential (Get-Credential)   # weiterer Forest
```

Ohne bestehende Sitzung meldet es sich selbst an Exchange Online an (gemeinsamer Prolog,
`-UserPrincipalName` optional). Durchsucht werden alle Domänen des eigenen Forests, weitere über
`-ExtraServer`. Das Ergebnis steht zusätzlich als CSV im aktuellen Ordner (`-OutFile`).

## Warum

Microsoft warnt: Wird ein Postfach zu früh nach dem Move auf cloud-verwaltet umgestellt, können
veraltete Empfängertyp-Werte aus der Cloud die richtigen Remote-Mailbox-Werte im AD überschreiben,
sobald der Writeback läuft
([Cloud-based management of Exchange attributes](https://learn.microsoft.com/en-us/exchange/hybrid-deployment/enable-exchange-attributes-cloud-management)).
Umgekehrt zeigt der Vergleich, was sich im AD über Jahre auseinanderentwickelt hat – und was der
Writeback dort angleichen wird.

## Zuordnung Postfach ↔ AD-Objekt

In dieser Reihenfolge, der verwendete Schlüssel steht in der Spalte `MatchedBy`:

1. **Exchange-GUID** – `ExchangeGuid` gegen `msExchMailboxGuid`. Fehlt im AD oft bei Postfächern,
   die mit `New-RemoteMailbox` direkt in der Cloud angelegt wurden.
2. **Entra-Objekt-ID** – `ExternalDirectoryObjectId` gegen `msDS-ExternalDirectoryObjectId`
   (von Entra Connect per Exchange-Hybrid-Writeback als `User_<ID>` geschrieben).
3. **UPN** – als letzter Ausweg, ohne Rücksicht auf Groß-/Kleinschreibung.

## Ergebnisse je Postfach

| Ergebnis | Stufe | Bedeutung |
|---|---|---|
| `OK-Cloud` | OK | cloud-verwaltet, Typ und Adressen stimmen |
| `OK-OnPrem` | OK | noch on-prem verwaltet, Werte stimmen |
| `TYPE-WRONG-Cloud` | FAIL (vorab: WARN) | cloud-verwaltet, aber im AD falscher Typ – Microsofts Warnfall |
| `ADDRESS-DIFF` | WARN | primäre Adresse in `proxyAddresses` weicht ab |
| `MAIL-DIFF` | WARN | `mail` weicht ab (typisch: Writeback-Job vor dem 03.08.2026 angelegt) |
| `TYPE-MISMATCH-OnPrem` | INFO | on-prem verwaltet, Typ passt nicht (z. B. in der Cloud zu Shared umgewandelt) – Aufräum-Empfehlung, kein Einfluss aufs Urteil |
| `TYPE-NO-TARGET` | INFO | für diesen Postfachtyp gibt es keinen Sollwert |
| `NOT-IN-AD` | – | kein passendes AD-Objekt gefunden |

Gewichtung: Über GRÜN oder GELB entscheiden die **Adressen**, denn die lesen lokale Anwendungen.
Der Empfängertyp im AD zählt nur, solange ein Exchange-Server on-prem läuft oder falls die
Verwaltung zurückgestellt wird. Ausnahme ist der Warnfall: Er hat Vorrang vor allen anderen
Ergebnissen und macht das Urteil ROT.

## Urteil und Exitcode

| Urteil | Exitcode | Bedingung |
|---|---|---|
| GRÜN | 0 | jedes Postfach zugeordnet, keine Warnung, kein Fehler |
| GELB | 1 | mindestens eine Adressabweichung (oder vorab ein falscher Typ) |
| ROT | 2 | mindestens ein Warnfall nach dem Writeback |
| UNVOLLSTÄNDIG | 3 | Postfächer ohne AD-Gegenstück, sonst unauffällig – ausdrücklich kein Grün |
| NICHT BEWERTBAR | 3 | AD nicht lesbar oder kein einziges Postfach zugeordnet |

Grenze der Prüfung: Sie zeigt den heutigen Zustand. Ob der Writeback `mail` aktiv schreibt, belegt
sie nur, wenn sich die Adresse seit der Umstellung geändert hat. Vorab gilt außerdem: Der Writeback
liest seine Werte aus Entra ID, und ein veralteter Empfängertyp dort ist von außen nicht sichtbar.
Als letzte Gegenprobe vor „Start provisioning“ deshalb „Provision on demand“ für ein, zwei
Testbenutzer – das schreibt für diese Benutzer wirklich.

## Voraussetzungen

Windows PowerShell 5.1 oder PowerShell 7, das Modul `ExchangeOnlineManagement` **ab 3.6.0**
(prüft der gemeinsame Prolog) und das AD-Modul aus den RSAT. Lesende Rechte im AD und in Exchange
Online genügen.

## Tests

`tests\Test-WritebackLogik.ps1` prüft Zuordnung und Urteil mit Attrappen statt Exchange und AD:
alle drei Schlüssel, der Warnfall vor jeder Adressabweichung, Typabweichung on-prem nur als Hinweis,
Vorschau-Modus, und dass nicht zugeordnete Postfächer nie GRÜN ergeben. Der letzte Punkt stammt aus
dem ersten Einsatz: Dort ergab 1 bewertetes von 1261 Postfächern noch GRÜN.
