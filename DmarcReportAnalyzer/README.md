# DmarcReportAnalyzer

Wertet DMARC-Aggregate-Reports (RUA) aus und schreibt einen **HTML-Bericht** samt **CSV** mit
allen Einzeldatensätzen. Der Bericht beantwortet: Wer sendet in meinem Namen? Besteht das DMARC –
und woran hängt es, an DKIM oder nur an SPF? Ist die Domain reif für die nächste Policy-Stufe?

Die Reports werden dort gelesen, wo sie ohnehin ankommen. Ein externer Dienst ist nicht nötig, und
die Daten verlassen den Rechner nicht.

```powershell
.\Invoke-DmarcReportAnalyzer.ps1 -Path C:\Reports                              # Ordner mit .xml/.gz/.zip
.\Invoke-DmarcReportAnalyzer.ps1 -GroupMailbox dmarc@contoso.onmicrosoft.com   # Microsoft-365-Gruppe
.\Invoke-DmarcReportAnalyzer.ps1 -Mailbox dmarc-reports@contoso.com -Days 14   # (freigegebenes) Postfach
```

| Parameter | Bedeutung |
|---|---|
| `-Path` | Ordner, Unterordner eingeschlossen |
| `-GroupMailbox` | Gruppenpostfach einer Microsoft-365-Gruppe; delegiert `Group.Read.All` (einmalig Administratorzustimmung, siehe unten) |
| `-Mailbox` | persönliches oder freigegebenes Postfach mit Vollzugriff; delegiert `Mail.Read.Shared` |
| `-Days` | Zeitraum, Standard 30 |
| `-OutputPath` | HTML-Datei; die CSV entsteht daneben. Fehlende Ordner werden angelegt |
| `-AccessToken`, `-Tenant` | vorhandenes Graph-Token bzw. Tenant für die Anmeldung |
| `-NoReverseDns` | keine PTR-Abfragen – schneller, aber ohne Zuordnung zu Diensten |

**Anmeldung.** Bei `-GroupMailbox` und `-Mailbox` meldet sich das Skript per Gerätecode an, und
zwar mit der App *Microsoft Graph Command Line Tools*. Anmelden muss sich **der Benutzer, der
Zugriff auf die Gruppe bzw. das Postfach hat**, kein Administrator. Ein Graph-Modul ist nicht
nötig.

**Einmalige Administratorzustimmung bei `-GroupMailbox`.** `Group.Read.All` braucht laut
Microsoft auch delegiert eine Administratorzustimmung. Hat im Tenant noch niemand der App
*Microsoft Graph Command Line Tools* dieses Recht gewährt, muss ein Administrator einmal
zustimmen, z. B. bei seiner eigenen ersten Anmeldung mit „Im Namen der Organisation zustimmen“.
Danach genügt jedem berechtigten Benutzer die eigene Anmeldung. `Mail.Read.Shared` für
`-Mailbox` braucht keine Administratorzustimmung. Der Gerätecode läuft bewusst über REST, weil `Connect-MgGraph` in Hosts ohne Konsole mit
„An error occurred when writing to a listener“ scheitert.

**Exitcodes.** `0` = Bericht geschrieben, `2` = keine DMARC-Daten im Zeitraum.

## Den Bericht lesen

Die Lesehilfe steht auch im Bericht selbst, aufklappbar unter „So liest du den Bericht“.

- **Übersicht je Domain:** Anteil der Mails mit DMARC-Pass, Policy, Zeitraum, Reporter.
- **Empfehlungen:** nach Dringlichkeit sortiert. Jede quellenbezogene Empfehlung nennt die Tage
  und springt per Link zur betroffenen Zeile.
- **Quellen:** je Domain und IP-Adresse. Sie zeigen Zeitraum, Mengen, DKIM- bzw. SPF-Alignment, das
  rohe DKIM-Ergebnis und die Spalte **Auffällig**. Auffällige Zeilen sind gelb hinterlegt.

| Auffällig | Bedeutung | Was tun |
|---|---|---|
| kein DKIM, nur SPF | Der Versandweg signiert nicht mit der eigenen Domain | DKIM einrichten oder über einen signierenden Dienst senden |
| DKIM gebrochen, nur SPF | Signiert, aber die Mail wurde danach verändert | Das verändernde System (Signatur-/Disclaimer-Gateway, Scanner, Verteiler) vor das Signieren legen |
| zweite Signatur gebrochen | Eine gültige Signatur blieb übrig | nichts – harmlos |
| Weiterleitung | Ein bekannter Weiterleiter (Google, GMX/WEB.DE) hat die Mail weitergereicht | nichts – vom Absender aus nicht zu beheben |
| DMARC fail | Nicht bestanden | Dienst identifizieren: gewollt → DKIM/SPF einrichten, sonst Missbrauch |

**Policy-Empfehlung.** Der nächste Schritt (`none` → `quarantine` → `reject`) wird erst
empfohlen, wenn über mindestens 14 Tage 98 % bzw. 99 % bestehen **und** keine Mail mehr allein an
SPF hängt. Unter einer strengeren Policy würde eine solche Mail nach jeder Weiterleitung
abgewiesen.

**Wie oft?** Wöchentlich, solange an der Policy gedreht wird. Sonst monatlich und nach jeder
Änderung am Mailfluss, etwa an einem Connector, einem Gateway oder einem neuen Versanddienst.

## Sicherheit

Report-Inhalte sind Fremddaten, jeder kann einen Report an die RUA-Adresse schicken. Das Skript
liest XML deshalb ohne DTD-Verarbeitung (kein XXE) und HTML-kodiert alles, was in den Bericht
geht. Der Bericht ist eine einzelne Datei ohne externe Skripte, Schriften oder Bilder.

## Tests

```powershell
.\tests\Test-DmarcAnalyse.ps1
```

Die Tests arbeiten mit künstlichen Reports, also Beispieldomains und Dokumentations-Adressen nach
RFC 5737, als `.xml`, `.gz` und `.zip`, ohne Netz und ohne Postfach. Echte Report-Daten gehören
nicht ins Repo.

## Release

Bei jeder neuen Versionsnummer in `$script:Version` entsteht automatisch ein Release mit dem
Ordner als ZIP (`DmarcReportAnalyzer/v<Version>`).
