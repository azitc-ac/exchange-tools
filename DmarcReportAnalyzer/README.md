# DmarcReportAnalyzer

Wertet DMARC-Aggregate-Reports (RUA) aus und schreibt einen **HTML-Bericht** samt **CSV** mit
allen Einzeldatensätzen. Der Bericht beantwortet: Wer sendet in meinem Namen? Besteht das DMARC –
und woran hängt es, an DKIM oder nur an SPF? Ist die Domain reif für die nächste Policy-Stufe?

Die Reports werden dort gelesen, wo sie ohnehin ankommen. Ein externer Dienst ist nicht nötig, und
die Daten verlassen den Rechner nicht. Abgerufen wird nur die öffentliche Liste der
Exchange-Online-Adressbereiche von Microsoft (abschaltbar mit `-NoMicrosoftRanges`).

```powershell
.\Invoke-DmarcReportAnalyzer.ps1 -Path C:\Reports                              # Ordner mit .xml/.gz/.zip
.\Invoke-DmarcReportAnalyzer.ps1 -GroupMailbox dmarc@contoso.onmicrosoft.com   # Microsoft-365-Gruppe
.\Invoke-DmarcReportAnalyzer.ps1 -Mailbox dmarc-reports@contoso.com -Days 14   # (freigegebenes) Postfach
.\Invoke-DmarcReportAnalyzer.ps1 -Path C:\Reports -MessageTrace               # dazu die Mails aus dem Message Trace
```

| Parameter | Bedeutung |
|---|---|
| `-Path` | Ordner, Unterordner eingeschlossen |
| `-GroupMailbox` | Gruppenpostfach einer Microsoft-365-Gruppe; delegiert `Group.Read.All` (einmalig Administratorzustimmung, siehe unten) |
| `-Mailbox` | persönliches oder freigegebenes Postfach mit Vollzugriff; delegiert `Mail.Read.Shared` |
| `-Days` | Zeitraum, Standard 30 |
| `-OutputPath` | HTML-Datei; die CSV entsteht daneben. Fehlende Ordner werden angelegt |
| `-UseDeviceCode` | Anmeldung per Gerätecode statt im Browser |
| `-AccessToken`, `-Tenant` | vorhandenes Graph-Token bzw. Tenant für die Anmeldung |
| `-NoReverseDns` | keine PTR-Abfragen – schneller, aber ohne Zuordnung zu Diensten |
| `-NoMicrosoftRanges` | die Adressbereiche von Exchange Online nicht laden; Microsoft 365 wird dann nur am Rückwärtsnamen erkannt |
| `-MessageTrace` | zu jeder auffälligen Quelle die passenden Mails aus dem Message Trace suchen, siehe unten; braucht `ExchangeOnlineManagement` |

**Anmeldung.** Bei `-GroupMailbox` und `-Mailbox` öffnet das Skript den Browser zur Anmeldung,
und zwar mit der App *Microsoft Graph Command Line Tools*. Die Antwort nimmt es auf
`http://localhost` an einem freien Port entgegen; der Port lauscht nur lokal, braucht keine
Adminrechte und keine Firewall-Freigabe und schließt nach der Anmeldung wieder. Der Ablauf ist
der Autorisierungscode mit PKCE. Öffnet sich kein Browser, steht die Adresse auch im Fenster.
Nach 5 Minuten ohne Anmeldung bricht das Skript mit einem Hinweis ab.

Auf Rechnern ohne Browser, etwa auf einem Server Core oder in einer Sitzung ohne Desktop, gibt
es den Gerätecode als Ausweichweg mit `-UseDeviceCode`. Dann meldest du dich auf einem
beliebigen Gerät unter microsoft.com/devicelogin mit dem angezeigten Code an. Der Code gilt etwa
15 Minuten.

Anmelden muss sich in beiden Fällen **der Benutzer, der Zugriff auf die Gruppe bzw. das Postfach
hat**, kein Administrator. Ein Graph-Modul ist nicht nötig.

**Einmalige Administratorzustimmung bei `-GroupMailbox`.** `Group.Read.All` braucht laut
Microsoft auch delegiert eine Administratorzustimmung. Hat im Tenant noch niemand der App
*Microsoft Graph Command Line Tools* dieses Recht gewährt, muss ein Administrator einmal
zustimmen, z. B. bei seiner eigenen ersten Anmeldung mit „Im Namen der Organisation zustimmen“.
Danach genügt jedem berechtigten Benutzer die eigene Anmeldung. `Mail.Read.Shared` für
`-Mailbox` braucht keine Administratorzustimmung.

Beide Anmeldewege laufen bewusst über REST statt über das Graph-Modul, weil `Connect-MgGraph` in
Hosts ohne Konsole mit „An error occurred when writing to a listener“ scheitert.

**Exitcodes.** `0` = Bericht geschrieben, `2` = keine DMARC-Daten im Zeitraum.

## Den Bericht lesen

Die Lesehilfe steht auch im Bericht selbst, aufklappbar unter „So liest du den Bericht“.

- **Übersicht je Domain:** Anteil der Mails mit DMARC-Pass, Policy, Zeitraum, Reporter.
- **Empfehlungen:** nach Dringlichkeit sortiert. Jede quellenbezogene Empfehlung nennt die Tage
  und springt per Link zur betroffenen Zeile.
- **Quellen:** je Domain und IP-Adresse. Sie zeigen Zeitraum, Mengen, DKIM- bzw. SPF-Alignment, das
  rohe DKIM-Ergebnis und die Spalte **Auffällig**. Auffällige Zeilen sind gelb hinterlegt.
  Quellen in den Adressbereichen von Exchange Online heißen „Microsoft 365“, auch ohne
  passenden Rückwärtsnamen.
- **Einzelheiten:** unter jeder auffälligen Quelle aufklappbar, je Datensatz Tag, Melder, Anzahl,
  Empfängerdomain (falls der Melder sie angibt), SPF, DKIM mit Selektor, Return-Path und der vom
  Empfänger gemeldete Grund (etwa `forwarded`). Ein Report nennt weder Uhrzeit noch
  Empfängeradresse noch Betreff; mehr als den Tag und den empfangenden Anbieter gibt er nicht her.

| Auffällig | Bedeutung | Was tun |
|---|---|---|
| kein DKIM, nur SPF | Der Versandweg signiert nicht mit der eigenen Domain | DKIM einrichten oder über einen signierenden Dienst senden |
| DKIM gebrochen, nur SPF | Signiert, aber die Mail wurde danach verändert | Das verändernde System (Signatur-/Disclaimer-Gateway, Scanner, Verteiler) vor das Signieren legen |
| zweite Signatur gebrochen | Eine gültige Signatur blieb übrig | nichts – harmlos |
| Weiterleitung | Ein bekannter Weiterleiter (Google, GMX/WEB.DE) hat die Mail weitergereicht, oder der Empfänger meldet als Grund `forwarded`, `mailing_list` bzw. `trusted_forwarder` | nichts – vom Absender aus nicht zu beheben |
| DMARC fail | Nicht bestanden | Dienst identifizieren: gewollt → DKIM/SPF einrichten, sonst Missbrauch |

**Welche Mail war es? (`-MessageTrace`)** Mit dem Schalter sucht das Skript zu jeder auffälligen
Quelle (DMARC fail, nur SPF, DKIM gebrochen) im Message Trace von Exchange Online die Mails mit
Absender `@Domain` an den Tagen des Befunds. Oben stehen die zugestellten Mails an Empfänger, die
zum Melder passen: die Empfängerdomain aus dem Report, sonst die Domains des Anbieters (etwa
`gmail.com` für Google). „Enterprise Outlook“ (Microsoft 365) lässt sich keiner Domain zuordnen,
dann stehen alle Mails des Tages da. Das ist eine Kandidatenliste, kein Nachweis. Der Message
Trace kennt nur Mails, die durch den eigenen Tenant liefen, und reicht 90 Tage zurück. Das Skript
nutzt eine bestehende Verbindung (`Connect-ExchangeOnline`), sonst meldet es sich selbst an; nötig
ist ein Recht zum Lesen des Message Trace. Je Lauf fragt es höchstens 40 Tage ab. Die Kandidaten
stehen im HTML unter „Einzelheiten“ und zusätzlich in `<Bericht>.trace.csv`.

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
RFC 5737, als `.xml`, `.gz` und `.zip`, ohne Netz und ohne Postfach. Den Message Trace ersetzt
eine Attrappe, die die Abfragen mitschreibt. Echte Report-Daten gehören
nicht ins Repo. Die Browser-Anmeldung wird ohne Browser und ohne Microsoft geprüft: PKCE gegen
den Testvektor aus RFC 7636, und der lokale Port mit einer simulierten Rückkehr. Die echte
Anmeldung bei Microsoft deckt der Test nicht ab.

## Release

Bei jeder neuen Versionsnummer in `$script:Version` entsteht automatisch ein Release mit dem
Ordner als ZIP (`DmarcReportAnalyzer/v<Version>`).
