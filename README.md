# exchange-tools

Werkzeuge für den Betrieb von Exchange – on-premises und Exchange Online. Ein Ordner je
Werkzeug, jeder für sich lauffähig – es gibt keine gemeinsame Laufzeit und keine Installation.

| Ordner | Werkzeug | Kurz |
|---|---|---|
| [LogViewer](LogViewer/) | Log-Betrachter | CMTrace, CSV, W3C und Text; Live-Aktualisierung, Suche, Filter. Zweisprachig, als EXE verfügbar |
| [AcmeExchange](AcmeExchange/) | Zertifikats-Lebenszyklus | Let's Encrypt über DNS-01 samt Verteilung auf die Exchange-Server, oder Import eines anderswo ausgestellten PFX |
| [ExchangeCertFix](ExchangeCertFix/) | Zertifikatsprobleme | Wenn das neue Zertifikat nicht SMTP-Standard wird oder das alte sich nicht löschen lässt |
| [DeliveryDiagnostics](DeliveryDiagnostics/) | Zustellung ins Postfach | Zustellweg über Port 475 prüfen, Laufzeitzustand sichern, SMTP-Zertifikate von innen und außen |
| [ExchangeDocumentation](ExchangeDocumentation/) | Bestandsaufnahme | Organisation vollständig als CSV herausschreiben – on-premises und Exchange Online |
| [PublicFolderMigration](PublicFolderMigration/) | Öffentliche Ordner → Postfächer | Struktur, Inhalte und Berechtigungen per EWS, inkrementell, einzeln oder als Stapel |
| [ExchangeTester](ExchangeTester/) | Erreichbarkeit und Hybrid | AutoDiscover wie Outlook, MRS-Proxy, Hybrid-Endpunkte und Frei/Gebucht über die Grenze hinweg |
| [ExchangeLogAnalyzer](ExchangeLogAnalyzer/) | Protokollauswertung | Wertet SMTP-Protokolle aus (EHLO/HELO je Gegenstelle, Connector- und Zeitraumfilter) |
| [ExchangeQueueViewer](ExchangeQueueViewer/) | Warteschlangen | Ersatz für den Queue Viewer der Toolbox – Warteschlangen und Nachrichten ansehen und bearbeiten |
| [ImportExchangePfx](ImportExchangePfx/) | PFX-Import | Kleines Hilfsskript: PFX auswählen und auf dem lokalen Server importieren |
| [EopMdoHardening](EopMdoHardening/) | EOP- und Defender-Assessment | 120 Prüfpunkte für Exchange Online Protection und Defender for Office 365: Ist-Zustand auslesen, Abweichungen begründen, Baseline setzen |
| [MailContactEditor](MailContactEditor/) | E-Mail-Kontakte (Exchange Online) | Kontakt aus filterbarer Liste wählen, primäre und externe Adresse ändern |

## Voraussetzungen

Windows PowerShell 5.1. `AcmeExchange`, `ExchangeQueueViewer` und `ImportExchangePfx` erwarten die
Exchange-Verwaltungsshell auf dem Server, auf dem sie laufen. `EopMdoHardening` und
`MailContactEditor` arbeiten gegen Exchange Online und erwarten das Modul
`ExchangeOnlineManagement` **ab Version 3.6.0**. `LogViewer`, `ExchangeLogAnalyzer` und
`ExchangeTester` sind davon unabhängig – der Tester läuft bewusst auch auf einem
Arbeitsplatzrechner, weil er die Sicht des Clients einnimmt.

Die Werkzeuge prüfen das selbst, statt mitten in der Arbeit mit „Begriff wird nicht erkannt"
abzubrechen: PowerShell-Ausgabe, TLS 1.2, Modulversion, bestehende Anmeldung – und ob die
Cmdlets, die sie wirklich brauchen, nach der Anmeldung auch da sind. Letzteres ist eine Frage
der RBAC-Rolle, nicht der Installation, und wird als solche gemeldet. Bei den Werkzeugen mit
Fenster erscheint die Meldung als Dialog, weil eine mit `-noConsole` gebaute EXE sonst wortlos
nicht startet.

Dieser Prüfblock steht einmal in `build\Prolog.Exo.ps1` bzw. `build\Prolog.OnPrem.ps1` und wird
von `build\Sync-Prolog.ps1` in die Hauptskripte kopiert – die Werkzeuge bleiben damit einzeln
lauffähig. `build\Test-Prolog.ps1` meldet jede abgewichene Kopie und jedes Werkzeug, das Exchange
ungeprüft aufruft; noch nicht umgestellte Altfälle stehen namentlich in
`build\prolog-map.psd1`.

## Releases

Fertige EXE-Dateien liegen unter [Releases](../../releases), nicht im Repo. Ein Release je
Werkzeug, ausgelöst durch einen Tag der Form `<Werkzeug>/v<Version>`:

```
git tag MailContactEditor/v1.0.0
git push origin MailContactEditor/v1.0.0
```

**Tags einzeln pushen.** Wer mehr als drei Tags in einem `git push --tags` schickt,
bekommt von GitHub für die überzähligen *keine* Push-Events – die Releases bleiben
dann stillschweigend aus. Nachholen lässt sich das über *Actions → Release → Run
workflow* mit dem Tag als Eingabe, oder per `gh workflow run Release -f tag=<Tag>`.

`.github/workflows/release.yml` baut daraufhin die EXE mit ps2exe und hängt sie an. Welche
Dateien ins Release gehören, steht in der `release.psd1` des Werkzeugs; ein Werkzeug ohne
diese Datei bekommt kein Release. Die Versionsnummer steht an genau einer Stelle – als
`$script:Version` im Hauptskript – und der Workflow bricht ab, wenn der Tag etwas anderes
behauptet.

## Prüfungen

```powershell
.\build\Invoke-AllTests.ps1            # alles
.\build\Invoke-AllTests.ps1 -SkipBuild # ohne EXE-Bau, falls ps2exe fehlt
```

Vier Prüfungen hinter einem Einstiegspunkt: **Syntax** (jede `.ps1` wird geparst),
**Encoding** (Umlaute nur mit UTF-8-BOM, sonst zeigt PowerShell 5.1 Mojibake),
**Prolog** (die eingebetteten Prüfblöcke stimmen mit ihrer Quelle überein, die
Mindestversion wird erzwungen) und **Release-Logik** (die Schritte aus
`release.yml`, inklusive echtem EXE-Bau).

`.github/workflows/ci.yml` ruft bei jedem Push und jedem Pull Request genau dieses
Skript auf – „lokal grün" und „CI grün" bedeuten damit dasselbe.

Die EXE-Dateien sind **nicht signiert**: Windows SmartScreen meldet sich beim ersten Start,
und manche Virenscanner stufen mit ps2exe erzeugte Dateien als verdächtig ein. Wer das
vermeiden will, nimmt das PowerShell-Skript aus dem Quellordner statt der EXE.

## Herkunft

Dieses Repo führt bisher getrennte Ablagen zusammen: `AcmeExchange` (vormals
`azitc-ac/Exchange-ACME-Cert`), `ExchangeTester` (vormals `azitc-ac/Exchange-Tester`) und
`PublicFolderMigration` (vormals `azitc-ac/PF2SharedMBXOnprem`). Hier liegt jeweils der Stand
zum Zeitpunkt der Zusammenführung; die frühere Entwicklungsgeschichte steht nicht in diesem
Repo, sondern in den drei alten Ablagen. Die übrigen Werkzeuge kamen als Einzelskripte dazu –
jeweils der jüngste Stand aus Servern und OneDrive.

## Lizenz

MIT, soweit in den einzelnen Ordnern nicht anders angegeben.
