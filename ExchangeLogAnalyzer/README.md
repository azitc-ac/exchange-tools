# Exchange Log Analyzer

Wertet die Protokolle eines Exchange-Servers aus, ohne dass man sich durch Textdateien arbeiten
muss: SMTP-Empfangsprotokolle, Nachrichtenverfolgung und IIS-Protokolle. Oberfläche mit
Verzeichnisliste, Filtern und CSV-Ausgabe.

## Starten

```powershell
powershell.exe -STA -ExecutionPolicy Bypass -File ".\Exchange Log Analyzer.ps1"
```

Die Exchange-Verwaltungsshell wird **nicht** gebraucht – gelesen werden nur Dateien. Läuft daher
auch auf einer Kopie der Protokolle abseits des Servers.

## Auswertungen

| Abfrage | Grundlage | Ergebnis |
|---|---|---|
| SMTP-Receive Hits | ProtocolLog\SmtpReceive | Gegenstellen nach IP, getrennt nach EHLO und HELO |
| Server Statistics | MessageTracking | Aufkommen je Server |
| Mails per Day | MessageTracking | Nachrichten je Tag |
| Top 20 Recipients / Senders | MessageTracking | häufigste Empfänger bzw. Absender |
| IIS Hits by IP / by User | IIS-W3C-Protokolle | Zugriffe je Adresse bzw. Konto |

**SMTP-Receive Hits** beantwortet die Frage, wer über einen Connector tatsächlich einliefert –
etwa vor dem Abschalten eines Relay-Connectors. Die Trennung nach EHLO und HELO zeigt nebenbei,
welche Gegenstellen noch mit altem SMTP sprechen.

## Bedienung

- **Verzeichnisse** werden beim Wechsel der Abfrage automatisch vorbelegt: die Protokollpfade
  unter `TransportRoles\Logs\...` (FrontEnd, Hub, Edge) bzw. `inetpub\logs\LogFiles` auf allen
  festen Laufwerken. Eigene Pfade lassen sich ergänzen, auch UNC.
- **Unterverzeichnisse einbeziehen** ist voreingestellt und nötig, damit die `W3SVC*`-Ordner der
  IIS-Websites mitkommen.
- **Reverse DNS** löst die gefundenen Adressen auf – hilfreich zum Zuordnen, kostet aber Zeit.
- **Connector-Filter** (Teilzeichenfolge, Groß-/Kleinschreibung egal, Vorgabe `Relay`) und
  **Zeitraumfilter** grenzen ein. Der Zeitraum greift zweistufig: zuerst über die Dateinamen
  (`…yyyyMMdd-n.LOG` bzw. `u_exYYMMDD.log`), dann zeilenweise.
- **CSV Export** schreibt das angezeigte Ergebnis heraus.

## Zum Tempo

Gelesen und gezählt wird in .NET, parallel über alle Kerne; die Oberfläche bleibt dabei bedienbar.
Der Balken folgt den gelesenen **Bytes**, die Textzeile den fertigen Dateien – bei IIS-Protokollen
liegen zwischen der kleinsten und der größten Datei leicht zwei Zehnerpotenzen, eine Anzeige nach
Dateien würde dort springen. Zwei Dinge sparen den Großteil der Zeit:

- **Ein Durchgang füllt alle Sichten einer Familie.** Teuer ist das Lesen und Zerlegen der Zeile,
  nicht das Zählen – also entstehen Server, Tage, Empfänger und Absender gemeinsam, ebenso die
  beiden IIS-Auswertungen.
- **Die Zählwerke bleiben stehen.** Ein Wechsel der Abfrage im selben Zeitraum zeigt sofort an,
  ohne noch einmal zu lesen. Der Schlüssel enthält jede Datei mit Größe und Änderungszeit sowie
  Zeitraum und Connector-Filter: wächst eine Datei oder kommt eine dazu, wird von selbst neu
  gelesen. Die letzten drei Stände je Familie werden gehalten, sodass auch das Hin und Her
  zwischen zwei Zeiträumen ohne Wartezeit auskommt. Aufgelöste Namen (Reverse DNS) bleiben
  ebenfalls stehen – nur neue Adressen kosten Zeit.

Gemessen auf einem Exchange-Server (Windows Server 2025, 32 Kerne):

| Datenbestand | vorher | jetzt |
|---|---|---|
| 366 MB Message Tracking, erste Abfrage | 11,4 s | 0,9 s |
| dieselben Daten, alle vier Abfragen | 44,1 s | 1,0 s |
| 74 MB SMTP-Receive, 217 Dateien | 5,6 s | 0,6 s |
| 220 MB IIS, beide Auswertungen | 12,1 s | 0,8 s |

Die Ergebnisse sind dieselben wie vorher – nachgeprüft über alle Sichten. Bei gleichem Zählwert
entscheidet jetzt der Name über die Reihenfolge, damit eine Auswertung zweimal dasselbe liefert.

## Test

```powershell
powershell.exe -STA -ExecutionPolicy Bypass -File .\tests\Test-AnalyzerCore.ps1
```

Legt kleine Protokolldateien mit von Hand nachgerechneten Werten an und lässt den echten Code
darauf laufen: alle vier Tracking-Sichten, die beiden Tücken des Formats (ein Betreff mit Komma
verschiebt jedes Feld dahinter; ein Eintrag kann mehrere Empfänger tragen), EHLO/HELO samt
Connector-Filter, IIS nach Adresse und Konto, Zeitraumfilter und das Verhalten des
Zwischenspeichers vor und nach einer Dateiänderung.

## Voraussetzungen

Windows PowerShell 5.1, Start mit `-STA`. Leserecht auf den Protokollverzeichnissen.
