# LogViewer

Betrachter für Protokolldateien: CMTrace, CSV, W3C (IIS) und einfacher Text. Mit Live-
Aktualisierung, Suche und Filter. Läuft eigenständig – keine Exchange-Shell, keine Installation.

![Version](https://img.shields.io/badge/Version-3.0.1-informational)

## Starten

```powershell
LogViewer.exe                                   # leer starten, Datei über "Öffnen" wählen
LogViewer.exe -Path "D:\Logs\app.log"
LogViewer.exe -Path "D:\Logs\iis.log" -Mode W3C -Language de
```

Ohne EXE genauso mit `LogViewer.ps1` (benötigt `-STA`):

```powershell
powershell.exe -STA -ExecutionPolicy Bypass -File .\LogViewer.ps1 -Path "D:\Logs\app.log"
```

| Parameter | |
|---|---|
| `-Path` | Datei, die geöffnet werden soll |
| `-Mode` | `CSV`, `CMTrace`, `DateTime`, `W3C`, `Text` oder `Auto` (Standard: erkennt das Format) |
| `-Language` | `de`, `en` oder `auto` (Standard: folgt der Anzeigesprache von Windows, sonst Englisch) |
| `-Register` / `-Unregister` | Trägt die Dateitypen ein bzw. entfernt sie, ohne das Fenster zu öffnen |

## Was es kann

- **Fünf Formate.** CMTrace mit farbiger Kennzeichnung von Warnungen und Fehlern, CSV nach
  RFC 4180 (auch Felder mit Zeilenumbrüchen), W3C mit den Spalten aus der `#Fields`-Zeile,
  Zeitstempel/Text und reiner Text.
- **Live.** Wächst die Datei, werden nur die neuen Zeilen nachgeladen – ab dem zuletzt gelesenen
  Byte, nicht die ganze Datei erneut. Auto-Scroll hält das Ende im Blick. Wird die Datei
  abgeschnitten oder rotiert, lädt der Betrachter sie neu. Das gilt auch für eine Datei, die
  beim Öffnen noch leer war und erst danach entsteht – etwa ein laufendes MSI-Setup, das in
  UTF-16 schreibt: die Codierung wird nachbestimmt, sobald genug dasteht.
- **Suchen und filtern** über alle Spalten. Bei großen Dateien blockweise mit Fortschritt,
  geschätzter Restzeit und Abbrechen.
- **Zweisprachig** (Deutsch/Englisch), Umschaltung über `-Language`.

## Warnungen und Fehler markieren

Im CMTrace-Format färbt der Betrachter Zeilen anhand der Typspalte: gelb für Warnungen, rot für
Fehler. Andere Protokolle haben keine solche Spalte – dort tragen die Zeilen ihre Schwere im Text.

Die Option **Warnungen/Fehler markieren** in der Werkzeugleiste färbt diese Zeilen in denselben
Farben:

| Farbe | Erkannt an |
|---|---|
| rot | `ERROR`, `FEHLER`, `FATAL`, `CRITICAL`, `SEVERE` – dazu `*ERROR*`, wie der Hybrid Configuration Wizard seine Fehler schreibt |
| gelb | `WARN`, `WARNING`, `WARNUNG` |

Gesucht wird ohne Rücksicht auf Groß- und Kleinschreibung, aber an Wortgrenzen – `Terrorliste`,
`Warner` oder `errorhandling.dll` färben also nicht. Enthält eine Zeile beides, gewinnt Rot.

Die Option ist beim Start aus und wirkt nur in den Formaten ohne Typspalte; CMTrace-Dateien
behalten ihre Färbung unabhängig davon.

## Verhalten bei großen Dateien

Das Lesen und Zerlegen erledigt ein eingebetteter .NET-Kern, nicht PowerShell. Gemessen auf einem
Exchange-Server (Windows Server 2025):

| Datei | Laden | Filtern |
|---|---|---|
| 300.000 Zeilen CMTrace (48 MB) | 2,0 s | 0,17 s |
| 2.000.000 Zeilen CMTrace (283 MB) | 15 s | 2,6 s |

Der Speicherbedarf liegt bei rund 500 MB je Million Zeilen, weil alle Zeilen gehalten werden –
das hält Suche und Filter schnell.

## Dateitypen eintragen

Beim ersten Start fragt der LogViewer einmalig, ob er für `.log`, `.csv` und `.txt` angeboten
werden soll. Danach steht "Mit LogViewer öffnen" im Kontextmenü und der LogViewer unter
"Öffnen mit". Nachholen oder rückgängig machen:

```powershell
.\LogViewer-register.ps1
.\LogViewer-register.ps1 -Remove
```

Alles landet unter `HKCU` – keine Administratorrechte nötig. Welches Programm beim **Doppelklick**
startet, entscheidet Windows selbst; das lässt sich von außen nicht setzen und muss einmalig über
"Öffnen mit > Andere App auswählen" gewählt werden.

## Selbst bauen

```powershell
.\Build-LogViewerExe.ps1        # braucht ps2exe (Install-Module ps2exe -Scope CurrentUser)
.\New-LogViewerIcon.ps1         # zeichnet LogViewer.ico neu
```

Die Versionsnummer steht in `LogViewer.ps1` (`$script:Version`) und wird beim Bauen von dort
übernommen.

## Test

```powershell
powershell.exe -STA -ExecutionPolicy Bypass -File .\tests\Test-GrowingUtf16.ps1
powershell.exe -STA -ExecutionPolicy Bypass -File .\tests\Test-GrowingUtf16.ps1 -Rueckbau
```

Der Test lädt den echten Code (nur ohne Fenster) und lässt eine UTF-16-Datei wachsen, die beim
Öffnen noch leer war. Mit `-Rueckbau` werden die beiden Gegenmaßnahmen im geladenen Code
zurückgenommen – dann muss er fehlschlagen, sonst prüft er nichts.

## Voraussetzungen

Windows PowerShell 5.1 und .NET Framework 4.x – beides auf jedem aktuellen Windows vorhanden.
