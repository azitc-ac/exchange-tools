#Requires -Version 5.1
<#
.SYNOPSIS
    Log Viewer mit CSV-, CMTrace-, DateTime-, W3C- und Text-Modus, Live-Updates, Suche und Filter.

.DESCRIPTION
    Das Lesen, Zerlegen, Filtern und Suchen erledigt ein eingebetteter .NET-Kern (LogCore), nicht
    PowerShell. Grund: pro Zeile ein PowerShell-Funktionsaufruf samt Regex und PSCustomObject
    kostete rund 180 Mikrosekunden - bei 300.000 Zeilen also knapp eine Minute. Derselbe Ablauf in
    C# braucht dafür ein bis zwei Sekunden.

    Ebenfalls neu: Nachladen bei Live-Updates geschieht ab dem zuletzt gelesenen Byte-Offset,
    auch im CSV-Modus. Vorher wurde dort bei jeder Dateiänderung die komplette Datei neu
    eingelesen - alle 500 ms.

.PARAMETER Path
    Pfad zur Datei (optional - ohne Angabe erscheint der Öffnen-Dialog).
.PARAMETER Mode
    Startmodus: CSV, CMTrace, DateTime, W3C, Text oder Auto (Standard).
.PARAMETER Language
    Oberflächensprache: de, en oder auto (Standard). "auto" folgt der Anzeigesprache von Windows
    und fällt auf Englisch zurück, wenn diese nicht Deutsch ist.
.PARAMETER Register
    Trägt die Dateitypen ein (Kontextmenü und "Öffnen mit") und beendet sich wieder.
.PARAMETER Unregister
    Entfernt diese Einträge wieder und beendet sich.
.PARAMETER Extensions
    Dateiendungen für -Register/-Unregister und für die Rückfrage beim ersten Start.
#>
Param(
    [string]$Path = '',
    [ValidateSet('CSV', 'CMTrace', 'DateTime', 'W3C', 'Text', 'Auto')]
    [string]$Mode = 'Auto',
    [ValidateSet('de', 'en', 'auto')]
    [string]$Language = 'auto',
    [switch]$Register,
    [switch]$Unregister,
    [string[]]$Extensions = @('.log', '.csv', '.txt'),
    [string]$TargetPath = ''   # Programm, das registriert werden soll (Standard: dieses)
)

# Beim Aufruf über "powershell.exe -File" kommt eine Liste als EINE Zeichenkette an
# (".log,.csv,.txt"). Deshalb hier auftrennen - so funktionieren beide Aufrufformen.
$Extensions = @($Extensions | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })

# ── Sprache ───────────────────────────────────────────────────────────────────
# Alle sichtbaren Texte stehen hier; Deutsch nur, wenn Windows auf Deutsch läuft.
$script:Lang = if ($Language -eq 'auto') {
    if ([System.Globalization.CultureInfo]::CurrentUICulture.TwoLetterISOLanguageName -eq 'de') { 'de' } else { 'en' }
} else { $Language }

$script:TextTable = @{
    de = @{
        WindowTitle    = 'Log Viewer'
        Open           = 'Öffnen…'
        ModeLabel      = 'Modus:'
        ModeDateTime   = 'Datum/Text'
        AutoScroll     = 'Auto-Scroll'
        SearchLabel    = 'Suchen:'
        SearchNext     = 'Weiter ▶'
        FilterLabel    = 'Filter:'
        FilterApply    = 'Anwenden'
        FilterClearTip = 'Filter aufheben'
        CopyRows       = 'Ausgewählte Zeilen kopieren'
        NoFile         = 'Keine Datei geladen.  Öffnen-Button oder Drag & Drop verwenden.'
        NotFoundTitle  = 'Suchen'
        NotFound       = "'{0}' nicht gefunden."
        ReadError      = "Die Datei konnte nicht gelesen werden:`n{0}"
        Rows           = '{0} Zeilen'
        RowsFiltered   = '{0} von {1} Zeilen (gefiltert)'
        StatusMode     = 'Modus: {0}'
        StatusDelim    = "Trennzeichen: '{0}'"
        StatusColumns  = 'Spalten: {0}'
        FileFilter     = 'Log & CSV (*.log;*.csv;*.txt)|*.log;*.csv;*.txt|Alle Dateien (*.*)|*.*'
        ColDate        = 'Datum';       ColTime    = 'Zeit'
        ColType        = 'Typ';         ColComp    = 'Komponente'
        ColMessage     = 'Nachricht';   ColThread  = 'Thread'
        ColTimestamp   = 'Zeitstempel'; ColLine    = 'Zeile'
        TypeInfo       = 'Info';        TypeWarn   = 'Warnung';  TypeError = 'Fehler'
        AssocTitle     = 'LogViewer für Protokolldateien registrieren?'
        AssocQuestion  = @'
Soll LogViewer künftig für {0} angeboten werden?

Dann erscheint "Mit LogViewer öffnen" im Kontextmenü, und LogViewer steht unter "Öffnen mit" zur Auswahl.

Hinweis: Welches Programm beim Doppelklick startet, entscheidet Windows selbst - das lässt sich von außen nicht mehr setzen. Die Umstellung nehmen Sie bei Bedarf einmalig über "Öffnen mit > Andere App auswählen" vor.
'@
        AssocDone      = 'Eingetragen für: {0}'
        AssocFailed    = 'Das Eintragen ist fehlgeschlagen:{0}{1}'
        AssocRemoved   = 'Einträge entfernt für: {0}'
        About          = 'Über'
        AboutTitle     = 'Über LogViewer'
        AboutVersion   = 'Version {0}'
        AboutText      = 'Log-Betrachter für CMTrace-, CSV-, W3C- und Textprotokolle.{0}Live-Aktualisierung, Suche und Filter.'
        AboutBlog      = 'Blog:'
        AboutClose     = 'Schließen'
        Cancel         = 'Abbrechen'
        FilterProgress = 'Filtern…'
        SearchProgress = 'Suchen…'
        ProgressRows   = '{0:N0} von {1:N0} Zeilen'
        RemainSoon     = 'gleich fertig'
        RemainSec      = 'noch ca. {0} s'
        RemainMin      = 'noch ca. {0} min'
    }
    en = @{
        WindowTitle    = 'Log Viewer'
        Open           = 'Open…'
        ModeLabel      = 'Mode:'
        ModeDateTime   = 'Date/Text'
        AutoScroll     = 'Auto-scroll'
        SearchLabel    = 'Find:'
        SearchNext     = 'Next ▶'
        FilterLabel    = 'Filter:'
        FilterApply    = 'Apply'
        FilterClearTip = 'Clear filter'
        CopyRows       = 'Copy selected rows'
        NoFile         = 'No file loaded.  Use the Open button or drag and drop.'
        NotFoundTitle  = 'Find'
        NotFound       = "'{0}' not found."
        ReadError      = "The file could not be read:`n{0}"
        Rows           = '{0} rows'
        RowsFiltered   = '{0} of {1} rows (filtered)'
        StatusMode     = 'Mode: {0}'
        StatusDelim    = "Delimiter: '{0}'"
        StatusColumns  = 'Columns: {0}'
        FileFilter     = 'Log & CSV (*.log;*.csv;*.txt)|*.log;*.csv;*.txt|All files (*.*)|*.*'
        ColDate        = 'Date';        ColTime    = 'Time'
        ColType        = 'Type';        ColComp    = 'Component'
        ColMessage     = 'Message';     ColThread  = 'Thread'
        ColTimestamp   = 'Timestamp';   ColLine    = 'Line'
        TypeInfo       = 'Info';        TypeWarn   = 'Warning'; TypeError = 'Error'
        AssocTitle     = 'Register LogViewer for log files?'
        AssocQuestion  = @'
Would you like LogViewer to be offered for {0}?

"Open with LogViewer" will then appear in the context menu, and LogViewer becomes selectable under "Open with".

Note: which program starts on a double click is decided by Windows itself and can no longer be set from outside. Change it once via "Open with > Choose another app" if you want that.
'@
        AssocDone      = 'Registered for: {0}'
        AssocFailed    = 'Registration failed:{0}{1}'
        AssocRemoved   = 'Entries removed for: {0}'
        About          = 'About'
        AboutTitle     = 'About LogViewer'
        AboutVersion   = 'Version {0}'
        AboutText      = 'Log viewer for CMTrace, CSV, W3C and plain text logs.{0}Live updates, search and filter.'
        AboutBlog      = 'Blog:'
        AboutClose     = 'Close'
        Cancel         = 'Cancel'
        FilterProgress = 'Filtering…'
        SearchProgress = 'Searching…'
        ProgressRows   = '{0:N0} of {1:N0} rows'
        RemainSoon     = 'almost done'
        RemainSec      = 'about {0} s left'
        RemainMin      = 'about {0} min left'
    }
}
$script:T = $script:TextTable[$script:Lang]

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ── Dateitypen ────────────────────────────────────────────────────────────────
# Eintragen, Prüfen und Entfernen liegen hier zusammen; LogViewer-register.ps1 ruft nur noch
# dieses Programm auf, damit es keine zweite Fassung derselben Logik gibt.
$script:Version  = '3.0.1'
$script:BlogUrl  = 'https://blog.zarenko.net'
$script:ProgId   = 'LogViewer.LogFile'
$script:StateKey = 'HKCU:\Software\LogViewer'

function Get-LauncherCommand {
    <#
      Liefert den Startbefehl - als EXE direkt, als Skript über powershell.exe.
      Mit -TargetPath lässt sich ein anderes Programm eintragen; das nutzt der Wrapper, um die
      EXE zu registrieren, während die Registrierung selbst im Skript läuft (in einer ps2exe-EXE
      beendet sich der Vorgang nicht zuverlässig von selbst).
    #>
    if ($TargetPath) {
        if ($TargetPath.ToLower().EndsWith('.exe')) {
            return [pscustomobject]@{ Command = ('"{0}" -Path "%1"' -f $TargetPath); Icon = "$TargetPath,0"; Target = $TargetPath }
        }
        return [pscustomobject]@{
            Command = ('powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "{0}" -Path "%1"' -f $TargetPath)
            Icon = 'powershell.exe'; Target = $TargetPath }
    }
    $proc = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $isExe = ([IO.Path]::GetFileNameWithoutExtension($proc)) -notin @('powershell', 'pwsh', 'powershell_ise')
    if ($isExe) {
        return [pscustomobject]@{ Command = ('"{0}" -Path "%1"' -f $proc); Icon = "$proc,0"; Target = $proc }
    }
    $script = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
    return [pscustomobject]@{
        Command = ('powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "{0}" -Path "%1"' -f $script)
        Icon    = 'powershell.exe'
        Target  = $script
    }
}

function Get-AssocKeys {
    param([string[]]$Ext)
    @(foreach ($e in $Ext) { "HKCU:\Software\Classes\SystemFileAssociations\$e\shell\OpenWithLogViewer" })
}

function Test-FileTypesRegistered {
    param([string[]]$Ext)
    foreach ($k in (Get-AssocKeys $Ext)) { if (-not (Test-Path $k)) { return $false } }
    return $true
}

function Register-FileTypes {
    param([string[]]$Ext)
    $l = Get-LauncherCommand
    # 1. Eigene ProgId - damit taucht LogViewer unter "Öffnen mit" auf.
    $progRoot = "HKCU:\Software\Classes\$script:ProgId"
    New-Item -Path "$progRoot\shell\open\command" -Force | Out-Null
    New-Item -Path "$progRoot\DefaultIcon" -Force | Out-Null
    Set-ItemProperty -Path $progRoot -Name '(default)' -Value $script:T.WindowTitle
    Set-Item -Path "$progRoot\shell\open\command" -Value $l.Command
    Set-Item -Path "$progRoot\DefaultIcon" -Value $l.Icon
    # 2. Kontextmenü je Endung + ProgId als Angebot eintragen
    foreach ($e in $Ext) {
        $k = "HKCU:\Software\Classes\SystemFileAssociations\$e\shell\OpenWithLogViewer"
        New-Item -Path "$k\command" -Force | Out-Null
        Set-ItemProperty -Path $k -Name 'MUIVerb' -Value ('{0} {1}' -f $script:T.Open.TrimEnd([char]0x2026), $script:T.WindowTitle)
        Set-ItemProperty -Path $k -Name 'Icon' -Value $l.Icon
        Set-Item -Path "$k\command" -Value $l.Command
        $owp = "HKCU:\Software\Classes\$e\OpenWithProgids"
        New-Item -Path $owp -Force | Out-Null
        New-ItemProperty -Path $owp -Name $script:ProgId -Value ([byte[]]@()) -PropertyType None -Force | Out-Null
    }
}

function Unregister-FileTypes {
    param([string[]]$Ext)
    foreach ($e in $Ext) {
        $k = "HKCU:\Software\Classes\SystemFileAssociations\$e\shell\OpenWithLogViewer"
        if (Test-Path $k) { Remove-Item $k -Recurse -Force }
        $owp = "HKCU:\Software\Classes\$e\OpenWithProgids"
        if (Test-Path $owp) { Remove-ItemProperty -Path $owp -Name $script:ProgId -ErrorAction SilentlyContinue }
    }
    $progRoot = "HKCU:\Software\Classes\$script:ProgId"
    if (Test-Path $progRoot) { Remove-Item $progRoot -Recurse -Force }
}

# Nicht-interaktive Wege: eintragen bzw. entfernen und beenden.
# "exit" reicht dafür nicht: in einer mit ps2exe erzeugten EXE läuft das Skript danach weiter
# und öffnet das Fenster. Deshalb zusätzlich ein Schalter, der den Start der Oberfläche verhindert.
$script:HeadlessOnly = $false
if ($Register -or $Unregister) {
    $script:HeadlessOnly = $true
    try {
        if ($Unregister) {
            Unregister-FileTypes -Ext $Extensions
            Write-Host ($script:T.AssocRemoved -f ($Extensions -join ', '))
        } else {
            Register-FileTypes -Ext $Extensions
            Write-Host ($script:T.AssocDone -f ($Extensions -join ', '))
            Write-Host ("  {0}" -f (Get-LauncherCommand).Command)
        }
        # [Environment]::Exit beendet den Prozess sofort - "exit" allein wirkt in einer mit
        # ps2exe erzeugten EXE nicht und das Programm liefe danach weiter.
        [Environment]::Exit(0)
    } catch {
        Write-Host ($script:T.AssocFailed -f ' ', $_.Exception.Message) -ForegroundColor Red
        [Environment]::Exit(1)
    }
}

# ── Kern ──────────────────────────────────────────────────────────────────────
# Liest blockweise ab einem Byte-Offset, damit die Oberfläche zwischendurch den Fortschritt
# zeichnen kann und Live-Updates nur das Neue lesen.
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

public class LogChunk
{
    public List<string[]> Rows;
    public long EndOffset;
    public bool Eof;
    public string[] Headers;      // nur beim W3C-Modus belegt
    public string Delimiter;      // nur bei CSV/W3C belegt
    public Encoding Encoding;     // was der Leser tatsächlich benutzt hat (siehe Load)
}

public static class LogCore
{
    const int BufferSize = 1 << 16;

    static readonly Regex CmTrace = new Regex(
        "<!\\[LOG\\[(?<msg>.*?)\\]LOG\\]!><time=\"(?<time>[^\"]+)\" date=\"(?<date>[^\"]+)\" component=\"(?<comp>[^\"]*)\" context=\"[^\"]*\" type=\"(?<type>\\d)\" thread=\"(?<thread>[^\"]*)\"",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);
    static readonly Regex TzSuffix = new Regex("[+-]\\d+$", RegexOptions.Compiled);
    static readonly Regex DtA = new Regex(
        @"^(?<dt>\d{1,4}[-./]\d{1,2}[-./]\d{1,4}[\sT]\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:[+-]\d{2}:?\d{2}|Z)?):?\s+(?<msg>.*)$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);
    static readonly Regex DtB = new Regex(
        @"^\[(?<dt>[^\]]+)\]\s*(?:\[\d+\])?\s*(?<msg>.*)$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);
    static readonly Regex W3cSpace = new Regex(" +", RegexOptions.Compiled);

    /// <summary>Beschriftung der CMTrace-Typen (Index 1..3). Wird beim Start aus der
    /// Sprachtabelle gesetzt, damit die Umschaltung keine Kosten pro Zeile verursacht.</summary>
    public static string[] TypeNames = { "", "Info", "Warning", "Error" };

    // ---------------------------------------------------------------- Dateiinfo
    public static Encoding DetectEncoding(string path)
    {
        try
        {
            var buf = new byte[4];
            int n;
            using (var fs = File.OpenRead(path)) { n = fs.Read(buf, 0, 4); }
            if (n >= 2 && buf[0] == 0xFF && buf[1] == 0xFE) return Encoding.Unicode;
            if (n >= 2 && buf[0] == 0xFE && buf[1] == 0xFF) return Encoding.BigEndianUnicode;
            if (n >= 3 && buf[0] == 0xEF && buf[1] == 0xBB && buf[2] == 0xBF) return new UTF8Encoding(true);
            if (n >= 4 && buf[1] == 0 && buf[3] == 0) return Encoding.Unicode;
        }
        catch { }
        return new UTF8Encoding(false);
    }

    public static string[] FirstLines(string path, Encoding enc, int count)
    {
        var res = new List<string>();
        try
        {
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, BufferSize))
            using (var sr = new StreamReader(fs, enc, true))
            {
                string line;
                while (res.Count < count && (line = sr.ReadLine()) != null) res.Add(line);
            }
        }
        catch { }
        return res.ToArray();
    }

    public static string[] ReadW3cFields(string path, Encoding enc, out string delimiter)
    {
        delimiter = " ";
        try
        {
            foreach (var line in FirstLines(path, enc, 50))
            {
                var t = line.TrimEnd();
                if (!t.TrimStart().StartsWith("#")) break;
                var m = Regex.Match(t, @"(?i)^#Fields:\s+(.+)$");
                if (m.Success)
                {
                    var fp = m.Groups[1].Value.Trim();
                    delimiter = fp.Contains(",") ? "," : " ";
                    var parts = delimiter == " " ? W3cSpace.Split(fp) : fp.Split(',');
                    var list = new List<string>();
                    foreach (var p in parts) { var v = p.Trim(); if (v.Length > 0) list.Add(v); }
                    return list.ToArray();
                }
            }
        }
        catch { }
        return new string[0];
    }

    // ---------------------------------------------------------------- CSV
    // Eigener RFC-4180-Leser. Er beherrscht Felder in Anführungszeichen samt eingebetteten
    // Zeilenumbrüchen und doppelten Anführungszeichen - und ist um ein Vielfaches schneller als
    // der TextFieldParser aus Microsoft.VisualBasic.
    static List<string> ParseCsvRecord(StreamReader sr, char delim, ref bool eof)
    {
        var fields = new List<string>();
        var sb = new StringBuilder(64);
        bool inQuotes = false, any = false;
        while (true)
        {
            int ci = sr.Read();
            if (ci < 0) { eof = true; break; }
            char c = (char)ci;
            any = true;
            if (inQuotes)
            {
                if (c == '"')
                {
                    if (sr.Peek() == '"') { sr.Read(); sb.Append('"'); }
                    else inQuotes = false;
                }
                else sb.Append(c);
            }
            else
            {
                if (c == '"' && sb.Length == 0) inQuotes = true;
                else if (c == delim) { fields.Add(sb.ToString()); sb.Length = 0; }
                else if (c == '\r') { if (sr.Peek() == '\n') sr.Read(); break; }
                else if (c == '\n') break;
                else sb.Append(c);
            }
        }
        if (!any && fields.Count == 0 && sb.Length == 0) return null;
        fields.Add(sb.ToString());
        return fields;
    }

    public static char DetectCsvDelimiter(string path, Encoding enc)
    {
        var lines = FirstLines(path, enc, 1);
        if (lines.Length > 0 && lines[0].Contains(";") && !lines[0].Contains(",")) return ';';
        return ',';
    }

    // ---------------------------------------------------------------- Laden
    public static LogChunk Load(string path, string mode, Encoding enc, long startOffset,
                                int maxLines, string[] w3cHeaders, string w3cDelimiter, char csvDelimiter)
    {
        var res = new LogChunk { Rows = new List<string[]>(), EndOffset = startOffset, Eof = true,
                                 Delimiter = w3cDelimiter, Encoding = enc };
        using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, BufferSize))
        {
            if (startOffset > fs.Length) startOffset = 0;
            fs.Seek(startOffset, SeekOrigin.Begin);
            using (var sr = new StreamReader(fs, enc, startOffset == 0))
            {
                int n = 0;
                if (mode == "CSV")
                {
                    bool eof = false;
                    while (n < maxLines)
                    {
                        var rec = ParseCsvRecord(sr, csvDelimiter, ref eof);
                        if (rec == null) break;
                        // eine letzte leere Zeile am Dateiende nicht als Datensatz zählen
                        if (eof && rec.Count == 1 && rec[0].Length == 0) break;
                        res.Rows.Add(rec.ToArray());
                        n++;
                        if (eof) break;
                    }
                    res.Eof = eof || sr.EndOfStream;
                }
                else
                {
                    string line;
                    while (n < maxLines && (line = sr.ReadLine()) != null)
                    {
                        switch (mode)
                        {
                            case "CMTrace":
                            {
                                var m = CmTrace.Match(line);
                                if (m.Success)
                                {
                                    int t; int.TryParse(m.Groups["type"].Value, out t);
                                    res.Rows.Add(new string[] {
                                        m.Groups["date"].Value,
                                        TzSuffix.Replace(m.Groups["time"].Value, ""),
                                        (t >= 0 && t < TypeNames.Length) ? TypeNames[t] : "",
                                        m.Groups["comp"].Value,
                                        m.Groups["msg"].Value,
                                        m.Groups["thread"].Value,
                                        t.ToString() });
                                    n++;
                                }
                                else if (line.Trim().Length > 0)
                                {
                                    res.Rows.Add(new string[] { "", "", "", "", line, "", "1" });
                                    n++;
                                }
                                break;
                            }
                            case "W3C":
                            {
                                if (line.Length == 0 || line[0] == '#' || line.Trim().Length == 0) break;
                                var parts = w3cDelimiter == " " ? W3cSpace.Split(line) : line.Split(w3cDelimiter[0]);
                                int want = w3cHeaders.Length;
                                var row = new string[want];
                                for (int i = 0; i < want; i++) row[i] = i < parts.Length ? parts[i] : "";
                                res.Rows.Add(row);
                                n++;
                                break;
                            }
                            case "Text":
                                res.Rows.Add(new string[] { line });
                                n++;
                                break;
                            default:
                            {
                                var m = DtA.Match(line);
                                if (!m.Success) m = DtB.Match(line);
                                if (m.Success) { res.Rows.Add(new string[] { m.Groups["dt"].Value, m.Groups["msg"].Value }); n++; }
                                else if (line.Trim().Length > 0) { res.Rows.Add(new string[] { "", line }); n++; }
                                break;
                            }
                        }
                    }
                    res.Eof = sr.EndOfStream;
                }
                // Ab Offset 0 darf der Leser die Codierung am BOM selbst bestimmen - er kann also
                // eine andere benutzen als die uebergebene. Diese hier zurueckmelden, sonst liest
                // der naechste Chunk (ohne BOM-Erkennung) mit der falschen und die Byteposition
                // unten wird ebenfalls falsch berechnet.
                res.Encoding = sr.CurrentEncoding;
                // Position des Lesers, nicht des Streams - der StreamReader puffert voraus.
                res.EndOffset = GetReaderPosition(sr, fs, sr.CurrentEncoding);
            }
        }
        return res;
    }

    // Der StreamReader liest in Blöcken voraus; fs.Position zeigt deshalb zu weit. Die tatsächlich
    // verarbeitete Byte-Position ergibt sich aus Streamposition minus dem, was noch im Puffer liegt.
    static long GetReaderPosition(StreamReader sr, FileStream fs, Encoding enc)
    {
        try
        {
            var t = sr.GetType();
            var charLen = (int)t.GetField("charLen", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance).GetValue(sr);
            var charPos = (int)t.GetField("charPos", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance).GetValue(sr);
            var charBuffer = (char[])t.GetField("charBuffer", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance).GetValue(sr);
            int pending = enc.GetByteCount(charBuffer, charPos, charLen - charPos);
            return fs.Position - pending;
        }
        catch { return fs.Position; }
    }

    // ---------------------------------------------------------------- Filter und Suche
    public static int Filter(List<string[]> src, List<string[]> dst, string needle, int columns)
    {
        dst.Clear();
        if (string.IsNullOrEmpty(needle)) return 0;
        foreach (var row in src)
        {
            int n = Math.Min(columns, row.Length);
            for (int c = 0; c < n; c++)
            {
                var v = row[c];
                if (v != null && v.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0) { dst.Add(row); break; }
            }
        }
        return dst.Count;
    }

    /// <summary>Filtert nur den Abschnitt [from, from+count) und hängt Treffer an dst an.
    /// Gibt den nächsten Startindex zurück, damit die Oberfläche zwischendurch zeichnen,
    /// eine Restzeit schätzen und abbrechen kann.</summary>
    public static int FilterRange(List<string[]> src, List<string[]> dst, string needle, int columns, int from, int count)
    {
        if (string.IsNullOrEmpty(needle)) return src.Count;
        int end = Math.Min(from + count, src.Count);
        for (int i = from; i < end; i++)
        {
            var row = src[i];
            int n = Math.Min(columns, row.Length);
            for (int c = 0; c < n; c++)
            {
                var v = row[c];
                if (v != null && v.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0) { dst.Add(row); break; }
            }
        }
        return end;
    }

    /// <summary>Sucht nur im Abschnitt [from, from+count) - für abbrechbare Suchläufe.
    /// Rückgabe: Trefferindex, oder -1 wenn im Abschnitt nichts gefunden wurde.</summary>
    public static int FindRange(List<string[]> rows, string needle, int from, int count, int columns)
    {
        if (string.IsNullOrEmpty(needle)) return -1;
        int total = rows.Count;
        if (total == 0) return -1;
        for (int k = 0; k < count; k++)
        {
            int idx = (from + k) % total;
            if (Matches(rows[idx], needle, columns)) return idx;
        }
        return -1;
    }

    public static bool Matches(string[] row, string needle, int columns)
    {
        if (string.IsNullOrEmpty(needle)) return true;
        int n = Math.Min(columns, row.Length);
        for (int c = 0; c < n; c++)
        {
            var v = row[c];
            if (v != null && v.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0) return true;
        }
        return false;
    }

    /// <summary>Sucht ab startIndex und läuft einmal umlaufend durch. -1 = nicht gefunden.</summary>
    public static int Find(List<string[]> rows, string needle, int startIndex, int columns)
    {
        int count = rows.Count;
        if (count == 0 || string.IsNullOrEmpty(needle)) return -1;
        for (int i = 0; i < count; i++)
        {
            int idx = (startIndex + i) % count;
            if (Matches(rows[idx], needle, columns)) return idx;
        }
        return -1;
    }
}
'@ -Language CSharp

# Typbeschriftungen an den Kern durchreichen (einmalig, nicht pro Zeile)
[LogCore]::TypeNames = [string[]]@('', $script:T.TypeInfo, $script:T.TypeWarn, $script:T.TypeError)

# ── Zustand ───────────────────────────────────────────────────────────────────
$s = [PSCustomObject]@{
    FilePath      = [string]''
    Mode          = 'CSV'
    Delimiter     = ','
    ByteOffset    = 0L
    LastWriteTime = [datetime]::MinValue
    LastLength    = -1L
    ManualMode    = $false
    W3cDelimiter  = ' '
    W3cHeaders    = [string[]]@()
    Encoding      = [System.Text.Encoding]::UTF8
    EncodingProvisional = $false   # true, solange die Datei zum Erkennen noch zu kurz war
    VirtualRows   = [System.Collections.Generic.List[string[]]]::new()
    FilteredRows  = [System.Collections.Generic.List[string[]]]::new()
    FilterActive  = $false
    SearchFrom    = 0
    Loading       = $false
}
# Direkt gehaltene Referenz auf die gerade angezeigte Liste: CellValueNeeded feuert pro Zelle,
# ein Funktionsaufruf an dieser Stelle kostet spürbar Zeit.
$script:CurrentRows = $s.VirtualRows

function Set-GridRowCount {
    <#
      Setzt die Zeilenzahl des Grids - beim VERKLEINERN immer über den Zwischenschritt 0.

      Grund (gemessen): Das DataGridView entfernt im VirtualMode jede wegfallende Zeile einzeln.
      Von 2.000.000 auf 400.000 direkt dauert das rund 173 Sekunden, über 0 nur 0,14 Sekunden.
      Genau das ließ die Anwendung nach einem Filterlauf scheinbar endlos hängen.
    #>
    param([int]$Count)
    if ($Count -lt $grid.RowCount) { $grid.RowCount = 0 }
    $grid.RowCount = $Count
}

function Set-CurrentRows {
    # ACHTUNG: "$x = if (...) { $liste }" gibt die Liste durch die Pipeline und PowerShell entrollt
    # sie dabei zu einem Array-Schnappschuss. Die Anzeige hinge dann an einer Kopie und bekäme
    # nachgeladene Zeilen nie zu sehen. Deshalb zwei getrennte Zuweisungen.
    if ($s.FilterActive) { $script:CurrentRows = $s.FilteredRows }
    else                 { $script:CurrentRows = $s.VirtualRows }
}
# Aus demselben Grund gibt es keine Get-Rows-Funktion mehr: Ein Rückgabewert würde bei jedem
# Aufruf die komplette Liste kopieren. Überall wird direkt $script:CurrentRows verwendet.

# ── Modus erkennen ───────────────────────────────────────────────────────────
function Get-AutoMode([string]$path) {
    $ext = [System.IO.Path]::GetExtension($path).ToLower()
    if ($ext -eq '.csv') { return 'CSV' }
    try {
        $enc = [LogCore]::DetectEncoding($path)
        $head = [LogCore]::FirstLines($path, $enc, 20)
        if ($head.Count -eq 0) { return 'Text' }
        $first = $head[0]
        if ($first -match '<!\[LOG\[') { return 'CMTrace' }
        if ($first.StartsWith('#')) {
            foreach ($l in $head) { if ($l -match '^#Fields:') { return 'W3C' } }
        }
        if ($first -match '^\d{1,4}[-./]\d{1,2}[-./]\d{1,4}[\sT]\d{2}:\d{2}:\d{2}' -or $first -match '^\[') { return 'DateTime' }
    } catch {}
    return 'Text'
}

# ── Grid-Spalten ──────────────────────────────────────────────────────────────
function New-Column([string]$name, [int]$width, [bool]$fill = $false) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.HeaderText = $name; $col.Name = $name
    $col.AutoSizeMode = if ($fill) { 'Fill' } else { 'None' }
    $col.Width = $width
    $col.SortMode = 'NotSortable'
    return $col
}
function Set-CsvColumns([string[]]$h) {
    $grid.Columns.Clear(); foreach ($n in $h) { [void]$grid.Columns.Add((New-Column $n 150)) }
}
function Set-CmTraceColumns {
    $grid.Columns.Clear()
    foreach ($c in @(@{N=$script:T.ColDate;W=85}, @{N=$script:T.ColTime;W=100}, @{N=$script:T.ColType;W=65},
                     @{N=$script:T.ColComp;W=120}, @{N=$script:T.ColMessage;W=600}, @{N=$script:T.ColThread;W=60})) {
        [void]$grid.Columns.Add((New-Column $c.N $c.W))
    }
}
function Set-DateTimeColumns {
    $grid.Columns.Clear()
    [void]$grid.Columns.Add((New-Column $script:T.ColTimestamp 175))
    [void]$grid.Columns.Add((New-Column $script:T.ColMessage 200 $true))
}
function Set-W3cColumns([string[]]$h) {
    $grid.Columns.Clear(); foreach ($n in $h) { [void]$grid.Columns.Add((New-Column $n 120)) }
}

# ── Laden ─────────────────────────────────────────────────────────────────────
function Invoke-FullLoad {
    if ($s.Loading) { return }
    $s.Loading = $true
    $form.Cursor = [System.Windows.Forms.Cursors]::AppStarting
    $progressBar.Height = 10
    $pbChunk.Height = 10; $pbChunk.Width = 0
    [System.Windows.Forms.Application]::DoEvents()

    $grid.RowCount = 0
    $s.VirtualRows.Clear()
    $s.FilteredRows.Clear()
    $s.FilterActive = $false
    $s.ByteOffset = 0L
    $s.SearchFrom = 0
    $filterBox.Text = ''
    Set-CurrentRows

    try {
        $s.Encoding = [LogCore]::DetectEncoding($s.FilePath)
        $fileSize = (New-Object System.IO.FileInfo($s.FilePath)).Length
        # Bei einer Datei, die gerade erst angelegt wurde, ist noch nichts zu erkennen. Die
        # Entscheidung gilt dann nur vorläufig und wird beim Nachladen wiederholt - sonst bliebe
        # ein UTF-16-Protokoll ohne BOM (msiexec schreibt so) dauerhaft 8-bittig gelesen.
        $s.EncodingProvisional = ($fileSize -lt 4)

        switch ($s.Mode) {
            'CSV' {
                $s.Delimiter = [LogCore]::DetectCsvDelimiter($s.FilePath, $s.Encoding)
                # Kopfzeile separat lesen, damit die Spalten stehen, bevor Daten kommen.
                $head = [LogCore]::Load($s.FilePath, 'CSV', $s.Encoding, 0, 1, @(), ' ', $s.Delimiter)
                if ($head.Rows.Count -gt 0) { Set-CsvColumns $head.Rows[0] } else { $grid.Columns.Clear() }
                $s.ByteOffset = $head.EndOffset
                if ($head.Encoding) { $s.Encoding = $head.Encoding }
            }
            'CMTrace' { Set-CmTraceColumns }
            'W3C' {
                $delim = ' '
                $s.W3cHeaders = [LogCore]::ReadW3cFields($s.FilePath, $s.Encoding, [ref]$delim)
                $s.W3cDelimiter = $delim
                if ($s.W3cHeaders.Count -gt 0) { Set-W3cColumns $s.W3cHeaders } else { $grid.Columns.Clear() }
            }
            'Text' {
                $grid.Columns.Clear()
                [void]$grid.Columns.Add((New-Column $script:T.ColLine 200 $true))
            }
            default { Set-DateTimeColumns }
        }

        if ($s.Mode -ne 'W3C' -or $s.W3cHeaders.Count -gt 0) {
            # Blockweise lesen: schnell, aber die Oberfläche kann zwischendurch zeichnen.
            $chunkSize = 50000
            while ($true) {
                $chunk = [LogCore]::Load($s.FilePath, $s.Mode, $s.Encoding, $s.ByteOffset, $chunkSize,
                                         $s.W3cHeaders, $s.W3cDelimiter, $s.Delimiter)
                if ($chunk.Rows.Count -gt 0) { $s.VirtualRows.AddRange($chunk.Rows) }
                $s.ByteOffset = $chunk.EndOffset
                # Der erste Block liest ab Offset 0 und erkennt die Codierung am BOM selbst.
                # Sie muss übernommen werden, sonst liest jeder weitere Block falsch.
                if ($chunk.Encoding) { $s.Encoding = $chunk.Encoding }
                if ($fileSize -gt 0) {
                    $pbChunk.Width = [int]($progressBar.Width * [Math]::Min(1.0, [double]$s.ByteOffset / $fileSize))
                }
                [System.Windows.Forms.Application]::DoEvents()
                if ($chunk.Eof -or $chunk.Rows.Count -eq 0) { break }
            }
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show(($script:T.ReadError -f $_.Exception.Message),
            $script:T.WindowTitle, 'OK', 'Warning') | Out-Null
    }

    Set-CurrentRows
    Set-GridRowCount $s.VirtualRows.Count
    $pbChunk.Width = $progressBar.Width
    $progressBar.Height = 0
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
    $s.Loading = $false
    Update-Status
    Invoke-AutoScroll
}

function Invoke-Append {
    if ($s.Loading) { return }
    try {
        if ($s.EncodingProvisional) {
            # Die Datei war beim Öffnen noch leer. Jetzt, wo etwas drinsteht, neu bestimmen.
            $len = (New-Object System.IO.FileInfo($s.FilePath)).Length
            if ($len -ge 4) {
                $enc = [LogCore]::DetectEncoding($s.FilePath)
                $s.EncodingProvisional = $false
                if ($enc.CodePage -ne $s.Encoding.CodePage) {
                    # Andere Codierung als angenommen: von vorn lesen, sonst bleiben die
                    # bereits angezeigten Zeilen und die Byteposition falsch.
                    Invoke-FullLoad
                    return
                }
            }
        }
        $chunk = [LogCore]::Load($s.FilePath, $s.Mode, $s.Encoding, $s.ByteOffset, [int]::MaxValue,
                                 $s.W3cHeaders, $s.W3cDelimiter, $s.Delimiter)
        if ($chunk.Rows.Count -eq 0) { return }
        $s.ByteOffset = $chunk.EndOffset
        if ($chunk.Encoding) { $s.Encoding = $chunk.Encoding }
        $s.VirtualRows.AddRange($chunk.Rows)
        if ($s.FilterActive) {
            $needle = $filterBox.Text
            $cols = $grid.Columns.Count
            foreach ($row in $chunk.Rows) {
                if ([LogCore]::Matches($row, $needle, $cols)) { $s.FilteredRows.Add($row) }
            }
        }
        Set-GridRowCount $script:CurrentRows.Count
        Update-Status
        Invoke-AutoScroll
    } catch {}
}

# ── Filter / Suche ────────────────────────────────────────────────────────────
# Ab dieser Zeilenzahl wird blockweise gearbeitet und ein Fortschrittsfenster gezeigt.
$script:ProgressThreshold = 150000
$script:ChunkRows         = 50000

function New-ProgressDialog {
    <# Kleines Fenster mit Balken, Restzeit und Abbrechen. Gibt die Steuerelemente zurück. #>
    param([string]$Caption)
    $p = New-Object System.Windows.Forms.Form
    $p.Text = $Caption
    $p.FormBorderStyle = 'FixedDialog'
    $p.MaximizeBox = $false; $p.MinimizeBox = $false; $p.ControlBox = $false
    $p.StartPosition = 'CenterParent'
    $p.ClientSize = New-Object System.Drawing.Size(420, 120)
    $p.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Location = New-Object System.Drawing.Point(16, 18)
    $bar.Size = New-Object System.Drawing.Size(388, 20)
    $bar.Minimum = 0; $bar.Maximum = 1000

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(16, 46)
    $lbl.Size = New-Object System.Drawing.Size(388, 20)
    $lbl.ForeColor = [System.Drawing.Color]::DimGray

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $script:T.Cancel
    $btn.Size = New-Object System.Drawing.Size(110, 28)
    $btn.Location = New-Object System.Drawing.Point(294, 76)
    $btn.Add_Click({ $script:CancelRequested = $true })

    $p.Controls.AddRange(@($bar, $lbl, $btn))
    $p.Show($form)
    [System.Windows.Forms.Application]::DoEvents()
    return [pscustomobject]@{ Form = $p; Bar = $bar; Label = $lbl; Button = $btn }
}

function Format-Remaining {
    <# Restzeit menschenlesbar - Sekunden bis knapp eine Minute, darüber Minuten. #>
    param([double]$Seconds)
    if ($Seconds -lt 1)  { return $script:T.RemainSoon }
    if ($Seconds -lt 90) { return ($script:T.RemainSec -f [math]::Ceiling($Seconds)) }
    return ($script:T.RemainMin -f [math]::Ceiling($Seconds / 60))
}

function Invoke-Filter {
    # Riegel gegen Mehrfachauslösung: ohne ihn stapeln sich Tastendrücke, die während eines
    # laufenden Filters in die Warteschlange fallen, und werden anschließend alle abgearbeitet.
    if ($script:Busy) { return }
    $text = $filterBox.Text
    if (-not $text) { Invoke-ClearFilter; return }

    $script:Busy = $true
    $script:CancelRequested = $false
    $total = $s.VirtualRows.Count
    $cols = $grid.Columns.Count
    $dlg = $null
    try {
        if ($total -le $script:ProgressThreshold) {
            $form.Cursor = [System.Windows.Forms.Cursors]::AppStarting
            [void][LogCore]::Filter($s.VirtualRows, $s.FilteredRows, $text, $cols)
        } else {
            $s.FilteredRows.Clear()
            $dlg = New-ProgressDialog -Caption $script:T.FilterProgress
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            $pos = 0
            while ($pos -lt $total) {
                $pos = [LogCore]::FilterRange($s.VirtualRows, $s.FilteredRows, $text, $cols, $pos, $script:ChunkRows)
                $frac = [double]$pos / $total
                $dlg.Bar.Value = [int]([math]::Min(1000, $frac * 1000))
                $rest = if ($frac -gt 0.01) { ($sw.Elapsed.TotalSeconds / $frac) - $sw.Elapsed.TotalSeconds } else { 0 }
                $dlg.Label.Text = ('{0}   -   {1}' -f ($script:T.ProgressRows -f $pos, $total), (Format-Remaining $rest))
                [System.Windows.Forms.Application]::DoEvents()
                if ($script:CancelRequested) { break }
            }
            if ($script:CancelRequested) {
                # Abbruch lässt die Ansicht unverändert - ein halber Filter wäre irreführend.
                $s.FilteredRows.Clear()
                return
            }
        }
        $s.FilterActive = $true
        $s.SearchFrom = 0
        Set-CurrentRows
        Set-GridRowCount $s.FilteredRows.Count
        Update-Status
    } finally {
        if ($dlg) { $dlg.Form.Close(); $dlg.Form.Dispose() }
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $script:Busy = $false
    }
}

function Invoke-ClearFilter {
    $s.FilterActive = $false
    $s.FilteredRows.Clear()
    $filterBox.Text = ''
    $s.SearchFrom = 0
    Set-CurrentRows
    Set-GridRowCount $s.VirtualRows.Count
    Update-Status
}

function Invoke-Search {
    if ($script:Busy) { return }
    $text = $searchBox.Text
    if (-not $text) { return }
    $rows = $script:CurrentRows
    if ($rows.Count -eq 0) { return }

    $script:Busy = $true
    $script:CancelRequested = $false
    $dlg = $null
    try {
        if ($rows.Count -le $script:ProgressThreshold) {
            $idx = [LogCore]::Find($rows, $text, $s.SearchFrom, $grid.Columns.Count)
        } else {
            $dlg = New-ProgressDialog -Caption $script:T.SearchProgress
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            $idx = -1; $done = 0; $total = $rows.Count
            while ($done -lt $total) {
                $take = [math]::Min($script:ChunkRows, $total - $done)
                $idx = [LogCore]::FindRange($rows, $text, ($s.SearchFrom + $done), $take, $grid.Columns.Count)
                if ($idx -ge 0) { break }
                $done += $take
                $frac = [double]$done / $total
                $dlg.Bar.Value = [int]([math]::Min(1000, $frac * 1000))
                $rest = if ($frac -gt 0.01) { ($sw.Elapsed.TotalSeconds / $frac) - $sw.Elapsed.TotalSeconds } else { 0 }
                $dlg.Label.Text = ('{0}   -   {1}' -f ($script:T.ProgressRows -f $done, $total), (Format-Remaining $rest))
                [System.Windows.Forms.Application]::DoEvents()
                if ($script:CancelRequested) { return }
            }
        }
    } finally {
        if ($dlg) { $dlg.Form.Close(); $dlg.Form.Dispose() }
        $script:Busy = $false
    }

    if ($idx -lt 0) {
        [System.Windows.Forms.MessageBox]::Show(($script:T.NotFound -f $text), $script:T.NotFoundTitle, 'OK', 'Information') | Out-Null
        $s.SearchFrom = 0
        return
    }
    $grid.ClearSelection()
    try { $grid.Rows[$idx].Selected = $true } catch {}
    $grid.FirstDisplayedScrollingRowIndex = $idx
    $s.SearchFrom = ($idx + 1) % $rows.Count
}

# ── Detail / Status / Scroll ─────────────────────────────────────────────────
function Update-Detail {
    if ($grid.SelectedRows.Count -eq 0) { $detailBox.Text = ''; return }
    $idx = $grid.SelectedRows[0].Index
    foreach ($r in $grid.SelectedRows) { if ($r.Index -lt $idx) { $idx = $r.Index } }
    $rows = $script:CurrentRows
    if ($idx -lt 0 -or $idx -ge $rows.Count) { return }
    $row = $rows[$idx]
    $colCount = [Math]::Min($row.Length, $grid.Columns.Count)
    $sb = [System.Text.StringBuilder]::new()
    for ($c = 0; $c -lt $colCount; $c++) {
        if ($c -gt 0) { [void]$sb.Append('   |   ') }
        [void]$sb.Append($grid.Columns[$c].HeaderText).Append(': ').Append($row[$c])
    }
    $detailBox.Text = $sb.ToString()
}

function Invoke-AutoScroll {
    if ($autoScrollCheck.Checked -and $grid.RowCount -gt 0) {
        try { $grid.FirstDisplayedScrollingRowIndex = $grid.RowCount - 1 } catch {}
    }
}

function Update-Status {
    $name  = [System.IO.Path]::GetFileName($s.FilePath)
    $extra = switch ($s.Mode) {
        'CSV' { '  |  ' + ($script:T.StatusDelim -f $s.Delimiter) }
        'W3C' { '  |  ' + ($script:T.StatusDelim -f $s.W3cDelimiter) + '  |  ' + ($script:T.StatusColumns -f $s.W3cHeaders.Count) }
        default { '' }
    }
    $rowInfo = if ($s.FilterActive) {
        $script:T.RowsFiltered -f $s.FilteredRows.Count, $s.VirtualRows.Count
    } else {
        $script:T.Rows -f $s.VirtualRows.Count
    }
    $statusLabel.Text = "$name  |  $rowInfo  |  $($script:T.StatusMode -f $s.Mode)$extra  |  $(Get-Date -Format 'HH:mm:ss')"
}

# ── Modus ─────────────────────────────────────────────────────────────────────
function Set-ModeButtons([string]$active) {
    $btnCsv.Checked=$active-eq'CSV'; $btnCmTrace.Checked=$active-eq'CMTrace'
    $btnDateTime.Checked=$active-eq'DateTime'; $btnW3c.Checked=$active-eq'W3C'
    $btnText.Checked=$active-eq'Text'
}
function Switch-To([string]$m) {
    $s.Mode=$m; $s.ManualMode=$true; Set-ModeButtons $m; if ($s.FilePath) { Invoke-FullLoad }
}
function Open-File([string]$fp) {
    if (-not (Test-Path -LiteralPath $fp)) { return }
    $s.FilePath=$fp; $form.Text="$($script:T.WindowTitle)  –  $([System.IO.Path]::GetFileName($fp))  -  $script:BlogUrl"
    if (-not $s.ManualMode) { $s.Mode=Get-AutoMode $fp; Set-ModeButtons $s.Mode }
    $s.LastWriteTime=[datetime]::MinValue; $s.LastLength=-1L
    Invoke-FullLoad; $timer.Start()
}

# ── UI ───────────────────────────────────────────────────────────────────────
$form = New-Object System.Windows.Forms.Form
$form.Text="$($script:T.WindowTitle)  -  $script:BlogUrl"; $form.Size=New-Object System.Drawing.Size(1200,740)
$form.StartPosition='CenterScreen'; $form.MinimumSize=New-Object System.Drawing.Size(800,500)
# Fenstersymbol: aus der eigenen EXE, beim Skriptstart aus LogViewer.ico daneben.
try {
    $selfPath = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    if (([IO.Path]::GetFileNameWithoutExtension($selfPath)) -notin @('powershell', 'pwsh', 'powershell_ise')) {
        $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($selfPath)
    } else {
        $icoPath = Join-Path (Split-Path $PSCommandPath -Parent) 'LogViewer.ico'
        if (Test-Path $icoPath) { $form.Icon = New-Object System.Drawing.Icon($icoPath) }
    }
} catch { }

$tools = New-Object System.Windows.Forms.ToolStrip
$tools.GripStyle='Hidden'; $tools.Padding=New-Object System.Windows.Forms.Padding(4,2,4,2)

$btnOpen=New-Object System.Windows.Forms.ToolStripButton; $btnOpen.Text=$script:T.Open; $btnOpen.DisplayStyle='Text'
[void]$tools.Items.Add($btnOpen); [void]$tools.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
[void]$tools.Items.Add((New-Object System.Windows.Forms.ToolStripLabel $script:T.ModeLabel))

foreach ($def in @(@{Var='btnCsv';L='CSV'},@{Var='btnCmTrace';L='CMTrace'},
                   @{Var='btnDateTime';L=$script:T.ModeDateTime},@{Var='btnW3c';L='W3C'},
                   @{Var='btnText';L='Text'})) {
    $b=New-Object System.Windows.Forms.ToolStripButton; $b.Text=$def.L; $b.DisplayStyle='Text'; $b.CheckOnClick=$false
    [void]$tools.Items.Add($b); Set-Variable -Name $def.Var -Value $b
}
$btnCsv.Checked=$true
[void]$tools.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
$autoScrollCheck=New-Object System.Windows.Forms.CheckBox; $autoScrollCheck.Text=$script:T.AutoScroll
$autoScrollCheck.Checked=$true; $autoScrollCheck.AutoSize=$true
[void]$tools.Items.Add((New-Object System.Windows.Forms.ToolStripControlHost $autoScrollCheck))
[void]$tools.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
$btnAbout=New-Object System.Windows.Forms.ToolStripButton
$btnAbout.Text=$script:T.About; $btnAbout.DisplayStyle='Text'
[void]$tools.Items.Add($btnAbout)

$tools2 = New-Object System.Windows.Forms.ToolStrip
$tools2.GripStyle='Hidden'; $tools2.Padding=New-Object System.Windows.Forms.Padding(4,1,4,1)

[void]$tools2.Items.Add((New-Object System.Windows.Forms.ToolStripLabel $script:T.SearchLabel))
$searchBox=New-Object System.Windows.Forms.ToolStripTextBox; $searchBox.Width=200
[void]$tools2.Items.Add($searchBox)
$btnSearch=New-Object System.Windows.Forms.ToolStripButton; $btnSearch.Text=$script:T.SearchNext; $btnSearch.DisplayStyle='Text'
[void]$tools2.Items.Add($btnSearch)
[void]$tools2.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
[void]$tools2.Items.Add((New-Object System.Windows.Forms.ToolStripLabel $script:T.FilterLabel))
$filterBox=New-Object System.Windows.Forms.ToolStripTextBox; $filterBox.Width=200
[void]$tools2.Items.Add($filterBox)
$btnFilter=New-Object System.Windows.Forms.ToolStripButton; $btnFilter.Text=$script:T.FilterApply; $btnFilter.DisplayStyle='Text'
[void]$tools2.Items.Add($btnFilter)
$btnClearFilter=New-Object System.Windows.Forms.ToolStripButton; $btnClearFilter.Text='✕'; $btnClearFilter.DisplayStyle='Text'; $btnClearFilter.ToolTipText=$script:T.FilterClearTip
[void]$tools2.Items.Add($btnClearFilter)

$grid=New-Object System.Windows.Forms.DataGridView
$grid.Dock='Fill'; $grid.VirtualMode=$true; $grid.ReadOnly=$true
$grid.AllowUserToAddRows=$false; $grid.AllowUserToDeleteRows=$false; $grid.RowHeadersVisible=$false
$grid.AllowUserToResizeRows=$false
$grid.SelectionMode='FullRowSelect'; $grid.MultiSelect=$true
$grid.BackgroundColor=[System.Drawing.SystemColors]::Window; $grid.BorderStyle='None'
$grid.ColumnHeadersHeightSizeMode='AutoSize'; $grid.AutoSizeColumnsMode='None'
$grid.ClipboardCopyMode='EnableWithoutHeaderText'
# Doppeltes Puffern spart beim Scrollen langer Listen sichtbar Zeichenzeit.
try {
    $dbProp = [System.Windows.Forms.DataGridView].GetProperty('DoubleBuffered',
              [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic)
    $dbProp.SetValue($grid, $true, $null)
} catch {}

$detailPanel=New-Object System.Windows.Forms.Panel; $detailPanel.Dock='Bottom'; $detailPanel.Height=72
$detailPanel.BackColor=[System.Drawing.SystemColors]::ControlLight
$detailPanel.MinimumSize=New-Object System.Drawing.Size(0,30)
$detailBox=New-Object System.Windows.Forms.TextBox
$detailBox.Dock='Fill'; $detailBox.ReadOnly=$true; $detailBox.Multiline=$true
$detailBox.WordWrap=$true; $detailBox.ScrollBars='Vertical'; $detailBox.BorderStyle='None'
$detailBox.BackColor=[System.Drawing.SystemColors]::ControlLight
$detailBox.Font=New-Object System.Drawing.Font('Consolas',9)
$detailBox.Margin=New-Object System.Windows.Forms.Padding(4)
$sep=New-Object System.Windows.Forms.Panel; $sep.Dock='Top'; $sep.Height=1
$sep.BackColor=[System.Drawing.Color]::FromArgb(180,180,180)
$detailPanel.Controls.Add($detailBox); $detailPanel.Controls.Add($sep)

$detailSplitter=New-Object System.Windows.Forms.Splitter
$detailSplitter.Dock='Bottom'; $detailSplitter.Height=4
$detailSplitter.BackColor=[System.Drawing.Color]::FromArgb(160,160,160)
$detailSplitter.Cursor=[System.Windows.Forms.Cursors]::HSplit
$detailSplitter.MinSize=30; $detailSplitter.MinExtra=60

$status=New-Object System.Windows.Forms.StatusStrip
$statusLabel=New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text=$script:T.NoFile
$statusLabel.TextAlign='MiddleLeft'; $statusLabel.Spring=$true
[void]$status.Items.Add($statusLabel)
$progressBar=New-Object System.Windows.Forms.Panel
$progressBar.Dock='Bottom'; $progressBar.Height=0
$progressBar.BackColor=[System.Drawing.SystemColors]::ControlLight

$pbChunk=New-Object System.Windows.Forms.Panel
$pbChunk.Size=New-Object System.Drawing.Size(0,0)
$pbChunk.BackColor=[System.Drawing.Color]::SteelBlue
$pbChunk.Top=0; $pbChunk.Left=0
$progressBar.Controls.Add($pbChunk)

$ctx=New-Object System.Windows.Forms.ContextMenuStrip
$ctxCopy=New-Object System.Windows.Forms.ToolStripMenuItem $script:T.CopyRows
$ctxCopy.ShortcutKeys=[System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::C
[void]$ctx.Items.Add($ctxCopy)
$grid.ContextMenuStrip=$ctx

# Layout: zuletzt hinzugefügt = zuerst an der Kante verankert
$form.Controls.Add($grid)
$form.Controls.Add($detailSplitter)
$form.Controls.Add($detailPanel)
$form.Controls.Add($progressBar)
$form.Controls.Add($status)
$form.Controls.Add($tools2)
$form.Controls.Add($tools)

# ── Events ───────────────────────────────────────────────────────────────────
# Läuft pro sichtbarer Zelle - hier zählt jeder eingesparte Zugriff.
$grid.Add_CellValueNeeded({
    $rows = $script:CurrentRows
    $r = $_.RowIndex
    if ($r -ge 0 -and $r -lt $rows.Count) {
        $row = $rows[$r]
        $c = $_.ColumnIndex
        if ($c -ge 0 -and $c -lt $row.Length) { $_.Value = $row[$c] }
    }
})

$grid.Add_CellFormatting({
    # ACHTUNG: switch setzt $_ auf den geprüften Wert. Das Ereignisargument muss deshalb vorher
    # festgehalten werden - sonst zeigt $_ im switch-Zweig auf die Typnummer und der Zugriff auf
    # CellStyle schlägt fehl (genau daran scheiterte die Einfärbung bisher unbemerkt).
    $e = $_
    if ($s.Mode -ne 'CMTrace') { return }
    $r = $e.RowIndex
    if ($r -lt 0) { return }
    $rows = $script:CurrentRows
    if ($r -ge $rows.Count) { return }
    $row = $rows[$r]
    if ($row.Length -lt 7) { return }
    switch ($row[6]) {
        '2' { $e.CellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,248,200) }
        '3' { $e.CellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,180,180) }
    }
})

$grid.Add_SelectionChanged({ Update-Detail })

$ctxCopy.Add_Click({
    $rows = $script:CurrentRows
    $colCount = $grid.Columns.Count
    $sb = [System.Text.StringBuilder]::new()
    $indices = New-Object System.Collections.Generic.List[int]
    foreach ($r in $grid.SelectedRows) { $indices.Add($r.Index) }
    $indices.Sort()
    foreach ($i in $indices) {
        if ($i -lt 0 -or $i -ge $rows.Count) { continue }
        $data = $rows[$i]
        $take = [Math]::Min($colCount, $data.Length)
        [void]$sb.AppendLine(($data[0..($take-1)] -join "`t"))
    }
    $text = $sb.ToString().TrimEnd()
    if ($text) { [System.Windows.Forms.Clipboard]::SetText($text) }
})

$form.AllowDrop=$true
$form.Add_DragEnter({ if ($_.Data.GetDataPresent('FileDrop')) { $_.Effect='Copy' } })
$form.Add_DragDrop({
    $files=$_.Data.GetData('FileDrop')
    if ($files -and $files.Count -gt 0) { Open-File $files[0] }
})

# Länge und Zeitstempel prüfen: manche Schreiber aktualisieren LastWriteTime verzögert.
$timer=New-Object System.Windows.Forms.Timer; $timer.Interval=500
$timer.Add_Tick({
    if (-not $s.FilePath -or $s.Loading) { return }
    $item=Get-Item -LiteralPath $s.FilePath -ErrorAction SilentlyContinue
    if (-not $item) { return }
    if ($item.Length -lt $s.ByteOffset) {
        # Datei wurde abgeschnitten oder rotiert - komplett neu laden.
        $s.LastWriteTime=$item.LastWriteTime; $s.LastLength=$item.Length
        Invoke-FullLoad
        return
    }
    if ($item.LastWriteTime -ne $s.LastWriteTime -or $item.Length -ne $s.LastLength) {
        $s.LastWriteTime=$item.LastWriteTime; $s.LastLength=$item.Length
        Invoke-Append
    }
})

$btnOpen.Add_Click({
    $dlg=New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter=$script:T.FileFilter
    if ($dlg.ShowDialog() -eq 'OK') { Open-File $dlg.FileName }
})

$btnCsv.Add_Click({      Switch-To 'CSV'      })
$btnCmTrace.Add_Click({  Switch-To 'CMTrace'  })
$btnDateTime.Add_Click({ Switch-To 'DateTime' })
$btnW3c.Add_Click({      Switch-To 'W3C'      })
$btnText.Add_Click({     Switch-To 'Text'     })

$btnSearch.Add_Click({ Invoke-Search })
$searchBox.Add_KeyDown({ if ($_.KeyCode -eq 'Return') { Invoke-Search; $_.SuppressKeyPress=$true } })

$btnFilter.Add_Click({ Invoke-Filter })
$filterBox.Add_KeyDown({ if ($_.KeyCode -eq 'Return') { Invoke-Filter; $_.SuppressKeyPress=$true } })
$btnClearFilter.Add_Click({ Invoke-ClearFilter })
$btnAbout.Add_Click({ Show-About })

$form.Add_FormClosed({ $timer.Stop() })

# ── Start ─────────────────────────────────────────────────────────────────────
function Show-About {
    <# Kleines Fenster mit Version und anklickbarer Blog-Adresse. #>
    $a = New-Object System.Windows.Forms.Form
    $a.Text = $script:T.AboutTitle
    $a.FormBorderStyle = 'FixedDialog'
    $a.MaximizeBox = $false; $a.MinimizeBox = $false
    $a.StartPosition = 'CenterParent'
    $a.ClientSize = New-Object System.Drawing.Size(420, 190)
    $a.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    $lblName = New-Object System.Windows.Forms.Label
    $lblName.Text = $script:T.WindowTitle
    $lblName.Font = New-Object System.Drawing.Font('Segoe UI', 14, [System.Drawing.FontStyle]::Bold)
    $lblName.Location = New-Object System.Drawing.Point(20, 18)
    $lblName.Size = New-Object System.Drawing.Size(380, 30)

    $lblVer = New-Object System.Windows.Forms.Label
    $lblVer.Text = $script:T.AboutVersion -f $script:Version
    $lblVer.Location = New-Object System.Drawing.Point(22, 50)
    $lblVer.Size = New-Object System.Drawing.Size(380, 20)
    $lblVer.ForeColor = [System.Drawing.Color]::DimGray

    $lblText = New-Object System.Windows.Forms.Label
    $lblText.Text = $script:T.AboutText -f [Environment]::NewLine
    $lblText.Location = New-Object System.Drawing.Point(22, 78)
    $lblText.Size = New-Object System.Drawing.Size(380, 40)

    $lblBlog = New-Object System.Windows.Forms.Label
    $lblBlog.Text = $script:T.AboutBlog
    $lblBlog.Location = New-Object System.Drawing.Point(22, 126)
    $lblBlog.Size = New-Object System.Drawing.Size(40, 20)

    $link = New-Object System.Windows.Forms.LinkLabel
    $link.Text = $script:BlogUrl
    $link.Location = New-Object System.Drawing.Point(62, 126)
    $link.Size = New-Object System.Drawing.Size(340, 20)
    $link.LinkBehavior = 'HoverUnderline'
    $link.Add_LinkClicked({
        try { Start-Process $script:BlogUrl } catch {
            [System.Windows.Forms.MessageBox]::Show($script:BlogUrl, $script:T.AboutTitle, 'OK', 'Information') | Out-Null
        }
    })

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = $script:T.AboutClose
    $btnClose.Size = New-Object System.Drawing.Size(100, 28)
    $btnClose.Location = New-Object System.Drawing.Point(302, 152)
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $a.Controls.AddRange(@($lblName, $lblVer, $lblText, $lblBlog, $link, $btnClose))
    $a.AcceptButton = $btnClose
    $a.CancelButton = $btnClose
    [void]$a.ShowDialog($form)
    $a.Dispose()
}

function Invoke-FirstRunAssocPrompt {
    <#
      Einmalige Rückfrage beim ersten Start, ob LogViewer für Protokolldateien angeboten werden
      soll. Danach wird nicht mehr gefragt - auch nicht nach einem "Nein", sonst wäre es lästig.
      Nachholen lässt sich das jederzeit mit -Register.
    #>
    try {
        if (Test-Path $script:StateKey) {
            $asked = (Get-ItemProperty -Path $script:StateKey -Name 'AssocPrompt' -ErrorAction SilentlyContinue).AssocPrompt
            if ($asked) { return }
        }
        if (Test-FileTypesRegistered -Ext $Extensions) {
            New-Item -Path $script:StateKey -Force | Out-Null
            Set-ItemProperty -Path $script:StateKey -Name 'AssocPrompt' -Value 'already'
            return
        }
        $answer = [System.Windows.Forms.MessageBox]::Show(
            ($script:T.AssocQuestion -f ($Extensions -join ', ')),
            $script:T.AssocTitle,
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)

        New-Item -Path $script:StateKey -Force | Out-Null
        if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                Register-FileTypes -Ext $Extensions
                Set-ItemProperty -Path $script:StateKey -Name 'AssocPrompt' -Value 'yes'
                [System.Windows.Forms.MessageBox]::Show(
                    ($script:T.AssocDone -f ($Extensions -join ', ')),
                    $script:T.WindowTitle, 'OK', 'Information') | Out-Null
            } catch {
                [System.Windows.Forms.MessageBox]::Show(
                    ($script:T.AssocFailed -f "`n", $_.Exception.Message),
                    $script:T.WindowTitle, 'OK', 'Warning') | Out-Null
            }
        } else {
            Set-ItemProperty -Path $script:StateKey -Name 'AssocPrompt' -Value 'no'
        }
    } catch { }   # Die Rückfrage darf den Start des Betrachters niemals verhindern.
}

$resolved = $null
if ($Path -and (Test-Path -LiteralPath $Path)) { $resolved = (Resolve-Path -LiteralPath $Path).Path }
$form.Add_Shown({
    if ($resolved) {
        if ($Mode -ne 'Auto') { $s.Mode=$Mode; $s.ManualMode=$true; Set-ModeButtons $Mode }
        Open-File $resolved
    }
    # Erst nach dem Laden fragen - der Betrachter soll zuerst zu sehen sein.
    Invoke-FirstRunAssocPrompt
})

if (-not $script:HeadlessOnly) { [System.Windows.Forms.Application]::Run($form) }
