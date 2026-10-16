<#
.SYNOPSIS
    Prüft, dass die READMEs keine Dateien nennen, die es nicht (mehr) gibt.

.DESCRIPTION
    Entstanden aus einem Fehler: Beim Zusammenführen der Einzel-Build-Skripte auf
    build\Build-ToolExe.ps1 wurden die alten gelöscht, die READMEs aber nicht nachgezogen.
    LogViewer\README.md nannte weiter

        .\Build-LogViewerExe.ps1

    und AcmeExchange\README.md ein `build-exe.ps1`. Beide gab es nicht mehr. Aufgefallen ist
    es erst einem Nutzer, der nach einem git pull danach suchte.

    Geprüft wird jede .ps1-Erwähnung in den READMEs. Gesucht wird an drei Stellen, weil
    READMEs aus verschiedenen Blickwinkeln schreiben: im Werkzeugordner selbst, im
    Wurzelverzeichnis des Repositorys und - als letzter Ausweg - irgendwo im Repository.

    Die zweite Regel ist die Umkehrung: Eine EXE darf ein README nur dann als Datei im
    Repository darstellen, wenn sie auch dort liegt. Sie liegt nie dort (.gitignore: *.exe),
    also muss an jeder Stelle, die das Bauen beschreibt, auch stehen, woher sie sonst kommt.

.PARAMETER Rueckbau
    Baut den Fehler zurück (prüft gegen ein README mit einem erfundenen Skriptnamen) und
    erwartet, dass die Prüfung FEHLSCHLÄGT. Ohne diese Gegenprobe wäre der Test wertlos.

.EXAMPLE
    .\build\Test-Readme.ps1
    .\build\Test-Readme.ps1 -Rueckbau
#>
[CmdletBinding()]
param([string]$Repo, [switch]$Rueckbau)

# $PSScriptRoot kommt beim Start ueber -File in manchen Shells leer an; dann den
# eigenen Pfad anders ermitteln, sonst scheitert schon die Parameterbindung.
if (-not $Repo) {
    $hier = if ($PSScriptRoot) { $PSScriptRoot }
            elseif ($MyInvocation.MyCommand.Path) { Split-Path $MyInvocation.MyCommand.Path -Parent }
            else { (Get-Location).Path }
    $Repo = Split-Path $hier -Parent
}
$ErrorActionPreference = 'Stop'

$pass = 0; $fail = 0
function Ok($t)  { $script:pass++; Write-Output "  OK   $t" }
function Bad($t) { $script:fail++; Write-Output "  FEHL $t" }

# Verweise auf FREMDE Skripte, die es hier nicht geben soll. Diese Liste darf nur kuerzer
# werden: Prüfung 1b schlaegt fehl, wenn ein Eintrag ueberfluessig geworden ist, aber
# stehenblieb - sonst verrottet sie und deckt irgendwann einen echten Fehler zu.
$Fremd = @(
    @{ Datei = 'ConfigureExchangeHybridApplication.ps1'; Grund = 'Skript von Microsoft, wird im Pruefpunkt HYB-02 nur genannt' }
    @{ Datei = 'Populate-TestMailboxByDate.ps1';         Grund = 'Vorlage von Joe Palarchio, Herkunftsangabe des TestMailboxFiller' }
)

# Alle .ps1 im Repository einmal einsammeln - danach nur noch Mengenvergleich.
$alleSkripte = @(Get-ChildItem $Repo -Filter *.ps1 -Recurse -File |
                 Where-Object { $_.FullName -notmatch '\\\.git\\' })
$alleNamen = @($alleSkripte.Name)

$readmes = @(Get-ChildItem $Repo -Filter 'README.md' -Recurse -File |
             Where-Object { $_.FullName -notmatch '\\\.git\\|\\_source\\' })

if ($Rueckbau) {
    # Eine zusaetzliche README mit einem erfundenen Skript - die Pruefung MUSS anschlagen.
    $rbOrdner = Join-Path $env:TEMP ("readme-rueckbau-{0}" -f $PID)
    New-Item -ItemType Directory -Path $rbOrdner -Force | Out-Null
    $rbDatei = Join-Path $rbOrdner 'README.md'
    [IO.File]::WriteAllText($rbDatei, "## Selbst bauen`n`n``````powershell`n.\Build-GibtsNichtExe.ps1`n``````n",
                            (New-Object Text.UTF8Encoding $true))
    $readmes += Get-Item $rbDatei
    Write-Output '  (Rückbau: eine README mit erfundenem Skript - die Prüfung MUSS fehlschlagen)'
}

Write-Output "`n######## 1. Jede genannte .ps1 existiert ########"
foreach ($r in $readmes) {
    $kurz = $r.FullName.Replace("$Repo\", '')
    $txt  = [IO.File]::ReadAllText($r.FullName, [Text.Encoding]::UTF8)

    # Drei Schreibweisen, in dieser Reihenfolge: in doppelten Anfuehrungszeichen, in einfachen,
    # und nackt. Die Anfuehrungszeichen muessen zuerst drankommen, weil mehrere Werkzeuge
    # LEERZEICHEN im Dateinamen haben ("Exchange Queue Viewer replacement.ps1") - ein Muster
    # ohne Anfuehrungszeichen schneidet dort auf das letzte Wort ab und meldet es als fehlend.
    $muster = '"(?<q>[^"\r\n]+\.ps1)"' +
              "|'(?<q>[^'\r\n]+\.ps1)'" +
              '|(?<q>(?<![\w.\-])(?:\.[\\/])?(?:[\w.\-]+[\\/])*[\w.\-]+\.ps1)'
    $genannt = @([regex]::Matches($txt, $muster) | ForEach-Object { $_.Groups['q'].Value } | Sort-Object -Unique)
    if (-not $genannt) { continue }

    $fehlend = @()
    foreach ($g in $genannt) {
        $rel  = $g -replace '^\.[\\/]', '' -replace '/', '\'
        $name = Split-Path $rel -Leaf
        if ($name -in $Fremd.Datei) { continue }

        # Drei Blickwinkel, in dieser Reihenfolge: relativ zum Werkzeugordner, relativ zur
        # Wurzel, und zuletzt nur der Dateiname irgendwo im Repository.
        $gefunden = (Test-Path -LiteralPath (Join-Path $r.DirectoryName $rel)) -or
                    (Test-Path -LiteralPath (Join-Path $Repo $rel)) -or
                    ($name -in $alleNamen)

        if (-not $gefunden) { $fehlend += $g }
    }

    if ($fehlend) { $fehlend | ForEach-Object { Bad "${kurz}: nennt '$_' - gibt es im Repository nicht" } }
    else { Ok "$kurz ($($genannt.Count) Skriptnennungen, alle vorhanden)" }
}

Write-Output "`n######## 1b. Die Fremd-Liste enthält nichts Überflüssiges ########"
# Zwei Wege, auf denen ein Eintrag faul wird: das Skript ist inzwischen doch im Repository
# (dann gehoert es geprueft), oder keine README nennt es mehr (dann ist der Eintrag Ballast).
$alleReadmeTexte = ($readmes | Where-Object { $_.FullName -notmatch 'readme-rueckbau' } |
                    ForEach-Object { [IO.File]::ReadAllText($_.FullName, [Text.Encoding]::UTF8) }) -join "`n"
foreach ($f in $Fremd) {
    if ($f.Datei -in $alleNamen) { Bad "Fremd-Liste nennt '$($f.Datei)', aber die Datei liegt jetzt im Repository - Eintrag entfernen" }
    elseif ($alleReadmeTexte -notmatch [regex]::Escape($f.Datei)) { Bad "Fremd-Liste nennt '$($f.Datei)', aber keine README erwähnt sie noch - Eintrag entfernen" }
    else { Ok "$($f.Datei) - $($f.Grund)" }
}

Write-Output "`n######## 2. Wer das Bauen beschreibt, sagt auch woher die EXE sonst kommt ########"
# Die EXE liegt nie im Repository (.gitignore: *.exe). Ein README, das nur erklaert wie man
# baut, laesst den Leser nach einem git pull vergeblich nach der Datei suchen - genau so ist
# dieser Test entstanden.
foreach ($r in ($readmes | Where-Object { $_.FullName -notmatch 'readme-rueckbau' })) {
    $kurz = $r.FullName.Replace("$Repo\", '')
    $txt  = [IO.File]::ReadAllText($r.FullName, [Text.Encoding]::UTF8)

    # Betrifft nur Werkzeug-READMEs, deren release.psd1 eine Exe nennt.
    $psd = Join-Path $r.DirectoryName 'release.psd1'
    if (-not (Test-Path $psd)) { continue }
    if (-not (Import-PowerShellDataFile $psd).Exe) { continue }

    if ($txt -notmatch 'Build-ToolExe|ps2exe|[Ss]elbst bauen|[Rr]ebuild') { continue }   # beschreibt kein Bauen

    if ($txt -match '[Rr]eleases?\b|releases/') { Ok "$kurz verweist auf den Release" }
    else { Bad "$kurz beschreibt das Bauen, nennt aber den Release nicht - die EXE ist nicht im Repository" }
}

if ($Rueckbau) { Remove-Item (Split-Path $rbDatei -Parent) -Recurse -Force -ErrorAction SilentlyContinue }

Write-Output "`n================ Bestanden: $pass   Fehlgeschlagen: $fail ================"
if ($Rueckbau) {
    if ($fail -gt 0) { Write-Output '  Gegenprobe erfolgreich: mit dem zurueckgebauten Fehler schlaegt die Pruefung fehl.'; exit 0 }
    Write-Output '  Gegenprobe GESCHEITERT: die Pruefung ist auch mit dem Fehler gruen - sie prueft nichts.'
    exit 1
}
if ($fail -gt 0) { exit 1 }
