<#
.SYNOPSIS
    Prüft, dass jedes Werkzeug seine Voraussetzungen erzwingt - und zwar einheitlich.

.DESCRIPTION
    Drei Regeln, jede mit einer Gegenprobe im Hinterkopf: baut man den Fehler zurück,
    muss die Prüfung fehlschlagen.

      1. Jede eingebettete Prolog-Kopie ist zeichengleich mit build/Prolog.*.ps1.
      2. Jedes Werkzeug gegen Exchange Online erzwingt dieselbe Mindestversion -
         egal ob über den Prolog oder über eigene Logik.
      3. Kein Werkzeug ruft Exchange-Cmdlets auf, ohne die Voraussetzung vorher
         zu prüfen (kein nacktes Add-PSSnapin, kein Connect ohne Sessionprüfung).

.EXAMPLE
    .\build\Test-Prolog.ps1
#>
[CmdletBinding()]
param([string]$Repo)
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
function Ok($t)   { $script:pass++; Write-Output "  OK   $t" }
function Bad($t)  { $script:fail++; Write-Output "  FEHL $t" }

$map = Import-PowerShellDataFile (Join-Path $Repo 'build\prolog-map.psd1')
$min = $map.MinimumExoModuleVersion

Write-Output "`n######## 1. Prolog-Kopien stimmen mit der Quelle überein ########"
foreach ($t in $map.Tools) {
    $path = Join-Path $Repo (Join-Path $t.Tool $t.Main)
    if (-not (Test-Path $path)) { Bad "$($t.Tool): Hauptskript fehlt ($($t.Main))"; continue }

    $srcPath = Join-Path $Repo ("build\Prolog.{0}.ps1" -f (Get-Culture).TextInfo.ToTitleCase($t.Prolog))
    $src = ([IO.File]::ReadAllText($srcPath, [Text.Encoding]::UTF8)).TrimEnd("`r", "`n")
    $txt = [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8)

    $open = "# <prolog:$($t.Prolog) "; $close = "# </prolog:$($t.Prolog)>"
    $i = $txt.IndexOf($open); $j = if ($i -ge 0) { $txt.IndexOf($close, $i) } else { -1 }
    if ($i -lt 0 -or $j -lt 0) { Bad "$($t.Tool): Prolog-Marker fehlen"; continue }

    $eingebettet = $txt.Substring($i, ($j + $close.Length) - $i)
    if ($eingebettet -ceq $src) { Ok "$($t.Tool) ($($t.Prolog))" }
    else { Bad "$($t.Tool): Prolog weicht von build/Prolog.$($t.Prolog).ps1 ab - build\Sync-Prolog.ps1 laufen lassen" }
}

Write-Output "`n######## 2. Mindestversion $min wird überall erzwungen ########"
# Der Prolog holt sie aus $script:RequiredModuleVersion, die Eigenbau-Werkzeuge
# aus $script:MinimumExoModuleVersion. Beide Wege sind recht, der Wert muss stimmen.
$exoAlle = @()
$exoAlle += $map.Tools | Where-Object { $_.Prolog -eq 'exo' } | ForEach-Object { [pscustomobject]@{ Tool = $_.Tool; Main = $_.Main; Var = 'RequiredModuleVersion' } }
$exoAlle += $map.ExoOwnLogic | ForEach-Object { [pscustomobject]@{ Tool = $_.Tool; Main = $_.Main; Var = 'MinimumExoModuleVersion' } }

foreach ($e in $exoAlle) {
    $path = Join-Path $Repo (Join-Path $e.Tool $e.Main)
    if (-not (Test-Path $path)) { Bad "$($e.Tool)/$($e.Main): fehlt"; continue }
    $txt = [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8)
    $m = [regex]::Match($txt, ('\$script:{0}\s*=\s*''([0-9]+\.[0-9]+\.[0-9]+)''' -f $e.Var))
    if (-not $m.Success) { Bad "$($e.Tool)/$($e.Main): `$script:$($e.Var) nicht gesetzt"; continue }
    if ($m.Groups[1].Value -ne $min) { Bad "$($e.Tool)/$($e.Main): fordert $($m.Groups[1].Value), erwartet $min"; continue }
    # Gesetzt reicht nicht - der Wert muss auch gegen die gefundene Version geprüft werden.
    # Beide Schreibweisen gelten: "-lt [Version]$x" und "-lt $min" nach vorheriger Umwandlung.
    if ($txt -notmatch '\.Version\s+-lt\s') { Bad "$($e.Tool)/$($e.Main): Version gesetzt, aber nirgends verglichen"; continue }
    Ok "$($e.Tool)/$($e.Main) fordert $min und vergleicht"
}

Write-Output "`n######## 3. Keine Exchange-Aufrufe ohne vorherige Prüfung ########"
$verdaechtig = @()
foreach ($d in (Get-ChildItem $Repo -Directory | Where-Object { $_.Name -notmatch '^\.|^build$' })) {
    foreach ($f in (Get-ChildItem $d.FullName -Filter *.ps1 -Recurse -File |
                    Where-Object { $_.FullName -notmatch '\\lib\\|\\tests\\|\\_source\\' -and $_.Name -notlike 'Build-*' -and $_.Name -ne 'build-exe.ps1' })) {
        $txt = [IO.File]::ReadAllText($f.FullName, [Text.Encoding]::UTF8)

        # Nacktes Add-PSSnapin ohne Registered-Prüfung oder Prolog
        if ($txt -match 'Add-PSSnapin' -and $txt -notmatch 'prolog:onprem' -and $txt -notmatch 'Get-PSSnapin\s+-Registered') {
            $verdaechtig += [pscustomobject]@{ Datei = $f.Name; Grund = 'Add-PSSnapin ohne Prüfung und ohne Prolog' }
        }
        # Connect-ExchangeOnline ohne vorherige Sessionprüfung oder Prolog
        if ($txt -match 'Connect-ExchangeOnline' -and $txt -notmatch 'prolog:exo' -and $txt -notmatch 'Get-ConnectionInformation') {
            $verdaechtig += [pscustomobject]@{ Datei = $f.Name; Grund = 'Connect-ExchangeOnline ohne Sessionprüfung' }
        }
    }
}
$gaps = @($map.KnownGaps)
$gapNamen = @($gaps.Datei)

# a) Neu aufgetretene Faelle sind ein Fehler - so wächst die Schuld nicht unbemerkt.
$neu = @($verdaechtig | Where-Object { $_.Datei -notin $gapNamen })
if ($neu) { $neu | ForEach-Object { Bad "$($_.Datei): $($_.Grund) - nicht in KnownGaps" } }
else { Ok "kein neues Werkzeug ruft Exchange ungeprüft auf ($($gaps.Count) bekannte Altfälle)" }

# b) Ein Eintrag, dessen Problem behoben ist, muss raus - sonst verrottet die Liste
#    und behauptet Schuld, die es nicht mehr gibt.
$erledigt = @($gapNamen | Where-Object { $_ -notin @($verdaechtig.Datei) })
if ($erledigt) { $erledigt | ForEach-Object { Bad "KnownGaps nennt '$_', aber dort ist nichts mehr offen - Eintrag entfernen" } }
else { Ok 'KnownGaps enthält keine erledigten Einträge' }

if ($gaps) {
    Write-Output "`n  Noch offen (aus build/prolog-map.psd1):"
    $gaps | ForEach-Object { Write-Output ("    {0,-38} {1}" -f $_.Datei, $_.Grund) }
}

Write-Output "`n================ Bestanden: $pass   Fehlgeschlagen: $fail ================"
if ($fail -gt 0) { exit 1 }
