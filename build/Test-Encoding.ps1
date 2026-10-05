<#
.SYNOPSIS
    Prüft, dass jedes Skript mit Umlauten ein UTF-8-BOM trägt.

.DESCRIPTION
    Windows PowerShell 5.1 liest eine .ps1 ohne BOM in der ANSI-Codepage. Aus "ä"
    wird dann "Ã¤" - in Dialogtexten, Meldungen und Protokollzeilen. Das fällt beim
    Schreiben nicht auf, sondern erst dem Anwender.

    Geprüft wird nur, was Umlaute enthält: eine reine ASCII-Datei braucht kein BOM.
    Markdown und YAML bleiben außen vor - GitHub erwartet dort UTF-8 ohne BOM.
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

$fail = 0; $geprueft = 0
$dateien = Get-ChildItem $Repo -Include *.ps1, *.psd1, *.psm1 -Recurse -File |
           Where-Object { $_.FullName -notmatch '\\lib\\Posh-ACME\\|\\\.git\\' }

foreach ($f in $dateien) {
    $bytes = [IO.File]::ReadAllBytes($f.FullName)
    if ($bytes.Length -lt 3) { continue }
    $hatBom = ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)

    # Als UTF-8 lesen und sehen, ob überhaupt Umlaute vorkommen.
    $text = [Text.Encoding]::UTF8.GetString($bytes)
    if ($text -notmatch '[äöüÄÖÜß]') { continue }
    $geprueft++

    if (-not $hatBom) {
        $fail++
        $rel = $f.FullName.Substring($Repo.Length).TrimStart('\')
        Write-Output ("  FEHL {0}: Umlaute ohne UTF-8-BOM - PowerShell 5.1 zeigt Mojibake" -f $rel)
    }
}

# Eine leere Datei ist fast immer ein Unfall beim Schreiben, kein Vorsatz.
$leer = @($dateien | Where-Object { $_.Length -le 3 })
foreach ($l in $leer) {
    $fail++
    Write-Output ("  FEHL {0}: Datei ist leer" -f $l.FullName.Substring($Repo.Length).TrimStart('\'))
}

# Dasselbe Problem eine Ebene höher: GitHub Actions schreibt den Inhalt eines
# "run: |"-Blocks als .ps1 OHNE BOM auf den Runner. Windows PowerShell 5.1 liest
# die dann in der ANSI-Codepage - aus "ausführen" wird "ausfÃ¼hren", sichtbar in
# Release-Notes und Fehlermeldungen. Genau so ist das erste Release erschienen.
# Umlaute in YAML-Kommentaren und name:-Feldern sind unkritisch (die liest GitHub
# selbst als UTF-8); geprüft wird nur, was PowerShell ausführt.
$wfOrdner = Join-Path $Repo '.github\workflows'
if (Test-Path $wfOrdner) {
    foreach ($w in (Get-ChildItem $wfOrdner -Include *.yml, *.yaml -File -Recurse)) {
        $inRun = $false; $runEinzug = 0; $n = 0
        foreach ($z in [IO.File]::ReadAllLines($w.FullName, [Text.Encoding]::UTF8)) {
            $n++
            if ($z -match '^\s*run:\s*\|') {
                $inRun = $true; $runEinzug = ($z -replace '\S.*$', '').Length; continue
            }
            if (-not $inRun) { continue }
            $einzug = ($z -replace '\S.*$', '').Length
            if ($z.Trim() -and $einzug -le $runEinzug) { $inRun = $false; continue }
            if ($z -match '[äöüÄÖÜß]') {
                $fail++
                Write-Output ("  FEHL {0}:{1}: Umlaut im run-Block - Actions legt ihn ohne BOM ab, PS 5.1 zeigt Mojibake" -f $w.Name, $n)
            }
        }
    }
}

Write-Output ("  {0} Skripte mit Umlauten geprüft, {1} Beanstandungen" -f $geprueft, $fail)
if ($fail -gt 0) { exit 1 }
