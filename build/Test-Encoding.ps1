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

Write-Output ("  {0} Skripte mit Umlauten geprüft, {1} Beanstandungen" -f $geprueft, $fail)
if ($fail -gt 0) { exit 1 }
