<#
.SYNOPSIS
    Prüft, dass jedes PowerShell-Skript im Repo fehlerfrei geparst wird.

.DESCRIPTION
    Ein Skript mit Syntaxfehler fällt sonst erst dem auf, der es herunterlädt und
    startet. Der Parser findet das in Sekunden, ohne irgendetwas auszuführen.

    Die mitgelieferte Fremdbibliothek (lib\Posh-ACME) bleibt außen vor - sie ist
    nicht unser Code und wird unverändert durchgereicht.

    Ausgabe über Write-Output statt Write-Host: Write-Host schreibt in Windows
    PowerShell 5.1 an den Host und fehlt dann in Protokollen und Umleitungen -
    also genau dort, wo man in der CI nachsieht.
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

$fail = 0
$dateien = Get-ChildItem $Repo -Filter *.ps1 -Recurse -File |
           Where-Object { $_.FullName -notmatch '\\lib\\Posh-ACME\\|\\\.git\\' }

foreach ($f in $dateien) {
    $err = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$null, [ref]$err)
    $rel = $f.FullName.Substring($Repo.Length).TrimStart('\')
    if ($err.Count) {
        $fail++
        Write-Output ("  FEHL {0}" -f $rel)
        $err | Select-Object -First 3 | ForEach-Object {
            Write-Output ("         Zeile {0}: {1}" -f $_.Extent.StartLineNumber, $_.Message)
        }
    }
}

Write-Output ("  {0} Skripte geparst, {1} mit Fehlern" -f $dateien.Count, $fail)
if ($fail -gt 0) { exit 1 }
