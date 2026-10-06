<#
.SYNOPSIS
    Führt alle Prüfungen des Repos aus - lokal und in der CI derselbe Befehl.

.DESCRIPTION
    Damit "bei mir lief es durch" und "die CI ist grün" dasselbe bedeuten, gibt es
    genau einen Einstiegspunkt. Der Workflow .github/workflows/ci.yml ruft dieses
    Skript auf, sonst nichts.

    Test-ReleaseLogic baut dabei die EXE-Dateien - dafür muss ps2exe da sein.
    Mit -SkipBuild bleibt diese Prüfung aus, etwa für einen schnellen Durchlauf
    auf einem Rechner ohne das Modul.

.EXAMPLE
    .\build\Invoke-AllTests.ps1
    .\build\Invoke-AllTests.ps1 -SkipBuild
#>
[CmdletBinding()]
param(
    [string]$Repo,
    [switch]$SkipBuild
)
$ErrorActionPreference = 'Stop'

# $PSScriptRoot kommt beim Start ueber -File in manchen Shells leer an; dann den
# eigenen Pfad anders ermitteln, sonst scheitert schon die Parameterbindung.
if (-not $Repo) {
    $hier = if ($PSScriptRoot) { $PSScriptRoot }
            elseif ($MyInvocation.MyCommand.Path) { Split-Path $MyInvocation.MyCommand.Path -Parent }
            else { (Get-Location).Path }
    $Repo = Split-Path $hier -Parent
}
Set-Location $Repo

$pruefungen = @(
    @{ Name = 'Syntax';        Skript = 'build\Test-Syntax.ps1';       Build = $false }
    @{ Name = 'Encoding';      Skript = 'build\Test-Encoding.ps1';     Build = $false }
    @{ Name = 'Prolog';        Skript = 'build\Test-Prolog.ps1';       Build = $false }
    @{ Name = 'Release-Logik'; Skript = 'build\Test-ReleaseLogic.ps1'; Build = $true  }
    # Startet die gebauten EXE-Dateien wirklich. Überspringt sich in der CI selbst,
    # weil ein Runner keine brauchbare Desktop-Sitzung hat.
    @{ Name = 'EXE-Start';     Skript = 'build\Test-ExeSmoke.ps1';     Build = $true  }
)

$ergebnis = @()
foreach ($p in $pruefungen) {
    if ($p.Build -and $SkipBuild) {
        Write-Output "`n=== $($p.Name): uebersprungen (-SkipBuild) ==="
        $ergebnis += [pscustomobject]@{ Pruefung = $p.Name; Status = 'uebersprungen' }
        continue
    }

    Write-Output "`n==================== $($p.Name) ===================="
    $pfad = Join-Path $Repo $p.Skript
    if (-not (Test-Path $pfad)) { throw "Prüfskript fehlt: $pfad" }

    # Jede Prüfung in einem eigenen Prozess: ihr "exit 1" soll nicht diese Sitzung
    # beenden, und gesetzte Variablen sollen die nächste Prüfung nicht beeinflussen.
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $pfad
    $code = $LASTEXITCODE
    $ergebnis += [pscustomobject]@{
        Pruefung = $p.Name
        Status   = $(if ($code -eq 0) { 'bestanden' } else { "FEHLGESCHLAGEN ($code)" })
    }
}

Write-Output "`n==================== Ergebnis ===================="
foreach ($e in $ergebnis) { Write-Output ("  {0,-16} {1}" -f $e.Pruefung, $e.Status) }

if (@($ergebnis | Where-Object { $_.Status -like 'FEHLGESCHLAGEN*' }).Count -gt 0) {
    Write-Output "`nMindestens eine Prüfung ist fehlgeschlagen."
    exit 1
}
Write-Output "`nAlles bestanden."
