<#
.SYNOPSIS
    Baut LogViewer.exe aus LogViewer.ps1 mit ps2exe.

.DESCRIPTION
    Nur zur Bauzeit nötig: ps2exe (Install-Module ps2exe -Scope CurrentUser).
    Die fertige EXE ist eigenständig - sie enthält das Skript samt .NET-Kern und braucht
    weder LogViewer.ps1 noch eine Ausführungsrichtlinie.

    Ohne Konsolenfenster (-noConsole) und im STA-Modus, weil WinForms das verlangt.
    Keine Administratorrechte: ein Log-Betrachter braucht sie nicht.

.EXAMPLE
    .\Build-LogViewerExe.ps1
    .\Build-LogViewerExe.ps1 -Version 2.1.0
#>
[CmdletBinding()]
param(
    [string]$Root = $PSScriptRoot,
    [string]$Version = ''   # leer = aus LogViewer.ps1 uebernehmen
)
$ErrorActionPreference = 'Stop'
Import-Module ps2exe -ErrorAction Stop

if (-not $Root) {
    $Root = if ($PSScriptRoot) { $PSScriptRoot }
            elseif ($MyInvocation.MyCommand.Path) { Split-Path $MyInvocation.MyCommand.Path -Parent }
            else { (Get-Location).Path }
}

$in  = Join-Path $Root 'LogViewer.ps1'
$out = Join-Path $Root 'LogViewer.exe'
$ico = Join-Path $Root 'LogViewer.ico'
if (-not (Test-Path $in)) { throw "Nicht gefunden: $in" }

# Version genau einmal pflegen: sie steht in LogViewer.ps1 und wird hier ausgelesen.
. (Join-Path $Root '..\build\Get-ToolVersion.ps1')
if (-not $Version) {
    $Version = Get-ToolVersion -Path $in
    Write-Host "Version aus LogViewer.ps1: $Version" -ForegroundColor DarkGray
} else {
    # Explizit übergebene Version muss zum Skript passen, sonst lügt die Dateiversion.
    $null = Assert-ToolVersion -Path $in -Expected $Version
}

# Eine laufende EXE lässt sich nicht überschreiben - vorher beenden.
Get-Process -Name 'LogViewer' -ErrorAction SilentlyContinue | ForEach-Object {
    Write-Host "Beende laufende Instanz (PID $($_.Id)) ..." -ForegroundColor Yellow
    $_.Kill(); $_.WaitForExit(5000)
}

$p2e = @{
    inputFile   = $in
    outputFile  = $out
    noConsole   = $true
    STA         = $true
    title       = 'Log Viewer'
    description = 'Log Viewer für CSV, CMTrace, W3C und Textprotokolle'
    company     = 'azitc'
    product     = 'LogViewer'
    version     = $Version
    copyright   = ''
}
if (Test-Path $ico) { $p2e.iconFile = $ico }

Invoke-ps2exe @p2e

if (-not (Test-Path $out)) { throw 'Build fehlgeschlagen: LogViewer.exe wurde nicht erzeugt.' }
Write-Host ("Gebaut: {0} ({1} KB)" -f $out, [math]::Round((Get-Item $out).Length / 1KB, 1)) -ForegroundColor Green
Write-Host 'Aufruf: LogViewer.exe -Path "C:\pfad\zur\datei.log" [-Mode CMTrace]'
