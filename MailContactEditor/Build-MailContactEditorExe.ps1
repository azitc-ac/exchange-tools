<#
.SYNOPSIS
    Baut MailContactEditor.exe aus Edit-MailContactAddresses.ps1 mit ps2exe.

.DESCRIPTION
    Nur zur Bauzeit nötig: ps2exe (Install-Module ps2exe -Scope CurrentUser).

    Ohne Konsolenfenster (-noConsole) und im STA-Modus, weil WinForms das verlangt.
    Keine Administratorrechte: das Werkzeug arbeitet gegen Exchange Online, nicht
    gegen den lokalen Rechner.

    Die fertige EXE enthält das Skript, aber nicht das Modul ExchangeOnlineManagement -
    das muss auf dem Zielrechner installiert sein (V3, wegen Get-ConnectionInformation).

.EXAMPLE
    .\Build-MailContactEditorExe.ps1
    .\Build-MailContactEditorExe.ps1 -Version 1.1.0
#>
[CmdletBinding()]
param(
    [string]$Root = $PSScriptRoot,
    [string]$Version = ''   # leer = aus Edit-MailContactAddresses.ps1 übernehmen
)
$ErrorActionPreference = 'Stop'
Import-Module ps2exe -ErrorAction Stop

if (-not $Root) {
    $Root = if ($PSScriptRoot) { $PSScriptRoot }
            elseif ($MyInvocation.MyCommand.Path) { Split-Path $MyInvocation.MyCommand.Path -Parent }
            else { (Get-Location).Path }
}

$in  = Join-Path $Root 'Edit-MailContactAddresses.ps1'
$out = Join-Path $Root 'MailContactEditor.exe'
$ico = Join-Path $Root 'MailContactEditor.ico'
if (-not (Test-Path $in)) { throw "Nicht gefunden: $in" }

# Version genau einmal pflegen: sie steht im Hauptskript und wird hier ausgelesen.
. (Join-Path $Root '..\build\Get-ToolVersion.ps1')
if (-not $Version) {
    $Version = Get-ToolVersion -Path $in
    Write-Host "Version aus Edit-MailContactAddresses.ps1: $Version" -ForegroundColor DarkGray
} else {
    # Explizit übergebene Version muss zum Skript passen, sonst lügt die Dateiversion.
    $null = Assert-ToolVersion -Path $in -Expected $Version
}

# Eine laufende EXE lässt sich nicht überschreiben - vorher beenden.
Get-Process -Name 'MailContactEditor' -ErrorAction SilentlyContinue | ForEach-Object {
    Write-Host "Beende laufende Instanz (PID $($_.Id)) ..." -ForegroundColor Yellow
    $_.Kill(); $_.WaitForExit(5000)
}

$p2e = @{
    inputFile   = $in
    outputFile  = $out
    noConsole   = $true
    STA         = $true
    title       = 'Mail Contact Editor'
    description = 'E-Mail-Kontakte in Exchange Online bearbeiten'
    company     = 'azitc'
    product     = 'MailContactEditor'
    version     = $Version
    copyright   = ''
}
if (Test-Path $ico) { $p2e.iconFile = $ico }

Invoke-ps2exe @p2e

if (-not (Test-Path $out)) { throw 'Build fehlgeschlagen: MailContactEditor.exe wurde nicht erzeugt.' }
Write-Host ("Gebaut: {0} ({1} KB)" -f $out, [math]::Round((Get-Item $out).Length / 1KB, 1)) -ForegroundColor Green
Write-Host 'Aufruf: MailContactEditor.exe [-UserPrincipalName admin@example.com]'
