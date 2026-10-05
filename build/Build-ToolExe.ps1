<#
.SYNOPSIS
    Baut die EXE eines Werkzeugs aus seinem Hauptskript - für alle Werkzeuge derselbe Weg.

.DESCRIPTION
    Vorher hatte jedes Werkzeug mit EXE sein eigenes, fast gleiches Build-Skript, und
    die übrigen hatten gar keins. Dieses Skript ist der gemeinsame Weg; was ein Werkzeug
    unterscheidet, steht in seiner release.psd1:

        Exe          Dateiname der EXE (fehlt er, wird keine gebaut)
        Title        Fenster-/Dateititel
        Description  Beschreibung in den Dateieigenschaften
        RequireAdmin $true, wenn die EXE erhöhte Rechte anfordern soll
        Icon         optionale .ico neben dem Hauptskript

    Die Version kommt immer aus $script:Version im Hauptskript - eine Quelle, kein
    zweiter Ort, an dem sie driften könnte.

.EXAMPLE
    .\build\Build-ToolExe.ps1 -Tool LogViewer
    .\build\Build-ToolExe.ps1 -Tool LogViewer -Version 3.0.2
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Tool,
    [string]$Repo,
    [string]$Version
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

$ordner = Join-Path $Repo $Tool
$metaPfad = Join-Path $ordner 'release.psd1'
if (-not (Test-Path $metaPfad)) { throw "Es gibt keine release.psd1 für '$Tool': $metaPfad" }
$meta = Import-PowerShellDataFile $metaPfad

if (-not $meta.Exe) {
    Write-Output "$Tool liefert keine EXE aus (kein Exe-Eintrag in release.psd1) - nichts zu bauen."
    return
}

Import-Module ps2exe -ErrorAction Stop
. (Join-Path $Repo 'build\Get-ToolVersion.ps1')

$in  = Join-Path $ordner $meta.Main
$out = Join-Path $ordner $meta.Exe
if (-not (Test-Path $in)) { throw "Hauptskript nicht gefunden: $in" }

if (-not $Version) {
    $Version = Get-ToolVersion -Path $in
    Write-Output "Version aus $($meta.Main): $Version"
}
else {
    # Eine ausdrücklich übergebene Version muss zum Skript passen, sonst lügt die Dateiversion.
    $null = Assert-ToolVersion -Path $in -Expected $Version
}

# Eine laufende EXE lässt sich nicht überschreiben - vorher beenden.
$prozess = [IO.Path]::GetFileNameWithoutExtension($meta.Exe)
Get-Process -Name $prozess -ErrorAction SilentlyContinue | ForEach-Object {
    Write-Output "Beende laufende Instanz (PID $($_.Id)) ..."
    $_.Kill(); $_.WaitForExit(5000)
}

$p2e = @{
    inputFile   = $in
    outputFile  = $out
    noConsole   = $true     # alle EXE-Werkzeuge hier haben eine Oberfläche
    STA         = $true     # WinForms und WPF verlangen das
    title       = $(if ($meta.Title) { $meta.Title } else { $Tool })
    description = $(if ($meta.Description) { $meta.Description } else { $Tool })
    company     = 'azitc'
    product     = $Tool
    version     = $Version
}
if ($meta.RequireAdmin) { $p2e.requireAdmin = $true }

$ico = Join-Path $ordner $(if ($meta.Icon) { $meta.Icon } else { "$Tool.ico" })
if (Test-Path $ico) { $p2e.iconFile = $ico }

Invoke-ps2exe @p2e

if (-not (Test-Path $out)) { throw "Build fehlgeschlagen: $($meta.Exe) wurde nicht erzeugt." }
Write-Output ("Gebaut: {0} ({1} KB)" -f $out, [math]::Round((Get-Item $out).Length / 1KB, 1))
