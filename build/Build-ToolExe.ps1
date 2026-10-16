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

    Ohne -Tool werden ALLE Werkzeuge gebaut, die eine Exe ausliefern. Das ist der Weg nach
    einem git pull: ein Aufruf, danach liegt neben jedem Hauptskript die passende EXE. Die
    EXE-Dateien stehen in .gitignore und kommen deshalb nie per git mit.

.EXAMPLE
    .\build\Build-ToolExe.ps1
    Baut alle Werkzeuge mit EXE. Ein Fehler bricht den Lauf nicht ab; am Ende steht eine Bilanz.

.EXAMPLE
    .\build\Build-ToolExe.ps1 -Tool LogViewer
    .\build\Build-ToolExe.ps1 -Tool LogViewer,ExchangeTester
    .\build\Build-ToolExe.ps1 -Tool LogViewer -Version 3.0.2
#>
[CmdletBinding()]
param(
    # Leer gelassen: alle Werkzeuge mit Exe-Eintrag in ihrer release.psd1.
    [string[]]$Tool,
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

# ------------------------------------------------------------------ Sammelbau
# Ohne -Tool ruft sich das Skript je Werkzeug selbst auf. Ein Fehler darf den Lauf nicht
# beenden: wer alles baut, will wissen WAS fehlschlug, nicht beim ersten Stolpern stehen.
if (-not $Tool) {
    $alle = @(Get-ChildItem $Repo -Directory |
              Where-Object { $_.Name -notmatch '^\.|^build$|^dist$' } |
              Where-Object { Test-Path (Join-Path $_.FullName 'release.psd1') } |
              Where-Object { (Import-PowerShellDataFile (Join-Path $_.FullName 'release.psd1')).Exe } |
              Sort-Object Name | Select-Object -ExpandProperty Name)

    if (-not $alle) { Write-Output 'Kein Werkzeug mit Exe-Eintrag gefunden.'; return }

    if (-not (Get-Module ps2exe -ListAvailable)) {
        throw "Das Modul ps2exe fehlt. Einmalig installieren mit: Install-Module ps2exe -Scope CurrentUser"
    }

    Write-Output ("Baue {0} Werkzeuge: {1}`n" -f $alle.Count, ($alle -join ', '))
    $ok = @(); $fehler = @()
    foreach ($t in $alle) {
        Write-Output "---- $t ----"
        try {
            & $MyInvocation.MyCommand.Path -Tool $t -Repo $Repo
            $ok += $t
        }
        catch {
            $fehler += [pscustomobject]@{ Tool = $t; Grund = $_.Exception.Message }
            Write-Output "  FEHL $t : $($_.Exception.Message)"
        }
    }

    Write-Output "`n==================== Bilanz ===================="
    Write-Output ("  gebaut:        {0}" -f ($ok.Count))
    if ($fehler) {
        Write-Output ("  fehlgeschlagen: {0}" -f $fehler.Count)
        $fehler | ForEach-Object { Write-Output ("    {0,-22} {1}" -f $_.Tool, $_.Grund) }
        exit 1
    }
    Write-Output '  alle Werkzeuge gebaut.'
    return
}

# Mehrere Werkzeuge ausdruecklich benannt: derselbe Weg, nur ohne Suchlauf.
if ($Tool.Count -gt 1) {
    if ($Version) { throw '-Version gilt nur fuer ein einzelnes Werkzeug.' }
    foreach ($t in $Tool) {
        Write-Output "---- $t ----"
        & $MyInvocation.MyCommand.Path -Tool $t -Repo $Repo
    }
    return
}
# Eigene Variable statt $Tool = $Tool[0]: die Zuweisung an einen als [string[]] deklarierten
# Parameter macht PowerShell sofort wieder zum Array, und Join-Path scheitert daran.
$Werkzeug = $Tool[0]

$ordner = Join-Path $Repo $Werkzeug
$metaPfad = Join-Path $ordner 'release.psd1'
if (-not (Test-Path $metaPfad)) { throw "Es gibt keine release.psd1 für '$Werkzeug': $metaPfad" }
$meta = Import-PowerShellDataFile $metaPfad

if (-not $meta.Exe) {
    Write-Output "$Werkzeug liefert keine EXE aus (kein Exe-Eintrag in release.psd1) - nichts zu bauen."
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
    title       = $(if ($meta.Title) { $meta.Title } else { $Werkzeug })
    description = $(if ($meta.Description) { $meta.Description } else { $Werkzeug })
    company     = 'azitc'
    product     = $Werkzeug
    version     = $Version
}
if ($meta.RequireAdmin) { $p2e.requireAdmin = $true }

$ico = Join-Path $ordner $(if ($meta.Icon) { $meta.Icon } else { "$Werkzeug.ico" })
if (Test-Path $ico) { $p2e.iconFile = $ico }

Invoke-ps2exe @p2e

if (-not (Test-Path $out)) { throw "Build fehlgeschlagen: $($meta.Exe) wurde nicht erzeugt." }
Write-Output ("Gebaut: {0} ({1} KB)" -f $out, [math]::Round((Get-Item $out).Length / 1KB, 1))
