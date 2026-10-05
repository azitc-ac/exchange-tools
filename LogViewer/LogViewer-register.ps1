<#
.SYNOPSIS
    Trägt "Mit LogViewer öffnen" für .log-, .csv- und .txt-Dateien ein - oder entfernt es wieder.

.DESCRIPTION
    Dünner Aufruf des LogViewers selbst: die Registrierungslogik steht ausschließlich dort
    (-Register / -Unregister), damit es keine zweite Fassung davon gibt, die auseinanderlaufen
    kann. Bevorzugt wird die EXE verwendet, sonst das Skript.

    Alles liegt unter HKCU - keine Administratorrechte nötig.

    Dasselbe fragt der LogViewer beim ersten Start von sich aus; dieses Skript ist der Weg für
    Ausrollung, Nachholen oder Entfernen.

.PARAMETER Remove
    Entfernt die Einträge.
.PARAMETER Extensions
    Dateiendungen (Standard: .log, .csv, .txt).

.EXAMPLE
    .\LogViewer-register.ps1
    .\LogViewer-register.ps1 -Remove
    .\LogViewer-register.ps1 -Extensions .log,.trace
#>
[CmdletBinding()]
param(
    [switch]$Remove,
    [string[]]$Extensions = @('.log', '.csv', '.txt')
)
$ErrorActionPreference = 'Stop'

$root = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path $MyInvocation.MyCommand.Path -Parent }
$exe = Join-Path $root 'LogViewer.exe'
$ps1 = Join-Path $root 'LogViewer.ps1'

$action = if ($Remove) { '-Unregister' } else { '-Register' }

if (-not (Test-Path $ps1)) { throw "LogViewer.ps1 nicht gefunden in $root" }

# Ausgeführt wird immer das Skript - eine mit ps2exe erzeugte EXE beendet sich bei dieser
# Aufgabe nicht zuverlässig. Eingetragen wird trotzdem die EXE, sofern sie vorhanden ist.
$target = if (Test-Path $exe) { $exe } else { $ps1 }
# Bei -File zerlegt PowerShell ein Array in einzelne Argumente; das zweite landete dadurch beim
# Parameter Mode. Deshalb als eine kommaseparierte Zeichenkette übergeben.
$extArg = $Extensions -join ','
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $ps1 $action -Extensions $extArg -TargetPath $target
