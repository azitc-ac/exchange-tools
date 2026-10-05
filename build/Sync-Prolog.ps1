<#
.SYNOPSIS
    Schreibt den Prolog aus build/Prolog.*.ps1 in die Hauptskripte der Werkzeuge.

.DESCRIPTION
    Die Werkzeuge sollen einzeln lauffähig bleiben - deshalb wird der Prolog
    einkopiert statt zur Laufzeit geladen. Damit die Kopien nicht auseinanderlaufen,
    ist build/Prolog.*.ps1 die Quelle und dieses Skript der einzige Weg, sie zu
    verteilen. build/Test-Prolog.ps1 meldet jede Abweichung.

    Das Zielskript muss die Marker bereits enthalten:

        # <prolog:exo v1 ...>
        ...
        # </prolog:exo>

    Beim ersten Mal werden sie von Hand gesetzt; danach pflegt dieses Skript den Inhalt.

.EXAMPLE
    .\build\Sync-Prolog.ps1              # schreibt
    .\build\Sync-Prolog.ps1 -WhatIf      # zeigt nur, was sich änderte
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$Repo
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

function Get-PrologText([string]$Repo, [string]$Art) {
    $p = Join-Path $Repo ("build\Prolog.{0}.ps1" -f (Get-Culture).TextInfo.ToTitleCase($Art))
    if (-not (Test-Path $p)) { throw "Prolog-Quelle fehlt: $p" }
    # Ohne abschließenden Zeilenumbruch, damit der Vergleich stabil bleibt.
    ([IO.File]::ReadAllText($p, [Text.Encoding]::UTF8)).TrimEnd("`r", "`n")
}

function Set-PrologInFile {
    param([string]$Path, [string]$Art, [string]$Text)

    $orig = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
    $open = "# <prolog:$Art "
    $close = "# </prolog:$Art>"

    $i = $orig.IndexOf($open)
    if ($i -lt 0) { throw "In '$Path' fehlt der Startmarker '$open'. Einmalig von Hand setzen." }
    $j = $orig.IndexOf($close, $i)
    if ($j -lt 0) { throw "In '$Path' fehlt der Endmarker '$close'." }
    $j += $close.Length

    $neu = $orig.Substring(0, $i) + $Text + $orig.Substring($j)
    if ($neu -ceq $orig) { return 'unverändert' }

    if ($PSCmdlet.ShouldProcess($Path, 'Prolog aktualisieren')) {
        # UTF-8 mit BOM: Windows PowerShell 5.1 liest Umlaute sonst als Mojibake.
        [IO.File]::WriteAllText($Path, $neu, (New-Object Text.UTF8Encoding $true))
    }
    'aktualisiert'
}

$map = Import-PowerShellDataFile (Join-Path $Repo 'build\prolog-map.psd1')
foreach ($t in $map.Tools) {
    $path = Join-Path $Repo (Join-Path $t.Tool $t.Main)
    if (-not (Test-Path $path)) { throw "Hauptskript aus der Map fehlt: $path" }
    $text = Get-PrologText -Repo $Repo -Art $t.Prolog
    $r = Set-PrologInFile -Path $path -Art $t.Prolog -Text $text
    "{0,-22} {1,-8} {2}" -f $t.Tool, $t.Prolog, $r
}
