<#
.SYNOPSIS
    Liest die Version eines Werkzeugs aus seinem Hauptskript.

.DESCRIPTION
    Eine einzige Stelle, an der die Versionsnummer steht: die Zeile

        $script:Version = '1.2.3'

    im Hauptskript des Werkzeugs. Build-Skripte und der Release-Workflow lesen sie
    von hier, damit Dateiversion, Release-Tag und Skript nicht auseinanderlaufen.

    Dot-sourcen und dann Get-ToolVersion aufrufen:

        . (Join-Path $PSScriptRoot '..\build\Get-ToolVersion.ps1')
        $v = Get-ToolVersion -Path .\LogViewer\LogViewer.ps1
#>

function Get-ToolVersion {
    [CmdletBinding()]
    param(
        # Hauptskript des Werkzeugs
        [Parameter(Mandatory)]
        [string]$Path
    )
    if (-not (Test-Path -LiteralPath $Path)) { throw "Hauptskript nicht gefunden: $Path" }

    # Einfache Anführungszeichen: sonst würde PowerShell $script:Version hier selbst ersetzen.
    $m = [regex]::Match([IO.File]::ReadAllText($Path), '\$script:Version\s*=\s*''([0-9]+\.[0-9]+\.[0-9]+)''')
    if (-not $m.Success) {
        throw "In '$Path' fehlt die Zeile `$script:Version = 'x.y.z' - ohne sie ist kein Release möglich."
    }
    $m.Groups[1].Value
}

function Assert-ToolVersion {
    <#
        Vergleicht die Version im Hauptskript mit der aus dem Release-Tag erwarteten.
        Wirft bei Abweichung - genau dann, wenn jemand einen Tag setzt, ohne die
        Version im Skript nachzuziehen (oder umgekehrt).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Expected
    )
    $actual = Get-ToolVersion -Path $Path
    if ($actual -ne $Expected) {
        throw ("Versionsdrift: Tag sagt '{0}', `$script:Version in '{1}' sagt '{2}'. " -f $Expected, $Path, $actual) +
              'Entweder den Tag korrigieren oder die Version im Skript nachziehen.'
    }
    Write-Host "Version stimmt überein: $actual ($Path)" -ForegroundColor Green
    $actual
}
