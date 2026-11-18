<#
.SYNOPSIS
    Zeigt, welche Werkzeuge eine Version haben, zu der es noch kein Release gibt.

.DESCRIPTION
    Jedes Werkzeug mit release.psd1 trägt seine Version als $script:Version im
    Hauptskript. Gibt es zu dieser Version keinen Tag <Werkzeug>/v<Version>, fehlt das
    Release - die EXE bzw. das ZIP auf GitHub ist dann älter als das Skript.

    Der Workflow .github/workflows/auto-release.yml ruft dieses Skript nach jeder grünen
    CI auf main mit -MissingOnly auf und legt die fehlenden Releases selbst an.

    Zum Dot-sourcen (nur die Funktion) oder direkt ausführen (Tabelle).

.EXAMPLE
    .\build\Get-ReleaseDrift.ps1                 # alle Werkzeuge, Stand der Releases
    .\build\Get-ReleaseDrift.ps1 -MissingOnly    # nur die ohne Release
#>
[CmdletBinding()]
param(
    [string]$Repo,
    [switch]$MissingOnly
)

function Get-ReleaseDrift {
    <#
        Reine Logik: Werkzeugordner + vorhandene Tags -> je Werkzeug Version, erwarteter
        Tag und ob er existiert. Ein Werkzeug ohne $script:Version wirft (Get-ToolVersion),
        denn dann wäre gar kein Release möglich.
    #>
    param(
        [Parameter(Mandatory)][string]$Repo,
        [string[]]$Tags = @()
    )
    . (Join-Path $Repo 'build\Get-ToolVersion.ps1')
    foreach ($dir in (Get-ChildItem -LiteralPath $Repo -Directory | Sort-Object Name)) {
        $psd = Join-Path $dir.FullName 'release.psd1'
        if (-not (Test-Path -LiteralPath $psd)) { continue }
        $meta = Import-PowerShellDataFile -LiteralPath $psd
        $ver  = Get-ToolVersion -Path (Join-Path $dir.FullName $meta.Main)
        $tag  = "$($dir.Name)/v$ver"
        [pscustomobject]@{
            Tool     = $dir.Name
            Version  = $ver
            Tag      = $tag
            Released = ($Tags -contains $tag)
        }
    }
}

# Direkt aufgerufen (nicht dot-gesourct): Tags aus git lesen und ausgeben.
if ($MyInvocation.InvocationName -ne '.') {
    $ErrorActionPreference = 'Stop'
    if (-not $Repo) {
        $hier = if ($PSScriptRoot) { $PSScriptRoot }
                elseif ($MyInvocation.MyCommand.Path) { Split-Path $MyInvocation.MyCommand.Path -Parent }
                else { (Get-Location).Path }
        $Repo = Split-Path $hier -Parent
    }
    $tags = @(git -C $Repo tag -l)
    if ($LASTEXITCODE -ne 0) { throw 'git tag -l ist fehlgeschlagen.' }
    $drift = @(Get-ReleaseDrift -Repo $Repo -Tags $tags)
    if ($MissingOnly) { $drift | Where-Object { -not $_.Released } }
    else { $drift | Format-Table Tool, Version, Tag, Released -AutoSize | Out-String -Width 200 | Write-Output }
}
