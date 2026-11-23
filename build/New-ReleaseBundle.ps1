<#
.SYNOPSIS
    Packs a tool folder into <Tool>-<Version>.zip for its release.

.DESCRIPTION
    Single implementation for the release workflow and build/Test-ReleaseBundle.ps1,
    so the test checks exactly what ships.

    Left out of the ZIP:
    - release.psd1 and .gitignore (build metadata)
    - every file listed in Artifacts: it is attached to the release on its own,
      e.g. the EXE. A file that is only named in Exe but not in Artifacts stays in
      the ZIP (AcmeExchange ships Setup.exe inside the bundle).

.OUTPUTS
    Full path of the ZIP.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Tool,
    [Parameter(Mandatory)][string]$Version,
    [string]$OutDir = 'dist',
    [string]$Repo = (Get-Location).Path,
    [string]$StageRoot = $(if ($env:RUNNER_TEMP) { $env:RUNNER_TEMP } else { $env:TEMP })
)
$ErrorActionPreference = 'Stop'

$src = Join-Path $Repo $Tool
$meta = Import-PowerShellDataFile (Join-Path $src 'release.psd1')

$stage = Join-Path $StageRoot "bundle/$Tool"
if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
New-Item -ItemType Directory -Force -Path $stage | Out-Null
Copy-Item (Join-Path $src '*') $stage -Recurse -Force

$exclude = @('release.psd1', '.gitignore') + @($meta.Artifacts | Where-Object { $_ })
# -Include only works with a wildcard in the path; -Force also catches hidden files.
Get-ChildItem "$stage\*" -Include $exclude -Recurse -Force | Remove-Item -Force

if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Force -Path $OutDir | Out-Null }
$zip = Join-Path (Resolve-Path $OutDir) "$Tool-$Version.zip"
Compress-Archive -Path "$stage/*" -DestinationPath $zip -Force
$zip
