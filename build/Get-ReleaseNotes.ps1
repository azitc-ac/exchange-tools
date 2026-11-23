<#
.SYNOPSIS
    Builds the release notes for one tool (used by the release workflow and by
    build/Test-ReleaseBundle.ps1).

.DESCRIPTION
    Visible text lives in build/release-notes-*.md, read as UTF-8 (a workflow block
    is saved without BOM and PowerShell 5.1 would garble umlauts).
    - release-notes-footer.md: always; {ZIP} becomes <Tool>-<Version>.zip
    - release-notes-exe.md:    only when the tool ships an EXE (Exe in release.psd1)
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Tool,
    [Parameter(Mandatory)][string]$Version,
    [string]$Repo = (Get-Location).Path
)
$ErrorActionPreference = 'Stop'
$meta = Import-PowerShellDataFile (Join-Path $Repo "$Tool/release.psd1")

function Read-Utf8([string]$name) {
    [IO.File]::ReadAllText((Join-Path $Repo "build\$name"), [Text.Encoding]::UTF8).TrimEnd()
}

$parts = @(
    "Werkzeug ``$Tool`` Version $Version aus dem Sammelrepo."
    ''
    $meta.Notes
    ''
    ((Read-Utf8 'release-notes-footer.md') -replace '\{ZIP\}', "$Tool-$Version.zip")
)
if ($meta.Exe) {
    $parts += ''
    $parts += (Read-Utf8 'release-notes-exe.md') -replace '\{EXE\}', $meta.Exe
}
$parts -join "`n"
