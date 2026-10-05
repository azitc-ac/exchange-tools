<#
    Build Setup.exe from AcmeExchangeSetup.ps1 using ps2exe.
    Build-time only dependency: ps2exe (Install-Module ps2exe -Scope CurrentUser).
    The resulting Setup.exe is a launcher for the GUI; it must stay in the bundle
    folder next to Invoke-AcmeExchangeCert.ps1 and lib\.
#>
[CmdletBinding()]
param(
    [string]$Root = $PSScriptRoot,
    [string]$Version = ''   # empty = take it from AcmeExchangeSetup.ps1
)
$ErrorActionPreference = 'Stop'
Import-Module ps2exe -ErrorAction Stop

# $PSScriptRoot can arrive empty when launched via -File in some shells; fall back to the invocation path.
if (-not $Root) {
    $Root = if ($PSScriptRoot) { $PSScriptRoot }
            elseif ($MyInvocation.MyCommand.Path) { Split-Path $MyInvocation.MyCommand.Path -Parent }
            else { (Get-Location).Path }
}

$in  = Join-Path $Root 'AcmeExchangeSetup.ps1'
$out = Join-Path $Root 'Setup.exe'
$ico = Join-Path $Root 'icon.ico'
if (-not (Test-Path $in)) { throw "Not found: $in" }

# Maintain the version in exactly one place: it lives in AcmeExchangeSetup.ps1.
. (Join-Path $Root '..\build\Get-ToolVersion.ps1')
if (-not $Version) {
    $Version = Get-ToolVersion -Path $in
    Write-Host "Version from AcmeExchangeSetup.ps1: $Version" -ForegroundColor DarkGray
} else {
    # An explicitly passed version must match the script, otherwise the file version lies.
    $null = Assert-ToolVersion -Path $in -Expected $Version
}

$p2e = @{
    inputFile   = $in
    outputFile  = $out
    noConsole   = $true
    STA         = $true
    requireAdmin = $true
    title       = 'Exchange ACME Certificate Setup'
    description = 'Setup and management GUI for Exchange ACME certificate renewal'
    company     = 'azitc'
    product     = 'Exchange-ACME-Cert'
    version     = $Version
}
if (Test-Path $ico) { $p2e.iconFile = $ico }   # embed the app/window icon when present
Invoke-ps2exe @p2e

if (Test-Path $out) {
    Write-Host "Built: $out ($([math]::Round((Get-Item $out).Length/1KB,1)) KB)" -ForegroundColor Green
} else {
    throw 'Build failed: Setup.exe not produced.'
}
