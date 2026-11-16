<#
.SYNOPSIS
    Runs Publish-InstallResult of Invoke-AcmeExchangeCert.ps1 under StrictMode 2 with the result
    shapes a real install produces. No Exchange server, mail server or event log is touched.

.DESCRIPTION
    A function that returns an empty array returns $null, and one element comes back as a scalar.
    Under StrictMode in Windows PowerShell 5.1 both throw on .Count. A run without connector
    changes (the normal case) therefore failed AFTER the certificate was installed - exit code 1,
    error mail instead of the success mail.

.EXAMPLE
    .\tests\Test-InstallReport.ps1
#>
[CmdletBinding()]
param([string]$Engine)
$ErrorActionPreference = 'Stop'
if (-not $Engine) { $Engine = Join-Path (Split-Path $PSScriptRoot -Parent) 'Invoke-AcmeExchangeCert.ps1' }

$script:pass = 0; $script:fail = 0
function Check([string]$Name, [scriptblock]$Test) {
    try { if (& $Test) { Write-Host "[OK  ] $Name" -ForegroundColor Green; $script:pass++ }
          else { Write-Host "[FAIL] $Name" -ForegroundColor Red; $script:fail++ } }
    catch { Write-Host "[FAIL] $Name -> $($_.Exception.Message)" -ForegroundColor Red; $script:fail++ }
}

$ast = [Management.Automation.Language.Parser]::ParseFile($Engine, [ref]$null, [ref]$null)
$fn = $ast.FindAll({ $args[0] -is [Management.Automation.Language.FunctionDefinitionAst] -and $args[0].Name -eq 'Publish-InstallResult' }, $false) | Select-Object -First 1
if (-not $fn) { throw 'Publish-InstallResult not found in engine' }
. ([scriptblock]::Create($fn.Extent.Text))

# stubs: capture instead of writing the event log / sending mail
$script:Logged = $null
function Write-EventLogEntry { param([string]$Message, $Type, $Id) $script:Logged = $Message }
function Send-Notification { param($Subject, $Body, $Level) }
function Write-Log { param($Message, $Level) }
$script:Config = [pscustomobject]@{ Notify = [pscustomobject]@{ SendSuccessMail = $true }; Exchange = [pscustomobject]@{ Services = 'IIS,SMTP' } }

function New-Result($Connectors, $Removed, $Fallback) {
    [pscustomobject]@{ Thumbprint = 'AB'; Subject = 'CN=mail.example.com'; Issuer = 'CN=R10'; NotAfter = (Get-Date)
                       Servers = 'EX01'; Connectors = $Connectors; Removed = $Removed; Backup = 'C:\b.pfx'; Fallback = $Fallback }
}

Set-StrictMode -Version 2
Check 'nothing changed ($null everywhere) - the normal renewal' {
    Publish-InstallResult -Result (New-Result $null $null $null) -Headline 'ok'
    $script:Logged -match 'Connectors updated: none' -and $script:Logged -match 'Old certificates removed: none' }
Check 'one element each (scalars)' {
    Publish-InstallResult -Result (New-Result "ReceiveConnector 'X'" 'old' 'EX01') -Headline 'ok'
    $script:Logged -match "ReceiveConnector 'X'" -and $script:Logged -match 'KSP fallback used on: EX01' }
Check 'several elements (arrays)' {
    Publish-InstallResult -Result (New-Result @('A', 'B') @('o1', 'o2') @('EX01', 'EX02')) -Headline 'ok'
    $script:Logged -match 'EX01, EX02' }
Set-StrictMode -Off

Write-Host ''
Write-Host ("{0} passed, {1} failed" -f $script:pass, $script:fail)
if ($script:fail) { exit 1 }
