<#
.SYNOPSIS
    Checks that Invoke-AcmeExchangeCert.ps1 only creates ACME accounts with Posh-ACME's portable
    (AES) plugin encryption. No ACME server is touched.

.DESCRIPTION
    Without -UseAltPluginEncryption, Posh-ACME stores secret plugin args (DNS API keys) DPAPI-
    encrypted for the current user. The certificate is usually issued by an admin in the GUI and
    renewed by the SYSTEM task, which then cannot decrypt the DNS credentials - renewal fails
    after ~60 days, long after setup looked fine.

.EXAMPLE
    .\tests\Test-PluginEncryption.ps1
#>
[CmdletBinding()]
param([string]$Engine)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if (-not $Engine) { $Engine = Join-Path $root 'Invoke-AcmeExchangeCert.ps1' }

$script:pass = 0; $script:fail = 0
function Check([string]$Name, [scriptblock]$Test) {
    try { if (& $Test) { Write-Host "[OK  ] $Name" -ForegroundColor Green; $script:pass++ }
          else { Write-Host "[FAIL] $Name" -ForegroundColor Red; $script:fail++ } }
    catch { Write-Host "[FAIL] $Name -> $($_.Exception.Message)" -ForegroundColor Red; $script:fail++ }
}

$ast = [Management.Automation.Language.Parser]::ParseFile($Engine, [ref]$null, [ref]$null)
$calls = @($ast.FindAll({ $args[0] -is [Management.Automation.Language.CommandAst] -and $args[0].GetCommandName() -in 'New-PAAccount', 'New-PACertificate' }, $true))
$acct = @($calls | Where-Object { $_.GetCommandName() -eq 'New-PAAccount' })

Check 'New-PAAccount is called (account creation found)' { $acct.Count -ge 1 }
foreach ($c in $acct) {
    $line = $c.Extent.StartLineNumber
    Check "New-PAAccount (line $line) uses -UseAltPluginEncryption" {
        @($c.CommandElements | Where-Object { $_ -is [Management.Automation.Language.CommandParameterAst] -and $_.ParameterName -eq 'UseAltPluginEncryption' }).Count -eq 1 }
}
# New-PACertificate creates the account itself (without AES) if none exists - every call must be
# preceded by Initialize-PAAccountSafe in the same function.
foreach ($c in @($calls | Where-Object { $_.GetCommandName() -eq 'New-PACertificate' })) {
    $fn = $c.Parent; while ($fn -and $fn -isnot [Management.Automation.Language.FunctionDefinitionAst]) { $fn = $fn.Parent }
    $line = $c.Extent.StartLineNumber
    Check "New-PACertificate (line $line) in $($fn.Name) runs after Initialize-PAAccountSafe" {
        $init = @($fn.FindAll({ $args[0] -is [Management.Automation.Language.CommandAst] -and $args[0].GetCommandName() -eq 'Initialize-PAAccountSafe' }, $true))
        $init.Count -ge 1 -and $init[0].Extent.StartOffset -lt $c.Extent.StartOffset }
}
# an existing account from an older version must be migrated on renewal, too
$ren = $ast.FindAll({ $args[0] -is [Management.Automation.Language.FunctionDefinitionAst] -and $args[0].Name -eq 'Invoke-Renewal' }, $false) | Select-Object -First 1
Check 'Invoke-Renewal migrates existing accounts (Initialize-PAAccountSafe before Submit-Renewal)' {
    $i = @($ren.FindAll({ $args[0] -is [Management.Automation.Language.CommandAst] -and $args[0].GetCommandName() -eq 'Initialize-PAAccountSafe' }, $true))
    $s = @($ren.FindAll({ $args[0] -is [Management.Automation.Language.CommandAst] -and $args[0].GetCommandName() -eq 'Submit-Renewal' }, $true))
    $i.Count -ge 1 -and $s.Count -ge 1 -and $i[-1].Extent.StartOffset -lt $s[0].Extent.StartOffset }

# Set-PAOrder -PluginArgs REPLACES all args of the plugin (Posh-ACME Export-PluginArgs). Passing only
# the non-secret args from config wiped the stored API secret on the first unattended renewal.
$mf = $ast.FindAll({ $args[0] -is [Management.Automation.Language.FunctionDefinitionAst] -and $args[0].Name -eq 'Merge-PluginArgs' }, $false) | Select-Object -First 1
Check 'Merge-PluginArgs exists' { $null -ne $mf }
if ($mf) {
    . ([scriptblock]::Create($mf.Extent.Text))
    $sec = ConvertTo-SecureString 'x' -AsPlainText -Force
    Check 'Merge-PluginArgs keeps the stored secret when only the prefix is new' {
        $m = Merge-PluginArgs @{ IONOSKeyPrefix = 'old'; IONOSKeySecret = $sec } @{ IONOSKeyPrefix = 'new' }
        $m.IONOSKeyPrefix -eq 'new' -and $m.IONOSKeySecret -eq $sec -and $m.Count -eq 2 }
    Check 'Merge-PluginArgs: new secret wins, nothing stored is fine' {
        $m = Merge-PluginArgs $null @{ A = 1 }; $n = Merge-PluginArgs @{ A = 1 } @{ A = 2 }
        $m.A -eq 1 -and $n.A -eq 2 }
}
foreach ($c in @($ast.FindAll({ $args[0] -is [Management.Automation.Language.CommandAst] -and $args[0].GetCommandName() -eq 'Set-PAOrder' }, $true))) {
    $txt = $c.Extent.Text
    if ($txt -notmatch '-PluginArgs') { continue }
    Check "Set-PAOrder (line $($c.Extent.StartLineNumber)) passes the complete arg set (Merge-PluginArgs or Azure's full set)" {
        $txt -match '-PluginArgs\s+\(\s*(Merge-PluginArgs|Get-AzurePluginArgs)\b' }
}

Write-Host ''
Write-Host ("{0} passed, {1} failed" -f $script:pass, $script:fail)
if ($script:fail) { exit 1 }
