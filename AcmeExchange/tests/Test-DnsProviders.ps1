<#
.SYNOPSIS
    Checks the DNS provider table (Get-DnsProviders) of Invoke-AcmeExchangeCert.ps1 against the
    bundled Posh-ACME plugins. No ACME server, DNS API or Exchange server is touched.

.DESCRIPTION
    A provider entry is only usable if
      - its plugin file exists in lib\Posh-ACME\Plugins,
      - every field key is a parameter of the plugin's Add-DnsTxt,
      - every mandatory Add-DnsTxt parameter (besides RecordName/TxtValue) has a field,
      - Secure matches the parameter type (SecureString <=> Secure = $true),
      - the README provider table names it.
    A typo in a field key otherwise only shows up as a failed order on the customer's server.

.EXAMPLE
    .\tests\Test-DnsProviders.ps1
#>
[CmdletBinding()]
param([string]$Engine)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if (-not $Engine) { $Engine = Join-Path $root 'Invoke-AcmeExchangeCert.ps1' }
$pluginDir = Join-Path $root 'lib\Posh-ACME\Plugins'
$readme = Get-Content (Join-Path $root 'README.md') -Raw -Encoding UTF8

$script:pass = 0; $script:fail = 0
function Check([string]$Name, [scriptblock]$Test) {
    try { if (& $Test) { Write-Host "[OK  ] $Name" -ForegroundColor Green; $script:pass++ }
          else { Write-Host "[FAIL] $Name" -ForegroundColor Red; $script:fail++ } }
    catch { Write-Host "[FAIL] $Name -> $($_.Exception.Message)" -ForegroundColor Red; $script:fail++ }
}

# load only Get-DnsProviders from the engine (the engine itself runs top to bottom)
$ast = [Management.Automation.Language.Parser]::ParseFile($Engine, [ref]$null, [ref]$null)
$fn = $ast.FindAll({ $args[0] -is [Management.Automation.Language.FunctionDefinitionAst] -and $args[0].Name -eq 'Get-DnsProviders' }, $false) | Select-Object -First 1
if (-not $fn) { throw 'Get-DnsProviders not found in engine' }
. ([scriptblock]::Create($fn.Extent.Text))
$providers = Get-DnsProviders

Check 'provider table is not empty' { $providers.Count -ge 2 }

foreach ($prov in $providers.Keys) {
    $def = $providers[$prov]
    # README may space a trailing number ("Route 53")
    $rx = ([regex]::Escape($prov)) -replace '(?<=[A-Za-z])(?=\d)', '\s?'
    Check "$prov - README provider table lists it" { $readme -match "\|\s*\*\*$rx\*\*" }
    if ($def.Bootstrap) { continue }   # Azure has its own bootstrap and no fields

    $file = Join-Path $pluginDir "$($def.Plugin).ps1"
    Check "$prov - plugin file $($def.Plugin).ps1 bundled" { Test-Path $file }
    if (-not (Test-Path $file)) { continue }

    # real cmdlet metadata instead of parsing attributes by hand (Mandatory = $false, several
    # [Parameter()] per parameter and parameter sets are all resolved by PowerShell itself)
    $add = & { . $file; Get-Command Add-DnsTxt -CommandType Function }
    $fieldKeys = @($def.Fields | ForEach-Object { $_.Key })
    # AcmeExchangeSetup.ps1 (Set-ProviderView) has exactly two credential boxes
    Check "$prov - at most 2 fields (GUI has two credential boxes)" { $fieldKeys.Count -le 2 }

    foreach ($f in $def.Fields) {
        $prm = $add.Parameters[$f.Key]
        Check "$prov - field $($f.Key) is an Add-DnsTxt parameter" { $null -ne $prm }
        if ($prm) {
            Check "$prov - field $($f.Key) Secure=$($f.Secure) matches parameter type" { [bool]$f.Secure -eq ($prm.ParameterType -eq [securestring]) }
        }
    }
    # a plugin may offer alternative parameter sets (token vs. key+secret, deprecated insecure).
    # The fields must fit ONE set completely: all its mandatory parameters present, nothing foreign.
    $fits = @($add.ParameterSets | Where-Object {
        $set = $_
        $names = @($set.Parameters | ForEach-Object { $_.Name })
        $mand  = @($set.Parameters | Where-Object { $_.IsMandatory -and $_.Name -notin 'RecordName', 'TxtValue' } | ForEach-Object { $_.Name })
        (@($mand | Where-Object { $_ -notin $fieldKeys }).Count -eq 0) -and (@($fieldKeys | Where-Object { $_ -notin $names }).Count -eq 0)
    })
    Check "$prov - fields [$($fieldKeys -join ', ')] cover one parameter set completely" { $fits.Count -ge 1 }
}

Write-Host ''
Write-Host ("{0} passed, {1} failed" -f $script:pass, $script:fail)
if ($script:fail) { exit 1 }
