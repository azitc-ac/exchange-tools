<#
.SYNOPSIS
    Tests for the -ImportPfx mode of Invoke-AcmeExchangeCert.ps1. No Exchange server is touched:
    everything that would talk to Exchange is replaced by a stub.

.DESCRIPTION
    Creates its own test CA and three PFX files (complete chain / leaf only / expired), then checks
    the password store, the PFX validation, the mode switch and the import flow.

    -Mutate additionally runs a rollback test: each guard is removed from a COPY of the script and
    the matching test must then fail. A test that still passes proves nothing.

.EXAMPLE
    .\tests\Test-PfxMode.ps1
    .\tests\Test-PfxMode.ps1 -Mutate
#>
[CmdletBinding()]
param(
    [string]$Engine,
    [string]$WorkDir = (Join-Path $env:TEMP 'acme-pfx-tests'),
    [switch]$Mutate
)
$ErrorActionPreference = 'Stop'
# $PSScriptRoot is not populated in parameter defaults under PS 5.1 - resolve it here.
if (-not $Engine) { $Engine = Join-Path (Split-Path $PSScriptRoot -Parent) 'Invoke-AcmeExchangeCert.ps1' }
$TestPw = 'Test-Pfx-2026!'

# ---------------------------------------------------------------- helpers
$script:pass = 0; $script:fail = 0
function Check([string]$Name, [scriptblock]$Test) {
    try { if (& $Test) { Write-Host "[OK  ] $Name" -ForegroundColor Green; $script:pass++ }
          else { Write-Host "[FAIL] $Name" -ForegroundColor Red; $script:fail++ } }
    catch { Write-Host "[FAIL] $Name -> $($_.Exception.Message)" -ForegroundColor Red; $script:fail++ }
}
function Get-EngineFunctionText([string]$Path) {
    # The engine runs top to bottom, so it cannot be dot-sourced. Return just its function
    # definitions as text; the caller dot-sources them, so they land in the SCRIPT scope (dot-
    # sourcing inside this function would drop them again when it returns).
    $ast = [Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
    $fns = $ast.FindAll({ $args[0] -is [Management.Automation.Language.FunctionDefinitionAst] }, $false)
    return [pscustomobject]@{ Count = $fns.Count; Text = (($fns | ForEach-Object { $_.Extent.Text }) -join "`r`n") }
}
function Remove-TestCertsFromStores {
    # The test root must not be reachable through ANY store, otherwise the "missing intermediate"
    # case cannot occur (the chain would build with status UntrustedRoot instead of PartialChain).
    # Collect first, delete afterwards: deleting while the certificate store is being enumerated
    # silently leaves part of the matches behind.
    $stores = 'Cert:\CurrentUser\My','Cert:\CurrentUser\CA','Cert:\CurrentUser\Root',
              'Cert:\LocalMachine\My','Cert:\LocalMachine\CA','Cert:\LocalMachine\Root'
    for ($round = 1; $round -le 3; $round++) {
        $hits = @()
        foreach ($store in $stores) {
            $hits += @(Get-ChildItem $store -ErrorAction SilentlyContinue |
                       Where-Object { $_.Subject -like '*acmepfxtest*' } |
                       ForEach-Object { "$store\$($_.Thumbprint)" })
        }
        if (-not $hits.Count) { return 0 }
        foreach ($p in $hits) { Remove-Item $p -Force -ErrorAction SilentlyContinue }
    }
    # report what is still there so a test can fail loudly instead of silently testing nothing
    $left = 0
    foreach ($store in $stores) {
        $left += @(Get-ChildItem $store -ErrorAction SilentlyContinue | Where-Object { $_.Subject -like '*acmepfxtest*' }).Count
    }
    return $left
}
function New-TestPfxSet([string]$Dir) {
    New-Item -ItemType Directory -Force $Dir | Out-Null
    $pw = ConvertTo-SecureString $TestPw -AsPlainText -Force
    $ca = New-SelfSignedCertificate -Subject 'CN=acmepfxtest Root CA' -KeyUsage CertSign,CRLSign,DigitalSignature `
            -KeyLength 2048 -NotAfter (Get-Date).AddYears(5) -CertStoreLocation Cert:\CurrentUser\My `
            -TextExtension @('2.5.29.19={critical}{text}ca=1&pathlength=1')
    $leaf = New-SelfSignedCertificate -Subject 'CN=acmepfxtest.example.net' -DnsName 'acmepfxtest.example.net' `
            -Signer $ca -KeyLength 2048 -NotAfter (Get-Date).AddYears(1) -CertStoreLocation Cert:\CurrentUser\My
    $exp = New-SelfSignedCertificate -Subject 'CN=acmepfxtest-expired.example.net' -Signer $ca -KeyLength 2048 `
            -NotBefore (Get-Date).AddDays(-400) -NotAfter (Get-Date).AddDays(-10) -CertStoreLocation Cert:\CurrentUser\My
    Export-PfxCertificate -Cert $leaf -FilePath "$Dir\good-chain.pfx" -Password $pw -ChainOption BuildChain | Out-Null
    Export-PfxCertificate -Cert $leaf -FilePath "$Dir\leaf-only.pfx"  -Password $pw -ChainOption EndEntityCertOnly | Out-Null
    Export-PfxCertificate -Cert $exp  -FilePath "$Dir\expired.pfx"    -Password $pw -ChainOption BuildChain | Out-Null
    Remove-TestCertsFromStores
}
function New-TestConfig([string]$HomeDir) {
    [pscustomobject]@{
        Acme     = [pscustomobject]@{ Domains = @() }
        Exchange = [pscustomobject]@{ Servers = @('TESTSRV'); Services = 'IIS,SMTP' }
        Paths    = [pscustomobject]@{ Home = $HomeDir; BackupKeep = 3 }
        Pfx      = [pscustomobject]@{ Enabled = $true; DropFolder = "$HomeDir\pfx"
                                      PasswordFile = "$HomeDir\pfxpass.dat"; ArchiveAfterImport = $true; Subject = '' }
        Notify   = [pscustomobject]@{ SendSuccessMail = $true }
    }
}

# ---------------------------------------------------------------- setup
if (-not (Test-Path $Engine)) { throw "Engine not found: $Engine" }
Remove-Item $WorkDir -Recurse -Force -ErrorAction SilentlyContinue
$pfxSrc = Join-Path $WorkDir 'src'; $homeDir = Join-Path $WorkDir 'home'
New-Item -ItemType Directory -Force "$homeDir\pfx" | Out-Null
Write-Host "Creating test certificates in $pfxSrc ..." -ForegroundColor DarkGray
New-TestPfxSet $pfxSrc
$engineFns = Get-EngineFunctionText $Engine
. ([scriptblock]::Create($engineFns.Text))
Write-Host "Loaded $($engineFns.Count) functions from $(Split-Path $Engine -Leaf)" -ForegroundColor DarkGray
if (-not (Get-Command New-CertFromPfx -ErrorAction SilentlyContinue)) { throw 'Engine functions did not load into the script scope.' }

$script:Config = New-TestConfig $homeDir
$script:LogFile = $null
$script:NonInteractive = $true
$script:SkipInstall = $false
$ImportPfx = $true; $SetPfxPassword = $false; $PfxPath = ''
# Stub MUST write to the host, not to the pipeline - otherwise the log line would be appended to the
# return value of the function under test (New-CertFromPfx would return an array).
function Write-Log { param($Message, $Level = 'INFO') Write-Host "      $Message" -ForegroundColor DarkGray }

# stubs for everything that needs a real Exchange organisation
$script:Calls = @{ Install = 0; Expiry = 0; Mail = 0 }
$script:Deployed = $false
$script:LastMailSubject = ''
function Test-CertDeployed { param([string]$Thumbprint) return $script:Deployed }
function Install-Certificate {
    param($PACert)
    $script:Calls.Install++
    [pscustomobject]@{ Thumbprint = $PACert.Thumbprint; Subject = $PACert.Subject; Issuer = 'CN=acmepfxtest Root CA'
        NotAfter = $PACert.NotAfter; Servers = @('TESTSRV'); Connectors = @(); Removed = @(); Backup = 'n/a'; Fallback = @() }
}
function Invoke-ExpiryCheck { $script:Calls.Expiry++ }
function Write-EventLogEntry { param($Message, $Type, $Id) }
function Send-Notification { param($Subject, $Body, $Level) $script:Calls.Mail++; $script:LastMailSubject = $Subject }
function Save-Config { param($Cfg) }
function Get-LocalServerName { 'TESTSRV' }

function Reset-Case {
    $script:Calls = @{ Install = 0; Expiry = 0; Mail = 0 }
    Get-ChildItem "$homeDir\pfx" -Recurse -Filter *.pfx -ErrorAction SilentlyContinue | Remove-Item -Force
    Copy-Item "$pfxSrc\good-chain.pfx" "$homeDir\pfx\" -Force
    $script:Config.Pfx.Subject = ''
}

# ---------------------------------------------------------------- tests
Write-Host "`n=== Password store (DPAPI, machine scope) ===" -ForegroundColor Cyan
Check 'Set-PfxPasswordFile writes the file' {
    Set-PfxPasswordFile -Password (ConvertTo-SecureString $TestPw -AsPlainText -Force) | Out-Null
    Test-Path "$homeDir\pfxpass.dat"
}
Check 'the password survives a round trip' {
    $sec = Get-PfxPasswordFromFile
    $b = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
    try { [Runtime.InteropServices.Marshal]::PtrToStringBSTR($b) -eq $TestPw }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($b) }
}
Check 'the file holds no clear text' {
    -not ([Text.Encoding]::UTF8.GetString([IO.File]::ReadAllBytes("$homeDir\pfxpass.dat")) -match [regex]::Escape($TestPw))
}
Check 'ACL is SYSTEM + Administrators only' {
    $ids = @((Get-Acl "$homeDir\pfxpass.dat").Access | ForEach-Object { $_.IdentityReference.Value })
    ($ids -contains 'NT AUTHORITY\SYSTEM') -and ($ids -contains 'BUILTIN\Administrators') -and $ids.Count -eq 2
}

Write-Host "`n=== PFX validation ===" -ForegroundColor Cyan
Copy-Item "$pfxSrc\good-chain.pfx" "$homeDir\pfx\" -Force
Check 'a complete chain is accepted' {
    $c = New-CertFromPfx -Path "$homeDir\pfx\good-chain.pfx"
    $c.Thumbprint -and $c.Subject -eq 'CN=acmepfxtest.example.net' -and $c.PfxPass -is [securestring]
}
Check 'the password reaches Get-PfxPasswordSecure without Posh-ACME' {
    (Get-PfxPasswordSecure (New-CertFromPfx -Path "$homeDir\pfx\good-chain.pfx")) -is [securestring]
}
Check 'a missing intermediate is rejected' {
    # Precondition: the test CA must not be findable in any store - otherwise this would test nothing.
    $left = Remove-TestCertsFromStores
    if ($left) { throw "precondition failed: $left test certificate(s) still in a store" }
    try { New-CertFromPfx -Path "$pfxSrc\leaf-only.pfx" | Out-Null; $false } catch { $_.Exception.Message -match 'intermediate' }
}
Check 'an expired certificate is rejected' {
    try { New-CertFromPfx -Path "$pfxSrc\expired.pfx" | Out-Null; $false } catch { $_.Exception.Message -match 'expired' }
}
Check 'a wrong password is rejected' {
    Set-PfxPasswordFile -Password (ConvertTo-SecureString 'wrong' -AsPlainText -Force) | Out-Null
    try { New-CertFromPfx -Path "$homeDir\pfx\good-chain.pfx" | Out-Null; $false }
    catch { $_.Exception.Message -match 'Wrong password|Could not open' }
}
Set-PfxPasswordFile -Password (ConvertTo-SecureString $TestPw -AsPlainText -Force) | Out-Null
Check 'a missing password file gives an actionable message' {
    Move-Item "$homeDir\pfxpass.dat" "$homeDir\pfxpass.bak" -Force
    try { New-CertFromPfx -Path "$homeDir\pfx\good-chain.pfx" | Out-Null; $false }
    catch { $_.Exception.Message -match 'SetPfxPassword' }
    finally { Move-Item "$homeDir\pfxpass.bak" "$homeDir\pfxpass.dat" -Force }
}
Check 'the newest PFX in the drop folder is picked' { (Resolve-PfxFile -Path '').Name -eq 'good-chain.pfx' }

Write-Host "`n=== Mode switch ===" -ForegroundColor Cyan
Check 'Test-PfxMode follows the config' { Test-PfxMode }
Check 'Get-MainDomain derives from the remembered subject' {
    $script:Config.Pfx.Subject = 'CN=acmepfxtest.example.net'
    (Get-MainDomain) -eq 'acmepfxtest.example.net'
}
Check 'Acme.Domains still wins when set' {
    $script:Config.Acme.Domains = @('mail.example.net')
    $r = (Get-MainDomain) -eq 'mail.example.net'; $script:Config.Acme.Domains = @(); $r
}
Check 'Get-MainDomainOrNull returns null instead of throwing' {
    $s = $script:Config.Pfx.Subject; $script:Config.Pfx.Subject = ''
    try { $null -eq (Get-MainDomainOrNull) } finally { $script:Config.Pfx.Subject = $s }
}

Write-Host "`n=== Import flow ===" -ForegroundColor Cyan
Reset-Case; $script:Deployed = $false; $script:SkipInstall = $false
Check 'not deployed -> installs once, remembers the subject, archives the PFX, mails the result' {
    Invoke-PfxImport
    $script:Calls.Install -eq 1 -and $script:Config.Pfx.Subject -eq 'CN=acmepfxtest.example.net' `
        -and $script:Calls.Mail -eq 1 -and $script:LastMailSubject -match 'acmepfxtest.example.net' `
        -and (-not (Test-Path "$homeDir\pfx\good-chain.pfx")) -and @(Get-ChildItem "$homeDir\pfx\archive" -Filter *.pfx).Count -ge 1
}
Reset-Case; $script:Deployed = $true
Check 'already deployed -> no install (no iisreset), but the expiry check runs' {
    Invoke-PfxImport
    $script:Calls.Install -eq 0 -and $script:Calls.Expiry -eq 1
}
Reset-Case; $script:Deployed = $false; $script:SkipInstall = $true
Check 'SkipInstall -> nothing installed, the PFX stays put' {
    Invoke-PfxImport
    $script:Calls.Install -eq 0 -and (Test-Path "$homeDir\pfx\good-chain.pfx")
}
$script:SkipInstall = $false
Reset-Case; $script:Deployed = $false
Remove-Item "$homeDir\pfxpass.dat" -Force
Check 'no password file -> aborts before any installation' {
    try { Invoke-PfxImport; $false } catch { $script:Calls.Install -eq 0 -and $_.Exception.Message -match 'SetPfxPassword' }
}
Set-PfxPasswordFile -Password (ConvertTo-SecureString $TestPw -AsPlainText -Force) | Out-Null
Reset-Case; $script:Deployed = $false
Get-ChildItem "$homeDir\pfx" -Filter *.pfx | Remove-Item -Force
Copy-Item "$pfxSrc\expired.pfx" "$homeDir\pfx\" -Force
Check 'an expired PFX aborts before any installation' {
    try { Invoke-PfxImport; $false } catch { $script:Calls.Install -eq 0 -and $_.Exception.Message -match 'expired' }
}

Write-Host ""
Write-Host "Result: $script:pass passed, $script:fail failed" -ForegroundColor $(if ($script:fail) { 'Red' } else { 'Green' })
Remove-TestCertsFromStores

# ---------------------------------------------------------------- rollback test
if ($Mutate) {
    Write-Host "`n=== Rollback test (each guard removed from a copy) ===" -ForegroundColor Cyan
    $mutPath = Join-Path $WorkDir 'engine-mutated.ps1'
    $cases = @(
        @{ Name = 'idempotency guard';        Expect = 'already deployed'
           Find  = 'if (Test-CertDeployed $cert.Thumbprint) {'; Replace = 'if ($false) {' }
        @{ Name = 'expiry check';             Expect = 'expired PFX'
           Find  = 'if ($leaf.NotAfter -lt (Get-Date))'; Replace = 'if ($false)' }
        @{ Name = 'chain check (both parts)'; Expect = 'missing intermediate'
           Find  = "if (`$stati -contains 'PartialChain') {"; Replace = 'if ($false) {'
           Find2 = 'if (-not $selfSigned -and $chain.ChainElements.Count -lt 2) {'; Replace2 = 'if ($false) {' }
    )
    $ok = $true
    foreach ($c in $cases) {
        $text = [IO.File]::ReadAllText($Engine)
        if ($text -notmatch [regex]::Escape($c.Find)) { Write-Host "[FAIL] pattern not found: $($c.Find)" -ForegroundColor Red; $ok = $false; continue }
        $text = $text.Replace($c.Find, $c.Replace)
        if ($c.ContainsKey('Find2')) { $text = $text.Replace($c.Find2, $c.Replace2) }
        [IO.File]::WriteAllText($mutPath, $text, (New-Object Text.UTF8Encoding $true))
        $out = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath -Engine $mutPath -WorkDir (Join-Path $WorkDir 'mut') 2>&1 | Out-String
        $hit = @($out -split "`n" | Where-Object { $_ -match '^\[FAIL\]' -and $_ -match [regex]::Escape($c.Expect) })
        if ($hit.Count) { Write-Host "[OK  ] $($c.Name) -> test fails as it should" -ForegroundColor Green }
        else { Write-Host "[FAIL] $($c.Name) -> no test noticed the removal; the guard is unverified" -ForegroundColor Red; $ok = $false }
    }
    Remove-Item $mutPath -Force -ErrorAction SilentlyContinue
    Write-Host ""
    if ($ok) { Write-Host 'Rollback test passed.' -ForegroundColor Green } else { Write-Host 'Rollback test FAILED.' -ForegroundColor Red; exit 1 }
}

if ($script:fail) { exit 1 }
