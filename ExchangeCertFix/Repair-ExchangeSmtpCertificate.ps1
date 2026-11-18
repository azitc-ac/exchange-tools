<#
.SYNOPSIS
    Diagnoses and remediates issues preventing the default SMTP certificate
    from being replaced on Exchange Server.

.DESCRIPTION
    Common root causes covered by this script:
      1. Private key stored in a CNG Key Storage Provider (KSP) instead of a
         legacy CSP. Exchange transport cannot use CNG keys reliably.
      2. Missing ACL on the private key for NETWORK SERVICE. Import-ExchangeCertificate
         does not grant it; Enable-ExchangeCertificate -Services SMTP does (measured on
         Exchange SE: no ACE after import, Read ACE right after enabling SMTP).
      3. The "Overwrite the existing default SMTP certificate?" question. Without -Force
         Exchange asks it AFTER it has already added SMTP and the ACL; answering No keeps
         the internal transport certificate (msExchServerInternalTLSCert), Yes or -Force
         replaces it. In a non-interactive session the question throws instead.
      4. Receive Connectors with a hard-coded TlsCertificateName pointing to an
         old thumbprint, which overrides the default SMTP certificate.
      5. Stale expired certificates still carrying the SMTP service flag.
      6. Transport services holding a handle on the previous certificate.
      7. Internal transport certificate pointing to a certificate that is gone from
         the store or expired (transport event 12013 "could not load the certificate
         with thumbprint ... from the personal store"). Fix: make the target the
         default, restore a valid self-signed Exchange certificate as default
         (Microsoft's recommendation), or create a new self-signed one.
      8. "Exchange Back End" site (port 444) bound to a certificate that is expired
         or gone - typically after the self-signed certificate reached its 5 years.

    The script is read-only by default. Every write operation is presented as an
    explicit prompt. Nothing is changed without confirmation.

.PARAMETER Thumbprint
    Target certificate thumbprint. If omitted, an interactive selection dialog
    is shown listing all Exchange certificates.

.PARAMETER WorkingFolder
    Folder used for temporary PFX export during CSP conversion.
    Default: $env:TEMP

.PARAMETER DiagnoseOnly
    Runs all checks and prints the report, but never offers any remediation.

.PARAMETER KeepDefaultSmtp
    Enables SMTP on the target certificate WITHOUT making it the internal transport
    (default SMTP) certificate. Without this switch the script recommends Yes or No
    based on the current default and asks.

.PARAMETER RestoreSelfSignedDefault
    Makes a valid self-signed Exchange certificate (CN=<server>) the internal transport
    certificate again - or creates a new one if none is valid. Runs even if no other
    issue was found. Microsoft: "Typically, you don't need to replace the default
    certificate with a certificate from a commercial CA".

.EXAMPLE
    .\Repair-ExchangeSmtpCertificate.ps1

.EXAMPLE
    .\Repair-ExchangeSmtpCertificate.ps1 -Thumbprint AABBCC... -DiagnoseOnly

.NOTES
    Author  : AZITC - Alexander Zarenko IT Consulting
    Requires: Exchange Management Shell, PowerShell 5.1, local administrator
    Tested  : Exchange 2016 / 2019 / SE
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidatePattern('^[0-9A-Fa-f]{40}$')]
    [string]$Thumbprint,

    [Parameter(Mandatory = $false)]
    [string]$WorkingFolder = $env:TEMP,

    [Parameter(Mandatory = $false)]
    [switch]$DiagnoseOnly,

    [Parameter(Mandatory = $false)]
    [switch]$KeepDefaultSmtp,

    [Parameter(Mandatory = $false)]
    [switch]$RestoreSelfSignedDefault
)

if ($KeepDefaultSmtp -and $RestoreSelfSignedDefault) {
    throw '-KeepDefaultSmtp and -RestoreSelfSignedDefault exclude each other.'
}

# Einzige Stelle fuer die Versionsnummer; Build und Release-Tag lesen sie hier aus.
$script:Version = '1.2.0'


#region ---------------------------------------------------------- Helper functions

$script:ChangeLog = New-Object System.Collections.ArrayList

function Write-Section {
    param([string]$Text)
    Write-Host ''
    Write-Host ('=' * 78) -ForegroundColor DarkCyan
    Write-Host "  $Text" -ForegroundColor Cyan
    Write-Host ('=' * 78) -ForegroundColor DarkCyan
}

function Write-Info    { param([string]$m) Write-Host "[INFO]  $m" -ForegroundColor Gray }
function Write-Ok      { param([string]$m) Write-Host "[OK]    $m" -ForegroundColor Green }
function Write-Warn    { param([string]$m) Write-Host "[WARN]  $m" -ForegroundColor Yellow }
function Write-Problem { param([string]$m) Write-Host "[ISSUE] $m" -ForegroundColor Red }
function Write-Action  { param([string]$m) Write-Host "[WRITE] $m" -ForegroundColor Magenta }

function Confirm-Write {
    <#
        Central confirmation gate. Every write operation in this script must be
        routed through this function. Returns $true only on explicit consent.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Action,
        [Parameter(Mandatory = $false)][string]$Detail,
        [Parameter(Mandatory = $false)][string]$Impact
    )

    if ($DiagnoseOnly) {
        Write-Info "DiagnoseOnly mode - skipping remediation: $Action"
        return $false
    }

    Write-Host ''
    Write-Host '---------------------------------------------------------------' -ForegroundColor DarkYellow
    Write-Action "Proposed change: $Action"
    if ($Detail) { Write-Host "        Detail : $Detail" -ForegroundColor Gray }
    if ($Impact) { Write-Host "        Impact : $Impact" -ForegroundColor Yellow }
    Write-Host '---------------------------------------------------------------' -ForegroundColor DarkYellow

    do {
        $answer = Read-Host 'Apply this change? [y] Yes  [n] No  [a] Abort script'
        switch ($answer.ToLower()) {
            'y' { return $true }
            'n' { Write-Info 'Skipped by user.'; return $false }
            'a' { Write-Warn 'Aborted by user.'; exit 1 }
            default { Write-Host 'Please enter y, n or a.' -ForegroundColor Red }
        }
    } while ($true)
}

function Add-ChangeLogEntry {
    param([string]$Entry)
    [void]$script:ChangeLog.Add($Entry)
}

function Test-ExchangeShell {
    if (-not (Get-Command Get-ExchangeCertificate -ErrorAction SilentlyContinue)) {
        Write-Problem 'Exchange cmdlets are not available in this session.'
        Write-Info    'Run this script from the Exchange Management Shell.'
        return $false
    }
    return $true
}

function Test-Elevation {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $pr = New-Object Security.Principal.WindowsPrincipal($id)
    return $pr.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-CertificateProviderName {
    <#
        Returns the cryptographic provider name of the private key.
        A CNG provider ("Key Storage Provider") is the primary root cause for
        SMTP certificates that cannot be activated by Exchange.
    #>
    param([Parameter(Mandatory = $true)][string]$Thumbprint)

    $result = [PSCustomObject]@{
        ProviderName = $null
        IsCng        = $false
        HasPrivateKey = $false
        KeyContainer = $null
    }

    $cert = Get-Item -Path ("Cert:\LocalMachine\My\$Thumbprint") -ErrorAction SilentlyContinue
    if (-not $cert) { return $result }

    $result.HasPrivateKey = $cert.HasPrivateKey
    if (-not $cert.HasPrivateKey) { return $result }

    # certutil is the most reliable way to read the provider on PS 5.1 without
    # touching the CNG/CAPI interop APIs directly.
    $certutil = certutil -store My $Thumbprint 2>&1 | Out-String

    $match = [regex]::Match($certutil, 'Provider\s*=\s*(.+)')
    if ($match.Success) {
        $result.ProviderName = $match.Groups[1].Value.Trim()
    }

    $containerMatch = [regex]::Match($certutil, 'Key Container\s*=\s*(.+)')
    if ($containerMatch.Success) {
        $result.KeyContainer = $containerMatch.Groups[1].Value.Trim()
    }

    if ($result.ProviderName -and $result.ProviderName -match 'Key Storage Provider') {
        $result.IsCng = $true
    }

    # Fallback detection when certutil output is unexpected
    if (-not $result.ProviderName) {
        try {
            if ($cert.PrivateKey -eq $null) {
                # No legacy CSP object exposed -> almost certainly CNG
                $result.IsCng = $true
                $result.ProviderName = 'Unknown (no legacy CSP handle - likely CNG)'
            }
            else {
                $result.ProviderName = $cert.PrivateKey.CspKeyContainerInfo.ProviderName
            }
        }
        catch {
            $result.IsCng = $true
            $result.ProviderName = 'Unknown (private key not accessible via CSP)'
        }
    }

    return $result
}

function Get-PrivateKeyFilePath {
    <#
        Resolves the on-disk private key file for a certificate so that its ACL
        can be inspected and corrected.
    #>
    param([Parameter(Mandatory = $true)][string]$Thumbprint)

    $cert = Get-Item -Path ("Cert:\LocalMachine\My\$Thumbprint") -ErrorAction SilentlyContinue
    if (-not $cert -or -not $cert.HasPrivateKey) { return $null }

    $containerName = $null

    try {
        if ($cert.PrivateKey -and $cert.PrivateKey.CspKeyContainerInfo) {
            $containerName = $cert.PrivateKey.CspKeyContainerInfo.UniqueKeyContainerName
        }
    }
    catch { }

    if (-not $containerName) {
        try {
            $cng = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
            if ($cng -and $cng.Key) {
                $containerName = $cng.Key.UniqueName
            }
        }
        catch { }
    }

    if (-not $containerName) { return $null }

    $searchRoots = @(
        (Join-Path $env:ProgramData 'Microsoft\Crypto\RSA\MachineKeys'),
        (Join-Path $env:ProgramData 'Microsoft\Crypto\Keys')
    )

    foreach ($root in $searchRoots) {
        if (-not (Test-Path $root)) { continue }
        $file = Join-Path $root $containerName
        if (Test-Path $file) { return $file }
    }

    return $null
}

function Test-PrivateKeyAcl {
    <#
        Exchange transport runs as NETWORK SERVICE and requires read access to
        the private key file.
    #>
    param([Parameter(Mandatory = $true)][string]$KeyFilePath)

    $result = [PSCustomObject]@{
        Path              = $KeyFilePath
        NetworkServiceOk  = $false
        Identities        = @()
    }

    try {
        $acl = Get-Acl -Path $KeyFilePath -ErrorAction Stop
        $result.Identities = $acl.Access | ForEach-Object {
            '{0} : {1}' -f $_.IdentityReference, $_.FileSystemRights
        }

        foreach ($ace in $acl.Access) {
            $identity = $ace.IdentityReference.Value
            if ($identity -match 'NETWORK SERVICE|NETZWERKDIENST|S-1-5-20') {
                $read = [System.Security.AccessControl.FileSystemRights]::Read
                if ($ace.AccessControlType -eq 'Allow' -and
                    ((($ace.FileSystemRights -band $read) -eq $read) -or
                     $ace.FileSystemRights.ToString() -match 'FullControl')) {
                    $result.NetworkServiceOk = $true
                }
            }
        }
    }
    catch {
        Write-Warn "Could not read ACL of '$KeyFilePath': $($_.Exception.Message)"
    }

    return $result
}

function Get-InternalTransportCertificate {
    <#
        Reads the internal transport ("default SMTP") certificate of a server from AD
        (msExchServerInternalTLSCert). Get-TransportService caches this value per
        session, so AD is the authoritative source. Returns $null if the attribute is
        empty, or an object with Error set if it cannot be read.
    #>
    param([string]$Server = $env:COMPUTERNAME)
    try {
        $cfg = ([ADSI]'LDAP://RootDSE').configurationNamingContext.ToString()
        $searcher = New-Object System.DirectoryServices.DirectorySearcher(
            [ADSI]"LDAP://$cfg", "(&(objectClass=msExchExchangeServer)(cn=$Server))")
        [void]$searcher.PropertiesToLoad.Add('msexchserverinternaltlscert')
        $hit = $searcher.FindOne()
        if (-not $hit) { return [PSCustomObject]@{ Error = "Exchange server object '$Server' not found in AD." } }
        $raw = $hit.Properties['msexchserverinternaltlscert']
        if (-not $raw -or $raw.Count -eq 0) { return $null }
        $c = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(, [byte[]]$raw[0])
        return [PSCustomObject]@{
            Thumbprint   = $c.Thumbprint
            Subject      = $c.Subject
            NotAfter     = $c.NotAfter
            IsSelfSigned = ($c.Subject -eq $c.Issuer)
            InStore      = [bool](Test-Path -LiteralPath "Cert:\LocalMachine\My\$($c.Thumbprint)")
            Error        = $null
        }
    }
    catch {
        return [PSCustomObject]@{ Error = $_.Exception.Message }
    }
}

function Get-DefaultSmtpRecommendation {
    <#
        What to do with the internal transport (default SMTP) certificate. Pure function
        (no Exchange calls) so it can be tested.
          Keep              - current default is a valid self-signed Exchange certificate
          RestoreSelfSigned - default is CA-issued or broken, and a valid self-signed
                              Exchange certificate exists (Microsoft's recommendation)
          CreateSelfSigned  - default is broken and no valid self-signed one exists
          Overwrite         - default is a valid CA-issued certificate and no valid
                              self-signed one exists: move it to the target
          AlreadyDefault    - the target already is the default (and it is self-signed,
                              or no self-signed alternative exists)
          Broken            - $true when the current default is missing, gone or expired
    #>
    param(
        $Current,
        [Parameter(Mandatory = $true)]$Target,
        [bool]$SelfSignedAvailable = $false,
        [datetime]$Now = (Get-Date)
    )
    $r = { param($a, $why, $broken) [PSCustomObject]@{ Action = $a; Reason = $why; Broken = [bool]$broken } }
    if ($Current -and $Current.Error) {
        return & $r 'Unknown' "The internal transport certificate could not be read: $($Current.Error)" $false
    }

    $why = $null
    if ($null -eq $Current)            { $why = 'No internal transport certificate is set for this server.' }
    elseif (-not $Current.InStore)     { $why = "The default points to $($Current.Thumbprint), which is no longer in the local store (transport event 12013)." }
    elseif ($Current.NotAfter -lt $Now){ $why = "The default $($Current.Thumbprint) expired on $($Current.NotAfter.ToString('yyyy-MM-dd'))." }
    if ($why) {
        if ($SelfSignedAvailable) { return & $r 'RestoreSelfSigned' "$why A valid self-signed Exchange certificate exists - make it the default." $true }
        return & $r 'CreateSelfSigned' "$why No valid self-signed Exchange certificate exists - create a new one." $true
    }

    if ($Current.IsSelfSigned) {
        if ($Current.Thumbprint -eq $Target.Thumbprint) { return & $r 'AlreadyDefault' 'The target already is the internal transport certificate.' $false }
        return & $r 'Keep' 'The default is a valid self-signed Exchange certificate. Keep it for TLS between Exchange servers; it does not follow the public renewal cycle.' $false
    }

    # Valid, CA-issued default.
    if ($SelfSignedAvailable) {
        return & $r 'RestoreSelfSigned' "The default is the CA-issued certificate $($Current.Subject). Microsoft: a commercial certificate is typically not needed for internal SMTP - a valid self-signed Exchange certificate exists, make it the default again." $false
    }
    if ($Current.Thumbprint -eq $Target.Thumbprint) { return & $r 'AlreadyDefault' 'The target already is the internal transport certificate.' $false }
    return & $r 'Overwrite' "The default is the CA-issued certificate $($Current.Subject) - usually the one being replaced, and no valid self-signed alternative exists. Move the default to the target." $false
}

function ConvertFrom-NetshSslCert {
    <#
        Parses "netsh http show sslcert" output into bindings (IpPort, Thumbprint,
        AppId). Pure function so it can be tested; field names are matched loosely
        because netsh labels are localized only in parts.
    #>
    param([string]$Text)
    $result = @()
    $cur = $null
    foreach ($line in ($Text -split "`r?`n")) {
        if ($line -match '^\s*(IP:port|Hostname:port)\s*:\s*(\S+)') {
            if ($cur) { $result += [PSCustomObject]$cur }
            $cur = [ordered]@{ IpPort = $Matches[2]; Thumbprint = $null; AppId = $null }
        }
        elseif ($cur -and $line -match '^\s*Certificate Hash\s*:\s*([0-9a-fA-F]{40})') { $cur.Thumbprint = $Matches[1].ToUpper() }
        elseif ($cur -and $line -match '^\s*Application ID\s*:\s*(\{[0-9a-fA-F-]{36}\})') { $cur.AppId = $Matches[1] }
    }
    if ($cur) { $result += [PSCustomObject]$cur }
    return $result
}

function Get-BackendBinding {
    <#
        HTTPS bindings of the "Exchange Back End" site (port 444) with the state of the
        bound certificate.
    #>
    $text = (netsh http show sslcert 2>$null | Out-String)
    $bindings = @(ConvertFrom-NetshSslCert -Text $text | Where-Object { $_.IpPort -match ':444$' })
    foreach ($b in $bindings) {
        $c = if ($b.Thumbprint) { Get-Item -LiteralPath "Cert:\LocalMachine\My\$($b.Thumbprint)" -ErrorAction SilentlyContinue }
        $b | Add-Member -NotePropertyName InStore  -NotePropertyValue ([bool]$c)
        $b | Add-Member -NotePropertyName NotAfter -NotePropertyValue $(if ($c) { $c.NotAfter } else { $null })
        $b | Add-Member -NotePropertyName Subject  -NotePropertyValue $(if ($c) { $c.Subject } else { $null })
    }
    return $bindings
}

function Get-SelfSignedCandidate {
    <#
        A valid self-signed Exchange server certificate (CN=<server>, private key,
        more than 30 days left). Prefers the one bound to the Exchange Back End site.
    #>
    param([string[]]$PreferThumbprint = @())
    $limit = (Get-Date).AddDays(30)
    $all = @(Get-ExchangeCertificate -ErrorAction SilentlyContinue | Where-Object {
        $_.IsSelfSigned -and $_.Subject -eq "CN=$env:COMPUTERNAME" -and $_.HasPrivateKey -and $_.NotAfter -gt $limit
    })
    if ($all.Count -eq 0) { return $null }
    $preferred = @($all | Where-Object { $PreferThumbprint -contains $_.Thumbprint })
    if ($preferred.Count -gt 0) { return $preferred[0] }
    return ($all | Sort-Object NotAfter -Descending | Select-Object -First 1)
}

function New-SelfSignedExchangeCertificate {
    <#
        Creates a new self-signed Exchange certificate like Setup does (CN=<server>,
        SAN server name + FQDN, 5 years), but WITHOUT services - so it cannot take
        over the IIS front end binding. Services are assigned explicitly afterwards.
    #>
    $fqdn = (Get-ExchangeServer -Identity $env:COMPUTERNAME -ErrorAction Stop).Fqdn
    $names = @($env:COMPUTERNAME, $fqdn) | Select-Object -Unique
    return New-ExchangeCertificate -SubjectName "CN=$env:COMPUTERNAME" -DomainName $names `
        -FriendlyName 'Microsoft Exchange' -Services None -PrivateKeyExportable $true -Force -ErrorAction Stop
}

function Get-SmtpCertificateSnapshot {
    <#
        Returns the current state of all Exchange certificates carrying the SMTP
        service flag, plus the TlsCertificateName of every Receive Connector.
        Used for the before/after comparison.
    #>
    $snapshot = [PSCustomObject]@{
        Timestamp          = Get-Date
        SmtpCertificates   = @()
        ReceiveConnectors  = @()
        InternalTransport  = Get-InternalTransportCertificate
        Backend            = @(Get-BackendBinding)
    }

    $snapshot.SmtpCertificates = @(
        Get-ExchangeCertificate -ErrorAction SilentlyContinue |
            Where-Object { $_.Services -match 'SMTP' } |
            Select-Object Thumbprint,
                          Subject,
                          NotAfter,
                          IsSelfSigned,
                          @{ Name = 'Services'; Expression = { $_.Services.ToString() } }
    )

    $snapshot.ReceiveConnectors = @(
        Get-ReceiveConnector -ErrorAction SilentlyContinue |
            Select-Object @{ Name = 'Identity';           Expression = { $_.Identity.ToString() } },
                          @{ Name = 'TlsCertificateName'; Expression = { if ($_.TlsCertificateName) { $_.TlsCertificateName.ToString() } else { '<not set>' } } }
    )

    return $snapshot
}

function Show-SmtpCertificateSnapshot {
    param(
        [Parameter(Mandatory = $true)]$Snapshot,
        [Parameter(Mandatory = $true)][string]$Label
    )

    Write-Section "$Label - default SMTP certificate state"

    if ($Snapshot.SmtpCertificates.Count -eq 0) {
        Write-Warn 'No certificate currently carries the SMTP service flag.'
    }
    else {
        Write-Host ''
        Write-Host 'Certificates with SMTP service flag:' -ForegroundColor White
        $Snapshot.SmtpCertificates |
            Format-Table Thumbprint, Subject, NotAfter, IsSelfSigned, Services -AutoSize |
            Out-String -Width 200 |
            Write-Host
    }

    $it = $Snapshot.InternalTransport
    Write-Host 'Internal transport (default SMTP) certificate, from AD:' -ForegroundColor White
    if ($null -eq $it)   { Write-Host '  <not set>' -ForegroundColor Yellow }
    elseif ($it.Error)   { Write-Host "  <unreadable: $($it.Error)>" -ForegroundColor Yellow }
    else {
        $flags = @()
        if ($it.IsSelfSigned)              { $flags += 'self-signed' }
        if (-not $it.InStore)              { $flags += 'NOT IN STORE' }
        if ($it.NotAfter -lt (Get-Date))   { $flags += 'EXPIRED' }
        Write-Host ("  {0}  {1}  until {2}  {3}" -f $it.Thumbprint, $it.Subject, $it.NotAfter.ToString('yyyy-MM-dd'), ($flags -join ', '))
    }
    Write-Host ''

    Write-Host 'Exchange Back End site (port 444) binding:' -ForegroundColor White
    if ($Snapshot.Backend.Count -eq 0) { Write-Host '  <no binding on port 444 found>' -ForegroundColor Yellow }
    foreach ($b in $Snapshot.Backend) {
        $state = if (-not $b.InStore) { 'NOT IN STORE' } elseif ($b.NotAfter -lt (Get-Date)) { 'EXPIRED' } else { "until $($b.NotAfter.ToString('yyyy-MM-dd'))" }
        Write-Host ("  {0}  {1}  {2}  {3}" -f $b.IpPort, $b.Thumbprint, $b.Subject, $state)
    }
    Write-Host ''

    Write-Host 'Receive Connector TLS certificate bindings:' -ForegroundColor White
    $Snapshot.ReceiveConnectors |
        Format-Table Identity, TlsCertificateName -AutoSize |
        Out-String -Width 200 |
        Write-Host
}

function Select-ExchangeCertificate {
    <#
        Interactive selection dialog for the target certificate.
    #>
    $certs = @(Get-ExchangeCertificate -ErrorAction Stop | Sort-Object NotAfter -Descending)

    if ($certs.Count -eq 0) {
        Write-Problem 'No Exchange certificates found on this server.'
        return $null
    }

    Write-Section 'Select target certificate'

    $index = 0
    $table = foreach ($c in $certs) {
        $index++
        $daysLeft = [math]::Round(($c.NotAfter - (Get-Date)).TotalDays)
        [PSCustomObject]@{
            '#'          = $index
            Thumbprint   = $c.Thumbprint
            Subject      = $c.Subject
            NotAfter     = $c.NotAfter.ToString('yyyy-MM-dd')
            DaysLeft     = $daysLeft
            Services     = $c.Services.ToString()
            SelfSigned   = $c.IsSelfSigned
        }
    }

    $table | Format-Table -AutoSize | Out-String -Width 220 | Write-Host

    do {
        $choice = Read-Host "Enter number 1-$($certs.Count) (or 'a' to abort)"
        if ($choice -eq 'a') { Write-Warn 'Aborted by user.'; exit 1 }
        $parsed = 0
        $valid = [int]::TryParse($choice, [ref]$parsed) -and $parsed -ge 1 -and $parsed -le $certs.Count
        if (-not $valid) { Write-Host 'Invalid selection.' -ForegroundColor Red }
    } while (-not $valid)

    return $certs[$parsed - 1]
}

function Get-TlsCertificateName {
    <#
        Builds the "<I>Issuer<S>Subject" string required by Set-ReceiveConnector.
    #>
    param([Parameter(Mandatory = $true)]$Certificate)
    return ('<I>{0}<S>{1}' -f $Certificate.Issuer, $Certificate.Subject)
}

#endregion

#region ---------------------------------------------------------- Preflight

Clear-Host
Write-Section 'Exchange SMTP certificate - diagnostics and remediation'
Write-Info "Server        : $env:COMPUTERNAME"
Write-Info "Run as        : $env:USERDOMAIN\$env:USERNAME"
Write-Info "Mode          : $(if ($DiagnoseOnly) { 'DIAGNOSE ONLY (read-only)' } else { 'INTERACTIVE (write operations require confirmation)' })"
Write-Info "Working folder: $WorkingFolder"

if (-not (Test-ExchangeShell)) { exit 1 }

if (-not (Test-Elevation)) {
    Write-Problem 'This script must run elevated (Run as Administrator).'
    exit 1
}
Write-Ok 'Exchange Management Shell detected and session is elevated.'

if (-not (Test-Path $WorkingFolder)) {
    Write-Problem "Working folder does not exist: $WorkingFolder"
    exit 1
}

#endregion

#region ---------------------------------------------------------- BEFORE snapshot

$before = Get-SmtpCertificateSnapshot
Show-SmtpCertificateSnapshot -Snapshot $before -Label 'BEFORE'

#endregion

#region ---------------------------------------------------------- Target selection

if ($Thumbprint) {
    $target = Get-ExchangeCertificate -Thumbprint $Thumbprint -ErrorAction SilentlyContinue
    if (-not $target) {
        Write-Problem "Certificate with thumbprint '$Thumbprint' not found."
        exit 1
    }
}
else {
    $target = Select-ExchangeCertificate
    if (-not $target) { exit 1 }
}

Write-Section 'Target certificate'
Write-Host "  Thumbprint : $($target.Thumbprint)"
Write-Host "  Subject    : $($target.Subject)"
Write-Host "  Issuer     : $($target.Issuer)"
Write-Host "  Valid until: $($target.NotAfter)"
Write-Host "  Services   : $($target.Services)"
Write-Host "  Self-signed: $($target.IsSelfSigned)"
Write-Host "  SAN        : $(($target.CertificateDomains | ForEach-Object { $_.Address }) -join ', ')"

#endregion

#region ---------------------------------------------------------- Diagnostics

Write-Section 'Diagnostics'

$issues = New-Object System.Collections.ArrayList

# --- Check 1: certificate validity -------------------------------------------
if ($target.NotAfter -lt (Get-Date)) {
    Write-Problem "Certificate is EXPIRED (NotAfter: $($target.NotAfter))."
    [void]$issues.Add('Expired')
}
elseif ($target.NotBefore -gt (Get-Date)) {
    Write-Problem "Certificate is not yet valid (NotBefore: $($target.NotBefore))."
    [void]$issues.Add('NotYetValid')
}
else {
    Write-Ok "Certificate validity period is OK (expires $($target.NotAfter.ToString('yyyy-MM-dd')))."
}

# --- Check 2: private key present --------------------------------------------
$providerInfo = Get-CertificateProviderName -Thumbprint $target.Thumbprint

if (-not $providerInfo.HasPrivateKey) {
    Write-Problem 'Certificate has NO private key. It cannot be used for SMTP.'
    [void]$issues.Add('NoPrivateKey')
}
else {
    Write-Ok 'Private key is present.'
    Write-Info "Provider: $($providerInfo.ProviderName)"
    if ($providerInfo.KeyContainer) {
        Write-Info "Key container: $($providerInfo.KeyContainer)"
    }
}

# --- Check 3: CNG vs CSP (primary root cause) --------------------------------
if ($providerInfo.HasPrivateKey) {
    if ($providerInfo.IsCng) {
        Write-Problem 'Private key uses a CNG Key Storage Provider (KSP).'
        Write-Info    'Exchange transport requires a legacy CSP such as'
        Write-Info    '"Microsoft RSA SChannel Cryptographic Provider".'
        Write-Info    'This is the most common reason why the default SMTP'
        Write-Info    'certificate cannot be overwritten.'
        [void]$issues.Add('CngProvider')
    }
    else {
        Write-Ok "Private key uses a legacy CSP - compatible with Exchange transport."
    }
}

# --- Check 4: private key ACL -------------------------------------------------
$keyFile = $null
if ($providerInfo.HasPrivateKey) {
    $keyFile = Get-PrivateKeyFilePath -Thumbprint $target.Thumbprint
    if (-not $keyFile) {
        Write-Warn 'Could not resolve the private key file on disk. ACL check skipped.'
    }
    else {
        Write-Info "Private key file: $keyFile"
        $aclInfo = Test-PrivateKeyAcl -KeyFilePath $keyFile
        if ($aclInfo.NetworkServiceOk) {
            Write-Ok 'NETWORK SERVICE has read access to the private key.'
        }
        else {
            Write-Problem 'NETWORK SERVICE has NO read access to the private key.'
            Write-Info    'Exchange transport will fail to load this certificate.'
            [void]$issues.Add('MissingAcl')
        }
    }
}

# --- Check 5: SMTP flag already set on target --------------------------------
$targetHasSmtp = ($target.Services -match 'SMTP')
if ($targetHasSmtp) {
    Write-Ok 'Target certificate already carries the SMTP service flag.'
}
else {
    Write-Warn 'Target certificate does NOT carry the SMTP service flag yet.'
    [void]$issues.Add('SmtpFlagMissing')
}

# --- Check 6: competing certificates with SMTP flag --------------------------
$competing = @(
    Get-ExchangeCertificate |
        Where-Object { $_.Services -match 'SMTP' -and $_.Thumbprint -ne $target.Thumbprint }
)

if ($competing.Count -gt 0) {
    Write-Warn "$($competing.Count) other certificate(s) also carry the SMTP flag:"
    foreach ($c in $competing) {
        $expiredTag = if ($c.NotAfter -lt (Get-Date)) { ' [EXPIRED]' } else { '' }
        $selfTag    = if ($c.IsSelfSigned) { ' [self-signed]' } else { '' }
        Write-Host "         $($c.Thumbprint) - $($c.Subject)$expiredTag$selfTag" -ForegroundColor Yellow
    }
    Write-Info 'Note: the Exchange self-signed certificate legitimately keeps SMTP'
    Write-Info 'for internal transport TLS. Expired third-party certificates do not.'

    $expiredCompeting = @($competing | Where-Object { $_.NotAfter -lt (Get-Date) -and -not $_.IsSelfSigned })
    if ($expiredCompeting.Count -gt 0) {
        [void]$issues.Add('StaleSmtpCertificates')
    }
}
else {
    Write-Ok 'No competing certificates hold the SMTP flag.'
}

# --- Check 7: Receive Connector TlsCertificateName ---------------------------
$expectedTlsName = Get-TlsCertificateName -Certificate $target
$connectors = @(Get-ReceiveConnector -ErrorAction SilentlyContinue)
$mismatchedConnectors = New-Object System.Collections.ArrayList

foreach ($rc in $connectors) {
    $current = if ($rc.TlsCertificateName) { $rc.TlsCertificateName.ToString() } else { $null }
    if ($current -and $current -ne $expectedTlsName) {
        [void]$mismatchedConnectors.Add(
            [PSCustomObject]@{
                Identity = $rc.Identity.ToString()
                Current  = $current
            }
        )
    }
}

if ($mismatchedConnectors.Count -gt 0) {
    Write-Problem "$($mismatchedConnectors.Count) Receive Connector(s) have a hard-coded TlsCertificateName that does NOT match the target certificate."
    Write-Info 'Enable-ExchangeCertificate has no effect on these connectors.'
    foreach ($m in $mismatchedConnectors) {
        Write-Host "         $($m.Identity)" -ForegroundColor Yellow
        Write-Host "           current : $($m.Current)" -ForegroundColor DarkGray
    }
    [void]$issues.Add('ConnectorTlsMismatch')
}
else {
    Write-Ok 'No Receive Connector has a conflicting TlsCertificateName binding.'
}

# --- Check 8: internal transport (default SMTP) certificate ------------------
$internalTransport = $before.InternalTransport
$backendThumbs     = @($before.Backend | ForEach-Object { $_.Thumbprint })
$selfSignedCand    = Get-SelfSignedCandidate -PreferThumbprint $backendThumbs
$defaultAdvice     = Get-DefaultSmtpRecommendation -Current $internalTransport -Target $target -SelfSignedAvailable ([bool]$selfSignedCand)
if ($defaultAdvice.Broken) {
    Write-Problem "Internal transport certificate is broken. $($defaultAdvice.Reason)"
    [void]$issues.Add('InternalTransportCertBroken')
}
elseif ($defaultAdvice.Action -in @('AlreadyDefault', 'Keep')) { Write-Ok "Internal transport certificate: $($defaultAdvice.Reason)" }
elseif ($defaultAdvice.Action -eq 'Unknown')                  { Write-Warn $defaultAdvice.Reason }
else {
    Write-Info "Internal transport certificate: $($defaultAdvice.Reason)"
    if ($defaultAdvice.Action -eq 'RestoreSelfSigned') {
        Write-Info "Run with -RestoreSelfSignedDefault to switch the default to $($selfSignedCand.Thumbprint)."
    }
}

# --- Check 9: Exchange Back End binding (port 444) --------------------------
$brokenBackend = @($before.Backend | Where-Object { -not $_.InStore -or $_.NotAfter -lt (Get-Date) })
if ($before.Backend.Count -eq 0) {
    Write-Warn 'No HTTPS binding on port 444 found (Exchange Back End). Not checked further.'
}
elseif ($brokenBackend.Count -gt 0) {
    foreach ($b in $brokenBackend) {
        Write-Problem "Exchange Back End binding $($b.IpPort) uses $($b.Thumbprint), which is $(if (-not $b.InStore) { 'no longer in the store' } else { 'expired' })."
    }
    [void]$issues.Add('BackendCertBroken')
}
else {
    Write-Ok 'Exchange Back End (port 444) is bound to a valid certificate.'
}

# --- Summary ------------------------------------------------------------------
Write-Section 'Diagnostic summary'
if ($issues.Count -eq 0) {
    Write-Ok 'No issues detected. The certificate should be usable for SMTP.'
}
else {
    Write-Host "Issues found: $($issues.Count)" -ForegroundColor Red
    foreach ($i in $issues) { Write-Host "  - $i" -ForegroundColor Red }
}

if ($DiagnoseOnly) {
    Write-Section 'DiagnoseOnly mode - no changes were made'
    return
}

if ($issues.Count -eq 0 -and -not $RestoreSelfSignedDefault) {
    Write-Info 'Nothing to remediate.'
    $after = Get-SmtpCertificateSnapshot
    Show-SmtpCertificateSnapshot -Snapshot $after -Label 'AFTER (unchanged)'
    return
}

#endregion

#region ---------------------------------------------------------- Remediation

Write-Section 'Remediation'

$activeThumbprint = $target.Thumbprint
$activeCert       = $target
$restartRequired  = $false

# --- Blocking conditions -----------------------------------------------------
if ($issues -contains 'NoPrivateKey') {
    Write-Problem 'Cannot remediate: the certificate has no private key.'
    Write-Info    'Re-request the certificate using New-ExchangeCertificate -GenerateRequest'
    Write-Info    'so the key is created in the correct CSP from the start.'
    return
}

if ($issues -contains 'Expired') {
    Write-Problem 'Cannot remediate: the certificate is expired. Obtain a new certificate first.'
    return
}

# --- Remediation 1: CNG -> CSP conversion via export/reimport ----------------
if ($issues -contains 'CngProvider') {

    $pfxPath = Join-Path $WorkingFolder ("ExchCert_{0}_{1}.pfx" -f $target.Thumbprint.Substring(0, 8), (Get-Date -Format 'yyyyMMdd_HHmmss'))

    $ok = Confirm-Write `
        -Action 'Convert private key from CNG (KSP) to legacy CSP' `
        -Detail "Export certificate to '$pfxPath', then reimport via Import-ExchangeCertificate. A NEW thumbprint may be generated." `
        -Impact 'Temporary PFX file containing the private key will be written to disk and deleted afterwards. Existing certificate object is not removed.'

    if ($ok) {
        try {
            $pfxPassword = Read-Host -AsSecureString -Prompt 'Enter a temporary PFX password'
            $pfxConfirm  = Read-Host -AsSecureString -Prompt 'Confirm the PFX password'

            $p1 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($pfxPassword))
            $p2 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($pfxConfirm))

            if ($p1 -ne $p2) {
                Write-Problem 'Passwords do not match. Conversion aborted.'
            }
            else {
                Write-Info 'Exporting certificate including private key...'
                Export-ExchangeCertificate -Thumbprint $target.Thumbprint `
                                           -FileName $pfxPath `
                                           -BinaryEncoded `
                                           -Password $pfxPassword `
                                           -ErrorAction Stop | Out-Null
                Write-Ok "Exported to $pfxPath"
                Add-ChangeLogEntry "Exported certificate $($target.Thumbprint) to $pfxPath"

                Write-Info 'Reimporting via Import-ExchangeCertificate (key will be recreated in a legacy CSP)...'
                $fileData = [Byte[]](Get-Content -Path $pfxPath -Encoding Byte -ReadCount 0)

                $imported = Import-ExchangeCertificate -FileData $fileData `
                                                       -Password $pfxPassword `
                                                       -PrivateKeyExportable $true `
                                                       -ErrorAction Stop

                if ($imported -and $imported.Thumbprint) {
                    $activeThumbprint = $imported.Thumbprint
                    Write-Ok "Reimported. Active thumbprint is now: $activeThumbprint"
                    Add-ChangeLogEntry "Reimported certificate as $activeThumbprint (CSP conversion)"
                }
                else {
                    Write-Warn 'Import returned no thumbprint. Continuing with the original thumbprint.'
                }

                # Verify the conversion actually worked
                $newProvider = Get-CertificateProviderName -Thumbprint $activeThumbprint
                if ($newProvider.IsCng) {
                    Write-Warn 'Private key is STILL reported as CNG after reimport.'
                    Write-Info 'Attempting explicit CSP import via certutil as a fallback...'

                    $ok2 = Confirm-Write `
                        -Action 'Force CSP import via certutil' `
                        -Detail "certutil -csp `"Microsoft RSA SChannel Cryptographic Provider`" -importpfx `"$pfxPath`"" `
                        -Impact 'Installs the certificate into LocalMachine\My with an explicitly forced legacy CSP.'

                    if ($ok2) {
                        $certutilOut = certutil -p $p1 -csp 'Microsoft RSA SChannel Cryptographic Provider' -importpfx $pfxPath 2>&1 | Out-String
                        Write-Host $certutilOut -ForegroundColor DarkGray
                        Add-ChangeLogEntry 'Forced CSP import via certutil'

                        $recheck = Get-CertificateProviderName -Thumbprint $activeThumbprint
                        if ($recheck.IsCng) {
                            Write-Problem 'CSP conversion failed. Manual intervention required.'
                        }
                        else {
                            Write-Ok "Provider is now: $($recheck.ProviderName)"
                        }
                    }
                }
                else {
                    Write-Ok "Provider is now: $($newProvider.ProviderName)"
                }

                $restartRequired = $true
            }
        }
        catch {
            Write-Problem "CSP conversion failed: $($_.Exception.Message)"
        }
        finally {
            # Always remove the PFX - it contains the private key
            if (Test-Path $pfxPath) {
                Remove-Item -Path $pfxPath -Force -ErrorAction SilentlyContinue
                Write-Info 'Temporary PFX file removed.'
            }
            # Clear plaintext passwords from memory
            $p1 = $null
            $p2 = $null
            [GC]::Collect()
        }

        # Refresh the certificate object after a possible thumbprint change
        $activeCert = Get-ExchangeCertificate -Thumbprint $activeThumbprint -ErrorAction SilentlyContinue
        if (-not $activeCert) {
            Write-Problem 'Could not reload the certificate after conversion. Aborting.'
            return
        }
    }
}

# --- Remediation 2: private key ACL ------------------------------------------
$currentKeyFile = Get-PrivateKeyFilePath -Thumbprint $activeThumbprint
if ($currentKeyFile) {
    $aclCheck = Test-PrivateKeyAcl -KeyFilePath $currentKeyFile
    if (-not $aclCheck.NetworkServiceOk) {

        $ok = Confirm-Write `
            -Action 'Grant NETWORK SERVICE read access to the private key' `
            -Detail "Key file: $currentKeyFile" `
            -Impact 'Adds a Read ACE for NT AUTHORITY\NETWORK SERVICE. No existing permissions are removed.'

        if ($ok) {
            try {
                $acl  = Get-Acl -Path $currentKeyFile
                $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                    'NT AUTHORITY\NETWORK SERVICE',
                    'Read',
                    'Allow'
                )
                $acl.AddAccessRule($rule)
                Set-Acl -Path $currentKeyFile -AclObject $acl -ErrorAction Stop
                Write-Ok 'ACL updated - NETWORK SERVICE now has read access.'
                Add-ChangeLogEntry "Granted NETWORK SERVICE Read on $currentKeyFile"
                $restartRequired = $true
            }
            catch {
                Write-Problem "Failed to update ACL: $($_.Exception.Message)"
            }
        }
    }
}

# --- Remediation 3: remove SMTP flag from stale certificates -----------------
$staleCerts = @(
    Get-ExchangeCertificate |
        Where-Object {
            $_.Services -match 'SMTP' -and
            $_.Thumbprint -ne $activeThumbprint -and
            $_.NotAfter -lt (Get-Date) -and
            -not $_.IsSelfSigned
        }
)

foreach ($stale in $staleCerts) {

    $ok = Confirm-Write `
        -Action 'Remove SMTP service flag from an expired certificate' `
        -Detail "Thumbprint: $($stale.Thumbprint)  Subject: $($stale.Subject)  Expired: $($stale.NotAfter.ToString('yyyy-MM-dd'))" `
        -Impact 'Sets Services to None on this certificate. The certificate itself is NOT deleted.'

    if ($ok) {
        try {
            Enable-ExchangeCertificate -Thumbprint $stale.Thumbprint -Services None -Force -ErrorAction Stop
            Write-Ok "SMTP flag removed from $($stale.Thumbprint)."
            Add-ChangeLogEntry "Removed SMTP flag from expired certificate $($stale.Thumbprint)"
            $restartRequired = $true
        }
        catch {
            Write-Problem "Failed to remove SMTP flag: $($_.Exception.Message)"
        }
    }
}

# --- Remediation 4: SMTP on the target + internal transport (default) certificate
$activeCert  = Get-ExchangeCertificate -Thumbprint $activeThumbprint -ErrorAction SilentlyContinue
$needEnable  = [bool]($activeCert -and $activeCert.Services -notmatch 'SMTP')
$needDefault = [bool]($activeCert -and (($issues -contains 'InternalTransportCertBroken') -or $RestoreSelfSignedDefault))

if ($needEnable -or $needDefault) {

    # Services: asked only if SMTP is missing on the target.
    $services = $null
    if ($needEnable) {
        Write-Host ''
        Write-Host 'Which services should be enabled on the target certificate?' -ForegroundColor White
        Write-Host '  [1] SMTP only'
        Write-Host '  [2] SMTP and IIS  (recommended for a standard Exchange certificate)'
        Write-Host '  [3] SMTP, IIS, IMAP, POP'
        Write-Host '  [4] Skip'

        do {
            $svcChoice = Read-Host 'Selection [1-4]'
        } while ($svcChoice -notin @('1', '2', '3', '4'))

        $services = switch ($svcChoice) {
            '1' { 'SMTP' }
            '2' { 'SMTP,IIS' }
            '3' { 'SMTP,IIS,IMAP,POP' }
            '4' { $null }
        }
    }

    # Default SMTP certificate: keep / target / restore self-signed / new self-signed
    $selfSigned = Get-SelfSignedCandidate -PreferThumbprint @($before.Backend | ForEach-Object { $_.Thumbprint })
    $advice     = Get-DefaultSmtpRecommendation -Current (Get-InternalTransportCertificate) -Target $activeCert -SelfSignedAvailable ([bool]$selfSigned)

    if ($KeepDefaultSmtp) {
        $defaultMode = 'keep'
        Write-Info '-KeepDefaultSmtp: the internal transport certificate will NOT be changed.'
        if ($advice.Broken) { Write-Warn "Note: $($advice.Reason)" }
    }
    elseif ($RestoreSelfSignedDefault) {
        $defaultMode = if ($selfSigned) { 'selfsigned' } else { 'newselfsigned' }
    }
    elseif ($advice.Action -eq 'AlreadyDefault' -and -not $needDefault) {
        $defaultMode = 'keep'
    }
    else {
        $options = [ordered]@{}
        if (-not $advice.Broken) { $options['k'] = 'Keep the current default' }
        $options['t'] = 'Make the TARGET certificate the default'
        if ($selfSigned) { $options['s'] = "Make the self-signed $($selfSigned.Thumbprint) (until $($selfSigned.NotAfter.ToString('yyyy-MM-dd'))) the default" }
        $options['n'] = 'Create a NEW self-signed Exchange certificate (5 years) and make it the default'

        $recommended = switch ($advice.Action) {
            'Keep'              { 'k' }
            'AlreadyDefault'    { 'k' }
            'Overwrite'         { 't' }
            'RestoreSelfSigned' { 's' }
            'CreateSelfSigned'  { 'n' }
            default             { if ($options.Contains('k')) { 'k' } else { 't' } }
        }

        Write-Host ''
        Write-Host 'Internal transport (default SMTP) certificate:' -ForegroundColor White
        Write-Host "  Recommendation: [$recommended] - $($advice.Reason)"
        Write-Host '  It is used for TLS between Exchange servers. Inbound STARTTLS from the internet'
        Write-Host '  is chosen by connector FQDN / TlsCertificateName and does not need it.'
        foreach ($k in $options.Keys) { Write-Host "  [$k] $($options[$k])" }
        do {
            $answer = Read-Host "Selection [$(@($options.Keys) -join '/')] (Enter = $recommended)"
            if ([string]::IsNullOrWhiteSpace($answer)) { $answer = $recommended }
        } while (-not $options.Contains($answer))
        $defaultMode = switch ($answer) { 'k' { 'keep' } 't' { 'target' } 's' { 'selfsigned' } 'n' { 'newselfsigned' } }
    }

    # Step A: SMTP on the target. -Force only if the target shall become the default;
    # for self-signed modes the default is moved explicitly in step B right after.
    if ($needEnable -and $services) {
        $useForce = ($defaultMode -ne 'keep')
        $cmd    = "Enable-ExchangeCertificate -Thumbprint $activeThumbprint -Services $services$(if ($useForce) { ' -Force' })"
        $impact = switch ($defaultMode) {
            'target' { 'The target becomes the internal transport (default SMTP) certificate (-Force answers the overwrite question with Yes).' }
            'keep'   { 'The internal transport certificate stays as it is. Exchange will ask "Overwrite the existing default SMTP certificate?" - answer N. SMTP and the NETWORK SERVICE key ACL are applied either way.' }
            default  { 'The target briefly becomes the default; the next step moves the default to the self-signed certificate.' }
        }
        if ($services -match 'IIS') { $impact += ' The IIS binding for the Default Web Site will be replaced.' }

        if (Confirm-Write -Action "Enable services '$services' on the target certificate" -Detail $cmd -Impact $impact) {
            try {
                if ($useForce) { Enable-ExchangeCertificate -Thumbprint $activeThumbprint -Services $services -Force -ErrorAction Stop }
                else           { Enable-ExchangeCertificate -Thumbprint $activeThumbprint -Services $services -ErrorAction Stop }
            }
            catch {
                # Measured: without -Force the overwrite question comes AFTER SMTP has been
                # applied. In a non-interactive host it throws; the result still counts.
                Write-Warn "Enable-ExchangeCertificate reported: $($_.Exception.Message)"
            }
            $check = Get-ExchangeCertificate -Thumbprint $activeThumbprint -ErrorAction SilentlyContinue
            if ($check -and $check.Services -match 'SMTP') {
                Write-Ok "Services now on $activeThumbprint : $($check.Services)"
                Add-ChangeLogEntry "Enabled services '$services' on $activeThumbprint (default mode: $defaultMode)"
                $restartRequired = $true
            }
            else {
                Write-Problem "SMTP is not set on $activeThumbprint after Enable-ExchangeCertificate."
            }
        }
    }
    elseif ($defaultMode -eq 'target') {
        # SMTP already present - only move the default.
        if (Confirm-Write -Action 'Make the target the internal transport certificate' `
                -Detail "Enable-ExchangeCertificate -Thumbprint $activeThumbprint -Services SMTP -Force" `
                -Impact 'The target becomes the internal transport (default SMTP) certificate.') {
            try {
                Enable-ExchangeCertificate -Thumbprint $activeThumbprint -Services SMTP -Force -ErrorAction Stop
                Add-ChangeLogEntry "Made $activeThumbprint the internal transport certificate"
                $restartRequired = $true
            }
            catch { Write-Problem "Enable-ExchangeCertificate failed: $($_.Exception.Message)" }
        }
    }

    # Step B: self-signed certificate as default (existing or new).
    if ($defaultMode -eq 'newselfsigned') {
        if (Confirm-Write -Action 'Create a new self-signed Exchange certificate' `
                -Detail "New-ExchangeCertificate -SubjectName CN=$env:COMPUTERNAME -DomainName <server>,<fqdn> -Services None -PrivateKeyExportable `$true" `
                -Impact 'Adds a certificate valid for 5 years. -Services None: it takes over nothing on its own (no IIS front end binding).') {
            try {
                $selfSigned = New-SelfSignedExchangeCertificate
                Write-Ok "Created self-signed certificate $($selfSigned.Thumbprint) (until $($selfSigned.NotAfter.ToString('yyyy-MM-dd')))."
                Add-ChangeLogEntry "Created self-signed certificate $($selfSigned.Thumbprint)"
            }
            catch {
                Write-Problem "New-ExchangeCertificate failed: $($_.Exception.Message)"
                $selfSigned = $null
            }
        }
        else { $selfSigned = $null }
    }

    if ($defaultMode -in @('selfsigned', 'newselfsigned') -and $selfSigned) {
        if (Confirm-Write -Action 'Make the self-signed certificate the internal transport certificate' `
                -Detail "Enable-ExchangeCertificate -Thumbprint $($selfSigned.Thumbprint) -Services SMTP -Force" `
                -Impact 'TLS between Exchange servers uses the self-signed certificate again. Connectors with TlsCertificateName (hybrid) are not affected.') {
            try {
                Enable-ExchangeCertificate -Thumbprint $selfSigned.Thumbprint -Services SMTP -Force -ErrorAction Stop
            }
            catch { Write-Problem "Enable-ExchangeCertificate failed: $($_.Exception.Message)" }
            $nowDefault = Get-InternalTransportCertificate
            if ($nowDefault -and $nowDefault.Thumbprint -eq $selfSigned.Thumbprint) {
                Write-Ok "Internal transport certificate is now $($selfSigned.Thumbprint) (read back from AD)."
                Add-ChangeLogEntry "Made self-signed $($selfSigned.Thumbprint) the internal transport certificate"
                $restartRequired = $true
            }
            else {
                Write-Problem "AD still shows $(if ($nowDefault) { $nowDefault.Thumbprint } else { '<none>' }) as internal transport certificate."
            }
        }
    }
}

# --- Remediation 4b: Exchange Back End binding (port 444) -------------------
if ($issues -contains 'BackendCertBroken') {
    if (-not $selfSigned) { $selfSigned = Get-SelfSignedCandidate }
    if (-not $selfSigned) {
        Write-Warn 'No valid self-signed Exchange certificate to bind to the Back End site.'
        Write-Info 'Run again with -RestoreSelfSignedDefault to create one, then repeat.'
    }
    else {
        foreach ($b in @(Get-BackendBinding | Where-Object { -not $_.InStore -or $_.NotAfter -lt (Get-Date) })) {
            $appId = if ($b.AppId) { $b.AppId } else { '{4dc3e181-e14b-4a21-b022-59fc669b0914}' }
            $cmd   = "netsh http update sslcert ipport=$($b.IpPort) certhash=$($selfSigned.Thumbprint) appid=$appId certstorename=MY"
            if (Confirm-Write -Action "Bind the Exchange Back End site ($($b.IpPort)) to the self-signed certificate" -Detail $cmd `
                    -Impact 'Front end to back end proxying (OWA, ECP, EWS, ...) uses this certificate. Only port 444 is touched.') {
                $out = & netsh http update sslcert "ipport=$($b.IpPort)" "certhash=$($selfSigned.Thumbprint)" "appid=$appId" 'certstorename=MY' 2>&1 | Out-String
                $nb  = @(Get-BackendBinding | Where-Object { $_.IpPort -eq $b.IpPort })
                if ($nb.Count -gt 0 -and $nb[0].Thumbprint -eq $selfSigned.Thumbprint) {
                    Write-Ok "Back End $($b.IpPort) now uses $($selfSigned.Thumbprint)."
                    Add-ChangeLogEntry "Bound Exchange Back End $($b.IpPort) to $($selfSigned.Thumbprint)"
                    Write-Info 'Recommended afterwards: iisreset /noforce'
                }
                else {
                    Write-Problem "Back End binding unchanged. netsh said: $($out.Trim())"
                }
            }
        }
    }
}

# --- Remediation 5: Receive Connector TlsCertificateName ---------------------
$activeCert = Get-ExchangeCertificate -Thumbprint $activeThumbprint -ErrorAction SilentlyContinue

if ($activeCert) {
    $newTlsName = Get-TlsCertificateName -Certificate $activeCert

    $connectorsToFix = @(
        Get-ReceiveConnector -ErrorAction SilentlyContinue |
            Where-Object {
                $_.TlsCertificateName -and
                $_.TlsCertificateName.ToString() -ne $newTlsName
            }
    )

    foreach ($rc in $connectorsToFix) {

        $ok = Confirm-Write `
            -Action 'Update TlsCertificateName on a Receive Connector' `
            -Detail ("Connector: {0}`n                 old: {1}`n                 new: {2}" -f $rc.Identity, $rc.TlsCertificateName, $newTlsName) `
            -Impact 'Binds this connector explicitly to the target certificate. A hard-coded stale binding here overrides the default SMTP certificate.'

        if ($ok) {
            try {
                Set-ReceiveConnector -Identity $rc.Identity `
                                     -TlsCertificateName $newTlsName `
                                     -ErrorAction Stop
                Write-Ok "TlsCertificateName updated on '$($rc.Identity)'."
                Add-ChangeLogEntry "Set TlsCertificateName on Receive Connector '$($rc.Identity)'"
                $restartRequired = $true
            }
            catch {
                Write-Problem "Failed to update Receive Connector: $($_.Exception.Message)"
            }
        }
    }
}

# --- Remediation 6: transport service restart --------------------------------
if ($restartRequired) {
    Write-Host ''
    Write-Warn 'Transport services still hold a handle on the previous certificate.'
    Write-Warn 'Without a restart, the new certificate will not be presented on port 25.'

    $ok = Confirm-Write `
        -Action 'Restart Exchange transport services' `
        -Detail 'MSExchangeTransport, MSExchangeFrontEndTransport' `
        -Impact 'SHORT MAIL FLOW INTERRUPTION (typically 10-60 seconds). Inbound SMTP connections will be refused during the restart.'

    if ($ok) {
        foreach ($svc in @('MSExchangeTransport', 'MSExchangeFrontEndTransport')) {
            $service = Get-Service -Name $svc -ErrorAction SilentlyContinue
            if (-not $service) {
                Write-Info "Service '$svc' not present on this server - skipped."
                continue
            }
            try {
                Write-Info "Restarting $svc ..."
                Restart-Service -Name $svc -Force -ErrorAction Stop
                Write-Ok "$svc restarted."
                Add-ChangeLogEntry "Restarted service $svc"
            }
            catch {
                Write-Problem "Failed to restart ${svc}: $($_.Exception.Message)"
            }
        }

        Write-Info 'Waiting 10 seconds for transport to initialize...'
        Start-Sleep -Seconds 10
    }
    else {
        Write-Warn 'Transport NOT restarted. The change will not take effect until you restart it manually:'
        Write-Host '        Restart-Service MSExchangeTransport' -ForegroundColor Gray
        Write-Host '        Restart-Service MSExchangeFrontEndTransport' -ForegroundColor Gray
    }
}

#endregion

#region ---------------------------------------------------------- AFTER snapshot

$after = Get-SmtpCertificateSnapshot
Show-SmtpCertificateSnapshot -Snapshot $after -Label 'AFTER'

Write-Section 'Change log'
if ($script:ChangeLog.Count -eq 0) {
    Write-Info 'No changes were applied.'
}
else {
    foreach ($entry in $script:ChangeLog) {
        Write-Host "  * $entry" -ForegroundColor Green
    }
}

Write-Section 'Comparison'

$beforeThumbs = @($before.SmtpCertificates | ForEach-Object { $_.Thumbprint })
$afterThumbs  = @($after.SmtpCertificates  | ForEach-Object { $_.Thumbprint })

$added   = @($afterThumbs  | Where-Object { $beforeThumbs -notcontains $_ })
$removed = @($beforeThumbs | Where-Object { $afterThumbs  -notcontains $_ })

if ($added.Count -eq 0 -and $removed.Count -eq 0) {
    Write-Info 'SMTP certificate assignment unchanged.'
}
else {
    foreach ($t in $added)   { Write-Host "  + SMTP flag ADDED   : $t" -ForegroundColor Green }
    foreach ($t in $removed) { Write-Host "  - SMTP flag REMOVED : $t" -ForegroundColor Yellow }
}

Write-Section 'Verification'
Write-Info 'Verify the certificate actually presented on port 25 from a remote host:'
Write-Host '  openssl s_client -connect mail.example.com:25 -starttls smtp -servername mail.example.com' -ForegroundColor Gray
Write-Host ''
Write-Info 'Or check the Exchange transport log for TLS negotiation:'
Write-Host '  Get-Content "$env:ExchangeInstallPath\TransportRoles\Logs\FrontEnd\ProtocolLog\SmtpReceive\*.log" -Tail 50' -ForegroundColor Gray
Write-Host ''

#endregion
