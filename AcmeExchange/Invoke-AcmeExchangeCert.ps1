#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Let's Encrypt (ACME) certificate lifecycle for Exchange Server on-premises,
    using Azure DNS (DNS-01) via Posh-ACME. One script, two roles:
    an interactive one-time setup wizard and an unattended daily renewal job.

.DESCRIPTION
    -Setup        Interactive wizard. Bootstraps Azure (app registration with
                  certificate credential, least-privilege role on the DNS zone),
                  collects Exchange and notification settings, requests the
                  first certificate, installs it and registers the scheduled task.
                  Re-runnable: existing resources are detected and reused.

    -Renew        Default. Calls Submit-Renewal; if a new certificate is issued it
                  is installed on every configured Exchange server (Import-/
                  Enable-ExchangeCertificate), verified with real TLS handshakes,
                  connector TlsCertificateName values are updated and the old
                  certificate is removed. Sends success/warning/error mail.

    -ImportPfx    Install a PFX produced elsewhere (no ACME, no DNS, no Azure) through the exact
                  same deployment pipeline as -Renew: import, enable, loopback binding, iisreset,
                  TLS verification, connector update, removal of the superseded certificate.
                  Takes -PfxPath, or the newest *.pfx in the drop folder (<Home>\pfx).
                  Unattended: the password comes from -SetPfxPassword, never from a prompt.
                  Idempotent - a certificate already presented by every server is not reinstalled.

    -SetPfxPassword  Store the PFX password once, DPAPI-encrypted for this machine, so the SYSTEM
                  task can read it. Interactive prompt, or -PfxPassword <securestring>.

    -ForceRenew   Force a new certificate even if not due.
    -Status       Show bound certificates, days to expiry, connectors, order state.
    -TestMail     Send a test notification.
    -InstallTask  (Re-)register the scheduled task (SYSTEM, daily).
    -Staging      Use the Let's Encrypt staging environment (with -Setup/-ForceRenew).

.NOTES
    Only dependency: Posh-ACME (https://github.com/rmbolger/Posh-ACME).
    Azure/Graph calls are plain REST. Az.Accounts is used only as an optional
    login fallback during -Setup.

    Design notes and the reasoning behind Import-/Enable-ExchangeCertificate
    are documented in README.md.

    License: MIT
#>
[CmdletBinding(DefaultParameterSetName = 'Renew')]
param(
    [Parameter(ParameterSetName = 'Setup')]       [switch]$Setup,
    [Parameter(ParameterSetName = 'Renew')]       [switch]$Renew,
    [Parameter(ParameterSetName = 'Renew')]       [switch]$ForceRenew,
    [Parameter(ParameterSetName = 'Status')]      [switch]$Status,
    [Parameter(ParameterSetName = 'TestMail')]    [switch]$TestMail,
    [Parameter(ParameterSetName = 'InstallTask')] [switch]$InstallTask,
    [Parameter(ParameterSetName = 'RemoveTask')]  [switch]$RemoveTask,
    [Parameter(ParameterSetName = 'Teardown')]    [switch]$Teardown,
    [Parameter(ParameterSetName = 'ImportPfx')]   [switch]$ImportPfx,
    [Parameter(ParameterSetName = 'ImportPfx')]   [string]$PfxPath,
    [Parameter(ParameterSetName = 'SetPfxPassword', Mandatory = $true)] [switch]$SetPfxPassword,
    [Parameter(ParameterSetName = 'SetPfxPassword')] [securestring]$PfxPassword,
    [switch]$Staging,
    [switch]$NonInteractive,
    [switch]$SkipInstall,
    [switch]$Yes,
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config.json')
)

Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$script:ScriptVersion   = '1.0.0'
$script:EventSource     = 'AcmeExchangeCert'
$script:LogFile         = $null
$script:Config          = $null
$script:MailPrefix      = '[ACME Exchange]'
$script:UsedFallback    = @()
$script:ExchangeLoaded  = $false
$script:BundleRoot      = $null   # set by the GUI/exe launcher; else the script's own folder is used
$script:NonInteractive  = [bool]$NonInteractive   # when set, Read-* helpers return their defaults (GUI-driven)
$script:SkipInstall     = [bool]$SkipInstall      # when set, -Setup issues the cert but does not install it on any server
$script:AuthKeyChanged  = $false                  # set when a new app auth key was just registered (needs Entra propagation)

# First-party public client that can obtain tokens for ARM and Graph via device code.
# (Azure PowerShell). Override with -Setup prompt if Microsoft restricts it in future.
$script:BootstrapClientId = '1950a258-227b-4e31-a9cf-717495945fc2'
$script:ArmBase   = 'https://management.azure.com'
$script:GraphBase = 'https://graph.microsoft.com/v1.0'

#region ---------------------------------------------------------------- Logging
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR', 'STEP')][string]$Level = 'INFO'
    )
    $ts   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "$ts [$Level] $Message"
    $color = switch ($Level) { 'OK' { 'Green' } 'WARN' { 'Yellow' } 'ERROR' { 'Red' } 'STEP' { 'Cyan' } default { 'Gray' } }
    Write-Host $line -ForegroundColor $color
    if ($script:LogFile) {
        try { Add-Content -Path $script:LogFile -Value $line -Encoding UTF8 } catch { }
    }
}

function Write-EventLogEntry {
    param([string]$Message, [ValidateSet('Information', 'Warning', 'Error')][string]$Type = 'Information', [int]$Id = 9000)
    try {
        if (-not [Diagnostics.EventLog]::SourceExists($script:EventSource)) {
            New-EventLog -LogName Application -Source $script:EventSource
        }
        Write-EventLog -LogName Application -Source $script:EventSource -EntryType $Type -EventId $Id -Message $Message
    } catch {
        Write-Log "Event log write failed: $($_.Exception.Message)" WARN
    }
}

function Initialize-Logging {
    param([string]$HomePath)
    $logDir = Join-Path $HomePath 'logs'
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $script:LogFile = Join-Path $logDir ("acme-{0}.log" -f (Get-Date -Format 'yyyyMM'))
    # keep 12 months
    Get-ChildItem $logDir -Filter 'acme-*.log' | Where-Object { $_.LastWriteTime -lt (Get-Date).AddMonths(-12) } |
        Remove-Item -Force -ErrorAction SilentlyContinue
}
#endregion

#region ---------------------------------------------------------------- Config
function Get-DefaultConfig {
    [ordered]@{
        Version  = 1
        Acme     = [ordered]@{
            Server          = 'LE_PROD'
            Domains         = @()
            Contact         = ''
            KeyLength       = '2048'
            DnsProvider     = 'Azure'   # Azure | Cloudflare | Route53 | GoDaddy | DigitalOcean
            DnsPluginArgs   = @{}       # NON-secret plugin args only (e.g. Route53 access key); secrets live encrypted in the Posh-ACME order
        }
        Azure    = [ordered]@{
            TenantId           = ''
            SubscriptionId     = ''
            AppId              = ''
            AppObjectId        = ''
            SpObjectId         = ''
            AppDisplayName     = ''
            AuthCertThumbprint = ''
            DnsZoneIds         = @()
            RoleName           = 'DNS TXT Contributor'
            AuthCertWarnDays   = 60      # warn by mail this many days before the app auth cert expires (re-run -Setup to rotate)
            BootstrapClientId  = $script:BootstrapClientId
        }
        Exchange = [ordered]@{
            Servers               = @()
            Services              = 'IIS,SMTP'
            PrivateKeyExportable  = $true
            UpdateConnectors      = $true
            RemoveOldCertificate  = $true
            LoopbackBinding       = $true
            RestartTransportIfStale = $true
        }
        Notify   = [ordered]@{
            SmtpServer           = 'localhost'
            Port                 = 25
            UseSsl               = $false
            From                 = ''
            To                   = @()
            WarnDaysBeforeExpiry = 14
            SendSuccessMail      = $true
        }
        Pfx      = [ordered]@{
            # Alternative to ACME: a PFX produced elsewhere is dropped into DropFolder and installed
            # by the same pipeline. Enabled is set by -ImportPfx; the password lives DPAPI-encrypted
            # in PasswordFile (machine scope), never in this file.
            Enabled            = $false
            DropFolder         = ''      # default: <Home>\pfx
            PasswordFile       = ''      # default: <Home>\pfxpass.dat
            ArchiveAfterImport = $true   # move the PFX to <DropFolder>\archive after a successful install
            Subject            = ''      # remembered from the last import so -Status works without the file
        }
        Paths    = [ordered]@{
            Home          = 'C:\Tools\AcmeExchange'
            BackupKeep    = 3
        }
        Task     = [ordered]@{
            Name = 'ACME Exchange Certificate Renewal'
            Time = '03:00'
        }
    }
}

function Get-DnsProviders {
    <#
      Supported DNS-01 providers. 'Azure' is special (its own app/cert/role bootstrap in -Setup);
      all others just need API credentials, which are passed once and then kept encrypted in the
      Posh-ACME order. Each field maps to the exact Posh-ACME plugin parameter name; Secure fields
      are passed as SecureString and never written to config.json. Extend by adding a table entry -
      Posh-ACME ships ~100 plugins (Get-PAPlugin) so most providers are a one-line addition here.
    #>
    [ordered]@{
        'Azure'        = @{ Plugin = 'Azure';      Bootstrap = $true;  Fields = @() }
        'Cloudflare'   = @{ Plugin = 'Cloudflare'; Bootstrap = $false; Fields = @(
                                @{ Key = 'CFToken';       Label = 'Cloudflare API Token';    Secure = $true }) }
        'Route53'      = @{ Plugin = 'Route53';    Bootstrap = $false; Fields = @(
                                @{ Key = 'R53AccessKey';  Label = 'AWS Access Key ID';       Secure = $false },
                                @{ Key = 'R53SecretKey';  Label = 'AWS Secret Access Key';   Secure = $true }) }
        'GoDaddy'      = @{ Plugin = 'GoDaddy';    Bootstrap = $false; Fields = @(
                                @{ Key = 'GDKey';         Label = 'GoDaddy API Key';         Secure = $false },
                                @{ Key = 'GDSecretSecure';Label = 'GoDaddy API Secret';      Secure = $true }) }
        'DigitalOcean' = @{ Plugin = 'DOcean';     Bootstrap = $false; Fields = @(
                                @{ Key = 'DOTokenSecure'; Label = 'DigitalOcean API Token';  Secure = $true }) }
    }
}

function Read-Config {
    if (-not (Test-Path $ConfigPath)) { return $null }
    $raw = Get-Content $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
    # ConvertTo-Json collapses single-element arrays to scalars, so list fields can come back as a
    # plain string. Force the known list fields back to arrays so .Count and [0] behave everywhere.
    if ($raw) {
        if ($raw.PSObject.Properties['Acme']     -and $raw.Acme.PSObject.Properties['Domains'])   { $raw.Acme.Domains       = @($raw.Acme.Domains) }
        if ($raw.PSObject.Properties['Exchange'] -and $raw.Exchange.PSObject.Properties['Servers']){ $raw.Exchange.Servers   = @($raw.Exchange.Servers) }
        if ($raw.PSObject.Properties['Notify']   -and $raw.Notify.PSObject.Properties['To'])       { $raw.Notify.To          = @($raw.Notify.To) }
        if ($raw.PSObject.Properties['Azure']    -and $raw.Azure.PSObject.Properties['DnsZoneIds']){ $raw.Azure.DnsZoneIds   = @($raw.Azure.DnsZoneIds) }
    }
    return $raw
}

function Save-Config {
    param($Cfg)
    $json = $Cfg | ConvertTo-Json -Depth 6
    [IO.File]::WriteAllText($ConfigPath, $json, [Text.UTF8Encoding]::new($false))
    Write-Log "Configuration saved to $ConfigPath" OK
}

function Get-PoshAcmeHome { Join-Path $script:Config.Paths.Home 'Posh-ACME' }
#endregion

#region ---------------------------------------------------------------- Console helpers
function Read-Default {
    param([string]$Prompt, [string]$Default = '')
    if ($script:NonInteractive) { return $Default }
    if ($Default) { $in = Read-Host "$Prompt [$Default]" } else { $in = Read-Host $Prompt }
    if ([string]::IsNullOrWhiteSpace($in)) { return $Default }
    return $in.Trim()
}

function Read-YesNo {
    param([string]$Prompt, [bool]$Default = $true)
    if ($script:NonInteractive) { return $Default }
    $d = if ($Default) { 'Y/n' } else { 'y/N' }
    $in = Read-Host "$Prompt [$d]"
    if ([string]::IsNullOrWhiteSpace($in)) { return $Default }
    return $in.Trim() -match '^(y|yes|j|ja)$'
}

function Read-List {
    param([string]$Prompt, [string[]]$Default = @())
    $d = ($Default -join ', ')
    $in = Read-Default "$Prompt (comma separated)" $d
    return @($in -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Show-Banner {
    param([string]$Text)
    Write-Host ''
    Write-Host ('=' * 78) -ForegroundColor Cyan
    Write-Host "  $Text" -ForegroundColor Cyan
    Write-Host ('=' * 78) -ForegroundColor Cyan
}
#endregion

#region ---------------------------------------------------------------- Auth (device code / Az fallback)
function ConvertFrom-Jwt {
    param([string]$Token)
    $payload = $Token.Split('.')[1].Replace('-', '+').Replace('_', '/')
    switch ($payload.Length % 4) { 2 { $payload += '==' } 3 { $payload += '=' } }
    return ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload)) | ConvertFrom-Json)
}

function Get-DeviceCodeToken {
    param([string]$ClientId, [string]$Tenant, [string]$Scope)
    $dc = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/devicecode" `
        -Body @{ client_id = $ClientId; scope = $Scope }
    Write-Host ''
    Write-Host $dc.message -ForegroundColor Yellow
    Write-Host ''
    # convenience: copy the user code to the clipboard and open the sign-in page
    $copied = $false
    try { Set-Clipboard -Value $dc.user_code -ErrorAction Stop; $copied = $true }
    catch { try { $dc.user_code | clip.exe; $copied = $true } catch { } }
    Write-Host ("  Code : {0}   {1}" -f $dc.user_code, $(if ($copied) { '(in die Zwischenablage kopiert / copied to clipboard)' } else { '' })) -ForegroundColor Cyan
    Write-Host ("  URL  : {0}" -f $dc.verification_uri) -ForegroundColor Cyan
    try { Start-Process $dc.verification_uri | Out-Null; Write-Host '  (Browser wird geoeffnet / opening browser ...)' -ForegroundColor DarkGray }
    catch { Write-Host '  (Browser konnte nicht automatisch geoeffnet werden - URL bitte manuell aufrufen)' -ForegroundColor DarkGray }
    Write-Host ''
    $deadline = (Get-Date).AddSeconds($dc.expires_in)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds $dc.interval
        try {
            $tok = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/token" `
                -Body @{ grant_type = 'urn:ietf:params:oauth:grant-type:device_code'; client_id = $ClientId; device_code = $dc.device_code }
            return $tok
        } catch {
            $err = $null
            try { $err = ($_.ErrorDetails.Message | ConvertFrom-Json).error } catch { }
            if ($err -eq 'authorization_pending' -or $err -eq 'slow_down') { continue }
            throw "Device code login failed: $err $($_.Exception.Message)"
        }
    }
    throw 'Device code expired before sign-in completed.'
}

function Get-TokenForResource {
    param([string]$ClientId, [string]$Tenant, [string]$RefreshToken, [string]$Scope)
    $tok = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/token" `
        -Body @{ grant_type = 'refresh_token'; client_id = $ClientId; refresh_token = $RefreshToken; scope = $Scope }
    return $tok.access_token
}

function Connect-Bootstrap {
    <# Returns @{ Arm = <token>; Graph = <token>; TenantId; Upn } #>
    param([string]$ClientId, [string]$TenantHint = 'organizations')
    Write-Log 'Signing in (device code). Use a Global Administrator who is also Owner of the subscription.' STEP
    try {
        $first = Get-DeviceCodeToken -ClientId $ClientId -Tenant $TenantHint -Scope "$script:ArmBase/.default offline_access openid profile"
        $claims = ConvertFrom-Jwt $first.access_token
        $tenant = $claims.tid
        $graph  = Get-TokenForResource -ClientId $ClientId -Tenant $tenant -RefreshToken $first.refresh_token -Scope 'https://graph.microsoft.com/.default'
        $upn = if ($claims.PSObject.Properties['upn']) { $claims.upn } elseif ($claims.PSObject.Properties['unique_name']) { $claims.unique_name } else { '(unknown)' }
        Write-Log "Signed in as $upn, tenant $tenant" OK
        return @{ Arm = $first.access_token; Graph = $graph; TenantId = $tenant; Upn = $upn }
    } catch {
        Write-Log "First-party device code login failed: $($_.Exception.Message)" WARN
        Write-Log 'Falling back to Az.Accounts (Connect-AzAccount -UseDeviceAuthentication).' WARN
        if (-not (Get-Module Az.Accounts -ListAvailable)) {
            if (Read-YesNo 'Az.Accounts is not installed. Install it now (CurrentUser scope)?' $true) {
                Install-Module Az.Accounts -Scope CurrentUser -Force -AllowClobber
            } else { throw 'No usable login method.' }
        }
        Import-Module Az.Accounts
        $azParams = @{ UseDeviceAuthentication = $true }
        if ($TenantHint -ne 'organizations') { $azParams.Tenant = $TenantHint }
        Connect-AzAccount @azParams | Out-Null
        $ctx = Get-AzContext
        $arm   = Get-AzAccessToken -ResourceUrl "$script:ArmBase/"
        $graph = Get-AzAccessToken -ResourceUrl 'https://graph.microsoft.com/'
        $armT = if ($arm.Token -is [securestring]) { [Net.NetworkCredential]::new('', $arm.Token).Password } else { $arm.Token }
        $grT  = if ($graph.Token -is [securestring]) { [Net.NetworkCredential]::new('', $graph.Token).Password } else { $graph.Token }
        Write-Log "Signed in via Az as $($ctx.Account.Id), tenant $($ctx.Tenant.Id)" OK
        return @{ Arm = $armT; Graph = $grT; TenantId = $ctx.Tenant.Id; Upn = $ctx.Account.Id }
    }
}
#endregion

#region ---------------------------------------------------------------- REST helpers
function Invoke-Rest {
    param([string]$Method, [string]$Uri, [string]$Token, $Body = $null, [int]$Retries = 0)
    $headers = @{ Authorization = "Bearer $Token"; 'Content-Type' = 'application/json' }
    $attempt = 0
    while ($true) {
        try {
            if ($null -ne $Body) {
                $json = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 10 }
                return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -Body ([Text.Encoding]::UTF8.GetBytes($json))
            }
            return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers
        } catch {
            $detail = ''
            try { $detail = $_.ErrorDetails.Message } catch { }
            if ($attempt -lt $Retries) {
                $attempt++
                Write-Log "REST $Method $Uri failed (attempt $attempt/$Retries): $detail - retrying in 10s" WARN
                Start-Sleep -Seconds 10
                continue
            }
            throw "REST $Method $Uri failed: $($_.Exception.Message) $detail"
        }
    }
}

function Invoke-Arm   { param([string]$Method, [string]$Path, [string]$Token, $Body = $null, [string]$ApiVersion, [int]$Retries = 0)
    $sep = if ($Path.Contains('?')) { '&' } else { '?' }
    Invoke-Rest -Method $Method -Uri "$script:ArmBase$Path${sep}api-version=$ApiVersion" -Token $Token -Body $Body -Retries $Retries
}
function Invoke-Graph { param([string]$Method, [string]$Path, [string]$Token, $Body = $null, [int]$Retries = 0)
    Invoke-Rest -Method $Method -Uri "$script:GraphBase$Path" -Token $Token -Body $Body -Retries $Retries
}
#endregion

#region ---------------------------------------------------------------- Exchange helpers
function Connect-ExchangeManagement {
    if ($script:ExchangeLoaded) { return }
    if (Get-Command Get-ExchangeCertificate -ErrorAction SilentlyContinue) { $script:ExchangeLoaded = $true; return }
    try {
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
        $script:ExchangeLoaded = $true
        Write-Log 'Exchange management snap-in loaded.' INFO
        return
    } catch {
        Write-Log "Snap-in not available ($($_.Exception.Message)); trying RemoteExchange.ps1" WARN
    }
    $rx = Join-Path $env:ExchangeInstallPath 'bin\RemoteExchange.ps1'
    if (Test-Path $rx) {
        . $rx
        Connect-ExchangeServer -auto -ClientApplication:ManagementShell | Out-Null
        $script:ExchangeLoaded = $true
        return
    }
    throw 'Exchange management tools not found on this machine. Run the script on an Exchange server.'
}

function Get-LocalServerName { $env:COMPUTERNAME.ToUpper() }

function Resolve-ExchangeServer {
    param([string]$Name)
    $srv = Get-ExchangeServer -Identity $Name
    [pscustomobject]@{
        Name    = $srv.Name.ToUpper()
        Fqdn    = $srv.Fqdn
        IsLocal = ($srv.Name.ToUpper() -eq (Get-LocalServerName))
    }
}

function Test-WinRm {
    param([string]$Fqdn)
    try { Test-WSMan -ComputerName $Fqdn -ErrorAction Stop | Out-Null; return $true } catch { return $false }
}

function Invoke-OnServer {
    <# Runs a script block locally or via WinRM. Returns $null and logs a warning if WinRM is unavailable. #>
    param([pscustomobject]$Server, [scriptblock]$ScriptBlock, [object[]]$ArgumentList = @())
    if ($Server.IsLocal) { return (& $ScriptBlock @ArgumentList) }
    if (-not (Test-WinRm $Server.Fqdn)) {
        Write-Log "WinRM not reachable on $($Server.Fqdn); skipping remote step." WARN
        return $null
    }
    return Invoke-Command -ComputerName $Server.Fqdn -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList
}
#endregion

#region ---------------------------------------------------------------- TLS probes
function Test-TlsEndpoint {
    <# Returns thumbprint of the certificate presented on host:port, or $null. Hard timeouts on
       connect and the TLS handshake - a stalled port (e.g. mid-iisreset) must not block for minutes. #>
    param([string]$HostName, [int]$Port, [string]$Sni = $HostName, [int]$TimeoutMs = 8000)
    $tcp = $null; $ssl = $null
    try {
        $tcp = New-Object Net.Sockets.TcpClient
        $iar = $tcp.BeginConnect($HostName, $Port, $null, $null)
        if (-not $iar.AsyncWaitHandle.WaitOne($TimeoutMs)) { throw 'connect timeout' }
        $tcp.EndConnect($iar)
        $ns = $tcp.GetStream(); $ns.ReadTimeout = $TimeoutMs; $ns.WriteTimeout = $TimeoutMs
        $ssl = New-Object Net.Security.SslStream($ns, $false, ({ $true }))
        $ssl.AuthenticateAsClient($Sni)
        return $ssl.RemoteCertificate.GetCertHashString()
    } catch {
        Write-Log "TLS probe ${HostName}:${Port} failed: $($_.Exception.Message)" WARN
        return $null
    } finally {
        if ($ssl) { $ssl.Dispose() }; if ($tcp) { $tcp.Dispose() }
    }
}

function Test-StartTls {
    <# SMTP EHLO + STARTTLS; returns thumbprint or $null. Hard timeouts throughout. #>
    param([string]$HostName, [int]$Port = 25, [string]$Sni = $HostName, [int]$TimeoutMs = 8000)
    $tcp = $null; $ssl = $null
    try {
        $tcp = New-Object Net.Sockets.TcpClient
        $iar = $tcp.BeginConnect($HostName, $Port, $null, $null)
        if (-not $iar.AsyncWaitHandle.WaitOne($TimeoutMs)) { throw 'connect timeout' }
        $tcp.EndConnect($iar)
        $ns = $tcp.GetStream(); $ns.ReadTimeout = $TimeoutMs; $ns.WriteTimeout = $TimeoutMs
        $sr = New-Object IO.StreamReader($ns); $sw = New-Object IO.StreamWriter($ns); $sw.AutoFlush = $true
        $null = $sr.ReadLine()
        $sw.WriteLine("EHLO acme-probe.local")
        do { $l = $sr.ReadLine() } while ($l -match '^250-')
        $sw.WriteLine('STARTTLS')
        $resp = $sr.ReadLine()
        if ($resp -notmatch '^220') { throw "STARTTLS refused: $resp" }
        $ssl = New-Object Net.Security.SslStream($ns, $false, ({ $true }))
        $ssl.AuthenticateAsClient($Sni)
        return $ssl.RemoteCertificate.GetCertHashString()
    } catch {
        Write-Log "STARTTLS probe ${HostName}:${Port} failed: $($_.Exception.Message)" WARN
        return $null
    } finally {
        if ($ssl) { $ssl.Dispose() }; if ($tcp) { $tcp.Dispose() }
    }
}

function Wait-ForThumbprint {
    param([scriptblock]$Probe, [string]$Expected, [int]$Attempts = 6, [int]$DelaySeconds = 10)
    for ($i = 1; $i -le $Attempts; $i++) {
        $tp = & $Probe
        if ($tp -and $tp.ToUpper() -eq $Expected.ToUpper()) { return $true }
        if ($i -lt $Attempts) { Start-Sleep -Seconds $DelaySeconds }
    }
    return $false
}
#endregion

#region ---------------------------------------------------------------- Notification
function Send-Notification {
    param([string]$Subject, [string]$Body, [ValidateSet('Info', 'Warning', 'Error')][string]$Level = 'Info')
    $n = $script:Config.Notify
    $recipients = @($n.To | Where-Object { $_ })
    if (-not $recipients.Count -or -not $n.From) { Write-Log 'Notification not configured; skipping mail.' WARN; return }
    try {
        $msg = New-Object Net.Mail.MailMessage
        $msg.From = $n.From
        foreach ($t in $recipients) { $msg.To.Add($t) }
        $msg.Subject = "$script:MailPrefix $Subject"
        $msg.Body = $Body + "`r`n`r`n-- `r`nInvoke-AcmeExchangeCert v$script:ScriptVersion on $(Get-LocalServerName)`r`nLog: $script:LogFile"
        $msg.IsBodyHtml = $false
        $client = New-Object Net.Mail.SmtpClient($n.SmtpServer, [int]$n.Port)
        $client.EnableSsl = [bool]$n.UseSsl
        $client.Send($msg)
        Write-Log "Notification mail sent to $($recipients -join ', ') ($Level): $Subject" OK
    } catch {
        Write-Log "Sending notification failed: $($_.Exception.Message)" ERROR
    }
}
#endregion

#region ---------------------------------------------------------------- Posh-ACME runtime
function Get-BundleRoot {
    # Directory the bundle lives in: explicit override (set by the exe launcher), else the
    # script's own folder, else the current directory.
    if ($script:BundleRoot) { return $script:BundleRoot }
    if ($PSScriptRoot)      { return $PSScriptRoot }
    return (Get-Location).Path
}

function Get-BundledPoshAcme {
    # Path to a Posh-ACME manifest shipped alongside the script (lib\Posh-ACME\Posh-ACME.psd1),
    # so no module needs to be installed for any user. Returns $null when not bundled.
    $manifest = Join-Path (Get-BundleRoot) 'lib\Posh-ACME\Posh-ACME.psd1'
    if (Test-Path $manifest) { return $manifest }
    return $null
}

function Initialize-PoshAcme {
    param([switch]$Quiet)
    $acmeHome = Get-PoshAcmeHome
    if (-not (Test-Path $acmeHome)) { New-Item -ItemType Directory -Path $acmeHome -Force | Out-Null }
    $env:POSHACME_HOME = $acmeHome
    if (Get-Module Posh-ACME) { Remove-Module Posh-ACME -Force }
    $bundled = Get-BundledPoshAcme
    if ($bundled) {
        Import-Module $bundled -ErrorAction Stop
        if (-not $Quiet) { Write-Log "Posh-ACME imported from bundle: $bundled" INFO }
    } else {
        Import-Module Posh-ACME -ErrorAction Stop
    }
    $server = $script:Config.Acme.Server
    if ($Staging) { $server = 'LE_STAGE' }
    $cur = $null
    try { $cur = Get-PAServer } catch { }
    if (-not $cur -or $cur.Name -ne $server) {
        Set-PAServer $server | Out-Null
        if (-not $Quiet) { Write-Log "Posh-ACME server set to $server" INFO }
    }
    if (-not $Quiet) { Write-Log "Posh-ACME $((Get-Module Posh-ACME).Version) ready, home = $acmeHome, server = $server" INFO }
}

function Test-PfxMode {
    # True when the tool runs in "import a PFX produced elsewhere" mode instead of issuing via ACME.
    if ($ImportPfx -or $SetPfxPassword) { return $true }
    if (-not $script:Config) { return $false }
    if (-not $script:Config.PSObject.Properties['Pfx']) { return $false }
    return [bool]$script:Config.Pfx.Enabled
}

function Get-MainDomain {
    # ACME mode: the first configured domain. PFX mode: the subject of the last imported certificate
    # (remembered in config), so -Status and the expiry check work without touching the PFX file.
    $domains = @()
    if ($script:Config.PSObject.Properties['Acme'] -and $script:Config.Acme.PSObject.Properties['Domains']) {
        $domains = @($script:Config.Acme.Domains | Where-Object { $_ })
    }
    if ($domains.Count) { return $domains[0] }
    if ($script:Config.PSObject.Properties['Pfx'] -and $script:Config.Pfx.Subject) {
        return ($script:Config.Pfx.Subject -replace '^CN=', '')
    }
    throw 'No certificate subject known: configure Acme.Domains, or import a PFX once (-ImportPfx).'
}

function Get-MainDomainOrNull {
    # For display/monitoring paths: before the first PFX import there is no subject yet, and that
    # must not abort -Status.
    try { return (Get-MainDomain) } catch { return $null }
}

function Get-PAOrderSafe {
    <# Get-PAOrder throws a terminating error when no account/order exists; normalise to $null. #>
    param([string]$MainDomain)
    try { return (Get-PAOrder -MainDomain $MainDomain -ErrorAction Stop) } catch { return $null }
}

function Get-CertSubject { param([string]$Domain) "CN=$Domain" }
#endregion

#region ---------------------------------------------------------------- Certificate installation
function ConvertTo-SecureStringSafe {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [System.Security.SecureString]) { return $Value }
    if ($Value -is [string] -and $Value.Length -gt 0) { return (ConvertTo-SecureString -String $Value -AsPlainText -Force) }
    return $null
}

function Get-PfxPasswordSecure {
    param($PACert)
    # PfxPass is a SecureString on a Get-PACertificate object and plaintext on a New-PACertificate
    # result; the Posh-ACME order always has the password too. Try the cert first, then the order.
    $sec = $null
    if ($PACert.PSObject.Properties['PfxPass']) { $sec = ConvertTo-SecureStringSafe $PACert.PfxPass }
    # In PFX mode there is no Posh-ACME order to fall back to - the password comes from the
    # DPAPI-protected file and is already on the object built by New-CertFromPfx.
    if (-not $sec -and -not (Test-PfxMode)) {
        try { $o = Get-PAOrder -MainDomain (Get-MainDomain); if ($o) { $sec = ConvertTo-SecureStringSafe $o.PfxPass } } catch { }
    }
    if (-not $sec) { throw 'Could not determine the PFX password from the certificate or the Posh-ACME order.' }
    return $sec
}

function Install-CertificateOnServer {
    <#
      Import -> Enable -> loopback binding -> iisreset -> verify. Fallback: re-import into CNG KSP.
      Returns $true on success.
    #>
    param([pscustomobject]$Server, [byte[]]$PfxBytes, [securestring]$PfxPassword, [string]$Thumbprint)
    $ex = $script:Config.Exchange
    Write-Log "[$($Server.Name)] Installing certificate $Thumbprint" STEP

    # 1. Import (skip if already present)
    $existing = Get-ExchangeCertificate -Server $Server.Name -Thumbprint $Thumbprint -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Log "[$($Server.Name)] Certificate already present in store; skipping import." INFO
    } else {
        Import-ExchangeCertificate -Server $Server.Name -FileData $PfxBytes -Password $PfxPassword `
            -PrivateKeyExportable:([bool]$ex.PrivateKeyExportable) | Out-Null
        Write-Log "[$($Server.Name)] Import-ExchangeCertificate done." OK
    }
    $provider = Get-KeyProviderInfo -Server $Server -Thumbprint $Thumbprint
    Write-Log "[$($Server.Name)] Key provider: $provider" INFO

    # 2. Enable (and keep the server's own certificate as internal transport certificate)
    Enable-CertificateServices -Server $Server -Thumbprint $Thumbprint

    # 3. Loopback binding 127.0.0.1:443 (only if such a binding exists)
    if ($ex.LoopbackBinding) { Update-LoopbackBinding -Server $Server -Thumbprint $Thumbprint }

    # 4. iisreset
    Invoke-IisReset -Server $Server

    # 5. Verify HTTPS (SMTP is verified later, after connectors are updated)
    if (Test-HttpsPresented -Server $Server -Thumbprint $Thumbprint) { return $true }

    # 6. Fallback: re-import with CNG KSP (only when HTTPS could not present the key)
    Write-Log "[$($Server.Name)] HTTPS verification failed. Trying fallback: re-import into CNG Key Storage Provider." WARN
    $script:UsedFallback += $Server.Name
    $ok = Invoke-KspFallback -Server $Server -PfxBytes $PfxBytes -PfxPassword $PfxPassword -Thumbprint $Thumbprint
    if (-not $ok) { return $false }
    Enable-CertificateServices -Server $Server -Thumbprint $Thumbprint
    if ($ex.LoopbackBinding) { Update-LoopbackBinding -Server $Server -Thumbprint $Thumbprint }
    Invoke-IisReset -Server $Server
    return (Test-HttpsPresented -Server $Server -Thumbprint $Thumbprint)
}

function Enable-CertificateServices {
    <#
      Enable-ExchangeCertificate -Services ... -Force. With SMTP, -Force also makes the certificate the
      internal transport certificate (server-to-server TLS) - that is the Exchange default behaviour
      and is intentionally left as is. Event 12017 ("internal transport certificate will expire soon")
      is therefore expected with short-lived certificates; the renewal job keeps it from expiring.
    #>
    param([pscustomobject]$Server, [string]$Thumbprint)
    $ex = $script:Config.Exchange
    Enable-ExchangeCertificate -Server $Server.Name -Thumbprint $Thumbprint -Services $ex.Services -Force -WarningAction SilentlyContinue
    Write-Log "[$($Server.Name)] Enable-ExchangeCertificate -Services $($ex.Services) done." OK
}

function Get-KeyProviderInfo {
    param([pscustomobject]$Server, [string]$Thumbprint)
    $sb = {
        param($tp)
        $out = certutil -store My $tp 2>&1 | Select-String 'Provider ='
        if ($out) { return ($out[0].ToString().Trim() -replace '^Provider = ', '') }
        return 'unknown'
    }
    $r = Invoke-OnServer -Server $Server -ScriptBlock $sb -ArgumentList @($Thumbprint)
    if ($null -eq $r) { return 'n/a (no WinRM)' }
    return [string]$r
}

function Update-LoopbackBinding {
    param([pscustomobject]$Server, [string]$Thumbprint)
    $sb = {
        param($tp)
        $show = netsh http show sslcert ipport=127.0.0.1:443 2>&1
        if ($LASTEXITCODE -ne 0 -or -not ($show -match 'Certificate Hash')) { return 'none' }
        $hash  = (($show | Select-String 'Certificate Hash').ToString() -split ':\s+')[1].Trim()
        $appid = (($show | Select-String 'Application ID').ToString() -split ':\s+')[1].Trim()
        if ($hash -ieq $tp) { return 'current' }
        netsh http delete sslcert ipport=127.0.0.1:443 | Out-Null
        netsh http add sslcert ipport=127.0.0.1:443 certhash=$tp appid="$appid" certstorename=MY | Out-Null
        return "updated (was $hash)"
    }
    $r = Invoke-OnServer -Server $Server -ScriptBlock $sb -ArgumentList @($Thumbprint)
    if ($null -ne $r) { Write-Log "[$($Server.Name)] Loopback binding 127.0.0.1:443: $r" INFO }
}

function Invoke-IisReset {
    param([pscustomobject]$Server)
    if ($Server.IsLocal) { $out = iisreset /noforce /timeout:90 2>&1 } else { $out = iisreset $Server.Fqdn /noforce /timeout:90 2>&1 }
    Write-Log "[$($Server.Name)] iisreset: $(($out | Select-Object -Last 1))" INFO
    # On busy Exchange servers iisreset can report a start timeout (8007041d) and leave W3SVC stopped -
    # http.sys still answers TLS on :443 so the cert probe passes while the web sites are actually down.
    # Make sure W3SVC (and thus IIS) is running before moving on.
    try {
        $svc = if ($Server.IsLocal) { Get-Service -Name W3SVC -ErrorAction Stop }
               else { Get-Service -ComputerName $Server.Fqdn -Name W3SVC -ErrorAction Stop }
        if ($svc.Status -ne 'Running') {
            Write-Log "[$($Server.Name)] W3SVC is $($svc.Status) after iisreset - starting it." WARN
            $svc | Start-Service -ErrorAction Stop
            $svc.WaitForStatus('Running', (New-TimeSpan -Seconds 60))
            Write-Log "[$($Server.Name)] W3SVC is now Running." OK
        }
    } catch { Write-Log "[$($Server.Name)] Could not verify/start W3SVC after iisreset: $($_.Exception.Message)" WARN }
}

function Test-HttpsPresented {
    <#
      Per-server HTTPS verification only. HTTPS :443 presenting the new thumbprint is the real
      "Schannel can build a server credential from this key" test, so a failure here is what should
      trigger the CNG KSP fallback. SMTP :25 is verified separately AFTER connectors are updated,
      because what :25 presents is governed by the connector's TlsCertificateName, not the store.
    #>
    param([pscustomobject]$Server, [string]$Thumbprint)
    if ($script:Config.Exchange.Services -notmatch 'IIS') { return $true }
    $fqdn = $Server.Fqdn
    $probe = { Test-TlsEndpoint -HostName $fqdn -Port 443 }.GetNewClosure()
    $ok = Wait-ForThumbprint -Expected $Thumbprint -Probe $probe
    Write-Log "[$($Server.Name)] HTTPS :443 presents new certificate: $ok" $(if ($ok) { 'OK' } else { 'ERROR' })
    return $ok
}

function Confirm-SmtpPresented {
    <#
      Soft SMTP verification, run once per server AFTER connectors reference the new certificate.
      A mismatch is a warning, not a failure (it does not mean the key is unusable) - so it never
      triggers the KSP fallback or a reinstall. One optional transport restart to nudge it.
    #>
    param([pscustomobject[]]$Servers, [string]$Thumbprint)
    if ($script:Config.Exchange.Services -notmatch 'SMTP') { return }
    foreach ($s in $Servers) {
        $fqdn = $s.Fqdn
        $probe = { Test-StartTls -HostName $fqdn -Port 25 }.GetNewClosure()
        $ok = Wait-ForThumbprint -Expected $Thumbprint -Probe $probe -Attempts 3 -DelaySeconds 5
        if (-not $ok -and $script:Config.Exchange.RestartTransportIfStale) {
            Write-Log "[$($s.Name)] SMTP :25 not yet presenting the new certificate; restarting MSExchangeTransport once." WARN
            try {
                Get-Service -ComputerName $fqdn -Name MSExchangeTransport | Restart-Service -Force
                Start-Sleep -Seconds 15
                $ok = Wait-ForThumbprint -Expected $Thumbprint -Probe $probe -Attempts 3 -DelaySeconds 5
            } catch { Write-Log "[$($s.Name)] Transport restart failed: $($_.Exception.Message)" WARN }
        }
        Write-Log "[$($s.Name)] SMTP :25 presents new certificate: $ok" $(if ($ok) { 'OK' } else { 'WARN' })
    }
}

function Invoke-KspFallback {
    param([pscustomobject]$Server, [byte[]]$PfxBytes, [securestring]$PfxPassword, [string]$Thumbprint)
    if (-not $Server.IsLocal -and -not (Test-WinRm $Server.Fqdn)) {
        Write-Log "[$($Server.Name)] Fallback requires WinRM, which is not reachable. Manual intervention needed." ERROR
        return $false
    }
    try { Remove-ExchangeCertificate -Server $Server.Name -Thumbprint $Thumbprint -Confirm:$false -ErrorAction Stop }
    catch { Write-Log "[$($Server.Name)] Remove before fallback import: $($_.Exception.Message)" WARN }
    $plain = [Net.NetworkCredential]::new('', $PfxPassword).Password
    $sb = {
        param($bytes, $pw, $tp)
        $tmp = Join-Path $env:TEMP ("acme-{0}.pfx" -f [guid]::NewGuid())
        try {
            [IO.File]::WriteAllBytes($tmp, $bytes)
            $out = certutil -f -csp 'Microsoft Software Key Storage Provider' -p $pw -importPFX My $tmp 2>&1
            $prov = certutil -store My $tp 2>&1 | Select-String 'Provider ='
            return "$($out | Select-Object -Last 1) / $prov"
        } finally { if (Test-Path $tmp) { Remove-Item $tmp -Force } }
    }
    $r = Invoke-OnServer -Server $Server -ScriptBlock $sb -ArgumentList @($PfxBytes, $plain, $Thumbprint)
    Write-Log "[$($Server.Name)] KSP fallback import: $r" INFO
    return ($null -ne $r)
}

function Update-ConnectorCertificates {
    param($NewCert)   # Exchange certificate object
    $subject = $NewCert.Subject
    $newName = "<I>$($NewCert.Issuer)<S>$subject"
    $changed = @()
    foreach ($sc in Get-SendConnector) {
        $cur = if ($sc.TlsCertificateName) { $sc.TlsCertificateName.ToString() } else { '' }
        if ($cur -like "*<S>$subject" -and $cur -ne $newName) {
            Set-SendConnector -Identity $sc.Identity -TlsCertificateName $newName
            $changed += "SendConnector '$($sc.Name)'"
        }
    }
    foreach ($srvName in $script:Config.Exchange.Servers) {
        foreach ($rc in Get-ReceiveConnector -Server $srvName) {
            $cur = if ($rc.TlsCertificateName) { $rc.TlsCertificateName.ToString() } else { '' }
            if ($cur -like "*<S>$subject" -and $cur -ne $newName) {
                Set-ReceiveConnector -Identity $rc.Identity -TlsCertificateName $newName
                $changed += "ReceiveConnector '$($rc.Identity)'"
            }
        }
    }
    if ($changed.Count) { Write-Log "Connectors updated to '$newName': $($changed -join '; ')" OK }
    else { Write-Log 'No connector TlsCertificateName needed updating.' INFO }
    return $changed
}

function Remove-SupersededCertificates {
    <#
      Removes older certificates with the same subject (never self-signed ones).
      A send connector's TlsCertificateName can block removal, but only when it references the exact
      certificate being removed (issuer+subject). Update-ConnectorCertificates has normally already
      moved every connector to the new certificate, so nothing needs touching here. As a safety net,
      a send connector still pointing at a certificate being removed is repointed to the NEW one
      (never cleared-and-restored, which would leave it on a certificate that no longer exists and
      would needlessly disturb a connector that already uses the new certificate).
    #>
    param($NewCert, [pscustomobject[]]$Servers)
    $removed = @()
    $subject = $NewCert.Subject
    $newRef  = "<I>$($NewCert.Issuer)<S>$subject"
    $candidates = @()
    foreach ($s in $Servers) {
        $old = @(Get-ExchangeCertificate -Server $s.Name | Where-Object {
            $_.Subject -eq $subject -and $_.Thumbprint -ne $NewCert.Thumbprint -and $_.Issuer -ne $_.Subject
        })
        foreach ($o in $old) { $candidates += [pscustomobject]@{ Server = $s; Cert = $o } }
    }
    if (-not $candidates.Count) { Write-Log 'No superseded certificates to remove.' INFO; return $removed }

    # only these exact references (issuer+subject of a certificate being removed) can block removal
    $oldRefs = @($candidates | ForEach-Object { "<I>$($_.Cert.Issuer)<S>$($_.Cert.Subject)" } | Select-Object -Unique)
    $repointed = $false
    foreach ($sc in Get-SendConnector) {
        if ($sc.TlsCertificateName -and ($oldRefs -contains $sc.TlsCertificateName.ToString())) {
            Set-SendConnector -Identity $sc.Identity -TlsCertificateName $newRef
            Write-Log "Send connector '$($sc.Name)' repointed to the new certificate (it still referenced one being removed)." INFO
            $repointed = $true
        }
    }

    foreach ($c in $candidates) {
        $srv = $c.Server; $o = $c.Cert
        $ok = $false
        for ($i = 1; $i -le 6 -and -not $ok; $i++) {
            if ($repointed -and $i -eq 1) { Start-Sleep -Seconds 20 }   # let the connector change replicate
            try {
                Remove-ExchangeCertificate -Server $srv.Name -Thumbprint $o.Thumbprint -Confirm:$false -ErrorAction Stop
                $ok = $true
            } catch {
                if ($i -lt 6 -and $_.Exception.Message -match 'Connector') { Start-Sleep -Seconds 20; continue }
                Write-Log "[$($srv.Name)] Could not remove $($o.Thumbprint): $($_.Exception.Message)" WARN
                break
            }
        }
        if ($ok) {
            $removed += "$($srv.Name): $($o.Thumbprint) (exp $($o.NotAfter.ToString('yyyy-MM-dd')))"
            Write-Log "[$($srv.Name)] Removed superseded certificate $($o.Thumbprint)" OK
        }
    }
    return $removed
}

function Backup-Pfx {
    param([string]$PfxPath, [string]$Domain)
    $dir = Join-Path $script:Config.Paths.Home 'backup'
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $safe = $Domain -replace '[*]', 'wildcard' -replace '[^A-Za-z0-9.-]', '_'
    $dest = Join-Path $dir ("{0}_{1}.pfx" -f $safe, (Get-Date -Format 'yyyyMMdd_HHmmss'))
    Copy-Item $PfxPath $dest -Force
    $keep = [int]$script:Config.Paths.BackupKeep
    Get-ChildItem $dir -Filter "$safe`_*.pfx" | Sort-Object LastWriteTime -Descending | Select-Object -Skip $keep |
        Remove-Item -Force -ErrorAction SilentlyContinue
    Write-Log "PFX backup: $dest (password is the order PfxPass in Posh-ACME home)" INFO
    return $dest
}

function Install-Certificate {
    <# Full install pipeline for a freshly issued PACertificate. Returns a result object. #>
    param($PACert)
    Connect-ExchangeManagement
    $tp      = $PACert.Thumbprint.ToUpper()
    $pfxPath = if ($PACert.PfxFullChain) { $PACert.PfxFullChain } else { $PACert.PfxFile }
    $bytes   = [IO.File]::ReadAllBytes($pfxPath)
    $pw      = Get-PfxPasswordSecure $PACert

    $servers = @()
    foreach ($n in $script:Config.Exchange.Servers) { $servers += Resolve-ExchangeServer $n }
    # local server first
    $servers = @($servers | Sort-Object { -not $_.IsLocal })

    $failed = @()
    foreach ($s in $servers) {
        try {
            if (-not (Install-CertificateOnServer -Server $s -PfxBytes $bytes -PfxPassword $pw -Thumbprint $tp)) {
                $failed += $s.Name
            }
        } catch {
            Write-Log "[$($s.Name)] Installation error: $($_.Exception.Message)" ERROR
            $failed += $s.Name
        }
    }
    if ($failed.Count) {
        throw "Certificate installation failed on: $($failed -join ', '). Old certificate left in place where possible."
    }

    $newCert = Get-ExchangeCertificate -Server $servers[0].Name -Thumbprint $tp
    # Update the connectors BEFORE checking SMTP, because what :25 presents follows the connector's
    # TlsCertificateName (issuer+subject), which changes when the CA's intermediate changes.
    $changed = @()
    if ($script:Config.Exchange.UpdateConnectors) { $changed = Update-ConnectorCertificates -NewCert $newCert }
    Confirm-SmtpPresented -Servers $servers -Thumbprint $tp
    $removed = @()
    if ($script:Config.Exchange.RemoveOldCertificate) { $removed = Remove-SupersededCertificates -NewCert $newCert -Servers $servers }
    # Name the backup after the certificate that was actually installed - works in both modes and
    # does not depend on the configured ACME domain list.
    $backup = Backup-Pfx -PfxPath $pfxPath -Domain ($newCert.Subject -replace '^CN=', '' -replace ',.*$', '')

    [pscustomobject]@{
        Thumbprint = $tp; Subject = $newCert.Subject; Issuer = $newCert.Issuer; NotAfter = $newCert.NotAfter
        Servers = ($servers | ForEach-Object { $_.Name }); Connectors = $changed; Removed = $removed
        Backup = $backup; Fallback = $script:UsedFallback
    }
}
#endregion

#region ---------------------------------------------------------------- PFX import (no ACME)
function Get-PfxSettings {
    <# Pfx config with defaults filled in. Tolerates a config.json written before this feature. #>
    $homeDir = $script:Config.Paths.Home
    $p = if ($script:Config.PSObject.Properties['Pfx']) { $script:Config.Pfx } else { $null }
    $get = {
        param($name, $default)
        if ($p -and $p.PSObject.Properties[$name] -and "$($p.$name)") { return $p.$name }
        return $default
    }
    [pscustomobject]@{
        DropFolder         = & $get 'DropFolder'   (Join-Path $homeDir 'pfx')
        PasswordFile       = & $get 'PasswordFile' (Join-Path $homeDir 'pfxpass.dat')
        ArchiveAfterImport = if ($p -and $p.PSObject.Properties['ArchiveAfterImport']) { [bool]$p.ArchiveAfterImport } else { $true }
    }
}

function Set-PfxPasswordFile {
    <#
      Store the PFX password encrypted with DPAPI in the LOCAL MACHINE scope, so the SYSTEM
      scheduled task can read it without any interactive logon. The blob is useless on another
      machine. Protection against local admins is NOT the goal and not achievable here - they can
      export the private key from the certificate store anyway.
    #>
    param([securestring]$Password)
    $s = Get-PfxSettings
    if (-not $Password) {
        if ($script:NonInteractive) { throw 'No password supplied. Use -SetPfxPassword -PfxPassword (Read-Host -AsSecureString).' }
        $Password = Read-Host 'PFX password' -AsSecureString
    }
    $bstr  = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
    try { $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }

    Add-Type -AssemblyName System.Security
    $bytes = [Text.Encoding]::UTF8.GetBytes($plain)
    $blob  = [Security.Cryptography.ProtectedData]::Protect($bytes, $null, [Security.Cryptography.DataProtectionScope]::LocalMachine)
    [Array]::Clear($bytes, 0, $bytes.Length)

    $dir = Split-Path $s.PasswordFile -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    [IO.File]::WriteAllBytes($s.PasswordFile, $blob)
    Protect-PfxPath -Path $s.PasswordFile
    Write-Log "PFX password stored DPAPI-encrypted (machine scope) in $($s.PasswordFile)" OK

    # verify by reading it back - a password file that cannot be decrypted is worse than none
    $check = Get-PfxPasswordFromFile
    if (-not $check) { throw "Password file was written but could not be read back: $($s.PasswordFile)" }
    Write-Log 'Read-back check passed.' OK
}

function Get-PfxPasswordFromFile {
    <# Returns the stored password as SecureString, or $null when there is none. #>
    $s = Get-PfxSettings
    if (-not (Test-Path $s.PasswordFile)) { return $null }
    try {
        Add-Type -AssemblyName System.Security
        $blob  = [IO.File]::ReadAllBytes($s.PasswordFile)
        $bytes = [Security.Cryptography.ProtectedData]::Unprotect($blob, $null, [Security.Cryptography.DataProtectionScope]::LocalMachine)
        $plain = [Text.Encoding]::UTF8.GetString($bytes)
        [Array]::Clear($bytes, 0, $bytes.Length)
        if (-not $plain) { return $null }
        return (ConvertTo-SecureString -String $plain -AsPlainText -Force)
    } catch {
        Write-Log "Could not decrypt $($s.PasswordFile): $($_.Exception.Message). Re-run -SetPfxPassword (the file is bound to this machine)." ERROR
        return $null
    }
}

function Protect-PfxPath {
    <# SYSTEM + Administrators only, inheritance off - same hardening the working directory gets. #>
    param([string]$Path)
    try {
        $acl = Get-Acl $Path
        $acl.SetAccessRuleProtection($true, $false)
        foreach ($r in @($acl.Access)) { [void]$acl.RemoveAccessRule($r) }
        $isDir = (Get-Item $Path).PSIsContainer
        foreach ($id in 'NT AUTHORITY\SYSTEM', 'BUILTIN\Administrators') {
            $rule = if ($isDir) {
                New-Object Security.AccessControl.FileSystemAccessRule($id, 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
            } else {
                New-Object Security.AccessControl.FileSystemAccessRule($id, 'FullControl', 'Allow')
            }
            $acl.AddAccessRule($rule)
        }
        Set-Acl $Path $acl
    } catch { Write-Log "Could not harden ACL on ${Path}: $($_.Exception.Message)" WARN }
}

function Resolve-PfxFile {
    <# Explicit -PfxPath wins; otherwise the newest *.pfx in the drop folder. #>
    param([string]$Path)
    if ($Path) {
        if (-not (Test-Path $Path)) { throw "PFX not found: $Path" }
        return (Get-Item $Path)
    }
    $s = Get-PfxSettings
    if (-not (Test-Path $s.DropFolder)) {
        New-Item -ItemType Directory -Path $s.DropFolder -Force | Out-Null
        Protect-PfxPath -Path $s.DropFolder
        throw "Drop folder $($s.DropFolder) was empty (just created). Put the PFX there and run again."
    }
    $f = Get-ChildItem $s.DropFolder -Filter *.pfx -File -ErrorAction SilentlyContinue |
         Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $f) { throw "No *.pfx in $($s.DropFolder)." }
    return $f
}

function New-CertFromPfx {
    <#
      Builds the same shape Install-Certificate expects from a PFX on disk, after checking that the
      file really is usable: password fits, private key present, not expired, chain complete.
      A PFX without its intermediate would import and bind, but clients would reject the chain -
      so that is an error here, not a surprise in production.
    #>
    param([string]$Path)
    $file = Resolve-PfxFile -Path $Path
    $pw   = Get-PfxPasswordFromFile
    if (-not $pw) { throw "No PFX password stored. Run once: -SetPfxPassword (see README)." }

    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($pw)
    try { $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }

    $coll = New-Object Security.Cryptography.X509Certificates.X509Certificate2Collection
    try {
        # EphemeralKeySet is not available on PS 5.1/.NET 4.x, so this touches the user key store
        # briefly; the certificate itself is installed later by Import-ExchangeCertificate.
        $coll.Import($file.FullName, $plain, 'PersistKeySet,Exportable')
    } catch {
        throw "Could not open $($file.Name): $($_.Exception.Message). Wrong password, or the file is not a PFX."
    }
    $leaf = @($coll | Where-Object { $_.HasPrivateKey }) | Sort-Object NotAfter -Descending | Select-Object -First 1
    if (-not $leaf) { throw "$($file.Name) contains no certificate with a private key." }
    if ($leaf.NotAfter -lt (Get-Date)) { throw "$($file.Name) expired on $($leaf.NotAfter.ToString('yyyy-MM-dd')) - refusing to install it." }
    if ($leaf.NotBefore -gt (Get-Date)) { throw "$($file.Name) is not valid before $($leaf.NotBefore.ToString('yyyy-MM-dd'))." }

    # Chain check. Only a MISSING intermediate is fatal - clients would reject such a certificate
    # even though Exchange binds it happily. An untrusted root (internal CA) or an unreachable CRL
    # is the administrator's decision, not an error, so those are logged and accepted.
    $chain = New-Object Security.Cryptography.X509Certificates.X509Chain
    $chain.ChainPolicy.RevocationMode = 'NoCheck'
    foreach ($c in $coll) { if (-not $c.HasPrivateKey) { [void]$chain.ChainPolicy.ExtraStore.Add($c) } }
    [void]$chain.Build($leaf)
    $stati = @($chain.ChainStatus | ForEach-Object { $_.Status })
    if ($stati -contains 'PartialChain') {
        throw "Chain for $($file.Name) is incomplete - an intermediate certificate is missing. Export the PFX including its intermediates."
    }
    $selfSigned = ($leaf.Subject -eq $leaf.Issuer)
    if (-not $selfSigned -and $chain.ChainElements.Count -lt 2) {
        throw "Chain for $($file.Name) has only the leaf certificate - export the PFX including its intermediates."
    }
    foreach ($st in ($stati | Where-Object { $_ -ne 'NoError' } | Select-Object -Unique)) {
        Write-Log "PFX chain note: $st (accepted - only a missing intermediate is treated as an error)" WARN
    }
    Write-Log "PFX $($file.Name): subject=$($leaf.Subject) issuer=$($leaf.Issuer) notAfter=$($leaf.NotAfter.ToString('yyyy-MM-dd')) chain=$($chain.ChainElements.Count) elements" INFO

    [pscustomobject]@{
        Thumbprint   = $leaf.Thumbprint.ToUpper()
        Subject      = $leaf.Subject
        NotAfter     = $leaf.NotAfter
        PfxFile      = $file.FullName
        PfxFullChain = $file.FullName
        PfxPass      = $pw          # SecureString - Get-PfxPasswordSecure takes it straight from here
        SourceFile   = $file
    }
}

function Resolve-LongPath {
    <#
      Full path in its long form. Resolve-Path alone is not enough: it normalises
      "..", drive-relative paths and casing, but it leaves an 8.3 short name short.
      A configured drop folder may well be short (a build agent's TEMP is
      C:\Users\RUNNER~1\...), while a FileInfo coming from Get-ChildItem is long -
      comparing the two as strings then fails even though both mean the same folder.
      Only GetLongPathName settles that.
    #>
    param([string]$Path)
    if (-not $Path) { return $null }
    $resolved = (Resolve-Path -LiteralPath $Path -ErrorAction SilentlyContinue).Path
    if (-not $resolved) { return $null }

    if (-not ('Native.LongPath' -as [type])) {
        try {
            Add-Type -Namespace Native -Name LongPath -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("kernel32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode, SetLastError = true)]
public static extern uint GetLongPathName(string lpszShortPath, System.Text.StringBuilder lpszLongPath, uint cchBuffer);
'@ -ErrorAction Stop
        } catch { return $resolved }   # ohne die API bleibt der aufgeloeste Pfad das Beste
    }
    try {
        $sb = New-Object System.Text.StringBuilder 1024
        $len = [Native.LongPath]::GetLongPathName($resolved, $sb, 1024)
        if ($len -gt 0 -and $len -lt 1024) { return $sb.ToString() }
    } catch { }
    $resolved
}

function Move-ImportedPfx {
    <# Keep key material out of the drop folder once it is installed. #>
    param($File)
    $s = Get-PfxSettings
    if (-not $s.ArchiveAfterImport) {
        # Used to return in silence. If archiving is off, say so - otherwise a PFX
        # that quietly stays in the drop folder looks like a bug somewhere else.
        Write-Log "PFX not archived: ArchiveAfterImport is off." INFO
        return
    }

    # Bring BOTH sides to the same long form before comparing - see Resolve-LongPath.
    # Otherwise a short-named drop folder never matches the file's long directory, the
    # function returns, and the PFX silently stays put with its private key - exactly
    # what archiving is meant to prevent.
    $dropDir = Resolve-LongPath $s.DropFolder
    $fileDir = Resolve-LongPath $File.DirectoryName
    if (-not $dropDir -or -not $fileDir) {
        Write-Log "PFX not archived: cannot resolve '$($File.DirectoryName)' or '$($s.DropFolder)'." WARN
        return
    }
    if ($fileDir.TrimEnd('\') -ine $dropDir.TrimEnd('\')) {
        # -PfxPath from elsewhere: leave it alone. Say so instead of returning in silence.
        Write-Log "PFX left in place: '$fileDir' is not the drop folder '$dropDir'." INFO
        return
    }

    $archive = Join-Path $s.DropFolder 'archive'
    if (-not (Test-Path $archive)) { New-Item -ItemType Directory -Path $archive -Force | Out-Null; Protect-PfxPath -Path $archive }
    $dest = Join-Path $archive ("{0}_{1}{2}" -f $File.BaseName, (Get-Date -Format 'yyyyMMdd_HHmmss'), $File.Extension)
    Move-Item $File.FullName $dest -Force
    Write-Log "PFX moved out of the drop folder: $dest" OK
}

function Invoke-PfxImport {
    <# Same deployment pipeline as -Renew, but the certificate comes from a file instead of ACME. #>
    $cert = New-CertFromPfx -Path $PfxPath
    Write-Log "Certificate run from PFX $(Split-Path $cert.PfxFile -Leaf) (thumbprint $($cert.Thumbprint), skipInstall=$($script:SkipInstall))" STEP

    # Idempotency guard: without it a daily task would re-run iisreset on every single run.
    if (Test-CertDeployed $cert.Thumbprint) {
        Write-Log "Certificate $($cert.Thumbprint) is already presented by every configured server - nothing to do." INFO
        Save-PfxSubject -Subject $cert.Subject   # also on this path: Get-MainDomain needs it below
        Move-ImportedPfx -File $cert.SourceFile
        Invoke-ExpiryCheck
        return
    }
    if ($script:SkipInstall) {
        Write-Log "SkipInstall: PFX $($cert.Thumbprint) validated but NOT installed on any server (rehearsal)." WARN
        return
    }

    $result = Install-Certificate -PACert $cert
    Save-PfxSubject -Subject $result.Subject
    Move-ImportedPfx -File $cert.SourceFile
    Publish-InstallResult -Result $result -Headline 'A certificate was imported from PFX and installed successfully.'
}

function Save-PfxSubject {
    <# Remember the subject so -Status and the expiry check work without the PFX file. #>
    param([string]$Subject)
    try {
        if (-not $script:Config.PSObject.Properties['Pfx']) {
            $script:Config | Add-Member -NotePropertyName Pfx -NotePropertyValue ([pscustomobject](Get-DefaultConfig).Pfx)
        }
        if ($script:Config.Pfx.Subject -eq $Subject -and $script:Config.Pfx.Enabled) { return }
        $script:Config.Pfx.Subject = $Subject
        $script:Config.Pfx.Enabled = $true
        Save-Config $script:Config
    } catch { Write-Log "Could not remember the PFX subject in the config: $($_.Exception.Message)" WARN }
}
#endregion

#region ---------------------------------------------------------------- Renew
function Publish-InstallResult {
    <# One reporting path for both modes (ACME renewal and PFX import). #>
    param($Result, [string]$Headline)
    $body = @"
$Headline

Subject:     $($Result.Subject)
Issuer:      $($Result.Issuer)
Thumbprint:  $($Result.Thumbprint)
Valid until: $($Result.NotAfter)
Servers:     $($Result.Servers -join ', ')
Services:    $($script:Config.Exchange.Services)

Connectors updated: $(if ($Result.Connectors.Count) { "`r`n  " + ($Result.Connectors -join "`r`n  ") } else { 'none' })
Old certificates removed: $(if ($Result.Removed.Count) { "`r`n  " + ($Result.Removed -join "`r`n  ") } else { 'none' })
KSP fallback used on: $(if ($Result.Fallback.Count) { $Result.Fallback -join ', ' } else { 'none' })
Backup: $($Result.Backup)
"@
    Write-EventLogEntry -Message $body -Type Information -Id 9000
    if ($script:Config.Notify.SendSuccessMail) {
        $what = ($Result.Subject -replace '^CN=', '' -replace ',.*$', '')
        Send-Notification -Subject "Certificate deployed for $what (valid until $($Result.NotAfter.ToString('yyyy-MM-dd')))" -Body $body -Level Info
    }
}

function Get-BoundCertificateStatus {
    <#
      Per server: what HTTPS (:443) and SMTP STARTTLS (:25) really present (probe), plus the
      certificate Exchange *thinks* is enabled for IIS (metadata). DaysLeft is the minimum over
      the probed certificates so that a stale manual binding is caught.
    #>
    Connect-ExchangeManagement
    $md = Get-MainDomainOrNull
    $subject = if ($md) { Get-CertSubject $md } else { $null }   # unknown before the first PFX import
    $rows = @()
    foreach ($n in $script:Config.Exchange.Servers) {
        $srv = Resolve-ExchangeServer $n
        $store = @(Get-ExchangeCertificate -Server $srv.Name)
        $flag = if ($subject) { $store | Where-Object { $_.Subject -eq $subject -and $_.Services -match 'IIS' } | Sort-Object NotAfter -Descending | Select-Object -First 1 } else { $null }
        $httpsTp = Test-TlsEndpoint -HostName $srv.Fqdn -Port 443
        $smtpTp  = Test-StartTls   -HostName $srv.Fqdn -Port 25
        $days = @()
        foreach ($tp in @($httpsTp, $smtpTp)) {
            if (-not $tp) { $days += -1; continue }
            $c = $store | Where-Object { $_.Thumbprint -ieq $tp } | Select-Object -First 1
            if ($c) { $days += [math]::Floor(($c.NotAfter - (Get-Date)).TotalDays) } else { $days += -1 }
        }
        $rows += [pscustomobject]@{
            Server    = $srv.Name
            Https     = if ($httpsTp) { $httpsTp } else { 'FAIL' }
            Smtp      = if ($smtpTp)  { $smtpTp }  else { 'FAIL' }
            DaysLeft  = ($days | Measure-Object -Minimum).Minimum
            IISFlag   = if ($flag) { $flag.Thumbprint } else { '-' }
            FlagExp   = if ($flag) { $flag.NotAfter.ToString('yyyy-MM-dd') } else { '-' }
            Consistent = ($httpsTp -and $flag -and $httpsTp -ieq $flag.Thumbprint -and $smtpTp -ieq $httpsTp)
        }
    }
    return $rows
}

function Wait-AuthCertPropagation {
    # If the app auth certificate was created very recently (e.g. by "1. Azure einrichten"),
    # give Entra ID a moment to propagate the new key before the first token request (avoids 401).
    try {
        $c = Get-Item "Cert:\LocalMachine\My\$($script:Config.Azure.AuthCertThumbprint)" -ErrorAction Stop
        if ($c.NotBefore -gt (Get-Date).AddMinutes(-3)) {
            Write-Log 'App auth certificate is fresh; waiting 30s for Entra ID propagation ...' INFO
            Start-Sleep -Seconds 30
        }
    } catch { }
}

function New-ManagedCertificate {
    # Request a certificate via the configured DNS-01 provider (first issuance or forced re-issue).
    # Retries on auth errors, which right after Azure setup are credential-propagation lag.
    if (-not (Get-PAAccount)) {
        New-PAAccount -Contact $script:Config.Acme.Contact -AcceptTOS | Out-Null
        Write-Log "ACME account created for $($script:Config.Acme.Contact)" OK
    }
    $provider = Get-DnsProviderName
    $plugin   = (Get-DnsProviders)[$provider].Plugin
    if ($provider -eq 'Azure') { Wait-AuthCertPropagation }
    $pluginArgs = Resolve-PluginArgs
    if (-not $pluginArgs) { throw "No DNS credentials available for provider '$provider'. Enter them in the GUI and try again." }
    $main = Get-MainDomain
    $pfxPass = -join ((48..57) + (65..90) + (97..122) | Get-Random -Count 24 | ForEach-Object { [char]$_ })
    $params = @{
        Domain = $script:Config.Acme.Domains; Plugin = $plugin; PluginArgs = $pluginArgs
        Contact = $script:Config.Acme.Contact; AcceptTOS = $true; PfxPass = $pfxPass
        FriendlyName = "$main (Let's Encrypt)"; CertKeyLength = $script:Config.Acme.KeyLength; Force = $true
    }
    Write-Log "Requesting certificate for $($script:Config.Acme.Domains -join ', ') via $provider DNS-01 ..." STEP
    Write-Log 'Publishing the DNS TXT record and waiting for DNS propagation + Let''s Encrypt validation - this usually takes 1-3 minutes with no further output. Please wait ...' INFO
    $cert = $null
    for ($try = 1; $try -le 4; $try++) {
        try { $cert = New-PACertificate @params; break }
        catch {
            if ($try -lt 4 -and ($_.Exception.Message -match '401|Unauthorized|AADSTS')) {
                Write-Log "Request failed with an auth error (attempt $try/4), likely credential propagation - retrying in 30s ..." WARN
                Start-Sleep -Seconds 30; continue
            }
            throw
        }
    }
    if (-not $cert) { throw 'New-PACertificate returned nothing.' }
    return $cert
}

function Test-CertDeployed {
    param([string]$Thumbprint)
    Connect-ExchangeManagement
    foreach ($n in $script:Config.Exchange.Servers) {
        $srv = Resolve-ExchangeServer $n
        $tp = Test-TlsEndpoint -HostName $srv.Fqdn -Port 443
        if (-not $tp -or $tp.ToUpper() -ne $Thumbprint.ToUpper()) { return $false }
    }
    return $true
}

function Invoke-Renewal {
    $mainDomain = Get-MainDomain
    $forced = $ForceRenew.IsPresent
    Write-Log "Certificate run for $mainDomain (force=$forced, skipInstall=$($script:SkipInstall))" STEP
    Initialize-PoshAcme
    Invoke-AuthCertCheck   # Azure DNS-01 auth cert cannot renew unattended - warn by mail to re-run -Setup

    $order = Get-PAOrderSafe -MainDomain $mainDomain
    $newCert = $null

    if (-not $order) {
        Write-Log 'No existing order - requesting the first certificate.' INFO
        $newCert = New-ManagedCertificate
    } else {
        Write-Log "Order status: $($order.status), expires $($order.CertExpires), renew after $($order.RenewAfter)" INFO
        # keep the stored plugin args current (Azure: cert-based, always; others: only when new creds were supplied)
        try { $pa = Resolve-PluginArgs; if ($pa) { Set-PAOrder -MainDomain $mainDomain -PluginArgs $pa | Out-Null } } catch { }
        Write-Log 'Renewing: publishing the DNS TXT record and waiting for DNS propagation + Let''s Encrypt validation - this usually takes 1-3 minutes with no further output. Please wait ...' INFO
        $warnings = @()
        $newCert = Submit-Renewal -MainDomain $mainDomain -Force:$forced -WarningVariable warnings -WarningAction SilentlyContinue
        foreach ($w in $warnings) { Write-Log "Posh-ACME: $w" INFO }
        if (-not $newCert) {
            # nothing renewed: if the current cert is issued but not yet on the servers, install it now
            $pa = Get-PACertificate -MainDomain $mainDomain -ErrorAction SilentlyContinue
            if ($pa -and -not (Test-CertDeployed $pa.Thumbprint)) {
                Write-Log "Certificate $($pa.Thumbprint) is issued but not installed on all servers - installing it now." INFO
                $newCert = $pa
            } else {
                Write-Log 'Nothing to do - certificate is current and already deployed.' INFO
                Invoke-ExpiryCheck
                return
            }
        }
    }

    if ($script:SkipInstall) {
        Write-Log "SkipInstall: certificate $($newCert.Thumbprint) is ready but NOT installed on any server (rehearsal)." WARN
        return
    }

    Write-Log "Deploying certificate $($newCert.Thumbprint), valid until $($newCert.NotAfter)" OK
    $result = Install-Certificate -PACert $newCert
    Publish-InstallResult -Result $result -Headline 'A new certificate was issued and installed successfully.'
}

function Invoke-ExpiryCheck {
    $rows = Get-BoundCertificateStatus
    $warn = [int]$script:Config.Notify.WarnDaysBeforeExpiry
    $bad = @($rows | Where-Object { $_.DaysLeft -lt 0 -or $_.DaysLeft -le $warn -or -not $_.Consistent })
    foreach ($r in $rows) { Write-Log "[$($r.Server)] https=$($r.Https) smtp=$($r.Smtp) iisFlag=$($r.IISFlag) daysLeft=$($r.DaysLeft) consistent=$($r.Consistent)" INFO }
    if ($bad.Count) {
        $body = "The certificate presented by the following servers expires in $warn days or less, could not be probed, or is inconsistent with the Exchange configuration - and no renewal was performed:`r`n`r`n" +
            (($bad | ForEach-Object { "  $($_.Server): https=$($_.Https) smtp=$($_.Smtp) iisFlag=$($_.IISFlag) (exp $($_.FlagExp)) daysLeft=$($_.DaysLeft) consistent=$($_.Consistent)" }) -join "`r`n") +
            "`r`n`r`nCheck with: Invoke-AcmeExchangeCert.ps1 -Status`r`nForce a renewal (re-installs on all servers): Invoke-AcmeExchangeCert.ps1 -ForceRenew"
        Write-Log 'Expiry warning threshold reached.' WARN
        Write-EventLogEntry -Message $body -Type Warning -Id 9002
        Send-Notification -Subject "WARNING: certificate for $(if (Get-MainDomainOrNull) { Get-MainDomainOrNull } else { Get-LocalServerName }) expires soon" -Body $body -Level Warning
    }
}

function Invoke-AuthCertCheck {
    # The Azure app auth certificate (DNS-01) can NOT be renewed unattended (chosen model: warn +
    # re-run -Setup). Warn by mail well before it expires; the idempotent Step 1 rotates it (it
    # reuses only when >90 days remain, else creates and re-registers a new one). Azure provider only.
    if ($script:Config.Acme.DnsProvider -and $script:Config.Acme.DnsProvider -ne 'Azure') { return }
    $tp = [string]$script:Config.Azure.AuthCertThumbprint
    if (-not $tp) { return }
    $warnDays = 60
    if ($script:Config.Azure.PSObject.Properties['AuthCertWarnDays'] -and $script:Config.Azure.AuthCertWarnDays) { $warnDays = [int]$script:Config.Azure.AuthCertWarnDays }
    $cert = Get-ChildItem Cert:\LocalMachine\My -ErrorAction SilentlyContinue | Where-Object { $_.Thumbprint -eq $tp } | Select-Object -First 1
    $missing = -not $cert
    $daysLeft = if ($cert) { [math]::Floor(($cert.NotAfter - (Get-Date)).TotalDays) } else { -1 }
    if (-not $missing -and $daysLeft -gt $warnDays) { return }   # healthy, nothing to do

    $appName = if ($script:Config.Azure.AppDisplayName) { $script:Config.Azure.AppDisplayName } else { "ACME-DNS-$(Get-MainDomain)" }
    $server  = Get-LocalServerName
    $howto = "Fix: re-run the setup wizard (Step 1) on $server - a Global Administrator sign-in is required.`r`n" +
             "  Invoke-AcmeExchangeCert.ps1 -Setup      (or in the GUI: '1. Set up Azure')`r`n" +
             "It creates a new non-exportable auth certificate and re-registers it on the app; the old key is replaced."
    if ($missing) {
        $subject = "WARNING: Azure auth certificate missing for $(Get-MainDomain)"
        $body = "The Azure app authentication certificate for DNS-01 validation was NOT found in the local " +
                "machine store on $server. Certificate renewals will fail with an Azure authentication error.`r`n`r`n" +
                "App:                    CN=$appName`r`nConfigured thumbprint:  $tp`r`n`r`n$howto"
    } else {
        $subject = "WARNING: Azure auth certificate for $(Get-MainDomain) expires in $daysLeft days"
        $body = "The Azure app authentication certificate used for DNS-01 validation expires on " +
                "$($cert.NotAfter.ToString('yyyy-MM-dd')) ($daysLeft days left). It cannot be renewed unattended.`r`n`r`n" +
                "App:        CN=$appName`r`nThumbprint: $tp`r`n`r`n$howto`r`n`r`n" +
                "If it expires, certificate renewals will start failing with an Azure authentication error."
    }

    Write-Log ("Azure auth certificate " + $(if ($missing) { 'is MISSING' } else { "expires in $daysLeft days" }) + " - re-run -Setup to rotate it.") WARN

    # throttle: at most one mail per 7 days per thumbprint, so the daily job does not spam the warn window
    $stateDir  = if ($script:LogFile) { Split-Path $script:LogFile -Parent } else { Join-Path $script:Config.Paths.Home 'logs' }
    $stateFile = Join-Path $stateDir '.authcert-warn.json'
    try {
        if (Test-Path $stateFile) {
            $st = Get-Content $stateFile -Raw | ConvertFrom-Json
            if ($st.Thumbprint -eq $tp -and $st.LastSent -and ((Get-Date) - [datetime]$st.LastSent).TotalDays -lt 7) {
                Write-Log 'Auth-cert warning already mailed within the last 7 days; not repeating.' INFO
                return
            }
        }
    } catch { }

    Write-EventLogEntry -Message $body -Type Warning -Id 9003
    Send-Notification -Subject $subject -Body $body -Level Warning
    try {
        if (-not (Test-Path $stateDir)) { New-Item -ItemType Directory -Force $stateDir | Out-Null }
        @{ Thumbprint = $tp; LastSent = (Get-Date).ToString('o') } | ConvertTo-Json | Set-Content -Path $stateFile -Encoding UTF8
    } catch { }
}
#endregion

#region ---------------------------------------------------------------- Status
function Show-Status {
    $md = Get-MainDomainOrNull
    Show-Banner "Status - $(if ($md) { $md } else { 'no certificate deployed yet' })"
    if (Test-PfxMode) {
        $s = Get-PfxSettings
        $pending = @(Get-ChildItem $s.DropFolder -Filter *.pfx -File -ErrorAction SilentlyContinue)
        Write-Host ("Certificate src : PFX import (no ACME)")
        Write-Host ("Drop folder     : {0}  ({1} PFX waiting)" -f $s.DropFolder, $pending.Count)
        $pwState = if (Test-Path $s.PasswordFile) { if (Get-PfxPasswordFromFile) { 'stored, decrypts OK' } else { 'PRESENT BUT UNREADABLE - re-run -SetPfxPassword' } } else { 'MISSING - run -SetPfxPassword' }
        Write-Host ("PFX password    : {0}" -f $pwState) -ForegroundColor $(if ($pwState -like 'stored*') { 'Gray' } else { 'Red' })
    } else {
        Initialize-PoshAcme -Quiet
        $order = if ($md) { Get-PAOrderSafe -MainDomain $md } else { $null }
        if ($order) {
            Write-Host ("Posh-ACME order : {0}  status={1}  expires={2}  renewAfter={3}" -f $order.Name, $order.status, $order.CertExpires, $order.RenewAfter)
            Write-Host ("ACME server     : {0}" -f (Get-PAServer).Name)
        } else { Write-Host 'Posh-ACME order : none' -ForegroundColor Yellow }
    }
    Write-Host ''
    Write-Host 'Presented certificates (probe) vs. Exchange IIS flag (metadata):'
    Get-BoundCertificateStatus | Format-Table Server, Https, Smtp, DaysLeft, IISFlag, FlagExp, Consistent -AutoSize -Wrap | Out-String | Write-Host
    Connect-ExchangeManagement
    if ($md) {
        $subject = Get-CertSubject $md
        Write-Host 'Connectors referencing this subject:'
        Get-SendConnector | Where-Object { $_.TlsCertificateName -and $_.TlsCertificateName.ToString() -like "*<S>$subject" } |
            ForEach-Object { Write-Host ("  Send    {0,-45} {1}" -f $_.Name, $_.TlsCertificateName) }
        foreach ($n in $script:Config.Exchange.Servers) {
            Get-ReceiveConnector -Server $n | Where-Object { $_.TlsCertificateName -and $_.TlsCertificateName.ToString() -like "*<S>$subject" } |
                ForEach-Object { Write-Host ("  Receive {0,-45} {1}" -f $_.Identity, $_.TlsCertificateName) }
        }
    } else {
        Write-Host 'Connectors: skipped - no certificate imported yet, so there is no subject to match.' -ForegroundColor Yellow
    }
    Write-Host ''
    $task = Get-ScheduledTask -TaskName $script:Config.Task.Name -ErrorAction SilentlyContinue
    if ($task) {
        $info = $task | Get-ScheduledTaskInfo
        Write-Host ("Scheduled task  : {0}  state={1}  lastRun={2}  lastResult={3}  nextRun={4}" -f $task.TaskName, $task.State, $info.LastRunTime, $info.LastTaskResult, $info.NextRunTime)
    } else { Write-Host 'Scheduled task  : not registered (run -InstallTask)' -ForegroundColor Yellow }
    if (-not (Test-PfxMode) -and (-not $script:Config.Acme.DnsProvider -or $script:Config.Acme.DnsProvider -eq 'Azure') -and $script:Config.Azure.AuthCertThumbprint) {
        $ac = Get-ChildItem Cert:\LocalMachine\My -ErrorAction SilentlyContinue | Where-Object { $_.Thumbprint -eq $script:Config.Azure.AuthCertThumbprint } | Select-Object -First 1
        if ($ac) {
            $d = [math]::Floor(($ac.NotAfter - (Get-Date)).TotalDays)
            $hint = if ($d -le 60) { '  <-- re-run -Setup to rotate' } else { '' }
            Write-Host ("Azure auth cert : {0}  expires={1} ({2} days){3}" -f $ac.Thumbprint, $ac.NotAfter.ToString('yyyy-MM-dd'), $d, $hint) -ForegroundColor $(if ($d -le 60) { 'Yellow' } else { 'Gray' })
        } else {
            Write-Host ("Azure auth cert : {0}  NOT FOUND in machine store - re-run -Setup" -f $script:Config.Azure.AuthCertThumbprint) -ForegroundColor Red
        }
    }
    if ($script:LogFile -and (Test-Path $script:LogFile)) {
        Write-Host ''; Write-Host "Last log lines ($script:LogFile):"
        Get-Content $script:LogFile -Tail 10 | ForEach-Object { Write-Host "  $_" -ForegroundColor DarkGray }
    }
}
#endregion

#region ---------------------------------------------------------------- Scheduled task
function Register-RenewalTask {
    $t = $script:Config.Task
    $scriptPath = $PSCommandPath
    # In PFX mode the daily run watches the drop folder instead of renewing via ACME.
    $mode = if (Test-PfxMode) { '-ImportPfx' } else { '-Renew' }
    $arg = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$scriptPath`" $mode -ConfigPath `"$ConfigPath`""
    $action    = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $arg
    $trigger   = New-ScheduledTaskTrigger -Daily -At $t.Time -RandomDelay (New-TimeSpan -Minutes 30)
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    $settings  = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 2) -MultipleInstances IgnoreNew
    Register-ScheduledTask -TaskName $t.Name -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
    Write-Log "Scheduled task '$($t.Name)' registered (SYSTEM, daily at $($t.Time) + up to 30 min random delay, mode $mode)." OK
}

function Unregister-RenewalTask {
    $name = $script:Config.Task.Name
    if (Get-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $name -Confirm:$false
        Write-Log "Scheduled task '$name' removed." OK
    } else {
        Write-Log "Scheduled task '$name' does not exist." INFO
    }
}
#endregion

#region ---------------------------------------------------------------- Setup wizard
function Step-Preflight {
    param($Cfg)
    Show-Banner 'Step 1/10 - Preflight'
    Connect-ExchangeManagement
    Write-Log "Exchange server: $(Get-LocalServerName)" OK

    $Cfg.Paths.Home = Read-Default 'Working directory (config, Posh-ACME state, logs, backups)' $Cfg.Paths.Home
    if (-not (Test-Path $Cfg.Paths.Home)) { New-Item -ItemType Directory -Path $Cfg.Paths.Home -Force | Out-Null }
    # restrict ACL: SYSTEM + Administrators only (private keys live here)
    $acl = Get-Acl $Cfg.Paths.Home
    $acl.SetAccessRuleProtection($true, $false)
    foreach ($id in 'NT AUTHORITY\SYSTEM', 'BUILTIN\Administrators') {
        $acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule($id, 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')))
    }
    Set-Acl $Cfg.Paths.Home $acl
    Write-Log "Working directory $($Cfg.Paths.Home) secured (SYSTEM, Administrators)." OK
    $script:Config = $Cfg
    Initialize-Logging $Cfg.Paths.Home

    # machine-wide POSHACME_HOME so the SYSTEM task and interactive admins share state
    $acmeHome = Get-PoshAcmeHome
    [Environment]::SetEnvironmentVariable('POSHACME_HOME', $acmeHome, 'Machine')
    $env:POSHACME_HOME = $acmeHome
    Write-Log "POSHACME_HOME = $acmeHome (machine environment variable)" OK

    # Posh-ACME: prefer the bundled copy (no install, works for every account incl. the SYSTEM task).
    $bundled = Get-BundledPoshAcme
    if ($bundled) {
        Write-Log "Posh-ACME bundled with the tool: $bundled (no module install needed)." OK
    } else {
        $allUsers = Get-Module Posh-ACME -ListAvailable | Where-Object { $_.Path -like "$env:ProgramFiles\*" }
        if (-not $allUsers) {
            Write-Log 'Posh-ACME is neither bundled nor installed for AllUsers (the SYSTEM task needs one of these).' WARN
            if (Read-YesNo 'Install Posh-ACME from the PowerShell Gallery (AllUsers)?' $true) {
                if (-not (Get-PackageProvider NuGet -ErrorAction SilentlyContinue)) { Install-PackageProvider NuGet -Force | Out-Null }
                Install-Module Posh-ACME -Scope AllUsers -Force -AllowClobber
            } else { throw 'Posh-ACME (bundled or AllUsers) is required.' }
        }
        Write-Log "Posh-ACME $((Get-Module Posh-ACME -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1).Version) available (installed)." OK
    }
}

function Step-Login {
    param($Cfg)
    Show-Banner 'Step 2/10 - Azure sign-in (bootstrap)'
    $Cfg.Azure.BootstrapClientId = Read-Default 'Bootstrap client id (first-party public client)' $Cfg.Azure.BootstrapClientId
    $hint = Read-Default 'Tenant id or domain (leave empty for the signed-in user''s home tenant)' $Cfg.Azure.TenantId
    if (-not $hint) { $hint = 'organizations' }
    $tok = Connect-Bootstrap -ClientId $Cfg.Azure.BootstrapClientId -TenantHint $hint
    $Cfg.Azure.TenantId = $tok.TenantId
    return $tok
}

function Step-Discovery {
    param($Cfg, $Tok)
    Show-Banner 'Step 3/10 - Discovery: DNS zone and subscription (auto)'
    $subs = @((Invoke-Arm GET '/subscriptions' $Tok.Arm -ApiVersion '2022-12-01').value | Where-Object { $_.state -eq 'Enabled' })
    if (-not $subs.Count) { throw 'No enabled subscriptions visible for this account.' }

    $defDomains = @($Cfg.Acme.Domains)
    $domains = Read-List 'Certificate names (first = main domain; wildcard and apex recommended)' $defDomains
    if (-not $domains.Count) { throw 'No certificate names configured.' }
    $Cfg.Acme.Domains = $domains

    # If a subscription is pinned in config, search only there; otherwise search all subscriptions
    # and auto-select the one that actually holds the DNS zone for the domain.
    # plain-if (not $x = if(){}else{@()}): the expression form yields $null for the empty/single
    # case here, which then fails on .Count under Set-StrictMode.
    $subScope = @()
    if ($Cfg.Azure.SubscriptionId) { $subScope = @($subs | Where-Object { $_.subscriptionId -eq $Cfg.Azure.SubscriptionId }) }
    if (-not $subScope.Count) { $subScope = $subs }

    $allZones = @()
    foreach ($s in $subScope) {
        # a stale / cross-tenant subscription that ARM cannot resolve returns 400 - skip it and
        # keep searching the others instead of aborting the whole discovery.
        try {
            $zs = @((Invoke-Arm GET "/subscriptions/$($s.subscriptionId)/providers/Microsoft.Network/dnszones" $Tok.Arm -ApiVersion '2018-05-01').value)
        } catch {
            Write-Log "Skipping subscription $($s.subscriptionId) ($($s.displayName)): $($_.Exception.Message)" WARN
            continue
        }
        foreach ($z in $zs) { $allZones += [pscustomobject]@{ Name = $z.name; Id = $z.id; Sub = $s } }
    }
    if (-not $allZones.Count) { throw 'No Azure DNS zones visible for this account.' }

    $zoneIds = @(); $chosenSub = $null
    foreach ($d in $domains) {
        $bare = $d -replace '^\*\.', ''
        $match = $allZones | Where-Object { $bare -eq $_.Name -or $bare.EndsWith(".$($_.Name)") } |
            Sort-Object { $_.Name.Length } -Descending | Select-Object -First 1
        if (-not $match) { throw "No Azure DNS zone found for '$d' (searched $($subScope.Count) subscription(s))." }
        if ($zoneIds -notcontains $match.Id) { $zoneIds += $match.Id }
        if (-not $chosenSub) { $chosenSub = $match.Sub }
        Write-Log "'$d' -> zone '$($match.Name)' in subscription '$($match.Sub.displayName)'" OK
    }
    $Cfg.Azure.SubscriptionId = $chosenSub.subscriptionId
    $Cfg.Azure.DnsZoneIds = $zoneIds
    Write-Log "Subscription auto-selected from the DNS zone: $($chosenSub.displayName) ($($chosenSub.subscriptionId))" OK

    $Cfg.Acme.Contact = Read-Default 'ACME account contact e-mail' $Cfg.Acme.Contact
    if (-not $Cfg.Acme.Contact) { throw 'Contact e-mail is required.' }
}

function Step-AppRegistration {
    param($Cfg, $Tok)
    Show-Banner 'Step 4/10 - App registration'
    $mainZone = ($Cfg.Azure.DnsZoneIds[0] -split '/')[-1]
    $Cfg.Azure.AppDisplayName = Read-Default 'App registration display name' $(if ($Cfg.Azure.AppDisplayName) { $Cfg.Azure.AppDisplayName } else { "ACME-DNS-$mainZone" })
    $name = $Cfg.Azure.AppDisplayName
    $existing = @((Invoke-Graph GET "/applications?`$filter=displayName eq '$($name -replace "'", "''")'" $Tok.Graph).value)
    if ($existing.Count) {
        $app = $existing[0]
        Write-Log "Reusing existing app '$name' (appId $($app.appId))" OK
    } else {
        $app = Invoke-Graph POST '/applications' $Tok.Graph -Body @{ displayName = $name; signInAudience = 'AzureADMyOrg'; notes = 'Created by Invoke-AcmeExchangeCert. DNS-01 validation for Let''s Encrypt via Posh-ACME.' }
        Write-Log "Created app '$name' (appId $($app.appId))" OK
    }
    $Cfg.Azure.AppId = $app.appId
    $Cfg.Azure.AppObjectId = $app.id

    $sp = @((Invoke-Graph GET "/servicePrincipals?`$filter=appId eq '$($app.appId)'" $Tok.Graph).value)
    if ($sp.Count) { $sp = $sp[0]; Write-Log "Service principal exists (objectId $($sp.id))" OK }
    else {
        $sp = Invoke-Graph POST '/servicePrincipals' $Tok.Graph -Body @{ appId = $app.appId } -Retries 5
        Write-Log "Service principal created (objectId $($sp.id))" OK
    }
    $Cfg.Azure.SpObjectId = $sp.id
}

function Step-AuthCertificate {
    param($Cfg, $Tok)
    Show-Banner 'Step 5/10 - App credential: certificate (no secret)'
    $cert = $null
    if ($Cfg.Azure.AuthCertThumbprint) {
        $cert = Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.Thumbprint -eq $Cfg.Azure.AuthCertThumbprint -and $_.NotAfter -gt (Get-Date).AddDays(90) -and $_.HasPrivateKey }
    }
    if ($cert) {
        Write-Log "Reusing auth certificate $($cert.Thumbprint) (valid until $($cert.NotAfter))" OK
    } else {
        $subject = "CN=$($Cfg.Azure.AppDisplayName)"
        # Legacy CSP (not CNG): the Posh-ACME Azure plugin signs the client assertion via the
        # X509Certificate2.PrivateKey property, which returns $null for CNG keys on .NET Framework
        # (Windows PowerShell 5.1). The "...AES..." provider (PROV_RSA_AES) is required for SHA256/RS256;
        # the base "Enhanced Cryptographic Provider v1.0" cannot do SHA256. This is deliberately the
        # opposite of the server TLS certificate, which must live in the CNG KSP.
        $cert = New-SelfSignedCertificate -Subject $subject -CertStoreLocation Cert:\LocalMachine\My `
            -KeyExportPolicy NonExportable -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 `
            -Provider 'Microsoft Enhanced RSA and AES Cryptographic Provider' -KeySpec Signature `
            -KeyUsage DigitalSignature `
            -NotAfter (Get-Date).AddYears(3) -FriendlyName "$($Cfg.Azure.AppDisplayName) auth (Invoke-AcmeExchangeCert)"
        $Cfg.Azure.AuthCertThumbprint = $cert.Thumbprint
        $script:AuthKeyChanged = $true
        Write-Log "Created auth certificate $($cert.Thumbprint) in LocalMachine\My (legacy CSP for Posh-ACME signing, non-exportable, until $($cert.NotAfter))" OK
    }

    # Register the public key on the app. This is a dedicated single-purpose app, so we set
    # keyCredentials to exactly this one certificate rather than appending: appending would require
    # re-sending existing entries with key=null, which Entra rejects/drops - which would break the
    # 3-year auth-cert rollover. Replacing is idempotent and self-healing.
    $tpB64 = [Convert]::ToBase64String($cert.GetCertHash())
    $app = Invoke-Graph GET "/applications/$($Cfg.Azure.AppObjectId)?`$select=keyCredentials" $Tok.Graph
    $present = @(@($app.keyCredentials) | Where-Object { $_ -and $_.customKeyIdentifier -eq $tpB64 })
    if ($present.Count -and @($app.keyCredentials).Count -eq 1) {
        Write-Log 'Public key already registered on the app.' OK
    } else {
        $keys = @(@{
            type = 'AsymmetricX509Cert'; usage = 'Verify'
            key = [Convert]::ToBase64String($cert.GetRawCertData())
            displayName = "Invoke-AcmeExchangeCert on $(Get-LocalServerName)"
            startDateTime = $cert.NotBefore.ToUniversalTime().ToString('o')
            endDateTime   = $cert.NotAfter.ToUniversalTime().ToString('o')
            customKeyIdentifier = $tpB64
        })
        Invoke-Graph PATCH "/applications/$($Cfg.Azure.AppObjectId)" $Tok.Graph -Body @{ keyCredentials = $keys } | Out-Null
        $script:AuthKeyChanged = $true
        Write-Log 'Public key registered on the app registration (keyCredentials set to this certificate).' OK
    }
}

function Step-Role {
    param($Cfg, $Tok)
    Show-Banner 'Step 6/10 - Least privilege: DNS TXT role on the zone(s)'
    $sub = $Cfg.Azure.SubscriptionId
    $roleName = $Cfg.Azure.RoleName
    $subScope = "/subscriptions/$sub"
    $defs = @((Invoke-Arm GET "$subScope/providers/Microsoft.Authorization/roleDefinitions?`$filter=roleName eq '$roleName'" $Tok.Arm -ApiVersion '2022-04-01').value)
    if ($defs.Count) {
        $role = $defs[0]
        Write-Log "Role '$roleName' exists ($($role.id))" OK
    } else {
        $roleId = [guid]::NewGuid().ToString()
        $body = @{ properties = @{
            roleName = $roleName; type = 'CustomRole'
            description = 'Read DNS zones and manage TXT records only (ACME DNS-01 validation).'
            permissions = @(@{ actions = @('Microsoft.Network/dnsZones/read', 'Microsoft.Network/dnsZones/TXT/*', 'Microsoft.Resources/subscriptions/resourceGroups/read'); notActions = @() })
            assignableScopes = @($subScope)
        } }
        $role = Invoke-Arm PUT "$subScope/providers/Microsoft.Authorization/roleDefinitions/$roleId" $Tok.Arm -Body $body -ApiVersion '2022-04-01'
        Write-Log "Role '$roleName' created ($($role.id))" OK
    }

    foreach ($zoneId in $Cfg.Azure.DnsZoneIds) {
        $zoneName = ($zoneId -split '/')[-1]
        $asg = @((Invoke-Arm GET "$zoneId/providers/Microsoft.Authorization/roleAssignments?`$filter=principalId eq '$($Cfg.Azure.SpObjectId)'" $Tok.Arm -ApiVersion '2022-04-01').value)
        $have = @($asg | Where-Object { $_ -and $_.properties.roleDefinitionId -like "*$($role.name)" -and $_.properties.scope -eq $zoneId })
        if ($have.Count) { Write-Log "Assignment on zone $zoneName exists." OK; continue }
        $aid = [guid]::NewGuid().ToString()
        $body = @{ properties = @{ roleDefinitionId = $role.id; principalId = $Cfg.Azure.SpObjectId; principalType = 'ServicePrincipal' } }
        # SP propagation can take a while -> retries
        Invoke-Arm PUT "$zoneId/providers/Microsoft.Authorization/roleAssignments/$aid" $Tok.Arm -Body $body -ApiVersion '2022-04-01' -Retries 8 | Out-Null
        Write-Log "Assigned '$roleName' to the app on zone $zoneName." OK
    }
}

function Step-ExchangeParams {
    param($Cfg)
    Show-Banner 'Step 7/10 - Exchange parameters'
    Connect-ExchangeManagement
    $all = Get-ExchangeServer | Where-Object { $_.ServerRole -match 'Mailbox' -and $_.AdminDisplayVersion.Major -ge 15 } | ForEach-Object { $_.Name.ToUpper() }
    Write-Host "  Mailbox servers in the organization: $($all -join ', ')"
    $def = @($Cfg.Exchange.Servers); if (-not $def.Count) { $def = @((Get-LocalServerName)) }
    $Cfg.Exchange.Servers = Read-List 'Servers to install the certificate on (this server is handled first)' $def
    foreach ($s in $Cfg.Exchange.Servers) { if ($all -notcontains $s.ToUpper()) { throw "'$s' is not a mailbox server in this organization." } }
    $Cfg.Exchange.Services = Read-Default 'Exchange services (IIS,SMTP recommended; also IMAP,POP)' $Cfg.Exchange.Services
    $Cfg.Exchange.PrivateKeyExportable = Read-YesNo 'Mark private key exportable (allows repair/export later)?' ([bool]$Cfg.Exchange.PrivateKeyExportable)
    $Cfg.Exchange.UpdateConnectors = Read-YesNo 'Update TlsCertificateName on send/receive connectors that reference this subject?' ([bool]$Cfg.Exchange.UpdateConnectors)
    $Cfg.Exchange.RemoveOldCertificate = Read-YesNo 'Remove superseded certificates with the same subject after successful install?' ([bool]$Cfg.Exchange.RemoveOldCertificate)
    $Cfg.Exchange.LoopbackBinding = Read-YesNo 'Maintain a manual 127.0.0.1:443 http.sys binding if one exists?' ([bool]$Cfg.Exchange.LoopbackBinding)
    foreach ($s in $Cfg.Exchange.Servers) {
        $r = Resolve-ExchangeServer $s
        if (-not $r.IsLocal) {
            $w = Test-WinRm $r.Fqdn
            Write-Log "[$($r.Name)] remote; WinRM reachable: $w $(if (-not $w) { '(loopback binding and KSP fallback will be skipped there)' })" $(if ($w) { 'OK' } else { 'WARN' })
        }
    }
}

function Step-Notify {
    param($Cfg)
    Show-Banner 'Step 8/10 - Notifications'
    $Cfg.Notify.SmtpServer = Read-Default 'SMTP server' $Cfg.Notify.SmtpServer
    $Cfg.Notify.Port = [int](Read-Default 'SMTP port' $Cfg.Notify.Port)
    $Cfg.Notify.UseSsl = Read-YesNo 'Use STARTTLS/SSL?' ([bool]$Cfg.Notify.UseSsl)
    $Cfg.Notify.From = Read-Default 'Sender address' $(if ($Cfg.Notify.From) { $Cfg.Notify.From } else { "acme@$(($Cfg.Acme.Domains[0]) -replace '^\*\.', '')" })
    $Cfg.Notify.To = Read-List 'Recipient address(es)' $Cfg.Notify.To
    $Cfg.Notify.WarnDaysBeforeExpiry = [int](Read-Default 'Warn when bound certificate expires within N days' $Cfg.Notify.WarnDaysBeforeExpiry)
    $Cfg.Notify.SendSuccessMail = Read-YesNo 'Send a mail after each successful renewal?' ([bool]$Cfg.Notify.SendSuccessMail)
    $script:Config = $Cfg
    if (Read-YesNo 'Send a test mail now?' $true) {
        Send-Notification -Subject 'Test notification' -Body "This is a test from the setup wizard on $(Get-LocalServerName)." -Level Info
    }
}

function Step-FirstCertificate {
    param($Cfg)
    Show-Banner 'Step 9/10 - First certificate'
    $script:Config = $Cfg
    Save-Config $Cfg
    Initialize-PoshAcme
    $main = Get-MainDomain
    $existing = Get-PAOrderSafe -MainDomain $main
    if ($existing -and $existing.status -eq 'valid' -and -not (Read-YesNo "An order for $main already exists (expires $($existing.CertExpires)). Request a new certificate anyway?" $false)) {
        Write-Log 'Keeping existing order. Use -Renew / -ForceRenew later.' INFO
        # still make sure plugin args are current
        Set-PAOrder -MainDomain $main -PluginArgs (Get-AzurePluginArgs $Cfg) | Out-Null
        return
    }
    if (-not (Get-PAAccount)) {
        New-PAAccount -Contact $Cfg.Acme.Contact -AcceptTOS | Out-Null
        Write-Log "ACME account created for $($Cfg.Acme.Contact)" OK
    }
    $pfxPass = -join ((48..57) + (65..90) + (97..122) | Get-Random -Count 24 | ForEach-Object { [char]$_ })
    $params = @{
        Domain        = $Cfg.Acme.Domains
        Plugin        = 'Azure'
        PluginArgs    = Get-AzurePluginArgs $Cfg
        Contact       = $Cfg.Acme.Contact
        AcceptTOS     = $true
        PfxPass       = $pfxPass
        FriendlyName  = "$main (Let's Encrypt)"
        CertKeyLength = $Cfg.Acme.KeyLength
        Force         = $true
        Verbose       = $false
    }
    if ($script:AuthKeyChanged) {
        Write-Log 'A new app credential was just registered; waiting 30s for Entra ID to propagate it before requesting the certificate ...' INFO
        Start-Sleep -Seconds 30
    }
    Write-Log "Requesting certificate for $($Cfg.Acme.Domains -join ', ') via Azure DNS-01 ..." STEP
    $cert = $null
    for ($try = 1; $try -le 4; $try++) {
        try { $cert = New-PACertificate @params; break }
        catch {
            # 401/unauthorized right after a credential change is Entra propagation lag - wait and retry
            if ($try -lt 4 -and ($_.Exception.Message -match '401|Unauthorized|AADSTS')) {
                Write-Log "Certificate request failed with an auth error (attempt $try/4), likely credential propagation - retrying in 30s ..." WARN
                Start-Sleep -Seconds 30
                continue
            }
            throw
        }
    }
    if (-not $cert) { throw 'New-PACertificate returned nothing.' }
    Write-Log "Issued: $($cert.Thumbprint), valid until $($cert.NotAfter)" OK
    if ($script:SkipInstall) {
        Write-Log "SkipInstall set: certificate issued but NOT installed on any server (rehearsal). PFX in the Posh-ACME state." WARN
        Write-Log "Do not point production servers at a staging certificate - run production setup to install for real." WARN
        return
    }
    $result = Install-Certificate -PACert $cert
    Write-Log "Installed on $($result.Servers -join ', '); connectors changed: $($result.Connectors.Count); removed: $($result.Removed.Count); fallback: $($result.Fallback -join ',')" OK
    Write-EventLogEntry -Message "Initial certificate $($cert.Thumbprint) installed by setup wizard." -Type Information -Id 9000
}

function Get-AzurePluginArgs {
    param($Cfg)
    # AZSubscriptionId must never be empty. If it is (e.g. a config edited/overwritten without it),
    # fall back to the subscription embedded in the DNS zone resource id.
    $sub = [string]$Cfg.Azure.SubscriptionId
    if (-not $sub -and @($Cfg.Azure.DnsZoneIds).Count) {
        $sub = (@($Cfg.Azure.DnsZoneIds)[0] -split '/')[2]
    }
    @{
        AZSubscriptionId  = $sub
        AZTenantId        = $Cfg.Azure.TenantId
        AZAppUsername     = $Cfg.Azure.AppId
        AZCertThumbprint  = $Cfg.Azure.AuthCertThumbprint
    }
}

function Get-DnsProviderName {
    $p = if ($script:Config.Acme.PSObject.Properties['DnsProvider'] -and $script:Config.Acme.DnsProvider) { $script:Config.Acme.DnsProvider } else { 'Azure' }
    return $p
}

function Resolve-PluginArgs {
    <#
      Plugin arguments for the configured DNS provider. Azure derives them from the app + auth
      certificate (no secret). Other providers use their non-secret args from config plus any secret
      values passed transiently via the ACME_PLUGIN_SECRETS file (written by the GUI, consumed once);
      after the first issuance Posh-ACME keeps the whole set encrypted in the order, so later renewals
      need nothing here. Returns $null when a non-Azure provider has no args to (re)apply this run.
    #>
    $provider = Get-DnsProviderName
    if ($provider -eq 'Azure') { return Get-AzurePluginArgs $script:Config }

    $def = (Get-DnsProviders)[$provider]
    if (-not $def) { throw "Unknown DNS provider '$provider'." }

    $args = @{}
    # non-secret args stored in config
    if ($script:Config.Acme.PSObject.Properties['DnsPluginArgs'] -and $script:Config.Acme.DnsPluginArgs) {
        foreach ($p in $script:Config.Acme.DnsPluginArgs.PSObject.Properties) { $args[$p.Name] = $p.Value }
    }
    # transient secrets from the GUI (plaintext JSON, consumed and deleted); Secure fields -> SecureString
    if ($env:ACME_PLUGIN_SECRETS -and (Test-Path $env:ACME_PLUGIN_SECRETS)) {
        try {
            $sec = Get-Content $env:ACME_PLUGIN_SECRETS -Raw | ConvertFrom-Json
            foreach ($f in $def.Fields) {
                if ($sec.PSObject.Properties[$f.Key] -and "$($sec.$($f.Key))".Length) {
                    $args[$f.Key] = if ($f.Secure) { ConvertTo-SecureString -String ([string]$sec.$($f.Key)) -AsPlainText -Force } else { [string]$sec.$($f.Key) }
                }
            }
        } finally { Remove-Item $env:ACME_PLUGIN_SECRETS -Force -ErrorAction SilentlyContinue }
    }
    if (-not $args.Keys.Count) { return $null }
    return $args
}

function Invoke-SetupWizard {
    Show-Banner "Invoke-AcmeExchangeCert v$script:ScriptVersion - Setup"
    Write-Host '  This wizard is re-runnable. Existing Azure resources and settings are detected and reused.'
    Write-Host "  Config file: $ConfigPath"
    if ($Staging) { Write-Host '  Using Let''s Encrypt STAGING.' -ForegroundColor Yellow }

    $cfg = Get-DefaultConfig
    $old = Read-Config
    if ($old) {
        # merge stored values over defaults
        foreach ($sec in $cfg.Keys) {
            if ($sec -eq 'Version') { continue }
            if ($old.PSObject.Properties[$sec]) {
                foreach ($k in @($cfg[$sec].Keys)) {
                    if ($old.$sec.PSObject.Properties[$k]) { $cfg[$sec][$k] = $old.$sec.$k }
                }
            }
        }
        Write-Log 'Existing configuration loaded as defaults.' INFO
    }
    if ($Staging) { $cfg.Acme.Server = 'LE_STAGE' } else { $cfg.Acme.Server = 'LE_PROD' }

    $provider = if ($cfg.Acme.DnsProvider) { $cfg.Acme.DnsProvider } else { 'Azure' }
    Step-Preflight $cfg
    if ((Get-DnsProviders)[$provider].Bootstrap) {
        # Azure: sign in and provision app registration, auth certificate and DNS role
        $tok = Step-Login $cfg
        Step-Discovery $cfg $tok
        Step-AppRegistration $cfg $tok
        Step-AuthCertificate $cfg $tok
        Step-Role $cfg $tok
    } else {
        # Other providers need no cloud bootstrap - their API credentials are supplied at issue time
        # (GUI) and then kept encrypted in the Posh-ACME order.
        Write-Log "DNS provider '$provider' selected - no cloud bootstrap needed. Credentials are entered in the GUI and stored encrypted by Posh-ACME on first issuance." INFO
    }
    Step-ExchangeParams $cfg
    Step-Notify $cfg
    Save-Config $cfg

    # Setup prepares the DNS/cloud side and the configuration only. Issuing and installing the
    # certificate is a separate, repeatable step (GUI step 2 / -Renew), which is also what the
    # scheduled task runs on its first execution.
    Show-Banner 'Setup complete'
    Write-Host '  Azure app, credential, role and configuration are ready. No certificate was issued yet.'
    Write-Host ''
    Write-Host '  >>> NEXT: create the scheduled task now. <<<' -ForegroundColor Yellow
    Write-Host '      It issues and installs the certificate on its first run, then renews automatically.' -ForegroundColor Yellow
    Write-Host "      GUI: create the scheduled task     CLI: .\$(Split-Path $PSCommandPath -Leaf) -InstallTask"
    Write-Host ''
    Write-Host '  To confirm it works right away (recommended) - run it once and check the result:'
    Write-Host "      Start-ScheduledTask -TaskName '$($cfg.Task.Name)'"
    Write-Host "      .\$(Split-Path $PSCommandPath -Leaf) -Status"
    Write-Host '  Issuing + installing by hand is optional (the task does it for you):'
    Write-Host "      GUI: step 2 (get + install)        CLI: .\$(Split-Path $PSCommandPath -Leaf) -Renew"
    Write-Host "  Config:       $ConfigPath (contains no secrets)"
    Write-Host "  Posh-ACME:    $(Get-PoshAcmeHome) (contains private keys - keep the ACL tight)"
}
#endregion

#region ---------------------------------------------------------------- Teardown
function Invoke-Teardown {
    # Remove everything the tool created, mirroring -Setup, for a completely fresh start.
    # Needs an interactive admin sign-in for the Azure side (the app's own credential is not
    # allowed to delete itself). The certificate already bound on the Exchange servers is left
    # in place - removing it would break TLS until a fresh -Renew installs a new one.
    Show-Banner "Invoke-AcmeExchangeCert v$script:ScriptVersion - TEARDOWN"
    $cfg = $script:Config
    if (-not $cfg) {
        Write-Log "No configuration found at '$ConfigPath'. Teardown needs the real config to know what to remove - nothing was changed." ERROR
        Write-Host "  Pass the real config, e.g.:  .\Invoke-AcmeExchangeCert.ps1 -Teardown -ConfigPath C:\Tools\AcmeExchange\config.json"
        return
    }
    $homeDir  = if ($cfg.Paths.Home) { $cfg.Paths.Home } else { 'C:\Tools\AcmeExchange' }
    $taskName = if ($cfg -and $cfg.Task.Name)  { $cfg.Task.Name }  else { 'ACME Exchange Certificate Renewal' }
    $prov = if ($cfg -and $cfg.Acme.PSObject.Properties['DnsProvider'] -and $cfg.Acme.DnsProvider) { $cfg.Acme.DnsProvider } else { 'Azure' }
    $isAzure = ($prov -eq 'Azure')

    Write-Host '  This removes everything the tool created (as if it had never been used):'
    if ($cfg -and $isAzure -and $cfg.Azure.AppObjectId) {
        Write-Host "    - Entra app registration + service principal ($($cfg.Azure.AppDisplayName), appId $($cfg.Azure.AppId))"
        Write-Host "    - role assignment(s) on the DNS zone(s) and the custom role '$($cfg.Azure.RoleName)'"
        Write-Host "    - local auth certificate $($cfg.Azure.AuthCertThumbprint)"
    }
    Write-Host "    - scheduled task '$taskName'"
    Write-Host '    - POSHACME_HOME environment variable'
    Write-Host "    - state folder '$homeDir' (moved aside as a backup, not deleted)"
    Write-Host '  The certificate already bound on the Exchange servers is left untouched.'
    Write-Host ''
    if (-not $Yes -and -not (Read-YesNo 'Proceed with the complete teardown? This cannot be undone.' $false)) {
        Write-Host 'Aborted - nothing was changed.'
        return
    }

    # ---- Azure (interactive admin sign-in) ----
    $azureClean = $true
    if ($isAzure -and $cfg.Azure.AppObjectId) {
        try {
            $tok = Step-Login $cfg
            foreach ($zoneId in @($cfg.Azure.DnsZoneIds)) {
                try {
                    $asg = @((Invoke-Arm GET "$zoneId/providers/Microsoft.Authorization/roleAssignments?`$filter=principalId eq '$($cfg.Azure.SpObjectId)'" $tok.Arm -ApiVersion '2022-04-01').value)
                    foreach ($a in @($asg | Where-Object { $_ -and $_.properties.scope -eq $zoneId })) {
                        try { Invoke-Arm DELETE $a.id $tok.Arm -ApiVersion '2022-04-01' | Out-Null; Write-Log "Removed role assignment on zone $(($zoneId -split '/')[-1])." OK }
                        catch { Write-Log "Assignment delete failed: $($_.Exception.Message)" WARN }
                    }
                } catch { Write-Log "Listing assignments on a zone failed: $($_.Exception.Message)" WARN }
            }
            try {
                $sub = $cfg.Azure.SubscriptionId
                $defs = @((Invoke-Arm GET "/subscriptions/$sub/providers/Microsoft.Authorization/roleDefinitions?`$filter=roleName eq '$($cfg.Azure.RoleName)'" $tok.Arm -ApiVersion '2022-04-01').value)
                foreach ($d in @($defs | Where-Object { $_ -and $_.properties.type -eq 'CustomRole' })) {
                    # remove EVERY assignment that still references this custom role (other scopes / orphaned
                    # after deleting the SP), otherwise the role cannot be deleted
                    try {
                        $all = @((Invoke-Arm GET "/subscriptions/$sub/providers/Microsoft.Authorization/roleAssignments" $tok.Arm -ApiVersion '2022-04-01').value)
                        foreach ($a in @($all | Where-Object { $_ -and $_.properties.roleDefinitionId -eq $d.id })) {
                            try { Invoke-Arm DELETE $a.id $tok.Arm -ApiVersion '2022-04-01' | Out-Null; Write-Log "Removed lingering role assignment ($($a.properties.scope))." OK }
                            catch { Write-Log "Lingering assignment delete failed: $($_.Exception.Message)" WARN }
                        }
                    } catch { Write-Log "Listing role assignments for cleanup failed: $($_.Exception.Message)" WARN }
                    # RBAC is eventually consistent - retry the role deletion a few times after removing assignments
                    $done = $false
                    for ($i = 1; $i -le 6 -and -not $done; $i++) {
                        try { Invoke-Arm DELETE $d.id $tok.Arm -ApiVersion '2022-04-01' | Out-Null; Write-Log "Removed custom role '$($cfg.Azure.RoleName)'." OK; $done = $true }
                        catch { if ($i -lt 6) { Start-Sleep -Seconds 10 } else { $azureClean = $false; Write-Log "Role definition delete failed after retries: $($_.Exception.Message)" WARN } }
                    }
                }
            } catch { $azureClean = $false; Write-Log "Role definition cleanup failed: $($_.Exception.Message)" WARN }
            if ($cfg.Azure.SpObjectId) {
                try { Invoke-Graph DELETE "/servicePrincipals/$($cfg.Azure.SpObjectId)" $tok.Graph | Out-Null; Write-Log 'Removed service principal.' OK }
                catch { if ($_.Exception.Message -match '\(404\)|NotFound') { Write-Log 'Service principal already gone.' OK } else { $azureClean = $false; Write-Log "Service principal delete: $($_.Exception.Message)" WARN } }
            }
            try { Invoke-Graph DELETE "/applications/$($cfg.Azure.AppObjectId)" $tok.Graph | Out-Null; Write-Log 'Removed app registration.' OK }
            catch { if ($_.Exception.Message -match '\(404\)|NotFound') { Write-Log 'App registration already gone.' OK } else { $azureClean = $false; Write-Log "App registration delete: $($_.Exception.Message)" WARN } }
        } catch {
            $azureClean = $false
            Write-Log "Azure sign-in or cleanup failed: $($_.Exception.Message)" WARN
            Write-Log 'Remove the app registration, the role assignment and the custom role in the Azure portal manually.' WARN
        }
    } else {
        Write-Log 'No Azure app registration in config (or non-Azure provider) - skipping the Azure teardown.' INFO
    }

    # ---- Local ----
    if ($cfg -and $cfg.Azure.AppDisplayName) {
        $tp = [string]$cfg.Azure.AuthCertThumbprint
        $subj = "CN=$($cfg.Azure.AppDisplayName)"
        foreach ($c in @(Get-ChildItem Cert:\LocalMachine\My -ErrorAction SilentlyContinue | Where-Object { ($tp -and $_.Thumbprint -eq $tp) -or $_.Subject -eq $subj })) {
            try { Remove-Item $c.PSPath -Force; Write-Log "Removed auth certificate $($c.Thumbprint) ($($c.Subject))." OK }
            catch { Write-Log "Certificate remove failed: $($_.Exception.Message)" WARN }
        }
    }
    if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
        try { Unregister-ScheduledTask -TaskName $taskName -Confirm:$false; Write-Log "Removed scheduled task '$taskName'." OK }
        catch { Write-Log "Task remove failed: $($_.Exception.Message)" WARN }
    }
    try {
        if ([Environment]::GetEnvironmentVariable('POSHACME_HOME', 'Machine')) {
            [Environment]::SetEnvironmentVariable('POSHACME_HOME', $null, 'Machine')
            Write-Log 'Removed POSHACME_HOME machine environment variable.' OK
        }
    } catch { Write-Log "POSHACME_HOME remove failed: $($_.Exception.Message)" WARN }
    if ($azureClean) {
        if (Test-Path $homeDir) {
            $bak = "$homeDir.removed-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
            try { Move-Item -LiteralPath $homeDir -Destination $bak; Write-Log "Moved state folder to '$bak' (delete it once you are satisfied)." OK }
            catch { Write-Log "Could not move '$homeDir': $($_.Exception.Message)" WARN }
        }
        if ($ConfigPath -and (Test-Path $ConfigPath)) {
            try { Remove-Item -LiteralPath $ConfigPath -Force -ErrorAction SilentlyContinue } catch { }
        }
    } else {
        Write-Log 'Azure was NOT fully cleaned up - keeping config.json and the state folder so you can re-run -Teardown after fixing the Azure side.' WARN
    }

    Show-Banner 'Teardown complete'
    if ($azureClean) {
        Write-Host '  Everything the tool created has been removed. Run -Setup for a completely fresh start.'
    } else {
        Write-Host '  Local parts removed, but the Azure cleanup was incomplete (see the warnings above). Config and state were kept for a re-run.'
    }
    Write-Host '  The Exchange servers keep their currently bound certificate until a fresh -Renew installs a new one.'
}
#endregion

#region ---------------------------------------------------------------- Main
try {
    if ($PSCmdlet.ParameterSetName -eq 'Setup') {
        Invoke-SetupWizard
        exit 0
    }
    if ($PSCmdlet.ParameterSetName -eq 'Teardown') {
        $script:Config = Read-Config
        if ($script:Config) { try { Initialize-Logging $script:Config.Paths.Home } catch { } }
        Invoke-Teardown
        exit 0
    }

    $script:Config = Read-Config
    if (-not $script:Config) { throw "No configuration at $ConfigPath. Run with -Setup first." }
    Initialize-Logging $script:Config.Paths.Home

    switch ($PSCmdlet.ParameterSetName) {
        'Status'         { Show-Status }
        'TestMail'       { Send-Notification -Subject 'Test notification' -Body "Manual test from $(Get-LocalServerName)." -Level Info }
        'InstallTask'    { Register-RenewalTask }
        'RemoveTask'     { Unregister-RenewalTask }
        'SetPfxPassword' { Set-PfxPasswordFile -Password $PfxPassword }
        'ImportPfx'      { Invoke-PfxImport }
        default          { Invoke-Renewal }
    }
    exit 0
} catch {
    $msg = "$($_.Exception.Message)`r`n`r`nAt: $($_.InvocationInfo.PositionMessage)"
    Write-Log $msg ERROR
    if ($script:Config -and $PSCmdlet.ParameterSetName -in @('Renew', 'ImportPfx')) {
        Write-EventLogEntry -Message $msg -Type Error -Id 9001
        Send-Notification -Subject "ERROR during certificate renewal on $(Get-LocalServerName)" -Body $msg -Level Error
    }
    exit 1
}
#endregion
