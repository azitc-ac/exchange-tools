<#
    ExchangeOnlineDocumention_v3.1.ps1
    Purpose: Export a structured Exchange Online documentation set (CSV/TXT) in one run.

    Style:
      - Functions instead of regions
      - English comments
      - PowerShell 5.1 compatible
      - Concise Write-Host logging
      - File is intentionally pure ASCII; the multi-value separator is built from
        a char code so the script is immune to encoding damage.

    Output:
      .\EXO-Documentation\<TenantName>\*.csv / *.txt

    Auto-Connect:
      - If no active Exchange Online session is found, Connect-ExchangeOnline -ShowBanner:$false is invoked.

    Changes vs. v3.0
      1. FIX ManagementObjectAmbiguousException ("... is not a unique recipient"):
         every per-mailbox cmdlet call now uses ExchangeGuid (mailbox scope) or
         PrimarySmtpAddress (recipient scope) instead of Identity/Name, which
         resolve to the display name and are ambiguous as soon as two objects
         share it.
      2. FIX silent data loss: the script functions Get-EXOMailboxStatistics /
         Get-EXOArchiveMailboxStatistics shadowed the EXO cmdlet of the same name
         and called themselves recursively. The parameter binding error was
         swallowed by the surrounding try/catch, so the statistics CSVs stayed
         empty. Renamed to Export-*.
      3. FIX SendAs export: it used $u.Name, which was never requested via
         -Properties and was therefore always $null -> empty SendAs CSV.
      4. FIX folder permissions: the old helper used "break" instead of
         "continue" (the first skipped folder aborted the whole mailbox), split
         the folder identity on the first backslash only (nested paths were
         truncated) and contained a no-op .replace("","/"). Replaced by the
         on-prem v3.2 logic that builds "smtp:\relative\path".
      5. Errors are no longer silently swallowed - every failed object is written
         to <Tenant>_Errors.csv.
      6. Ported from on-prem v3.2: richer mailbox properties, multi-value fields
         joined with a real separator, ConvertTo-FlatObject for wide objects,
         List[T] instead of += (O(n^2)), numeric guard in the size summary,
         distribution groups + members, mail contacts / mail users, resource
         mailbox CalendarProcessing.
#>

param(
    [bool]$IncludeMailboxFolderPermissions = $false,
    [bool]$IncludeMailboxPermissions       = $true,
    [bool]$IncludeRecipients               = $true
)

# Prevent truncation if someone pipes to Format-List / Format-Table later
$FormatEnumerationLimit = -1

# Multi-value separator inside a CSV field (paragraph sign), built from char code
$script:Sep         = [char]0x00A7
$script:CsvDelim    = ';'
$script:CsvEncoding = if ($PSVersionTable.PSVersion.Major -ge 7) { 'utf8BOM' } else { 'UTF8' }  # PS 5.1 UTF8 already has a BOM

# Collects every non-fatal failure so problems are visible instead of swallowed
$script:ErrorLog = New-Object 'System.Collections.Generic.List[PSObject]'

# -----------------------------
# Utility: Simple timers
# -----------------------------
function Start-Stopwatch {
    [Diagnostics.Stopwatch]::StartNew()
}

function Stop-StopwatchString {
    param([Diagnostics.Stopwatch]$Watch)
    if ($null -eq $Watch) { return "0s" }
    $Watch.Stop()
    "{0:D2}m:{1:D2}s" -f $Watch.Elapsed.Minutes, $Watch.Elapsed.Seconds
}

# -----------------------------
# Utility: Console markers
# -----------------------------
function Write-Section {
    param([Parameter(Mandatory=$true)][string]$Text)
    Write-Host ""
    Write-Host ("## " + $Text + " ##") -ForegroundColor Cyan
}

function Add-EXOError {
    param(
        [string]$Scope,
        [string]$Target,
        [string]$Message
    )
    $script:ErrorLog.Add([pscustomobject]@{
        TimeStamp = (Get-Date)
        Scope     = $Scope
        Target    = $Target
        Message   = $Message
    })
}

# -----------------------------
# Utility: Identity resolution
# -----------------------------
# The Identity / Name of a mailbox is the display name in Exchange Online.
# As soon as two directory objects share it (mailbox + mail user, mailbox +
# contact, active + soft-deleted mailbox) every cmdlet that takes -Identity
# fails with ManagementObjectAmbiguousException. ExchangeGuid and
# PrimarySmtpAddress are unique, so we always bind through them.
function Get-EXOMailboxKey {
    param([Parameter(Mandatory=$true)]$Mailbox)

    $guid = $Mailbox.ExchangeGuid
    if ($guid -and ("$guid" -ne '00000000-0000-0000-0000-000000000000')) { return "$guid" }

    if ($Mailbox.ExternalDirectoryObjectId) { return [string]$Mailbox.ExternalDirectoryObjectId }
    if ($Mailbox.PrimarySmtpAddress)        { return [string]$Mailbox.PrimarySmtpAddress }

    return [string]$Mailbox.Identity
}

function Get-EXORecipientKey {
    param([Parameter(Mandatory=$true)]$Mailbox)

    # Recipient-scoped cmdlets (Get-RecipientPermission) resolve SMTP reliably,
    # the mailbox ExchangeGuid is not always accepted there.
    if ($Mailbox.PrimarySmtpAddress) { return [string]$Mailbox.PrimarySmtpAddress }
    return (Get-EXOMailboxKey -Mailbox $Mailbox)
}

# Resolve the fastest available cmdlet (REST-based EXO* variants when present)
function Resolve-EXOCmdlet {
    param(
        [Parameter(Mandatory=$true)][string]$Preferred,
        [Parameter(Mandatory=$true)][string]$Fallback
    )
    if (Get-Command $Preferred -ErrorAction SilentlyContinue) { return $Preferred }
    return $Fallback
}

# -----------------------------
# Utility: Value formatting
# -----------------------------
function ConvertTo-JoinedString {
    param(
        $Value,
        [string]$Separator = $script:Sep
    )
    if ($null -eq $Value) { return "" }
    if ($Value -is [string]) { return $Value }

    if ($Value -is [System.Collections.IEnumerable]) {
        # Safe enumeration - avoids the PS 5.1 SZArrayEnumerator pipeline issue
        $parts = New-Object 'System.Collections.Generic.List[string]'
        try {
            $enum = ([System.Collections.IEnumerable]$Value).GetEnumerator()
            while ($enum.MoveNext()) { $parts.Add([string]$enum.Current) }
        } catch {
            return [string]$Value
        }
        return ($parts -join $Separator)
    }
    return [string]$Value
}

function ConvertTo-FlatObject {
    param(
        [Parameter(ValueFromPipeline)]
        $InputObject,
        [string]$Separator = $script:Sep
    )
    process {
        if ($null -eq $InputObject) { return }
        $hash = [ordered]@{}
        foreach ($prop in $InputObject.PSObject.Properties) {
            $hash[$prop.Name] = ConvertTo-JoinedString -Value $prop.Value -Separator $Separator
        }
        [pscustomobject]$hash
    }
}

function Convert-BytesToSizeString {
    param([Parameter(Mandatory=$true)][long]$Bytes)
    $units = @("B","KB","MB","GB","TB","PB")
    $size  = [double]$Bytes
    $i     = 0
    while ($size -ge 1024 -and $i -lt ($units.Count - 1)) {
        $size = $size / 1024
        $i    = $i + 1
    }
    $sizeStr = [System.String]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:0.##}", $size)
    return "$sizeStr $($units[$i])"
}

# Convert an Exchange size value into bytes. Works for both the typed
# ByteQuantifiedSize (classic cmdlets) and the plain string the REST cmdlets
# return, e.g. "1.5 GB (1,610,612,736 bytes)".
function ConvertFrom-EXOSizeString {
    param($SizeValue)

    if ($null -eq $SizeValue) { return [int64]0 }

    try {
        if ($SizeValue.Value -and ($SizeValue.Value | Get-Member -Name ToBytes -MemberType Method -ErrorAction SilentlyContinue)) {
            return [int64]$SizeValue.Value.ToBytes()
        }
    } catch { }

    $text = [string]$SizeValue
    if ([string]::IsNullOrWhiteSpace($text)) { return [int64]0 }

    $m = [regex]::Match($text, '\((?<bytes>[0-9,\.\s]+)\s*bytes\)', 'IgnoreCase')
    if ($m.Success) {
        $raw = $m.Groups['bytes'].Value -replace '[^0-9]', ''
        if ($raw) { return [int64]$raw }
    }

    $n = $text -replace '[^0-9]', ''
    if ($n) { return [int64]$n }

    return [int64]0
}

function ConvertTo-Int64Safe {
    param($Value)
    if ($null -eq $Value) { return [int64]0 }
    $s = [string]$Value
    if ([string]::IsNullOrWhiteSpace($s)) { return [int64]0 }
    $s = ($s -replace '[^0-9\-]', '')
    $out = [int64]0
    [void][int64]::TryParse($s, [ref]$out)
    return $out
}

function Get-ProgressPercent {
    param([int]$Current, [int]$Total)
    if ($Total -le 0) { return 0 }
    $p = [int](($Current / $Total) * 100)
    if ($p -lt 0)   { return 0 }
    if ($p -gt 100) { return 100 }
    return $p
}

# -----------------------------
# Utility: Ensure EXO connection
# -----------------------------
function Connect-EXOIfNeeded {
    $connected = $false

    # EXO V3 does not create a PSSession in REST mode - Get-ConnectionInformation
    # is the only reliable check there.
    if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
        try {
            $ci = Get-ConnectionInformation -ErrorAction SilentlyContinue |
                  Where-Object { $_.State -eq 'Connected' -and $_.TokenStatus -ne 'Expired' }
            if ($ci) { $connected = $true }
        } catch { }
    }

    if (-not $connected) {
        try {
            $sess = Get-PSSession -ErrorAction SilentlyContinue |
                    Where-Object { $_.ConfigurationName -eq 'Microsoft.Exchange' -and $_.State -eq 'Opened' }
            if ($sess) { $connected = $true }
        } catch { }
    }

    if (-not $connected) {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
        try {
            Connect-ExchangeOnline -ShowBanner:$false
            Write-Host "Connected to Exchange Online." -ForegroundColor Green
        } catch {
            Write-Host "Failed to connect to Exchange Online: $($_.Exception.Message)" -ForegroundColor Red
            throw
        }
    } else {
        Write-Host "Exchange Online session already available." -ForegroundColor Green
    }
}

# -----------------------------
# Utility: Create export path
# -----------------------------
function New-EXOExportPath {
    param([Parameter(Mandatory=$true)][string]$BaseName)

    if ($PSScriptRoot) {
        $root = Join-Path -Path $PSScriptRoot -ChildPath $BaseName
    } else {
        $root = Join-Path -Path "C:\Tools\Scripts" -ChildPath $BaseName
    }

    if (-not (Test-Path -LiteralPath $root)) {
        [void](New-Item -Path $root -ItemType Directory -Force)
    }
    return $root
}

# -----------------------------
# Helper: CSV / TXT export
# -----------------------------
function Export-EXOCsv {
    param(
        $InputObject,
        [Parameter(Mandatory=$true)][string]$Path,
        [switch]$Quiet
    )
    $dir = Split-Path -Path $Path -Parent
    if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -Path $dir -ItemType Directory -Force) }

    $items = @($InputObject | Where-Object { $null -ne $_ })
    if ($items.Count -gt 0) {
        $items | Export-Csv -Path $Path -NoTypeInformation -Delimiter $script:CsvDelim -Encoding $script:CsvEncoding -Force
    } else {
        # Keep a marker file so a missing file always means "step did not run"
        [pscustomobject]@{ Info = "no data" } |
            Export-Csv -Path $Path -NoTypeInformation -Delimiter $script:CsvDelim -Encoding $script:CsvEncoding -Force
    }
    if (-not $Quiet) {
        Write-Host ("    -> " + (Split-Path -Path $Path -Leaf) + "  (" + $items.Count + " records)")
    }
}

function Export-EXOTxt {
    param(
        $InputObject,
        [Parameter(Mandatory=$true)][string]$Path
    )
    $dir = Split-Path -Path $Path -Parent
    if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -Path $dir -ItemType Directory -Force) }
    $InputObject | Out-File -FilePath $Path -Encoding utf8 -Force
}

# -----------------------------
# Function: Get tenant info (basic)
# -----------------------------
function Get-EXOTenantInfo {
    <#
        .SYNOPSIS
            Retrieves basic tenant info used for path naming and context.
    #>
    Write-Section "Getting basic tenant data"
    $watch = Start-Stopwatch

    try {
        $org        = Get-OrganizationConfig
        $tenantName = $org.Name
        Write-Host (" Tenant: " + $tenantName)
        Write-Host (" Finished in " + (Stop-StopwatchString $watch)) -ForegroundColor Green

        [pscustomobject]@{
            Name = $tenantName
            Raw  = $org
        }
    } catch {
        Write-Host "Failed to get organization configuration: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

# =============================
# PART 2: Mailbox data & permissions
# =============================

function Get-EXOMailboxData {
    param([Parameter(Mandatory=$true)][string]$ExportPath)

    Write-Section "Getting basic mailbox data"
    $watchAll   = Start-Stopwatch
    $tenantName = Split-Path -Path $ExportPath -Leaf

    # 1) Load all user mailboxes.
    #    ExchangeGuid / ExternalDirectoryObjectId are the unambiguous handles used
    #    by every per-mailbox call further down - they must be requested here.
    Write-Host " Loading ALL user mailboxes..." -NoNewline
    $watch = Start-Stopwatch

    # UserPrincipalName, Identity, Name, Alias, DisplayName, PrimarySmtpAddress,
    # EmailAddresses, ExternalDirectoryObjectId, RecipientType(Details) are part of
    # the Minimum property set and are always returned - only the extras are listed
    # here, because a single unsupported name would abort the whole call.
    $userProps = @(
        "ExchangeGuid",
        "AccountDisabled","IsInactiveMailbox","ExchangeUserAccountControl",
        "HiddenFromAddressListsEnabled",
        "ForwardingAddress","ForwardingSmtpAddress","DeliverToMailboxAndForward",
        "GrantSendOnBehalfTo",
        "LitigationHoldEnabled","LitigationHoldDuration","LitigationHoldDate","LitigationHoldOwner",
        "RetentionPolicy","RetentionHoldEnabled","SingleItemRecoveryEnabled",
        "AuditEnabled","AuditAdmin","AuditDelegate",
        "MaxSendSize","MaxReceiveSize",
        "ProhibitSendQuota","ProhibitSendReceiveQuota","IssueWarningQuota",
        "ArchiveStatus","ArchiveName","ArchiveGuid","ArchiveQuota",
        "Office","WhenCreated","WhenChanged",
        "CustomAttribute1","CustomAttribute2","CustomAttribute3","CustomAttribute4","CustomAttribute5",
        "ExtensionCustomAttribute1","ExtensionCustomAttribute2"
    )

    $users = $null
    try {
        $users = Get-EXOMailbox -ResultSize Unlimited -Properties $userProps
    } catch {
        # A single unsupported property name aborts the whole call - fall back to
        # the complete property set rather than losing the run.
        Write-Host ""
        Write-Host ("  Property set rejected (" + $_.Exception.Message + ") - falling back to -PropertySets All") -ForegroundColor Yellow
        Add-EXOError -Scope "Get-EXOMailbox" -Target "(all)" -Message $_.Exception.Message
        $users = Get-EXOMailbox -ResultSize Unlimited -PropertySets All
    }
    Write-Host (" done in " + (Stop-StopwatchString $watch) + ". " + @($users).Count + " found.")

    # 2) Export mailbox list (flat, multi-value fields joined)
    Write-Host " Exporting mailbox list to CSV..."
    $watch = Start-Stopwatch

    $mbxExport = foreach ($u in $users) {
        [pscustomobject][ordered]@{
            UserPrincipalName             = [string]$u.UserPrincipalName
            DisplayName                   = [string]$u.DisplayName
            Alias                         = [string]$u.Alias
            PrimarySmtpAddress            = [string]$u.PrimarySmtpAddress
            EmailAddresses                = ConvertTo-JoinedString $u.EmailAddresses
            RecipientTypeDetails          = [string]$u.RecipientTypeDetails
            ExchangeGuid                  = [string]$u.ExchangeGuid
            ExternalDirectoryObjectId     = [string]$u.ExternalDirectoryObjectId
            AccountDisabled               = $u.AccountDisabled
            IsInactiveMailbox             = $u.IsInactiveMailbox
            ExchangeUserAccountControl    = [string]$u.ExchangeUserAccountControl
            HiddenFromAddressListsEnabled = $u.HiddenFromAddressListsEnabled
            ForwardingAddress             = [string]$u.ForwardingAddress
            ForwardingSmtpAddress         = [string]$u.ForwardingSmtpAddress
            DeliverToMailboxAndForward    = $u.DeliverToMailboxAndForward
            GrantSendOnBehalfTo           = ConvertTo-JoinedString $u.GrantSendOnBehalfTo
            LitigationHoldEnabled         = $u.LitigationHoldEnabled
            LitigationHoldDuration        = [string]$u.LitigationHoldDuration
            LitigationHoldDate            = [string]$u.LitigationHoldDate
            LitigationHoldOwner           = [string]$u.LitigationHoldOwner
            RetentionPolicy               = [string]$u.RetentionPolicy
            RetentionHoldEnabled          = $u.RetentionHoldEnabled
            SingleItemRecoveryEnabled     = $u.SingleItemRecoveryEnabled
            AuditEnabled                  = $u.AuditEnabled
            AuditAdmin                    = ConvertTo-JoinedString $u.AuditAdmin
            AuditDelegate                 = ConvertTo-JoinedString $u.AuditDelegate
            MaxSendSize                   = [string]$u.MaxSendSize
            MaxReceiveSize                = [string]$u.MaxReceiveSize
            ProhibitSendQuota             = [string]$u.ProhibitSendQuota
            ProhibitSendReceiveQuota      = [string]$u.ProhibitSendReceiveQuota
            IssueWarningQuota             = [string]$u.IssueWarningQuota
            ArchiveStatus                 = [string]$u.ArchiveStatus
            ArchiveName                   = ConvertTo-JoinedString $u.ArchiveName
            ArchiveGuid                   = [string]$u.ArchiveGuid
            ArchiveQuota                  = [string]$u.ArchiveQuota
            Office                        = [string]$u.Office
            WhenCreated                   = [string]$u.WhenCreated
            WhenChanged                   = [string]$u.WhenChanged
            CustomAttribute1              = [string]$u.CustomAttribute1
            CustomAttribute2              = [string]$u.CustomAttribute2
            CustomAttribute3              = [string]$u.CustomAttribute3
            CustomAttribute4              = [string]$u.CustomAttribute4
            CustomAttribute5              = [string]$u.CustomAttribute5
            ExtensionCustomAttribute1     = ConvertTo-JoinedString $u.ExtensionCustomAttribute1
            ExtensionCustomAttribute2     = ConvertTo-JoinedString $u.ExtensionCustomAttribute2
        }
    }
    Export-EXOCsv -InputObject $mbxExport -Path (Join-Path $ExportPath ($tenantName + "_Mailboxes.csv"))
    Write-Host (" done in " + (Stop-StopwatchString $watch) + ".")

    # 3) Count by mailbox type
    Write-Host " Counting mailbox types..." -NoNewline
    $watch     = Start-Stopwatch
    $countPath = Join-Path $ExportPath ($tenantName + "_MBXCount.csv")
    if (Test-Path -LiteralPath $countPath) { Remove-Item -LiteralPath $countPath -Force }

    $users | Group-Object -Property RecipientTypeDetails | Select-Object Count,Name |
        Export-Csv -Path $countPath -Delimiter $script:CsvDelim -NoTypeInformation -Encoding $script:CsvEncoding -Force

    foreach ($variant in @(
        @{ Switch = 'GroupMailbox'; Label = $null },
        @{ Switch = 'PublicFolder'; Label = $null },
        @{ Switch = 'Archive';      Label = "MBX with Archive" }
    )) {
        $splat = @{ ResultSize = 'Unlimited'; ErrorAction = 'Stop' }
        $splat[$variant.Switch] = $true

        $set = $null
        try {
            $set = @(Get-Mailbox @splat)
        } catch {
            Add-EXOError -Scope "Get-Mailbox (count)" -Target $variant.Switch -Message $_.Exception.Message
            continue
        }
        if ($set.Count -eq 0) { continue }

        $label = $variant.Label
        if ($label) {
            $set | Group-Object -Property RecipientTypeDetails |
                Select-Object Count,@{Name="Name";Expression={$label}} |
                Export-Csv -Path $countPath -Delimiter $script:CsvDelim -NoTypeInformation -Encoding $script:CsvEncoding -Append
        } else {
            $set | Group-Object -Property RecipientTypeDetails | Select-Object Count,Name |
                Export-Csv -Path $countPath -Delimiter $script:CsvDelim -NoTypeInformation -Encoding $script:CsvEncoding -Append
        }
    }
    Write-Host (" done in " + (Stop-StopwatchString $watch) + ".")

    Write-Host ("Finished mailbox data in " + (Stop-StopwatchString $watchAll)) -ForegroundColor Green
    Write-Host ""

    return $users
}

function Get-EXOMailboxPermissions {
    param(
        [Parameter(Mandatory=$true)][string]$ExportPath,
        [Parameter(Mandatory=$true)][object]$Users
    )

    Write-Section "Getting mailbox permission data"
    $watchAll   = Start-Stopwatch
    $tenantName = Split-Path -Path $ExportPath -Leaf

    $cmdMbxPerm = Resolve-EXOCmdlet -Preferred 'Get-EXOMailboxPermission'   -Fallback 'Get-MailboxPermission'
    $cmdRcpPerm = Resolve-EXOCmdlet -Preferred 'Get-EXORecipientPermission' -Fallback 'Get-RecipientPermission'

    $fullAccess = New-Object 'System.Collections.Generic.List[PSObject]'
    $sendAs     = New-Object 'System.Collections.Generic.List[PSObject]'

    $total = @($Users).Count
    $i     = 0

    Write-Host " Collecting FullAccess and SendAs..."
    foreach ($user in $Users) {
        $i++
        Write-Progress -Activity "Mailbox permissions" `
            -PercentComplete (Get-ProgressPercent -Current $i -Total $total) `
            -Status ("$i/$total - " + $user.Alias)

        # Bind through ExchangeGuid / PrimarySmtpAddress, never through Identity:
        # Identity is the display name and is ambiguous whenever two objects share it.
        $mbxKey = Get-EXOMailboxKey   -Mailbox $user
        $rcpKey = Get-EXORecipientKey -Mailbox $user

        # --- FullAccess ---
        try {
            $perms = & $cmdMbxPerm -Identity $mbxKey -ErrorAction Stop |
                     Where-Object { ($_.User -notmatch "SELF") -and ($_.IsInherited -eq $false) }
            foreach ($p in $perms) {
                $fullAccess.Add([pscustomobject]@{
                    UserPrincipalName  = [string]$user.UserPrincipalName
                    PrimarySmtpAddress = [string]$user.PrimarySmtpAddress
                    Identity           = [string]$p.Identity
                    User               = [string]$p.User
                    AccessRights       = ConvertTo-JoinedString $p.AccessRights
                    Deny               = $p.Deny
                })
            }
        } catch {
            Add-EXOError -Scope "Get-MailboxPermission" -Target $mbxKey -Message $_.Exception.Message
        }

        # --- SendAs ---
        try {
            $perms = & $cmdRcpPerm -Identity $rcpKey -AccessRights SendAs -ErrorAction Stop |
                     Where-Object { $_.Trustee -notlike "SELF*" }
            foreach ($p in $perms) {
                $sendAs.Add([pscustomobject]@{
                    UserPrincipalName  = [string]$user.UserPrincipalName
                    PrimarySmtpAddress = [string]$user.PrimarySmtpAddress
                    Identity           = [string]$p.Identity
                    Trustee            = [string]$p.Trustee
                    AccessRights       = ConvertTo-JoinedString $p.AccessRights
                })
            }
        } catch {
            Add-EXOError -Scope "Get-RecipientPermission" -Target $rcpKey -Message $_.Exception.Message
        }
    }
    Write-Progress -Activity "Mailbox permissions" -Completed

    Export-EXOCsv -InputObject $fullAccess -Path (Join-Path $ExportPath ($tenantName + "_MailboxPermissions.csv"))
    Export-EXOCsv -InputObject $sendAs     -Path (Join-Path $ExportPath ($tenantName + "_MailboxesSendAs.csv"))

    # --- Forwarding ---
    $forwarding = $Users |
        Where-Object { $_.ForwardingAddress -or $_.ForwardingSmtpAddress } |
        Select-Object @{Name="UserPrincipalName";Expression={[string]$_.UserPrincipalName}},
                      Alias,
                      @{Name="ForwardingAddress";Expression={[string]$_.ForwardingAddress}},
                      @{Name="ForwardingSmtpAddress";Expression={[string]$_.ForwardingSmtpAddress}},
                      DeliverToMailboxAndForward
    Export-EXOCsv -InputObject $forwarding -Path (Join-Path $ExportPath ($tenantName + "_MailboxesForwardTo.csv"))

    # --- SendOnBehalf ---
    $sob = $Users |
        Where-Object { $_.GrantSendOnBehalfTo -and @($_.GrantSendOnBehalfTo).Count -gt 0 } |
        Select-Object @{Name="UserPrincipalName";Expression={[string]$_.UserPrincipalName}},
                      Alias,
                      @{Name="GrantSendOnBehalfTo";Expression={ ConvertTo-JoinedString $_.GrantSendOnBehalfTo }}
    Export-EXOCsv -InputObject $sob -Path (Join-Path $ExportPath ($tenantName + "_MailboxesSendOnBehalf.csv"))

    Write-Host ("Finished mailbox permissions in " + (Stop-StopwatchString $watchAll)) -ForegroundColor Green
    Write-Host ""
}

# ------------------------------------------------------------
# Folder permissions (logic ported from on-prem v3.2)
# ------------------------------------------------------------
function Get-EXOFolderPermissionsForMailbox {
    param([Parameter(Mandatory=$true)]$Mailbox)

    $smtp   = [string]$Mailbox.PrimarySmtpAddress
    $mbxKey = Get-EXOMailboxKey -Mailbox $Mailbox

    $skip    = @("Recoverable Items","SubstrateHolds","Purges","Deletions","Calendar Logging","Versions")
    $special = @("Top of Information Store","Oberste Ebene des Informationsspeichers",
                 "Haut de la banque d'informations","Bovenste map van gegevensarchief")

    $cmdStats = Resolve-EXOCmdlet -Preferred 'Get-EXOMailboxFolderStatistics' -Fallback 'Get-MailboxFolderStatistics'
    $cmdPerm  = Resolve-EXOCmdlet -Preferred 'Get-EXOMailboxFolderPermission' -Fallback 'Get-MailboxFolderPermission'

    $results = New-Object 'System.Collections.Generic.List[PSObject]'

    $stats = $null
    try {
        $stats = & $cmdStats -Identity $mbxKey -ErrorAction Stop
    } catch {
        Add-EXOError -Scope "Get-MailboxFolderStatistics" -Target $mbxKey -Message $_.Exception.Message
        return $results
    }

    $total = @($stats).Count
    $i     = 0

    foreach ($folder in $stats) {
        $i++
        Write-Progress -Activity ("Folder permissions [" + $smtp + "]") -ParentId 1 `
            -PercentComplete (Get-ProgressPercent -Current $i -Total $total) `
            -Status ("$i/$total - " + $folder.Name)

        # Identity of Get-MailboxFolderStatistics is "{DisplayName|GUID}\{Folder}\...".
        # Take everything after the FIRST backslash as the relative path and always
        # prefix the SMTP address - that avoids display-name / GUID ambiguity and
        # keeps nested paths intact.
        $identityStr  = $folder.Identity.ToString()
        $backslashPos = $identityStr.IndexOf('\')
        if ($backslashPos -lt 0) { continue }

        $relativePart = $identityStr.Substring($backslashPos + 1)
        $topFolder    = $relativePart.Split('\')[0]

        # "continue", not "break" - a single skipped system folder must not abort
        # the whole mailbox (that was the v3.0 bug).
        if ($skip -contains $topFolder) { continue }

        $folderPath = if ($special -contains $topFolder) { "$($smtp):\" } else { "$($smtp):\$relativePart" }

        $perms = $null
        try {
            $perms = & $cmdPerm -Identity $folderPath -ErrorAction Stop
        } catch {
            # Folders that cannot be addressed (special chars, deleted mid-run)
            Add-EXOError -Scope "Get-MailboxFolderPermission" -Target $folderPath -Message $_.Exception.Message
            continue
        }
        if (-not $perms) { continue }

        foreach ($p in $perms) {
            if ([string]$p.User -match "Default|Standard|Anonym") { continue }
            $results.Add([pscustomobject]@{
                Mailbox                = $smtp
                FolderName             = $folder.Name
                FolderPath             = $folderPath
                User                   = [string]$p.User
                AccessRights           = ConvertTo-JoinedString $p.AccessRights
                SharingPermissionFlags = [string]$p.SharingPermissionFlags
            })
        }
    }

    Write-Progress -Activity ("Folder permissions [" + $smtp + "]") -ParentId 1 -Completed
    Write-Host ("    [" + $smtp + "] " + $total + " folders scanned, " + $results.Count + " custom permissions found")
    return $results
}

function Get-EXOMailboxFolderPermissions {
    param(
        [Parameter(Mandatory=$true)][string]$ExportPath,
        [Parameter(Mandatory=$true)][object]$Users
    )

    Write-Section "Getting mailbox FOLDER permission data"
    $watchAll = Start-Stopwatch

    $permDir = Join-Path $ExportPath "MBXPerms"
    if (-not (Test-Path -LiteralPath $permDir)) {
        [void](New-Item -Path $permDir -ItemType Directory -Force)
    }

    $total    = @($Users).Count
    $i        = 0
    $activity = "Mailbox folder permissions"

    foreach ($mbx in $Users) {
        $i++
        Write-Progress -Activity $activity -Id 1 `
            -PercentComplete (Get-ProgressPercent -Current $i -Total $total) `
            -Status ("$i/$total - " + $mbx.Alias)

        $result = Get-EXOFolderPermissionsForMailbox -Mailbox $mbx

        if ($result -and $result.Count -gt 0) {
            # Alias can contain characters that are illegal in file names
            $safeName = ($mbx.Alias -replace '[\\/:*?"<>|]', '_')
            if ([string]::IsNullOrWhiteSpace($safeName)) { $safeName = (Get-EXOMailboxKey -Mailbox $mbx) }
            $result | Export-Csv -Path (Join-Path $permDir ($safeName + ".csv")) `
                -NoTypeInformation -Delimiter $script:CsvDelim -Encoding $script:CsvEncoding -Force
        }
    }
    Write-Progress -Activity $activity -Completed -Id 1

    Write-Host ("Finished folder permissions in " + (Stop-StopwatchString $watchAll)) -ForegroundColor Green
    Write-Host ""
}

# =============================
# PART 3: Mailbox Statistics + Public Folders
# =============================

# NOTE: deliberately named Export-* and not Get-EXOMailboxStatistics. In v3.0 the
# script function shadowed the EXO cmdlet of the same name, so the call inside the
# function recursed into itself, failed parameter binding and was swallowed by the
# try/catch - the statistics CSV came out empty for every mailbox.
function Export-EXOMailboxStatistics {
    param(
        [Parameter(Mandatory=$true)][string]$ExportPath,
        [Parameter(Mandatory=$true)][object]$Users,
        [switch]$Archive
    )

    $label = if ($Archive) { "archive mailbox statistics" } else { "mailbox statistics" }
    Write-Section ("Getting " + $label)
    $watchAll   = Start-Stopwatch
    $tenantName = Split-Path -Path $ExportPath -Leaf

    $cmdStats = Resolve-EXOCmdlet -Preferred 'Get-EXOMailboxStatistics' -Fallback 'Get-MailboxStatistics'

    $target = if ($Archive) { $Users | Where-Object { $_.ArchiveGuid -and "$($_.ArchiveGuid)" -ne '00000000-0000-0000-0000-000000000000' } } else { $Users }
    $target = @($target)

    $csvData  = New-Object 'System.Collections.Generic.List[PSObject]'
    $total    = $target.Count
    $i        = 0
    $activity = if ($Archive) { "Get-MailboxStatistics (Archive)" } else { "Get-MailboxStatistics" }

    foreach ($user in $target) {
        $i++
        Write-Progress -Activity $activity `
            -PercentComplete (Get-ProgressPercent -Current $i -Total $total) `
            -Status ("$i/$total - " + $user.Alias)

        $mbxKey = Get-EXOMailboxKey -Mailbox $user

        try {
            $stats = if ($Archive) {
                & $cmdStats -Identity $mbxKey -Archive -ErrorAction Stop
            } else {
                & $cmdStats -Identity $mbxKey -ErrorAction Stop
            }

            foreach ($s in @($stats)) {
                $csvData.Add([pscustomobject][ordered]@{
                    UserPrincipalName       = [string]$user.UserPrincipalName
                    PrimarySmtpAddress      = [string]$user.PrimarySmtpAddress
                    DisplayName             = [string]$s.DisplayName
                    MailboxType             = [string]$s.MailboxType
                    MailboxTypeDetail       = [string]$s.MailboxTypeDetail
                    ItemCount               = ConvertTo-Int64Safe $s.ItemCount
                    TotalItemSize           = [string]$s.TotalItemSize
                    TotalItemSizeBytes      = ConvertFrom-EXOSizeString $s.TotalItemSize
                    DeletedItemCount        = ConvertTo-Int64Safe $s.DeletedItemCount
                    TotalDeletedItemSize    = [string]$s.TotalDeletedItemSize
                    LastLoggedOnUserAccount = [string]$s.LastLoggedOnUserAccount
                    LastLogonTime           = [string]$s.LastLogonTime
                })
            }
        } catch {
            Add-EXOError -Scope $activity -Target $mbxKey -Message $_.Exception.Message
        }
    }
    Write-Progress -Activity $activity -Completed

    $fileName = if ($Archive) { $tenantName + "_MailboxStatistics-ArchiveMBX.csv" } else { $tenantName + "_MailboxStatistics.csv" }
    Export-EXOCsv -InputObject $csvData -Path (Join-Path $ExportPath $fileName)

    Write-Host ("Finished " + $label + " in " + (Stop-StopwatchString $watchAll)) -ForegroundColor Green
    Write-Host ""
}

function Get-EXOPublicFolderData {
    param([Parameter(Mandatory=$true)][string]$ExportPath)

    Write-Section "Getting Public Folder data"
    $watchAll   = Start-Stopwatch
    $tenantName = Split-Path -Path $ExportPath -Leaf

    # 1) Structure (loaded once and reused for the permission export)
    $pf = $null
    try {
        Write-Host " Exporting Public Folder structure..."
        $pf = Get-PublicFolder -Recurse -ResultSize Unlimited -ErrorAction Stop
        Export-EXOCsv -InputObject ($pf | Select-Object Name,ParentPath) `
            -Path (Join-Path $ExportPath ($tenantName + "_PFStructure.csv"))
    } catch {
        Write-Host " No public folders detected." -ForegroundColor Yellow
        Add-EXOError -Scope "Get-PublicFolder" -Target "(all)" -Message $_.Exception.Message
        Export-EXOCsv -InputObject $null -Path (Join-Path $ExportPath ($tenantName + "_PFStructure.csv")) -Quiet
    }

    # 2) Statistics
    try {
        Write-Host " Exporting Public Folder statistics..."
        $pfStats = Get-PublicFolderStatistics -ResultSize Unlimited -ErrorAction Stop |
            Select-Object Name,
                @{Name="Identity";Expression={ $_.Identity.ToString() }},
                ItemCount,
                TotalItemSize,
                @{Name="TotalItemSizeBytes";Expression={ ConvertFrom-EXOSizeString $_.TotalItemSize }},
                LastModificationTime,
                @{Name="MailboxOwnerId";Expression={ [string]$_.MailboxOwnerId }}
        Export-EXOCsv -InputObject $pfStats -Path (Join-Path $ExportPath ($tenantName + "_PFStatistics.csv"))
    } catch {
        Add-EXOError -Scope "Get-PublicFolderStatistics" -Target "(all)" -Message $_.Exception.Message
        Export-EXOCsv -InputObject $null -Path (Join-Path $ExportPath ($tenantName + "_PFStatistics.csv")) -Quiet
    }

    # 3) Client permissions
    try {
        Write-Host " Exporting Public Folder permissions..."
        $pfPerms = New-Object 'System.Collections.Generic.List[PSObject]'
        foreach ($folder in @($pf)) {
            if ($null -eq $folder) { continue }
            try {
                $perms = Get-PublicFolderClientPermission -Identity $folder.Identity -ErrorAction Stop
                foreach ($p in $perms) {
                    $pfPerms.Add([pscustomobject]@{
                        Identity     = [string]$p.Identity
                        User         = [string]$p.User
                        AccessRights = ConvertTo-JoinedString $p.AccessRights
                    })
                }
            } catch {
                Add-EXOError -Scope "Get-PublicFolderClientPermission" -Target ([string]$folder.Identity) -Message $_.Exception.Message
            }
        }
        Export-EXOCsv -InputObject $pfPerms -Path (Join-Path $ExportPath ($tenantName + "_PFPerms.csv"))
    } catch {
        Export-EXOCsv -InputObject $null -Path (Join-Path $ExportPath ($tenantName + "_PFPerms.csv")) -Quiet
    }

    # 4) Mail-enabled public folders + SendAs / SendOnBehalf
    try {
        Write-Host " Exporting Mail-Enabled Public Folder data..."
        $mepf = Get-MailPublicFolder -ResultSize Unlimited -ErrorAction Stop

        Export-EXOCsv -InputObject ($mepf | Select-Object Name,Alias,
                @{Name="PrimarySmtpAddress";Expression={[string]$_.PrimarySmtpAddress}},
                @{Name="EmailAddresses";Expression={ ConvertTo-JoinedString $_.EmailAddresses }},
                HiddenFromAddressListsEnabled,
                @{Name="ContentMailbox";Expression={[string]$_.ContentMailbox}}) `
            -Path (Join-Path $ExportPath ($tenantName + "_MailPFAddresses.csv"))

        # SendAs - resolve each PF by its (unique) SMTP address
        $pfSendAs = New-Object 'System.Collections.Generic.List[PSObject]'
        foreach ($mpf in $mepf) {
            $key = [string]$mpf.PrimarySmtpAddress
            if ([string]::IsNullOrWhiteSpace($key)) { continue }
            try {
                $perms = Get-RecipientPermission -Identity $key -AccessRights SendAs -ErrorAction Stop |
                         Where-Object { $_.Trustee -notlike "SELF*" }
                foreach ($p in $perms) {
                    $pfSendAs.Add([pscustomobject]@{
                        PublicFolder = [string]$mpf.Name
                        Identity     = [string]$p.Identity
                        Trustee      = [string]$p.Trustee
                        AccessRights = ConvertTo-JoinedString $p.AccessRights
                    })
                }
            } catch {
                Add-EXOError -Scope "Get-RecipientPermission (MailPF)" -Target $key -Message $_.Exception.Message
            }
        }
        Export-EXOCsv -InputObject $pfSendAs -Path (Join-Path $ExportPath ($tenantName + "_MailPFSendAs.csv"))

        $pfSoB = $mepf | Where-Object { $_.GrantSendOnBehalfTo -and @($_.GrantSendOnBehalfTo).Count -gt 0 } |
                 Select-Object Name,
                     @{Name="PrimarySmtpAddress";Expression={[string]$_.PrimarySmtpAddress}},
                     @{Name="GrantSendOnBehalfTo";Expression={ ConvertTo-JoinedString $_.GrantSendOnBehalfTo }}
        Export-EXOCsv -InputObject $pfSoB -Path (Join-Path $ExportPath ($tenantName + "_MailPFSoB.csv"))
    } catch {
        Add-EXOError -Scope "Get-MailPublicFolder" -Target "(all)" -Message $_.Exception.Message
        Export-EXOCsv -InputObject $null -Path (Join-Path $ExportPath ($tenantName + "_MailPFAddresses.csv")) -Quiet
        Export-EXOCsv -InputObject $null -Path (Join-Path $ExportPath ($tenantName + "_MailPFSendAs.csv"))    -Quiet
        Export-EXOCsv -InputObject $null -Path (Join-Path $ExportPath ($tenantName + "_MailPFSoB.csv"))       -Quiet
    }

    Write-Host ("Finished public folder data in " + (Stop-StopwatchString $watchAll)) -ForegroundColor Green
    Write-Host ""
}

# =============================
# PART 4: Groups and other recipients (ported from on-prem v3.2)
# =============================

function Get-EXOGroupData {
    param([Parameter(Mandatory=$true)][string]$ExportPath)

    Write-Section "Getting distribution group data"
    $watchAll   = Start-Stopwatch
    $tenantName = Split-Path -Path $ExportPath -Leaf

    Write-Host " Loading distribution / mail-enabled security groups..." -NoNewline
    $allGroups = @()
    try {
        $allGroups = @(Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop)
    } catch {
        Add-EXOError -Scope "Get-DistributionGroup" -Target "(all)" -Message $_.Exception.Message
    }
    Write-Host (" " + $allGroups.Count + " found.")

    Export-EXOCsv -InputObject ($allGroups | Select-Object Name,DisplayName,Alias,
            @{Name="PrimarySmtpAddress";Expression={[string]$_.PrimarySmtpAddress}},
            @{Name="EmailAddresses";Expression={ ConvertTo-JoinedString $_.EmailAddresses }},
            @{Name="GroupType";Expression={[string]$_.GroupType}},
            RecipientTypeDetails,
            @{Name="ManagedBy";Expression={ ConvertTo-JoinedString $_.ManagedBy }},
            MemberJoinRestriction,MemberDepartRestriction,
            RequireSenderAuthenticationEnabled,HiddenFromAddressListsEnabled,
            @{Name="AcceptMessagesOnlyFrom";Expression={ ConvertTo-JoinedString $_.AcceptMessagesOnlyFrom }},
            @{Name="AcceptMessagesOnlyFromDLMembers";Expression={ ConvertTo-JoinedString $_.AcceptMessagesOnlyFromDLMembers }},
            @{Name="RejectMessagesFrom";Expression={ ConvertTo-JoinedString $_.RejectMessagesFrom }},
            @{Name="GrantSendOnBehalfTo";Expression={ ConvertTo-JoinedString $_.GrantSendOnBehalfTo }}) `
        -Path (Join-Path $ExportPath ($tenantName + "_DistributionGroups.csv"))

    # Members - bind through PrimarySmtpAddress, group names are not unique either
    Write-Host " Collecting group members..."
    $memberList = New-Object 'System.Collections.Generic.List[PSObject]'
    $total = $allGroups.Count
    $i     = 0
    foreach ($grp in $allGroups) {
        $i++
        Write-Progress -Activity "Get-DistributionGroupMember" `
            -PercentComplete (Get-ProgressPercent -Current $i -Total $total) `
            -Status ("$i/$total - " + $grp.Alias)

        $key = [string]$grp.PrimarySmtpAddress
        if ([string]::IsNullOrWhiteSpace($key)) { $key = [string]$grp.Guid }

        try {
            $members = Get-DistributionGroupMember -Identity $key -ResultSize Unlimited -ErrorAction Stop
            foreach ($m in $members) {
                $memberList.Add([pscustomobject]@{
                    GroupName                  = [string]$grp.Name
                    GroupSmtpAddress           = [string]$grp.PrimarySmtpAddress
                    GroupType                  = [string]$grp.GroupType
                    MemberName                 = [string]$m.Name
                    MemberAlias                = [string]$m.Alias
                    MemberSmtpAddress          = [string]$m.PrimarySmtpAddress
                    MemberRecipientType        = [string]$m.RecipientType
                    MemberRecipientTypeDetails = [string]$m.RecipientTypeDetails
                })
            }
        } catch {
            Add-EXOError -Scope "Get-DistributionGroupMember" -Target $key -Message $_.Exception.Message
        }
    }
    Write-Progress -Activity "Get-DistributionGroupMember" -Completed
    Export-EXOCsv -InputObject $memberList -Path (Join-Path $ExportPath ($tenantName + "_DistributionGroupMembers.csv"))

    # Dynamic distribution groups
    Write-Host " Loading dynamic distribution groups..." -NoNewline
    $dynGroups = @()
    try {
        $dynGroups = @(Get-DynamicDistributionGroup -ResultSize Unlimited -ErrorAction Stop)
    } catch {
        Add-EXOError -Scope "Get-DynamicDistributionGroup" -Target "(all)" -Message $_.Exception.Message
    }
    Write-Host (" " + $dynGroups.Count + " found.")

    Export-EXOCsv -InputObject ($dynGroups | Select-Object Name,DisplayName,Alias,
            @{Name="PrimarySmtpAddress";Expression={[string]$_.PrimarySmtpAddress}},
            @{Name="EmailAddresses";Expression={ ConvertTo-JoinedString $_.EmailAddresses }},
            @{Name="RecipientFilter";Expression={[string]$_.RecipientFilter}},
            @{Name="RecipientContainer";Expression={[string]$_.RecipientContainer}},
            @{Name="ManagedBy";Expression={ ConvertTo-JoinedString $_.ManagedBy }},
            HiddenFromAddressListsEnabled) `
        -Path (Join-Path $ExportPath ($tenantName + "_DynamicDistributionGroups.csv"))

    # Microsoft 365 groups
    Write-Host " Loading Microsoft 365 groups..." -NoNewline
    $m365Groups = @()
    try {
        $m365Groups = @(Get-UnifiedGroup -ResultSize Unlimited -ErrorAction Stop)
    } catch {
        Add-EXOError -Scope "Get-UnifiedGroup" -Target "(all)" -Message $_.Exception.Message
    }
    Write-Host (" " + $m365Groups.Count + " found.")

    Export-EXOCsv -InputObject ($m365Groups | Select-Object DisplayName,Alias,
            @{Name="PrimarySmtpAddress";Expression={[string]$_.PrimarySmtpAddress}},
            @{Name="EmailAddresses";Expression={ ConvertTo-JoinedString $_.EmailAddresses }},
            AccessType,
            @{Name="ManagedBy";Expression={ ConvertTo-JoinedString $_.ManagedBy }},
            HiddenFromAddressListsEnabled,HiddenFromExchangeClientsEnabled,
            GroupMemberCount,GroupExternalMemberCount,
            SharePointSiteUrl,WhenCreated) `
        -Path (Join-Path $ExportPath ($tenantName + "_M365Groups.csv"))

    Write-Host ("Finished group data in " + (Stop-StopwatchString $watchAll)) -ForegroundColor Green
    Write-Host ""
}

function Get-EXOOtherRecipients {
    param(
        [Parameter(Mandatory=$true)][string]$ExportPath,
        [Parameter(Mandatory=$true)][object]$Users
    )

    Write-Section "Getting other recipients"
    $watchAll   = Start-Stopwatch
    $tenantName = Split-Path -Path $ExportPath -Leaf

    Write-Host " Mail contacts..." -NoNewline
    $contacts = @()
    try { $contacts = @(Get-MailContact -ResultSize Unlimited -ErrorAction Stop) }
    catch { Add-EXOError -Scope "Get-MailContact" -Target "(all)" -Message $_.Exception.Message }
    Write-Host (" " + $contacts.Count + " found.")

    Export-EXOCsv -InputObject ($contacts | Select-Object Name,Alias,DisplayName,
            @{Name="ExternalEmailAddress";Expression={[string]$_.ExternalEmailAddress}},
            @{Name="EmailAddresses";Expression={ ConvertTo-JoinedString $_.EmailAddresses }},
            HiddenFromAddressListsEnabled,
            CustomAttribute1,CustomAttribute2,CustomAttribute3) `
        -Path (Join-Path $ExportPath ($tenantName + "_MailContacts.csv"))

    Write-Host " Mail users..." -NoNewline
    $mailUsers = @()
    try { $mailUsers = @(Get-MailUser -ResultSize Unlimited -ErrorAction Stop) }
    catch { Add-EXOError -Scope "Get-MailUser" -Target "(all)" -Message $_.Exception.Message }
    Write-Host (" " + $mailUsers.Count + " found.")

    Export-EXOCsv -InputObject ($mailUsers | Select-Object Name,Alias,DisplayName,UserPrincipalName,
            @{Name="ExternalEmailAddress";Expression={[string]$_.ExternalEmailAddress}},
            @{Name="EmailAddresses";Expression={ ConvertTo-JoinedString $_.EmailAddresses }},
            HiddenFromAddressListsEnabled,RecipientTypeDetails) `
        -Path (Join-Path $ExportPath ($tenantName + "_MailUsers.csv"))

    # Resource mailboxes + CalendarProcessing
    $resourceMbx = @($Users | Where-Object { @("RoomMailbox","EquipmentMailbox") -contains [string]$_.RecipientTypeDetails })
    Write-Host (" Resource mailboxes (Room + Equipment): " + $resourceMbx.Count)

    $calProc = New-Object 'System.Collections.Generic.List[PSObject]'
    foreach ($r in $resourceMbx) {
        $key = Get-EXOMailboxKey -Mailbox $r
        try {
            $cp = Get-CalendarProcessing -Identity $key -ErrorAction Stop | ConvertTo-FlatObject
            foreach ($c in @($cp)) {
                $c | Add-Member -MemberType NoteProperty -Name 'PrimarySmtpAddress'   -Value ([string]$r.PrimarySmtpAddress)   -Force
                $c | Add-Member -MemberType NoteProperty -Name 'DisplayName'          -Value ([string]$r.DisplayName)          -Force
                $c | Add-Member -MemberType NoteProperty -Name 'RecipientTypeDetails' -Value ([string]$r.RecipientTypeDetails) -Force
                $calProc.Add($c)
            }
        } catch {
            Add-EXOError -Scope "Get-CalendarProcessing" -Target $key -Message $_.Exception.Message
        }
    }
    Export-EXOCsv -InputObject $calProc -Path (Join-Path $ExportPath ($tenantName + "_ResourceMBX_CalendarProcessing.csv"))

    Write-Host ("Finished other recipients in " + (Stop-StopwatchString $watchAll)) -ForegroundColor Green
    Write-Host ""
}

# =============================
# PART 5: Tenant basics + Merge + Summaries
# =============================

function Get-EXOTenantBasics {
    <#
        .SYNOPSIS
            Exports tenant-wide objects: Accepted Domains, Connectors, Transport Rules,
            OWA policies, Mobile device policies, Organization relationships.
    #>
    param([Parameter(Mandatory=$true)][string]$ExportPath)

    Write-Section "Getting basic tenant-wide configuration data"
    $watchAll   = Start-Stopwatch
    $tenantName = Split-Path -Path $ExportPath -Leaf

    # Wide config objects are exported completely via ConvertTo-FlatObject, so a
    # future property does not silently drop out of the documentation.
    $exports = @(
        @{ Name = "AcceptedDomains";        Cmd = "Get-AcceptedDomain"           }
        @{ Name = "RemoteDomains";          Cmd = "Get-RemoteDomain"             }
        @{ Name = "InboundConnectors";      Cmd = "Get-InboundConnector"         }
        @{ Name = "OutboundConnectors";     Cmd = "Get-OutboundConnector"        }
        @{ Name = "TransportRules";         Cmd = "Get-TransportRule"            }
        @{ Name = "JournalRules";           Cmd = "Get-JournalRule"              }
        @{ Name = "OwaMailboxPolicies";     Cmd = "Get-OwaMailboxPolicy"         }
        @{ Name = "MobileDevicePolicies";   Cmd = "Get-MobileDeviceMailboxPolicy"}
        @{ Name = "RetentionPolicies";      Cmd = "Get-RetentionPolicy"          }
        @{ Name = "RetentionPolicyTags";    Cmd = "Get-RetentionPolicyTag"       }
        @{ Name = "Free-Busy-OrgRels";      Cmd = "Get-OrganizationRelationship" }
        @{ Name = "SharingPolicies";        Cmd = "Get-SharingPolicy"            }
        @{ Name = "AddressLists";           Cmd = "Get-AddressList"              }
        @{ Name = "OfflineAddressBooks";    Cmd = "Get-OfflineAddressBook"       }
        @{ Name = "OrganizationConfig";     Cmd = "Get-OrganizationConfig"       }
    )

    foreach ($e in $exports) {
        $outPath = Join-Path $ExportPath ($tenantName + "_" + $e.Name + ".csv")
        try {
            Write-Host (" Exporting " + $e.Name + "...")
            $data = & $e.Cmd -ErrorAction Stop | ConvertTo-FlatObject
            Export-EXOCsv -InputObject $data -Path $outPath
        } catch {
            Write-Host ("  Failed: " + $_.Exception.Message) -ForegroundColor Red
            Add-EXOError -Scope $e.Cmd -Target "(all)" -Message $_.Exception.Message
            Export-EXOCsv -InputObject $null -Path $outPath -Quiet
        }
    }

    # MX records for all accepted domains (ported from on-prem v3.2)
    if (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue) {
        Write-Host " Resolving MX records for accepted domains..."
        $mxList = New-Object 'System.Collections.Generic.List[PSObject]'
        try {
            foreach ($domain in (Get-AcceptedDomain -ErrorAction Stop)) {
                $mx = Resolve-DnsName -Name $domain.DomainName -Type MX -ErrorAction SilentlyContinue |
                      Where-Object { $_.Type -eq 'MX' } | Sort-Object Preference
                if ($mx) {
                    foreach ($r in $mx) {
                        $mxList.Add([pscustomobject]@{
                            AcceptedDomain = [string]$domain.Name
                            DomainName     = [string]$domain.DomainName
                            DomainType     = [string]$domain.DomainType
                            IsDefault      = $domain.Default
                            MXHost         = [string]$r.NameExchange
                            Preference     = $r.Preference
                        })
                    }
                } else {
                    $mxList.Add([pscustomobject]@{
                        AcceptedDomain = [string]$domain.Name
                        DomainName     = [string]$domain.DomainName
                        DomainType     = [string]$domain.DomainType
                        IsDefault      = $domain.Default
                        MXHost         = "(no MX record)"
                        Preference     = ""
                    })
                }
            }
        } catch {
            Add-EXOError -Scope "Resolve-DnsName" -Target "(accepted domains)" -Message $_.Exception.Message
        }
        Export-EXOCsv -InputObject $mxList -Path (Join-Path $ExportPath ($tenantName + "_MXRecords.csv"))
    }

    Write-Host ("Finished basic tenant data in " + (Stop-StopwatchString $watchAll)) -ForegroundColor Green
    Write-Host ""
}

function Merge-EXOMailboxData {
    <#
        .SYNOPSIS
            Merges mailbox inventory + primary stats + archive stats into one CSV.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$ExportPath,
        [Parameter(Mandatory=$true)][string]$TenantName
    )

    Write-Section "Merging mailbox information and statistics"
    $watchAll = Start-Stopwatch

    $mbxPath     = Join-Path $ExportPath ($TenantName + "_Mailboxes.csv")
    $statPath    = Join-Path $ExportPath ($TenantName + "_MailboxStatistics.csv")
    $arcStatPath = Join-Path $ExportPath ($TenantName + "_MailboxStatistics-ArchiveMBX.csv")

    $mbxInfos = @(); $mbxStatInfos = @(); $arcStatInfos = @()
    if (Test-Path -LiteralPath $mbxPath)     { $mbxInfos     = @(Import-Csv -Path $mbxPath     -Delimiter $script:CsvDelim) }
    if (Test-Path -LiteralPath $statPath)    { $mbxStatInfos = @(Import-Csv -Path $statPath    -Delimiter $script:CsvDelim) }
    if (Test-Path -LiteralPath $arcStatPath) { $arcStatInfos = @(Import-Csv -Path $arcStatPath -Delimiter $script:CsvDelim) }

    $statsIndex = @{}
    foreach ($s in $mbxStatInfos) {
        if ($s.UserPrincipalName) { $statsIndex[$s.UserPrincipalName.ToLower()] = $s }
    }
    $arcIndex = @{}
    foreach ($a in $arcStatInfos) {
        if ($a.UserPrincipalName) { $arcIndex[$a.UserPrincipalName.ToLower()] = $a }
    }

    $merged = foreach ($m in $mbxInfos) {
        $key     = if ($m.UserPrincipalName) { $m.UserPrincipalName.ToLower() } else { "" }
        $stat    = if ($key -and $statsIndex.ContainsKey($key)) { $statsIndex[$key] } else { $null }
        $arcStat = if ($key -and $arcIndex.ContainsKey($key))   { $arcIndex[$key]   } else { $null }

        $props = [ordered]@{}
        foreach ($p in $m.PSObject.Properties.Name) { $props[$p] = $m.$p }
        if ($stat) {
            foreach ($p in $stat.PSObject.Properties.Name) {
                if (-not $props.Contains($p)) { $props[$p] = $stat.$p }
            }
        }

        $primCount = 0; $primBytes = 0; $arcCount = 0; $arcBytes = 0
        if ($stat) {
            $primCount = ConvertTo-Int64Safe $stat.ItemCount
            $primBytes = ConvertTo-Int64Safe $stat.TotalItemSizeBytes
        }
        if ($arcStat) {
            $arcCount = ConvertTo-Int64Safe $arcStat.ItemCount
            $arcBytes = ConvertTo-Int64Safe $arcStat.TotalItemSizeBytes
        }

        $props['CombinedItemCount']     = $primCount + $arcCount
        $props['CombinedItemSizeBytes'] = $primBytes + $arcBytes
        $props['CombinedItemSize']      = Convert-BytesToSizeString -Bytes ($primBytes + $arcBytes)
        $props['Archive_ItemCount']     = $arcCount
        $props['Archive_ItemSizeBytes'] = $arcBytes

        [pscustomobject]$props
    }

    $outFile = Join-Path $ExportPath ($TenantName + "_MailboxesAndMailboxStatistics_Merged.csv")
    Export-EXOCsv -InputObject $merged -Path $outFile

    Write-Host ("Finished merging in " + (Stop-StopwatchString $watchAll)) -ForegroundColor Green
    Write-Host ""
}

function Get-EXOMailboxSizeSummary {
    param(
        [Parameter(Mandatory=$true)][string]$ExportPath,
        [Parameter(Mandatory=$true)][string]$TenantName
    )

    Write-Section "Calculating mailbox size summary"
    $watchAll = Start-Stopwatch

    $mergedPath = Join-Path $ExportPath ($TenantName + "_MailboxesAndMailboxStatistics_Merged.csv")
    if (-not (Test-Path -LiteralPath $mergedPath)) {
        Write-Host "Merged CSV not found. Skipping summary." -ForegroundColor Yellow
        return
    }

    # Import-Csv returns strings - Measure-Object needs numeric rows only,
    # otherwise a single non-numeric cell breaks the whole measurement.
    $rows = Import-Csv -Path $mergedPath -Delimiter $script:CsvDelim |
            Where-Object { $_.CombinedItemSizeBytes -match '^\d+$' }

    if (-not $rows) {
        Write-Host "No numeric size data found. Skipping summary." -ForegroundColor Yellow
        return
    }

    $measure = $rows | Measure-Object -Property CombinedItemSizeBytes -Sum -Average -Maximum -Minimum

    $result = [pscustomobject]@{
        Count  = $measure.Count
        Min_GB = "{0:N2}" -f ([double]$measure.Minimum / 1GB)
        Max_GB = "{0:N2}" -f ([double]$measure.Maximum / 1GB)
        Sum_GB = "{0:N2}" -f ([double]$measure.Sum     / 1GB)
        Avg_GB = "{0:N2}" -f ([double]$measure.Average / 1GB)
        Min_MB = "{0:N2}" -f ([double]$measure.Minimum / 1MB)
        Max_MB = "{0:N2}" -f ([double]$measure.Maximum / 1MB)
        Sum_MB = "{0:N2}" -f ([double]$measure.Sum     / 1MB)
        Avg_MB = "{0:N2}" -f ([double]$measure.Average / 1MB)
    }

    Export-EXOTxt -InputObject $result -Path (Join-Path $ExportPath ($TenantName + "_MailboxesSizeSummary.txt"))
    Write-Host ("    -> " + $TenantName + "_MailboxesSizeSummary.txt")
    Write-Host ("Finished mailbox size summary in " + (Stop-StopwatchString $watchAll)) -ForegroundColor Green
    Write-Host ""
}

function Get-EXOPFAggregation {
    <#
        .SYNOPSIS
            Aggregates Public Folder statistics per PF mailbox and joins PF-mailbox sizes.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$ExportPath,
        [Parameter(Mandatory=$true)][string]$TenantName
    )

    Write-Section "Aggregating Public Folder statistics per PF mailbox"
    $watchAll = Start-Stopwatch

    $csvPath = Join-Path $ExportPath ($TenantName + "_PFStatistics.csv")
    if (-not (Test-Path -LiteralPath $csvPath)) {
        Write-Host "PFStatistics CSV not found. Skipping PF aggregation." -ForegroundColor Yellow
        return
    }

    $data = @(Import-Csv -Path $csvPath -Delimiter $script:CsvDelim |
              Where-Object { $_.PSObject.Properties.Match('MailboxOwnerId').Count -gt 0 })
    if (-not $data -or $data.Count -eq 0) {
        Write-Host "No public folder statistics to aggregate." -ForegroundColor Yellow
        return
    }

    $agg = foreach ($group in ($data | Group-Object -Property MailboxOwnerId)) {
        $sumItems = [int64]0
        $sumBytes = [int64]0
        foreach ($r in $group.Group) {
            $sumItems += ConvertTo-Int64Safe $r.ItemCount
            $sumBytes += ConvertTo-Int64Safe $r.TotalItemSizeBytes
        }
        [pscustomobject]@{
            MailboxOwnerIdName    = $group.Name
            FolderCount           = $group.Count
            TotalItemCount        = $sumItems
            TotalItemSizeBytes    = $sumBytes
            TotalItemSizeReadable = Convert-BytesToSizeString -Bytes $sumBytes
        }
    }

    # PF mailbox sizes
    $pfMailboxes = @()
    try { $pfMailboxes = @(Get-Mailbox -PublicFolder -ResultSize Unlimited -ErrorAction Stop) }
    catch { Add-EXOError -Scope "Get-Mailbox -PublicFolder" -Target "(all)" -Message $_.Exception.Message }

    $mbxSizes = @{}
    foreach ($mbx in $pfMailboxes) {
        $bytes = [int64]0
        $key   = Get-EXOMailboxKey -Mailbox $mbx
        try {
            $stats = Get-MailboxStatistics -Identity $key -ErrorAction Stop
            if ($stats) { $bytes = ConvertFrom-EXOSizeString $stats.TotalItemSize }
        } catch {
            Add-EXOError -Scope "Get-MailboxStatistics (PF)" -Target $key -Message $_.Exception.Message
        }
        $mbxSizes[[string]$mbx.Name] = @{ Bytes = $bytes; Readable = (Convert-BytesToSizeString -Bytes $bytes) }
    }

    $result = foreach ($row in $agg) {
        $name        = [string]$row.MailboxOwnerIdName
        $mbxBytes    = [int64]0
        $mbxReadable = "0 B"
        if ($mbxSizes.ContainsKey($name)) {
            $mbxBytes    = [int64]$mbxSizes[$name].Bytes
            $mbxReadable = [string]$mbxSizes[$name].Readable
        }
        [pscustomobject]@{
            MailboxOwnerIdName    = $name
            FolderCount           = $row.FolderCount
            TotalItemCount        = $row.TotalItemCount
            TotalItemSizeBytes    = $row.TotalItemSizeBytes
            TotalItemSizeReadable = $row.TotalItemSizeReadable
            MailboxSizeBytes      = $mbxBytes
            MailboxSizeReadable   = $mbxReadable
        }
    }

    Export-EXOCsv -InputObject ($result | Sort-Object -Property MailboxSizeBytes -Descending) `
        -Path (Join-Path $ExportPath ($TenantName + "_PFAggregation_WithMailboxSize.csv"))

    Write-Host ("Finished PF aggregation in " + (Stop-StopwatchString $watchAll)) -ForegroundColor Green
    Write-Host ""
}

# -----------------------------
# MAIN: Orchestrator
# -----------------------------
function Start-EXODocumentation {
    param(
        [bool]$IncludeMailboxFolderPermissions = $false,
        [bool]$IncludeMailboxPermissions       = $true,
        [bool]$IncludeRecipients               = $true
    )

    Write-Section "Starting Exchange Online documentation export (v3.1)"
    $runStart = Get-Date

    Connect-EXOIfNeeded

    $baseExport = New-EXOExportPath -BaseName "EXO-Documentation"

    $tenant     = Get-EXOTenantInfo
    $tenantName = $tenant.Name

    $exportPath = Join-Path -Path $baseExport -ChildPath $tenantName
    if (-not (Test-Path -LiteralPath $exportPath)) {
        [void](New-Item -Path $exportPath -ItemType Directory -Force)
    }

    Push-Location -Path $exportPath
    try {
        Write-Host (" Export directory: " + $exportPath)

        Get-EXOTenantBasics -ExportPath $exportPath

        $users = Get-EXOMailboxData -ExportPath $exportPath

        if ($IncludeMailboxPermissions) {
            Get-EXOMailboxPermissions -ExportPath $exportPath -Users $users
        }
        if ($IncludeMailboxFolderPermissions) {
            Get-EXOMailboxFolderPermissions -ExportPath $exportPath -Users $users
        }

        Export-EXOMailboxStatistics -ExportPath $exportPath -Users $users
        Export-EXOMailboxStatistics -ExportPath $exportPath -Users $users -Archive

        if ($IncludeRecipients) {
            Get-EXOGroupData       -ExportPath $exportPath
            Get-EXOOtherRecipients -ExportPath $exportPath -Users $users
        }

        Get-EXOPublicFolderData -ExportPath $exportPath

        Merge-EXOMailboxData      -ExportPath $exportPath -TenantName $tenantName
        Get-EXOMailboxSizeSummary -ExportPath $exportPath -TenantName $tenantName
        Get-EXOPFAggregation      -ExportPath $exportPath -TenantName $tenantName

        # Every swallowed failure ends up here instead of disappearing
        if ($script:ErrorLog.Count -gt 0) {
            Write-Section "Warnings"
            Write-Host (" " + $script:ErrorLog.Count + " object(s) could not be read - see the error CSV.") -ForegroundColor Yellow
            $script:ErrorLog | Group-Object Scope | Sort-Object Count -Descending |
                ForEach-Object { Write-Host ("   " + $_.Count + "x " + $_.Name) -ForegroundColor Yellow }
            Export-EXOCsv -InputObject $script:ErrorLog -Path (Join-Path $exportPath ($tenantName + "_Errors.csv"))
        }
    }
    finally {
        Pop-Location
    }

    $totalSecs = [math]::Round(((Get-Date) - $runStart).TotalSeconds)
    Write-Host ""
    Write-Host ("=" * 60) -ForegroundColor White
    Write-Host ("All tasks completed. Total runtime: " + $totalSecs + "s") -ForegroundColor Green
    Write-Host ("Output: " + $exportPath) -ForegroundColor Green
}

# ========================================
# AUTO-START (uses top-level parameters)
# ========================================
Start-EXODocumentation `
    -IncludeMailboxFolderPermissions $IncludeMailboxFolderPermissions `
    -IncludeMailboxPermissions       $IncludeMailboxPermissions `
    -IncludeRecipients               $IncludeRecipients
