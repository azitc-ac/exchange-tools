param(
    [switch]$includeMailboxFolderPermissions,
    [switch]$includeInboxRules,
    [switch]$includeMobileDevices
)

# Einzige Stelle fuer die Versionsnummer; Build und Release-Tag lesen sie hier aus.
$script:Version = '1.0.0'


$FormatEnumerationLimit = -1
$sep         = [char]0x00A7  # §  — multi-value separator within CSV fields
$csvDelim    = ";"
$csvEncoding = if ($PSVersionTable.PSVersion.Major -ge 7) { "utf8BOM" } else { "UTF8" }  # PS5.1 UTF8 already includes BOM

#region Functions

function ConvertTo-FlatObject {
    param(
        [Parameter(ValueFromPipeline)]
        $InputObject,
        [string]$Separator = $sep
    )
    process {
        if ($null -eq $InputObject) { return }
        $hash = [ordered]@{}
        foreach ($prop in $InputObject.PSObject.Properties) {
            $val = $prop.Value
            $hash[$prop.Name] = switch ($true) {
                ($null -eq $val)                             { "" }
                ($val -is [string])                          { $val }
                ($val -is [array])                           {
                    $parts = [System.Collections.Generic.List[string]]::new()
                    for ($idx = 0; $idx -lt $val.Length; $idx++) { $parts.Add([string]$val[$idx]) }
                    $parts -join $Separator
                }
                ($val -is [System.Collections.IEnumerable]) {
                    try {
                        $parts = [System.Collections.Generic.List[string]]::new()
                        $enum = ([System.Collections.IEnumerable]$val).GetEnumerator()
                        while ($enum.MoveNext()) { $parts.Add([string]$enum.Current) }
                        $parts -join $Separator
                    } catch { [string]$val }
                }
                default                                      { "$val" }
            }
        }
        [pscustomobject]$hash
    }
}

function Convert-BytesToSizeString {
    param([long]$Bytes)
    $units = "B","KB","MB","GB","TB","PB"
    $size  = [double]$Bytes
    $i     = 0
    while ($size -ge 1024 -and $i -lt ($units.Count - 1)) { $size /= 1024; $i++ }
    $s = [System.String]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:0.##}", $size)
    return "$s $($units[$i])"
}

function Export-EXData {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline)]
        $Data,
        [Parameter(Mandatory)]
        [string]$FileName
    )
    begin   { $buf = [System.Collections.Generic.List[PSObject]]::new() }
    process { if ($null -ne $Data) { $buf.Add($Data) } }
    end {
        $path = Join-Path $logpath ($Kunde + "_" + $FileName)
        $buf | Export-Csv $path -Delimiter $csvDelim -NoTypeInformation -Encoding $csvEncoding -Force
        Write-Host "    -> $($Kunde)_$FileName  ($($buf.Count) records)"
    }
}

function Get-MailboxFolderPermissionsForAllFolders {
    param(
        [Parameter(Mandatory)]
        $Mailbox
    )
    $smtp    = [string]$Mailbox.PrimarySmtpAddress
    $skip    = "Recoverable Items","SubstrateHolds","Purges","Deletions","Calendar Logging","Versions"
    $special = "Top of Information Store","Oberste Ebene des Informationsspeichers",
               "Haut de la banque d'informations","Bovenste map van gegevensarchief"

    $stats   = Get-MailboxFolderStatistics $Mailbox -ErrorAction SilentlyContinue
    $results = [System.Collections.Generic.List[PSObject]]::new()
    $total   = $stats.Count
    $i       = 0

    foreach ($folder in $stats) {
        $i++
        Write-Progress -Activity "Folder permissions [$smtp]" -ParentId 1 `
            -PercentComplete ($i / $total * 100) -Status "$i/$total - $($folder.Name)"

        # Extract relative folder path: everything after the first backslash in Identity.
        # Identity format from Get-MailboxFolderStatistics is "{DisplayName|GUID}\{FolderName}\..."
        # We always use $smtp as the mailbox part — avoids display-name / GUID ambiguity.
        $identityStr  = $folder.Identity.ToString()
        $backslashPos = $identityStr.IndexOf('\')
        if ($backslashPos -lt 0) { continue }

        $relativePart = $identityStr.Substring($backslashPos + 1)
        $topFolder    = $relativePart.Split('\')[0]

        if ($topFolder -in $skip) { continue }

        # Top-of-store maps to "smtp:\", all others to "smtp:\relative\path"
        $folderPath = if ($topFolder -in $special) { "$($smtp):\" } else { "$($smtp):\$relativePart" }

        $perms = Get-MailboxFolderPermission $folderPath -ErrorAction SilentlyContinue
        if (-not $perms) { continue }

        foreach ($p in $perms) {
            if ([string]$p.User -match "Default|Standard|Anonym") { continue }
            # Safe AccessRights enumeration — avoids PS 5.1 SZArrayEnumerator pipeline issue
            $arParts = [System.Collections.Generic.List[string]]::new()
            try {
                $arEnum = ([System.Collections.IEnumerable]$p.AccessRights).GetEnumerator()
                while ($arEnum.MoveNext()) { $arParts.Add([string]$arEnum.Current) }
            } catch { $arParts.Add([string]$p.AccessRights) }

            $results.Add([pscustomobject]@{
                Mailbox                = $smtp
                FolderName             = $folder.Name
                FolderPath             = $folderPath
                User                   = [string]$p.User
                AccessRights           = $arParts -join $sep
                SharingPermissionFlags = [string]$p.SharingPermissionFlags
            })
        }
    }
    Write-Host "    [$smtp] $total folders scanned, $($results.Count) custom permissions found"
    Write-Progress -Activity "Folder permissions [$smtp]" -ParentId 1 -Completed
    return $results
}

#endregion Functions

#region Init

$snapin = Get-PSSnapin -Registered Microsoft.Exchange.Management.PowerShell.E* -ErrorAction SilentlyContinue
if ($null -eq $snapin) {
    Write-Warning "Exchange Snap-in not available"
} else {
    if ($null -eq (Get-Command "Get-Mailbox" -ErrorAction SilentlyContinue)) {
        Write-Host "Loading Exchange Commands"
        Add-PSSnapin $snapin -ErrorAction SilentlyContinue
    } else {
        Write-Host "Exchange Snap-in already loaded"
    }
}

$OrgSettings = Get-OrganizationConfig
$Kunde       = $OrgSettings.Name
$logpath     = "$PSScriptRoot\EXOnprem-Doku"
if (-not (Test-Path $logpath)) { New-Item -ItemType Directory -Path $logpath | Out-Null }
Push-Location $logpath

$runStart = Get-Date
Write-Host "`nExchange On-Premises Documentation v3.2" -ForegroundColor White
Write-Host "Organisation : $Kunde"
Write-Host "Output path  : $logpath"
Write-Host "Started      : $runStart"
Write-Host ("-" * 60)

#endregion Init

# ============================================================
#region Infrastructure
# ============================================================
Write-Host "`n## [1/9] Infrastructure ##" -ForegroundColor Cyan
$sectionStart = Get-Date

$exServers = Get-ExchangeServer
$exServers | ConvertTo-FlatObject | Export-EXData -FileName "Infrastructure_ExchangeServers.csv"

Get-MailboxDatabase -Status -ErrorAction SilentlyContinue |
    ConvertTo-FlatObject | Export-EXData -FileName "Infrastructure_MailboxDatabases.csv"

Get-DatabaseAvailabilityGroup -ErrorAction SilentlyContinue |
    ConvertTo-FlatObject | Export-EXData -FileName "Infrastructure_DAG.csv"

# Certificates from all Exchange servers
$certList = [System.Collections.Generic.List[PSObject]]::new()
foreach ($srv in $exServers) {
    $certs = Get-ExchangeCertificate -Server $srv.Name -ErrorAction SilentlyContinue
    foreach ($c in $certs) {
        $flat = $c | ConvertTo-FlatObject
        $flat | Add-Member -MemberType NoteProperty -Name 'Server' -Value $srv.Name -Force
        $certList.Add($flat)
    }
}
$certList | Export-EXData -FileName "Infrastructure_ExchangeCertificates.csv"

# Client Access URLs (virtual directories) from all CAS servers
$vDirList  = [System.Collections.Generic.List[PSObject]]::new()
$casServers = $exServers | Where-Object { $_.IsClientAccessServer }

# vDir types accessible via -Server parameter
$vDirTypes = @(
    [pscustomobject]@{ Type = "OWA";        Cmd = "Get-OwaVirtualDirectory"           }
    [pscustomobject]@{ Type = "ECP";        Cmd = "Get-EcpVirtualDirectory"           }
    [pscustomobject]@{ Type = "EWS";        Cmd = "Get-WebServicesVirtualDirectory"   }
    [pscustomobject]@{ Type = "MAPI";       Cmd = "Get-MapiVirtualDirectory"          }
    [pscustomobject]@{ Type = "ActiveSync"; Cmd = "Get-ActiveSyncVirtualDirectory"    }
    [pscustomobject]@{ Type = "OAB";        Cmd = "Get-OabVirtualDirectory"           }
    [pscustomobject]@{ Type = "PowerShell"; Cmd = "Get-PowerShellVirtualDirectory"    }
)

foreach ($srv in $casServers) {
    foreach ($vt in $vDirTypes) {
        $vdirs = & $vt.Cmd -Server $srv.Name -ErrorAction SilentlyContinue
        foreach ($vd in $vdirs) {
            $vDirList.Add([pscustomobject]@{
                Server      = $srv.Name
                vDirType    = $vt.Type
                Name        = $vd.Name
                InternalUrl = [string]$vd.InternalUrl
                ExternalUrl = [string]$vd.ExternalUrl
            })
        }
    }
    # Autodiscover is not a classic vDir — lives on Get-ClientAccessService
    $cas = Get-ClientAccessService $srv.Name -ErrorAction SilentlyContinue
    if ($cas) {
        $vDirList.Add([pscustomobject]@{
            Server      = $srv.Name
            vDirType    = "Autodiscover"
            Name        = "Autodiscover"
            InternalUrl = [string]$cas.AutoDiscoverServiceInternalUri
            ExternalUrl = ""
        })
    }
}
$vDirList | Export-EXData -FileName "Infrastructure_ClientAccessURLs.csv"

Write-Host "  Done in $([math]::Round(((Get-Date)-$sectionStart).TotalSeconds))s"
#endregion Infrastructure

# ============================================================
#region HybridConfig
# ============================================================
Write-Host "`n## [2/9] Hybrid Configuration ##" -ForegroundColor Cyan
$sectionStart = Get-Date

$hybridConfig = Get-HybridConfiguration -ErrorAction SilentlyContinue

if ($hybridConfig) {
    $hybridConfig | ConvertTo-FlatObject | Export-EXData -FileName "Hybrid_HybridConfiguration.csv"
} else {
    [pscustomobject]@{ Info = "No hybrid configuration found" } |
        Export-EXData -FileName "Hybrid_HybridConfiguration.csv"
}

$intraOrgConns = Get-IntraOrganizationConnector -ErrorAction SilentlyContinue
$intraOrgConns | ConvertTo-FlatObject | Export-EXData -FileName "Hybrid_IntraOrgConnectors.csv"

Get-AuthServer          -ErrorAction SilentlyContinue | ConvertTo-FlatObject | Export-EXData -FileName "Hybrid_AuthServers.csv"
Get-FederationTrust     -ErrorAction SilentlyContinue | ConvertTo-FlatObject | Export-EXData -FileName "Hybrid_FederationTrusts.csv"
Get-OnPremisesOrganization -ErrorAction SilentlyContinue | ConvertTo-FlatObject | Export-EXData -FileName "Hybrid_OnPremisesOrganization.csv"
Get-OrganizationRelationship -ErrorAction SilentlyContinue | ConvertTo-FlatObject | Export-EXData -FileName "Hybrid_OrganizationRelationships.csv"

# Human-readable hybrid summary
$hmaEnabled = $intraOrgConns | Where-Object { $_.Enabled -eq $true }
$oAuthServer = Get-AuthServer -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*ACS*" -or $_.AuthMetadataUrl -like "*.microsoft.com*" }
$summaryLines = @(
    "=== HYBRID SUMMARY ==="
    "Organisation             : $Kunde"
    "Generated                : $(Get-Date)"
    "---"
    "Hybrid deployed          : $(if ($hybridConfig) { 'YES' } else { 'NO' })"
)
if ($hybridConfig) {
    $summaryLines += "Hybrid Features          : $($hybridConfig.Features)"
    $summaryLines += "Centralized Mail Transport: $($hybridConfig.CentralizedTransportEnabled)"
    $summaryLines += "HybridDomains            : $(($hybridConfig.Domains | ForEach-Object {"$_"}) -join ', ')"
    $summaryLines += "SendingTransportServers  : $(($hybridConfig.SendingTransportServers | ForEach-Object {"$_"}) -join ', ')"
    $summaryLines += "ReceivingTransportServers: $(($hybridConfig.ReceivingTransportServers | ForEach-Object {"$_"}) -join ', ')"
}
$summaryLines += "---"
$summaryLines += "Hybrid Modern Auth (HMA)  : $(if ($hmaEnabled) { 'ENABLED (IntraOrgConnector active)' } else { 'DISABLED or not configured' })"
$summaryLines += "OAuth / AuthServer        : $(if ($oAuthServer) { 'YES' } else { 'NO' })"
$summaryLines += "Federation Trust          : $(if (Get-FederationTrust -ErrorAction SilentlyContinue) { 'YES' } else { 'NO' })"
$summaryLines | Out-File (Join-Path $logpath ($Kunde + "_Hybrid_Summary.txt")) -Encoding $csvEncoding -Force
$summaryLines | ForEach-Object { Write-Host "  $_" }

Write-Host "  Done in $([math]::Round(((Get-Date)-$sectionStart).TotalSeconds))s"
#endregion HybridConfig

# ============================================================
#region Organization
# ============================================================
Write-Host "`n## [3/9] Organization ##" -ForegroundColor Cyan
$sectionStart = Get-Date

# Full OrgConfig (all properties — ConvertTo-FlatObject handles the wide output)
$OrgSettings | ConvertTo-FlatObject | Export-EXData -FileName "Org_OrganizationConfig.csv"

$acceptedDomains = Get-AcceptedDomain
$acceptedDomains | ConvertTo-FlatObject | Export-EXData -FileName "Org_AcceptedDomains.csv"

Get-RemoteDomain                  | ConvertTo-FlatObject | Export-EXData -FileName "Org_RemoteDomains.csv"
Get-EmailAddressPolicy            -ErrorAction SilentlyContinue | ConvertTo-FlatObject | Export-EXData -FileName "Org_EmailAddressPolicies.csv"
Get-AddressList                   -ErrorAction SilentlyContinue | ConvertTo-FlatObject | Export-EXData -FileName "Org_AddressLists.csv"
Get-GlobalAddressList             -ErrorAction SilentlyContinue | ConvertTo-FlatObject | Export-EXData -FileName "Org_GlobalAddressLists.csv"
Get-OfflineAddressBook            -ErrorAction SilentlyContinue | ConvertTo-FlatObject | Export-EXData -FileName "Org_OfflineAddressBooks.csv"

# MX records for all accepted domains
$mxList = [System.Collections.Generic.List[PSObject]]::new()
foreach ($domain in $acceptedDomains) {
    $mxRecords = Resolve-DnsName $domain.DomainName -Type MX -ErrorAction SilentlyContinue |
                 Where-Object { $_.Type -eq 'MX' } |
                 Sort-Object Preference
    if ($mxRecords) {
        foreach ($mx in $mxRecords) {
            $mxList.Add([pscustomobject]@{
                AcceptedDomain = $domain.Name
                DomainName     = [string]$domain.DomainName
                DomainType     = [string]$domain.DomainType
                IsDefault      = $domain.Default
                MXHost         = $mx.NameExchange
                Preference     = $mx.Preference
            })
        }
    } else {
        $mxList.Add([pscustomobject]@{
            AcceptedDomain = $domain.Name
            DomainName     = [string]$domain.DomainName
            DomainType     = [string]$domain.DomainType
            IsDefault      = $domain.Default
            MXHost         = "(no MX record)"
            Preference     = ""
        })
    }
}
$mxList | Export-EXData -FileName "Org_MXRecords.csv"

Write-Host "  Done in $([math]::Round(((Get-Date)-$sectionStart).TotalSeconds))s"
#endregion Organization

# ============================================================
#region Transport
# ============================================================
Write-Host "`n## [4/9] Transport ##" -ForegroundColor Cyan
$sectionStart = Get-Date

Get-ReceiveConnector | ConvertTo-FlatObject | Export-EXData -FileName "Transport_ReceiveConnectors.csv"
Get-SendConnector    | ConvertTo-FlatObject | Export-EXData -FileName "Transport_SendConnectors.csv"

# Full transport rule export — conditions + actions expandiert
Get-TransportRule -ResultSize Unlimited | ConvertTo-FlatObject | Export-EXData -FileName "Transport_TransportRules.csv"

Get-JournalRule -ErrorAction SilentlyContinue | ConvertTo-FlatObject | Export-EXData -FileName "Transport_JournalRules_Global.csv"

# Per-database journaling (second journaling level in Exchange, separate from transport rules)
$dbJournal = Get-MailboxDatabase -ErrorAction SilentlyContinue |
    Where-Object { $_.JournalRecipient } |
    Select-Object Name, @{N="Server";E={[string]$_.Server}},
        @{N="JournalRecipient";E={[string]$_.JournalRecipient}}
if ($dbJournal) {
    $dbJournal | Export-EXData -FileName "Transport_JournalRules_PerDB.csv"
} else {
    Write-Host "    No per-database journaling configured"
}

Write-Host "  Done in $([math]::Round(((Get-Date)-$sectionStart).TotalSeconds))s"
#endregion Transport

# ============================================================
#region Policies
# ============================================================
Write-Host "`n## [5/9] Policies ##" -ForegroundColor Cyan
$sectionStart = Get-Date

Get-OwaMailboxPolicy         | ConvertTo-FlatObject | Export-EXData -FileName "Policies_OwaMailboxPolicies.csv"
Get-MobileDeviceMailboxPolicy | ConvertTo-FlatObject | Export-EXData -FileName "Policies_MobileDevicePolicies.csv"

Get-RetentionPolicy          -ErrorAction SilentlyContinue | ConvertTo-FlatObject | Export-EXData -FileName "Policies_RetentionPolicies.csv"
Get-RetentionPolicyTag       -ErrorAction SilentlyContinue | ConvertTo-FlatObject | Export-EXData -FileName "Policies_RetentionPolicyTags.csv"

Get-IRMConfiguration         -ErrorAction SilentlyContinue | ConvertTo-FlatObject | Export-EXData -FileName "Policies_IRMConfiguration.csv"
Get-RMSTemplate              -ErrorAction SilentlyContinue | ConvertTo-FlatObject | Export-EXData -FileName "Policies_RMSTemplates.csv"

Write-Host "  Done in $([math]::Round(((Get-Date)-$sectionStart).TotalSeconds))s"
#endregion Policies

# ============================================================
#region Mailboxes
# ============================================================
Write-Host "`n## [6/9] Mailboxes ##" -ForegroundColor Cyan
$sectionStart = Get-Date

Write-Host "  Loading all mailboxes..." -NoNewline
$users = Get-Mailbox -ResultSize Unlimited
Write-Host " $($users.Count) found."

$users | Select-Object `
    UserPrincipalName, Identity, DisplayName, Alias,
    PrimarySmtpAddress,
    @{N="EmailAddresses";           E={ ($_.EmailAddresses            | ForEach-Object {"$_"}) -join $sep }},
    RecipientTypeDetails, AccountDisabled, IsInactiveMailbox,
    ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward,
    LitigationHoldEnabled, LitigationHoldDuration, LitigationHoldDate, LitigationHoldOwner,
    RetentionPolicy, RetentionHoldEnabled, EndDateForRetentionHold, StartDateForRetentionHold,
    SingleItemRecoveryEnabled,
    AuditEnabled,
    @{N="AuditAdmin";               E={ ($_.AuditAdmin               | ForEach-Object {"$_"}) -join $sep }},
    @{N="AuditDelegate";            E={ ($_.AuditDelegate            | ForEach-Object {"$_"}) -join $sep }},
    HiddenFromAddressListsEnabled,
    MaxSendSize, MaxReceiveSize,
    ProhibitSendQuota, ProhibitSendReceiveQuota, IssueWarningQuota, UseDatabaseQuotaDefaults,
    ArchiveStatus, ArchiveName, ArchiveDatabase,
    ExchangeUserAccountControl, Office, Database,
    CustomAttribute1, CustomAttribute2, CustomAttribute3,
    CustomAttribute4, CustomAttribute5,
    @{N="ExtensionCustomAttribute1";E={ ($_.ExtensionCustomAttribute1 | ForEach-Object {"$_"}) -join $sep }},
    @{N="ExtensionCustomAttribute2";E={ ($_.ExtensionCustomAttribute2 | ForEach-Object {"$_"}) -join $sep }} `
    | Export-EXData -FileName "Mailboxes_Mailboxes.csv"

# Mailbox counts by type
$countFile = Join-Path $logpath ($Kunde + "_Mailboxes_CountByType.csv")
$users | Group-Object RecipientTypeDetails | Select-Object Count, Name |
    Export-Csv $countFile -Delimiter $csvDelim -NoTypeInformation -Encoding $csvEncoding
Get-Mailbox -GroupMailbox  -ResultSize Unlimited -ErrorAction SilentlyContinue |
    Group-Object RecipientTypeDetails | Select-Object Count, Name |
    Export-Csv $countFile -Delimiter $csvDelim -NoTypeInformation -Encoding $csvEncoding -Append
Get-Mailbox -PublicFolder  -ResultSize Unlimited -ErrorAction SilentlyContinue |
    Group-Object RecipientTypeDetails | Select-Object Count, Name |
    Export-Csv $countFile -Delimiter $csvDelim -NoTypeInformation -Encoding $csvEncoding -Append
Get-Mailbox -Archive       -ResultSize Unlimited -ErrorAction SilentlyContinue |
    Group-Object RecipientTypeDetails | Select-Object Count, @{N="Name";E={"MBX with Archive"}} |
    Export-Csv $countFile -Delimiter $csvDelim -NoTypeInformation -Encoding $csvEncoding -Append
Write-Host "    -> $($Kunde)_Mailboxes_CountByType.csv"

# --- Mailbox Permissions ---
Write-Host "  Collecting mailbox permissions (FullAccess, SendAs, SendOnBehalf, Forwarding)..."
$total = $users.Count
$i     = 0

$fullAccessList = [System.Collections.Generic.List[PSObject]]::new()
$sendAsList     = [System.Collections.Generic.List[PSObject]]::new()

foreach ($user in $users) {
    $i++
    Write-Progress -Activity "Get-MailboxPermission / Get-RecipientPermission" `
        -PercentComplete ($i / $total * 100) -Status "$i/$total - $($user.Alias)"

    # FullAccess
    $fa = Get-MailboxPermission $user -ErrorAction SilentlyContinue |
          Where-Object { $_.User -notmatch "SELF" -and $_.IsInherited -eq $false }
    foreach ($p in $fa) {
        $fullAccessList.Add([pscustomobject]@{
            UserPrincipalName = $user.UserPrincipalName
            Identity          = [string]$p.Identity
            User              = [string]$p.User
            AccessRights      = ($p.AccessRights | ForEach-Object {"$_"}) -join $sep
        })
    }

    # SendAs
    $sa = Get-RecipientPermission $user.Name -AccessRights SendAs -ErrorAction SilentlyContinue |
          Where-Object { $_.Trustee -notmatch "SELF" }
    foreach ($p in $sa) {
        $sendAsList.Add([pscustomobject]@{
            UserPrincipalName = $user.UserPrincipalName
            Identity          = [string]$p.Identity
            Trustee           = [string]$p.Trustee
            AccessRights      = ($p.AccessRights | ForEach-Object {"$_"}) -join $sep
        })
    }
}
Write-Progress -Activity "Get-MailboxPermission / Get-RecipientPermission" -Completed

$fullAccessList | Export-EXData -FileName "Mailboxes_FullAccess.csv"
$sendAsList     | Export-EXData -FileName "Mailboxes_SendAs.csv"

$users |
    Select-Object UserPrincipalName, Alias,
        @{N="GrantSendOnBehalfTo";E={ ($_.GrantSendOnBehalfTo | ForEach-Object {"$_"}) -join $sep }} |
    Where-Object { $_.GrantSendOnBehalfTo } |
    Export-EXData -FileName "Mailboxes_SendOnBehalf.csv"

$users |
    Select-Object UserPrincipalName, Alias, ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward |
    Where-Object { $_.ForwardingAddress -or $_.ForwardingSmtpAddress } |
    Export-EXData -FileName "Mailboxes_ForwardTo.csv"

# --- Mailbox Statistics ---
Write-Host "  Collecting mailbox statistics..."
$i = 0
$mbxStatList = [System.Collections.Generic.List[PSObject]]::new()
$arcStatList = [System.Collections.Generic.List[PSObject]]::new()
$archiveUsers = $users | Where-Object { $_.ArchiveName }

foreach ($user in $users) {
    $i++
    Write-Progress -Activity "Get-MailboxStatistics" `
        -PercentComplete ($i / $total * 100) -Status "$i/$total - $($user.Alias)"
    $stats = Get-MailboxStatistics $user -ErrorAction SilentlyContinue |
        Select-Object DisplayName, MailboxType, MailboxTypeDetail, Database,
            ItemCount, TotalItemSize,
            @{N="TotalItemSizeBytes";E={$_.TotalItemSize.Value.ToBytes()}},
            LastLoggedOnUserAccount, LastLogonTime
    if ($stats) {
        $stats | Add-Member -MemberType NoteProperty -Name 'UserPrincipalName' -Value $user.UserPrincipalName
        $mbxStatList.Add($stats)
    }
}
Write-Progress -Activity "Get-MailboxStatistics" -Completed
$mbxStatList | Select-Object UserPrincipalName, DisplayName, MailboxType, MailboxTypeDetail, Database,
    ItemCount, TotalItemSize, TotalItemSizeBytes, LastLoggedOnUserAccount, LastLogonTime |
    Export-EXData -FileName "Mailboxes_MailboxStatistics.csv"

$i = 0
foreach ($user in $archiveUsers) {
    $i++
    Write-Progress -Activity "Get-MailboxStatistics (Archive)" `
        -PercentComplete ($i / [Math]::Max($archiveUsers.Count,1) * 100) -Status "$i/$($archiveUsers.Count) - $($user.Alias)"
    $stats = Get-MailboxStatistics $user -Archive -ErrorAction SilentlyContinue |
        Select-Object DisplayName, MailboxType, MailboxTypeDetail, Database,
            ItemCount, TotalItemSize,
            @{N="TotalItemSizeBytes";E={$_.TotalItemSize.Value.ToBytes()}},
            LastLoggedOnUserAccount, LastLogonTime
    if ($stats) {
        $stats | Add-Member -MemberType NoteProperty -Name 'UserPrincipalName' -Value $user.UserPrincipalName
        $arcStatList.Add($stats)
    }
}
Write-Progress -Activity "Get-MailboxStatistics (Archive)" -Completed
$arcStatList | Select-Object UserPrincipalName, DisplayName, MailboxType, MailboxTypeDetail, Database,
    ItemCount, TotalItemSize, TotalItemSizeBytes, LastLoggedOnUserAccount, LastLogonTime |
    Export-EXData -FileName "Mailboxes_MailboxStatistics_Archive.csv"

Write-Host "  Done in $([math]::Round(((Get-Date)-$sectionStart).TotalSeconds))s"
#endregion Mailboxes

# ============================================================
#region Groups
# ============================================================
Write-Host "`n## [7/9] Groups ##" -ForegroundColor Cyan
$sectionStart = Get-Date

Write-Host "  Loading distribution / mail-enabled security groups..." -NoNewline
$allGroups = Get-DistributionGroup -ResultSize Unlimited
Write-Host " $($allGroups.Count) found."

$allGroups | Select-Object `
    Name, DisplayName, Alias, PrimarySmtpAddress,
    @{N="EmailAddresses";                    E={ ($_.EmailAddresses                    | ForEach-Object {"$_"}) -join $sep }},
    GroupType, RecipientTypeDetails,
    @{N="ManagedBy";                         E={ ($_.ManagedBy                         | ForEach-Object {"$_"}) -join $sep }},
    MemberJoinRestriction, MemberDepartRestriction,
    RequireSenderAuthenticationEnabled,
    HiddenFromAddressListsEnabled,
    @{N="AcceptMessagesOnlyFrom";            E={ ($_.AcceptMessagesOnlyFrom            | ForEach-Object {"$_"}) -join $sep }},
    @{N="AcceptMessagesOnlyFromDLMembers";   E={ ($_.AcceptMessagesOnlyFromDLMembers   | ForEach-Object {"$_"}) -join $sep }},
    @{N="RejectMessagesFrom";               E={ ($_.RejectMessagesFrom               | ForEach-Object {"$_"}) -join $sep }},
    @{N="GrantSendOnBehalfTo";              E={ ($_.GrantSendOnBehalfTo              | ForEach-Object {"$_"}) -join $sep }},
    CustomAttribute1, CustomAttribute2, CustomAttribute3 `
    | Export-EXData -FileName "Groups_DistributionGroups.csv"

# Members
Write-Host "  Collecting group members..."
$memberList = [System.Collections.Generic.List[PSObject]]::new()
$gi = 0
foreach ($grp in $allGroups) {
    $gi++
    Write-Progress -Activity "Get-DistributionGroupMember" `
        -PercentComplete ($gi / $allGroups.Count * 100) -Status "$gi/$($allGroups.Count) - $($grp.Alias)"
    $members = Get-DistributionGroupMember $grp -ResultSize Unlimited -ErrorAction SilentlyContinue
    foreach ($m in $members) {
        $memberList.Add([pscustomobject]@{
            GroupName                  = $grp.Name
            GroupSmtpAddress           = [string]$grp.PrimarySmtpAddress
            GroupType                  = [string]$grp.GroupType
            MemberName                 = $m.Name
            MemberAlias                = $m.Alias
            MemberSmtpAddress          = [string]$m.PrimarySmtpAddress
            MemberRecipientType        = [string]$m.RecipientType
            MemberRecipientTypeDetails = [string]$m.RecipientTypeDetails
        })
    }
}
Write-Progress -Activity "Get-DistributionGroupMember" -Completed
$memberList | Export-EXData -FileName "Groups_DistributionGroupMembers.csv"

# Dynamic distribution groups
Write-Host "  Loading dynamic distribution groups..." -NoNewline
$dynGroups = Get-DynamicDistributionGroup -ResultSize Unlimited -ErrorAction SilentlyContinue
Write-Host " $($dynGroups.Count) found."
$dynGroups | Select-Object `
    Name, DisplayName, Alias, PrimarySmtpAddress,
    @{N="EmailAddresses";E={ ($_.EmailAddresses | ForEach-Object {"$_"}) -join $sep }},
    RecipientFilter, LdapRecipientFilter, RecipientContainer,
    @{N="ManagedBy";E={ ($_.ManagedBy | ForEach-Object {"$_"}) -join $sep }},
    HiddenFromAddressListsEnabled `
    | Export-EXData -FileName "Groups_DynamicDistributionGroups.csv"

Write-Host "  Done in $([math]::Round(((Get-Date)-$sectionStart).TotalSeconds))s"
#endregion Groups

# ============================================================
#region Other Recipients
# ============================================================
Write-Host "`n## [8/9] Other Recipients ##" -ForegroundColor Cyan
$sectionStart = Get-Date

Write-Host "  Mail contacts..." -NoNewline
$contacts = Get-MailContact -ResultSize Unlimited -ErrorAction SilentlyContinue
Write-Host " $($contacts.Count) found."
$contacts | Select-Object `
    Name, Alias, DisplayName, ExternalEmailAddress,
    @{N="EmailAddresses";E={ ($_.EmailAddresses | ForEach-Object {"$_"}) -join $sep }},
    HiddenFromAddressListsEnabled, UsePreferMessageFormat,
    CustomAttribute1, CustomAttribute2, CustomAttribute3 `
    | Export-EXData -FileName "Recipients_MailContacts.csv"

Write-Host "  Mail users..." -NoNewline
$mailUsers = Get-MailUser -ResultSize Unlimited -ErrorAction SilentlyContinue
Write-Host " $($mailUsers.Count) found."
$mailUsers | Select-Object `
    Name, Alias, DisplayName, UserPrincipalName, ExternalEmailAddress,
    @{N="EmailAddresses";E={ ($_.EmailAddresses | ForEach-Object {"$_"}) -join $sep }},
    HiddenFromAddressListsEnabled, RecipientTypeDetails `
    | Export-EXData -FileName "Recipients_MailUsers.csv"

# Resource mailboxes + CalendarProcessing
Write-Host "  Resource mailboxes (Room + Equipment) + CalendarProcessing..." -NoNewline
$resourceMBX = $users | Where-Object { $_.RecipientTypeDetails -in "RoomMailbox","EquipmentMailbox" }
Write-Host " $($resourceMBX.Count) found."
$calProcList = [System.Collections.Generic.List[PSObject]]::new()
foreach ($r in $resourceMBX) {
    $cp = Get-CalendarProcessing $r -ErrorAction SilentlyContinue | ConvertTo-FlatObject
    if ($cp) {
        $cp | Add-Member -MemberType NoteProperty -Name 'PrimarySmtpAddress'   -Value ([string]$r.PrimarySmtpAddress)   -Force
        $cp | Add-Member -MemberType NoteProperty -Name 'RecipientTypeDetails' -Value ([string]$r.RecipientTypeDetails) -Force
        $cp | Add-Member -MemberType NoteProperty -Name 'DisplayName'          -Value ([string]$r.DisplayName)          -Force
        $calProcList.Add($cp)
    }
}
$calProcList | Export-EXData -FileName "Recipients_ResourceMBX_CalendarProcessing.csv"

Write-Host "  Done in $([math]::Round(((Get-Date)-$sectionStart).TotalSeconds))s"
#endregion Other Recipients

# ============================================================
#region Public Folders
# ============================================================
Write-Host "`n## [9/9] Public Folders ##" -ForegroundColor Cyan
$sectionStart = Get-Date

try {
    $PF = Get-PublicFolder -Recurse -ResultSize Unlimited -ErrorAction Stop

    $PF | Select-Object Name, ParentPath | Export-EXData -FileName "PF_Structure.csv"

    Get-PublicFolderStatistics -ResultSize Unlimited |
        Select-Object Name, @{N="Identity";E={$_.Identity.ToString()}},
            ItemCount, TotalItemSize, LastModificationTime |
        Export-EXData -FileName "PF_Statistics.csv"

    $PF | Get-PublicFolderClientPermission |
        Select-Object Identity, User,
            @{N="AccessRights";E={ ($_.AccessRights | ForEach-Object {"$_"}) -join $sep }} |
        Export-EXData -FileName "PF_Permissions.csv"

    # Mail-enabled public folders
    $mepf = Get-MailPublicFolder -ResultSize Unlimited -ErrorAction SilentlyContinue
    $mepf | Select-Object `
        Name, Alias, PrimarySmtpAddress,
        @{N="EmailAddresses";       E={ ($_.EmailAddresses       | ForEach-Object {"$_"}) -join $sep }},
        @{N="GrantSendOnBehalfTo";  E={ ($_.GrantSendOnBehalfTo  | ForEach-Object {"$_"}) -join $sep }},
        HiddenFromAddressListsEnabled, ContentMailbox `
        | Export-EXData -FileName "PF_MailPublicFolders.csv"

    # SendAs on mail-enabled PFs (via ADPermission)
    $mepfSendAs = [System.Collections.Generic.List[PSObject]]::new()
    foreach ($mpf in $mepf) {
        $adPerms = Get-ADPermission $mpf -ErrorAction SilentlyContinue |
                   Where-Object { $_.ExtendedRights -like "*Send-As*" }
        foreach ($p in $adPerms) {
            $mepfSendAs.Add([pscustomobject]@{
                Identity = [string]$p.Identity
                User     = [string]$p.User
                Deny     = $p.Deny
            })
        }
    }
    $mepfSendAs | Export-EXData -FileName "PF_MailPublicFolders_SendAs.csv"

    # SendOnBehalf on mail-enabled PFs
    $mepf | Where-Object { $_.GrantSendOnBehalfTo } |
        Select-Object Name, PrimarySmtpAddress,
            @{N="GrantSendOnBehalfTo";E={ ($_.GrantSendOnBehalfTo | ForEach-Object {"$_"}) -join $sep }} |
        Export-EXData -FileName "PF_MailPublicFolders_SendOnBehalf.csv"
}
catch {
    if (-not $PF) {
        Write-Host "  No public folders found." -ForegroundColor Yellow
        [pscustomobject]@{ Info = "No public folders found" } | Export-EXData -FileName "PF_Structure.csv"
    } else {
        Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}
Write-Host "  Done in $([math]::Round(((Get-Date)-$sectionStart).TotalSeconds))s"
#endregion Public Folders

# ============================================================
#region Optional — Inbox Rules
# ============================================================
if ($includeInboxRules) {
    Write-Host "`n## [Opt] Inbox Rules ##" -ForegroundColor Cyan
    $sectionStart = Get-Date
    $inboxRulesDir = Join-Path $logpath "InboxRules"
    if (-not (Test-Path $inboxRulesDir)) { New-Item -ItemType Directory -Path $inboxRulesDir | Out-Null }

    $total = $users.Count; $i = 0
    foreach ($user in $users) {
        $i++
        Write-Progress -Activity "Get-InboxRule" `
            -PercentComplete ($i / $total * 100) -Status "$i/$total - $($user.Alias)"
        $rules = Get-InboxRule -Mailbox $user -ErrorAction SilentlyContinue
        if ($rules) {
            $rules | ConvertTo-FlatObject |
                Export-Csv (Join-Path $inboxRulesDir "$($user.Alias).csv") `
                    -Delimiter $csvDelim -NoTypeInformation -Encoding $csvEncoding
        }
    }
    Write-Progress -Activity "Get-InboxRule" -Completed
    Write-Host "  Done in $([math]::Round(((Get-Date)-$sectionStart).TotalSeconds))s"
}
#endregion Optional — Inbox Rules

# ============================================================
#region Optional — Mobile Devices
# ============================================================
if ($includeMobileDevices) {
    Write-Host "`n## [Opt] Mobile Devices ##" -ForegroundColor Cyan
    $sectionStart = Get-Date
    $mobileList = [System.Collections.Generic.List[PSObject]]::new()
    $total = $users.Count; $i = 0

    foreach ($user in $users) {
        $i++
        Write-Progress -Activity "Get-MobileDeviceStatistics" `
            -PercentComplete ($i / $total * 100) -Status "$i/$total - $($user.Alias)"
        $devices = Get-MobileDeviceStatistics -Mailbox $user -ErrorAction SilentlyContinue
        foreach ($d in $devices) {
            $mobileList.Add([pscustomobject]@{
                UserPrincipalName   = $user.UserPrincipalName
                DeviceFriendlyName  = $d.DeviceFriendlyName
                DeviceModel         = $d.DeviceModel
                DeviceOS            = $d.DeviceOS
                DeviceType          = $d.DeviceType
                LastSyncAttemptTime = $d.LastSyncAttemptTime
                LastSuccessSync     = $d.LastSuccessSync
                Status              = $d.Status
                DevicePolicyApplied = $d.DevicePolicyApplied
            })
        }
    }
    Write-Progress -Activity "Get-MobileDeviceStatistics" -Completed
    $mobileList | Export-EXData -FileName "Optional_MobileDevices.csv"
    Write-Host "  Done in $([math]::Round(((Get-Date)-$sectionStart).TotalSeconds))s"
}
#endregion Optional — Mobile Devices

# ============================================================
#region Optional — Mailbox Folder Permissions
# ============================================================
if ($includeMailboxFolderPermissions) {
    Write-Host "`n## [Opt] Mailbox Folder Permissions ##" -ForegroundColor Cyan
    $sectionStart = Get-Date
    $mbxPermsDir = Join-Path $logpath "MBXPerms"
    if (-not (Test-Path $mbxPermsDir)) { New-Item -ItemType Directory -Path $mbxPermsDir | Out-Null }

    $total = $users.Count; $i = 0
    foreach ($mbx in $users) {
        $i++
        Write-Progress -Activity "Mailbox folder permissions" `
            -PercentComplete ($i / $total * 100) -Status "$i/$total - $($mbx.Alias)" -Id 1
        $result = Get-MailboxFolderPermissionsForAllFolders -Mailbox $mbx
        if ($result.Count -gt 0) {
            $result | Export-Csv (Join-Path $mbxPermsDir "$($mbx.Alias).csv") `
                -Delimiter $csvDelim -NoTypeInformation -Encoding $csvEncoding
        }
    }
    Write-Progress -Activity "Mailbox folder permissions" -Completed -Id 1
    Write-Host "  Done in $([math]::Round(((Get-Date)-$sectionStart).TotalSeconds))s"
}
#endregion Optional — Mailbox Folder Permissions

# ============================================================
#region Merge & Summary
# ============================================================
Write-Host "`n## Merge & Summary ##" -ForegroundColor Cyan

$mbxInfos     = Import-Csv (Join-Path $logpath ($Kunde + "_Mailboxes_Mailboxes.csv"))              -Delimiter $csvDelim
$mbxStatInfos = Import-Csv (Join-Path $logpath ($Kunde + "_Mailboxes_MailboxStatistics.csv"))       -Delimiter $csvDelim
$arcStatInfos = Import-Csv (Join-Path $logpath ($Kunde + "_Mailboxes_MailboxStatistics_Archive.csv")) -Delimiter $csvDelim

$statsIndex = @{}
foreach ($s in $mbxStatInfos) { if ($s.UserPrincipalName) { $statsIndex[$s.UserPrincipalName.ToLower()] = $s } }

$arcIndex = @{}
foreach ($a in $arcStatInfos) { if ($a.UserPrincipalName) { $arcIndex[$a.UserPrincipalName.ToLower()] = $a } }

$merged = foreach ($m in $mbxInfos) {
    $key     = if ($m.UserPrincipalName) { $m.UserPrincipalName.ToLower() } else { "" }
    $stat    = if ($key -and $statsIndex.ContainsKey($key)) { $statsIndex[$key] } else { $null }
    $arcStat = if ($key -and $arcIndex.ContainsKey($key))   { $arcIndex[$key]   } else { $null }

    $props = [ordered]@{}
    foreach ($p in $m.PSObject.Properties.Name)    { $props[$p] = $m.$p }
    if ($stat) {
        foreach ($p in $stat.PSObject.Properties.Name) { if (-not $props.Contains($p)) { $props[$p] = $stat.$p } }
    }

    $primBytes = [long]0; $arcBytes = [long]0; $primCount = [long]0; $arcCount = [long]0
    if ($stat)    {
        [void][int64]::TryParse(($stat.TotalItemSizeBytes    -replace '[^0-9\-]',''), [ref]$primBytes)
        [void][int64]::TryParse(($stat.ItemCount              -replace '[^0-9\-]',''), [ref]$primCount)
    }
    if ($arcStat) {
        [void][int64]::TryParse(($arcStat.TotalItemSizeBytes -replace '[^0-9\-]',''), [ref]$arcBytes)
        [void][int64]::TryParse(($arcStat.ItemCount           -replace '[^0-9\-]',''), [ref]$arcCount)
    }
    $props['CombinedItemCount']      = $primCount + $arcCount
    $props['CombinedItemSizeBytes']  = $primBytes + $arcBytes
    $props['CombinedItemSize']       = Convert-BytesToSizeString -Bytes ($primBytes + $arcBytes)
    $props['Archive_ItemCount']      = $arcCount
    $props['Archive_ItemSizeBytes']  = $arcBytes

    [pscustomobject]$props
}

$merged | Export-Csv (Join-Path $logpath ($Kunde + "_Merged_MailboxesAndStatistics.csv")) `
    -Delimiter $csvDelim -NoTypeInformation -Encoding $csvEncoding -Force
Write-Host "    -> $($Kunde)_Merged_MailboxesAndStatistics.csv ($($merged.Count) records)"

$measure = $merged | Where-Object { $_.CombinedItemSizeBytes -match '^\d+$' } |
    Measure-Object -Property CombinedItemSizeBytes -Sum -Average -Maximum -Minimum

[pscustomobject]@{
    Count  = $measure.Count
    Min_GB = "{0:N2}" -f ([double]$measure.Minimum / 1GB)
    Max_GB = "{0:N2}" -f ([double]$measure.Maximum / 1GB)
    Sum_GB = "{0:N2}" -f ([double]$measure.Sum     / 1GB)
    Avg_GB = "{0:N2}" -f ([double]$measure.Average / 1GB)
    Min_MB = "{0:N2}" -f ([double]$measure.Minimum / 1MB)
    Max_MB = "{0:N2}" -f ([double]$measure.Maximum / 1MB)
    Sum_MB = "{0:N2}" -f ([double]$measure.Sum     / 1MB)
    Avg_MB = "{0:N2}" -f ([double]$measure.Average / 1MB)
} | Out-File (Join-Path $logpath ($Kunde + "_Merged_MailboxSizeSummary.txt")) -Encoding $csvEncoding -Force
Write-Host "    -> $($Kunde)_Merged_MailboxSizeSummary.txt"

#endregion Merge & Summary

Pop-Location

$totalSecs = [math]::Round(((Get-Date) - $runStart).TotalSeconds)
Write-Host "`n$("=" * 60)" -ForegroundColor White
Write-Host "All done.  Total runtime: ${totalSecs}s" -ForegroundColor Green
Write-Host "Output: $logpath" -ForegroundColor Green
