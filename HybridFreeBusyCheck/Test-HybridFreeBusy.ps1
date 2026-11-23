#Requires -Version 5.1
<#
.SYNOPSIS
    Checks the on-premises side of hybrid free/busy (on-premises -> Exchange Online) step by step.

.DESCRIPTION
    Walks the prerequisites an Exchange server needs to look up free/busy of an Exchange Online
    mailbox with the dedicated Exchange hybrid app, bottom-up. The step numbers match the
    decision tree in the README:

      2  target is a remote mailbox routed to the tenant        (only with -ExoMailbox)
      3  IntraOrganizationConnector enabled for the tenant domain
      4  server build supports the dedicated hybrid app (and Graph)
      5  dedicated hybrid app enabled (setting override)
      6  EvoSts auth server points to the app ID
      7  Exchange gets a token for that app, auth certificate valid, login.microsoftonline.com reachable
      8  Graph or EWS route, endpoints reachable, recent errors in the EWS log

    Steps 1 and 9 are the client's view - use ExchangeTester (Free/Busy, On-Prem -> EXO) for those.

    The script only reads. Commands to fix a finding are printed as suggestions.
    The OAuth test, the EWS log and the auth certificate are checked on THIS server; run the
    script on each Mailbox server that answers availability requests.

.PARAMETER OnPremMailbox
    An on-premises mailbox for Test-OAuthConnectivity (step 7). Without it, step 7 skips the token test.

.PARAMETER ExoMailbox
    An Exchange Online mailbox (its on-premises remote mailbox object) for step 2. Also used to
    filter the EWS log.

.PARAMETER OutputPath
    Also write all results to this CSV file.

.PARAMETER SkipNetwork
    Skip the TCP 443 tests.

.PARAMETER EwsLogFiles
    Number of most recent EWS log files to search. Default 3.

.PARAMETER PassThru
    Return the result objects.

.EXAMPLE
    .\Test-HybridFreeBusy.ps1 -OnPremMailbox anna@contoso.com -ExoMailbox ben@contoso.com
    .\Test-HybridFreeBusy.ps1 -OnPremMailbox anna@contoso.com -OutputPath C:\Temp\fb-check.csv

.NOTES
    Exchange Server 2016/2019/SE, Windows PowerShell 5.1, Exchange Management Shell.
    Exit code: 0 = no FAIL, 1 = at least one FAIL.
    Source for builds, override and OAuth test: Microsoft Learn, "Deploy dedicated Exchange hybrid app".
#>
[CmdletBinding()]
param(
    [string]$OnPremMailbox,
    [string]$ExoMailbox,
    [string]$OutputPath,
    [switch]$SkipNetwork,
    [ValidateRange(1, 50)][int]$EwsLogFiles = 3,
    [switch]$PassThru
)

$script:Version = '1.0.0'
$ErrorActionPreference = 'Stop'

# Voraussetzungen. Der Block zwischen den Markern stammt aus build/Prolog.OnPrem.ps1
# und wird von build/Sync-Prolog.ps1 gepflegt - nicht von Hand aendern.
$script:RequiredCmdlets = @('Get-ExchangeServer', 'Get-IntraOrganizationConnector', 'Get-SettingOverride', 'Get-AuthServer', 'Get-AuthConfig')
$script:ToolIsGui       = $false
# <prolog:onprem v1 - Quelle: build/Prolog.OnPrem.ps1, eingefügt von build/Sync-Prolog.ps1.
#                     NICHT von Hand ändern - build/Test-Prolog.ps1 meldet jede Abweichung.>
# Erwartet davor gesetzt:
#   $script:RequiredCmdlets = @('Get-Queue')      (Cmdlets, die das Werkzeug wirklich braucht)
#   $script:ToolIsGui       = $true|$false        (bei $true kommen Fehler als MessageBox)
function Initialize-OnPremPrerequisite {
    [CmdletBinding()]
    param(
        [string[]]$Cmdlets = $script:RequiredCmdlets,
        [bool]$Gui         = [bool]$script:ToolIsGui
    )

    function Stop-WithReason([string]$Text) {
        # Eine mit -noConsole gebaute EXE hat kein Fenster für Write-Host oder throw;
        # ohne MessageBox bliebe der Start wortlos erfolglos.
        if ($Gui) {
            try {
                Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
                [void][System.Windows.Forms.MessageBox]::Show($Text, 'Voraussetzung fehlt', 'OK', 'Warning')
            } catch { Write-Warning $Text }
        }
        throw $Text
    }

    # 1. Das Exchange-Snap-in ist .NET-Framework-Code und lädt nur in Windows PowerShell 5.1.
    #    In PowerShell 7 scheitert Add-PSSnapin mit einer Meldung, die das nicht verrät.
    if ($PSVersionTable.PSEdition -eq 'Core') {
        Stop-WithReason ("Dieses Werkzeug braucht die Exchange-Verwaltungsshell und läuft deshalb nur in " +
                         "Windows PowerShell 5.1, nicht in PowerShell $($PSVersionTable.PSVersion).`n`n" +
                         'Über "Exchange Management Shell" starten oder powershell.exe statt pwsh.exe verwenden.')
    }

    # 2. Snap-in laden, falls die Sitzung es noch nicht hat (etwa beim Start per Doppelklick
    #    oder aus einer EXE heraus - dort ist die EMS-Umgebung nicht vorhanden).
    if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
        if (-not (Get-PSSnapin -Registered -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
            Stop-WithReason ("Auf diesem Rechner ist keine Exchange-Verwaltungsshell installiert.`n`n" +
                             'Dieses Werkzeug gehört auf einen Exchange-Server oder einen Rechner mit den Exchange-Verwaltungswerkzeugen.')
        }
        try { Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop }
        catch { Stop-WithReason "Das Exchange-Snap-in ließ sich nicht laden: $($_.Exception.Message)" }
    }

    # 3. Snap-in geladen heißt noch nicht berechtigt: auch on-premises blendet RBAC
    #    Cmdlets aus, für die die Rolle nicht reicht.
    $fehlt = @($Cmdlets | Where-Object { -not (Get-Command $_ -ErrorAction SilentlyContinue) })
    if ($fehlt) {
        Stop-WithReason ("Die Exchange-Verwaltungsshell ist geladen, aber diese Cmdlets fehlen:`n" +
                         ("    " + ($fehlt -join "`n    ")) +
                         "`n`nDas ist eine Frage der RBAC-Rolle, nicht der Installation.")
    }

    # 4. Exchange gibt Datumsangaben und Zahlen in der Sprache des Servers zurück.
    #    Ohne feste Kultur brechen Vergleiche und Parser auf deutschen Systemen.
    [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::InvariantCulture
}
Initialize-OnPremPrerequisite
# </prolog:onprem>

# ---------------------------------------------------------------------------------------------
# Pure evaluation functions: plain objects in, result rows out. No Exchange calls in here, so
# tests can feed them with mock objects.
# ---------------------------------------------------------------------------------------------

$script:FirstPartyAppId = '00000002-0000-0ff1-ce00-000000000000'
$script:GraphMinBuild   = [version]'15.2.2562.41'

function New-CheckResult {
    param([int]$Step, [string]$Check, [ValidateSet('OK', 'WARN', 'FAIL', 'INFO', 'SKIP')][string]$Status,
          [string]$Finding, [string]$Fix = '')
    [pscustomobject]@{ Step = $Step; Status = $Status; Check = $Check; Finding = $Finding; Fix = $Fix }
}

function ConvertTo-ExchangeBuild {
    # "Version 15.2 (Build 2562.17)" -> 15.2.2562.17; a plain "15.2.2562.17" also works.
    param($Value)
    $s = "$Value"
    $m = [regex]::Match($s, 'Version\s+(\d+)\.(\d+)\s+\(Build\s+(\d+)\.(\d+)\)')
    if ($m.Success) { return [version]('{0}.{1}.{2}.{3}' -f $m.Groups[1].Value, $m.Groups[2].Value, $m.Groups[3].Value, $m.Groups[4].Value) }
    $m = [regex]::Match($s, '(\d+)\.(\d+)\.(\d+)\.(\d+)')
    if ($m.Success) { return [version]$m.Value }
    $null
}

function Get-DedicatedAppMinBuild {
    # Minimum builds from Microsoft Learn, "Supported Exchange Server versions".
    # Returns the minimum for the build's product line, or $null when that line has none.
    param([version]$Build)
    if (-not $Build) { return $null }
    if ($Build.Major -ne 15) { return $null }
    if ($Build.Minor -eq 2 -and $Build.Build -ge 2562) { return [version]'15.2.2562.17' }   # SE
    if ($Build.Minor -eq 2 -and $Build.Build -eq 1748) { return [version]'15.2.1748.24' }   # 2019 CU15 + Apr 2025 HU
    if ($Build.Minor -eq 2 -and $Build.Build -eq 1544) { return [version]'15.2.1544.25' }   # 2019 CU14 + Apr 2025 HU
    if ($Build.Minor -eq 1 -and $Build.Build -eq 2507) { return [version]'15.1.2507.55' }   # 2016 CU23 + Apr 2025 HU
    $null
}

function Test-ServerBuild {
    # Step 4. $Servers: objects with Name, DisplayBuild ([version]), FileBuild ([version] or $null).
    # FileBuild comes from ExSetup.exe on this server; AdminDisplayVersion does not include
    # hotfix updates, so a low DisplayBuild alone is only a warning.
    param([object[]]$Servers)
    foreach ($s in $Servers) {
        if (-not $s.FileBuild -and -not $s.DisplayBuild) {
            New-CheckResult 4 "Build $($s.Name)" 'WARN' 'Build could not be read' 'Check ExSetup.exe on that server: (Get-Command ExSetup.exe).FileVersionInfo.ProductVersion'
            continue
        }
        $build  = if ($s.FileBuild) { [version]$s.FileBuild } else { [version]$s.DisplayBuild }
        $source = if ($s.FileBuild) { 'ExSetup.exe' } else { 'AdminDisplayVersion' }
        $min    = Get-DedicatedAppMinBuild $build
        $graph  = if ($build -ge $script:GraphMinBuild) { 'Graph route supported' } else { 'EWS route only (Graph needs 15.2.2562.41)' }
        $check  = "Build $($s.Name)"
        if (-not $min) {
            if ($s.FileBuild) {
                New-CheckResult 4 $check 'FAIL' "$build ($source) cannot use the dedicated hybrid app" 'Update to a supported build: SE, 2019 CU14/CU15 or 2016 CU23 with the April 2025 HU or later.'
            } else {
                New-CheckResult 4 $check 'WARN' "$build ($source) is not a known dedicated-app build line" 'AdminDisplayVersion does not show hotfix updates - check ExSetup.exe on that server: (Get-Command ExSetup.exe).FileVersionInfo.ProductVersion'
            }
        } elseif ($build -ge $min) {
            New-CheckResult 4 $check 'OK' "$build ($source) - dedicated app supported, $graph"
        } elseif ($s.FileBuild) {
            New-CheckResult 4 $check 'FAIL' "$build ($source) is below $min - the shared service principal is blocked since 31 Oct 2025" "Install the hotfix update that brings the server to $min or later."
        } else {
            New-CheckResult 4 $check 'WARN' "$build ($source) is below $min, but AdminDisplayVersion does not show hotfix updates" 'Check ExSetup.exe on that server: (Get-Command ExSetup.exe).FileVersionInfo.ProductVersion'
        }
    }
}

function Get-TenantRoutingDomain {
    # *.mail.onmicrosoft.com from the connector domains or a remote routing address.
    param([string[]]$Candidates)
    @($Candidates | ForEach-Object { "$_".Trim().TrimStart('*').TrimStart('.') } |
      Where-Object { $_ -match '\.mail\.onmicrosoft\.com$' } | Select-Object -Unique)
}

function Test-RemoteRouting {
    # Step 2. $Recipient: Get-RemoteMailbox output (RemoteRoutingAddress) or $null.
    param($Recipient, [string]$Address)
    if (-not $Address) { return New-CheckResult 2 'Target is a remote mailbox' 'SKIP' 'Pass -ExoMailbox to check the target object.' }
    if (-not $Recipient) {
        return New-CheckResult 2 'Target is a remote mailbox' 'FAIL' "$Address is not a remote mailbox on-premises" 'Free/busy for a cloud mailbox needs an on-premises remote mailbox object (Enable-RemoteMailbox / directory sync).'
    }
    $rra = "$($Recipient.RemoteRoutingAddress)" -replace '^(?i)smtp:', ''
    if ($rra -match '@[^@]+\.mail\.onmicrosoft\.com$') {
        New-CheckResult 2 'Target is a remote mailbox' 'OK' "$Address routes to $rra"
    } else {
        New-CheckResult 2 'Target is a remote mailbox' 'FAIL' "$Address routes to '$rra', not to <tenant>.mail.onmicrosoft.com" "Set-RemoteMailbox '$Address' -RemoteRoutingAddress <alias>@<tenant>.mail.onmicrosoft.com"
    }
}

function Test-IntraOrgConnector {
    # Step 3. $Connectors: objects with Name, Enabled, TargetAddressDomains, DiscoveryEndpoint.
    # $RoutingDomain: the tenant routing domain of the target, if known.
    param([object[]]$Connectors, [string]$RoutingDomain)
    $check = 'IntraOrganizationConnector'
    $all = @($Connectors | Where-Object { $_ })
    if (-not $all) {
        return New-CheckResult 3 $check 'FAIL' 'No IntraOrganizationConnector - lookups fall back to the organization relationship (DAuth), which no longer reaches Exchange Online' 'Run the Hybrid Configuration Wizard with "OAuth, Intra Organization Connector and Organization Relationship".'
    }
    $enabled = @($all | Where-Object { "$($_.Enabled)" -eq 'True' })
    if (-not $enabled) {
        return New-CheckResult 3 $check 'FAIL' "Connector(s) disabled: $(@($all.Name) -join ', ')" "Set-IntraOrganizationConnector '$($all[0].Name)' -Enabled `$true"
    }
    $want = if ($RoutingDomain) { $RoutingDomain } else { $null }
    $hit = @($enabled | Where-Object {
        $doms = Get-TenantRoutingDomain @($_.TargetAddressDomains | ForEach-Object { "$_" })
        if ($want) { $doms -contains $want } else { $doms.Count -gt 0 }
    })
    if ($hit) {
        $c = $hit[0]
        New-CheckResult 3 $check 'OK' ("'{0}' enabled for {1}, discovery {2}" -f $c.Name, ((@($c.TargetAddressDomains) | ForEach-Object { "$_" }) -join ', '), $c.DiscoveryEndpoint)
    } else {
        $missing = if ($want) { $want } else { '<tenant>.mail.onmicrosoft.com' }
        New-CheckResult 3 $check 'FAIL' "No enabled connector lists $missing" "Set-IntraOrganizationConnector '$($enabled[0].Name)' -TargetAddressDomains @{Add='$missing'}"
    }
}

function Test-DedicatedAppOverride {
    # Step 5. $Overrides: Get-SettingOverride output. Matched by section, never by name:
    # the name is free text and differs between HCW, script and hand-made overrides.
    param([object[]]$Overrides, [datetime]$Now = (Get-Date))
    $check = 'Dedicated hybrid app enabled'
    $ov = @($Overrides | Where-Object { "$($_.SectionName)" -eq 'ExchangeOnpremAsThirdPartyAppId' })
    $enable = 'New-SettingOverride -Name "EnableExchangeHybrid3PAppFeature" -Component Global -Section ExchangeOnpremAsThirdPartyAppId -Parameters @("Enabled=true") -Reason "Enable dedicated Exchange hybrid app"; Get-ExchangeDiagnosticInfo -Process Microsoft.Exchange.Directory.TopologyService -Component VariantConfiguration -Argument Refresh'
    if (-not $ov) {
        return New-CheckResult 5 $check 'FAIL' 'No setting override for section ExchangeOnpremAsThirdPartyAppId - Exchange still uses the shared service principal (blocked since 31 Oct 2025). HCW creates the app but does not enable it.' $enable
    }
    $on = @($ov | Where-Object { @($_.Parameters | ForEach-Object { "$_" }) -match '^\s*Enabled\s*=\s*true\s*$' })
    if (-not $on) {
        return New-CheckResult 5 $check 'FAIL' "Override '$($ov[0].Name)' exists but is not Enabled=true" "Set-SettingOverride '$($ov[0].Name)' -Parameters @('Enabled=true'), then refresh VariantConfiguration."
    }
    $o = $on[0]
    $scope = @($o.Server | Where-Object { "$_" })
    if ($scope) {
        return New-CheckResult 5 $check 'WARN' "Override '$($o.Name)' applies only to: $($scope -join ', ')" 'Other servers still use the shared service principal - create the override without -Server.'
    }
    if ($o.WhenChanged -and ($Now - [datetime]$o.WhenChanged).TotalMinutes -lt 60) {
        $min = [int]($Now - [datetime]$o.WhenChanged).TotalMinutes
        return New-CheckResult 5 $check 'WARN' "Override '$($o.Name)' enabled, but changed only $min minute(s) ago" 'Exchange can take up to 60 minutes to pick it up. Refresh VariantConfiguration, recycle MSExchangeServicesAppPool, test again later.'
    }
    New-CheckResult 5 $check 'OK' "Override '$($o.Name)' enabled organization-wide"
}

function Test-AuthServerAppId {
    # Step 6. $AuthServers: Get-AuthServer output. Returns the result rows; the app IDs found
    # are in the Finding text and also returned by Get-DedicatedAppId.
    param([object[]]$AuthServers)
    $check = 'EvoSts auth server app ID'
    $evo = @($AuthServers | Where-Object { "$($_.Name)" -like '*EvoSts*' })
    if (-not $evo) {
        return New-CheckResult 6 $check 'FAIL' 'No EvoSts auth server - hybrid OAuth is not configured' 'Run the Hybrid Configuration Wizard (OAuth), then ConfigureExchangeHybridApplication.ps1.'
    }
    foreach ($a in $evo) {
        $id = "$($a.ApplicationIdentifier)".Trim()
        if ("$($a.Enabled)" -ne 'True') {
            New-CheckResult 6 $check 'FAIL' "'$($a.Name)' is disabled" "Set-AuthServer '$($a.Name)' -Enabled `$true"
        } elseif (-not $id) {
            New-CheckResult 6 $check 'FAIL' "'$($a.Name)' has no ApplicationIdentifier - Exchange requests tokens as the shared first-party app" '.\ConfigureExchangeHybridApplication.ps1 -FullyConfigureExchangeHybridApplication (or -ConfigureAuthServer -CustomAppId <appId> -TenantId <tenantId>)'
        } elseif ($id -eq $script:FirstPartyAppId) {
            New-CheckResult 6 $check 'FAIL' "'$($a.Name)' points to the shared first-party app $id" '.\ConfigureExchangeHybridApplication.ps1 -FullyConfigureExchangeHybridApplication'
        } else {
            New-CheckResult 6 $check 'OK' "'$($a.Name)' uses app $id"
        }
    }
}

function Get-DedicatedAppId {
    param([object[]]$AuthServers)
    @($AuthServers | Where-Object { "$($_.Name)" -like '*EvoSts*' -and "$($_.Enabled)" -eq 'True' } |
      ForEach-Object { "$($_.ApplicationIdentifier)".Trim() } |
      Where-Object { $_ -and $_ -ne $script:FirstPartyAppId })
}

function Get-OAuthAppId {
    # Microsoft Learn: Detail.FullId contains "L:<appId>-AS:" for the app the token was issued to.
    param($Result)
    $full = if ($Result.Detail -and $Result.Detail.PSObject.Properties['FullId']) { "$($Result.Detail.FullId)" } else { "$($Result.Detail)" }
    $m = [regex]::Match($full, 'L:(?<guid>[0-9a-fA-F-]{36})-AS:')
    if ($m.Success) { $m.Groups['guid'].Value } else { $null }
}

function Test-OAuthResult {
    # Step 7. $Result: Test-OAuthConnectivity output (ResultType, Detail). $ExpectedAppIds from step 6.
    param($Result, [string[]]$ExpectedAppIds, [string]$Mailbox, [string]$ErrorText)
    $check = 'Token for the dedicated app'
    if (-not $Mailbox) { return New-CheckResult 7 $check 'SKIP' 'Pass -OnPremMailbox to run Test-OAuthConnectivity.' }
    if ($ErrorText) { return New-CheckResult 7 $check 'FAIL' "Test-OAuthConnectivity failed: $ErrorText" 'Run it by hand and read $r.Detail.FullId.' }
    $type = "$($Result.ResultType)"
    $app  = Get-OAuthAppId $Result
    $certHint = 'Check: auth certificate uploaded to the app (ConfigureExchangeHybridApplication.ps1 -UpdateCertificate), API permissions consented, outbound login.microsoftonline.com, Entra service principal sign-ins / Conditional Access.'
    if ($type -ne 'Success') {
        $detail = ("$($Result.Detail)" -replace '\s+', ' ').Trim()
        if ($detail.Length -gt 300) { $detail = $detail.Substring(0, 300) + '...' }
        return New-CheckResult 7 $check 'FAIL' "ResultType $type. $detail" $certHint
    }
    if (-not $app) {
        return New-CheckResult 7 $check 'WARN' 'Success, but no app ID found in Detail.FullId' 'Read $r.Detail.FullId by hand.'
    }
    if ($app -eq $script:FirstPartyAppId) {
        return New-CheckResult 7 $check 'FAIL' "Token issued to the shared first-party app $app" 'Enable the dedicated app (step 5) and set the auth server app ID (step 6).'
    }
    if ($ExpectedAppIds -and ($ExpectedAppIds -notcontains $app)) {
        return New-CheckResult 7 $check 'FAIL' "Token issued to $app, but the EvoSts auth server names $($ExpectedAppIds -join ', ')" 'Check which app is configured; rerun ConfigureExchangeHybridApplication.ps1.'
    }
    New-CheckResult 7 $check 'OK' "Success, token issued to app $app (tested from this server)"
}

function Test-AuthCertificate {
    # Step 7. $AuthConfig: Get-AuthConfig output. $Certificates: certificates of this server
    # (Thumbprint, NotAfter).
    param($AuthConfig, [object[]]$Certificates, [datetime]$Now = (Get-Date))
    $check = 'Auth certificate'
    $tp = "$($AuthConfig.CurrentCertificateThumbprint)".Trim()
    if (-not $tp) { return New-CheckResult 7 $check 'FAIL' 'Get-AuthConfig has no CurrentCertificateThumbprint' 'Create and publish a new auth certificate (Microsoft Learn: Maintain the Exchange Server OAuth certificate).' }
    $c = @($Certificates | Where-Object { "$($_.Thumbprint)" -eq $tp })
    if (-not $c) { return New-CheckResult 7 $check 'WARN' "Auth certificate $tp is not in this server's store" 'Every server needs the auth certificate - check Get-ExchangeCertificate -Thumbprint on this server.' }
    $days = [int]([datetime]$c[0].NotAfter - $Now).TotalDays
    if ($days -lt 0) { return New-CheckResult 7 $check 'FAIL' "Auth certificate $tp expired $(-$days) day(s) ago" 'Renew the auth certificate, then ConfigureExchangeHybridApplication.ps1 -UpdateCertificate.' }
    $next = "$($AuthConfig.NextCertificateThumbprint)".Trim()
    if ($days -lt 30) {
        $extra = if ($next) { " Next certificate $next is staged - upload it to the app as well." } else { '' }
        return New-CheckResult 7 $check 'WARN' "Auth certificate $tp expires in $days day(s).$extra" 'Renew it and run ConfigureExchangeHybridApplication.ps1 -UpdateCertificate.'
    }
    New-CheckResult 7 $check 'OK' "Auth certificate $tp valid for $days more day(s)"
}

function Test-GraphRoute {
    # Step 8. Which route does the server use to Exchange Online? The Graph override is matched
    # by name or section, because Microsoft documents only its name.
    param([object[]]$Overrides, [object[]]$Servers)
    $check = 'Route to Exchange Online'
    $g = @($Overrides | Where-Object { "$($_.Name)" -match 'RouteThroughMSGraph' -or "$($_.SectionName)" -match 'RouteThroughMSGraph' })
    $on = @($g | Where-Object { @($_.Parameters | ForEach-Object { "$_" }) -match '^\s*Enabled\s*=\s*true\s*$' })
    # Below the Graph build for sure (ExSetup.exe on this server) or only by AdminDisplayVersion,
    # which does not include hotfix updates - the second is only a warning.
    $lowSure   = @($Servers | Where-Object { $_.FileBuild -and [version]$_.FileBuild -lt $script:GraphMinBuild })
    $lowUnsure = @($Servers | Where-Object { -not $_.FileBuild -and (-not $_.DisplayBuild -or [version]$_.DisplayBuild -lt $script:GraphMinBuild) })
    if ($on) {
        if ($lowSure) {
            return New-CheckResult 8 $check 'FAIL' "Graph route enabled, but below 15.2.2562.41: $(@($lowSure.Name) -join ', ')" 'Install the May 2026 HU (or later) on every server, or remove the override.'
        }
        if ($lowUnsure) {
            return New-CheckResult 8 $check 'WARN' "Graph route enabled; AdminDisplayVersion shows below 15.2.2562.41 for: $(@($lowUnsure.Name) -join ', ')" 'AdminDisplayVersion does not show hotfix updates - check ExSetup.exe on those servers, or run this script there.'
        }
        return New-CheckResult 8 $check 'OK' "Graph route enabled ('$($on[0].Name)') - lookups go to graph.microsoft.com"
    }
    $how = if ($g) { "Graph override '$($g[0].Name)' is not Enabled=true" } else { 'No Graph override' }
    $fix = if ($lowSure) { 'Graph needs Exchange SE with the May 2026 HU (15.2.2562.41) on all servers.' } else { 'To switch to Graph: .\ConfigureExchangeHybridApplication.ps1 -FullyConfigureExchangeHybridApplication (adds Graph permissions and the override).' }
    New-CheckResult 8 $check 'INFO' "$how - lookups use EWS in Exchange Online (outlook.office365.com)" $fix
}

function Test-EndpointResult {
    # Steps 7/8. $Results: objects with Host, Ok (bool), Error. $Proxy: InternetWebProxy, if set -
    # then a failed direct test is expected and only a warning.
    param([object[]]$Results, [string]$Proxy)
    foreach ($r in $Results) {
        $step = if ($r.Host -eq 'login.microsoftonline.com') { 7 } else { 8 }
        $check = "TCP 443 $($r.Host)"
        if ($r.Ok) { New-CheckResult $step $check 'OK' 'reachable directly' }
        elseif ($Proxy) { New-CheckResult $step $check 'WARN' "not reachable directly ($($r.Error)); InternetWebProxy is $Proxy" 'Direct test only - verify the proxy allows this host and does not inspect TLS.' }
        else { New-CheckResult $step $check 'FAIL' "not reachable ($($r.Error))" 'Allow outbound 443 to this host (firewall, proxy, TLS inspection).' }
    }
}

function Test-EwsLogHit {
    # Step 8. $Hits: Select-String results (Line, Filename, LineNumber). $FileCount: files searched.
    param([object[]]$Hits, [int]$FileCount, [string]$Address)
    $check = 'EWS log errors (this server)'
    if ($FileCount -eq 0) { return New-CheckResult 8 $check 'SKIP' 'No EWS log files found on this server.' }
    $h = @($Hits | Where-Object { $_ })
    if ($Address) { $h = @($h | Where-Object { "$($_.Line)" -match [regex]::Escape($Address) }) }
    $scope = if ($Address) { " for $Address" } else { '' }
    if (-not $h) { return New-CheckResult 8 $check 'OK' "No EXNL/ErrorProxy/Forbidden entries$scope in the last $FileCount file(s)" }
    $graph = @($h | Where-Object { "$($_.Line)" -match 'GraphWebRequest' }).Count
    $last = $h[-1]
    $line = ("$($last.Line)" -replace '\s+', ' ')
    if ($line.Length -gt 250) { $line = $line.Substring(0, 250) + '...' }
    $fix = if ($graph) { 'Graph request errors: check Graph API permissions/consent on the app and outbound graph.microsoft.com; delegate scenario see KB 5127092.' } else { 'Read the full lines in the EWS log; the inner error names the failing hop.' }
    New-CheckResult 8 $check 'WARN' ("{0} hit(s){1}, {2} Graph-related. Last: {3}:{4} {5}" -f $h.Count, $scope, $graph, $last.Filename, $last.LineNumber, $line) $fix
}

function Get-OverallExitCode {
    param([object[]]$Results)
    if (@($Results | Where-Object Status -eq 'FAIL').Count) { 1 } else { 0 }
}

# ---------------------------------------------------------------------------------------------
# Collection: Exchange calls. Each block is guarded - one failing cmdlet must not hide the rest.
# ---------------------------------------------------------------------------------------------

function Invoke-Guarded {
    param([scriptblock]$Block)
    try { , @(& $Block) } catch { $script:LastGuardError = $_.Exception.Message; $null }
}

function Get-LocalExSetupBuild {
    $exe = if ($env:ExchangeInstallPath) { Join-Path $env:ExchangeInstallPath 'Bin\ExSetup.exe' } else { $null }
    if ($exe -and (Test-Path $exe)) {
        try { return ConvertTo-ExchangeBuild (Get-Item $exe).VersionInfo.ProductVersion } catch { }
    }
    $null
}

function Test-Tcp443 {
    param([string]$HostName, [int]$TimeoutMs = 4000)
    $c = New-Object System.Net.Sockets.TcpClient
    try {
        $iar = $c.BeginConnect($HostName, 443, $null, $null)
        if (-not $iar.AsyncWaitHandle.WaitOne($TimeoutMs)) { return [pscustomobject]@{ Host = $HostName; Ok = $false; Error = 'timeout' } }
        $c.EndConnect($iar)
        [pscustomobject]@{ Host = $HostName; Ok = $true; Error = '' }
    } catch {
        $msg = if ($_.Exception.InnerException) { $_.Exception.InnerException.Message } else { $_.Exception.Message }
        [pscustomobject]@{ Host = $HostName; Ok = $false; Error = $msg }
    } finally { $c.Close() }
}

function Invoke-HybridFreeBusyCheck {
    $results = New-Object System.Collections.Generic.List[object]
    $add = { param($r) foreach ($x in @($r)) { if ($x) { $results.Add($x) } } }
    $local = $env:COMPUTERNAME

    Write-Host "Hybrid free/busy check $script:Version on $local" -ForegroundColor Cyan

    # Step 4 data: all Mailbox servers.
    $servers = @()
    $exs = Invoke-Guarded { Get-ExchangeServer | Where-Object { "$($_.ServerRole)" -match 'Mailbox' } }
    if ($null -eq $exs) { & $add (New-CheckResult 4 'Server builds' 'FAIL' "Get-ExchangeServer failed: $script:LastGuardError") }
    else {
        $localBuild = Get-LocalExSetupBuild
        $servers = @($exs | ForEach-Object {
            [pscustomobject]@{
                Name         = "$($_.Name)"
                DisplayBuild = ConvertTo-ExchangeBuild $_.AdminDisplayVersion
                FileBuild    = if ("$($_.Name)" -eq $local) { $localBuild } else { $null }
                Proxy        = "$($_.InternetWebProxy)"
            } })
    }

    # Step 2
    $routing = $null
    if ($ExoMailbox) {
        if (Get-Command Get-RemoteMailbox -ErrorAction SilentlyContinue) {
            $rm = Invoke-Guarded { Get-RemoteMailbox -Identity $ExoMailbox -ErrorAction Stop }
            $rmObj = if ($rm) { $rm[0] } else { $null }
            & $add (Test-RemoteRouting -Recipient $rmObj -Address $ExoMailbox)
            if ($rmObj) { $routing = @(Get-TenantRoutingDomain @(("$($rmObj.RemoteRoutingAddress)" -replace '^(?i)smtp:', '' -split '@')[-1]))[0] }
        } else {
            & $add (New-CheckResult 2 'Target is a remote mailbox' 'SKIP' 'Get-RemoteMailbox is not available (RBAC role).')
        }
    } else { & $add (Test-RemoteRouting -Recipient $null -Address '') }

    # Step 3
    $ioc = Invoke-Guarded { Get-IntraOrganizationConnector }
    if ($null -eq $ioc) { & $add (New-CheckResult 3 'IntraOrganizationConnector' 'FAIL' "Get-IntraOrganizationConnector failed: $script:LastGuardError") }
    else { & $add (Test-IntraOrgConnector -Connectors $ioc -RoutingDomain $routing) }

    # Step 4
    if ($servers) { & $add (Test-ServerBuild -Servers $servers) }

    # Step 5 + 8 data
    $ov = Invoke-Guarded { Get-SettingOverride }
    if ($null -eq $ov) { & $add (New-CheckResult 5 'Dedicated hybrid app enabled' 'FAIL' "Get-SettingOverride failed: $script:LastGuardError") }
    else { & $add (Test-DedicatedAppOverride -Overrides $ov) }

    # Step 6
    $as = Invoke-Guarded { Get-AuthServer }
    $appIds = @()
    if ($null -eq $as) { & $add (New-CheckResult 6 'EvoSts auth server app ID' 'FAIL' "Get-AuthServer failed: $script:LastGuardError") }
    else { & $add (Test-AuthServerAppId -AuthServers $as); $appIds = Get-DedicatedAppId -AuthServers $as }

    # Step 7: token, certificate, login endpoint
    if ($OnPremMailbox) {
        if (Get-Command Test-OAuthConnectivity -ErrorAction SilentlyContinue) {
            Write-Host "Running Test-OAuthConnectivity for $OnPremMailbox ..." -ForegroundColor DarkGray
            $oa = Invoke-Guarded { Test-OAuthConnectivity -Service EWS -TargetUri 'https://outlook.office365.com' -Mailbox $OnPremMailbox -ErrorAction Stop }
            if ($null -eq $oa) { & $add (Test-OAuthResult -Mailbox $OnPremMailbox -ErrorText $script:LastGuardError) }
            else { & $add (Test-OAuthResult -Result $oa[0] -ExpectedAppIds $appIds -Mailbox $OnPremMailbox) }
        } else {
            & $add (New-CheckResult 7 'Token for the dedicated app' 'SKIP' 'Test-OAuthConnectivity is not available (RBAC role).')
        }
    } else { & $add (Test-OAuthResult -Mailbox '') }

    $ac = Invoke-Guarded { Get-AuthConfig }
    if ($null -eq $ac) { & $add (New-CheckResult 7 'Auth certificate' 'FAIL' "Get-AuthConfig failed: $script:LastGuardError") }
    else {
        $certs = @(Get-ChildItem Cert:\LocalMachine\My -ErrorAction SilentlyContinue | Select-Object Thumbprint, NotAfter)
        & $add (Test-AuthCertificate -AuthConfig $ac[0] -Certificates $certs)
    }

    # Step 8: route, endpoints, EWS log
    if ($null -ne $ov) { & $add (Test-GraphRoute -Overrides $ov -Servers $servers) }

    if (-not $SkipNetwork) {
        $proxy = @($servers | Where-Object Name -eq $local | ForEach-Object Proxy | Where-Object { $_ })[0]
        $net = foreach ($h in 'login.microsoftonline.com', 'graph.microsoft.com', 'outlook.office365.com', 'autodiscover-s.outlook.com') { Test-Tcp443 $h }
        & $add (Test-EndpointResult -Results $net -Proxy $proxy)
    }

    $logDir = if ($env:ExchangeInstallPath) { Join-Path $env:ExchangeInstallPath 'Logging\Ews' } else { $null }
    $files = @(if ($logDir -and (Test-Path $logDir)) { Get-ChildItem $logDir -Filter *.log | Sort-Object LastWriteTime | Select-Object -Last $EwsLogFiles })
    $hits = @(if ($files) { $files | Select-String -Pattern 'EXNL|ErrorProxy|Forbidden' })
    & $add (Test-EwsLogHit -Hits $hits -FileCount $files.Count -Address $ExoMailbox)

    # Sort-Object is not stable in Windows PowerShell 5.1: group by step, keep the order inside.
    $sorted = @(foreach ($st in @($results | ForEach-Object Step | Sort-Object -Unique)) { $results | Where-Object Step -eq $st })
    $sorted | Format-Table Step, Status, Check, Finding -AutoSize -Wrap | Out-Host
    $todo = @($sorted | Where-Object { $_.Status -in 'FAIL', 'WARN' -and $_.Fix })
    if ($todo) {
        Write-Host 'What to do:' -ForegroundColor Yellow
        foreach ($t in $todo) {
            $col = if ($t.Status -eq 'FAIL') { 'Red' } else { 'Yellow' }
            Write-Host ("  [{0}] {1}: {2}" -f $t.Step, $t.Check, $t.Fix) -ForegroundColor $col
        }
    }
    Write-Host 'Not checked here: steps 1 and 9 (client view) - use ExchangeTester, Free/Busy, On-Prem -> EXO.' -ForegroundColor DarkGray

    if ($OutputPath) {
        $dir = Split-Path $OutputPath -Parent
        if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $sorted | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "CSV: $OutputPath" -ForegroundColor DarkGray
    }
    $sorted
}

$all = Invoke-HybridFreeBusyCheck
if ($PassThru) { $all }
exit (Get-OverallExitCode $all)
