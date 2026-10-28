#Requires -Version 5.1
<#
.SYNOPSIS
    Read-only check: do the Exchange attributes in on-premises AD match Exchange Online?
    Explains every finding and ends with one verdict: GREEN, YELLOW, RED or INCOMPLETE.

.DESCRIPTION
    After moving Exchange-attribute management to the cloud (IsExchangeCloudManaged), the Cloud Sync
    job "EXO to AD attribute sync" writes values back to AD. Microsoft warns: if a mailbox is switched
    too early after a move, stale recipient-type values from the cloud can overwrite the correct
    remote-mailbox values in AD.

    For every directory-synced mailbox the script compares:
      - msExchRecipientTypeDetails in AD with the expected value for the mailbox type in Exchange Online
      - mail in AD with the primary address in Exchange Online
      - the primary SMTP address in proxyAddresses with the primary address in Exchange Online

    Mailboxes are matched to AD objects in this order (column "MatchedBy"):
      1. Exchange GUID     (EXO: ExchangeGuid, AD: msExchMailboxGuid). Often missing in AD for mailboxes
                           that were created in the cloud with New-RemoteMailbox.
      2. Entra object ID   (EXO: ExternalDirectoryObjectId, AD: msDS-ExternalDirectoryObjectId, written
                           by Entra Connect's Exchange hybrid writeback as "User_<id>").
      3. UPN               (userPrincipalName), as a last resort.

    Weighting: addresses decide the verdict, because local applications read them. The recipient type
    in AD only matters while an on-premises Exchange server still runs or if management is switched
    back. Exception: a cloud-managed mailbox with a wrong type AFTER writeback is exactly Microsoft's
    warning case and turns the verdict RED.

    The script changes nothing, neither in AD nor in Exchange Online.

.PARAMETER BeforeWriteback
    Run BEFORE setting up the writeback. Findings are worded as a preview ("the writeback would change
    this in AD") instead of a result. Limit: the writeback reads its source values from Entra ID; a
    stale recipient type there is not visible from outside. As a last check before "Start provisioning",
    run "Provision on demand" for one or two test users (this really writes for those users).

.PARAMETER AdServer
    Domains or DCs to search. Default: all domains of the current forest.

.PARAMETER ExtraServer
    Additional domains/DCs, e.g. from another forest. Queried with -Credential.

.PARAMETER Credential
    Credentials for -ExtraServer.

.PARAMETER OutFile
    Result as CSV. Default: .\Writeback-Check_<timestamp>.csv

.PARAMETER UserPrincipalName
    Optional: admin account for the Exchange Online sign-in.

.EXAMPLE
    .\Test-ExchangeAttributeWriteback.ps1 -BeforeWriteback
    .\Test-ExchangeAttributeWriteback.ps1
    .\Test-ExchangeAttributeWriteback.ps1 -ExtraServer dc01.fabrikam.com -Credential (Get-Credential)

.NOTES
    Windows PowerShell 5.1 or PowerShell 7, ExchangeOnlineManagement >= 3.6.0, RSAT ActiveDirectory module.
    Exit code: 0 = green, 1 = yellow, 2 = red, 3 = incomplete or not evaluable.
#>
[CmdletBinding()]
param(
    [switch]       $BeforeWriteback,
    [string[]]     $AdServer,
    [string[]]     $ExtraServer,
    [pscredential] $Credential,
    [string]       $OutFile = ".\Writeback-Check_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
    [string]       $UserPrincipalName
)

$script:Version = '1.0.0'
$ErrorActionPreference = 'Stop'

# Prerequisites. The block between the markers comes from build/Prolog.Exo.ps1
# and is maintained by build/Sync-Prolog.ps1 - do not edit by hand.
$script:RequiredModuleVersion = '3.6.0'
$script:RequiredCmdlets       = @('Get-Mailbox')
$script:ToolIsGui             = $false
# <prolog:exo v1 - Quelle: build/Prolog.Exo.ps1, eingefügt von build/Sync-Prolog.ps1.
#                  NICHT von Hand ändern - build/Test-Prolog.ps1 meldet jede Abweichung.>
# Erwartet davor gesetzt:
#   $script:RequiredModuleVersion = '3.6.0'            (Mindestversion ExchangeOnlineManagement)
#   $script:RequiredCmdlets       = @('Get-MailContact')  (Cmdlets, die das Werkzeug wirklich braucht)
#   $script:ToolIsGui             = $true|$false       (bei $true kommen Fehler als MessageBox)
#   $UserPrincipalName                                  (optional, aus dem param-Block)
function Initialize-ExoPrerequisite {
    [CmdletBinding()]
    param(
        [string]$MinimumVersion = $script:RequiredModuleVersion,
        [string[]]$Cmdlets      = $script:RequiredCmdlets,
        [string]$Upn            = $UserPrincipalName,
        [bool]$Gui              = [bool]$script:ToolIsGui
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

    # 1. Windows PowerShell 5.1 oder PowerShell 7+ - dazwischen gibt es nichts Brauchbares.
    if ($PSVersionTable.PSVersion -lt [Version]'5.1') {
        Stop-WithReason "Dieses Werkzeug braucht mindestens Windows PowerShell 5.1, gefunden: $($PSVersionTable.PSVersion)."
    }

    # 2. PowerShell 5.1 spricht ohne Zutun noch TLS 1.0 - die PowerShell Gallery und
    #    Exchange Online nehmen das nicht mehr an.
    if ($PSVersionTable.PSEdition -ne 'Core') {
        [Net.ServicePointManager]::SecurityProtocol =
            [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    }

    # 3. Modul vorhanden und neu genug?
    $min = [Version]$MinimumVersion
    $mod = Get-Module -ListAvailable ExchangeOnlineManagement |
           Sort-Object Version -Descending | Select-Object -First 1
    if (-not $mod) {
        Stop-WithReason ("Das Modul ExchangeOnlineManagement fehlt (mindestens $MinimumVersion nötig).`n`n" +
                         "Installieren mit:`n    Install-Module ExchangeOnlineManagement -Scope CurrentUser -MinimumVersion $MinimumVersion")
    }
    if ($mod.Version -lt $min) {
        Stop-WithReason ("ExchangeOnlineManagement $($mod.Version) ist zu alt, mindestens $MinimumVersion nötig.`n`n" +
                         "Aktualisieren mit:`n    Update-Module ExchangeOnlineManagement`n" +
                         "oder:`n    Install-Module ExchangeOnlineManagement -Scope CurrentUser -MinimumVersion $MinimumVersion -Force")
    }
    # Gezielt die passende Version laden - sonst greift bei mehreren installierten
    # Fassungen die zuerst gefundene, die auch die alte sein kann.
    if (-not (Get-Module ExchangeOnlineManagement)) {
        Import-Module ExchangeOnlineManagement -MinimumVersion $MinimumVersion -ErrorAction Stop
    }

    # 4. Besteht schon eine brauchbare Verbindung? Dann nicht erneut anmelden.
    $live = Get-ConnectionInformation -ErrorAction SilentlyContinue |
            Where-Object { $_.State -eq 'Connected' -and $_.TokenStatus -eq 'Active' }
    if (-not $live) {
        $p = @{ ShowBanner = $false; ErrorAction = 'Stop' }
        if ($Upn) { $p.UserPrincipalName = $Upn }
        Connect-ExchangeOnline @p
        $live = Get-ConnectionInformation -ErrorAction SilentlyContinue |
                Where-Object { $_.State -eq 'Connected' -and $_.TokenStatus -eq 'Active' }
        if (-not $live) { Stop-WithReason 'Die Anmeldung an Exchange Online ist nicht zustande gekommen.' }
    }

    # 5. Modul geladen heißt noch nicht berechtigt: in Exchange Online bringt RBAC nur die
    #    Cmdlets mit, für die die Rolle reicht. Fehlt eines, ist die Rolle das Problem -
    #    das soll hier stehen und nicht später als "Begriff wird nicht erkannt".
    $fehlt = @($Cmdlets | Where-Object { -not (Get-Command $_ -ErrorAction SilentlyContinue) })
    if ($fehlt) {
        Stop-WithReason ("Angemeldet als $($live.UserPrincipalName) an $($live.Organization), aber diese Cmdlets fehlen:`n" +
                         ("    " + ($fehlt -join "`n    ")) +
                         "`n`nDas ist eine Frage der RBAC-Rolle, nicht der Installation.")
    }

    $live | Select-Object -First 1
}
$script:ExoConnection = Initialize-ExoPrerequisite
# </prolog:exo>

# --- Lookup tables --------------------------------------------------------------------------

# Expected msExchRecipientTypeDetails in AD for a mailbox that lives in Exchange Online.
function Get-ExpectedAdType {
    param([string]$ExoType)
    switch ($ExoType) {
        'UserMailbox'      { [int64]2147483648 }    # RemoteUserMailbox
        'SharedMailbox'    { [int64]34359738368 }   # RemoteSharedMailbox
        'RoomMailbox'      { [int64]8589934592 }    # RemoteRoomMailbox
        'EquipmentMailbox' { [int64]17179869184 }   # RemoteEquipmentMailbox
        default            { $null }
    }
}

function Get-AdTypeName {
    param($Value)
    if ($null -eq $Value) { return '' }
    $name = switch ([int64]$Value) {
        1           { 'UserMailbox, i.e. mailbox on-premises!' }
        4           { 'SharedMailbox on-premises!' }
        16          { 'RoomMailbox on-premises!' }
        32          { 'EquipmentMailbox on-premises!' }
        128         { 'MailUser' }
        2147483648  { 'RemoteUserMailbox' }
        8589934592  { 'RemoteRoomMailbox' }
        17179869184 { 'RemoteEquipmentMailbox' }
        34359738368 { 'RemoteSharedMailbox' }
        default     { 'unknown' }
    }
    "$Value ($name)"
}

# One entry per finding code: severity and what it means / what to do.
# Severity: OK | INFO (shown, does not affect the verdict) | WARN | FAIL | NOMATCH
function Get-FindingCatalog {
    param([switch]$BeforeWriteback)
    $c = [ordered]@{
        'OK-Cloud' = @{ Severity = 'OK'; Text =
            'Cloud-managed; AD has the right remote type, mail and primary address. ' +
            'The writeback has not written anything wrong to AD.' }
        'OK-OnPrem' = @{ Severity = 'OK'; Text =
            'Still managed on-premises (IsExchangeCloudManaged = False) and the values match. ' +
            'The writeback only affects this mailbox after it is switched.' }
        'TYPE-WRONG-Cloud' = @{ Severity = 'FAIL'; Text =
            'Cloud-managed, but AD does not hold the remote type that matches the mailbox in Exchange Online. ' +
            'This is the case Microsoft warns about: stale cloud values overwrote the remote-mailbox values in AD. ' +
            'Matters while an on-premises Exchange server still runs (it treats the object as local) and if management is ever switched back. ' +
            'To do: find out when the mailbox was moved and when it was switched, correct the AD values, only then switch further mailboxes.' }
        'TYPE-MISMATCH-OnPrem' = @{ Severity = 'INFO'; Text =
            'Still managed on-premises, and the type in AD does not match Exchange Online, e.g. converted to a shared mailbox ' +
            'in the cloud without updating on-premises. Not a writeback problem and irrelevant for mail flow. ' +
            'Clean-up recommendation while Exchange still runs: e.g. Set-RemoteMailbox -Type Shared on-premises.' }
        'TYPE-NO-TARGET' = @{ Severity = 'INFO'; Text =
            'The script has no expected AD value for this Exchange Online mailbox type. Not evaluated.' }
        'MAIL-DIFF' = @{ Severity = 'WARN'; Text =
            'Type is fine, but mail in AD differs from the primary address in Exchange Online (primary proxyAddress matches). ' +
            'For cloud-managed mailboxes this is typical when the writeback job was created before 2026-08-03: ' +
            'Microsoft says the Mail -> mail mapping is missing then ("Restore default mappings" adds it, but resets your own customizations).' }
        'ADDRESS-DIFF' = @{ Severity = 'WARN'; Text =
            'The primary SMTP address in proxyAddresses differs from Exchange Online. Local applications that read AD see the old address. ' +
            'Cloud-managed: check the writeback (provisioning log of the job). On-premises managed: the address was probably changed ' +
            'directly in the cloud, or synchronization is stuck.' }
        'NOT-IN-AD' = @{ Severity = 'NOMATCH'; Text =
            'No AD object with an Exchange recipient type matches by Exchange GUID, Entra object ID or UPN in the searched domains. Not evaluable. ' +
            'Common reasons: the object lives in a domain or forest that was not searched (-ExtraServer); ' +
            'the account was created in AD without Exchange attributes and the mailbox only exists through a license in the cloud.' }
    }
    if ($BeforeWriteback) {
        $c['OK-Cloud'].Text =
            'Cloud-managed; AD, mail and primary address already match Exchange Online. The writeback would change nothing here.'
        $c['TYPE-WRONG-Cloud'].Severity = 'WARN'
        $c['TYPE-WRONG-Cloud'].Text =
            'Cloud-managed, but AD does not yet hold the remote type that matches Exchange Online. Without writeback it stays like this; ' +
            'with writeback it is overwritten, and with which value cannot be predicted reliably from outside. ' +
            'To do: clarify why AD differs and use "Provision on demand" for this mailbox as the first check.'
        $c['MAIL-DIFF'].Text =
            'mail in AD differs from the primary address in Exchange Online. For cloud-managed mailboxes the writeback will set mail ' +
            'in AD to the cloud value (jobs created from 2026-08-03). Check whether that is wanted, e.g. if local applications read mail.'
        $c['ADDRESS-DIFF'].Text =
            'The primary SMTP address in proxyAddresses differs from Exchange Online. For cloud-managed mailboxes the writeback will bring ' +
            'proxyAddresses in AD to the cloud state. Addresses that exist only in AD will be gone afterwards. Check whether that is wanted.'
    }
    $c
}

# --- AD index -------------------------------------------------------------------------------

function New-AdIndex {
    # Builds three lookup tables from AD user objects. Keys: Guid, Entra object ID (lower case, without
    # the "User_" prefix), UPN (lower case).
    param([object[]]$AdUser)
    $ix = @{ Guid = @{}; Ext = @{}; Upn = @{} }
    foreach ($u in $AdUser) {
        if ($u.msExchMailboxGuid) { $ix.Guid[[guid][byte[]]$u.msExchMailboxGuid] = $u }
        $ext = [string]$u.'msDS-ExternalDirectoryObjectId'
        if ($ext) { $ix.Ext[($ext -replace '^User_').ToLower()] = $u }
        if ($u.userPrincipalName) { $ix.Upn[([string]$u.userPrincipalName).ToLower()] = $u }
    }
    $ix
}

function Find-AdUser {
    param($Mailbox, $Index)
    if ($Mailbox.ExchangeGuid) {
        $u = $Index.Guid[[guid]$Mailbox.ExchangeGuid]
        if ($u) { return @{ User = $u; MatchedBy = 'ExchangeGuid' } }
    }
    if ($Mailbox.ExternalDirectoryObjectId) {
        $u = $Index.Ext[([string]$Mailbox.ExternalDirectoryObjectId).ToLower()]
        if ($u) { return @{ User = $u; MatchedBy = 'EntraObjectId' } }
    }
    if ($Mailbox.UserPrincipalName) {
        $u = $Index.Upn[([string]$Mailbox.UserPrincipalName).ToLower()]
        if ($u) { return @{ User = $u; MatchedBy = 'UPN' } }
    }
    @{ User = $null; MatchedBy = '' }
}

# --- Comparison -----------------------------------------------------------------------------

function Get-MailboxFinding {
    param($Mailbox, $Index, $Catalog)
    $cloud = [string]$Mailbox.IsExchangeCloudManaged -eq 'True'
    $hit = Find-AdUser -Mailbox $Mailbox -Index $Index
    $u = $hit.User
    $expected = Get-ExpectedAdType -ExoType ([string]$Mailbox.RecipientTypeDetails)
    $actual = if ($u -and $null -ne $u.msExchRecipientTypeDetails) { [int64]$u.msExchRecipientTypeDetails } else { $null }
    $primary = [string]$Mailbox.PrimarySmtpAddress
    $proxyPrimary = if ($u) { [string](@($u.proxyAddresses) | Where-Object { $_ -cmatch '^SMTP:' } | Select-Object -First 1) -replace '^SMTP:' } else { '' }
    $mail = if ($u) { [string]$u.mail } else { '' }

    # Microsoft's warning case wins over everything else; then addresses (they matter for operation);
    # an on-premises type mismatch is only a clean-up note.
    $typeKnown = $null -ne $expected
    $code = if (-not $u) { 'NOT-IN-AD' }
        elseif ($cloud -and $typeKnown -and $actual -ne $expected) { 'TYPE-WRONG-Cloud' }
        elseif ($proxyPrimary -ne $primary) { 'ADDRESS-DIFF' }
        elseif ($mail -ne $primary) { 'MAIL-DIFF' }
        elseif (-not $typeKnown) { 'TYPE-NO-TARGET' }
        elseif ($actual -ne $expected) { 'TYPE-MISMATCH-OnPrem' }
        elseif ($cloud) { 'OK-Cloud' } else { 'OK-OnPrem' }

    [pscustomobject]@{
        Finding         = $code
        Severity        = $Catalog[$code].Severity
        UPN             = [string]$Mailbox.UserPrincipalName
        MatchedBy       = $hit.MatchedBy
        CloudManaged    = $cloud
        ExoType         = [string]$Mailbox.RecipientTypeDetails
        AdTypeActual    = Get-AdTypeName $actual
        AdTypeExpected  = Get-AdTypeName $expected
        PrimaryExo      = $primary
        MailAd          = $mail
        PrimaryProxyAd  = $proxyPrimary
        Meaning         = $Catalog[$code].Text
    }
}

function Get-OverallVerdict {
    # Returns verdict, exit code and the sentence to print.
    param([object[]]$Finding, [int]$SearchedServers, [switch]$BeforeWriteback)
    $all       = @($Finding)
    $fail      = @($all | Where-Object Severity -eq 'FAIL').Count
    $warn      = @($all | Where-Object Severity -eq 'WARN').Count
    $nomatch   = @($all | Where-Object Severity -eq 'NOMATCH').Count
    $evaluated = @($all | Where-Object { $_.Severity -ne 'NOMATCH' -and $_.Finding -ne 'TYPE-NO-TARGET' }).Count
    $cloudOk   = @($all | Where-Object Finding -eq 'OK-Cloud').Count
    $total     = $all.Count

    if (-not $SearchedServers -or -not $evaluated) {
        return @{ Verdict = 'NOT EVALUABLE'; ExitCode = 3; Text =
            'Not a single mailbox could be compared with AD (AD not readable or no matches). This is NOT green.' }
    }
    if ($fail) {
        return @{ Verdict = 'RED'; ExitCode = 2; Text =
            "$fail cloud-managed mailbox(es) with a wrong recipient type in AD. See TYPE-WRONG-Cloud." }
    }
    if ($warn) {
        $t = "$warn address difference(s), no failure. Each warning above says what it means and what to do."
        if ($nomatch) { $t += " In addition, $nomatch of $total mailboxes could not be matched (NOT-IN-AD)." }
        return @{ Verdict = 'YELLOW'; ExitCode = 1; Text = $t }
    }
    if ($nomatch) {
        return @{ Verdict = 'INCOMPLETE'; ExitCode = 3; Text =
            "Only $evaluated of $total mailboxes evaluated, without failures or warnings. $nomatch could not be matched (NOT-IN-AD). " +
            'This is NOT green: add the missing domains/forests (-ExtraServer) or clarify the reason first.' }
    }
    $t = if ($BeforeWriteback) {
        "All $cloudOk cloud-managed mailboxes already match Exchange Online in AD; addresses match for all $evaluated evaluated mailboxes. The writeback would change nothing that matters. Last check before 'Start provisioning': 'Provision on demand' for one or two test users."
    } else {
        "All $cloudOk cloud-managed mailboxes have the right remote type and matching addresses in AD; addresses match for all $evaluated evaluated mailboxes."
    }
    @{ Verdict = 'GREEN'; ExitCode = 0; Text = $t }
}

# --- Main -----------------------------------------------------------------------------------

if (-not (Get-Command Get-ADUser -ErrorAction SilentlyContinue)) {
    try { Import-Module ActiveDirectory -ErrorAction Stop }
    catch { throw 'The ActiveDirectory module is missing (RSAT: Active Directory Domain Services tools). Install it or run the script on a DC/admin host.' }
}

$catalog = Get-FindingCatalog -BeforeWriteback:$BeforeWriteback

$exo = @(Get-Mailbox -ResultSize Unlimited -Filter "IsDirSynced -eq `$true" |
    Select-Object UserPrincipalName, ExchangeGuid, ExternalDirectoryObjectId, RecipientTypeDetails, IsExchangeCloudManaged, PrimarySmtpAddress)
Write-Host "Exchange Online: $($exo.Count) directory-synced mailboxes read"

if (-not $AdServer) { $AdServer = @((Get-ADForest).Domains) }
$props = 'msExchMailboxGuid', 'msExchRecipientTypeDetails', 'mail', 'proxyAddresses', 'msDS-ExternalDirectoryObjectId', 'userPrincipalName'
$adUsers = New-Object System.Collections.Generic.List[object]
$searched = @()
foreach ($srv in @($AdServer) + @($ExtraServer)) {
    if (-not $srv) { continue }
    $p = @{ Server = $srv; Filter = 'msExchRecipientTypeDetails -like "*"'; Properties = $props }
    if ($ExtraServer -contains $srv -and $Credential) { $p.Credential = $Credential }
    try {
        $found = @(Get-ADUser @p)
        $nGuid = @($found | Where-Object { $_.msExchMailboxGuid }).Count
        $nExt  = @($found | Where-Object { $_.'msDS-ExternalDirectoryObjectId' }).Count
        Write-Host "AD $srv : $($found.Count) users with an Exchange recipient type ($nGuid with Exchange GUID, $nExt with Entra object ID)"
        $found | ForEach-Object { $adUsers.Add($_) }
        $searched += $srv
    } catch {
        Write-Warning "AD $srv not readable: $($_.Exception.Message)"
    }
}
$index = New-AdIndex -AdUser $adUsers.ToArray()

$result = @(foreach ($m in $exo) { Get-MailboxFinding -Mailbox $m -Index $index -Catalog $catalog })
$result | Export-Csv $OutFile -NoTypeInformation -Encoding UTF8

# --- Report ---------------------------------------------------------------------------------
$color = @{ OK = 'Green'; INFO = 'Cyan'; WARN = 'Yellow'; FAIL = 'Red'; NOMATCH = 'Magenta' }
Write-Host ''
Write-Host "Searched AD domains/servers: $($searched -join ', ')"
if ($BeforeWriteback) { Write-Host 'Mode: preview BEFORE setting up the writeback' }
Write-Host ''
foreach ($code in $catalog.Keys) {
    $hits = @($result | Where-Object Finding -eq $code)
    if (-not $hits) { continue }
    $e = $catalog[$code]
    Write-Host ("[{0}] {1}: {2} mailbox(es)" -f $e.Severity, $code, $hits.Count) -ForegroundColor $color[$e.Severity]
    Write-Host ("  Meaning: {0}" -f $e.Text)
    if ($e.Severity -ne 'OK') {
        $hits | Select-Object -First 20 | ForEach-Object {
            Write-Host ("    - {0}  (EXO: {1}, AD actual: {2}, AD expected: {3}, mail: {4}, primary proxy: {5}, matched by: {6})" -f
                $_.UPN, $_.ExoType, $_.AdTypeActual, $_.AdTypeExpected, $_.MailAd, $_.PrimaryProxyAd, $(if ($_.MatchedBy) { $_.MatchedBy } else { '-' }))
        }
        if ($hits.Count -gt 20) { Write-Host "    ... more in $OutFile" }
    }
    Write-Host ''
}

$verdict = Get-OverallVerdict -Finding $result -SearchedServers $searched.Count -BeforeWriteback:$BeforeWriteback
$vColor = @{ GREEN = 'Green'; YELLOW = 'Yellow'; RED = 'Red'; INCOMPLETE = 'Magenta'; 'NOT EVALUABLE' = 'Magenta' }
Write-Host '=== Verdict ==='
Write-Host ("{0}: {1}" -f $verdict.Verdict, $verdict.Text) -ForegroundColor $vColor[$verdict.Verdict]
$cleanup = @($result | Where-Object Finding -eq 'TYPE-MISMATCH-OnPrem').Count
if ($cleanup) {
    Write-Host "Note: $cleanup mailbox(es) with a type mismatch that does not affect the verdict (TYPE-MISMATCH-OnPrem, clean-up recommendation)." -ForegroundColor Cyan
}
Write-Host ('Limit of this check: it shows the current state. Whether the writeback actively writes mail is only proven ' +
            'if the address changed after the switch.')
Write-Host "Details: $OutFile"
exit $verdict.ExitCode
