#Requires -Version 5.1
<#
.SYNOPSIS
    Exchange AutoDiscover Tester
.DESCRIPTION
    Mirrors Outlook's "Test E-Mail AutoConfiguration" function (Ctrl+right-click tray icon).
    Tests AutoDiscover endpoints, shows HTTP sequence, redirects and the resulting XML.
.NOTES
    No installation required.
    Run: powershell.exe -ExecutionPolicy Bypass -File ExchangeTester.ps1
    Requires: Windows PowerShell 5.1, .NET Framework 4.x (standard on Windows 10/11)
    DNS SRV lookup requires Windows 8.1+ (Resolve-DnsName)
#>

# Einzige Stelle fuer die Versionsnummer; Build und Release-Tag lesen sie hier aus.
$script:Version = '1.0.0'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

# TLS 1.2 / 1.1 / 1.0 for maximum compatibility with older Exchange servers
[System.Net.ServicePointManager]::SecurityProtocol =
    [System.Net.SecurityProtocolType]::Tls12 -bor
    [System.Net.SecurityProtocolType]::Tls11 -bor
    [System.Net.SecurityProtocolType]::Tls

#region ======================================================================
#  FORM
#==============================================================================

$form = New-Object System.Windows.Forms.Form
$form.Text            = "Test E-Mail AutoConfiguration"
$form.ClientSize      = New-Object System.Drawing.Size(775, 595)
$form.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox     = $false
$form.MinimizeBox     = $true

#region --- Row 1: E-Mail Address ---
$lblEmail = New-Object System.Windows.Forms.Label
$lblEmail.Text      = "E-Mail Address"
$lblEmail.Location  = New-Object System.Drawing.Point(8, 14)
$lblEmail.Size      = New-Object System.Drawing.Size(100, 20)
$lblEmail.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$form.Controls.Add($lblEmail)

$txtEmail = New-Object System.Windows.Forms.TextBox
$txtEmail.Location = New-Object System.Drawing.Point(112, 11)
$txtEmail.Size     = New-Object System.Drawing.Size(650, 22)
$txtEmail.TabIndex = 0
$form.Controls.Add($txtEmail)
#endregion

#region --- Row 2: Password ---
$lblPass = New-Object System.Windows.Forms.Label
$lblPass.Text      = "Password"
$lblPass.Location  = New-Object System.Drawing.Point(8, 42)
$lblPass.Size      = New-Object System.Drawing.Size(100, 20)
$lblPass.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lblPass.Enabled   = $false
$form.Controls.Add($lblPass)

$txtPass = New-Object System.Windows.Forms.TextBox
$txtPass.Location     = New-Object System.Drawing.Point(112, 39)
$txtPass.Size         = New-Object System.Drawing.Size(280, 22)
$txtPass.PasswordChar = [char]0x25CF   # ●
$txtPass.Enabled      = $false
$txtPass.TabIndex     = 1
$form.Controls.Add($txtPass)
#endregion

#region --- Auth mode (top-level radio buttons) ---
$radModernAuth = New-Object System.Windows.Forms.RadioButton
$radModernAuth.Text     = "Modern Auth (OAuth2)"
$radModernAuth.Location = New-Object System.Drawing.Point(8, 65)
$radModernAuth.Size     = New-Object System.Drawing.Size(185, 20)
$radModernAuth.Checked  = $true
$radModernAuth.TabIndex = 2
$form.Controls.Add($radModernAuth)

$radWIA = New-Object System.Windows.Forms.RadioButton
$radWIA.Text     = "Windows Integrated Auth"
$radWIA.Location = New-Object System.Drawing.Point(390, 65)
$radWIA.Size     = New-Object System.Drawing.Size(205, 20)
$radWIA.Checked  = $false
$radWIA.TabIndex = 3
$form.Controls.Add($radWIA)
#endregion

#region --- Buttons (top-right) ---
$btnTest = New-Object System.Windows.Forms.Button
$btnTest.Text     = "Test"
$btnTest.Location = New-Object System.Drawing.Point(620, 62)
$btnTest.Size     = New-Object System.Drawing.Size(68, 26)
$btnTest.TabIndex = 9
$form.Controls.Add($btnTest)
$form.AcceptButton = $btnTest

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text     = "Cancel"
$btnCancel.Location = New-Object System.Drawing.Point(696, 62)
$btnCancel.Size     = New-Object System.Drawing.Size(68, 26)
$btnCancel.Enabled  = $false
$btnCancel.TabIndex = 10
$form.Controls.Add($btnCancel)

$btnAddTests = New-Object System.Windows.Forms.Button
$btnAddTests.Text     = "Additional Tests"
$btnAddTests.Location = New-Object System.Drawing.Point(620, 92)
$btnAddTests.Size     = New-Object System.Drawing.Size(144, 26)
$btnAddTests.Enabled  = $false
$btnAddTests.TabIndex = 11
$form.Controls.Add($btnAddTests)
#endregion

#region --- Modern Auth sub-options (Panel keeps radDCF/radACF exclusive with each other only) ---
$pnlModernSub = New-Object System.Windows.Forms.Panel
$pnlModernSub.Location = New-Object System.Drawing.Point(22, 89)
$pnlModernSub.Size     = New-Object System.Drawing.Size(358, 42)
$form.Controls.Add($pnlModernSub)

$radDCF = New-Object System.Windows.Forms.RadioButton
$radDCF.Text     = "Public Office app and Device Code Flow"
$radDCF.Location = New-Object System.Drawing.Point(2, 1)
$radDCF.Size     = New-Object System.Drawing.Size(354, 18)
$radDCF.Checked  = $true
$radDCF.TabIndex = 0
$pnlModernSub.Controls.Add($radDCF)

$radACF = New-Object System.Windows.Forms.RadioButton
$radACF.Text     = "Own 'Exchange Tester' app and Auth Code Flow"
$radACF.Location = New-Object System.Drawing.Point(2, 22)
$radACF.Size     = New-Object System.Drawing.Size(354, 18)
$radACF.Checked  = $false
$radACF.TabIndex = 1
$pnlModernSub.Controls.Add($radACF)
#endregion

#region --- WIA sub-option ---
$chkUseCurrentUser = New-Object System.Windows.Forms.CheckBox
$chkUseCurrentUser.Text     = "Use logged-in user"
$chkUseCurrentUser.Location = New-Object System.Drawing.Point(408, 91)
$chkUseCurrentUser.Size     = New-Object System.Drawing.Size(175, 20)
$chkUseCurrentUser.Checked  = $true
$chkUseCurrentUser.Enabled  = $false
$chkUseCurrentUser.TabIndex = 4
$form.Controls.Add($chkUseCurrentUser)
#endregion

#region --- Standalone options ---
$chkIgnoreCert = New-Object System.Windows.Forms.CheckBox
$chkIgnoreCert.Text     = "Ignore certificate errors"
$chkIgnoreCert.Location = New-Object System.Drawing.Point(8, 133)
$chkIgnoreCert.Size     = New-Object System.Drawing.Size(200, 20)
$chkIgnoreCert.Checked  = $false
$chkIgnoreCert.TabIndex = 5
$form.Controls.Add($chkIgnoreCert)

$chkUseSCP = New-Object System.Windows.Forms.CheckBox
$chkUseSCP.Text     = "Use SCP (domain-joined)"
$chkUseSCP.Location = New-Object System.Drawing.Point(390, 133)
$chkUseSCP.Size     = New-Object System.Drawing.Size(200, 20)
$chkUseSCP.Checked  = $false
$chkUseSCP.TabIndex = 6
$form.Controls.Add($chkUseSCP)
#endregion

#region --- ACF fields: Client ID + Tenant ID (shown only when ACF selected) ---
$lblClientId = New-Object System.Windows.Forms.Label
$lblClientId.Text      = "Client ID:"
$lblClientId.Location  = New-Object System.Drawing.Point(22, 135)
$lblClientId.Size      = New-Object System.Drawing.Size(100, 20)
$lblClientId.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lblClientId.Visible   = $false
$form.Controls.Add($lblClientId)

$txtClientId = New-Object System.Windows.Forms.TextBox
$txtClientId.Location  = New-Object System.Drawing.Point(126, 133)
$txtClientId.Size      = New-Object System.Drawing.Size(356, 22)
$txtClientId.Text      = ''
$txtClientId.Visible   = $false
$txtClientId.TabIndex  = 7
$txtClientId.Font      = New-Object System.Drawing.Font("Consolas", 8.5)
$form.Controls.Add($txtClientId)

$btnCreateApp = New-Object System.Windows.Forms.Button
$btnCreateApp.Text     = "Register App"
$btnCreateApp.Location = New-Object System.Drawing.Point(486, 133)
$btnCreateApp.Size     = New-Object System.Drawing.Size(100, 22)
$btnCreateApp.Visible  = $false
$btnCreateApp.TabIndex = 99
$form.Controls.Add($btnCreateApp)

$lblTenantId = New-Object System.Windows.Forms.Label
$lblTenantId.Text      = "Tenant ID:"
$lblTenantId.Location  = New-Object System.Drawing.Point(22, 159)
$lblTenantId.Size      = New-Object System.Drawing.Size(100, 20)
$lblTenantId.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lblTenantId.Visible   = $false
$form.Controls.Add($lblTenantId)

$txtTenantId = New-Object System.Windows.Forms.TextBox
$txtTenantId.Location  = New-Object System.Drawing.Point(126, 157)
$txtTenantId.Size      = New-Object System.Drawing.Size(460, 22)
$txtTenantId.Text      = ''
$txtTenantId.Visible   = $false
$txtTenantId.TabIndex  = 8
$txtTenantId.Font      = New-Object System.Drawing.Font("Consolas", 8.5)
$form.Controls.Add($txtTenantId)

$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.SetToolTip($lblClientId, "Application (client) ID of your 'Exchange Tester' app registration in Azure AD.")
$toolTip.SetToolTip($txtClientId, "Application (client) ID of your 'Exchange Tester' app registration in Azure AD.")
$toolTip.SetToolTip($lblTenantId, "Directory (tenant) ID — GUID or domain, e.g. contoso.onmicrosoft.com.")
$toolTip.SetToolTip($txtTenantId, "Directory (tenant) ID — GUID or domain, e.g. contoso.onmicrosoft.com.")
#endregion

#region --- Separator + Progress bar (y adjusted dynamically by Update-Layout) ---
$pnlSep = New-Object System.Windows.Forms.Panel
$pnlSep.Location  = New-Object System.Drawing.Point(0, 156)
$pnlSep.Size      = New-Object System.Drawing.Size(775, 2)
$pnlSep.BackColor = [System.Drawing.SystemColors]::ControlDark
$form.Controls.Add($pnlSep)

$prgBar = New-Object System.Windows.Forms.ProgressBar
$prgBar.Location = New-Object System.Drawing.Point(8, 164)
$prgBar.Size     = New-Object System.Drawing.Size(757, 14)
$prgBar.Minimum  = 0
$prgBar.Maximum  = 100
$prgBar.Value    = 0
$form.Controls.Add($prgBar)
#endregion

#region --- TabControl ---
$tabCtrl = New-Object System.Windows.Forms.TabControl
$tabCtrl.Location = New-Object System.Drawing.Point(8, 183)
$tabCtrl.Size     = New-Object System.Drawing.Size(757, 403)
$form.Controls.Add($tabCtrl)

# Tab: Results
$tabResults = New-Object System.Windows.Forms.TabPage
$tabResults.Text = "Results"
$tabCtrl.Controls.Add($tabResults)

$lvwResults = New-Object System.Windows.Forms.ListView
$lvwResults.Dock          = [System.Windows.Forms.DockStyle]::Fill
$lvwResults.View          = [System.Windows.Forms.View]::Details
$lvwResults.FullRowSelect = $true
$lvwResults.GridLines     = $true
$lvwResults.ShowGroups    = $true
$lvwResults.HeaderStyle   = [System.Windows.Forms.ColumnHeaderStyle]::Nonclickable
[void]$lvwResults.Columns.Add("Setting", 220)
[void]$lvwResults.Columns.Add("Value",   510)
$tabResults.Controls.Add($lvwResults)

# Tab: Log
$tabLog = New-Object System.Windows.Forms.TabPage
$tabLog.Text = "Log"
$tabCtrl.Controls.Add($tabLog)

$rtbLog = New-Object System.Windows.Forms.RichTextBox
$rtbLog.Dock        = [System.Windows.Forms.DockStyle]::Fill
$rtbLog.ReadOnly    = $true
$rtbLog.Font        = New-Object System.Drawing.Font("Consolas", 9)
$rtbLog.BackColor   = [System.Drawing.Color]::White
$rtbLog.ScrollBars  = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
$rtbLog.WordWrap    = $true
$tabLog.Controls.Add($rtbLog)

# Tab: XML
$tabXml = New-Object System.Windows.Forms.TabPage
$tabXml.Text = "XML"
$tabCtrl.Controls.Add($tabXml)

$rtbXml = New-Object System.Windows.Forms.RichTextBox
$rtbXml.Dock       = [System.Windows.Forms.DockStyle]::Fill
$rtbXml.ReadOnly   = $true
$rtbXml.Font       = New-Object System.Drawing.Font("Consolas", 9)
$rtbXml.BackColor  = [System.Drawing.Color]::White
$rtbXml.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Both
$rtbXml.WordWrap   = $false
$tabXml.Controls.Add($rtbXml)
#endregion

#region --- Context menus for Log and XML ---

# Shared items: Copy selection / Select All
$ctxLog = New-Object System.Windows.Forms.ContextMenuStrip
$miLogCopy = New-Object System.Windows.Forms.ToolStripMenuItem("Copy")
$miLogCopy.ShortcutKeyDisplayString = "Ctrl+C"
$miLogSelAll = New-Object System.Windows.Forms.ToolStripMenuItem("Select All")
$miLogSelAll.ShortcutKeyDisplayString = "Ctrl+A"
$miLogClear = New-Object System.Windows.Forms.ToolStripMenuItem("Clear Log")
[void]$ctxLog.Items.Add($miLogCopy)
[void]$ctxLog.Items.Add($miLogSelAll)
[void]$ctxLog.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
[void]$ctxLog.Items.Add($miLogClear)
$rtbLog.ContextMenuStrip = $ctxLog

$miLogCopy.Add_Click({
    $t = $rtbLog.SelectedText
    if (-not $t) { $t = $rtbLog.Text }
    if ($t) { [System.Windows.Forms.Clipboard]::SetText($t) }
})
$miLogSelAll.Add_Click({ $rtbLog.SelectAll() })
$miLogClear.Add_Click({ $rtbLog.Clear() })

$ctxXml = New-Object System.Windows.Forms.ContextMenuStrip
$miXmlCopy   = New-Object System.Windows.Forms.ToolStripMenuItem("Copy")
$miXmlCopy.ShortcutKeyDisplayString = "Ctrl+C"
$miXmlSelAll = New-Object System.Windows.Forms.ToolStripMenuItem("Select All")
$miXmlSelAll.ShortcutKeyDisplayString = "Ctrl+A"
$miXmlSave   = New-Object System.Windows.Forms.ToolStripMenuItem("Save XML As...")
[void]$ctxXml.Items.Add($miXmlCopy)
[void]$ctxXml.Items.Add($miXmlSelAll)
[void]$ctxXml.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
[void]$ctxXml.Items.Add($miXmlSave)
$rtbXml.ContextMenuStrip = $ctxXml

$miXmlCopy.Add_Click({
    $t = $rtbXml.SelectedText
    if (-not $t) { $t = $rtbXml.Text }
    if ($t) { [System.Windows.Forms.Clipboard]::SetText($t) }
})
$miXmlSelAll.Add_Click({ $rtbXml.SelectAll() })
$miXmlSave.Add_Click({
    if (-not $rtbXml.Text) { return }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter   = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*"
    $sfd.FileName = "autodiscover.xml"
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        [System.IO.File]::WriteAllText($sfd.FileName, $rtbXml.Text, [System.Text.Encoding]::UTF8)
        [System.Windows.Forms.MessageBox]::Show(
            "Saved to:`n$($sfd.FileName)", "Saved",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
})
#endregion

#endregion ===================================================================
#  XML PARSER
#==============================================================================

function ConvertFrom-AutodiscoverXml {
    param([string]$RawXml)

    $rows = [System.Collections.Generic.List[PSCustomObject]]::new()

    try {
        $xd = [xml]$RawXml
        $ns = New-Object System.Xml.XmlNamespaceManager($xd.NameTable)
        $ns.AddNamespace("ad", "http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a")

        # Returns inner text of first matching node, or $null
        $txt = {
            param($node, [string]$xpath)
            $n = $node.SelectSingleNode($xpath, $ns)
            if ($n) { $n.InnerText } else { $null }
        }

        $add = {
            param([string]$group, [string]$setting, $value)
            if ($value -and $value.Trim() -ne '') {
                $rows.Add([PSCustomObject]@{ Group = $group; Setting = $setting; Value = $value.Trim() })
            }
        }

        # --- User ---
        $user = $xd.SelectSingleNode("//ad:User", $ns)
        if ($user) {
            & $add "User" "Display Name"  (& $txt $user "ad:DisplayName")
            & $add "User" "Legacy DN"      (& $txt $user "ad:LegacyDN")
            & $add "User" "Deployment ID"  (& $txt $user "ad:DeploymentId")
        }

        # --- Account ---
        $account = $xd.SelectSingleNode("//ad:Account", $ns)
        if ($account) {
            $action = & $txt $account "ad:Action"
            & $add "Account" "Account Type"     (& $txt $account "ad:AccountType")
            & $add "Account" "Action"             $action
            & $add "Account" "Redirect Address"  (& $txt $account "ad:RedirectAddr")
            & $add "Account" "Redirect URL"      (& $txt $account "ad:RedirectUrl")
        }

        # --- Protocol sections (EXCH, EXPR, EXHTTP, mapiHttp, WEB, ...) ---
        $protocols = $xd.SelectNodes("//ad:Protocol", $ns)
        foreach ($proto in $protocols) {
            $typeNode = $proto.SelectSingleNode("ad:Type", $ns)
            # mapiHttp uses Type as an XML attribute instead of a child element
            $type = if ($typeNode) { $typeNode.InnerText } else { $proto.GetAttribute("Type") }
            if (-not $type) { $type = "Unknown" }
            $grp = "Protocol: $type"

            & $add $grp "Server"                  (& $txt $proto "ad:Server")
            & $add $grp "Port"                    (& $txt $proto "ad:Port")
            & $add $grp "SSL"                     (& $txt $proto "ad:SSL")
            & $add $grp "Encryption"              (& $txt $proto "ad:Encryption")
            & $add $grp "Login Name"              (& $txt $proto "ad:LoginName")
            & $add $grp "Domain Required"         (& $txt $proto "ad:DomainRequired")
            & $add $grp "Auth Required"           (& $txt $proto "ad:AuthRequired")
            & $add $grp "Auth Package"            (& $txt $proto "ad:AuthPackage")
            & $add $grp "EWS URL"                 (& $txt $proto "ad:EwsUrl")
            & $add $grp "OAB URL"                 (& $txt $proto "ad:OABUrl")
            & $add $grp "OOF URL"                 (& $txt $proto "ad:OOFUrl")
            & $add $grp "Availability Service URL" (& $txt $proto "ad:ASUrl")
            & $add $grp "EMWS URL"                (& $txt $proto "ad:EmwsUrl")
            & $add $grp "ECP URL"                 (& $txt $proto "ad:EcpUrl")
            & $add $grp "Sharing URL"             (& $txt $proto "ad:SharingUrl")
            & $add $grp "UM URL"                  (& $txt $proto "ad:UMUrl")
            & $add $grp "Public Folder Server"    (& $txt $proto "ad:PublicFolderServer")
            & $add $grp "Server Exclusive Connect" (& $txt $proto "ad:ServerExclusiveConnect")
            & $add $grp "Autodiscover Internal"   (& $txt $proto "ad:AutodiscoverServiceInternalUri")

            # WEB protocol: OWA URLs are nested under Internal/External
            $inOwa  = $proto.SelectSingleNode("ad:Internal/ad:OWAUrl",  $ns)
            $extOwa = $proto.SelectSingleNode("ad:External/ad:OWAUrl", $ns)
            if ($inOwa)  { & $add $grp "OWA URL (Internal)" $inOwa.InnerText }
            if ($extOwa) { & $add $grp "OWA URL (External)" $extOwa.InnerText }

            # mapiHttp: MailStore and AddressBook have InternalUrl/ExternalUrl children
            foreach ($store in @("MailStore", "AddressBook")) {
                $nInt = $proto.SelectSingleNode("ad:$store/ad:InternalUrl", $ns)
                $nExt = $proto.SelectSingleNode("ad:$store/ad:ExternalUrl", $ns)
                if ($nInt) { & $add $grp "$store Internal URL" $nInt.InnerText }
                if ($nExt) { & $add $grp "$store External URL" $nExt.InnerText }
            }
        }

        # --- Alternative Mailboxes (shared/delegate mailboxes returned by AutoDiscover) ---
        $altBoxes = $xd.SelectNodes("//ad:Account/ad:AlternativeMailbox", $ns)
        $altIdx = 0
        foreach ($alt in $altBoxes) {
            $altIdx++
            $dn = & $txt $alt "ad:DisplayName"
            $gname = if ($dn) { "Alternative Mailbox: $dn" } else { "Alternative Mailbox $altIdx" }
            & $add $gname "Type"              (& $txt $alt "ad:Type")
            & $add $gname "Display Name"      $dn
            & $add $gname "SMTP Address"      (& $txt $alt "ad:SmtpAddress")
            & $add $gname "Legacy DN"         (& $txt $alt "ad:LegacyDN")
            & $add $gname "Server"            (& $txt $alt "ad:Server")
            & $add $gname "Owner SMTP Address" (& $txt $alt "ad:OwnerSmtpAddress")
        }

        # --- Public Folder Information ---
        $pfi = $xd.SelectSingleNode("//ad:Account/ad:PublicFolderInformation", $ns)
        if ($pfi) {
            & $add "Public Folder Information" "SMTP Address" (& $txt $pfi "ad:SmtpAddress")
        }
    } catch {
        $rows.Add([PSCustomObject]@{
            Group   = "Error"
            Setting = "XML parse error"
            Value   = $_.Exception.Message
        })
    }

    return ,$rows   # comma forces List to stay as-is, not unroll
}

# Extract all testable (absolute https?://) URLs from AutoDiscover XML.
# Dedup is per-protocol so the same URL appearing under different protocol
# sections (e.g. EXCH, EXPR and EXHTTP all pointing at ews/exchange.asmx)
# is shown once per section. Within a section, fields sharing a URL are merged.
function Get-AutodiscoverUrls {
    param([string]$RawXml)

    $results  = [System.Collections.Generic.List[PSCustomObject]]::new()
    # key "protocol`n url" -> row object (so we can append field names)
    $rowMap   = @{}
    $allHosts = [System.Collections.Generic.List[string]]::new()   # scheme://host, in order seen
    $hostSeen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $segSeen  = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    $addUrl = {
        param([string]$protoType, [string]$field, [string]$url)
        if (-not $url) { return }
        $url = $url.Trim()
        if ($url -notmatch '^https?://') { return }
        # MAPI/HTTP AutoDiscover URLs carry ?MailboxId=… which causes 500 on a
        # plain GET without MAPI headers; strip the query string for probing
        if ($protoType -eq 'mapiHttp' -and $url -match '\?') {
            $url = ($url -split '\?')[0]
        }
        $key = "$protoType`n$url"
        if ($rowMap.ContainsKey($key)) {
            # Same URL already listed for this protocol — merge the field label
            $existing = $rowMap[$key]
            if (($existing.Field -split ', ') -notcontains $field) {
                $existing.Field = "$($existing.Field), $field"
            }
        } else {
            $row = [PSCustomObject]@{ Protocol = $protoType; Field = $field; Url = $url }
            $rowMap[$key] = $row
            $results.Add($row)
        }
        # Track host + first path segment (vdir) for the healthcheck basis test
        if ($url -match '^(https?://[^/]+)(?:/([^/?#]+))?') {
            if ($hostSeen.Add($Matches[1])) { $allHosts.Add($Matches[1]) }
            if ($Matches[2]) { [void]$segSeen.Add($Matches[2]) }
        }
    }

    try {
        $xd = [xml]$RawXml
        $ns = New-Object System.Xml.XmlNamespaceManager($xd.NameTable)
        $ns.AddNamespace("ad", "http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a")

        $protocols = $xd.SelectNodes("//ad:Protocol", $ns)
        foreach ($proto in $protocols) {
            $typeNode  = $proto.SelectSingleNode("ad:Type", $ns)
            $protoType = if ($typeNode) { $typeNode.InnerText } else { $proto.GetAttribute("Type") }
            if (-not $protoType) { $protoType = "?" }

            foreach ($f in @("EwsUrl","ASUrl","OABUrl","OOFUrl","EcpUrl","SharingUrl","EmwsUrl","UMUrl","EwsPartnerUrl")) {
                $n = $proto.SelectSingleNode("ad:$f", $ns)
                if ($n) { & $addUrl $protoType $f $n.InnerText }
            }

            # WEB: OWA URLs nested under Internal/External
            foreach ($dir in @("Internal","External")) {
                $n = $proto.SelectSingleNode("ad:$dir/ad:OWAUrl", $ns)
                if ($n) { & $addUrl $protoType "OWA ($dir)" $n.InnerText }
            }

            # mapiHttp: MailStore / AddressBook (query string stripped inside $addUrl)
            foreach ($store in @("MailStore","AddressBook")) {
                foreach ($uField in @("InternalUrl","ExternalUrl")) {
                    $n = $proto.SelectSingleNode("ad:$store/ad:$uField", $ns)
                    if ($n) { & $addUrl $protoType "$store.$uField" $n.InnerText }
                }
            }
        }

        # ── Healthcheck basis test ────────────────────────────────────────────
        # Every Exchange IIS virtual directory exposes /<vdir>/healthcheck.htm
        # (the page load balancers probe; returns HTTP 200 when the vdir's app
        # pool is healthy). Build one healthcheck per host for the standard set
        # of vdirs, plus any vdir actually seen in the AutoDiscover URLs.
        $vdirs = [System.Collections.Specialized.OrderedDictionary]::new()  # lowercase -> canonical
        foreach ($v in @('autodiscover','ews','oab','owa','ecp','mapi','rpc','Microsoft-Server-ActiveSync')) {
            $vdirs[$v.ToLower()] = $v
        }
        foreach ($seg in $segSeen) {
            $lc = $seg.ToLower()
            if (-not $vdirs.Contains($lc)) { $vdirs[$lc] = $seg }
        }
        foreach ($h in $allHosts) {
            foreach ($vd in $vdirs.Values) {
                $hcUrl = "$h/$vd/healthcheck.htm"
                $key   = "Health`n$hcUrl"
                if (-not $rowMap.ContainsKey($key)) {
                    $row = [PSCustomObject]@{ Protocol = 'Health'; Field = $vd; Url = $hcUrl }
                    $rowMap[$key] = $row
                    $results.Add($row)
                }
            }
        }
    } catch {}

    return ,$results
}

function Show-EndpointTestDialog {
    param($Title, $Entries, $AuthHeader, $UseWinAuth, $NetCred, $IgnoreCert, $Owner, $CsvName = 'endpoint-tests.csv')

    # ── Dialog ───────────────────────────────────────────────────────────────
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = $Title
    $dlg.ClientSize      = New-Object System.Drawing.Size(952, 490)
    $dlg.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dlg.MinimizeBox     = $false
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable

    $lv = New-Object System.Windows.Forms.ListView
    $lv.Location      = New-Object System.Drawing.Point(0, 0)
    $lv.Size          = New-Object System.Drawing.Size(952, 444)
    $lv.Anchor        = ([System.Windows.Forms.AnchorStyles]::Top    -bor
                          [System.Windows.Forms.AnchorStyles]::Left   -bor
                          [System.Windows.Forms.AnchorStyles]::Right  -bor
                          [System.Windows.Forms.AnchorStyles]::Bottom)
    $lv.View          = [System.Windows.Forms.View]::Details
    $lv.FullRowSelect = $true
    $lv.GridLines     = $true
    $lv.HeaderStyle   = [System.Windows.Forms.ColumnHeaderStyle]::Nonclickable
    [void]$lv.Columns.Add("Protocol",  68)
    [void]$lv.Columns.Add("Field",    130)
    [void]$lv.Columns.Add("Status",    48)
    [void]$lv.Columns.Add("Result",   150)
    [void]$lv.Columns.Add("URL",      370)
    [void]$lv.Columns.Add("Auth / Info", 170)

    # Right-click context menu: copy URL / copy all as CSV
    $ctxLv     = New-Object System.Windows.Forms.ContextMenuStrip
    $miCopyUrl = New-Object System.Windows.Forms.ToolStripMenuItem("Copy URL")
    [void]$ctxLv.Items.Add($miCopyUrl)
    $miCopyUrl.Add_Click({
        if ($lv.SelectedItems.Count -gt 0) {
            [System.Windows.Forms.Clipboard]::SetText($lv.SelectedItems[0].SubItems[4].Text)
        }
    })
    [void]$ctxLv.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    $miCsvCtx = New-Object System.Windows.Forms.ToolStripMenuItem("Copy all as CSV")
    [void]$ctxLv.Items.Add($miCsvCtx)
    $lv.ContextMenuStrip = $ctxLv

    # CSV builder (shared by Copy and Save buttons)
    $makeCsv = {
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine('"Protocol";"Field";"Status";"Result";"URL";"Auth / Info"')
        foreach ($row in $lv.Items) {
            $cols = @(
                $row.Text,
                $row.SubItems[1].Text,
                $row.SubItems[2].Text,
                $row.SubItems[3].Text,
                $row.SubItems[4].Text,
                $row.SubItems[5].Text
            )
            $line = ($cols | ForEach-Object { '"' + ($_ -replace '"','""') + '"' }) -join ';'
            [void]$sb.AppendLine($line)
        }
        return $sb.ToString()
    }

    $lblProg = New-Object System.Windows.Forms.Label
    $lblProg.Text      = "Connecting…"
    $lblProg.Location  = New-Object System.Drawing.Point(8, 457)
    $lblProg.Size      = New-Object System.Drawing.Size(590, 18)
    $lblProg.Anchor    = ([System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom)
    $lblProg.ForeColor = [System.Drawing.Color]::Gray

    $btnSaveCsv = New-Object System.Windows.Forms.Button
    $btnSaveCsv.Text     = "Save CSV…"
    $btnSaveCsv.Location = New-Object System.Drawing.Point(688, 453)
    $btnSaveCsv.Size     = New-Object System.Drawing.Size(80, 26)
    $btnSaveCsv.Anchor   = ([System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom)
    $btnSaveCsv.Add_Click({
        $csv = & $makeCsv
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter   = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        $sfd.FileName = $CsvName
        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            [System.IO.File]::WriteAllText($sfd.FileName, $csv, [System.Text.Encoding]::UTF8)
        }
    })

    $btnCopyCsv = New-Object System.Windows.Forms.Button
    $btnCopyCsv.Text     = "Copy CSV"
    $btnCopyCsv.Location = New-Object System.Drawing.Point(776, 453)
    $btnCopyCsv.Size     = New-Object System.Drawing.Size(80, 26)
    $btnCopyCsv.Anchor   = ([System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom)
    $btnCopyCsv.Add_Click({
        $csv = & $makeCsv
        [System.Windows.Forms.Clipboard]::SetText($csv)
    })

    $miCsvCtx.Add_Click({ $btnCopyCsv.PerformClick() })

    $btnDlgClose = New-Object System.Windows.Forms.Button
    $btnDlgClose.Text     = "Close"
    $btnDlgClose.Location = New-Object System.Drawing.Point(864, 453)
    $btnDlgClose.Size     = New-Object System.Drawing.Size(80, 26)
    $btnDlgClose.Anchor   = ([System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom)
    $btnDlgClose.Add_Click({ $dlg.Close() })
    $dlg.CancelButton = $btnDlgClose

    $dlg.Controls.AddRange(@($lv, $lblProg, $btnSaveCsv, $btnCopyCsv, $btnDlgClose))

    # Pre-populate rows with placeholder status. Key by index, not URL, because
    # the same URL can legitimately appear under several protocol sections.
    $lvItems = @{}
    for ($i = 0; $i -lt $Entries.Count; $i++) {
        $entry = $Entries[$i]
        $entry | Add-Member -NotePropertyName Index -NotePropertyValue $i -Force
        $item = New-Object System.Windows.Forms.ListViewItem($entry.Protocol)
        [void]$item.SubItems.Add($entry.Field)
        [void]$item.SubItems.Add("…")   # Status
        [void]$item.SubItems.Add("")     # Result
        [void]$item.SubItems.Add($entry.Url)
        [void]$item.SubItems.Add("")     # Auth / Info
        $item.ForeColor = [System.Drawing.Color]::Gray
        [void]$lv.Items.Add($item)
        $lvItems[$i] = $item
    }

    # Sync bridge for runspace → UI
    $uSync = [hashtable]::Synchronized(@{
        Queue  = [System.Collections.Concurrent.ConcurrentQueue[hashtable]]::new()
        Done   = $false
        Cancel = $false
        Tested = 0
        Total  = $Entries.Count
    })

    # Background runspace: probe each URL with the selected auth method
    $uRs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $uRs.ApartmentState = [System.Threading.ApartmentState]::STA
    $uRs.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::UseNewThread
    $uRs.Open()
    $uRs.SessionStateProxy.SetVariable('uSync',      $uSync)
    $uRs.SessionStateProxy.SetVariable('urlList',    $Entries)
    $uRs.SessionStateProxy.SetVariable('ignoreCert', $IgnoreCert)
    $uRs.SessionStateProxy.SetVariable('authHeader', $AuthHeader)
    $uRs.SessionStateProxy.SetVariable('useWinAuth', $UseWinAuth)
    $uRs.SessionStateProxy.SetVariable('netCred',    $NetCred)

    $uPs = [System.Management.Automation.PowerShell]::Create()
    $uPs.Runspace = $uRs
    [void]$uPs.AddScript({
        if ($ignoreCert) {
            [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
        }
        [System.Net.ServicePointManager]::SecurityProtocol =
            [System.Net.SecurityProtocolType]::Tls12 -bor
            [System.Net.SecurityProtocolType]::Tls11 -bor
            [System.Net.SecurityProtocolType]::Tls

        foreach ($entry in $urlList) {
            if ($uSync.Cancel) { break }
            $url    = $entry.Url
            $idx    = $entry.Index
            $status = 0
            $info   = ''
            try {
                $req = [System.Net.HttpWebRequest]::Create($url)
                $req.Method            = "GET"
                $req.AllowAutoRedirect = $false
                $req.Timeout           = 10000
                # Use a browser-like User-Agent and Accept header. The MAPI/HTTP
                # endpoints (/mapi/emsmdb, /mapi/nspi) serve their friendly
                # "Connectivity Endpoint" HTML page (HTTP 200) to browsers, but
                # return HTTP 500 to an Office User-Agent because the server then
                # routes the request into the MAPI handler, which expects a POST.
                $req.UserAgent         = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
                $req.Accept            = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
                if ($authHeader) {
                    $req.Headers["Authorization"] = $authHeader
                } elseif ($useWinAuth) {
                    $req.UseDefaultCredentials = $true
                } elseif ($netCred) {
                    $req.Credentials = $netCred
                }
                try {
                    $rp     = $req.GetResponse()
                    $status = [int]$rp.StatusCode
                    $loc    = $rp.Headers["Location"]
                    if ($loc) { $info = "-> $loc" }
                    $rp.Close()
                } catch [System.Net.WebException] {
                    $ex = $_.Exception
                    if ($ex.Response) {
                        $status = [int]$ex.Response.StatusCode
                        $loc    = $ex.Response.Headers["Location"]
                        $wwwA   = try { $ex.Response.Headers.GetValues("WWW-Authenticate") -join " | " } catch { $null }
                        if ($wwwA)    { $info = $wwwA }
                        elseif ($loc) { $info = "-> $loc" }
                        $ex.Response.Close()
                    } else {
                        $status = -1
                        $inner  = $ex.InnerException
                        $info   = if ($inner -and $inner.Message) { $inner.Message } else { $ex.Message }
                    }
                }
            } catch {
                $status = -1
                $info   = $_.Exception.Message
            }
            $uSync.Queue.Enqueue(@{ Index = $idx; Status = $status; Info = $info })
            $uSync.Tested++
        }
        $uSync.Done = $true
    })
    [void]$uPs.BeginInvoke()

    # Timer: drain result queue and update ListView on UI thread
    $uTimer = New-Object System.Windows.Forms.Timer
    $uTimer.Interval = 150
    $uTimer.Add_Tick({
        $upd = $null
        while ($uSync.Queue.TryDequeue([ref]$upd)) {
            if ($lvItems.ContainsKey($upd.Index)) {
                $item = $lvItems[$upd.Index]
                $s    = $upd.Status
                # Verdict: interpret the status as reachability, not pass/fail. A
                # service that answers with an auth challenge or a redirect is
                # reachable and healthy — only real server errors / no-connect
                # are failures. This keeps Exchange Online results (mostly 401 /
                # 302 / 404 by design) readable instead of an alarming wall of red.
                if     ($s -eq 200)                 { $verdict = 'OK';                        $col = [System.Drawing.Color]::DarkGreen }
                elseif ($s -ge 300 -and $s -lt 400) { $verdict = 'Reachable (redirect)';      $col = [System.Drawing.Color]::SeaGreen }
                elseif ($s -eq 401)                 { $verdict = 'Reachable — auth required';  $col = [System.Drawing.Color]::FromArgb(160,100,0) }
                elseif ($s -eq 403)                 { $verdict = 'Reachable — forbidden';      $col = [System.Drawing.Color]::FromArgb(160,100,0) }
                elseif ($s -eq 405)                 { $verdict = 'Reachable — needs POST';     $col = [System.Drawing.Color]::FromArgb(160,100,0) }
                elseif ($s -eq 404)                 { $verdict = 'Not present';                $col = [System.Drawing.Color]::DimGray }
                elseif ($s -ge 400 -and $s -lt 500) { $verdict = "Reachable (HTTP $s)";        $col = [System.Drawing.Color]::FromArgb(160,100,0) }
                elseif ($s -ge 500)                 { $verdict = 'Server error';               $col = [System.Drawing.Color]::DarkRed }
                elseif ($s -lt 0)                   { $verdict = 'Unreachable';                $col = [System.Drawing.Color]::Red }
                else                                { $verdict = "HTTP $s";                    $col = [System.Drawing.Color]::Black }
                $item.SubItems[2].Text = if ($s -lt 0) { "ERR" } else { "$s" }
                $item.SubItems[3].Text = $verdict
                $item.SubItems[5].Text = $upd.Info
                $item.ForeColor = $col
            }
        }
        $lblProg.Text = "Tested $($uSync.Tested) of $($uSync.Total)…"
        if ($uSync.Done) {
            $uTimer.Stop()
            $lblProg.Text      = "Done  —  $($uSync.Total) URL(s) tested."
            $lblProg.ForeColor = [System.Drawing.Color]::Black
            try { $uPs.Dispose() } catch {}
            try { $uRs.Close(); $uRs.Dispose() } catch {}
        }
    })

    $dlg.Add_FormClosed({
        $uSync.Cancel = $true
        $uTimer.Stop()
        try { $uPs.Stop() }    catch {}
        try { $uPs.Dispose() } catch {}
        try { $uRs.Close(); $uRs.Dispose() } catch {}
    })

    $uTimer.Start()
    [void]$dlg.ShowDialog($Owner)
    $uTimer.Stop()
    try { $uTimer.Dispose() } catch {}
}

# Build the endpoint list for the Hybrid Connectivity test: AutoDiscover
# (on-prem + Exchange Online), OAuth/modern-auth metadata, free/busy (EWS),
# and per-vdir healthcheck.htm on the on-prem host.
function Get-HybridConnectivityEndpoints {
    param([string]$Email, [string]$OnPremHost)

    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $add = {
        param([string]$cat, [string]$label, [string]$url)
        if ($url) { $results.Add([PSCustomObject]@{ Protocol = $cat; Field = $label; Url = $url }) }
    }

    $domain = ''
    if ($Email -match '@([^@\s]+)$') { $domain = $Matches[1] }
    if (-not $OnPremHost -and $domain) { $OnPremHost = "autodiscover.$domain" }

    # AutoDiscover — on-prem and Exchange Online
    if ($domain) {
        & $add 'AutoDiscover' "On-prem (autodiscover.$domain)" "https://autodiscover.$domain/autodiscover/autodiscover.xml"
        & $add 'AutoDiscover' 'On-prem (root domain)'          "https://$domain/autodiscover/autodiscover.xml"
    }
    & $add 'AutoDiscover' 'Exchange Online (V1)' 'https://outlook.office365.com/autodiscover/autodiscover.xml'
    if ($Email) {
        & $add 'AutoDiscover' 'Exchange Online (V2 JSON)' `
            "https://outlook.office365.com/autodiscover/autodiscover.json?Email=$([Uri]::EscapeDataString($Email))&Protocol=Autodiscoverv1"
    }

    # OAuth / modern-auth metadata (hybrid modern authentication)
    if ($domain) {
        & $add 'OAuth' 'On-prem auth metadata' "https://autodiscover.$domain/autodiscover/metadata/json/1"
        & $add 'OAuth' 'Tenant OpenID config'  "https://login.microsoftonline.com/$domain/.well-known/openid-configuration"
    }

    # Free/Busy federation runs over EWS
    if ($OnPremHost) { & $add 'EWS' 'On-prem EWS'        "https://$OnPremHost/EWS/Exchange.asmx" }
    & $add 'EWS' 'Exchange Online EWS' 'https://outlook.office365.com/EWS/Exchange.asmx'

    # Per-vdir healthcheck.htm on the on-prem host
    if ($OnPremHost) {
        foreach ($vd in @('autodiscover','ews','oab','owa','ecp','mapi','rpc','Microsoft-Server-ActiveSync')) {
            & $add 'Health' $vd "https://$OnPremHost/$vd/healthcheck.htm"
        }
    }

    return ,$results
}

#endregion ===================================================================
#  TEST ENGINE  (PowerShell Runspace + WinForms Timer — avoids ThreadPool runspace issue)
#==============================================================================

# Script-level state for the currently running test
$script:CurrentPS      = $null
$script:CurrentRS      = $null
$script:CurrentSync    = $null
$script:PollTimer      = $null
$script:LastXml        = $null
$script:LastToken      = $null
$script:LastTokenExpiry = $null

# The actual test logic runs inside a dedicated PS runspace.
# $sync is the only bridge between the runspace and the UI thread.
$script:TestScript = {
    param($sync)

    $email      = $sync.Email
    $password   = $sync.Password
    $useWinAuth = $sync.UseWindowsAuth
    $domain     = ($email -split '@')[1]

    # Credentials
    $netCred = $null
    if (-not $useWinAuth -and $password -ne '') {
        $netCred = New-Object System.Net.NetworkCredential($email, $password)
    }

    # AutoDiscover POST body
    $bodyXml = "<?xml version=""1.0"" encoding=""utf-8""?>" +
        "<Autodiscover xmlns=""http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006"">" +
        "<Request><EMailAddress>$email</EMailAddress>" +
        "<AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema>" +
        "</Request></Autodiscover>"
    $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($bodyXml)

    # Helpers — all run inside the dedicated runspace, no closure issues
    $logLine = { param([string]$msg) $sync.Queue.Enqueue($msg) }
    $setPct  = { param([int]$pct)   $sync.Pct = $pct }

    # HTTP POST → @{Code; Body; Location; WwwAuth; Error}
    # Pass $authHeader to override credentials with an explicit Authorization value (e.g. Bearer token)
    $doPost = {
        param([string]$url, [string]$authHeader = '')
        try {
            $req = [System.Net.HttpWebRequest]::Create($url)
            $req.Method            = "POST"
            $req.ContentType       = "text/xml; charset=utf-8"
            $req.ContentLength     = $bodyBytes.Length
            $req.AllowAutoRedirect = $false
            $req.Timeout           = 30000
            $req.UserAgent                    = "Microsoft Office/16.0 (Windows NT 10.0)"
            $req.Headers["X-MapiHttpCapability"] = "1"   # request mapiHttp protocol block
            $req.Headers["X-ClientCanHandle"]    = "Negotiate"
            if ($authHeader -ne '') {
                $req.Headers["Authorization"] = $authHeader
            } elseif ($useWinAuth) {
                $req.UseDefaultCredentials = $true
            } elseif ($netCred) {
                $req.Credentials = $netCred
            }
            $s = $req.GetRequestStream()
            $s.Write($bodyBytes, 0, $bodyBytes.Length)
            $s.Close()
            try {
                $resp = $req.GetResponse()
                $code = [int]$resp.StatusCode
                $body = $null
                if ($code -eq 200) {
                    $rs   = $resp.GetResponseStream()
                    $rdr  = New-Object System.IO.StreamReader($rs, [System.Text.Encoding]::UTF8)
                    $body = $rdr.ReadToEnd()
                    $rdr.Close()
                }
                $resp.Close()
                return @{ Code = $code; Body = $body; Location = $null; WwwAuth = $null; Error = $null }
            } catch [System.Net.WebException] {
                $ex = $_.Exception
                if ($ex.Response) {
                    $code    = [int]$ex.Response.StatusCode
                    $loc          = $ex.Response.Headers["Location"]
                    $wwwAuthVals  = try { $ex.Response.Headers.GetValues("WWW-Authenticate") } catch { $null }
                    $wwwAuth      = if ($wwwAuthVals) { $wwwAuthVals -join ' ' } else { $null }
                    $ex.Response.Close()
                    return @{ Code = $code; Body = $null; Location = $loc; WwwAuth = $wwwAuth; Error = $null }
                }
                $emx = $ex.Message
                if ($ex.Status -eq [System.Net.WebExceptionStatus]::TrustFailure -or
                    $ex.Status -eq [System.Net.WebExceptionStatus]::SecureChannelFailure) {
                    $emx = "TLS/certificate not trusted — enable 'Ignore certificate errors' for on-premises servers with a self-signed or untrusted certificate. ($emx)"
                }
                return @{ Code = -1; Body = $null; Location = $null; WwwAuth = $null; Error = $emx }
            }
        } catch {
            $exo = $_.Exception
            $emo = $exo.Message
            if ($exo -is [System.Net.WebException] -and
                ($exo.Status -eq [System.Net.WebExceptionStatus]::TrustFailure -or
                 $exo.Status -eq [System.Net.WebExceptionStatus]::SecureChannelFailure)) {
                $emo = "TLS/certificate not trusted — enable 'Ignore certificate errors' for on-premises servers with a self-signed or untrusted certificate. ($emo)"
            }
            return @{ Code = -1; Body = $null; Location = $null; WwwAuth = $null; Error = $emo }
        }
    }

    # HTTP GET → @{Code; Location; Error}
    $doGet = {
        param([string]$url)
        try {
            $req = [System.Net.HttpWebRequest]::Create($url)
            $req.Method            = "GET"
            $req.AllowAutoRedirect = $false
            $req.Timeout           = 15000
            $req.UserAgent         = "Microsoft Office/16.0 (Windows NT 10.0)"
            try {
                $resp = $req.GetResponse()
                $code = [int]$resp.StatusCode
                $loc  = $resp.Headers["Location"]
                $resp.Close()
                return @{ Code = $code; Location = $loc; Error = $null }
            } catch [System.Net.WebException] {
                $ex = $_.Exception
                if ($ex.Response) {
                    $code = [int]$ex.Response.StatusCode
                    $loc  = $ex.Response.Headers["Location"]
                    $ex.Response.Close()
                    return @{ Code = $code; Location = $loc; Error = $null }
                }
                return @{ Code = -1; Location = $null; Error = $ex.Message }
            }
        } catch {
            return @{ Code = -1; Location = $null; Error = $_.Exception.Message }
        }
    }

    $tryModernAuth = $sync.ModernAuth
    $useSCP        = $sync.UseSCP

    # Token acquisition: Device Code Flow or Auth Code Flow + PKCE
    $getToken = {
        param([string]$wwwAuthHeader)

        if ($sync.UseDeviceCode) {
            # ── Device Code Flow ──────────────────────────────────────────────
            # Determine the authority. The .default scope cannot be combined with
            # the /common endpoint (AADSTS50059: no tenant-identifying info), so
            # resolve a concrete tenant where possible and use a specific EWS
            # delegated scope. Order: known TenantId → OIDC discovery on the
            # e-mail domain → /organizations fallback.
            $tenant = $sync.TenantId
            if (-not $tenant) {
                $dom = ($sync.Email -split '@')[1]
                if ($dom) {
                    try {
                        $oidcReq = [System.Net.HttpWebRequest]::Create(
                            "https://login.microsoftonline.com/$([Uri]::EscapeDataString($dom))/.well-known/openid-configuration")
                        $oidcReq.Method = "GET"; $oidcReq.Timeout = 8000
                        $oidcRp = $oidcReq.GetResponse()
                        $oidcJ  = (New-Object System.IO.StreamReader($oidcRp.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
                        $oidcRp.Close()
                        if ($oidcJ.token_endpoint -match '/([0-9a-fA-F-]{36})/') { $tenant = $Matches[1] }
                    } catch {}
                }
            }
            $authUri = if ($tenant) {
                "https://login.microsoftonline.com/$([Uri]::EscapeDataString($tenant))/oauth2/v2.0/authorize"
            } else {
                'https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize'
            }
            if ($wwwAuthHeader -match 'authorization_uri\s*=\s*"([^"]+)"') {
                $authUri = $Matches[1] -replace '/oauth2(?:/v2\.0)?/authorize.*', '/oauth2/v2.0/authorize'
            }
            $deviceCodeUrl = $authUri -replace '/authorize', '/devicecode'
            $tokenUrl      = $authUri -replace '/authorize', '/token'
            $clientId      = if ($sync.ClientId) { $sync.ClientId } else { 'd3590ed6-52b3-4102-aeff-aad2292ab01c' }
            $scope         = 'https://outlook.office365.com/EWS.AccessAsUser.All offline_access'

            & $logLine "Modern Auth: requesting device code…"

            # Step 1 — get device code
            $dcBytes = [System.Text.Encoding]::UTF8.GetBytes(
                "client_id=$([Uri]::EscapeDataString($clientId))&scope=$([Uri]::EscapeDataString($scope))")
            $dcJson = $null
            try {
                $rq = [System.Net.HttpWebRequest]::Create($deviceCodeUrl)
                $rq.Method        = "POST"
                $rq.ContentType   = "application/x-www-form-urlencoded"
                $rq.ContentLength = $dcBytes.Length
                $rq.Timeout       = 15000
                $ss = $rq.GetRequestStream(); $ss.Write($dcBytes, 0, $dcBytes.Length); $ss.Close()
                $rp = $rq.GetResponse()
                $dcJson = (New-Object System.IO.StreamReader($rp.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
                $rp.Close()
            } catch [System.Net.WebException] {
                $exT = $_.Exception
                if ($exT.Response) {
                    try {
                        $ej = (New-Object System.IO.StreamReader($exT.Response.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
                        $exT.Response.Close()
                        & $logLine "Modern Auth: device code request failed — $($ej.error): $($ej.error_description)"
                    } catch { & $logLine "Modern Auth: device code request failed — $($exT.Message)" }
                } else { & $logLine "Modern Auth: device code request failed — $($exT.Message)" }
                return $null
            } catch {
                & $logLine "Modern Auth: device code request error — $($_.Exception.Message)"
                return $null
            }

            $userCode   = $dcJson.user_code
            $deviceCode = $dcJson.device_code
            $verifyUri  = if ($dcJson.verification_uri) { $dcJson.verification_uri } else { $dcJson.verification_url }
            $pollSec    = [int]$dcJson.interval; if ($pollSec -lt 5) { $pollSec = 5 }

            & $logLine "Modern Auth: visit $verifyUri — enter code: $userCode"

            $sync.DeviceToken  = $null
            $sync.DeviceError  = $null
            $sync.DeviceCancel = $false

            # Step 2 — show dialog with code while polling
            $dcForm = New-Object System.Windows.Forms.Form
            $dcForm.Text            = "Sign in to Microsoft"
            $dcForm.Size            = New-Object System.Drawing.Size(440, 210)
            $dcForm.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
            $dcForm.MinimizeBox     = $false
            $dcForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog

            $lbl1 = New-Object System.Windows.Forms.Label
            $lbl1.Text     = "1.  Open a browser and go to:"
            $lbl1.Location = New-Object System.Drawing.Point(12, 14)
            $lbl1.AutoSize = $true

            $lnk = New-Object System.Windows.Forms.LinkLabel
            $lnk.Text     = $verifyUri
            $lnk.Location = New-Object System.Drawing.Point(28, 34)
            $lnk.AutoSize = $true
            $lnk.Add_LinkClicked({ [System.Diagnostics.Process]::Start($lnk.Text) })

            $lbl2 = New-Object System.Windows.Forms.Label
            $lbl2.Text     = "2.  Enter this code:"
            $lbl2.Location = New-Object System.Drawing.Point(12, 62)
            $lbl2.AutoSize = $true

            $lblCode = New-Object System.Windows.Forms.Label
            $lblCode.Text      = $userCode
            $lblCode.Font      = New-Object System.Drawing.Font("Consolas", 22, [System.Drawing.FontStyle]::Bold)
            $lblCode.Location  = New-Object System.Drawing.Point(28, 80)
            $lblCode.AutoSize  = $true
            $lblCode.ForeColor = [System.Drawing.Color]::DarkBlue

            $btnCopy = New-Object System.Windows.Forms.Button
            $btnCopy.Text     = "Copy"
            $btnCopy.Location = New-Object System.Drawing.Point(340, 82)
            $btnCopy.Size     = New-Object System.Drawing.Size(72, 26)
            $btnCopy.Add_Click({ [System.Windows.Forms.Clipboard]::SetText($userCode) })

            $lblWait = New-Object System.Windows.Forms.Label
            $lblWait.Text      = "Waiting for sign-in…"
            $lblWait.Location  = New-Object System.Drawing.Point(12, 144)
            $lblWait.AutoSize  = $true
            $lblWait.ForeColor = [System.Drawing.Color]::Gray

            $btnCancelDC = New-Object System.Windows.Forms.Button
            $btnCancelDC.Text     = "Cancel"
            $btnCancelDC.Location = New-Object System.Drawing.Point(340, 140)
            $btnCancelDC.Size     = New-Object System.Drawing.Size(72, 26)
            $btnCancelDC.Add_Click({ $sync.DeviceCancel = $true; $dcForm.Close() })

            $dcForm.Controls.AddRange(@($lbl1, $lnk, $lbl2, $lblCode, $btnCopy, $lblWait, $btnCancelDC))

            # Poll token endpoint on a timer; runs on UI thread so keep HTTP timeout short
            $pollTimer = New-Object System.Windows.Forms.Timer
            $pollTimer.Interval = $pollSec * 1000
            $pollTimer.Add_Tick({
                if ($sync.DeviceToken -or $sync.DeviceError -or $sync.DeviceCancel) { return }
                $pb = [System.Text.Encoding]::UTF8.GetBytes(
                    "grant_type=urn:ietf:params:oauth:grant-type:device_code" +
                    "&client_id=$([Uri]::EscapeDataString($clientId))" +
                    "&device_code=$([Uri]::EscapeDataString($deviceCode))")
                try {
                    $rq2 = [System.Net.HttpWebRequest]::Create($tokenUrl)
                    $rq2.Method        = "POST"
                    $rq2.ContentType   = "application/x-www-form-urlencoded"
                    $rq2.ContentLength = $pb.Length
                    $rq2.Timeout       = 4000
                    $ss2 = $rq2.GetRequestStream(); $ss2.Write($pb, 0, $pb.Length); $ss2.Close()
                    try {
                        $rp2  = $rq2.GetResponse()
                        $tokJ = (New-Object System.IO.StreamReader($rp2.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
                        $rp2.Close()
                        if ($tokJ.access_token) {
                            $sync.DeviceToken = "Bearer $($tokJ.access_token)"
                            $dcForm.Close()
                        }
                    } catch [System.Net.WebException] {
                        $ex2 = $_.Exception
                        if ($ex2.Response) {
                            $ej2 = (New-Object System.IO.StreamReader($ex2.Response.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
                            $ex2.Response.Close()
                            switch ($ej2.error) {
                                'authorization_pending' {}
                                'slow_down'             { $pollTimer.Interval += 5000 }
                                default {
                                    $sync.DeviceError = "$($ej2.error): $($ej2.error_description)"
                                    $dcForm.Close()
                                }
                            }
                        }
                    }
                } catch {}
            })

            $dcForm.Add_Shown({ $pollTimer.Start() })
            $dcForm.Add_FormClosed({ $pollTimer.Stop() })
            [void]$dcForm.ShowDialog()
            $pollTimer.Dispose()
            $dcForm.Dispose()

            if ($sync.DeviceCancel -or (-not $sync.DeviceToken -and -not $sync.DeviceError)) {
                & $logLine "Modern Auth: sign-in cancelled."
                return $null
            }
            if ($sync.DeviceError) {
                & $logLine "Modern Auth: sign-in error — $($sync.DeviceError)"
                $sync.DeviceError = $null
                return $null
            }
            & $logLine "Modern Auth: access token acquired."
            $tok = $sync.DeviceToken; $sync.DeviceToken = $null
            $sync.LastToken       = $tok
            $sync.LastTokenExpiry = (Get-Date).AddMinutes(55)
            return $tok

        } else {
            # ── Authorization Code Flow + PKCE ────────────────────────────────
            $clientId = $sync.ClientId
            $tenantId = $sync.TenantId

            # PKCE: 48-byte random verifier → SHA-256 → base64url challenge
            $rng = [System.Security.Cryptography.RNGCryptoServiceProvider]::new()
            $rngB = [byte[]]::new(48)
            $rng.GetBytes($rngB); $rng.Dispose()
            $verifier  = [Convert]::ToBase64String($rngB) -replace '\+','-' -replace '/','_' -replace '=',''
            $sha256    = [System.Security.Cryptography.SHA256Managed]::new()
            $chalB     = $sha256.ComputeHash([System.Text.Encoding]::ASCII.GetBytes($verifier)); $sha256.Dispose()
            $challenge = [Convert]::ToBase64String($chalB) -replace '\+','-' -replace '/','_' -replace '=',''

            # Find a free ephemeral port
            $tcpT = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
            $tcpT.Start(); $freePort = $tcpT.LocalEndpoint.Port; $tcpT.Stop()

            $redirectUri = "http://localhost:$freePort/"
            $scope       = 'https://outlook.office365.com/EWS.AccessAsUser.All offline_access'
            $stateVal    = [Guid]::NewGuid().ToString('N')
            $tokenUrl    = "https://login.microsoftonline.com/$([Uri]::EscapeDataString($tenantId))/oauth2/v2.0/token"
            $authUrl2    = "https://login.microsoftonline.com/$([Uri]::EscapeDataString($tenantId))/oauth2/v2.0/authorize" +
                "?client_id=$([Uri]::EscapeDataString($clientId))" +
                "&response_type=code" +
                "&redirect_uri=$([Uri]::EscapeDataString($redirectUri))" +
                "&scope=$([Uri]::EscapeDataString($scope))" +
                "&state=$([Uri]::EscapeDataString($stateVal))" +
                "&code_challenge=$([Uri]::EscapeDataString($challenge))" +
                "&code_challenge_method=S256" +
                "&prompt=select_account"

            # Synchronized hashtable for callback listener ↔ dialog thread
            $tSync = [hashtable]::Synchronized(@{
                Code = $null; Err = $null; Cancel = $false; Done = $false; DotCount = 0
            })

            # Background runspace: HttpListener waits for the OAuth2 redirect callback
            $tRs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
            $tRs.ApartmentState = [System.Threading.ApartmentState]::STA
            $tRs.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
            $tRs.Open()
            $tRs.SessionStateProxy.SetVariable('tSync',    $tSync)
            $tRs.SessionStateProxy.SetVariable('freePort', $freePort)
            $tRs.SessionStateProxy.SetVariable('stateVal', $stateVal)

            $tPs = [System.Management.Automation.PowerShell]::Create()
            $tPs.Runspace = $tRs
            [void]$tPs.AddScript({
                $hl = [System.Net.HttpListener]::new()
                $hl.Prefixes.Add("http://localhost:$freePort/")
                $hl.Start()
                try {
                    $ar       = $hl.BeginGetContext($null, $null)
                    $deadline = (Get-Date).AddMinutes(5)
                    while (-not $ar.IsCompleted -and (Get-Date) -lt $deadline) {
                        if ($tSync.Cancel) { $hl.Stop(); return }
                        [System.Threading.Thread]::Sleep(200)
                    }
                    if (-not $ar.IsCompleted) {
                        $hl.Stop()
                        if (-not $tSync.Cancel) { $tSync.Err = 'Sign-in timed out (5 min)' }
                        return
                    }
                    $ctx  = $hl.EndGetContext($ar)
                    $qs   = $ctx.Request.QueryString
                    $code = $qs['code']; $retState = $qs['state']
                    $oErr = $qs['error']; $oErrD = $qs['error_description']
                    $ok   = $code -and $retState -eq $stateVal
                    $html = if ($ok) {
                        '<html><body style="font-family:sans-serif;padding:40px"><h2 style="color:green">&#10003; Signed in successfully</h2><p>You may close this tab and return to Exchange Tester.</p></body></html>'
                    } else {
                        '<html><body style="font-family:sans-serif;padding:40px"><h2 style="color:red">&#10007; Sign-in failed or cancelled</h2><p>You may close this tab.</p></body></html>'
                    }
                    $hb = [System.Text.Encoding]::UTF8.GetBytes($html)
                    $ctx.Response.ContentType = 'text/html; charset=utf-8'
                    $ctx.Response.ContentLength64 = $hb.Length
                    $ctx.Response.OutputStream.Write($hb, 0, $hb.Length)
                    $ctx.Response.OutputStream.Close()
                    $ctx.Response.Close()
                    if ($ok)       { $tSync.Code = $code }
                    elseif ($oErr) { $tSync.Err  = if ($oErrD) { "${oErr}: $oErrD" } else { $oErr } }
                    else           { $tSync.Err  = 'No authorization code received' }
                } catch {
                    if (-not $tSync.Cancel) { $tSync.Err = $_.Exception.Message }
                } finally {
                    try { $hl.Stop(); $hl.Close() } catch {}
                    $tSync.Done = $true
                }
            })
            [void]$tPs.BeginInvoke()

            & $logLine "Modern Auth: opening browser for sign-in…"
            [System.Diagnostics.Process]::Start($authUrl2) | Out-Null

            # Waiting dialog on this STA thread
            $dcForm = New-Object System.Windows.Forms.Form
            $dcForm.Text            = "Sign in to Microsoft"
            $dcForm.Size            = New-Object System.Drawing.Size(480, 155)
            $dcForm.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
            $dcForm.MinimizeBox     = $false
            $dcForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog

            $lbl0 = New-Object System.Windows.Forms.Label
            $lbl0.Text     = "Sign in to Microsoft in your browser."
            $lbl0.Location = New-Object System.Drawing.Point(12, 12)
            $lbl0.Size     = New-Object System.Drawing.Size(450, 18)
            $lbl0.Font     = New-Object System.Drawing.Font($dcForm.Font, [System.Drawing.FontStyle]::Bold)

            $lblWait = New-Object System.Windows.Forms.Label
            $lblWait.Text      = "The sign-in page has been opened in your browser. Waiting for authentication…"
            $lblWait.Location  = New-Object System.Drawing.Point(12, 38)
            $lblWait.Size      = New-Object System.Drawing.Size(450, 18)
            $lblWait.ForeColor = [System.Drawing.Color]::Gray

            $lblDots = New-Object System.Windows.Forms.Label
            $lblDots.Text      = ""
            $lblDots.Location  = New-Object System.Drawing.Point(12, 62)
            $lblDots.AutoSize  = $true
            $lblDots.ForeColor = [System.Drawing.Color]::SteelBlue

            $btnCancelACF = New-Object System.Windows.Forms.Button
            $btnCancelACF.Text     = "Cancel"
            $btnCancelACF.Location = New-Object System.Drawing.Point(390, 90)
            $btnCancelACF.Size     = New-Object System.Drawing.Size(72, 26)
            $btnCancelACF.Add_Click({ $tSync.Cancel = $true; $dcForm.Close() })

            $dcForm.Controls.AddRange(@($lbl0, $lblWait, $lblDots, $btnCancelACF))

            $pollACF = New-Object System.Windows.Forms.Timer; $pollACF.Interval = 300
            $pollACF.Add_Tick({
                if ($tSync.Done -or $tSync.Cancel) { $dcForm.Close(); return }
                $tSync.DotCount = (($tSync.DotCount + 1) % 6)
                $lblDots.Text = '.' * ($tSync.DotCount + 1)
            })
            $dcForm.Add_Shown({ $tSync.DotCount = 0; $pollACF.Start() })
            $dcForm.Add_FormClosed({ $pollACF.Stop() })
            [void]$dcForm.ShowDialog()
            $pollACF.Dispose(); $dcForm.Dispose()

            try { $tPs.Stop() } catch {}
            $tPs.Dispose(); $tRs.Close(); $tRs.Dispose()

            if ($tSync.Cancel -or (-not $tSync.Code -and -not $tSync.Err)) {
                & $logLine "Modern Auth: sign-in cancelled."
                return $null
            }
            if ($tSync.Err) {
                & $logLine "Modern Auth: sign-in error — $($tSync.Err)"
                return $null
            }

            # Exchange authorization code for access token
            & $logLine "Modern Auth: exchanging code for access token…"
            $tokBody  = "grant_type=authorization_code" +
                "&client_id=$([Uri]::EscapeDataString($clientId))" +
                "&code=$([Uri]::EscapeDataString($tSync.Code))" +
                "&redirect_uri=$([Uri]::EscapeDataString($redirectUri))" +
                "&code_verifier=$([Uri]::EscapeDataString($verifier))"
            $tokBytes = [System.Text.Encoding]::UTF8.GetBytes($tokBody)
            $accessTok = $null
            try {
                $rq2 = [System.Net.HttpWebRequest]::Create($tokenUrl)
                $rq2.Method = "POST"; $rq2.ContentType = "application/x-www-form-urlencoded"
                $rq2.ContentLength = $tokBytes.Length; $rq2.Timeout = 15000
                $ss2 = $rq2.GetRequestStream(); $ss2.Write($tokBytes, 0, $tokBytes.Length); $ss2.Close()
                $rp2 = $rq2.GetResponse()
                $tJ  = (New-Object System.IO.StreamReader($rp2.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
                $rp2.Close()
                $accessTok = $tJ.access_token
            } catch [System.Net.WebException] {
                $exT = $_.Exception
                $m2 = if ($exT.Response) {
                    try { $ej=(New-Object System.IO.StreamReader($exT.Response.GetResponseStream())).ReadToEnd()|ConvertFrom-Json; $exT.Response.Close(); "$($ej.error): $($ej.error_description)" } catch { $exT.Message }
                } else { $exT.Message }
                & $logLine "Modern Auth: token exchange failed — $m2"
                return $null
            } catch {
                & $logLine "Modern Auth: token exchange failed — $($_.Exception.Message)"
                return $null
            }
            if (-not $accessTok) {
                & $logLine "Modern Auth: token exchange returned no access token."
                return $null
            }
            & $logLine "Modern Auth: access token acquired."
            $sync.LastToken       = "Bearer $accessTok"
            $sync.LastTokenExpiry = (Get-Date).AddMinutes(55)
            return "Bearer $accessTok"
        }
    }

    # Try one AutoDiscover POST URL; returns XML string on 200, $null otherwise
    $tryUrl = {
        param([string]$url)
        & $logLine "AutoDiscover via $url starting."
        $res = & $doPost $url

        if ($res.Code -eq 200) {
            & $logLine "GetLastError=0; httpStatus=200."
            & $logLine "AutoDiscover via $url succeeded."
            return $res.Body
        }

        if ($res.Code -ge 0) {
            & $logLine "GetLastError=0; httpStatus=$($res.Code)."
        } else {
            & $logLine "AutoDiscover via $url failed: $($res.Error)"
            return $null
        }

        if ($res.Code -eq 301 -or $res.Code -eq 302) {
            $loc = $res.Location
            if ($loc) { & $logLine "URL redirect to $loc at AutoDiscover." }
            if ($loc -and $loc -match '^https://' -and $loc -match 'autodiscover') {
                # Redirect points directly at an autodiscover endpoint — follow it
                & $logLine "Redirect check for $loc starting."
                $res2 = & $doPost $loc
                if ($res2.Code -ge 0) { & $logLine "GetLastError=0; httpStatus=$($res2.Code)." }
                if ($res2.Code -eq 200) {
                    & $logLine "AutoDiscover via $loc succeeded."
                    return $res2.Body
                }
                & $logLine "Redirect check for $loc failed (0x800C8209)."
            } elseif ($loc) {
                # Non-autodiscover redirect target — GET to verify it's a real endpoint
                & $logLine "Redirect check for $loc starting."
                $res2 = & $doGet $loc
                if ($res2.Code -ge 0) { & $logLine "GetLastError=0; httpStatus=$($res2.Code)." }
                & $logLine "Redirect check for $loc failed (0x800C8209)."
            }
            & $logLine "AutoDiscover via $url failed (0x800C8204)."
            return $null
        }

        if ($res.Code -eq 401) {
            if ($res.WwwAuth) {
                & $logLine "  WWW-Authenticate: $($res.WwwAuth)"
            } else {
                & $logLine "  WWW-Authenticate: (not present)"
            }
            $hasBearerChallenge = $res.WwwAuth -and $res.WwwAuth -match 'Bearer'
            $isO365Endpoint     = $url -match 'outlook\.office365\.com'
            if ($tryModernAuth -and ($hasBearerChallenge -or $isO365Endpoint)) {
                if ($isO365Endpoint -and -not $hasBearerChallenge) {
                    & $logLine "  O365 endpoint detected — attempting Modern Auth proactively (Bearer not advertised)."
                }
                & $logLine "Modern Auth: initiating OAuth2 Authorization Code Flow…"
                $token = & $getToken $res.WwwAuth
                if ($token) {
                    & $logLine "Retrying AutoDiscover with Bearer token."
                    $res2 = & $doPost $url $token
                    if ($res2.Code -ge 0) { & $logLine "GetLastError=0; httpStatus=$($res2.Code)." }
                    if ($res2.Code -eq 200) {
                        & $logLine "AutoDiscover via $url succeeded (Modern Auth)."
                        return $res2.Body
                    }
                    & $logLine "AutoDiscover via $url failed after Modern Auth (httpStatus=$($res2.Code))."
                    return $null
                }
            } elseif ($tryModernAuth) {
                & $logLine "  No Bearer challenge — HMA/OAuth2 not offered by this endpoint."
            }
            & $logLine "AutoDiscover via $url failed (0x800C820E)."
            return $null
        }

        & $logLine "AutoDiscover via $url failed (httpStatus=$($res.Code))."
        return $null
    }

    #endregion

    $foundXml = $null

    # --- Step 1: SCP — runs first when "Use SCP" is checked (domain-joined priority) ---
    & $setPct 5
    if (-not $sync.Cancel -and $useSCP) {
        & $logLine "Local AutoDiscover for $domain starting (SCP)."
        try {
            # Exchange AutoDiscover SCPs are stored in the Configuration partition,
            # not the default domain partition — must set search root explicitly.
            $rootDSE    = New-Object System.DirectoryServices.DirectoryEntry("LDAP://RootDSE")
            $configNC   = $rootDSE.Properties["configurationNamingContext"].Value
            & $logLine "  SCP search root: $configNC"
            $searchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$configNC")

            $searcher = New-Object System.DirectoryServices.DirectorySearcher($searchRoot)
            # Match on the Exchange AutoDiscover GUID keyword (with or without braces)
            # OR on serviceBindingInformation containing "autodiscover" as a fallback.
            $searcher.Filter      = "(&(objectClass=serviceConnectionPoint)(|(serviceBindingInformation=*autodiscover*)(keywords=67661d7F-8FC4-4fa7-BFAC-E1D7794C1F68)(keywords={67661d7F-8FC4-4fa7-BFAC-E1D7794C1F68})))"
            $searcher.SearchScope = [System.DirectoryServices.SearchScope]::Subtree
            [void]$searcher.PropertiesToLoad.Add("serviceBindingInformation")

            $scpHits = $searcher.FindAll()
            & $logLine "  SCP records found: $($scpHits.Count)"
            if ($scpHits.Count -eq 0) {
                & $logLine "Local AutoDiscover for $domain failed (0x8004010F)."
            } else {
                foreach ($hit in $scpHits) {
                    $scpUrl = $hit.Properties["serviceBindingInformation"][0]
                    & $logLine "SCP record: $scpUrl"
                    $res = & $doPost $scpUrl
                    if ($res.Code -ge 0) { & $logLine "GetLastError=0; httpStatus=$($res.Code)." }
                    if ($res.Code -eq 200) {
                        $foundXml = $res.Body
                        & $logLine "Local AutoDiscover via SCP succeeded."
                        break
                    }
                }
                if (-not $foundXml) { & $logLine "Local AutoDiscover for $domain failed." }
            }
        } catch {
            & $logLine "Local AutoDiscover for $domain failed (0x8004010F): $($_.Exception.Message)"
        }
    }

    # --- Step 2: O365 ---
    & $setPct 20
    if (-not $foundXml -and -not $sync.Cancel) {
        $foundXml = & $tryUrl "https://outlook.office365.com/autodiscover/autodiscover.xml"
    }

    # --- Step 3: https://<domain>/autodiscover/autodiscover.xml ---
    & $setPct 40
    if (-not $foundXml -and -not $sync.Cancel) {
        $foundXml = & $tryUrl "https://$domain/autodiscover/autodiscover.xml"
    }

    # --- Step 4: https://autodiscover.<domain>/autodiscover/autodiscover.xml ---
    & $setPct 58
    if (-not $foundXml -and -not $sync.Cancel) {
        $foundXml = & $tryUrl "https://autodiscover.$domain/autodiscover/autodiscover.xml"
    }

    # --- Step 5: HTTP redirect check (well-known URL) ---
    & $setPct 74
    if (-not $foundXml -and -not $sync.Cancel) {
        $rdUrl = "http://autodiscover.$domain/autodiscover/autodiscover.xml"
        & $logLine "Redirect check for $rdUrl starting."
        $res = & $doGet $rdUrl
        if ($res.Code -ge 0) { & $logLine "GetLastError=0; httpStatus=$($res.Code)." }
        if ($res.Location -and $res.Location -match '^https://') {
            & $logLine "Redirect to $($res.Location) found."
            $res2 = & $doPost $res.Location
            if ($res2.Code -ge 0) { & $logLine "GetLastError=0; httpStatus=$($res2.Code)." }
            if ($res2.Code -eq 200) {
                $foundXml = $res2.Body
                & $logLine "AutoDiscover via HTTP redirect succeeded."
            } else {
                & $logLine "Redirect check for $rdUrl failed (0x80004005)."
            }
        } else {
            & $logLine "Redirect check for $rdUrl failed (0x80004005)."
        }
    }

    # --- Step 6: DNS SRV _autodiscover._tcp.<domain> ---
    & $setPct 88
    if (-not $foundXml -and -not $sync.Cancel) {
        & $logLine "DNS SRV lookup for $domain starting."
        try {
            $srvRecs = Resolve-DnsName -Name "_autodiscover._tcp.$domain" -Type SRV -ErrorAction Stop
            $srvHit = $false
            foreach ($srv in $srvRecs) {
                if (-not $srv.NameTarget) { continue }
                $srvUrl = "https://$($srv.NameTarget):$($srv.Port)/autodiscover/autodiscover.xml"
                & $logLine "DNS SRV: $($srv.NameTarget):$($srv.Port)"
                $res = & $doPost $srvUrl
                if ($res.Code -ge 0) { & $logLine "GetLastError=0; httpStatus=$($res.Code)." }
                if ($res.Code -eq 200) {
                    $foundXml = $res.Body
                    $srvHit   = $true
                    & $logLine "AutoDiscover via DNS SRV succeeded."
                    break
                }
            }
            if (-not $srvHit) { & $logLine "DNS SRV lookup for $domain failed (0x8004010F)." }
        } catch [System.Management.Automation.CommandNotFoundException] {
            & $logLine "DNS SRV skipped (Resolve-DnsName requires Windows 8.1+)."
        } catch {
            & $logLine "DNS SRV lookup for $domain failed (0x8004010F)."
        }
    }

    if ($sync.Cancel) { $sync.Done = $true; return }

    & $setPct 100
    $sync.Xml  = $foundXml
    $sync.Done = $true
}

# Called from the poll timer when $sync.Done becomes $true
function Complete-Test {
    $script:PollTimer.Stop()

    # Propagate cached auth token so Additional Tests can reuse it without re-auth
    if ($script:CurrentSync.LastToken) {
        $script:LastToken       = $script:CurrentSync.LastToken
        $script:LastTokenExpiry = $script:CurrentSync.LastTokenExpiry
    }

    # Drain any remaining log lines
    $msg = $null
    while ($script:CurrentSync.Queue.TryDequeue([ref]$msg)) {
        $rtbLog.AppendText("$msg`r`n")
    }

    $btnTest.Enabled     = $true
    $btnCancel.Enabled   = $false
    $btnAddTests.Enabled = $script:LastXml -ne $null

    if ($script:CurrentSync.Cancel) {
        $form.Text = "Test E-Mail AutoConfiguration"
        $rtbLog.AppendText("`r`nTest cancelled.`r`n")
    } elseif ($script:CurrentSync.Xml) {
        $xml = $script:CurrentSync.Xml
        $script:LastXml      = $xml
        $btnAddTests.Enabled = $true

        # XML tab: pretty-print
        try {
            $xd = New-Object System.Xml.XmlDocument
            $xd.LoadXml($xml)
            $sb = New-Object System.Text.StringBuilder
            $sw = New-Object System.IO.StringWriter($sb)
            $xw = New-Object System.Xml.XmlTextWriter($sw)
            $xw.Formatting  = [System.Xml.Formatting]::Indented
            $xw.Indentation = 2
            $xd.WriteTo($xw)
            $xw.Flush()
            $rtbXml.Text = $sb.ToString()
        } catch {
            $rtbXml.Text = $xml
        }
        $rtbXml.SelectionStart = 0
        $rtbXml.ScrollToCaret()

        # Results tab: parsed rows grouped by section
        $rows = ConvertFrom-AutodiscoverXml -RawXml $xml
        $lvwResults.BeginUpdate()
        $lvwResults.Items.Clear()
        $lvwResults.Groups.Clear()
        $groupMap = @{}
        foreach ($row in $rows) {
            if (-not $groupMap.ContainsKey($row.Group)) {
                $lvg = New-Object System.Windows.Forms.ListViewGroup($row.Group, $row.Group)
                [void]$lvwResults.Groups.Add($lvg)
                $groupMap[$row.Group] = $lvg
            }
            $item = New-Object System.Windows.Forms.ListViewItem($row.Setting)
            [void]$item.SubItems.Add($row.Value)
            $item.Group = $groupMap[$row.Group]
            [void]$lvwResults.Items.Add($item)
        }
        $lvwResults.EndUpdate()

        $form.Text = "Test E-Mail AutoConfiguration  —  OK"
        $rtbLog.AppendText("`r`nAutoDiscover completed successfully.`r`n")
        if ($rows.Count -gt 0) { $tabCtrl.SelectedTab = $tabResults }
        else                    { $tabCtrl.SelectedTab = $tabXml }
    } else {
        $form.Text = "Test E-Mail AutoConfiguration  —  No configuration found"
        $rtbLog.AppendText("`r`nAutoDiscover failed for all tested methods.`r`n")
        $tabCtrl.SelectedTab = $tabLog
    }

    # Clean up runspace
    try { $script:CurrentPS.Dispose() }   catch {}
    try { $script:CurrentRS.Close();  $script:CurrentRS.Dispose() } catch {}
    $script:CurrentPS   = $null
    $script:CurrentRS   = $null
    $script:CurrentSync = $null
}

#endregion ===================================================================
#  CONTROL INTERACTIONS
#==============================================================================

# Adjusts visible/enabled state of all sub-controls based on current auth selection
function Update-AuthMode {
    $isModern = $radModernAuth.Checked
    $pnlModernSub.Enabled      = $isModern
    $chkUseCurrentUser.Enabled = -not $isModern
    $useExplicit = (-not $isModern) -and (-not $chkUseCurrentUser.Checked)
    $lblPass.Enabled = $useExplicit
    $txtPass.Enabled = $useExplicit
    if (-not $useExplicit) { $txtPass.Clear() }
    Update-Layout
}

# Repositions standalone options, separator, progress bar, and tab control
# based on whether ACF rows are visible
function Update-Layout {
    $showACF = $radACF.Checked -and $radModernAuth.Checked
    $lblClientId.Visible  = $showACF
    $txtClientId.Visible  = $showACF
    $btnCreateApp.Visible = $showACF
    $lblTenantId.Visible  = $showACF
    $txtTenantId.Visible  = $showACF
    $y = if ($showACF) { 179 } else { 133 }
    $chkIgnoreCert.Top = $y
    $chkUseSCP.Top     = $y
    $pnlSep.Top        = $y + 23
    $prgBar.Top        = $y + 31
    $tabCtrl.Top       = $y + 50
    $tabCtrl.Height    = $form.ClientSize.Height - ($y + 50) - 9
}

# Looks up the tenant ID for the domain in $txtEmail via the OIDC discovery endpoint.
# Runs synchronously with DoEvents so "Detecting…" is visible before the HTTP call.
# Triggered when ACF is selected or when the email field loses focus while ACF is active.
function Resolve-TenantId {
    if (-not $radACF.Checked) { return }
    $email = $txtEmail.Text.Trim()
    if ($email -notmatch '@([^@\s]+)$') { return }
    $domain = $Matches[1]
    $txtTenantId.Text    = "Detecting…"
    $txtTenantId.Enabled = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()
    $tid = $null
    try {
        $req = [System.Net.HttpWebRequest]::Create(
            "https://login.microsoftonline.com/$([Uri]::EscapeDataString($domain))/.well-known/openid-configuration")
        $req.Method  = "GET"
        $req.Timeout = 8000
        $rp = $req.GetResponse()
        $j  = (New-Object System.IO.StreamReader($rp.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
        $rp.Close()
        if ($j.token_endpoint -match '/([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})/') {
            $tid = $Matches[1]
        }
    } catch {}
    $txtTenantId.Text    = if ($tid) { $tid } else { '' }
    $txtTenantId.Enabled = $true
    $form.Cursor         = [System.Windows.Forms.Cursors]::Default
}

$radModernAuth.Add_CheckedChanged({ if ($radModernAuth.Checked) { Update-AuthMode } })
$radWIA.Add_CheckedChanged({        if ($radWIA.Checked)        { Update-AuthMode } })
$radDCF.Add_CheckedChanged({        if ($radDCF.Checked)        { Update-Layout  } })
$radACF.Add_CheckedChanged({        if ($radACF.Checked)        { Update-Layout; Resolve-TenantId } })
$txtEmail.Add_Leave({ Resolve-TenantId })

$chkUseCurrentUser.Add_CheckedChanged({
    $useExplicit     = -not $chkUseCurrentUser.Checked
    $lblPass.Enabled = $useExplicit
    $txtPass.Enabled = $useExplicit
    if (-not $useExplicit) { $txtPass.Clear() }
})

# Config file path (same directory as this script)
$_scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path $MyInvocation.MyCommand.Path }
$script:configPath = Join-Path $_scriptDir "ExchangeTester.config"

$btnCreateApp.Add_Click({
    $tenantId = $txtTenantId.Text.Trim()
    if (-not $tenantId) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please detect or enter the Tenant ID first.",
            "Tenant ID Required",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    # ── Auth Code Flow + PKCE via local HTTP listener ────────────────────────
    $graphApp = '14d82eec-204b-4c2f-b7e8-296a70dab67e'   # Microsoft Graph Command Line Tools

    # PKCE: 48-byte random verifier → SHA-256 → base64url challenge
    $rng = [System.Security.Cryptography.RNGCryptoServiceProvider]::new()
    $rngBytes = [byte[]]::new(48)
    $rng.GetBytes($rngBytes); $rng.Dispose()
    $verifier  = [Convert]::ToBase64String($rngBytes) -replace '\+','-' -replace '/','_' -replace '=',''
    $sha256    = [System.Security.Cryptography.SHA256Managed]::new()
    $chalBytes = $sha256.ComputeHash([System.Text.Encoding]::ASCII.GetBytes($verifier)); $sha256.Dispose()
    $challenge = [Convert]::ToBase64String($chalBytes) -replace '\+','-' -replace '/','_' -replace '=',''

    # Find a free ephemeral port
    $tcpT = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
    $tcpT.Start(); $freePort = $tcpT.LocalEndpoint.Port; $tcpT.Stop()

    $redirectUri = "http://localhost:$freePort/"
    $graphScope  = 'https://graph.microsoft.com/Application.ReadWrite.All ' +
                   'https://graph.microsoft.com/AppRoleAssignment.ReadWrite.All ' +
                   'https://graph.microsoft.com/Directory.ReadWrite.All ' +
                   'https://graph.microsoft.com/Directory.AccessAsUser.All ' +
                   'offline_access'
    $stateVal    = [Guid]::NewGuid().ToString('N')
    $authUrl     = "https://login.microsoftonline.com/$([Uri]::EscapeDataString($tenantId))/oauth2/v2.0/authorize" +
        "?client_id=$([Uri]::EscapeDataString($graphApp))" +
        "&response_type=code" +
        "&redirect_uri=$([Uri]::EscapeDataString($redirectUri))" +
        "&scope=$([Uri]::EscapeDataString($graphScope))" +
        "&state=$([Uri]::EscapeDataString($stateVal))" +
        "&code_challenge=$([Uri]::EscapeDataString($challenge))" +
        "&code_challenge_method=S256" +
        "&prompt=select_account"

    # Synchronized hashtable for listener ↔ UI thread communication
    $lsync = [hashtable]::Synchronized(@{ Code = $null; Err = $null; Cancel = $false; Done = $false })

    # Background runspace: listens for the OAuth2 redirect callback
    $lRs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $lRs.ApartmentState = [System.Threading.ApartmentState]::STA
    $lRs.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $lRs.Open()
    $lRs.SessionStateProxy.SetVariable('lsync',    $lsync)
    $lRs.SessionStateProxy.SetVariable('freePort', $freePort)
    $lRs.SessionStateProxy.SetVariable('stateVal', $stateVal)

    $lPs = [System.Management.Automation.PowerShell]::Create()
    $lPs.Runspace = $lRs
    [void]$lPs.AddScript({
        $hl = [System.Net.HttpListener]::new()
        $hl.Prefixes.Add("http://localhost:$freePort/")
        $hl.Start()
        try {
            $ar       = $hl.BeginGetContext($null, $null)
            $deadline = (Get-Date).AddMinutes(5)
            while (-not $ar.IsCompleted -and (Get-Date) -lt $deadline) {
                if ($lsync.Cancel) { $hl.Stop(); return }
                [System.Threading.Thread]::Sleep(200)
            }
            if (-not $ar.IsCompleted) {
                $hl.Stop()
                if (-not $lsync.Cancel) { $lsync.Err = 'Sign-in timed out (5 min)' }
                return
            }
            $ctx  = $hl.EndGetContext($ar)
            $qs   = $ctx.Request.QueryString
            $code = $qs['code']; $retState = $qs['state']
            $oErr = $qs['error']; $oErrD = $qs['error_description']
            $ok   = $code -and $retState -eq $stateVal
            $html = if ($ok) {
                '<html><body style="font-family:sans-serif;padding:40px"><h2 style="color:green">&#10003; Signed in successfully</h2><p>You may close this tab and return to Exchange Tester.</p></body></html>'
            } else {
                '<html><body style="font-family:sans-serif;padding:40px"><h2 style="color:red">&#10007; Sign-in failed or cancelled</h2><p>You may close this tab.</p></body></html>'
            }
            $hb = [System.Text.Encoding]::UTF8.GetBytes($html)
            $ctx.Response.ContentType = 'text/html; charset=utf-8'
            $ctx.Response.ContentLength64 = $hb.Length
            $ctx.Response.OutputStream.Write($hb, 0, $hb.Length)
            $ctx.Response.OutputStream.Close()
            $ctx.Response.Close()
            if ($ok)       { $lsync.Code = $code }
            elseif ($oErr) { $lsync.Err  = if ($oErrD) { "${oErr}: $oErrD" } else { $oErr } }
            else           { $lsync.Err  = 'No authorization code received' }
        } catch {
            if (-not $lsync.Cancel) { $lsync.Err = $_.Exception.Message }
        } finally {
            try { $hl.Stop(); $hl.Close() } catch {}
            $lsync.Done = $true
        }
    })
    [void]$lPs.BeginInvoke()

    # Open the system browser at the auth URL
    [System.Diagnostics.Process]::Start($authUrl) | Out-Null

    # ── "Waiting for browser sign-in" dialog ───────────────────────────────────
    $aForm = New-Object System.Windows.Forms.Form
    $aForm.Text            = "Sign in  —  Create App Registration"
    $aForm.Size            = New-Object System.Drawing.Size(520, 155)
    $aForm.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $aForm.MinimizeBox     = $false
    $aForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog

    $aL0 = New-Object System.Windows.Forms.Label
    $aL0.Text     = "Sign in as Application Administrator or Global Admin."
    $aL0.Location = New-Object System.Drawing.Point(12, 12)
    $aL0.Size     = New-Object System.Drawing.Size(490, 18)
    $aL0.Font     = New-Object System.Drawing.Font($aForm.Font, [System.Drawing.FontStyle]::Bold)

    $aWait = New-Object System.Windows.Forms.Label
    $aWait.Text      = "The sign-in page has been opened in your browser. Waiting for authentication…"
    $aWait.Location  = New-Object System.Drawing.Point(12, 38)
    $aWait.Size      = New-Object System.Drawing.Size(490, 18)
    $aWait.ForeColor = [System.Drawing.Color]::Gray

    $aDots = New-Object System.Windows.Forms.Label
    $aDots.Text      = ""
    $aDots.Location  = New-Object System.Drawing.Point(12, 62)
    $aDots.AutoSize  = $true
    $aDots.ForeColor = [System.Drawing.Color]::SteelBlue

    $aCancel = New-Object System.Windows.Forms.Button
    $aCancel.Text     = "Cancel"
    $aCancel.Location = New-Object System.Drawing.Point(430, 90)
    $aCancel.Size     = New-Object System.Drawing.Size(72, 26)
    $aCancel.Add_Click({ $lsync.Cancel = $true; $aForm.Close() })

    $aForm.Controls.AddRange(@($aL0, $aWait, $aDots, $aCancel))

    $aPoll = New-Object System.Windows.Forms.Timer; $aPoll.Interval = 300
    $aPoll.Add_Tick({
        if ($lsync.Done -or $lsync.Cancel) { $aForm.Close(); return }
        $lsync.DotCount = (($lsync.DotCount + 1) % 6)
        $aDots.Text = '.' * ($lsync.DotCount + 1)
    })
    $aForm.Add_Shown({ $lsync.DotCount = 0; $aPoll.Start() })
    $aForm.Add_FormClosed({ $aPoll.Stop() })
    [void]$aForm.ShowDialog()
    $aPoll.Dispose(); $aForm.Dispose()

    # Clean up listener runspace
    try { $lPs.Stop() } catch {}
    $lPs.Dispose(); $lRs.Close(); $lRs.Dispose()

    if ($lsync.Cancel -or (-not $lsync.Code -and -not $lsync.Err)) { return }
    if ($lsync.Err) {
        [System.Windows.Forms.MessageBox]::Show("Sign-in error: $($lsync.Err)", "Auth Error",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }

    # ── Exchange authorization code for access token ───────────────────────────
    $tokUrl   = "https://login.microsoftonline.com/$([Uri]::EscapeDataString($tenantId))/oauth2/v2.0/token"
    $tokBody  = "grant_type=authorization_code" +
        "&client_id=$([Uri]::EscapeDataString($graphApp))" +
        "&code=$([Uri]::EscapeDataString($lsync.Code))" +
        "&redirect_uri=$([Uri]::EscapeDataString($redirectUri))" +
        "&code_verifier=$([Uri]::EscapeDataString($verifier))"
    $tokBytes = [System.Text.Encoding]::UTF8.GetBytes($tokBody)
    $adminToken = $null
    try {
        $rq2 = [System.Net.HttpWebRequest]::Create($tokUrl)
        $rq2.Method = "POST"; $rq2.ContentType = "application/x-www-form-urlencoded"
        $rq2.ContentLength = $tokBytes.Length; $rq2.Timeout = 15000
        $ss2 = $rq2.GetRequestStream(); $ss2.Write($tokBytes, 0, $tokBytes.Length); $ss2.Close()
        $rp2 = $rq2.GetResponse()
        $tJ  = (New-Object System.IO.StreamReader($rp2.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
        $rp2.Close()
        $adminToken = $tJ.access_token
    } catch [System.Net.WebException] {
        $exT = $_.Exception
        $msg = if ($exT.Response) {
            try { $ej=(New-Object System.IO.StreamReader($exT.Response.GetResponseStream())).ReadToEnd()|ConvertFrom-Json; $exT.Response.Close(); "$($ej.error): $($ej.error_description)" } catch { $exT.Message }
        } else { $exT.Message }
        [System.Windows.Forms.MessageBox]::Show("Token exchange failed: $msg", "Auth Error",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Token exchange failed: $($_.Exception.Message)", "Auth Error",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }
    if (-not $adminToken) {
        [System.Windows.Forms.MessageBox]::Show("Token exchange returned no access token.", "Auth Error",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }

    # ── Create app registration via Microsoft Graph ───────────────────────────
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()

    $appBody  = ('{"displayName":"Exchange Tester","isFallbackPublicClient":true,' +
                '"publicClient":{"redirectUris":["http://localhost"]},' +
                '"requiredResourceAccess":[' +
                  '{"resourceAppId":"00000002-0000-0ff1-ce00-000000000000",' +   # Exchange Online
                   '"resourceAccess":[{"id":"3b5f3d61-589b-4a3c-a359-5dd4b5ee5bd5","type":"Scope"}]},' +  # EWS.AccessAsUser.All
                  '{"resourceAppId":"00000003-0000-0000-c000-000000000000",' +   # Microsoft Graph
                   '"resourceAccess":[' +
                     '{"id":"e1fe6dd8-ba31-4d61-89e7-88639da4683d","type":"Scope"},' +  # User.Read
                     '{"id":"7427e0e9-2fba-42fe-b0c0-848c9e6a8182","type":"Scope"}' +   # offline_access
                   ']}' +
                ']}')
    $appBytes = [System.Text.Encoding]::UTF8.GetBytes($appBody)
    $newId    = $null
    try {
        $rq3 = [System.Net.HttpWebRequest]::Create("https://graph.microsoft.com/v1.0/applications")
        $rq3.Method = "POST"; $rq3.ContentType = "application/json"
        $rq3.ContentLength = $appBytes.Length; $rq3.Timeout = 20000
        $rq3.Headers["Authorization"] = "Bearer $adminToken"
        $ss3 = $rq3.GetRequestStream(); $ss3.Write($appBytes, 0, $appBytes.Length); $ss3.Close()
        $rp3  = $rq3.GetResponse()
        $aJ   = (New-Object System.IO.StreamReader($rp3.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
        $rp3.Close()
        $newId = $aJ.appId
    } catch [System.Net.WebException] {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $exT3 = $_.Exception
        $m3 = if ($exT3.Response) {
            try { $e3=(New-Object System.IO.StreamReader($exT3.Response.GetResponseStream())).ReadToEnd()|ConvertFrom-Json; $exT3.Response.Close(); "$($e3.error.code): $($e3.error.message)" } catch { $exT3.Message }
        } else { $exT3.Message }
        [System.Windows.Forms.MessageBox]::Show($m3, "App Registration Error",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    } catch {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }
    $form.Cursor = [System.Windows.Forms.Cursors]::Default

    # ── Service principal + admin consent ─────────────────────────────────────
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()

    $spId      = $null
    $consentOk = $false
    $consentErr = ''

    # Helper: GET from Graph → parsed JSON
    $graphGet = {
        param([string]$url)
        $rq = [System.Net.HttpWebRequest]::Create($url)
        $rq.Method = "GET"; $rq.Timeout = 12000
        $rq.Headers["Authorization"] = "Bearer $adminToken"
        $rp = $rq.GetResponse()
        $j = (New-Object System.IO.StreamReader($rp.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
        $rp.Close(); return $j
    }
    # Helper: POST JSON to Graph → parsed JSON (or $null on error, sets $consentErr)
    $graphPost = {
        param([string]$url, [string]$body)
        $b = [System.Text.Encoding]::UTF8.GetBytes($body)
        $rq = [System.Net.HttpWebRequest]::Create($url)
        $rq.Method = "POST"; $rq.ContentType = "application/json"
        $rq.ContentLength = $b.Length; $rq.Timeout = 20000
        $rq.Headers["Authorization"] = "Bearer $adminToken"
        $ss = $rq.GetRequestStream(); $ss.Write($b, 0, $b.Length); $ss.Close()
        try {
            $rp = $rq.GetResponse()
            $j = (New-Object System.IO.StreamReader($rp.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
            $rp.Close(); return $j
        } catch [System.Net.WebException] {
            $ex = $_.Exception
            if ($ex.Response -and [int]$ex.Response.StatusCode -eq 409) {
                $ex.Response.Close(); return 'conflict'   # signal: already exists
            }
            $m = if ($ex.Response) {
                try { $ej=(New-Object System.IO.StreamReader($ex.Response.GetResponseStream())).ReadToEnd()|ConvertFrom-Json; $ex.Response.Close()
                      "$($ej.error.code): $($ej.error.message)" } catch { $ex.Message }
            } else { $ex.Message }
            throw $m
        }
    }

    try {
        # 1. Create service principal for the new app
        $spRes = & $graphPost "https://graph.microsoft.com/v1.0/servicePrincipals" "{`"appId`":`"$newId`"}"
        if ($spRes -eq 'conflict') {
            # Already exists — look it up
            $spRes = (& $graphGet "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId+eq+'$newId'&`$select=id").value[0]
        }
        $spId = $spRes.id

        # 2. Resolve resource service principal IDs (Exchange Online + Microsoft Graph)
        $exoSp = (& $graphGet "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId+eq+'00000002-0000-0ff1-ce00-000000000000'&`$select=id").value[0].id
        $grSp  = (& $graphGet "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId+eq+'00000003-0000-0000-c000-000000000000'&`$select=id").value[0].id

        # 3. Grant admin consent (AllPrincipals = tenant-wide)
        foreach ($grant in @(
            @{ Res=$exoSp; Scope='EWS.AccessAsUser.All' },
            @{ Res=$grSp;  Scope='User.Read offline_access' }
        )) {
            $gBody = "{`"clientId`":`"$spId`",`"consentType`":`"AllPrincipals`",`"resourceId`":`"$($grant.Res)`",`"scope`":`"$($grant.Scope)`"}"
            & $graphPost "https://graph.microsoft.com/v1.0/oauth2PermissionGrants" $gBody | Out-Null
        }
        $consentOk = $true
    } catch {
        $consentErr = $_.Exception.Message
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::Default

    # ── Save config and update UI ─────────────────────────────────────────────
    try {
        [System.IO.File]::WriteAllText($script:configPath,
            (ConvertTo-Json @{ ClientId = $newId }),
            [System.Text.Encoding]::UTF8)
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "App created but config could not be saved:`n$($_.Exception.Message)",
            "Warning", [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    }

    $txtClientId.Text = $newId

    $summaryMsg = if ($consentOk) {
        "App registration 'Exchange Tester' created successfully.`n`nClient ID:`n$newId`n`n" +
        "API permissions granted with tenant-wide admin consent:`n" +
        "  • EWS.AccessAsUser.All`n  • User.Read`n  • offline_access`n`n" +
        "Saved to ExchangeTester.config."
    } else {
        "App registration 'Exchange Tester' created successfully.`n`nClient ID:`n$newId`n`n" +
        "Saved to ExchangeTester.config.`n`n" +
        "Note: admin consent could not be granted automatically:`n$consentErr`n`n" +
        "Please grant consent manually in the Azure portal."
    }
    [System.Windows.Forms.MessageBox]::Show(
        $summaryMsg, "App Registration Created",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
})

$btnTest.Add_Click({
    $email = $txtEmail.Text.Trim()
    if ($email -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a valid e-mail address.",
            "Input Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    if ($radModernAuth.Checked -and $radACF.Checked) {
        if (-not $txtClientId.Text.Trim()) {
            [System.Windows.Forms.MessageBox]::Show(
                "Please enter or register a Client ID first.`nUse the 'Register App' button to create an app registration automatically.",
                "Client ID Required",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }
        if (-not $txtTenantId.Text.Trim()) {
            [System.Windows.Forms.MessageBox]::Show(
                "Please enter or detect the Tenant ID first.",
                "Tenant ID Required",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }
    }


    # Certificate validation callback
    if ($chkIgnoreCert.Checked) {
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    } else {
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null
    }

    # Reset UI
    $form.Text = "Test E-Mail AutoConfiguration  —  $email"
    $rtbLog.Clear()
    $rtbXml.Clear()
    $lvwResults.Groups.Clear()
    $lvwResults.Items.Clear()
    $prgBar.Value        = 0
    $btnTest.Enabled     = $false
    $btnCancel.Enabled   = $true
    $btnAddTests.Enabled = $false
    $tabCtrl.SelectedTab = $tabLog

    # Build the sync hash that bridges UI thread and test runspace
    $sync = [hashtable]::Synchronized(@{
        Email          = $email
        Password       = $txtPass.Text
        UseWindowsAuth = $radWIA.Checked -and $chkUseCurrentUser.Checked
        ModernAuth     = $radModernAuth.Checked
        UseDeviceCode  = $radDCF.Checked
        ClientId       = $txtClientId.Text.Trim()
        TenantId       = $txtTenantId.Text.Trim()
        UseSCP         = $chkUseSCP.Checked
        Cancel          = $false
        Done            = $false
        Xml             = $null
        Pct             = 0
        LastToken       = $null
        LastTokenExpiry = $null
        Queue           = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
    })
    $script:CurrentSync = $sync

    # Create a dedicated runspace so PS script blocks work without issues
    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = [System.Threading.ApartmentState]::STA
    $rs.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::UseNewThread
    $rs.Open()
    $script:CurrentRS = $rs

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript($script:TestScript).AddArgument($sync)
    $script:CurrentPS = $ps
    [void]$ps.BeginInvoke()

    # Poll timer: drain log queue and check for completion (runs on UI thread)
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 100
    $timer.Add_Tick({
        $msg = $null
        while ($script:CurrentSync -and $script:CurrentSync.Queue.TryDequeue([ref]$msg)) {
            $rtbLog.AppendText("$msg`r`n")
            $rtbLog.ScrollToCaret()
        }
        if ($script:CurrentSync) {
            $v = $script:CurrentSync.Pct
            if ($v -ge 0 -and $v -le 100) { $prgBar.Value = $v }
        }
        if ($script:CurrentSync -and $script:CurrentSync.Done) {
            Complete-Test
        }
    })
    $script:PollTimer = $timer
    $timer.Start()
})

$btnCancel.Add_Click({
    if ($script:CurrentSync) { $script:CurrentSync.Cancel = $true }
    $btnCancel.Enabled = $false
    $rtbLog.AppendText("Cancelling...`r`n")
})

$btnAddTests.Add_Click({
    if (-not $script:LastXml) { return }

    $urlList = Get-AutodiscoverUrls -RawXml $script:LastXml
    if ($urlList.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No testable URLs found in the AutoDiscover response.",
            "Additional Tests",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }

    $ignoreCert = $chkIgnoreCert.Checked

    # ── Determine auth for URL probing ────────────────────────────────────────
    $authHeader = $null   # "Bearer <token>" for Modern Auth
    $useWinAuth = $false  # UseDefaultCredentials (current Windows user)
    $netCred    = $null   # explicit NetworkCredential

    if ($radModernAuth.Checked) {
        $maScope = 'https://outlook.office365.com/EWS.AccessAsUser.All offline_access'

        # Reuse cached token from the most recent successful test if still valid
        if ($script:LastToken -and $script:LastTokenExpiry -and (Get-Date) -lt $script:LastTokenExpiry) {
            $authHeader = $script:LastToken
        } else {
            # Token absent or expired — re-acquire
            if ($radDCF.Checked) {
                # ── Device Code Flow ─────────────────────────────────────────
                $dcClientId2 = if ($txtClientId.Text.Trim()) { $txtClientId.Text.Trim() } else { 'd3590ed6-52b3-4102-aeff-aad2292ab01c' }
                $dcUrl2 = 'https://login.microsoftonline.com/common/oauth2/v2.0/devicecode'
                $dtUrl2 = 'https://login.microsoftonline.com/common/oauth2/v2.0/token'
                $dcBytes2 = [System.Text.Encoding]::UTF8.GetBytes(
                    "client_id=$([Uri]::EscapeDataString($dcClientId2))&scope=$([Uri]::EscapeDataString($maScope))")
                $dcJson2 = $null
                try {
                    $rqD = [System.Net.HttpWebRequest]::Create($dcUrl2)
                    $rqD.Method = "POST"; $rqD.ContentType = "application/x-www-form-urlencoded"
                    $rqD.ContentLength = $dcBytes2.Length; $rqD.Timeout = 15000
                    $ssD = $rqD.GetRequestStream(); $ssD.Write($dcBytes2, 0, $dcBytes2.Length); $ssD.Close()
                    $rpD = $rqD.GetResponse()
                    $dcJson2 = (New-Object System.IO.StreamReader($rpD.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
                    $rpD.Close()
                } catch {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Device code request failed:`n$($_.Exception.Message)", "Auth Error",
                        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
                    return
                }
                $dcUserCode2   = $dcJson2.user_code
                $dcDevCode2    = $dcJson2.device_code
                $dcVerify2     = if ($dcJson2.verification_uri) { $dcJson2.verification_uri } else { $dcJson2.verification_url }
                $dcPoll2       = [int]$dcJson2.interval; if ($dcPoll2 -lt 5) { $dcPoll2 = 5 }
                # Use hashtable to share state with timer/button closures (plain variables don't mutate across scriptblock scopes)
                $dcState = @{ Tok = $null; Err = $null; Cancel = $false }

                $dcF2 = New-Object System.Windows.Forms.Form
                $dcF2.Text            = "Sign in to Microsoft  —  Additional Tests"
                $dcF2.Size            = New-Object System.Drawing.Size(440, 210)
                $dcF2.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
                $dcF2.MinimizeBox     = $false
                $dcF2.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
                $dL1 = New-Object System.Windows.Forms.Label;    $dL1.Text = "1.  Open a browser and go to:"; $dL1.Location = New-Object System.Drawing.Point(12,14); $dL1.AutoSize = $true
                $dLnk = New-Object System.Windows.Forms.LinkLabel; $dLnk.Text = $dcVerify2; $dLnk.Location = New-Object System.Drawing.Point(28,34); $dLnk.AutoSize = $true
                $dLnk.Add_LinkClicked({ [System.Diagnostics.Process]::Start($dLnk.Text) })
                $dL2 = New-Object System.Windows.Forms.Label;    $dL2.Text = "2.  Enter this code:"; $dL2.Location = New-Object System.Drawing.Point(12,62); $dL2.AutoSize = $true
                $dCodeLbl = New-Object System.Windows.Forms.Label; $dCodeLbl.Text = $dcUserCode2
                $dCodeLbl.Font = New-Object System.Drawing.Font("Consolas",22,[System.Drawing.FontStyle]::Bold)
                $dCodeLbl.Location = New-Object System.Drawing.Point(28,80); $dCodeLbl.AutoSize = $true; $dCodeLbl.ForeColor = [System.Drawing.Color]::DarkBlue
                $dCopy = New-Object System.Windows.Forms.Button; $dCopy.Text = "Copy"; $dCopy.Location = New-Object System.Drawing.Point(340,82); $dCopy.Size = New-Object System.Drawing.Size(72,26)
                $dCopy.Add_Click({ [System.Windows.Forms.Clipboard]::SetText($dcUserCode2) })
                $dWait = New-Object System.Windows.Forms.Label;  $dWait.Text = "Waiting for sign-in…"; $dWait.Location = New-Object System.Drawing.Point(12,144); $dWait.AutoSize = $true; $dWait.ForeColor = [System.Drawing.Color]::Gray
                $dCancelBtn = New-Object System.Windows.Forms.Button; $dCancelBtn.Text = "Cancel"; $dCancelBtn.Location = New-Object System.Drawing.Point(340,140); $dCancelBtn.Size = New-Object System.Drawing.Size(72,26)
                $dCancelBtn.Add_Click({ $dcState.Cancel = $true; $dcF2.Close() })
                $dcF2.Controls.AddRange(@($dL1,$dLnk,$dL2,$dCodeLbl,$dCopy,$dWait,$dCancelBtn))

                $dpTimer = New-Object System.Windows.Forms.Timer; $dpTimer.Interval = $dcPoll2 * 1000
                $dpTimer.Add_Tick({
                    if ($dcState.Tok -or $dcState.Err -or $dcState.Cancel) { return }
                    $pb2 = [System.Text.Encoding]::UTF8.GetBytes(
                        "grant_type=urn:ietf:params:oauth:grant-type:device_code" +
                        "&client_id=$([Uri]::EscapeDataString($dcClientId2))" +
                        "&device_code=$([Uri]::EscapeDataString($dcDevCode2))")
                    try {
                        $rq2 = [System.Net.HttpWebRequest]::Create($dtUrl2)
                        $rq2.Method = "POST"; $rq2.ContentType = "application/x-www-form-urlencoded"
                        $rq2.ContentLength = $pb2.Length; $rq2.Timeout = 4000
                        $ss2 = $rq2.GetRequestStream(); $ss2.Write($pb2, 0, $pb2.Length); $ss2.Close()
                        try {
                            $rp2  = $rq2.GetResponse()
                            $tj2  = (New-Object System.IO.StreamReader($rp2.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
                            $rp2.Close()
                            if ($tj2.access_token) { $dcState.Tok = "Bearer $($tj2.access_token)"; $dcF2.Close() }
                        } catch [System.Net.WebException] {
                            $ex2 = $_.Exception
                            if ($ex2.Response) {
                                $ej2 = (New-Object System.IO.StreamReader($ex2.Response.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
                                $ex2.Response.Close()
                                switch ($ej2.error) {
                                    'authorization_pending' {}
                                    'slow_down'             { $dpTimer.Interval += 5000 }
                                    default                 { $dcState.Err = "$($ej2.error): $($ej2.error_description)"; $dcF2.Close() }
                                }
                            }
                        }
                    } catch {}
                })
                $dcF2.Add_Shown({ $dpTimer.Start() })
                $dcF2.Add_FormClosed({ $dpTimer.Stop() })
                [void]$dcF2.ShowDialog()
                $dpTimer.Dispose(); $dcF2.Dispose()

                if ($dcState.Cancel -or (-not $dcState.Tok -and -not $dcState.Err)) { return }
                if ($dcState.Err) {
                    [System.Windows.Forms.MessageBox]::Show("Sign-in error: $($dcState.Err)", "Auth Error",
                        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
                    return
                }
                $authHeader             = $dcState.Tok
                $script:LastToken       = $authHeader
                $script:LastTokenExpiry = (Get-Date).AddMinutes(55)

            } else {
                # ── Auth Code Flow + PKCE ────────────────────────────────────
                $acfCid = $txtClientId.Text.Trim()
                $acfTid = $txtTenantId.Text.Trim()
                if (-not $acfCid -or -not $acfTid) {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Client ID and Tenant ID are required for Auth Code Flow.",
                        "Config Required", [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                    return
                }
                $rng3 = [System.Security.Cryptography.RNGCryptoServiceProvider]::new()
                $rngB3 = [byte[]]::new(48); $rng3.GetBytes($rngB3); $rng3.Dispose()
                $ver3  = [Convert]::ToBase64String($rngB3) -replace '\+','-' -replace '/','_' -replace '=',''
                $sha3  = [System.Security.Cryptography.SHA256Managed]::new()
                $chal3 = [Convert]::ToBase64String($sha3.ComputeHash([System.Text.Encoding]::ASCII.GetBytes($ver3))) -replace '\+','-' -replace '/','_' -replace '=',''
                $sha3.Dispose()
                $tcpT3 = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
                $tcpT3.Start(); $port3 = $tcpT3.LocalEndpoint.Port; $tcpT3.Stop()
                $redir3   = "http://localhost:$port3/"
                $state3   = [Guid]::NewGuid().ToString('N')
                $tokUrl3  = "https://login.microsoftonline.com/$([Uri]::EscapeDataString($acfTid))/oauth2/v2.0/token"
                $authUrl3 = "https://login.microsoftonline.com/$([Uri]::EscapeDataString($acfTid))/oauth2/v2.0/authorize" +
                    "?client_id=$([Uri]::EscapeDataString($acfCid))" +
                    "&response_type=code&redirect_uri=$([Uri]::EscapeDataString($redir3))" +
                    "&scope=$([Uri]::EscapeDataString($maScope))&state=$([Uri]::EscapeDataString($state3))" +
                    "&code_challenge=$([Uri]::EscapeDataString($chal3))&code_challenge_method=S256&prompt=select_account"

                $a3Sync = [hashtable]::Synchronized(@{ Code=$null; Err=$null; Cancel=$false; Done=$false; DotCount=0 })
                $a3Rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
                $a3Rs.ApartmentState = [System.Threading.ApartmentState]::STA
                $a3Rs.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
                $a3Rs.Open()
                $a3Rs.SessionStateProxy.SetVariable('a3Sync',  $a3Sync)
                $a3Rs.SessionStateProxy.SetVariable('port3',   $port3)
                $a3Rs.SessionStateProxy.SetVariable('state3',  $state3)
                $a3Ps = [System.Management.Automation.PowerShell]::Create(); $a3Ps.Runspace = $a3Rs
                [void]$a3Ps.AddScript({
                    $hl3 = [System.Net.HttpListener]::new(); $hl3.Prefixes.Add("http://localhost:$port3/"); $hl3.Start()
                    try {
                        $ar3 = $hl3.BeginGetContext($null,$null); $dl3 = (Get-Date).AddMinutes(5)
                        while (-not $ar3.IsCompleted -and (Get-Date) -lt $dl3) {
                            if ($a3Sync.Cancel) { $hl3.Stop(); return }
                            [System.Threading.Thread]::Sleep(200)
                        }
                        if (-not $ar3.IsCompleted) { $hl3.Stop(); if (-not $a3Sync.Cancel) { $a3Sync.Err = 'Timed out' }; return }
                        $ctx3 = $hl3.EndGetContext($ar3); $qs3 = $ctx3.Request.QueryString
                        $code3 = $qs3['code']; $ok3 = $code3 -and $qs3['state'] -eq $state3
                        $htm3 = if ($ok3) {'<html><body style="font-family:sans-serif;padding:40px"><h2 style="color:green">&#10003; Signed in</h2><p>You may close this tab.</p></body></html>'} else {'<html><body style="font-family:sans-serif;padding:40px"><h2 style="color:red">&#10007; Sign-in failed</h2></body></html>'}
                        $hb3 = [System.Text.Encoding]::UTF8.GetBytes($htm3)
                        $ctx3.Response.ContentType = 'text/html; charset=utf-8'; $ctx3.Response.ContentLength64 = $hb3.Length
                        $ctx3.Response.OutputStream.Write($hb3,0,$hb3.Length); $ctx3.Response.OutputStream.Close(); $ctx3.Response.Close()
                        if ($ok3) { $a3Sync.Code = $code3 } else { $a3Sync.Err = $qs3['error'] }
                    } catch { if (-not $a3Sync.Cancel) { $a3Sync.Err = $_.Exception.Message } }
                    finally { try { $hl3.Stop(); $hl3.Close() } catch {}; $a3Sync.Done = $true }
                })
                [void]$a3Ps.BeginInvoke()
                [System.Diagnostics.Process]::Start($authUrl3) | Out-Null

                $a3F = New-Object System.Windows.Forms.Form
                $a3F.Text = "Sign in — Additional Tests"; $a3F.Size = New-Object System.Drawing.Size(480,155)
                $a3F.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
                $a3F.MinimizeBox = $false; $a3F.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
                $a3L0 = New-Object System.Windows.Forms.Label; $a3L0.Text = "Sign in to Microsoft in your browser."
                $a3L0.Location = New-Object System.Drawing.Point(12,12); $a3L0.Size = New-Object System.Drawing.Size(450,18)
                $a3L0.Font = New-Object System.Drawing.Font($a3F.Font,[System.Drawing.FontStyle]::Bold)
                $a3LW = New-Object System.Windows.Forms.Label; $a3LW.Text = "Waiting for authentication…"
                $a3LW.Location = New-Object System.Drawing.Point(12,38); $a3LW.Size = New-Object System.Drawing.Size(450,18); $a3LW.ForeColor = [System.Drawing.Color]::Gray
                $a3Dots = New-Object System.Windows.Forms.Label; $a3Dots.Text = ""; $a3Dots.Location = New-Object System.Drawing.Point(12,62); $a3Dots.AutoSize = $true; $a3Dots.ForeColor = [System.Drawing.Color]::SteelBlue
                $a3CancelBtn = New-Object System.Windows.Forms.Button; $a3CancelBtn.Text = "Cancel"; $a3CancelBtn.Location = New-Object System.Drawing.Point(390,90); $a3CancelBtn.Size = New-Object System.Drawing.Size(72,26)
                $a3CancelBtn.Add_Click({ $a3Sync.Cancel = $true; $a3F.Close() })
                $a3F.Controls.AddRange(@($a3L0,$a3LW,$a3Dots,$a3CancelBtn))
                $a3Poll = New-Object System.Windows.Forms.Timer; $a3Poll.Interval = 300
                $a3Poll.Add_Tick({
                    if ($a3Sync.Done -or $a3Sync.Cancel) { $a3F.Close(); return }
                    $a3Sync.DotCount = (($a3Sync.DotCount + 1) % 6); $a3Dots.Text = '.' * ($a3Sync.DotCount + 1)
                })
                $a3F.Add_Shown({ $a3Sync.DotCount = 0; $a3Poll.Start() })
                $a3F.Add_FormClosed({ $a3Poll.Stop() })
                [void]$a3F.ShowDialog()
                $a3Poll.Dispose(); $a3F.Dispose()
                try { $a3Ps.Stop() } catch {}; $a3Ps.Dispose(); $a3Rs.Close(); $a3Rs.Dispose()

                if ($a3Sync.Cancel -or (-not $a3Sync.Code -and -not $a3Sync.Err)) { return }
                if ($a3Sync.Err) {
                    [System.Windows.Forms.MessageBox]::Show("Sign-in error: $($a3Sync.Err)", "Auth Error",
                        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
                    return
                }
                $a3TokBody  = "grant_type=authorization_code&client_id=$([Uri]::EscapeDataString($acfCid))" +
                    "&code=$([Uri]::EscapeDataString($a3Sync.Code))&redirect_uri=$([Uri]::EscapeDataString($redir3))" +
                    "&code_verifier=$([Uri]::EscapeDataString($ver3))"
                $a3TokBytes = [System.Text.Encoding]::UTF8.GetBytes($a3TokBody)
                $a3Tok = $null
                try {
                    $a3Rq = [System.Net.HttpWebRequest]::Create($tokUrl3)
                    $a3Rq.Method = "POST"; $a3Rq.ContentType = "application/x-www-form-urlencoded"
                    $a3Rq.ContentLength = $a3TokBytes.Length; $a3Rq.Timeout = 15000
                    $a3Ss = $a3Rq.GetRequestStream(); $a3Ss.Write($a3TokBytes,0,$a3TokBytes.Length); $a3Ss.Close()
                    $a3Rp = $a3Rq.GetResponse()
                    $a3TJ = (New-Object System.IO.StreamReader($a3Rp.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
                    $a3Rp.Close(); $a3Tok = $a3TJ.access_token
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Token exchange failed: $($_.Exception.Message)", "Auth Error",
                        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
                    return
                }
                if (-not $a3Tok) {
                    [System.Windows.Forms.MessageBox]::Show("Token exchange returned no access token.", "Auth Error",
                        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
                    return
                }
                $authHeader             = "Bearer $a3Tok"
                $script:LastToken       = $authHeader
                $script:LastTokenExpiry = (Get-Date).AddMinutes(55)
            }
        }
    } elseif ($radWIA.Checked) {
        if (-not $chkUseCurrentUser.Checked -and $txtPass.Text) {
            $netCred = New-Object System.Net.NetworkCredential($txtEmail.Text.Trim(), $txtPass.Text)
        } else {
            $useWinAuth = $true
        }
    }

    Show-EndpointTestDialog -Title "Additional Tests  —  URL Connectivity" -Entries $urlList -AuthHeader $authHeader -UseWinAuth $useWinAuth -NetCred $netCred -IgnoreCert $ignoreCert -Owner $form -CsvName 'additional-tests.csv'
})

$form.Add_FormClosing({
    if ($script:CurrentSync) { $script:CurrentSync.Cancel = $true }
    if ($script:PollTimer)   { $script:PollTimer.Stop() }
    try { $script:CurrentPS.Dispose() }   catch {}
    try { $script:CurrentRS.Close();  $script:CurrentRS.Dispose() } catch {}
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null
})

#region ======================================================================
#  HYBRID DEPLOYMENT TEST  (Migration Endpoint / MRS Proxy)
#==============================================================================

$hybForm = New-Object System.Windows.Forms.Form
$hybForm.Text            = "Hybrid Deployment  —  Migration Endpoint (MRS Proxy)"
$hybForm.ClientSize      = New-Object System.Drawing.Size(775, 595)
$hybForm.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
$hybForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$hybForm.MaximizeBox     = $true
$hybForm.MinimizeBox     = $true
$hybForm.MinimumSize     = New-Object System.Drawing.Size(640, 420)

#region --- Row 1: On-prem mailbox ---
$hybLblEmail = New-Object System.Windows.Forms.Label
$hybLblEmail.Text      = "On-Prem Mailbox"
$hybLblEmail.Location  = New-Object System.Drawing.Point(8, 14)
$hybLblEmail.Size      = New-Object System.Drawing.Size(128, 20)
$hybLblEmail.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$hybForm.Controls.Add($hybLblEmail)

$hybTxtEmail = New-Object System.Windows.Forms.TextBox
$hybTxtEmail.Location = New-Object System.Drawing.Point(140, 11)
$hybTxtEmail.Size     = New-Object System.Drawing.Size(622, 22)
$hybTxtEmail.TabIndex = 0
$hybTxtEmail.Anchor   = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right)
$hybForm.Controls.Add($hybTxtEmail)
#endregion

#region --- Row 2: MRS endpoint FQDN ---
$hybLblFqdn = New-Object System.Windows.Forms.Label
$hybLblFqdn.Text      = "MRS Endpoint (FQDN)"
$hybLblFqdn.Location  = New-Object System.Drawing.Point(8, 42)
$hybLblFqdn.Size      = New-Object System.Drawing.Size(128, 20)
$hybLblFqdn.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$hybForm.Controls.Add($hybLblFqdn)

$hybTxtFqdn = New-Object System.Windows.Forms.TextBox
$hybTxtFqdn.Location = New-Object System.Drawing.Point(140, 39)
$hybTxtFqdn.Size     = New-Object System.Drawing.Size(340, 22)
$hybTxtFqdn.TabIndex = 1
$hybForm.Controls.Add($hybTxtFqdn)

$hybLblFqdnHint = New-Object System.Windows.Forms.Label
$hybLblFqdnHint.Text      = "(leave empty to discover via AutoDiscover)"
$hybLblFqdnHint.Location  = New-Object System.Drawing.Point(486, 42)
$hybLblFqdnHint.Size      = New-Object System.Drawing.Size(276, 20)
$hybLblFqdnHint.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$hybLblFqdnHint.ForeColor = [System.Drawing.Color]::Gray
$hybLblFqdnHint.Anchor    = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)
$hybForm.Controls.Add($hybLblFqdnHint)
#endregion

#region --- Row 3: On-prem credentials ---
$hybLblUser = New-Object System.Windows.Forms.Label
$hybLblUser.Text      = "On-Prem User"
$hybLblUser.Location  = New-Object System.Drawing.Point(8, 70)
$hybLblUser.Size      = New-Object System.Drawing.Size(128, 20)
$hybLblUser.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$hybForm.Controls.Add($hybLblUser)

$hybTxtUser = New-Object System.Windows.Forms.TextBox
$hybTxtUser.Location = New-Object System.Drawing.Point(140, 67)
$hybTxtUser.Size     = New-Object System.Drawing.Size(240, 22)
$hybTxtUser.TabIndex = 2
$hybForm.Controls.Add($hybTxtUser)

$hybLblPass = New-Object System.Windows.Forms.Label
$hybLblPass.Text      = "Password"
$hybLblPass.Location  = New-Object System.Drawing.Point(390, 70)
$hybLblPass.Size      = New-Object System.Drawing.Size(70, 20)
$hybLblPass.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$hybForm.Controls.Add($hybLblPass)

$hybTxtPass = New-Object System.Windows.Forms.TextBox
$hybTxtPass.Location     = New-Object System.Drawing.Point(466, 67)
$hybTxtPass.Size         = New-Object System.Drawing.Size(180, 22)
$hybTxtPass.PasswordChar = [char]0x25CF   # ●
$hybTxtPass.TabIndex     = 3
$hybForm.Controls.Add($hybTxtPass)
#endregion

#region --- Row 4: options + buttons ---
$hybChkIgnoreCert = New-Object System.Windows.Forms.CheckBox
$hybChkIgnoreCert.Text     = "Ignore certificate errors"
$hybChkIgnoreCert.Location = New-Object System.Drawing.Point(8, 96)
$hybChkIgnoreCert.Size     = New-Object System.Drawing.Size(200, 20)
$hybChkIgnoreCert.Checked  = $false
$hybChkIgnoreCert.TabIndex = 4
$hybForm.Controls.Add($hybChkIgnoreCert)

$hybChkExo = New-Object System.Windows.Forms.CheckBox
$hybChkExo.Text     = "Verify from Exchange Online"
$hybChkExo.Location = New-Object System.Drawing.Point(240, 96)
$hybChkExo.Size     = New-Object System.Drawing.Size(370, 20)
$hybChkExo.Checked  = $false
$hybChkExo.TabIndex = 5
$hybForm.Controls.Add($hybChkExo)

$hybBtnTest = New-Object System.Windows.Forms.Button
$hybBtnTest.Text     = "Test"
$hybBtnTest.Location = New-Object System.Drawing.Point(620, 92)
$hybBtnTest.Size     = New-Object System.Drawing.Size(68, 26)
$hybBtnTest.TabIndex = 6
$hybBtnTest.Anchor   = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)
$hybForm.Controls.Add($hybBtnTest)
$hybForm.AcceptButton = $hybBtnTest

$hybBtnCancel = New-Object System.Windows.Forms.Button
$hybBtnCancel.Text     = "Cancel"
$hybBtnCancel.Location = New-Object System.Drawing.Point(696, 92)
$hybBtnCancel.Size     = New-Object System.Drawing.Size(68, 26)
$hybBtnCancel.Enabled  = $false
$hybBtnCancel.TabIndex = 7
$hybBtnCancel.Anchor   = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)
$hybForm.Controls.Add($hybBtnCancel)

$toolTip.SetToolTip($hybTxtEmail, "Primary SMTP address of an on-premises mailbox. Used for AutoDiscover-based endpoint discovery when the FQDN field is empty.")
$toolTip.SetToolTip($hybTxtFqdn,  "External FQDN of the on-prem MRS endpoint, e.g. mail.contoso.com (the host serving /EWS/mrsproxy.svc). Leave empty to discover it via AutoDiscover.")
$toolTip.SetToolTip($hybTxtUser,  "On-premises account, DOMAIN\user or UPN. Leave empty to use the logged-in Windows user (local probes only).")
$toolTip.SetToolTip($hybChkExo,   "Additionally run Test-MigrationServerAvailability inside Exchange Online (AdminAPI, device code sign-in as an Exchange admin). This is the only check that proves Microsoft's datacenter can reach the endpoint.")
#endregion

#region --- Separator + progress bar + tabs ---
$hybPnlSep = New-Object System.Windows.Forms.Panel
$hybPnlSep.Location  = New-Object System.Drawing.Point(0, 126)
$hybPnlSep.Size      = New-Object System.Drawing.Size(775, 2)
$hybPnlSep.BackColor = [System.Drawing.SystemColors]::ControlDark
$hybPnlSep.Anchor    = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right)
$hybForm.Controls.Add($hybPnlSep)

$hybPrgBar = New-Object System.Windows.Forms.ProgressBar
$hybPrgBar.Location = New-Object System.Drawing.Point(8, 134)
$hybPrgBar.Size     = New-Object System.Drawing.Size(757, 14)
$hybPrgBar.Minimum  = 0
$hybPrgBar.Maximum  = 100
$hybPrgBar.Anchor   = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right)
$hybForm.Controls.Add($hybPrgBar)

$hybTabCtrl = New-Object System.Windows.Forms.TabControl
$hybTabCtrl.Location = New-Object System.Drawing.Point(8, 157)
$hybTabCtrl.Size     = New-Object System.Drawing.Size(757, 429)
$hybTabCtrl.Anchor   = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right)
$hybForm.Controls.Add($hybTabCtrl)

$hybTabResults = New-Object System.Windows.Forms.TabPage
$hybTabResults.Text = "Results"
$hybTabCtrl.Controls.Add($hybTabResults)

$hybLvwResults = New-Object System.Windows.Forms.ListView
$hybLvwResults.Dock            = [System.Windows.Forms.DockStyle]::Fill
$hybLvwResults.View            = [System.Windows.Forms.View]::Details
$hybLvwResults.FullRowSelect   = $true
$hybLvwResults.GridLines       = $true
$hybLvwResults.ShowItemToolTips = $true
$hybLvwResults.HeaderStyle     = [System.Windows.Forms.ColumnHeaderStyle]::Nonclickable
[void]$hybLvwResults.Columns.Add("Step",    200)
[void]$hybLvwResults.Columns.Add("Result",   60)
[void]$hybLvwResults.Columns.Add("Details", 900)
$hybTabResults.Controls.Add($hybLvwResults)

# Double-click a row to read the full details (also copies them to clipboard)
$hybLvwResults.Add_DoubleClick({
    if ($hybLvwResults.SelectedItems.Count -gt 0) {
        $it  = $hybLvwResults.SelectedItems[0]
        $txt = "$($it.Text)  —  $($it.SubItems[1].Text)`r`n`r`n$($it.SubItems[2].Text)"
        try { [System.Windows.Forms.Clipboard]::SetText($txt) } catch {}
        [System.Windows.Forms.MessageBox]::Show($txt, "Step detail (copied to clipboard)",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    }
})

$hybTabLog = New-Object System.Windows.Forms.TabPage
$hybTabLog.Text = "Log"
$hybTabCtrl.Controls.Add($hybTabLog)

$hybRtbLog = New-Object System.Windows.Forms.RichTextBox
$hybRtbLog.Dock       = [System.Windows.Forms.DockStyle]::Fill
$hybRtbLog.ReadOnly   = $true
$hybRtbLog.Font       = New-Object System.Drawing.Font("Consolas", 9)
$hybRtbLog.BackColor  = [System.Drawing.Color]::White
$hybRtbLog.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
$hybRtbLog.WordWrap   = $true
$hybTabLog.Controls.Add($hybRtbLog)

$ctxHybLog = New-Object System.Windows.Forms.ContextMenuStrip
$miHybCopy = New-Object System.Windows.Forms.ToolStripMenuItem("Copy")
$miHybCopy.ShortcutKeyDisplayString = "Ctrl+C"
$miHybSelAll = New-Object System.Windows.Forms.ToolStripMenuItem("Select All")
$miHybSelAll.ShortcutKeyDisplayString = "Ctrl+A"
$miHybClear = New-Object System.Windows.Forms.ToolStripMenuItem("Clear Log")
[void]$ctxHybLog.Items.Add($miHybCopy)
[void]$ctxHybLog.Items.Add($miHybSelAll)
[void]$ctxHybLog.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
[void]$ctxHybLog.Items.Add($miHybClear)
$hybRtbLog.ContextMenuStrip = $ctxHybLog

$miHybCopy.Add_Click({
    $t = $hybRtbLog.SelectedText
    if (-not $t) { $t = $hybRtbLog.Text }
    if ($t) { [System.Windows.Forms.Clipboard]::SetText($t) }
})
$miHybSelAll.Add_Click({ $hybRtbLog.SelectAll() })
$miHybClear.Add_Click({ $hybRtbLog.Clear() })
#endregion

# Script-level state for the currently running hybrid test
$script:HybPS        = $null
$script:HybRS        = $null
$script:HybSync      = $null
$script:HybPollTimer = $null

# The hybrid test logic runs inside a dedicated PS runspace (same pattern as
# the AutoDiscover test). $sync is the only bridge to the UI thread.
$script:HybridTestScript = {
    param($sync)

    $email  = $sync.Email
    $fqdn   = $sync.Fqdn
    $user   = $sync.User
    $pass   = $sync.Password
    $runExo = $sync.RunExo
    $domain = if ($email -match '@') { ($email -split '@')[1] } else { $null }

    $logLine = { param([string]$msg) $sync.Queue.Enqueue($msg) }
    $setPct  = { param([int]$pct)    $sync.Pct = $pct }
    $addRow  = {
        param([string]$step, [string]$result, [string]$details)
        if ($result -eq 'FAIL') { $sync.Failed = $true }
        $sync.RowQueue.Enqueue(@{ Step = $step; Result = $result; Details = $details })
    }

    [System.Net.ServicePointManager]::SecurityProtocol =
        [System.Net.SecurityProtocolType]::Tls12 -bor
        [System.Net.SecurityProtocolType]::Tls11 -bor
        [System.Net.SecurityProtocolType]::Tls
    if ($sync.IgnoreCert) {
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    }

    $netCred = $null
    if ($user) {
        # Split DOMAIN\user into (user, domain) so NTLM/Negotiate binds correctly;
        # a UPN (user@domain) is passed through as-is.
        if ($user -match '^([^\\]+)\\(.+)$') {
            $netCred = New-Object System.Net.NetworkCredential($Matches[2], $pass, $Matches[1])
        } else {
            $netCred = New-Object System.Net.NetworkCredential($user, $pass)
        }
    }

    # HTTP GET → @{Code; WwwAuth; Location; Body; Error}. $cred: NetworkCredential,
    # the string 'default' (logged-in Windows user) or $null (anonymous).
    $mrsGet = {
        param([string]$url, $cred)
        try {
            $req = [System.Net.HttpWebRequest]::Create($url)
            $req.Method            = "GET"
            $req.AllowAutoRedirect = $false
            $req.Timeout           = 20000
            $req.UserAgent         = "ExchangeMigrationTester/1.0"
            if ($cred -is [System.Net.CredentialCache]) {
                # Caller pre-selected the auth scheme(s) — use as-is.
                $req.Credentials = $cred
            }
            elseif ($cred -is [System.Net.NetworkCredential]) {
                # Bind the credential to both Windows auth schemes the endpoint
                # offers. A CredentialCache is the correct way to supply explicit
                # creds for these schemes.
                $cc = New-Object System.Net.CredentialCache
                $u  = New-Object System.Uri($url)
                $cc.Add($u, "Negotiate", $cred)
                $cc.Add($u, "NTLM",      $cred)
                $req.Credentials = $cc
            }
            elseif ($cred -eq 'default')                  { $req.UseDefaultCredentials = $true }
            try {
                $rp   = $req.GetResponse()
                $code = [int]$rp.StatusCode
                $rdr  = New-Object System.IO.StreamReader($rp.GetResponseStream())
                $body = $rdr.ReadToEnd(); $rdr.Close()
                $rp.Close()
                return @{ Code = $code; WwwAuth = $null; Location = $null; Body = $body; Error = $null }
            } catch [System.Net.WebException] {
                $ex = $_.Exception
                if ($ex.Response) {
                    $code = [int]$ex.Response.StatusCode
                    $wa   = try { ($ex.Response.Headers.GetValues("WWW-Authenticate")) -join ', ' } catch { $null }
                    $loc  = $ex.Response.Headers["Location"]
                    $ex.Response.Close()
                    return @{ Code = $code; WwwAuth = $wa; Location = $loc; Body = $null; Error = $null }
                }
                return @{ Code = -1; WwwAuth = $null; Location = $null; Body = $null; Error = $ex.Message }
            }
        } catch {
            return @{ Code = -1; WwwAuth = $null; Location = $null; Body = $null; Error = $_.Exception.Message }
        }
    }

    # Authenticated GET via WinHTTP. Unlike .NET's HttpWebRequest, WinHTTP
    # supplies the TLS channel-binding token, so it satisfies Extended
    # Protection for Authentication (EPA) on patched Exchange — where a .NET
    # NTLM/Negotiate probe would 401 despite correct credentials.
    $mrsGetWinHttp = {
        param([string]$url, [string]$user, [string]$pass, [bool]$ignoreCert)
        $wh = $null
        try {
            $wh = New-Object -ComObject 'WinHttp.WinHttpRequest.5.1'
            $wh.Open('GET', $url, $false)
            $wh.SetTimeouts(20000, 20000, 20000, 20000)
            if ($ignoreCert) { $wh.Option(4) = 13056 }   # SslErrorIgnoreFlags = 0x3300
            $wh.Option(6) = $false                        # EnableRedirects = false
            if ($user) {
                $wh.SetCredentials($user, $pass, 0)       # 0 = for server
            } else {
                $wh.SetAutoLogonPolicy(0)                 # 0 = Always: use logged-in Windows creds
            }
            $wh.Send()
            return @{ Code = [int]$wh.Status; Error = $null }
        } catch {
            $hr = $_.Exception.Message
            return @{ Code = -1; Error = "WinHTTP: $hr" }
        } finally {
            if ($wh) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wh) }
        }
    }

    # Authenticated SOAP POST to the MRS proxy (WinHTTP). Sends a minimal SOAP 1.2
    # envelope; the WCF service answers with a SOAP fault whose text tells us
    # definitively whether the MRS proxy is live/enabled. Returns Code + Body.
    $mrsSoapPost = {
        param([string]$url, [string]$user, [string]$pass, [bool]$ignoreCert)
        $wh = $null
        try {
            # SOAP 1.1 (text/xml) — the EWS/MRS family uses SOAP 1.1, not 1.2.
            $soap = '<?xml version="1.0" encoding="utf-8"?>' +
                    '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                    '<soap:Body/></soap:Envelope>'
            $wh = New-Object -ComObject 'WinHttp.WinHttpRequest.5.1'
            $wh.Open('POST', $url, $false)
            $wh.SetTimeouts(20000, 20000, 20000, 20000)
            if ($ignoreCert) { $wh.Option(4) = 13056 }
            $wh.Option(6) = $false
            $wh.SetRequestHeader('Content-Type', 'text/xml; charset=utf-8')
            $wh.SetRequestHeader('SOAPAction', '""')
            if ($user) { $wh.SetCredentials($user, $pass, 0) } else { $wh.SetAutoLogonPolicy(0) }
            $wh.Send($soap)
            $bodyTxt = try { "$($wh.ResponseText)" } catch { '' }
            return @{ Code = [int]$wh.Status; Body = $bodyTxt; Error = $null }
        } catch {
            return @{ Code = -1; Body = ''; Error = "WinHTTP: $($_.Exception.Message)" }
        } finally {
            if ($wh) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wh) }
        }
    }

    $abort = $false

    # ── Step 1: Endpoint discovery via AutoDiscover (when no FQDN given) ──────
    if (-not $fqdn -and -not $sync.Cancel) {
        & $setPct 5
        & $logLine "Endpoint discovery via AutoDiscover for $domain starting."

        $adXml   = "<?xml version=""1.0"" encoding=""utf-8""?>" +
            "<Autodiscover xmlns=""http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006"">" +
            "<Request><EMailAddress>$email</EMailAddress>" +
            "<AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema>" +
            "</Request></Autodiscover>"
        $adBytes = [System.Text.Encoding]::UTF8.GetBytes($adXml)

        $doAdPost = {
            param([string]$url)
            try {
                $req = [System.Net.HttpWebRequest]::Create($url)
                $req.Method            = "POST"
                $req.ContentType       = "text/xml; charset=utf-8"
                $req.ContentLength     = $adBytes.Length
                $req.AllowAutoRedirect = $false
                $req.Timeout           = 20000
                $req.UserAgent         = "Microsoft Office/16.0 (Windows NT 10.0)"
                if ($netCred) { $req.Credentials = $netCred } else { $req.UseDefaultCredentials = $true }
                $s = $req.GetRequestStream(); $s.Write($adBytes, 0, $adBytes.Length); $s.Close()
                try {
                    $resp = $req.GetResponse()
                    $code = [int]$resp.StatusCode
                    $body = $null
                    if ($code -eq 200) {
                        $rdr = New-Object System.IO.StreamReader($resp.GetResponseStream(), [System.Text.Encoding]::UTF8)
                        $body = $rdr.ReadToEnd(); $rdr.Close()
                    }
                    $loc = $resp.Headers["Location"]
                    $resp.Close()
                    return @{ Code = $code; Body = $body; Location = $loc; Error = $null }
                } catch [System.Net.WebException] {
                    $ex = $_.Exception
                    if ($ex.Response) {
                        $code = [int]$ex.Response.StatusCode
                        $loc  = $ex.Response.Headers["Location"]
                        $ex.Response.Close()
                        return @{ Code = $code; Body = $null; Location = $loc; Error = $null }
                    }
                    return @{ Code = -1; Body = $null; Location = $null; Error = $ex.Message }
                }
            } catch {
                return @{ Code = -1; Body = $null; Location = $null; Error = $_.Exception.Message }
            }
        }

        $adBody = $null
        foreach ($u in @("https://autodiscover.$domain/autodiscover/autodiscover.xml",
                         "https://$domain/autodiscover/autodiscover.xml")) {
            if ($sync.Cancel) { break }
            & $logLine "AutoDiscover via $u starting."
            $res = & $doAdPost $u
            if ($res.Code -ge 0) { & $logLine "GetLastError=0; httpStatus=$($res.Code)." }
            else                 { & $logLine "AutoDiscover via $u failed: $($res.Error)" }
            if ($res.Code -eq 200 -and $res.Body) { $adBody = $res.Body; break }
            if (($res.Code -eq 301 -or $res.Code -eq 302) -and $res.Location -match '^https://') {
                & $logLine "Redirect to $($res.Location)."
                $res2 = & $doAdPost $res.Location
                if ($res2.Code -ge 0) { & $logLine "GetLastError=0; httpStatus=$($res2.Code)." }
                if ($res2.Code -eq 200 -and $res2.Body) { $adBody = $res2.Body; break }
            }
        }

        if ($adBody) {
            try {
                $xd  = [xml]$adBody
                $nsm = New-Object System.Xml.XmlNamespaceManager($xd.NameTable)
                $nsm.AddNamespace("ad", "http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a")
                $ews = $null; $src = ''
                # Prefer the external (EXPR) EWS URL — that host is the MRS endpoint
                foreach ($want in @('EXPR', 'EXCH')) {
                    foreach ($proto in $xd.SelectNodes("//ad:Protocol", $nsm)) {
                        $tn = $proto.SelectSingleNode("ad:Type", $nsm)
                        if ($tn -and $tn.InnerText -eq $want) {
                            $n = $proto.SelectSingleNode("ad:EwsUrl", $nsm)
                            if ($n -and $n.InnerText) { $ews = $n.InnerText.Trim(); $src = $want; break }
                        }
                    }
                    if ($ews) { break }
                }
                if (-not $ews) {
                    $n = $xd.SelectSingleNode("//ad:EwsUrl", $nsm)
                    if ($n -and $n.InnerText) { $ews = $n.InnerText.Trim(); $src = 'first found' }
                }
                if ($ews) {
                    $fqdn = ([Uri]$ews).Host
                    & $logLine "EWS URL ($src): $ews"
                    & $addRow "Endpoint discovery" "OK" "$fqdn (from $src EWS URL: $ews)"
                }
            } catch {
                & $logLine "AutoDiscover XML parse error: $($_.Exception.Message)"
            }
        }
        if (-not $fqdn) {
            & $addRow "Endpoint discovery" "FAIL" "AutoDiscover returned no usable EWS URL — enter the MRS endpoint FQDN manually"
            $abort = $true
        }
    }

    # ── Step 2: DNS resolution ────────────────────────────────────────────────
    if (-not $abort -and -not $sync.Cancel) {
        & $setPct 25
        & $logLine ""
        & $logLine "DNS resolution for $fqdn starting."
        try {
            $ips = [System.Net.Dns]::GetHostAddresses($fqdn) | ForEach-Object { $_.IPAddressToString }
            & $logLine "DNS: $($ips -join ', ')"
            & $addRow "DNS resolution" "OK" ($ips -join ', ')
        } catch {
            & $addRow "DNS resolution" "FAIL" $_.Exception.Message
            $abort = $true
        }
    }

    # ── Step 3: TCP 443 ───────────────────────────────────────────────────────
    if (-not $abort -and -not $sync.Cancel) {
        & $setPct 35
        & $logLine "TCP connect to $fqdn`:443 starting."
        try {
            $tcp = New-Object System.Net.Sockets.TcpClient
            if (-not $tcp.ConnectAsync($fqdn, 443).Wait(10000)) {
                $tcp.Close()
                throw "Connection timed out (10 s)"
            }
            $tcp.Close()
            & $logLine "TCP 443: connected."
            & $addRow "TCP port 443" "OK" "Connected"
        } catch {
            $m = if ($_.Exception.InnerException) { $_.Exception.InnerException.Message } else { $_.Exception.Message }
            & $addRow "TCP port 443" "FAIL" $m
            $abort = $true
        }
    }

    # ── Step 4: TLS handshake + certificate inspection ────────────────────────
    # EXO only connects to migration endpoints with a publicly trusted, valid,
    # name-matching certificate — so this is checked explicitly, independent of
    # the "Ignore certificate errors" option (which only affects the HTTP probes).
    if (-not $abort -and -not $sync.Cancel) {
        & $setPct 48
        & $logLine "TLS handshake with $fqdn`:443 starting."
        $cert2 = $null; $sslProto = ''
        try {
            $tcp2 = New-Object System.Net.Sockets.TcpClient
            if (-not $tcp2.ConnectAsync($fqdn, 443).Wait(10000)) { $tcp2.Close(); throw "TCP connect timeout" }
            # Accept-all callback: the handshake must succeed so the certificate
            # can be captured; trust and name are validated manually below.
            $ssl = New-Object System.Net.Security.SslStream($tcp2.GetStream(), $false, { $true })
            $ssl.AuthenticateAsClient($fqdn)
            $sslProto = $ssl.SslProtocol.ToString()
            $cert2 = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($ssl.RemoteCertificate)
            $ssl.Dispose(); $tcp2.Close()
        } catch {
            $m = if ($_.Exception.InnerException) { $_.Exception.InnerException.Message } else { $_.Exception.Message }
            & $logLine "TLS handshake failed: $m"
            & $addRow "TLS handshake" "FAIL" $m
            $abort = $true
        }

        if ($cert2) {
            & $logLine "TLS: $sslProto"
            & $logLine "Certificate subject: $($cert2.Subject)"
            & $logLine "Certificate issuer:  $($cert2.Issuer)"
            & $logLine "Certificate valid:   $($cert2.NotBefore.ToString('yyyy-MM-dd')) - $($cert2.NotAfter.ToString('yyyy-MM-dd'))"
            & $addRow "TLS handshake" "OK" "$sslProto — $($cert2.Subject)"

            # Trust chain (revocation not checked to stay fast/offline-capable)
            $chain = New-Object System.Security.Cryptography.X509Certificates.X509Chain
            $chain.ChainPolicy.RevocationMode = [System.Security.Cryptography.X509Certificates.X509RevocationMode]::NoCheck
            $chainOk  = $chain.Build($cert2)
            $chainErr = @($chain.ChainStatus | ForEach-Object { $_.Status.ToString() }) -join ', '
            & $logLine "Certificate chain: $(if ($chainOk) { 'valid' } else { $chainErr })"
            if ($chainOk) {
                & $addRow "Certificate trust" "OK" "Chain valid (issuer: $($cert2.Issuer))"
            } else {
                & $addRow "Certificate trust" "FAIL" "$chainErr — Exchange Online requires a publicly trusted certificate on the MRS endpoint"
            }

            # Name match against SAN entries (with wildcard support)
            $sanNames = @()
            foreach ($ext in $cert2.Extensions) {
                if ($ext.Oid.Value -eq '2.5.29.17') {
                    # Format() output is localized ("DNS Name=", "DNS-Name=", ...) —
                    # extract hostname-shaped tokens instead of parsing labels
                    $sanNames = @([regex]::Matches($ext.Format($false), '(?i)((?:\*\.)?(?:[a-z0-9\-]+\.)+[a-z0-9\-]{2,})') |
                        ForEach-Object { $_.Groups[1].Value })
                }
            }
            if ($sanNames.Count -eq 0 -and $cert2.Subject -match 'CN=([^,]+)') { $sanNames = @($Matches[1].Trim()) }
            $nameMatch = $false
            foreach ($n in $sanNames) {
                if ($n -ieq $fqdn) { $nameMatch = $true; break }
                if ($n.StartsWith('*.')) {
                    $suffix = $n.Substring(1)   # ".contoso.com"
                    if ($fqdn.Length -gt $suffix.Length -and
                        $fqdn.EndsWith($suffix, [System.StringComparison]::OrdinalIgnoreCase) -and
                        $fqdn.Substring(0, $fqdn.Length - $suffix.Length) -notmatch '\.') {
                        $nameMatch = $true; break
                    }
                }
            }
            & $logLine "Certificate names:   $($sanNames -join ', ')"
            if ($nameMatch) {
                & $addRow "Certificate name match" "OK" "$fqdn is covered ($($sanNames -join ', '))"
            } else {
                & $addRow "Certificate name match" "FAIL" "$fqdn not in certificate names: $($sanNames -join ', ')"
            }

            # Validity period
            $daysLeft = [int]($cert2.NotAfter - (Get-Date)).TotalDays
            if ((Get-Date) -lt $cert2.NotBefore -or $daysLeft -lt 0) {
                & $addRow "Certificate validity" "FAIL" "Not valid: $($cert2.NotBefore.ToString('yyyy-MM-dd')) - $($cert2.NotAfter.ToString('yyyy-MM-dd'))"
            } elseif ($daysLeft -lt 30) {
                & $addRow "Certificate validity" "WARN" "Expires in $daysLeft day(s): $($cert2.NotAfter.ToString('yyyy-MM-dd'))"
            } else {
                & $addRow "Certificate validity" "OK" "Valid until $($cert2.NotAfter.ToString('yyyy-MM-dd')) ($daysLeft days left)"
            }
        }
    }

    # ── Step 5: MRS Proxy endpoint (unauthenticated challenge) ────────────────
    $mrsUrl = "https://$fqdn/EWS/mrsproxy.svc"
    if (-not $abort -and -not $sync.Cancel) {
        & $setPct 62
        & $logLine ""
        & $logLine "Probing MRS Proxy endpoint: $mrsUrl"
        $r = & $mrsGet $mrsUrl $null
        if ($r.Code -ge 0) { & $logLine "GetLastError=0; httpStatus=$($r.Code)." }
        if ($r.Code -eq 401) {
            $wa = if ($r.WwwAuth) { $r.WwwAuth } else { '(no WWW-Authenticate header)' }
            & $logLine "  WWW-Authenticate: $wa"
            if ($r.WwwAuth -match 'Negotiate|NTLM') {
                & $addRow "MRS Proxy endpoint" "OK" "Published — 401 challenge offers: $wa"
            } else {
                & $addRow "MRS Proxy endpoint" "WARN" "401 without Negotiate/NTLM ($wa) — EXO migration endpoints authenticate via NTLM/Negotiate"
            }
        } elseif ($r.Code -eq 404) {
            & $addRow "MRS Proxy endpoint" "FAIL" "HTTP 404 — mrsproxy.svc not found under /EWS. Check external EWS publishing and MRSProxyEnabled (Set-WebServicesVirtualDirectory -MRSProxyEnabled `$true)"
        } elseif ($r.Code -eq 403) {
            & $addRow "MRS Proxy endpoint" "FAIL" "HTTP 403 — access blocked. MRS Proxy may be disabled on the EWS virtual directory, or a reverse proxy is filtering the path"
        } elseif ($r.Code -eq 503) {
            & $addRow "MRS Proxy endpoint" "FAIL" "HTTP 503 — service unavailable (EWS app pool stopped or backend down)"
        } elseif ($r.Code -eq 200) {
            & $addRow "MRS Proxy endpoint" "WARN" "HTTP 200 without authentication — unusual, the endpoint should challenge with 401"
        } elseif ($r.Code -ge 300 -and $r.Code -lt 400) {
            & $addRow "MRS Proxy endpoint" "WARN" "HTTP $($r.Code) redirect to $($r.Location) — EXO does not follow redirects on the MRS endpoint"
        } elseif ($r.Code -lt 0) {
            & $addRow "MRS Proxy endpoint" "FAIL" $r.Error
            $abort = $true
        } else {
            & $addRow "MRS Proxy endpoint" "WARN" "HTTP $($r.Code) — unexpected response"
        }
    }

    # ── Step 6: MRS Proxy authenticated probe ─────────────────────────────────
    if (-not $abort -and -not $sync.Cancel) {
        & $setPct 72
        $who = if ($netCred) { $user } else { "logged-in user ($env:USERDOMAIN\$env:USERNAME)" }
        if ($netCred) {
            # Exchange Online authenticates to the MRS proxy with NTLM (it cannot
            # reach the on-prem KDC for Kerberos). A local client offered
            # "Negotiate" tries Kerberos first, which can fail on an SPN
            # misconfiguration and return 401 even though NTLM would succeed.
            # So probe NTLM explicitly (mirrors EXO) and Negotiate separately.
            $uriObj    = New-Object System.Uri($mrsUrl)
            $ntlmCache = New-Object System.Net.CredentialCache
            $ntlmCache.Add($uriObj, "NTLM", $netCred)
            & $logLine "Authenticated probe as $who — NTLM (like Exchange Online) starting."
            $rN = & $mrsGet $mrsUrl $ntlmCache
            & $logLine "  NTLM: httpStatus=$($rN.Code)."

            & $logLine "Authenticated probe as $who — Negotiate/Kerberos (WinHTTP) starting."
            $rW = & $mrsGetWinHttp $mrsUrl $user $pass $sync.IgnoreCert
            & $logLine "  Negotiate: httpStatus=$($rW.Code)."

            # Prefer an authenticated response from either scheme. HTTP 400 counts
            # as success here: it means auth completed and the MRS proxy processed
            # the request but rejected the plain GET (it expects a SOAP POST).
            if ($rN.Code -eq 200 -or $rN.Code -eq 400)      { $r2 = $rN }
            elseif ($rW.Code -eq 200 -or $rW.Code -eq 400)  { $r2 = $rW }
            else { $r2 = @{ Code = $rN.Code; Error = "NTLM=$($rN.Code), Negotiate=$($rW.Code)" } }
        } else {
            & $logLine "Authenticated probe as $who (logged-in user) starting."
            $r2 = & $mrsGetWinHttp $mrsUrl '' '' $sync.IgnoreCert
            if ($r2.Code -lt 0) { $r2 = & $mrsGet $mrsUrl 'default' }
        }
        if ($r2.Code -ge 0) { & $logLine "GetLastError=0; httpStatus=$($r2.Code)." }
        if ($r2.Code -eq 200 -or $r2.Code -eq 400) {
            & $addRow "MRS Proxy authentication" "OK" "HTTP $($r2.Code) as $who — authenticated successfully. The MRS proxy accepted the credentials; a GET returns $($r2.Code) because the endpoint expects a SOAP POST, which is normal and confirms the endpoint is a working MRS proxy."
        } elseif ($r2.Code -eq 401) {
            $detail = if ($r2.Error) { " ($($r2.Error))" } else { "" }
            & $addRow "MRS Proxy authentication" "INFO" "HTTP 401 as $who$detail — both NTLM and Negotiate were rejected by a local probe. This tests plain Windows auth, which may not match how Exchange Online reached the endpoint (different source/trust or credential path). It does not by itself mean the endpoint or your migration is broken. Authoritative check: 'Verify from Exchange Online' (Test-MigrationServerAvailability) — a successful EXO migration-endpoint creation already confirms it. Tip: browse the EWS URL and sign in with the same account to see whether it can do Windows auth at all."
        } elseif ($r2.Code -eq 403) {
            & $addRow "MRS Proxy authentication" "WARN" "HTTP 403 as $who — authenticated but access denied (check MRSProxyEnabled on the EWS vdir)"
        } elseif ($r2.Code -lt 0) {
            & $addRow "MRS Proxy authentication" "FAIL" $r2.Error
        } else {
            & $addRow "MRS Proxy authentication" "WARN" "HTTP $($r2.Code) as $who — unexpected response"
        }
    }

    # ── Step 6b: MRS Proxy SOAP probe ─────────────────────────────────────────
    # A SOAP POST goes further than a GET: the WCF service's own fault text tells
    # us definitively whether the MRS proxy is live and enabled.
    if (-not $abort -and -not $sync.Cancel -and ($r2.Code -eq 200 -or $r2.Code -eq 400)) {
        & $setPct 76
        & $logLine "SOAP probe (POST text/xml, SOAP 1.1) as $who starting."
        $probeUser = if ($netCred) { $user } else { '' }
        $probePass = if ($netCred) { $pass } else { '' }
        $rs = & $mrsSoapPost $mrsUrl $probeUser $probePass $sync.IgnoreCert
        & $logLine "  SOAP: httpStatus=$($rs.Code)."
        if ($rs.Body) {
            $snip = ($rs.Body -replace '\s+', ' ').Trim()
            if ($snip.Length -gt 400) { $snip = $snip.Substring(0, 400) + '…' }
            & $logLine "  SOAP response: $snip"
        }
        $b = "$($rs.Body)"
        if ($rs.Code -lt 0) {
            & $addRow "MRS Proxy SOAP" "FAIL" $rs.Error
        } elseif ($rs.Code -eq 401) {
            & $addRow "MRS Proxy SOAP" "INFO" "HTTP 401 — SOAP request was not authenticated (see the authentication row above)."
        } elseif ($b -match '(?i)MRSProxy|MailboxReplication|Mailbox Replication|ServiceVersion|WrongServerVersion') {
            & $addRow "MRS Proxy SOAP" "OK" "The Mailbox Replication proxy answered the SOAP request (HTTP $($rs.Code)) — the MRS proxy is live and enabled. See the Log tab for the server's response."
        } elseif ($b -match '(?i)disabled|not enabled') {
            & $addRow "MRS Proxy SOAP" "WARN" "HTTP $($rs.Code) — the server indicates the MRS proxy may be disabled (Set-WebServicesVirtualDirectory -MRSProxyEnabled `$true). See the Log tab."
        } elseif ($b -match '(?i)Fault|ContractFilter|cannot be processed|Action|a:Sender|a:Receiver') {
            & $addRow "MRS Proxy SOAP" "OK" "The WCF endpoint processed the SOAP request and returned a fault to the minimal envelope (HTTP $($rs.Code)) — expected, and it confirms a live SOAP service. See the Log tab for the fault text."
        } elseif ($rs.Code -eq 503 -or $rs.Code -eq 415) {
            & $addRow "MRS Proxy SOAP" "INFO" "HTTP $($rs.Code) — the endpoint did not service this SOAP request. The MRSProxy WCF binding is internal/undocumented, so a hand-crafted local SOAP call is inconclusive. This is not a fault of the endpoint — the authenticated 400 above already proves it is a live MRS proxy that accepts your credentials; 'Verify from Exchange Online' is the end-to-end proof."
        } else {
            & $addRow "MRS Proxy SOAP" "INFO" "HTTP $($rs.Code) — see the Log tab for the raw response."
        }
    }

    # ── Step 7 (optional): Test-MigrationServerAvailability from EXO ──────────
    # Runs the real cmdlet inside Exchange Online via the AdminAPI (the same REST
    # endpoint the EXO V3 module uses) — the only check that proves Microsoft's
    # datacenter can reach the on-prem endpoint.
    if ($runExo -and -not $abort -and -not $sync.Cancel) {
        & $setPct 80
        & $logLine ""
        & $logLine "Exchange Online verification (Test-MigrationServerAvailability):"

        $clientId = 'fb78d390-0c51-40cd-8e17-fdbfab77341b'   # 'Microsoft Exchange REST API Based Powershell' public client
        $scope    = 'https://outlook.office365.com/.default offline_access'
        $dcUrl    = 'https://login.microsoftonline.com/organizations/oauth2/v2.0/devicecode'
        $tokenUrl = 'https://login.microsoftonline.com/organizations/oauth2/v2.0/token'

        & $logLine "EXO: requesting device code…"
        $dcJson = $null
        try {
            $dcBytes = [System.Text.Encoding]::UTF8.GetBytes(
                "client_id=$([Uri]::EscapeDataString($clientId))&scope=$([Uri]::EscapeDataString($scope))")
            $rq = [System.Net.HttpWebRequest]::Create($dcUrl)
            $rq.Method = "POST"; $rq.ContentType = "application/x-www-form-urlencoded"
            $rq.ContentLength = $dcBytes.Length; $rq.Timeout = 15000
            $ss = $rq.GetRequestStream(); $ss.Write($dcBytes, 0, $dcBytes.Length); $ss.Close()
            $rp = $rq.GetResponse()
            $dcJson = (New-Object System.IO.StreamReader($rp.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
            $rp.Close()
        } catch {
            & $logLine "EXO: device code request failed — $($_.Exception.Message)"
        }

        $exoToken = $null
        if ($dcJson) {
            $userCode   = $dcJson.user_code
            $deviceCode = $dcJson.device_code
            $verifyUri  = if ($dcJson.verification_uri) { $dcJson.verification_uri } else { $dcJson.verification_url }
            $pollSec    = [int]$dcJson.interval; if ($pollSec -lt 5) { $pollSec = 5 }
            & $logLine "EXO: visit $verifyUri — enter code: $userCode  (sign in as an Exchange admin)"

            $sync.DeviceToken  = $null
            $sync.DeviceError  = $null
            $sync.DeviceCancel = $false

            $dcForm = New-Object System.Windows.Forms.Form
            $dcForm.Text            = "Sign in to Microsoft  —  Exchange Online"
            $dcForm.Size            = New-Object System.Drawing.Size(440, 210)
            $dcForm.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
            $dcForm.MinimizeBox     = $false
            $dcForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog

            $lbl1 = New-Object System.Windows.Forms.Label
            $lbl1.Text     = "1.  Open a browser and go to:"
            $lbl1.Location = New-Object System.Drawing.Point(12, 14)
            $lbl1.AutoSize = $true

            $lnk = New-Object System.Windows.Forms.LinkLabel
            $lnk.Text     = $verifyUri
            $lnk.Location = New-Object System.Drawing.Point(28, 34)
            $lnk.AutoSize = $true
            $lnk.Add_LinkClicked({ [System.Diagnostics.Process]::Start($lnk.Text) })

            $lbl2 = New-Object System.Windows.Forms.Label
            $lbl2.Text     = "2.  Enter this code (sign in as an Exchange admin):"
            $lbl2.Location = New-Object System.Drawing.Point(12, 62)
            $lbl2.AutoSize = $true

            $lblCode = New-Object System.Windows.Forms.Label
            $lblCode.Text      = $userCode
            $lblCode.Font      = New-Object System.Drawing.Font("Consolas", 22, [System.Drawing.FontStyle]::Bold)
            $lblCode.Location  = New-Object System.Drawing.Point(28, 80)
            $lblCode.AutoSize  = $true
            $lblCode.ForeColor = [System.Drawing.Color]::DarkBlue

            $btnCopy = New-Object System.Windows.Forms.Button
            $btnCopy.Text     = "Copy"
            $btnCopy.Location = New-Object System.Drawing.Point(340, 82)
            $btnCopy.Size     = New-Object System.Drawing.Size(72, 26)
            $btnCopy.Add_Click({ [System.Windows.Forms.Clipboard]::SetText($userCode) })

            $lblWait = New-Object System.Windows.Forms.Label
            $lblWait.Text      = "Waiting for sign-in…"
            $lblWait.Location  = New-Object System.Drawing.Point(12, 144)
            $lblWait.AutoSize  = $true
            $lblWait.ForeColor = [System.Drawing.Color]::Gray

            $btnCancelDC = New-Object System.Windows.Forms.Button
            $btnCancelDC.Text     = "Cancel"
            $btnCancelDC.Location = New-Object System.Drawing.Point(340, 140)
            $btnCancelDC.Size     = New-Object System.Drawing.Size(72, 26)
            $btnCancelDC.Add_Click({ $sync.DeviceCancel = $true; $dcForm.Close() })

            $dcForm.Controls.AddRange(@($lbl1, $lnk, $lbl2, $lblCode, $btnCopy, $lblWait, $btnCancelDC))

            $pollTimer = New-Object System.Windows.Forms.Timer
            $pollTimer.Interval = $pollSec * 1000
            $pollTimer.Add_Tick({
                if ($sync.DeviceToken -or $sync.DeviceError -or $sync.DeviceCancel) { return }
                $pb = [System.Text.Encoding]::UTF8.GetBytes(
                    "grant_type=urn:ietf:params:oauth:grant-type:device_code" +
                    "&client_id=$([Uri]::EscapeDataString($clientId))" +
                    "&device_code=$([Uri]::EscapeDataString($deviceCode))")
                try {
                    $rq2 = [System.Net.HttpWebRequest]::Create($tokenUrl)
                    $rq2.Method        = "POST"
                    $rq2.ContentType   = "application/x-www-form-urlencoded"
                    $rq2.ContentLength = $pb.Length
                    $rq2.Timeout       = 4000
                    $ss2 = $rq2.GetRequestStream(); $ss2.Write($pb, 0, $pb.Length); $ss2.Close()
                    try {
                        $rp2  = $rq2.GetResponse()
                        $tokJ = (New-Object System.IO.StreamReader($rp2.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
                        $rp2.Close()
                        if ($tokJ.access_token) {
                            $sync.DeviceToken = "Bearer $($tokJ.access_token)"
                            $dcForm.Close()
                        }
                    } catch [System.Net.WebException] {
                        $ex2 = $_.Exception
                        if ($ex2.Response) {
                            $ej2 = (New-Object System.IO.StreamReader($ex2.Response.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
                            $ex2.Response.Close()
                            switch ($ej2.error) {
                                'authorization_pending' {}
                                'slow_down'             { $pollTimer.Interval += 5000 }
                                default {
                                    $sync.DeviceError = "$($ej2.error): $($ej2.error_description)"
                                    $dcForm.Close()
                                }
                            }
                        }
                    }
                } catch {}
            })

            $dcForm.Add_Shown({ $pollTimer.Start() })
            $dcForm.Add_FormClosed({ $pollTimer.Stop() })
            [void]$dcForm.ShowDialog()
            $pollTimer.Dispose()
            $dcForm.Dispose()

            if ($sync.DeviceCancel -or (-not $sync.DeviceToken -and -not $sync.DeviceError)) {
                & $logLine "EXO: sign-in cancelled."
                & $addRow "EXO sign-in" "SKIP" "Sign-in cancelled — Exchange Online verification skipped"
            } elseif ($sync.DeviceError) {
                & $logLine "EXO: sign-in error — $($sync.DeviceError)"
                & $addRow "EXO sign-in" "FAIL" $sync.DeviceError
                $sync.DeviceError = $null
            } else {
                $exoToken = $sync.DeviceToken; $sync.DeviceToken = $null
                & $logLine "EXO: access token acquired."
            }
        } else {
            & $addRow "EXO sign-in" "FAIL" "Device code request failed — see Log tab"
        }

        if ($exoToken -and -not $sync.Cancel) {
            # Tenant ID + admin UPN from the token claims
            $tid = $null; $adminUpn = $null
            try {
                $p = ($exoToken -replace '^Bearer\s+', '').Split('.')[1].Replace('-', '+').Replace('_', '/')
                switch ($p.Length % 4) { 2 { $p += '==' } 3 { $p += '=' } }
                $claims = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($p)) | ConvertFrom-Json
                $tid      = $claims.tid
                $adminUpn = if ($claims.upn) { $claims.upn } else { $claims.unique_name }
            } catch {
                & $logLine "EXO: could not decode token claims — $($_.Exception.Message)"
            }

            if ($tid) {
                & $addRow "EXO sign-in" "OK" "Signed in as $adminUpn"
                & $setPct 88
                & $logLine "EXO: invoking Test-MigrationServerAvailability -ExchangeRemoteMove -RemoteServer $fqdn (can take a minute)…"
                $payload = @{
                    CmdletInput = @{
                        CmdletName = 'Test-MigrationServerAvailability'
                        Parameters = @{
                            ExchangeRemoteMove = $true
                            RemoteServer       = $fqdn
                            Credentials        = @{ UserName = $user; Password = $pass }
                        }
                    }
                } | ConvertTo-Json -Depth 5
                $pb = [System.Text.Encoding]::UTF8.GetBytes($payload)
                try {
                    $rq = [System.Net.HttpWebRequest]::Create("https://outlook.office365.com/adminapi/beta/$tid/InvokeCommand")
                    $rq.Method        = "POST"
                    $rq.ContentType   = "application/json"
                    $rq.Accept        = "application/json"
                    $rq.ContentLength = $pb.Length
                    $rq.Timeout       = 180000
                    $rq.Headers["Authorization"]   = $exoToken
                    $rq.Headers["X-AnchorMailbox"] = "UPN:$adminUpn"
                    $ss = $rq.GetRequestStream(); $ss.Write($pb, 0, $pb.Length); $ss.Close()
                    $rp   = $rq.GetResponse()
                    $body = (New-Object System.IO.StreamReader($rp.GetResponseStream())).ReadToEnd()
                    $rp.Close()
                    & $logLine "EXO response: $body"
                    $j = $body | ConvertFrom-Json
                    $v = if ($j.value) { @($j.value)[0] } else { $null }
                    if ($v -and "$($v.Result)" -match 'Success') {
                        & $addRow "Test-MigrationServerAvailability" "OK" "Result: $($v.Result)$(if ($v.Message) { " — $($v.Message)" })"
                    } elseif ($v) {
                        & $addRow "Test-MigrationServerAvailability" "FAIL" "Result: $($v.Result) — $($v.Message)"
                    } else {
                        & $addRow "Test-MigrationServerAvailability" "WARN" "No result object returned — see Log tab"
                    }
                } catch [System.Net.WebException] {
                    $ex = $_.Exception
                    if ($ex.Response) {
                        $eBody = ''
                        try { $eBody = (New-Object System.IO.StreamReader($ex.Response.GetResponseStream())).ReadToEnd() } catch {}
                        $code = [int]$ex.Response.StatusCode
                        $ex.Response.Close()
                        & $logLine "EXO error (HTTP $code): $eBody"
                        $eMsg = try { ($eBody | ConvertFrom-Json).error.message } catch { $null }
                        if (-not $eMsg) { $eMsg = "HTTP $code — see Log tab" }
                        if ($eBody -match 'Unable to cast|InvalidCastException') {
                            $psCmd = "Connect-ExchangeOnline; Test-MigrationServerAvailability -ExchangeRemoteMove -RemoteServer $fqdn -Credentials (Get-Credential)"
                            & $logLine "EXO: this cmdlet is not callable via the REST AdminAPI. Run it in Exchange Online PowerShell:"
                            & $logLine "     $psCmd"
                            & $addRow "Test-MigrationServerAvailability" "INFO" "Not callable via the Exchange Online REST AdminAPI (server-side serialization error on the Credentials parameter — a limitation of this call path, not an endpoint problem). Run it in Exchange Online PowerShell instead (double-click this row to copy the command):  $psCmd  — or rely on a successful New-MigrationEndpoint, which runs the same check server-side."
                        } else {
                            $hint = if ($code -eq 401 -or $code -eq 403) { " (account needs an Exchange admin role with migration permissions)" } else { '' }
                            & $addRow "Test-MigrationServerAvailability" "FAIL" "$eMsg$hint"
                        }
                    } else {
                        & $logLine "EXO error: $($ex.Message)"
                        & $addRow "Test-MigrationServerAvailability" "FAIL" $ex.Message
                    }
                } catch {
                    & $logLine "EXO error: $($_.Exception.Message)"
                    & $addRow "Test-MigrationServerAvailability" "FAIL" $_.Exception.Message
                }
            } else {
                & $addRow "EXO sign-in" "FAIL" "Could not determine tenant from token"
            }
        }
    }

    # Reset the global validation callback if this runspace set it — the
    # scriptblock delegate dies with the runspace and would break later probes
    if ($sync.IgnoreCert) {
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null
    }

    & $setPct 100
    $sync.Done = $true
}

# Adds one result row to the hybrid Results ListView (UI thread only)
function Add-HybridRow {
    param($row)
    $item = New-Object System.Windows.Forms.ListViewItem($row.Step)
    [void]$item.SubItems.Add($row.Result)
    [void]$item.SubItems.Add($row.Details)
    $item.ToolTipText = "$($row.Step) — $($row.Result)`r`n$($row.Details)"
    $item.ForeColor = switch ($row.Result) {
        'OK'    { [System.Drawing.Color]::DarkGreen }
        'WARN'  { [System.Drawing.Color]::DarkOrange }
        'FAIL'  { [System.Drawing.Color]::DarkRed }
        'INFO'  { [System.Drawing.Color]::SteelBlue }
        'SKIP'  { [System.Drawing.Color]::Gray }
        default { [System.Drawing.Color]::Black }
    }
    [void]$hybLvwResults.Items.Add($item)
}

# Called from the poll timer when $sync.Done becomes $true
function Complete-HybridTest {
    $script:HybPollTimer.Stop()

    $msg = $null
    while ($script:HybSync.Queue.TryDequeue([ref]$msg)) { $hybRtbLog.AppendText("$msg`r`n") }
    $row = $null
    while ($script:HybSync.RowQueue.TryDequeue([ref]$row)) { Add-HybridRow $row }

    $hybBtnTest.Enabled   = $true
    $hybBtnCancel.Enabled = $false

    if ($script:HybSync.Cancel) {
        $hybForm.Text = "Hybrid Deployment  —  Migration Endpoint (MRS Proxy)"
        $hybRtbLog.AppendText("`r`nTest cancelled.`r`n")
    } elseif ($script:HybSync.Failed) {
        $hybForm.Text = "Hybrid Deployment  —  Migration Endpoint (MRS Proxy)  —  Issues found"
        $hybRtbLog.AppendText("`r`nMigration endpoint test finished with issues.`r`n")
    } else {
        $hybForm.Text = "Hybrid Deployment  —  Migration Endpoint (MRS Proxy)  —  OK"
        $hybRtbLog.AppendText("`r`nMigration endpoint test completed successfully.`r`n")
    }
    if ($hybLvwResults.Items.Count -gt 0) { $hybTabCtrl.SelectedTab = $hybTabResults }

    try { $script:HybPS.Dispose() }   catch {}
    try { $script:HybRS.Close(); $script:HybRS.Dispose() } catch {}
    $script:HybPS   = $null
    $script:HybRS   = $null
    $script:HybSync = $null
}

$hybBtnTest.Add_Click({
    $email = $hybTxtEmail.Text.Trim()
    $fqdn  = ($hybTxtFqdn.Text.Trim() -replace '^https?://', '') -replace '/.*$', ''
    $user  = $hybTxtUser.Text.Trim()
    $pass  = $hybTxtPass.Text

    if (-not $fqdn -and $email -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
        [System.Windows.Forms.MessageBox]::Show(
            "Enter the MRS endpoint FQDN, or a valid on-prem e-mail address for AutoDiscover-based discovery.",
            "Input Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }
    if ($hybChkExo.Checked -and (-not $user -or -not $pass)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Explicit on-prem credentials (user + password) are required for the Exchange Online verification" +
            " — Test-MigrationServerAvailability passes them to the on-prem endpoint.",
            "Credentials Required",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    # Clear any leftover global callback; the test runspace sets its own when
    # "Ignore certificate errors" is checked
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null

    # Reset UI
    $hybForm.Text = "Hybrid Deployment  —  Migration Endpoint (MRS Proxy)"
    $hybRtbLog.Clear()
    $hybLvwResults.Items.Clear()
    $hybPrgBar.Value      = 0
    $hybBtnTest.Enabled   = $false
    $hybBtnCancel.Enabled = $true
    $hybTabCtrl.SelectedTab = $hybTabLog

    $sync = [hashtable]::Synchronized(@{
        Email        = $email
        Fqdn         = $fqdn
        User         = $user
        Password     = $pass
        RunExo       = $hybChkExo.Checked
        IgnoreCert   = $hybChkIgnoreCert.Checked
        Cancel       = $false
        Done         = $false
        Pct          = 0
        Failed       = $false
        Queue        = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
        RowQueue     = [System.Collections.Concurrent.ConcurrentQueue[hashtable]]::new()
        DeviceToken  = $null
        DeviceError  = $null
        DeviceCancel = $false
    })
    $script:HybSync = $sync

    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = [System.Threading.ApartmentState]::STA
    $rs.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::UseNewThread
    $rs.Open()
    $script:HybRS = $rs

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript($script:HybridTestScript).AddArgument($sync)
    $script:HybPS = $ps
    [void]$ps.BeginInvoke()

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 100
    $timer.Add_Tick({
        $msg = $null
        while ($script:HybSync -and $script:HybSync.Queue.TryDequeue([ref]$msg)) {
            $hybRtbLog.AppendText("$msg`r`n")
            $hybRtbLog.ScrollToCaret()
        }
        $row = $null
        while ($script:HybSync -and $script:HybSync.RowQueue.TryDequeue([ref]$row)) {
            Add-HybridRow $row
        }
        if ($script:HybSync) {
            $v = $script:HybSync.Pct
            if ($v -ge 0 -and $v -le 100) { $hybPrgBar.Value = $v }
        }
        if ($script:HybSync -and $script:HybSync.Done) {
            Complete-HybridTest
        }
    })
    $script:HybPollTimer = $timer
    $timer.Start()
})

$hybBtnCancel.Add_Click({
    if ($script:HybSync) { $script:HybSync.Cancel = $true }
    $hybBtnCancel.Enabled = $false
    $hybRtbLog.AppendText("Cancelling...`r`n")
})

$hybForm.Add_FormClosing({
    if ($script:HybSync)      { $script:HybSync.Cancel = $true }
    if ($script:HybPollTimer) { $script:HybPollTimer.Stop() }
    try { $script:HybPS.Dispose() } catch {}
    try { $script:HybRS.Close(); $script:HybRS.Dispose() } catch {}
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null
})

#endregion ===================================================================
#  START
#==============================================================================

# Load saved config (Client ID from a previous "Register App")
if (Test-Path $script:configPath) {
    try {
        $cfg = Get-Content $script:configPath -Raw | ConvertFrom-Json
        if ($cfg.ClientId) { $txtClientId.Text = $cfg.ClientId }
    } catch {}
}

# Detect domain-joined status and set SCP checkbox accordingly
$chkUseSCP.Checked = $false
try {
    [void][System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain()
    $chkUseSCP.Checked = $true
} catch {}

# Pre-fill e-mail with the logged-in user's UPN (multiple fallback methods)
$upn = ''

# Method 1: WindowsIdentity UPN claim (works on domain-joined with Kerberos)
if (-not $upn) {
    try {
        $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $claim = $id.Claims | Where-Object {
            $_.Type -eq 'http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn'
        } | Select-Object -First 1
        if ($claim -and $claim.Value) { $upn = $claim.Value }
    } catch {}
}

# Method 2: USERPRINCIPALNAME environment variable (set by some SSO/MDM solutions)
if (-not $upn) {
    try {
        if ($env:USERPRINCIPALNAME -and $env:USERPRINCIPALNAME -match '@') {
            $upn = $env:USERPRINCIPALNAME
        }
    } catch {}
}

# Method 3: ADSI lookup in Active Directory (domain-joined machines)
if (-not $upn) {
    try {
        $searcher = [adsisearcher]"samaccountname=$env:USERNAME"
        [void]$searcher.PropertiesToLoad.Add("userprincipalname")
        $result = $searcher.FindOne()
        if ($result) {
            $val = $result.Properties['userprincipalname']
            if ($val -and $val.Count -gt 0 -and $val[0] -match '@') { $upn = $val[0] }
        }
    } catch {}
}

# Method 4: whoami /upn (works on domain-joined, may be slow on non-domain machines)
if (-not $upn) {
    try {
        $w = & whoami.exe /upn 2>$null
        if ($w -and $w.Trim() -match '@') { $upn = $w.Trim() }
    } catch {}
}

if ($upn) {
    $txtEmail.Text    = $upn
    $hybTxtEmail.Text = $upn
}

#region ======================================================================
#  HYBRID CONNECTIVITY TEST  (AutoDiscover / OAuth / Free-Busy / Health)
#==============================================================================

$hcForm = New-Object System.Windows.Forms.Form
$hcForm.Text            = "Hybrid Connectivity  —  AutoDiscover · OAuth · Free/Busy · Health"
$hcForm.ClientSize      = New-Object System.Drawing.Size(620, 150)
$hcForm.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
$hcForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$hcForm.MaximizeBox     = $false
$hcForm.MinimizeBox     = $true

$hcLblEmail = New-Object System.Windows.Forms.Label
$hcLblEmail.Text      = "E-Mail Address"
$hcLblEmail.Location  = New-Object System.Drawing.Point(8, 14)
$hcLblEmail.Size      = New-Object System.Drawing.Size(120, 20)
$hcLblEmail.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$hcForm.Controls.Add($hcLblEmail)

$hcTxtEmail = New-Object System.Windows.Forms.TextBox
$hcTxtEmail.Location = New-Object System.Drawing.Point(134, 11)
$hcTxtEmail.Size     = New-Object System.Drawing.Size(478, 22)
$hcTxtEmail.TabIndex = 0
$hcForm.Controls.Add($hcTxtEmail)

$hcLblHost = New-Object System.Windows.Forms.Label
$hcLblHost.Text      = "On-Prem Host"
$hcLblHost.Location  = New-Object System.Drawing.Point(8, 42)
$hcLblHost.Size      = New-Object System.Drawing.Size(120, 20)
$hcLblHost.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$hcForm.Controls.Add($hcLblHost)

$hcTxtHost = New-Object System.Windows.Forms.TextBox
$hcTxtHost.Location = New-Object System.Drawing.Point(134, 39)
$hcTxtHost.Size     = New-Object System.Drawing.Size(280, 22)
$hcTxtHost.TabIndex = 1
$hcForm.Controls.Add($hcTxtHost)

$hcLblHostHint = New-Object System.Windows.Forms.Label
$hcLblHostHint.Text      = "(optional; defaults to autodiscover.<domain>)"
$hcLblHostHint.Location  = New-Object System.Drawing.Point(420, 42)
$hcLblHostHint.Size      = New-Object System.Drawing.Size(192, 20)
$hcLblHostHint.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$hcLblHostHint.ForeColor = [System.Drawing.Color]::Gray
$hcForm.Controls.Add($hcLblHostHint)

$hcChkIgnore = New-Object System.Windows.Forms.CheckBox
$hcChkIgnore.Text     = "Ignore certificate errors"
$hcChkIgnore.Location = New-Object System.Drawing.Point(134, 68)
$hcChkIgnore.Size     = New-Object System.Drawing.Size(220, 20)
$hcChkIgnore.TabIndex = 2
$hcForm.Controls.Add($hcChkIgnore)

$hcLblInfo = New-Object System.Windows.Forms.Label
$hcLblInfo.Text      = "Probes on-prem + Exchange Online endpoints with the current Windows user."
$hcLblInfo.Location  = New-Object System.Drawing.Point(8, 100)
$hcLblInfo.Size      = New-Object System.Drawing.Size(430, 34)
$hcLblInfo.ForeColor = [System.Drawing.Color]::Gray
$hcForm.Controls.Add($hcLblInfo)

$hcBtnTest = New-Object System.Windows.Forms.Button
$hcBtnTest.Text     = "Test"
$hcBtnTest.Location = New-Object System.Drawing.Point(452, 112)
$hcBtnTest.Size     = New-Object System.Drawing.Size(76, 26)
$hcBtnTest.TabIndex = 3
$hcForm.Controls.Add($hcBtnTest)
$hcForm.AcceptButton = $hcBtnTest

$hcBtnClose = New-Object System.Windows.Forms.Button
$hcBtnClose.Text     = "Close"
$hcBtnClose.Location = New-Object System.Drawing.Point(536, 112)
$hcBtnClose.Size     = New-Object System.Drawing.Size(76, 26)
$hcBtnClose.TabIndex = 4
$hcBtnClose.Add_Click({ $hcForm.Close() })
$hcForm.Controls.Add($hcBtnClose)
$hcForm.CancelButton = $hcBtnClose

$toolTip.SetToolTip($hcTxtEmail, "Primary SMTP / UPN. Its domain drives on-prem AutoDiscover, OAuth metadata and the tenant OpenID lookup.")
$toolTip.SetToolTip($hcTxtHost,  "External FQDN of the on-prem Exchange host (e.g. mail.contoso.com). Leave empty to use autodiscover.<domain>.")

$hcBtnTest.Add_Click({
    $e = $hcTxtEmail.Text.Trim()
    if ($e -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a valid e-mail address.", "Input Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }
    $entries = Get-HybridConnectivityEndpoints -Email $e -OnPremHost $hcTxtHost.Text.Trim()
    if ($entries.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No endpoints could be derived.", "Hybrid Connectivity",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }
    Show-EndpointTestDialog -Title "Hybrid Connectivity  —  $e" -Entries $entries `
        -AuthHeader $null -UseWinAuth $true -NetCred $null `
        -IgnoreCert $hcChkIgnore.Checked -Owner $hcForm -CsvName 'hybrid-connectivity.csv'
})

if ($upn) { $hcTxtEmail.Text = $upn }
#endregion

#region ======================================================================
#  FREE/BUSY CROSS-PREMISES TEST
#==============================================================================

# Write one line to the Free/Busy log tab. The test runs on the UI thread, so the
# text box can be written directly — no queue/timer like the AutoDiscover test needs.
# $Level only picks a colour; the text itself carries the information.
function Add-FbLog {
    param([string]$Msg, [string]$Level = '')
    if (-not $fbRtbLog) { return }
    $stamp = (Get-Date).ToString('HH:mm:ss')
    $fbRtbLog.SelectionStart  = $fbRtbLog.TextLength
    $fbRtbLog.SelectionLength = 0
    $fbRtbLog.SelectionColor  = switch ($Level) {
        'ok'   { [System.Drawing.Color]::DarkGreen }
        'warn' { [System.Drawing.Color]::DarkOrange }
        'fail' { [System.Drawing.Color]::DarkRed }
        'step' { [System.Drawing.Color]::DarkBlue }
        'dim'  { [System.Drawing.Color]::Gray }
        default { [System.Drawing.Color]::Black }
    }
    $fbRtbLog.AppendText("$stamp  $Msg`r`n")
    $fbRtbLog.SelectionColor = [System.Drawing.Color]::Black
    $fbRtbLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

# Explain what the server does behind the single request we send, and — when the
# answer is empty or an error — where to look. The probe only ever talks to one
# endpoint; the cross-organisation hop happens server-side and is invisible from here.
function Add-FbExplain {
    param([string]$Direction)
    if ($Direction -eq 'A') {
        Add-FbLog "  What happens behind this single request:" 'dim'
        Add-FbLog "    1. the on-prem Availability service receives GetUserAvailability for a foreign address" 'dim'
        Add-FbLog "    2. it looks the target domain up in its address spaces (IntraOrganizationConnector, or" 'dim'
        Add-FbLog "       the older OrganizationRelationship / AvailabilityAddressSpace)" 'dim'
        Add-FbLog "    3. it obtains a token for Exchange Online — OAuth via AuthServer + partner application" 'dim'
        Add-FbLog "       (modern hybrid), or DAuth via the federation trust (legacy)" 'dim'
        Add-FbLog "    4. it calls https://outlook.office365.com/EWS/Exchange.asmx itself and asks there" 'dim'
        Add-FbLog "    5. Exchange Online checks its own organisation relationship and the mailbox's calendar" 'dim'
        Add-FbLog "       permissions, then answers" 'dim'
        Add-FbLog "    6. the on-prem server hands the result back to us" 'dim'
        Add-FbLog "    Steps 3-5 run server-to-server and leave no trace in this log — only their outcome does." 'dim'
    } else {
        Add-FbLog "  What happens behind this single request:" 'dim'
        Add-FbLog "    1. we authenticate to Exchange Online ourselves (device code, see above)" 'dim'
        Add-FbLog "    2. Exchange Online receives GetUserAvailability for an on-prem address" 'dim'
        Add-FbLog "    3. it resolves the target via its organisation relationship to the on-prem org" 'dim'
        Add-FbLog "    4. it calls the on-prem EWS endpoint published to the internet and asks there" 'dim'
        Add-FbLog "    5. the on-prem server validates the OAuth token (partner application, AuthServer)" 'dim'
        Add-FbLog "       and answers" 'dim'
        Add-FbLog "    Steps 3-5 run server-to-server — if they fail, the error text below comes from EXO." 'dim'
    }
}

# Where to look when the call succeeded but came back empty, or failed outright.
function Add-FbWhereToLook {
    param([string]$Direction, $R)
    $onprem = ($Direction -eq 'A')     # A = on-prem did the outbound lookup
    if ($R.ResponseCode -eq 'NoError' -and ($R.ViewType -eq 'None' -or ($R.SlotsTotal -eq 0 -and $R.EventCount -eq 0))) {
        Add-FbLog "  Answer received, but without data. That rules out connectivity and authentication —" 'warn'
        Add-FbLog "  the lookup reached the other side and came back clean. Candidates, in this order:" 'warn'
        Add-FbLog "    - the mailbox genuinely has nothing in the next 7 days (check with a known-busy mailbox)" 'dim'
        if ($onprem) {
            Add-FbLog "    - sharing level in Exchange Online: Get-OrganizationRelationship | fl Name,Enabled," 'dim'
            Add-FbLog "      FreeBusyAccessEnabled,FreeBusyAccessLevel,DomainNames  (AvailabilityOnly returns" 'dim'
            Add-FbLog "      busy/free only, None returns nothing at all)" 'dim'
            Add-FbLog "    - calendar permission of the EXO mailbox: Get-MailboxFolderPermission <mbx>:\Calendar" 'dim'
            Add-FbLog "      — the Default entry must be at least AvailabilityOnly" 'dim'
        } else {
            Add-FbLog "    - sharing level on-premises: Get-OrganizationRelationship and Get-SharingPolicy" 'dim'
            Add-FbLog "    - calendar permission of the on-prem mailbox: Get-MailboxFolderPermission <mbx>:\Calendar" 'dim'
            Add-FbLog "    - Get-IntraOrganizationConnector — Enabled, and TargetAddressDomains must contain the" 'dim'
            Add-FbLog "      on-prem domain" 'dim'
        }
        return
    }
    if ($R.Http -eq 401) {
        Add-FbLog "  HTTP 401 — the request never got past authentication. Candidates:" 'fail'
        if ($onprem) {
            Add-FbLog "    - running as a user without a mailbox, or Extended Protection rejecting the token" 'dim'
            Add-FbLog "    - the EWS virtual directory requires an authentication method we did not offer" 'dim'
            Add-FbLog "      (see the WWW-Authenticate header logged above)" 'dim'
        } else {
            Add-FbLog "    - the access token was not accepted — wrong audience, or the sign-in used an account" 'dim'
            Add-FbLog "      without an Exchange Online mailbox" 'dim'
        }
        return
    }
    if ($R.ResponseCode -or $R.Fault) {
        $txt = "$($R.ResponseCode) $($R.Message) $($R.Fault)"
        Add-FbLog "  The other side answered with an error. Candidates:" 'fail'
        if ($txt -match '(?i)proxy|partner|oauth|token|unauthoriz|access denied|RBAC') {
            Add-FbLog "    - hybrid OAuth is the usual cause for this wording. On the on-prem side:" 'dim'
            Add-FbLog "      Get-PartnerApplication | fl Name,Enabled,LinkedAccount   <- an empty LinkedAccount" 'dim'
            Add-FbLog "      breaks exactly this call; Get-AuthServer | fl Name,Enabled,AuthMetadataUrl" 'dim'
            Add-FbLog "      Get-IntraOrganizationConnector | fl Name,Enabled,TargetAddressDomains" 'dim'
            Add-FbLog "    - verify end to end: Test-OAuthConnectivity -Service EWS -TargetUri <ews url> -Mailbox <mbx>" 'dim'
            Add-FbLog "    - after any change recycle MSExchangeServicesAppPool, the token cache is sticky" 'dim'
        } elseif ($txt -match '(?i)not found|no mailbox|recipient') {
            Add-FbLog "    - the target address is unknown to the answering side — check spelling, and that the" 'dim'
            Add-FbLog "      mail user / remote mailbox object carries this address" 'dim'
        } elseif ($txt -match '(?i)certificate|SSL|TLS|trust') {
            Add-FbLog "    - certificate trust on the path — try 'Ignore certificate errors' to confirm, then fix" 'dim'
            Add-FbLog "      the chain rather than leaving it switched off" 'dim'
        } else {
            Add-FbLog "    - read the MessageText above; it is generated by the answering server, not by this tool" 'dim'
        }
        return
    }
    if ($R.Http -lt 0) {
        Add-FbLog "  No HTTP response at all — name resolution, firewall or TLS handshake. Candidates:" 'fail'
        Add-FbLog "    - resolve the host and test port 443 from this machine" 'dim'
        Add-FbLog "    - for on-prem: is the EWS endpoint reachable under the name AutoDiscover returned?" 'dim'
    }
}

# Build a standard EWS GetUserAvailability (SOAP 1.1) request for one target.
function New-FreeBusySoap {
    param([string]$Target)
    $start = (Get-Date).Date
    $end   = $start.AddDays(7)
    $s = $start.ToString('yyyy-MM-ddTHH:mm:ss')
    $e = $end.ToString('yyyy-MM-ddTHH:mm:ss')
    return ('<?xml version="1.0" encoding="utf-8"?>' +
        '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">' +
        '<soap:Header><t:RequestServerVersion Version="Exchange2013"/></soap:Header>' +
        '<soap:Body><m:GetUserAvailabilityRequest>' +
        '<t:TimeZone><t:Bias>0</t:Bias>' +
        '<t:StandardTime><t:Bias>0</t:Bias><t:Time>03:00:00</t:Time><t:DayOrder>5</t:DayOrder><t:Month>10</t:Month><t:DayOfWeek>Sunday</t:DayOfWeek></t:StandardTime>' +
        '<t:DaylightTime><t:Bias>-60</t:Bias><t:Time>02:00:00</t:Time><t:DayOrder>5</t:DayOrder><t:Month>3</t:Month><t:DayOfWeek>Sunday</t:DayOfWeek></t:DaylightTime>' +
        '</t:TimeZone>' +
        '<m:MailboxDataArray><t:MailboxData><t:Email><t:Address>' + $Target + '</t:Address></t:Email>' +
        '<t:AttendeeType>Required</t:AttendeeType><t:ExcludeConflicts>false</t:ExcludeConflicts></t:MailboxData></m:MailboxDataArray>' +
        '<t:FreeBusyViewOptions><t:TimeWindow><t:StartTime>' + $s + '</t:StartTime><t:EndTime>' + $e + '</t:EndTime></t:TimeWindow>' +
        '<t:MergedFreeBusyIntervalInMinutes>60</t:MergedFreeBusyIntervalInMinutes><t:RequestedView>DetailedMerged</t:RequestedView></t:FreeBusyViewOptions>' +
        '</m:GetUserAvailabilityRequest></soap:Body></soap:Envelope>')
}

# POST GetUserAvailability via WinHTTP. $Mode: 'win' (Windows creds / logged-in)
# or 'bearer' (OAuth). Returns @{Http; ResponseCode; ViewType; Message; Fault; Raw}.
function Invoke-FreeBusyProbe {
    param([string]$EwsUrl, [string]$Target, [string]$Mode, [string]$User, [string]$Pass, [string]$Token, [bool]$IgnoreCert)
    $soap = New-FreeBusySoap -Target $Target
    $wh = $null
    Add-FbLog "POST $EwsUrl" 'step'
    Add-FbLog "  SOAP action : GetUserAvailability (Exchange2013), target = $Target, window = 7 days, view = DetailedMerged" 'dim'
    switch ($Mode) {
        'bearer' { Add-FbLog "  Authorization: Bearer token (OAuth2, length $($Token.Length))" 'dim' }
        default  {
            if ($User) { Add-FbLog "  Authorization: explicit credentials for $User (NTLM/Negotiate via WinHTTP)" 'dim' }
            else       { Add-FbLog "  Authorization: logged-in Windows user, SetAutoLogonPolicy(0) — NTLM/Negotiate via WinHTTP (handles Extended Protection)" 'dim' }
        }
    }
    if ($IgnoreCert) { Add-FbLog "  TLS          : certificate errors ignored (Option 4 = 13056)" 'warn' }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $wh = New-Object -ComObject 'WinHttp.WinHttpRequest.5.1'
        $wh.Open('POST', $EwsUrl, $false)
        $wh.SetTimeouts(30000, 30000, 30000, 30000)
        if ($IgnoreCert) { $wh.Option(4) = 13056 }
        $wh.Option(6) = $false
        $wh.SetRequestHeader('Content-Type', 'text/xml; charset=utf-8')
        if ($Mode -eq 'bearer') { $wh.SetRequestHeader('Authorization', $Token) }
        elseif ($User)          { $wh.SetCredentials($User, $Pass, 0) }
        else                    { $wh.SetAutoLogonPolicy(0) }
        $wh.Send($soap)
        $code = [int]$wh.Status
        $body = try { "$($wh.ResponseText)" } catch { '' }
        $sw.Stop()
        $lvl = if ($code -eq 200) { 'ok' } elseif ($code -eq 401) { 'fail' } else { 'warn' }
        Add-FbLog ("  GetLastError=0; httpStatus={0}. {1}  ({2} ms, {3} bytes)" -f $code, "$($wh.StatusText)", $sw.ElapsedMilliseconds, $body.Length) $lvl
        $srvHdr = try { "$($wh.GetResponseHeader('Server'))" } catch { '' }
        if ($srvHdr) { Add-FbLog "  Server header: $srvHdr" 'dim' }
        if ($code -eq 401) {
            $wwwAuth = try { "$($wh.GetResponseHeader('WWW-Authenticate'))" } catch { '' }
            if ($wwwAuth) { Add-FbLog "  WWW-Authenticate: $wwwAuth" 'dim' }
        }
        $rc = ''; $view = ''; $msg = ''; $fault = ''
        if ($body -match '<(?:\w+:)?ResponseCode>([^<]+)<')     { $rc = $Matches[1] }
        if ($body -match '<(?:\w+:)?FreeBusyViewType>([^<]+)<') { $view = $Matches[1] }
        if ($body -match '<(?:\w+:)?MessageText>([^<]+)<')      { $msg = $Matches[1] }
        if ($body -match '<(?:\w+:)?faultstring[^>]*>([^<]+)<') { $fault = $Matches[1] }

        # Merged free/busy map (one digit per interval: 0=free 1=tentative 2=busy 3=OOF 4=elsewhere)
        $merged = ''; if ($body -match '<(?:\w+:)?MergedFreeBusy>([^<]*)<') { $merged = $Matches[1] }
        $slotsTotal = $merged.Length
        $slotsBusy  = ($merged.ToCharArray() | Where-Object { $_ -ne '0' } | Measure-Object).Count

        # Detailed calendar events (only present when the view includes detail)
        $evLines = New-Object System.Collections.Generic.List[string]
        foreach ($m in [regex]::Matches($body, '(?s)<(?:\w+:)?CalendarEvent>(.*?)</(?:\w+:)?CalendarEvent>')) {
            $seg = $m.Groups[1].Value
            $st = ''; $et = ''; $bt = ''
            if ($seg -match '<(?:\w+:)?StartTime>([^<]+)<') { $st = $Matches[1] }
            if ($seg -match '<(?:\w+:)?EndTime>([^<]+)<')   { $et = $Matches[1] }
            if ($seg -match '<(?:\w+:)?BusyType>([^<]+)<')  { $bt = $Matches[1] }
            $stShort = ($st -replace 'T', ' '); if ($stShort.Length -ge 16) { $stShort = $stShort.Substring(0,16) }
            $etShort = ($et -replace '^.*T', ''); if ($etShort.Length -ge 5) { $etShort = $etShort.Substring(0,5) }
            $evLines.Add(("  {0}–{1}  {2}" -f $stShort, $etShort, $bt))
        }
        if ($rc)    { Add-FbLog "  ResponseCode : $rc" $(if ($rc -eq 'NoError') { 'ok' } else { 'fail' }) }
        if ($view)  { Add-FbLog "  FreeBusyViewType: $view" $(if ($view -eq 'None') { 'warn' } else { 'ok' }) }
        if ($msg)   { Add-FbLog "  MessageText  : $msg" 'warn' }
        if ($fault) { Add-FbLog "  SOAP fault   : $fault" 'fail' }
        if ($slotsTotal -gt 0) { Add-FbLog "  MergedFreeBusy: $slotsBusy of $slotsTotal hourly slots not free" 'dim' }
        if ($evLines.Count -gt 0) { Add-FbLog "  CalendarEvents: $($evLines.Count) returned" 'dim' }

        return @{ Http = $code; ResponseCode = $rc; ViewType = $view; Message = $msg; Fault = $fault; Raw = $body;
                  Merged = $merged; SlotsTotal = $slotsTotal; SlotsBusy = $slotsBusy;
                  EventCount = $evLines.Count; Events = ($evLines -join "`r`n") }
    } catch {
        $sw.Stop()
        Add-FbLog "  WinHTTP error after $($sw.ElapsedMilliseconds) ms: $($_.Exception.Message)" 'fail'
        return @{ Http = -1; ResponseCode = ''; ViewType = ''; Message = ''; Fault = "WinHTTP: $($_.Exception.Message)"; Raw = '';
                  Merged = ''; SlotsTotal = 0; SlotsBusy = 0; EventCount = 0; Events = '' }
    } finally {
        if ($wh) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wh) }
    }
}

# Discover the on-prem EWS URL for a mailbox via AutoDiscover (Windows auth via
# WinHTTP, which handles Extended Protection). Returns the URL or $null.
function Get-OnPremEwsUrl {
    param([string]$Mailbox, [bool]$IgnoreCert)
    if ($Mailbox -notmatch '@([^@\s]+)$') { return $null }
    $domain = $Matches[1]
    $body = '<?xml version="1.0" encoding="utf-8"?>' +
        '<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006">' +
        '<Request><EMailAddress>' + $Mailbox + '</EMailAddress>' +
        '<AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema>' +
        '</Request></Autodiscover>'
    Add-FbLog "AutoDiscover for $Mailbox (domain $domain), as logged-in Windows user" 'step'
    foreach ($u in @("https://autodiscover.$domain/autodiscover/autodiscover.xml",
                     "https://$domain/autodiscover/autodiscover.xml")) {
        $wh = $null
        try {
            $wh = New-Object -ComObject 'WinHttp.WinHttpRequest.5.1'
            $wh.Open('POST', $u, $false)
            $wh.SetTimeouts(15000, 15000, 15000, 15000)
            if ($IgnoreCert) { $wh.Option(4) = 13056 }
            $wh.Option(6) = $false
            $wh.SetRequestHeader('Content-Type', 'text/xml; charset=utf-8')
            $wh.SetAutoLogonPolicy(0)
            $wh.Send($body)
            $st = [int]$wh.Status
            Add-FbLog "  POST $u" 'dim'
            Add-FbLog "  GetLastError=0; httpStatus=$st." $(if ($st -eq 200) { 'ok' } else { 'warn' })
            if ($st -eq 200) {
                $resp = "$($wh.ResponseText)"
                if ($resp -match '<(?:\w+:)?EwsUrl>([^<]+)<') {
                    Add-FbLog "  EwsUrl = $($Matches[1])" 'ok'
                    return $Matches[1]
                }
                Add-FbLog "  200 but no EwsUrl element in the response." 'warn'
            }
        } catch {
            Add-FbLog "  POST $u" 'dim'
            Add-FbLog "  failed: $($_.Exception.Message)" 'warn'
        } finally {
            if ($wh) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wh) }
        }
    }
    Add-FbLog "  AutoDiscover did not yield an EWS URL." 'fail'
    return $null
}

# Device-code sign-in for an Exchange Online EWS token (UI thread). Returns "Bearer …" or $null.
function Get-ExoEwsToken {
    $clientId = 'd3590ed6-52b3-4102-aeff-aad2292ab01c'   # Microsoft Office (public client)
    $scope    = 'https://outlook.office365.com/EWS.AccessAsUser.All offline_access'
    $dcUrl    = 'https://login.microsoftonline.com/organizations/oauth2/v2.0/devicecode'
    $tkUrl    = 'https://login.microsoftonline.com/organizations/oauth2/v2.0/token'
    $dcJson = $null
    Add-FbLog "Exchange Online sign-in: device code flow" 'step'
    Add-FbLog "  client_id : $clientId  (Microsoft Office, public client)" 'dim'
    Add-FbLog "  scope     : $scope" 'dim'
    Add-FbLog "  POST $dcUrl" 'dim'
    try {
        $b = [System.Text.Encoding]::UTF8.GetBytes("client_id=$([Uri]::EscapeDataString($clientId))&scope=$([Uri]::EscapeDataString($scope))")
        $rq = [System.Net.HttpWebRequest]::Create($dcUrl); $rq.Method = 'POST'; $rq.ContentType = 'application/x-www-form-urlencoded'; $rq.ContentLength = $b.Length; $rq.Timeout = 15000
        $s = $rq.GetRequestStream(); $s.Write($b, 0, $b.Length); $s.Close()
        $rp = $rq.GetResponse(); $dcJson = (New-Object System.IO.StreamReader($rp.GetResponseStream())).ReadToEnd() | ConvertFrom-Json; $rp.Close()
        Add-FbLog "  GetLastError=0; httpStatus=200. device code issued, valid $($dcJson.expires_in)s, poll every $($dcJson.interval)s" 'ok'
    } catch {
        Add-FbLog "  device code request failed: $($_.Exception.Message)" 'fail'
        [System.Windows.Forms.MessageBox]::Show("Device code request failed:`n$($_.Exception.Message)", "Auth Error", 0, 16) | Out-Null
        return $null
    }
    $st = @{ Tok = $null; Err = $null; Cancel = $false }
    $uc = $dcJson.user_code; $dcode = $dcJson.device_code
    $vuri = if ($dcJson.verification_uri) { $dcJson.verification_uri } else { $dcJson.verification_url }
    $poll = [int]$dcJson.interval; if ($poll -lt 5) { $poll = 5 }

    $f = New-Object System.Windows.Forms.Form
    $f.Text = "Sign in to Exchange Online (EWS)"; $f.Size = New-Object System.Drawing.Size(450, 210)
    $f.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen; $f.MinimizeBox = $false
    $f.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $l1 = New-Object System.Windows.Forms.Label; $l1.Text = "1.  Open a browser and go to:"; $l1.Location = New-Object System.Drawing.Point(12, 14); $l1.AutoSize = $true
    $lnk = New-Object System.Windows.Forms.LinkLabel; $lnk.Text = $vuri; $lnk.Location = New-Object System.Drawing.Point(28, 34); $lnk.AutoSize = $true
    $lnk.Add_LinkClicked({ [System.Diagnostics.Process]::Start($lnk.Text) })
    $l2 = New-Object System.Windows.Forms.Label; $l2.Text = "2.  Enter this code (sign in as the EXO mailbox user):"; $l2.Location = New-Object System.Drawing.Point(12, 62); $l2.AutoSize = $true
    $lc = New-Object System.Windows.Forms.Label; $lc.Text = $uc; $lc.Font = New-Object System.Drawing.Font("Consolas", 22, [System.Drawing.FontStyle]::Bold); $lc.Location = New-Object System.Drawing.Point(28, 80); $lc.AutoSize = $true; $lc.ForeColor = [System.Drawing.Color]::DarkBlue
    $bc = New-Object System.Windows.Forms.Button; $bc.Text = "Copy"; $bc.Location = New-Object System.Drawing.Point(348, 82); $bc.Size = New-Object System.Drawing.Size(72, 26)
    $bc.Add_Click({ [System.Windows.Forms.Clipboard]::SetText($uc) })
    $lw = New-Object System.Windows.Forms.Label; $lw.Text = "Waiting for sign-in…"; $lw.Location = New-Object System.Drawing.Point(12, 144); $lw.AutoSize = $true; $lw.ForeColor = [System.Drawing.Color]::Gray
    $bx = New-Object System.Windows.Forms.Button; $bx.Text = "Cancel"; $bx.Location = New-Object System.Drawing.Point(348, 140); $bx.Size = New-Object System.Drawing.Size(72, 26)
    $bx.Add_Click({ $st.Cancel = $true; $f.Close() })
    $f.Controls.AddRange(@($l1, $lnk, $l2, $lc, $bc, $lw, $bx))
    $t = New-Object System.Windows.Forms.Timer; $t.Interval = $poll * 1000
    $t.Add_Tick({
        if ($st.Tok -or $st.Err -or $st.Cancel) { return }
        $pb = [System.Text.Encoding]::UTF8.GetBytes("grant_type=urn:ietf:params:oauth:grant-type:device_code&client_id=$([Uri]::EscapeDataString($clientId))&device_code=$([Uri]::EscapeDataString($dcode))")
        try {
            $rq2 = [System.Net.HttpWebRequest]::Create($tkUrl); $rq2.Method = 'POST'; $rq2.ContentType = 'application/x-www-form-urlencoded'; $rq2.ContentLength = $pb.Length; $rq2.Timeout = 4000
            $s2 = $rq2.GetRequestStream(); $s2.Write($pb, 0, $pb.Length); $s2.Close()
            try {
                $rp2 = $rq2.GetResponse(); $tj = (New-Object System.IO.StreamReader($rp2.GetResponseStream())).ReadToEnd() | ConvertFrom-Json; $rp2.Close()
                if ($tj.access_token) { $st.Tok = "Bearer $($tj.access_token)"; $f.Close() }
            } catch [System.Net.WebException] {
                $e2 = $_.Exception
                if ($e2.Response) {
                    $ej = (New-Object System.IO.StreamReader($e2.Response.GetResponseStream())).ReadToEnd() | ConvertFrom-Json; $e2.Response.Close()
                    switch ($ej.error) { 'authorization_pending' {} 'slow_down' { $t.Interval += 5000 } default { $st.Err = "$($ej.error): $($ej.error_description)"; $f.Close() } }
                }
            }
        } catch {}
    })
    $f.Add_Shown({ $t.Start() }); $f.Add_FormClosed({ $t.Stop() })
    Add-FbLog "  user_code = $uc  —  waiting for sign-in at $vuri" 'dim'
    [void]$f.ShowDialog(); $t.Dispose(); $f.Dispose()
    if ($st.Err) {
        Add-FbLog "  sign-in error: $($st.Err)" 'fail'
        [System.Windows.Forms.MessageBox]::Show("Sign-in error: $($st.Err)", "Auth Error", 0, 16) | Out-Null
    } elseif ($st.Cancel) {
        Add-FbLog "  sign-in cancelled by user." 'warn'
    } elseif ($st.Tok) {
        Add-FbLog "  access token acquired (POST $tkUrl)." 'ok'
    }
    return $st.Tok
}

$fbForm = New-Object System.Windows.Forms.Form
$fbForm.Text            = "Free/Busy  —  Cross-Premises Availability"
$fbForm.ClientSize      = New-Object System.Drawing.Size(860, 520)
$fbForm.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
$fbForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$fbForm.MinimizeBox     = $true
$fbForm.MaximizeBox     = $true
$fbForm.MinimumSize     = New-Object System.Drawing.Size(700, 420)

$fbAnchLR = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right)
$fbAnchTR = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)

$fbLblOnp = New-Object System.Windows.Forms.Label
$fbLblOnp.Text = "On-Prem Mailbox"; $fbLblOnp.Location = New-Object System.Drawing.Point(8, 14); $fbLblOnp.Size = New-Object System.Drawing.Size(140, 20)
$fbForm.Controls.Add($fbLblOnp)
$fbTxtOnp = New-Object System.Windows.Forms.TextBox
$fbTxtOnp.Location = New-Object System.Drawing.Point(152, 11); $fbTxtOnp.Size = New-Object System.Drawing.Size(552, 22); $fbTxtOnp.Anchor = $fbAnchLR; $fbTxtOnp.TabIndex = 0
$fbForm.Controls.Add($fbTxtOnp)
$fbLblOnpHint = New-Object System.Windows.Forms.Label
$fbLblOnpHint.Text = "(on-prem user)"; $fbLblOnpHint.Location = New-Object System.Drawing.Point(712, 14); $fbLblOnpHint.Size = New-Object System.Drawing.Size(140, 20); $fbLblOnpHint.ForeColor = [System.Drawing.Color]::Gray; $fbLblOnpHint.Anchor = $fbAnchTR
$fbForm.Controls.Add($fbLblOnpHint)

$fbLblExo = New-Object System.Windows.Forms.Label
$fbLblExo.Text = "EXO Mailbox"; $fbLblExo.Location = New-Object System.Drawing.Point(8, 42); $fbLblExo.Size = New-Object System.Drawing.Size(140, 20)
$fbForm.Controls.Add($fbLblExo)
$fbTxtExo = New-Object System.Windows.Forms.TextBox
$fbTxtExo.Location = New-Object System.Drawing.Point(152, 39); $fbTxtExo.Size = New-Object System.Drawing.Size(552, 22); $fbTxtExo.Anchor = $fbAnchLR; $fbTxtExo.TabIndex = 1
$fbForm.Controls.Add($fbTxtExo)
$fbLblExoHint = New-Object System.Windows.Forms.Label
$fbLblExoHint.Text = "(Exchange Online user)"; $fbLblExoHint.Location = New-Object System.Drawing.Point(712, 42); $fbLblExoHint.Size = New-Object System.Drawing.Size(140, 20); $fbLblExoHint.ForeColor = [System.Drawing.Color]::Gray; $fbLblExoHint.Anchor = $fbAnchTR
$fbForm.Controls.Add($fbLblExoHint)

$fbChkIgnore = New-Object System.Windows.Forms.CheckBox
$fbChkIgnore.Text = "Ignore certificate errors"; $fbChkIgnore.Location = New-Object System.Drawing.Point(152, 68); $fbChkIgnore.Size = New-Object System.Drawing.Size(220, 20); $fbChkIgnore.TabIndex = 2
$fbForm.Controls.Add($fbChkIgnore)

$fbBtnTest = New-Object System.Windows.Forms.Button
$fbBtnTest.Text = "Test"; $fbBtnTest.Location = New-Object System.Drawing.Point(692, 64); $fbBtnTest.Size = New-Object System.Drawing.Size(76, 26); $fbBtnTest.Anchor = $fbAnchTR; $fbBtnTest.TabIndex = 3
$fbForm.Controls.Add($fbBtnTest); $fbForm.AcceptButton = $fbBtnTest
$fbBtnClose = New-Object System.Windows.Forms.Button
$fbBtnClose.Text = "Close"; $fbBtnClose.Location = New-Object System.Drawing.Point(776, 64); $fbBtnClose.Size = New-Object System.Drawing.Size(76, 26); $fbBtnClose.Anchor = $fbAnchTR; $fbBtnClose.TabIndex = 4
$fbBtnClose.Add_Click({ $fbForm.Close() }); $fbForm.CancelButton = $fbBtnClose
$fbForm.Controls.Add($fbBtnClose)

# Two tabs: the verdict per direction, and the protocol log behind it.
$fbTabCtrl = New-Object System.Windows.Forms.TabControl
# Size must match the form's ClientSize (860 x 520) minus the margins, otherwise the
# anchors keep the control at its old size and the window shows empty space.
$fbTabCtrl.Location = New-Object System.Drawing.Point(8, 100); $fbTabCtrl.Size = New-Object System.Drawing.Size(844, 412)
$fbTabCtrl.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right)
$fbForm.Controls.Add($fbTabCtrl)

$fbTabResults = New-Object System.Windows.Forms.TabPage
$fbTabResults.Text = "Results"
$fbTabCtrl.Controls.Add($fbTabResults)

$fbTabLog = New-Object System.Windows.Forms.TabPage
$fbTabLog.Text = "Log"
$fbTabCtrl.Controls.Add($fbTabLog)

$fbRtbLog = New-Object System.Windows.Forms.RichTextBox
$fbRtbLog.Dock       = [System.Windows.Forms.DockStyle]::Fill
$fbRtbLog.ReadOnly   = $true
$fbRtbLog.Font       = New-Object System.Drawing.Font("Consolas", 9)
$fbRtbLog.BackColor  = [System.Drawing.Color]::White
$fbRtbLog.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Both
$fbRtbLog.WordWrap   = $false
$fbTabLog.Controls.Add($fbRtbLog)

# Right-click: copy selection (or everything) / select all / clear
$fbCtxLog = New-Object System.Windows.Forms.ContextMenuStrip
$fbMiCopy = $fbCtxLog.Items.Add("Copy")
$fbMiCopy.Add_Click({
    $t = $fbRtbLog.SelectedText
    if (-not $t) { $t = $fbRtbLog.Text }
    if ($t) { try { [System.Windows.Forms.Clipboard]::SetText($t) } catch {} }
})
$fbMiSel = $fbCtxLog.Items.Add("Select All")
$fbMiSel.Add_Click({ $fbRtbLog.SelectAll() })
$fbMiClr = $fbCtxLog.Items.Add("Clear")
$fbMiClr.Add_Click({ $fbRtbLog.Clear() })
$fbRtbLog.ContextMenuStrip = $fbCtxLog

$fbLvw = New-Object System.Windows.Forms.ListView
$fbLvw.Dock = [System.Windows.Forms.DockStyle]::Fill
$fbLvw.View = [System.Windows.Forms.View]::Details; $fbLvw.FullRowSelect = $true; $fbLvw.GridLines = $true; $fbLvw.ShowItemToolTips = $true
$fbLvw.HeaderStyle = [System.Windows.Forms.ColumnHeaderStyle]::Nonclickable
[void]$fbLvw.Columns.Add("Direction", 150)
[void]$fbLvw.Columns.Add("Result", 70)
[void]$fbLvw.Columns.Add("Details", 608)
$fbLvw.Add_DoubleClick({
    if ($fbLvw.SelectedItems.Count -gt 0) {
        $it = $fbLvw.SelectedItems[0]
        $txt = "$($it.Text) — $($it.SubItems[1].Text)`r`n`r`n$($it.SubItems[2].Text)`r`n`r`n$($it.ToolTipText)"
        try { [System.Windows.Forms.Clipboard]::SetText($txt) } catch {}
        [System.Windows.Forms.MessageBox]::Show($txt, "Free/Busy detail (copied to clipboard)", 0, 64) | Out-Null
    }
})
$fbTabResults.Controls.Add($fbLvw)

$toolTip.SetToolTip($fbTxtOnp, "An on-premises mailbox SMTP address. Target for EXO → On-Prem, and its domain is used to auto-discover the on-prem EWS endpoint.")
$toolTip.SetToolTip($fbTxtExo, "An Exchange Online mailbox SMTP address. Target for On-Prem → EXO.")

# Add one result row (interpreting a probe result)
function Add-FbRow {
    param([string]$Dir, $R, [string]$Ctx)
    $res = ''; $det = ''
    if ($R.Http -eq 401) {
        $res = 'FAIL'; $det = "HTTP 401 — authentication failed. $Ctx"
    } elseif ($R.Http -lt 0) {
        $res = 'FAIL'; $det = "$($R.Fault) $Ctx"
    } elseif ($R.ResponseCode -eq 'NoError') {
        # NoError = the availability request was served (the cross-org path worked).
        # Summarise the actual data returned over the 7-day window.
        $data = ''
        if ($R.SlotsTotal -gt 0) { $data = "$($R.SlotsBusy)/$($R.SlotsTotal) hrs busy" }
        if ($R.EventCount -gt 0) {
            if ($data) { $data += ", " }
            $data += "$($R.EventCount) appointment(s) visible"
        }
        if ($R.ViewType -eq 'None' -and -not $data) {
            $res = 'WARN'; $det = "Availability lookup succeeded (NoError) but returned no data (view = None) — the target may restrict free/busy detail or sharing is limited. $Ctx"
        } else {
            $vt = if ($R.ViewType) { $R.ViewType } else { 'returned' }
            $sum = if ($data) { " — $data (next 7 days)" } else { '' }
            $res = 'OK'; $det = "Cross-premises free/busy works — view = $vt$sum. Double-click for the appointment list. $Ctx"
        }
    } elseif ($R.ResponseCode) {
        $extra = if ($R.Message) { " — $($R.Message)" } else { '' }
        $hint = ''
        if ($R.ResponseCode -match 'Proxy' -or "$($R.Message)" -match 'linked account|RBAC|OAuth|partner') {
            $hint = "  Hint: hybrid OAuth/free-busy authorization issue on the target side — verify the on-prem OAuth config (IntraOrganizationConnector, AuthServer, and especially the partner application's LinkedAccount / RBAC role assignments); run Test-OAuthConnectivity, and recycle the EWS app pool after any change."
        }
        $res = 'FAIL'; $det = "$($R.ResponseCode)$extra.$hint $Ctx"
    } elseif ($R.Fault) {
        $res = 'FAIL'; $det = "$($R.Fault). $Ctx"
    } else {
        $res = 'INFO'; $det = "HTTP $($R.Http) — see detail (double-click). $Ctx"
    }
    $it = New-Object System.Windows.Forms.ListViewItem($Dir)
    [void]$it.SubItems.Add($res)
    [void]$it.SubItems.Add($det)
    $tip = "$Ctx"
    if ($R.ViewType)   { $tip += "`r`nView: $($R.ViewType)" }
    if ($R.SlotsTotal -gt 0) { $tip += "`r`nBusy hours (next 7 days): $($R.SlotsBusy) of $($R.SlotsTotal)" }
    if ($R.EventCount -gt 0) { $tip += "`r`nAppointments ($($R.EventCount)):`r`n$($R.Events)" }
    if ($R.Merged)     { $tip += "`r`nFree/Busy map: $($R.Merged)" }
    $tip += "`r`n`r`n--- raw response ---`r`n$($R.Raw)"
    $it.ToolTipText = $tip
    $it.ForeColor = switch ($res) {
        'OK'   { [System.Drawing.Color]::DarkGreen }
        'WARN' { [System.Drawing.Color]::DarkOrange }
        'FAIL' { [System.Drawing.Color]::DarkRed }
        'INFO' { [System.Drawing.Color]::SteelBlue }
        'SKIP' { [System.Drawing.Color]::Gray }
        default { [System.Drawing.Color]::Black }
    }
    [void]$fbLvw.Items.Add($it)
}
function Add-FbSkip { param([string]$Dir, [string]$Why)
    $it = New-Object System.Windows.Forms.ListViewItem($Dir)
    [void]$it.SubItems.Add('SKIP'); [void]$it.SubItems.Add($Why)
    $it.ForeColor = [System.Drawing.Color]::Gray
    [void]$fbLvw.Items.Add($it)
}

$fbBtnTest.Add_Click({
    $onpMbx = $fbTxtOnp.Text.Trim()
    $exoMbx = $fbTxtExo.Text.Trim()
    $ign    = $fbChkIgnore.Checked
    if (-not $onpMbx -or -not $exoMbx) {
        [System.Windows.Forms.MessageBox]::Show("Enter both mailboxes — one on-premises and one Exchange Online.", "Input Required", 0, 48) | Out-Null
        return
    }
    $fbLvw.Items.Clear()
    $fbRtbLog.Clear()
    $fbBtnTest.Enabled = $false
    $fbForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    Add-FbLog "Free/Busy cross-premises test started." 'step'
    Add-FbLog "  on-prem mailbox : $onpMbx" 'dim'
    Add-FbLog "  EXO mailbox     : $exoMbx" 'dim'
    Add-FbLog "  running as      : $env:USERDOMAIN\$env:USERNAME on $env:COMPUTERNAME" 'dim'
    Add-FbLog "" ''
    try {
        # Auto-discover the on-prem EWS URL from the on-prem mailbox; ask if it can't be found.
        [System.Windows.Forms.Application]::DoEvents()
        $ewsOnprem = Get-OnPremEwsUrl -Mailbox $onpMbx -IgnoreCert $ign
        if (-not $ewsOnprem) {
            $fbForm.Cursor = [System.Windows.Forms.Cursors]::Default
            $dom = ''; if ($onpMbx -match '@([^@\s]+)$') { $dom = $Matches[1] }
            $ans = [Microsoft.VisualBasic.Interaction]::InputBox(
                "AutoDiscover could not determine the on-prem EWS host for $onpMbx.`r`nEnter the on-prem Exchange host (e.g. mail.contoso.com):",
                "On-Prem EWS Host", "mail.$dom")
            $fbForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $ans = ($ans -replace '^https?://', '') -replace '/.*$', ''
            if ($ans) {
                $ewsOnprem = "https://$ans/EWS/Exchange.asmx"
                Add-FbLog "  EWS host entered by hand: $ewsOnprem" 'warn'
            }
        }

        # Direction A: On-Prem → EXO (on-prem EWS, logged-in Windows user, target = EXO mailbox)
        Add-FbLog "" ''
        Add-FbLog "Direction A — On-Prem -> EXO: the on-prem server looks up the EXO mailbox" 'step'
        Add-FbExplain 'A'
        if ($ewsOnprem) {
            [System.Windows.Forms.Application]::DoEvents()
            $rA = Invoke-FreeBusyProbe -EwsUrl $ewsOnprem -Target $exoMbx -Mode 'win' -User '' -Pass '' -Token '' -IgnoreCert $ign
            Add-FbWhereToLook 'A' $rA
            Add-FbRow "On-Prem → EXO" $rA "$ewsOnprem  →  $exoMbx  (as logged-in Windows user)"
        } else {
            Add-FbLog "  skipped — no on-prem EWS host." 'warn'
            Add-FbSkip "On-Prem → EXO" "On-prem EWS host not provided."
        }

        # Direction B: EXO → On-Prem (EXO EWS with OAuth, target = on-prem mailbox)
        Add-FbLog "" ''
        Add-FbLog "Direction B — EXO -> On-Prem: Exchange Online looks up the on-prem mailbox" 'step'
        $fbForm.Cursor = [System.Windows.Forms.Cursors]::Default
        $tok = Get-ExoEwsToken
        $fbForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        if ($tok) {
            Add-FbExplain 'B'
            [System.Windows.Forms.Application]::DoEvents()
            $rB = Invoke-FreeBusyProbe -EwsUrl 'https://outlook.office365.com/EWS/Exchange.asmx' -Target $onpMbx -Mode 'bearer' -User '' -Pass '' -Token $tok -IgnoreCert $ign
            Add-FbWhereToLook 'B' $rB
            Add-FbRow "EXO → On-Prem" $rB "Exchange Online EWS  →  $onpMbx"
        } else {
            Add-FbLog "  skipped — no access token." 'warn'
            Add-FbSkip "EXO → On-Prem" "Exchange Online sign-in cancelled or failed."
        }
        Add-FbLog "" ''
        Add-FbLog "Test finished. The Results tab holds the verdict per direction." 'step'
    } finally {
        $fbForm.Cursor = [System.Windows.Forms.Cursors]::Default
        $fbBtnTest.Enabled = $true
    }
})

# Pre-fill the on-prem mailbox from the detected UPN
if ($upn) { $fbTxtOnp.Text = $upn }
#endregion

#region ======================================================================
#  LAUNCHER
#==============================================================================

$lForm = New-Object System.Windows.Forms.Form
$lForm.Text            = "Exchange Tester"
$lForm.ClientSize      = New-Object System.Drawing.Size(408, 318)
$lForm.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
$lForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$lForm.MaximizeBox     = $false
$lForm.MinimizeBox     = $false

$lLbl = New-Object System.Windows.Forms.Label
$lLbl.Text     = "Please select a test:"
$lLbl.Location = New-Object System.Drawing.Point(12, 12)
$lLbl.AutoSize = $true
$lLbl.Font     = New-Object System.Drawing.Font($lForm.Font, [System.Drawing.FontStyle]::Bold)
$lForm.Controls.Add($lLbl)

$lBtnAuto = New-Object System.Windows.Forms.Button
$lBtnAuto.Text     = "E-Mail AutoConfiguration`nAutoDiscover test — like Outlook's 'Test E-Mail AutoConfiguration'"
$lBtnAuto.Location = New-Object System.Drawing.Point(12, 40)
$lBtnAuto.Size     = New-Object System.Drawing.Size(384, 58)
$lBtnAuto.TabIndex = 0
$lForm.Controls.Add($lBtnAuto)

$lBtnHyb = New-Object System.Windows.Forms.Button
$lBtnHyb.Text     = "Hybrid Deployment`nMigration endpoint availability — MRS Proxy, on-prem mailbox"
$lBtnHyb.Location = New-Object System.Drawing.Point(12, 104)
$lBtnHyb.Size     = New-Object System.Drawing.Size(384, 58)
$lBtnHyb.TabIndex = 1
$lForm.Controls.Add($lBtnHyb)

$lBtnHcon = New-Object System.Windows.Forms.Button
$lBtnHcon.Text     = "Hybrid Connectivity`nAutoDiscover · OAuth · Free/Busy (EWS) · vdir health — on-prem + EXO"
$lBtnHcon.Location = New-Object System.Drawing.Point(12, 168)
$lBtnHcon.Size     = New-Object System.Drawing.Size(384, 58)
$lBtnHcon.TabIndex = 2
$lForm.Controls.Add($lBtnHcon)

$lBtnFb = New-Object System.Windows.Forms.Button
$lBtnFb.Text     = "Free/Busy (Cross-Premises)`nAvailability both directions with real mailboxes — On-Prem ↔ EXO"
$lBtnFb.Location = New-Object System.Drawing.Point(12, 232)
$lBtnFb.Size     = New-Object System.Drawing.Size(384, 58)
$lBtnFb.TabIndex = 3
$lForm.Controls.Add($lBtnFb)

$script:LauncherChoice = $null
$lBtnAuto.Add_Click({ $script:LauncherChoice = 'auto';   $lForm.DialogResult = [System.Windows.Forms.DialogResult]::OK })
$lBtnHyb.Add_Click({  $script:LauncherChoice = 'hybrid'; $lForm.DialogResult = [System.Windows.Forms.DialogResult]::OK })
$lBtnHcon.Add_Click({ $script:LauncherChoice = 'hcon';   $lForm.DialogResult = [System.Windows.Forms.DialogResult]::OK })
$lBtnFb.Add_Click({   $script:LauncherChoice = 'freebusy'; $lForm.DialogResult = [System.Windows.Forms.DialogResult]::OK })

# Launcher loop: closing a test window returns to the launcher; closing the
# launcher (X) exits. Modal forms are only hidden on close, so all test
# windows can be reopened without rebuilding them.
while ($true) {
    $script:LauncherChoice = $null
    [void]$lForm.ShowDialog()
    if (-not $script:LauncherChoice) { break }
    switch ($script:LauncherChoice) {
        'auto'     { [void]$form.ShowDialog() }
        'hybrid'   { [void]$hybForm.ShowDialog() }
        'hcon'     { [void]$hcForm.ShowDialog() }
        'freebusy' { [void]$fbForm.ShowDialog() }
    }
}
#endregion
