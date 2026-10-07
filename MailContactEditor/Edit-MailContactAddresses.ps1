param(
    # Konto für Connect-ExchangeOnline; ohne Angabe fragt die Anmeldung nach
    [string]$UserPrincipalName
)

# Einzige Stelle für die Versionsnummer; Build und Release-Tag lesen sie hier aus.
$script:Version = '1.1.0'

# Voraussetzungen und Anmeldung. Der Block zwischen den Markern stammt aus
# build/Prolog.Exo.ps1 und wird von build/Sync-Prolog.ps1 gepflegt - nicht von Hand ändern.
$script:RequiredModuleVersion = '3.6.0'
$script:RequiredCmdlets       = @('Get-MailContact', 'Set-MailContact')
$script:ToolIsGui             = $true
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

Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$font    = New-Object System.Drawing.Font('Segoe UI', 9)
$columns = 'Name', 'Alias', 'PrimarySmtpAddress', 'ExternalEmailAddress'

function ConvertTo-PlainAddress($address) {
    "$address" -replace '^smtp:', ''
}

# --- EmailAddresses: Zerlegen und Zusammensetzen ---------------------------------
# Eine Proxy-Adresse ist "<Praefix>:<Wert>". Gross geschriebenes SMTP markiert die
# primaere Adresse; davon darf es genau eine geben. X500 und SIP stammen meist aus
# Migrationen - sie gehoeren nicht angefasst, sonst brechen alte Antwortadressen
# und die Lync/Teams-Zuordnung.
function ConvertFrom-ProxyAddress([string]$Proxy) {
    $i = $Proxy.IndexOf(':')
    if ($i -lt 1) {
        # Ohne Praefix behandelt Exchange den Wert als sekundaere SMTP-Adresse.
        return [pscustomobject]@{ Prefix = 'smtp'; Address = $Proxy; IsPrimary = $false; IsSmtp = $true }
    }
    $p = $Proxy.Substring(0, $i)
    $a = $Proxy.Substring($i + 1)
    [pscustomobject]@{
        Prefix    = $p
        Address   = $a
        IsPrimary = ($p -ceq 'SMTP')
        IsSmtp    = ($p -ieq 'smtp')
    }
}

function ConvertTo-ProxyAddress($Entry) {
    if ($Entry.IsSmtp) { "$(if ($Entry.IsPrimary) { 'SMTP' } else { 'smtp' }):$($Entry.Address)" }
    else               { "$($Entry.Prefix):$($Entry.Address)" }
}

function Test-SmtpAddress([string]$Address) {
    if ([string]::IsNullOrWhiteSpace($Address)) { return $false }
    try { $null = [System.Net.Mail.MailAddress]::new($Address); $true } catch { $false }
}

function Set-ContactRow([System.Data.DataRow]$Row, $Contact) {
    $Row['Name']                 = $Contact.Name
    $Row['Alias']                = $Contact.Alias
    $Row['PrimarySmtpAddress']   = "$($Contact.PrimarySmtpAddress)"
    $Row['ExternalEmailAddress'] = ConvertTo-PlainAddress $Contact.ExternalEmailAddress
}

function ConvertTo-LikeLiteral([string]$Text) {
    # Sonderzeichen für DataView.RowFilter (LIKE) maskieren
    $Text -replace '([\[\]\*%])', '[$1]' -replace "'", "''"
}

#region Adress-Eingabe
# Kleiner Prompt für Hinzufügen und Ändern einer SMTP-Adresse. Prüft schon hier,
# damit eine unbrauchbare Eingabe gar nicht erst in die Liste gelangt.
function Show-AddressPrompt {
    param([string]$Title, [string]$Value = '', [string[]]$Vorhanden = @())

    $dlg = New-Object System.Windows.Forms.Form -Property @{
        Text            = $Title
        ClientSize      = New-Object System.Drawing.Size(430, 108)
        StartPosition   = 'CenterParent'
        FormBorderStyle = 'FixedDialog'
        MaximizeBox     = $false
        MinimizeBox     = $false
        Font            = $font
        AutoScaleMode   = 'Font'
    }
    $dlg.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
                Text = 'SMTP-Adresse:'; Location = New-Object System.Drawing.Point(12, 18); AutoSize = $true }))
    $tb = New-Object System.Windows.Forms.TextBox -Property @{
        Text = $Value; Location = New-Object System.Drawing.Point(110, 15); Width = 308 }
    $dlg.Controls.Add($tb)

    $lblHint = New-Object System.Windows.Forms.Label -Property @{
        Location = New-Object System.Drawing.Point(110, 44); AutoSize = $true
        ForeColor = [System.Drawing.Color]::Firebrick }
    $dlg.Controls.Add($lblHint)

    $ok = New-Object System.Windows.Forms.Button -Property @{
        Text = 'OK'; Location = New-Object System.Drawing.Point(243, 70); Width = 85 }
    $ab = New-Object System.Windows.Forms.Button -Property @{
        Text = 'Abbrechen'; Location = New-Object System.Drawing.Point(333, 70); Width = 85; DialogResult = 'Cancel' }
    $dlg.Controls.AddRange(@($ok, $ab))
    $dlg.AcceptButton = $ok
    $dlg.CancelButton = $ab

    $ergebnis = $null
    $ok.Add_Click({
            $wert = $tb.Text.Trim()
            if (-not (Test-SmtpAddress $wert)) { $lblHint.Text = 'Keine gültige SMTP-Adresse.'; return }
            # Gross-/Kleinschreibung ist bei Adressen unerheblich - Dubletten lehnt Exchange ab.
            if ($Vorhanden -contains $wert.ToLowerInvariant()) { $lblHint.Text = 'Diese Adresse ist bereits eingetragen.'; return }
            $script:promptErgebnis = $wert
            $dlg.DialogResult = 'OK'
        })

    # Groesse aus dem Inhalt, gleicher Grund wie im Bearbeiten-Dialog.
    $rand   = 12
    $rechts = ($dlg.Controls | ForEach-Object { $_.Right }  | Measure-Object -Maximum).Maximum
    $unten  = ($dlg.Controls | ForEach-Object { $_.Bottom } | Measure-Object -Maximum).Maximum
    $dlg.ClientSize = New-Object System.Drawing.Size(($rechts + $rand), ($unten + $rand))

    $script:promptErgebnis = $null
    $r = $dlg.ShowDialog()
    $ergebnis = $script:promptErgebnis
    $dlg.Dispose()
    if ($r -eq 'OK') { $ergebnis } else { $null }
}
#endregion

#region Bearbeiten-Dialog
function Show-ContactEditDialog([System.Data.DataRow]$Row) {
    # Frisch lesen statt aus der Tabelle: die Adressliste kann sich seit dem Laden
    # geändert haben, und sie ist hier die Arbeitsgrundlage.
    try { $kontakt = Get-MailContact -Identity $Row['Guid'] -ErrorAction Stop }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Kontakt konnte nicht gelesen werden:`n`n$($_.Exception.Message)",
            'Fehler', 'OK', 'Error') | Out-Null
        return $false
    }

    # Arbeitskopie der Adressen; das Original wird erst beim Speichern angefasst.
    $adressen = New-Object System.Collections.ArrayList
    foreach ($p in $kontakt.EmailAddresses) { [void]$adressen.Add((ConvertFrom-ProxyAddress "$p")) }

    $form = New-Object System.Windows.Forms.Form -Property @{
        Text            = "Kontakt bearbeiten - $($Row['Name'])"
        ClientSize      = New-Object System.Drawing.Size(620, 466)
        StartPosition   = 'CenterParent'
        FormBorderStyle = 'FixedDialog'
        MaximizeBox     = $false
        MinimizeBox     = $false
        Font            = $font
        AutoScaleMode   = 'Font'
    }

    $y = 15
    $fields = @{}
    foreach ($f in @(
            @{ Key = 'Name';                 Label = 'Name:';                 ReadOnly = $true  }
            @{ Key = 'PrimarySmtpAddress';   Label = 'PrimarySmtpAddress:';   ReadOnly = $false }
            @{ Key = 'ExternalEmailAddress'; Label = 'ExternalEmailAddress:'; ReadOnly = $false })) {
        $form.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
                    Text = $f.Label; Location = New-Object System.Drawing.Point(12, ($y + 3)); AutoSize = $true }))
        $tb = New-Object System.Windows.Forms.TextBox -Property @{
            Text     = $Row[$f.Key]
            ReadOnly = $f.ReadOnly
            Location = New-Object System.Drawing.Point(160, $y)
            Width    = 448
        }
        $form.Controls.Add($tb)
        $fields[$f.Key] = $tb
        $y += 32
    }

    $chkAddProxy = New-Object System.Windows.Forms.CheckBox -Property @{
        Text     = 'Neue externe Adresse auch in EmailAddresses eintragen'
        Checked  = $true
        Location = New-Object System.Drawing.Point(160, $y)
        AutoSize = $true
    }
    $form.Controls.Add($chkAddProxy)
    $y += 30

    $form.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
                Text = 'EmailAddresses:'; Location = New-Object System.Drawing.Point(12, $y); AutoSize = $true }))
    $y += 23   # Labelhöhe 21 plus Luft; bei 20 ueberlappte die Liste das Label um 1 px

    $lv = New-Object System.Windows.Forms.ListView -Property @{
        Location      = New-Object System.Drawing.Point(12, $y)
        Size          = New-Object System.Drawing.Size(596, 196)
        View          = 'Details'
        FullRowSelect = $true
        MultiSelect   = $false
        HideSelection = $false
        GridLines     = $true
    }
    [void]$lv.Columns.Add('Typ', 70)
    [void]$lv.Columns.Add('Adresse', 500)
    $form.Controls.Add($lv)
    $y += 204

    $btnAdd     = New-Object System.Windows.Forms.Button -Property @{ Text = 'Hinzufügen'; Location = New-Object System.Drawing.Point(12, $y);  Width = 95 }
    $btnEdit    = New-Object System.Windows.Forms.Button -Property @{ Text = 'Ändern';     Location = New-Object System.Drawing.Point(113, $y); Width = 95 }
    $btnDel     = New-Object System.Windows.Forms.Button -Property @{ Text = 'Entfernen';  Location = New-Object System.Drawing.Point(214, $y); Width = 95 }
    $btnPrimary = New-Object System.Windows.Forms.Button -Property @{ Text = 'Als primär'; Location = New-Object System.Drawing.Point(315, $y); Width = 105 }
    $form.Controls.AddRange(@($btnAdd, $btnEdit, $btnDel, $btnPrimary))
    $y += 34

    # +2 statt +6: sonst sitzt das Label tiefer als die Schaltflaechen daneben und
    # verschiebt den unteren Rand des Dialogs.
    $lblInfo = New-Object System.Windows.Forms.Label -Property @{
        Location = New-Object System.Drawing.Point(12, ($y + 2)); AutoSize = $true
        ForeColor = [System.Drawing.SystemColors]::GrayText }
    $form.Controls.Add($lblInfo)

    $btnSave = New-Object System.Windows.Forms.Button -Property @{
        Text = 'Speichern'; Location = New-Object System.Drawing.Point(430, $y); Width = 85 }
    $btnCancel = New-Object System.Windows.Forms.Button -Property @{
        Text = 'Abbrechen'; Location = New-Object System.Drawing.Point(523, $y); Width = 85; DialogResult = 'Cancel' }
    $form.Controls.AddRange(@($btnSave, $btnCancel))
    $form.AcceptButton = $btnSave
    $form.CancelButton = $btnCancel

    # --- Liste und Primärfeld halten sich gegenseitig aktuell ---------------------
    function Update-AddressList {
        $lv.BeginUpdate()
        $lv.Items.Clear()
        foreach ($e in $adressen) {
            $typ = if ($e.IsSmtp) { if ($e.IsPrimary) { 'SMTP' } else { 'smtp' } } else { $e.Prefix }
            $it = New-Object System.Windows.Forms.ListViewItem($typ)
            [void]$it.SubItems.Add($e.Address)
            $it.Tag = $e
            if ($e.IsPrimary) { $it.Font = New-Object System.Drawing.Font($font, [System.Drawing.FontStyle]::Bold) }
            if (-not $e.IsSmtp) { $it.ForeColor = [System.Drawing.SystemColors]::GrayText }
            [void]$lv.Items.Add($it)
        }
        $lv.EndUpdate()

        $prim = @($adressen | Where-Object { $_.IsPrimary }) | Select-Object -First 1
        if ($prim) { $fields['PrimarySmtpAddress'].Text = $prim.Address }

        $andere = @($adressen | Where-Object { -not $_.IsSmtp }).Count
        $lblInfo.Text = "$($adressen.Count) Adressen" + $(if ($andere) { ", davon $andere nicht bearbeitbar (X500/SIP)" } else { '' })
    }

    function Get-SelectedEntry {
        if ($lv.SelectedItems.Count -eq 0) { return $null }
        $lv.SelectedItems[0].Tag
    }

    function Update-ButtonState {
        $e = Get-SelectedEntry
        $istSmtp = ($null -ne $e -and $e.IsSmtp)
        $btnEdit.Enabled    = $istSmtp
        $btnDel.Enabled     = ($istSmtp -and -not $e.IsPrimary)   # die primäre nie ersatzlos entfernen
        $btnPrimary.Enabled = ($istSmtp -and -not $e.IsPrimary)
    }

    $lv.Add_SelectedIndexChanged({ Update-ButtonState })

    $btnAdd.Add_Click({
            $vorhanden = @($adressen | ForEach-Object { $_.Address.ToLowerInvariant() })
            $neu = Show-AddressPrompt -Title 'Adresse hinzufügen' -Vorhanden $vorhanden
            if (-not $neu) { return }
            [void]$adressen.Add([pscustomobject]@{ Prefix = 'smtp'; Address = $neu; IsPrimary = $false; IsSmtp = $true })
            Update-AddressList; Update-ButtonState
        })

    $btnEdit.Add_Click({
            $e = Get-SelectedEntry
            if (-not $e -or -not $e.IsSmtp) { return }
            $vorhanden = @($adressen | Where-Object { $_ -ne $e } | ForEach-Object { $_.Address.ToLowerInvariant() })
            $neu = Show-AddressPrompt -Title 'Adresse ändern' -Value $e.Address -Vorhanden $vorhanden
            if (-not $neu) { return }
            $e.Address = $neu
            Update-AddressList; Update-ButtonState
        })

    $btnDel.Add_Click({
            $e = Get-SelectedEntry
            if (-not $e -or -not $e.IsSmtp -or $e.IsPrimary) { return }
            $adressen.Remove($e)
            Update-AddressList; Update-ButtonState
        })

    $btnPrimary.Add_Click({
            $e = Get-SelectedEntry
            if (-not $e -or -not $e.IsSmtp) { return }
            foreach ($a in $adressen) { if ($a.IsSmtp) { $a.IsPrimary = $false } }
            $e.IsPrimary = $true
            Update-AddressList; Update-ButtonState
        })

    # Primärfeld geändert -> die primäre Adresse in der Liste zieht nach.
    $fields['PrimarySmtpAddress'].Add_Leave({
            $wert = $fields['PrimarySmtpAddress'].Text.Trim()
            if (-not (Test-SmtpAddress $wert)) { return }
            $prim = @($adressen | Where-Object { $_.IsPrimary }) | Select-Object -First 1
            if ($prim -and $prim.Address -ieq $wert) { return }
            # Steht der Wert schon als sekundäre Adresse drin, wird sie befördert
            # statt ein zweites Mal angelegt.
            $treffer = @($adressen | Where-Object { $_.IsSmtp -and $_.Address -ieq $wert }) | Select-Object -First 1
            foreach ($a in $adressen) { if ($a.IsSmtp) { $a.IsPrimary = $false } }
            if ($treffer) { $treffer.IsPrimary = $true }
            else { [void]$adressen.Add([pscustomobject]@{ Prefix = 'SMTP'; Address = $wert; IsPrimary = $true; IsSmtp = $true }) }
            Update-AddressList; Update-ButtonState
        })

    # Externe Adresse geändert -> auf Wunsch gleich sichtbar in die Liste, statt
    # sie nach dem Speichern unbemerkt nachzutragen.
    $fields['ExternalEmailAddress'].Add_Leave({
            if (-not $chkAddProxy.Checked) { return }
            $wert = $fields['ExternalEmailAddress'].Text.Trim()
            if (-not (Test-SmtpAddress $wert)) { return }
            if (@($adressen | Where-Object { $_.IsSmtp -and $_.Address -ieq $wert }).Count) { return }
            [void]$adressen.Add([pscustomobject]@{ Prefix = 'smtp'; Address = $wert; IsPrimary = $false; IsSmtp = $true })
            Update-AddressList; Update-ButtonState
        })

    $btnSave.Add_Click({
            $newExternal = $fields['ExternalEmailAddress'].Text.Trim()
            if (-not (Test-SmtpAddress $newExternal)) {
                [System.Windows.Forms.MessageBox]::Show("Ungültige Adresse: '$newExternal'", 'Fehler', 'OK', 'Warning') | Out-Null
                return
            }
            $prim = @($adressen | Where-Object { $_.IsPrimary })
            if ($prim.Count -ne 1) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Es muss genau eine primäre Adresse geben, gefunden: $($prim.Count).", 'Fehler', 'OK', 'Warning') | Out-Null
                return
            }

            $neueListe = @($adressen | ForEach-Object { ConvertTo-ProxyAddress $_ })
            $alteListe = @($kontakt.EmailAddresses | ForEach-Object { "$_" })
            $adressenGeaendert = @(Compare-Object $alteListe $neueListe -CaseSensitive).Count -gt 0
            $externalChanged   = $newExternal -cne (ConvertTo-PlainAddress $kontakt.ExternalEmailAddress)

            if (-not ($adressenGeaendert -or $externalChanged)) { $form.DialogResult = 'Cancel'; return }

            # EXO kennt bei Set-MailContact kein -PrimarySmtpAddress. Die vollständige
            # Sammlung mit genau einem gross geschriebenen SMTP: setzt sie mit - und
            # bildet zugleich Hinzufügen, Ändern und Entfernen in einem Schritt ab.
            $params = @{ Identity = $Row['Guid']; ErrorAction = 'Stop' }
            if ($externalChanged)   { $params.ExternalEmailAddress = $newExternal }
            if ($adressenGeaendert) { $params.EmailAddresses       = $neueListe }

            $form.Cursor = 'WaitCursor'
            try {
                Set-MailContact @params
                # neu einlesen, da Exchange ggf. abhängige Werte mitändert
                $current = Get-MailContact -Identity $Row['Guid'] -ErrorAction Stop
                Set-ContactRow $Row $current
                $form.DialogResult = 'OK'
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Speichern fehlgeschlagen:`n`n$($_.Exception.Message)", 'Fehler', 'OK', 'Error') | Out-Null
            }
            finally { $form.Cursor = 'Default' }
        })

    # Fenstergroesse aus dem Inhalt ableiten statt sie zu raten: mit
    # AutoScaleMode='Font' skaliert WinForms die Form, nicht aber fest
    # positionierte Controls - eine gesetzte ClientSize liess je nach
    # Bildschirmskalierung einen breiten leeren Streifen rechts und unten stehen.
    # So bleibt ringsum derselbe Rand, unabhaengig von DPI und Schriftgroesse.
    $rand   = 12
    $rechts = ($form.Controls | ForEach-Object { $_.Right }  | Measure-Object -Maximum).Maximum
    $unten  = ($form.Controls | ForEach-Object { $_.Bottom } | Measure-Object -Maximum).Maximum
    $form.ClientSize = New-Object System.Drawing.Size(($rechts + $rand), ($unten + $rand))

    $form.Add_Shown({ Update-AddressList; Update-ButtonState })

    $result = $form.ShowDialog()
    $form.Dispose()
    $result -eq 'OK'
}
#endregion

#region Daten laden
$table = New-Object System.Data.DataTable
$columns + 'Guid' | ForEach-Object { [void]$table.Columns.Add($_) }

foreach ($contact in Get-MailContact -ResultSize Unlimited) {
    $row = $table.NewRow()
    $row['Guid'] = "$($contact.Guid)"
    Set-ContactRow $row $contact
    $table.Rows.Add($row)
}
#endregion

#region Auswahl-Dialog
$mainForm = New-Object System.Windows.Forms.Form -Property @{
    Text          = 'Zu ändernden Kontakt auswählen'
    ClientSize    = New-Object System.Drawing.Size(1100, 600)
    MinimumSize   = New-Object System.Drawing.Size(600, 300)
    StartPosition = 'CenterScreen'
    Font          = $font
    AutoScaleMode = 'Font'
}

$grid = New-Object System.Windows.Forms.DataGridView -Property @{
    Dock                  = 'Fill'
    DataSource            = $table.DefaultView
    ReadOnly              = $true
    AllowUserToAddRows    = $false
    AllowUserToDeleteRows = $false
    SelectionMode         = 'FullRowSelect'
    MultiSelect           = $false
    RowHeadersVisible     = $false
    AutoSizeColumnsMode   = 'Fill'
    BackgroundColor       = [System.Drawing.SystemColors]::Window
}

$topPanel = New-Object System.Windows.Forms.Panel -Property @{ Dock = 'Top'; Size = New-Object System.Drawing.Size(1100, 38) }
$topPanel.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
            Text = 'Filter:'; Location = New-Object System.Drawing.Point(10, 11); AutoSize = $true }))
$txtFilter = New-Object System.Windows.Forms.TextBox -Property @{
    Location = New-Object System.Drawing.Point(60, 8)
    Width    = 1000
    Anchor   = 'Top, Left, Right'
}
$btnClearFilter = New-Object System.Windows.Forms.Button -Property @{
    Text     = '✕'
    Location = New-Object System.Drawing.Point(1064, 7)
    Size     = New-Object System.Drawing.Size(26, 25)
    Anchor   = 'Top, Right'
    TabStop  = $false
}
$topPanel.Controls.AddRange(@($txtFilter, $btnClearFilter))

$bottomPanel = New-Object System.Windows.Forms.Panel -Property @{ Dock = 'Bottom'; Size = New-Object System.Drawing.Size(1100, 42) }
$lblCount = New-Object System.Windows.Forms.Label -Property @{
    Location = New-Object System.Drawing.Point(10, 14); AutoSize = $true }
$lblStatus = New-Object System.Windows.Forms.Label -Property @{
    Location  = New-Object System.Drawing.Point(220, 14)
    AutoSize  = $true
    ForeColor = [System.Drawing.Color]::Green
    Font      = New-Object System.Drawing.Font($font, [System.Drawing.FontStyle]::Bold)
}
$statusTimer = New-Object System.Windows.Forms.Timer -Property @{ Interval = 5000 }
$statusTimer.Add_Tick({ $statusTimer.Stop(); $lblStatus.Text = '' })
$btnEdit = New-Object System.Windows.Forms.Button -Property @{
    Text = 'OK'; Width = 100; Location = New-Object System.Drawing.Point(885, 8); Anchor = 'Bottom, Right' }
$btnClose = New-Object System.Windows.Forms.Button -Property @{
    Text = 'Abbrechen'; Width = 100; Location = New-Object System.Drawing.Point(990, 8); Anchor = 'Bottom, Right'; DialogResult = 'Cancel' }
$bottomPanel.Controls.AddRange(@($lblCount, $lblStatus, $btnEdit, $btnClose))

# Reihenfolge wichtig: Fill-Control zuerst hinzufügen, damit Top/Bottom korrekt docken
$mainForm.Controls.AddRange(@($grid, $topPanel, $bottomPanel))
$mainForm.CancelButton = $btnClose

function Update-Count {
    $lblCount.Text = "$($table.DefaultView.Count) von $($table.Rows.Count) Kontakten"
}

function Edit-SelectedContact {
    if ($null -eq $grid.CurrentRow) { return }
    $dataRow = $grid.CurrentRow.DataBoundItem.Row
    if (Show-ContactEditDialog $dataRow) {
        Update-Count
        # kurzes, nicht-blockierendes Erfolgs-Feedback; Fehler kommen als MessageBox aus dem Dialog
        $lblStatus.Text = "✔ Kontakt '$($dataRow['Name'])' gespeichert"
        $statusTimer.Stop()
        $statusTimer.Start()
    }
}

$txtFilter.Add_TextChanged({
        $t = ConvertTo-LikeLiteral $txtFilter.Text
        $table.DefaultView.RowFilter = if ($t) { ($columns | ForEach-Object { "[$_] LIKE '%$t%'" }) -join ' OR ' } else { '' }
        Update-Count
    })
$txtFilter.Add_KeyDown({
        if ($_.KeyCode -eq 'Escape' -and $txtFilter.Text) { $txtFilter.Clear(); $_.SuppressKeyPress = $true }
        elseif ($_.KeyCode -eq 'Down') { $grid.Focus(); $_.Handled = $true }
        elseif ($_.KeyCode -eq 'Enter') { Edit-SelectedContact; $_.SuppressKeyPress = $true }
    })
$btnClearFilter.Add_Click({ $txtFilter.Clear(); $txtFilter.Focus() })
$grid.Add_CellDoubleClick({ if ($_.RowIndex -ge 0) { Edit-SelectedContact } })
$grid.Add_KeyDown({ if ($_.KeyCode -eq 'Enter') { Edit-SelectedContact; $_.Handled = $true } })
$btnEdit.Add_Click({ Edit-SelectedContact })

$mainForm.Add_Shown({
        $grid.Columns['Guid'].Visible = $false
        Update-Count
        $txtFilter.Focus()
    })

[void]$mainForm.ShowDialog()
$statusTimer.Dispose()
$mainForm.Dispose()
#endregion
