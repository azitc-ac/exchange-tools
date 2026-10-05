param(
    # Konto für Connect-ExchangeOnline; ohne Angabe fragt die Anmeldung nach
    [string]$UserPrincipalName
)

# Einzige Stelle für die Versionsnummer; Build und Release-Tag lesen sie hier aus.
$script:Version = '1.0.0'

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

#region Bearbeiten-Dialog
function Show-ContactEditDialog([System.Data.DataRow]$Row) {
    $form = New-Object System.Windows.Forms.Form -Property @{
        Text            = "Kontakt bearbeiten - $($Row['Name'])"
        ClientSize      = New-Object System.Drawing.Size(520, 178)
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
            Width    = 345
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
    $y += 28

    $btnSave = New-Object System.Windows.Forms.Button -Property @{
        Text = 'Speichern'; Location = New-Object System.Drawing.Point(330, ($y + 8)); Width = 85 }
    $btnCancel = New-Object System.Windows.Forms.Button -Property @{
        Text = 'Abbrechen'; Location = New-Object System.Drawing.Point(420, ($y + 8)); Width = 85; DialogResult = 'Cancel' }
    $form.Controls.AddRange(@($btnSave, $btnCancel))
    $form.AcceptButton = $btnSave
    $form.CancelButton = $btnCancel

    $btnSave.Add_Click({
            $newPrimary  = $fields['PrimarySmtpAddress'].Text.Trim()
            $newExternal = $fields['ExternalEmailAddress'].Text.Trim()

            foreach ($addr in $newPrimary, $newExternal) {
                try { $null = [System.Net.Mail.MailAddress]::new($addr) }
                catch {
                    [System.Windows.Forms.MessageBox]::Show("Ungültige Adresse: '$addr'", 'Fehler', 'OK', 'Warning') | Out-Null
                    return
                }
            }

            $primaryChanged  = $newPrimary  -cne $Row['PrimarySmtpAddress']
            $externalChanged = $newExternal -cne $Row['ExternalEmailAddress']
            if (-not ($primaryChanged -or $externalChanged)) { $form.DialogResult = 'Cancel'; return }

            # EXO kennt bei Set-MailContact kein -PrimarySmtpAddress; WindowsEmailAddress setzt sie mit
            $params = @{ Identity = $Row['Guid']; ErrorAction = 'Stop' }
            if ($externalChanged) { $params.ExternalEmailAddress = $newExternal }
            if ($primaryChanged)  { $params.WindowsEmailAddress  = $newPrimary }

            $form.Cursor = 'WaitCursor'
            try {
                Set-MailContact @params
                # neu einlesen, da Exchange ggf. abhängige Werte mitändert
                $current = Get-MailContact -Identity $Row['Guid'] -ErrorAction Stop

                if ($externalChanged -and $chkAddProxy.Checked -and
                    -not ($current.EmailAddresses | Where-Object { "$_" -eq "smtp:$newExternal" })) {
                    Set-MailContact -Identity $Row['Guid'] -EmailAddresses @{ Add = "smtp:$newExternal" } -ErrorAction Stop
                    $current = Get-MailContact -Identity $Row['Guid'] -ErrorAction Stop
                }

                Set-ContactRow $Row $current
                $form.DialogResult = 'OK'
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Speichern fehlgeschlagen:`n`n$($_.Exception.Message)", 'Fehler', 'OK', 'Error') | Out-Null
            }
            finally { $form.Cursor = 'Default' }
        })

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
