# Voraussetzungen. Bis hierher rief das Skript Get-Queue und Get-Message auf, ohne zu
# prüfen, ob die Exchange-Verwaltungsshell überhaupt geladen ist - außerhalb der EMS
# brach es mit "Get-Queue wird nicht erkannt" ab. Der Block zwischen den Markern stammt
# aus build/Prolog.OnPrem.ps1 und wird von build/Sync-Prolog.ps1 gepflegt.
$script:RequiredCmdlets = @('Get-Queue', 'Get-Message')
$script:ToolIsGui       = $true
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

function Open-QueueViewerDialog {
    param (
        [string]$Title = "Exchange Queue Viewer"
    )

    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore

    # ---------- Window ----------
    $window = New-Object Windows.Window
    $window.Title  = $Title
    $window.Width  = 1100
    $window.Height = 650
    $window.MinWidth  = 800
    $window.MinHeight = 400
    $window.WindowStartupLocation = 'CenterScreen'

    # ---------- Root Grid: Row 0 = Toolbar (Auto), Row 1 = Action Bar (Auto), Row 2 = DataGrid (*) ----------
    $root = New-Object Windows.Controls.Grid
    $rowTop = New-Object Windows.Controls.RowDefinition
    $rowTop.Height = [Windows.GridLength]::Auto
    $root.RowDefinitions.Add($rowTop)
    $rowAction = New-Object Windows.Controls.RowDefinition
    $rowAction.Height = [Windows.GridLength]::Auto
    $root.RowDefinitions.Add($rowAction)
    $rowMain = New-Object Windows.Controls.RowDefinition
    $rowMain.Height = New-Object Windows.GridLength(1, [Windows.GridUnitType]::Star)
    $root.RowDefinitions.Add($rowMain)

    # ===========================
    # Toolbar
    # Col: 0=Label | 1=ComboBox(200) | 2=Spacer(12) | 3=Connect | 4=Refresh |
    #      5=Spacer(12) | 6=AutoRefreshChk | 7=IntervalTxt(44) | 8="s" | 9=Status(*)
    # ===========================
    $toolbar = New-Object Windows.Controls.Grid
    $toolbar.Margin = "10,8,10,6"

    foreach ($w in @('Auto', 200, 12, 'Auto', 'Auto', 12, 'Auto', 44, 'Auto', 'Star')) {
        $col = New-Object Windows.Controls.ColumnDefinition
        switch ($w) {
            'Auto' { $col.Width = [Windows.GridLength]::Auto }
            'Star' { $col.Width = New-Object Windows.GridLength(1, [Windows.GridUnitType]::Star) }
            default { $col.Width = New-Object Windows.GridLength($w) }
        }
        $toolbar.ColumnDefinitions.Add($col) | Out-Null
    }

    $lbl = New-Object Windows.Controls.TextBlock
    $lbl.Text = "Server:"
    $lbl.VerticalAlignment = 'Center'
    $lbl.Margin = "0,0,6,0"
    $toolbar.Children.Add($lbl) | Out-Null
    [Windows.Controls.Grid]::SetColumn($lbl, 0)

    $combo = New-Object Windows.Controls.ComboBox
    $combo.IsEditable = $true
    $combo.VerticalAlignment = 'Center'
    $combo.MinHeight = 24
    $combo.MaxDropDownHeight = 300
    $toolbar.Children.Add($combo) | Out-Null
    [Windows.Controls.Grid]::SetColumn($combo, 1)

    $btnConnect = New-Object Windows.Controls.Button
    $btnConnect.Content = "Connect"
    $btnConnect.MinWidth = 100; $btnConnect.Height = 26; $btnConnect.Margin = "0,0,8,0"
    $toolbar.Children.Add($btnConnect) | Out-Null
    [Windows.Controls.Grid]::SetColumn($btnConnect, 3)

    $btnRefresh = New-Object Windows.Controls.Button
    $btnRefresh.Content = "Refresh"
    $btnRefresh.MinWidth = 90; $btnRefresh.Height = 26
    $toolbar.Children.Add($btnRefresh) | Out-Null
    [Windows.Controls.Grid]::SetColumn($btnRefresh, 4)

    $chkAutoRefresh = New-Object Windows.Controls.CheckBox
    $chkAutoRefresh.Content = "Auto-Refresh"
    $chkAutoRefresh.VerticalAlignment = 'Center'
    $chkAutoRefresh.Margin = "0,0,4,0"
    $toolbar.Children.Add($chkAutoRefresh) | Out-Null
    [Windows.Controls.Grid]::SetColumn($chkAutoRefresh, 6)

    $txtInterval = New-Object Windows.Controls.TextBox
    $txtInterval.Text = "30"
    $txtInterval.Height = 24
    $txtInterval.VerticalAlignment = 'Center'
    $txtInterval.TextAlignment = 'Center'
    $txtInterval.Margin = "0,0,2,0"
    $toolbar.Children.Add($txtInterval) | Out-Null
    [Windows.Controls.Grid]::SetColumn($txtInterval, 7)

    $lblSec = New-Object Windows.Controls.TextBlock
    $lblSec.Text = "s"
    $lblSec.VerticalAlignment = 'Center'
    $lblSec.Margin = "2,0,0,0"
    $toolbar.Children.Add($lblSec) | Out-Null
    [Windows.Controls.Grid]::SetColumn($lblSec, 8)

    $txtStatus = New-Object Windows.Controls.TextBlock
    $txtStatus.Margin = "12,0,0,0"
    $txtStatus.VerticalAlignment = 'Center'
    $txtStatus.Opacity = 0.8
    $txtStatus.Text = ""
    $txtStatus.TextTrimming = 'CharacterEllipsis'
    $toolbar.Children.Add($txtStatus) | Out-Null
    [Windows.Controls.Grid]::SetColumn($txtStatus, 9)

    # ===========================
    # Queue Action Bar
    # ===========================
    $actionBar = New-Object Windows.Controls.StackPanel
    $actionBar.Orientation = 'Horizontal'
    $actionBar.Margin = "10,0,10,6"

    $btnSuspendQueue = New-Object Windows.Controls.Button
    $btnSuspendQueue.Content = "Suspend Queue"
    $btnSuspendQueue.MinWidth = 100; $btnSuspendQueue.Height = 26; $btnSuspendQueue.Margin = "0,0,6,0"
    $btnSuspendQueue.IsEnabled = $false
    $actionBar.Children.Add($btnSuspendQueue) | Out-Null

    $btnResumeQueue = New-Object Windows.Controls.Button
    $btnResumeQueue.Content = "Resume Queue"
    $btnResumeQueue.MinWidth = 100; $btnResumeQueue.Height = 26; $btnResumeQueue.Margin = "0,0,6,0"
    $btnResumeQueue.IsEnabled = $false
    $actionBar.Children.Add($btnResumeQueue) | Out-Null

    $btnRetryQueue = New-Object Windows.Controls.Button
    $btnRetryQueue.Content = "Retry Queue"
    $btnRetryQueue.MinWidth = 90; $btnRetryQueue.Height = 26; $btnRetryQueue.Margin = "0,0,14,0"
    $btnRetryQueue.IsEnabled = $false
    $actionBar.Children.Add($btnRetryQueue) | Out-Null

    $btnViewMsgs = New-Object Windows.Controls.Button
    $btnViewMsgs.Content = "View Messages..."
    $btnViewMsgs.MinWidth = 110; $btnViewMsgs.Height = 26; $btnViewMsgs.Margin = "0,0,14,0"
    $btnViewMsgs.IsEnabled = $false
    $actionBar.Children.Add($btnViewMsgs) | Out-Null

    $btnDeleteAllMsgs = New-Object Windows.Controls.Button
    $btnDeleteAllMsgs.Content = "Delete All Messages..."
    $btnDeleteAllMsgs.MinWidth = 140; $btnDeleteAllMsgs.Height = 26
    $btnDeleteAllMsgs.IsEnabled = $false
    $actionBar.Children.Add($btnDeleteAllMsgs) | Out-Null

    # ===========================
    # DataGrid
    # ===========================
    $dataGrid = New-Object Windows.Controls.DataGrid
    $dataGrid.Margin = "10,0,10,10"
    $dataGrid.AutoGenerateColumns = $false
    $dataGrid.CanUserSortColumns  = $true
    $dataGrid.IsReadOnly = $true
    $dataGrid.HeadersVisibility = 'Column'
    $dataGrid.GridLinesVisibility = 'Horizontal'
    $dataGrid.RowHeaderWidth = 0
    $dataGrid.AlternationCount = 2
    $dataGrid.SelectionMode = 'Extended'
    $dataGrid.SelectionUnit = 'FullRow'
    $dataGrid.EnableRowVirtualization = $true
    $dataGrid.EnableColumnVirtualization = $true
    $dataGrid.ColumnWidth = 'SizeToCells'
    $dataGrid.MinColumnWidth = 80
    $dataGrid.ItemsSource = @()

    # ===========================
    # Context Menu
    # ===========================
    $contextMenu = New-Object System.Windows.Controls.ContextMenu

    $miSuspend = New-Object System.Windows.Controls.MenuItem
    $miSuspend.Header = "Suspend Queue"; $miSuspend.IsEnabled = $false
    $contextMenu.Items.Add($miSuspend) | Out-Null

    $miResume = New-Object System.Windows.Controls.MenuItem
    $miResume.Header = "Resume Queue"; $miResume.IsEnabled = $false
    $contextMenu.Items.Add($miResume) | Out-Null

    $miRetry = New-Object System.Windows.Controls.MenuItem
    $miRetry.Header = "Retry Queue"; $miRetry.IsEnabled = $false
    $contextMenu.Items.Add($miRetry) | Out-Null

    $contextMenu.Items.Add((New-Object System.Windows.Controls.Separator)) | Out-Null

    $miView = New-Object System.Windows.Controls.MenuItem
    $miView.Header = "View Messages..."
    $miView.IsEnabled = $false
    $contextMenu.Items.Add($miView) | Out-Null

    $contextMenu.Items.Add((New-Object System.Windows.Controls.Separator)) | Out-Null

    $miDelete = New-Object System.Windows.Controls.MenuItem
    $miDelete.Header = "Delete All Messages..."; $miDelete.IsEnabled = $false
    $contextMenu.Items.Add($miDelete) | Out-Null

    $dataGrid.ContextMenu = $contextMenu

    $dataGrid.Add_SelectionChanged({
        $count  = $dataGrid.SelectedItems.Count
        $any    = $count -gt 0
        $single = $count -eq 1
        $btnSuspendQueue.IsEnabled  = $any
        $btnResumeQueue.IsEnabled   = $any
        $btnRetryQueue.IsEnabled    = $any
        $btnViewMsgs.IsEnabled      = $single
        $btnDeleteAllMsgs.IsEnabled = $any
        $miSuspend.IsEnabled  = $any
        $miResume.IsEnabled   = $any
        $miRetry.IsEnabled    = $any
        $miView.IsEnabled     = $single
        $miDelete.IsEnabled   = $any
    })

    # ===========================
    # Column helper
    # ===========================
    $createColumns = {
        param($sample)
        $dataGrid.Columns.Clear()
        if ($null -eq $sample) { return }
        foreach ($prop in $sample.PSObject.Properties.Name) {
            $col = New-Object Windows.Controls.DataGridTextColumn
            $col.Header  = $prop
            $col.Binding = New-Object Windows.Data.Binding($prop)
            $col.CanUserSort = $true

            if ($prop -eq "Identity") { $col.Visibility = "Collapsed" }

            switch ($prop) {
                'NextHopDomain' { $col.Width = 220 }
                'DeliveryType'  { $col.Width = 140 }
                'Status'        { $col.Width = 110 }
                'MessageCount'  { $col.Width = 100 }
                'NextRetryTime' { $col.Width = 160 }
                'LastError'     { $col.Width = New-Object Windows.Controls.DataGridLength(1, [Windows.Controls.DataGridLengthUnitType]::Star) }
                default         { $col.Width = [Windows.Controls.DataGridLength]::Auto }
            }
            [void]$dataGrid.Columns.Add($col)
        }
    }

    # ===========================
    # Load queues helper
    # ===========================
    $flushUI = { $window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [Action]{}) }

    $loadQueues = {
        param([string]$server)
        $txtStatus.Text = "Connecting to $server..."
        & $flushUI
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        try {
            $txtStatus.Text = "Loading queues on $server..."
            & $flushUI
            $localFQDN = "$($env:COMPUTERNAME).$((Get-WmiObject Win32_ComputerSystem).Domain)"
            if ($server -and ($server -eq $env:COMPUTERNAME -or $server -eq $localFQDN)) {
                $qs = Get-Queue
            } else {
                $qs = Get-Queue -Server $server
            }

            $txtStatus.Text = "Processing data..."
            & $flushUI
            $data = @($qs | Select-Object Identity, NextHopDomain, DeliveryType, Status, MessageCount, NextRetryTime, LastError)

            if ($data.Count -gt 0) {
                & $createColumns $data[0]
            } else {
                $dataGrid.Columns.Clear()
            }

            $dataGrid.ItemsSource = @($data)
            $txtStatus.Text = "$(Get-Date -Format 'HH:mm:ss')  |  Queues: $($data.Count)  |  Server: $server"
        }
        catch {
            $dataGrid.ItemsSource = @()
            $dataGrid.Columns.Clear()
            $txtStatus.Text = "Error: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show(
                "Failed to load queues on $server`n$($_.Exception.Message)",
                "Get-Queue Error", 'OK', 'Error'
            ) | Out-Null
        }
        finally {
            $window.Cursor = [System.Windows.Input.Cursors]::Arrow
        }
    }

    # ===========================
    # Auto-Refresh Timer
    # ===========================
    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromSeconds(30)
    $timer.Add_Tick({
        $server = if ($combo.SelectedItem) { $combo.SelectedItem } else { $combo.Text }
        if ($server) { & $loadQueues $server }
    })

    $chkAutoRefresh.Add_Checked({
        $interval = 30
        if ([int]::TryParse($txtInterval.Text, [ref]$interval) -and $interval -ge 5) {
            $timer.Interval = [TimeSpan]::FromSeconds($interval)
        }
        $timer.Start()
    })
    $chkAutoRefresh.Add_Unchecked({ $timer.Stop() })

    $txtInterval.Add_TextChanged({
        if ($chkAutoRefresh.IsChecked -ne $true) { return }
        $interval = 30
        if ([int]::TryParse($txtInterval.Text, [ref]$interval) -and $interval -ge 5) {
            $timer.Interval = [TimeSpan]::FromSeconds($interval)
        }
    })

    # ===========================
    # Queue Action Handlers (shared by toolbar buttons + context menu)
    # ===========================
    $doSuspendQueue = {
        $item = $dataGrid.SelectedItem
        if ($null -eq $item) { return }
        $id = $item.Identity.ToString()
        try {
            Suspend-Queue -Identity $id -Confirm:$false
            $txtStatus.Text = "Queue suspended: $id"
            $server = if ($combo.SelectedItem) { $combo.SelectedItem } else { $combo.Text }
            & $loadQueues $server
        }
        catch { [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", 'OK', 'Error') | Out-Null }
    }
    $miSuspend.Add_Click({ & $doSuspendQueue })
    $btnSuspendQueue.Add_Click({ & $doSuspendQueue })

    $doResumeQueue = {
        $item = $dataGrid.SelectedItem
        if ($null -eq $item) { return }
        $id = $item.Identity.ToString()
        try {
            Resume-Queue -Identity $id -Confirm:$false
            $txtStatus.Text = "Queue resumed: $id"
            $server = if ($combo.SelectedItem) { $combo.SelectedItem } else { $combo.Text }
            & $loadQueues $server
        }
        catch { [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", 'OK', 'Error') | Out-Null }
    }
    $miResume.Add_Click({ & $doResumeQueue })
    $btnResumeQueue.Add_Click({ & $doResumeQueue })

    $doRetryQueue = {
        $item = $dataGrid.SelectedItem
        if ($null -eq $item) { return }
        $id = $item.Identity.ToString()
        try {
            Retry-Queue -Identity $id -Resubmit $false -Confirm:$false
            $txtStatus.Text = "Queue retry triggered: $id"
            $server = if ($combo.SelectedItem) { $combo.SelectedItem } else { $combo.Text }
            & $loadQueues $server
        }
        catch { [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", 'OK', 'Error') | Out-Null }
    }
    $miRetry.Add_Click({ & $doRetryQueue })
    $btnRetryQueue.Add_Click({ & $doRetryQueue })

    # Shared: open message window (context menu + double-click)
    $doViewMessages = {
        $row = $dataGrid.SelectedItem
        if ($null -eq $row) { return }
        $queueId = $row.Identity.ToString()

        try {
            $messages = @(Get-Message -Queue $queueId |
                Select-Object Identity, Subject, FromAddress, Status, Size, MessageLatency, LastError, DateReceived)

            $msgWin = New-Object Windows.Window
            $msgWin.Title = "Messages in $queueId  ($($messages.Count))"
            $msgWin.Width = 980
            $msgWin.Height = 600
            $msgWin.MinWidth = 600
            $msgWin.MinHeight = 300
            $msgWin.WindowStartupLocation = "CenterOwner"
            $msgWin.Owner = $window

            $msgRoot = New-Object Windows.Controls.Grid
            $msgRoot.RowDefinitions.Add((New-Object Windows.Controls.RowDefinition)) | Out-Null
            $msgRoot.RowDefinitions[0].Height = [Windows.GridLength]::Auto
            $msgRoot.RowDefinitions.Add((New-Object Windows.Controls.RowDefinition)) | Out-Null
            $msgRoot.RowDefinitions[1].Height = New-Object Windows.GridLength(1, [Windows.GridUnitType]::Star)

            $msgBar = New-Object Windows.Controls.StackPanel
            $msgBar.Orientation = 'Horizontal'
            $msgBar.Margin = "10,8,10,6"

            $btnSuspendMsg = New-Object Windows.Controls.Button
            $btnSuspendMsg.Content = "Suspend"
            $btnSuspendMsg.MinWidth = 80; $btnSuspendMsg.Height = 26; $btnSuspendMsg.Margin = "0,0,6,0"
            $btnSuspendMsg.IsEnabled = $false
            $msgBar.Children.Add($btnSuspendMsg) | Out-Null

            $btnResumeMsg = New-Object Windows.Controls.Button
            $btnResumeMsg.Content = "Resume"
            $btnResumeMsg.MinWidth = 80; $btnResumeMsg.Height = 26; $btnResumeMsg.Margin = "0,0,14,0"
            $btnResumeMsg.IsEnabled = $false
            $msgBar.Children.Add($btnResumeMsg) | Out-Null

            $btnViewDetails = New-Object Windows.Controls.Button
            $btnViewDetails.Content = "View Details..."
            $btnViewDetails.MinWidth = 100; $btnViewDetails.Height = 26; $btnViewDetails.Margin = "0,0,14,0"
            $msgBar.Children.Add($btnViewDetails) | Out-Null

            $btnDeleteNdr = New-Object Windows.Controls.Button
            $btnDeleteNdr.Content = "Delete with NDR..."
            $btnDeleteNdr.MinWidth = 110; $btnDeleteNdr.Height = 26; $btnDeleteNdr.Margin = "0,0,6,0"
            $btnDeleteNdr.IsEnabled = $false
            $msgBar.Children.Add($btnDeleteNdr) | Out-Null

            $btnDeleteNoNdr = New-Object Windows.Controls.Button
            $btnDeleteNoNdr.Content = "Delete without NDR..."
            $btnDeleteNoNdr.MinWidth = 130; $btnDeleteNoNdr.Height = 26; $btnDeleteNoNdr.Margin = "0,0,14,0"
            $btnDeleteNoNdr.IsEnabled = $false
            $msgBar.Children.Add($btnDeleteNoNdr) | Out-Null

            $btnExportCsv = New-Object Windows.Controls.Button
            $btnExportCsv.Content = "Export CSV..."
            $btnExportCsv.MinWidth = 90; $btnExportCsv.Height = 26
            $msgBar.Children.Add($btnExportCsv) | Out-Null

            $msgRoot.Children.Add($msgBar) | Out-Null
            [Windows.Controls.Grid]::SetRow($msgBar, 0)

            $msgGrid = New-Object Windows.Controls.DataGrid
            $msgGrid.AutoGenerateColumns = $true
            $msgGrid.IsReadOnly = $true
            $msgGrid.Margin = "10,0,10,10"
            $msgGrid.SelectionMode = 'Extended'
            $msgGrid.SelectionUnit = 'FullRow'
            $msgGrid.EnableRowVirtualization = $true
            $msgGrid.ItemsSource = $messages

            $msgRoot.Children.Add($msgGrid) | Out-Null
            [Windows.Controls.Grid]::SetRow($msgGrid, 1)

            $reloadMessages = {
                $updated = @(Get-Message -Queue $queueId |
                    Select-Object Identity, Subject, FromAddress, Status, Size, MessageLatency, LastError, DateReceived)
                $msgGrid.ItemsSource = $updated
                $msgWin.Title = "Messages in $queueId  ($($updated.Count))"
            }

            $btnViewDetails.IsEnabled = $false

            $msgGrid.Add_SelectionChanged({
                $count  = $msgGrid.SelectedItems.Count
                $any    = $count -gt 0
                $single = $count -eq 1
                $btnViewDetails.IsEnabled   = $single
                $btnSuspendMsg.IsEnabled    = $any
                $btnResumeMsg.IsEnabled     = $any
                $btnDeleteNdr.IsEnabled     = $any
                $btnDeleteNoNdr.IsEnabled   = $any
                $miMsgViewDetails.IsEnabled = $single
                $miMsgSuspend.IsEnabled     = $any
                $miMsgResume.IsEnabled      = $any
                $miMsgDelNdr.IsEnabled      = $any
                $miMsgDelNoNdr.IsEnabled    = $any
            })

            # Shared: open detail window (button + context menu + double-click)
            $doViewDetails = {
                $msg = $msgGrid.SelectedItem
                if ($null -eq $msg) { return }

                        $detailWin = New-Object Windows.Window
                        $detailWin.Title = "Message Details"
                        $detailWin.Width = 660
                        $detailWin.Height = 460
                        $detailWin.MinWidth = 400
                        $detailWin.MinHeight = 300
                        $detailWin.WindowStartupLocation = "CenterOwner"
                        $detailWin.Owner = $msgWin

                        $detailRoot = New-Object Windows.Controls.Grid
                        $detailRoot.RowDefinitions.Add((New-Object Windows.Controls.RowDefinition)) | Out-Null
                        $detailRoot.RowDefinitions[0].Height = [Windows.GridLength]::Auto
                        $detailRoot.RowDefinitions.Add((New-Object Windows.Controls.RowDefinition)) | Out-Null
                        $detailRoot.RowDefinitions[1].Height = New-Object Windows.GridLength(1, [Windows.GridUnitType]::Star)

                        $detailBar = New-Object Windows.Controls.StackPanel
                        $detailBar.Orientation = 'Horizontal'
                        $detailBar.Margin = "10,8,10,6"
                        $btnExportEml = New-Object Windows.Controls.Button
                        $btnExportEml.Content = "Export as .eml..."
                        $btnExportEml.MinWidth = 150; $btnExportEml.Height = 26
                        $detailBar.Children.Add($btnExportEml) | Out-Null
                        $detailRoot.Children.Add($detailBar) | Out-Null
                        [Windows.Controls.Grid]::SetRow($detailBar, 0)

                        $scroll = New-Object Windows.Controls.ScrollViewer
                        $scroll.Margin = "10,0,10,10"
                        $scroll.VerticalScrollBarVisibility = 'Auto'

                        $propGrid = New-Object Windows.Controls.Grid
                        $propGrid.ColumnDefinitions.Add((New-Object Windows.Controls.ColumnDefinition)) | Out-Null
                        $propGrid.ColumnDefinitions[0].Width = New-Object Windows.GridLength(150)
                        $propGrid.ColumnDefinitions.Add((New-Object Windows.Controls.ColumnDefinition)) | Out-Null
                        $propGrid.ColumnDefinitions[1].Width = New-Object Windows.GridLength(1, [Windows.GridUnitType]::Star)

                        $rowIdx = 0
                        foreach ($prop in $msg.PSObject.Properties) {
                            $rdef = New-Object Windows.Controls.RowDefinition
                            $rdef.Height = [Windows.GridLength]::Auto
                            $propGrid.RowDefinitions.Add($rdef) | Out-Null

                            $lblP = New-Object Windows.Controls.TextBlock
                            $lblP.Text = "$($prop.Name):"
                            $lblP.FontWeight = 'SemiBold'
                            $lblP.Margin = "4,4,8,4"
                            $lblP.VerticalAlignment = 'Top'
                            $propGrid.Children.Add($lblP) | Out-Null
                            [Windows.Controls.Grid]::SetRow($lblP, $rowIdx)
                            [Windows.Controls.Grid]::SetColumn($lblP, 0)

                            $valP = New-Object Windows.Controls.TextBlock
                            $valP.Text = "$($prop.Value)"
                            $valP.TextWrapping = 'Wrap'
                            $valP.Margin = "0,4,4,4"
                            $propGrid.Children.Add($valP) | Out-Null
                            [Windows.Controls.Grid]::SetRow($valP, $rowIdx)
                            [Windows.Controls.Grid]::SetColumn($valP, 1)

                            $rowIdx++
                        }

                        $scroll.Content = $propGrid
                        $detailRoot.Children.Add($scroll) | Out-Null
                        [Windows.Controls.Grid]::SetRow($scroll, 1)

                        $msgIdentity = $msg.Identity.ToString()
                        $btnExportEml.Add_Click({
                            $dlg = New-Object Microsoft.Win32.SaveFileDialog
                            $dlg.Filter = "EML Files|*.eml|All Files|*.*"
                            $dlg.FileName = "message_$(Get-Date -Format 'yyyyMMdd_HHmmss').eml"
                            if ($dlg.ShowDialog() -ne $true) { return }
                            try {
                                $exServer      = $msgIdentity.Split('\')[0]
                                $hasAssemble   = $null -ne (Get-Command 'AssembleMessage' -ErrorAction SilentlyContinue)
                                $isLocalServer = $exServer -ieq $env:COMPUTERNAME

                                if ($hasAssemble -and $isLocalServer) {
                                    # Local EMS on the Exchange server — path is local
                                    Suspend-Message -Identity $msgIdentity -Confirm:$false -ErrorAction SilentlyContinue
                                    try {
                                        Export-Message -Identity $msgIdentity -ErrorAction Stop | AssembleMessage -Path $dlg.FileName
                                    } finally {
                                        Resume-Message -Identity $msgIdentity -Confirm:$false -ErrorAction SilentlyContinue
                                    }
                                    [System.Windows.MessageBox]::Show(
                                        "Exported to:`n$($dlg.FileName)",
                                        "Export Successful", 'OK', 'Information') | Out-Null
                                } elseif ($hasAssemble) {
                                    # Remote machine (e.g. EMC): run as SYSTEM via scheduled task, then UNC copy
                                    $eid        = $msgIdentity -replace "'", "''"
                                    $tempName   = "exq_$([System.Guid]::NewGuid().ToString('N')).eml"
                                    $remoteTemp = "C:\Windows\Temp\$tempName"
                                    $uncTemp    = "\\$exServer\C$\Windows\Temp\$tempName"
                                    $psCmd      = "Add-PSSnapin *Exchange* -EA SilentlyContinue; Export-Message -Identity '$eid' | AssembleMessage -Path '$remoteTemp'"
                                    $encodedCmd = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($psCmd))
                                    $taskName   = "ExQ_$([System.Guid]::NewGuid().ToString('N').Substring(0,8))"
                                    $exported   = $false
                                    try {
                                        $sched = New-Object -ComObject Schedule.Service
                                        $sched.Connect($exServer)
                                        $folder  = $sched.GetFolder('\')
                                        $taskDef = $sched.NewTask(0)
                                        $taskDef.Settings.ExecutionTimeLimit = 'PT2M'
                                        $action           = $taskDef.Actions.Create(0)
                                        $action.Path      = 'powershell.exe'
                                        $action.Arguments = "-NonInteractive -NoProfile -EncodedCommand $encodedCmd"
                                        $folder.RegisterTask($taskName, $taskDef.XmlText, 6, 'NT AUTHORITY\SYSTEM', $null, 5) | Out-Null
                                        try {
                                            $folder.GetTask($taskName).Run($null) | Out-Null
                                            $sw = [System.Diagnostics.Stopwatch]::StartNew()
                                            while ($folder.GetTask($taskName).State -eq 4 -and $sw.Elapsed.TotalSeconds -lt 45) {
                                                [System.Threading.Thread]::Sleep(500)
                                            }
                                            if (Test-Path $uncTemp) {
                                                Copy-Item $uncTemp $dlg.FileName -ErrorAction Stop
                                                Remove-Item $uncTemp -ErrorAction SilentlyContinue
                                                $exported = $true
                                            }
                                        } finally {
                                            try { $folder.DeleteTask($taskName, 0) | Out-Null } catch {}
                                        }
                                    } catch {}
                                    if ($exported) {
                                        [System.Windows.MessageBox]::Show(
                                            "Exported to:`n$($dlg.FileName)",
                                            "Export Successful", 'OK', 'Information') | Out-Null
                                    } else {
                                        $clipCmd = "Export-Message -Identity '$eid' | AssembleMessage -Path 'C:\TEMP\$([System.IO.Path]::GetFileName($dlg.FileName))'"
                                        try { [System.Windows.Clipboard]::SetText($clipCmd) } catch {}
                                        [System.Windows.MessageBox]::Show(
                                            "Automatic export failed.`n`nCommand copied to clipboard - run on $exServer in EMS:`n`n$clipCmd",
                                            "Run Manually", 'OK', 'Information') | Out-Null
                                    }
                                } else {
                                    $eid     = $msgIdentity -replace "'", "''"
                                    $clipCmd = "Export-Message -Identity '$eid' | AssembleMessage -Path 'C:\TEMP\message.eml'"
                                    try { [System.Windows.Clipboard]::SetText($clipCmd) } catch {}
                                    [System.Windows.MessageBox]::Show(
                                        "AssembleMessage not available.`nRun the script in the Exchange Management Shell (EMS).`n`nCommand copied to clipboard:`n`n$clipCmd",
                                        "Run Manually", 'OK', 'Information') | Out-Null
                                }
                            }
                            catch {
                                [System.Windows.MessageBox]::Show(
                                    "Export failed:`n$($_.Exception.Message)",
                                    "Error", 'OK', 'Error') | Out-Null
                            }
                        })

                $detailWin.Content = $detailRoot
                $detailWin.ShowDialog() | Out-Null
            }

            # Double-click on message row -> detail window
            $msgGrid.Add_MouseDoubleClick({
                $el = $_.OriginalSource -as [Windows.DependencyObject]
                while ($null -ne $el) {
                    if ($el -is [System.Windows.Controls.Primitives.DataGridColumnHeader]) { return }
                    if ($el -is [Windows.Controls.DataGridRow]) { & $doViewDetails; return }
                    $el = [Windows.Media.VisualTreeHelper]::GetParent($el) -as [Windows.DependencyObject]
                }
            })

            $btnViewDetails.Add_Click({ & $doViewDetails })

            $btnSuspendMsg.Add_Click({
                foreach ($msg in @($msgGrid.SelectedItems)) {
                    try { Suspend-Message -Identity $msg.Identity.ToString() -Confirm:$false } catch {}
                }
                & $reloadMessages
            })
            $btnResumeMsg.Add_Click({
                foreach ($msg in @($msgGrid.SelectedItems)) {
                    try { Resume-Message -Identity $msg.Identity.ToString() -Confirm:$false } catch {}
                }
                & $reloadMessages
            })
            $btnDeleteNdr.Add_Click({
                $sel = @($msgGrid.SelectedItems)
                if ($sel.Count -eq 0) { return }
                $r = [System.Windows.MessageBox]::Show(
                    "Delete $($sel.Count) message(s) with NDR?",
                    "Delete Messages", 'YesNo', 'Warning')
                if ($r -eq 'Yes') {
                    foreach ($msg in $sel) {
                        try { Remove-Message -Identity $msg.Identity.ToString() -WithNDR $true -Confirm:$false } catch {}
                    }
                    & $reloadMessages
                }
            })
            $btnDeleteNoNdr.Add_Click({
                $sel = @($msgGrid.SelectedItems)
                if ($sel.Count -eq 0) { return }
                $r = [System.Windows.MessageBox]::Show(
                    "Delete $($sel.Count) message(s) without NDR?",
                    "Delete Messages", 'YesNo', 'Warning')
                if ($r -eq 'Yes') {
                    foreach ($msg in $sel) {
                        try { Remove-Message -Identity $msg.Identity.ToString() -WithNDR $false -Confirm:$false } catch {}
                    }
                    & $reloadMessages
                }
            })
            $btnExportCsv.Add_Click({
                $dlg = New-Object Microsoft.Win32.SaveFileDialog
                $dlg.Filter = "CSV Files|*.csv"
                $dlg.FileName = "messages_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
                if ($dlg.ShowDialog() -eq $true) {
                    $msgGrid.ItemsSource | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
                }
            })

            # Context menu mirrors toolbar buttons
            $msgCtx = New-Object System.Windows.Controls.ContextMenu
            $miMsgViewDetails = New-Object System.Windows.Controls.MenuItem; $miMsgViewDetails.Header = "View Details..."; $miMsgViewDetails.IsEnabled = $false
            $miMsgSuspend    = New-Object System.Windows.Controls.MenuItem; $miMsgSuspend.Header    = "Suspend";               $miMsgSuspend.IsEnabled    = $false
            $miMsgResume     = New-Object System.Windows.Controls.MenuItem; $miMsgResume.Header     = "Resume";                $miMsgResume.IsEnabled     = $false
            $miMsgDelNdr     = New-Object System.Windows.Controls.MenuItem; $miMsgDelNdr.Header     = "Delete with NDR...";    $miMsgDelNdr.IsEnabled     = $false
            $miMsgDelNoNdr   = New-Object System.Windows.Controls.MenuItem; $miMsgDelNoNdr.Header   = "Delete without NDR..."; $miMsgDelNoNdr.IsEnabled   = $false
            $msgCtx.Items.Add($miMsgViewDetails) | Out-Null
            $msgCtx.Items.Add((New-Object System.Windows.Controls.Separator)) | Out-Null
            $msgCtx.Items.Add($miMsgSuspend)     | Out-Null
            $msgCtx.Items.Add($miMsgResume)      | Out-Null
            $msgCtx.Items.Add((New-Object System.Windows.Controls.Separator)) | Out-Null
            $msgCtx.Items.Add($miMsgDelNdr)      | Out-Null
            $msgCtx.Items.Add($miMsgDelNoNdr)    | Out-Null
            $msgGrid.ContextMenu = $msgCtx

            $miMsgViewDetails.Add_Click({ & $doViewDetails })

            $miMsgSuspend.Add_Click({
                foreach ($msg in @($msgGrid.SelectedItems)) {
                    try { Suspend-Message -Identity $msg.Identity.ToString() -Confirm:$false } catch {}
                }
                & $reloadMessages
            })
            $miMsgResume.Add_Click({
                foreach ($msg in @($msgGrid.SelectedItems)) {
                    try { Resume-Message -Identity $msg.Identity.ToString() -Confirm:$false } catch {}
                }
                & $reloadMessages
            })
            $miMsgDelNdr.Add_Click({
                $sel = @($msgGrid.SelectedItems)
                if ($sel.Count -eq 0) { return }
                $r = [System.Windows.MessageBox]::Show(
                    "Delete $($sel.Count) message(s) with NDR?",
                    "Delete Messages", 'YesNo', 'Warning')
                if ($r -eq 'Yes') {
                    foreach ($msg in $sel) {
                        try { Remove-Message -Identity $msg.Identity.ToString() -WithNDR $true -Confirm:$false } catch {}
                    }
                    & $reloadMessages
                }
            })
            $miMsgDelNoNdr.Add_Click({
                $sel = @($msgGrid.SelectedItems)
                if ($sel.Count -eq 0) { return }
                $r = [System.Windows.MessageBox]::Show(
                    "Delete $($sel.Count) message(s) without NDR?",
                    "Delete Messages", 'YesNo', 'Warning')
                if ($r -eq 'Yes') {
                    foreach ($msg in $sel) {
                        try { Remove-Message -Identity $msg.Identity.ToString() -WithNDR $false -Confirm:$false } catch {}
                    }
                    & $reloadMessages
                }
            })

            $msgWin.Content = $msgRoot
            $msgWin.ShowDialog() | Out-Null
        }
        catch {
            [System.Windows.MessageBox]::Show("Message view error: $($_.Exception.Message)", "Error", 'OK', 'Error') | Out-Null
        }
    }

    $miView.Add_Click({ & $doViewMessages })
    $btnViewMsgs.Add_Click({ & $doViewMessages })

    # "Delete All Messages" on the queue: Yes = with NDR, No = without NDR, Cancel = abort
    $doDeleteAllMessages = {
        $item = $dataGrid.SelectedItem
        if ($null -eq $item) { return }
        $id = $item.Identity.ToString()
        $r = [System.Windows.MessageBox]::Show(
            "Delete all messages in queue '$id'?`n`nYes = with NDR     No = without NDR",
            "Delete All Messages", 'YesNoCancel', 'Warning')
        if ($r -eq 'Cancel') { return }
        $withNdr = ($r -eq 'Yes')
        try {
            Get-Message -Queue $id | Remove-Message -WithNDR $withNdr -Confirm:$false
            $txtStatus.Text = "Messages deleted from: $id"
            $server = if ($combo.SelectedItem) { $combo.SelectedItem } else { $combo.Text }
            & $loadQueues $server
        }
        catch { [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", 'OK', 'Error') | Out-Null }
    }
    $miDelete.Add_Click({ & $doDeleteAllMessages })
    $btnDeleteAllMsgs.Add_Click({ & $doDeleteAllMessages })

    # ===========================
    # Button Events
    # ===========================
    $btnConnect.Add_Click({
        $server = if ($combo.SelectedItem) { $combo.SelectedItem } else { $combo.Text }
        if ($server) { & $loadQueues $server }
    })

    $btnRefresh.Add_Click({
        $server = if ($combo.SelectedItem) { $combo.SelectedItem } else { $combo.Text }
        if ($server) { & $loadQueues $server }
    })

    # Double-click on queue row -> View Messages
    $dataGrid.Add_MouseDoubleClick({
        $el = $_.OriginalSource -as [Windows.DependencyObject]
        while ($null -ne $el) {
            if ($el -is [System.Windows.Controls.Primitives.DataGridColumnHeader]) { return }
            if ($el -is [Windows.Controls.DataGridRow]) { & $doViewMessages; return }
            $el = [Windows.Media.VisualTreeHelper]::GetParent($el) -as [Windows.DependencyObject]
        }
    })

    # Stop timer on close
    $window.Add_Closed({ $timer.Stop() })

    # ===========================
    # Compose & Show
    # ===========================
    $root.Children.Add($toolbar) | Out-Null
    [Windows.Controls.Grid]::SetRow($toolbar, 0)

    $root.Children.Add($actionBar) | Out-Null
    [Windows.Controls.Grid]::SetRow($actionBar, 1)

    $root.Children.Add($dataGrid) | Out-Null
    [Windows.Controls.Grid]::SetRow($dataGrid, 2)

    $window.Content = $root

    # Deferred init: let window render before blocking Exchange cmdlets run
    $window.Add_Loaded({
        $initTimer = New-Object System.Windows.Threading.DispatcherTimer
        $initTimer.Interval = [TimeSpan]::FromMilliseconds(150)
        $initTimer.Add_Tick({
            $this.Stop()

            $txtStatus.Text = "Querying Exchange servers..."
            & $flushUI
            try {
                $servers = @(Get-ExchangeServer |
                    Where-Object { $_.IsHubTransportServer -or $_.IsMailboxServer } |
                    Sort-Object Name |
                    Select-Object -ExpandProperty Name)

                $txtStatus.Text = "Found $($servers.Count) server(s), populating list..."
                & $flushUI
                $combo.Items.Clear()
                foreach ($s in $servers) { [void]$combo.Items.Add($s) }

                if ($combo.Items.Count -gt 0) {
                    $combo.SelectedIndex = 0
                    & $loadQueues ($combo.SelectedItem)
                } else {
                    $txtStatus.Text = "No Exchange servers found - enter server name manually"
                }
            }
            catch {
                $txtStatus.Text = "Get-ExchangeServer failed - enter server name manually and click Connect"
            }
        })
        $initTimer.Start()
    })

    $null = $window.ShowDialog()
}

### Load Exchange shell
$snapin = Get-PSSnapin -Registered Microsoft.Exchange.Management.PowerShell.E* -ErrorAction SilentlyContinue
if ($null -eq $snapin) {
    Write-Warning "Exchange snapin not available"
} else {
    if ($null -eq (Get-Command "Get-Queue" -ErrorAction SilentlyContinue)) {
        Write-Host "Loading Exchange commands..."
        Add-PSSnapin $snapin -ErrorAction SilentlyContinue
    } else {
        Write-Host "Exchange snapin already loaded"
    }
}

Open-QueueViewerDialog
