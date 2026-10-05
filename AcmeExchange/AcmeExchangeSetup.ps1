<#
.SYNOPSIS
    GUI front-end for Invoke-AcmeExchangeCert.ps1 - edit configuration and run the
    setup / renewal / task actions with a few clicks. No console needed for config;
    long-running actions (Azure setup, issuance) open a console so you can watch progress
    and complete the Azure device-code sign-in.

    Layout: a menu bar (File / Language / Help) over grouped sections
    (Certificate, DNS provider, Exchange, Notifications, "Run now", Scheduled task),
    and a log pane that grows when the window is resized taller. The scheduled task is
    presented as the recommended steady-state; the two "Run now" buttons are the
    one-time / manual path. UI language (de/en) switches at runtime and is stored in config.

.NOTES
    Run elevated. Part of the Exchange-ACME-Cert bundle; keep it next to
    Invoke-AcmeExchangeCert.ps1 and the lib\ folder. License: MIT.
#>
[CmdletBinding()]
param(
    [string]$ConfigPath = 'C:\Tools\AcmeExchange\config.json',
    [string]$EnginePath,
    [switch]$SelfTest    # build the form and exit (no window) - for validation
)

$ErrorActionPreference = 'Stop'

# Single place for the version number; build and release tag read it from here.
$script:Version = '1.0.0'

# --- locate bundle + engine (works as .ps1 and when wrapped into an .exe by PS2EXE) ---
$BundleRoot =
    if ($PSScriptRoot) { $PSScriptRoot }
    elseif ($MyInvocation.MyCommand.Path) { Split-Path $MyInvocation.MyCommand.Path -Parent }
    else { Split-Path ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -Parent }
if (-not $EnginePath) { $EnginePath = Join-Path $BundleRoot 'Invoke-AcmeExchangeCert.ps1' }
if (-not (Test-Path $EnginePath)) { throw "Engine script not found: $EnginePath" }

# --- load the engine's pure helpers (config skeleton + DNS provider table) ---
$engineAst = [System.Management.Automation.Language.Parser]::ParseFile($EnginePath, [ref]$null, [ref]$null)
$script:BootstrapClientId = '1950a258-227b-4e31-a9cf-717495945fc2'
foreach ($fn in 'Get-DefaultConfig', 'Get-DnsProviders') {
    $def = $engineAst.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $args[0].Name -eq $fn }, $false)
    if ($def) { Invoke-Expression $def[0].Extent.Text }
}

function New-Config { if (Get-Command Get-DefaultConfig -ErrorAction SilentlyContinue) { return Get-DefaultConfig } else { return [ordered]@{} } }
$script:Providers = if (Get-Command Get-DnsProviders -ErrorAction SilentlyContinue) { Get-DnsProviders } else { [ordered]@{ 'Azure' = @{ Plugin = 'Azure'; Bootstrap = $true; Fields = @() } } }

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:BlogUrl = 'https://blog.zarenko.net'

# grey placeholder text ("cue banner") for empty text boxes - shown as fill-in help on a
# fresh start, never stored. Uses the Win32 EM_SETCUEBANNER edit-control message.
Add-Type -Namespace Native -Name Edit -MemberDefinition @'
[DllImport("user32.dll", CharSet=CharSet.Unicode)]
public static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);
'@
$script:EM_SETCUEBANNER = 0x1501
function Set-Cue($ctrl, [string]$text) {
    if ($ctrl -and $ctrl.IsHandleCreated) { [void][Native.Edit]::SendMessage($ctrl.Handle, $script:EM_SETCUEBANNER, [IntPtr]1, $text) }
}

# --------------------------------------------------------------- localization
$script:Lang = 'de'
$script:Strings = @{
    de = @{
        title      = 'Exchange ACME Zertifikat Setup - https://blog.zarenko.net'
        mFile='Datei'; mLoad='Konfiguration laden'; mSave='Konfiguration speichern'; mExit='Beenden'
        mLang='Sprache'; mDe='Deutsch'; mEn='English'
        mHelp='Hilfe'; mStatus='Status anzeigen'; mLogDir='Log-Verzeichnis öffnen'; mAbout='Über ...'
        gCert='Zertifikat (ACME)'; gDns='DNS-Anbieter'; gExch='Exchange'; gNotify='Benachrichtigungen'
        gRun='Jetzt ausführen (manuell)'; gTaskRec='Geplante Aufgabe (empfohlen)'; gLog='Protokoll'
        lDomains='Domains (Komma):'; lContact='Kontakt-E-Mail:'
        lProvider='DNS-Anbieter:'; lSub='Azure Subscription-ID:'; lApp='App-Anzeigename:'; subHint='(leer = automatisch aus der Zone)'
        lServers='Server (Komma):'; lServices='Dienste:'; lHome='Arbeitsverzeichnis:'
        lSmtp='SMTP-Server:Port:'; lFrom='Absender:'; lTo='Empfänger:'; lWarn='Warnung (Tage):'
        lTaskName='Aufgabenname:'; lTaskTime='Startzeit (HH:mm):'
        cStaging='Staging (Test-Zertifikat, keine Installation)'; cSuccess='Erfolgsmail senden'
        bDiscover='Ermitteln'; bSetupAzure='1. Azure einrichten (einmalig)'; bSetupPrep='1. Vorbereiten (einmalig)'
        bRenew='2. Zertifikat holen + installieren'
        bTestMail='Test-Mail senden'; bTaskAdd='Task anlegen'; bTaskDel='Task entfernen'
        hintRun='Für die Ersteinrichtung "1" einmal ausführen; "2" holt und installiert das Zertifikat sofort. Für den Dauerbetrieb stattdessen unten den Task anlegen.'
        hintTask='Empfohlen: Einmal "Task anlegen" - der Task verlängert und installiert das Zertifikat danach automatisch (täglich geprüft). Manuelle Läufe sind dann nicht mehr nötig.'
        logSaved='Konfiguration gespeichert: {0}'; logLoaded='Konfiguration geladen: {0}'
        logNoConfig='Keine Konfig unter {0} - Standardwerte.'; logBundle='Bundle: {0}'; logConfig='Konfig: {0}'
        logDiscover='Ermittle Exchange-Mailbox-Server ...'; logFound='Gefunden: {0}'
        logSetupStarted='Azure-Einrichtung in eigener Konsole gestartet. Gerätecode dort eingeben. Danach einmalig "2. Zertifikat holen + installieren" oder gleich den Task anlegen.'
        logRenewStarted='Zertifikatslauf in eigener Konsole gestartet (Modus: {0}).'
        logStatus='Status wird abgerufen ...'; logTaskAdd='Task-Anlage ausgeführt.'; logTaskDel='Task-Entfernung ausgeführt.'
        logTestMail='Sende Test-Mail an: {0} über {1}:{2} ...'; logSecrets='DNS-Zugangsdaten werden einmalig an die Engine übergeben (danach verschlüsselt in Posh-ACME).'
        logLangSet='Sprache: Deutsch'; err='FEHLER: {0}'; logBusy='Ein Vorgang läuft noch - bitte warten.'; logNoLogDir='Log-Verzeichnis noch nicht vorhanden: {0}'; logCfgReloaded='Konfiguration aktualisiert - Subscription/App aus dem Setup übernommen.'
        modeStaging='STAGING (nur holen)'; modeProd='PRODUKTIV (holen + installieren)'
        mbSetupTitle='Schritt 1: Azure einrichten'
        mbSetupBody="Schritt 1 - Azure einrichten (einmalige Vorbereitung):`n`nAnmeldung, App-Registrierung, Zertifikat-Credential, DNS-Rolle und Konfiguration.`n`nEs wird noch KEIN Zertifikat ausgestellt und nichts auf den Servern geändert.`nEs öffnet sich eine Konsole - melde dich dort per Gerätecode an (URL + Code kommen automatisch)."
        mbRenewTitle='Schritt 2: Zertifikat holen + installieren'
        mbRenewStaging="STAGING (Test-CA): Es wird nur ein TEST-Zertifikat geholt, um Azure/DNS zu prüfen.`nDie Server ({0}) werden NICHT verändert."
        mbRenewProd="PRODUKTIV: Es wird ein echtes Zertifikat geholt und auf den Servern ({0}) installiert.`nDas ersetzt das bisherige Zertifikat gleichen Namens, passt die Connector-Zuordnung an`nund macht je Server einen kurzen iisreset. (Setzt voraus, dass Schritt 1 gelaufen ist.)`n`nHinweis: Für den Dauerbetrieb ist der geplante Task der bequemere Weg - er macht das automatisch."
        mbConsoleTail="`n`nEs öffnet sich eine Konsole mit dem Fortschritt.`n`nFortfahren?"
        mbSetupTail="`n`nFortfahren?"
        aboutTitle='Über'; aboutBody="Exchange ACME Zertifikat - Setup`n`nGUI für Invoke-AcmeExchangeCert.ps1`nBundle: {0}`nKonfig: {1}"
    }
    en = @{
        title      = 'Exchange ACME Certificate Setup - https://blog.zarenko.net'
        mFile='File'; mLoad='Load configuration'; mSave='Save configuration'; mExit='Exit'
        mLang='Language'; mDe='Deutsch'; mEn='English'
        mHelp='Help'; mStatus='Show status'; mLogDir='Open log directory'; mAbout='About ...'
        gCert='Certificate (ACME)'; gDns='DNS provider'; gExch='Exchange'; gNotify='Notifications'
        gRun='Run now (manual)'; gTaskRec='Scheduled task (recommended)'; gLog='Log'
        lDomains='Domains (comma):'; lContact='Contact e-mail:'
        lProvider='DNS provider:'; lSub='Azure Subscription ID:'; lApp='App display name:'; subHint='(empty = auto-detect from zone)'
        lServers='Servers (comma):'; lServices='Services:'; lHome='Working directory:'
        lSmtp='SMTP server:port:'; lFrom='Sender:'; lTo='Recipient:'; lWarn='Warn (days):'
        lTaskName='Task name:'; lTaskTime='Start time (HH:mm):'
        cStaging='Staging (test certificate, no installation)'; cSuccess='Send success mail'
        bDiscover='Discover'; bSetupAzure='1. Set up Azure (once)'; bSetupPrep='1. Prepare (once)'
        bRenew='2. Get + install certificate'
        bTestMail='Send test mail'; bTaskAdd='Create task'; bTaskDel='Remove task'
        hintRun='For first-time setup run "1" once; "2" issues and installs the certificate immediately. For continuous operation create the task below instead.'
        hintTask='Recommended: click "Create task" once - it then renews and installs the certificate automatically (checked daily). Manual runs are no longer needed.'
        logSaved='Configuration saved: {0}'; logLoaded='Configuration loaded: {0}'
        logNoConfig='No config at {0} - using defaults.'; logBundle='Bundle: {0}'; logConfig='Config: {0}'
        logDiscover='Discovering Exchange mailbox servers ...'; logFound='Found: {0}'
        logSetupStarted='Azure setup started in its own console. Enter the device code there. Then run "2. Get + install certificate" once, or just create the task.'
        logRenewStarted='Certificate run started in its own console (mode: {0}).'
        logStatus='Retrieving status ...'; logTaskAdd='Task creation executed.'; logTaskDel='Task removal executed.'
        logTestMail='Sending test mail to: {0} via {1}:{2} ...'; logSecrets='DNS credentials are handed to the engine once (kept encrypted by Posh-ACME afterwards).'
        logLangSet='Language: English'; err='ERROR: {0}'; logBusy='A task is still running - please wait.'; logNoLogDir='Log directory does not exist yet: {0}'; logCfgReloaded='Configuration refreshed - subscription/app from setup applied.'
        modeStaging='STAGING (issue only)'; modeProd='PRODUCTION (issue + install)'
        mbSetupTitle='Step 1: Set up Azure'
        mbSetupBody="Step 1 - Set up Azure (one-time preparation):`n`nSign-in, app registration, certificate credential, DNS role and configuration.`n`nNo certificate is issued yet and nothing on the servers is changed.`nA console opens - sign in there with the device code (URL + code appear automatically)."
        mbRenewTitle='Step 2: Get + install certificate'
        mbRenewStaging="STAGING (test CA): Only a TEST certificate is issued to verify Azure/DNS.`nThe servers ({0}) are NOT changed."
        mbRenewProd="PRODUCTION: A real certificate is issued and installed on the servers ({0}).`nThis replaces the previous certificate of the same name, adjusts the connector binding`nand does a short iisreset per server. (Requires that step 1 has run.)`n`nNote: for continuous operation the scheduled task is the easier path - it does this automatically."
        mbConsoleTail="`n`nA console opens showing the progress.`n`nContinue?"
        mbSetupTail="`n`nContinue?"
        aboutTitle='About'; aboutBody="Exchange ACME Certificate - Setup`n`nGUI for Invoke-AcmeExchangeCert.ps1`nBundle: {0}`nConfig: {1}"
    }
}
function T([string]$key) { $t = $script:Strings[$script:Lang][$key]; if ($null -eq $t) { $key } else { $t } }

# --------------------------------------------------------------- control helpers
function New-Label($text, $x, $y, $w = 125) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $text; $l.Location = New-Object System.Drawing.Point($x, ($y + 3)); $l.Size = New-Object System.Drawing.Size($w, 20)
    return $l
}
function New-Hint($text, $x, $y, $w, $h = 30) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $text; $l.AutoSize = $false
    $l.Location = New-Object System.Drawing.Point($x, $y); $l.Size = New-Object System.Drawing.Size($w, $h)
    $l.ForeColor = [System.Drawing.Color]::DimGray
    $l.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    return $l
}
function New-Text($x, $y, $w) {
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location = New-Object System.Drawing.Point($x, $y); $t.Size = New-Object System.Drawing.Size($w, 22)
    return $t
}
function New-Group($title, $x, $y, $w, $h) {
    $g = New-Object System.Windows.Forms.GroupBox
    $g.Text = $title; $g.Location = New-Object System.Drawing.Point($x, $y); $g.Size = New-Object System.Drawing.Size($w, $h)
    return $g
}
function New-Btn($text, $x, $y, $w = 150, $h = 28) {
    $b = New-Object System.Windows.Forms.Button; $b.Text = $text
    $b.Location = New-Object System.Drawing.Point($x, $y); $b.Size = New-Object System.Drawing.Size($w, $h)
    return $b
}

# --------------------------------------------------------------- form + menu
$FormW = 720; $InnerW = $FormW - 24; $Gx = 12
$boldFont = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$form = New-Object System.Windows.Forms.Form
$form.Text = T 'title'
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.MinimizeBox = $true
$iconPath = Join-Path $BundleRoot 'icon.ico'
if (Test-Path $iconPath) { try { $form.Icon = New-Object System.Drawing.Icon($iconPath) } catch { } }

$menu = New-Object System.Windows.Forms.MenuStrip
$miFile   = New-Object System.Windows.Forms.ToolStripMenuItem
$miLoad   = New-Object System.Windows.Forms.ToolStripMenuItem
$miSave   = New-Object System.Windows.Forms.ToolStripMenuItem
$miExit   = New-Object System.Windows.Forms.ToolStripMenuItem
[void]$miFile.DropDownItems.Add($miLoad); [void]$miFile.DropDownItems.Add($miSave)
[void]$miFile.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator)); [void]$miFile.DropDownItems.Add($miExit)
$miLang   = New-Object System.Windows.Forms.ToolStripMenuItem
$miDe     = New-Object System.Windows.Forms.ToolStripMenuItem
$miEn     = New-Object System.Windows.Forms.ToolStripMenuItem
[void]$miLang.DropDownItems.Add($miDe); [void]$miLang.DropDownItems.Add($miEn)
$miHelp   = New-Object System.Windows.Forms.ToolStripMenuItem
$miStatusM= New-Object System.Windows.Forms.ToolStripMenuItem
$miLogDir = New-Object System.Windows.Forms.ToolStripMenuItem
$miAbout  = New-Object System.Windows.Forms.ToolStripMenuItem
[void]$miHelp.DropDownItems.Add($miStatusM); [void]$miHelp.DropDownItems.Add($miLogDir)
[void]$miHelp.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator)); [void]$miHelp.DropDownItems.Add($miAbout)
[void]$menu.Items.Add($miFile); [void]$menu.Items.Add($miLang); [void]$menu.Items.Add($miHelp)
$form.MainMenuStrip = $menu
$form.Controls.Add($menu)

$topY = 28   # below the menu strip

# --------------------------------------------------------------- section: Certificate
$grpCert = New-Group (T 'gCert') $Gx $topY $InnerW 88
$lblDomains = New-Label (T 'lDomains') 12 24 120; $grpCert.Controls.Add($lblDomains)
$txtDomains = New-Text 140 24 540;                $grpCert.Controls.Add($txtDomains)
$lblContact = New-Label (T 'lContact') 12 54 120; $grpCert.Controls.Add($lblContact)
$txtContact = New-Text 140 54 250;                $grpCert.Controls.Add($txtContact)
$chkStaging = New-Object System.Windows.Forms.CheckBox
$chkStaging.Text = T 'cStaging'; $chkStaging.Location = New-Object System.Drawing.Point(410, 55); $chkStaging.Size = New-Object System.Drawing.Size(280, 22)
$grpCert.Controls.Add($chkStaging)
$form.Controls.Add($grpCert)

# --------------------------------------------------------------- section: DNS provider
$grpDns = New-Group (T 'gDns') $Gx ($topY + 96) $InnerW 118
$lblProvider = New-Label (T 'lProvider') 12 24 150; $grpDns.Controls.Add($lblProvider)
$cboProvider = New-Object System.Windows.Forms.ComboBox
$cboProvider.DropDownStyle = 'DropDownList'
$cboProvider.Location = New-Object System.Drawing.Point(170, 24); $cboProvider.Size = New-Object System.Drawing.Size(200, 24)
foreach ($k in $script:Providers.Keys) { [void]$cboProvider.Items.Add($k) }
$grpDns.Controls.Add($cboProvider)
# Row 2: Azure Subscription-ID (+ hint)  OR  provider credential field 1
$lblSub     = New-Label (T 'lSub') 12 54 150;                 $grpDns.Controls.Add($lblSub)
$txtSub     = New-Text 170 54 280;                            $grpDns.Controls.Add($txtSub)
$lblSubHint = New-Label (T 'subHint') 460 57 230;             $grpDns.Controls.Add($lblSubHint)
$lblCred1   = New-Label 'Feld 1:' 12 54 160;                  $grpDns.Controls.Add($lblCred1)
$txtCred1   = New-Text 180 54 480;                            $grpDns.Controls.Add($txtCred1)
# Row 3: Azure App-Anzeigename  OR  provider credential field 2
$lblApp   = New-Label (T 'lApp') 12 84 150;                   $grpDns.Controls.Add($lblApp)
$txtApp   = New-Text 170 84 280;                              $grpDns.Controls.Add($txtApp)
$lblCred2 = New-Label 'Feld 2:' 12 84 160;                    $grpDns.Controls.Add($lblCred2)
$txtCred2 = New-Text 180 84 480;                              $grpDns.Controls.Add($txtCred2)
$form.Controls.Add($grpDns)

# --------------------------------------------------------------- section: Exchange
$grpExch = New-Group (T 'gExch') $Gx ($topY + 224) $InnerW 118
$lblServers = New-Label (T 'lServers') 12 24 120; $grpExch.Controls.Add($lblServers)
$txtServers = New-Text 140 24 400;                $grpExch.Controls.Add($txtServers)
$btnDiscover = New-Btn (T 'bDiscover') 550 23 130 24; $grpExch.Controls.Add($btnDiscover)
$lblServices = New-Label (T 'lServices') 12 54 120; $grpExch.Controls.Add($lblServices)
$chkIIS  = New-Object System.Windows.Forms.CheckBox; $chkIIS.Text='IIS';  $chkIIS.Location=New-Object System.Drawing.Point(140,54); $chkIIS.Size=New-Object System.Drawing.Size(60,22);  $chkIIS.Checked=$true
$chkSMTP = New-Object System.Windows.Forms.CheckBox; $chkSMTP.Text='SMTP';$chkSMTP.Location=New-Object System.Drawing.Point(210,54); $chkSMTP.Size=New-Object System.Drawing.Size(70,22); $chkSMTP.Checked=$true
$grpExch.Controls.Add($chkIIS); $grpExch.Controls.Add($chkSMTP)
$lblHome = New-Label (T 'lHome') 12 84 120; $grpExch.Controls.Add($lblHome)
$txtHome = New-Text 140 84 520;             $grpExch.Controls.Add($txtHome)
$form.Controls.Add($grpExch)

# --------------------------------------------------------------- section: Notifications (+ test mail)
$grpNotify = New-Group (T 'gNotify') $Gx ($topY + 352) $InnerW 118
$lblSmtp = New-Label (T 'lSmtp') 12 24 120; $grpNotify.Controls.Add($lblSmtp)
$txtSmtp = New-Text 140 24 300;             $grpNotify.Controls.Add($txtSmtp)
$txtPort = New-Text 450 24 70;              $grpNotify.Controls.Add($txtPort)
$lblFrom = New-Label (T 'lFrom') 12 54 120; $grpNotify.Controls.Add($lblFrom)
$txtFrom = New-Text 140 54 190;             $grpNotify.Controls.Add($txtFrom)
$lblTo   = New-Label (T 'lTo') 345 54 80;   $grpNotify.Controls.Add($lblTo)
$txtTo   = New-Text 430 54 250;             $grpNotify.Controls.Add($txtTo)
$lblWarn = New-Label (T 'lWarn') 12 84 115; $grpNotify.Controls.Add($lblWarn)
$txtWarn = New-Text 140 84 55;              $grpNotify.Controls.Add($txtWarn)
$chkSuccess = New-Object System.Windows.Forms.CheckBox; $chkSuccess.Text = T 'cSuccess'; $chkSuccess.Location=New-Object System.Drawing.Point(210,84); $chkSuccess.Size=New-Object System.Drawing.Size(180,22); $chkSuccess.Checked=$true
$grpNotify.Controls.Add($chkSuccess)
$btnTestMail = New-Btn (T 'bTestMail') 510 82 170 26; $grpNotify.Controls.Add($btnTestMail)
$form.Controls.Add($grpNotify)

# --------------------------------------------------------------- section: Run now (manual, one-time path)
$grpRun = New-Group (T 'gRun') $Gx ($topY + 480) $InnerW 90
$lblHintRun = New-Hint (T 'hintRun') 12 18 ($InnerW - 24) 28; $grpRun.Controls.Add($lblHintRun)
$w2 = [int](($InnerW - 24 - 8) / 2)
$btnSetup = New-Btn (T 'bSetupAzure') 12 50 $w2 32; $btnSetup.Font = $boldFont; $grpRun.Controls.Add($btnSetup)
$btnRenew = New-Btn (T 'bRenew') (12 + $w2 + 8) 50 $w2 32; $btnRenew.Font = $boldFont; $grpRun.Controls.Add($btnRenew)
$form.Controls.Add($grpRun)

# --------------------------------------------------------------- section: Scheduled task (recommended)
$grpTask = New-Group (T 'gTaskRec') $Gx ($topY + 578) $InnerW 122
$lblHintTask = New-Hint (T 'hintTask') 12 18 ($InnerW - 24) 28
$lblHintTask.ForeColor = [System.Drawing.Color]::FromArgb(0, 110, 0)
$lblHintTask.Font = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)
$grpTask.Controls.Add($lblHintTask)
$lblTaskName = New-Label (T 'lTaskName') 12 52 120; $grpTask.Controls.Add($lblTaskName)
$txtTaskName = New-Text 140 52 300;                 $grpTask.Controls.Add($txtTaskName)
$lblTaskTime = New-Label (T 'lTaskTime') 460 52 110;$grpTask.Controls.Add($lblTaskTime)
$txtTaskTime = New-Text 575 52 90;                  $grpTask.Controls.Add($txtTaskTime)
$wTask = [int](($InnerW - 24 - 8) / 2)
$btnTaskAdd = New-Btn (T 'bTaskAdd') 12 84 $wTask 30; $btnTaskAdd.Font = $boldFont; $grpTask.Controls.Add($btnTaskAdd)
$btnTaskDel = New-Btn (T 'bTaskDel') (12 + $wTask + 8) 84 $wTask 30; $grpTask.Controls.Add($btnTaskDel)
$form.Controls.Add($grpTask)

# --------------------------------------------------------------- log (grows with the window)
$logH = 150
$logY = $topY + 708
$grpLog = New-Group (T 'gLog') $Gx $logY $InnerW $logH   # no anchor: keep its gray margin on all sides
# RichTextBox, word wrap OFF: long lines stay on one line and a horizontal scrollbar
# lets you scroll to their end. Scrollbars auto-hide (unlike a TextBox), so a short
# log stays clean - the horizontal bar shows only when a line is too long, the vertical
# only when there are many lines.
$log = New-Object System.Windows.Forms.RichTextBox
$log.ReadOnly = $true
$log.WordWrap = $false
$log.ScrollBars = 'Both'
$log.BorderStyle = 'FixedSingle'
$log.BackColor = [System.Drawing.Color]::White
$log.Font = New-Object System.Drawing.Font('Consolas', 8)
$log.Location = New-Object System.Drawing.Point(12, 20)
$log.Size = New-Object System.Drawing.Size(($InnerW - 24), ($logH - 32))
$grpLog.Controls.Add($log)
$form.Controls.Add($grpLog)

# fixed-size dialog: content sits inside a uniform gray margin on all four sides
# (bottom margin = 14 px below the log group).
$form.ClientSize = New-Object System.Drawing.Size($FormW, ($logY + $logH + 14))

[void]$log.Handle   # force native handle so text appended before the form is shown is retained

function Write-GuiLog($msg) {
    $ts = (Get-Date).ToString('HH:mm:ss')
    $log.AppendText("$ts  $msg`r`n")
    if ($log.IsHandleCreated) {
        # scroll vertically to the newest line but keep the horizontal position at the
        # line START (AppendText would otherwise jump to the right end of a long line).
        $idx = $log.Text.TrimEnd("`r", "`n").LastIndexOf("`n") + 1
        if ($idx -lt 0) { $idx = 0 }
        $log.SelectionStart = $idx; $log.SelectionLength = 0; $log.ScrollToCaret()
    }
}

# Background-capture plumbing: quick actions run hidden and their output is tailed from
# temp files by this timer (on the UI thread) so the window never freezes.
$script:Cap = @{ Proc = $null; Out = $null; Err = $null; OutPos = 0; ErrPos = 0; DoneMsg = $null }
$logTimer = New-Object System.Windows.Forms.Timer
$logTimer.Interval = 150
$logTimer.Add_Tick({ Drain-Capture })

# After "1. Set up Azure" the engine writes the discovered subscription/app into config.json in its own
# console; watch the file and reload the form so those values actually show up (and are not lost).
$script:ExpectReload = $false
$script:CfgWatchMtime = $null
$cfgWatchTimer = New-Object System.Windows.Forms.Timer
$cfgWatchTimer.Interval = 1000
$cfgWatchTimer.Add_Tick({
    if (-not $script:ExpectReload) { return }
    if (Test-Path $ConfigPath) {
        $m = (Get-Item $ConfigPath).LastWriteTime
        if (($null -eq $script:CfgWatchMtime) -or ($m -gt $script:CfgWatchMtime)) {
            try { Load-GuiConfig; $script:ExpectReload = $false; Write-GuiLog (T 'logCfgReloaded') } catch { }  # may be mid-write; retry next tick
        }
    }
})
$cfgWatchTimer.Start()

# --------------------------------------------------------------- provider view toggle
$script:CurFields = @()
function Set-ProviderView([string]$provider) {
    if (-not $provider) { $provider = 'Azure' }
    $azure = ($provider -eq 'Azure')
    $lblSub.Visible = $azure; $txtSub.Visible = $azure; $lblSubHint.Visible = $azure
    $lblApp.Visible = $azure; $txtApp.Visible = $azure
    $def = $script:Providers[$provider]
    $script:CurFields = @()
    if ($def -and $def.Fields) { $script:CurFields = @($def.Fields) }   # plain-if: the $x=if(){}else{} form mis-collects here
    $f1 = $null; if ($script:CurFields.Count -ge 1) { $f1 = $script:CurFields[0] }
    $f2 = $null; if ($script:CurFields.Count -ge 2) { $f2 = $script:CurFields[1] }
    $c1 = [bool](-not $azure -and $f1); $lblCred1.Visible = $c1; $txtCred1.Visible = $c1
    $c2 = [bool](-not $azure -and $f2); $lblCred2.Visible = $c2; $txtCred2.Visible = $c2
    if ($f1) { $lblCred1.Text = "$($f1.Label):"; $txtCred1.UseSystemPasswordChar = [bool]$f1.Secure }
    if ($f2) { $lblCred2.Text = "$($f2.Label):"; $txtCred2.UseSystemPasswordChar = [bool]$f2.Secure }
    $btnSetup.Text = if ($azure) { T 'bSetupAzure' } else { T 'bSetupPrep' }
}
$cboProvider.Add_SelectedIndexChanged({ Set-ProviderView $cboProvider.SelectedItem })

# --------------------------------------------------------------- apply language
function Set-Language([string]$lang) {
    if ($lang -ne 'en') { $lang = 'de' }
    $script:Lang = $lang
    $miDe.Checked = ($lang -eq 'de'); $miEn.Checked = ($lang -eq 'en')
    $form.Text = T 'title'
    $miFile.Text=T 'mFile'; $miLoad.Text=T 'mLoad'; $miSave.Text=T 'mSave'; $miExit.Text=T 'mExit'
    $miLang.Text=T 'mLang'; $miDe.Text=T 'mDe'; $miEn.Text=T 'mEn'
    $miHelp.Text=T 'mHelp'; $miStatusM.Text=T 'mStatus'; $miLogDir.Text=T 'mLogDir'; $miAbout.Text=T 'mAbout'
    $grpCert.Text=T 'gCert'; $grpDns.Text=T 'gDns'; $grpExch.Text=T 'gExch'; $grpNotify.Text=T 'gNotify'
    $grpRun.Text=T 'gRun'; $grpTask.Text=T 'gTaskRec'; $grpLog.Text=T 'gLog'
    $lblDomains.Text=T 'lDomains'; $lblContact.Text=T 'lContact'; $chkStaging.Text=T 'cStaging'
    $lblProvider.Text=T 'lProvider'; $lblSub.Text=T 'lSub'; $lblApp.Text=T 'lApp'; $lblSubHint.Text=T 'subHint'
    $lblServers.Text=T 'lServers'; $lblServices.Text=T 'lServices'; $lblHome.Text=T 'lHome'; $btnDiscover.Text=T 'bDiscover'
    $lblSmtp.Text=T 'lSmtp'; $lblFrom.Text=T 'lFrom'; $lblTo.Text=T 'lTo'; $lblWarn.Text=T 'lWarn'; $chkSuccess.Text=T 'cSuccess'; $btnTestMail.Text=T 'bTestMail'
    $lblHintRun.Text=T 'hintRun'; $lblHintTask.Text=T 'hintTask'
    $lblTaskName.Text=T 'lTaskName'; $lblTaskTime.Text=T 'lTaskTime'; $btnTaskAdd.Text=T 'bTaskAdd'; $btnTaskDel.Text=T 'bTaskDel'
    $btnRenew.Text=T 'bRenew'
    Set-ProviderView $cboProvider.SelectedItem   # refreshes the setup button + credential labels
}

# --------------------------------------------------------------- config <-> form
function Set-FormFromConfig($cfg) {
    $txtDomains.Text  = (@($cfg.Acme.Domains) -join ', ')
    $txtContact.Text  = [string]$cfg.Acme.Contact
    $txtServers.Text  = (@($cfg.Exchange.Servers) -join ', ')
    $svc = [string]$cfg.Exchange.Services
    $chkIIS.Checked   = $svc -match 'IIS'
    $chkSMTP.Checked  = $svc -match 'SMTP'
    $txtHome.Text     = [string]$cfg.Paths.Home
    $txtSmtp.Text     = [string]$cfg.Notify.SmtpServer
    $txtPort.Text     = [string]$cfg.Notify.Port
    $txtFrom.Text     = [string]$cfg.Notify.From
    $txtTo.Text       = (@($cfg.Notify.To) -join ', ')
    $txtWarn.Text     = [string]$cfg.Notify.WarnDaysBeforeExpiry
    $chkSuccess.Checked = [bool]$cfg.Notify.SendSuccessMail
    $txtTaskName.Text = [string]$cfg.Task.Name
    $txtTaskTime.Text = [string]$cfg.Task.Time
    $txtSub.Text      = [string]$cfg.Azure.SubscriptionId
    $txtApp.Text      = [string]$cfg.Azure.AppDisplayName

    # DNS provider + its view
    $prov = if ($cfg.Acme.PSObject.Properties['DnsProvider'] -and $cfg.Acme.DnsProvider) { [string]$cfg.Acme.DnsProvider } else { 'Azure' }
    if ($cboProvider.Items -contains $prov) { $cboProvider.SelectedItem = $prov } else { $cboProvider.SelectedItem = 'Azure' }
    Set-ProviderView $cboProvider.SelectedItem
    # fill the non-secret credential fields from DnsPluginArgs (secrets are never stored here)
    if (-not ($cboProvider.SelectedItem -eq 'Azure') -and $cfg.Acme.PSObject.Properties['DnsPluginArgs'] -and $cfg.Acme.DnsPluginArgs) {
        $pa = $cfg.Acme.DnsPluginArgs
        if ($script:CurFields.Count -ge 1 -and -not $script:CurFields[0].Secure -and $pa.PSObject.Properties[$script:CurFields[0].Key]) { $txtCred1.Text = [string]$pa.$($script:CurFields[0].Key) }
        if ($script:CurFields.Count -ge 2 -and -not $script:CurFields[1].Secure -and $pa.PSObject.Properties[$script:CurFields[1].Key]) { $txtCred2.Text = [string]$pa.$($script:CurFields[1].Key) }
    }
}

function Split-Csv($s) { @($s -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }

function Get-ConfigFromForm {
    $cfg = New-Config
    # Preserve fields the form does not manage (Azure app identity, credential, zone ids, tenant,
    # ACME server) so saving from the GUI never wipes what -Setup created.
    if (Test-Path $ConfigPath) {
        try {
            $old = Get-Content $ConfigPath -Raw | ConvertFrom-Json
            if ($old.PSObject.Properties['Azure']) {
                foreach ($p in 'SubscriptionId','TenantId','AppId','AppObjectId','SpObjectId','AuthCertThumbprint','RoleName','BootstrapClientId') {
                    if ($old.Azure.PSObject.Properties[$p]) { $cfg.Azure[$p] = $old.Azure.$p }
                }
                if ($old.Azure.PSObject.Properties['DnsZoneIds']) { $cfg.Azure.DnsZoneIds = @($old.Azure.DnsZoneIds) }
            }
            if ($old.PSObject.Properties['Acme'] -and $old.Acme.PSObject.Properties['Server']) { $cfg.Acme.Server = $old.Acme.Server }
        } catch { }
    }
    $cfg.Acme.Domains  = Split-Csv $txtDomains.Text
    $cfg.Acme.Contact  = $txtContact.Text.Trim()
    $svc = @(); if ($chkIIS.Checked) { $svc += 'IIS' }; if ($chkSMTP.Checked) { $svc += 'SMTP' }
    $cfg.Exchange.Servers  = Split-Csv $txtServers.Text
    $cfg.Exchange.Services = ($svc -join ',')
    if ($txtHome.Text.Trim()) { $cfg.Paths.Home = $txtHome.Text.Trim() }
    $cfg.Notify.SmtpServer = $txtSmtp.Text.Trim()
    $p = 25; [void][int]::TryParse($txtPort.Text.Trim(), [ref]$p); $cfg.Notify.Port = $p
    $cfg.Notify.From = $txtFrom.Text.Trim()
    $cfg.Notify.To   = Split-Csv $txtTo.Text
    $w = 14; [void][int]::TryParse($txtWarn.Text.Trim(), [ref]$w); $cfg.Notify.WarnDaysBeforeExpiry = $w
    $cfg.Notify.SendSuccessMail = [bool]$chkSuccess.Checked
    if ($txtTaskName.Text.Trim()) { $cfg.Task.Name = $txtTaskName.Text.Trim() }
    if ($txtTaskTime.Text.Trim()) { $cfg.Task.Time = $txtTaskTime.Text.Trim() }

    # DNS provider
    $prov = if ($cboProvider.SelectedItem) { [string]$cboProvider.SelectedItem } else { 'Azure' }
    $cfg.Acme.DnsProvider = $prov
    if ($prov -eq 'Azure') {
        # only overwrite the (possibly auto-detected) stored subscription when the user actually typed one;
        # an empty field must NOT wipe the value discovered during -Setup
        if ($txtSub.Text.Trim()) { $cfg.Azure.SubscriptionId = $txtSub.Text.Trim() }
        if ($txtApp.Text.Trim()) { $cfg.Azure.AppDisplayName = $txtApp.Text.Trim() }
    } else {
        # store only NON-secret plugin args in config; secrets go to Posh-ACME at issue time
        $pa = @{}
        if ($script:CurFields.Count -ge 1 -and -not $script:CurFields[0].Secure -and $txtCred1.Text.Trim()) { $pa[$script:CurFields[0].Key] = $txtCred1.Text.Trim() }
        if ($script:CurFields.Count -ge 2 -and -not $script:CurFields[1].Secure -and $txtCred2.Text.Trim()) { $pa[$script:CurFields[1].Key] = $txtCred2.Text.Trim() }
        $cfg.Acme.DnsPluginArgs = $pa
    }
    # remember the chosen UI language (ignored by the engine)
    $cfg.Ui = [ordered]@{ Language = $script:Lang }
    return $cfg
}

function Get-ProviderSecrets {
    # secret credential values currently entered (non-Azure) - for transient hand-off to the engine
    $s = @{}
    if ($cboProvider.SelectedItem -and $cboProvider.SelectedItem -ne 'Azure') {
        if ($script:CurFields.Count -ge 1 -and $script:CurFields[0].Secure -and $txtCred1.Text.Length) { $s[$script:CurFields[0].Key] = $txtCred1.Text }
        if ($script:CurFields.Count -ge 2 -and $script:CurFields[1].Secure -and $txtCred2.Text.Length) { $s[$script:CurFields[1].Key] = $txtCred2.Text }
    }
    return $s
}

function Save-GuiConfig {
    $cfg = Get-ConfigFromForm
    $dir = Split-Path $ConfigPath -Parent
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Force $dir | Out-Null }
    ($cfg | ConvertTo-Json -Depth 6) | Set-Content -Path $ConfigPath -Encoding UTF8
    Write-GuiLog ((T 'logSaved') -f $ConfigPath)
}

function Load-GuiConfig {
    if (-not (Test-Path $ConfigPath)) { Write-GuiLog ((T 'logNoConfig') -f $ConfigPath); Set-FormFromConfig (New-Config); return }
    $cfg = Get-Content $ConfigPath -Raw | ConvertFrom-Json
    Set-FormFromConfig $cfg
    if ($cfg.PSObject.Properties['Ui'] -and $cfg.Ui.PSObject.Properties['Language'] -and $cfg.Ui.Language) { Set-Language ([string]$cfg.Ui.Language) }
    Write-GuiLog ((T 'logLoaded') -f $ConfigPath)
}

# --------------------------------------------------------------- run engine
# Long actions (setup, renew): keep a console open so the user can watch progress and
# complete the Azure device-code sign-in.
function Start-Engine([string[]]$engineArgs) {
    $psArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$EnginePath`"") + $engineArgs + @('-ConfigPath', "`"$ConfigPath`"")
    $full = @('-NoExit') + $psArgs
    Start-Process powershell.exe -ArgumentList ($full -join ' ')
}

# Quick actions (status, task, test mail) run hidden and stream their output into the
# log WITHOUT blocking the UI: the child writes to temp files, and the WinForms timer
# above tails them on the UI thread, so the window stays responsive and lines appear
# as they arrive (instead of the window freezing and dumping everything at the end).
function Read-CaptureStream([string]$PathKey, [string]$PosKey) {
    $path = $script:Cap[$PathKey]
    if (-not $path -or -not (Test-Path $path)) { return }
    try {
        $fs = [System.IO.File]::Open($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        [void]$fs.Seek([long]$script:Cap[$PosKey], [System.IO.SeekOrigin]::Begin)
        $sr = New-Object System.IO.StreamReader($fs)
        $text = $sr.ReadToEnd()
        $script:Cap[$PosKey] = $fs.Position
        $sr.Dispose(); $fs.Dispose()
        if ($text) { ($text -split "`r?`n") | Where-Object { $_.Trim() } | ForEach-Object { Write-GuiLog $_ } }
    } catch { }
}
function Drain-Capture {
    if (-not $script:Cap.Proc) { return }
    Read-CaptureStream 'Out' 'OutPos'
    Read-CaptureStream 'Err' 'ErrPos'
    if ($script:Cap.Proc.HasExited) {
        Read-CaptureStream 'Out' 'OutPos'   # final drain for anything flushed at exit
        Read-CaptureStream 'Err' 'ErrPos'
        $logTimer.Stop()
        if ($script:Cap.DoneMsg) { Write-GuiLog $script:Cap.DoneMsg }
        foreach ($k in 'Out', 'Err') { $p = $script:Cap[$k]; if ($p -and (Test-Path $p)) { Remove-Item $p -Force -ErrorAction SilentlyContinue } }
        $script:Cap.Proc = $null
    }
}
function Start-EngineCapture([string[]]$engineArgs, [string]$DoneMsg) {
    if ($script:Cap.Proc -and -not $script:Cap.Proc.HasExited) { Write-GuiLog (T 'logBusy'); return }
    $psArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$EnginePath`"") + $engineArgs + @('-ConfigPath', "`"$ConfigPath`"")
    $out = Join-Path $env:TEMP ("acme-out-{0}.log" -f [guid]::NewGuid())
    $err = Join-Path $env:TEMP ("acme-err-{0}.log" -f [guid]::NewGuid())
    $script:Cap.Out = $out; $script:Cap.Err = $err; $script:Cap.OutPos = 0; $script:Cap.ErrPos = 0; $script:Cap.DoneMsg = $DoneMsg
    $script:Cap.Proc = Start-Process powershell.exe -ArgumentList ($psArgs -join ' ') -WindowStyle Hidden -PassThru -RedirectStandardOutput $out -RedirectStandardError $err
    $logTimer.Start()
}

function Invoke-Status { Write-GuiLog (T 'logStatus'); Start-EngineCapture @('-Status') }

# Open an Explorer window on the engine's log directory (<Home>\logs); if this month's
# log file exists, select it so the focus lands on the log itself.
function Open-LogDir {
    $homeDir = if ($txtHome.Text.Trim()) { $txtHome.Text.Trim() } else { 'C:\Tools\AcmeExchange' }
    $logDir = Join-Path $homeDir 'logs'
    $logFile = Join-Path $logDir ("acme-{0}.log" -f (Get-Date -Format 'yyyyMM'))
    if (Test-Path $logFile)     { Start-Process explorer.exe "/select,`"$logFile`"" }
    elseif (Test-Path $logDir)  { Start-Process explorer.exe "`"$logDir`"" }
    else { Write-GuiLog ((T 'logNoLogDir') -f $logDir) }
}

# About dialog with a clickable blog link.
function Show-About {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = T 'aboutTitle'
    $dlg.FormBorderStyle = 'FixedDialog'; $dlg.StartPosition = 'CenterParent'
    $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false; $dlg.ShowInTaskbar = $false
    $dlg.ClientSize = New-Object System.Drawing.Size(470, 200)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = (T 'aboutBody') -f $BundleRoot, $ConfigPath
    $lbl.Location = New-Object System.Drawing.Point(16, 16); $lbl.Size = New-Object System.Drawing.Size(438, 120)
    $dlg.Controls.Add($lbl)
    $link = New-Object System.Windows.Forms.LinkLabel
    $link.Text = $script:BlogUrl; $link.AutoSize = $true
    $link.Location = New-Object System.Drawing.Point(16, 140)
    $link.Add_LinkClicked({ try { Start-Process $script:BlogUrl } catch { } })
    $dlg.Controls.Add($link)
    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'OK'; $ok.DialogResult = 'OK'
    $ok.Location = New-Object System.Drawing.Point(374, 162); $ok.Size = New-Object System.Drawing.Size(80, 28)
    $dlg.Controls.Add($ok); $dlg.AcceptButton = $ok
    [void]$dlg.ShowDialog($form); $dlg.Dispose()
}

# --------------------------------------------------------------- events
$miLoad.Add_Click({ try { Load-GuiConfig } catch { Write-GuiLog ((T 'err') -f $_.Exception.Message) } })
$miSave.Add_Click({ try { Save-GuiConfig } catch { Write-GuiLog ((T 'err') -f $_.Exception.Message) } })
$miExit.Add_Click({ $form.Close() })
$miDe.Add_Click({ Set-Language 'de'; Write-GuiLog (T 'logLangSet') })
$miEn.Add_Click({ Set-Language 'en'; Write-GuiLog (T 'logLangSet') })
$miStatusM.Add_Click({ try { Invoke-Status } catch { Write-GuiLog ((T 'err') -f $_.Exception.Message) } })
$miLogDir.Add_Click({ try { Open-LogDir } catch { Write-GuiLog ((T 'err') -f $_.Exception.Message) } })
$miAbout.Add_Click({ try { Show-About } catch { Write-GuiLog ((T 'err') -f $_.Exception.Message) } })

$btnDiscover.Add_Click({
    try {
        Write-GuiLog (T 'logDiscover')
        if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
            Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
        }
        $srv = @(Get-ExchangeServer | Where-Object { $_.ServerRole -match 'Mailbox' } | ForEach-Object { $_.Name.ToUpper() })
        $txtServers.Text = ($srv -join ', ')
        Write-GuiLog ((T 'logFound') -f ($srv -join ', '))
    } catch { Write-GuiLog ((T 'err') -f $_.Exception.Message) }
})

# Run now - Step 1: prepare Azure only (no certificate, no server change). Needs the device-code sign-in.
$btnSetup.Add_Click({
    try {
        Save-GuiConfig
        $r = [System.Windows.Forms.MessageBox]::Show((T 'mbSetupBody') + (T 'mbSetupTail'), (T 'mbSetupTitle'), 'YesNo', 'Question')
        if ($r -ne 'Yes') { return }
        $script:CfgWatchMtime = if (Test-Path $ConfigPath) { (Get-Item $ConfigPath).LastWriteTime } else { [datetime]'1970-01-01' }
        $script:ExpectReload = $true
        Start-Engine @('-Setup', '-NonInteractive')
        Write-GuiLog (T 'logSetupStarted')
    } catch { Write-GuiLog ((T 'err') -f $_.Exception.Message) }
})

# Run now - Step 2: issue the certificate and install it on the servers (also the renewal path).
$btnRenew.Add_Click({
    try {
        Save-GuiConfig
        $srv = (Split-Csv $txtServers.Text) -join ', '
        if ($chkStaging.Checked) { $detail = (T 'mbRenewStaging') -f $srv; $icon = 'Question' }
        else                     { $detail = (T 'mbRenewProd')    -f $srv; $icon = 'Warning' }
        $r = [System.Windows.Forms.MessageBox]::Show($detail + (T 'mbConsoleTail'), (T 'mbRenewTitle'), 'YesNo', $icon)
        if ($r -ne 'Yes') { return }
        $a = @('-Renew')
        if ($chkStaging.Checked) { $a += @('-Staging', '-SkipInstall') }   # staging only issues, never installs
        # hand non-Azure API secrets to the engine transiently (consumed + deleted there; then kept
        # encrypted by Posh-ACME). The child process inherits ACME_PLUGIN_SECRETS from us.
        $secrets = Get-ProviderSecrets
        try {
            if ($secrets.Count) {
                $secFile = Join-Path $env:TEMP ("acme-sec-{0}.json" -f [guid]::NewGuid())
                ($secrets | ConvertTo-Json) | Set-Content -Path $secFile -Encoding UTF8
                $env:ACME_PLUGIN_SECRETS = $secFile
                Write-GuiLog (T 'logSecrets')
            }
            Start-Engine $a
        } finally { $env:ACME_PLUGIN_SECRETS = $null }
        $mode = if ($chkStaging.Checked) { T 'modeStaging' } else { T 'modeProd' }
        Write-GuiLog ((T 'logRenewStarted') -f $mode)
    } catch { Write-GuiLog ((T 'err') -f $_.Exception.Message) }
})

$btnTaskAdd.Add_Click({
    try { Save-GuiConfig; Start-EngineCapture @('-InstallTask') (T 'logTaskAdd') }
    catch { Write-GuiLog ((T 'err') -f $_.Exception.Message) }
})
$btnTaskDel.Add_Click({
    try { Start-EngineCapture @('-RemoveTask') (T 'logTaskDel') }
    catch { Write-GuiLog ((T 'err') -f $_.Exception.Message) }
})
$btnTestMail.Add_Click({
    try {
        Save-GuiConfig
        Write-GuiLog ((T 'logTestMail') -f $txtTo.Text, $txtSmtp.Text, $txtPort.Text)
        Start-EngineCapture @('-TestMail')
    } catch { Write-GuiLog ((T 'err') -f $_.Exception.Message) }
})

# first-start fill-in help: grey placeholder cues (not saved) + a computed App-name suggestion
function Update-AppCue {
    $d = Split-Csv $txtDomains.Text
    $first = if ($d.Count) { $d[0] -replace '^\*\.', '' } else { 'contoso.com' }
    Set-Cue $txtApp "ACME-DNS-$first"
}
function Set-InitialCues {
    Set-Cue $txtContact 'acme@contoso.com'
    Set-Cue $txtDomains '*.contoso.com, contoso.com'
    Set-Cue $txtServers 'EX01, EX02'
    Set-Cue $txtSub 'abcdef12-3456-7890-abcd-ef1234567890'
    Set-Cue $txtFrom 'acme@contoso.com'
    Set-Cue $txtTo 'admin@contoso.com, ops@contoso.com'
    Update-AppCue
}
$txtDomains.Add_TextChanged({ Update-AppCue })
$form.Add_Shown({ Set-InitialCues })

# initial language + load
Set-Language $script:Lang
try { Load-GuiConfig } catch { }
Write-GuiLog ((T 'logBundle') -f $BundleRoot)
Write-GuiLog ((T 'logConfig') -f $ConfigPath)

if ($SelfTest) {
    Write-Host 'GUI self-test: form built OK'
    if ($env:ACMEGUI_SHOT) {
        try {
            if ($env:ACMEGUI_FILLLOG) {
                # dev-only: exercise the log pane with long lines to check the horizontal scrollbar
                1..12 | ForEach-Object { Write-GuiLog "[$_] Sample very long log line to test horizontal scrolling: C:\Program Files\Vendor\A Very Long Folder Name\Invoke-AcmeExchangeCert.ps1 -Renew -ConfigPath C:\Tools\AcmeExchange\config.json (event=$_)" }
            }
            if ($env:ACMEGUI_REALSHOT) {
                # dev-only: show the window on-screen and grab the REAL pixels (incl. native
                # scrollbars, which DrawToBitmap does not paint).
                $form.StartPosition = 'Manual'; $form.Location = New-Object System.Drawing.Point(80, 40); $form.ShowInTaskbar = $true; $form.TopMost = $true
                $form.Show(); 1..20 | ForEach-Object { [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 80 }
                Write-Host ("Form bounds: {0}x{1} at {2},{3}  Screen WA: {4}" -f $form.Width, $form.Height, $form.Left, $form.Top, [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea)
                $sb = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
                $bmp = New-Object System.Drawing.Bitmap($sb.Width, $sb.Height)
                $g = [System.Drawing.Graphics]::FromImage($bmp)
                $g.CopyFromScreen(0, 0, 0, 0, $sb.Size)
                $g.Dispose()
            } else {
                $form.StartPosition = 'Manual'; $form.Location = New-Object System.Drawing.Point(-4000, -4000); $form.ShowInTaskbar = $false
                $form.Show(); 1..8 | ForEach-Object { [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 60 }
                $bmp = New-Object System.Drawing.Bitmap($form.Width, $form.Height)
                $form.DrawToBitmap($bmp, (New-Object System.Drawing.Rectangle(0, 0, $form.Width, $form.Height)))
            }
            $bmp.Save($env:ACMEGUI_SHOT, [System.Drawing.Imaging.ImageFormat]::Png)
            $form.Close()
            Write-Host "Screenshot: $($env:ACMEGUI_SHOT)"
        } catch { Write-Host "Screenshot failed: $($_.Exception.Message)" }
    }
    $form.Dispose(); return
}
[void]$form.ShowDialog()
$form.Dispose()
