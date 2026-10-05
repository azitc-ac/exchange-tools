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
