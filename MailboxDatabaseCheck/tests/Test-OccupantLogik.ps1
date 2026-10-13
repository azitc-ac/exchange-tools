# Prüft die Sammellogik von Test-MailboxDatabaseEmpty.ps1 ohne Exchange: die Funktionen
# werden per AST herausgelöst (wie in MailContactEditor/tests), Get-Mailbox und
# Get-MailboxStatistics durch Attrappen ersetzt, die aus einer Tabelle antworten.
$ErrorActionPreference = 'Stop'
$hier = if ($PSScriptRoot) { $PSScriptRoot }
        elseif ($MyInvocation.MyCommand.Path) { Split-Path $MyInvocation.MyCommand.Path -Parent }
        else { (Get-Location).Path }
$skript = Join-Path (Split-Path $hier -Parent) 'Test-MailboxDatabaseEmpty.ps1'

$ast = [System.Management.Automation.Language.Parser]::ParseFile($skript, [ref]$null, [ref]$null)
foreach ($fn in 'ConvertTo-PsLiteral', 'New-Occupant', 'Get-DatabaseOccupant') {
    $def = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $args[0].Name -eq $fn }, $false)
    if (-not $def) { throw "Funktion nicht gefunden: $fn" }
    Invoke-Expression $def[0].Extent.Text
}

# --- Attrappen -------------------------------------------------------------------------
# $script:Fx[<Datenbank>][<Art>] = Objekte; Art: '' (ohne Schalter), Arbitration, AuditLog,
# Monitoring, PublicFolder, Archive. $script:Stats[<Datenbank>] = Statistik-Objekte.
$script:Fx = @{}; $script:Stats = @{}; $script:StatsAufrufe = 0
function Get-Mailbox {
    param($Database, $ResultSize, [switch]$Arbitration, [switch]$AuditLog, [switch]$Monitoring, [switch]$PublicFolder, [switch]$Archive)
    $art = ''
    foreach ($s in 'Arbitration', 'AuditLog', 'Monitoring', 'PublicFolder', 'Archive') {
        if ((Get-Variable $s -ValueOnly).IsPresent) { $art = $s }
    }
    $db = $script:Fx[$Database]
    if ($db -and $db.ContainsKey($art)) { $db[$art] }
}
function Get-MailboxStatistics {
    param($Database)
    $script:StatsAufrufe++
    $script:Stats[$Database]
}
function Mbx($name, $typ = 'UserMailbox', $archivDb = $null) {
    [pscustomobject]@{ DisplayName = $name; Identity = "contoso.com/Users/$name"; RecipientTypeDetails = $typ; ArchiveDatabase = $archivDb }
}

$pass = 0; $fail = 0
function T($name, $ist, $soll) {
    if ("$ist" -eq "$soll") { $script:pass++; Write-Output "  OK   $name" }
    else { $script:fail++; Write-Output "  FEHL $name : ist '$ist', soll '$soll'" }
}

Write-Output "`n######## Leere Datenbank ########"
$r = @(Get-DatabaseOccupant -Database 'DB-LEER')
T 'keine Funde'                       $r.Count 0

Write-Output "`n######## Discovery nur einmal, unter eigenem Typ ########"
$script:Fx['DB1'] = @{ '' = @((Mbx 'Erika'), (Mbx 'DiscoverySearchMailbox' 'DiscoveryMailbox')) }
$r = @(Get-DatabaseOccupant -Database 'DB1')
T 'zwei Funde, nicht drei'            $r.Count 2
T 'Discovery als Discovery'           (@($r | Where-Object Typ -eq 'Discovery').Count) 1
T 'Erika als Postfach'                (@($r | Where-Object Typ -eq 'Postfach').Name) 'Erika'

Write-Output "`n######## Sonderpostfächer und ihre Befehle ########"
$script:Fx['DB2'] = @{
    'Arbitration'  = @((Mbx 'SystemMailbox{1f05a927}' 'ArbitrationMailbox'))
    'AuditLog'     = @((Mbx 'SystemMailbox{8cc370d3}' 'AuditLogMailbox'))
    'PublicFolder' = @((Mbx 'PF-Root' 'PublicFolderMailbox'))
    'Monitoring'   = @((Mbx 'HealthMailbox1' 'MonitoringMailbox'))
}
$r = @(Get-DatabaseOccupant -Database 'DB2')
T 'vier Funde'                        $r.Count 4
$arb = $r | Where-Object Typ -eq 'Arbitration'
T 'Arbitration blockiert'             $arb.Blockiert 'ja'
T 'Arbitration: letzter erlaubt'      ($arb.Befehl -match '-RemoveLastArbitrationMailboxAllowed') $true
T 'Arbitration: OAB erlaubt'          ($arb.Befehl -match '-RemoveArbitrationMailboxWithOABsAllowed') $true
T 'AuditLog-Befehl'                   (($r | Where-Object Typ -eq 'AuditLog').Befehl -match '^Remove-Mailbox .* -AuditLog$') $true
T 'PublicFolder-Befehl'               (($r | Where-Object Typ -eq 'PublicFolder').Befehl -match '-PublicFolder') $true
T 'Monitoring ist unklar'             ($r | Where-Object Typ -eq 'Monitoring').Blockiert 'unklar'

Write-Output "`n######## Archive: beide Wege, keine Doppelten ########"
$erika = Mbx 'Erika' 'UserMailbox' 'DB3'
$max   = Mbx 'Max'   'UserMailbox' 'DB3'
$script:Fx['DB3'] = @{ 'Archive' = @($erika) }
# Max' Hauptpostfach liegt woanders, sein Archiv in DB3: nur über ArchiveDatabase zu finden.
$r = @(Get-DatabaseOccupant -Database 'DB3' -AllMailboxes @($erika, $max, (Mbx 'Fremd' 'UserMailbox' 'DB9')))
$arch = @($r | Where-Object Typ -eq 'Archiv')
T 'zwei Archive'                      $arch.Count 2
T 'Max über ArchiveDatabase gefunden' (@($arch.Name) -contains 'Max') $true
T 'fremdes Archiv nicht dabei'        (@($arch.Name) -contains 'Fremd') $false
T 'Archiv-Befehl mit -Archive'        ($arch[0].Befehl -match '^Disable-Mailbox .* -Archive$') $true

Write-Output "`n######## Getrennte Postfächer ########"
$script:Stats['DB4'] = @(
    [pscustomobject]@{ DisplayName = 'Alt';    MailboxGuid = [guid]::NewGuid(); DisconnectReason = 'SoftDeleted' }
    [pscustomobject]@{ DisplayName = 'Weg';    MailboxGuid = [guid]::NewGuid(); DisconnectReason = 'Disabled' }
    [pscustomobject]@{ DisplayName = 'Aktiv';  MailboxGuid = [guid]::NewGuid(); DisconnectReason = $null }
)
$r = @(Get-DatabaseOccupant -Database 'DB4')
T 'soft-deleted und disabled'         $r.Count 2
T 'blockieren nicht'                  (@($r | Where-Object Blockiert -ne 'nein').Count) 0

$script:StatsAufrufe = 0
$r = @(Get-DatabaseOccupant -Database 'DB4' -Mounted $false)
T 'nicht eingehängt: keine Statistik' $script:StatsAufrufe 0
T 'nicht eingehängt: keine Funde'     $r.Count 0

Write-Output "`n######## Quoting ########"
T 'Apostroph verdoppelt'              (ConvertTo-PsLiteral "contoso.com/Users/O'Brien") "'contoso.com/Users/O''Brien'"
$script:Fx['DB5'] = @{ '' = @((Mbx "O'Brien")) }
$r = @(Get-DatabaseOccupant -Database 'DB5')
T 'Befehl mit sicherem Quoting'       ($r[0].Befehl -match "'contoso\.com/Users/O''Brien'") $true

Write-Output "`n================ Bestanden: $pass   Fehlgeschlagen: $fail ================"
if ($fail -gt 0) { exit 1 }
