# Prüft Get-DefaultSmtpRecommendation aus Repair-ExchangeSmtpCertificate.ps1 ohne Exchange:
# Soll beim Aktivieren für SMTP auch das interne Transportzertifikat (Standard-SMTP)
# überschrieben werden? Die Funktion wird per AST herausgelöst.
$ErrorActionPreference = 'Stop'
$hier = if ($PSScriptRoot) { $PSScriptRoot }
        elseif ($MyInvocation.MyCommand.Path) { Split-Path $MyInvocation.MyCommand.Path -Parent }
        else { (Get-Location).Path }
$skript = Join-Path (Split-Path $hier -Parent) 'Repair-ExchangeSmtpCertificate.ps1'

$ast = [System.Management.Automation.Language.Parser]::ParseFile($skript, [ref]$null, [ref]$null)
$def = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $args[0].Name -eq 'Get-DefaultSmtpRecommendation' }, $false)
if (-not $def) { throw 'Funktion nicht gefunden: Get-DefaultSmtpRecommendation' }
Invoke-Expression $def[0].Extent.Text

$pass = 0; $fail = 0
function T($name, $ist, $soll) {
    if ("$ist" -eq "$soll") { $script:pass++; Write-Output "  OK   $name" }
    else { $script:fail++; Write-Output "  FEHL $name : ist '$ist', soll '$soll'" }
}

$jetzt  = [datetime]'2026-11-18'
$ziel   = [pscustomobject]@{ Thumbprint = 'NEU0000000000000000000000000000000000000' }
function Std($tp, $selbst, $imSpeicher, $bis, $fehler = $null) {
    [pscustomobject]@{ Thumbprint = $tp; Subject = "CN=$tp"; IsSelfSigned = $selbst; InStore = $imSpeicher; NotAfter = [datetime]$bis; Error = $fehler }
}
function A($current) { (Get-DefaultSmtpRecommendation -Current $current -Target $ziel -Now $jetzt).Action }

Write-Output "`n######## Normalfall: selbstsigniert und gültig -> behalten ########"
T 'selbstsigniert gültig'              (A (Std 'SELF' $true $true '2030-01-01')) 'Keep'

Write-Output "`n######## Kaputter Standard -> überschreiben ########"
T 'nicht gesetzt'                      (A $null) 'Overwrite'
T 'nicht mehr im Speicher (12013)'     (A (Std 'WEG' $true $false '2030-01-01')) 'Overwrite'
T 'abgelaufen, selbstsigniert'         (A (Std 'ALT' $true $true '2026-01-01')) 'Overwrite'

Write-Output "`n######## Öffentliches Zertifikat als Standard -> mitnehmen ########"
T 'CA-Zertifikat gültig, wird ersetzt' (A (Std 'CA' $false $true '2027-01-01')) 'Overwrite'

Write-Output "`n######## Sonderfälle ########"
T 'Ziel ist schon Standard'            (A (Std $ziel.Thumbprint $false $true '2027-01-01')) 'AlreadyDefault'
T 'nicht lesbar'                       (A ([pscustomobject]@{ Error = 'LDAP down' })) 'Unknown'
# Reihenfolge: ein selbstsigniertes, aber verschwundenes Zertifikat darf NICHT als "behalten" durchgehen.
T 'selbstsigniert, aber weg'           (A (Std 'WEG2' $true $false '2030-01-01')) 'Overwrite'

Write-Output "`n================ Bestanden: $pass   Fehlgeschlagen: $fail ================"
if ($fail -gt 0) { exit 1 }
