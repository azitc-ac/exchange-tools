# Prüft ohne Exchange zwei reine Funktionen aus Repair-ExchangeSmtpCertificate.ps1:
#  - Get-DefaultSmtpRecommendation: was mit dem internen Transportzertifikat (Standard-SMTP)
#    geschehen soll - behalten, auf das Ziel, auf ein selbstsigniertes zurück, neu anlegen.
#  - ConvertFrom-NetshSslCert: liest die Bindung der Website "Exchange Back End" (Port 444).
# Die Funktionen werden per AST herausgelöst.
$ErrorActionPreference = 'Stop'
$hier = if ($PSScriptRoot) { $PSScriptRoot }
        elseif ($MyInvocation.MyCommand.Path) { Split-Path $MyInvocation.MyCommand.Path -Parent }
        else { (Get-Location).Path }
$skript = Join-Path (Split-Path $hier -Parent) 'Repair-ExchangeSmtpCertificate.ps1'

$ast = [System.Management.Automation.Language.Parser]::ParseFile($skript, [ref]$null, [ref]$null)
foreach ($fn in 'Get-DefaultSmtpRecommendation', 'ConvertFrom-NetshSslCert') {
    $def = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $args[0].Name -eq $fn }, $false)
    if (-not $def) { throw "Funktion nicht gefunden: $fn" }
    Invoke-Expression $def[0].Extent.Text
}

$pass = 0; $fail = 0
function T($name, $ist, $soll) {
    if ("$ist" -eq "$soll") { $script:pass++; Write-Output "  OK   $name" }
    else { $script:fail++; Write-Output "  FEHL $name : ist '$ist', soll '$soll'" }
}

$jetzt = [datetime]'2026-11-18'
$ziel  = [pscustomobject]@{ Thumbprint = 'NEU0000000000000000000000000000000000000' }
function Std($tp, $selbst, $imSpeicher, $bis, $fehler = $null) {
    [pscustomobject]@{ Thumbprint = $tp; Subject = "CN=$tp"; IsSelfSigned = $selbst; InStore = $imSpeicher; NotAfter = [datetime]$bis; Error = $fehler }
}
function E($current, $selbstDa) { Get-DefaultSmtpRecommendation -Current $current -Target $ziel -SelfSignedAvailable $selbstDa -Now $jetzt }

Write-Output "`n######## Normalfall: selbstsigniert und gültig -> behalten ########"
T 'selbstsigniert gültig'                       (E (Std 'SELF' $true $true '2030-01-01') $true).Action 'Keep'
T '... ist nicht kaputt'                        (E (Std 'SELF' $true $true '2030-01-01') $true).Broken $false

Write-Output "`n######## Kaputter Standard ########"
T 'nicht gesetzt, Selbstsigniertes da'          (E $null $true).Action 'RestoreSelfSigned'
T 'nicht gesetzt, keins da -> neu anlegen'      (E $null $false).Action 'CreateSelfSigned'
T 'nicht mehr im Speicher (12013)'              (E (Std 'WEG' $true $false '2030-01-01') $true).Action 'RestoreSelfSigned'
T '... gilt als kaputt'                         (E (Std 'WEG' $true $false '2030-01-01') $true).Broken $true
T 'abgelaufen nach 5 Jahren, keins da'          (E (Std 'ALT' $true $true '2026-01-01') $false).Action 'CreateSelfSigned'
T 'abgelaufenes CA-Zertifikat, Selbstsign. da'  (E (Std 'CA' $false $true '2026-01-01') $true).Action 'RestoreSelfSigned'

Write-Output "`n######## Öffentliches Zertifikat als Standard ########"
T 'CA gültig, Selbstsigniertes da -> zurück'    (E (Std 'CA' $false $true '2027-01-01') $true).Action 'RestoreSelfSigned'
T '... ist nicht kaputt'                        (E (Std 'CA' $false $true '2027-01-01') $true).Broken $false
T 'CA gültig, keins da -> aufs Ziel'            (E (Std 'CA' $false $true '2027-01-01') $false).Action 'Overwrite'
T 'Ziel ist CA-Standard, Selbstsign. da'        (E (Std $ziel.Thumbprint $false $true '2027-01-01') $true).Action 'RestoreSelfSigned'
T 'Ziel ist CA-Standard, keins da'              (E (Std $ziel.Thumbprint $false $true '2027-01-01') $false).Action 'AlreadyDefault'

Write-Output "`n######## Sonderfälle ########"
T 'Ziel ist selbstsignierter Standard'          (E (Std $ziel.Thumbprint $true $true '2030-01-01') $true).Action 'AlreadyDefault'
T 'nicht lesbar'                                (E ([pscustomobject]@{ Error = 'LDAP down' }) $true).Action 'Unknown'
T 'selbstsigniert, aber weg: nicht behalten'    ((E (Std 'WEG2' $true $false '2030-01-01') $true).Action -ne 'Keep') $true

Write-Output "`n######## netsh: Bindung Port 444 ########"
$netsh = @'

SSL Certificate bindings:
-------------------------

    IP:port                      : 0.0.0.0:443
    Certificate Hash             : 5dcc1c305dd2bd87097498903378f4fbb9b9a181
    Application ID               : {4dc3e181-e14b-4a21-b022-59fc669b0914}
    Certificate Store Name       : My

    IP:port                      : 0.0.0.0:444
    Certificate Hash             : 0254b76578c5ba4d55942b2662f4db9eded766b8
    Application ID               : {4dc3e181-e14b-4a21-b022-59fc669b0914}
    Certificate Store Name       : My

    Hostname:port                : mail.contoso.com:443
    Certificate Hash             : 1111111111111111111111111111111111111111
    Application ID               : {00000000-0000-0000-0000-000000000000}
'@
$b = @(ConvertFrom-NetshSslCert -Text $netsh)
T 'drei Bindungen'                              $b.Count 3
$b444 = @($b | Where-Object { $_.IpPort -match ':444$' })
T 'eine auf 444'                                $b444.Count 1
T '444: Fingerabdruck groß'                     $b444[0].Thumbprint '0254B76578C5BA4D55942B2662F4DB9EDED766B8'
T '444: AppId'                                  $b444[0].AppId '{4dc3e181-e14b-4a21-b022-59fc669b0914}'
T 'Hostname-Bindung erkannt'                    $b[2].IpPort 'mail.contoso.com:443'

Write-Output "`n================ Bestanden: $pass   Fehlgeschlagen: $fail ================"
if ($fail -gt 0) { exit 1 }
