# Prüft die Adress-Zerlegung von Edit-MailContactAddresses.ps1, ohne das Werkzeug
# zu starten: die Funktionen werden per AST herausgelöst (dasselbe Muster, das
# AcmeExchangeSetup.ps1 für die Engine nutzt), damit der Prolog nicht anläuft und
# keine Anmeldung an Exchange Online verlangt.
$ErrorActionPreference = 'Stop'
$hier = if ($PSScriptRoot) { $PSScriptRoot }
        elseif ($MyInvocation.MyCommand.Path) { Split-Path $MyInvocation.MyCommand.Path -Parent }
        else { (Get-Location).Path }
$skript = Join-Path (Split-Path $hier -Parent) 'Edit-MailContactAddresses.ps1'

$ast = [System.Management.Automation.Language.Parser]::ParseFile($skript, [ref]$null, [ref]$null)
foreach ($fn in 'ConvertFrom-ProxyAddress', 'ConvertTo-ProxyAddress', 'Test-SmtpAddress', 'ConvertTo-PlainAddress') {
    $def = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $args[0].Name -eq $fn }, $false)
    if (-not $def) { throw "Funktion nicht gefunden: $fn" }
    Invoke-Expression $def[0].Extent.Text
}

$pass = 0; $fail = 0
function T($name, $ist, $soll) {
    if ($ist -eq $soll) { $script:pass++; Write-Output "  OK   $name" }
    else { $script:fail++; Write-Output "  FEHL $name : ist '$ist', soll '$soll'" }
}

Write-Output "`n######## Zerlegen ########"
$p = ConvertFrom-ProxyAddress 'SMTP:max@kunde.de'
T 'SMTP gross -> primaer'        $p.IsPrimary $true
T 'SMTP gross -> ist SMTP'       $p.IsSmtp    $true
T 'SMTP gross -> Adresse'        $p.Address   'max@kunde.de'

$p = ConvertFrom-ProxyAddress 'smtp:zweit@kunde.de'
T 'smtp klein -> nicht primaer'  $p.IsPrimary $false
T 'smtp klein -> ist SMTP'       $p.IsSmtp    $true

$p = ConvertFrom-ProxyAddress 'X500:/o=ExchangeLabs/ou=Exchange Administrative Group/cn=Recipients/cn=abc'
T 'X500 -> kein SMTP'            $p.IsSmtp    $false
T 'X500 -> nicht primaer'        $p.IsPrimary $false
T 'X500 -> Praefix erhalten'     $p.Prefix    'X500'

$p = ConvertFrom-ProxyAddress 'SIP:max@kunde.de'
T 'SIP -> kein SMTP'             $p.IsSmtp    $false
T 'SIP -> Praefix erhalten'      $p.Prefix    'SIP'

$p = ConvertFrom-ProxyAddress 'ohnepraefix@kunde.de'
T 'ohne Praefix -> smtp'         $p.IsSmtp    $true
T 'ohne Praefix -> sekundaer'    $p.IsPrimary $false

Write-Output "`n######## Hin und zurueck - nichts darf sich veraendern ########"
$proben = @(
    'SMTP:max@kunde.de'
    'smtp:m.mustermann@kunde.de'
    'smtp:max@partner.de'
    'X500:/o=ExchangeLabs/ou=Exchange Administrative Group/cn=Recipients/cn=7f2a'
    'SIP:max@kunde.de'
)
foreach ($roh in $proben) {
    $zurueck = ConvertTo-ProxyAddress (ConvertFrom-ProxyAddress $roh)
    T "Rundlauf $roh" $zurueck $roh
}

Write-Output "`n######## Primaer umschalten ########"
$liste = $proben | ForEach-Object { ConvertFrom-ProxyAddress $_ }
# zweite SMTP-Adresse zur primaeren machen, wie es der Knopf 'Als primaer' tut
foreach ($a in $liste) { if ($a.IsSmtp) { $a.IsPrimary = $false } }
($liste | Where-Object { $_.Address -eq 'm.mustermann@kunde.de' }).IsPrimary = $true
$neu = @($liste | ForEach-Object { ConvertTo-ProxyAddress $_ })
T 'genau eine grosse SMTP'       (@($neu | Where-Object { $_ -cmatch '^SMTP:' }).Count) 1
T 'neue primaere stimmt'         ($neu | Where-Object { $_ -cmatch '^SMTP:' }) 'SMTP:m.mustermann@kunde.de'
T 'alte primaere ist sekundaer'  ($neu -contains 'smtp:max@kunde.de') $true
T 'X500 unveraendert dabei'      ($neu -contains 'X500:/o=ExchangeLabs/ou=Exchange Administrative Group/cn=Recipients/cn=7f2a') $true
T 'SIP unveraendert dabei'       ($neu -contains 'SIP:max@kunde.de') $true
T 'Anzahl unveraendert'          $neu.Count 5

Write-Output "`n######## Adressen pruefen ########"
T 'gueltige Adresse'             (Test-SmtpAddress 'a@b.de')  $true
T 'leer ist ungueltig'           (Test-SmtpAddress '')        $false
T 'nur Text ist ungueltig'       (Test-SmtpAddress 'kein at') $false
T 'Leerzeichen ist ungueltig'    (Test-SmtpAddress '   ')     $false

Write-Output "`n================ Bestanden: $pass   Fehlgeschlagen: $fail ================"
if ($fail -gt 0) { exit 1 }
