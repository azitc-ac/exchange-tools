# Checks the evaluation logic of Test-HybridFreeBusy.ps1 without Exchange: the pure functions
# are extracted via AST (as in MailboxDatabaseCheck/tests) and fed with mock objects shaped like
# the output of Get-SettingOverride, Get-AuthServer, Test-OAuthConnectivity and friends.
[CmdletBinding()]
param([string]$Skript)
$ErrorActionPreference = 'Stop'
$hier = if ($PSScriptRoot) { $PSScriptRoot }
        elseif ($MyInvocation.MyCommand.Path) { Split-Path $MyInvocation.MyCommand.Path -Parent }
        else { (Get-Location).Path }
if (-not $Skript) { $Skript = Join-Path (Split-Path $hier -Parent) 'Test-HybridFreeBusy.ps1' }

$ast = [System.Management.Automation.Language.Parser]::ParseFile($Skript, [ref]$null, [ref]$null)
$fns = 'New-CheckResult', 'ConvertTo-ExchangeBuild', 'Get-DedicatedAppMinBuild', 'Test-ServerBuild',
       'Get-TenantRoutingDomain', 'Test-RemoteRouting', 'Test-IntraOrgConnector', 'Test-DedicatedAppOverride',
       'Test-AuthServerAppId', 'Get-DedicatedAppId', 'Get-OAuthAppId', 'Test-OAuthResult', 'Test-AuthCertificate',
       'Test-GraphRoute', 'Test-EndpointResult', 'Test-EwsLogHit', 'Get-OverallExitCode'
foreach ($fn in $fns) {
    $def = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $args[0].Name -eq $fn }, $false)
    if (-not $def) { throw "Function not found: $fn" }
    Invoke-Expression $def[0].Extent.Text
}
# The two constants the functions use, taken from the script itself.
foreach ($v in 'FirstPartyAppId', 'GraphMinBuild') {
    $a = $ast.Find({ $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] -and "$($args[0].Left)" -eq "`$script:$v" }, $true)
    if (-not $a) { throw "Constant not found: $v" }
    Invoke-Expression $a.Extent.Text
}

$pass = 0; $fail = 0
function T($name, $ist, $soll) {
    if ("$ist" -eq "$soll") { $script:pass++; Write-Output "  OK   $name" }
    else { $script:fail++; Write-Output "  FEHL $name : ist '$ist', soll '$soll'" }
}
function Srv($name, $display, $file = $null) {
    [pscustomobject]@{ Name = $name; DisplayBuild = $(if ($display) { [version]$display }); FileBuild = $(if ($file) { [version]$file }) }
}
function Ov($name, $section, $params, $server = @(), $changed = '2026-01-01') {
    [pscustomobject]@{ Name = $name; ComponentName = 'Global'; SectionName = $section; Parameters = @($params); Server = @($server); WhenChanged = [datetime]$changed }
}
$now = [datetime]'2026-11-23 15:00'
$app = 'a1b2c3d4-0000-4000-8000-000000000001'

Write-Output "`n######## Build lesen ########"
T 'AdminDisplayVersion'               (ConvertTo-ExchangeBuild 'Version 15.2 (Build 2562.17)') '15.2.2562.17'
T 'Dateiversion'                      (ConvertTo-ExchangeBuild '15.2.2562.41') '15.2.2562.41'
T 'Unsinn ergibt nichts'              ([string](ConvertTo-ExchangeBuild 'unknown')) ''

Write-Output "`n######## Schritt 4: Build ########"
$r = @(Test-ServerBuild @((Srv 'EX1' '15.2.2562.17' '15.2.2562.17')))
T 'SE RTM: OK'                        $r[0].Status 'OK'
T 'SE RTM: nur EWS'                   ($r[0].Finding -match 'EWS route only') $true
$r = @(Test-ServerBuild @((Srv 'EX1' '15.2.2562.17' '15.2.2562.41')))
T 'SE HU6 aus ExSetup: Graph'         ($r[0].Finding -match 'Graph route supported') $true
T 'ExSetup schlaegt Anzeige'          ($r[0].Finding -match 'ExSetup') $true
$r = @(Test-ServerBuild @((Srv 'EX2' '15.2.1748.10' '15.2.1748.10')))
T '2019 CU15 ohne HU, Datei: FAIL'    $r[0].Status 'FAIL'
$r = @(Test-ServerBuild @((Srv 'EX2' '15.2.1748.10')))
T '2019 CU15 nur Anzeige: WARN'       $r[0].Status 'WARN'
$r = @(Test-ServerBuild @((Srv 'EX3' '15.2.1748.24' '15.2.1748.24')))
T '2019 CU15 April-HU: OK'            $r[0].Status 'OK'
$r = @(Test-ServerBuild @((Srv 'EX4' '15.2.1258.12' '15.2.1258.12')))
T '2019 CU12 (Datei): FAIL'           $r[0].Status 'FAIL'
$r = @(Test-ServerBuild @((Srv 'EX5' '15.1.2507.55' '15.1.2507.55')))
T '2016 CU23 April-HU: OK'            $r[0].Status 'OK'
$r = @(Test-ServerBuild @((Srv 'EX6' $null $null)))
T 'Build unbekannt: WARN'             $r[0].Status 'WARN'
$r = @(Test-ServerBuild @((Srv 'A' '15.2.2562.17' '15.2.2562.17'), (Srv 'B' '15.2.1748.10')))
T 'eine Zeile je Server'              $r.Count 2

Write-Output "`n######## Schritt 2: Remote-Postfach ########"
T 'ohne Adresse: SKIP'                (Test-RemoteRouting $null '').Status 'SKIP'
T 'kein Objekt: FAIL'                 (Test-RemoteRouting $null 'ben@contoso.com').Status 'FAIL'
T 'Routing in den Tenant: OK'         (Test-RemoteRouting ([pscustomobject]@{ RemoteRoutingAddress = 'SMTP:ben@contoso.mail.onmicrosoft.com' }) 'ben@contoso.com').Status 'OK'
T 'Routing woandershin: FAIL'         (Test-RemoteRouting ([pscustomobject]@{ RemoteRoutingAddress = 'ben@contoso.com' }) 'ben@contoso.com').Status 'FAIL'
T 'Routing-Domain ableiten'           ((Get-TenantRoutingDomain @('contoso.com', '*.contoso.mail.onmicrosoft.com')) -join ',') 'contoso.mail.onmicrosoft.com'

Write-Output "`n######## Schritt 3: IntraOrganizationConnector ########"
function Ioc($en, $doms) { [pscustomobject]@{ Name = 'HybridIOC - 1'; Enabled = $en; TargetAddressDomains = @($doms); DiscoveryEndpoint = 'https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc' } }
$r = Test-IntraOrgConnector @() ''
T 'keiner: FAIL'                      $r.Status 'FAIL'
T 'keiner: nennt DAuth'               ($r.Finding -match 'DAuth') $true
T 'deaktiviert: FAIL'                 (Test-IntraOrgConnector @((Ioc $false 'contoso.mail.onmicrosoft.com')) '').Status 'FAIL'
T 'aktiv mit Tenant: OK'              (Test-IntraOrgConnector @((Ioc $true 'contoso.mail.onmicrosoft.com')) '').Status 'OK'
T 'aktiv, nur eigene Domain: FAIL'    (Test-IntraOrgConnector @((Ioc $true 'contoso.com')) '').Status 'FAIL'
T 'anderer Tenant als Ziel: FAIL'     (Test-IntraOrgConnector @((Ioc $true 'fabrikam.mail.onmicrosoft.com')) 'contoso.mail.onmicrosoft.com').Status 'FAIL'
T 'passender Tenant als Ziel: OK'     (Test-IntraOrgConnector @((Ioc $true 'contoso.mail.onmicrosoft.com')) 'contoso.mail.onmicrosoft.com').Status 'OK'

Write-Output "`n######## Schritt 5: Override der Dedicated App ########"
$r = Test-DedicatedAppOverride @((Ov 'EnableSigningVerification' 'EnableSerializationDataSigning' 'Enabled=true')) -Now $now
T 'nur fremder Override: FAIL'        $r.Status 'FAIL'
T 'FAIL nennt den Befehl'             ($r.Fix -match 'New-SettingOverride .*ExchangeOnpremAsThirdPartyAppId') $true
T 'HCW-Hinweis'                       ($r.Finding -match 'HCW') $true
T 'anderer Name, richtige Sektion: OK' (Test-DedicatedAppOverride @((Ov 'Mein Override' 'ExchangeOnpremAsThirdPartyAppId' 'Enabled=True')) -Now $now).Status 'OK'
T 'richtiger Name, falsche Sektion: FAIL' (Test-DedicatedAppOverride @((Ov 'EnableExchangeHybrid3PAppFeature' 'Irgendwas' 'Enabled=true')) -Now $now).Status 'FAIL'
T 'Enabled=false: FAIL'               (Test-DedicatedAppOverride @((Ov 'X' 'ExchangeOnpremAsThirdPartyAppId' 'Enabled=false')) -Now $now).Status 'FAIL'
T 'nur ein Server: WARN'              (Test-DedicatedAppOverride @((Ov 'X' 'ExchangeOnpremAsThirdPartyAppId' 'Enabled=true' @('EX1'))) -Now $now).Status 'WARN'
$r = Test-DedicatedAppOverride @((Ov 'X' 'ExchangeOnpremAsThirdPartyAppId' 'Enabled=true' @() '2026-11-23 14:40')) -Now $now
T 'vor 20 Minuten: WARN'              $r.Status 'WARN'
T 'nennt 60 Minuten'                  ($r.Fix -match '60 minutes') $true

Write-Output "`n######## Schritt 6: AuthServer ########"
function As($name, $en, $id) { [pscustomobject]@{ Name = $name; Enabled = $en; ApplicationIdentifier = $id } }
$acs = As 'ACS - 1' $true ''
$r = @(Test-AuthServerAppId @($acs, (As 'EvoSts - 1' $true $app)))
T 'EvoSts mit App: OK, ACS ignoriert' (($r | ForEach-Object Status) -join ',') 'OK'
T 'App-ID ermittelt'                  ((Get-DedicatedAppId @($acs, (As 'EvoSts - 1' $true $app))) -join ',') $app
T 'leer: FAIL'                        (@(Test-AuthServerAppId @((As 'EvoSts - 1' $true '')))[0].Status) 'FAIL'
T 'First-Party: FAIL'                 (@(Test-AuthServerAppId @((As 'EvoSts - 1' $true '00000002-0000-0ff1-ce00-000000000000')))[0].Status) 'FAIL'
T 'First-Party ist keine App-ID'      ((Get-DedicatedAppId @((As 'EvoSts - 1' $true '00000002-0000-0ff1-ce00-000000000000'))).Count) 0
T 'deaktiviert: FAIL'                 (@(Test-AuthServerAppId @((As 'EvoSts - 1' $false $app)))[0].Status) 'FAIL'
T 'kein EvoSts: FAIL'                 (@(Test-AuthServerAppId @($acs))[0].Status) 'FAIL'

Write-Output "`n######## Schritt 7: Token ########"
function Oa($type, $fullId) { [pscustomobject]@{ ResultType = $type; Detail = [pscustomobject]@{ FullId = $fullId } } }
T 'ohne Postfach: SKIP'               (Test-OAuthResult -Mailbox '').Status 'SKIP'
T 'Cmdlet-Fehler: FAIL'               (Test-OAuthResult -Mailbox 'anna@contoso.com' -ErrorText 'boom').Status 'FAIL'
T 'Success mit App: OK'               (Test-OAuthResult -Result (Oa 'Success' "x L:$app-AS:y") -ExpectedAppIds @($app) -Mailbox 'a').Status 'OK'
T 'Success, First-Party: FAIL'        (Test-OAuthResult -Result (Oa 'Success' 'x L:00000002-0000-0ff1-ce00-000000000000-AS:y') -ExpectedAppIds @($app) -Mailbox 'a').Status 'FAIL'
# Without an app ID on the auth server only the first-party check catches the old app.
T 'First-Party ohne AuthServer-ID: FAIL' (Test-OAuthResult -Result (Oa 'Success' 'x L:00000002-0000-0ff1-ce00-000000000000-AS:y') -ExpectedAppIds @() -Mailbox 'a').Status 'FAIL'
T 'Success, fremde App: FAIL'        (Test-OAuthResult -Result (Oa 'Success' 'x L:11111111-2222-3333-4444-555555555555-AS:y') -ExpectedAppIds @($app) -Mailbox 'a').Status 'FAIL'
T 'Success ohne ID: WARN'             (Test-OAuthResult -Result (Oa 'Success' 'nothing here') -ExpectedAppIds @($app) -Mailbox 'a').Status 'WARN'
T 'Fehlschlag: FAIL'                  (Test-OAuthResult -Result (Oa 'Error' 'The remote server returned (401)') -Mailbox 'a').Status 'FAIL'
T 'FullId als Text'                   (Get-OAuthAppId ([pscustomobject]@{ ResultType = 'Success'; Detail = "L:$app-AS:" })) $app

Write-Output "`n######## Schritt 7: Auth-Zertifikat ########"
$tp = 'AB12CD34'
function AuthCfg($cur, $next = '') { [pscustomobject]@{ CurrentCertificateThumbprint = $cur; NextCertificateThumbprint = $next } }
function Ce($t, $na) { [pscustomobject]@{ Thumbprint = $t; NotAfter = [datetime]$na } }
T 'kein Thumbprint: FAIL'             (Test-AuthCertificate (AuthCfg '') @() -Now $now).Status 'FAIL'
T 'nicht im Speicher: WARN'           (Test-AuthCertificate (AuthCfg $tp) @((Ce 'FFFF' '2030-01-01')) -Now $now).Status 'WARN'
T 'abgelaufen: FAIL'                  (Test-AuthCertificate (AuthCfg $tp) @((Ce $tp '2026-11-01')) -Now $now).Status 'FAIL'
T 'bald ablaufend: WARN'              (Test-AuthCertificate (AuthCfg $tp) @((Ce $tp '2026-12-03')) -Now $now).Status 'WARN'
T 'gueltig: OK'                       (Test-AuthCertificate (AuthCfg $tp) @((Ce $tp '2028-01-01')) -Now $now).Status 'OK'

Write-Output "`n######## Schritt 8: Weg nach EXO ########"
$g = Ov 'EnableRouteThroughMSGraphFeature' 'RouteThroughMSGraph' 'Enabled=true'
$hu6 = Srv 'EX1' '15.2.2562.17' '15.2.2562.41'
$rtm = Srv 'EX2' '15.2.2562.17'
T 'ohne Graph-Override: INFO'         (Test-GraphRoute @() @($hu6)).Status 'INFO'
T 'INFO nennt EWS'                    ((Test-GraphRoute @() @($hu6)).Finding -match 'EWS') $true
T 'Graph an, alle HU6: OK'            (Test-GraphRoute @($g) @($hu6)).Status 'OK'
T 'Graph an, anderer Server nur Anzeige RTM: WARN' (Test-GraphRoute @($g) @($hu6, $rtm)).Status 'WARN'
T 'Graph an, Datei RTM: FAIL'         (Test-GraphRoute @($g) @($hu6, (Srv 'EX3' '15.2.2562.17' '15.2.2562.17'))).Status 'FAIL'
T 'Graph per Name gefunden'           (Test-GraphRoute @((Ov 'EnableRouteThroughMSGraphFeature' 'Unbekannt' 'Enabled=true')) @($hu6)).Status 'OK'
T 'Graph aus: INFO'                   (Test-GraphRoute @((Ov 'EnableRouteThroughMSGraphFeature' 'X' 'Enabled=false')) @($hu6)).Status 'INFO'

Write-Output "`n######## Schritt 7/8: Endpunkte ########"
$net = @([pscustomobject]@{ Host = 'login.microsoftonline.com'; Ok = $true; Error = '' },
         [pscustomobject]@{ Host = 'graph.microsoft.com'; Ok = $false; Error = 'timeout' })
$r = @(Test-EndpointResult $net '')
T 'login gehoert zu Schritt 7'        $r[0].Step 7
T 'graph gehoert zu Schritt 8'        $r[1].Step 8
T 'erreichbar: OK'                    $r[0].Status 'OK'
T 'nicht erreichbar: FAIL'            $r[1].Status 'FAIL'
T 'mit Proxy nur WARN'                (@(Test-EndpointResult $net 'http://192.0.2.10:8080')[1].Status) 'WARN'

Write-Output "`n######## Schritt 8: EWS-Log ########"
function Hit($line, $n = 1) { [pscustomobject]@{ Line = $line; Filename = 'EWS_20261123-1.LOG'; LineNumber = $n } }
T 'keine Dateien: SKIP'               (Test-EwsLogHit @() 0 '').Status 'SKIP'
T 'keine Treffer: OK'                 (Test-EwsLogHit @() 3 '').Status 'OK'
$h = @((Hit 'GetUserAvailability;(403) Forbidden;EXNL-GraphWebRequestProcessingException:WebException;ben@contoso.com' 7), (Hit 'ErrorProxyRequestProcessingFailed;erika@contoso.com' 9))
$r = Test-EwsLogHit $h 3 ''
T 'Treffer: WARN'                     $r.Status 'WARN'
T 'zaehlt Graph-Treffer'              ($r.Finding -match '2 hit\(s\), 1 Graph-related') $true
T 'Graph-Hinweis'                     ($r.Fix -match 'Graph API permissions') $true
$r = Test-EwsLogHit $h 3 'ben@contoso.com'
T 'Filter auf Adresse'                ($r.Finding -match '^1 hit') $true
T 'fremde Adresse: OK'                (Test-EwsLogHit $h 3 'nobody@contoso.com').Status 'OK'

Write-Output "`n######## Exitcode ########"
T 'ohne FAIL: 0'                      (Get-OverallExitCode @((New-CheckResult 1 'a' 'WARN' 'x'), (New-CheckResult 2 'b' 'OK' 'y'))) 0
T 'mit FAIL: 1'                       (Get-OverallExitCode @((New-CheckResult 1 'a' 'FAIL' 'x'))) 1

Write-Output ("`n================ Bestanden: {0}   Fehlgeschlagen: {1} ================" -f $pass, $fail)
if ($fail) { exit 1 }
