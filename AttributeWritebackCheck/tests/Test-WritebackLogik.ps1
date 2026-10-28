# Prüft die Vergleichs- und Urteilslogik von Test-ExchangeAttributeWriteback.ps1 ohne Exchange
# und ohne AD: die Funktionen werden per AST herausgelöst (wie in MailboxDatabaseCheck/tests),
# Postfächer und AD-Benutzer sind einfache Objekte.
$ErrorActionPreference = 'Stop'
$hier = if ($PSScriptRoot) { $PSScriptRoot }
        elseif ($MyInvocation.MyCommand.Path) { Split-Path $MyInvocation.MyCommand.Path -Parent }
        else { (Get-Location).Path }
$skript = Join-Path (Split-Path $hier -Parent) 'Test-ExchangeAttributeWriteback.ps1'

$ast = [System.Management.Automation.Language.Parser]::ParseFile($skript, [ref]$null, [ref]$null)
foreach ($fn in 'Get-ExpectedAdType', 'Get-AdTypeName', 'Get-FindingCatalog', 'New-AdIndex',
                'Find-AdUser', 'Get-MailboxFinding', 'Get-OverallVerdict') {
    $def = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $args[0].Name -eq $fn }, $false)
    if (-not $def) { throw "Funktion nicht gefunden: $fn" }
    Invoke-Expression $def[0].Extent.Text
}

$pass = 0; $fail = 0
function T($name, $ist, $soll) {
    if ("$ist" -eq "$soll") { $script:pass++; Write-Output "  OK   $name" }
    else { $script:fail++; Write-Output "  FEHL $name : ist '$ist', soll '$soll'" }
}

# Postfach in Exchange Online und Benutzer im AD
function Mbx($upn, $typ, $cloud, $guid, $ext, $primary = $upn) {
    [pscustomobject]@{ UserPrincipalName = $upn; ExchangeGuid = $guid; ExternalDirectoryObjectId = $ext
                       RecipientTypeDetails = $typ; IsExchangeCloudManaged = $cloud; PrimarySmtpAddress = $primary }
}
function Usr($upn, $typ, $guid, $ext, $mail = $upn, $proxy = "SMTP:$upn") {
    [pscustomobject]@{ userPrincipalName = $upn; msExchRecipientTypeDetails = $typ; mail = $mail
                       proxyAddresses = @($proxy, 'smtp:alias@contoso.com')
                       msExchMailboxGuid = $(if ($guid) { ([guid]$guid).ToByteArray() } else { $null })
                       'msDS-ExternalDirectoryObjectId' = $(if ($ext) { "User_$ext" } else { $null }) }
}
$g1 = '11111111-1111-1111-1111-111111111111'; $g2 = '22222222-2222-2222-2222-222222222222'
$e2 = 'bbbbbbbb-0000-0000-0000-000000000002'; $e3 = 'cccccccc-0000-0000-0000-000000000003'
$remoteUser = [int64]2147483648; $remoteShared = [int64]34359738368
$cat = Get-FindingCatalog

function Finding($m, $users, $c = $cat) {
    Get-MailboxFinding -Mailbox $m -Index (New-AdIndex -AdUser $users) -Catalog $c
}

Write-Output "`n######## Zuordnung: GUID, Entra-Objekt-ID, UPN ########"
$users = @(
    (Usr 'a@contoso.com' $remoteUser $g1 $null)
    (Usr 'b@contoso.com' $remoteUser $null $e2)                  # ohne GUID (New-RemoteMailbox)
    (Usr 'C@Contoso.com' $remoteShared $null $null)              # nur UPN, andere Schreibweise
)
T 'GUID'                  (Finding (Mbx 'a@contoso.com' 'UserMailbox' 'True' $g1 $null) $users).MatchedBy 'ExchangeGuid'
T 'Entra-ID ohne GUID'    (Finding (Mbx 'b@contoso.com' 'UserMailbox' 'False' $g2 $e2.ToUpper()) $users).MatchedBy 'EntraObjectId'
T 'UPN ohne Gross/klein'  (Finding (Mbx 'c@contoso.com' 'SharedMailbox' 'False' $null $e3) $users).MatchedBy 'UPN'
T 'nicht gefunden'        (Finding (Mbx 'z@contoso.com' 'UserMailbox' 'False' $null $null) $users).Finding 'NOT-IN-AD'

Write-Output "`n######## Bewertung je Postfach ########"
T 'OK cloud'              (Finding (Mbx 'a@contoso.com' 'UserMailbox' 'True' $g1 $null) $users).Finding 'OK-Cloud'
T 'OK on-prem'            (Finding (Mbx 'b@contoso.com' 'UserMailbox' 'False' $null $e2) $users).Finding 'OK-OnPrem'
$u = @(Usr 'd@contoso.com' 1 $g1 $null)
T 'Warnfall: cloud, AD-Typ on-prem'   (Finding (Mbx 'd@contoso.com' 'UserMailbox' 'True' $g1 $null) $u).Finding 'TYPE-WRONG-Cloud'
T 'Warnfall ist FAIL'                 (Finding (Mbx 'd@contoso.com' 'UserMailbox' 'True' $g1 $null) $u).Severity 'FAIL'
# Gegenprobe zur Reihenfolge: Weicht zusätzlich die Adresse ab, darf der Warnfall nicht hinter
# ADDRESS-DIFF verschwinden (sonst würde aus ROT ein GELB).
$u = @(Usr 'd@contoso.com' 1 $g1 $null 'alt@contoso.com' 'SMTP:alt@contoso.com')
T 'Warnfall vor Adressabweichung'     (Finding (Mbx 'd@contoso.com' 'UserMailbox' 'True' $g1 $null) $u).Finding 'TYPE-WRONG-Cloud'
$u = @(Usr 's@contoso.com' $remoteUser $g1 $null)
T 'Shared in Cloud, on-prem nur Hinweis' (Finding (Mbx 's@contoso.com' 'SharedMailbox' 'False' $g1 $null) $u).Finding 'TYPE-MISMATCH-OnPrem'
T 'Hinweis ist INFO'                  (Finding (Mbx 's@contoso.com' 'SharedMailbox' 'False' $g1 $null) $u).Severity 'INFO'
$u = @(Usr 'n@contoso.com' $remoteUser $g1 $null 'n@contoso.com' 'SMTP:alt@contoso.com')
T 'primaere proxyAddress weicht ab'   (Finding (Mbx 'n@contoso.com' 'UserMailbox' 'True' $g1 $null) $u).Finding 'ADDRESS-DIFF'
$u = @(Usr 'm@contoso.com' $remoteUser $g1 $null 'alt@contoso.com')
T 'nur mail weicht ab'                (Finding (Mbx 'm@contoso.com' 'UserMailbox' 'True' $g1 $null) $u).Finding 'MAIL-DIFF'
$u = @(Usr 'x@contoso.com' $remoteUser $g1 $null)
T 'Typ ohne Sollwert'                 (Finding (Mbx 'x@contoso.com' 'SchedulingMailbox' 'True' $g1 $null) $u).Finding 'TYPE-NO-TARGET'
T 'Klartext Typ'                      (Get-AdTypeName 128) '128 (MailUser)'

Write-Output "`n######## Vorschau vor dem Writeback ########"
$vor = Get-FindingCatalog -BeforeWriteback
$u = @(Usr 'd@contoso.com' 1 $g1 $null)
T 'Warnfall vorab nur WARN'           (Finding (Mbx 'd@contoso.com' 'UserMailbox' 'True' $g1 $null) $u $vor).Severity 'WARN'
T 'Katalog nachher unverändert'       $cat['TYPE-WRONG-Cloud'].Severity 'FAIL'

Write-Output "`n######## Gesamturteil ########"
function F($code, $sev) { [pscustomobject]@{ Finding = $code; Severity = $sev } }
$ok = @((F 'OK-Cloud' 'OK'), (F 'OK-OnPrem' 'OK'))
T 'grün'                   (Get-OverallVerdict -Finding $ok -SearchedServers 1).Verdict 'GREEN'
T 'grün Exit 0'            (Get-OverallVerdict -Finding $ok -SearchedServers 1).ExitCode 0
T 'Hinweis bleibt grün'    (Get-OverallVerdict -Finding ($ok + (F 'TYPE-MISMATCH-OnPrem' 'INFO')) -SearchedServers 1).Verdict 'GREEN'
T 'Adresse gelb'           (Get-OverallVerdict -Finding ($ok + (F 'ADDRESS-DIFF' 'WARN')) -SearchedServers 1).ExitCode 1
T 'Warnfall rot'           (Get-OverallVerdict -Finding ($ok + (F 'ADDRESS-DIFF' 'WARN') + (F 'TYPE-WRONG-Cloud' 'FAIL')) -SearchedServers 1).ExitCode 2
# Entstanden aus dem ersten Einsatz: 1 von 1261 bewertet ergab GRÜN.
T 'nicht zugeordnet: unvollständig' (Get-OverallVerdict -Finding ($ok + (F 'NOT-IN-AD' 'NOMATCH')) -SearchedServers 1).Verdict 'INCOMPLETE'
T 'unvollständig Exit 3'   (Get-OverallVerdict -Finding ($ok + (F 'NOT-IN-AD' 'NOMATCH')) -SearchedServers 1).ExitCode 3
T 'kein AD: nicht bewertbar' (Get-OverallVerdict -Finding $ok -SearchedServers 0).Verdict 'NOT EVALUABLE'
T 'nichts zugeordnet: nicht bewertbar' (Get-OverallVerdict -Finding @((F 'NOT-IN-AD' 'NOMATCH')) -SearchedServers 1).ExitCode 3

Write-Output ("`n{0} bestanden, {1} fehlgeschlagen" -f $pass, $fail)
if ($fail) { exit 1 } else { exit 0 }
