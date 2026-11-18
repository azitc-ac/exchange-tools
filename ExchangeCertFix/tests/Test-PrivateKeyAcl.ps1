# Prüft Test-PrivateKeyAcl aus Repair-ExchangeSmtpCertificate.ps1 gegen echte Dateien:
# eine mit Leserecht für NETWORK SERVICE, eine ohne, eine mit Deny. Hintergrund: Bis 1.0.0
# warf die Prüfung bei JEDEM Lauf (ungültiger Typausdruck), der Fehler wurde abgefangen und
# als "kein Leserecht" gemeldet - das Werkzeug hätte jedes Zertifikat als defekt eingestuft.
$ErrorActionPreference = 'Stop'
$hier = if ($PSScriptRoot) { $PSScriptRoot }
        elseif ($MyInvocation.MyCommand.Path) { Split-Path $MyInvocation.MyCommand.Path -Parent }
        else { (Get-Location).Path }
$skript = Join-Path (Split-Path $hier -Parent) 'Repair-ExchangeSmtpCertificate.ps1'

$ast = [System.Management.Automation.Language.Parser]::ParseFile($skript, [ref]$null, [ref]$null)
$def = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $args[0].Name -eq 'Test-PrivateKeyAcl' }, $false)
if (-not $def) { throw 'Funktion nicht gefunden: Test-PrivateKeyAcl' }
Invoke-Expression $def[0].Extent.Text

# Attrappe: eine Warnung heißt hier, dass die Prüfung selbst gescheitert ist.
$script:Warnungen = 0
function Write-Warn { param([string]$m) $script:Warnungen++; Write-Output "  (Warnung: $m)" }

$pass = 0; $fail = 0
function T($name, $ist, $soll) {
    if ("$ist" -eq "$soll") { $script:pass++; Write-Output "  OK   $name" }
    else { $script:fail++; Write-Output "  FEHL $name : ist '$ist', soll '$soll'" }
}

$dir = Join-Path $env:TEMP ("acltest-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory $dir | Out-Null
$ns = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-20')   # NETWORK SERVICE, sprachunabhängig
function New-Datei($name, $recht, $typ) {
    $f = Join-Path $dir $name
    Set-Content -Path $f -Value 'x'
    if ($recht) {
        $acl = Get-Acl $f
        $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($ns, $recht, $typ)))
        Set-Acl -Path $f -AclObject $acl
    }
    $f
}

try {
    Write-Output "`n######## Leserecht vorhanden ########"
    $script:Warnungen = 0
    $r = Test-PrivateKeyAcl -KeyFilePath (New-Datei 'lesen.key' 'Read' 'Allow')
    T 'Read erkannt'                 $r.NetworkServiceOk $true
    T 'Prüfung ohne Warnung'         $script:Warnungen 0

    $script:Warnungen = 0
    $r = Test-PrivateKeyAcl -KeyFilePath (New-Datei 'voll.key' 'FullControl' 'Allow')
    T 'FullControl erkannt'          $r.NetworkServiceOk $true

    Write-Output "`n######## Kein Leserecht ########"
    $script:Warnungen = 0
    $r = Test-PrivateKeyAcl -KeyFilePath (New-Datei 'ohne.key' $null $null)
    T 'ohne Eintrag: kein Recht'     $r.NetworkServiceOk $false
    T 'Prüfung ohne Warnung'         $script:Warnungen 0

    $r = Test-PrivateKeyAcl -KeyFilePath (New-Datei 'deny.key' 'Read' 'Deny')
    T 'Deny ist kein Leserecht'      $r.NetworkServiceOk $false

    $r = Test-PrivateKeyAcl -KeyFilePath (New-Datei 'attr.key' 'ReadAttributes' 'Allow')
    T 'nur ReadAttributes reicht nicht' $r.NetworkServiceOk $false
}
finally {
    Remove-Item $dir -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Output "`n================ Bestanden: $pass   Fehlgeschlagen: $fail ================"
if ($fail -gt 0) { exit 1 }
