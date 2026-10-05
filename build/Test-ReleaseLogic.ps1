# Spielt die PowerShell-Schritte aus .github/workflows/release.yml lokal nach -
# alles ausser "gh release create". Zweck: die Logik prüfen, bevor ein Tag gesetzt wird.
[CmdletBinding()]
param([string]$Repo = (Split-Path $PSScriptRoot -Parent))
$ErrorActionPreference = 'Stop'
Set-Location $Repo

$pass = 0; $fail = 0
function Check($name, [scriptblock]$sb, [bool]$sollWerfen = $false) {
    $threw = $false; $msg = ''
    try { & $sb | Out-Null } catch { $threw = $true; $msg = $_.Exception.Message }
    if ($threw -eq $sollWerfen) { $script:pass++; Write-Host "  OK   $name" -ForegroundColor Green }
    else { $script:fail++; Write-Host "  FEHL $name (warf=$threw, erwartet=$sollWerfen) $msg" -ForegroundColor Red }
}

# --- Schritt "Tag zerlegen" als Funktion, wie im Workflow ---
function Split-ReleaseTag([string]$ref) {
    if ($ref -notmatch '^(?<tool>[^/]+)/v(?<ver>\d+\.\d+\.\d+)$') {
        throw "Tag '$ref' passt nicht auf <Werkzeug>/v<Version>"
    }
    $tool = $Matches['tool']; $ver = $Matches['ver']
    if (-not (Test-Path $tool -PathType Container)) { throw "Ordner '$tool' gibt es nicht." }
    if (-not (Test-Path "$tool/release.psd1")) { throw "'$tool/release.psd1' fehlt." }
    [pscustomobject]@{ Tool = $tool; Version = $ver }
}

Write-Host "`n######## A. Tag-Zerlegung ########"
Check 'gültiger Tag MailContactEditor/v1.0.0' { Split-ReleaseTag 'MailContactEditor/v1.0.0' }
Check 'gültiger Tag LogViewer/v3.0.1'         { Split-ReleaseTag 'LogViewer/v3.0.1' }
Check 'RUECKBAU: Tag ohne Werkzeug (v1.0.0)'   { Split-ReleaseTag 'v1.0.0' }                     $true
Check 'RUECKBAU: Tag ohne Patch (Tool/v1.0)'   { Split-ReleaseTag 'LogViewer/v1.0' }             $true
Check 'RUECKBAU: unbekannter Ordner'           { Split-ReleaseTag 'GibtsNicht/v1.0.0' }          $true
Check 'RUECKBAU: Ordner ohne release.psd1'     { Split-ReleaseTag 'ExchangeTester/v1.0.0' }      $true

Write-Host "`n######## B. Metadaten lesbar + vollständig ########"
foreach ($t in 'MailContactEditor', 'LogViewer', 'AcmeExchange') {
    $meta = Import-PowerShellDataFile "$t/release.psd1"
    $ok = $meta.Main -and (Test-Path "$t/$($meta.Main)")
    $buildOk = (-not $meta.Build) -or (Test-Path "$t/$($meta.Build)")
    if ($ok -and $buildOk) { $pass++; Write-Host "  OK   $t -> Main=$($meta.Main) Build=$($meta.Build) Bundle=$($meta.Bundle)" -ForegroundColor Green }
    else { $fail++; Write-Host "  FEHL $t : Main oder Build zeigt ins Leere" -ForegroundColor Red }
}

Write-Host "`n######## C. Versionscheck je Werkzeug ########"
. ./build/Get-ToolVersion.ps1
foreach ($t in 'MailContactEditor', 'LogViewer', 'AcmeExchange') {
    $meta = Import-PowerShellDataFile "$t/release.psd1"
    $v = Get-ToolVersion -Path "$t/$($meta.Main)"
    Check "$t : Tag v$v passt" { Assert-ToolVersion -Path "$t/$($meta.Main)" -Expected $v }
    Check "$t : RUECKBAU Tag v0.0.1 wirft" { Assert-ToolVersion -Path "$t/$($meta.Main)" -Expected '0.0.1' } $true
}

Write-Host "`n######## D0. Bauen, was die folgenden Schritte brauchen ########"
# Die EXE-Dateien sind nicht versioniert - auf einem frischen Klon gibt es sie nicht.
# Dieser Test muss sie deshalb selbst erzeugen, so wie der Workflow es tut; sonst
# prüft er nur, was zufällig noch im Arbeitsverzeichnis liegt.
if (-not (Get-Module -ListAvailable ps2exe)) {
    throw 'Das Modul ps2exe fehlt - ohne es lässt sich die Release-Kette nicht prüfen. Install-Module ps2exe -Scope CurrentUser'
}
foreach ($t in 'MailContactEditor', 'LogViewer', 'AcmeExchange') {
    $m = Import-PowerShellDataFile "$t/release.psd1"
    if (-not $m.Build) { continue }
    try {
        & "./$t/$($m.Build)" -Root (Resolve-Path $t).Path *>&1 | Out-Null
        $pass++; Write-Host "  OK   $t gebaut ($($m.Build))" -ForegroundColor Green
    }
    catch {
        $fail++; Write-Host "  FEHL $t : Build scheiterte - $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "`n######## D. Artefakte einsammeln (MailContactEditor, einzelne EXE) ########"
$tool = 'MailContactEditor'
$meta = Import-PowerShellDataFile "$tool/release.psd1"
$ver = Get-ToolVersion -Path "$tool/$($meta.Main)"
if (Test-Path dist) { Remove-Item dist -Recurse -Force }
New-Item -ItemType Directory -Force -Path dist | Out-Null
foreach ($a in @($meta.Artifacts)) {
    $p = Join-Path $tool $a
    if (-not (Test-Path $p)) { throw "Artefakt fehlt nach dem Build: $p" }
    Copy-Item $p dist/
}
$files = Get-ChildItem dist -File
if ($files) { $pass++; Write-Host "  OK   dist enthält: $(($files.Name) -join ', ')" -ForegroundColor Green }
else { $fail++; Write-Host '  FEHL dist ist leer' -ForegroundColor Red }

Write-Host "`n######## E. Bundle-Zweig (AcmeExchange, ZIP) ########"
$tool = 'AcmeExchange'
$meta = Import-PowerShellDataFile "$tool/release.psd1"
$ver = Get-ToolVersion -Path "$tool/$($meta.Main)"
$stage = Join-Path $env:TEMP "bundle-test/$tool"
if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
New-Item -ItemType Directory -Force -Path $stage | Out-Null
Copy-Item "$tool/*" $stage -Recurse -Force
Get-ChildItem "$stage\*" -Include 'release.psd1', 'build-exe.ps1', 'Build-*.ps1', '.gitignore' -Recurse -Force | Remove-Item -Force
$zip = "dist/$tool-$ver.zip"
Compress-Archive -Path "$stage/*" -DestinationPath $zip -Force

# Das ZIP muss das Nötige enthalten und das UnNötige nicht.
Add-Type -AssemblyName System.IO.Compression.FileSystem
# Handle ausdrücklich schließen - sonst bleibt das ZIP gesperrt und der nächste Lauf
# kann dist\ nicht aufräumen.
$archive = [IO.Compression.ZipFile]::OpenRead((Resolve-Path $zip))
try { $names = @($archive.Entries.FullName) } finally { $archive.Dispose() }
foreach ($must in 'Setup.exe', 'Invoke-AcmeExchangeCert.ps1', 'AcmeExchangeSetup.ps1') {
    if ($names -contains $must) { $pass++; Write-Host "  OK   ZIP enthält $must" -ForegroundColor Green }
    else { $fail++; Write-Host "  FEHL ZIP ohne $must" -ForegroundColor Red }
}
foreach ($darfNicht in 'release.psd1', 'build-exe.ps1', '.gitignore') {
    if ($names -notcontains $darfNicht) { $pass++; Write-Host "  OK   $darfNicht nicht im ZIP" -ForegroundColor Green }
    else { $fail++; Write-Host "  FEHL $darfNicht noch im ZIP" -ForegroundColor Red }
}
# Compress-Archive schreibt unter Windows '\' als Trenner - auf beides pruefen.
$libCount = ($names | Where-Object { $_ -match '^lib[\\/]' }).Count
if ($libCount -gt 0) { $pass++; Write-Host "  OK   lib\ ist im ZIP ($libCount Dateien)" -ForegroundColor Green }
else { $fail++; Write-Host '  FEHL lib\ fehlt im ZIP - Setup.exe wäre unbrauchbar' -ForegroundColor Red }

Write-Host "`n================ Bestanden: $pass   Fehlgeschlagen: $fail ================"
if ($fail -gt 0) { exit 1 }
