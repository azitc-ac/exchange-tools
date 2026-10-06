# Spielt die PowerShell-Schritte aus .github/workflows/release.yml lokal nach -
# alles ausser "gh release create". Zweck: die Logik prüfen, bevor ein Tag gesetzt wird.
[CmdletBinding()]
param([string]$Repo)
# $PSScriptRoot kommt beim Start ueber -File in manchen Shells leer an; dann den
# eigenen Pfad anders ermitteln, sonst scheitert schon die Parameterbindung.
if (-not $Repo) {
    $hier = if ($PSScriptRoot) { $PSScriptRoot }
            elseif ($MyInvocation.MyCommand.Path) { Split-Path $MyInvocation.MyCommand.Path -Parent }
            else { (Get-Location).Path }
    $Repo = Split-Path $hier -Parent
}
$ErrorActionPreference = 'Stop'
Set-Location $Repo

$pass = 0; $fail = 0
function Check($name, [scriptblock]$sb, [bool]$sollWerfen = $false) {
    $threw = $false; $msg = ''
    try { & $sb | Out-Null } catch { $threw = $true; $msg = $_.Exception.Message }
    if ($threw -eq $sollWerfen) { $script:pass++; Write-Output "  OK   $name" }
    else { $script:fail++; Write-Output "  FEHL $name (warf=$threw, erwartet=$sollWerfen) $msg" }
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

Write-Output "`n######## A. Tag-Zerlegung ########"
Check 'gültiger Tag MailContactEditor/v1.0.0' { Split-ReleaseTag 'MailContactEditor/v1.0.0' }
Check 'gültiger Tag LogViewer/v3.0.1'         { Split-ReleaseTag 'LogViewer/v3.0.1' }
Check 'RUECKBAU: Tag ohne Werkzeug (v1.0.0)'   { Split-ReleaseTag 'v1.0.0' }                     $true
Check 'RUECKBAU: Tag ohne Patch (Tool/v1.0)'   { Split-ReleaseTag 'LogViewer/v1.0' }             $true
Check 'RUECKBAU: unbekannter Ordner'           { Split-ReleaseTag 'GibtsNicht/v1.0.0' }          $true
# 'build' ist kein Werkzeug und bekommt nie eine release.psd1 - taugt also dauerhaft
# als Gegenprobe, anders als ein Werkzeugordner, der spaeter doch eine bekommt.
Check 'RUECKBAU: Ordner ohne release.psd1'     { Split-ReleaseTag 'build/v1.0.0' }               $true

Write-Output "`n######## B. Jedes Werkzeug ist release-faehig ########"
# Nicht mehr drei fest verdrahtete Namen: geprüft wird jeder Werkzeugordner. Wer
# keine release.psd1 hat, faellt hier auf - sonst bliebe er stillschweigend ohne
# Release, so wie neun von zwoelf es monatelang waren.
$werkzeuge = @(Get-ChildItem . -Directory |
               Where-Object { $_.Name -notmatch '^\.|^build$|^dist$' } |
               Sort-Object Name)
$ohneMeta = @($werkzeuge | Where-Object { -not (Test-Path (Join-Path $_.Name 'release.psd1')) })
if ($ohneMeta) {
    foreach ($o in $ohneMeta) { $fail++; Write-Output "  FEHL $($o.Name): keine release.psd1 - bekaeme nie ein Release" }
}
else { $pass++; Write-Output "  OK   alle $($werkzeuge.Count) Werkzeuge haben eine release.psd1" }

foreach ($w in $werkzeuge) {
    $t = $w.Name
    $meta = Import-PowerShellDataFile "$t/release.psd1"
    if (-not ($meta.Main -and (Test-Path "$t/$($meta.Main)"))) {
        $fail++; Write-Output "  FEHL $t : Main zeigt ins Leere ($($meta.Main))"; continue
    }
    # Entweder eine EXE oder ein Bundle - ohne beides bliebe das Release leer.
    if (-not $meta.Exe -and -not $meta.Bundle) {
        $fail++; Write-Output "  FEHL $t : weder Exe noch Bundle - das Release waere leer"; continue
    }
    $art = if ($meta.Exe) { "Exe=$($meta.Exe)" } else { 'ZIP' }
    $pass++; Write-Output "  OK   $t -> Main=$($meta.Main), $art, Bundle=$($meta.Bundle)"
}

Write-Output "`n######## B2. EXE-tauglicher Pfad-Fallback ########"
# In einer mit ps2exe gebauten EXE sind $PSScriptRoot UND $MyInvocation.MyCommand.Path
# leer. Wer daraus einen Pfad ableitet, bekommt beim Start
# "Cannot bind argument to parameter 'Path' because it is null" - und zwar bevor ein
# Fenster erscheint. Genau so ist ExchangeTester.exe ausgeliefert worden.
# Einziger verlaesslicher Weg dort: der Prozesspfad.
foreach ($w in $werkzeuge) {
    $t = $w.Name
    $meta = Import-PowerShellDataFile "$t/release.psd1"
    if (-not $meta.Exe) { continue }          # nur was als EXE ausgeliefert wird
    $pfad = "$t/$($meta.Main)"
    if (-not (Test-Path $pfad)) { continue }
    $txt = [IO.File]::ReadAllText((Resolve-Path $pfad), [Text.Encoding]::UTF8)

    if ($txt -notmatch '\$PSScriptRoot|\$MyInvocation\.MyCommand\.Path|\$PSCommandPath') {
        $pass++; Write-Output "  OK   $t leitet keinen Pfad aus dem Skriptort ab"
        continue
    }
    if ($txt -match 'MainModule\.FileName') {
        $pass++; Write-Output "  OK   $t hat den Prozesspfad als Rueckfallebene"
    }
    else {
        $fail++
        Write-Output "  FEHL $t ($($meta.Main)): nutzt `$PSScriptRoot ohne Rueckfall auf MainModule.FileName - die EXE bricht beim Start ab"
    }
}

Write-Output "`n######## C. Versionscheck je Werkzeug ########"
. ./build/Get-ToolVersion.ps1
foreach ($w in $werkzeuge) {
    $t = $w.Name
    $meta = Import-PowerShellDataFile "$t/release.psd1"
    if (-not (Test-Path "$t/$($meta.Main)")) { continue }
    $v = Get-ToolVersion -Path "$t/$($meta.Main)"
    Check "$t : Tag v$v passt" { Assert-ToolVersion -Path "$t/$($meta.Main)" -Expected $v }
    Check "$t : RUECKBAU Tag v0.0.1 wirft" { Assert-ToolVersion -Path "$t/$($meta.Main)" -Expected '0.0.1' } $true
}

Write-Output "`n######## D0. Bauen, was die folgenden Schritte brauchen ########"
# Die EXE-Dateien sind nicht versioniert - auf einem frischen Klon gibt es sie nicht.
# Dieser Test muss sie deshalb selbst erzeugen, so wie der Workflow es tut; sonst
# prüft er nur, was zufällig noch im Arbeitsverzeichnis liegt.
if (-not (Get-Module -ListAvailable ps2exe)) {
    throw 'Das Modul ps2exe fehlt - ohne es lässt sich die Release-Kette nicht prüfen. Install-Module ps2exe -Scope CurrentUser'
}
foreach ($w in $werkzeuge) {
    $t = $w.Name
    $m = Import-PowerShellDataFile "$t/release.psd1"
    if (-not $m.Exe) { continue }
    try {
        & ./build/Build-ToolExe.ps1 -Tool $t *>&1 | Out-Null
        $exe = Join-Path $t $m.Exe
        if (Test-Path $exe) {
            $fv = (Get-Item $exe).VersionInfo.FileVersion
            $soll = Get-ToolVersion -Path "$t/$($m.Main)"
            if ($fv -eq $soll) { $pass++; Write-Output ("  OK   {0,-22} {1} v{2}" -f $t, $m.Exe, $fv) }
            else { $fail++; Write-Output "  FEHL $t : EXE traegt v$fv, Skript sagt v$soll" }
        }
        else { $fail++; Write-Output "  FEHL $t : $($m.Exe) wurde nicht erzeugt" }
    }
    catch {
        $fail++; Write-Output "  FEHL $t : Build scheiterte - $($_.Exception.Message)"
    }
}

Write-Output "`n######## D. Artefakte einsammeln (MailContactEditor, einzelne EXE) ########"
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
if ($files) { $pass++; Write-Output "  OK   dist enthält: $(($files.Name) -join ', ')" }
else { $fail++; Write-Output '  FEHL dist ist leer' }

# Die Übergabe der Assets an gh - hier ist der erste Release-Lauf gescheitert.
# Bei genau EINER Datei liefert .FullName einen String, und @string splattet
# zeichenweise: gh bekam 'D' statt 'D:\...\MailContactEditor.exe' und meldete
# "no matches found for `D`". @(...) erzwingt das Array.
$assets = @((Get-ChildItem dist -File).FullName)
$sammler = { param([Parameter(ValueFromRemainingArguments)]$rest) $rest }
$uebergeben = @(& $sammler @assets)
if ($uebergeben.Count -eq $assets.Count -and $uebergeben[0] -eq $assets[0]) {
    $pass++; Write-Output "  OK   Assets kommen als $($uebergeben.Count) vollständige(r) Pfad(e) an"
}
else {
    $fail++
    Write-Output ("  FEHL Assets zerfallen beim Splatting: {0} Argumente statt {1}, erstes = '{2}'" -f `
                  $uebergeben.Count, $assets.Count, $uebergeben[0])
}

Write-Output "`n######## E. Bundle-Zweig (AcmeExchange, ZIP) ########"
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
    if ($names -contains $must) { $pass++; Write-Output "  OK   ZIP enthält $must" }
    else { $fail++; Write-Output "  FEHL ZIP ohne $must" }
}
foreach ($darfNicht in 'release.psd1', 'build-exe.ps1', '.gitignore') {
    if ($names -notcontains $darfNicht) { $pass++; Write-Output "  OK   $darfNicht nicht im ZIP" }
    else { $fail++; Write-Output "  FEHL $darfNicht noch im ZIP" }
}
# Compress-Archive schreibt unter Windows '\' als Trenner - auf beides pruefen.
$libCount = ($names | Where-Object { $_ -match '^lib[\\/]' }).Count
if ($libCount -gt 0) { $pass++; Write-Output "  OK   lib\ ist im ZIP ($libCount Dateien)" }
else { $fail++; Write-Output '  FEHL lib\ fehlt im ZIP - Setup.exe wäre unbrauchbar' }

Write-Output "`n================ Bestanden: $pass   Fehlgeschlagen: $fail ================"
if ($fail -gt 0) { exit 1 }
