# Prüft die Erkennung "Version ohne Release" (build\Get-ReleaseDrift.ps1) und die beiden
# Workflow-Stellen, an denen sie hängt. Hintergrund: Am 18.11.2026 hatten sieben Werkzeuge
# eine Version ohne Release - die EXE auf GitHub war älter als das Skript.
[CmdletBinding()]
param([string]$Repo)
if (-not $Repo) {
    $hier = if ($PSScriptRoot) { $PSScriptRoot }
            elseif ($MyInvocation.MyCommand.Path) { Split-Path $MyInvocation.MyCommand.Path -Parent }
            else { (Get-Location).Path }
    $Repo = Split-Path $hier -Parent
}
$ErrorActionPreference = 'Stop'

$pass = 0; $fail = 0
function T($name, $ist, $soll) {
    if ("$ist" -eq "$soll") { $script:pass++; Write-Output "  OK   $name" }
    else { $script:fail++; Write-Output "  FEHL $name : ist '$ist', soll '$soll'" }
}

# -Repo mitgeben: beim Dot-Sourcen setzt der param()-Block des Skripts seine Parameter im
# Scope des Aufrufers - ohne Wert wäre danach unser $Repo leer.
. (Join-Path $Repo 'build\Get-ReleaseDrift.ps1') -Repo $Repo

# --- Attrappen-Repo: zwei Werkzeuge mit release.psd1, eins ohne -------------------------
$tmp = Join-Path $env:TEMP ('drift-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path (Join-Path $tmp 'build') -Force | Out-Null
Copy-Item (Join-Path $Repo 'build\Get-ToolVersion.ps1') (Join-Path $tmp 'build')
function New-Tool($name, $ver, [switch]$OhneRelease) {
    $d = Join-Path $tmp $name; New-Item -ItemType Directory $d | Out-Null
    Set-Content (Join-Path $d "$name.ps1") "`$script:Version = '$ver'"
    if (-not $OhneRelease) { Set-Content (Join-Path $d 'release.psd1') "@{ Main = '$name.ps1'; Bundle = `$true }" }
}
try {
    New-Tool 'Alpha' '1.2.0'
    New-Tool 'Beta'  '2.0.1'
    New-Tool 'Gamma' '0.9.0' -OhneRelease

    Write-Output "`n######## Erkennung ########"
    $d = @(Get-ReleaseDrift -Repo $tmp -Tags @('Alpha/v1.2.0', 'Beta/v2.0.0', 'Gamma/v0.9.0'))
    T 'nur Werkzeuge mit release.psd1'        $d.Count 2
    T 'Alpha 1.2.0 hat Release'               ($d | Where-Object Tool -eq 'Alpha').Released $true
    T 'Beta 2.0.1 ohne Release (nur 2.0.0)'   ($d | Where-Object Tool -eq 'Beta').Released $false
    T 'erwarteter Tag'                        ($d | Where-Object Tool -eq 'Beta').Tag 'Beta/v2.0.1'
    T 'ohne Tags: alles fehlt'                @(Get-ReleaseDrift -Repo $tmp | Where-Object { -not $_.Released }).Count 2
    # Präfix darf nicht reichen: "Alpha/v1.2.0-test" ist nicht "Alpha/v1.2.0".
    T 'ähnlicher Tag zählt nicht'             @(Get-ReleaseDrift -Repo $tmp -Tags @('Alpha/v1.2.0-test') | Where-Object { $_.Tool -eq 'Alpha' -and $_.Released }).Count 0

    Set-Content (Join-Path $tmp 'Beta\Beta.ps1') '# keine Version'
    $warf = $false; try { $null = @(Get-ReleaseDrift -Repo $tmp) } catch { $warf = $true }
    T 'Werkzeug ohne $script:Version wirft'    $warf $true
}
finally { Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue }

# --- Workflows ---------------------------------------------------------------------------
Write-Output "`n######## Workflows ########"
$rel  = Get-Content (Join-Path $Repo '.github\workflows\release.yml') -Raw
$auto = Get-Content (Join-Path $Repo '.github\workflows\auto-release.yml') -Raw
T 'release.yml baut beim Nachholen den Tag-Stand' ($rel -match 'ref:\s*\$\{\{\s*github\.event_name == ''workflow_dispatch'' && github\.event\.inputs\.tag') $true
T 'auto-release nutzt Get-ReleaseDrift'           ($auto -match 'Get-ReleaseDrift\.ps1 -MissingOnly') $true
T 'auto-release startet mit --ref auf den Tag'    ($auto -match 'gh workflow run release\.yml --ref \$m\.Tag') $true
T 'auto-release nur nach grüner CI'               ($auto -match "workflow_run\.conclusion == 'success'") $true
# Run-Blöcke laufen unter PS 5.1 ohne BOM: Umlaute würden zu Mojibake (siehe release.yml).
T 'auto-release.yml ist reines ASCII'             ([bool]($auto -cmatch '[^\x00-\x7F]')) $false

Write-Output "`n================ Bestanden: $pass   Fehlgeschlagen: $fail ================"
if ($fail -gt 0) { exit 1 }
