<#
.SYNOPSIS
    Prueft, was ein Release ausliefert: eigenes ZIP je Werkzeug, Inhalt des ZIP,
    Release-Notizen. Ohne ps2exe lauffaehig (laeuft auch mit -SkipBuild).

.DESCRIPTION
    Anlass: Beim ExchangeTester-Release gab es nur die EXE. Wer das Skript wollte,
    landete beim "Source code"-Archiv von GitHub - dem ganzen Sammelrepo - und die
    Notizen schickten ihn auch noch dorthin ("aus dem Quellordner").
    Geprueft werden die Skripte, die der Workflow selbst aufruft:
    build/New-ReleaseBundle.ps1 und build/Get-ReleaseNotes.ps1.
#>
[CmdletBinding()]
param([string]$Repo)
if (-not $Repo) {
    $hier = if ($PSScriptRoot) { $PSScriptRoot }
            elseif ($MyInvocation.MyCommand.Path) { Split-Path $MyInvocation.MyCommand.Path -Parent }
            else { (Get-Location).Path }
    $Repo = Split-Path $hier -Parent
}
$ErrorActionPreference = 'Stop'
Set-Location $Repo
Add-Type -AssemblyName System.IO.Compression.FileSystem

$pass = 0; $fail = 0
function Ok($cond, [string]$text) {
    if ($cond) { $script:pass++; Write-Output "  OK   $text" }
    else { $script:fail++; Write-Output "  FEHL $text" }
}
function Get-ZipNames([string]$zip) {
    $a = [IO.Compression.ZipFile]::OpenRead($zip)
    try { @($a.Entries.FullName | ForEach-Object { $_ -replace '\\', '/' }) } finally { $a.Dispose() }
}

$werkzeuge = @(Get-ChildItem . -Directory |
               Where-Object { $_.Name -notmatch '^\.|^build$|^dist$' -and (Test-Path (Join-Path $_.FullName 'release.psd1')) } |
               Sort-Object Name)

Write-Output "`n######## A. Jedes Werkzeug bekommt ein eigenes ZIP ########"
foreach ($w in $werkzeuge) {
    $meta = Import-PowerShellDataFile (Join-Path $w.FullName 'release.psd1')
    Ok ($meta.Bundle -eq $true) "$($w.Name): Bundle = `$true (sonst bleibt fuer das Skript nur das Komplett-Archiv von GitHub)"
}

# Arbeitskopie, damit eine gebaute EXE darin liegt wie im Runner - auch auf einem
# frischen Klon, auf dem es keine EXE gibt.
$tmp = Join-Path $env:TEMP ("relbundle-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
$tmpRepo = Join-Path $tmp 'repo'; $out = Join-Path $tmp 'dist'
New-Item -ItemType Directory -Force -Path $tmpRepo, $out | Out-Null
try {
    Write-Output "`n######## B. ExchangeTester: Skript im ZIP, EXE nur als eigenes Asset ########"
    Copy-Item (Join-Path $Repo 'ExchangeTester') $tmpRepo -Recurse -Force
    Set-Content (Join-Path $tmpRepo 'ExchangeTester\ExchangeTester.exe') 'dummy'
    $zip = & (Join-Path $Repo 'build\New-ReleaseBundle.ps1') -Tool ExchangeTester -Version 9.9.9 -OutDir $out -Repo $tmpRepo -StageRoot $tmp
    Ok ((Split-Path $zip -Leaf) -eq 'ExchangeTester-9.9.9.zip') "Name ExchangeTester-9.9.9.zip"
    $names = Get-ZipNames $zip
    foreach ($must in 'ExchangeTester.ps1', 'ExchangeTester.bat', 'README.md') { Ok ($names -contains $must) "ZIP enthaelt $must" }
    foreach ($not in 'ExchangeTester.exe', 'release.psd1', '.gitignore') { Ok ($names -notcontains $not) "$not nicht im ZIP" }

    Write-Output "`n######## C. AcmeExchange: Setup.exe gehoert INS ZIP (nicht in Artifacts) ########"
    Copy-Item (Join-Path $Repo 'AcmeExchange') $tmpRepo -Recurse -Force
    Set-Content (Join-Path $tmpRepo 'AcmeExchange\Setup.exe') 'dummy'
    $zip = & (Join-Path $Repo 'build\New-ReleaseBundle.ps1') -Tool AcmeExchange -Version 9.9.9 -OutDir $out -Repo $tmpRepo -StageRoot $tmp
    $names = Get-ZipNames $zip
    foreach ($must in 'Setup.exe', 'Invoke-AcmeExchangeCert.ps1', 'AcmeExchangeSetup.ps1') { Ok ($names -contains $must) "ZIP enthaelt $must" }
    Ok (@($names | Where-Object { $_ -like 'lib/*' }).Count -gt 0) "lib/ ist im ZIP"
    Ok ($names -notcontains 'release.psd1') "release.psd1 nicht im ZIP"
}
finally { Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue }

Write-Output "`n######## D. Release-Notizen ########"
$n = & (Join-Path $Repo 'build\Get-ReleaseNotes.ps1') -Tool ExchangeTester -Version 9.9.9 -Repo $Repo
Ok ($n -match [regex]::Escape('`ExchangeTester-9.9.9.zip`')) "ExchangeTester: nennt das eigene ZIP"
Ok ($n -match 'Sammelrepo mit allen Werkzeugen') "ExchangeTester: erklaert die Source-code-Archive"
Ok ($n -match 'ExchangeTester\.exe.*nicht signiert') "ExchangeTester: EXE-Hinweis mit Dateiname"
Ok ($n -notmatch '\{ZIP\}|\{EXE\}|Quellordner') "ExchangeTester: keine Platzhalter, kein Verweis auf den Quellordner"
$n = & (Join-Path $Repo 'build\Get-ReleaseNotes.ps1') -Tool DmarcReportAnalyzer -Version 9.9.9 -Repo $Repo
Ok ($n -match 'DmarcReportAnalyzer-9\.9\.9\.zip') "DmarcReportAnalyzer: nennt das eigene ZIP"
Ok ($n -notmatch 'nicht signiert') "DmarcReportAnalyzer: kein EXE-Hinweis ohne EXE"

Write-Output "`n================ Bestanden: $pass   Fehlgeschlagen: $fail ================"
if ($fail -gt 0) { exit 1 }
