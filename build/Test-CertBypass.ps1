<#
.SYNOPSIS
    Prüft, dass kein Skript die Zertifikatsprüfung prozessweit abschaltet.

.DESCRIPTION
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    (oder ::CertificatePolicy = TrustAll) gilt für jede Anfrage im ganzen Prozess,
    auch für Anmeldungen bei Microsoft. Genau diese Zeile steht außerdem in fast
    jedem bekannten Angriffsskript - EDR-Produkte wie CrowdStrike Falcon haben den
    ExchangeTester mit "This script contains malicious content" geblockt (16.11.2026).

    Richtig ist die Ausnahme pro Anfrage:
        if ($ignoreCert) { $req.ServerCertificateValidationCallback = $acceptAnyCert }

    $Bekannt führt Altlasten mit Begründung. Wie KnownGaps im Prolog-Test schlägt
    die Prüfung auch fehl, wenn ein Eintrag erledigt ist und stehenbleibt.

    Ausgabe über Write-Output statt Write-Host, damit sie in Protokollen landet.
#>
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

$Bekannt = @{
    'TestMailboxFiller\Populate-TestMailbox.ps1' = 'EWS-Zweig mit -IgnoreCert, noch nicht auf pro Anfrage umgestellt'
    'PublicFolderMigration\PFmig-Functions.ps1'  = 'TrustAll-ICertificatePolicy per Add-Type, noch nicht umgestellt'
}

$muster = 'ServicePointManager\]::(ServerCertificateValidationCallback|CertificatePolicy)\s*(\+|-)?='

$fail = 0
$treffer = @{}
$dateien = Get-ChildItem $Repo -Filter *.ps1 -Recurse -File |
           Where-Object { $_.FullName -notmatch '\\lib\\Posh-ACME\\|\\\.git\\' }

foreach ($f in $dateien) {
    $rel = $f.FullName.Substring($Repo.Length).TrimStart('\')
    if ($rel -eq 'build\Test-CertBypass.ps1') { continue }
    $zeilen = @(Select-String -LiteralPath $f.FullName -Pattern $muster |
                Where-Object { $_.Line -notmatch '^\s*#' })
    if (-not $zeilen.Count) { continue }
    $treffer[$rel] = $true
    if ($Bekannt.ContainsKey($rel)) {
        Write-Output ("  bekannt {0} ({1})" -f $rel, $Bekannt[$rel])
        continue
    }
    $fail++
    Write-Output ("  FEHL {0}" -f $rel)
    $zeilen | ForEach-Object { Write-Output ("         Zeile {0}: {1}" -f $_.LineNumber, $_.Line.Trim()) }
}

foreach ($k in $Bekannt.Keys) {
    if (-not $treffer.ContainsKey($k)) {
        $fail++
        Write-Output ("  FEHL {0} steht in `$Bekannt, ist aber erledigt - Eintrag entfernen" -f $k)
    }
}

Write-Output ("  {0} Skripte geprüft, {1} Funde" -f $dateien.Count, $fail)
if ($fail -gt 0) { exit 1 }
