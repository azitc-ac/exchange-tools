<#
.SYNOPSIS
    Prüft, dass -MsgSize auch beim Start über -File richtig umgerechnet wird.

.DESCRIPTION
    Entstanden aus einem Fehler beim ersten Testlauf: der Parameter war als [int64]
    deklariert, und "5MB" kam über

        powershell.exe -File .\Populate-TestMailbox.ps1 -MsgSize 5MB

    als Text an - PowerShell wertet Argumente hinter -File nicht aus. Die Bindung
    scheiterte, bevor überhaupt etwas passierte. In der Shell aufgerufen fiel das nicht
    auf, weil dort 5MB zu 5242880 ausgewertet wird. Beide Wege müssen funktionieren.

    Geprüft wird am echten Skript, nicht an einer nachgebauten Kopie: der Lauf bricht
    beim Verbindungsversuch ab, aber die Verbose-Zeile mit der errechneten Bytezahl
    kommt vorher - und genau die ist das Ergebnis der Umrechnung.

.PARAMETER Rueckbau
    Baut den Fehler zurück (Parameter wieder als [int64]) und erwartet, dass die
    Prüfung FEHLSCHLÄGT. Ohne diese Gegenprobe wäre der Test wertlos.

.EXAMPLE
    .\tests\Test-MsgSize.ps1
    .\tests\Test-MsgSize.ps1 -Rueckbau
#>
[CmdletBinding()]
param([switch]$Rueckbau)

$ErrorActionPreference = 'Stop'
$hier = if ($PSScriptRoot) { $PSScriptRoot }
        elseif ($MyInvocation.MyCommand.Path) { Split-Path $MyInvocation.MyCommand.Path -Parent }
        else { (Get-Location).Path }
$skript = Join-Path (Split-Path $hier -Parent) 'Populate-TestMailbox.ps1'
if (-not (Test-Path $skript)) { throw "Hauptskript nicht gefunden: $skript" }

# Beim Rückbau auf einer Kopie arbeiten - das Werkzeug selbst bleibt unangetastet.
if ($Rueckbau) {
    $kopie = Join-Path $env:TEMP ('Populate-TestMailbox.rueckbau.{0}.ps1' -f $PID)
    $txt = [IO.File]::ReadAllText($skript, [Text.Encoding]::UTF8)
    $alt = "[string]`$MsgSize = '1000KB',"
    if ($txt -notmatch [regex]::Escape($alt)) { throw "Rückbau nicht möglich: '$alt' nicht gefunden." }
    $txt = $txt.Replace($alt, '[int64]$MsgSize = 1024000,')
    [IO.File]::WriteAllText($kopie, $txt, (New-Object Text.UTF8Encoding $true))
    $skript = $kopie
    Write-Output '  (Rückbau: MsgSize wieder als [int64] - die Prüfung MUSS fehlschlagen)'
}

$pass = 0; $fail = 0
function Ok($t)  { $script:pass++; Write-Output "  OK   $t" }
function Bad($t) { $script:fail++; Write-Output "  FEHL $t" }

# Das Skript scheitert danach am nicht auflösbaren Endpunkt - gewollt: die Umrechnung
# ist zu diesem Zeitpunkt längst gelaufen und als Verbose-Zeile heraus.
function Get-Ausgabe([string]$Groesse) {
    $alt = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
    try {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $script:skript `
            -TargetMailbox 'test@invalid.invalid' `
            -EwsUrl 'https://invalid.invalid/EWS/Exchange.asmx' `
            -MsgSize $Groesse -Verbose 2>&1 | Out-String
    }
    finally { $ErrorActionPreference = $alt }
}

Write-Output "`n######## Einheiten werden umgerechnet ########"
# 0,001GB statt 1GB: prüft den GB-Faktor, ohne ein Gigabyte Zufallsdaten zu erzeugen.
# Erwartet werden 1073742, nicht 1073741: der Cast nach [int64] rundet, er schneidet nicht ab.
$faelle = @(
    @{ Ein = '5MB';     Soll = 5242880 }
    @{ Ein = '100KB';   Soll = 102400 }
    @{ Ein = '0,001GB'; Soll = 1073742 }
    @{ Ein = '5242880'; Soll = 5242880 }
    @{ Ein = '1,5MB';   Soll = 1572864 }
)
foreach ($f in $faelle) {
    $aus = Get-Ausgabe $f.Ein
    $m = [regex]::Match($aus, 'Erzeuge\s+(\d+)\s+Bytes')
    if (-not $m.Success) {
        Bad "$($f.Ein): keine Bytezahl in der Ausgabe (Parameterbindung gescheitert?)"
        continue
    }
    $ist = [int64]$m.Groups[1].Value
    if ($ist -eq $f.Soll) { Ok "$($f.Ein) -> $ist Bytes" }
    else { Bad "$($f.Ein) -> $ist Bytes, erwartet $($f.Soll)" }
}

Write-Output "`n######## Unsinn wird abgewiesen ########"
foreach ($u in @('abc', '0', '-5MB', '5TB')) {
    $aus = Get-Ausgabe $u
    # Entweder der eigene Fehlertext oder - bei negativen Zahlen - die Parameterprüfung.
    if ($aus -match 'nicht verstanden|ergibt\s+-?\d+\s+Bytes|Cannot convert|kann nicht') {
        Ok "'$u' abgewiesen"
    }
    elseif ($aus -match 'Erzeuge\s+(\d+)\s+Bytes') {
        Bad "'$u' wurde als $($Matches[1]) Bytes akzeptiert"
    }
    else { Bad "'$u': unklare Reaktion" }
}

if ($Rueckbau) { Remove-Item $skript -Force -ErrorAction SilentlyContinue }

Write-Output "`n================ Bestanden: $pass   Fehlgeschlagen: $fail ================"
if ($Rueckbau) {
    if ($fail -gt 0) { Write-Output '  Gegenprobe erfolgreich: mit dem zurueckgebauten Fehler schlaegt die Pruefung fehl.'; exit 0 }
    Write-Output '  Gegenprobe GESCHEITERT: die Pruefung ist auch mit dem Fehler gruen - sie prueft nichts.'
    exit 1
}
if ($fail -gt 0) { exit 1 }
