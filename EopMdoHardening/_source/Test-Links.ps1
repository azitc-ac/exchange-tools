<#
.SYNOPSIS
    Prueft die Fundort-Links des EOP/MDO-Hardening-Standards.

.DESCRIPTION
    Zwei Stufen:

    OFFLINE (immer, schnell, fuer CI geeignet)
      1. Jede ID aus der Kuerzel-Legende der README hat genau eine Zeile in den Fundort-Tabellen,
         und die Punktzahlen der Legende stimmen mit den Zeilen ueberein.
      2. Jede Fundort-Zeile hat einen Learn-Link.
      3. Die Menge der Portal-/Learn-URLs in der README ist identisch mit der im Audit-Skript
         (Invoke-EopAudit.ps1, $script:Links). Driftet eine Seite, schlaegt der Test fehl.

    ONLINE (-Online, braucht Internet, nicht fuer CI)
      4. Jede Learn-URL antwortet mit HTTP 200.
      5. Der Sprunganker (#...) existiert als id-Attribut in der Seite.
      Portal-Links (security.microsoft.com usw.) verlangen eine Anmeldung und werden NICHT geprueft.

    Exitcode 0 = alles gut, 1 = mindestens ein Fehler.

.PARAMETER Online
    Zusaetzlich die Learn-Seiten und -Anker abrufen.

.PARAMETER ThrottleMs
    Pause zwischen Abrufen im Online-Modus (Standard 250 ms).

.EXAMPLE
    .\Test-Links.ps1
    .\Test-Links.ps1 -Online
#>
[CmdletBinding()]
param(
    [switch]$Online,
    [int]$ThrottleMs = 250
)

$ErrorActionPreference = 'Stop'
$root   = Split-Path -Parent $PSScriptRoot
$readme = Join-Path $root 'README.md'
$audit  = Join-Path $root 'Invoke-EopAudit.ps1'
$errors = New-Object System.Collections.Generic.List[string]
function Fail([string]$m) { $script:errors.Add($m); Write-Host "  FEHLER  $m" -ForegroundColor Red }
function Ok([string]$m)   { Write-Host "  ok      $m" -ForegroundColor Green }

$rd = Get-Content -LiteralPath $readme -Raw -Encoding UTF8
$au = Get-Content -LiteralPath $audit  -Raw -Encoding UTF8

# --- 1. Legende gegen Fundort-Zeilen -------------------------------------------------------------
Write-Host "`n[1] Legende gegen Fundort-Tabellen"
$legend = @{}
foreach ($m in [regex]::Matches($rd, '(?m)^\| \*\*([A-Z0-9]+)\*\* \|[^|]*\|\s*(\d+)\s*\|')) {
    $legend[$m.Groups[1].Value] = [int]$m.Groups[2].Value
}
$rows = @{}      # ID -> Zeilentext
$dupe = @()
foreach ($m in [regex]::Matches($rd, '(?m)^\| `([A-Z0-9]+-\d+)` \|(.*)$')) {
    $id = $m.Groups[1].Value
    if ($rows.ContainsKey($id)) { $dupe += $id }
    $rows[$id] = $m.Groups[2].Value
}
# Skriptgenerierte IDs (INV-*, TABL-FileHash ...) stehen in einer anderen Tabelle und enthalten
# keine Ziffern am Ende bzw. haben das Praefix INV - sie fallen durch das Muster \d+ bereits raus.
if ($dupe) { Fail "ID mehrfach in den Fundort-Tabellen: $($dupe -join ', ')" }
if ($legend.Count -eq 0) { Fail 'Kuerzel-Legende nicht gefunden (Format der README geaendert?)' }
$sum = ($legend.Values | Measure-Object -Sum).Sum
if ($rows.Count -eq $sum) { Ok "$($rows.Count) Fundort-Zeilen = Summe der Legende ($sum)" }
else { Fail "Fundort-Zeilen ($($rows.Count)) <> Summe der Legende ($sum)" }
foreach ($p in ($legend.Keys | Sort-Object)) {
    $n = @($rows.Keys | Where-Object { $_ -match "^$([regex]::Escape($p))-\d+$" }).Count
    if ($n -ne $legend[$p]) { Fail "Praefix $p : Legende sagt $($legend[$p]), Tabellen haben $n" }
}

# --- 2. Learn-Link je Zeile ----------------------------------------------------------------------
Write-Host "`n[2] Learn-Link je Fundort-Zeile"
$learn = @{}     # ID -> URL
foreach ($id in $rows.Keys) {
    $m = [regex]::Match($rows[$id], '\[Learn\]\((https://learn\.microsoft\.com/[^)\s]+)\)')
    if ($m.Success) { $learn[$id] = $m.Groups[1].Value } else { Fail "$id : kein Learn-Link" }
}
if ($learn.Count -eq $rows.Count) { Ok "alle $($rows.Count) Zeilen haben einen Learn-Link" }

# --- 3. README gegen Audit-Skript ----------------------------------------------------------------
Write-Host "`n[3] URL-Menge README gegen Invoke-EopAudit.ps1"
$re = 'https://(?:security\.microsoft\.com|admin\.exchange\.microsoft\.com|learn\.microsoft\.com|admin\.microsoft\.com|entra\.microsoft\.com|purview\.microsoft\.com)[^\s)|''"`<>]*'
function Get-Urls([string]$t) {
    [regex]::Matches($t, $re) | ForEach-Object { $_.Value.TrimEnd('.', ',', ';') } | Sort-Object -Unique
}
$ur = @(Get-Urls $rd)
$us = @(Get-Urls $au)
# Die nackte EAC-Startadresse steht in der README als Fliesstext, nicht als Fundort.
$onlyR = @($ur | Where-Object { $_ -notin $us -and $_ -ne 'https://admin.exchange.microsoft.com' })
$onlyS = @($us | Where-Object { $_ -notin $ur })
if (-not $onlyR -and -not $onlyS) { Ok "$($us.Count) URLs im Skript, alle auch in der README (und umgekehrt)" }
foreach ($u in $onlyR) { Fail "nur in README, nicht im Audit-Skript: $u" }
foreach ($u in $onlyS) { Fail "nur im Audit-Skript, nicht in README: $u" }

# --- 4./5. Online ---------------------------------------------------------------------------------
if ($Online) {
    Write-Host "`n[4/5] Learn-Seiten und Anker (online)"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $pages = @{}   # URL ohne Anker -> HTML oder $null
    $urls  = $learn.GetEnumerator() | Sort-Object Name
    $i = 0
    foreach ($e in $urls) {
        $i++
        $u = $e.Value
        $base, $anchor = $u -split '#', 2
        if (-not $pages.ContainsKey($base)) {
            try {
                $r = Invoke-WebRequest -Uri $base -UseBasicParsing -TimeoutSec 30
                $pages[$base] = [string]$r.Content
            } catch {
                $pages[$base] = $null
                Fail "$($e.Name) : $base nicht abrufbar ($($_.Exception.Message))"
            }
            Start-Sleep -Milliseconds $ThrottleMs
        }
        $html = $pages[$base]
        if ($null -eq $html) { continue }
        if ($anchor) {
            if ($html -notmatch ('id="' + [regex]::Escape($anchor) + '"')) {
                Fail "$($e.Name) : Anker #$anchor fehlt in $base"
            }
        }
    }
    $bad = $errors.Count
    if ($bad -eq 0) { Ok "$($learn.Count) Learn-Links, $($pages.Count) Seiten: alle erreichbar, alle Anker vorhanden" }
}
else {
    Write-Host "`n[4/5] uebersprungen (ohne -Online). Learn-Seiten und Anker wurden NICHT geprueft." -ForegroundColor Yellow
}

Write-Host ''
if ($errors.Count) { Write-Host "$($errors.Count) Fehler." -ForegroundColor Red; exit 1 }
Write-Host 'Alles in Ordnung.' -ForegroundColor Green
exit 0
