<#
.SYNOPSIS
    Regressionstest: wachsende UTF-16-Protokolle beim Nachladen.

.DESCRIPTION
    msiexec schreibt seine Verbose-Protokolle als UTF-16LE. Wird eine solche Datei geoeffnet,
    waehrend sie noch leer ist (Installation laeuft gerade an), kann die Codierung zu diesem
    Zeitpunkt nicht erkannt werden. Ohne Gegenmassnahme bleibt es bei der Annahme UTF-8 und
    alles Nachgeladene erscheint 8-bittig - im Fenster als Text mit einem Leerzeichen zwischen
    jedem Zeichen und einer Leerzeile nach jeder Zeile.

    Der Test laedt den echten Code aus LogViewer.ps1 (nur ohne Application::Run) und ruft
    Open-File und Invoke-Append auf, waehrend die Datei in Schueben waechst.

.PARAMETER Rueckbau
    Nimmt beide Gegenmassnahmen im geladenen Code zurueck. Der Test MUSS damit fehlschlagen -
    sonst prueft er nichts.

.EXAMPLE
    .\Test-GrowingUtf16.ps1
    .\Test-GrowingUtf16.ps1 -Rueckbau
#>
param([switch]$Rueckbau)
$ErrorActionPreference = 'Stop'

$root = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path $MyInvocation.MyCommand.Path -Parent }
$orig = Join-Path (Split-Path $root -Parent) 'LogViewer.ps1'
if (-not (Test-Path $orig)) { throw "LogViewer.ps1 nicht gefunden: $orig" }

$text = [IO.File]::ReadAllText($orig)
$text = $text.Replace('if (-not $script:HeadlessOnly) { [System.Windows.Forms.Application]::Run($form) }',
                      '# Im Test laeuft kein Fenster')
if ($Rueckbau) {
    $text = $text.Replace('if ($chunk.Encoding) { $s.Encoding = $chunk.Encoding }', '# rueckgebaut')
    $text = $text.Replace('if ($s.EncodingProvisional) {', 'if ($false) {')
}

$tmpScript = Join-Path $env:TEMP 'LogViewer-headless.ps1'
[IO.File]::WriteAllText($tmpScript, $text, (New-Object Text.UTF8Encoding $true))

$log = Join-Path $env:TEMP 'GrowingUtf16-Test.log'
if (Test-Path $log) { Remove-Item $log -Force }
[IO.File]::WriteAllBytes($log, @())        # leer, so legt msiexec sie an

. $tmpScript                                # baut Fenster und Funktionen auf, zeigt nichts

Open-File $log
"nach dem Oeffnen : $($s.VirtualRows.Count) Zeilen, Codierung $($s.Encoding.EncodingName), vorlaeufig=$($s.EncodingProvisional)"

foreach ($schub in 1..2) {
    # UTF-16LE, BOM nur beim ersten Schreiben, Zeilenende LF CR LF wie bei msiexec
    $bom = ((Get-Item $log).Length -eq 0)
    $sw = New-Object IO.StreamWriter($log, $true, (New-Object Text.UnicodeEncoding($false, $bom)))
    foreach ($i in 1..20) { $sw.Write("MSI (s) (D8:5C): Component: __cmp_Schub$schub`_$i; Installed: Null`n`r`n") }
    $sw.Close()
    Invoke-Append                           # im Betrieb ruft das der Poll-Timer
    "nach Schub $schub    : $($s.VirtualRows.Count) Zeilen, Codierung $($s.Encoding.EncodingName)"
}

$alle    = ($s.VirtualRows | ForEach-Object { $_[0] }) -join "`n"
$nul     = ([char[]]$alle | Where-Object { $_ -eq [char]0 }).Count
$treffer = (1..20 | Where-Object { $alle -match "__cmp_Schub2_$_;" }).Count
""
"NUL-Zeichen in allen Zeilen    : $nul        (erwartet 0)"
"Eintraege aus Schub 2 sichtbar : $treffer von 20"

Remove-Item $log, $tmpScript -Force
$ok = ($nul -eq 0) -and ($treffer -eq 20)
if ($ok) { 'ERGEBNIS: sauber'; exit 0 } else { 'ERGEBNIS: defekt'; exit 1 }
