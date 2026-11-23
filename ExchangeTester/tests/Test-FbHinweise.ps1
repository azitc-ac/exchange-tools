<#
.SYNOPSIS
    Prueft die Deutung im Free/Busy-Test: Richtung A (On-Prem -> EXO) verweist auf
    die Dedicated Hybrid App, Richtung B behaelt die Hinweise zu Partner-App/LinkedAccount.
.DESCRIPTION
    Anlass: Richtung A schickte bei Fehlern zu LinkedAccount und Partner-App. Die gehoeren
    zur Gegenrichtung; On-Prem -> EXO holt das Token seit 31.10.2025 nur noch ueber die
    Dedicated Hybrid App (HCW legt sie an, schaltet sie aber nicht ein).
    Laedt ExchangeTester.ps1 ohne Launcher-Schleife wie Test-FbLogTab.ps1.
.PARAMETER Source
    Anderes Skript pruefen (fuer Gegenproben mit eingebautem Fehler).
#>
param([string]$Source)
$ErrorActionPreference = 'Stop'
$root = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path $MyInvocation.MyCommand.Path -Parent }
$src  = if ($Source) { $Source } else { Join-Path (Split-Path $root -Parent) 'ExchangeTester.ps1' }
$t = [IO.File]::ReadAllText($src)
$i = $t.IndexOf('while ($true) {')
if ($i -lt 0) { throw 'Launcher-Schleife nicht gefunden' }
$t = $t.Substring(0, $i) + "# Launcher im Test deaktiviert`r`n"
$tmp = Join-Path $env:TEMP ('et-hinweise-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.ps1')
[IO.File]::WriteAllText($tmp, $t, (New-Object Text.UTF8Encoding $true))
. $tmp

$fehler = 0; $ok = 0
function Pruefe([string]$Name, [bool]$Bedingung) {
    if ($Bedingung) { $script:ok++; "[OK]   $Name" } else { $script:fehler++; "[FEHL] $Name" }
}
function Log-Von([scriptblock]$sb) { $fbRtbLog.Clear(); & $sb; $fbRtbLog.Text }

$err   = @{ Http = 200; ResponseCode = 'ErrorProxyRequestProcessingFailed'; Message = 'The OAuth token request failed.'; Fault = ''; ViewType = '' }
$nf    = @{ Http = 200; ResponseCode = 'ErrorMailRecipientNotFound'; Message = 'Recipient not found.'; Fault = ''; ViewType = '' }
$empty = @{ Http = 200; ResponseCode = 'NoError'; ViewType = 'None'; SlotsTotal = 0; EventCount = 0 }
$full  = @{ Http = 200; ResponseCode = 'NoError'; ViewType = 'DetailedMerged'; SlotsTotal = 168; SlotsBusy = 11; EventCount = 3 }

'--- Ablauf, Richtung A ---'
$l = Log-Von { Add-FbExplain 'A' }
Pruefe 'A: nennt die Dedicated Hybrid App'              ($l -match 'dedicated hybrid app')
Pruefe 'A: nennt Graph-Weg mit Override'                ($l -match 'EnableRouteThroughMSGraphFeature')
Pruefe 'A: DAuth fuehrt nicht mehr zu EXO'              ($l -match 'no longer reaches Exchange Online')
Pruefe 'A: kein Partner-App-Token mehr'                 ($l -notmatch 'partner application')

'--- Fehler, Richtung A ---'
$l = Log-Von { Add-FbWhereToLook 'A' $err }
Pruefe 'A-Fehler: Override-Section genannt'             ($l -match 'ExchangeOnpremAsThirdPartyAppId')
Pruefe 'A-Fehler: HCW schaltet nicht ein, 60 min'       ($l -match 'does not enable it' -and $l -match '60 min')
Pruefe 'A-Fehler: AuthServer mit App-ID'                ($l -match 'Get-AuthServer' -and $l -match 'EvoSts')
Pruefe 'A-Fehler: Test-OAuthConnectivity mit FullId'    ($l -match 'outlook\.office365\.com -Mailbox' -and $l -match 'L:<app ID>-AS:')
Pruefe 'A-Fehler: EWS-Log nach EXNL-Graph'              ($l -match 'EXNL-GraphWebRequestProcessingException')
Pruefe 'A-Fehler: ausgehend login + graph'              ($l -match 'login\.microsoftonline\.com' -and $l -match 'graph\.microsoft\.com')
Pruefe 'A-Fehler: KEIN LinkedAccount-Hinweis'           ($l -notmatch 'LinkedAccount')

$l = Log-Von { Add-FbWhereToLook 'A' $nf }
Pruefe 'A-unbekanntes Ziel: Empfaenger-Hinweis'         ($l -match 'target address is unknown')
Pruefe 'A-unbekanntes Ziel: keine App-Checkliste'       ($l -notmatch 'ExchangeOnpremAsThirdPartyAppId')

'--- Fehler, Richtung B bleibt ---'
$l = Log-Von { Add-FbWhereToLook 'B' $err }
Pruefe 'B-Fehler: LinkedAccount-Hinweis bleibt'         ($l -match 'LinkedAccount')
Pruefe 'B-Fehler: keine Dedicated-App-Checkliste'       ($l -notmatch 'ExchangeOnpremAsThirdPartyAppId')

'--- leer und OK, Richtung A ---'
$l = Log-Von { Add-FbWhereToLook 'A' $empty }
Pruefe 'A-leer: Kalenderrecht Default'                  ($l -match 'Default entry')
Pruefe 'A-leer: FreeBusyAccessLevel nur fuer EWS belegt' ($l -match 'EWS route' -and $l -match 'not verified')
$l = Log-Von { Add-FbWhereToLook 'A' $full }
Pruefe 'A-OK: Zeiten gegen den Kalender (KB 5125804)'    ($l -match 'KB 5125804')
$l = Log-Von { Add-FbWhereToLook 'B' $full }
Pruefe 'B-OK: kein Graph-Zeithinweis'                   ($l -notmatch 'KB 5125804')

'--- Ergebniszeile ---'
$fbLvw.Items.Clear()
Add-FbRow 'On-Prem → EXO' $err 'ctx' -Direction A
Add-FbRow 'EXO → On-Prem' $err 'ctx' -Direction B
Add-FbRow 'On-Prem → EXO' $full 'ctx' -Direction A
$dA = $fbLvw.Items[0].SubItems[2].Text; $dB = $fbLvw.Items[1].SubItems[2].Text; $dOk = $fbLvw.Items[2].SubItems[2].Text
Pruefe 'Zeile A FAIL'                                   ($fbLvw.Items[0].SubItems[1].Text -eq 'FAIL')
Pruefe 'Zeile A: Hinweis auf Dedicated App'             ($dA -match 'ExchangeOnpremAsThirdPartyAppId')
Pruefe 'Zeile A: kein LinkedAccount'                    ($dA -notmatch 'LinkedAccount')
Pruefe 'Zeile B: LinkedAccount bleibt'                  ($dB -match 'LinkedAccount')
Pruefe 'Zeile A OK: Zeiten pruefen'                     ($fbLvw.Items[2].SubItems[1].Text -eq 'OK' -and $dOk -match 'Compare the times')

Remove-Item $tmp -Force
''
"Bestanden: $ok   Fehlgeschlagen: $fehler"
if ($fehler -eq 0) { exit 0 } else { exit 1 }
