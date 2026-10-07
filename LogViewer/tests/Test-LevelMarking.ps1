<#
.SYNOPSIS
    Prüft die Erkennung von Warnungen und Fehlern für die Markierungsoption.

.DESCRIPTION
    Die Option "Warnungen/Fehler markieren" färbt Zeilen in Formaten ohne Typspalte.
    Welche Zeile gelb und welche rot wird, hängt allein an zwei Ausdrücken im
    Hauptskript - die werden hier mit echten Protokollzeilen geprüft, nicht mit
    ausgedachten.

    Zwei Dinge sollen dabei halten: dass die üblichen Schreibweisen erkannt werden
    (auch das *ERROR* des Hybrid Configuration Wizard), und dass harmlose Wörter mit
    "error" oder "warn" darin NICHT einfärben.
#>
[CmdletBinding()]
param([string]$Skript)

# $PSScriptRoot kommt beim Start ueber -File in manchen Shells leer an.
if (-not $Skript) {
    $hier = if ($PSScriptRoot) { $PSScriptRoot }
            elseif ($MyInvocation.MyCommand.Path) { Split-Path $MyInvocation.MyCommand.Path -Parent }
            else { (Get-Location).Path }
    $Skript = Join-Path (Split-Path $hier -Parent) 'LogViewer.ps1'
}
$ErrorActionPreference = 'Stop'

# Die beiden Ausdrücke aus dem Hauptskript holen, statt sie hier zu wiederholen -
# sonst prüft der Test seine eigene Kopie und merkt eine Änderung nie.
$text = [IO.File]::ReadAllText($Skript, [Text.Encoding]::UTF8)
foreach ($name in 'RxError', 'RxWarn') {
    $m = [regex]::Match($text, ('\$script:{0}\s*=\s*\[regex\]::new\(''([^'']+)''' -f $name))
    if (-not $m.Success) { throw "Ausdruck `$script:$name nicht in $Skript gefunden." }
    Set-Variable -Name $name -Value ([regex]::new($m.Groups[1].Value,
        [System.Text.RegularExpressions.RegexOptions]'IgnoreCase, Compiled'))
}

function Get-Level([string]$Zeile) {
    if ($RxError.IsMatch($Zeile)) { return 3 }
    if ($RxWarn.IsMatch($Zeile))  { return 2 }
    0
}

$pass = 0; $fail = 0
function T([string]$Name, [string]$Zeile, [int]$Soll) {
    $ist = Get-Level $Zeile
    if ($ist -eq $Soll) { $script:pass++; Write-Output "  OK   $Name" }
    else { $script:fail++; Write-Output ("  FEHL {0}: Stufe {1}, erwartet {2}  <{3}>" -f $Name, $ist, $Soll, $Zeile) }
}

Write-Output "`n######## Fehler (Stufe 3) ########"
T 'HCW mit Sternen'        '2026.09.17 23:20:00.089 *ERROR* 10085 [Client=UX] Connecting to remote server failed' 3
T 'eckige Klammern'        '2026-10-05 11:48:40 [ERROR] Das Zertifikat konnte nicht geladen werden.'              3
T 'deutsches FEHLER'       '12:03:11  FEHLER  Der Dienst antwortet nicht.'                                        3
T 'FATAL'                  'FATAL: unrecoverable state, aborting'                                                 3
T 'CRITICAL'               'severity=CRITICAL component=transport'                                                3
T 'Error am Zeilenende'    'Get-Queue returned an Error'                                                          3

Write-Output "`n######## Warnungen (Stufe 2) ########"
T 'WARN in Klammern'       '2026-10-05 11:48:41 [WARN] Posh-ACME: Order is not recommended for renewal yet.'       2
T 'WARNING ausgeschrieben' '2026.09.17 23:21:19 WARNING: certificate expires in 12 days'                          2
T 'deutsches WARNUNG'      '09:14:22  WARNUNG  Postfachdatenbank fast voll.'                                      2

Write-Output "`n######## Weder noch (Stufe 0) ########"
T 'reine Informationszeile' '2026-10-05 11:48:38 [INFO] Posh-ACME 4.34.0 ready, home = C:\Tools\AcmeExchange'      0
T 'Erfolgsmeldung'          'FINISH Time=1714.0ms Results=NoContent'                                              0
T 'leere Zeile'             ''                                                                                    0

Write-Output "`n######## Keine Fehlalarme bei aehnlichen Woertern ########"
T 'Terrorliste'            'Die Terrorliste wurde aktualisiert.'                                                   0
T 'Warner als Name'        'Benutzer Warner, Thomas angemeldet'                                                    0
T 'errorless'              'The run completed errorless'                                                           0
T 'Dateiname mit error'    'Lade C:\Logs\errorhandling.dll'                                                        0

Write-Output "`n######## Fehler gewinnt gegen Warnung ########"
T 'beides in einer Zeile'  '[WARN] retry failed with ERROR 0x80070005'                                             3

Write-Output "`n================ Bestanden: $pass   Fehlgeschlagen: $fail ================"
if ($fail -gt 0) { exit 1 }
