<#
.SYNOPSIS
    Prueft den Log-Tab des Free/Busy-Tests.
.DESCRIPTION
    Laedt ExchangeTester.ps1 ohne die Launcher-Schleife, damit kein Fenster modal
    blockiert, und prueft Aufbau und Funktion des Log-Tabs: beide Registerkarten,
    die Zuordnung der Steuerelemente, das Kontextmenue und ob Add-FbLog mit
    Zeitstempel und Einfaerbung schreibt.
.EXAMPLE
    powershell.exe -STA -ExecutionPolicy Bypass -File .\tests\Test-FbLogTab.ps1
#>
# Laedt den ExchangeTester ohne Launcher-Schleife und prueft den neuen Log-Tab.
$ErrorActionPreference = 'Stop'
$root = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path $MyInvocation.MyCommand.Path -Parent }
$src  = Join-Path (Split-Path $root -Parent) 'ExchangeTester.ps1'
if (-not (Test-Path $src)) { throw "ExchangeTester.ps1 nicht gefunden: $src" }
$t = [IO.File]::ReadAllText($src)

# Launcher-Schleife entfernen, damit nichts modal blockiert
$i = $t.IndexOf('while ($true) {')
if ($i -lt 0) { throw 'Launcher-Schleife nicht gefunden' }
$t = $t.Substring(0, $i) + "# Launcher im Test deaktiviert`r`n"

$tmp = Join-Path $env:TEMP 'et-headless.ps1'
[IO.File]::WriteAllText($tmp, $t, (New-Object Text.UTF8Encoding $true))
. $tmp

$fehler = 0
function Pruefe($Name, $Erwartet, $Erhalten) {
    $ok = ("$Erwartet" -eq "$Erhalten")
    '{0} {1,-46} erwartet {2,-24} erhalten {3}' -f $(if ($ok) { '[OK]  ' } else { '[FEHL]' }), $Name, $Erwartet, $Erhalten
    if (-not $ok) { $script:fehler++ }
}

Pruefe 'TabControl im Free/Busy-Fenster vorhanden' $true  ($null -ne $fbTabCtrl)
Pruefe 'Anzahl Tabs'                               2      $fbTabCtrl.TabPages.Count
Pruefe 'Tab 1 heisst Results'                      'Results' $fbTabCtrl.TabPages[0].Text
Pruefe 'Tab 2 heisst Log'                          'Log'     $fbTabCtrl.TabPages[1].Text
Pruefe 'ListView liegt im Results-Tab'             'Results' $fbLvw.Parent.Text
Pruefe 'Logfeld liegt im Log-Tab'                  'Log'     $fbRtbLog.Parent.Text
Pruefe 'Logfeld ist schreibgeschuetzt'             $true     $fbRtbLog.ReadOnly
Pruefe 'ListView-Spalten erhalten'                 3         $fbLvw.Columns.Count
Pruefe 'Kontextmenue am Logfeld'                   3         $fbRtbLog.ContextMenuStrip.Items.Count

# Fuellen die Steuerelemente das Fenster aus? (Anker wirken erst beim Vergroessern -
# wenn die Ausgangsgroesse nicht zur ClientSize passt, bleibt der Bereich leer.)
$rand = 8
Pruefe 'TabControl reicht bis zum rechten Rand'    ($fbForm.ClientSize.Width - $rand)  $fbTabCtrl.Right
Pruefe 'TabControl reicht bis zum unteren Rand'    ($fbForm.ClientSize.Height - $rand) $fbTabCtrl.Bottom
Pruefe 'Close-Knopf am rechten Rand'               ($fbForm.ClientSize.Width - $rand)  $fbBtnClose.Right
Pruefe 'Hinweistext am rechten Rand'               ($fbForm.ClientSize.Width - $rand)  $fbLblOnpHint.Right
Pruefe 'Eingabefeld reicht bis zum Hinweistext'    $true ($fbTxtOnp.Right -le $fbLblOnpHint.Left)

# Schreibt Add-FbLog tatsaechlich in das Feld?
$fbRtbLog.Clear()
Add-FbLog 'POST https://mail.example.com/EWS/Exchange.asmx' 'step'
Add-FbLog '  GetLastError=0; httpStatus=200.' 'ok'
$txt = $fbRtbLog.Text
Pruefe 'Add-FbLog schreibt ins Logfeld'            $true  ($txt -match 'httpStatus=200')
Pruefe 'Zeitstempel vorangestellt'                 $true  ($txt -match '^\d{2}:\d{2}:\d{2}\s')
Pruefe 'zwei Zeilen geschrieben'                   2      (@($fbRtbLog.Lines | Where-Object { $_ -ne '' }).Count)

# Faerbt es die Zeilen unterschiedlich ein?
$fbRtbLog.SelectionStart = 0; $fbRtbLog.SelectionLength = 10
$c1 = $fbRtbLog.SelectionColor
$fbRtbLog.SelectionStart = $txt.IndexOf('GetLastError'); $fbRtbLog.SelectionLength = 8
$c2 = $fbRtbLog.SelectionColor
Pruefe 'Zeilen unterschiedlich eingefaerbt'        $true  ($c1 -ne $c2)

Remove-Item $tmp -Force
''
if ($fehler -eq 0) { 'ERGEBNIS: alle Pruefungen bestanden'; exit 0 } else { "ERGEBNIS: $fehler fehlgeschlagen"; exit 1 }
