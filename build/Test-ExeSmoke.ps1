<#
.SYNOPSIS
    Startet jede gebaute EXE und prüft, ob wirklich ein Fenster erscheint.

.DESCRIPTION
    Entstanden aus einem Fehler: Sieben EXE-Werkzeuge wurden gebaut, veröffentlicht
    und nie gestartet. Der ExchangeTester brach beim Start mit

        Das Argument kann nicht an den Parameter "Path" gebunden werden, da es NULL ist.

    ab, weil in einer ps2exe-EXE sowohl $PSScriptRoot als auch
    $MyInvocation.MyCommand.Path leer sind. Ein erfolgreicher Build sagt darüber nichts.

    Geprüft wird nicht "der Prozess lebt" - ein hängender Prozess lebt auch -, sondern
    ob ein sichtbares Fenster auftaucht. Dafür reicht MainWindowHandle nicht: Ein
    Werkzeug, das als Erstes einen Dialog öffnet (ImportExchangePfx), hat keins. Darum
    die Fensterliste des Prozesses über die Win32-API.

    Welche EXE sich automatisch starten lässt, steht in der release.psd1 des Werkzeugs
    unter SmokeTest; wo das nicht geht, nennt SmokeSkipGrund den Grund.

    Läuft nicht in der CI: Ein GitHub-Runner hat keine brauchbare Desktop-Sitzung.
#>
[CmdletBinding()]
param([string]$Repo, [int]$TimeoutSekunden = 25)

# $PSScriptRoot kommt beim Start ueber -File in manchen Shells leer an; dann den
# eigenen Pfad anders ermitteln, sonst scheitert schon die Parameterbindung.
if (-not $Repo) {
    $hier = if ($PSScriptRoot) { $PSScriptRoot }
            elseif ($MyInvocation.MyCommand.Path) { Split-Path $MyInvocation.MyCommand.Path -Parent }
            else { (Get-Location).Path }
    $Repo = Split-Path $hier -Parent
}
$ErrorActionPreference = 'Stop'

if ($env:GITHUB_ACTIONS -eq 'true' -or $env:CI -eq 'true') {
    Write-Output '  uebersprungen: kein Desktop in der CI'
    exit 0
}

Add-Type @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public class SmokeWin {
  delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, IntPtr l);
  // Nur Fenster MIT Titel: ein Werkzeug bringt oft unsichtbare Hilfsfenster mit, und
  // deren Symbol sagt nichts ueber die Titelleiste des Hauptfensters.
  public static IntPtr HandleForPid(uint target) {
    IntPtr found = IntPtr.Zero;
    EnumWindows((h, l) => {
      uint pid; GetWindowThreadProcessId(h, out pid);
      if (pid == target && IsWindowVisible(h)) {
        var sb = new StringBuilder(300); GetWindowText(h, sb, 300);
        if (sb.Length > 0) { found = h; return false; }
      }
      return true;
    }, IntPtr.Zero);
    return found; }
  public static List<string> ForPid(uint target) {
    var res = new List<string>();
    EnumWindows((h, l) => {
      uint pid; GetWindowThreadProcessId(h, out pid);
      if (pid == target && IsWindowVisible(h)) {
        var sb = new StringBuilder(300); GetWindowText(h, sb, 300);
        res.Add(sb.Length > 0 ? sb.ToString() : "(ohne Titel)");
      }
      return true;
    }, IntPtr.Zero);
    return res;
  }
}
'@ -ErrorAction SilentlyContinue

$pass = 0; $fail = 0
$werkzeuge = Get-ChildItem $Repo -Directory |
             Where-Object { $_.Name -notmatch '^\.|^build$|^dist$' } | Sort-Object Name

foreach ($w in $werkzeuge) {
    $meta = Import-PowerShellDataFile (Join-Path $w.FullName 'release.psd1')
    if (-not $meta.Exe) { continue }

    if (-not $meta.SmokeTest) {
        $grund = if ($meta.SmokeSkipGrund) { $meta.SmokeSkipGrund } else { 'kein Grund hinterlegt' }
        Write-Output ("  uebersprungen {0,-22} {1}" -f $w.Name, $grund)
        continue
    }

    $exe = Join-Path $w.FullName $meta.Exe
    if (-not (Test-Path $exe)) {
        $fail++; Write-Output "  FEHL $($w.Name): $($meta.Exe) nicht gebaut"; continue
    }

    $p = Start-Process -FilePath $exe -PassThru
    $fenster = @(); $abgestuerzt = $false; $sek = 0; $fensterHandle = [IntPtr]::Zero
    for ($i = 1; $i -le $TimeoutSekunden; $i++) {
        Start-Sleep -Seconds 1
        $p.Refresh()
        if ($p.HasExited) { $abgestuerzt = $true; $sek = $i; break }
        $fenster = @([SmokeWin]::ForPid([uint32]$p.Id))
        if ($fenster.Count -gt 0) { $sek = $i; $fensterHandle = [SmokeWin]::HandleForPid([uint32]$p.Id); break }
    }

    # Symbolgroesse noch am LEBENDEN Fenster ermitteln - nach dem Beenden gibt es nichts
    # mehr zu fragen, und WM_GETICON liefert dann stumm eine Null.
    $iconGroesse = -1
    if ($meta.Icon -and $fensterHandle -ne [IntPtr]::Zero -and -not $p.HasExited) {
        # Das Symbol wird ueblicherweise im Shown-Ereignis gesetzt, also erst kurz nachdem
        # das Fenster sichtbar wurde. Ohne diese Pause misst man zu frueh.
        Start-Sleep -Milliseconds 1200
        try {
            Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue
            $hIco = [SmokeWin]::SendMessage($fensterHandle, 0x7F, [IntPtr]0, [IntPtr]::Zero)   # WM_GETICON, ICON_SMALL
            if ($hIco -ne [IntPtr]::Zero) { $iconGroesse = [System.Drawing.Icon]::FromHandle($hIco).Width }
        }
        catch { }
    }
    if (-not $p.HasExited) { $p.Kill(); $p.WaitForExit(5000) | Out-Null }

    if ($abgestuerzt) {
        $fail++; Write-Output ("  FEHL {0,-22} abgestuerzt nach {1}s (ExitCode {2})" -f $w.Name, $sek, $p.ExitCode)
    }
    elseif ($fenster.Count -eq 0) {
        $fail++; Write-Output ("  FEHL {0,-22} kein Fenster nach {1}s" -f $w.Name, $TimeoutSekunden)
    }
    else {
        $pass++; Write-Output ("  OK   {0,-22} Fenster nach {1}s: '{2}'" -f $w.Name, $sek, $fenster[0])

        # Titelleisten-Symbol: Windows holt sich dafuer ICON_SMALL. Wer das Fenstersymbol
        # ueber ExtractAssociatedIcon setzt, bekommt nur die 32er-Fassung und Windows
        # rechnet sie herunter - das Symbol wirkt dann matschig. Hier faellt das auf.
        if ($meta.Icon) {
            if ($iconGroesse -lt 0) { $fail++; Write-Output ("  FEHL {0,-22} Fenster ohne kleines Symbol" -f $w.Name) }
            elseif ($iconGroesse -le 16) { $pass++; Write-Output ("  OK   {0,-22} Titelleisten-Symbol {1}x{1}" -f $w.Name, $iconGroesse) }
            else { $fail++; Write-Output ("  FEHL {0,-22} Titelleisten-Symbol ist {1}x{1} statt 16x16 - heruntergerechnet" -f $w.Name, $iconGroesse) }
        }
    }
}

Write-Output ("  {0} gestartet, {1} ohne Fenster" -f ($pass + $fail), $fail)
if ($fail -gt 0) { exit 1 }
