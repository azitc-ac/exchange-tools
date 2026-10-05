<#
.SYNOPSIS
    Erzeugt LogViewer.ico - ein Klemmbrett mit Textzeilen.

.DESCRIPTION
    Zeichnet das Symbol in mehreren Größen und schreibt sie als PNG-Einträge in eine ICO-Datei
    (seit Vista zulässig und für 256er-Größen üblich). Ohne Zusatzwerkzeuge, nur .NET.

.EXAMPLE
    .\New-LogViewerIcon.ps1
#>
[CmdletBinding()]
param([string]$OutFile)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

if (-not $OutFile) {
    $root = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path $MyInvocation.MyCommand.Path -Parent }
    $OutFile = Join-Path $root 'LogViewer.ico'
}

function New-ClipboardBitmap {
    param([int]$Size)
    $bmp = New-Object Drawing.Bitmap($Size, $Size, [Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.Clear([Drawing.Color]::Transparent)
    $u = $Size / 32.0          # alles relativ zu einem 32er-Raster

    function Rect([double]$x, [double]$y, [double]$w, [double]$h) {
        New-Object Drawing.RectangleF(($x * $u), ($y * $u), ($w * $u), ($h * $u))
    }
    function RoundedPath($r, [double]$radius) {
        $rad = $radius * $u
        $p = New-Object Drawing.Drawing2D.GraphicsPath
        $d = $rad * 2
        $p.AddArc($r.X, $r.Y, $d, $d, 180, 90)
        $p.AddArc(($r.Right - $d), $r.Y, $d, $d, 270, 90)
        $p.AddArc(($r.Right - $d), ($r.Bottom - $d), $d, $d, 0, 90)
        $p.AddArc($r.X, ($r.Bottom - $d), $d, $d, 90, 90)
        $p.CloseFigure()
        return $p
    }

    # Brett
    $board = RoundedPath (Rect 4 3 24 27) 2.5
    $brushBoard = New-Object Drawing.Drawing2D.LinearGradientBrush(
        (Rect 4 3 24 27), [Drawing.Color]::FromArgb(255,150,99,56), [Drawing.Color]::FromArgb(255,113,72,39), 60.0)
    $g.FillPath($brushBoard, $board)
    $penBoard = New-Object Drawing.Pen([Drawing.Color]::FromArgb(255,86,54,28), (0.9 * $u))
    $g.DrawPath($penBoard, $board)

    # Papier
    $paper = RoundedPath (Rect 6.5 6 19 21.5) 1.2
    $g.FillPath([Drawing.Brushes]::White, $paper)
    $penPaper = New-Object Drawing.Pen([Drawing.Color]::FromArgb(255,200,200,200), (0.4 * $u))
    $g.DrawPath($penPaper, $paper)

    # Klammer
    $clipTop = RoundedPath (Rect 12 1.5 8 4.5) 1.2
    $g.FillPath((New-Object Drawing.SolidBrush([Drawing.Color]::FromArgb(255,176,180,186))), $clipTop)
    $penClip = New-Object Drawing.Pen([Drawing.Color]::FromArgb(255,120,125,132), (0.5 * $u))
    $g.DrawPath($penClip, $clipTop)

    # Textzeilen - eine davon rot, damit man die Fehlermarkierung wiedererkennt
    $lines = @(
        @{ Y = 10.5; W = 13; C = [Drawing.Color]::FromArgb(255,120,130,140) }
        @{ Y = 13.5; W = 10; C = [Drawing.Color]::FromArgb(255,120,130,140) }
        @{ Y = 16.5; W = 14; C = [Drawing.Color]::FromArgb(255,196,60,52) }
        @{ Y = 19.5; W = 11; C = [Drawing.Color]::FromArgb(255,120,130,140) }
        @{ Y = 22.5; W = 13; C = [Drawing.Color]::FromArgb(255,120,130,140) }
    )
    foreach ($l in $lines) {
        $h = [Math]::Max(1.0, 1.4 * $u)
        $brush = New-Object Drawing.SolidBrush($l.C)
        $g.FillRectangle($brush, (9 * $u), ($l.Y * $u), ($l.W * $u), $h)
        $brush.Dispose()
    }

    $g.Dispose()
    return $bmp
}

function ConvertTo-IconDib {
    <#
      Wandelt ein Bitmap in einen klassischen ICO-Eintrag (BITMAPINFOHEADER + BGRA von unten nach
      oben + AND-Maske). Nötig, weil nicht jeder Verbraucher PNG-Einträge versteht - Icon.ToBitmap
      im .NET Framework etwa scheitert daran.
    #>
    param([Drawing.Bitmap]$Bitmap)
    $w = $Bitmap.Width; $h = $Bitmap.Height
    $ms = New-Object IO.MemoryStream
    $bw = New-Object IO.BinaryWriter($ms)
    # BITMAPINFOHEADER - Höhe doppelt, weil Farb- und Maskenteil zusammen gezählt werden
    $bw.Write([UInt32]40); $bw.Write([Int32]$w); $bw.Write([Int32]($h * 2))
    $bw.Write([UInt16]1); $bw.Write([UInt16]32); $bw.Write([UInt32]0)
    $bw.Write([UInt32]($w * $h * 4)); $bw.Write([Int32]0); $bw.Write([Int32]0)
    $bw.Write([UInt32]0); $bw.Write([UInt32]0)

    $data = $Bitmap.LockBits((New-Object Drawing.Rectangle(0, 0, $w, $h)),
            [Drawing.Imaging.ImageLockMode]::ReadOnly, [Drawing.Imaging.PixelFormat]::Format32bppArgb)
    try {
        $row = New-Object byte[] ($w * 4)
        for ($y = $h - 1; $y -ge 0; $y--) {     # von unten nach oben
            [Runtime.InteropServices.Marshal]::Copy(
                [IntPtr]($data.Scan0.ToInt64() + ($y * $data.Stride)), $row, 0, $row.Length)
            $bw.Write($row)
        }
    } finally { $Bitmap.UnlockBits($data) }

    # AND-Maske: alles sichtbar (Transparenz steckt im Alphakanal), Zeilen auf 4 Byte ausgerichtet
    $maskRow = [math]::Ceiling($w / 8.0)
    if ($maskRow % 4 -ne 0) { $maskRow += 4 - ($maskRow % 4) }
    $zero = New-Object byte[] $maskRow
    for ($y = 0; $y -lt $h; $y++) { $bw.Write($zero) }

    $bw.Flush()
    $bytes = $ms.ToArray()
    $bw.Dispose(); $ms.Dispose()
    return , $bytes
}

# Kleine Größen klassisch (überall lesbar), große als PNG (dafür ist das Format gedacht)
$sizes = @(16, 24, 32, 48, 64, 128, 256)
$entries = @()
foreach ($s in $sizes) {
    $bmp = New-ClipboardBitmap -Size $s
    if ($s -le 64) {
        $entries += , (ConvertTo-IconDib -Bitmap $bmp)
    } else {
        $ms = New-Object IO.MemoryStream
        $bmp.Save($ms, [Drawing.Imaging.ImageFormat]::Png)
        $entries += , $ms.ToArray()
        $ms.Dispose()
    }
    $bmp.Dispose()
}

# ICO zusammensetzen: Header (6 Byte) + je Größe ein Verzeichniseintrag (16 Byte) + die Daten
$fs = New-Object IO.MemoryStream
$bw = New-Object IO.BinaryWriter($fs)
$bw.Write([UInt16]0); $bw.Write([UInt16]1); $bw.Write([UInt16]$sizes.Count)
$offset = 6 + (16 * $sizes.Count)
for ($i = 0; $i -lt $sizes.Count; $i++) {
    $s = $sizes[$i]
    $bw.Write([byte]$(if ($s -ge 256) { 0 } else { $s }))   # 0 bedeutet 256
    $bw.Write([byte]$(if ($s -ge 256) { 0 } else { $s }))
    $bw.Write([byte]0)        # Farbanzahl
    $bw.Write([byte]0)        # reserviert
    $bw.Write([UInt16]1)      # Ebenen
    $bw.Write([UInt16]32)     # Bit pro Pixel
    $bw.Write([UInt32]$entries[$i].Length)
    $bw.Write([UInt32]$offset)
    $offset += $entries[$i].Length
}
foreach ($e in $entries) { $bw.Write($e) }
$bw.Flush()
[IO.File]::WriteAllBytes($OutFile, $fs.ToArray())
$bw.Dispose(); $fs.Dispose()

$fi = Get-Item $OutFile
Write-Host ("Symbol erzeugt: {0} ({1:N1} KB, Größen: {2})" -f $fi.FullName, ($fi.Length / 1KB), ($sizes -join ', ')) -ForegroundColor Green

# Gegenprobe: lässt es sich als Symbol laden?
try {
    $ico = New-Object Drawing.Icon($OutFile)
    Write-Host ("Gegenprobe: ladbar, Standardgröße {0}x{1}" -f $ico.Width, $ico.Height) -ForegroundColor Green
    $ico.Dispose()
} catch {
    throw "Die erzeugte Datei ließ sich nicht als Symbol laden: $($_.Exception.Message)"
}
