# The TIFF pictures the engine tests read, written by Windows' own encoder
# (the Windows Imaging Component, through WPF), so the reader is held to
# files another writer made: CCITT fax Group 3 and 4, PackBits, LZW, Deflate
# and none, and a page of each in one multi-page file. Beside each, a PNG of
# the same pixels from the same encoder, for the tests to compare against.
#
#   powershell -ExecutionPolicy Bypass -File tools/make-picture-fixtures.ps1
#
# Made by this project from content drawn here; nothing is downloaded.

Add-Type -AssemblyName PresentationCore, WindowsBase
$out = Join-Path $PSScriptRoot '..\tests\fixtures\pictures'
New-Item -ItemType Directory -Force $out | Out-Null

function Draw([int]$w, [int]$h, [int]$seed) {
  $visual = New-Object System.Windows.Media.DrawingVisual
  $dc = $visual.RenderOpen()
  $dc.DrawRectangle([System.Windows.Media.Brushes]::White, $null, (New-Object System.Windows.Rect 0, 0, $w, $h))
  $grad = New-Object System.Windows.Media.LinearGradientBrush ([System.Windows.Media.Color]::FromRgb(30, 90, 200)), ([System.Windows.Media.Color]::FromRgb(240, 180, 40)), 0
  $dc.DrawRectangle($grad, $null, (New-Object System.Windows.Rect 8, 8, ($w - 16), ($h / 3)))
  $pen = New-Object System.Windows.Media.Pen ([System.Windows.Media.Brushes]::Black), 3
  $dc.DrawEllipse([System.Windows.Media.Brushes]::Crimson, $pen, (New-Object System.Windows.Point ($w / 3), ($h * 0.66)), ($w / 6), ($h / 6))
  $dc.DrawLine($pen, (New-Object System.Windows.Point 4, ($h - 6)), (New-Object System.Windows.Point ($w - 4), (12 + $seed * 9)))
  $type = New-Object System.Windows.Media.Typeface 'Segoe UI'
  $text = New-Object System.Windows.Media.FormattedText ("Page $seed fax", [System.Globalization.CultureInfo]::InvariantCulture, [System.Windows.FlowDirection]::LeftToRight, $type, 22, [System.Windows.Media.Brushes]::Black, 1.0)
  $dc.DrawText($text, (New-Object System.Windows.Point ($w / 2), ($h / 2)))
  $dc.Close()
  $bmp = New-Object System.Windows.Media.Imaging.RenderTargetBitmap $w, $h, 96, 96, ([System.Windows.Media.PixelFormats]::Pbgra32)
  $bmp.Render($visual)
  $opaque = New-Object System.Windows.Media.Imaging.FormatConvertedBitmap $bmp, ([System.Windows.Media.PixelFormats]::Bgr24), $null, 0
  return $opaque
}

function Save($encoder, $frames, $name) {
  foreach ($f in $frames) { $encoder.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($f)) }
  $stream = [System.IO.File]::Create((Join-Path $out $name))
  $encoder.Save($stream)
  $stream.Close()
}

$colour = Draw 203 131 1
$bw = New-Object System.Windows.Media.Imaging.FormatConvertedBitmap $colour, ([System.Windows.Media.PixelFormats]::BlackWhite), $null, 0

foreach ($c in @('None', 'Rle', 'Lzw', 'Zip')) {
  $e = New-Object System.Windows.Media.Imaging.TiffBitmapEncoder
  $e.Compression = [System.Windows.Media.Imaging.TiffCompressOption]::$c
  Save $e @($colour) "wic-$($c.ToLower()).tif"
}
foreach ($c in @('Ccitt3', 'Ccitt4', 'None', 'Rle', 'Lzw')) {
  $e = New-Object System.Windows.Media.Imaging.TiffBitmapEncoder
  $e.Compression = [System.Windows.Media.Imaging.TiffCompressOption]::$c
  Save $e @($bw) "wic-bw-$($c.ToLower()).tif"
}
$pages = @((Draw 160 120 1), (Draw 120 160 2), (Draw 200 100 3))
$e = New-Object System.Windows.Media.Imaging.TiffBitmapEncoder
$e.Compression = [System.Windows.Media.Imaging.TiffCompressOption]::Lzw
Save $e $pages 'wic-pages.tif'

$e = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
Save $e @($colour) 'wic-colour.png'
$e = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
Save $e @($bw) 'wic-bw.png'
for ($i = 0; $i -lt 3; $i++) {
  $e = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
  Save $e @($pages[$i]) "wic-page-$($i + 1).png"
}
Get-ChildItem $out | Select-Object Name, Length | Format-Table -AutoSize
