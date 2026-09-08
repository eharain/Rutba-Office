# Microsoft Excel's rendering of a workbook, as a PDF, for comparison with ours.
#
#   powershell -File tools/compare-excel.ps1 -Files "a.xlsx","b.xlsx" -Out "D:\compare"
#
# Excel opens each file read-only and exports the active sheet (or every
# sheet, with -AllSheets) as PDF — Excel's own layout of the workbook, the
# thing our grid is measured against. Rasterise the pages with
# tools/pdf-pages.ps1 and capture ours with:
#
#   RUTBA_SMOKE_FILE="a.xlsx" RUTBA_SMOKE_OUT=<dir> node apps/desktop/build/smoke-run.js sheets
#
# One Excel per file in a job with a time limit, and an unblocked copy of the
# file, for the same reasons compare-word.ps1 does it: Protected View wedges
# automation silently, and a workbook Excel cannot finish must cost minutes,
# not the afternoon. Nothing here changes the workbooks.

param(
  [Parameter(Mandatory = $true)] [string[]] $Files,
  [Parameter(Mandatory = $true)] [string] $Out,
  [switch] $AllSheets
)

New-Item -ItemType Directory -Force -Path $Out | Out-Null

$exportOne = {
  param($copy, $pdf, $allSheets)
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.AutomationSecurity = 3   # msoAutomationSecurityForceDisable: never run a macro
  try {
    # Open(Filename, UpdateLinks=0, ReadOnly=true)
    $book = $excel.Workbooks.Open($copy, 0, $true)
    if ($allSheets) {
      # 0 = xlTypePDF; the whole workbook, every sheet in order.
      $book.ExportAsFixedFormat(0, $pdf)
    } else {
      $book.ActiveSheet.ExportAsFixedFormat(0, $pdf)
    }
    $sheets = $book.Worksheets.Count
    $book.Close($false)
    "ok    {0,3} sheets  {1}" -f $sheets, $pdf
  } catch {
    "FAIL  {0}  {1}" -f $copy, $_.Exception.Message
  } finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  }
}

foreach ($file in $Files) {
  $full = (Resolve-Path -LiteralPath $file).Path
  $name = [System.IO.Path]::GetFileNameWithoutExtension($full)
  $pdf = Join-Path $Out ($name + '.excel.pdf')
  $copyDir = Join-Path $Out 'src'
  New-Item -ItemType Directory -Force -Path $copyDir | Out-Null
  $copy = Join-Path $copyDir ([System.IO.Path]::GetFileName($full))
  Copy-Item -LiteralPath $full -Destination $copy -Force
  Unblock-File -LiteralPath $copy -ErrorAction SilentlyContinue

  $job = Start-Job -ScriptBlock $exportOne -ArgumentList $copy, $pdf, [bool]$AllSheets
  if (Wait-Job $job -Timeout 240) {
    Receive-Job $job
  } else {
    Stop-Job $job
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
    Write-Output ("STUCK {0}  Excel did not finish in 240 s" -f $full)
  }
  Remove-Job $job -Force
}
