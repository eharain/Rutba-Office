# Microsoft Word's rendering of a document, as a PDF, for comparison with ours.
#
#   powershell -File tools/compare-word.ps1 -Files "a.docx","b.docx" -Out "D:\compare"
#
# Word is asked to open each file and export it as PDF, silently. That PDF is
# Word's own layout of the document — the thing our page is measured against.
# Both renderings are then photographed by the same smoke harness:
#
#   RUTBA_SMOKE_FILE="a.docx" npm run smoke -- word       # ours
#   RUTBA_SMOKE_FILE="a.pdf"  npm run smoke -- pictures   # Word's, through the PDF viewer
#
# Nothing here changes the documents. Word opens them read-only and is quit
# afterwards, whether or not the export worked.

param(
  [Parameter(Mandatory = $true)] [string[]] $Files,
  [Parameter(Mandatory = $true)] [string] $Out
)

New-Item -ItemType Directory -Force -Path $Out | Out-Null

# One document, one Word, one job — so a document Word cannot finish (a
# dialog it will not answer, a layout it never converges on) costs the run
# three minutes and a killed process, not the whole afternoon.
$exportOne = {
  param($copy, $pdf)
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  # 3 = msoAutomationSecurityForceDisable: never run a macro, never ask about one.
  $word.AutomationSecurity = 3
  try {
    # ReadOnly, no confirmations, no macros, no recovery prompts.
    $doc = $word.Documents.Open($copy, $false, $true, $false)
    # 17 = wdExportFormatPDF
    $doc.ExportAsFixedFormat($pdf, 17)
    $pages = $doc.ComputeStatistics(2)  # wdStatisticPages
    $doc.Close($false)
    "ok    {0,4} pages  {1}" -f $pages, $pdf
  } catch {
    "FAIL  {0}  {1}" -f $copy, $_.Exception.Message
  } finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
  }
}

foreach ($file in $Files) {
  $full = (Resolve-Path $file).Path
  $name = [System.IO.Path]::GetFileNameWithoutExtension($full)
  $pdf = Join-Path $Out ($name + '.word.pdf')
  # A file that came from the internet carries the mark of the web, and Word
  # opens it in Protected View — a window automation cannot see, so the call
  # never returns. Work from an unblocked copy; the original is not touched.
  $copyDir = Join-Path $Out 'src'
  New-Item -ItemType Directory -Force -Path $copyDir | Out-Null
  $copy = Join-Path $copyDir ([System.IO.Path]::GetFileName($full))
  Copy-Item -LiteralPath $full -Destination $copy -Force
  Unblock-File -LiteralPath $copy -ErrorAction SilentlyContinue

  $job = Start-Job -ScriptBlock $exportOne -ArgumentList $copy, $pdf
  if (Wait-Job $job -Timeout 180) {
    Receive-Job $job
  } else {
    Stop-Job $job
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
    Write-Output ("STUCK {0}  Word did not finish in 180 s" -f $full)
  }
  Remove-Job $job -Force
}
