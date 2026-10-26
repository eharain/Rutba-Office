# One Word document saved again as Strict Open XML (or Transitional), by
# Word. Run by make-strict-fixtures.ps1 and check-office-reopen.ps1 in a
# PowerShell of their own with a time limit.
#
# Written straight through, at the top level, on purpose: inside a function
# or script block Word's Save As has been seen to wait for ever in an
# automated session on this machine (make-binary-word.ps1).
param([string]$In, [string]$Out, [int]$Format = 24)

$ErrorActionPreference = 'Stop'
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
foreach ($a in $w.COMAddIns) { try { if ($a.Connect) { $a.Connect = $false } } catch {} }
$d = $w.Documents.Open($In, $false, $true, $false)
if (Test-Path $Out) { Remove-Item $Out -Force }
$d.SaveAs2($Out, $Format)
"$(Split-Path -Leaf $Out) written by Word ($((Get-Item $Out).Length) bytes)"
$d.Close(0)
$w.Quit()
