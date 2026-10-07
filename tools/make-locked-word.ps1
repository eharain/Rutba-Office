# One password-protected Word 97-2003 document, made at the top level of a
# PowerShell of its own as make-binary-word.ps1 makes its documents (Word's
# Save As hangs here when called from inside a function). Called by
# make-locked-fixtures.ps1; the password is "Rutba-1".
param([Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
if (Test-Path $Out) { Remove-Item $Out -Force }
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
$d = $w.Documents.Add()
$r = $d.Content
$r.Text = 'A locked document'
$r.InsertParagraphAfter()
$r.InsertAfter('Its words are kept from anyone without the password, and open with it.')
$r.InsertParagraphAfter()
$r.InsertAfter('A second paragraph, in bold.')
$d.Paragraphs.Item(1).Range.Font.Size = 20
$d.Paragraphs.Item(3).Range.Font.Bold = $true
$d.Password = 'Rutba-1'
$d.SaveAs2($Out, 0)
"locked.doc written ($((Get-Item $Out).Length) bytes)"
$d.Close(0)
$w.Quit()
