# One Word 97-2003 fixture, made by Word: the characters, the paragraphs and
# lists, or the structure. Run by make-binary-fixtures.ps1, once per kind, in
# a PowerShell of its own with a time limit.
#
# Written straight through, at the top level, on purpose: the same steps
# inside a PowerShell function or script block leave Word's Save As waiting
# for ever in an automated session on this machine, and at the top level
# they save.
param([string]$Out, [string]$Kind, [string]$Picture)

$ErrorActionPreference = 'Stop'
function RGB($r, $g, $b) { return $r + ($g * 256) + ($b * 65536) }

$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
# One at a time: an add-in installed for all users cannot be disconnected
# without an administrator, and must not keep the rest connected.
foreach ($a in $w.COMAddIns) { try { if ($a.Connect) { $a.Connect = $false } } catch {} }
$d = $w.Documents.Add()
$s = $w.Selection

if ($Kind -eq 'text') {
  $s.Style = 'Heading 1'; $s.TypeText('Characters'); $s.TypeParagraph(); $s.Style = 'Normal'
  $s.TypeText('Plain, ')
  $s.Font.Bold = $true; $s.TypeText('bold'); $s.Font.Bold = $false; $s.TypeText(', ')
  $s.Font.Italic = $true; $s.TypeText('italic'); $s.Font.Italic = $false; $s.TypeText(', ')
  $s.Font.Underline = 1; $s.TypeText('underlined'); $s.Font.Underline = 0; $s.TypeText(', ')
  $s.Font.StrikeThrough = $true; $s.TypeText('struck'); $s.Font.StrikeThrough = $false; $s.TypeText(', ')
  $s.Font.Color = RGB 192 0 0; $s.TypeText('red'); $s.Font.Color = -16777216; $s.TypeText(', ')
  $s.Font.Size = 16; $s.TypeText('sixteen point'); $s.Font.Size = 11; $s.TypeText(', ')
  $s.Font.Name = 'Courier New'; $s.TypeText('Courier New'); $s.Font.Name = 'Calibri'; $s.TypeText(', ')
  $s.Font.Superscript = $true; $s.TypeText('raised'); $s.Font.Superscript = $false; $s.TypeText(' and ')
  $s.Font.Subscript = $true; $s.TypeText('lowered'); $s.Font.Subscript = $false; $s.TypeText(', ')
  $s.Range.HighlightColorIndex = 7; $s.TypeText('highlighted'); $s.Range.HighlightColorIndex = 0; $s.TypeText(', ')
  $s.Font.SmallCaps = $true; $s.TypeText('small capitals'); $s.Font.SmallCaps = $false
  $s.TypeText('.'); $s.TypeParagraph()
  # Other alphabets, built from code points so the script's own encoding cannot garble them.
  $other = 'Accents and other alphabets: caf' + [char]0xE9 + ', na' + [char]0xEF + 've, Z' + [char]0xFC + 'rich, ' + (-join ([char[]](0x395, 0x3BB, 0x3BB, 0x3B7, 0x3BD, 0x3B9, 0x3BA, 0x3AC))) + ', ' + (-join ([char[]](0x420, 0x443, 0x441, 0x441, 0x43A, 0x438, 0x439))) + '.'
  $s.TypeText($other); $s.TypeParagraph()
}

if ($Kind -eq 'paragraphs') {
  $s.Style = 'Heading 1'; $s.TypeText('Paragraphs'); $s.TypeParagraph(); $s.Style = 'Normal'
  $s.ParagraphFormat.Alignment = 1; $s.TypeText('This paragraph is centred.'); $s.TypeParagraph()
  $s.ParagraphFormat.Alignment = 2; $s.TypeText('This paragraph is aligned right.'); $s.TypeParagraph()
  $s.ParagraphFormat.Alignment = 3; $s.TypeText('This paragraph is justified, and long enough to wrap onto a second line so that the justification shows: the words spread to meet both margins on every line but the last one.'); $s.TypeParagraph()
  $s.ParagraphFormat.Alignment = 0
  $s.ParagraphFormat.LeftIndent = 36; $s.TypeText('Indented half an inch on the left.'); $s.TypeParagraph(); $s.ParagraphFormat.LeftIndent = 0
  $s.ParagraphFormat.FirstLineIndent = 36; $s.TypeText('A first line indented half an inch, the rest of the paragraph at the margin, which shows once the paragraph is long enough to wrap.'); $s.TypeParagraph(); $s.ParagraphFormat.FirstLineIndent = 0
  $s.ParagraphFormat.SpaceBefore = 18; $s.ParagraphFormat.SpaceAfter = 18; $s.TypeText('Eighteen points before and after.'); $s.TypeParagraph()
  $s.ParagraphFormat.SpaceBefore = 0; $s.ParagraphFormat.SpaceAfter = 8
  $s.TypeText('A line'); $s.InsertBreak(6); $s.TypeText("broken by a line break, and a tab:`tafter the tab."); $s.TypeParagraph()
  $s.Style = 'Heading 2'; $s.TypeText('Lists'); $s.TypeParagraph(); $s.Style = 'Normal'
  $s.Range.ListFormat.ApplyBulletDefault()
  $s.TypeText('First bullet'); $s.TypeParagraph(); $s.TypeText('Second bullet'); $s.TypeParagraph()
  $s.Range.ListFormat.RemoveNumbers()
  $s.Range.ListFormat.ApplyNumberDefault()
  $s.TypeText('First number'); $s.TypeParagraph(); $s.TypeText('Second number'); $s.TypeParagraph(); $s.TypeText('Third number'); $s.TypeParagraph()
  $s.Range.ListFormat.RemoveNumbers()
  $s.TypeText('After the lists.')
}

if ($Kind -eq 'structure') {
  $s.Style = 'Title'; $s.TypeText('Structure'); $s.TypeParagraph()
  $s.Style = 'Heading 2'; $s.TypeText('A table'); $s.TypeParagraph(); $s.Style = 'Normal'
  $t = $d.Tables.Add($s.Range, 3, 3)
  $t.Borders.Enable = $true
  $cells = @(@('Region', 'Q1', 'Q2'), @('North', '120', '135'), @('South', '98', '110'))
  for ($r = 1; $r -le 3; $r++) { for ($c = 1; $c -le 3; $c++) { $t.Cell($r, $c).Range.Text = $cells[$r - 1][$c - 1] } }
  $t.Rows.Item(1).Range.Font.Bold = $true
  $s.EndKey(6) | Out-Null
  $s.TypeParagraph()
  $s.Style = 'Heading 2'; $s.TypeText('A link, a footnote and a picture'); $s.TypeParagraph(); $s.Style = 'Normal'
  $s.TypeText('A link to ')
  [void]$d.Hyperlinks.Add($s.Range, 'https://example.com/', '', '', 'example.com')
  $s.TypeText(', a footnote here')
  [void]$d.Footnotes.Add($s.Range, [Type]::Missing, 'The footnote''s own words.')
  $s.TypeText(', and a picture: ')
  [void]$s.InlineShapes.AddPicture($Picture)
  $s.TypeParagraph()
  $d.Sections.Item(1).Headers.Item(1).Range.Text = 'Structure header'
  $f = $d.Sections.Item(1).Footers.Item(1).Range
  $f.Text = 'Page '
  $f.Collapse(0)
  [void]$d.Fields.Add($f, 33)
  $s.InsertBreak(7)
  $s.Style = 'Heading 1'; $s.TypeText('Second page'); $s.TypeParagraph(); $s.Style = 'Normal'
  $s.TypeText('After a page break, the last paragraph.')
}

if ($Kind -eq 'floats') {
  $s.Style = 'Heading 1'; $s.TypeText('Floating drawings'); $s.TypeParagraph(); $s.Style = 'Normal'
  $s.TypeText(('A paragraph the drawings float beside, long enough to wrap round them. ' * 6)); $s.TypeParagraph()
  $anchor = $d.Paragraphs.Item(2).Range
  $pic = $d.Shapes.AddPicture($Picture, $false, $true, 300, 60, 120, 80, $anchor)
  $pic.WrapFormat.Type = 0
  $tb = $d.Shapes.AddTextbox(1, 72, 220, 200, 60, $anchor)
  $tb.TextFrame.TextRange.Text = 'Words in a text box'
  $tb.Fill.ForeColor.RGB = RGB 255 242 204
  $tb.Line.ForeColor.RGB = RGB 192 0 0
  $rect = $d.Shapes.AddShape(1, 300, 220, 100, 60, $anchor)
  $rect.Fill.ForeColor.RGB = RGB 68 114 196
  $oval = $d.Shapes.AddShape(9, 420, 220, 80, 60, $anchor)
  $oval.Fill.ForeColor.RGB = RGB 112 173 71
  $s.EndKey(6) | Out-Null
  $s.TypeParagraph(); $s.TypeText('After the drawings.')
}

# In Word 2003's mode first, so saving as 97-2003 has nothing to ask.
$d.SetCompatibilityMode(11)
if (Test-Path $Out) { Remove-Item $Out -Force }
$d.SaveAs2($Out, 0)
"$(Split-Path -Leaf $Out) written ($((Get-Item $Out).Length) bytes)"
$d.Close(0)
$w.Quit()
