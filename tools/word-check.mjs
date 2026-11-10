// Held to Word itself. Builds a document in the real program, makes the same
// edits there and here, and has Word open what we write. The fixture and the
// answers are committed, so tests/word-edits.test.js runs anywhere; running
// this needs a PC with Word.
//
//   node tools/word-check.mjs            build the fixture, record the answers
//   node tools/word-check.mjs --verify   and have Word open our results
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { OoxmlPackage } from '@rutba/ooxml/package';

const unesc = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

/** What a run holds, in the order it holds it. */
function tokensOf(inner) {
  const re = /<w:(?:t|delText)(?:\s[^>]*)?>([^<]*)<\/w:(?:t|delText)>|<w:tab\/>|<w:noBreakHyphen\/>|<w:softHyphen\/>|<w:br\b[^>]*\/>|<w:sym\b[^>]*\/>|<w:(footnote|endnote)Reference\b[^>]*\/>/g;
  let out = '';
  for (const m of inner.matchAll(re)) {
    const tag = m[0];
    if (/^<w:(?:t|delText)[\s>]/.test(tag)) out += unesc(m[1]);
    else if (tag === '<w:tab/>') out += '{tab}';
    else if (tag === '<w:noBreakHyphen/>') out += '{nbh}';
    else if (tag === '<w:softHyphen/>') out += '{shy}';
    else if (tag.startsWith('<w:br')) out += /w:type="page"/.test(tag) ? '{page}' : /w:type="column"/.test(tag) ? '{col}' : '{br}';
    else if (tag.startsWith('<w:sym')) out += `{sym:${/w:font="([^"]*)"/.exec(tag)[1]}:${/w:char="([^"]*)"/.exec(tag)[1].toUpperCase()}}`;
    else out += '{note}';
  }
  return out;
}

/** A paragraph as a line: its words and the things among them, tracked marks shown. */
function lineOf(xml) {
  const noBoxes = xml.replace(/<mc:AlternateContent\b[\s\S]*?<\/mc:AlternateContent>/g, '{box}').replace(/<w:pict\b[\s\S]*?<\/w:pict>/g, '{box}');
  const pPr = /<w:pPr\b[^>]*>(?:(?!<\/w:pPr>)[\s\S])*<\/w:pPr>/.exec(noBoxes)?.[0] ?? '';
  const body = noBoxes.replace(pPr, '');
  const parts = [];
  const re = /<w:(ins|del|moveFrom|moveTo)\b(?![^>]*\/>)[^>]*>([\s\S]*?)<\/w:\1>|<w:r\b(?![a-zA-Z])[^>]*>([\s\S]*?)<\/w:r>|\{box\}/g;
  for (const m of body.matchAll(re)) {
    if (m[0] === '{box}') parts.push('{box}');
    else if (m[1]) {
      const words = [...m[2].matchAll(/<w:r\b(?![a-zA-Z])[^>]*>([\s\S]*?)<\/w:r>/g)].map((r) => tokensOf(r[1])).join('');
      if (words) parts.push(`[${m[1]}:${words}]`);
    } else parts.push(tokensOf(m[3]));
  }
  const style = /<w:pStyle\b[^>]*w:val="([^"]*)"/.exec(pPr)?.[1];
  const marks = [...pPr.matchAll(/<w:rPr\b[\s\S]*?<\/w:rPr>/g)].flatMap((r) => [...r[0].matchAll(/<w:(ins|del|moveFrom|moveTo)\b[^>]*\/>/g)].map((x) => x[1]));
  return [style ? `(${style})` : '', parts.join('').replace(/\]\[(ins|del|moveFrom|moveTo):/g, (all, k) => all), /<w:sectPr\b/.test(pPr) ? '{section}' : '', marks.length ? `{mark:${marks.join()}}` : ''].join('');
}

const mergeSame = (line) => line.replace(/\[(ins|del|moveFrom|moveTo):([^\]]*)\]\[\1:/g, '[$1:$2');

/** A document's body as lines, table cells included, a text box's words left out. */
export function shapeOf(bytes) {
  const xml = OoxmlPackage.read(bytes).text('word/document.xml');
  const body = /<w:body\b[^>]*>([\s\S]*)<\/w:body>/.exec(xml)[1]
    .replace(/<w:txbxContent\b[\s\S]*?<\/w:txbxContent>/g, '');
  return [...body.matchAll(/<w:p\b(?![a-zA-Z])[^>]*\/>|<w:p\b(?![a-zA-Z])[^>]*>[\s\S]*?<\/w:p>/g)].map((m) => mergeSame(lineOf(m[0])));
}

/** Every edit, as the same gesture in both programs. */
export const EDITS = [
  { name: 'typed at the start of a paragraph that opens with a page break', find: 'After the break', at: 'start', type: 'X' },
  { name: 'typed inside words after a page break', find: 'After the break', at: 5, type: 'X' },
  { name: 'Enter just before a footnote reference', find: 'Words before', at: 'wordsEnd', enter: true },
  { name: 'Enter just after a footnote reference', find: ' and after', at: 'phraseStart', enter: true },
  { name: 'typed in a paragraph with a footnote', find: ' and after', at: 'phraseEnd', type: 'Z' },
  { name: 'typed in words with a no-break hyphen', find: 'Well', at: 2, type: 'X' },
  { name: 'typed beside a symbol', find: 'Symbol here', at: 'phraseEnd', type: 'Q' },
  { name: 'Enter just before a symbol', find: 'Symbol here', at: 'phraseEnd', enter: true },
  { name: 'Enter in a paragraph that ends a section', find: 'Last of section', at: 5, enter: true },
  { name: 'Delete at the end of a paragraph that ends a section', find: 'Last of section', at: 'paraEnd', del: true },
  { name: 'Delete at the end of the paragraph before a section\'s last', find: 'Second last', at: 'paraEnd', del: true },
  { name: 'typed in the paragraph with a text box', find: 'Box anchor', at: 3, type: 'X' },
  { name: 'Enter in the words after a page break', find: 'After the break', at: 5, enter: true },
];

export const TRACKED = [
  { name: 'Accept All', accept: true },
  { name: 'Reject All', accept: false },
];

/** The same edit, here. */
export function applyEdit(bytes, edit) {
  const view = openDocx(bytes);
  const blocks = view.blocks;
  const block = blocks.findIndex((b) => b.text.includes(edit.find));
  if (block < 0) throw new Error(`no paragraph has "${edit.find}"`);
  const text = blocks[block].text;
  const found = text.indexOf(edit.find);
  const offset = edit.at === 'start' ? 0
    : edit.at === 'wordsEnd' || edit.at === 'phraseEnd' ? found + edit.find.length
      : edit.at === 'phraseStart' ? found
        : edit.at === 'paraEnd' ? text.length
          : found + edit.at;
  view.collapseTo({ block, offset });
  if (edit.type) view.insertText(edit.type);
  else if (edit.enter) view.splitParagraph();
  else if (edit.del) view.deleteForward();
  return view.save();
}

export function applyTracked(bytes, { accept }) {
  const view = openDocx(bytes);
  if (accept) view.acceptChanges({ all: true }); else view.rejectChanges({ all: true });
  return view.save();
}

/** A page break and the words after it in one paragraph, as older files have them: Word now writes the break a paragraph of its own. */
export function withLeadingBreak(bytes) {
  const pkg = OoxmlPackage.read(bytes);
  const part = 'word/document.xml';
  const xml = pkg.text(part);
  const close = xml.indexOf('</w:p>', xml.indexOf('<w:body'));
  const next = /^<w:p\b[^>]*>(?:<w:pPr\b[\s\S]*?<\/w:pPr>)?/.exec(xml.slice(close + 6));
  pkg.write_(part, xml.slice(0, close) + xml.slice(close + 6 + next[0].length));
  return pkg.write();
}

/** A tracked move, as Word writes one, put into a document that has none: Word records it only from the clipboard. */
export function withMove(bytes) {
  const pkg = OoxmlPackage.read(bytes);
  const part = 'word/document.xml';
  const attrs = 'w:author="Reviewer" w:date="2026-01-02T03:04:00Z"';
  const moved = '<w:p><w:pPr><w:rPr><w:moveFrom w:id="901" ' + attrs + '/></w:rPr></w:pPr>'
    + '<w:moveFromRangeStart w:id="902" ' + attrs + ' w:name="move1"/>'
    + '<w:moveFrom w:id="903" ' + attrs + '><w:r><w:t>Moved line</w:t></w:r></w:moveFrom>'
    + '<w:moveFromRangeEnd w:id="902"/></w:p>';
  const landed = '<w:p><w:pPr><w:rPr><w:moveTo w:id="904" ' + attrs + '/></w:rPr></w:pPr>'
    + '<w:moveToRangeStart w:id="905" ' + attrs + ' w:name="move1"/>'
    + '<w:moveTo w:id="906" ' + attrs + '><w:r><w:t>Moved line</w:t></w:r></w:moveTo>'
    + '<w:moveToRangeEnd w:id="905"/></w:p>';
  const xml = pkg.text(part);
  const at = xml.indexOf('<w:body') + xml.slice(xml.indexOf('<w:body')).search(/<w:p[ >]/);
  const end = xml.lastIndexOf('<w:sectPr');
  pkg.write_(part, xml.slice(0, at) + moved + xml.slice(at, end) + landed + xml.slice(end));
  return pkg.write();
}

const DESCRIBE = `
function Describe($d) {
  $o = @{ paragraphs = @(); sections = $d.Sections.Count; notes = $d.Footnotes.Count; shapes = $d.Shapes.Count; revisions = $d.Revisions.Count }
  foreach ($p in $d.Paragraphs) { $o.paragraphs += ,[string]$p.Range.Text }
  return $o
}
function Place($d, $edit) {
  $target = $null
  foreach ($p in $d.Paragraphs) { if ($p.Range.Text.Contains($edit.find)) { $target = $p; break } }
  if (-not $target) { throw ('no paragraph has ' + $edit.find) }
  $text = $target.Range.Text
  $found = $text.IndexOf($edit.find)
  $start = $target.Range.Start
  switch ($edit.at) {
    'start' { $pos = $start + $(if ($text[0] -eq [char]12) { 1 } else { 0 }) }
    'wordsEnd' { $pos = $start + $found + $edit.find.Length }
    'phraseEnd' { $pos = $start + $found + $edit.find.Length }
    'phraseStart' { $pos = $start + $found }
    'paraEnd' { $pos = $target.Range.End - 1 }
    default { $pos = $start + $found + [int]$edit.at }
  }
  $d.Range([ref]$pos, [ref]$pos).Select()
}
`;

const BUILD = `
$d = $w.Documents.Add()
$s = $w.Selection
$s.InsertBreak([ref]7)
$s.TypeText('After the break line')
$s.TypeParagraph()
$s.TypeText('Words before')
$n = $d.Footnotes.Add($s.Range)
$n.Range.Text = 'The note'
$s.EndKey([ref]6) | Out-Null
$s.TypeText(' and after')
$s.TypeParagraph()
$s.TypeText('Well' + [string][char]30 + 'known words')
$s.TypeParagraph()
$s.TypeText('Symbol here')
$s.InsertSymbol(-3872, [ref]'Wingdings', [ref]$true)
$s.TypeText(' tail')
$s.TypeParagraph()
$s.TypeText('Second last of it')
$s.TypeParagraph()
$s.TypeText('Last of section one')
$s.InsertBreak([ref]2)
$s.TypeText('Box anchor words')
$r = $s.Range
$box = $d.Shapes.AddTextbox(1, 200, 200, 120, 60, [ref]$r)
$box.TextFrame.TextRange.Text = 'Words in the box'
$s.EndKey([ref]6) | Out-Null
$s.TypeParagraph()
$s.TypeText('Chapter one')
$s.Style = $d.Styles.Item([ref](-2))
$s.TypeParagraph()
$s.Style = $d.Styles.Item([ref](-1))
$s.TypeText('Body under it')
$s.TypeParagraph()
$s.TypeText('Chapter two')
$s.Style = $d.Styles.Item([ref](-3))
$s.TypeParagraph()
$s.Style = $d.Styles.Item([ref](-1))
$s.TypeText('Closing words')
$d.Sections.Item(1).PageSetup.Orientation = 1
$d.RemoveDocumentInformation(99)
$d.SaveAs2([ref]'$DIR\\built.docx', [ref]16)
$d.Close([ref]0)

$d = $w.Documents.Add()
$s = $w.Selection
$s.TypeText('Alpha' + [string][char]9 + 'Beta Gamma')
$s.TypeParagraph()
$s.TypeText('Keep this line')
$s.TypeParagraph()
$s.TypeText('Second' + [string][char]11 + 'line of it')
$s.TypeParagraph()
$s.TypeText('Well' + [string][char]30 + 'known end')
$d.TrackRevisions = $true
$p = $d.Paragraphs.Item(1).Range
$d.Range([ref]($p.Start + 3), [ref]($p.End - 3)).Delete() | Out-Null
$p = $d.Paragraphs.Item(3).Range
$d.Range([ref]($p.Start + 2), [ref]($p.End - 3)).Delete() | Out-Null
$p = $d.Paragraphs.Item(4).Range
$d.Range([ref]($p.Start), [ref]($p.End - 5)).Delete() | Out-Null
$at = $d.Paragraphs.Item(2).Range.Start + 4
$d.Range([ref]$at, [ref]$at).InsertAfter('NEW')
$d.TrackRevisions = $false
$d.RemoveDocumentInformation(4)
$d.RemoveDocumentInformation(8)
$d.SaveAs2([ref]'$DIR\\tracked.docx', [ref]16)
$d.Close([ref]0)
`;

function runWord(body, needs) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-word-'));
  for (const [name, content] of Object.entries(needs ?? {})) fs.writeFileSync(path.join(dir, name), content);
  const script = path.join(dir, 'ask.ps1');
  fs.writeFileSync(script, `${String.fromCharCode(0xfeff)}$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
${DESCRIBE}
try {
${body.replaceAll('$DIR', dir)}
} finally { $w.Quit([ref]0) }
`);
  const r = spawnSync('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', script], { encoding: 'utf8', maxBuffer: 1 << 26, timeout: 300000 });
  if (r.status !== 0) throw new Error(`Word did not answer: ${r.stderr || r.stdout || r.error}`);
  return { dir, stdout: r.stdout };
}

const open = (file) => `$w.Documents.Open([ref]'${file}', [ref]$false, [ref]$false, [ref]$false)`;
const quote = (s) => `'${String(s).replace(/'/g, "''")}'`;

function main() {
  const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
  const fixture = path.join(root, 'tests/fixtures/word-edits.docx');
  const trackedFixture = path.join(root, 'tests/fixtures/word-tracked.docx');
  const answers = path.join(root, 'tests/fixtures/word-edits.json');
  const built = runWord(BUILD);
  fs.writeFileSync(fixture, withLeadingBreak(fs.readFileSync(path.join(built.dir, 'built.docx'))));
  fs.writeFileSync(trackedFixture, withMove(fs.readFileSync(path.join(built.dir, 'tracked.docx'))));
  const win = (p) => p.replaceAll('/', '\\');

  const editScript = EDITS.map((e, i) => `$d = ${open(win(fixture))}
$edit = @{ find = ${quote(e.find)}; at = ${quote(e.at)} }
Place $d $edit
${e.type ? `$w.Selection.TypeText(${quote(e.type)})` : e.enter ? '$w.Selection.TypeParagraph()' : '$w.Selection.Delete([ref]1, [ref]1) | Out-Null'}
$d.SaveAs2([ref]'$DIR\\edit${i}.docx', [ref]16)
$d.Close([ref]0)`).join('\n');
  const trackedScript = TRACKED.map((e, i) => `$d = ${open(win(trackedFixture))}
${e.accept ? '$d.Revisions.AcceptAll()' : '$d.Revisions.RejectAll()'}
$d.SaveAs2([ref]'$DIR\\tracked${i}.docx', [ref]16)
$d.Close([ref]0)`).join('\n');
  const done = runWord(`${editScript}\n${trackedScript}`);
  const read = (name) => shapeOf(fs.readFileSync(path.join(done.dir, name)));
  const out = {
    base: shapeOf(fs.readFileSync(fixture)),
    edits: EDITS.map((e, i) => ({ name: e.name, word: read(`edit${i}.docx`) })),
    tracked: TRACKED.map((e, i) => ({ name: e.name, word: read(`tracked${i}.docx`) })),
  };
  fs.writeFileSync(answers, `${JSON.stringify(out, null, 1)}\n`);
  console.log(`word-check: ${out.edits.length + out.tracked.length} edits answered`);

  if (process.argv.includes('--verify')) {
    const bytes = fs.readFileSync(fixture);
    const tracked = fs.readFileSync(trackedFixture);
    const results = {};
    for (const e of EDITS) results[e.name] = applyEdit(bytes, e);
    for (const e of TRACKED) results[e.name] = applyTracked(tracked, e);
    const names = Object.keys(results);
    const needs = Object.fromEntries(names.map((n, i) => [`ours${i}.docx`, results[n]]));
    const check = names.map((n, i) => `try { $d = ${open(`$DIR\\ours${i}.docx`)}; $o = Describe $d; ${quote(`${n}: OK `)} + ($o.sections) + ' sections, ' + ($o.notes) + ' notes, ' + ($o.paragraphs.Count) + ' paragraphs, ' + ($o.revisions) + ' revisions'; $d.Close([ref]0) } catch { ${quote(`${n}: REFUSED `)} + $_.Exception.Message }`).join('\n');
    console.log(runWord(check, needs).stdout);
  }
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) main();
