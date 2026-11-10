// Held to PowerPoint itself. Builds a deck in the real program, makes the same
// edits there and here, and has PowerPoint read both. The fixture and the
// answers are committed, so tests/powerpoint-edits.test.js runs anywhere;
// running this needs a PC with PowerPoint.
//
//   node tools/powerpoint-check.mjs            build the fixture, record the answers
//   node tools/powerpoint-check.mjs --verify   and have PowerPoint open our results
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { Deck } from '@rutba/presentation';

const TABLE = 2;
export const COLUMN_EDITS = [
  { name: 'column inserted through a merged cell', op: 'insert', at: 2 },
  { name: 'column inserted before a merged cell', op: 'insert', at: 1 },
  { name: 'column deleted that holds a merged cell\'s words', op: 'delete', at: 1 },
  { name: 'column deleted that a merged cell covers', op: 'delete', at: 2 },
];
export const ROW_EDITS = [
  { name: 'row inserted through a merged cell', op: 'insert', at: 2 },
  { name: 'row deleted that holds a merged cell\'s words', op: 'delete', at: 1 },
  { name: 'row deleted that a merged cell covers', op: 'delete', at: 2 },
];

const tableOf = (deck) => deck.slide(1).shapes.find((s) => s.kind === 'table' || s.table);

/** The deck after one of the edits, made here. */
export function applyTableEdit(bytes, axis, edit) {
  const deck = Deck.open(bytes);
  const id = tableOf(deck).id;
  if (axis === 'col') {
    if (edit.op === 'insert') deck.insertTableColumn(1, id, edit.at);
    else deck.removeTableColumn(1, id, edit.at);
  } else if (edit.op === 'insert') deck.insertTableRow(1, id, edit.at);
  else deck.removeTableRow(1, id, edit.at);
  return deck.save();
}

/** The merges of a table, as `<a:tc>` attributes say them, row by row. */
export function spansOf(bytes) {
  const deck = Deck.open(bytes);
  const xml = deck.pkg.text(deck.slideParts[1].part);
  return [...xml.matchAll(/<a:tr\b[\s\S]*?<\/a:tr>/g)].map(([row]) => [...row.matchAll(/<a:tc\b([^>]*)>([\s\S]*?)<\/a:tc>/g)].map((m) => {
    const n = (name) => Number(new RegExp(`\\b${name}="(\\d+)"`).exec(m[1])?.[1] ?? 0);
    const words = [...m[2].matchAll(/<a:t(?:\s[^>]*)?>([^<]*)<\/a:t>/g)].map((x) => x[1]).join('');
    return [n('gridSpan'), n('rowSpan'), n('hMerge'), n('vMerge'), words].join('|');
  }));
}

/** Deck edits only we can be asked for; PowerPoint is asked whether it can open them. */
export function ourEdits(bytes) {
  const out = {};
  const edit = (name, fn) => { const deck = Deck.open(bytes); fn(deck); out[name] = deck.save(); };
  edit('frame moved', (d) => d.setGeometry(1, tableOf(d).id, { x: 100, y: 120, w: 500, h: 200 }));
  edit('slide duplicated', (d) => { const i = d.duplicateSlide(0); d.setNotes(i, 'Notes of the copy'); });
  edit('slide deleted', (d) => d.removeSlide(2));
  edit('first slide deleted', (d) => d.removeSlide(0));
  edit('words edited', (d) => {
    const s = d.slide(0).shapes.find((x) => x.text?.paragraphs?.length > 1);
    d.setText(0, s.id, s.text.paragraphs.map((p) => ({ ...p, runs: p.runs.map((r) => (r.field || r.break ? r : { ...r, text: r.text.replace('First', 'Changed') })) })));
  });
  edit('slide added after a slide', (d) => d.insertSlide(1, { title: 'New' }));
  return out;
}

const DESCRIBE = `
function Describe($p) {
  $o = @{ slides = @() }
  foreach ($s in $p.Slides) {
    $notes = ''
    if ($s.HasNotesPage) { foreach ($sh in $s.NotesPage.Shapes) { if ($sh.Type -eq 14 -and $sh.PlaceholderFormat.Type -eq 2) { $notes = $sh.TextFrame.TextRange.Text } } }
    $o.slides += ,@{ layout = [string]$s.CustomLayout.Name; notes = $notes; shapes = $s.Shapes.Count }
  }
  return $o
}
function Grid($t) {
  $rows = @()
  for ($r = 1; $r -le $t.Rows.Count; $r++) {
    $line = @()
    for ($c = 1; $c -le $t.Columns.Count; $c++) { $line += [string]$t.Cell($r, $c).Merged + '|' + $t.Cell($r, $c).Shape.TextFrame.TextRange.Text }
    $rows += ,$line
  }
  return ,$rows
}
function TableOf($s) { foreach ($sh in $s.Shapes) { if ($sh.HasTable) { return $sh } } }
`;

const BUILD = `
$p = $pp.Presentations.Add($false)
$s1 = $p.Slides.Add(1, 2)
$s1.Shapes.Item(1).TextFrame.TextRange.Text = 'Alpha title'
$b = $s1.Shapes.Item(2).TextFrame.TextRange
$b.Text = "First line of words\`rاردو کا جملہ\`rThird "
$b.Paragraphs(1).Font.Color.ObjectThemeColor = 5
$b.Paragraphs(2).Font.NameComplexScript = 'Noto Nastaliq Urdu'
$b.InsertAfter('0').InsertSlideNumber() | Out-Null
$s1.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = 'Notes of the first slide'
$s2 = $p.Slides.Add(2, 12)
$sh = $s2.Shapes.AddTable(4, 4, 40, 40, 480, 200)
$t = $sh.Table
for ($r = 1; $r -le 4; $r++) { for ($c = 1; $c -le 4; $c++) { $t.Cell($r, $c).Shape.TextFrame.TextRange.Text = "R$r" + "C$c" } }
$t.Cell(1, 2).Merge($t.Cell(1, 3))
$t.Cell(2, 1).Merge($t.Cell(3, 1))
$s3 = $p.Slides.Add(3, 1)
$s3.Shapes.Item(1).TextFrame.TextRange.Text = 'Secret plan'
$s3.Shapes.Item(2).TextFrame.TextRange.Text = 'Words nobody should keep'
$s3.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = 'Secret notes'
$link = $s1.Shapes.Item(1).ActionSettings.Item(1)
$link.Hyperlink.SubAddress = [string]$s3.SlideID + ',3,Secret plan'
$p.RemoveDocumentInformation(99)
$p.SaveAs('$DIR\\built.pptx', 24)
$p.Close()
`;

function runPowerPoint(body, needs) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-powerpoint-'));
  for (const [name, content] of Object.entries(needs ?? {})) fs.writeFileSync(path.join(dir, name), content);
  const script = path.join(dir, 'ask.ps1');
  fs.writeFileSync(script, `${String.fromCharCode(0xfeff)}$ErrorActionPreference = 'Stop'
$pp = New-Object -ComObject PowerPoint.Application
${DESCRIBE}
try {
${body.replaceAll('$DIR', dir)}
} finally { $pp.Quit() }
`);
  const r = spawnSync('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', script], { encoding: 'utf8', maxBuffer: 1 << 26, timeout: 300000 });
  if (r.status !== 0) throw new Error(`PowerPoint did not answer: ${r.stderr || r.stdout || r.error}`);
  return { dir, stdout: r.stdout };
}

const open = (file) => `$pp.Presentations.Open('${file}', $true, $false, $false)`;

function main() {
  const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
  const fixture = path.join(root, 'tests/fixtures/powerpoint-edits.pptx');
  const answers = path.join(root, 'tests/fixtures/powerpoint-edits.json');
  const built = runPowerPoint(BUILD);
  fs.copyFileSync(path.join(built.dir, 'built.pptx'), fixture);

  const editsThere = [...COLUMN_EDITS.map((e) => ['col', e]), ...ROW_EDITS.map((e) => ['row', e])];
  const editScript = editsThere.map(([axis, e], i) => `$p = ${open(fixture.replaceAll('/', '\\'))}
$p2 = $p
$sh = TableOf $p.Slides.Item(2)
$t = $sh.Table
${axis === 'col'
    ? (e.op === 'insert' ? `$t.Columns.Add(${e.at + 1}) | Out-Null` : `$t.Columns.Item(${e.at + 1}).Delete()`)
    : (e.op === 'insert' ? `$t.Rows.Add(${e.at + 1}) | Out-Null` : `$t.Rows.Item(${e.at + 1}).Delete()`)}
$p.SaveCopyAs('$DIR\\edit${i}.pptx')
$p.Close()`).join('\n');
  const edited = runPowerPoint(editScript);
  const out = { base: spansOf(fs.readFileSync(fixture)), edits: editsThere.map(([axis, e], i) => ({ axis, name: e.name, powerpoint: spansOf(fs.readFileSync(path.join(edited.dir, `edit${i}.pptx`))) })) };
  fs.writeFileSync(answers, `${JSON.stringify(out, null, 1)}\n`);
  console.log(`powerpoint-check: ${out.edits.length} table edits answered`);

  if (process.argv.includes('--verify')) {
    const bytes = fs.readFileSync(fixture);
    const results = { ...ourEdits(bytes) };
    editsThere.forEach(([axis, e]) => { results[`table: ${e.name}`] = applyTableEdit(bytes, axis, e); });
    const names = Object.keys(results);
    const needs = Object.fromEntries(names.map((n, i) => [`ours${i}.pptx`, results[n]]));
    const check = names.map((n, i) => `try { $p = ${open('$DIR\\ours' + i + '.pptx')}; $d = Describe $p; '${n.replace(/'/g, "''")}: OK ' + ($d | ConvertTo-Json -Depth 5 -Compress); $p.Close() } catch { '${n.replace(/'/g, "''")}: REFUSED ' + $_.Exception.Message }`).join('\n');
    console.log(runPowerPoint(check, needs).stdout);
  }
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) main();
