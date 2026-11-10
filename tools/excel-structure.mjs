// Builds a workbook in Excel, has Excel insert and delete rows and columns in
// it, and writes down what each edit left: every formula, conditional format,
// validation, link, table, note and chart reference. The test
// (tests/excel-structure.test.js) makes the same edits here and compares.
//
//   node tools/excel-structure.mjs           needs Excel; writes the fixtures
//   node tools/excel-structure.mjs --verify  also has Excel open our results
//
// Nothing is saved but the fixtures, and those carry no name or machine.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { Workbook, OoxmlPackage, parseRef } from '@rutba/ooxml';

/** The edits, 0-based as ours take them; Excel is told the same places by address. */
export const EDITS = [
  { name: 'insert a row above the data', sheet: 'Data', axis: 'row', op: 'insert', at: 0, count: 1 },
  { name: 'insert a row inside a filled column', sheet: 'Data', axis: 'row', op: 'insert', at: 3, count: 1 },
  { name: 'insert two rows inside a filled column', sheet: 'Data', axis: 'row', op: 'insert', at: 2, count: 2 },
  { name: 'delete a row inside a filled column', sheet: 'Data', axis: 'row', op: 'delete', at: 3, count: 1 },
  { name: 'delete the first filled row', sheet: 'Data', axis: 'row', op: 'delete', at: 1, count: 1 },
  { name: 'insert a row inside the table', sheet: 'Data', axis: 'row', op: 'insert', at: 11, count: 1 },
  { name: 'delete a row inside the table', sheet: 'Data', axis: 'row', op: 'delete', at: 11, count: 1 },
  { name: 'insert a column before everything', sheet: 'Data', axis: 'col', op: 'insert', at: 0, count: 1 },
  { name: 'insert a column inside the table', sheet: 'Data', axis: 'col', op: 'insert', at: 1, count: 1 },
  { name: 'delete a column inside the table', sheet: 'Data', axis: 'col', op: 'delete', at: 1, count: 1 },
  { name: 'insert a column right of the table', sheet: 'Data', axis: 'col', op: 'insert', at: 7, count: 1 },
  { name: 'delete a column the formulas read', sheet: 'Data', axis: 'col', op: 'delete', at: 2, count: 1 },
  { name: 'insert a row on the other sheet', sheet: 'Other', axis: 'row', op: 'insert', at: 0, count: 1 },
];

const sorted = (list) => list.sort((a, b) => (JSON.stringify(a) < JSON.stringify(b) ? -1 : 1));
const unesc = (s) => String(s ?? '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
/** Excel may or may not write a cell's sheet where the formula is on that sheet. */
const bare = (f) => String(f ?? '').replace(/\s+/g, '');

/**
 * What a workbook says, whoever wrote it. Numbers rounded so that a value and
 * its fifteen-digit cache compare.
 */
export function describe(bytes) {
  const wb = Workbook.open(bytes);
  const out = { sheets: {}, names: {}, tables: [], charts: [] };
  for (const { name, part: partName } of wb.sheets()) {
    const { part } = wb._sheetPart(name);
    const tail = part.suffix;
    const cells = {};
    for (const c of wb.cells(name)) {
      const v = c.value;
      cells[c.ref] = typeof v === 'number' ? Number(v.toPrecision(12)) : typeof v === 'string' && v.startsWith('=') ? bare(v) : v;
    }
    const s = { cells };
    s.merges = sorted(part.mergeRefs().slice());
    s.conditional = sorted([...tail.matchAll(/<conditionalFormatting\b[^>]*?\bsqref="([^"]*)"[^>]*>([\s\S]*?)<\/conditionalFormatting>/g)]
      .map((m) => [m[1], [...m[2].matchAll(/<formula\b[^>]*>([\s\S]*?)<\/formula>/g)].map((x) => bare(unesc(x[1])))]));
    s.validations = sorted([...tail.matchAll(/<dataValidation\b([^>]*?)>([\s\S]*?)<\/dataValidation>/g)]
      .map((m) => [/\bsqref="([^"]*)"/.exec(m[1])?.[1], ...[...m[2].matchAll(/<formula[12]>([\s\S]*?)<\/formula[12]>/g)].map((x) => bare(unesc(x[1])))]));
    s.links = sorted([...tail.matchAll(/<hyperlink\b([^>]*?)\/>/g)]
      .map((m) => [/\bref="([^"]*)"/.exec(m[1])?.[1], unesc(/\blocation="([^"]*)"/.exec(m[1])?.[1] ?? '')]));
    s.breaks = sorted([...tail.matchAll(/<rowBreaks\b[^>]*>([\s\S]*?)<\/rowBreaks>/g)].flatMap((m) => [...m[1].matchAll(/\bid="(\d+)"/g)].map((x) => Number(x[1]))));
    s.autoFilter = /<autoFilter\b[^>]*?\bref="([^"]*)"/.exec(tail)?.[1] ?? null;
    const related = wb.pkg.rels(partName).map((r) => ({ type: String(r.Type), name: OoxmlPackage.resolveTarget(partName, r.Target) })).filter((r) => wb.pkg.has(r.name));
    s.notes = sorted(related.filter((r) => r.type.endsWith('/comments')).flatMap((r) => [...wb.pkg.text(r.name).matchAll(/<comment\b[^>]*?\bref="([^"]+)"/g)].map((m) => m[1])));
    s.noteBoxes = sorted(related.filter((r) => r.type.endsWith('/vmlDrawing')).flatMap((r) => [...wb.pkg.text(r.name).matchAll(/<x:Row>(\d+)<\/x:Row>\s*<x:Column>(\d+)<\/x:Column>/g)].map((m) => [Number(m[1]), Number(m[2])])));
    s.anchors = sorted(related.filter((r) => r.type.endsWith('/drawing')).flatMap((r) => [...wb.pkg.text(r.name).matchAll(/<xdr:from><xdr:col>(\d+)<\/xdr:col>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/g)].map((m) => [Number(m[1]), Number(m[2])])));
    out.sheets[name] = s;
  }
  for (const n of wb.definedNames()) {
    if (!n.name.startsWith('_xlnm')) out.names[n.name] = bare(n.ref);
  }
  out.tables = sorted(wb.tables().map((t) => [t.name, t.sheet, t.ref, t.columns]));
  for (const name of wb.pkg.partNames()) {
    if (!/^xl\/charts\/chart\d+\.xml$/.test(name)) continue;
    out.charts.push(...[...wb.pkg.text(name).matchAll(/<c:f>([^<]*)<\/c:f>/g)].map((m) => bare(unesc(m[1]))));
  }
  sorted(out.charts);
  return out;
}

/** The same edit, on bytes, through our own code. */
export function applyEdit(bytes, e) {
  const wb = Workbook.open(bytes);
  if (e.axis === 'row' && e.op === 'insert') wb.insertRows(e.sheet, e.at, e.count);
  else if (e.axis === 'row') wb.deleteRows(e.sheet, e.at, e.count);
  else if (e.op === 'insert') wb.insertCols(e.sheet, e.at, e.count);
  else wb.deleteCols(e.sheet, e.at, e.count);
  return wb.save();
}

const letter = (n) => { let s = ''; let v = n + 1; while (v > 0) { const r = (v - 1) % 26; s = String.fromCharCode(65 + r) + s; v = (v - r - 1) / 26; } return s; };
function excelAddress(e) {
  if (e.axis === 'row') return `${e.at + 1}:${e.at + e.count}`;
  return `${letter(e.at)}:${letter(e.at + e.count - 1)}`;
}

const BUILD = `
$ws = $wb.Worksheets.Item(1); $ws.Name = 'Data'
$other = $wb.Worksheets.Add([Type]::Missing, $ws); $other.Name = 'Other'
$rows = @(@('Item','Qty','Price'),@('Nut',1,2),@('Bolt',2,3),@('Cog',3,4),@('Gear',4,5),@('Pin',5,6))
for ($r = 0; $r -lt $rows.Count; $r++) { for ($c = 0; $c -lt 3; $c++) { $v = $rows[$r][$c]; if ($v -is [string]) { $ws.Cells.Item($r + 1, $c + 1).Value2 = $v } else { $ws.Cells.Item($r + 1, $c + 1).Value2 = [double]$v } } }
$ws.Range('D1').Value2 = 'Total'; $ws.Range('E1').Value2 = 'Running'
$ws.Range('D2').Formula = '=B2*C2'; $ws.Range('D2:D6').FillDown() | Out-Null
$ws.Range('E2').Formula = '=SUM($B$2:B2)'; $ws.Range('E2:E6').FillDown() | Out-Null
$ws.Range('F1').Formula = '=SUM(D2:D6)'
$ws.Range('G1').Formula = '=Other!A1+1'
$ws.Range('H1').Value2 = 'yes'; $ws.Range('H2').Value2 = 'no'; $ws.Range('H3').Value2 = 'maybe'
$ws.Range('A9').Value2 = 'Q1'; $ws.Range('B9').Value2 = 'Q2'; $ws.Range('C9').Value2 = 'Q3'
$tb = @(@(1,2,3),@(4,5,6),@(7,8,9),@(10,11,12))
for ($r = 0; $r -lt 4; $r++) { for ($c = 0; $c -lt 3; $c++) { $ws.Cells.Item($r + 10, $c + 1).Value2 = [double]$tb[$r][$c] } }
$lo = $ws.ListObjects.Add(1, $ws.Range('A9:C13'), [Type]::Missing, 1); $lo.Name = 'Sales'
$ws.Range('E10').Formula = '=SUM(Sales[Q1])'
$ws.Range('E11').Formula = '=SUM(Sales[Q2],Sales[Q3])'
$ws.Range('I5:J6').Merge() | Out-Null
$ws.Activate() | Out-Null
$ws.Range('A1').Select() | Out-Null
$ws.Range('B2:B6').FormatConditions.Add(2, 0, '=B1>2') | Out-Null
$ws.Range('B2:B6').FormatConditions.Item(1).Interior.ColorIndex = 6
$ws.Range('C2:C6').Validation.Add(3, 1, 1, '=$H$1:$H$3') | Out-Null
$ws.Range('B15').Validation.Add(1, 1, 1, '1', '=$F$1') | Out-Null
$ws.Hyperlinks.Add($ws.Range('A2'), 'https://example.com/') | Out-Null
$ws.Hyperlinks.Add($ws.Range('A3'), '', 'Data!H2', [Type]::Missing, 'there') | Out-Null
$ws.Range('C3').AddComment('A note') | Out-Null
$ws.Range('D4').AddComment('Another') | Out-Null
$wb.Names.Add('Rate', '=Data!$C$2') | Out-Null
$wb.Names.Add('Span', '=Data!$B$2:$B$6') | Out-Null
$ws.HPageBreaks.Add($ws.Range('A8')) | Out-Null
$sh = $ws.Shapes.AddChart2(201, 51, 500, 150, 200, 120)
$sh.Chart.SetSourceData($ws.Range('B1:B6'))
$other.Range('A1').Formula = '=Data!B3'
$other.Range('A2').Formula = '=SUM(Data!B2:B6)'
$other.Range('A3').Formula = '=Data!D4*2'
$other.Range('B2').Value2 = 5
$other.Range('C1').Formula = '=A1+1'; $other.Range('C1:C3').FillDown() | Out-Null
$wb.RemovePersonalInformation = $true
`;

function runExcel(body, needs) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-excel-structure-'));
  const script = path.join(dir, 'ask.ps1');
  fs.writeFileSync(script, `${String.fromCharCode(0xfeff)}$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
try {
${body.replaceAll('$DIR', dir)}
} finally { $xl.Quit() | Out-Null; [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
`);
  for (const [name, content] of Object.entries(needs ?? {})) fs.writeFileSync(path.join(dir, name), content);
  const r = spawnSync('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', script], { encoding: 'utf8', maxBuffer: 1 << 26 });
  if (r.status !== 0) throw new Error(`Excel did not answer: ${r.stderr || r.stdout || r.error}`);
  return { dir, stdout: r.stdout };
}

function main() {
  const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
  const fixture = path.join(root, 'tests/fixtures/excel-structure.xlsx');
  const answers = path.join(root, 'tests/fixtures/excel-structure.json');
  let { dir } = runExcel(`$wb = $xl.Workbooks.Add()
while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item(1).Delete() }
${BUILD}
$wb.SaveAs('$DIR\\built.xlsx', 51) | Out-Null
$wb.Close($false) | Out-Null`);
  fs.copyFileSync(path.join(dir, 'built.xlsx'), fixture);

  const steps = EDITS.map((e, i) => `$wb = $xl.Workbooks.Open('${fixture}', 0, $false)
$ws = $wb.Worksheets.Item('${e.sheet}')
$rng = $ws.${e.axis === 'row' ? 'Rows' : 'Columns'}.Item('${excelAddress(e)}')
$rng.${e.op === 'insert' ? 'Insert' : 'Delete'}() | Out-Null
$xl.CalculateFull()
$wb.SaveCopyAs('$DIR\\edit${i}.xlsx') | Out-Null
$wb.Close($false) | Out-Null`).join('\n');
  ({ dir } = runExcel(steps));
  const out = { base: describe(fs.readFileSync(fixture)), edits: EDITS.map((e, i) => ({ name: e.name, excel: describe(fs.readFileSync(path.join(dir, `edit${i}.xlsx`))) })) };
  fs.writeFileSync(answers, `${JSON.stringify(out, null, 1)}\n`);
  console.log(`excel-structure: ${EDITS.length} edits answered`);

  if (process.argv.includes('--verify')) {
    const needs = {};
    EDITS.forEach((e, i) => { needs[`ours${i}.xlsx`] = applyEdit(fs.readFileSync(fixture), e); });
    const check = EDITS.map((e, i) => `$wb = $xl.Workbooks.Open('$DIR\\ours${i}.xlsx', 0, $false)
$xl.CalculateFull()
$wb.SaveCopyAs('$DIR\\back${i}.xlsx') | Out-Null
$wb.Close($false) | Out-Null`).join('\n');
    const r = runExcel(check, needs);
    let bad = 0;
    EDITS.forEach((e, i) => {
      const back = JSON.stringify(describe(fs.readFileSync(path.join(r.dir, `back${i}.xlsx`))));
      const ours = JSON.stringify(out.edits[i].excel);
      if (back !== ours) { bad += 1; console.log(`differs after Excel reads ours: ${e.name}`); }
    });
    console.log(`verify: ${EDITS.length - bad} of ${EDITS.length} read back as Excel's own`);
  }
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) main();
