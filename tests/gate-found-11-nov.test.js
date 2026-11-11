// Found by the gate's fuzzers and corpus on 11 Nov 2026, each written out here as the fuzz run cannot be.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Worker } from 'node:worker_threads';
import { Spreadsheet, parse } from '@rutba/formula';
import { readOdf } from '@rutba/office-formats/odf';
import { writeZip, ZipEntry } from '@rutba/ooxml/zip';

const odsOf = (content) => {
  const entry = (name, text) => {
    const e = new ZipEntry({ name, method: 0, crc: 0, compressedSize: 0, uncompressedSize: 0, compressed: Buffer.alloc(0), dosTime: 0, dosDate: 0x2821 });
    e.data = Buffer.from(text, 'utf8');
    return e;
  };
  return Buffer.from(writeZip([entry('mimetype', 'application/vnd.oasis.opendocument.spreadsheet'), entry('content.xml', content)]));
};
const NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"';
const sheetXml = (rows) => `<office:document-content ${NS}><office:body><office:spreadsheet><table:table table:name="S">${rows}</table:table></office:spreadsheet></office:body></office:document-content>`;

test('a formula past Excel\'s 8,192 characters is refused by the parser, and the cell keeps what its file stored', () => {
  const long = '=A1' + '+A1'.repeat(60000);
  assert.throws(() => parse(long), (e) => e.type === '#VALUE!' && /longer than 8192 characters/.test(e.detail));
  // A long one inside the limit is worked out (the deepest, on a worker's stack, in the next test).
  const s = new Spreadsheet();
  s.setCell('S', 0, 1, 1);
  s.setCell('S', 1, 0, '=1' + '+1'.repeat(999));
  assert.equal(s.getValue('S', 1, 0), 1000);
  s.setCell('S', 2, 0, long, { cached: 42 });
  assert.equal(s.getValue('S', 2, 0), 42);
});

test('a very long formula is worked out on a worker thread\'s stack, as the document service runs', async () => {
  const code = `
    const { parentPort } = require('node:worker_threads');
    import(${JSON.stringify(new URL('../packages/formula/src/index.js', import.meta.url).href)}).then(({ Spreadsheet }) => {
      const s = new Spreadsheet();
      s.setCell('S', 0, 1, 2);
      s.setCell('S', 1, 0, '=B1' + '+B1'.repeat(2729));
      s.setCell('S', 2, 0, '=A1' + '+A1'.repeat(60000));
      parentPort.postMessage([s.getValue('S', 1, 0), String(s.getValue('S', 2, 0)?.type ?? s.getValue('S', 2, 0))]);
    }, (e) => parentPort.postMessage(['threw', e.message]));`;
  const answer = await new Promise((resolve, reject) => {
    const w = new Worker(code, { eval: true });
    w.once('message', resolve);
    w.once('error', reject);
  });
  assert.deepEqual(answer, [5460, '#VALUE!']);
});

test('a merged span of 1e309, NaN or -1 in an .ods is no span, and one past the sheet is held to its edge', () => {
  const cell = (cs, rs) => `<table:table-cell table:number-columns-spanned="${cs}" table:number-rows-spanned="${rs}" office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell>`;
  const t0 = performance.now();
  // 1e309 is Infinity, whose column letters were worked out for ever.
  const odf = readOdf(odsOf(sheetXml(`<table:table-row>${cell('1e309', 'NaN')}${cell('-1', '1e309')}${cell('3', '2')}</table:table-row><table:table-row/><table:table-row>${cell('2147483647', '2147483647')}</table:table-row>`)));
  assert.ok(performance.now() - t0 < 2000);
  assert.deepEqual(odf.sheets[0].merges.sort(), ['A3:XFD1048576', 'C1:E2']);
});

test('a row repeated to the sheet\'s end with a long formula in it is refused in a sentence, not built to gigabytes', () => {
  const formula = `of:=${'('.repeat(5000)}1${')'.repeat(5000)}`;
  const xml = sheetXml(`<table:table-row table:number-rows-repeated="2147483647"><table:table-cell table:formula="${formula}" office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell></table:table-row>`);
  const t0 = performance.now();
  assert.throws(() => readOdf(odsOf(xml)), /more cells in it than this program can open/);
  assert.ok(performance.now() - t0 < 2000, `${Math.round(performance.now() - t0)} ms`);
  // A row of short formulas repeated a thousand times is still read.
  const fine = readOdf(odsOf(sheetXml(`<table:table-row table:number-rows-repeated="1000"><table:table-cell table:formula="of:=1+1" office:value-type="float" office:value="2"><text:p>2</text:p></table:table-cell></table:table-row>`)));
  assert.equal(fine.sheets[0].rows.length, 1000);
});

test('a span over cells an earlier span took is dropped, as Excel drops it, and spans side by side are kept', () => {
  const span = (cs, rs) => `<table:table-cell table:number-columns-spanned="${cs}" table:number-rows-spanned="${rs}" office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell>`;
  // The row repeated three times asks for A1:B5 on top of A2:B6 on top of A3:B7; C1:D1 and A9:B9 touch nothing.
  const covered = '<table:covered-table-cell/>';
  const odf = readOdf(odsOf(sheetXml(`<table:table-row table:number-rows-repeated="3">${span(2, 5)}${covered}${span(2, 1)}${covered}</table:table-row><table:table-row/><table:table-row/><table:table-row/><table:table-row/><table:table-row/><table:table-row>${span(2, 1)}${covered}</table:table-row>`)));
  assert.deepEqual(odf.sheets[0].merges, ['A1:B5', 'C1:D1', 'C2:D2', 'C3:D3', 'A9:B9']);
  // A span starting outside an earlier one but reaching into it is dropped too.
  const odf2 = readOdf(odsOf(sheetXml(`<table:table-row><table:table-cell/>${span(1, 4)}</table:table-row><table:table-row>${span(3, 1)}</table:table-row>`)));
  assert.deepEqual(odf2.sheets[0].merges, ['B1:B4']);
});

test('sixty-five thousand rows each merged across the sheet find their merge without looking through all of them', async () => {
  const { SheetView } = await import('../packages/sheet-view/src/view.js');
  const list = Array.from({ length: 65536 }, (_, r) => ({ top: r, bottom: r, left: 0, right: 16383 }));
  const view = { merges: new Map([['S', list]]), activeSheet: 'S' };
  const t0 = performance.now();
  for (let r = 0; r < 2000; r++) for (let c = 0; c < 20; c++) assert.equal(SheetView.prototype.mergeAt.call(view, r * 30, c * 800), list[r * 30]);
  assert.equal(SheetView.prototype.mergeAt.call(view, 70000, 5), null);
  assert.ok(performance.now() - t0 < 2000, `${Math.round(performance.now() - t0)} ms`);
});

test('thirty-two thousand formulas each reading the whole sheet, themselves inside it, open in seconds and are circular', () => {
  const s = new Spreadsheet();
  for (let c = 0; c < 16384; c++) { s.setCell('S', 0, c, '=1/0'); s.setCell('S', 1, c, '=SUM(A1:XFD1048576)'); }
  const t0 = performance.now();
  const out = s.recalculate({ all: true });
  assert.ok(performance.now() - t0 < 10000, `${Math.round(performance.now() - t0)} ms`);
  assert.equal(out.calculated, 16384);
  assert.equal(out.cycles.length, 16384);
  assert.equal(s.getValue('S', 1, 0).type, '#CIRCULAR!');
});

test('three thousand formulas each reading the same three thousand cells are worked out in a moment, and an edit reaches them all', () => {
  const s = new Spreadsheet();
  s.manual = true;
  // ROWS reads the range as SUM does, as far as the order goes; the first and last sum it, for an edit to reach.
  for (let r = 0; r < 3000; r++) { s.setCell('S', r, 0, 1); s.setCell('S', r, 1, `=${r === 0 || r === 2999 ? 'SUM' : 'ROWS'}($A$1:$A$3000)+${r}`); }
  // One reading its own column: it is circular, and its readers through the range are not held up.
  s.setCell('S', 0, 2, '=SUM(C1:C100)');
  const t0 = performance.now();
  s.recalculate({ all: true, force: true });
  // Nine million edges took eight seconds; the order is now built in a moment, and the rest is reading the cells.
  assert.ok(performance.now() - t0 < 4000, `${Math.round(performance.now() - t0)} ms`);
  assert.equal(s.getValue('S', 0, 1), 3000);
  assert.equal(s.getValue('S', 1, 1), 3001);
  assert.equal(s.getValue('S', 2999, 1), 5999);
  assert.equal(s.getValue('S', 0, 2).type, '#CIRCULAR!');
  s.manual = false;
  s.setCell('S', 1234, 0, 101);
  assert.equal(s.getValue('S', 0, 1), 3100);
  assert.equal(s.getValue('S', 2999, 1), 6099);
});

test('formulas inside a wide range are still worked out before its readers, and an edit reaches the readers', () => {
  const s = new Spreadsheet();
  // 200 formulas in B, more than make the range a node of its own; D1 sums all of A:B, D2 reads D1.
  for (let r = 0; r < 200; r++) { s.setCell('S', r, 0, r + 1); s.setCell('S', r, 1, `=A${r + 1}*2`); }
  s.setCell('S', 0, 3, '=SUM(A1:B1048576)');
  s.setCell('S', 1, 3, '=D1+1');
  s.setCell('S', 2, 3, '=INDIRECT("D1")');
  s.recalculate({ all: true });
  assert.equal(s.getValue('S', 0, 3), 60300);
  assert.equal(s.getValue('S', 1, 3), 60301);
  assert.equal(s.getValue('S', 2, 3), 60300);
  s.setCell('S', 0, 0, 1001);
  assert.equal(s.getValue('S', 0, 1), 2002);
  assert.equal(s.getValue('S', 0, 3), 63300);
  assert.equal(s.getValue('S', 1, 3), 63301);
  // Calculate Sheet keeps another sheet's waiting, and leaves no range node waiting with it.
  s.manual = true;
  s.setCell('S', 1, 0, 5);
  s.recalculate({ force: true, sheet: 'S' });
  assert.equal(s.getValue('S', 0, 3), 63300 - 6 + 15);
  assert.ok([...s.pending].every((k) => !k.includes('\u0000')));
  // A formula inside the range that reads its reader is circular, as it was.
  s.manual = false;
  s.setCell('S', 5, 1, '=D1');
  assert.equal(s.getValue('S', 0, 3).type, '#CIRCULAR!');
});

test('an .ods saved again keeps a formula as it was, with no @ written into the file', async () => {
  const { createDocumentService } = await import('../apps/desktop/main/documents.js');
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ods-at-'));
  try {
    const cells = (r) => `<table:table-row><table:table-cell office:value-type="float" office:value="${r}"><text:p>${r}</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="${r * 10}"><text:p>${r * 10}</text:p></table:table-cell></table:table-row>`;
    const formula = '<table:table-row><table:table-cell table:formula="of:=IFERROR(MATCH([.A1:.A3];[.B1:.B3];0);0)" office:value-type="float" office:value="0"><text:p>0</text:p></table:table-cell></table:table-row>';
    const file = path.join(dir, 'book.ods');
    fs.writeFileSync(file, odsOf(sheetXml(cells(1) + cells(2) + cells(3) + formula)));
    const service = createDocumentService({ holdBlob: () => ({ url: 'blob://held' }) });
    const opened = service.open({ path: file });
    const saved = path.join(dir, 'again.ods');
    await service.saveAsync({ id: opened.id, path: saved });
    service.close({ id: opened.id });
    const back = readOdf(fs.readFileSync(saved));
    const f = back.sheets[0].rows[3][0].formula;
    assert.equal(f, '=IFERROR(MATCH(A1:A3,B1:B3,0),0)');
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
