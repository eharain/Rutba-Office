/**
 * Word's binary documents, of every age, opened in full.
 *
 * The 97-2003 files here were written by Word itself (tools/make-binary-
 * fixtures.ps1): the reader is judged against what Word writes, not what we
 * think the format says. No Office on hand writes the older formats any
 * more, so Word 6.0/95, Word 2.0, Windows Write and Word for DOS files are
 * laid out from their published layouts in fixtures/old-word.js.
 *
 * Each must open as what it is — its words, their bold and italic, sizes,
 * colours and fonts, paragraphs' alignment and indents, lists, tables,
 * notes, headers and footers, pictures and pages — and write out as a .docx
 * this suite then opens.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { readWordDocument, wordKind } from '@rutba/office-formats/msword';
import { docModelToDocx } from '@rutba/office-formats/msdoc-docx';
import { sniff } from '@rutba/office-formats/sniff';
import { OoxmlPackage } from '@rutba/ooxml/package';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { createDocumentService } from '../apps/desktop/main/documents.js';
import { buildOldWord, buildWriteDoc, anld6, defTable6, tinyWmf, pic6 } from './fixtures/old-word.js';

const FIXTURES = path.join(path.dirname(fileURLToPath(import.meta.url)), 'fixtures', 'binary');
const fixture = (name) => new Uint8Array(fs.readFileSync(path.join(FIXTURES, name)));
const le16 = (v) => [v & 0xff, (v >> 8) & 0xff];

/** Every paragraph of a list of blocks, the cells' too. */
function paragraphs(blocks) {
  const out = [];
  for (const b of blocks) {
    if (b.type === 'table') for (const row of b.rows) for (const cell of row.cells) out.push(...paragraphs(cell));
    else out.push(b);
  }
  return out;
}
const textOf = (p) => p.runs.map((r) => r.text ?? '').join('');
const runWith = (model, text) => paragraphs(model.body).flatMap((p) => p.runs).find((r) => r.text === text);
const paraWith = (model, text) => paragraphs(model.body).find((p) => textOf(p).includes(text));
const styleName = (model, p) => model.styles.find((s) => s.istd === p.pap.istd)?.name;

/** The .docx a model writes, opened again: its blocks, and the parts' XML to look into. */
function roundTrip(model) {
  const bytes = docModelToDocx(model);
  const pkg = OoxmlPackage.read(bytes);
  return { view: openDocx(Buffer.from(bytes)), part: (name) => (pkg.has?.(name) === false ? '' : safe(() => pkg.text(name))), pkg };
}
const safe = (fn) => { try { return fn(); } catch { return ''; } };

/* ── Word 97-2003, as Word writes it ──────────────────────────────────── */

test('a Word 97-2003 document opens with its characters as Word set them', () => {
  const bytes = fixture('text.doc');
  assert.equal(wordKind(bytes), 'word97');
  const model = readWordDocument(bytes);
  assert.equal(model.format, 'word97');
  assert.equal(textOf(paragraphs(model.body)[0]), 'Characters');
  assert.match(styleName(model, paragraphs(model.body)[0]), /^heading 1$/i);
  assert.equal(runWith(model, 'bold').chp.bold, true);
  assert.equal(runWith(model, 'italic').chp.italic, true);
  assert.equal(runWith(model, 'underlined').chp.underline, 1);
  assert.equal(runWith(model, 'struck').chp.strike, true);
  assert.equal(runWith(model, 'red').chp.colour, 'C00000');
  assert.equal(runWith(model, 'sixteen point').chp.size, 32);
  assert.equal(model.fonts[runWith(model, 'Courier New').chp.font].name, 'Courier New');
  assert.equal(runWith(model, 'raised').chp.iss, 1);
  assert.equal(runWith(model, 'lowered').chp.iss, 2);
  assert.equal(runWith(model, 'small capitals').chp.smallCaps, true);
  assert.match(textOf(paraWith(model, 'Accents')), /café, naïve, Zürich, Ελληνικά, Русский\./);
});

test('a Word 97-2003 document keeps its paragraphs, breaks and lists', () => {
  const model = readWordDocument(fixture('paragraphs.doc'));
  assert.equal(paraWith(model, 'centred').pap.jc, 1);
  assert.equal(paraWith(model, 'aligned right').pap.jc, 2);
  assert.equal(paraWith(model, 'justified').pap.jc, 3);
  assert.equal(paraWith(model, 'Indented half an inch').pap.left, 720);
  assert.equal(paraWith(model, 'A first line indented').pap.firstLine, 720);
  assert.equal(paraWith(model, 'Eighteen points').pap.before, 360);
  const broken = paraWith(model, 'A line');
  assert.ok(broken.runs.some((r) => r.kind === 'break' && r.type === 'line'));
  assert.ok(broken.runs.some((r) => r.kind === 'tab'));
  const bullet = paraWith(model, 'First bullet');
  const number = paraWith(model, 'First number');
  assert.ok(bullet.pap.ilfo && number.pap.ilfo && bullet.pap.ilfo !== number.pap.ilfo);
  assert.equal(paraWith(model, 'Second bullet').pap.ilfo, bullet.pap.ilfo);
  const listOf = (p) => model.lists.lists.find((l) => l.lsid === model.lists.lfos[p.pap.ilfo - 1].lsid);
  assert.equal(listOf(bullet).levels[0].nfc, 23);
  assert.equal(listOf(number).levels[0].nfc, 0);
  assert.equal(paraWith(model, 'After the lists').pap.ilfo ?? 0, 0);
});

test('a Word 97-2003 document keeps its table, link, footnote, picture, header and footer', () => {
  const model = readWordDocument(fixture('structure.doc'));
  const table = model.body.find((b) => b.type === 'table');
  assert.equal(table.rows.length, 3);
  assert.deepEqual(table.rows[0].cells.map((c) => c.map(textOf).join('')), ['Region', 'Q1', 'Q2']);
  assert.equal(table.rows[0].cells[0][0].runs[0].chp.bold, true);
  assert.equal(model.notes.footnote.length, 1);
  assert.match(model.notes.footnote[0].blocks.map((b) => textOf(b)).join(''), /The footnote's own words\./);
  assert.equal(model.images.length, 1);
  assert.equal(model.images[0].contentType, 'image/png');
  assert.equal(model.headers[0].headers.default.map(textOf).join(''), 'Structure header');
  const footer = model.headers[0].footers.default.flatMap((p) => p.runs);
  assert.ok(footer.some((r) => r.text && /PAGE/.test(r.text)), 'the page number is a field');

  const { view, part } = roundTrip(model);
  assert.ok(view.blocks.some((b) => b.text === 'Second page'));
  assert.match(part('word/document.xml'), /<w:hyperlink r:id="[^"]+">/);
  assert.match(part('word/document.xml'), /<w:footnoteReference w:id="1"\/>/);
  assert.match(part('word/document.xml'), /<w:tbl>/);
  assert.match(part('word/footnotes.xml'), /The footnote/);
  assert.match(part('word/header1.xml'), /Structure header/);
  assert.match(part('word/footer1.xml'), /PAGE/);
});

/* ── Word 6.0/95 ──────────────────────────────────────────────────────── */

const SYMBOL_BULLET = anld6({ before: [0xb7], font: 1 });
const NUMBERED = anld6({ nfc: 0, after: [0x2e] });
function word6() {
  return buildOldWord({
    version: 6,
    fonts: [{ name: 'Times New Roman' }, { name: 'Symbol', charset: 2 }, { name: 'Arial' }, { name: 'Arial Cyr', charset: 204 }],
    styles: [
      { sti: 0, kind: 1, base: null, next: 0, name: 'Normal', papx: [], chpx: [93, 0, 0, 99, 24, 0] },
      { sti: 1, kind: 1, base: 0, next: 0, name: 'heading 1', papx: [21, ...le16(240)], chpx: [85, 1, 99, 32, 0] },
      { sti: 65, kind: 2, base: null, next: 2, name: 'Default Paragraph Font', chpx: [] },
    ],
    main: [
      { istd: 1, runs: [{ text: 'Old Word' }] },
      { runs: [{ text: 'Plain, ' }, { text: 'bold', chpx: [85, 1] }, { text: ', ' }, { text: 'italic', chpx: [86, 1] }, { text: ', ' }, { text: 'red', chpx: [98, 6] }, { text: ' and ' }, { text: 'large', chpx: [99, 40, 0] }, { text: '.' }] },
      { papx: [5, 1, 17, ...le16(720), 22, ...le16(240)], runs: [{ text: 'Centred and indented.' }] },
      { runs: [{ text: [0xcf, 0xf0, 0xe8, 0xe2, 0xe5, 0xf2], chpx: [93, 3, 0] }, { text: ' and caf\xe9' }] },
      { papx: [13, 11, ...SYMBOL_BULLET], runs: [{ text: 'First bullet' }] },
      { papx: [13, 11, ...SYMBOL_BULLET], runs: [{ text: 'Second bullet' }] },
      { runs: [{ text: 'Between the lists' }] },
      { papx: [13, 10, ...NUMBERED], runs: [{ text: 'One' }] },
      { papx: [13, 10, ...NUMBERED], runs: [{ text: 'Two' }] },
      { papx: [24, 1], runs: [{ text: 'A1' }], mark: 7 },
      { papx: [24, 1], runs: [{ text: 'B1' }], mark: 7 },
      { papx: [24, 1, 25, 1, ...defTable6([0, 2000, 4000])], runs: [], mark: 7 },
      { runs: [{ text: 'A note' }, { text: '\x02', chpx: [117, 1], noteRef: true }, { text: ' and a picture ' }, { text: '\x01', chpx: [117, 1], picture: 'pic' }] },
      { runs: [{ text: 'The last paragraph.' }] },
    ],
    footnotes: [{ runs: [{ text: '\x02', chpx: [117, 1] }, { text: ' The note.' }] }],
    header: [{ runs: [{ text: 'Header page ' }, { text: '\x13' }, { text: ' PAGE ' }, { text: '\x14' }, { text: '1' }, { text: '\x15' }] }],
    sepx: [153, 0x02, 164, ...le16(12240), 165, ...le16(15840), 166, ...le16(1440), 167, ...le16(1440)],
    pictures: [{ name: 'pic', pic: pic6(tinyWmf(), 1440, 720) }],
  });
}

test('a Word 6.0/95 document opens with its characters, paragraphs and styles', () => {
  const bytes = word6();
  assert.equal(wordKind(bytes), 'word6');
  const model = readWordDocument(bytes);
  assert.equal(model.format, 'word6');
  const heading = paragraphs(model.body)[0];
  assert.equal(textOf(heading), 'Old Word');
  assert.equal(styleName(model, heading), 'heading 1');
  assert.equal(heading.runs[0].chp.bold, true, 'the heading style is bold');
  assert.equal(heading.runs[0].chp.size, 32);
  assert.equal(runWith(model, 'bold').chp.bold, true);
  assert.equal(runWith(model, 'italic').chp.italic, true);
  assert.equal(runWith(model, 'red').chp.colour, 'FF0000');
  assert.equal(runWith(model, 'large').chp.size, 40);
  assert.equal(runWith(model, 'Plain, ').chp.bold ?? false, false);
  const centred = paraWith(model, 'Centred');
  assert.equal(centred.pap.jc, 1);
  assert.equal(centred.pap.left, 720);
  assert.equal(centred.pap.after, 240);
});

test('a Word 6 document\'s text is read in its own code page, a font\'s own character set included', () => {
  const model = readWordDocument(word6());
  assert.equal(textOf(paraWith(model, 'caf')), 'Привет and café');
});

test('Word 6\'s numbering becomes lists: bullets and numbers, each starting again after a break', () => {
  const model = readWordDocument(word6());
  const first = paraWith(model, 'First bullet');
  const second = paraWith(model, 'Second bullet');
  const one = paraWith(model, 'One');
  const two = paraWith(model, 'Two');
  assert.ok(first.pap.ilfo, 'the bullet paragraph is in a list');
  assert.equal(second.pap.ilfo, first.pap.ilfo, 'the two bullets are one list');
  assert.equal(two.pap.ilfo, one.pap.ilfo, 'the two numbers are one list');
  assert.notEqual(one.pap.ilfo, first.pap.ilfo);
  assert.equal(paraWith(model, 'Between').pap.ilfo ?? 0, 0);
  const listOf = (p) => model.lists.lists.find((l) => l.lsid === model.lists.lfos[p.pap.ilfo - 1].lsid);
  assert.deepEqual(listOf(first).levels[0].codes, [0xf0b7], 'the symbol font\'s bullet, where Word keeps it');
  assert.equal(listOf(first).levels[0].nfc, 23);
  assert.deepEqual(listOf(one).levels[0].codes, [0, 0x2e], 'the number, then a full stop');
  const { part } = roundTrip(model);
  assert.match(part('word/numbering.xml'), /<w:numFmt w:val="bullet"\/>/);
  assert.match(part('word/numbering.xml'), /<w:lvlText w:val="%1\."\/>/);
  assert.equal((part('word/document.xml').match(/<w:numPr>/g) || []).length, 4);
});

test('a Word 6 document keeps its table, footnote, picture, header and page', () => {
  const model = readWordDocument(word6());
  const table = model.body.find((b) => b.type === 'table');
  assert.ok(table, 'the table is a table');
  assert.deepEqual(table.rows[0].cells.map((c) => c.map(textOf).join('')), ['A1', 'B1']);
  assert.deepEqual(table.rows[0].tap.cells.map((c) => c.width), [2000, 2000]);
  assert.equal(table.rows[0].tap.cells[0].borders.top.type, 1);
  assert.equal(model.notes.footnote.length, 1);
  assert.equal(model.notes.footnote[0].blocks.map(textOf).join(''), ' The note.');
  assert.ok(paraWith(model, 'A note').runs.some((r) => r.kind === 'noteRef'));
  assert.equal(model.images.length, 1);
  assert.equal(model.images[0].contentType, 'image/x-wmf');
  assert.equal(model.images[0].widthTwips, 1440);
  assert.equal(model.images[0].heightTwips, 720);
  assert.deepEqual([...model.images[0].bytes.subarray(0, 4)], [0xd7, 0xcd, 0xc6, 0x9a], 'given a placeable header');
  assert.equal(model.sections[0].sep.width, 12240);
  assert.equal(model.sections[0].sep.left, 1440);
  const header = model.headers[0].headers.default.flatMap((p) => p.runs);
  assert.equal(header[0].text, 'Header page ');
  assert.ok(header.some((r) => r.kind === 'fieldBegin'));

  const { view, part, pkg } = roundTrip(model);
  assert.ok(view.blocks.some((b) => b.text === 'The last paragraph.'));
  assert.match(part('word/header1.xml'), /PAGE/);
  assert.match(part('word/footnotes.xml'), /The note\./);
  assert.match(part('word/document.xml'), /<w:drawing>/);
  assert.ok(pkg.text('[Content_Types].xml').includes('Extension="wmf"'));
});

/* ── Word 2.0 ─────────────────────────────────────────────────────────── */

function word2() {
  return buildOldWord({
    version: 2,
    fonts: [{ name: 'Tms Rmn' }, { name: 'Symbol', charset: 2 }, { name: 'Helv' }],
    styles: {
      cstcStd: 2,
      entries: [
        { name: '', chpx: [0x01], papx: { stc: 254, sprms: [5, 1] }, next: 0, base: 0 },
        { name: null, chpx: null, papx: null },
        { name: '', chpx: [], papx: { stc: 0, sprms: [] }, next: 0, base: 222 },
        { name: 'Quote', chpx: [0x02], papx: { stc: 1, sprms: [17, ...le16(720)] }, next: 1, base: 0 },
      ],
    },
    main: [
      { istd: 254, runs: [{ text: 'Word 2 heading' }] },
      { runs: [{ text: 'Plain and ' }, { text: 'bold', chpx: [0x01] }, { text: '.' }] },
      { istd: 1, runs: [{ text: 'A quote.' }] },
      { runs: [{ text: 'Large red', chpx: [0, 0, 0x05, 0, 0, 0, 36, 0, 0, 6] }] },
      { runs: [{ text: 'Helv text', chpx: [0, 0, 0x02, 0, 2, 0] }] },
    ],
    sepx: [139, ...le16(12240), 140, ...le16(15840)],
  });
}

test('a Word 2.0 document opens: its style sheet, its character structure and its sprms', () => {
  const bytes = word2();
  assert.equal(wordKind(bytes), 'word2');
  assert.equal(sniff(bytes, 'old.doc').kind, 'doc');
  const model = readWordDocument(bytes);
  assert.equal(model.format, 'word2');
  const heading = paraWith(model, 'Word 2 heading');
  assert.equal(styleName(model, heading), 'heading 1');
  assert.equal(heading.runs[0].chp.bold, true, 'the heading style turns bold on');
  assert.equal(heading.pap.jc, 1);
  assert.equal(runWith(model, 'bold').chp.bold, true);
  assert.equal(runWith(model, 'Plain and ').chp.bold ?? false, false);
  const quote = paraWith(model, 'A quote.');
  assert.equal(styleName(model, quote), 'Quote');
  assert.equal(quote.runs[0].chp.italic, true);
  assert.equal(quote.pap.left, 720);
  assert.equal(runWith(model, 'Large red').chp.size, 36);
  assert.equal(runWith(model, 'Large red').chp.colour, 'FF0000');
  assert.equal(model.fonts[runWith(model, 'Helv text').chp.font].name, 'Helv');
  assert.equal(model.sections[0].sep.width, 12240);
  const { view } = roundTrip(model);
  assert.deepEqual(view.blocks.map((b) => b.text), ['Word 2 heading', 'Plain and bold.', 'A quote.', 'Large red', 'Helv text']);
});

/* ── Windows Write and Word for DOS ───────────────────────────────────── */

function writeFile() {
  const wmf = tinyWmf();
  const object = new Array(40).fill(0);
  object[0] = 0x88;
  [[10, 1440], [12, 720], [36, 1000], [38, 1000]].forEach(([o, v]) => { object[o] = v & 0xff; object[o + 1] = v >> 8; });
  object[32] = wmf.length & 0xff;
  object[33] = wmf.length >> 8;
  const head = (rhc) => [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 240, 0, 0, 0, 0, 0, rhc];
  return buildWriteDoc({
    fonts: ['Arial', 'Times New Roman'],
    paragraphs: [
      { pap: [0, 1], runs: [{ text: 'Write document', prop: [1, 0x01 | (1 << 2), 28] }] },
      { runs: [{ text: 'Plain then ' }, { text: 'italic', prop: [1, 0x02] }, { text: '.' }] },
      { pap: head(0x07), runs: [{ text: 'Footer text' }] },
      { pap: head(0x10), object: [...object, ...wmf] },
      { runs: [{ text: 'Last line \x96 en dash' }] },
    ],
  });
}

test('a Windows Write document opens with its fonts, running foot and picture', () => {
  const bytes = writeFile();
  assert.equal(wordKind(bytes), 'write');
  assert.equal(sniff(bytes, 'letter.wri').kind, 'doc');
  const model = readWordDocument(bytes);
  assert.equal(model.format, 'write');
  const title = paraWith(model, 'Write document');
  assert.equal(title.pap.jc, 1);
  assert.equal(title.runs[0].chp.bold, true);
  assert.equal(title.runs[0].chp.size, 28);
  assert.equal(model.fonts[title.runs[0].chp.font].name, 'Times New Roman');
  assert.equal(runWith(model, 'italic').chp.italic, true);
  assert.equal(textOf(paraWith(model, 'Last line')), 'Last line – en dash');
  assert.ok(!paraWith(model, 'Footer text'), 'the running foot is not in the body');
  assert.equal(model.headers[0].footers.default.map(textOf).join(''), 'Footer text');
  assert.equal(model.sections[0].sep.titlePage, true, 'Write leaves the first page without its running heads unless told');
  assert.equal(model.images.length, 1);
  assert.equal(model.images[0].widthTwips, 1440);
  const { view, part } = roundTrip(model);
  assert.deepEqual(view.blocks.map((b) => b.text).filter(Boolean), ['Write document', 'Plain then italic.', 'Last line – en dash']);
  assert.match(part('word/footer1.xml'), /Footer text/);
  assert.match(part('word/document.xml'), /<w:drawing>/);
});

test('a Word for DOS document opens in its DOS code page, with its footnote', () => {
  const bytes = buildWriteDoc({
    dos: true,
    codepage: 850,
    paragraphs: [
      { runs: [{ text: 'DOS document', prop: [0, 0x01, 24] }] },
      { runs: [{ text: 'Caf\x82 and a note' }, { text: '1', prop: [(13 << 1) | 1] }] },
      { pap: [(39 << 1) | 1], runs: [{ text: '1 The note text.' }] },
    ],
    notes: [{ ref: [1, 'Caf\x82 and a note'.length], text: 2 }],
  });
  assert.equal(wordKind(bytes), 'worddos');
  const model = readWordDocument(bytes);
  assert.equal(model.format, 'worddos');
  assert.equal(paragraphs(model.body).length, 2, 'the footnote is not in the body');
  assert.equal(runWith(model, 'DOS document').chp.bold, true);
  const noted = paraWith(model, 'Café');
  assert.equal(textOf(noted), 'Café and a note');
  assert.ok(noted.runs.some((r) => r.kind === 'noteRef'));
  assert.equal(model.notes.footnote.length, 1);
  assert.equal(model.notes.footnote[0].blocks.map(textOf).join(''), '1 The note text.');
  const { part } = roundTrip(model);
  assert.match(part('word/footnotes.xml'), /The note text\./);
});

/* ── recognising them, and opening them in the suite ──────────────────── */

test('the older formats are told apart by their first bytes', () => {
  assert.equal(sniff(word2()).container, 'word2');
  assert.equal(sniff(writeFile()).container, 'write');
  assert.equal(sniff(Uint8Array.from([0x09, 0x00, 0x04, 0x00, 0x02, 0x00, 0x10, 0x00, ...new Array(120).fill(0)])).kind, 'xls');
  assert.equal(sniff(Uint8Array.from([0x09, 0x04, 0x06, 0x00, 0x00, 0x00, 0x10, 0x00, ...new Array(120).fill(0)])).kind, 'xls');
  assert.equal(sniff(new Uint8Array(0), 'old.wri').kind, 'doc');
  assert.equal(wordKind(new TextEncoder().encode('not a document at all, just some words for a while')), null);
});

test('the suite opens each kind as a document, and says it saves a .docx', () => {
  const doc = createDocumentService({ holdBlob: () => ({ url: 'blob://held' }) });
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-word-binary-'));
  const open = (name, bytes) => {
    const p = path.join(dir, name);
    fs.writeFileSync(p, bytes);
    const s = doc.open({ path: p });
    doc.close({ id: s.id });
    return s;
  };
  const text = (s) => (s.model.blocks || []).map((b) => (b.runs || []).map((r) => r.text).join('') || b.text || '');
  const w97 = open('structure.doc', fixture('structure.doc'));
  assert.equal(w97.kind, 'doc');
  assert.equal(w97.converted.from, 'doc');
  assert.ok(!w97.converted.partial, 'read in full, not scraped for text');
  assert.ok(text(w97).includes('Second page'));
  const w6 = open('word6.doc', word6());
  assert.ok(text(w6).includes('The last paragraph.'));
  const w2 = open('word2.doc', word2());
  assert.ok(text(w2).includes('A quote.'));
  const wri = open('letter.wri', writeFile());
  assert.equal(wri.converted.from, 'wri');
  assert.ok(text(wri).includes('Plain then italic.'));
});
