// The PDF app's engine: reading anybody's PDF and making new files from its
// pages (extract, reorder, rotate, delete, insert, merge, split) without
// drawing anything, each page's content, resources and fonts copied byte
// for byte.
//
// Every fixture is made here: by the suite's own writer, by the incremental
// updater, and by hand, so the structures other writers use (object
// streams, cross-reference streams with a predictor, a hybrid file, a page
// tree whose nodes pass attributes down, a damaged table) are each met.

import test from 'node:test';
import assert from 'node:assert/strict';
import zlib from 'node:zlib';
import { PdfDocument, appendExecutionPage } from '@rutba/pdf';
import { openPdf, PdfReadError, Stream, Dict, Ref, Name } from '@rutba/pdf/reader';
import { assemble, splitEvery, splitBefore, parsePageRange } from '@rutba/pdf/assemble';

const created = '2026-10-13T09:00:00Z';

/** A PDF of `count` pages from the suite's writer, each saying which it is. */
function written(count, prefix = 'Page') {
  const doc = new PdfDocument({ title: `${prefix} file`, created });
  for (let i = 0; i < count; i++) {
    const page = doc.addPage();
    page.text(`${prefix} ${i + 1}`, 72, 100, { font: 'Helvetica', size: 24 });
    page.text('Body text on the page.', 72, 140, { font: 'Courier', size: 12 });
  }
  return doc.toBuffer();
}

/**
 * A PDF written by hand. `objects` maps a number to its body (text, or
 * { dict, data } for a stream). Options put some objects in an object
 * stream, write the cross-references as a stream (with the PNG Up
 * predictor, as most writers do), or both in a hybrid file.
 */
function handmade(objects, { root = 1, objStm = [], xref = 'table', header = '%PDF-1.7', info = null } = {}) {
  const parts = [];
  let length = 0;
  const add = (data) => {
    const b = Buffer.isBuffer(data) ? data : Buffer.from(data, 'latin1');
    parts.push(b);
    length += b.length;
  };
  add(`${header}\n%\xE2\xE3\xCF\xD3\n`);
  const entries = new Map();
  const body = (n) => {
    const o = objects[n];
    if (typeof o === 'string') return o;
    return null;
  };
  let size = Math.max(...Object.keys(objects).map(Number)) + 1;
  for (const key of Object.keys(objects)) {
    const n = Number(key);
    if (objStm.includes(n)) continue;
    entries.set(n, [1, length, 0]);
    const o = objects[n];
    if (typeof o === 'string') add(`${n} 0 obj\n${o}\nendobj\n`);
    else {
      add(`${n} 0 obj\n${o.dict.replace('LENGTH', String(o.data.length))}\nstream\n`);
      add(o.data);
      add('\nendstream\nendobj\n');
    }
  }
  if (objStm.length) {
    const stmNum = size++;
    let head = '';
    let data = '';
    objStm.forEach((n, i) => {
      head += `${n} ${data.length} `;
      data += `${body(n)}\n`;
      entries.set(n, [2, stmNum, i]);
    });
    const raw = zlib.deflateSync(Buffer.from(head + data, 'latin1'));
    entries.set(stmNum, [1, length, 0]);
    add(`${stmNum} 0 obj\n<< /Type /ObjStm /N ${objStm.length} /First ${head.length} /Filter /FlateDecode /Length ${raw.length} >>\nstream\n`);
    add(raw);
    add('\nendstream\nendobj\n');
  }
  const trailerExtra = info ? ` /Info ${info} 0 R` : '';
  if (xref === 'stream' || xref === 'hybrid') {
    const xrefNum = size++;
    const xrefAt = length;
    entries.set(xrefNum, [1, xrefAt, 0]);
    const rows = [];
    for (let n = 0; n < size; n++) {
      const e = entries.get(n) || [0, 0, 65535];
      const row = Buffer.alloc(7);
      row[0] = e[0];
      row.writeUInt32BE(e[1], 1);
      row.writeUInt16BE(e[2], 5);
      rows.push(row);
    }
    // PNG "Up" predictor: each row the difference from the one above.
    const predicted = [];
    let prev = Buffer.alloc(7);
    for (const row of rows) {
      const out = Buffer.alloc(8);
      out[0] = 2;
      for (let i = 0; i < 7; i++) out[i + 1] = (row[i] - prev[i]) & 0xff;
      predicted.push(out);
      prev = row;
    }
    const raw = zlib.deflateSync(Buffer.concat(predicted));
    const stmDict = `<< /Type /XRef /Size ${size} /W [1 4 2] /Root ${root} 0 R${trailerExtra} /Filter /FlateDecode /DecodeParms << /Predictor 12 /Columns 7 >> /Length ${raw.length} >>`;
    if (xref === 'stream') {
      add(`${xrefNum} 0 obj\n${stmDict}\nstream\n`);
      add(raw);
      add('\nendstream\nendobj\n');
      add(`startxref\n${xrefAt}\n%%EOF\n`);
      return Buffer.concat(parts);
    }
    // Hybrid: the stream for new readers, a table (objects in the stream marked free) for old ones.
    add(`${xrefNum} 0 obj\n${stmDict}\nstream\n`);
    add(raw);
    add('\nendstream\nendobj\n');
    const tableAt = length;
    let table = `xref\n0 ${size}\n0000000000 65535 f \n`;
    for (let n = 1; n < size; n++) {
      const e = entries.get(n);
      table += e && e[0] === 1 ? `${String(e[1]).padStart(10, '0')} 00000 n \n` : '0000000000 00001 f \n';
    }
    add(table);
    add(`trailer\n<< /Size ${size} /Root ${root} 0 R${trailerExtra} /XRefStm ${xrefAt} >>\nstartxref\n${tableAt}\n%%EOF\n`);
    return Buffer.concat(parts);
  }
  const xrefAt = length;
  let table = `xref\n0 ${size}\n0000000000 65535 f \n`;
  for (let n = 1; n < size; n++) {
    const e = entries.get(n);
    table += e ? `${String(e[1]).padStart(10, '0')} 00000 n \n` : '0000000000 00001 f \n';
  }
  add(table);
  add(`trailer\n<< /Size ${size} /Root ${root} 0 R${trailerExtra} >>\nstartxref\n${xrefAt}\n%%EOF\n`);
  return Buffer.concat(parts);
}

const content = (text) => {
  const data = Buffer.from(`BT /F1 18 Tf 40 300 Td (${text}) Tj ET`, 'latin1');
  return { dict: '<< /Length LENGTH >>', data };
};
const deflated = (text) => {
  const data = zlib.deflateSync(Buffer.from(`BT /F1 18 Tf 40 300 Td (${text}) Tj ET`, 'latin1'));
  return { dict: '<< /Filter /FlateDecode /Length LENGTH >>', data };
};

/**
 * Three pages under a two-level tree whose nodes pass a MediaBox, the
 * resources (one font) and a rotation down; the last page has its own box.
 */
function treeObjects() {
  return {
    1: '<< /Type /Catalog /Pages 2 0 R /Lang (en-GB) >>',
    2: '<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 3 /MediaBox [0 0 300 400] /Resources << /Font << /F1 5 0 R >> >> /Rotate 90 >>',
    3: '<< /Type /Pages /Parent 2 0 R /Kids [6 0 R 7 0 R] /Count 2 >>',
    4: '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 500 500] /Rotate 0 /Contents 10 0 R >>',
    5: '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
    6: '<< /Type /Page /Parent 3 0 R /Contents 8 0 R >>',
    7: '<< /Type /Page /Parent 3 0 R /Contents 9 0 R >>',
    8: deflated('Alpha'),
    9: content('Bravo'),
    10: deflated('Charlie'),
  };
}

const texts = (file) => Array.from({ length: file.pageCount }, (_, i) => file.pageText(i).split('\n')[0]);
const firstWords = (file) => Array.from({ length: file.pageCount }, (_, i) => file.pageText(i));

test('the writer\'s own PDF reads back: pages, sizes and the words on each', () => {
  const file = openPdf(written(3));
  assert.equal(file.pageCount, 3);
  assert.equal(file.encrypted, false);
  assert.deepEqual(texts(file), ['Page 1', 'Page 2', 'Page 3']);
  const info = file.pageInfo(0);
  assert.equal(Math.round(info.width), 595);
  assert.equal(Math.round(info.height), 842);
  assert.equal(info.rotate, 0);
  assert.equal(file.info().title, 'Page file');
  assert.match(file.pageText(1), /Body text on the page\./);
});

test('a page tree\'s nodes pass the box, resources and rotation down to their pages', () => {
  const file = openPdf(handmade(treeObjects()));
  assert.equal(file.pageCount, 3);
  assert.deepEqual(texts(file), ['Alpha', 'Bravo', 'Charlie']);
  assert.deepEqual(file.pageInfo(0), { mediaBox: [0, 0, 300, 400], cropBox: [0, 0, 300, 400], width: 300, height: 400, rotate: 90 });
  assert.equal(file.pageInfo(2).width, 500, 'the last page keeps its own box');
  assert.equal(file.pageInfo(2).rotate, 0, 'and its own rotation');
  // Written out on their own, the pages carry what they inherited.
  const out = openPdf(assemble([{ file, page: 1 }]));
  const page = out.pages()[0].dict;
  assert.ok(page.get('MediaBox'), 'the box is on the page now');
  assert.equal(page.get('Rotate'), 90);
  const resources = out.resolve(page.get('Resources'));
  const font = out.resolve(out.resolve(resources.get('Font')).get('F1'));
  assert.equal(font.get('BaseFont').name, 'Helvetica');
  assert.deepEqual(texts(out), ['Bravo']);
});

test('object streams and a cross-reference stream with a predictor are read', () => {
  const bytes = handmade(treeObjects(), { objStm: [1, 2, 3, 5, 6, 7], xref: 'stream' });
  assert.ok(!bytes.includes('\nxref\n'), 'there is no table to fall back on');
  const file = openPdf(bytes);
  assert.equal(file.repaired, false, 'read through its own cross-references, not by scanning');
  assert.equal(file.pageCount, 3);
  assert.deepEqual(texts(file), ['Alpha', 'Bravo', 'Charlie']);
  const out = openPdf(assemble([{ file, page: 2 }, { file, page: 0 }]));
  assert.deepEqual(texts(out), ['Charlie', 'Alpha']);
});

test('a hybrid file finds the objects its table marks free in the stream its trailer names', () => {
  const file = openPdf(handmade(treeObjects(), { objStm: [2, 3, 6, 7], xref: 'hybrid' }));
  assert.equal(file.repaired, false);
  assert.deepEqual(texts(file), ['Alpha', 'Bravo', 'Charlie']);
});

test('an incremental update is read: its new page, and an object it redefines', () => {
  const original = written(2);
  const appended = appendExecutionPage(original, ({ page }) => page.text('Signed page', 72, 100, { size: 18 }));
  assert.ok(appended.subarray(0, original.length).equals(original), 'the update only appends');
  const file = openPdf(appended);
  assert.equal(file.pageCount, 3);
  assert.deepEqual(texts(file), ['Page 1', 'Page 2', 'Signed page']);

  // A second update by hand: page 1's content redefined at a new offset.
  const before = openPdf(appended);
  const size = before.trailer.get('Size');
  const contentsRef = before.pages()[0].dict.get('Contents');
  const startxref = /startxref\s+(\d+)\s+%%EOF\s*$/.exec(appended.toString('latin1'))[1];
  const stream = 'BT /F1 24 Tf 72 700 Td (Page one, changed) Tj ET';
  const at = appended.length;
  const tail = `${contentsRef.num} 0 obj\n<< /Length ${stream.length} >>\nstream\n${stream}\nendstream\nendobj\n`;
  const xrefAt = at + tail.length;
  const update = `${tail}xref\n${contentsRef.num} 1\n${String(at).padStart(10, '0')} 00000 n \ntrailer\n<< /Size ${size} /Root ${before.trailer.get('Root').num} 0 R /Prev ${startxref} >>\nstartxref\n${xrefAt}\n%%EOF\n`;
  const twice = Buffer.concat([appended, Buffer.from(update, 'latin1')]);
  const again = openPdf(twice);
  assert.equal(again.repaired, false);
  assert.deepEqual(texts(again), ['Page one, changed', 'Page 2', 'Signed page']);
});

test('reordering keeps every page\'s content stream and fonts byte for byte', () => {
  const source = openPdf(written(4));
  const out = openPdf(assemble([3, 0, 2, 1].map((page) => ({ file: source, page }))));
  assert.deepEqual(texts(out), ['Page 4', 'Page 1', 'Page 3', 'Page 2']);
  [3, 0, 2, 1].forEach((from, to) => {
    const a = source.resolve(source.pages()[from].dict.get('Contents'));
    const b = out.resolve(out.pages()[to].dict.get('Contents'));
    assert.ok(a instanceof Stream && b instanceof Stream);
    assert.ok(a.data.equals(b.data), `page ${to + 1}'s content is the source page ${from + 1}'s bytes`);
  });
  // The fonts the pages share are copied once, not once a page.
  const fonts = [];
  for (const [n] of out.entries) {
    const v = out.getObject(new Ref(n));
    if (v instanceof Dict && v.type === 'Font') fonts.push(v.get('BaseFont').name);
  }
  assert.deepEqual(fonts.sort(), ['Courier', 'Helvetica']);
});

test('rotation is written on the page, added to what it had', () => {
  const file = openPdf(handmade(treeObjects()));
  const out = openPdf(assemble([
    { file, page: 0, rotate: 180 },  // inherited 90, turned right twice
    { file, page: 2, rotate: 270 },
    { file, page: 1, rotate: 360 },
  ]));
  assert.deepEqual([0, 1, 2].map((i) => out.pageInfo(i).rotate), [180, 270, 0]);
  assert.equal(out.pages()[2].dict.has('Rotate'), false, 'an upright page carries no /Rotate');
});

test('deleting and extracting keep only the pages named, and nothing they alone used', () => {
  const doc = new PdfDocument({ created });
  const a = doc.addPage();
  a.text('Kept', 72, 100, { font: 'Helvetica' });
  const b = doc.addPage();
  b.text('Dropped', 72, 100, { font: 'Courier' });
  const file = openPdf(doc.toBuffer());
  const out = assemble([{ file, page: 0 }]);
  const back = openPdf(out);
  assert.deepEqual(texts(back), ['Kept']);
  assert.ok(!out.includes('Dropped'), 'the deleted page\'s content is not in the file');
});

test('inserting from another file and merging bring each file\'s pages, fonts and pictures', () => {
  const first = openPdf(written(2, 'First'));
  const second = openPdf(handmade(treeObjects(), { objStm: [2, 3, 6, 7], xref: 'stream' }));
  const merged = openPdf(assemble([
    { file: first, page: 0 },
    { file: second, page: 0 },
    { file: second, page: 1 },
    { file: first, page: 1 },
    { file: second, page: 2 },
  ]));
  assert.equal(merged.pageCount, 5);
  assert.deepEqual(texts(merged), ['First 1', 'Alpha', 'Bravo', 'First 2', 'Charlie']);
  assert.equal(merged.pageInfo(1).rotate, 90, 'an inserted page keeps the rotation it inherited');
  assert.equal(merged.pageInfo(0).rotate, 0);
  assert.equal(merged.pageInfo(1).width, 300);
  assert.equal(Math.round(merged.pageInfo(0).width), 595);
});

test('a link to a kept page follows it to its new place; one to a deleted page loses its destination', () => {
  const objects = {
    1: '<< /Type /Catalog /Pages 2 0 R >>',
    2: '<< /Type /Pages /Kids [3 0 R 4 0 R 5 0 R] /Count 3 /MediaBox [0 0 200 200] >>',
    3: '<< /Type /Page /Parent 2 0 R /Contents 6 0 R /Annots [9 0 R 10 0 R] >>',
    4: '<< /Type /Page /Parent 2 0 R /Contents 7 0 R >>',
    5: '<< /Type /Page /Parent 2 0 R /Contents 8 0 R >>',
    6: content('One'),
    7: content('Two'),
    8: content('Three'),
    9: '<< /Type /Annot /Subtype /Link /Rect [0 0 50 50] /Dest [5 0 R /Fit] /P 3 0 R >>',
    10: '<< /Type /Annot /Subtype /Link /Rect [60 0 110 50] /Dest [4 0 R /Fit] /P 3 0 R >>',
  };
  const file = openPdf(handmade(objects));
  const out = openPdf(assemble([{ file, page: 2 }, { file, page: 0 }]));
  assert.deepEqual(texts(out), ['Three', 'One']);
  const annots = out.resolve(out.pages()[1].dict.get('Annots')).map((a) => out.resolve(a));
  const toThree = annots[0].get('Dest')[0];
  assert.ok(toThree instanceof Ref);
  assert.equal(toThree.num, out.pages()[0].ref.num, 'the link to page three now names the new first page');
  assert.equal(annots[1].get('Dest'), undefined, 'page two was left out, so its link goes nowhere');
  assert.equal(annots[0].get('P').num, out.pages()[1].ref.num, 'an annotation names its own page');
});

test('a damaged file is read by finding its objects: wrong offsets, no startxref', () => {
  const good = written(3);
  // Every offset in the table pushed off by ten bytes.
  const shifted = Buffer.from(good.toString('latin1').replace(/^(\d{10}) 00000 n/gm, (m, o) => `${String(Number(o) + 10).padStart(10, '0')} 00000 n`), 'latin1');
  const a = openPdf(shifted);
  assert.equal(a.repaired, true);
  assert.deepEqual(texts(a), ['Page 1', 'Page 2', 'Page 3']);
  // The end of the file cut off: no table, no trailer, no startxref.
  const cut = good.subarray(0, good.indexOf('xref\n0 '));
  const b = openPdf(cut);
  assert.equal(b.repaired, true);
  assert.deepEqual(texts(b), ['Page 1', 'Page 2', 'Page 3']);
  // And what comes out of a repaired file is whole.
  const c = openPdf(assemble([{ file: b, page: 1 }]));
  assert.equal(c.repaired, false);
  assert.deepEqual(texts(c), ['Page 2']);
});

test('bytes before the header shift every offset, and are allowed for', () => {
  const file = openPdf(Buffer.concat([Buffer.from('X-Mailer junk\r\n'), written(2)]));
  assert.equal(file.repaired, false);
  assert.deepEqual(texts(file), ['Page 1', 'Page 2']);
});

test('an encrypted file is recognised, and its pages are not changed by guesswork', () => {
  const objects = { ...treeObjects(), 11: '<< /Filter /Standard /V 2 /R 3 /O <00> /U <00> /P -4 >>' };
  let bytes = handmade(objects).toString('latin1');
  bytes = bytes.replace(/trailer\n<< \/Size (\d+) \/Root 1 0 R >>/, (m, s) => `trailer\n<< /Size ${s} /Root 1 0 R /Encrypt 11 0 R /ID [<01> <01>] >>`);
  const file = openPdf(Buffer.from(bytes, 'latin1'));
  assert.equal(file.encrypted, true);
  assert.equal(file.pageCount, 3, 'its pages can still be counted');
  assert.throws(() => assemble([{ file, page: 0 }]), (err) => err instanceof PdfReadError && err.code === 'encrypted' && /password/.test(err.message));
});

test('something that is not a PDF says so', () => {
  assert.throws(() => openPdf(Buffer.from('PK\u0003\u0004 a zip file')), (err) => err.code === 'not-pdf');
  assert.throws(() => openPdf(Buffer.from('%PDF-1.4\nnothing in here\n')), (err) => err instanceof PdfReadError);
});

test('a page tree whose Kids is not a list has no pages there, rather than a fault', () => {
  for (const kids of ['7', '<< /A 1 >>', '(text)']) {
    const bytes = handmade({
      1: '<< /Type /Catalog /Pages 2 0 R >>',
      2: `<< /Type /Pages /Kids ${kids} /Count 1 >>`,
    });
    try {
      const file = openPdf(bytes);
      assert.equal(file.pages().length, 0);
    } catch (err) {
      assert.ok(err instanceof PdfReadError, `Kids ${kids}: ${err.name}: ${err.message}`);
    }
  }
});

test('strings, names and numbers survive the copy', () => {
  const objects = {
    1: '<< /Type /Catalog /Pages 2 0 R >>',
    2: '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    3: '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612.5 -0.25] /Contents 4 0 R /PieceInfo << /A#20B (paren \\( and \\) and \\\\ and \\101\\n) /H <48656c6c6f> /N -.5 >> >>',
    4: content('Strings'),
  };
  const file = openPdf(handmade(objects));
  const out = openPdf(assemble([{ file, page: 0 }]));
  const piece = out.resolve(out.pages()[0].dict.get('PieceInfo'));
  assert.equal(piece.get('A B').bytes.toString('latin1'), 'paren ( and ) and \\ and A\n');
  assert.equal(piece.get('H').bytes.toString('latin1'), 'Hello');
  assert.equal(piece.get('N'), -0.5);
  assert.deepEqual(out.resolve(out.pages()[0].dict.get('MediaBox')), [0, 0, 612.5, -0.25]);
});

test('split and page ranges', () => {
  assert.deepEqual(splitEvery(7, 3), [[0, 1, 2], [3, 4, 5], [6]]);
  assert.deepEqual(splitEvery(2, 1), [[0], [1]]);
  assert.deepEqual(splitBefore(10, [3, 7]), [[0, 1, 2], [3, 4, 5, 6], [7, 8, 9]]);
  assert.deepEqual(splitBefore(4, [0, 9, 2, 2]), [[0, 1], [2, 3]]);
  assert.deepEqual(parsePageRange('1-3, 5, 8-', 9), [0, 1, 2, 4, 7, 8]);
  assert.equal(parsePageRange('0', 5), null);
  assert.equal(parsePageRange('2-1', 5), null);
  assert.equal(parsePageRange('abc', 5), null);
  // A split written out: each part a whole PDF of its own pages.
  const file = openPdf(written(5));
  const parts = splitEvery(file.pageCount, 2).map((pages) => openPdf(assemble(pages.map((page) => ({ file, page })))));
  assert.deepEqual(parts.map(texts), [['Page 1', 'Page 2'], ['Page 3', 'Page 4'], ['Page 5']]);
});

test('the information dictionary goes with the first file, and the title can be given', () => {
  const file = openPdf(written(1, 'Report'));
  const out = openPdf(assemble([{ file, page: 0 }], { title: 'Extracted pages' }));
  assert.equal(out.info().title, 'Extracted pages');
  assert.equal(out.info().producer, 'Rutba Office');
  const unicode = openPdf(assemble([{ file, page: 0 }], { title: 'رپورٹ' }));
  assert.equal(unicode.info().title, 'رپورٹ');
  assert.equal(openPdf(assemble([{ file, page: 0 }])).info().title, 'Report file');
});

test('another reader opens what this writes (pdf.js, when it is installed)', async (t) => {
  let pdfjs;
  try { pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs'); } catch { t.skip('pdf.js is not installed'); return; }
  const second = openPdf(handmade(treeObjects(), { objStm: [2, 3, 6, 7], xref: 'stream' }));
  const first = openPdf(written(2, 'First'));
  const bytes = assemble([{ file: second, page: 2 }, { file: first, page: 1, rotate: 90 }, { file: second, page: 0 }]);
  const doc = await pdfjs.getDocument({ data: new Uint8Array(bytes), useSystemFonts: false, verbosity: 0 }).promise;
  assert.equal(doc.numPages, 3);
  const words = [];
  const rotations = [];
  for (let i = 1; i <= 3; i++) {
    const page = await doc.getPage(i);
    rotations.push(page.rotate);
    const tc = await page.getTextContent();
    words.push(tc.items.map((it) => it.str).join(' ').trim().split(/\s+/).slice(0, 2).join(' '));
  }
  assert.deepEqual(words, ['Charlie', 'First 2', 'Alpha']);
  assert.deepEqual(rotations, [0, 90, 90]);
  await doc.destroy();
});

test('values the reader hands out are the format\'s own', () => {
  const file = openPdf(handmade(treeObjects()));
  const catalog = file.catalog();
  assert.ok(catalog instanceof Dict);
  assert.ok(catalog.get('Type') instanceof Name);
  assert.equal(catalog.get('Lang').text(), 'en-GB');
  const out = openPdf(assemble([{ file, page: 0 }]));
  assert.equal(out.catalog().get('Lang').text(), 'en-GB', 'the language goes with the pages');
  assert.equal(firstWords(out)[0], 'Alpha');
});
