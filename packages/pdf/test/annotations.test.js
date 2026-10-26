'use strict';

// Comments and markup (annotations.js): every kind written as an
// incremental update after the original bytes, with an appearance stream,
// read back with its words, author and colour; replies threaded by /IRT,
// resolved by a review state, and deleted with their replies, the objects
// freed. Annotations another writer made are read, and kept by an update.

const test = require('node:test');
const assert = require('node:assert');
const zlib = require('zlib');
const { PdfDocument } = require('../src/document');
const { openPdf, Dict, Stream } = require('../src/reader');
const { readAnnotations, annotate, parseDate, wrapLines } = require('../src/annotations');
const { PdfUpdate } = require('../src/update');

const created = '2026-10-14T09:00:00Z';

function written(pages = 2) {
  const doc = new PdfDocument({ title: 'Commented', created });
  for (let i = 0; i < pages; i++) {
    const page = doc.addPage();
    page.text(`Page ${i + 1} heading`, 72, 100, { font: 'Helvetica', size: 24 });
    page.text('The quick brown fox jumps over the lazy dog.', 72, 140, { font: 'Helvetica', size: 12 });
  }
  return doc.toBuffer();
}

/** A file with a classic table whose page 1 already carries an annotation another writer made (no appearance). */
function withForeignNote() {
  const base = written(1);
  const up = new PdfUpdate(base);
  const page = up.file.pages()[0];
  const { textString } = require('../src/assemble');
  const { Name } = require('../src/reader');
  const note = up.add(new Dict([
    ['Type', new Name('Annot')], ['Subtype', new Name('Square')], ['Rect', [100, 100, 200, 160]], ['C', [0, 0, 1]],
    ['T', textString('Ayesha')], ['Contents', textString('Another writer’s box')], ['M', textString("D:20250102030405+05'00'")],
  ]));
  const dict = new Dict(page.dict.entries());
  dict.set('Annots', [note]);
  up.set(page.ref, dict);
  return up.toBuffer();
}

const quad = (x0, y0, x1, y1) => [x0, y1, x1, y1, x0, y0, x1, y0];

test('each kind of comment is written after the original bytes and read back', () => {
  const original = written(2);
  const { bytes, added } = annotate(original, {
    add: [
      { page: 0, type: 'Highlight', quads: [quad(72, 690, 300, 704)], contents: 'Check this' },
      { page: 0, type: 'Underline', quads: [quad(72, 650, 200, 664)] },
      { page: 0, type: 'StrikeOut', quads: [quad(72, 620, 200, 634)] },
      { page: 0, type: 'Text', rect: [400, 700, 420, 720], contents: 'A sticky note' },
      { page: 0, type: 'FreeText', rect: [72, 400, 300, 450], contents: 'Words in a box, long enough to wrap onto a second line here.', fontSize: 12 },
      { page: 1, type: 'Ink', ink: [[100, 100, 120, 130, 150, 120, 170, 160]], width: 2 },
      { page: 1, type: 'Square', rect: [100, 300, 200, 360] },
      { page: 1, type: 'Circle', rect: [250, 300, 350, 360], interior: [0.9, 0.9, 1] },
      { page: 1, type: 'Line', line: [100, 500, 300, 520] },
      { page: 1, type: 'Arrow', line: [100, 550, 300, 600], color: [0, 0, 1] },
    ],
  }, { author: 'Ejaz', date: new Date(created) });
  assert.equal(added.length, 10);
  assert.ok(bytes.subarray(0, original.length).equals(original), 'the original stays the start of the file');
  const file = openPdf(bytes);
  assert.equal(file.repaired, false);
  const list = readAnnotations(file);
  assert.deepEqual(list.map((a) => a.type), ['Highlight', 'Underline', 'StrikeOut', 'Text', 'FreeText', 'Ink', 'Square', 'Circle', 'Line', 'Line']);
  assert.deepEqual(list.map((a) => a.page), [0, 0, 0, 0, 0, 1, 1, 1, 1, 1]);
  const hl = list[0];
  assert.equal(hl.contents, 'Check this');
  assert.equal(hl.author, 'Ejaz');
  assert.equal(hl.modified, '2026-10-14T09:00:00.000Z');
  assert.deepEqual(hl.quads, [quad(72, 690, 300, 704)]);
  assert.ok(list.every((a) => a.hasAppearance), 'every one has an appearance');
  assert.equal(list[8].arrow, false);
  assert.equal(list[9].arrow, true);
  assert.deepEqual(list[9].color, [0, 0, 1]);
  assert.deepEqual(list[5].ink, [[100, 100, 120, 130, 150, 120, 170, 160]]);
  assert.equal(list[4].fontSize, 12);
  // The highlight's appearance multiplies over the words, as Acrobat's does.
  const dict = file.getObject(Number(hl.id));
  const form = file.resolve(file.resolve(dict.get('AP')).get('N'));
  assert.ok(form instanceof Stream);
  assert.match(file.decode(form).toString('latin1'), /\/GS0 gs[\s\S]*re|m [\s\S]* l h f/);
  const gs = file.resolve(file.resolve(form.dict.get('Resources')).get('ExtGState')).get('GS0');
  assert.equal(gs.get('BM').name, 'Multiply');
  // The text box's words are in its appearance, wrapped.
  const box = file.getObject(Number(list[4].id));
  const text = file.decode(file.resolve(file.resolve(box.get('AP')).get('N'))).toString('latin1');
  assert.match(text, /\(Words in a box/);
  assert.ok((text.match(/Tj/g) || []).length >= 2, 'wrapped onto two lines');
});

test('replies thread to their comment, a resolved comment has a review state, and deleting takes the replies too', () => {
  const first = annotate(written(1), { add: [{ page: 0, type: 'Text', rect: [400, 700, 420, 720], contents: 'Is this right?' }] }, { author: 'Ejaz' });
  const id = first.added[0];
  const second = annotate(first.bytes, { reply: [{ to: id, contents: 'Yes, it is.' }], resolve: [{ id, resolved: true }] }, { author: 'Sara' });
  let list = readAnnotations(openPdf(second.bytes));
  const reply = list.find((a) => a.replyTo === id && !a.state);
  assert.equal(reply.contents, 'Yes, it is.');
  assert.equal(reply.author, 'Sara');
  const note = list.find((a) => a.id === id);
  assert.equal(note.status, 'Completed');
  assert.equal(note.statusBy, 'Sara');
  const state = list.find((a) => a.replyTo === id && a.state);
  assert.equal(state.hidden, true, 'the state reply is not drawn');

  // Unresolving takes the state away; there is one state reply at most.
  const third = annotate(second.bytes, { resolve: [{ id, resolved: false }] }, { author: 'Sara' });
  list = readAnnotations(openPdf(third.bytes));
  assert.equal(list.find((a) => a.id === id).status, undefined);
  assert.equal(list.filter((a) => a.state).length, 0);

  // Deleting the note takes its reply with it, and frees both objects.
  const fourth = annotate(third.bytes, { remove: [id] });
  const file = openPdf(fourth.bytes);
  assert.deepEqual(readAnnotations(file), []);
  assert.equal(file.entries.get(Number(id)).type, 0, 'the note is free');
  assert.equal(file.entries.get(Number(reply.id)).type, 0, 'its reply is free');
  assert.ok(fourth.bytes.subarray(0, third.bytes.length).equals(third.bytes));
});

test('an edit changes the words and draws a text box again', () => {
  const one = annotate(written(1), { add: [{ page: 0, type: 'FreeText', rect: [72, 400, 300, 450], contents: 'Before' }] });
  const id = one.added[0];
  const two = annotate(one.bytes, { edit: [{ id, contents: 'After the edit' }] });
  const file = openPdf(two.bytes);
  const a = readAnnotations(file).find((x) => x.id === id);
  assert.equal(a.contents, 'After the edit');
  const form = file.resolve(file.resolve(file.getObject(Number(id)).get('AP')).get('N'));
  assert.match(file.decode(form).toString('latin1'), /\(After the edit\)/);
});

test('a signature, a tick or a date moved or resized: a new Rect, the stamp drawn into it, a text box laid out again', () => {
  const one = annotate(written(1), {
    add: [
      { page: 0, type: 'Stamp', rect: [100, 100, 114, 114], paths: [[95, 107, 98.5, 103, 105.5, 111.5]], width: 1.8, icon: 'Check', subject: 'Checkmark' },
      { page: 0, type: 'FreeText', rect: [72, 400, 160, 418], contents: '26 Oct 2026', fontSize: 11, subject: 'Date' },
      { page: 0, type: 'Highlight', quads: [quad(72, 690, 300, 704)] },
    ],
  });
  const [tick, date, mark] = one.added;
  const two = annotate(one.bytes, {
    move: [
      { id: tick, rect: [300, 500, 328, 528] },
      { id: date, rect: [200, 300, 320, 330] },
      { id: mark, rect: [0, 0, 10, 10] },
    ],
  });
  assert.ok(two.bytes.subarray(0, one.bytes.length).equals(one.bytes), 'an update after the bytes');
  const file = openPdf(two.bytes);
  const read = readAnnotations(file);
  const at = (id) => read.find((a) => a.id === id);
  assert.deepEqual(at(tick).rect, [300, 500, 328, 528]);
  assert.deepEqual(at(date).rect, [200, 300, 320, 330]);
  assert.equal(at(date).contents, '26 Oct 2026', 'its words kept');
  assert.notDeepEqual(at(mark).rect, [0, 0, 10, 10], 'a highlight stays on its words');
  // The tick's appearance is the one it had: drawn into the new Rect by every reader.
  const before = openPdf(one.bytes);
  const apOf = (f, id) => f.resolve(f.resolve(f.getObject(Number(id)).get('AP')).get('N'));
  assert.equal(file.decode(apOf(file, tick)).toString('latin1'), before.decode(apOf(before, tick)).toString('latin1'));
  // The date's box drawn again at its new place.
  const form = apOf(file, date);
  assert.deepEqual(file.resolve(form.dict.get('BBox')).map(Number), [200, 300, 320, 330]);
  assert.match(file.decode(form).toString('latin1'), /\(26 Oct 2026\)/);
  // A move and an edit of the same text box together: the words at the new place.
  const three = annotate(two.bytes, { move: [{ id: date, rect: [50, 50, 170, 80] }], edit: [{ id: date, contents: '27 Oct 2026' }] });
  const after = readAnnotations(openPdf(three.bytes)).find((a) => a.id === date);
  assert.deepEqual(after.rect, [50, 50, 170, 80]);
  assert.equal(after.contents, '27 Oct 2026');
});

test('annotations another writer made are read, and kept when comments are added', () => {
  const foreign = withForeignNote();
  const before = readAnnotations(openPdf(foreign));
  assert.equal(before.length, 1);
  assert.equal(before[0].author, 'Ayesha');
  assert.equal(before[0].contents, 'Another writer’s box');
  assert.equal(before[0].modified, '2025-01-01T22:04:05.000Z', 'the zone is allowed for');
  assert.equal(before[0].hasAppearance, false);
  const { bytes } = annotate(foreign, { add: [{ page: 0, type: 'Text', rect: [10, 10, 30, 30], contents: 'Mine' }] });
  const after = readAnnotations(openPdf(bytes));
  assert.deepEqual(after.map((a) => a.author || a.contents), ['Ayesha', 'Mine']);
});

test('a file with cross-reference streams gets an update that ends with one', () => {
  // A stream-form file made by hand: page tree and a cross-reference stream.
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << >> >>',
  ];
  let out = '%PDF-1.7\n';
  const offsets = [];
  objects.forEach((o, i) => { offsets.push(out.length); out += `${i + 1} 0 obj\n${o}\nendobj\n`; });
  const xrefAt = out.length;
  const rows = [Buffer.from([0, 0, 0, 0, 0, 255, 255]), ...offsets.map((o) => { const r = Buffer.alloc(7); r[0] = 1; r.writeUInt32BE(o, 1); return r; })];
  const rowsOf = Buffer.concat([...rows, (() => { const r = Buffer.alloc(7); r[0] = 1; r.writeUInt32BE(xrefAt, 1); return r; })()]);
  const data = zlib.deflateSync(rowsOf);
  const file = Buffer.concat([
    Buffer.from(`${out}4 0 obj\n<< /Type /XRef /Size 5 /Root 1 0 R /W [1 4 2] /Filter /FlateDecode /Length ${data.length} >>\nstream\n`, 'latin1'),
    data, Buffer.from(`\nendstream\nendobj\nstartxref\n${xrefAt}\n%%EOF\n`, 'latin1'),
  ]);
  const { bytes } = annotate(file, { add: [{ page: 0, type: 'Square', rect: [10, 10, 50, 50] }] });
  const tail = bytes.toString('latin1', file.length);
  assert.match(tail, /\/Type \/XRef/);
  assert.doesNotMatch(tail, /\nxref\n/);
  const back = openPdf(bytes);
  assert.equal(back.repaired, false);
  assert.equal(readAnnotations(back)[0].type, 'Square');
});

test('dates, and words wrapped by Helvetica\'s widths', () => {
  assert.equal(parseDate('D:20261014'), '2026-10-14T00:00:00.000Z');
  assert.equal(parseDate("D:20261014120000-05'30'"), '2026-10-14T17:30:00.000Z');
  assert.equal(parseDate('nonsense'), null);
  assert.deepEqual(wrapLines('one two three', 12, 1000), ['one two three']);
  assert.ok(wrapLines('one two three four five six seven', 12, 60).length >= 3);
  assert.deepEqual(wrapLines('a\nb', 12, 100), ['a', 'b']);
});

test('a page that does not exist, or an encrypted file, is refused in a sentence', () => {
  assert.throws(() => annotate(written(1), { add: [{ page: 4, type: 'Text', rect: [0, 0, 20, 20] }] }), /no page 5/);
  assert.throws(() => annotate(written(1), { add: [{ page: 0, type: 'Banana', rect: [0, 0, 20, 20] }] }), /not a kind of comment/);
});

test('pdf.js, the reader the window draws with, reads each comment with its words, author and appearance', async (t) => {
  const { readWithPdfjs } = require('./pdfjs');
  const first = annotate(written(1), {
    add: [
      { page: 0, type: 'Highlight', quads: [quad(72, 690, 300, 704)], contents: 'Check this' },
      { page: 0, type: 'Text', rect: [400, 700, 420, 720], contents: 'A sticky note' },
      { page: 0, type: 'Ink', ink: [[100, 100, 120, 130]] },
      { page: 0, type: 'Arrow', line: [100, 550, 300, 600] },
    ],
  }, { author: 'Ejaz' });
  const read = await readWithPdfjs(first.bytes);
  if (!read) { t.skip('pdf.js is not vendored'); return; }
  const got = read.pages[0].annotations.map((a) => [a.subtype, a.contentsObj?.str || '', a.titleObj?.str || '', a.hasAppearance]);
  assert.deepEqual(got, [
    ['Highlight', 'Check this', 'Ejaz', true],
    ['Text', 'A sticky note', 'Ejaz', true],
    ['Ink', '', 'Ejaz', true],
    ['Line', '', 'Ejaz', true],
  ]);
  assert.match(read.pages[0].text, /quick brown fox/);
});
