'use strict';

// A redaction leaves its words nowhere (redact.js's scrub, in assemble.js).
// One file carries a secret on its page and in every other place a person
// or a program could find it: the document's title, subject, keywords,
// author and creator, its XMP metadata and a page's, a bookmark, a comment
// and its reply, a text box lying under the mark, a link, an attached file
// (as an annotation and in the catalog), a form field's value, its tooltip,
// a field named with it, and a page label. Only the words on the page are
// marked; applying takes the secret out of all of it, and the file is read
// back by this package's reader, by pdf.js and by a plain search of every
// byte and every inflated stream: the secret is found nowhere, and what did
// not carry it (another bookmark, another comment, another field) is kept.

const test = require('node:test');
const assert = require('node:assert');
const zlib = require('zlib');
const { PdfDocument } = require('../src/document');
const { openPdf, Dict, Name, Ref, Stream } = require('../src/reader');
const { textString } = require('../src/assemble');
const { PdfUpdate, flateStream } = require('../src/update');
const { readAnnotations } = require('../src/annotations');
const { readFields } = require('../src/forms');
const { readOutline, pageLabels, labelText } = require('../src/navigation');
const { applyRedactions, bytesContain, scrubber } = require('../src/redact');
const { widthOfBytes } = require('../src/metrics');
const { encode } = require('../src/encoding');
const { readWithPdfjs } = require('./pdfjs');

const SECRET = 'SECRET-4417';
const H = 841.89;
const LINE = `Account ${SECRET} belongs to the client.`;

function leakyFile() {
  const doc = new PdfDocument({ title: 'Plain title', created: '2026-10-14T09:00:00Z' });
  doc.addPage().text(LINE, 72, 140, { font: 'Helvetica', size: 14 });
  doc.addPage().text('Nothing private is written on this page.', 72, 100, { font: 'Helvetica', size: 14 });
  const up = new PdfUpdate(doc.toBuffer());
  const { file } = up;
  const [one, two] = file.pages();
  const s = (x) => textString(x);
  const add = (entries) => up.add(new Dict(entries));
  const r = (n) => new Ref(n, 0);

  // The document's information and XMP metadata.
  const info = add([['Title', s(`Report on ${SECRET}`)], ['Subject', s(`About ${SECRET}`)], ['Keywords', s(`${SECRET}, accounts`)], ['Author', s(`${SECRET} office`)], ['Creator', s(`${SECRET} tool`)]]);
  const xmp = (words) => up.add(new Stream(new Dict([['Type', new Name('Metadata')], ['Subtype', new Name('XML')]]), Buffer.from(`<?xpacket begin=""?><x:xmpmeta xmlns:x="adobe:ns:meta/"><rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#"><rdf:Description xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>${words}</dc:title></rdf:Description></rdf:RDF></x:xmpmeta><?xpacket end="w"?>`)));
  // An attached file in the catalog, and one as an annotation.
  const embedded = up.add(flateStream([['Type', new Name('EmbeddedFile')]], `${SECRET} inside the attachment`));
  const spec = (name) => add([['Type', new Name('Filespec')], ['F', s(name)], ['UF', s(name)], ['EF', new Dict([['F', embedded]])]]);
  const names = new Dict([['EmbeddedFiles', new Dict([['Names', [s(`${SECRET}.txt`), spec(`${SECRET}.txt`)]]])]]);
  // Bookmarks: one with the secret, one without.
  const root = up.alloc();
  const bad = up.alloc();
  const good = up.alloc();
  up.set(bad, new Dict([['Title', s(`${SECRET} notes`)], ['Parent', r(root)], ['Next', r(good)], ['Dest', [two.ref, new Name('Fit')]]]));
  up.set(good, new Dict([['Title', s('Summary')], ['Parent', r(root)], ['Prev', r(bad)], ['Dest', [one.ref, new Name('Fit')]]]));
  up.set(root, new Dict([['Type', new Name('Outlines')], ['First', r(bad)], ['Last', r(good)], ['Count', 2]]));
  // Page one: a text box lying under the mark, saying something else.
  const at = 72 + widthOfBytes(encode('Account '), 'Helvetica', 14);
  const end = at + widthOfBytes(encode(SECRET), 'Helvetica', 14);
  const base = H - 140;
  const mark = add([['Type', new Name('Annot')], ['Subtype', new Name('Redact')], ['Rect', [at - 0.5, base - 4, end + 0.5, base + 12]], ['P', one.ref]]);
  const under = add([['Type', new Name('Annot')], ['Subtype', new Name('FreeText')], ['Rect', [at, base - 2, at + 40, base + 10]], ['Contents', s('A box on the mark')], ['DA', s('/Helv 9 Tf 0 g')], ['P', one.ref]]);
  // Page two: a comment and its reply, a link, an attached file, a harmless comment, a form.
  const comment = add([['Type', new Name('Annot')], ['Subtype', new Name('Text')], ['Rect', [400, 700, 420, 720]], ['Contents', s(`Call about ${SECRET}`)], ['T', s('Ayesha')], ['P', two.ref]]);
  const reply = add([['Type', new Name('Annot')], ['Subtype', new Name('Text')], ['Rect', [400, 700, 420, 720]], ['Contents', s('Done, as asked.')], ['IRT', comment], ['P', two.ref]]);
  const link = add([['Type', new Name('Annot')], ['Subtype', new Name('Link')], ['Rect', [72, 600, 200, 620]], ['A', new Dict([['S', new Name('URI')], ['URI', s(`https://example.com/${SECRET}`)]])], ['P', two.ref]]);
  const attach = add([['Type', new Name('Annot')], ['Subtype', new Name('FileAttachment')], ['Rect', [300, 600, 320, 620]], ['FS', spec(`${SECRET}.pdf`)], ['P', two.ref]]);
  const harmless = add([['Type', new Name('Annot')], ['Subtype', new Name('Text')], ['Rect', [100, 100, 120, 120]], ['Contents', s('Unrelated note')], ['P', two.ref]]);
  const look = (words) => up.add(flateStream([['Type', new Name('XObject')], ['Subtype', new Name('Form')], ['BBox', [0, 0, 200, 20]], ['Resources', new Dict()]], `/Tx BMC BT /Helv 10 Tf 2 6 Td (${words}) Tj ET EMC`));
  const field = add([['Type', new Name('Annot')], ['Subtype', new Name('Widget')], ['FT', new Name('Tx')], ['T', s('account')], ['TU', s(`Enter ${SECRET}`)], ['V', s(SECRET)], ['Rect', [72, 500, 272, 520]], ['AP', new Dict([['N', look(SECRET)]])], ['P', two.ref]]);
  const named = add([['Type', new Name('Annot')], ['Subtype', new Name('Widget')], ['FT', new Name('Tx')], ['T', s(`${SECRET} holder`)], ['V', s('x')], ['Rect', [72, 460, 272, 480]], ['P', two.ref]]);
  const other = add([['Type', new Name('Annot')], ['Subtype', new Name('Widget')], ['FT', new Name('Tx')], ['T', s('other')], ['V', s('fine')], ['Rect', [72, 420, 272, 440]], ['P', two.ref]]);
  const form = add([['Fields', [field, named, other]], ['DA', s('/Helv 0 Tf 0 g')]]);

  const p1 = new Dict(one.dict.entries());
  p1.set('Annots', [mark, under]);
  up.set(one.ref, p1);
  const p2 = new Dict(two.dict.entries());
  p2.set('Annots', [comment, reply, link, attach, harmless, field, named, other]);
  p2.set('Metadata', xmp(`Page about ${SECRET}`));
  up.set(two.ref, p2);
  const catalog = new Dict(file.catalog().entries());
  catalog.set('Outlines', r(root));
  catalog.set('Names', names);
  catalog.set('AcroForm', form);
  catalog.set('Metadata', xmp(`The ${SECRET} report`));
  catalog.set('PageLabels', new Dict([['Nums', [0, new Dict([['S', new Name('D')], ['P', s(`${SECRET}-`)]])]]]));
  up.set(file.trailer.get('Root'), catalog);
  // The information dictionary, in the trailer's own place.
  const infoRef = file.trailer.get('Info');
  up.set(infoRef, up.get(info));
  return up.toBuffer();
}

test('the leaky file has the secret everywhere before the redaction', () => {
  const bytes = leakyFile();
  assert.equal(bytesContain(bytes, SECRET), true);
  const file = openPdf(bytes);
  assert.match(file.info().title, /SECRET/);
  assert.equal(readOutline(file)[0].title, `${SECRET} notes`);
  assert.equal(labelText(pageLabels(file)[0]), `${SECRET}-1`);
});

test('applying the one mark takes the secret out of everything, and keeps what did not carry it', async () => {
  const { bytes, removed, elsewhere, words } = applyRedactions(leakyFile());
  assert.equal(removed.glyphs, SECRET.length);
  assert.ok(words.includes(SECRET), `learnt from the glyphs, as written: ${words}`);
  assert.deepEqual(elsewhere, {
    comment: 3, // the text box under the mark, the comment and its reply
    title: 1, author: 1, subject: 1, keywords: 1, creator: 1,
    'document metadata': 1, 'attached file': 2, link: 1, bookmark: 1, 'page label': 2, // both pages carried the prefix
    metadata: 1, 'form value': 1, tooltip: 1, 'field name': 1,
  });

  // A plain search of every byte and every inflated stream.
  assert.equal(bytesContain(bytes, SECRET), false);
  assert.equal(bytesContain(bytes, 'SECRET'), false);
  assert.equal(bytesContain(bytes, '4417'), false);

  // This package's reader.
  const file = openPdf(bytes);
  const info = file.info();
  assert.ok(!Object.values(info).some((v) => /SECRET/.test(v)), JSON.stringify(info));
  assert.deepEqual(readOutline(file).map((n) => n.title), ['Summary']);
  assert.deepEqual(readAnnotations(file).map((a) => a.contents), ['Unrelated note']);
  const fields = readFields(file);
  assert.deepEqual(fields.map((f) => `${f.name}=${f.value}`), ['account=', 'Field 1=x', 'other=fine']);
  assert.equal(file.getObject(Number(fields[0].widgets[0].id)).get('AP'), undefined, 'the widget no longer shows the old value');
  assert.deepEqual(pageLabels(file).map(labelText), ['1', '2']);
  assert.doesNotMatch(file.pageText(0) + file.pageText(1), /SECRET/);
  assert.match(file.pageText(0), /belongs to the client/);

  // pdf.js.
  const read = await readWithPdfjs(bytes);
  if (read) {
    const everything = JSON.stringify({
      text: read.pages.map((p) => p.text),
      annotations: read.pages.map((p) => p.annotations.map((a) => [a.contentsObj?.str, a.titleObj?.str, a.url, a.file?.filename, a.fieldValue, a.alternativeText, a.fieldName])),
      fields: read.fields, outline: read.outline, labels: read.labels, info: read.info, metadata: read.metadata,
    });
    assert.doesNotMatch(everything, /SECRET|4417/);
    assert.match(everything, /Unrelated note/);
    assert.match(everything, /Summary/);
  }
});

test('the words a mark records are searched for too, a phrase and its long words, case and spacing aside', () => {
  const s = scrubber(['Zainab  Rehman', 'x']);
  assert.deepEqual(s.needles, ['zainab rehman', 'zainab', 'rehman']);
  assert.ok(s.text('Letter to ZAINAB'));
  assert.ok(s.text('zainab\n rehman'));
  assert.ok(!s.text('Zain'));
  assert.ok(zlib && true);
});
