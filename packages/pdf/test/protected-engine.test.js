'use strict';

// The reader and the writers on PDFs protected with a password: each kind
// (RC4 40 and 128 bit, AES-128, AES-256, with classic cross-references and
// with object streams) opened with its user and its owner password and
// refused a wrong one; a file with only an owner password opened without
// one, its permissions kept; metadata and strings left plain by the file's
// own choice; comments, fields and bookmarks read; and every write (a
// comment, a form value, a text layer, pages into a new file, a redaction,
// a smaller file) kept under the same protection, as pdf.js confirms.

const test = require('node:test');
const assert = require('node:assert');
const { PdfDocument } = require('../src/document');
const { openPdf, Dict, Name, Stream, PdfReadError } = require('../src/reader');
const { assemble } = require('../src/assemble');
const { PdfUpdate } = require('../src/update');
const { readAnnotations, annotate } = require('../src/annotations');
const { readFields, fillFields } = require('../src/forms');
const { readOutline } = require('../src/navigation');
const { addTextLayer } = require('../src/textlayer');
const { applyRedactions } = require('../src/redact');
const { compressPdf } = require('../src/compress');
const { protectPdf, KINDS } = require('./protect-fixture');
const { formPdf } = require('./form-fixture');
const { book } = require('./nav-fixture');
const { loadPdfjs } = require('./pdfjs');

const USER = 'open sesame';
const OWNER = 'owner key 26';

function plainPdf() {
  const doc = new PdfDocument({ title: 'Board figures', created: '2026-10-26T09:00:00Z' });
  doc.addPage().text('Quarterly figures for the board', 72, 90, { font: 'Helvetica-Bold', size: 20 });
  doc.addPage().text('The second page lists the risks.', 72, 90, { font: 'Helvetica', size: 14 });
  return doc.toBuffer();
}

/** pdf.js's reading of a protected file, with its password. */
async function pdfjsRead(lib, bytes, password) {
  const task = lib.getDocument({ data: new Uint8Array(bytes), password, verbosity: 0 });
  try {
    const doc = await task.promise;
    const pages = [];
    for (let i = 1; i <= doc.numPages; i++) {
      const page = await doc.getPage(i);
      pages.push({
        text: (await page.getTextContent()).items.map((it) => it.str).join(''),
        annotations: (await page.getAnnotations()).map((a) => a.contentsObj?.str || ''),
      });
    }
    return { pages, fields: await doc.getFieldObjects(), title: (await doc.getMetadata()).info.Title };
  } finally {
    await task.destroy();
  }
}

// ── reading ────────────────────────────────────────────────────────────────

test('each kind opens with its user password and its owner password, and a wrong one leaves it locked', () => {
  const plain = plainPdf();
  for (const kind of Object.keys(KINDS)) {
    for (const objectStreams of [false, true]) {
      const what = `${kind}${objectStreams ? ', object streams' : ''}`;
      const bytes = protectPdf(plain, { kind, user: USER, owner: OWNER, objectStreams });
      const wrong = openPdf(bytes, { password: 'not it' });
      assert.equal(wrong.encrypted, true, what);
      assert.equal(wrong.locked, true, `${what}: a wrong password leaves it locked`);
      assert.equal(wrong.unlock('still not it'), null, what);
      assert.deepEqual(wrong.info(), {}, `${what}: nothing is read from a locked file`);
      assert.throws(() => assemble([{ file: wrong, page: 0 }]), (err) => err instanceof PdfReadError && err.code === 'encrypted' && /password/.test(err.message), what);
      const asUser = openPdf(bytes, { password: USER });
      assert.equal(asUser.locked, false, what);
      assert.equal(asUser.access, 'user', what);
      assert.equal(asUser.pageCount, 2, what);
      assert.equal(asUser.pageText(1), 'The second page lists the risks.', what);
      assert.equal(asUser.info().title, 'Board figures', `${what}: the information's strings are decrypted too`);
      const asOwner = openPdf(bytes, { password: OWNER });
      assert.equal(asOwner.access, 'owner', what);
      assert.equal(asOwner.pageText(0), 'Quarterly figures for the board', what);
      // A locked file opened with the password later reads the same.
      assert.equal(wrong.unlock(USER), 'user', what);
      assert.equal(wrong.pageText(1), 'The second page lists the risks.', what);
      assert.equal(wrong.unlock(OWNER), 'owner', `${what}: the owner password lifts a user's opening`);
    }
  }
});

test('a file with only an owner password opens without asking, keeps its permissions, and lifts them with the owner password', () => {
  for (const kind of Object.keys(KINDS)) {
    // Printing, copying and comments allowed; changing the document not (bit 4 cleared).
    const bytes = protectPdf(plainPdf(), { kind, user: '', owner: OWNER, permissions: -4 & ~(1 << 3) });
    const file = openPdf(bytes);
    assert.equal(file.locked, false, kind);
    assert.equal(file.access, 'user', kind);
    assert.equal(file.pageText(0), 'Quarterly figures for the board', kind);
    assert.equal(file.permissions.change, false, `${kind}: changing it is not allowed`);
    assert.equal(file.permissions.print, true, kind);
    assert.equal(file.unlock(OWNER), 'owner', kind);
    assert.equal(file.permissions.change, true, `${kind}: its owner may`);
  }
});

test('metadata left plain (EncryptMetadata false) and strings left plain (the Identity filter) are read as they are', async () => {
  const lib = await loadPdfjs();
  // The catalog given an XMP stream, as Acrobat writes one.
  const up = new PdfUpdate(plainPdf());
  const xmp = '<?xpacket begin="" id="W5M0MpCehiHzreSzNTczkc9d"?><x:xmpmeta xmlns:x="adobe:ns:meta/"><dc:title>Board figures</dc:title></x:xmpmeta><?xpacket end="w"?>';
  const meta = up.add(new Stream(new Dict([['Type', new Name('Metadata')], ['Subtype', new Name('XML')]]), Buffer.from(xmp)));
  const catalog = new Dict(up.file.catalog().entries());
  catalog.set('Metadata', meta);
  up.set(up.file.trailer.get('Root'), catalog);
  const withXmp = up.toBuffer();
  for (const kind of ['aes-128', 'aes-256']) {
    const bytes = protectPdf(withXmp, { kind, user: USER, owner: OWNER, encryptMetadata: false });
    assert.ok(bytes.includes('adobe:ns:meta'), `${kind}: the XMP is in the file as it was`);
    const file = openPdf(bytes, { password: USER });
    assert.equal(file.locked, false, kind);
    assert.match(file.decode(file.resolve(file.catalog().get('Metadata'))).toString(), /<dc:title>Board figures/, kind);
    assert.equal(file.pageText(1), 'The second page lists the risks.', kind);
    if (lib) assert.equal((await pdfjsRead(lib, bytes, USER)).pages[1].text, 'The second page lists the risks.', `${kind}: pdf.js reads it so too`);
    const identity = protectPdf(plainPdf(), { kind, user: USER, owner: OWNER, identityStrings: true });
    assert.ok(identity.includes('Board figures'), `${kind}: the title is in the file plain`);
    const read = openPdf(identity, { password: USER });
    assert.equal(read.info().title, 'Board figures', kind);
    assert.equal(read.pageText(0), 'Quarterly figures for the board', kind);
  }
});

test('comments, form fields and bookmarks are read from a protected file', () => {
  const commented = annotate(plainPdf(), { add: [{ page: 1, type: 'Text', rect: [400, 700, 420, 720], contents: 'Ask about the risks' }] }, { author: 'Ejaz' }).bytes;
  for (const kind of ['rc4-128', 'aes-256']) {
    const notes = readAnnotations(openPdf(protectPdf(commented, { kind, user: USER, owner: OWNER }), { password: USER }));
    assert.deepEqual(notes.map((a) => [a.page, a.type, a.contents, a.author]), [[1, 'Text', 'Ask about the risks', 'Ejaz']], kind);
    const fields = readFields(openPdf(protectPdf(formPdf(), { kind, user: USER, owner: OWNER, objectStreams: true }), { password: OWNER }));
    assert.ok(fields.some((f) => f.name === 'name' && f.type === 'text'), `${kind}: the form's fields`);
    assert.equal(fields.find((f) => f.name === 'locked')?.value, 'fixed', `${kind}: a field's value`);
    const outline = readOutline(openPdf(protectPdf(book(), { kind, user: USER, owner: OWNER }), { password: USER }));
    assert.deepEqual(outline.map((n) => n.title), ['Contents', 'Web'], `${kind}: the bookmarks`);
    assert.deepEqual(outline[0].kids.map((n) => n.title), ['Intro', 'Body'], kind);
  }
});

test('a handler other than the password one is recognised and refused in a sentence', () => {
  const bytes = protectPdf(plainPdf(), { kind: 'aes-128', user: USER });
  const pubsec = Buffer.from(bytes.toString('latin1').replace('/Filter /Standard', '/Filter /Adobe.PubSec'), 'latin1');
  const file = openPdf(pubsec, { password: USER });
  assert.equal(file.locked, true);
  assert.match(file.securityProblem, /certificate/);
  assert.throws(() => assemble([{ file, page: 0 }]), /certificate/);
});

// ── writing: the same protection kept ──────────────────────────────────────

test('a comment and a form value written into a protected file are encrypted with it, and the file still opens only with its passwords', async () => {
  const lib = await loadPdfjs();
  for (const kind of Object.keys(KINDS)) {
    const original = protectPdf(formPdf(), { kind, user: USER, owner: OWNER, objectStreams: kind === 'aes-256' });
    const file = openPdf(original, { password: USER });
    const noted = annotate(file, { add: [{ page: 0, type: 'Text', rect: [400, 700, 420, 720], contents: 'Signed off' }] }, { author: 'Ejaz' }).bytes;
    assert.ok(noted.subarray(0, original.length).equals(original), `${kind}: an update after the original bytes`);
    assert.ok(!noted.includes('Signed off'), `${kind}: the comment's words are not in the file plain`);
    assert.equal(openPdf(noted, { password: 'not it' }).locked, true, `${kind}: still protected`);
    const back = openPdf(noted, { password: OWNER });
    assert.deepEqual(readAnnotations(back).map((a) => a.contents), ['Signed off'], kind);
    const id = readFields(back).find((f) => f.name === 'name').id;
    const filled = fillFields(back, { [id]: 'Ejaz Arain' }).bytes;
    assert.ok(!filled.includes('Ejaz Arain'), `${kind}: the value is not in the file plain`);
    const again = openPdf(filled, { password: USER });
    assert.equal(readFields(again).find((f) => f.name === 'name').value, 'Ejaz Arain', kind);
    if (lib) {
      const seen = await pdfjsRead(lib, filled, USER);
      assert.ok(seen.pages[0].annotations.includes('Signed off'), `${kind}: pdf.js reads the comment`);
      assert.equal(seen.fields.name[0].value, 'Ejaz Arain', `${kind}: and the value`);
    }
  }
});

test('pages written into a new file (Save, Extract, Split) keep the protection: the same passwords, the same permissions', async () => {
  const lib = await loadPdfjs();
  for (const kind of Object.keys(KINDS)) {
    const original = protectPdf(book(), { kind, user: USER, owner: OWNER, permissions: -4 & ~(1 << 4) });
    const file = openPdf(original, { password: OWNER });
    const out = assemble([4, 0, 2].map((page) => ({ file, page })), { title: 'Reordered' });
    assert.ok(!out.includes('Reordered'), `${kind}: the title is encrypted`);
    assert.equal(openPdf(out).locked, true, `${kind}: the new file needs the password`);
    const read = openPdf(out, { password: USER });
    assert.equal(read.access, 'user', kind);
    assert.deepEqual([0, 1, 2].map((i) => read.pageText(i)), ['Page 5', 'Page 1', 'Page 3'], kind);
    assert.equal(read.info().title, 'Reordered', kind);
    assert.equal(read.permissions.copy, false, `${kind}: copying stays not allowed`);
    assert.equal(openPdf(out, { password: OWNER }).access, 'owner', kind);
    assert.ok(readOutline(read).length, `${kind}: the bookmarks to the kept pages go with them`);
    if (lib) assert.deepEqual((await pdfjsRead(lib, out, USER)).pages.map((p) => p.text), ['Page 5', 'Page 1', 'Page 3'], `${kind}: pdf.js reads the new file`);
    // Written without the protection only when that is asked for.
    assert.equal(openPdf(assemble([{ file, page: 0 }], { security: null })).encrypted, false, kind);
  }
});

test('a text layer, a redaction and a smaller file of a protected PDF keep its protection', async () => {
  const lib = await loadPdfjs();
  const doc = new PdfDocument({ title: 'Scan' });
  const image = doc.addImage({ width: 40, height: 20, colorSpace: 'DeviceGray', data: Buffer.alloc(800, 200) });
  doc.addPage().image(image, 0, 0, 595, 842);
  doc.addPage().text('Account SECRET-4417 closed', 72, 100, { font: 'Helvetica', size: 14 });
  for (const kind of ['rc4-128', 'aes-256']) {
    const original = protectPdf(doc.toBuffer(), { kind, user: USER, owner: OWNER });
    const file = openPdf(original, { password: USER });
    const layered = addTextLayer(file, [{ page: 0, lines: [{ words: [{ text: 'Recognised', x: 72, y: 700, width: 80, size: 12 }] }] }]).bytes;
    assert.equal(openPdf(layered).locked, true, kind);
    assert.match(openPdf(layered, { password: USER }).pageContent(0).toString('latin1'), /3 Tr/, `${kind}: the layer is there`);
    if (lib) assert.match((await pdfjsRead(lib, layered, USER)).pages[0].text, /Recognised/, `${kind}: pdf.js finds the layer's word`);
    const { bytes: redacted } = applyRedactions(file, { areas: [{ page: 1, rects: [[60, 730, 400, 760]] }] });
    assert.equal(openPdf(redacted).locked, true, `${kind}: the redacted copy is protected`);
    assert.doesNotMatch(openPdf(redacted, { password: USER }).pageText(1), /SECRET/, `${kind}: the words are gone`);
    const { bytes: smaller } = compressPdf(file, { dpi: 72 });
    assert.equal(openPdf(smaller).locked, true, `${kind}: the smaller file is protected`);
    assert.equal(openPdf(smaller, { password: USER }).pageText(1), 'Account SECRET-4417 closed', kind);
  }
});
