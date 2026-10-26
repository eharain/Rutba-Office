'use strict';

// PDFs protected with a password (the Standard security handler): RC4 of 40
// and 128 bits, AES-128 and AES-256, made at test time by the fixture's own
// encryptor (protect-fixture.js) with known user and owner passwords, and a
// file with only an owner password (its permissions), which opens without
// asking. pdf.js, which the window draws with, is the second reader: it
// checks the fixtures, and the way the window asks for a password.

const test = require('node:test');
const assert = require('node:assert');
const { PdfDocument } = require('../src/document');
const { protectPdf, KINDS } = require('./protect-fixture');
const { loadPdfjs } = require('./pdfjs');

const USER = 'open sesame';
const OWNER = 'owner key 26';

/** A plain two-page file, the words on each saying which page it is. */
function plainPdf() {
  const doc = new PdfDocument({ title: 'Board figures', created: '2026-10-26T09:00:00Z' });
  doc.addPage().text('Quarterly figures for the board', 72, 90, { font: 'Helvetica-Bold', size: 20 });
  doc.addPage().text('The second page lists the risks.', 72, 90, { font: 'Helvetica', size: 14 });
  return doc.toBuffer();
}

/** pdf.js's words on a page, opened as the window opens a file: a password asked for through onPassword. */
async function pdfjsOpen(lib, bytes, answers) {
  const task = lib.getDocument({ data: new Uint8Array(bytes), verbosity: 0 });
  const reasons = [];
  task.onPassword = (update, reason) => {
    reasons.push(reason === lib.PasswordResponses.INCORRECT_PASSWORD ? 'wrong' : 'needed');
    const next = answers.shift();
    if (next === undefined) task.destroy();
    else update(next);
  };
  try {
    const doc = await task.promise;
    const page = await doc.getPage(2);
    const text = (await page.getTextContent()).items.map((it) => it.str).join('');
    return { reasons, pages: doc.numPages, text };
  } catch (err) {
    return { reasons, error: err.name };
  } finally {
    await task.destroy();
  }
}

test('pdf.js asks for the password of each kind, again when it is wrong, and opens with the user or the owner password', async (t) => {
  const lib = await loadPdfjs();
  if (!lib) { t.skip('pdf.js is not vendored here'); return; }
  const plain = plainPdf();
  for (const kind of Object.keys(KINDS)) {
    for (const objectStreams of [false, true]) {
      const bytes = protectPdf(plain, { kind, user: USER, owner: OWNER, objectStreams });
      const byUser = await pdfjsOpen(lib, bytes, ['wrong one', USER]);
      assert.deepEqual(byUser, { reasons: ['needed', 'wrong'], pages: 2, text: 'The second page lists the risks.' }, `${kind}${objectStreams ? ', object streams' : ''}: the user password`);
      const byOwner = await pdfjsOpen(lib, bytes, [OWNER]);
      assert.deepEqual(byOwner, { reasons: ['needed'], pages: 2, text: 'The second page lists the risks.' }, `${kind}: the owner password`);
      const cancelled = await pdfjsOpen(lib, bytes, []);
      assert.deepEqual(cancelled.reasons, ['needed'], `${kind}: asked once, then cancelled`);
      assert.ok(cancelled.error, `${kind}: a cancelled open does not open`);
    }
  }
});

test('pdf.js opens a file with only an owner password without asking', async (t) => {
  const lib = await loadPdfjs();
  if (!lib) { t.skip('pdf.js is not vendored here'); return; }
  for (const kind of ['rc4-128', 'aes-256']) {
    const bytes = protectPdf(plainPdf(), { kind, user: '', owner: OWNER, permissions: -4 & ~(1 << 3) });
    assert.deepEqual(await pdfjsOpen(lib, bytes, []), { reasons: [], pages: 2, text: 'The second page lists the risks.' }, kind);
  }
});
