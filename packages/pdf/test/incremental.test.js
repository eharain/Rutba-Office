'use strict';

const test = require('node:test');
const assert = require('node:assert');
const zlib = require('zlib');
const { renderPaper } = require('../src/paper');
const { appendExecutionPage, Unsupported } = require('../src/incremental');
const { decodePng } = require('../src/images');

let pdfjs = null;
const load = async () => {
  if (pdfjs) return pdfjs;
  try {
    pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
  } catch {
    pdfjs = false;
  }
  return pdfjs;
};
const open = async (buffer) => {
  const lib = await load();
  if (!lib) return null;
  return lib.getDocument({ data: new Uint8Array(buffer), useSystemFonts: false }).promise;
};

const chunk = (type, body) => {
  const out = Buffer.alloc(12 + body.length);
  out.writeUInt32BE(body.length, 0);
  out.write(type, 4, 'latin1');
  body.copy(out, 8);
  return out;
};
const tinyPng = () => Buffer.concat([
  Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
  chunk('IHDR', (() => { const b = Buffer.alloc(13); b.writeUInt32BE(1, 0); b.writeUInt32BE(1, 4); b[8] = 8; b[9] = 6; return b; })()),
  chunk('IDAT', zlib.deflateSync(Buffer.from([0, 40, 60, 200, 255]))),
  chunk('IEND', Buffer.alloc(0)),
]);

const original = () => renderPaper({
  title: 'Agreement',
  created: '2026-09-02T00:00:00Z',
  letterhead: { name: 'Original Ltd' },
  sections: [{ heading: 'Terms', body: 'The page that must not move by one byte.' }],
});

test('the appended file keeps the original bytes as an exact prefix', () => {
  const before = original();
  const after = appendExecutionPage(before, ({ page }) => {
    page.text('EXECUTION PAGE', 60, 80, { size: 16, font: 'Helvetica-Bold' });
  });
  assert.ok(after.length > before.length);
  assert.ok(after.subarray(0, before.length).equals(before), 'incremental means APPEND');
  const tail = after.toString('latin1');
  assert.ok(/\/Count 2/.test(tail));
  assert.ok(/\/Prev \d+/.test(tail));
});

test('a real reader sees one more page carrying the drawn marks', async () => {
  const after = appendExecutionPage(original(), ({ page, addImage }) => {
    page.text('EXECUTION PAGE', 60, 80, { size: 16, font: 'Helvetica-Bold' });
    page.text('Signed by Alma Signer', 60, 110);
    const ref = addImage(decodePng(tinyPng()));
    page.image(ref, 60, 130, 120, 40);
  });
  const doc = await open(after);
  if (!doc) return; // pdfjs unavailable - the structural test above still ran
  assert.equal(doc.numPages, 2);
  const page2 = await doc.getPage(2);
  const text = (await page2.getTextContent()).items.map((i) => i.str).join(' ');
  assert.ok(text.includes('EXECUTION PAGE'));
  assert.ok(text.includes('Alma Signer'));
  const page1 = await doc.getPage(1);
  const original1 = (await page1.getTextContent()).items.map((i) => i.str).join(' ');
  assert.ok(original1.includes('must not move'), 'page one is still the original');
});

test('what it cannot append to it refuses by name', () => {
  assert.throws(() => appendExecutionPage(Buffer.from('not a pdf'), () => {}), Unsupported);
  // A fake pointing startxref at an xref STREAM object instead of a table.
  const streamish = Buffer.from('%PDF-1.5\n1 0 obj\n<< /Type /XRef >>\nstream\nx\nendstream\nendobj\nstartxref\n9\n%%EOF\n', 'latin1');
  assert.throws(() => appendExecutionPage(streamish, () => {}), /cross-reference stream/);
});

// ── overlays: drawing ON the original's pages ────────────────────────────

const { overlayPages, describePages } = require('../src/incremental');

test('an overlay keeps the original bytes as an exact prefix and redefines only the page', () => {
  const before = original();
  const after = overlayPages(before, {
    1: ({ page, width, height }) => {
      page.text('Ada Signer', 320, 700, { size: 14, font: 'Helvetica-Oblique' });
      page.rect(60, 60, width - 120, 20, { stroke: '#cccccc' });
      assert.ok(height > width, 'A4 portrait as the page was written');
    },
  });
  assert.ok(after.length > before.length);
  assert.ok(after.subarray(0, before.length).equals(before), 'incremental means APPEND');
  const tail = after.toString('latin1').slice(before.length);
  assert.ok(/\/Contents \[\d+ 0 R \d+ 0 R \d+ 0 R\]/.test(tail), 'q, the original stream, then Q + overlay');
  assert.ok(/\/RSF1 \d+ 0 R/.test(tail), 'our font merged into the resources under a non-clashing name');
  assert.ok(!/\/Count 2/.test(tail), 'no page was added');
});

test('a real reader sees the overlay AND the original text on the same page', async () => {
  const after = overlayPages(original(), {
    1: ({ page, addImage }) => {
      page.text('Ada Signer', 320, 700, { size: 14, font: 'Helvetica-Oblique' });
      const ref = addImage(decodePng(tinyPng()));
      page.image(ref, 320, 640, 120, 40);
    },
  });
  const doc = await open(after);
  if (!doc) return;
  assert.equal(doc.numPages, 1);
  const page1 = await doc.getPage(1);
  const text = (await page1.getTextContent()).items.map((i) => i.str).join(' ');
  assert.ok(text.includes('must not move'), 'the original content still renders');
  assert.ok(text.includes('Ada Signer'), 'the overlay renders on the same page');
});

test('overlay then execution page chain through /Prev and both survive', async () => {
  const before = original();
  const overlaid = overlayPages(before, { 1: ({ page }) => page.text('Ada Signer', 320, 700) });
  const after = appendExecutionPage(overlaid, ({ page }) => page.text('EXECUTION PAGE', 60, 80));
  assert.ok(after.subarray(0, before.length).equals(before));
  assert.ok(after.subarray(0, overlaid.length).equals(overlaid));
  const doc = await open(after);
  if (!doc) return;
  assert.equal(doc.numPages, 2);
  const text1 = (await (await doc.getPage(1)).getTextContent()).items.map((i) => i.str).join(' ');
  const text2 = (await (await doc.getPage(2)).getTextContent()).items.map((i) => i.str).join(' ');
  assert.ok(text1.includes('Ada Signer'));
  assert.ok(text2.includes('EXECUTION PAGE'));
});

test('overlays refuse what they cannot place, by name', () => {
  assert.throws(() => overlayPages(original(), { 3: ({ page }) => page.text('x', 1, 1) }), /page 3 does not exist/);
  assert.throws(() => overlayPages(Buffer.from('not a pdf'), { 1: () => {} }), Unsupported);
  const untouched = overlayPages(original(), {});
  assert.ok(Buffer.isBuffer(untouched));
});

test('describePages reports count and size', () => {
  const pages = describePages(original());
  assert.equal(pages.length, 1);
  assert.equal(pages[0].page, 1);
  assert.ok(pages[0].width > 500 && pages[0].height > 800);
  assert.equal(pages[0].rotate, 0);
});
