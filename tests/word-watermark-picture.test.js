// Design → Watermark → Picture watermark: a picture behind every page, as
// Word keeps one — a VML picture named WordPictureWatermark in the default
// header, washed out, centred on the margins — read back, drawn in the PDF,
// and sent to the window as a URL held once, not with every edit.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';
import { renderPdf } from '@rutba/doc-view/export/pdf';
import { gradientPng } from '../apps/desktop/main/sample-picture.js';
import { createDocumentService } from '../apps/desktop/main/documents.js';

const doc = () => openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Behind these words.' }] }));
const headerXml = (view) => {
  const d = view.doc.doc;
  return d.pkg.text(d.headerFooters().headers.default.part);
};

test('a picture watermark is written in the header as Word writes one, and read back', () => {
  const view = doc();
  view.setPictureWatermark(gradientPng(1200, 600, [30, 80, 160], [200, 220, 240]), { contentType: 'image/png', widthPx: 1200, heightPx: 600 });
  const xml = headerXml(view);
  assert.match(xml, /<v:shape id="WordPictureWatermark1"[^>]*type="#_x0000_t75"[^>]*style="position:absolute;[^"]*mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin"/);
  assert.match(xml, /<v:imagedata r:id="(rId\d+)" o:title="Watermark" gain="19661f" blacklevel="22938f"\/>/, 'washed out as Word washes one out');
  assert.match(xml, /^[\s\S]*<w:hdr\b[^>]*xmlns:r="/, 'the relationships namespace declared');
  const rid = /<v:imagedata r:id="(rId\d+)"/.exec(xml)[1];
  const d = view.doc.doc;
  const target = d.pkg.rels(d.headerFooters().headers.default.part).find((r) => r.Id === rid)?.Target;
  assert.match(target, /^media\/watermark\d+\.png$/);
  const width = Number(/width:([\d.]+)pt/.exec(xml)[1]);
  assert.ok(width < 1200 * 0.75 && width > 300, `shrunk to the room inside the margins: ${width}pt`);
  const mark = d.headerFooters().watermark;
  assert.equal(mark.picturePart, 'word/media/watermark1.png');
  assert.equal(mark.washout, true);
  assert.ok(mark.widthPx > 400 && mark.heightPx > 200 && Math.abs(mark.widthPx / mark.heightPx - 2) < 0.05, 'its shape kept');
  // Reopened from the saved file, it is still there.
  assert.equal(openDocx(view.save()).doc.doc.headerFooters().watermark.picturePart, 'word/media/watermark1.png');
});

test('words replace a picture watermark, a picture replaces words, and Remove takes either off', () => {
  const view = doc();
  view.setWatermark('DRAFT');
  view.setPictureWatermark(gradientPng(200, 100, [0, 0, 0], [255, 255, 255]), { widthPx: 200, heightPx: 100 });
  assert.doesNotMatch(headerXml(view), /<v:textpath\b/, 'the words gone');
  assert.equal((headerXml(view).match(/WordPictureWatermark/g) || []).length, 1);
  view.setWatermark('CONFIDENTIAL');
  assert.doesNotMatch(headerXml(view), /WordPictureWatermark/, 'the picture gone');
  assert.equal(view.doc.doc.headerFooters().watermark.text, 'CONFIDENTIAL');
  view.setPictureWatermark(gradientPng(200, 100, [0, 0, 0], [255, 255, 255]), { widthPx: 200, heightPx: 100, washout: false });
  assert.doesNotMatch(headerXml(view), /gain=/, 'not washed out when asked not to be');
  view.setWatermark(null);
  assert.equal(view.doc.doc.headerFooters().watermark ?? null, null);
  view.undo();
  assert.ok(view.doc.doc.headerFooters().watermark?.picturePart, 'one undo puts it back');
});

test('the PDF draws the picture behind the page', () => {
  const view = doc();
  view.setPictureWatermark(gradientPng(300, 150, [200, 40, 40], [250, 250, 250]), { widthPx: 300, heightPx: 150 });
  const { buffer } = renderPdf(view, { title: 'Watermark' });
  assert.match(Buffer.from(buffer).toString('latin1'), /\/Subtype\s*\/Image/, 'an image in the page');
  const plain = renderPdf(doc(), { title: 'Plain' });
  assert.doesNotMatch(Buffer.from(plain.buffer).toString('latin1'), /\/Subtype\s*\/Image/);
});

test('the window is sent the picture as one held URL, not with every edit', async () => {
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-wm-')), 'wm.docx');
  fs.writeFileSync(file, buildDocx({ styles: true, paragraphs: [{ text: 'One' }, { text: 'Two' }] }));
  let held = 0;
  const docs = createDocumentService({ holdBlob: () => { held += 1; return { id: `b${held}`, url: `rutba://blob/b${held}` }; }, releaseBlob: () => {} });
  const { id } = await docs.open({ path: file });
  const png = gradientPng(240, 120, [10, 120, 60], [240, 250, 240]);
  const first = await docs.apply({ id, ops: [{ op: 'setPictureWatermark', data: png }] });
  const bands = (first.patch || first.model).bands;
  assert.equal(bands.watermark.picture, 'rutba://blob/b1');
  assert.equal(bands.watermark.washout, true);
  assert.ok(bands.watermark.widthPx > 100);
  const again = await docs.apply({ id, ops: [{ op: 'setSelection', anchor: { block: 1, offset: 0 }, focus: { block: 1, offset: 0 } }, { op: 'insertText', text: 'x' }] });
  assert.equal((again.patch || again.model).bands.watermark.picture, 'rutba://blob/b1', 'the same URL: held once');
  assert.equal(held, 1);
});
