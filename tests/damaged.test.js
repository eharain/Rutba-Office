/**
 * Files that have been damaged, and the answers they must get.
 *
 * A downloaded file is eventually a truncated one, and a file on a failing
 * disk is eventually a file with a flipped byte. Neither may reach a person
 * as a stack from inside a parser: what a window can draw is either the
 * document or a sentence saying why not.
 *
 * The two cases pinned here were found by `node tools/fuzz-open.js`, which
 * damages good files at random and opens them until something answers with
 * something other than a sentence. They are written out longhand because a
 * fuzz run is not a test: it finds cases, and then they belong here.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx, buildDocx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml/package';
import { Workbook } from '@rutba/ooxml';
import { SheetView } from '@rutba/sheet-view';

test('a workbook that lists no sheets is refused, not carried to the first frame', () => {
  // A flipped byte in the part that lists the sheets is enough. The view took
  // an undefined active sheet all the way to render() and threw there —
  // "Cannot read properties of undefined (reading 'viewport')" — which is a
  // window that never draws rather than a file that could not be read.
  const pkg = OoxmlPackage.read(buildXlsx({ sheets: [{ name: 'S', rows: [[1, 2]] }] }));
  const main = pkg.mainDocument();
  pkg.write_(main, Buffer.from(pkg.text(main).replace(/<sheets>[\s\S]*?<\/sheets>/, '<sheets/>'), 'utf8'));
  const bytes = pkg.write();

  assert.deepEqual(Workbook.open(bytes).sheetNames(), [], 'the workbook itself says there are none');
  assert.throws(() => new SheetView(bytes), /not a workbook: it lists no sheets/);
});

test('a relationship with no target is passed over, or said in words', () => {
  // The Target attribute damaged away: `resolveTarget` called startsWith on
  // undefined, four frames below anything that knew what file it was.
  assert.throws(() => OoxmlPackage.resolveTarget('', undefined), /a relationship is missing its target/);
  assert.throws(() => OoxmlPackage.resolveTarget('word/document.xml', ''), /a relationship is missing its target/);

  // A package whose main relationship is damaged still opens, because the
  // part is where every file Office writes keeps it.
  const pkg = OoxmlPackage.read(buildDocx({ paragraphs: ['One.'] }));
  const rels = pkg.text('_rels/.rels');
  pkg.write_('_rels/.rels', Buffer.from(rels.replace(/Target="[^"]*"/, ''), 'utf8'));
  assert.equal(pkg.mainDocument(), 'word/document.xml', 'the known part answers instead');

  // And when there is no such part either, the refusal says so.
  const bare = OoxmlPackage.read(buildDocx({ paragraphs: ['One.'] }));
  bare.write_('_rels/.rels', Buffer.from(rels.replace(/Target="[^"]*"/, ''), 'utf8'));
  bare.removePart('word/document.xml');
  assert.throws(() => bare.mainDocument(), /missing/);
});

test('a deck whose slide will not unzip is refused in a sentence, with no library words and no session left open', async () => {
  // Found by the gate's fuzz run on 11 Nov 2026: a slide is unzipped only when the first model asks for it,
  // after the session was made and outside the open's sentences, so zlib's own words reached the window.
  const fs = await import('node:fs');
  const os = await import('node:os');
  const path = await import('node:path');
  const { buildPptx } = await import('@rutba/presentation');
  const { createDocumentService } = await import('../apps/desktop/main/documents.js');
  const { judge } = await import('../tools/fuzz-open.js');
  const bytes = Buffer.from(buildPptx({ title: 'Deck', slides: [{ layout: 'title', title: 'Rutba' }] }));
  // The slide's deflated bytes, every one 0xFF: a block of the reserved type, which no inflater reads.
  let at = 0;
  let done = false;
  while ((at = bytes.indexOf(Buffer.from([0x50, 0x4b, 0x03, 0x04]), at)) >= 0) {
    const method = bytes.readUInt16LE(at + 8);
    const size = bytes.readUInt32LE(at + 18);
    const nameLen = bytes.readUInt16LE(at + 26);
    const extraLen = bytes.readUInt16LE(at + 28);
    const name = bytes.toString('utf8', at + 30, at + 30 + nameLen);
    const data = at + 30 + nameLen + extraLen;
    if (name === 'ppt/slides/slide1.xml' && method === 8) { bytes.fill(0xff, data, data + size); done = true; break; }
    at = data + size;
  }
  assert.ok(done, 'the slide is a deflated part');
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'damaged-deck-'));
  const file = path.join(dir, 'deck.pptx');
  fs.writeFileSync(file, bytes);
  const service = createDocumentService({ holdBlob: () => ({ url: 'blob://held' }) });
  try {
    let thrown = null;
    try { service.open({ path: file }); } catch (err) { thrown = err; }
    assert.ok(thrown, 'the deck is refused');
    assert.match(thrown.message, /^deck\.pptx is damaged: part of it cannot be unpacked\./);
    assert.equal(judge(thrown).outcome, 'refused', thrown.message);
    assert.ok(thrown.cause, 'the library\'s words are kept in the cause, for a bug report');
    assert.deepEqual(service.sessions(), [], 'no session is left for a file that did not open');
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
