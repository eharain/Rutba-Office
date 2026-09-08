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
