// What File → Print starts from, kept in the file. A deck keeps its print's
// choices where PowerPoint keeps them (`p:prnPr` in presProps); a document's
// paper, orientation and margins are its own section's, which the print
// reads and changes, as Word's does. Both used to start from this build's
// defaults every time.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';
import { docPageSetup } from '../apps/desktop/main/documents.js';

test('a deck keeps what it prints in presProps, ahead of the show\'s settings, as PowerPoint does', () => {
  const d = Deck.open(buildPptx({ title: 'Print', slides: [{ layout: 'blank' }, { layout: 'blank' }] }));
  assert.equal(d.printSettings(), null, 'none kept: a print starts from its own defaults');
  d.setShowSettings({ loop: true });
  assert.deepEqual(d.setPrintSettings({ layout: 'handout', perPage: 3, frame: true }), { layout: 'handout', perPage: 3, frame: true });
  const part = 'ppt/presProps.xml';
  const xml = d.pkg.text(part);
  assert.match(xml, /<p:prnPr prnWhat="handouts3" frameSlides="1"\/><p:showPr\b/, 'prnPr ahead of showPr, as the schema orders them');
  assert.deepEqual(Deck.open(d.save()).printSettings(), { layout: 'handout', perPage: 3, frame: true }, 'read back from the saved file');
  // PowerPoint's own choices this print does not ask about are kept.
  d.pkg.write_(part, d.pkg.text(part).replace('<p:prnPr prnWhat="handouts3" frameSlides="1"/>', '<p:prnPr prnWhat="handouts3" clrMode="gray" hiddenSlides="1" frameSlides="1"/>'));
  d.setPrintSettings({ layout: 'notes', frame: false });
  assert.match(d.pkg.text(part), /<p:prnPr prnWhat="notes" clrMode="gray" hiddenSlides="1" frameSlides="0"\/>/);
  assert.equal((d.pkg.text(part).match(/<p:prnPr\b/g) || []).length, 1, 'one, replaced');
  assert.deepEqual(d.setPrintSettings({ layout: 'slides', perPage: 7 }), { layout: 'slides', perPage: 6, frame: false });
});

test('a deck with no presProps gets one for its print, related from the presentation', () => {
  const d = Deck.open(buildPptx({ title: 'Print', slides: [{ layout: 'blank' }] }));
  if (d.pkg.has('ppt/presProps.xml')) return; // the builder made one; the test above covers it
  d.setPrintSettings({ layout: 'notes' });
  const back = Deck.open(d.save());
  assert.equal(back.printSettings().layout, 'notes');
});

test('a document\'s page as File → Print shows it: its own paper, orientation and margins, changed in the document', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Page' }] }));
  const first = docPageSetup(view);
  assert.equal(first.paper, 'A4');
  assert.equal(first.orientation, 'portrait');
  view.setPageSetup({ size: 'Letter', orientation: 'landscape', margins: { top: 720, right: 720, bottom: 720, left: 720 } });
  assert.deepEqual(docPageSetup(view), { paper: 'Letter', orientation: 'landscape', margins: { top: 12.7, right: 12.7, bottom: 12.7, left: 12.7 } });
  view.setPageSetup({ size: 'A3' });
  assert.equal(docPageSetup(view).paper, 'A3', 'A3 and A5 as well as A4, Letter and Legal');
  assert.match(view.doc.doc.xml, /<w:pgSz w:w="23811" w:h="16838" w:orient="landscape"\/>/, 'A3 the wrong way round, as asked');
  assert.equal(docPageSetup(openDocx(view.save())).paper, 'A3', 'kept in the file');
});
