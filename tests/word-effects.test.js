// Text effects — outline, shadow, glow: the engine's half of the Home tab's
// "A" button.
//
// Outline and shadow are the legacy toggle elements (`w:outline`, `w:shadow`):
// bare when on, absent when off, sitting at their own slot in rPr — after the
// strike toggles, before colour — rather than at the front the way b/i/u/s
// do. Glow is Word 2010's `w14:glow`, an extension element that always rides
// at the very end of rPr, its own namespace declared inline unless the
// document root already carries it.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';

const doc = (paragraphs) => openDocx(buildDocx({ styles: true, paragraphs }));
const rPrOf = (view, block, run = 0) => view.doc.doc.editParagraph(block).runs[run].rPr;

test('outline is written after colour and size, and off takes the element away again', () => {
  const view = doc([{ text: 'Hello there' }]);
  view.setSelection({ block: 0, offset: 0 }, { block: 0, offset: 5 });
  view.setRunFormat({ fontColour: 'FF0000', fontSize: 14 });
  view.setRunFormat({ outline: true });

  const rPr = rPrOf(view, 0);
  assert.match(rPr, /<w:outline\/>/, 'the outline element');
  assert.ok(rPr.indexOf('<w:outline/>') < rPr.indexOf('<w:color'), 'outline before the colour already there');
  assert.ok(rPr.indexOf('<w:outline/>') < rPr.indexOf('<w:sz'), 'outline before the size already there');
  assert.equal(view.doc.readRunProps(rPr).outline, true, 'read back');

  view.setRunFormat({ outline: false });
  assert.doesNotMatch(rPrOf(view, 0), /w:outline/, 'gone, not merely turned off with w:val="0"');
});

test('shadow likewise, and outline then shadow when both are on', () => {
  const view = doc([{ text: 'Hello there' }]);
  view.setSelection({ block: 0, offset: 0 }, { block: 0, offset: 5 });
  view.setRunFormat({ fontColour: '0000FF' });
  view.setRunFormat({ shadow: true });

  let rPr = rPrOf(view, 0);
  assert.match(rPr, /<w:shadow\/>/);
  assert.ok(rPr.indexOf('<w:shadow/>') < rPr.indexOf('<w:color'), 'shadow before the colour already there');
  assert.equal(view.doc.readRunProps(rPr).shadow, true);

  view.setRunFormat({ outline: true });
  rPr = rPrOf(view, 0);
  assert.ok(rPr.indexOf('<w:outline/>') < rPr.indexOf('<w:shadow/>'), 'outline before shadow');

  view.setRunFormat({ outline: false, shadow: false });
  assert.doesNotMatch(rPrOf(view, 0), /w:outline|w:shadow/, 'both gone');
});

test('glow sits at the very end of rPr, reads back, and null removes it', () => {
  const view = doc([{ text: 'Hello there' }]);
  view.setSelection({ block: 0, offset: 0 }, { block: 0, offset: 5 });
  view.setRunFormat({ fontColour: '00FF00', outline: true });
  view.setRunFormat({ glow: { colour: 'FFC000', radiusPt: 4 } });

  const rPr = rPrOf(view, 0);
  assert.equal(
    rPr,
    '<w:rPr><w:outline/><w:color w:val="00FF00"/>'
    + '<w14:glow xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:rad="50800">'
    + '<w14:srgbClr w14:val="FFC000"/></w14:glow></w:rPr>',
    'glow last, points turned into EMU, its namespace declared inline',
  );
  assert.deepEqual(view.doc.readRunProps(rPr).glow, { colour: 'FFC000', radiusPt: 4 });

  view.setRunFormat({ glow: null });
  assert.doesNotMatch(rPrOf(view, 0), /w14:glow/, 'gone');
  assert.equal(view.doc.readRunProps(rPrOf(view, 0)).glow, null);
});

test('a document whose root already declares xmlns:w14 gets glow written bare', () => {
  const view = doc([{ text: 'Hello there' }]);
  // As if Word itself had written the file: the namespace already sits on
  // <w:document>, so the run need not repeat it.
  view.doc.doc.xml = view.doc.doc.xml.replace(
    /<w:document\b([^>]*)>/,
    (m, attrs) => `<w:document${attrs} xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">`,
  );
  view.setSelection({ block: 0, offset: 0 }, { block: 0, offset: 5 });
  view.setRunFormat({ glow: { colour: 'FF0000', radiusPt: 2 } });

  const rPr = rPrOf(view, 0);
  assert.equal(rPr, '<w:rPr><w14:glow w14:rad="25400"><w14:srgbClr w14:val="FF0000"/></w14:glow></w:rPr>', 'no repeated namespace');
});

test('a glow with a scheme colour reads back with colour null and the radius kept', () => {
  const view = doc([{ text: 'Hello there' }]);
  view.setSelection({ block: 0, offset: 0 }, { block: 0, offset: 5 });
  view.setRunFormat({ glow: { colour: 'FFC000', radiusPt: 4 } });
  const withScheme = rPrOf(view, 0).replace(/<w14:srgbClr[^/]*\/>/, '<w14:schemeClr w14:val="accent1"/>');
  assert.deepEqual(view.doc.readRunProps(withScheme).glow, { colour: null, radiusPt: 4 });
});

test('all three survive a save and reopen, and render into the frame', () => {
  const view = doc([{ text: 'Hello there' }]);
  view.setSelection({ block: 0, offset: 0 }, { block: 0, offset: 5 });
  view.setRunFormat({ outline: true, shadow: true, glow: { colour: '4472C4', radiusPt: 4 } });

  const reopened = openDocx(view.save());
  const runs = reopened.render({ pages: false }).blocks[0].runs;
  assert.equal(runs[0].text, 'Hello');
  assert.equal(runs[0].outline, true);
  assert.equal(runs[0].shadow, true);
  assert.deepEqual(runs[0].glow, { colour: '4472C4', radiusPt: 4 });
  // The rest of the paragraph carries none of it.
  assert.ok(!runs[1]?.outline && !runs[1]?.shadow && !runs[1]?.glow);
});

test('clearFormat takes all three off, and formatAtCaret reports them at the caret', () => {
  const view = doc([{ text: 'Hello there' }]);
  view.setSelection({ block: 0, offset: 0 }, { block: 0, offset: 5 });
  view.setRunFormat({ outline: true, shadow: true, glow: { colour: 'FF0000', radiusPt: 4 } });

  view.collapseTo({ block: 0, offset: 2 });
  const at = view.formatAtCaret();
  assert.equal(at.outline, true, 'outline at the caret');
  assert.equal(at.shadow, true, 'shadow at the caret');
  assert.deepEqual(at.glow, { colour: 'FF0000', radiusPt: 4 }, 'glow at the caret');

  view.setSelection({ block: 0, offset: 0 }, { block: 0, offset: 5 });
  view.clearFormat();
  assert.equal(rPrOf(view, 0), null, 'clearFormat drops the whole rPr, effects included');

  view.collapseTo({ block: 0, offset: 2 });
  const cleared = view.formatAtCaret();
  assert.equal(cleared.outline, false);
  assert.equal(cleared.shadow, false);
  assert.equal(cleared.glow, null);
});

test('a plain document reports the three as off, never undefined', () => {
  const view = doc([{ text: 'plain' }]);
  const at = view.formatAtCaret();
  assert.equal(at.outline, false);
  assert.equal(at.shadow, false);
  assert.equal(at.glow, null);
});
