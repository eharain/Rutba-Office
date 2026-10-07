// Delete on a selected shape, chart or picture takes it out.
//
// Insert → Shapes and Insert → Chart put a drawing in a paragraph of its
// own, painted on the page among the paragraph's pictures. Delete sent it
// to the engine as a picture by its place, and the engine — counting only
// drawings with a picture in them — answered "no picture 0 in paragraph 1"
// and left it there. A selected drawing is now taken out by its own id,
// whatever it is, and one in the line goes with the paragraph it leaves
// empty, as a picture does.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';

const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==', 'base64');

/** The id the page knows a block's painted drawing by — what a click picks. */
const pickedId = (view, block) => view.block(block).images?.[0]?.id;

test('a shape inserted from the ribbon is deleted by its id, and its empty paragraph with it', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'before' }, { text: 'after' }] }));
  view.collapseTo({ block: 0, offset: 0 });
  view.insertShape({ preset: 'rect', widthPx: 120, heightPx: 80 });
  assert.equal(view.blocks.length, 3);
  const id = pickedId(view, 1);
  assert.ok(id != null, 'the painted shape carries its drawing id');
  assert.equal(view.block(1).images[0].kind, 'shape');
  assert.throws(() => view.removeImage({ block: 1, image: 0 }), /no picture/, 'as a picture it was refused');

  view.removeDrawing([id], { emptyParagraph: true });
  assert.deepEqual(view.blocks.map((b) => b.text), ['before', 'after'], 'the shape and its paragraph are gone');
  assert.ok(!/<wps:wsp\b/.test(view.doc.doc.xml), 'no shape left in the file');
  view.undo();
  assert.equal(view.blocks.length, 3, 'one undo brings it back');
});

test('a chart in the line is deleted the same way', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Sales' }], tables: [] }));
  // A table to chart from, then the chart after it.
  view.collapseTo({ block: 0, offset: 5 });
  view.insertTable({ rows: 2, cols: 2 });
  const cells = view.blocks.map((b, i) => [b, i]).filter(([b]) => b.container);
  const fill = ['Region', 'Q1', 'North', '12'];
  cells.forEach(([, i], k) => { view.collapseTo({ block: i, offset: 0 }); view.insertText(fill[k]); });
  view.collapseTo({ block: cells[0][1], offset: 0 });
  view.insertChart({ kind: 'column' });
  const at = view.blocks.findIndex((b) => b.images?.some((img) => img.kind === 'chart'));
  assert.ok(at >= 0, 'the chart is painted');
  const id = view.block(at).images.find((img) => img.kind === 'chart').id;
  view.removeDrawing([id], { emptyParagraph: true });
  assert.ok(!view.blocks.some((b) => b.images?.some((img) => img.kind === 'chart')), 'the chart is gone');
  assert.ok(!/<c:chart\b/.test(view.doc.doc.xml));
});

test('a picture is deleted by its id too, and a paragraph of words keeps its words', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'one' }, { text: 'two' }] }));
  view.collapseTo({ block: 0, offset: 0 });
  view.insertImage({ name: 'dot.png', contentType: 'image/png', data: PNG, widthPx: 20, heightPx: 20 });
  const id = pickedId(view, 1);
  view.removeDrawing([id], { emptyParagraph: true });
  assert.deepEqual(view.blocks.map((b) => b.text), ['one', 'two']);

  // Without `emptyParagraph` (a floating drawing) the paragraph stays.
  view.collapseTo({ block: 0, offset: 0 });
  view.insertImage({ name: 'dot.png', contentType: 'image/png', data: PNG, widthPx: 20, heightPx: 20 });
  view.removeDrawing([pickedId(view, 1)]);
  assert.equal(view.blocks.length, 3, 'the paragraph it was anchored in stays');
});

test('the body\'s last paragraph is emptied, not removed', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: '' }] }));
  view.collapseTo({ block: 0, offset: 0 });
  view.insertShape({ preset: 'ellipse', widthPx: 60, heightPx: 60 });
  // The shape went in after the empty paragraph: take the empty one away first.
  view.collapseTo({ block: 0, offset: 0 });
  view.deleteForward();
  assert.equal(view.blocks.length, 1, 'only the shape\'s paragraph is left');
  view.removeDrawing([pickedId(view, 0)], { emptyParagraph: true });
  assert.equal(view.blocks.length, 1, 'a body keeps a paragraph');
  assert.ok(!view.block(0).images?.length);
});
