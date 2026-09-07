/**
 * The old cell paint gaps close: list labels and pictures in cell paragraphs.
 *
 * Cells have no fragments, so their labels and images could not ride the
 * pagination path the body uses. The frame now carries the full label map
 * (computed on the same counters pagination uses, so body and cells never
 * disagree about which item is third) and cell paragraphs with pictures
 * serialise them for the painter.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { OoxmlPackage } from '@rutba/ooxml';
import { buildDocx } from '@rutba/ooxml/build';

const PNG_1x1 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

test('a list in a table cell gets its labels, counted with the body', () => {
  const view = openDocx(buildDocx({ paragraphs: ['Agenda', ''] }));
  // a numbered item in the BODY first, so the cell items continue its list
  view.collapseTo({ block: 0, offset: 0 });
  view.setParagraphFormat({ list: 'number' });
  view.collapseTo({ block: 1, offset: 0 });
  view.insertTable({ rows: 2, cols: 1 });
  const cellA = view.focus.block;
  view.insertText('First cell item');
  view.setParagraphFormat({ list: 'number' });
  view.tabCell({});
  view.insertText('Second cell item');
  view.setParagraphFormat({ list: 'number' });

  const frame = view.render();
  assert.ok(frame.listLabels, 'the frame carries the label map');
  assert.equal(frame.listLabels[0].label, '1.', 'the body item leads');
  assert.equal(frame.listLabels[cellA].label, '2.', 'the first cell item continues the count');
  assert.equal(frame.listLabels[cellA + 1].label, '3.', 'and the second follows it');
});

test('a bullet in a cell is a bullet, not silence', () => {
  const view = openDocx(buildDocx({ paragraphs: ['', ''] }));
  view.collapseTo({ block: 0, offset: 0 });
  view.insertTable({ rows: 1, cols: 1 });
  const cell = view.focus.block;
  view.insertText('point one');
  view.setParagraphFormat({ list: 'bullet' });
  const frame = view.render();
  assert.equal(frame.listLabels[cell].label, '•');
  assert.equal(frame.listLabels[cell].bullet, true);
});

test('a picture in a cell paragraph reaches the frame as a data URI', () => {
  const view = openDocx(buildDocx({ paragraphs: ['', ''] }));
  view.collapseTo({ block: 0, offset: 0 });
  view.insertTable({ rows: 1, cols: 1 });
  const cell = view.focus.block;
  view.insertText('Logo cell');
  // insertImage refuses in cells (its own paragraph would not paint before);
  // graft the drawing into the cell paragraph the way a real file carries one.
  const pkg = OoxmlPackage.read(view.save());
  const media = 'word/media/rutba1.png';
  pkg.addPart(media, Buffer.from(PNG_1x1, 'base64'), 'image/png');
  pkg.addPart('word/_rels/document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId77" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/rutba1.png"/>' +
    '</Relationships>');
  pkg.write_('word/document.xml', pkg.text('word/document.xml').replace(
    'Logo cell</w:t></w:r>',
    'Logo cell</w:t></w:r><w:r><w:drawing><wp:inline>'
    + '<wp:extent cx="190500" cy="190500"/><wp:docPr id="9" name="logo"/>'
    + '<a:graphic><a:graphicData><pic:pic><pic:blipFill>'
    + '<a:blip r:embed="rId77"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>'
    + '</wp:inline></w:drawing></w:r>',
  ));
  const reopened = openDocx(pkg.write());

  const block = reopened.render().blocks.find((b) => b.text === 'Logo cell');
  assert.ok(block.container, 'it is a cell paragraph');
  assert.equal(block.images?.length, 1, 'the picture rides the frame');
  assert.ok(block.images[0].href.startsWith('data:image/png;base64,'), 'resolved to bytes the painter can draw');

  const plain = reopened.render().blocks.find((b) => b.container === null);
  assert.ok(!('images' in plain), 'prose paragraphs carry no ballast — theirs ride the fragments');
});
