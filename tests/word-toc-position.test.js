/**
 * A table of contents put in with the caret in a heading, or after one
 * (2.25.0). Each heading gets a bookmark as the table is built, which moved
 * the place the table was to go: it landed inside a heading, and the heading
 * lost its style and its words.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';

const paragraphs = () => [
  { text: 'First', style: 'Heading1' },
  'Words under it',
  { text: 'Second', style: 'Heading2' },
  'More words',
];

for (const at of [0, 1, 2, 3]) {
  test(`the table goes after paragraph ${at}, and every heading keeps its words and its style`, () => {
    const view = openDocx(buildDocx({ styles: true, paragraphs: paragraphs() }));
    view.collapseTo({ block: at, offset: 0 });
    view.insertTableOfContents({ levels: 3, pages: [1, 1] });
    const blocks = view.render().blocks;
    assert.deepEqual(blocks.filter((b) => /^Heading/.test(b.style || '')).map((b) => [b.style, b.text]), [['Heading1', 'First'], ['Heading2', 'Second']]);
    assert.deepEqual(blocks.filter((b) => !b.style && b.text).map((b) => b.text), ['Words under it', 'More words']);
    const entries = view.render().tableOfContents.entries.map((e) => e.text);
    assert.deepEqual(entries, ['First', 'Second']);
    const heading = blocks.findIndex((b) => b.style === 'Heading1');
    const first = blocks.findIndex((b) => /^TOC/.test(b.style || ''));
    assert.equal(first > heading, at >= 0, 'the table comes after the paragraph the caret was in');
    const reopened = openDocx(view.doc.doc.toBuffer ? view.doc.doc.toBuffer() : view.save());
    assert.deepEqual(reopened.render().blocks.filter((b) => /^Heading/.test(b.style || '')).map((b) => b.text), ['First', 'Second']);
  });
}

test('Update Table with a heading above the table and a new one below replaces the table and nothing else', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: paragraphs() }));
  view.collapseTo({ block: 1, offset: 0 });
  view.insertTableOfContents({ levels: 3, pages: [1, 1] });
  const last = view.render().blocks.length - 1;
  view.collapseTo({ block: last, offset: 0 });
  view.setParagraphStyle?.('Heading2');
  view.updateTableOfContents({ pages: [1, 1, 2] });
  const blocks = view.render().blocks;
  assert.equal(blocks[0].text, 'First');
  assert.equal(blocks[0].style, 'Heading1');
  assert.ok(blocks.some((b) => b.text === 'Words under it' && !b.style));
  assert.ok(blocks.some((b) => b.text === 'Second' && b.style === 'Heading2'));
  assert.equal(blocks.filter((b) => /^TOC/.test(b.style || '') && b.text.includes('First')).length, 1, 'one table, not two');
});
