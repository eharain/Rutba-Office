/**
 * A Markdown file opened in Documents and saved again (2.18.0, A3.1).
 *
 * In 2.17.1 the round trip changed a file nobody had edited: a raw HTML
 * block directly above a code block lost the blank line between them (and
 * the code block broke on GitHub), tables were padded anew, reference links
 * became inline ones and a definition list was joined into one line. Task
 * boxes were the text "☐" after a visible "- ". Now what was not edited goes
 * back as it was written, and a task's box is a check box a person clicks.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { parseMarkdown, serializeMarkdown } from '../packages/office-formats/src/markdown.js';
import { markdownToParagraphs, paragraphsToMarkdown } from '../apps/desktop/main/markdown-bridge.js';
import { buildDocx } from '@rutba/ooxml/build';
import { openDocx } from '@rutba/doc-view/backends/ooxml';

const README = `# Larkspur stock

<details>
<summary>How the list is kept</summary>
</details>

\`\`\`js
const shelf = 'A3';
\`\`\`

| Item  | Qty |
|-------|----:|
| Tea   |   2 |
| Bread |  10 |

See [the supplier list][suppliers] and [orders].

Reorder level
: The count below which an item is ordered again.

- [ ] Order milk
- [x] Pay the rent
- A plain item

[suppliers]: https://example.com/suppliers
[orders]: https://example.com/orders
`;

/** Open the way Documents opens a .md: the blocks become a document. */
function open(source) {
  const md = parseMarkdown(source);
  const view = openDocx(buildDocx({ paragraphs: markdownToParagraphs(md.blocks), styles: true }));
  return { view, meta: md.meta };
}

/** Save the way Documents saves one back. */
function save({ view, meta }) {
  const frame = view.render();
  return paragraphsToMarkdown(frame.blocks, meta, frame.listLabels);
}

test('a README opened and saved with no edit comes back byte for byte', () => {
  assert.equal(save(open(README)), README);
});

test('a raw HTML block directly above a code block keeps the blank line between them', () => {
  const source = '<div align="center">\n<img src="logo.png">\n</div>\n\n```sh\nnpm install\n```\n';
  assert.equal(save(open(source)), source);
  // And two code blocks one after the other stay two.
  const two = '```js\na()\n```\n\n```js\nb()\n```\n';
  assert.equal(save(open(two)), two);
});

test('a table is written back as it was written until a cell changes', () => {
  const doc = open(README);
  const out = save(doc);
  assert.ok(out.includes('| Item  | Qty |\n|-------|----:|\n| Tea   |   2 |'), 'padding and rule as written');
  // A cell edited: the table is written afresh, its alignment kept.
  const cell = doc.view.render().blocks.findIndex((b) => b.text === 'Tea');
  doc.view.collapseTo({ block: cell, offset: 3 });
  doc.view.insertText(' bags');
  const edited = save(doc);
  assert.match(edited, /\| Tea bags \| 2 +\|/);
  assert.match(edited, /\| -+ \| -+: \|/);
});

test('reference links stay reference links, in the form they were written', () => {
  const source = 'Read [the guide][g], [Docs][] and [home].\n\n[g]: https://example.com/guide\n[docs]: https://example.com/docs\n[home]: https://example.com\n';
  assert.equal(save(open(source)), source);
  assert.equal(serializeMarkdown(parseMarkdown(source).blocks, parseMarkdown(source).meta), source);
});

test('a definition list is kept line for line, not joined into one', () => {
  const source = 'Term\n: The meaning.\n: A second meaning.\n\nAnother term\n: Its meaning.\n';
  const doc = open(source);
  const blocks = doc.view.render().blocks;
  assert.deepEqual(blocks.map((b) => b.text), ['Term', ': The meaning.', ': A second meaning.', 'Another term', ': Its meaning.']);
  assert.equal(save(doc), source);
});

test('a task box is a check box a person clicks, its "- " not shown, and the tick is saved', () => {
  const doc = open('- [ ] Order milk\n- [x] Pay the rent\n');
  const blocks = doc.view.render().blocks;
  assert.ok(!blocks[0].text.startsWith('- '), 'no marker on the page: ' + JSON.stringify(blocks[0].text));
  const box = blocks[0].runs.find((r) => r.control);
  assert.equal(box?.control?.kind, 'checkBox');
  assert.equal(box.control.checked, false);
  assert.equal(blocks[1].runs.find((r) => r.control)?.control?.checked, true);
  // Ticked, as a click on the box does.
  doc.view.setControlValue({ block: 0, offset: 0, value: true });
  assert.equal(save(doc), '- [x] Order milk\n- [x] Pay the rent\n');
  // A new item typed after a task, in the same list, is a list item still.
  doc.view.collapseTo({ block: 1, offset: doc.view.block(1).text.length });
  doc.view.splitParagraph();
  doc.view.insertText('Ring the bakery');
  assert.equal(save(doc), '- [x] Order milk\n- [x] Pay the rent\n- Ring the bakery\n');
});
