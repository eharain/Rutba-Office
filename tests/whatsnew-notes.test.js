// The release notes, read for the What's new dialog.

import test from 'node:test';
import assert from 'node:assert/strict';
import { parseNotes, runsOf } from '../apps/desktop/renderer/whatsnew/notes.js';

const NOTES = `A scroll draws at once, and an update says so.

## Worksheets: a large workbook scrolls

- **A frame is 30 ms**, from a third of a second — the \`region\` went.
  Ten rows are drawn beyond the edge.
- **The rows above stay** where they are.

## The launcher

The list scrolls. It used to be clipped.

Two paragraphs, one block each.
`;

test('headings, bullets with their continuation lines, and paragraphs come out as blocks', () => {
  const blocks = parseNotes(NOTES);
  assert.deepEqual(blocks.map((b) => b.kind), ['p', 'h', 'li', 'li', 'h', 'p', 'p']);
  assert.equal(blocks[0].runs[0].text, 'A scroll draws at once, and an update says so.');
  assert.equal(blocks[1].text, 'Worksheets: a large workbook scrolls');
  assert.equal(blocks[1].level, 2);
  assert.equal(blocks[2].text, '**A frame is 30 ms**, from a third of a second — the `region` went. Ten rows are drawn beyond the edge.', 'the indented line continues the bullet');
  assert.deepEqual(blocks[2].runs, [
    { text: 'A frame is 30 ms', bold: true },
    { text: ', from a third of a second — the ' },
    { text: 'region', code: true },
    { text: ' went. Ten rows are drawn beyond the edge.' },
  ]);
  assert.equal(blocks[5].runs.map((r) => r.text).join(''), 'The list scrolls. It used to be clipped.');
  assert.equal(blocks[6].runs[0].text, 'Two paragraphs, one block each.');
});

test('a line with no marks is one run, and CRLF notes read the same', () => {
  assert.deepEqual(runsOf('plain words'), [{ text: 'plain words' }]);
  assert.deepEqual(parseNotes('a\r\n\r\n## b\r\n'), [{ kind: 'p', runs: [{ text: 'a' }] }, { kind: 'h', level: 2, text: 'b' }]);
  assert.deepEqual(parseNotes(''), []);
  assert.deepEqual(parseNotes(null), []);
});
