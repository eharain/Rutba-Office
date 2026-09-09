/**
 * Rich Text, written.
 *
 * The installer tells Windows this suite is the EDITOR of `.rtf`, and opening
 * one and pressing Ctrl+S refused: a trap the operating system set on our
 * behalf. Writing it is reading it backwards, and the awkward parts are the
 * same ones — a control word eats the space that delimits it, a group ends
 * the formatting inside it, and anything above ASCII is a signed 16-bit word.
 *
 * Every check here goes out through the writer and back in through the reader
 * this suite already had, because a format written to a specification nobody
 * reads back is a format nobody can open.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { readRtf, writeRtf, rtfToText } from '@rutba/office-formats/rtf';

const textOf = (blocks) =>
  blocks
    .map((b) =>
      b.type === 'table'
        ? b.rows.map((r) => r.map((c) => (c.runs || []).map((x) => x.text).join('')).join('|')).join(' / ')
        : (b.runs || []).map((r) => r.text).join('')
    )
    .join(' ~ ');

test('words, and the formatting on them, survive the round trip', () => {
  const rtf = writeRtf({
    title: 'Report',
    blocks: [
      { type: 'heading', level: 1, runs: [{ text: 'Quarterly review' }] },
      { type: 'paragraph', align: 'center', runs: [{ text: 'Bold', bold: true }, { text: ' and normal, ' }, { text: 'italic', italic: true }] },
      { type: 'paragraph', align: 'right', runs: [{ text: 'Georgia, 14pt', size: 14, font: 'Georgia' }] },
      { type: 'paragraph', runs: [{ text: 'red', colour: '#cc0000' }, { text: ' and ' }, { text: 'underlined', underline: true }] },
    ],
  });
  const { blocks } = readRtf(rtf);
  assert.equal(blocks.length, 4);
  assert.equal(textOf(blocks), 'Quarterly review ~ Bold and normal, italic ~ Georgia, 14pt ~ red and underlined');

  // The space after a formatted run is the one a naive writer loses: a
  // control word is delimited by a space, and that space is not text.
  assert.match(textOf(blocks), /Bold and normal/);

  assert.ok(blocks[1].runs.some((r) => r.bold), 'bold');
  assert.ok(blocks[1].runs.some((r) => r.italic), 'italic');
  assert.ok(blocks[3].runs.some((r) => r.underline), 'underline');
  assert.equal(blocks[1].align, 'center');
  assert.equal(blocks[2].align, 'right');
  assert.deepEqual(blocks[2].runs.map((r) => r.size), [14], 'type is written in half-points and read back in points');
});

test('everything that has to be escaped is', () => {
  const awkward = 'Braces {here} and \\ backslash, a\ttab, café — 日本語, and an emoji 😀';
  const { blocks } = readRtf(writeRtf({ blocks: [{ type: 'paragraph', runs: [{ text: awkward }] }] }));
  assert.equal(textOf(blocks), awkward);
});

test('a table comes back as one table with its rows', () => {
  // Wrapping a row in a group, or opening each cell with \pard, both read
  // back as something else: one empty cell, and one table per row.
  const rtf = writeRtf({
    blocks: [
      { type: 'paragraph', runs: [{ text: 'Before' }] },
      { type: 'table', rows: [[{ text: 'Region' }, { text: 'Total' }], [{ text: 'North' }, { text: '12' }], [{ text: 'South' }, { text: '8' }]] },
      { type: 'paragraph', runs: [{ text: 'After' }] },
    ],
  });
  const { blocks, tables } = readRtf(rtf);
  assert.equal(tables, 1, 'one table, not one per row');
  const table = blocks.find((b) => b.type === 'table');
  assert.ok(table, 'and it is a table');
  assert.equal(table.rows.length, 3);
  assert.deepEqual(table.rows[1].map((c) => c.runs.map((r) => r.text).join('')), ['North', '12']);
  assert.match(textOf(blocks), /^Before/);
  assert.match(textOf(blocks), /After$/);
});

test('what is written can be written again, unchanged', () => {
  // A format that drifts on every save is a format that loses a document
  // slowly rather than at once.
  const blocks = [
    { type: 'heading', level: 2, runs: [{ text: 'Terms' }] },
    { type: 'paragraph', runs: [{ text: 'The supplier ' }, { text: 'shall', bold: true }, { text: ' deliver.' }] },
    { type: 'table', rows: [[{ text: 'a' }, { text: 'b' }]] },
  ];
  const once = writeRtf({ blocks });
  const twice = writeRtf({ blocks: readRtf(once).blocks });
  assert.equal(rtfToText(twice), rtfToText(once));
  assert.equal(textOf(readRtf(twice).blocks), textOf(readRtf(once).blocks));
});

test('an empty document is still a valid file', () => {
  const rtf = writeRtf({});
  assert.match(rtf, /^\{\\rtf1/);
  assert.match(rtf, /\}$/);
  assert.equal(rtfToText(rtf).trim(), '');
});
