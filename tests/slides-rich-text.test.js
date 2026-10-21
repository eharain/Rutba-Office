// Presentations' box words as runs (slides/rich-text.js): a range formatted
// on its own, words retyped with the runs round them kept, the ribbon's state.
import test from 'node:test';
import assert from 'node:assert/strict';
import { plainOf, formatRange, withPlainText, formatState, replaceRange, formatParagraphsIn } from '../apps/desktop/renderer/apps/slides/rich-text.js';

test("a box's words: a range formatted alone, retyping keeps the runs, the ribbon reads the words chosen", () => {
  const box = [
    { align: 'left', runs: [{ text: 'Hello ', size: 24 }, { text: 'bold', size: 24, bold: true }, { text: ' world', size: 24 }] },
    { bullet: { type: 'char', char: '•' }, runs: [{ text: 'Second line', size: 20, color: '#C00000' }] },
  ];
  assert.equal(plainOf(box), 'Hello bold world\nSecond line');

  // Bold on "world" only.
  let b = formatRange(box, 11, 16, { bold: 'toggle' });
  assert.deepEqual(b[0].runs, [{ text: 'Hello ', size: 24 }, { text: 'bold', size: 24, bold: true }, { text: ' ', size: 24 }, { text: 'world', size: 24, bold: true }]);
  assert.deepEqual(b[1], box[1]);

  // Across paragraphs: "lo bold world\nSec" made red.
  b = formatRange(box, 3, 20, { color: '#0070C0' });
  assert.equal(b[0].runs[0].text, 'Hel');
  assert.equal(b[0].runs[0].color, undefined);
  assert.ok(b[0].runs.slice(1).every((r) => r.color === '#0070C0'));
  assert.deepEqual(b[1].runs.map((r) => [r.text, r.color]), [['Sec', '#0070C0'], ['ond line', '#C00000']]);

  // Toggle on a range that is all bold turns it off.
  b = formatRange(box, 6, 10, { bold: 'toggle' });
  assert.deepEqual(b[0].runs, [{ text: 'Hello bold world', size: 24 }]);

  // Retyping a word keeps the bold word bold.
  b = withPlainText(box, 'Hello bold planet\nSecond line');
  assert.deepEqual(b[0].runs, [{ text: 'Hello ', size: 24 }, { text: 'bold', size: 24, bold: true }, { text: ' planet', size: 24 }]);
  // A symbol typed after the bold word takes the bold word's look.
  b = withPlainText(box, 'Hello bold™ world\nSecond line');
  assert.deepEqual(b[0].runs[1], { text: 'bold™', size: 24, bold: true });
  // A new paragraph keeps the paragraph's settings.
  b = withPlainText(box, 'Hello bold world\nSecond line\nThird');
  assert.equal(b.length, 3);
  assert.deepEqual(b[2].bullet, { type: 'char', char: '•' });
  assert.equal(b[2].runs[0].color, '#C00000');
  // Joining two paragraphs.
  b = withPlainText(box, 'Hello bold worldSecond line');
  assert.equal(b.length, 1);
  assert.equal(plainOf(b), 'Hello bold worldSecond line');

  // The ribbon's state.
  assert.equal(formatState(box, 6, 10).bold, true);
  assert.equal(formatState(box, 0, 10).bold, false);
  assert.equal(formatState(box, 8, 8).bold, true);
  assert.equal(formatState(box, 17, 20).color, '#C00000');

  // Paragraph formatting in a range.
  b = formatParagraphsIn(box, 0, 3, { align: 'center' });
  assert.equal(b[0].align, 'center');
  assert.equal(b[1].align, undefined);

  // Line break inside a paragraph.
  b = replaceRange(box, 5, 6, '\v');
  assert.equal(plainOf(b), 'Hello\vbold world\nSecond line');
});
