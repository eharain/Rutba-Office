// Where a paragraph's lines end when the page decides where to split it.
//
// Documents splits a paragraph that crosses a page's foot between two of its
// lines (apps/desktop/renderer/apps/word/pages.js, `splitParagraph`). A line
// of words is held to its whole line box (`lineEndsAt`), with room above and
// below its glyphs, so a part that fits by its glyphs but not by its box is
// not left on a page it then has to leave again. A line of pictures with no
// words is held to its pictures' foot (`fitEndsAt`): the room the page draws
// round an inline picture (6 px above and below, and the font's descent) is
// not Word's, and with it two lines of 450 px pictures would not fit an A4
// page as they do in Word. The numbers are the ones measured in the window
// (window checks `cards` and `sidebyside`).
import test from 'node:test';
import assert from 'node:assert/strict';
import { lineEndsAt, fitEndsAt } from '../apps/desktop/renderer/apps/word/pages.js';

// A4 at 96 dpi with inch margins: 931 px inside, measured from the content's top.
const INSIDE = 931;
const fitsBy = (ends, lim) => ends.findIndex((bottom) => bottom > lim - 0.5);

test('a line of words ends where the next line box starts, and the last where the paragraph does', () => {
  // 15 px words on 22.5 px lines: each line's glyphs stand 2 px into its box.
  const lines = [{ top: 2, bottom: 20 }, { top: 24.5, bottom: 42.5 }, { top: 47, bottom: 65 }];
  assert.deepEqual(lineEndsAt(lines, 0, 67.5), [22.5, 45, 67.5]);
  // Words are never bare: their fit is their line box.
  assert.deepEqual(fitEndsAt(lines, [22.5, 45, 67.5], [false, false, false]), [22.5, 45, 67.5]);
});

test('a part whose glyphs fit but whose line box runs over the foot does not count as fitting', () => {
  const lines = [{ top: 2, bottom: 20 }, { top: 24.5, bottom: 42.5 }];
  const ends = lineEndsAt(lines, 0, 45);
  // The foot 44.6 px down: the second line's glyphs (42.5) fit, its box (45) does not.
  assert.equal(fitsBy(ends, 44.6), 1, 'the second line goes over');
  assert.equal(fitsBy(ends, 46), -1, 'with a pixel to spare it stays');
});

test('a line of pictures with no words fits by its pictures\' foot: two lines of 450 px cards fill an A4 page', () => {
  // Two cards a line, each 450 px with 6 px above and below, the descent under the baseline: 468.5 px a line.
  const lines = [{ top: 6, bottom: 456 }, { top: 474.5, bottom: 924.5 }];
  const ends = lineEndsAt(lines, 0, 937);
  assert.deepEqual(ends, [468.5, 937]);
  assert.equal(fitsBy(ends, INSIDE), 1, 'by line boxes the second line of cards would go to a third page');
  const fits = fitEndsAt(lines, ends, [true, true]);
  assert.deepEqual(fits, [456, 924.5]);
  assert.equal(fitsBy(fits, INSIDE), -1, 'by the pictures\' foot both lines stay');
  // Three lines do not fit: the split falls before the third, as Word's does.
  const three = [...lines, { top: 943, bottom: 1393 }];
  assert.equal(fitsBy(fitEndsAt(three, lineEndsAt(three, 0, 1405.5), [true, true, true]), INSIDE), 2);
});

test('a line of pictures beside words, and the paragraph\'s own border under its pictures, still count', () => {
  const lines = [{ top: 6, bottom: 456 }, { top: 474.5, bottom: 924.5 }];
  // Words on the second line: the whole box.
  assert.deepEqual(fitEndsAt(lines, [468.5, 937], [true, false]), [456, 937]);
  // A border and padding of 8 px under the last line of pictures has to fit too.
  assert.deepEqual(fitEndsAt(lines, [468.5, 945], [true, true], 8), [456, 932.5]);
  // Never past the line box.
  assert.deepEqual(fitEndsAt(lines, [468.5, 930], [true, true], 8), [456, 930]);
});
