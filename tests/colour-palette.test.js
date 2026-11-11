// The colour palette's colours (apps/desktop/renderer/apps/colour/grid.js):
// the theme grid's lighter and darker steps by Office's own rule and to
// Office's own numbers, the standard colours, Word's highlight set, the hex
// a person types, the picker's sums, and the recent list.
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  OFFICE_COLOURS, THEME_SLOTS, STANDARD, HIGHLIGHT, hexOf, steps, themeGrid, remember, sameColour,
  hsvOf, hexFromHsv, hslOf, hexFromHsl,
} from '../apps/desktop/renderer/apps/colour/grid.js';

// Office works the steps out in its own whole-number lightness, so a step can
// land one unit away from its hex in a channel; nearer than that is the same
// colour to the eye and to the file.
const near = (actual, expected, what) => {
  assert.equal(actual.length, expected.length, what);
  actual.forEach((a, i) => {
    const off = [0, 2, 4].map((k) => Math.abs(parseInt(a.slice(k, k + 2), 16) - parseInt(expected[i].slice(k, k + 2), 16)));
    assert.ok(Math.max(...off) <= 1, `${what}: step ${i} is ${a}, Office's is ${expected[i]}`);
  });
};

test('a hex is read however it is typed, and nothing else is a colour', () => {
  assert.equal(hexOf('#e03131'), 'E03131');
  assert.equal(hexOf('e03131'), 'E03131');
  assert.equal(hexOf('#abc'), 'AABBCC');
  assert.equal(hexOf(' 0F0 '), '00FF00');
  assert.equal(hexOf('FFE03131'), 'E03131', 'as a sheet reports it, alpha first');
  for (const bad of ['', null, undefined, '#12345', 'red', '#GGGGGG', '1234567', '123456789']) assert.equal(hexOf(bad), null, String(bad));
  assert.ok(sameColour('#ff0000', 'FF0000'));
  assert.ok(!sameColour(null, null), 'no colour is not a colour');
});

test('the theme grid has Office\'s ten columns in Office\'s order', () => {
  assert.deepEqual(THEME_SLOTS, ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']);
  const grid = themeGrid(null);
  assert.equal(grid.length, 10);
  assert.deepEqual(grid.map((c) => c.hex), THEME_SLOTS.map((s) => OFFICE_COLOURS[s]));
  for (const c of grid) assert.equal(c.steps.length, 5, c.slot);
  // A document's own theme, written either way, takes the columns; a slot it lacks is Office's.
  const own = themeGrid({ accent1: '#2b5fd9', dk1: '1C2B39' });
  assert.equal(own.find((c) => c.slot === 'accent1').hex, '2B5FD9');
  assert.equal(own.find((c) => c.slot === 'dk1').hex, '1C2B39');
  assert.equal(own.find((c) => c.slot === 'accent2').hex, 'ED7D31');
});

test('the steps are Office\'s: white darker by 5 to 50, black lighter by 50 to 5', () => {
  assert.deepEqual(steps('FFFFFF').map((s) => [s.darker, s.hex]), [[5, 'F2F2F2'], [15, 'D9D9D9'], [25, 'BFBFBF'], [35, 'A6A6A6'], [50, '808080']]);
  assert.deepEqual(steps('000000').map((s) => [s.lighter, s.hex]), [[50, '808080'], [35, '595959'], [25, '404040'], [15, '262626'], [5, '0D0D0D']]);
});

test('a middling colour goes 80, 60 and 40 lighter then 25 and 50 darker, to Office\'s hex', () => {
  // Blue, Accent 1 in Office's theme, as Word's palette shows it.
  const s = steps('4472C4');
  assert.deepEqual(s.map((x) => x.lighter ?? -x.darker), [80, 60, 40, -25, -50]);
  near(s.map((x) => x.hex), ['DAE3F3', 'B4C7E7', '8FAADC', '2F5597', '203864'], 'Accent 1');
  // Orange, Accent 2.
  near(steps('ED7D31').map((x) => x.hex), ['FBE5D6', 'F8CBAD', 'F4B183', 'C55A11', '843C0C'], 'Accent 2');
});

test('a very dark colour steps lighter, a very light one darker', () => {
  // Blue-grey, Text 2 is middling (lightness 0.35); a near-black navy is not.
  assert.deepEqual(steps('002060').map((x) => x.lighter), [90, 75, 50, 25, 10]);
  // Light grey, Background 2.
  assert.deepEqual(steps('E7E6E6').map((x) => x.darker), [10, 25, 50, 75, 90]);
  near(steps('E7E6E6').map((x) => x.hex), ['D0CECE', 'AEABAB', '767171', '3B3838', '171616'], 'Background 2');
});

test('the standard colours and the highlight colours are Office\'s', () => {
  assert.deepEqual(STANDARD.map(([h]) => h), ['C00000', 'FF0000', 'FFC000', 'FFFF00', '92D050', '00B050', '00B0F0', '0070C0', '002060', '7030A0']);
  assert.equal(HIGHLIGHT.length, 15, 'Word\'s fifteen');
  // Each a w:highlight value Word writes (ST_HighlightColor), once.
  const names = HIGHLIGHT.map(([k]) => k);
  assert.equal(new Set(names).size, 15);
  for (const n of names) assert.match(n, /^(yellow|green|cyan|magenta|blue|red|darkBlue|darkCyan|darkGreen|darkMagenta|darkRed|darkYellow|darkGray|lightGray|black)$/);
});

test('the picker\'s sums go there and back', () => {
  for (const hex of ['FF0000', '00FF00', '0000FF', '4472C4', 'ED7D31', '7F7F7F', '000000', 'FFFFFF', '12ABEF']) {
    assert.equal(hexFromHsv(hsvOf(hex)), hex, `hsv ${hex}`);
    assert.equal(hexFromHsl(hslOf(hex)), hex, `hsl ${hex}`);
  }
  assert.equal(hexFromHsv([0, 1, 1]), 'FF0000');
  assert.equal(hexFromHsv([1 / 3, 1, 1]), '00FF00');
  assert.equal(hexFromHsv([0, 0, 0.5]), '808080');
});

test('the recent colours are newest first, each once, ten at most', () => {
  let list = [];
  for (const h of ['FF0000', '#00ff00', 'ff0000', 'nope']) list = remember(list, h);
  assert.deepEqual(list, ['FF0000', '00FF00']);
  for (let i = 0; i < 15; i++) list = remember(list, i.toString(16).padStart(6, '0'));
  assert.equal(list.length, 10);
  assert.equal(list[0], '00000E');
});
