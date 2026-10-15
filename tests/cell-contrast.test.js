// A document's own colours in the dark theme: a cell's fill stays the file's,
// as Excel keeps it, and the words on it stay readable. Automatic text takes
// whichever of the suite's inks reads on the fill; a font colour the file set
// on a fill is the author's and is kept; with no fill a colour too dark for
// the night surface is lifted, keeping its hue. The window checks
// (RUTBA_VERIFY_ONLY=sheet-dark) measure it on screen.
import test from 'node:test';
import assert from 'node:assert/strict';
import { parseColour, contrast, inkOn, readableOn, cellInk, edgeInk, INK, SURFACE, AA, AA_LARGE } from '@rutba/office-ui/contrast';

test('colours are read in the forms files and the window give them', () => {
  assert.deepEqual(parseColour('#EEF2F7'), [0xee, 0xf2, 0xf7]);
  assert.deepEqual(parseColour('FFEEF2F7'), [0xee, 0xf2, 0xf7]);
  assert.deepEqual(parseColour('#fff'), [255, 255, 255]);
  assert.deepEqual(parseColour('rgb(30, 33, 39)'), [30, 33, 39]);
  assert.equal(parseColour('rgba(0, 0, 0, 0.1)'), null);
  assert.equal(parseColour('var(--ink)'), null);
  assert.equal(parseColour(null), null);
});

test('contrast is WCAG\'s ratio', () => {
  assert.equal(Math.round(contrast('#000000', '#ffffff')), 21);
  assert.equal(contrast('#777777', '#777777'), 1);
});

test('the sample\'s total row: a light fill keeps its fill and its automatic words go dark in either theme', () => {
  for (const dark of [false, true]) {
    const { ink, background } = cellInk({ fill: '#EEF2F7', font: null, dark });
    assert.equal(background, '#EEF2F7');
    assert.equal(ink, INK.light);
    assert.ok(contrast(ink, background) >= AA, `${dark ? 'dark' : 'light'}: ${contrast(ink, background)}`);
  }
});

test('automatic words on a dark fill go light', () => {
  const { ink } = cellInk({ fill: '#1F3A5F', dark: false });
  assert.equal(ink, INK.dark);
  assert.equal(inkOn('#0f9d58') === INK.light || inkOn('#0f9d58') === INK.dark, true);
});

test('a font colour the file set on a fill is the author\'s and is kept', () => {
  assert.equal(cellInk({ fill: '#1F3A5F', font: '#FFFFFF', dark: true }).ink, '#FFFFFF');
  assert.equal(cellInk({ fill: '#FFEB9C', font: '#9C5700', dark: true }).ink, '#9C5700');
});

test('with no fill, a dark file colour is lifted on the night surface and left alone by day', () => {
  const teal = '#0E7C66';
  assert.equal(cellInk({ font: teal, dark: false }).ink, teal);
  const night = cellInk({ font: teal, dark: true }).ink;
  assert.notEqual(night, teal);
  assert.ok(contrast(night, SURFACE.dark) >= AA, String(contrast(night, SURFACE.dark)));
  // Still teal: green the strongest, blue next, red least.
  const [r, g, b] = parseColour(night);
  assert.ok(g > b && b > r, night);
  // Automatic words with no fill are the suite's own ink: nothing to say.
  assert.equal(cellInk({ dark: true }).ink, null);
  // A colour that already reads at night is not touched.
  assert.equal(cellInk({ font: '#FFC000', dark: true }).ink, '#FFC000');
});

test('every pair a file can bring reads at night', () => {
  const fills = [null, '#EEF2F7', '#FFFFFF', '#C6EFCE', '#FFC7CE', '#FFEB9C', '#DDEBF7', '#1F3A5F', '#000000', '#0F9D58', '#4472C4', '#7F7F7F'];
  const fonts = [null, '#000000', '#1F3864', '#0E7C66', '#C00000', '#0000FF', '#7030A0', '#808080'];
  for (const fill of fills) {
    for (const font of fonts) {
      const { ink, background } = cellInk({ fill, font, dark: true });
      const shown = ink || INK.dark;
      // The author's own pair on a fill is theirs; every other pair reads.
      if (fill && font) continue;
      assert.ok(contrast(shown, background) >= AA - 0.01, `${fill} / ${font}: ${shown} at ${contrast(shown, background).toFixed(2)}`);
    }
  }
});

test('readableOn mirrors a dark colour to its light twin, keeping the hue', () => {
  assert.equal(readableOn('#ffffff', SURFACE.dark), '#ffffff');
  const navy = readableOn('#1F3864', SURFACE.dark);
  assert.ok(contrast(navy, SURFACE.dark) >= AA, navy);
  const [r, g, b] = parseColour(navy);
  assert.ok(b > g && g > r, navy);
  // Black the file set by hand reads as the night's light ink would.
  const black = readableOn('#000000', SURFACE.dark);
  assert.ok(contrast(black, SURFACE.dark) > 12, black);
  // And by day, on white, a pale colour goes the other way.
  const pale = readableOn('#FFEB9C', SURFACE.light);
  assert.ok(contrast(pale, SURFACE.light) >= AA, pale);
});

test('edges: automatic is the file\'s black by day and visible at night; a set colour on the night surface is lifted', () => {
  assert.equal(edgeInk({ dark: false }), '#000000');
  assert.ok(contrast(edgeInk({ dark: true }), SURFACE.dark) >= AA_LARGE);
  assert.equal(edgeInk({ dark: true, fill: '#EEF2F7' }), '#000000');
  assert.equal(edgeInk({ colour: '#1F3A5F', dark: true, fill: '#EEF2F7' }), '#1F3A5F');
  assert.ok(contrast(edgeInk({ colour: '#1F3A5F', dark: true }), SURFACE.dark) >= AA_LARGE);
});
