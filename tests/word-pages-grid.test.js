// Pages side by side: the arithmetic.
//
// Documents lays its pages across the window when there is room, as Word
// does at about half size and below, and in one row for View, Page
// Movement, Side to Side. The pages are only drawn moved (see
// apps/desktop/renderer/apps/word/pages.js, `gridOf`); how many go across,
// where each one is drawn and which page a point falls on is plain
// arithmetic, held here. The window check `pages` proves the drawing.
import test from 'node:test';
import assert from 'node:assert/strict';
import { gridOf, pagesAcross, pageAtPoint, gapAt, geometryOf, pageTopOf, PAGE_GAP, GAP_SCREEN, DESK_SIDE } from '../apps/desktop/renderer/apps/word/pages.js';

// A4 at 96 dpi: 794 by 1123, inch margins.
const section = { widthPx: 794, heightPx: 1123, margins: { top: 96, bottom: 96, left: 96, right: 96 } };
const geo = geometryOf(section);
const W = 794;

test('the gap between pages beside each other is never less than the one below a page, nor less than a fixed room on screen', () => {
  assert.equal(gapAt(1), PAGE_GAP);
  assert.equal(gapAt(0.5), Math.ceil(GAP_SCREEN / 0.5));
  assert.ok(gapAt(0.3) * 0.3 >= GAP_SCREEN);
  assert.ok(gapAt(2) >= PAGE_GAP);
});

test('as many pages go across as fit the desk at the zoom', () => {
  // A normal window at 100%: one.
  assert.equal(pagesAcross({ desk: 1280, zoom: 1, W }), 1);
  // Half size in the same window: two (each 397 px, with the gap and the desk's room).
  assert.equal(pagesAcross({ desk: 1280, zoom: 0.5, W }), 3);
  assert.equal(pagesAcross({ desk: 1000, zoom: 0.5, W }), 2);
  // A third of the size: four in a 1280 window.
  assert.ok(pagesAcross({ desk: 1280, zoom: 0.3, W }) >= 3);
  // Exactly enough room for two: two, and a pixel less: one.
  const z = 0.5;
  const g = gapAt(z);
  const exact = (2 * W + g) * z + 2 * DESK_SIDE;
  assert.equal(pagesAcross({ desk: exact, zoom: z, W }), 2);
  assert.equal(pagesAcross({ desk: exact - 1, zoom: z, W }), 1);
  // No room at all: still one.
  assert.equal(pagesAcross({ desk: 0, zoom: 1, W }), 1);
});

test('one page across is no grid at all: the pages stand one under another as before', () => {
  assert.equal(gridOf({ geo, count: 8, W, zoom: 1, desk: 1280 }), null);
  assert.equal(gridOf({ geo, count: 8, W, zoom: 0.5, desk: 1280, most: 1 }), null, 'One Page asks for one across');
  assert.equal(gridOf({ geo, count: 1, W, zoom: 0.3, desk: 1600 }), null, 'a one-page document has nothing to lay beside it');
});

test('pages in rows that wrap: reading order left to right, each row sharing its top edge', () => {
  const grid = gridOf({ geo, count: 7, W, zoom: 0.5, desk: 1280 });
  assert.equal(grid.cols, 3);
  assert.equal(grid.rows, 3);
  const g = grid.gap;
  assert.deepEqual(grid.places.map((p) => [p.row, p.col]), [[0, 0], [0, 1], [0, 2], [1, 0], [1, 1], [1, 2], [2, 0]]);
  assert.deepEqual(grid.places.slice(0, 3).map((p) => p.x), [0, W + g, 2 * (W + g)]);
  assert.ok(grid.places.slice(0, 3).every((p) => p.y === 0), 'the first row shares a top edge');
  assert.equal(grid.places[3].y, geo.H + g, 'the second row starts a page and a gap down');
  assert.equal(grid.width, 3 * W + 2 * g);
  assert.equal(grid.height, 3 * geo.H + 2 * g);
  // Each page's move is from where the layout puts it: under the one before.
  for (let k = 0; k < 7; k++) {
    assert.equal(grid.places[k].dx, grid.places[k].x);
    assert.equal(grid.places[k].dy, grid.places[k].y - pageTopOf(geo, k));
  }
  assert.equal(grid.places[0].dx, 0);
  assert.equal(grid.places[0].dy, 0, 'the first page stays where it is');
});

test('never more across than there are pages, nor than One Page or Multiple Pages ask', () => {
  assert.equal(gridOf({ geo, count: 2, W, zoom: 0.3, desk: 1800 }).cols, 2);
  assert.equal(gridOf({ geo, count: 9, W, zoom: 0.3, desk: 1800, most: 2 }).cols, 2);
});

test('a right-to-left document lays its pages from the right', () => {
  const grid = gridOf({ geo, count: 5, W, zoom: 0.5, desk: 1280, rtl: true });
  const g = grid.gap;
  assert.deepEqual(grid.places.slice(0, 3).map((p) => p.x), [2 * (W + g), W + g, 0]);
  assert.equal(grid.places[3].x, 2 * (W + g), 'the next row starts at the right again');
});

test('Side to Side: every page in one row, whatever the desk', () => {
  const grid = gridOf({ geo, count: 6, W, zoom: 1, desk: 900, side: true });
  assert.equal(grid.cols, 6);
  assert.equal(grid.rows, 1);
  assert.ok(grid.places.every((p) => p.y === 0));
  assert.equal(grid.width, 6 * W + 5 * grid.gap);
  assert.equal(grid.height, geo.H);
  // Even one page is a row of its own there.
  assert.equal(gridOf({ geo, count: 1, W, zoom: 1, desk: 900, side: true }).cols, 1);
});

test('an envelope in front: its row is as tall as the tallest page in it, and it keeps its own place in the layout', () => {
  const envGeo = geometryOf(section, { widthPx: 912, heightPx: 396, margins: { top: 24, bottom: 48, left: 38 }, endsAt: 1 });
  const grid = gridOf({ geo: envGeo, count: 4, W, zoom: 0.5, desk: 1000 });
  assert.equal(grid.cols, 2);
  assert.equal(grid.rowH[0], envGeo.H, 'the letter beside the envelope sets the row');
  assert.equal(grid.places[1].dy, 0 - pageTopOf(envGeo, 1));
  assert.equal(grid.places[2].y, envGeo.H + grid.gap);
});

test('the page at a point: the nearest one, gaps split down the middle, past the last page the last', () => {
  const grid = gridOf({ geo, count: 5, W, zoom: 0.5, desk: 1280 });
  const g = grid.gap;
  assert.equal(pageAtPoint(grid, 10, 10), 0);
  assert.equal(pageAtPoint(grid, W + 1, 10), 0, 'just into the gap: the page before it');
  assert.equal(pageAtPoint(grid, W + g - 1, 10), 1, 'just short of the next page: that page');
  assert.equal(pageAtPoint(grid, 2 * (W + g) + 5, geo.H - 5), 2);
  assert.equal(pageAtPoint(grid, 5, geo.H + g + 5), 3);
  assert.equal(pageAtPoint(grid, 2 * (W + g) + 5, geo.H + g + 5), 4, 'an empty place in the last row: the last page');
  assert.equal(pageAtPoint(grid, -50, -50), 0);
  const rtl = gridOf({ geo, count: 5, W, zoom: 0.5, desk: 1280, rtl: true });
  assert.equal(pageAtPoint(rtl, 5, 5), 2, 'right to left: the left-hand page of the first row is the third');
});
