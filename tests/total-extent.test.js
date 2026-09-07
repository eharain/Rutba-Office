/**
 * The canvas total is a floor, never a spring.
 *
 * The defect this pins: `total` followed the viewport, so a frame computed
 * behind a scrolled-ahead client came back SMALLER than the canvas the
 * client was already standing on. The browser clamped the scroll position to
 * fit, the clamp fired a scroll event, the event became an action, and that
 * action produced the next too-small frame — a feedback loop with no user in
 * it. In a foreground tab the sink's coalescing hid it at ~8 requests a
 * second; in a background tab, where the browser clamps timers to one tick a
 * second, it showed in the gateway log as a metronome: one POST per second,
 * forever, from an idle tab.
 *
 * A total that never shrinks while the session lives cannot clamp anybody.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { SheetView } from '@rutba/sheet-view';
import { buildComplexWorkbook } from './fixtures/complex-workbook.js';

const WORKBOOK = buildComplexWorkbook({ protect: false });

test('scrolling back does not shrink the canvas', () => {
  const view = SheetView.open(WORKBOOK, { viewportWidth: 900, viewportHeight: 500 });

  const atTop = view.render().total;

  view.scrollTo(0, 20_000);
  const farDown = view.render().total;
  assert.ok(farDown.height > atTop.height,
    'scrolling toward the edge extends the sheet - the spreadsheet feel');

  view.scrollTo(0, 0);
  const backAtTop = view.render().total;
  assert.equal(backAtTop.height, farDown.height,
    'and coming back leaves the explored grid in place - a shrinking total is '
    + 'the one thing that can move a client\'s scroll position by itself');
  assert.equal(backAtTop.width, farDown.width);
});

test('the total always clears the viewport bottom - a frame can never clamp the position it was asked from', () => {
  const view = SheetView.open(WORKBOOK, { viewportWidth: 900, viewportHeight: 1000 });

  // Walk down and back in uneven strides, the way a wheel and a scrollbar
  // actually move, asking at each stop whether the frame could clamp it.
  for (const y of [0, 3_000, 12_000, 7_000, 25_000, 400, 25_000]) {
    view.scrollTo(0, y);
    const { total } = view.render();
    assert.ok(total.height >= y + 1000,
      `at y=${y} the total (${total.height}) must cover the viewport bottom, `
      + 'or the browser clamps scrollTop and the clamp echoes as a new action');
  }
});

test('each sheet keeps its own extent', () => {
  const view = SheetView.open(WORKBOOK, { viewportWidth: 900, viewportHeight: 500 });
  const [first, second] = view.sheetNames();

  view.scrollTo(0, 30_000);
  const explored = view.render().total;

  view.selectSheet(second);
  const other = view.render().total;
  assert.ok(other.height < explored.height,
    'a sheet nobody scrolled must not inherit another sheet\'s explored canvas');

  view.selectSheet(first);
  assert.equal(view.render().total.height, explored.height,
    'and coming back finds the first sheet\'s floor where it was left');
});
