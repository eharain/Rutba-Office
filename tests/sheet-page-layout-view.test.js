/**
 * View → Page Layout: the sheet drawn as the pages it prints on — the same
 * page setup and the same cuts the printer makes (`planBands`, which
 * `paginate` crosses into pages), each band of rows and columns a sheet of
 * paper with its margins, the header and footer lines the print gives them,
 * a gap between pages, and blank pages past the printed range. The view is
 * `<sheetView view="pageLayout">`; the ruler is `showRuler`.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { writePageSetup, readPageSetup, paginate, parseArea, PX_PER_MM } from '@rutba/sheet-view/print';

const grid = (rows, cols) => Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => (r === 0 ? 'H' + c : r * 10 + c)));
const open = (rows = 150, cols = 14) => new SheetView(buildXlsx({ sheets: [{ name: 'Report', rows: grid(rows, cols) }] }), { viewportWidth: 1400, viewportHeight: 900 });
const sheetXml = (view) => {
  const part = view.workbook.partNameFor(view.activeSheet);
  return view.workbook.snapshotParts([part])[part];
};
/** Every page of the layout, however far the window would have to scroll. */
const allPages = (view) => {
  const L = view._pageLayout();
  view.viewportWidth = L.total.width;
  view.viewportHeight = L.total.height;
  view.scrollTo(0, 0);
  return view.render().pageLayout.pages;
};

test('Page Layout is written to the sheet view and read back; its frame is paper, not a grid', () => {
  const view = open();
  view.setViewMode('pageLayout');
  assert.match(sheetXml(view), /<sheetView\b[^>]*\bview="pageLayout"/);
  assert.equal(new SheetView(view.save()).viewMode(), 'pageLayout');
  const frame = view.render();
  assert.equal(frame.viewMode, 'pageLayout');
  assert.ok(frame.pageLayout.pages.length >= 2, 'pages in sight');
  const a4 = 210 * PX_PER_MM;
  assert.ok(Math.abs(frame.pageLayout.pageWidth - a4) < 1, 'an A4 page, at 100%');
  assert.equal(frame.frozen.rows, 0, 'nothing is frozen on paper');
  assert.equal(frame.split, null);
  assert.ok(frame.total.width >= frame.pageLayout.pageWidth * 2, 'two pages across for fourteen columns');
});

test('the pages are cut where the printer cuts them — the same count, the same rows and columns on each, numbered down then across', () => {
  const view = open();
  view.setViewMode('pageLayout');
  const setup = readPageSetup(view);
  const b = view.calc.usedBounds('Report');
  const plan = paginate({ geo: view.geo, range: parseArea(setup.area) || { top: 0, left: 0, bottom: b.maxRow, right: b.maxCol }, setup });
  const pages = allPages(view).filter((p) => p.n);
  assert.equal(pages.length, plan.pages.length, 'every printed page, and no more, is numbered');
  assert.equal(view.render().pageLayout.count, plan.pages.length);
  const L = view._pageLayout();
  for (const printed of plan.pages) {
    const n = plan.pages.indexOf(printed) + 1;
    const page = pages.find((p) => p.n === n);
    const firstRow = printed.rows[0].index;
    const firstCol = printed.cols[0].index;
    // Each printed page's first column and row sit at the top left of the laid-out page with its number.
    assert.equal(Math.round(L.mapX(firstCol)), Math.round(page.box.x), `page ${n}: column ${firstCol} opens it`);
    if (printed.rows[0].index !== 1 || n > 1) assert.ok(L.mapY(firstRow) >= page.box.y && L.mapY(firstRow) < page.box.y + page.box.height, `page ${n}: row ${firstRow} is on it`);
    const lastRow = printed.rows[printed.rows.length - 1].index;
    assert.ok(L.mapY(lastRow) + view.geo.rowHeight(lastRow) <= page.box.y + page.box.height + 0.5, `page ${n}: its last row fits on it`);
  }
  assert.ok(allPages(view).some((p) => p.blank), 'past the printed range, blank pages to type onto');
});

test('a manual break and Fit to width move the cuts the same way on screen as on paper', () => {
  const view = open();
  const setup = readPageSetup(view);
  writePageSetup(view, 'Report', { ...setup, rowBreaks: [20], fit: 'width' });
  view.setViewMode('pageLayout');
  const L = view._pageLayout();
  assert.equal(L.cols.length >= 1 && L.printedCols, 1, 'fitted: all the columns on one page across');
  assert.ok(L.scale < 1, 'at the print\'s own scale');
  assert.ok(L.pageW > 210 * PX_PER_MM, 'so a page holds more of the sheet at 100%');
  assert.equal(L.rows[1].start, 20, 'the manual break starts the second page');
  const page2 = allPages(view).find((p) => p.n === 2);
  assert.equal(page2.first.row, 20);
});

test('the screen and paper agree: a row maps onto its page and back to the same place in the sheet', () => {
  const view = open(400, 20);
  view.setViewMode('pageLayout');
  const L = view._pageLayout();
  let last = -1;
  for (const row of [0, 1, 40, 41, 120, 399]) {
    const y = L.mapY(row);
    assert.ok(y > last, 'rows go down the pages in order');
    last = y;
    assert.ok(Math.abs(L.toSheetY(y) - view.geo.rowOffset(row)) < 1e-6, `row ${row} round trip`);
  }
  // A scroll into the gap between two pages reads as the foot of the first.
  const gapY = L.pageY(1) - 10;
  assert.equal(L.toSheetY(gapY), view.geo.rowOffset(L.rows[0].end) + view.geo.rowHeight(L.rows[0].end));
  // Cells in the frame are drawn where they lie on paper.
  view.scrollTo(0, L.pageY(2));
  const f = view.render();
  const cell = f.cells.find((c) => c.row === L.rows[2].start && c.col === 0);
  assert.equal(Math.round(cell.y), Math.round(L.contentY(2)));
  assert.equal(f.rows.find((r) => r.index === L.rows[2].start).y, cell.y, 'the row heading beside it');
});

test('the header and footer take their lines on the page as they do in print; the ruler is kept in the sheet view', () => {
  const view = open();
  const setup = readPageSetup(view);
  writePageSetup(view, 'Report', { ...setup, header: '&LQuarterly&RPage &P', footer: '&CConfidential' });
  view.setViewMode('pageLayout');
  const pl = view.render().pageLayout;
  assert.equal(pl.header, '&LQuarterly&RPage &P');
  assert.equal(pl.footer, '&CConfidential');
  assert.equal(pl.head, 22);
  assert.equal(pl.foot, 22);
  const L = view._pageLayout();
  assert.equal(Math.round(L.contentY(0) - L.pageY(0)), Math.round(pl.margins.top + 22), 'the cells start under the header line');
  assert.equal(pl.ruler, true);
  view.setShowRuler(false);
  assert.match(sheetXml(view), /showRuler="0"/);
  assert.equal(view.render().pageLayout.ruler, false);
  view.setShowRuler(true);
  assert.doesNotMatch(sheetXml(view), /showRuler/);
});

test('a frame of a 60,000-row sheet in Page Layout stays quick, at the top and deep down', () => {
  const rows = Array.from({ length: 60000 }, (_, r) => [r === 0 ? 'Item' : 'Row ' + r, r, r * 2, r % 7]);
  const view = new SheetView(buildXlsx({ sheets: [{ name: 'Big', rows }] }), { viewportWidth: 1400, viewportHeight: 800 });
  view.setViewMode('pageLayout');
  view.render();
  const L = view._pageLayout();
  const times = [];
  for (const row of [0, 20000, 45000, 59990, 300]) {
    view.scrollTo(0, L.mapY(row) - 100);
    const t0 = process.hrtime.bigint();
    const f = view.render();
    times.push(Number(process.hrtime.bigint() - t0) / 1e6);
    assert.ok(f.cells.some((c) => c.row === row), `row ${row} is in the frame`);
  }
  // The median, not the slowest: a frame that walked the whole sheet would make
  // every one slow, while a single pause for garbage collection on a busy
  // machine (430 ms once, beside frames of 100) is not the view's doing.
  const later = times.slice(1).sort((x, y) => x - y);
  const median = (later[(later.length - 1) >> 1] + later[later.length >> 1]) / 2;
  assert.ok(median < 150, `frames ${times.map((t) => t.toFixed(1)).join(', ')} ms`);
  assert.ok(view.render().pageLayout.count > 1000);
});
