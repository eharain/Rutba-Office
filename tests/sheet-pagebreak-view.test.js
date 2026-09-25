/**
 * View → Page Break Preview: the sheet as the printer will cut it, from the
 * same page setup and pagination the print uses — the printed area, each
 * page's box and number, and the breaks between them, a manual one told
 * from one the paper made. The view is `<sheetView view="pageBreakPreview">`.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { writePageSetup, readPageSetup, printSummary } from '@rutba/sheet-view/print';

const grid = (rows, cols) => Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => (r === 0 ? 'H' + c : r * 10 + c)));
const open = (rows = 120, cols = 14, opts = {}) => new SheetView(buildXlsx({ sheets: [{ name: 'Report', rows: grid(rows, cols) }] }), { viewportWidth: 1200, viewportHeight: 800, ...opts });
const sheetXml = (view) => {
  const part = view.workbook.partNameFor(view.activeSheet);
  return view.workbook.snapshotParts([part])[part];
};

test('the view is written to sheetView and read back; Normal takes the attribute away', () => {
  const view = open();
  assert.equal(view.viewMode(), 'normal');
  view.setViewMode('pageBreakPreview');
  assert.match(sheetXml(view), /<sheetView\b[^>]*\bview="pageBreakPreview"/);
  assert.equal(view.render().viewMode, 'pageBreakPreview');
  const again = new SheetView(view.save());
  assert.equal(again.viewMode(), 'pageBreakPreview');
  again.setViewMode('normal');
  assert.doesNotMatch(sheetXml(again), /view="/);
  assert.equal(again.render().pageBreaks, null, 'Normal draws no pages');
  assert.throws(() => again.setViewMode('outline'), /normal, as a page break preview or as its page layout/);
});

test('the pages are the printer\'s own: the same count, numbered down and then across', () => {
  const view = open();
  view.setViewMode('pageBreakPreview');
  const pb = view.pageBreakPreview();
  const summary = printSummary(view, readPageSetup(view));
  assert.equal(pb.count, summary.pages);
  assert.ok(pb.count >= 4, 'a sheet two pages wide and two down: ' + pb.count);
  assert.deepEqual(pb.area, { ...pb.area, top: 0, left: 0, bottom: 119, right: 13 });
  const p1 = pb.pages.find((p) => p.n === 1);
  const p2 = pb.pages.find((p) => p.n === 2);
  assert.equal(p1.x, 0);
  assert.equal(p1.y, 0);
  assert.equal(p2.x, p1.x, 'page 2 is below page 1 — down first');
  assert.equal(p2.y, p1.y + p1.height);
  // Every break is automatic until a person puts one.
  assert.ok(pb.rows.length >= 1 && pb.rows.every((b) => !b.manual));
  assert.ok(pb.cols.length >= 1 && pb.cols.every((b) => !b.manual));
  assert.equal(pb.rows[0].y, view.geo.rowOffset(pb.rows[0].index));
});

test('a break put by hand is marked manual, and moves the pages after it', () => {
  const view = open();
  view.setViewMode('pageBreakPreview');
  const before = view.pageBreakPreview();
  writePageSetup(view, 'Report', { ...readPageSetup(view), rowBreaks: [20, 60] });
  const after = view.pageBreakPreview();
  const manual = after.rows.find((b) => b.index === 20);
  assert.ok(manual && manual.manual === true, JSON.stringify(after.rows));
  assert.ok(after.rows.find((b) => b.index === 60)?.manual);
  assert.ok(after.rows.filter((b) => b.index !== 20 && b.index !== 60).every((b) => !b.manual));
  assert.ok(after.count > before.count, 'breaks before the natural ones make more pages');
  assert.match(sheetXml(view), /<rowBreaks count="2" manualBreakCount="2"><brk id="20" max="16383" man="1"\/><brk id="60" max="16383" man="1"\/><\/rowBreaks>/);
});

test('a print area is what is printed: the area\'s box, and pages only inside it', () => {
  const view = open();
  view.setViewMode('pageBreakPreview');
  writePageSetup(view, 'Report', { ...readPageSetup(view), area: 'B3:E30' });
  const pb = view.pageBreakPreview();
  assert.deepEqual([pb.area.top, pb.area.left, pb.area.bottom, pb.area.right], [2, 1, 29, 4]);
  assert.equal(pb.count, 1);
  assert.equal(pb.area.x, view.geo.colOffset(1));
  assert.equal(pb.pages[0].y, view.geo.rowOffset(2));
});

test('in the frame only the pages near the window are drawn; the plan is kept between frames', () => {
  const view = open(60000, 6);
  view.setViewMode('pageBreakPreview');
  const first = view.render().pageBreaks;
  assert.ok(first.count > 500, 'a long sheet has many pages: ' + first.count);
  assert.ok(first.pages.length < 10, 'only those in view are drawn: ' + first.pages.length);
  const plan = view._breakPlan.laid;
  view.scrollTo(0, 400000);
  const later = view.render().pageBreaks;
  assert.equal(view._breakPlan.laid, plan, 'scrolling does not paginate again');
  assert.ok(later.pages.every((p) => p.y + p.height > 400000 - 400 && p.y < 400000 + 1400));
  // A row made taller is a new plan.
  view.setRowHeight(5, 60);
  view.render();
  assert.notEqual(view._breakPlan.laid, plan);
});
