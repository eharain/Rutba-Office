// Worksheets: Insert → PivotChart and Insert → Slicer.
//
// From a list of sales, Insert → PivotChart opens PivotChart & PivotTable on
// the list round the cursor; OK makes the pivot under the data and a chart
// of it beside it. With the cursor in the pivot, Insert → Slicer offers the
// pivot's fields; two panels are made. Pressing a button filters the pivot
// and its chart, Ctrl+click adds an item, the other panel greys the items
// with no data, Clear Filter shows them all; the panel is dragged by its
// header and resized by a handle. The list made a table gets a slicer too,
// and a press hides the table's other rows. Saved, the file carries the
// slicer parts and the chart's pivot source. Run alone with
// RUTBA_VERIFY_ONLY=slicers.
import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';
import { readSlicers } from '@rutba/ooxml/slicers';
import { readPivotChart } from '@rutba/ooxml/pivot';
import { SheetView } from '@rutba/sheet-view';

const ROWS = [
  ['Region', 'Product', 'Rep', 'Units', 'Revenue'],
  ['North', 'Pens', 'Amira', 120, 360],
  ['South', 'Ink', 'Ben', 80, 640],
  ['East', 'Paper', 'Chen', 300, 450],
  ['North', 'Ink', 'Dana', 45, 360],
  ['West', 'Pens', 'Ben', 60, 180],
  ['South', 'Paper', 'Amira', 210, 315],
  ['East', 'Pens', 'Dana', 90, 270],
  ['West', 'Ink', 'Chen', 30, 240],
  ['North', 'Paper', 'Ben', 150, 225],
];

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, capture, doc, sessionFor
 * @param {{ dir: string }} where the fixtures are written
 */
export async function verifySlicers(h, { dir }) {
  const { open, check, until, wait, errorsIn, capture, doc, sessionFor } = h;
  const file = path.join(dir, 'slicers.xlsx');
  fs.writeFileSync(file, buildXlsx({ sheets: [{
    name: 'Sales', rows: ROWS,
    styles: { 'A1:E1': { bold: true }, 'E2:E10': { numFmt: '#,##0' } },
  }] }));
  try {
    const win = await open('sheets', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = sessionFor('sheet');
    const model = () => doc.model({ id: session.id });
    await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="E10"]'))`), 'the grid', 8000);

    const clickIn = async (title) => {
      const find = `[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || n.textContent || '').trim().startsWith(${JSON.stringify(title)}) && !n.disabled)`;
      await until(() => js(`Boolean(${find})`), `the ${title} button`, 3000).catch(() => {});
      return js(`(() => { const b = ${find}; if (!b) return 'no button ' + ${JSON.stringify(title)}; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked'; })()`);
    };
    const tab = (name) => js(`(() => { [...document.querySelectorAll('.rw-ribbon button, .rw-tabs button, button')].find((b) => b.textContent.trim() === ${JSON.stringify(name)})?.click(); return 1; })()`);
    const setField = (selector, v) => js(`(() => {
      const el = document.querySelector(${JSON.stringify(selector)});
      if (!el) return 'no field ' + ${JSON.stringify(selector)};
      const proto = el.tagName === 'SELECT' ? window.HTMLSelectElement.prototype : window.HTMLInputElement.prototype;
      Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, ${JSON.stringify(v)});
      el.dispatchEvent(new Event(el.tagName === 'SELECT' ? 'change' : 'input', { bubbles: true }));
      return 'set';
    })()`);
    const clickSel = (selector) => js(`(() => { const b = document.querySelector(${JSON.stringify(selector)}); if (!b) return 'no ' + ${JSON.stringify(selector)}; if (b.disabled) return 'disabled'; b.click(); return 'clicked'; })()`);
    const refresh = async () => { win.webContents.invalidate(); await wait(700); };
    const nudge = () => js(`(() => { const g = document.querySelector('.sh-grid'); g.scrollTop += 1; g.scrollTop -= 1; g.dispatchEvent(new Event('scroll')); return 1; })()`);
    const select = async (row, col) => {
      doc.apply({ id: session.id, ops: [{ op: 'select', row, col }] });
      await nudge();
      await wait(300);
    };
    const rect = (selector) => js(`(() => { const n = document.querySelector(${JSON.stringify(selector)}); if (!n) return null; const b = n.getBoundingClientRect(); return { x: b.left, y: b.top, w: b.width, h: b.height }; })()`);
    const drag = async (selector, dx, dy, at = 'centre') => {
      const r = await rect(selector);
      if (!r) return 'no ' + selector;
      const x0 = Math.round(r.x + (at === 'head' ? 30 : r.w / 2));
      const y0 = Math.round(r.y + (at === 'head' ? 12 : r.h / 2));
      win.webContents.sendInputEvent({ type: 'mouseMove', x: x0, y: y0 });
      win.webContents.sendInputEvent({ type: 'mouseDown', x: x0, y: y0, button: 'left', clickCount: 1 });
      for (let i = 1; i <= 6; i++) {
        win.webContents.sendInputEvent({ type: 'mouseMove', x: Math.round(x0 + (dx * i) / 6), y: Math.round(y0 + (dy * i) / 6), button: 'left' });
        await wait(40);
      }
      win.webContents.sendInputEvent({ type: 'mouseUp', x: x0 + dx, y: y0 + dy, button: 'left', clickCount: 1 });
      return 'dragged';
    };
    const pressItem = (slicer, item, ctrl = false) => js(`(() => {
      const b = document.querySelector('.sh-slicer[data-slicer="${slicer}"] .sh-slicer-item[data-item="${item}"]');
      if (!b) return 'no item';
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0, ctrlKey: ${ctrl} }));
      return 'pressed';
    })()`);
    const items = (slicer) => js(`[...document.querySelectorAll('.sh-slicer[data-slicer="${slicer}"] .sh-slicer-item')].map((b) => ({ label: b.dataset.item, on: b.classList.contains('on'), nodata: b.classList.contains('nodata') }))`);
    const cellText = (ref) => js(`document.querySelector('.sh-cell[data-ref="${ref}"]')?.textContent.trim() ?? ''`);

    // ── Insert → PivotChart on data: PivotChart & PivotTable.
    await select(2, 1);
    await tab('Insert');
    const opened = await clickIn('PivotChart');
    const dialog = await until(() => js(`Boolean(document.querySelector('.sh-pivot-source'))`), 'PivotChart & PivotTable', 5000).catch(() => false);
    const prefilled = await js(`document.querySelector('.sh-pivot-source')?.value || ''`);
    const title = await js(`document.querySelector('.rw-dialog-head')?.textContent || ''`);
    await setField('.sh-pivot-rows', 'Region');
    await setField('.sh-pivot-values', 'Revenue');
    await clickSel('.sh-pivot-kind[data-kind="column"]');
    await wait(300);
    await refresh();
    await refresh();
    await capture(win, 'sheets-pivotchart-dialog.png');
    await clickSel('.sh-pivot-ok');
    const made = await until(() => js(`Boolean(document.querySelector('.sh-drawing[data-kind="chart"] svg rect'))`), 'the PivotChart', 8000).catch(() => false);
    const chartTip = await js(`document.querySelector('.sh-drawing[data-kind="chart"]')?.getAttribute('title') || ''`);
    const pivots = model().pivots || [];
    const grand = await until(async () => (await cellText('A17')) === 'Grand Total', 'the pivot', 5000).catch(() => false);
    check('sheets: Insert → PivotChart on data opens PivotChart & PivotTable on the list round the cell; OK makes the pivot under the data and a chart of it beside it',
      opened === 'clicked' && dialog === true && /Sales!A1:E10/.test(prefilled) && /PivotChart & PivotTable/.test(title) && made === true
        && /PivotChart of PivotTable1/.test(chartTip) && pivots.length === 1 && pivots[0].ref === 'A12:B17' && grand === true,
      `${opened}; dialog ${dialog} "${title}", source ${prefilled}; chart ${made} "${chartTip}"; pivots ${JSON.stringify(pivots.map((p) => p.ref))}; grand ${grand}`);
    await refresh();
    await capture(win, 'sheets-pivotchart.png');

    // ── Insert → Slicer with the cursor in the pivot.
    await select(13, 0);
    const slicerOpened = await clickIn('Slicer');
    const asked = await until(() => js(`document.querySelectorAll('.sh-slicers-field').length === 5`), 'Insert Slicers', 5000).catch(() => false);
    await js(`(() => { for (const f of ['Product', 'Rep']) document.querySelector('.sh-slicers-field[data-field="' + f + '"] input').click(); return 1; })()`);
    await wait(200);
    await refresh();
    await capture(win, 'sheets-insert-slicers.png');
    await clickSel('.sh-slicers-ok');
    const panels = await until(() => js(`document.querySelectorAll('.sh-slicer').length === 2 && document.querySelectorAll('.sh-slicer[data-slicer="Product"] .sh-slicer-item.on').length === 3`), 'the two slicers', 8000).catch(() => false);
    const product = await items('Product');
    const head = await js(`document.querySelector('.sh-slicer[data-slicer="Product"] .sh-slicer-caption')?.textContent || ''`);
    const noOverlap = await js(`(() => { const c = document.querySelector('.sh-drawing[data-kind="chart"]').getBoundingClientRect(); const s = document.querySelector('.sh-slicer[data-slicer="Product"]').getBoundingClientRect(); return s.left >= c.right - 1 || s.top >= c.bottom - 1; })()`);
    check('sheets: Insert → Slicer in a pivot offers its fields; each ticked field gets a panel with a button per item, all selected, clear of the chart',
      slicerOpened === 'clicked' && asked === true && panels === true && head === 'Product'
        && JSON.stringify(product.map((i) => i.label)) === JSON.stringify(['Ink', 'Paper', 'Pens']) && noOverlap === true,
      `${slicerOpened}; dialog ${asked}; panels ${panels}; ${JSON.stringify(product)}; caption "${head}"; clear of the chart ${noOverlap}`);

    // A press shows that item alone: the pivot and its chart follow.
    const before = await cellText('B17');
    await pressItem('Product', 'Pens');
    // South sells no pens: three regions left, the total a row higher.
    const filtered = await until(async () => (await cellText('B16')) === '810' && (await cellText('A16')) === 'Grand Total', 'the pivot filtered', 8000).catch(() => false);
    const afterItems = await items('Product');
    const reps = await items('Rep');
    const clearOn = await js(`!document.querySelector('.sh-slicer[data-slicer="Product"] .sh-slicer-clear').disabled`);
    const chartSays = await js(`document.querySelector('.sh-drawing[data-kind="chart"] svg')?.textContent || ''`);
    check('sheets: a slicer button filters the pivot to that item — totals and chart follow — and the other panel fades the items with no data',
      filtered === true && before === '3040' && JSON.stringify(afterItems.filter((i) => i.on).map((i) => i.label)) === '["Pens"]' && clearOn
        && reps.filter((i) => i.nodata).map((i) => i.label).join(',') === 'Chen' && /360/.test(chartSays),
      `B17 ${before} → B16 ${await cellText('B16')}; ${JSON.stringify(afterItems)}; reps ${JSON.stringify(reps)}; clear ${clearOn}`);
    await refresh();
    await capture(win, 'sheets-slicers-pivot.png');

    await pressItem('Product', 'Ink', true);
    const two = await until(async () => JSON.stringify((await items('Product')).filter((i) => i.on).map((i) => i.label)) === '["Ink","Pens"]', 'Ctrl+click adds Ink', 6000).catch(() => false);
    await clickSel('.sh-slicer[data-slicer="Product"] .sh-slicer-clear');
    const cleared = await until(async () => (await cellText('B17')) === '3040' && (await items('Product')).every((i) => i.on), 'Clear Filter', 6000).catch(() => false);
    check('sheets: Ctrl+click adds an item to the slicer\'s choice; Clear Filter shows every item and the whole pivot again', two === true && cleared === true, `two ${two}; cleared ${cleared}`);

    // Dragged by its header, resized by a handle.
    const d0 = (model().drawings || []).find((d) => d.kind === 'slicer' && d.slicer?.name === 'Product');
    await drag('.sh-slicer[data-slicer="Product"] .sh-slicer-head', 60, 40, 'head');
    const moved = await until(() => { const d = (model().drawings || []).find((x) => x.kind === 'slicer' && x.slicer?.name === 'Product'); return d && Math.abs(d.x - d0.x - 60) <= 2 && Math.abs(d.y - d0.y - 40) <= 2; }, 'the slicer moved', 6000).catch(() => false);
    const handles = await until(() => js(`document.querySelectorAll('.sh-obj-handle').length === 8`), 'the handles', 3000).catch(() => false);
    await drag('.sh-obj-handle[data-handle="se"]', 40, 30);
    const resized = await until(() => { const d = (model().drawings || []).find((x) => x.kind === 'slicer' && x.slicer?.name === 'Product'); return d && Math.abs(d.width - d0.width - 40) <= 2 && Math.abs(d.height - d0.height - 30) <= 2; }, 'the slicer resized', 6000).catch(() => false);
    check('sheets: a slicer is dragged by its header and resized by a corner handle, its anchor following the cells', moved === true && handles === true && resized === true, `moved ${moved}; handles ${handles}; resized ${resized}`);
    await refresh();
    await capture(win, 'sheets-slicer-picked.png');

    // ── A table's slicer: the list made a table, a Region panel, North pressed.
    await select(3, 0);
    await clickIn('Table');
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim().startsWith('Light')))`), 'the table menu', 4000).catch(() => {});
    await js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim().startsWith('Light'))?.click(); return 1; })()`);
    await until(() => (model().tables || []).length === 1, 'the table', 6000).catch(() => {});
    await select(3, 0);
    await clickIn('Slicer');
    await until(() => js(`Boolean(document.querySelector('.sh-slicers-field[data-field="Region"]'))`), 'Insert Slicers for the table', 5000).catch(() => {});
    await js(`(() => { document.querySelector('.sh-slicers-field[data-field="Region"] input').click(); return 1; })()`);
    await wait(150);
    await clickSel('.sh-slicers-ok');
    await until(() => js(`Boolean(document.querySelector('.sh-slicer[data-slicer="Region"]'))`), 'the table slicer', 6000).catch(() => {});
    await pressItem('Region', 'North');
    const hid = await until(() => { const shown = (model().rows || []).map((r) => r.index).filter((r) => r >= 1 && r <= 9); return JSON.stringify(shown) === '[1,4,9]'; }, 'the table filtered', 8000).catch(() => false);
    check('sheets: a table\'s slicer filters the table — a press on North hides every other region\'s rows',
      hid === true, `rows ${JSON.stringify((model().rows || []).map((r) => r.index).filter((r) => r <= 10))}`);
    await refresh();
    await capture(win, 'sheets-slicers-table.png');

    // Saved: the parts Excel reads.
    doc.save({ id: session.id });
    const saved = new SheetView(fs.readFileSync(file));
    const read = readSlicers(saved.workbook);
    const chartPart = saved.pkg.partNames().find((p) => /^xl\/charts\/chart\d+\.xml$/.test(p));
    const info = chartPart ? readPivotChart(saved.pkg.text(chartPart)) : null;
    check('sheets: saved, the workbook carries three slicers with their caches (two on the pivot, one on the table) and the chart\'s pivot source',
      read.length === 3 && read.filter((s) => s.cache?.kind === 'pivot').length === 2 && read.some((s) => s.cache?.kind === 'table')
        && info?.name === 'PivotTable1' && /^\[slicers\.xlsx\]Sales!PivotTable1$/.test(info?.source || ''),
      `${JSON.stringify(read.map((s) => [s.name, s.cache?.kind]))}; chart ${JSON.stringify(info)}`);
    // Put the fixture's own state back for a second run in the same folder.
    await pressItem('Region', 'North');
    await js(`(() => { document.querySelector('.sh-slicer[data-slicer="Region"] .sh-slicer-clear')?.click(); return 1; })()`);

    const complaints = await errorsIn(win);
    check('sheets: the PivotChart and slicer checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('sheets: the PivotChart and slicer checks ran', false, err.message);
  }
}
