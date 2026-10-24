// Worksheets: PivotTable Fields — a click in a pivot shows the pane beside
// the sheet, its fields ticked as the pivot uses them; Product moved to
// Columns lays the pivot out again at once, a ticked Price goes in Values,
// a value summarised by Max says so, and the saved file keeps the new
// layout for Excel to refresh. Then the Filters area shows (All) and one
// item, the pane's own Refresh recounts an edited figure, a new pivot under
// a long list is scrolled to with its pane open, and its dates go by month.

import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';
import { Workbook } from '@rutba/ooxml';
import { SheetView } from '@rutba/sheet-view';

/**
 * @param {object} h the harness: open, check, until, wait
 */
export async function verifySheetPivotFields(h, { dir }) {
  const { open, check, until, wait } = h;
  const file = path.join(dir, 'pivot-fields.xlsx');
  try {
    // A list and a pivot of it — Sum of Qty by Region — as a file.
    const rows = [['Region', 'Product', 'Qty', 'Price'], ['North', 'Ink', 4, 2], ['South', 'Ink', 6, 2], ['North', 'Paper', 3, 5], ['South', 'Paper', 1, 5]];
    const made = SheetView.open(buildXlsx({ sheets: [{ name: 'Data', rows }] }));
    made.createPivot({ name: 'Sales', source: 'Data!A1:D5', target: { sheet: 'Data', row: 0, col: 6 }, rowFields: ['Region'], dataFields: [{ field: 'Qty', subtotal: 'sum' }] });
    fs.writeFileSync(file, made.save());

    const win = await open('sheets', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const capture = async (name) => { if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG()); };
    const cell = (ref) => js(`(document.querySelector('.sh-cell[data-ref="${ref}"]')?.textContent || '').trim()`);
    const pane = () => js(`(() => {
      const p = document.querySelector('.sh-pf');
      if (!p) return null;
      const area = (a) => [...p.querySelectorAll('.sh-pf-area[data-area="' + a + '"] .sh-pf-item')].map((i) => i.dataset.field);
      return { pivot: p.dataset.pivot, ticked: [...p.querySelectorAll('.sh-pf-field input')].filter((i) => i.checked).map((i) => i.dataset.field), rows: area('rows'), cols: area('cols'), values: [...p.querySelectorAll('.sh-pf-area[data-area="values"] .sh-pf-name')].map((n) => n.textContent) };
    })()`);

    await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="G2"]'))`), 'the grid', 8000);
    await js(`(() => { document.querySelector('.sh-cell[data-ref="G2"]').dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 })); return 1; })()`);
    const shown = await until(async () => (await pane())?.pivot === 'Sales', 'the PivotTable Fields pane', 5000).then(() => pane()).catch(() => pane());
    check('worksheets: a click in a pivot shows PivotTable Fields beside the sheet, its fields ticked as the pivot uses them',
      shown?.pivot === 'Sales' && JSON.stringify(shown.ticked) === '["Region","Qty"]' && JSON.stringify(shown.rows) === '["Region"]' && shown.values[0] === 'Sum of Qty',
      JSON.stringify(shown));

    // Product ticked goes in Rows; then moved to Columns.
    await js(`(() => { document.querySelector('.sh-pf-field input[data-field="Product"]').click(); return 1; })()`);
    await until(async () => (await pane())?.rows?.includes('Product'), 'Product in Rows', 5000).catch(() => {});
    await js(`(() => {
      const item = document.querySelector('.sh-pf-area[data-area="rows"] .sh-pf-item[data-field="Product"] .sh-pf-to');
      const set = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set;
      set.call(item, 'cols');
      item.dispatchEvent(new Event('change', { bubbles: true }));
      return 1;
    })()`);
    const moved = await until(async () => JSON.stringify((await pane())?.cols) === '["Product"]' && (await cell('H1')) === 'Ink', 'the pivot laid out again', 6000).then(() => true).catch(() => false);
    const grid = { G1: await cell('G1'), H1: await cell('H1'), I1: await cell('I1'), G2: await cell('G2'), H2: await cell('H2'), J4: await cell('J4') };
    check('worksheets: a field moved to Columns lays the pivot out again at once — products across, regions down, totals at the end',
      moved && grid.G2 === 'North' && grid.H2 === '4' && grid.J4 === '14', JSON.stringify({ moved, grid, pane: await pane() }));

    // Price ticked goes in Values (it holds numbers), summarised by Max.
    await js(`(() => { document.querySelector('.sh-pf-field input[data-field="Price"]').click(); return 1; })()`);
    await until(async () => (await pane())?.values?.length === 2, 'Price in Values', 5000).catch(() => {});
    await js(`(() => {
      const sel = [...document.querySelectorAll('.sh-pf-area[data-area="values"] .sh-pf-item')].find((i) => i.dataset.field === 'Price').querySelector('.sh-pf-sum');
      const set = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set;
      set.call(sel, 'max');
      sel.dispatchEvent(new Event('change', { bubbles: true }));
      return 1;
    })()`);
    const maxed = await until(async () => (await pane())?.values?.includes('Max of Price'), 'Max of Price', 5000).then(() => true).catch(() => false);
    await capture('sheet-pivot-fields.png');
    check('worksheets: a ticked field of numbers goes in Values, and summarised by Max says so', maxed, JSON.stringify(await pane()));

    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => {
      try {
        const wb = Workbook.open(fs.readFileSync(file));
        const part = wb.pkg.partNames().find((p) => /pivotTables\/pivotTable\d+\.xml$/.test(p));
        const xml = wb.pkg.text(part);
        return /<colFields count="2"><field x="1"\/><field x="-2"\/><\/colFields>/.test(xml) && /<dataField name="Max of Price" fld="3" subtotal="max"/.test(xml);
      } catch { return false; }
    }, 'the saved layout', 8000).then(() => true).catch(() => false);
    check('worksheets: the saved pivot keeps its new layout for Excel to lay out on opening', saved, saved ? 'colFields and dataFields in the file' : 'not in the file');

    // The Filters area: Product moved there shows (All) above the pivot, then only the ink.
    const filters = () => js(`[...document.querySelectorAll('.sh-pf-area[data-area="filters"] .sh-pf-item')].map((i) => i.dataset.field)`);
    const setIn = (selector, value) => js(`(() => {
      const el = document.querySelector(${JSON.stringify(selector)});
      if (!el) return 'no ' + ${JSON.stringify(selector)};
      Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(el, ${JSON.stringify(value)});
      el.dispatchEvent(new Event('change', { bubbles: true }));
      return 'set';
    })()`);
    const toFilters = await setIn('.sh-pf-area[data-area="cols"] .sh-pf-item[data-field="Product"] .sh-pf-to', 'filters');
    const filtered = await until(async () => JSON.stringify(await filters()) === '["Product"]' && (await cell('G1')) === 'Product' && (await cell('H1')) === '(All)', 'Product in Filters', 6000).then(() => true).catch(() => false);
    check('worksheets: PivotTable Fields has a Filters area, and a field moved there shows (All) above the pivot',
      toFilters === 'set' && filtered, JSON.stringify({ toFilters, filters: await filters(), G1: await cell('G1'), H1: await cell('H1') }));
    const picked = await setIn('.sh-pf-area[data-area="filters"] .sh-pf-pick', 'Ink');
    const inkOnly = await until(async () => (await cell('H1')) === 'Ink' && (await cell('H6')) === '10', 'only the ink', 6000).then(() => true).catch(() => false);
    check('worksheets: a filter set to one item counts only that item',
      picked === 'set' && inkOnly, JSON.stringify({ picked, H1: await cell('H1'), H4: await cell('H4'), H5: await cell('H5'), H6: await cell('H6') }));

    // The pane's own Refresh: an edited source figure is counted when it is pressed, not before.
    const session = sessionFor('sheet');
    doc.apply({ id: session.id, ops: [{ op: 'setCell', row: 1, col: 2, value: '40' }] });
    await wait(400);
    const before = await cell('H4');
    const refreshed = await js(`(() => { const b = document.querySelector('.sh-pf-refresh'); if (!b) return 'no Refresh'; b.click(); return 'clicked'; })()`);
    const recounted = await until(async () => (await cell('H4')) === '40' && (await cell('H6')) === '46', 'the pivot refreshed', 6000).then(() => true).catch(() => false);
    check('worksheets: the PivotTable Fields pane has its own Refresh, which recounts the pivot from its source',
      before === '4' && refreshed === 'clicked' && recounted, JSON.stringify({ before, refreshed, H4: await cell('H4'), H6: await cell('H6') }));
    await capture('sheet-pivot-filters.png');

    // A new pivot under a long list: the sheet goes to it, its Fields pane open; its dates by month.
    const long = path.join(dir, 'pivot-long.xlsx');
    const serial = (m, day) => Math.round((Date.UTC(2026, m - 1, day) - Date.UTC(1899, 11, 30)) / 86400000);
    const list = [['Date', 'Item', 'Qty']];
    for (let i = 0; i < 60; i++) list.push([serial(1 + (i % 3), 1 + (i % 27)), 'Item ' + (i % 4), 1]);
    fs.writeFileSync(long, buildXlsx({ sheets: [{ name: 'List', rows: list, styles: { 'A2:A61': { numFmt: 'yyyy-mm-dd' } } }] }));
    const win2 = await open('sheets', long);
    const js2 = (code) => win2.webContents.executeJavaScript(code);
    const cell2 = (ref) => js2(`(document.querySelector('.sh-cell[data-ref="${ref}"]')?.textContent || '').trim()`);
    await until(() => js2(`Boolean(document.querySelector('.sh-cell[data-ref="A2"]'))`), 'the long list', 8000);
    const session2 = sessionFor('sheet');
    doc.apply({ id: session2.id, ops: [{ op: 'select', row: 1, col: 0 }] });
    await wait(300);
    await js2(`(() => { [...document.querySelectorAll('.rw-ribbon button, .rw-tabs button, button')].find((b) => b.textContent.trim() === 'Insert')?.click(); return 1; })()`);
    await until(() => js2(`Boolean([...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.textContent || '').trim() === 'PivotTable'))`), 'the PivotTable button', 4000).catch(() => {});
    await js2(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.textContent || '').trim() === 'PivotTable'); b?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b?.click(); return 1; })()`);
    const asked = await until(() => js2(`Boolean(document.querySelector('.sh-pivot-source'))`), 'the PivotTable dialog', 5000).then(() => true).catch(() => false);
    const setIn2 = (selector, value) => js2(`(() => {
      const el = document.querySelector(${JSON.stringify(selector)});
      if (!el) return 'no ' + ${JSON.stringify(selector)};
      Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(el, ${JSON.stringify(value)});
      el.dispatchEvent(new Event('change', { bubbles: true }));
      return 'set';
    })()`);
    await setIn2('.sh-pivot-rows', 'Date');
    await setIn2('.sh-pivot-values', 'Qty');
    await wait(200);
    await js2(`(() => { document.querySelector('.sh-pivot-ok')?.click(); return 1; })()`);
    const inSight = () => js2(`(() => {
      const g = document.querySelector('.sh-grid')?.getBoundingClientRect();
      const c = document.querySelector('.sh-cell[data-ref="A63"]')?.getBoundingClientRect();
      return Boolean(g && c && c.top >= g.top && c.bottom <= g.bottom);
    })()`);
    const shownNew = await until(async () => (await js2(`document.querySelector('.sh-pf')?.dataset.pivot || ''`)) === 'PivotTable1' && (await inSight()), 'the new pivot in sight', 8000).then(() => true).catch(() => false);
    const active = doc.model({ id: session2.id })?.selection?.active;
    check('worksheets: a new PivotTable made under a long list is scrolled to and picked, with PivotTable Fields open beside it',
      asked && shownNew && active?.row === 62 && active?.col === 0,
      JSON.stringify({ asked, shownNew, active, A63: await cell2('A63'), pane: await js2(`document.querySelector('.sh-pf')?.dataset.pivot || null`) }));

    const monthBox = await js2(`(() => { const b = document.querySelector('.sh-pf-area[data-area="rows"] .sh-pf-month input'); if (!b) return 'no By month'; b.click(); return 'clicked'; })()`);
    const months = await until(async () => (await cell2('A64')) === 'Jan' && (await cell2('A66')) === 'Mar' && (await cell2('A67')) === 'Grand Total', 'the months', 6000).then(() => true).catch(() => false);
    await win2.webContents.capturePage().then((img) => { if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'sheet-pivot-months.png'), img.toPNG()); });
    check('worksheets: a date field in a pivot can be shown by month, Jan to Mar in calendar order',
      monthBox === 'clicked' && months, JSON.stringify({ monthBox, A64: await cell2('A64'), A65: await cell2('A65'), A66: await cell2('A66'), A67: await cell2('A67'), B67: await cell2('B67') }));
    await wait(50);
  } catch (err) {
    check('worksheets: the PivotTable Fields checks ran', false, err.message);
  }
}
