// Worksheets: PivotTable Fields — a click in a pivot shows the pane beside
// the sheet, its fields ticked as the pivot uses them; Product moved to
// Columns lays the pivot out again at once, a ticked Price goes in Values,
// a value summarised by Max says so, and the saved file keeps the new
// layout for Excel to refresh.

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
    await wait(50);
  } catch (err) {
    check('worksheets: the PivotTable Fields checks ran', false, err.message);
  }
}
