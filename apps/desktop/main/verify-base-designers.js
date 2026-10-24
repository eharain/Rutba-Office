// Data Files' 2.18 designers (Part A2 of the plan), checked in the window:
// a form's Design moving and sizing its fields (dragged, and with the arrow
// keys), renaming a label and taking a field off, saved with the form; a
// subform's lines changed, added at its foot and deleted; and a report's
// Totals chosen for each number, a whole number and a query's too. Run alone
// with RUTBA_VERIFY_ONLY=base-designers.

import fs from 'node:fs';
import path from 'node:path';
import { Base } from '@rutba/database';

/** In-page helpers: a React field set as typing sets it, keys and drags as a person makes them, a box left. */
const HELPERS = `
  window.__d = window.__d || {
    el(sel) { return typeof sel === 'string' ? document.querySelector(sel) : sel; },
    set(sel, value) {
      const el = this.el(sel);
      if (!el) return false;
      const proto = el.tagName === 'SELECT' ? HTMLSelectElement.prototype : el.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
      Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, value);
      el.dispatchEvent(new Event(el.tagName === 'SELECT' ? 'change' : 'input', { bubbles: true }));
      return true;
    },
    leave(sel) { const el = this.el(sel); if (!el) return false; el.dispatchEvent(new FocusEvent('focusout', { bubbles: true })); return true; },
    key(sel, key, shiftKey = false) { const el = this.el(sel); if (!el) return false; el.dispatchEvent(new KeyboardEvent('keydown', { key, shiftKey, bubbles: true })); return true; },
    drag(sel, dx, dy) {
      const el = this.el(sel);
      if (!el) return false;
      const r = el.getBoundingClientRect();
      const at = (x, y) => ({ bubbles: true, cancelable: true, clientX: x, clientY: y, button: 0, buttons: 1, pointerId: 7, pointerType: 'mouse', isPrimary: true });
      const x0 = r.left + 6, y0 = r.top + 6;
      el.dispatchEvent(new PointerEvent('pointerdown', at(x0, y0)));
      el.dispatchEvent(new PointerEvent('pointermove', at(x0 + dx / 2, y0 + dy / 2)));
      el.dispatchEvent(new PointerEvent('pointermove', at(x0 + dx, y0 + dy)));
      el.dispatchEvent(new PointerEvent('pointerup', at(x0 + dx, y0 + dy)));
      return true;
    },
    click(sel) { const el = this.el(sel); if (!el) return false; el.click(); return true; },
    text(sel) { return this.el(sel)?.textContent || ''; },
    toasts() { return [...document.querySelectorAll('.rw-toast')].map((n) => n.textContent); },
  };
  true;
`;

export async function verifyBaseDesigners(h, { dir }) {
  const { open, check, until, wait } = h;
  const ok = (p) => p.then(() => true, () => false);
  const file = path.join(dir, 'Larkspur designers.rbase');
  const read = (fn) => { const d = Base.open(file); try { return fn(d); } finally { d.close(); } };
  try {
    for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
    const b = Base.open(file);
    b.createTable({ name: 'Products', fields: [{ name: 'Name', type: 'text' }, { name: 'Price', type: 'currency' }, { name: 'In Stock', type: 'integer' }, { name: 'Notes', type: 'longtext' }] });
    b.createTable({ name: 'Sales', fields: [{ name: 'Sold On', type: 'date' }, { name: 'Payment', type: 'text' }] });
    b.createTable({ name: 'Sale Lines', fields: [{ name: 'Sale', type: 'integer' }, { name: 'Product', type: 'integer' }, { name: 'Quantity', type: 'integer' }] });
    b.addRelationship({ from: { table: 'Sale Lines', field: 'Product' }, to: { table: 'Products', field: 'ID' } });
    b.addRelationship({ from: { table: 'Sale Lines', field: 'Sale' }, to: { table: 'Sales', field: 'ID' } });
    for (const [n, p, s] of [['Milk', 1.1, 3], ['Bread', 1.4, 10], ['Tea', 3, 4]]) b.insert('Products', { Name: n, Price: p, 'In Stock': s });
    for (const d of ['2026-10-21', '2026-10-22']) b.insert('Sales', { 'Sold On': d, Payment: 'Cash' });
    for (const [s, p, q] of [[1, 1, 2], [1, 2, 1], [2, 3, 5]]) b.insert('Sale Lines', { Sale: s, Product: p, Quantity: q });
    b.makeForm('Products');
    b.makeForm('Sales');
    b.makeReport({ table: 'Products' });
    b.saveQuery('Stock', { tables: ['Products'], columns: [{ table: 'Products', field: 'Name' }, { table: 'Products', field: 'In Stock' }, { table: 'Products', field: 'Price' }] });
    b.makeReport({ query: 'Stock' });
    const before = b.forms().find((f) => f.name === 'Products').layout;
    b.close();
    const was = Object.fromEntries(before.controls.map((c) => [c.field, c]));
    // Each step is snapped to the form's grid of 8: Currency's 140 wide, two steps wider, is 160.
    const grid = (v) => Math.round(v / 8) * 8;

    const win = await open('base', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await js(HELPERS);
    await until(() => js(`Boolean(document.querySelector('.bs-forms [data-form="Products"]'))`), 'the database', 8000);

    // A2.12: a form's Design.
    await js(`__d.click('.bs-forms [data-form="Products"]')`);
    await until(() => js(`Boolean(document.querySelector('.bs-form-design'))`), 'the form', 5000).catch(() => {});
    await js(`__d.click('.bs-form-design')`);
    await until(() => js(`Boolean(document.querySelector('.bs-fd-box[data-field="Price"]'))`), 'the design', 5000).catch(() => {});
    // Price: three grid steps right, two wider, by the arrow keys.
    for (let i = 0; i < 3; i++) await js(`__d.key('.bs-fd-box[data-field="Price"]', 'ArrowRight')`);
    for (let i = 0; i < 2; i++) await js(`__d.key('.bs-fd-box[data-field="Price"]', 'ArrowRight', true)`);
    // Name: dragged 83 across and 41 down, snapped to 80 and 40.
    await js(`__d.drag('.bs-fd-box[data-field="Name"]', 83, 41)`);
    await wait(150);
    // In Stock's label renamed; Notes taken off.
    await js(`__d.set('.bs-fd-label[data-field="In Stock"]', 'Units on the shelf')`);
    await js(`__d.leave('.bs-fd-label[data-field="In Stock"]')`);
    await js(`document.querySelector('.bs-fd-box[data-field="Notes"]').dispatchEvent(new FocusEvent('focusin', { bubbles: true }))`);
    await wait(100);
    await js(`__d.click('.bs-fd-remove')`);
    await wait(150);
    const offered = await js(`[...document.querySelectorAll('.bs-fd-add option')].map((o) => o.value).filter(Boolean)`);
    const where = await js(`__d.text('.bs-fd-where')`);
    await js(`__d.click('.bs-fd-save')`);
    const saved = await ok(until(() => js(`Boolean(document.querySelector('.bs-form-page .bs-fc[data-field="Price"]')) && !document.querySelector('.bs-fd')`), 'the form again', 6000));
    const kept = read((d) => d.forms().find((f) => f.name === 'Products').layout);
    const at = Object.fromEntries(kept.controls.map((c) => [c.field, c]));
    const shown = await js(`({ labels: [...document.querySelectorAll('.bs-form-page .bs-fl')].map((n) => n.textContent), notes: Boolean(document.querySelector('.bs-fc[data-field="Notes"]')), price: document.querySelector('.bs-fc[data-field="Price"]')?.style.left, width: document.querySelector('.bs-fc[data-field="Price"]')?.style.width })`);
    check("base: a form's Design moves a field with the arrow keys and sizes it with Shift, drags another on the grid, renames a label and takes a field off, kept with the form",
      saved && at.Price.x === was.Price.x + 24 && at.Price.w === grid(grid(was.Price.w + 8) + 8) && at.Name.x === was.Name.x + 80 && at.Name.y === was.Name.y + 40
        && at['In Stock'].label === 'Units on the shelf' && !at.Notes && JSON.stringify(offered) === JSON.stringify(['Notes'])
        && shown.labels.includes('Units on the shelf') && !shown.notes && shown.price === `${was.Price.x + 24}px` && shown.width === `${at.Price.w}px`,
      JSON.stringify({ saved, where, offered, shown, kept: kept.controls.map((c) => [c.field, c.x, c.y, c.w, c.label]) }));

    // A field put back, and Cancel leaves the form as it was kept.
    await js(`__d.click('.bs-form-design')`);
    await until(() => js(`Boolean(document.querySelector('.bs-fd-add'))`), 'the design again', 4000).catch(() => {});
    await js(`__d.set('.bs-fd-add', 'Notes')`);
    await until(() => js(`Boolean(document.querySelector('.bs-fd-box[data-field="Notes"]'))`), 'Notes back', 3000).catch(() => {});
    const putBack = await js(`Boolean(document.querySelector('.bs-fd-box[data-field="Notes"]'))`);
    await js(`__d.click('.bs-fd-cancel')`);
    await until(() => js(`!document.querySelector('.bs-fd')`), 'the form', 3000).catch(() => {});
    const still = read((d) => d.forms().find((f) => f.name === 'Products').layout.controls.some((c) => c.field === 'Notes'));
    check("base: Design's Put a field on puts a field taken off back, and Cancel leaves the form as it was kept",
      putBack && !still && !(await js(`Boolean(document.querySelector('.bs-fc[data-field="Notes"]'))`)), JSON.stringify({ putBack, still }));

    // A2.13: a subform's lines changed, added and deleted.
    await js(`__d.click('.bs-forms [data-form="Sales"]')`);
    await until(() => js(`document.querySelectorAll('.bs-subform[data-subform="Sale Lines"] tbody tr[data-rowid]').length === 2`), 'the lines', 5000).catch(() => {});
    // A value reads as words until it is clicked; then it is a box to change it in.
    const asWords = await js(`__d.text('.bs-subform[data-subform="Sale Lines"] tbody tr[data-rowid] td[data-col="Quantity"]')`);
    await js(`__d.click('.bs-subform[data-subform="Sale Lines"] tbody tr[data-rowid] td[data-col="Quantity"] .bs-sub-text')`);
    await until(() => js(`Boolean(document.querySelector('.bs-subform[data-subform="Sale Lines"] tbody .bs-sub-cell[data-col="Quantity"]'))`), 'the box', 3000).catch(() => {});
    const first = await js(`document.querySelector('.bs-subform[data-subform="Sale Lines"] tbody tr[data-rowid] .bs-sub-cell[data-col="Quantity"]')?.value`);
    await js(`__d.set('.bs-subform[data-subform="Sale Lines"] tbody tr[data-rowid] .bs-sub-cell[data-col="Quantity"]', '6')`);
    await js(`__d.leave('.bs-subform[data-subform="Sale Lines"] tbody tr[data-rowid] .bs-sub-cell[data-col="Quantity"]')`);
    const changed = await ok(until(() => read((d) => d.rows('Sale Lines')[0].Quantity === 6) || null, 'the change kept', 4000));
    const productChoices = await js(`[...document.querySelectorAll('.bs-sub-new[data-col="Product"] option')].map((o) => o.textContent)`);
    await js(`__d.set('.bs-sub-new[data-col="Product"]', '3')`);
    await wait(80);
    await js(`__d.set('.bs-sub-new[data-col="Quantity"]', '5')`);
    await wait(80);
    await js(`__d.click('.bs-sub-add')`);
    const added = await ok(until(() => js(`document.querySelectorAll('.bs-subform[data-subform="Sale Lines"] tbody tr[data-rowid]').length === 3`), 'the line added', 5000));
    const lines = read((d) => d.rows('Sale Lines').map((r) => [r.Sale, r.Product, r.Quantity]));
    const words = await js(`[...document.querySelectorAll('.bs-subform[data-subform="Sale Lines"] tbody tr[data-rowid] td[data-col="Product"]')].map((n) => n.textContent)`);
    const head = await js(`__d.text('.bs-subform[data-subform="Sale Lines"] .bs-subform-head small')`);
    check("base: a subform's line is changed where it stands, and a new one typed at its foot is added to the sale shown, its product a list by name",
      asWords === '2' && first === '2' && changed && added && JSON.stringify(lines) === JSON.stringify([[1, 1, 6], [1, 2, 1], [2, 3, 5], [1, 3, 5]])
        && JSON.stringify(productChoices) === JSON.stringify(['(none)', 'Bread', 'Milk', 'Tea']) && JSON.stringify(words) === JSON.stringify(['Milk', 'Bread', 'Tea']) && head === '3 records',
      JSON.stringify({ asWords, first, changed, added, lines, productChoices, words, head }));

    await js(`__d.click([...document.querySelectorAll('.bs-subform[data-subform="Sale Lines"] tbody tr[data-rowid] .bs-sub-remove')].at(-1))`);
    const asked = await ok(until(() => js(`Boolean(document.querySelector('.bs-sub-remove-yes'))`), 'the question', 3000));
    await js(`__d.click('.bs-sub-remove-yes')`);
    const gone = await ok(until(() => js(`document.querySelectorAll('.bs-subform[data-subform="Sale Lines"] tbody tr[data-rowid]').length === 2`), 'the line deleted', 5000));
    check("base: a subform's line is deleted after asking", asked && gone && read((d) => d.rows('Sale Lines').length) === 3, JSON.stringify({ asked, gone }));

    // A2.15: a report's totals chosen.
    await js(`__d.click('.bs-reports [data-report="Products"]')`);
    await until(() => js(`Boolean(document.querySelector('.bs-report-totals'))`), 'the report', 5000).catch(() => {});
    await js(`__d.click('.bs-report-totals')`);
    await until(() => js(`Boolean(document.querySelector('.bs-report-total-fn'))`), 'the totals', 3000).catch(() => {});
    const choices = await js(`[...document.querySelectorAll('[data-total-field]')].map((n) => [n.dataset.totalField, n.querySelector('select').value])`);
    await js(`__d.set('[data-total-field="In Stock"] select', 'avg')`);
    await until(() => js(`/Average/.test(__d.text('.bs-report-scroll .bs-report-total'))`), 'the average', 4000).catch(() => {});
    await js(`__d.set('[data-total-field="Price"] select', '')`);
    await until(() => js(`!/Sum/.test(__d.text('.bs-report-scroll .bs-report-total'))`), 'the sum gone', 4000).catch(() => {});
    await js(`__d.click('.bs-report-totals-done')`);
    const grand = await js(`__d.text('.bs-report-scroll .bs-report-total')`);
    const keptTotals = read((d) => d.reports().find((r) => r.name === 'Products').layout.totals);
    check("base: a report's Totals offers each number, a whole number too, and Average on In Stock with none on Price is drawn and kept",
      JSON.stringify(choices) === JSON.stringify([['Price', 'sum'], ['In Stock', '']]) && /Average\s*5\.67/.test(grand) && !/Sum/.test(grand)
        && JSON.stringify(keptTotals) === JSON.stringify([{ field: 'In Stock', fn: 'avg' }]),
      JSON.stringify({ choices, grand, keptTotals }));

    await js(`__d.click('.bs-reports [data-report="Stock"]')`);
    await until(() => js(`document.querySelector('.bs-report-page')?.dataset.report === 'Stock' && Boolean(document.querySelector('.bs-report-totals'))`), 'the query report', 5000).catch(() => {});
    await js(`__d.click('.bs-report-totals')`);
    await until(() => js(`Boolean(document.querySelector('.bs-report-total-fn'))`), 'its totals', 3000).catch(() => {});
    const fromQuery = await js(`[...document.querySelectorAll('[data-total-field]')].map((n) => n.dataset.totalField)`);
    await js(`__d.set('[data-total-field="In Stock"] select', 'sum')`);
    await until(() => js(`/17/.test(__d.text('.bs-report-scroll .bs-report-total'))`), 'the sum', 4000).catch(() => {});
    const queryGrand = await js(`__d.text('.bs-report-scroll .bs-report-total')`);
    await js(`__d.click('.bs-report-totals-done')`);
    check("base: a report of a query offers Totals for its numbers too, and sums its whole numbers",
      JSON.stringify(fromQuery) === JSON.stringify(['In Stock', 'Price']) && /Sum\s*17/.test(queryGrand), JSON.stringify({ fromQuery, queryGrand }));
  } catch (err) {
    check('base: the designer checks ran', false, err.stack || err.message);
  }
}
