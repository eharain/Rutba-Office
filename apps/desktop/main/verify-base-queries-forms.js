// Data Files' queries and forms, checked in the window: a query designed
// across two related tables with totals, run and kept under a name; the SQL
// it stands for shown; a form made by the wizard for Orders, its customer a
// list of the customers by name, moving between records, a change kept, a
// record added; and a form for Customers listing each one's orders below.

import fs from 'node:fs';
import path from 'node:path';
import { Base } from '@rutba/database';

/** In-page helpers: a React field set as typing sets it, left as a click elsewhere leaves it. */
const HELPERS = `
  window.__q = window.__q || {
    set(el, value) {
      if (typeof el === 'string') el = document.querySelector(el);
      if (!el) return false;
      const proto = el.tagName === 'SELECT' ? HTMLSelectElement.prototype : el.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
      Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, value);
      el.dispatchEvent(new Event(el.tagName === 'SELECT' ? 'change' : 'input', { bubbles: true }));
      return true;
    },
    leave(el) { if (typeof el === 'string') el = document.querySelector(el); el?.dispatchEvent(new FocusEvent('focusout', { bubbles: true })); return Boolean(el); },
    click(sel) { const el = document.querySelector(sel); if (!el) return false; el.click(); return true; },
    text(sel) { return document.querySelector(sel)?.textContent || ''; },
  };
  true;
`;

export async function verifyBaseQueriesForms(h, { dir }) {
  const { open, check, until, wait } = h;
  const file = path.join(dir, 'Shop2.rbase');
  try {
    for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
    const b = Base.open(file);
    b.createTable({ name: 'Customers', fields: [{ name: 'Name', type: 'text' }, { name: 'City', type: 'text' }] });
    b.createTable({ name: 'Orders', fields: [{ name: 'Customer', type: 'integer' }, { name: 'Amount', type: 'currency' }] });
    b.addRelationship({ from: { table: 'Orders', field: 'Customer' }, to: { table: 'Customers', field: 'ID' } });
    for (const [n, c] of [['Amira', 'Lahore'], ['Ben', 'London']]) b.insert('Customers', { Name: n, City: c });
    for (const [c, a] of [[1, 100], [1, 250], [2, 40]]) b.insert('Orders', { Customer: c, Amount: a });
    b.close();

    const win = await open('base', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await js(HELPERS);
    await until(() => js(`Boolean(document.querySelector('.bs-new-query'))`), 'the database', 8000);

    // A query: Customers and Orders, each customer's total, largest first.
    await js(`__q.click('.bs-new-query')`);
    await until(() => js(`Boolean(document.querySelector('.bs-q-add-table'))`), 'the designer', 4000).catch(() => {});
    await js(`__q.set('.bs-q-add-table', 'Customers')`);
    await wait(150);
    await js(`__q.set('.bs-q-add-table', 'Orders')`);
    await until(() => js(`document.querySelectorAll('.bs-q-field option').length >= 4`), 'the fields', 4000).catch(() => {});
    await js(`(() => { const t = document.querySelector('.bs-q-totals input'); t.click(); return 1; })()`);
    await wait(150);
    await js(`__q.set(document.querySelectorAll('.bs-q-field')[0], 'Customers\\u0001Name')`);
    await wait(150);
    await js(`__q.set(document.querySelectorAll('.bs-q-field')[1], 'Orders\\u0001Amount')`);
    await wait(150);
    await js(`__q.set(document.querySelectorAll('.bs-q-total')[1], 'sum'); __q.set(document.querySelectorAll('.bs-q-sort')[1], 'desc')`);
    await wait(150);
    await js(`__q.click('.bs-q-run')`);
    await until(() => js(`document.querySelectorAll('.bs-result tbody tr').length === 2`), 'the answer', 5000).catch(() => {});
    const answer = await js(`[...document.querySelectorAll('.bs-result tbody tr')].map((r) => [...r.cells].map((c) => c.textContent))`);
    const heads = await js(`[...document.querySelectorAll('.bs-result thead th')].map((c) => c.textContent)`);
    await js(`__q.click('.bs-q-sql')`);
    await until(() => js(`Boolean(document.querySelector('.bs-q-sql-text'))`), 'the SQL', 4000).catch(() => {});
    const sql = await js(`__q.text('.bs-q-sql-text')`);
    check("base: a query designed across two related tables, with totals and a sort, answers each customer's total, largest first, and shows the SQL it stands for",
      JSON.stringify(answer) === JSON.stringify([['Amira', '350'], ['Ben', '40']]) && heads.includes('Sum of Amount') && /JOIN "Customers"|JOIN "Orders"/.test(sql) && /GROUP BY/.test(sql),
      JSON.stringify({ answer, heads, sql: sql.slice(0, 160) }));

    await js(`__q.click('.bs-q-save')`);
    await until(() => js(`Boolean(document.querySelector('.bs-q-name'))`), 'the name box', 4000).catch(() => {});
    await js(`__q.set('.bs-q-name', 'Totals')`);
    await wait(100);
    await js(`__q.click('.bs-q-save-named')`);
    const kept = await until(() => js(`Boolean(document.querySelector('.bs-queries [data-query="Totals"]'))`), 'the query in the list', 5000).then(() => true, () => false);
    check('base: Save As keeps the query in the file under its name, listed with the queries', kept, String(kept));

    // Opened again from the list it runs: its records, then Edit query for its design, and Show results back (D-2026-10-22-06).
    await js(`__q.click('.bs-queries [data-query="Totals"]')`);
    await until(() => js(`document.querySelectorAll('.bs-query-run .bs-result tbody tr').length === 2`), 'the query run', 5000).catch(() => {});
    const ran = await js(`({ rows: document.querySelectorAll('.bs-query-run .bs-result tbody tr').length, title: __q.text('.bs-q-title'), count: __q.text('.bs-query-run .bs-count'), designer: Boolean(document.querySelector('.bs-q-grid')) })`);
    await js(`__q.click('.bs-q-edit')`);
    await until(() => js(`Boolean(document.querySelector('.bs-q-grid'))`), 'the design', 4000).catch(() => {});
    const designing = await js(`({ designer: Boolean(document.querySelector('.bs-q-grid')), tables: [...document.querySelectorAll('.bs-q-chip')].map((n) => n.textContent.trim()) })`);
    await js(`__q.click('.bs-q-results')`);
    const back = await until(() => js(`document.querySelectorAll('.bs-query-run .bs-result tbody tr').length === 2`), 'the records again', 5000).then(() => true, () => false);
    check('base: a saved query opens run, its records and how many, with Edit query opening its design and Show results going back',
      ran.rows === 2 && ran.title === 'Totals' && /2 rows/.test(ran.count) && !ran.designer && designing.designer && designing.tables.length === 2 && back,
      JSON.stringify({ ran, designing, back }));

    // A calculated column, typed as Access takes it in a Field cell.
    await js(`__q.click('.bs-new-query')`);
    await until(() => js(`Boolean(document.querySelector('.bs-q-add-table'))`), 'a new designer', 4000).catch(() => {});
    await js(`__q.set('.bs-q-add-table', 'Orders')`);
    await until(() => js(`document.querySelectorAll('.bs-q-field option').length >= 3`), 'the fields', 4000).catch(() => {});
    await js(`__q.set(document.querySelectorAll('.bs-q-field')[0], 'Orders\\u0001Amount')`);
    await wait(150);
    await js(`__q.set(document.querySelectorAll('.bs-q-field')[1], '\\u0002expression')`);
    await until(() => js(`Boolean(document.querySelector('.bs-q-expr'))`), 'the expression box', 4000).catch(() => {});
    await js(`__q.set('.bs-q-expr', 'Double: [Amount] * 2')`);
    await wait(150);
    await js(`__q.click('.bs-q-run')`);
    await until(() => js(`document.querySelectorAll('.bs-result tbody tr').length === 3`), 'the calculated answer', 5000).catch(() => {});
    const calc = await js(`({ heads: [...document.querySelectorAll('.bs-result thead th')].map((c) => c.textContent), rows: [...document.querySelectorAll('.bs-result tbody tr')].map((r) => [...r.cells].map((c) => c.textContent)) })`);
    check('base: Expression in a Field cell makes a calculated column named before its colon ("Double: [Amount] * 2"), and Run answers it',
      calc.heads.includes('Double') && JSON.stringify(calc.rows.map((r) => r.slice(-2))) === JSON.stringify([['100', '200'], ['250', '500'], ['40', '80']]),
      JSON.stringify(calc));

    // A form for Orders, by the wizard.
    await js(`__q.click('.bs-new-form')`);
    await until(() => js(`Boolean(document.querySelector('.bs-form-table'))`), 'the New Form box', 4000).catch(() => {});
    await js(`__q.set('.bs-form-table', 'Orders')`);
    await wait(100);
    await js(`__q.click('.bs-make-form')`);
    await until(() => js(`Boolean(document.querySelector('.bs-form-page[data-form="Orders"]'))`), 'the form', 6000).catch(() => {});
    await until(() => js(`document.querySelectorAll('.bs-fc-list option').length >= 3`), 'the customers list', 4000).catch(() => {});
    const shown = await js(`({ place: __q.text('.bs-nav-place'), amount: document.querySelector('.bs-fc[data-field="Amount"]')?.value, customer: document.querySelector('.bs-fc-list')?.selectedOptions[0]?.textContent, choices: [...document.querySelectorAll('.bs-fc-list option')].map((o) => o.textContent) })`);
    await js(`__q.click('.bs-nav-next')`);
    await until(() => js(`/Record 2 of 3/.test(__q.text('.bs-nav-place'))`), 'record 2', 4000).catch(() => {});
    await js(`(() => { const a = document.querySelector('.bs-fc[data-field="Amount"]'); __q.set(a, '275'); __q.leave(a); return 1; })()`);
    const changed = await until(() => {
      const d = Base.open(file);
      try { return d.rows('Orders', { offset: 1, limit: 1 })[0]?.Amount === 275 || null; } finally { d.close(); }
    }, 'the change in the file', 5000).then(() => true, () => false);
    check('base: the form wizard makes a form for Orders, its customer a list of the customers by name; Next moves to the next record, and a changed amount is kept when the box is left',
      shown.place === 'Record 1 of 3' && shown.amount === '100' && shown.customer === 'Amira' && shown.choices.includes('Ben') && changed,
      JSON.stringify({ shown, changed }));

    await js(`__q.click('.bs-nav-new')`);
    await until(() => js(`Boolean(document.querySelector('.bs-nav-add'))`), 'a new record', 4000).catch(() => {});
    await js(`(() => { const l = document.querySelector('.bs-fc-list'); __q.set(l, '2'); const a = document.querySelector('.bs-fc[data-field="Amount"]'); __q.set(a, '60'); __q.leave(a); return 1; })()`);
    await wait(150);
    await js(`__q.click('.bs-nav-add')`);
    const added = await until(() => js(`/Record 4 of 4/.test(__q.text('.bs-nav-place'))`), 'the record added', 5000).then(() => true, () => false);
    check('base: New Record and Add Record put a record in the table, and the form shows it', added, await js(`__q.text('.bs-nav-place')`));

    // A form for Customers: each customer's orders below.
    await js(`__q.click('.bs-new-form')`);
    await until(() => js(`Boolean(document.querySelector('.bs-form-table'))`), 'the New Form box', 4000).catch(() => {});
    await js(`__q.set('.bs-form-table', 'Customers')`);
    await wait(100);
    await js(`__q.click('.bs-make-form')`);
    await until(() => js(`document.querySelectorAll('.bs-subform[data-subform="Orders"] tbody tr').length === 2`), "Amira's orders", 6000).catch(() => {});
    const orders = await js(`[...document.querySelectorAll('.bs-subform[data-subform="Orders"] tbody tr')].map((r) => r.cells[0].textContent)`);
    check("base: a form for Customers lists each customer's orders below, from the relationship", JSON.stringify(orders.sort()) === JSON.stringify(['100', '275']), JSON.stringify(orders));
  } catch (err) {
    check('base: the query and form checks ran', false, err.stack || err.message);
  }
}
