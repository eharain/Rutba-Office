// Data Files' rough edges (2.18, Part A2 of the plan), checked in the
// window: a CSV of UK dates, its own IDs and its money imported, and a
// workbook's sheet chosen; a type change refused naming the records and
// cleared; a missing related record named; names in the datasheet; the
// default in the new row; a criterion naming another field; an average's
// heading and its Decimals; a query kept from the SQL view, its text kept;
// a form's Delete asking first; a sale named by its number and date; and a
// form, a template's record chooser and its Letterhead saying what they are
// for. Run alone with RUTBA_VERIFY_ONLY=base-fixes.

import fs from 'node:fs';
import path from 'node:path';
import { Base } from '@rutba/database';
import { buildXlsx } from '@rutba/ooxml/build';

const answer = (kind, value) => globalThis.__rutbaCheckDialogAnswers?.push({ kind, answer: value });

/** In-page helpers: a React field set as typing sets it, keys pressed on it, and what the toasts say. */
const HELPERS = `
  window.__f = window.__f || {
    set(el, value) {
      if (typeof el === 'string') el = document.querySelector(el);
      if (!el) return false;
      const proto = el.tagName === 'SELECT' ? HTMLSelectElement.prototype : el.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
      Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, value);
      el.dispatchEvent(new Event(el.tagName === 'SELECT' ? 'change' : 'input', { bubbles: true }));
      return true;
    },
    key(el, key) { if (typeof el === 'string') el = document.querySelector(el); el?.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true })); return Boolean(el); },
    click(sel) { const el = typeof sel === 'string' ? document.querySelector(sel) : sel; if (!el) return false; el.click(); return true; },
    text(sel) { return document.querySelector(sel)?.textContent || ''; },
    toasts() { return [...document.querySelectorAll('.rw-toast')].map((n) => n.textContent); },
  };
  true;
`;

export async function verifyBaseFixes(h, { dir }) {
  const { open, check, until, wait } = h;
  const ok = (p) => p.then(() => true, () => false);
  const file = path.join(dir, 'Larkspur fixes.rbase');
  const csv = path.join(dir, 'Members.csv');
  const xlsx = path.join(dir, 'Larkspur sales.xlsx');
  try {
    for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
    fs.writeFileSync(csv, 'ID,Name,Joined,Fee\n4,Amira,12/04/2025,1.10\n7,Ben,03/05/2025,12.00\n');
    fs.writeFileSync(xlsx, buildXlsx({ sheets: [
      { name: 'Sales', styles: { B2: { numFmt: 'dd/mm/yyyy' } }, rows: [['Number', 'Sold On'], [1, 45759]] },
      { name: 'Notes', rows: [['Note'], ['Open late on Friday'], ['Closed on Sunday']] },
    ] }));
    const b = Base.open(file);
    b.createTable({ name: 'Products', fields: [{ name: 'Name', type: 'text' }, { name: 'Price', type: 'currency', default: 1.5 }, { name: 'In Stock', type: 'integer' }, { name: 'Reorder Level', type: 'integer' }] });
    b.createTable({ name: 'Sales', fields: [{ name: 'Sold On', type: 'date' }, { name: 'Payment', type: 'text' }, { name: 'Total', type: 'currency' }] });
    b.createTable({ name: 'Sale Lines', fields: [{ name: 'Sale', type: 'integer' }, { name: 'Product', type: 'integer' }, { name: 'Quantity', type: 'integer' }] });
    b.createTable({ name: 'Visitors', fields: [{ name: 'Name', type: 'text' }, { name: 'Came', type: 'text' }] });
    b.addRelationship({ from: { table: 'Sale Lines', field: 'Product' }, to: { table: 'Products', field: 'ID' } });
    b.addRelationship({ from: { table: 'Sale Lines', field: 'Sale' }, to: { table: 'Sales', field: 'ID' } });
    for (const [n, p, s, r] of [['Milk', 1.1, 3, 5], ['Bread', 1.4, 10, 4], ['Tea', 3, 4, 4]]) b.insert('Products', { Name: n, Price: p, 'In Stock': s, 'Reorder Level': r });
    for (const [d, p, t] of [['2026-10-21', 'Cash', 2.2], ['2026-10-21', 'Cash', 1.4], ['2026-10-22', 'Card', 3]]) b.insert('Sales', { 'Sold On': d, Payment: p, Total: t });
    for (const [s, p, q] of [[1, 1, 2], [2, 2, 1], [3, 3, 1]]) b.insert('Sale Lines', { Sale: s, Product: p, Quantity: q });
    for (const [n, c] of [['Ann', '2025-04-12'], ['Bob', 'soon']]) b.insert('Visitors', { Name: n, Came: c });
    b.saveQuery('Low stock', { tables: ['Products'], columns: [{ table: 'Products', field: 'Name' }, { table: 'Products', field: 'In Stock', criteria: ['<= [Reorder Level]'] }] });
    b.makeForm('Products');
    b.makeForm('Sale Lines');
    b.makeDocument({ table: 'Sales' });
    b.close();

    const win = await open('base', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await js(HELPERS);
    await until(() => js(`Boolean(document.querySelector('.bs-tables [data-table="Sale Lines"]'))`), 'the database', 8000);

    // A2.4 and A2.6: names in the datasheet, the number kept behind them; the default in the new row.
    await js(`__f.click('.bs-tables [data-table="Sale Lines"]')`);
    await until(() => js(`[...document.querySelectorAll('.bs-body .bs-td[data-field="Product"]')].some((c) => c.textContent === 'Milk')`), 'the products by name', 5000).catch(() => {});
    const named = await js(`({ products: [...document.querySelectorAll('.bs-body .bs-td[data-field="Product"]')].map((c) => c.textContent), value: document.querySelector('.bs-body .bs-td[data-field="Product"] .bs-named')?.dataset.value, sales: [...document.querySelectorAll('.bs-body .bs-td[data-field="Sale"]')].map((c) => c.textContent) })`);
    check('base: the datasheet shows a related field as the record it points at (Milk, Sale 1, 21/10/2026), its number kept behind it',
      JSON.stringify(named.products) === JSON.stringify(['Milk', 'Bread', 'Tea']) && named.value === '1' && named.sales[0] === 'Sale 1, 21/10/2026' && named.sales[2] === 'Sale 3, 22/10/2026',
      JSON.stringify(named));

    // A2.3: a line pointing at product 99, refused naming it.
    await js(`__f.set('.bs-new-input[data-field="Product"]', '99')`);
    await wait(100);
    await js(`__f.set('.bs-new-input[data-field="Sale"]', '1')`);
    await wait(100);
    await js(`__f.key('.bs-new-input[data-field="Product"]', 'Enter')`);
    const missing = await ok(until(() => js(`__f.toasts().some((t) => /Products has no record 99, so Product cannot point at it/.test(t))`), 'the refusal', 5000));
    check('base: a line pointing at a product that is not there is refused as "Products has no record 99"', missing, JSON.stringify(await js(`__f.toasts()`)));

    await js(`__f.click('.bs-tables [data-table="Products"]')`);
    await until(() => js(`Boolean(document.querySelector('.bs-new-input[data-field="Price"]'))`), 'the products', 5000).catch(() => {});
    const placeholder = await js(`({ price: document.querySelector('.bs-new-input[data-field="Price"]')?.placeholder, name: document.querySelector('.bs-new-input[data-field="Name"]')?.placeholder })`);
    check("base: the new-record row shows a field's default (1.50) where a field without one shows its type",
      placeholder.price === (1.5).toLocaleString(undefined, { minimumFractionDigits: 2 }) && placeholder.name === 'Short text', JSON.stringify(placeholder));

    // A2.2: a type change a value does not fit, refused naming the record, then cleared.
    await js(`__f.click('.bs-tables [data-table="Visitors"]')`);
    await until(() => js(`Boolean(document.querySelector('.bs-tab[data-view="design"]'))`), 'the visitors', 4000).catch(() => {});
    await js(`__f.click('.bs-tab[data-view="design"]')`);
    await until(() => js(`Boolean(document.querySelector('.bs-fields tr[data-field="Came"] .bs-f-type'))`), 'the design', 4000).catch(() => {});
    await js(`__f.set('.bs-fields tr[data-field="Came"] .bs-f-type', 'date')`);
    await until(() => js(`Boolean(document.querySelector('.bs-unfit-words'))`), 'the question', 5000).catch(() => {});
    const unfit = await js(`__f.text('.bs-unfit-words')`);
    await js(`__f.click('.bs-unfit-clear')`);
    const cleared = await until(() => {
      const d = Base.open(file);
      try { return (d.fields('Visitors').find((f) => f.name === 'Came')?.type === 'date' && JSON.stringify(d.rows('Visitors').map((r) => r.Came)) === JSON.stringify(['2025-04-12', null])) || null; } finally { d.close(); }
    }, 'the change made', 6000).then(() => true, () => false);
    check('base: a type change a value does not fit says which record and value, and Clear them and change makes it with that one empty',
      unfit === '1 record of Visitors has a Came that is not a date: ID 2 ("soon"). Clear it and change the type, or put it right first.' && cleared,
      JSON.stringify({ unfit, cleared }));

    // A2.1, A2.9, A2.10: a CSV of UK dates, its own IDs and its fees with two places.
    answer('open', [csv]);
    await js(`__f.click('.bs-import')`);
    const csvIn = await ok(until(() => js(`Boolean(document.querySelector('.bs-tables [data-table="Members"]'))`), 'the members', 6000));
    const said = (await js(`__f.toasts()`)).join(' | ');
    await js(`__f.click('.bs-tables [data-table="Members"]')`);
    await until(() => js(`Boolean(document.querySelector('.bs-tab[data-view="design"]'))`), 'the members sheet', 3000).catch(() => {});
    await js(`__f.click('.bs-tab[data-view="design"]')`);
    await until(() => js(`Boolean(document.querySelector('.bs-fields tr[data-field="Joined"]'))`), 'the members design', 4000).catch(() => {});
    const types = await js(`Object.fromEntries([...document.querySelectorAll('.bs-fields tr[data-field]')].map((r) => [r.dataset.field, r.querySelector('select')?.value || 'autonumber']))`);
    const members = (() => { const d = Base.open(file); try { return d.rows('Members').map((r) => [r.ID, r.Joined, r.Fee]); } finally { d.close(); } })();
    check("base: Import reads a CSV's 12/04/2025 as 12 April and says so, keeps its own IDs as the key, and its fees with two places as Currency",
      csvIn && /Dates in Joined were read day first: 12\/04\/2025 is 12 April 2025/.test(said) && /keeps the file's own ID numbers/.test(said)
        && JSON.stringify(types) === JSON.stringify({ ID: 'autonumber', Name: 'text', Joined: 'date', Fee: 'currency' }) && JSON.stringify(members) === JSON.stringify([[4, '2025-04-12', 1.1], [7, '2025-05-03', 12]]),
      JSON.stringify({ csvIn, said, types, members }));

    // A2.5: a workbook of two sheets: the one chosen imported.
    answer('open', [xlsx]);
    await js(`__f.click('.bs-import')`);
    await until(() => js(`Boolean(document.querySelector('.bs-import-sheet'))`), 'the sheet chooser', 5000).catch(() => {});
    const offered = await js(`[...document.querySelectorAll('.bs-import-sheet option')].map((o) => o.value)`);
    await js(`__f.set('.bs-import-sheet', 'Notes')`);
    await wait(100);
    await js(`__f.click('.bs-import-sheets')`);
    const sheetIn = await ok(until(() => js(`/2/.test(document.querySelector('.bs-tables [data-table="Notes"] small')?.textContent || '')`), 'the Notes table', 6000));
    const noSales = await js(`!document.querySelector('.bs-tables [data-table="Sales 2"]')`);
    check('base: Import takes a workbook (.xlsx), asks which sheet (every one, or one), and makes a table of the one chosen',
      JSON.stringify(offered) === JSON.stringify(['', 'Sales', 'Notes']) && sheetIn && noSales, JSON.stringify({ offered, sheetIn, noSales }));

    // A2.7: a criterion naming another field.
    await js(`__f.click('.bs-queries [data-query="Low stock"]')`);
    await until(() => js(`document.querySelectorAll('.bs-query-run .bs-body .bs-row').length === 2`), 'the low stock', 5000).catch(() => {});
    const low = await js(`[...document.querySelectorAll('.bs-query-run .bs-body .bs-row')].map((r) => r.querySelector('[data-field="Name"]')?.textContent)`);
    check('base: a criterion naming another field (<= [Reorder Level]) compares with that field: Milk and Tea, not every product',
      JSON.stringify(low) === JSON.stringify(['Milk', 'Tea']), JSON.stringify(low));

    // A2.8: Decimals beside Total, and "Average of".
    await js(`__f.click('.bs-new-query')`);
    await until(() => js(`Boolean(document.querySelector('.bs-q-add-table'))`), 'the designer', 4000).catch(() => {});
    await js(`__f.set('.bs-q-add-table', 'Sales')`);
    await until(() => js(`document.querySelectorAll('.bs-q-field option').length >= 3`), 'the fields', 4000).catch(() => {});
    await js(`document.querySelector('.bs-q-totals input').click()`);
    await wait(150);
    await js(`__f.set(document.querySelectorAll('.bs-q-field')[0], 'Sales\\u0001Total')`);
    await wait(150);
    await js(`__f.set(document.querySelectorAll('.bs-q-total')[0], 'avg'); __f.set(document.querySelectorAll('.bs-q-decimals')[0], '2')`);
    await wait(150);
    await js(`__f.click('.bs-q-run')`);
    await until(() => js(`document.querySelectorAll('.bs-result tbody tr').length === 1`), 'the average', 5000).catch(() => {});
    const avg = await js(`({ head: __f.text('.bs-result thead th'), value: __f.text('.bs-result tbody td'), decimals: [...document.querySelectorAll('.bs-q-grid thead th')].map((th) => th.textContent) })`);
    check('base: with Totals ticked, Decimals stands beside Total and rounds an average, headed "Average of Total"',
      avg.head === 'Average of Total' && avg.value === '2.2' && avg.decimals.includes('Decimals'), JSON.stringify(avg));

    // A2.11: the SQL view keeps its text when left, and saves a query.
    await js(`__f.click('.bs-side-foot [data-view="sql"]')`);
    await until(() => js(`Boolean(document.querySelector('.bs-sql-text'))`), 'the SQL view', 4000).catch(() => {});
    const typed = 'SELECT "Payment", COUNT(*) AS "Sales" FROM "Sales" GROUP BY "Payment" ORDER BY "Payment"';
    await js(`__f.set('.bs-sql-text', ${JSON.stringify(typed)})`);
    await js(`__f.click('.bs-side-foot [data-view="relations"]')`);
    await wait(200);
    await js(`__f.click('.bs-side-foot [data-view="sql"]')`);
    await until(() => js(`Boolean(document.querySelector('.bs-sql-text'))`), 'the SQL view again', 4000).catch(() => {});
    const kept = await js(`document.querySelector('.bs-sql-text')?.value`);
    await js(`__f.click('.bs-sql-save')`);
    await until(() => js(`Boolean(document.querySelector('.bs-sql-name'))`), 'the name box', 4000).catch(() => {});
    await js(`__f.set('.bs-sql-name', 'By payment')`);
    await wait(100);
    await js(`__f.click('.bs-sql-save-named')`);
    const listed = await ok(until(() => js(`document.querySelectorAll('.bs-query-run .bs-body .bs-row').length === 2 && Boolean(document.querySelector('.bs-queries [data-query="By payment"]'))`), 'the saved query run', 6000));
    const ran = await js(`[...document.querySelectorAll('.bs-query-run .bs-body .bs-row')].map((r) => [...r.querySelectorAll('.bs-td')].map((c) => c.textContent))`);
    await js(`__f.click('.bs-q-edit')`);
    const back = await ok(until(() => js(`document.querySelector('.bs-sql-text')?.value === ${JSON.stringify(typed)}`), 'its SQL again', 4000));
    check('base: the SQL view keeps its text when it is left and come back to, Save As keeps the query in the list, run, and Edit query opens its SQL',
      kept === typed && listed && JSON.stringify(ran) === JSON.stringify([['Card', '1'], ['Cash', '2']]) && back, JSON.stringify({ kept, listed, ran, back }));

    // A2.14 and A2.17: a form's Delete asks first; a form with no print template says how to make one.
    await js(`__f.click('.bs-forms [data-form="Products"]')`);
    await until(() => js(`/Record 1 of 3/.test(__f.text('.bs-nav-place'))`), 'the Products form', 5000).catch(() => {});
    const noPrint = await js(`__f.text('.bs-form-no-print')`);
    await js(`__f.click('.bs-nav-delete')`);
    const asked = await ok(until(() => js(`Boolean(document.querySelector('.bs-delete-yes'))`), 'the question', 3000));
    await js(`__f.click('.bs-delete-cancel')`);
    await wait(200);
    const still = await js(`__f.text('.bs-nav-place')`);
    await js(`__f.click('.bs-nav-delete')`);
    await until(() => js(`Boolean(document.querySelector('.bs-delete-yes'))`), 'the question again', 3000).catch(() => {});
    await js(`__f.click('.bs-delete-yes')`);
    // Milk has a sale line: the database names it, in plain words.
    const refused = await ok(until(() => js(`__f.toasts().some((t) => /1 record of Sale Lines points at this one \\(ID 1\\)/.test(t))`), 'the refusal', 5000));
    check('base: a form\'s Delete asks first, Cancel keeps the record, and a product a sale line points at is refused naming the line; a form with no print template says how to make one',
      asked && /Record 1 of 3/.test(still) && refused && /make a print template for Products/.test(noPrint),
      JSON.stringify({ asked, still, refused, noPrint, toasts: await js(`__f.toasts()`) }));

    // A2.16: a sale named by its number and date in a form's list.
    await js(`__f.click('.bs-forms [data-form="Sale Lines"]')`);
    await until(() => js(`document.querySelectorAll('.bs-fc-list[data-field="Sale"] option').length >= 4`), 'the sales list', 5000).catch(() => {});
    const sales = await js(`[...document.querySelectorAll('.bs-fc-list[data-field="Sale"] option')].map((o) => o.textContent)`);
    check('base: a sale named by nothing of its own is listed by its number and date ("Sale 1, 21/10/2026"), not by its payment',
      sales.includes('Sale 1, 21/10/2026') && sales.includes('Sale 3, 22/10/2026') && !sales.includes('Cash'), JSON.stringify(sales));

    // A2.17: a template's record chooser names records; Letterhead says what it is for.
    await js(`__f.click('.bs-documents [data-document="Sale"]')`);
    await until(() => js(`document.querySelectorAll('.bs-doc-record option').length >= 4`), 'the record chooser', 5000).catch(() => {});
    const chooser = await js(`[...document.querySelectorAll('.bs-doc-record option')].map((o) => o.textContent)`);
    await js(`__f.click('.bs-doc-design')`);
    await until(() => js(`Boolean(document.querySelector('.bs-dd-letterhead'))`), 'the design panel', 4000).catch(() => {});
    const hint = await js(`__f.text('.bs-dd-letterhead-hint')`);
    check("base: a print template's record chooser names each record as its lists do, and Letterhead says what it is for when there is no Business table",
      JSON.stringify(chooser.slice(0, 2)) === JSON.stringify(['Every record', 'Sale 1, 21/10/2026']) && /make a table called Business/.test(hint), JSON.stringify({ chooser, hint }));
  } catch (err) {
    check('base: the rough-edge checks ran', false, err.stack || err.message);
  }
}
