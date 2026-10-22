// Searching and filtering Data Files' records in the window (owner,
// 2026-10-22, D-2026-10-22-07): a table's records and a saved query's
// results alike, searched for words, filtered by a column's type (a date
// range, a number range) and by selection from a cell's menu, each filter a
// chip with its cross, and the count saying how many were found. Run alone
// with RUTBA_VERIFY_ONLY=base-filter.

import fs from 'node:fs';
import path from 'node:path';
import { Base, buildFromSpec } from '@rutba/database';

const set = (sel, value) => `(() => { const el = document.querySelector(${JSON.stringify(sel)}); const proto = el instanceof HTMLSelectElement ? HTMLSelectElement.prototype : HTMLInputElement.prototype; Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, ${JSON.stringify(value)}); el.dispatchEvent(new Event(el instanceof HTMLSelectElement ? 'change' : 'input', { bubbles: true })); return 1; })()`;
const rows = (scope) => `[...document.querySelectorAll('${scope} .bs-body .bs-row')].map((r) => [...r.querySelectorAll('.bs-td')].map((c) => c.textContent.trim()))`;

export async function verifyBaseFilter({ open, check, until, wait }, { dir }) {
  const file = path.join(dir, 'Filters.rbase');
  try {
    for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
    const b = Base.open(file);
    buildFromSpec(b, {
      tables: [
        { name: 'Customers', fields: [{ name: 'Name', type: 'text' }, { name: 'City', type: 'text' }, { name: 'Joined', type: 'date' }] },
        { name: 'Orders', fields: [{ name: 'Customer', type: 'integer', points: 'Customers' }, { name: 'Amount', type: 'currency' }] },
      ],
      queries: { 'Order list': { tables: ['Orders', 'Customers'], columns: [{ table: 'Customers', field: 'Name' }, { table: 'Customers', field: 'City' }, { table: 'Orders', field: 'Amount' }] } },
    });
    for (const [n, c, j] of [['Amira Khan', 'Lahore', '2026-01-10'], ['Ben Ortiz', 'London', '2026-03-02'], ['Grace Okafor', 'Leeds', '2026-05-20'], ['Hiroshi Tanaka', 'London', '2026-08-15']]) b.insert('Customers', { Name: n, City: c, Joined: j });
    for (const [c, a] of [[1, 100], [1, 250], [2, 40], [3, 310], [4, 75]]) b.insert('Orders', { Customer: c, Amount: a });
    b.close();

    const win = await open('base', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('[data-table="Customers"]'))`), 'the database', 8000);
    await js(`document.querySelector('[data-table="Customers"]').click(), 1`);
    await until(() => js(`document.querySelectorAll('.bs-sheet .bs-body .bs-row').length === 4`), 'the customers', 6000).catch(() => {});

    // Words, in any column.
    await js(set('.bs-sheet .bs-search input', 'london'));
    await until(() => js(`document.querySelectorAll('.bs-sheet .bs-body .bs-row').length === 2`), 'the search', 5000).catch(() => {});
    const searched = { rows: await js(rows('.bs-sheet')), count: await js(`document.querySelector('.bs-sheet .bs-count')?.textContent || ''`) };
    await js(set('.bs-sheet .bs-search input', ''));

    // A date by its range.
    await js(`document.querySelector('.bs-sheet .bs-th-filter[data-filter-field="Joined"]').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.bs-filter-from'))`), 'the date filter', 4000);
    const dateBoxes = await js(`document.querySelector('.bs-filter-from').type`);
    await js(set('.bs-filter-from', '2026-03-01'));
    await js(set('.bs-filter-to', '2026-06-30'));
    await js(`document.querySelector('.bs-filter-apply').click(), 1`);
    await until(() => js(`document.querySelectorAll('.bs-sheet .bs-body .bs-row').length === 2`), 'the date range', 5000).catch(() => {});
    const snap = async (name) => { if (!process.env.RUTBA_VERIFY_CAPTURE) return; win.webContents.invalidate(); await wait(700); fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await win.webContents.capturePage()).toPNG()); };
    await snap('base-filter-table.png');
    const ranged = { rows: (await js(rows('.bs-sheet'))).map((r) => r[1]), chip: await js(`document.querySelector('.bs-sheet .bs-chip')?.textContent || ''`), lit: await js(`document.querySelector('.bs-th-filter[data-filter-field="Joined"]').classList.contains('on')`) };
    // Off by its cross.
    await js(`document.querySelector('.bs-sheet .bs-chip-x').click(), 1`);
    await until(() => js(`document.querySelectorAll('.bs-sheet .bs-body .bs-row').length === 4`), 'all again', 5000).catch(() => {});
    const offAgain = await js(`document.querySelectorAll('.bs-sheet .bs-body .bs-row').length`);

    // By selection: a city's cell, only this value.
    await js(`(() => { const c = [...document.querySelectorAll('.bs-sheet .bs-td[data-field="City"]')].find((n) => n.textContent.trim() === 'London'); c.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 200, clientY: 200 })); return 1; })()`);
    await until(() => js(`[...document.querySelectorAll('[role="menuitem"]')].some((n) => /Only this value/.test(n.textContent))`), 'the cell menu', 4000);
    await js(`[...document.querySelectorAll('[role="menuitem"]')].find((n) => /Only this value/.test(n.textContent)).click(), 1`);
    await until(() => js(`document.querySelectorAll('.bs-sheet .bs-body .bs-row').length === 2`), 'only London', 5000).catch(() => {});
    const selected = (await js(rows('.bs-sheet'))).map((r) => r[1]);
    check('base: a table\'s records are searched for words in any column, a date filtered by its range, a cell\'s value kept by selection, each filter a chip that takes it off, and the count says how many were found',
      JSON.stringify(searched.rows.map((r) => r[1])) === JSON.stringify(['Ben Ortiz', 'Hiroshi Tanaka']) && /2 records found/.test(searched.count)
        && dateBoxes === 'date' && JSON.stringify(ranged.rows) === JSON.stringify(['Ben Ortiz', 'Grace Okafor']) && /Joined from 2026-03-01 to 2026-06-30/.test(ranged.chip) && ranged.lit
        && offAgain === 4 && JSON.stringify(selected) === JSON.stringify(['Ben Ortiz', 'Hiroshi Tanaka']),
      JSON.stringify({ searched, dateBoxes, ranged, offAgain, selected }));

    // The saved query: run, then filtered the same way.
    await js(`document.querySelector('.bs-queries [data-query="Order list"]').click(), 1`);
    await until(() => js(`document.querySelectorAll('.bs-query-run .bs-body .bs-row').length === 5`), 'the query run', 6000).catch(() => {});
    await js(`document.querySelector('.bs-query-run .bs-th-filter[data-filter-field="Amount"]').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.bs-filter-from'))`), 'the amount filter', 4000);
    const numberBoxes = await js(`document.querySelector('.bs-filter-from').type`);
    await js(set('.bs-filter-from', '90'));
    await js(set('.bs-filter-to', '300'));
    await js(`document.querySelector('.bs-filter-apply').click(), 1`);
    await until(() => js(`document.querySelectorAll('.bs-query-run .bs-body .bs-row').length === 2`), 'the amount range', 5000).catch(() => {});
    const qRanged = (await js(rows('.bs-query-run'))).map((r) => r[2]);
    await js(set('.bs-query-run .bs-search input', 'lahore'));
    await until(() => js(`document.querySelectorAll('.bs-query-run .bs-body .bs-row').length === 2`), 'still two', 3000).catch(() => {});
    await js(set('.bs-query-run .bs-search input', 'amira 250'));
    await until(() => js(`document.querySelectorAll('.bs-query-run .bs-body .bs-row').length === 1`), 'one order', 5000).catch(() => {});
    await snap('base-filter-query.png');
    const qSearched = { rows: await js(rows('.bs-query-run')), count: await js(`document.querySelector('.bs-query-run .bs-count')?.textContent || ''`) };
    // Sorted by a column, as a table is.
    await js(`document.querySelector('.bs-query-run .bs-filter-clear').click(), 1`);
    await until(() => js(`document.querySelectorAll('.bs-query-run .bs-body .bs-row').length === 5`), 'all five', 5000).catch(() => {});
    await js(`document.querySelector('.bs-query-run .bs-th[data-field="Amount"]').click(), 1`);
    await wait(300);
    await js(`document.querySelector('.bs-query-run .bs-th[data-field="Amount"]').click(), 1`);
    await until(async () => (await js(rows('.bs-query-run')))[0]?.[2] === '310.00', 'largest first', 5000).catch(() => {});
    const sorted = (await js(rows('.bs-query-run'))).map((r) => r[2]);
    check('base: a saved query\'s results are filtered by a number range, searched for words, sorted by a column and counted as a table\'s records are',
      numberBoxes === 'number' && JSON.stringify(qRanged) === JSON.stringify(['100.00', '250.00']) && qSearched.rows.length === 1 && qSearched.rows[0][0] === 'Amira Khan' && /1 row found/.test(qSearched.count)
        && JSON.stringify(sorted) === JSON.stringify(['310.00', '250.00', '100.00', '75.00', '40.00']),
      JSON.stringify({ numberBoxes, qRanged, qSearched, sorted }));
    win.close();
  } catch (err) {
    check('base: the search and filter checks ran to the end', false, err?.message || String(err));
  }
}
