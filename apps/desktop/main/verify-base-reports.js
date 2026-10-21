// Data Files' reports, checked in the window: New Report makes a report of
// Orders grouped by customer, each customer shown by name with the total of
// their orders and the whole's below; Group by and Sort by change it and are
// kept in the file; Save as PDF writes the page; and a report of a saved
// query lists what the query answers.

import fs from 'node:fs';
import path from 'node:path';
import { Base } from '@rutba/database';

const HELPERS = `
  window.__r = window.__r || {
    set(el, value) {
      if (typeof el === 'string') el = document.querySelector(el);
      if (!el) return false;
      const proto = el.tagName === 'SELECT' ? HTMLSelectElement.prototype : HTMLInputElement.prototype;
      Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, value);
      el.dispatchEvent(new Event(el.tagName === 'SELECT' ? 'change' : 'input', { bubbles: true }));
      return true;
    },
    click(sel) { const el = document.querySelector(sel); if (!el) return false; el.click(); return true; },
    groups() {
      return [...document.querySelectorAll('.bs-report-scroll .bs-report-group')].map((g) => ({
        head: g.querySelector('.bs-report-group-head')?.textContent || null,
        rows: [...g.querySelectorAll('tr:not(.bs-report-group-head):not(.bs-report-subtotal)')].map((r) => r.cells[0].textContent),
        total: g.querySelector('.bs-report-subtotal')?.textContent || null,
      }));
    },
    grand() { return document.querySelector('.bs-report-scroll .bs-report-total')?.textContent || ''; },
  };
  true;
`;

export async function verifyBaseReports(h, { dir }) {
  const { open, check, until, wait } = h;
  const file = path.join(dir, 'Shop3.rbase');
  const pdf = path.join(dir, 'Orders report.pdf');
  try {
    for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
    fs.rmSync(pdf, { force: true });
    const b = Base.open(file);
    b.createTable({ name: 'Customers', fields: [{ name: 'Name', type: 'text' }] });
    b.createTable({ name: 'Orders', fields: [{ name: 'Customer', type: 'integer' }, { name: 'Item', type: 'text' }, { name: 'Amount', type: 'currency' }] });
    b.addRelationship({ from: { table: 'Orders', field: 'Customer' }, to: { table: 'Customers', field: 'ID' } });
    b.insert('Customers', { Name: 'Ben' });
    b.insert('Customers', { Name: 'Amira' });
    for (const [c, i, a] of [[1, 'pen', 2], [2, 'ink', 5], [2, 'pad', 3], [1, 'cap', 4]]) b.insert('Orders', { Customer: c, Item: i, Amount: a });
    b.saveQuery('Dear', { tables: ['Orders'], columns: [{ table: 'Orders', field: 'Item', sort: 'asc' }, { table: 'Orders', field: 'Amount', criteria: ['>= 4'] }] });
    b.close();

    const win = await open('base', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await js(HELPERS);
    await until(() => js(`Boolean(document.querySelector('.bs-new-report'))`), 'the database', 8000);

    await js(`__r.click('.bs-new-report')`);
    await until(() => js(`Boolean(document.querySelector('.bs-report-source'))`), 'the New Report box', 4000).catch(() => {});
    await js(`__r.set('.bs-report-source', 'table:Orders')`);
    await wait(100);
    await js(`__r.click('.bs-make-report')`);
    await until(() => js(`document.querySelectorAll('.bs-report-scroll .bs-report-group').length === 2`), 'the report', 6000).catch(() => {});
    if (process.env.RUTBA_VERIFY_CAPTURE) {
      await wait(400);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'base-report.png'), (await win.webContents.capturePage()).toPNG());
    }
    const grouped = await js(`__r.groups()`);
    const grand = await js(`__r.grand()`);
    const listed = await js(`Boolean(document.querySelector('.bs-reports [data-report="Orders"]'))`);
    check("base: New Report makes a report of Orders grouped by customer, each shown by name, with each customer's total and the whole's",
      listed && grouped.length === 2 && /Amira/.test(grouped[0].head) && /Ben/.test(grouped[1].head)
        && JSON.stringify(grouped.map((g) => g.rows)) === JSON.stringify([['ink', 'pad'], ['pen', 'cap']])
        && /8\.00/.test(grouped[0].total) && /6\.00/.test(grouped[1].total) && /14\.00/.test(grand),
      JSON.stringify({ listed, grouped, grand }));

    // No grouping, the dearest first: kept in the file.
    await js(`__r.set('.bs-report-group-by', '')`);
    await until(() => js(`document.querySelectorAll('.bs-report-scroll .bs-report-group').length === 1 && !document.querySelector('.bs-report-scroll .bs-report-group-head')`), 'no grouping', 4000).catch(() => {});
    await js(`__r.set('.bs-report-sort', 'Amount')`);
    await until(() => js(`!document.querySelector('.bs-report-desc').disabled`), 'the sort', 4000).catch(() => {});
    await js(`__r.click('.bs-report-desc')`);
    await until(() => js(`__r.groups()[0]?.rows[0] === 'Amira'`), 'largest first', 4000).catch(() => {});
    const flat = await js(`__r.groups()`);
    const kept = (() => { const d = Base.open(file); try { return d.reports()[0]?.layout; } finally { d.close(); } })();
    check('base: Group by (no grouping) and Sort by, largest first, change the report and are kept in the file, the field grouped before back among the columns',
      flat.length === 1 && JSON.stringify(flat[0].rows) === JSON.stringify(['Amira', 'Ben', 'Amira', 'Ben'])
        && kept?.group === null && kept?.sort === 'Amount' && kept?.desc === true && kept?.columns[0]?.field === 'Customer',
      JSON.stringify({ flat, kept }));

    // Save as PDF: the page, written where the person chose.
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: pdf });
    await js(`__r.click('.bs-report-pdf')`);
    const wrote = await until(() => (fs.existsSync(pdf) && fs.statSync(pdf).size > 1000) || null, 'the PDF', 10000).then(() => true, () => false);
    if (wrote && process.env.RUTBA_VERIFY_CAPTURE) fs.copyFileSync(pdf, path.join(process.env.RUTBA_VERIFY_CAPTURE, 'base-report.pdf'));
    const head = wrote ? fs.readFileSync(pdf).subarray(0, 5).toString('latin1') : '';
    const leftOver = await js(`Boolean(document.querySelector('.bs-report-print-root'))`);
    check('base: Save as PDF writes the report as a PDF, the page alone gone again from the window afterwards', wrote && head === '%PDF-' && !leftOver, JSON.stringify({ wrote, head, leftOver }));

    // A report of a saved query.
    await js(`__r.click('.bs-new-report')`);
    await until(() => js(`Boolean(document.querySelector('.bs-report-source'))`), 'the New Report box', 4000).catch(() => {});
    await js(`__r.set('.bs-report-source', 'query:Dear')`);
    await wait(100);
    await js(`__r.click('.bs-make-report')`);
    await until(() => js(`Boolean(document.querySelector('.bs-report-page[data-report="Dear"]'))`), 'the query report', 6000).catch(() => {});
    const dear = await js(`__r.groups()`);
    check('base: a report of a saved query lists what the query answers', dear.length === 1 && JSON.stringify(dear[0].rows) === JSON.stringify(['cap', 'ink']), JSON.stringify(dear));
  } catch (err) {
    check('base: the report checks ran', false, err.stack || err.message);
  }
}
