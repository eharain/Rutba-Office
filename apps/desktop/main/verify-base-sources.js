// Tables read from other sources (D-2026-10-22-13, step 5), in the window:
// Link to a file or web address reads a CSV file, shows its fields and first
// records, and links it as a table read only here, with the bar that says
// where it is from and how often it is read; the file changed, Refresh brings
// its new records and says its columns changed. Then a web address answering
// JSON, reached with a token, its records over two pages: linked, its nested
// fields flattened, and the token kept out of the file. The web address is a
// stand-in served here. Run alone with RUTBA_VERIFY_ONLY=base-sources.

import fs from 'node:fs';
import http from 'node:http';
import path from 'node:path';
import { Base } from '@rutba/database';

const set = (sel, value) => `(() => { const el = document.querySelector(${JSON.stringify(sel)}); const proto = el instanceof HTMLSelectElement ? HTMLSelectElement.prototype : HTMLInputElement.prototype; Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, ${JSON.stringify(value)}); el.dispatchEvent(new Event(el instanceof HTMLSelectElement ? 'change' : 'input', { bubbles: true })); return 1; })()`;

const TOKEN = 'Bearer window-check-stand-in';
const ORDERS = [
  { id: 1, total: 12.5, customer: { name: 'Amina', town: 'Leeds' } },
  { id: 2, total: 8, customer: { name: 'Bilal', town: 'York' } },
  { id: 3, total: 30.25, customer: { name: 'Cara', town: 'Hull' } },
  { id: 4, total: 4.75, customer: { name: 'Dev', town: 'Bath' } },
  { id: 5, total: 19, customer: { name: 'Esme', town: 'Ely' } },
];

function standIn() {
  const server = http.createServer((req, res) => {
    const u = new URL(req.url, 'http://127.0.0.1');
    if (req.headers.authorization !== TOKEN) { res.writeHead(401, { 'content-type': 'application/json' }); res.end('{"error":"no"}'); return; }
    const page = Number(u.searchParams.get('page') || 1);
    const items = page === 1 ? ORDERS.slice(0, 3) : ORDERS.slice(3);
    const next = page === 1 ? `http://127.0.0.1:${server.address().port}/orders?page=2` : null;
    res.writeHead(200, { 'content-type': 'application/json' });
    res.end(JSON.stringify({ data: { items }, next }));
  });
  return new Promise((resolve) => server.listen(0, '127.0.0.1', () => resolve(server)));
}

export async function verifyBaseSources({ open, check, until, wait }, { dir }) {
  const file = path.join(dir, 'Sources.rbase');
  const csv = path.join(dir, 'Prices.csv');
  let server = null;
  try {
    for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
    fs.writeFileSync(csv, 'Item,Price\r\nFlour,1.80\r\nSugar,0.95\r\n');
    const b = Base.open(file);
    b.createTable({ name: 'Notes', fields: [{ name: 'Text', type: 'text' }] });
    b.close();
    server = await standIn();

    const win = await open('base', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.bs-link-source'))`), 'Link to a file or web address', 8000);
    await js(`document.querySelector('.bs-link-source').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.bs-source-where'))`), 'the source box', 5000);
    await js(set('.bs-source-where', csv));
    await js(`document.querySelector('.bs-source-try').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.bs-source-preview'))`), 'the fields and first records', 8000);
    const tried = await js(`({
      fields: [...document.querySelectorAll('.bs-source-preview th')].map((n) => n.textContent),
      count: document.querySelector('.bs-source-preview')?.dataset.count,
      name: document.querySelector('.bs-source-name')?.value,
      every: document.querySelector('.bs-source-every-new')?.value,
    })`);
    await js(`document.querySelector('.bs-source-go').click(), 1`);
    await until(() => js(`document.querySelector('[data-linked="Prices"]')?.dataset.source === 'file'`), 'the source\'s bar', 10000);
    await until(() => js(`document.querySelectorAll('.bs-row[data-rowid]').length === 2`), 'its records', 8000);
    const linked = await js(`({
      bar: document.querySelector('[data-linked="Prices"]')?.textContent || '',
      items: [...document.querySelectorAll('.bs-row[data-rowid] [data-field="Item"]')].map((n) => n.textContent),
      designTab: Boolean(document.querySelector('.bs-tab[data-view="design"]')),
      mark: Boolean(document.querySelector('[data-linked-mark="Prices"]')),
    })`);
    // Read only here: a cell does not open to type in.
    await js(`document.querySelector('.bs-row[data-rowid] [data-field="Price"]').dispatchEvent(new MouseEvent('dblclick', { bubbles: true })), 1`);
    await wait(400);
    const typed = await js(`Boolean(document.querySelector('.bs-edit'))`);
    if (typed) await js(`document.querySelector('.bs-edit').dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })), 1`);

    // The file changed: Refresh brings it, and says its columns changed.
    fs.writeFileSync(csv, 'Item,Price,Unit\r\nFlour,1.80,kg\r\nSugar,0.99,kg\r\nSalt,0.40,kg\r\n');
    await js(`document.querySelector('.bs-source-refresh').click(), 1`);
    const refreshed = await until(() => js(`document.querySelectorAll('.bs-row[data-rowid]').length === 3`), 'the new records', 8000).then(() => true, () => false);
    const after = await js(`({
      note: document.querySelector('[data-note="Prices"]')?.textContent || '',
      units: [...document.querySelectorAll('.bs-row[data-rowid] [data-field="Unit"]')].map((n) => n.textContent),
    })`);
    if (process.env.RUTBA_VERIFY_CAPTURE) {
      win.webContents.invalidate();
      await wait(700);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'base-sources.png'), (await win.webContents.capturePage()).toPNG());
    }
    check('base: Link to a file or web address reads a CSV file, shows its fields and first records, and links it as a table read only here with the bar that says where it is from and that it is read every 15 minutes; the file changed, Refresh brings its new records and says its columns changed',
      JSON.stringify(tried.fields) === JSON.stringify(['Item', 'Price']) && tried.count === '2' && tried.name === 'Prices' && tried.every === '15'
        && /From Prices\.csv/.test(linked.bar) && /every 15 minutes/.test(linked.bar) && /Read only here/.test(linked.bar) && JSON.stringify(linked.items) === JSON.stringify(['Flour', 'Sugar']) && !linked.designTab && linked.mark && !typed
        && refreshed && /added Unit/.test(after.note) && JSON.stringify(after.units) === JSON.stringify(['kg', 'kg', 'kg']),
      JSON.stringify({ tried, linked, typed, refreshed, after }));

    // A web address answering JSON, reached with a token, two pages of records.
    const where = `http://127.0.0.1:${server.address().port}/orders`;
    await js(`document.querySelector('.bs-link-source').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.bs-source-kind'))`), 'the source box', 5000);
    await js(set('.bs-source-kind', 'json'));
    await until(() => js(`Boolean(document.querySelector('.bs-source-records'))`), 'the web address settings', 4000);
    await js(set('.bs-source-where', where));
    await js(set('.bs-source-header-name', 'Authorization'));
    await js(set('.bs-source-header-value', TOKEN));
    await js(set('.bs-source-records', 'data.items'));
    await js(set('.bs-source-paging', 'next'));
    await until(() => js(`Boolean(document.querySelector('.bs-source-next'))`), 'where the next page is', 4000);
    await js(set('.bs-source-next', 'next'));
    await js(`document.querySelector('.bs-source-try').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.bs-source-preview'))`), 'the records answered', 8000);
    const json = await js(`({
      fields: [...document.querySelectorAll('.bs-source-preview th')].map((n) => n.textContent),
      count: document.querySelector('.bs-source-preview')?.dataset.count,
    })`);
    await js(set('.bs-source-name', 'Orders'));
    await js(`document.querySelector('.bs-source-go').click(), 1`);
    await until(() => js(`document.querySelector('[data-linked="Orders"]')?.dataset.source === 'json'`), 'the web source\'s bar', 10000);
    await until(() => js(`document.querySelectorAll('.bs-row[data-rowid]').length === 5`), 'its records', 8000).catch(() => {});
    const orders = await js(`({
      count: document.querySelectorAll('.bs-row[data-rowid]').length,
      bar: document.querySelector('[data-linked="Orders"]')?.textContent || '',
    })`);
    const kept = (() => {
      const x = Base.open(file);
      try {
        const token = Buffer.from('window-check-stand-in');
        const leaks = ['', '-wal'].some((s) => fs.existsSync(file + s) && fs.readFileSync(file + s).includes(token));
        return { sources: x.links().map((l) => `${l.engine}:${l.name}`), leaks };
      } finally { x.close(); }
    })();
    check('base: a web address answering JSON, reached with a token, is read over two pages with its nested fields flattened, linked as a table, and the token is kept in the protected store, not the file',
      json.count === '5' && json.fields.includes('customer.name') && json.fields.includes('total')
        && orders.count === 5 && /From 127\.0\.0\.1/.test(orders.bar) && kept.sources.includes('json:Orders') && kept.sources.includes('file:Prices') && !kept.leaks,
      JSON.stringify({ json, orders, kept }));
  } catch (err) {
    check('base: the source checks ran to the end', false, err?.message || String(err));
  } finally {
    if (server) await new Promise((resolve) => server.close(() => resolve()));
  }
}
