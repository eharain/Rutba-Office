// Tables linked to a database engine (D-2026-10-22-13), in the window: Link to
// a database engine connects to a PostgreSQL, lists its tables, and links one;
// its records show with the bar that says they are the engine's; a value
// changed in the datasheet is changed in the engine; and Read again brings
// back what someone else changed there. Needs a real PostgreSQL named by
// OFFICE_PG_TEST_URL; without one it is not run. Run alone with
// RUTBA_VERIFY_ONLY=base-engine-link.

import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { Base } from '@rutba/database';

const set = (sel, value) => `(() => { const el = document.querySelector(${JSON.stringify(sel)}); const proto = el instanceof HTMLSelectElement ? HTMLSelectElement.prototype : HTMLInputElement.prototype; Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, ${JSON.stringify(value)}); el.dispatchEvent(new Event(el instanceof HTMLSelectElement ? 'change' : 'input', { bubbles: true })); return 1; })()`;

export async function verifyBaseEngineLink({ open, check, until, wait }, { dir }) {
  const url = process.env.OFFICE_PG_TEST_URL;
  if (!url) return;
  const pg = createRequire(import.meta.url)('pg');
  const u = new URL(url);
  const client = new pg.Client({ connectionString: url });
  const file = path.join(dir, 'Linked.rbase');
  try {
    await client.connect();
    await client.query('DROP SCHEMA IF EXISTS datafiles_link_window CASCADE');
    await client.query('CREATE SCHEMA datafiles_link_window');
    await client.query(`CREATE TABLE datafiles_link_window.stock (id serial PRIMARY KEY, item varchar(60) NOT NULL, on_hand integer, price numeric(8,2))`);
    await client.query(`INSERT INTO datafiles_link_window.stock (item, on_hand, price) VALUES ('Flour', 12, 1.80), ('Sugar', 4, 0.95)`);
    for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
    const b = Base.open(file);
    b.createTable({ name: 'Notes', fields: [{ name: 'Text', type: 'text' }] });
    b.close();

    const win = await open('base', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.bs-link-engine'))`), 'Link to a database engine', 8000);
    await js(`document.querySelector('.bs-link-engine').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.bs-link-host-input'))`), 'the link box', 5000);
    await js(set('.bs-link-host-input', u.hostname));
    await js(set('.bs-link-port-input', String(u.port || 5432)));
    await js(set('.bs-link-database', u.pathname.slice(1)));
    await js(set('.bs-link-user', decodeURIComponent(u.username)));
    await js(set('.bs-link-password', decodeURIComponent(u.password)));
    await js(`document.querySelector('.bs-link-connect').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('[data-remote="datafiles_link_window.stock"]'))`), 'the engine\'s tables', 10000);
    await js(`document.querySelector('[data-remote="datafiles_link_window.stock"] input').click(), 1`);
    await js(set('.bs-link-name', 'Warehouse'));
    await js(`document.querySelector('.bs-link-go').click(), 1`);
    const name = 'datafiles_link_window stock';
    await until(() => js(`Boolean(document.querySelector('[data-linked="${name}"]'))`), 'the linked table', 10000);
    await until(() => js(`document.querySelectorAll('.bs-row[data-rowid]').length === 2`), 'its records', 8000);
    const linked = await js(`({
      mark: Boolean(document.querySelector('[data-linked-mark="${name}"]')),
      bar: document.querySelector('[data-linked="${name}"]')?.textContent || '',
      items: [...document.querySelectorAll('.bs-row[data-rowid] [data-field="item"]')].map((n) => n.textContent),
      designTab: Boolean(document.querySelector('.bs-tab[data-view="design"]')),
    })`);
    if (process.env.RUTBA_VERIFY_CAPTURE) {
      win.webContents.invalidate();
      await wait(700);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'base-engine-link.png'), (await win.webContents.capturePage()).toPNG());
    }
    // Changed in the datasheet: changed in the engine.
    await js(`document.querySelector('.bs-row[data-rowid="2"] [data-field="on_hand"]').dispatchEvent(new MouseEvent('dblclick', { bubbles: true })), 1`);
    await until(() => js(`Boolean(document.querySelector('.bs-edit'))`), 'the cell open', 5000);
    await js(set('.bs-edit', '40'));
    await js(`document.querySelector('.bs-edit').dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })), 1`);
    let there = null;
    for (let i = 0; i < 40; i++) {
      there = (await client.query('SELECT on_hand FROM datafiles_link_window.stock WHERE id = 2')).rows[0]?.on_hand;
      if (there === 40) break;
      await wait(150);
    }
    // Changed in the engine by someone else: Read again brings it.
    await client.query(`UPDATE datafiles_link_window.stock SET item = 'Brown sugar' WHERE id = 2`);
    await js(`document.querySelector('.bs-linked-again').click(), 1`);
    const again = await until(() => js(`[...document.querySelectorAll('.bs-row[data-rowid] [data-field="item"]')].some((n) => n.textContent === 'Brown sugar')`), 'read again', 8000).then(() => true, () => false);
    // Live: changed in the engine again, it shows on its own within the next read (every 15 seconds).
    await client.query(`UPDATE datafiles_link_window.stock SET on_hand = 7 WHERE id = 1`);
    const live = await until(() => js(`document.querySelector('.bs-row[data-rowid="1"] [data-field="on_hand"]')?.textContent === '7'`), 'the change shown on its own', 25000).then(() => true, () => false);
    const kept = (() => { const x = Base.open(file); try { return { links: x.links().map((l) => l.name), linked: x.linked().map((l) => l.table), password: fs.readFileSync(file).includes(Buffer.from(decodeURIComponent(u.password))) }; } finally { x.close(); } })();
    check('base: Link to a database engine connects to PostgreSQL, links a table whose records show with the bar that says they are the engine\'s; a value changed here is changed in the engine, Read again brings what changed there, and it shows on its own within the next live read; the file keeps the link, not the password',
      linked.mark && /Warehouse/.test(linked.bar) && JSON.stringify(linked.items) === JSON.stringify(['Flour', 'Sugar']) && !linked.designTab
        && there === 40 && again && live && JSON.stringify(kept.links) === JSON.stringify(['Warehouse']) && kept.linked.includes(name) && !kept.password,
      JSON.stringify({ linked, there, again, live, kept }));
  } catch (err) {
    check('base: the engine link checks ran to the end', false, err?.message || String(err));
  } finally {
    await client.query('DROP SCHEMA IF EXISTS datafiles_link_window CASCADE').catch(() => {});
    await client.end().catch(() => {});
  }
}
