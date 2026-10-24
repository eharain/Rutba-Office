// A table linked to a database engine and kept as a copy in the file
// (D-2026-10-22-13, step 4), in the window: Link to a database engine with
// Keep a copy in this file chosen; its records show with the bar that says
// when it caught up; with the engine out of reach a value is changed, and the
// bar says one change waits and since when; with the engine back, Sync now
// sends it and brings what changed there meanwhile. The engine is reached
// through a relay here, stopped and started to take it out of reach. Needs a
// real PostgreSQL named by OFFICE_PG_TEST_URL; without one it is not run.
// Run alone with RUTBA_VERIFY_ONLY=base-engine-copy.

import fs from 'node:fs';
import net from 'node:net';
import path from 'node:path';
import { createRequire } from 'node:module';
import { Base } from '@rutba/database';

const set = (sel, value) => `(() => { const el = document.querySelector(${JSON.stringify(sel)}); const proto = el instanceof HTMLSelectElement ? HTMLSelectElement.prototype : HTMLInputElement.prototype; Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, ${JSON.stringify(value)}); el.dispatchEvent(new Event(el instanceof HTMLSelectElement ? 'change' : 'input', { bubbles: true })); return 1; })()`;

/** A relay to the engine, stopped to take it out of reach and started again on the same port. */
function relay(host, port) {
  let server = null;
  const sockets = new Set();
  let at = 0;
  const start = () => new Promise((resolve, reject) => {
    server = net.createServer((inbound) => {
      const outbound = net.connect(port, host);
      for (const s of [inbound, outbound]) { sockets.add(s); s.on('close', () => sockets.delete(s)); s.on('error', () => {}); }
      inbound.pipe(outbound);
      outbound.pipe(inbound);
    });
    server.once('error', reject);
    server.listen(at, '127.0.0.1', () => { at = server.address().port; resolve(at); });
  });
  const stop = () => new Promise((resolve) => {
    for (const s of sockets) s.destroy();
    sockets.clear();
    if (!server) { resolve(); return; }
    server.close(() => resolve());
    server = null;
  });
  return { start, stop, port: () => at };
}

export async function verifyBaseEngineCopy({ open, check, until, wait }, { dir }) {
  const url = process.env.OFFICE_PG_TEST_URL;
  if (!url) return;
  const pg = createRequire(import.meta.url)('pg');
  const u = new URL(url);
  const client = new pg.Client({ connectionString: url });
  const file = path.join(dir, 'Copied.rbase');
  const r = relay(u.hostname, Number(u.port || 5432));
  try {
    await client.connect();
    await client.query('DROP SCHEMA IF EXISTS datafiles_copy_window CASCADE');
    await client.query('CREATE SCHEMA datafiles_copy_window');
    await client.query(`CREATE TABLE datafiles_copy_window.stock (id serial PRIMARY KEY, item varchar(60) NOT NULL, on_hand integer)`);
    await client.query(`INSERT INTO datafiles_copy_window.stock (item, on_hand) VALUES ('Flour', 12), ('Sugar', 4)`);
    for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
    const b = Base.open(file);
    b.createTable({ name: 'Notes', fields: [{ name: 'Text', type: 'text' }] });
    b.close();
    const port = await r.start();

    const win = await open('base', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.bs-link-engine'))`), 'Link to a database engine', 8000);
    await js(`document.querySelector('.bs-link-engine').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.bs-link-host-input'))`), 'the link box', 5000);
    await js(set('.bs-link-host-input', '127.0.0.1'));
    await js(set('.bs-link-port-input', String(port)));
    await js(set('.bs-link-database', u.pathname.slice(1)));
    await js(set('.bs-link-user', decodeURIComponent(u.username)));
    await js(set('.bs-link-password', decodeURIComponent(u.password)));
    await js(`document.querySelector('.bs-link-connect').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('[data-remote="datafiles_copy_window.stock"]'))`), 'the engine\'s tables', 10000);
    await js(`document.querySelector('[data-remote="datafiles_copy_window.stock"] input').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('[data-mode-for="datafiles_copy_window.stock"]'))`), 'the Live or copy choice', 4000);
    await js(set('[data-mode-for="datafiles_copy_window.stock"]', 'copy'));
    await js(set('.bs-link-name', 'Warehouse'));
    await js(`document.querySelector('.bs-link-go').click(), 1`);
    const name = 'datafiles_copy_window stock';
    await until(() => js(`document.querySelector('[data-linked="${name}"]')?.dataset.mode === 'copy'`), 'the copy\'s bar', 10000);
    await until(() => js(`document.querySelectorAll('.bs-row[data-rowid]').length === 2`), 'its records', 8000);
    const linked = await js(`({
      bar: document.querySelector('[data-linked="${name}"]')?.textContent || '',
      items: [...document.querySelectorAll('.bs-row[data-rowid] [data-field="item"]')].map((n) => n.textContent),
      sync: Boolean(document.querySelector('.bs-linked-sync')),
    })`);
    const real = (() => { const x = Base.open(file); try { return x.db.prepare(`SELECT COUNT(*) AS n FROM "${name}"`).get().n; } finally { x.close(); } })();

    // Out of reach: a value changed here waits, and the bar says so.
    await r.stop();
    await client.query(`UPDATE datafiles_copy_window.stock SET item = 'Rye flour' WHERE id = 1`);
    const row2 = await js(`document.querySelector('.bs-row [data-field="item"]') && [...document.querySelectorAll('.bs-row[data-rowid]')].find((r) => r.querySelector('[data-field="item"]')?.textContent === 'Sugar')?.dataset.rowid`);
    await js(`document.querySelector('.bs-row[data-rowid="${row2}"] [data-field="on_hand"]').dispatchEvent(new MouseEvent('dblclick', { bubbles: true })), 1`);
    await until(() => js(`Boolean(document.querySelector('.bs-edit'))`), 'the cell open', 5000);
    await js(set('.bs-edit', '40'));
    await js(`document.querySelector('.bs-edit').dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })), 1`);
    await until(() => js(`document.querySelector('[data-linked="${name}"]')?.dataset.waiting === '1'`), 'one change waiting', 8000).catch(() => {});
    await js(`document.querySelector('.bs-linked-sync').click(), 1`);
    await until(() => js(`document.querySelector('[data-linked="${name}"]')?.classList.contains('bad')`), 'the engine out of reach', 15000).catch(() => {});
    const away = await js(`({
      waiting: document.querySelector('[data-linked="${name}"]')?.dataset.waiting,
      bar: document.querySelector('[data-linked="${name}"]')?.textContent || '',
      shown: document.querySelector('.bs-row[data-rowid="${row2}"] [data-field="on_hand"]')?.textContent,
    })`);
    if (process.env.RUTBA_VERIFY_CAPTURE) {
      win.webContents.invalidate();
      await wait(700);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'base-engine-copy.png'), (await win.webContents.capturePage()).toPNG());
    }
    const before = (await client.query('SELECT on_hand FROM datafiles_copy_window.stock WHERE id = 2')).rows[0]?.on_hand;

    // Back: Sync now sends the change and brings what changed there meanwhile.
    await r.start();
    await js(`document.querySelector('.bs-linked-sync').click(), 1`);
    const caught = await until(() => js(`document.querySelector('[data-linked="${name}"]')?.dataset.waiting === '0' && !document.querySelector('[data-linked="${name}"]').classList.contains('bad') && [...document.querySelectorAll('.bs-row[data-rowid] [data-field="item"]')].some((n) => n.textContent === 'Rye flour')`), 'caught up', 15000).then(() => true, () => false);
    let there = null;
    for (let i = 0; i < 40; i++) {
      there = (await client.query('SELECT on_hand FROM datafiles_copy_window.stock WHERE id = 2')).rows[0]?.on_hand;
      if (there === 40) break;
      await wait(150);
    }
    const back = await js(`({ bar: document.querySelector('[data-linked="${name}"]')?.textContent || '', items: [...document.querySelectorAll('.bs-row[data-rowid] [data-field="item"]')].map((n) => n.textContent) })`);
    const kept = (() => { const x = Base.open(file); try { return { mode: x.linked().find((l) => l.table === name)?.mode, password: fs.readFileSync(file).includes(Buffer.from(decodeURIComponent(u.password))) }; } finally { x.close(); } })();
    check('base: a table linked with Keep a copy in this file shows its records with the bar that says when it caught up; with the engine out of reach a value changed here waits and the bar says one change waits and since when; with the engine back, Sync now sends it and brings what changed there meanwhile; the file keeps the records, not the password',
      /caught up at/.test(linked.bar) && JSON.stringify(linked.items) === JSON.stringify(['Flour', 'Sugar']) && linked.sync && real === 2
        && away.waiting === '1' && /1 change waits to go to Warehouse, kept since/.test(away.bar) && /cannot be reached/.test(away.bar) && away.shown === '40' && before === 4
        && caught && there === 40 && back.items.includes('Rye flour') && kept.mode === 'copy' && !kept.password,
      JSON.stringify({ linked, real, away, before, caught, there, back, kept }));
  } catch (err) {
    check('base: the engine copy checks ran to the end', false, err?.message || String(err));
  } finally {
    await r.stop().catch(() => {});
    await client.query('DROP SCHEMA IF EXISTS datafiles_copy_window CASCADE').catch(() => {});
    await client.end().catch(() => {});
  }
}
