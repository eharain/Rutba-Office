// Pinning in Data Files' side panel (owner, 2026-10-22, D-2026-10-22-05: "the
// user should have way to create favourite list and pin either on top for
// easier access"): a query and a table pinned are listed first, above
// Tables, in the order pinned, and open from there; Unpin from an item's
// menu takes it off; a section folds away; both are kept in the file, so
// the database opens again as it was left. Run alone with
// RUTBA_VERIFY_ONLY=base-pin.

import fs from 'node:fs';
import path from 'node:path';
import { Base, buildFromSpec } from '@rutba/database';

export async function verifyBasePin({ open, check, until, wait }, { dir }) {
  const file = path.join(dir, 'Pinned.rbase');
  try {
    for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
    const b = Base.open(file);
    buildFromSpec(b, {
      tables: [{ name: 'Customers', fields: [{ name: 'Name', type: 'text' }] }, { name: 'Orders', fields: [{ name: 'Customer', type: 'integer', points: 'Customers' }, { name: 'Amount', type: 'currency' }] }],
      queries: { 'Big orders': { tables: ['Orders'], columns: [{ table: 'Orders', field: 'Amount', criteria: ['> 100'] }] }, 'All orders': { tables: ['Orders'], columns: [{ table: 'Orders', field: 'Amount' }] } },
    });
    b.close();

    let win = await open('base', file);
    let js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('[data-query="Big orders"]'))`), 'the database', 8000);
    const before = await js(`Boolean(document.querySelector('[data-section="pinned"]'))`);
    // Pinned: a query, then a table.
    await js(`document.querySelector('[data-pin="query:Big orders"]').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('[data-section="pinned"]'))`), 'the pinned list', 5000);
    await js(`document.querySelector('[data-pin="table:Orders"]').click(), 1`);
    await until(() => js(`document.querySelectorAll('.bs-pinned [data-pinned-item]').length === 2`), 'two pinned', 5000).catch(() => {});
    const listed = await js(`({ order: [...document.querySelectorAll('.bs-pinned [data-pinned-item]')].map((n) => n.dataset.pinnedItem), first: document.querySelector('.bs-side [data-section]')?.dataset.section, pressed: document.querySelector('[data-pin="query:Big orders"]').getAttribute('aria-pressed') })`);
    // Opened from there.
    await js(`document.querySelector('.bs-pinned [data-pinned-item="Big orders"]').click(), 1`);
    const opened = await until(() => js(`/Big orders/.test(document.querySelector('.bs-q-title')?.textContent || '')`), 'the query open', 5000).then(() => true, () => false);
    // Unpinned from the item's menu.
    await js(`document.querySelector('[data-table="Orders"]').dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 120, clientY: 200 })), 1`);
    await until(() => js(`[...document.querySelectorAll('[role="menuitem"]')].some((n) => /Unpin/.test(n.textContent))`), 'the menu', 4000);
    await js(`[...document.querySelectorAll('[role="menuitem"]')].find((n) => /Unpin/.test(n.textContent)).click(), 1`);
    await until(() => js(`document.querySelectorAll('.bs-pinned [data-pinned-item]').length === 1`), 'one pinned', 5000).catch(() => {});
    const left = await js(`[...document.querySelectorAll('.bs-pinned [data-pinned-item]')].map((n) => n.dataset.pinnedItem)`);
    // Forms folded away.
    await js(`document.querySelector('[data-section="forms"] .bs-fold').click(), 1`);
    await until(() => js(`document.querySelector('[data-section="forms"]').classList.contains('folded')`), 'forms folded', 5000).catch(() => {});
    const foldedNow = await js(`({ folded: document.querySelector('[data-section="forms"]').classList.contains('folded'), items: document.querySelectorAll('[data-section="forms"] [data-form]').length })`);
    // Opened again (in a window of its own, the first closed after: the last window closing ends the run): as it was left.
    const first = win;
    win = await open('base', file);
    first.close();
    js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('[data-query="Big orders"]'))`), 'the database again', 8000);
    const again = await js(`({ pinned: [...document.querySelectorAll('.bs-pinned [data-pinned-item]')].map((n) => n.dataset.pinnedItem), folded: document.querySelector('[data-section="forms"]')?.classList.contains('folded') })`);
    check('base: an item pinned is listed first, above Tables, in the order pinned, and opens from there; Unpin takes it off; a section folds away; both kept in the file',
      !before && JSON.stringify(listed.order) === JSON.stringify(['Big orders', 'Orders']) && listed.first === 'pinned' && listed.pressed === 'true' && opened
        && JSON.stringify(left) === JSON.stringify(['Big orders']) && foldedNow.folded && foldedNow.items === 0
        && JSON.stringify(again.pinned) === JSON.stringify(['Big orders']) && again.folded === true,
      JSON.stringify({ before, listed, opened, left, foldedNow, again }));
    if (process.env.RUTBA_VERIFY_CAPTURE) {
      win.webContents.invalidate();
      await wait(700);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'base-pin.png'), (await win.webContents.capturePage()).toPNG());
    }
  } catch (err) {
    check('base: the pinning checks ran to the end', false, err?.message || String(err));
  }
}
