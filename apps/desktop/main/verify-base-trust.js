// A database from someone else asks before it reads its sources (A2.18),
// in the window: a database made elsewhere, with a table read from a CSV
// file at a path here, opens with the question naming the file's path;
// Not now leaves the table's copy as it was, not read, and its Refresh asks
// again; Allow for this database reads it; opened again, it does not ask
// and reads the file as it opens. Run alone with RUTBA_VERIFY_ONLY=base-trust.

import fs from 'node:fs';
import path from 'node:path';
import { Base, replaceRows } from '@rutba/database';

export async function verifyBaseTrust({ open, check, until, wait }, { dir }) {
  const file = path.join(dir, 'Sent to me.rbase');
  const csv = path.join(dir, 'Their prices.csv');
  try {
    for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
    fs.writeFileSync(csv, 'Item,Price\r\nFlour,1.80\r\nSugar,0.95\r\n');
    // As someone else made it: a table read from the CSV, its copy kept from before.
    const b = Base.open(file);
    b.createTable({ name: 'Notes', fields: [{ name: 'Text', type: 'text' }] });
    const link = b.saveLink({ name: 'Prices', engine: 'file', settings: { where: csv, header: 1, every: 15 } });
    b.linkTable({ table: 'Prices', link, remote: { sheet: null }, key: null, fields: [{ name: 'Item', type: 'text' }, { name: 'Price', type: 'currency' }], mode: 'copy' });
    replaceRows(b, 'Prices', [{ Item: 'Kept flour', Price: 1 }], b.fields('Prices'));
    b.close();

    let win = await open('base', file);
    let js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.bs-ask'))`), 'the question', 8000);
    const asked = await js(`({
      title: document.querySelector('.bs-ask .rw-dialog-head')?.textContent || '',
      where: [...document.querySelectorAll('.bs-ask .bs-ask-where')].map((n) => n.textContent),
      kind: document.querySelector('.bs-ask .bs-ask-kind')?.textContent || '',
      place: document.querySelector('.bs-ask .bs-ask-item')?.dataset.place || '',
      allow: document.querySelector('.bs-ask-allow')?.textContent || '',
      later: document.querySelector('.bs-ask-later')?.textContent || '',
    })`);
    if (process.env.RUTBA_VERIFY_CAPTURE) {
      win.webContents.invalidate();
      await wait(700);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'base-trust.png'), (await win.webContents.capturePage()).toPNG());
    }
    // Not now: the copy as it was, the table not read.
    await js(`document.querySelector('.bs-ask-later').click(), 1`);
    await until(() => js(`!document.querySelector('.bs-ask')`), 'the question put away', 4000);
    await js(`document.querySelector('[data-table="Prices"]').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('[data-linked="Prices"]'))`), 'the table\'s bar', 5000);
    await until(() => js(`document.querySelectorAll('.bs-row[data-rowid]').length === 1`), 'the copy kept', 5000).catch(() => {});
    const later = await js(`({
      blocked: document.querySelector('[data-linked="Prices"]')?.dataset.blocked || '',
      bar: document.querySelector('[data-linked="Prices"]')?.textContent || '',
      items: [...document.querySelectorAll('.bs-row[data-rowid] [data-field="Item"]')].map((n) => n.textContent),
    })`);
    // Refresh asks the same question; Allow reads the file.
    await js(`document.querySelector('.bs-source-refresh').click(), 1`);
    const again = await until(() => js(`Boolean(document.querySelector('.bs-ask'))`), 'the question again', 4000).then(() => true, () => false);
    if (again) await js(`document.querySelector('.bs-ask-allow').click(), 1`);
    await until(() => js(`document.querySelectorAll('.bs-row[data-rowid]').length === 2`), 'the file\'s records', 8000).catch(() => {});
    const allowed = await js(`({
      question: Boolean(document.querySelector('.bs-ask')),
      blocked: document.querySelector('[data-linked="Prices"]')?.dataset.blocked || '',
      items: [...document.querySelectorAll('.bs-row[data-rowid] [data-field="Item"]')].map((n) => n.textContent),
    })`);
    // Opened again (in a window of its own, the first closed after: the last window closing ends the run).
    fs.writeFileSync(csv, 'Item,Price\r\nFlour,1.80\r\nSugar,0.95\r\nSalt,0.40\r\n');
    const first = win;
    win = await open('base', file);
    first.close();
    js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('[data-table="Prices"]'))`), 'the database again', 8000);
    await wait(800);
    const askedAgain = await js(`Boolean(document.querySelector('.bs-ask'))`);
    if (askedAgain) await js(`document.querySelector('.bs-ask-later').click(), 1`);
    await js(`document.querySelector('[data-table="Prices"]').click(), 1`);
    await until(() => js(`document.querySelectorAll('.bs-row[data-rowid]').length === 3`), 'the file read as it opened', 8000).catch(() => {});
    const reopened = await js(`({
      blocked: document.querySelector('[data-linked="Prices"]')?.dataset.blocked || '',
      items: [...document.querySelectorAll('.bs-row[data-rowid] [data-field="Item"]')].map((n) => n.textContent),
    })`);
    check('base: a database made elsewhere with a table read from a CSV file asks first, naming the file\'s path; Not now keeps its copy unread and Refresh asks again; Allow for this database reads it; opened again it does not ask and reads the file as it opens',
      asked.title === 'Read from these places?' && asked.where.length === 1 && asked.where[0] === path.resolve(csv) && /^A file (in|outside) your own folders$/.test(asked.kind) && ['mine', 'outside'].includes(asked.place)
        && asked.allow === 'Allow for this database' && asked.later === 'Not now'
        && later.blocked === '1' && /Not read: Their prices\.csv waits/.test(later.bar) && JSON.stringify(later.items) === JSON.stringify(['Kept flour'])
        && again && !allowed.question && !allowed.blocked && JSON.stringify(allowed.items) === JSON.stringify(['Flour', 'Sugar'])
        && !askedAgain && !reopened.blocked && JSON.stringify(reopened.items) === JSON.stringify(['Flour', 'Sugar', 'Salt']),
      JSON.stringify({ asked, later, again, allowed, askedAgain, reopened }));
  } catch (err) {
    check('base: the sources question checks ran to the end', false, err?.message || String(err));
  }
}
