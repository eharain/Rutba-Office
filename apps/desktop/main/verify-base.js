// Databases (2.9), checked in its window: a new database from the start
// screen, a table made and given fields in its design, records added and
// changed in the datasheet, a CSV imported as a table of its own, two tables
// related, a query run (and a change refused), and the file opened again from
// disk with all of it there. The Save and Open dialogs are answered by the
// check (ipc.js's check-run queue). Run alone with RUTBA_VERIFY_ONLY=base.

import fs from 'node:fs';
import path from 'node:path';
import { Base } from '@rutba/database';

const answer = (kind, value) => globalThis.__rutbaCheckDialogAnswers?.push({ kind, answer: value });

/** In-page helpers: set a React field's value as typing would, and press keys on it. */
const HELPERS = `
  window.__bs = window.__bs || {
    set(sel, value) {
      const el = typeof sel === 'string' ? document.querySelector(sel) : sel;
      if (!el) return false;
      const proto = el.tagName === 'SELECT' ? HTMLSelectElement.prototype : el.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
      Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, value);
      el.dispatchEvent(new Event(el.tagName === 'SELECT' ? 'change' : 'input', { bubbles: true }));
      return true;
    },
    key(sel, key, extra = {}) {
      const el = typeof sel === 'string' ? document.querySelector(sel) : sel;
      if (!el) return false;
      el.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, ...extra }));
      return true;
    },
    click(sel) { const el = document.querySelector(sel); if (!el) return false; el.click(); return true; },
    text(sel) { return document.querySelector(sel)?.textContent || ''; },
  };
  true;
`;

export async function verifyBase(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const ok = (p) => p.then(() => true, () => false);
  const file = path.join(dir, 'Shop.rbase');
  const csv = path.join(dir, 'people.csv');
  fs.writeFileSync(csv, 'Name,Age,Joined,Member\nAmira,34,2024-03-01,yes\nBen,41,2023-11-20,no\nCleo,29,2025-01-15,yes\n');

  // 1. A new database from the start screen.
  let win = await open('base');
  const js = (code) => win.webContents.executeJavaScript(code);
  await js(HELPERS);
  const start = await ok(until(() => js(`Boolean(document.querySelector('.bs-new-db'))`), 'the start screen', 6000));
  answer('save', file);
  await js(`__bs.click('.bs-new-db')`);
  const made = await ok(until(() => js(`Boolean(document.querySelector('.bs-app'))`), 'the new database', 6000));
  check('base: the start screen offers New Database, and New makes the file and opens it', start && made && fs.existsSync(file), JSON.stringify({ start, made, file: fs.existsSync(file) }));

  // 2. A table, and its fields in its design.
  await js(`__bs.click('.bs-new-table')`);
  await until(() => js(`Boolean(document.querySelector('.bs-table-name'))`), 'the New Table dialog', 4000).catch(() => {});
  await js(`__bs.set('.bs-table-name', 'Customers')`);
  await wait(100);
  await js(`__bs.click('.bs-make-table')`);
  const designed = await ok(until(() => js(`Boolean(document.querySelector('.bs-fields tr[data-field="ID"]'))`), 'the design with its ID', 5000));
  for (const [name, type] of [['Name', 'text'], ['Joined', 'date'], ['Member', 'yesno']]) {
    await js(`__bs.set('.bs-add-name', ${JSON.stringify(name)}); __bs.set('.bs-add-type', ${JSON.stringify(type)})`);
    await wait(80);
    await js(`__bs.click('.bs-add-field')`);
    await until(() => js(`Boolean(document.querySelector('.bs-fields tr[data-field=${JSON.stringify(name)}]'))`), `the ${name} field`, 4000).catch(() => {});
  }
  const fields = await js(`[...document.querySelectorAll('.bs-fields tr[data-field]')].map((r) => r.dataset.field)`);
  check('base: a new table starts with an ID, and Add Field gives it text, date and yes/no fields',
    designed && JSON.stringify(fields) === JSON.stringify(['ID', 'Name', 'Joined', 'Member']), JSON.stringify(fields));

  // 3. Records: two added from the new-record row, one changed in place.
  await js(`__bs.click('.bs-tab[data-view="data"]')`);
  await until(() => js(`Boolean(document.querySelector('.bs-new-input[data-field="Name"]'))`), 'the datasheet', 4000).catch(() => {});
  for (const [name, joined] of [['Dana', '2025-06-01'], ['Eli', '2025-07-04']]) {
    await js(`__bs.set('.bs-new-input[data-field="Name"]', ${JSON.stringify(name)}); __bs.set('.bs-new-input[data-field="Joined"]', ${JSON.stringify(joined)})`);
    await wait(80);
    await js(`__bs.key('.bs-new-input[data-field="Name"]', 'Enter')`);
    await until(() => js(`[...document.querySelectorAll('.bs-body .bs-td[data-field="Name"]')].some((c) => c.textContent === ${JSON.stringify(name)})`), `${name} in the sheet`, 4000).catch(() => {});
  }
  const counted = await js(`__bs.text('.bs-count')`);
  await js(`document.querySelector('.bs-body .bs-row[data-row="0"] .bs-td[data-field="Name"]').dispatchEvent(new MouseEvent('dblclick', { bubbles: true }))`);
  await until(() => js(`Boolean(document.querySelector('.bs-edit'))`), 'the cell editor', 3000).catch(() => {});
  await js(`__bs.set('.bs-edit', 'Dana Ortiz'); __bs.key('.bs-edit', 'Enter')`);
  const changed = await ok(until(() => js(`[...document.querySelectorAll('.bs-body .bs-td[data-field="Name"]')].some((c) => c.textContent === 'Dana Ortiz')`), 'the changed name', 4000));
  check('base: records added from the new-record row are counted, and a double-clicked cell is changed in place',
    /2 records/.test(counted) && changed, JSON.stringify({ counted, changed }));

  // 4. A CSV imported as a table of its own, its types read from its values.
  answer('open', [csv]);
  await js(`__bs.click('.bs-import')`);
  const imported = await ok(until(() => js(`/3/.test(document.querySelector('.bs-table[data-table="people"] small')?.textContent || '')`), 'the people table', 6000));
  await until(() => js(`Boolean(document.querySelector('.bs-tab[data-view="design"]'))`), 'the people sheet', 3000).catch(() => {});
  await js(`__bs.click('.bs-tab[data-view="design"]')`);
  await until(() => js(`Boolean(document.querySelector('.bs-fields tr[data-field="Age"]'))`), 'the people design', 4000).catch(() => {});
  const types = await js(`Object.fromEntries([...document.querySelectorAll('.bs-fields tr[data-field]')].map((r) => [r.dataset.field, r.querySelector('select')?.value || 'autonumber']))`);
  check('base: Import makes a table of a CSV, with whole numbers, dates and yes/no read from its values',
    imported && types.Age === 'integer' && types.Joined === 'date' && types.Member === 'yesno', JSON.stringify({ imported, types }));

  // 5. A relationship: each person points at a customer.
  await js(`__bs.click('.bs-tab[data-view="data"]')`);
  await js(`__bs.click('.bs-tab[data-view="design"]')`);
  await js(`__bs.set('.bs-add-name', 'Customer'); __bs.set('.bs-add-type', 'integer')`);
  await wait(80);
  await js(`__bs.click('.bs-add-field')`);
  await until(() => js(`Boolean(document.querySelector('.bs-fields tr[data-field="Customer"]'))`), 'the Customer field', 4000).catch(() => {});
  await js(`__bs.click('.bs-side-foot [data-view="relations"]')`);
  await until(() => js(`Boolean(document.querySelector('.bs-rel-from-table'))`), 'the relationships', 4000).catch(() => {});
  await until(() => js(`document.querySelectorAll('.bs-rel-from-table option').length > 2`), 'the tables to relate', 3000).catch(() => {});
  await js(`__bs.set('.bs-rel-from-table', 'people')`);
  await wait(80);
  await until(() => js(`[...document.querySelectorAll('.bs-rel-from-field option')].some((o) => o.value === 'Customer')`), 'the people fields', 3000).catch(() => {});
  await js(`__bs.set('.bs-rel-from-field', 'Customer'); __bs.set('.bs-rel-to-table', 'Customers')`);
  await wait(120);
  await js(`__bs.click('.bs-relate')`);
  const related = await ok(until(() => js(`/people.*Customer.*Customers.*ID/s.test(document.querySelector('.bs-rel-list')?.textContent || '')`), 'the relationship listed', 4000));
  check('base: two tables are related from Relationships, and the relationship is listed', related, await js(`__bs.text('.bs-rel-list').slice(0, 160)`));

  // 6. SQL: a query answers, and one that would change the database does not run.
  await js(`__bs.click('.bs-side-foot [data-view="sql"]')`);
  await until(() => js(`Boolean(document.querySelector('.bs-sql-text'))`), 'the SQL view', 4000).catch(() => {});
  await js(`__bs.set('.bs-sql-text', 'SELECT COUNT(*) AS n FROM "people" WHERE "Member" = 1')`);
  await wait(80);
  await js(`__bs.click('.bs-sql-run')`);
  const answered = await ok(until(() => js(`document.querySelector('.bs-result td')?.textContent === '2'`), 'the query result', 4000));
  await js(`__bs.set('.bs-sql-text', 'DELETE FROM "people"')`);
  await wait(80);
  await js(`__bs.click('.bs-sql-run')`);
  const refused = await ok(until(() => js(`Boolean(document.querySelector('.bs-error'))`), 'the refusal', 4000));
  const complaints = await errorsIn(win);
  check('base: a query answers in a grid, and a statement that would change the database is refused',
    answered && refused && !complaints.length, JSON.stringify({ answered, refused, error: await js(`__bs.text('.bs-error')`), complaints }));

  // 7. The file as it is on disk, read on its own, and opened again in a window.
  const disk = Base.open(file);
  let onDisk;
  try {
    onDisk = { tables: disk.tables().map((x) => x.name), people: disk.count('people'), names: disk.rows('Customers', { order: 'ID' }).map((r) => r.Name), relations: disk.relationships().length };
  } finally {
    disk.close();
  }
  win = await open('base', file);
  await js(HELPERS);
  const reopened = await ok(until(() => js(`Boolean(document.querySelector('.bs-table[data-table="Customers"]')) && Boolean(document.querySelector('.bs-table[data-table="people"]'))`), 'the tables on opening', 6000));
  check('base: the file keeps its tables, records and relationship, and opens again from disk',
    reopened && onDisk.people === 3 && JSON.stringify(onDisk.names) === JSON.stringify(['Dana Ortiz', 'Eli']) && onDisk.relations === 1, JSON.stringify({ reopened, onDisk }));
}
