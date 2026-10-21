// Data Files' ready-made databases, checked in the window: the start screen
// offers each one; choosing Lending library and a place for it makes the
// file and opens it with its tables, saved query, a form for each table and
// its report; and a loan put in the file shows in the report under
// its borrower's full name, with the book's title.

import fs from 'node:fs';
import path from 'node:path';
import { Base } from '@rutba/database';

export async function verifyBaseTemplates(h, { dir }) {
  const { open, check, until, wait } = h;
  const file = path.join(dir, 'Library.rbase');
  try {
    for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
    const win = await open('base');
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.bs-template').length === 5`), 'the ready-made databases', 6000).catch(() => {});
    const offered = await js(`[...document.querySelectorAll('.bs-template')].map((b) => b.dataset.template)`);
    if (process.env.RUTBA_VERIFY_CAPTURE) {
      await wait(400);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'base-templates.png'), (await win.webContents.capturePage()).toPNG());
    }
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: file });
    await js(`document.querySelector('.bs-template[data-template="library"]')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.bs-reports [data-report="Loans by borrower"]'))`), 'the library', 8000).catch(() => {});
    const made = await js(`({
      tables: [...document.querySelectorAll('.bs-tables:not(.bs-queries):not(.bs-forms):not(.bs-reports) [data-table]')].map((b) => b.dataset.table),
      queries: [...document.querySelectorAll('.bs-queries [data-query]')].map((b) => b.dataset.query),
      forms: [...document.querySelectorAll('.bs-forms [data-form]')].map((b) => b.dataset.form),
      reports: [...document.querySelectorAll('.bs-reports [data-report]')].map((b) => b.dataset.report),
    })`);
    const onDisk = (() => { const d = Base.open(file); try { return { relationships: d.relationships().length, tables: d.tables().length }; } finally { d.close(); } })();
    check('base: the start screen offers five ready-made databases, and Lending library makes its tables, query, forms and report in a file of its own',
      JSON.stringify(offered) === JSON.stringify(['contacts', 'inventory', 'membership', 'crm', 'library'])
        && JSON.stringify(made.tables) === JSON.stringify(['Books', 'Borrowers', 'Loans'])
        && JSON.stringify(made.queries) === JSON.stringify(['On loan'])
        && made.forms.length === 3 && JSON.stringify(made.reports) === JSON.stringify(['Loans by borrower'])
        && onDisk.relationships === 2,
      JSON.stringify({ offered, made, onDisk }));

    // A book, a borrower and a loan, put in from outside; then the report.
    const d = Base.open(file);
    d.insert('Books', { Title: 'Middlemarch', Author: 'George Eliot' });
    d.insert('Borrowers', { 'First Name': 'Amira', 'Last Name': 'Haddad' });
    d.insert('Loans', { Book: 1, Borrower: 1, 'Lent On': '2026-10-01', 'Due Back': '2026-10-22' });
    d.close();
    await js(`document.querySelector('.bs-reports [data-report="Loans by borrower"]')?.click(), 1`);
    await until(() => js(`/Middlemarch/.test(document.querySelector('.bs-report-scroll .bs-report-page')?.textContent || '')`), 'the loan in the report', 6000).catch(() => {});
    const page = await js(`({ head: document.querySelector('.bs-report-scroll .bs-report-group-head')?.textContent || '', row: document.querySelector('.bs-report-scroll .bs-report-group tr:not(.bs-report-group-head)')?.textContent || '' })`);
    check("base: a loan shows in the template's report under the borrower's first and last name, with the book's title",
      /Amira Haddad/.test(page.head) && /Middlemarch/.test(page.row), JSON.stringify(page));
  } catch (err) {
    check('base: the ready-made database checks ran', false, err.stack || err.message);
  }
}
