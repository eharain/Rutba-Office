// An app's entry page (the owner, 2026-10-21: "the entry page of
// applications should show recent files opened by the app beside current
// new and starter templates", and "user should be able to remove the already
// added ones from recent"), checked in the window: Documents opened as the
// launcher opens it shows Blank and its starter templates beside the files
// recently opened in Documents (and none of Worksheets'); the x takes a file
// off the list, and it stays off; Blank makes a new document in the window;
// and New, as every check and shortcut asks for it, still opens a document
// straight away.

import fs from 'node:fs';
import path from 'node:path';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';

export async function verifyStartPage(h, { dir }) {
  const { open, check, until, wait } = h;
  try {
    const a = path.join(dir, 'Start page letter.docx');
    const b = path.join(dir, 'Start page report.docx');
    const c = path.join(dir, 'Start page budget.xlsx');
    fs.writeFileSync(a, buildDocx({ paragraphs: ['A letter'] }));
    fs.writeFileSync(b, buildDocx({ paragraphs: ['A report'] }));
    fs.writeFileSync(c, buildXlsx({ sheets: [{ name: 'Sheet1', rows: [['a', 1]] }] }));

    const win = await open('word', null, { start: '1' });
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.sp[data-start="word"]'))`), 'the entry page', 8000);
    for (const [p, app] of [[c, 'sheets'], [b, 'word'], [a, 'word']]) await js(`window.rutbaOffice.app.addRecent(${JSON.stringify({ path: p, app })})`);
    // The page reads the list as it opens: open it again now the files are on it.
    await js(`location.reload(), 1`);
    await until(() => js(`document.querySelectorAll('.sp-row[data-path]').length >= 2 && document.querySelectorAll('.sp-card[data-template]').length >= 3`), 'the recent files and starters', 10000).catch(() => {});
    if (process.env.RUTBA_VERIFY_CAPTURE) {
      await wait(500);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'word-start.png'), (await win.webContents.capturePage()).toPNG());
    }
    const page = await js(`({
      blank: Boolean(document.querySelector('.sp-card[data-template="blank"]')),
      starters: document.querySelectorAll('.sp-card[data-template]:not([data-template="blank"])').length,
      more: Boolean(document.querySelector('.sp-more')),
      recent: [...document.querySelectorAll('.sp-row[data-path]')].map((r) => r.dataset.path.split(/[\\\\/]/).pop()),
      doc: Boolean(document.querySelector('.wd-page')),
    })`);
    check('word: opened as the app it shows its entry page: Blank, its starter templates and More templates beside the files recently opened in Documents, and no Worksheets file among them',
      page.blank && page.starters >= 3 && page.more && !page.doc
        && page.recent.includes('Start page letter.docx') && page.recent.includes('Start page report.docx') && !page.recent.includes('Start page budget.xlsx'),
      JSON.stringify(page));

    // The x takes one off the list, and it stays off.
    await js(`document.querySelector('.sp-row[data-path$="Start page report.docx"] .sp-remove')?.click(), 1`);
    const gone = await until(() => js(`!document.querySelector('.sp-row[data-path$="Start page report.docx"]')`), 'the file off the list', 5000).then(() => true, () => false);
    const kept = await js(`window.rutbaOffice.app.recent().then((l) => ({ report: l.some((r) => r.path.endsWith('Start page report.docx')), letter: l.some((r) => r.path.endsWith('Start page letter.docx')) }))`);
    check('word: the x on a recent file takes it off the list, the list kept without it and the other files left', gone && !kept.report && kept.letter, JSON.stringify({ gone, kept }));

    // Blank: a new document in this window.
    await js(`document.querySelector('.sp-card[data-template="blank"]')?.click(), 1`);
    const made = await until(() => js(`Boolean(document.querySelector('.wd-page')) && !document.querySelector('.sp')`), 'the new document', 10000).then(() => true, () => false);
    check('word: Blank on the entry page makes a new document in the same window', made, String(made));

    // Data Files: its start page lists its recent databases, and one opens in the same window.
    const db = path.join(dir, 'Start page club.rbase');
    const { Base } = await import('@rutba/database');
    for (const s of ['', '-wal', '-shm']) fs.rmSync(db + s, { force: true });
    const club = Base.open(db);
    club.createTable({ name: 'Members', fields: [{ name: 'Name', type: 'text' }] });
    club.close();
    await js(`window.rutbaOffice.app.addRecent(${JSON.stringify({ path: db, app: 'base' })})`);
    const base = await open('base');
    const bjs = (code) => base.webContents.executeJavaScript(code);
    await until(() => bjs(`Boolean(document.querySelector('.sp-recent[data-recent="base"] .sp-row[data-path$="Start page club.rbase"]'))`), 'the recent database', 8000).catch(() => {});
    const listed = await bjs(`[...document.querySelectorAll('.sp-recent[data-recent="base"] .sp-row')].map((r) => r.dataset.path.split(/[\\\\/]/).pop())`);
    await bjs(`document.querySelector('.sp-row[data-path$="Start page club.rbase"] .sp-row-main')?.click(), 1`);
    const inPlace = await until(() => bjs(`Boolean(document.querySelector('.bs-app [data-table="Members"]'))`), 'the database open', 8000).then(() => true, () => false);
    check('base: the start page lists the databases recently opened, and one opens in the same window', listed.includes('Start page club.rbase') && inPlace, JSON.stringify({ listed, inPlace }));

    // New, as checks and Ctrl+N ask: a document straight away.
    const plain = await open('word');
    const direct = await until(() => plain.webContents.executeJavaScript(`Boolean(document.querySelector('.wd-page')) && !document.querySelector('.sp')`), 'a document', 10000).then(() => true, () => false);
    check('word: New still opens a document straight away, with no entry page', direct, String(direct));
  } catch (err) {
    check('word: the entry page checks ran', false, err.stack || err.message);
  }
}
