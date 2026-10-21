// Word: a mail merge from a Data Files database. The letter is given a .rbase
// the way the OS hands it a file; Select Table lists the database's tables
// and saved queries, and the one chosen is the recipient list, its amounts
// with their two places. A merge field from its fields, then Preview Results,
// shows the first member's name on the page. The database is not changed.
// Run alone with RUTBA_VERIFY_ONLY=mailmerge-db.
import fs from 'node:fs';
import path from 'node:path';
import { DatabaseSync } from 'node:sqlite';
import { buildDocx } from '@rutba/ooxml/build';
import { Base } from '@rutba/database';
import { CHANNEL_PREFIX } from '@rutba/office-shell/contract';

export async function verifyWordMailMergeDatabase(h, { dir }) {
  const { open, check, until, wait, doc, sessionFor } = h;
  try {
    const letter = path.join(dir, 'club-letter.docx');
    fs.writeFileSync(letter, buildDocx({ styles: true, paragraphs: [{ text: 'Your fee is due, ' }, { text: 'The club' }] }));
    const db = path.join(dir, 'Club.rbase');
    for (const s of ['', '-wal', '-shm']) fs.rmSync(db + s, { force: true });
    const b = Base.open(db);
    b.createTable({ name: 'Members', fields: [{ name: 'First Name', type: 'text' }, { name: 'Fee', type: 'currency' }] });
    b.createTable({ name: 'Visits', fields: [{ name: 'When', type: 'date' }] });
    b.insert('Members', { 'First Name': 'Amira', Fee: 12.5 });
    b.insert('Members', { 'First Name': 'Ben', Fee: 20 });
    b.saveQuery('Dear fees', { tables: ['Members'], columns: [{ table: 'Members', field: 'First Name' }, { table: 'Members', field: 'Fee', criteria: ['> 15'] }] });
    b.close();
    const steps = () => { const d = new DatabaseSync(db); try { return d.prepare('SELECT COUNT(*) AS n FROM rutba_steps').get().n; } catch { return null; } finally { d.close(); } };
    const before = steps();

    const win = await open('word', letter);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = sessionFor('doc');
    const model = () => doc.model({ id: session.id });
    const clickRibbon = (title) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}) && !n.disabled);
      if (!b) return 'no button ' + ${JSON.stringify(title)};
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      b.click();
      return 'clicked';
    })()`);
    const clickMenu = (label) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(label)} && !n.disabled);
      if (!b) return 'no menu item ' + ${JSON.stringify(label)};
      b.click();
      return 'clicked';
    })()`);
    const blockText = (i) => js(`(() => { const b = document.querySelector('.wd-page [data-block="${i}"]'); return b ? b.innerText.replace(/\\u200b/g, '') : null; })()`);

    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="1"]'))`), 'the letter', 8000);
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Mailings')?.click(), 1`);
    await wait(250);
    await clickRibbon('Start Mail Merge');
    await until(() => js(`Boolean(document.querySelector('.rw-menu'))`), 'the Start Mail Merge menu', 3000).catch(() => {});
    await clickMenu('Letters');
    await until(() => model().mailMerge?.type === 'formLetters', 'a merge letter', 5000).catch(() => {});

    // The database, handed over by path: Select Table asks which list.
    win.webContents.send(`${CHANNEL_PREFIX}/event/app:command`, { command: 'mailings.useList', args: { path: db } });
    await until(() => js(`document.querySelectorAll('.wd-mm-pick [data-sheet]').length === 3`), 'Select Table', 6000).catch(() => {});
    const offered = await js(`[...document.querySelectorAll('.wd-mm-pick [data-sheet]')].map((b) => b.dataset.sheet)`);
    await js(`document.querySelector('.wd-mm-pick [data-sheet="Members"]')?.click(), 1`);
    await wait(100);
    await js(`(() => { const b = [...document.querySelectorAll('.rw-dialog .rw-dialog-foot button')].find((n) => n.textContent.trim() === 'OK'); b?.click(); return Boolean(b); })()`);
    const attached = await until(() => model().mailMerge?.source?.count === 2, 'the members attached', 6000).then(() => true, () => false);
    const src = model().mailMerge?.source || {};
    check("word: Use an Existing List on a Data Files database offers its tables and saved queries, and the table chosen is the recipient list",
      JSON.stringify(offered) === JSON.stringify(['Members', 'Visits', 'Dear fees']) && attached && src.kind === 'rbase' && src.sheet === 'Members' && JSON.stringify(src.fields) === JSON.stringify(['ID', 'First Name', 'Fee']),
      JSON.stringify({ offered, attached, src: { kind: src.kind, sheet: src.sheet, fields: src.fields, count: src.count } }));

    // «First_Name» at the end of the first paragraph, then Preview Results.
    await js(`(() => { const b = document.querySelector('.wd-page [data-block="0"]'); b.scrollIntoView({ block: 'center' }); return 1; })()`);
    const at = await js(`(() => {
      const b = document.querySelector('.wd-page [data-block="0"]');
      const w = document.createTreeWalker(b, NodeFilter.SHOW_TEXT); let last = null;
      for (let n = w.nextNode(); n; n = w.nextNode()) if (n.nodeValue.length) last = n;
      const r = document.createRange(); r.setStart(last, last.nodeValue.length - 1); r.setEnd(last, last.nodeValue.length);
      const rect = r.getClientRects()[0];
      return { x: Math.round(rect.right + 2), y: Math.round(rect.top + rect.height / 2) };
    })()`);
    win.webContents.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
    win.webContents.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1 });
    await until(() => model().selection?.focus?.block === 0, 'the caret', 5000).catch(() => {});
    await clickRibbon('Insert Merge Field');
    await until(() => js(`Boolean(document.querySelector('.rw-menu'))`), 'the field menu', 3000).catch(() => {});
    await clickMenu('First Name');
    await until(() => (model().blocks?.[0]?.runs || []).some((r) => r.field?.kind === 'mergefield'), 'the merge field', 5000).catch(() => {});
    await clickRibbon('Preview Results');
    const shown = await until(async () => (await blockText(0)) === 'Your fee is due, Amira', 'Amira on the page', 6000).then(() => true, () => false);
    const after = steps();
    check("word: a merge field from the database's table shows the first member's name in Preview Results, and the database is not changed",
      shown && before !== null && after === before, JSON.stringify({ shown, text: await blockText(0), before, after }));
  } catch (err) {
    check('word: the database mail merge checks ran', false, err.stack || err.message);
  }
}
