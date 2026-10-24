// Word: Save As in the Documents app menu (2.18.0, A3.7). In 2.17.1 the
// menu behind the app mark had Open and Save as Template but no Save As, so
// a copy under a new name was only Ctrl+Shift+S away. Now Save as… sits
// beside them, writes the document under the name chosen, and the window
// goes on with the new file. Run alone with RUTBA_VERIFY_ONLY=wordsaveas.
import fs from 'node:fs';
import path from 'node:path';
import { OoxmlPackage } from '@rutba/ooxml';
import { buildDocx } from '@rutba/ooxml/build';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn
 * @param {{ dir: string }} args where the fixture is written
 */
export async function verifyWordSaveAs({ open, check, until, wait, press, errorsIn }, { dir }) {
  const file = path.join(dir, 'rota.docx');
  const copy = path.join(dir, 'rota-december.docx');
  fs.writeFileSync(file, buildDocx({ paragraphs: ['Rota for November', 'Sam opens on Mondays.'] }));
  const original = fs.readFileSync(file);
  let win = null;
  try {
    win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const items = () => js(`[...document.querySelectorAll('.rw-menu button')].map((b) => b.textContent.trim())`);
    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="1"]'))`), 'the paragraphs', 8000);

    // A word changed, then the menu behind the app mark.
    await js(`(() => {
      const el = document.querySelector('.wd-page [data-block="0"]');
      const t = document.createTreeWalker(el, NodeFilter.SHOW_TEXT).nextNode();
      document.querySelector('.wd-page').focus();
      const r = document.createRange(); r.setStart(t, 9); r.setEnd(t, 17);
      const s = getSelection(); s.removeAllRanges(); s.addRange(r);
      document.querySelector('.wd-page').dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
      return 1;
    })()`);
    await wait(200);
    for (const ch of 'December') {
      wc.sendInputEvent({ type: 'keyDown', keyCode: ch });
      wc.sendInputEvent({ type: 'char', keyCode: ch });
      wc.sendInputEvent({ type: 'keyUp', keyCode: ch });
      await wait(40);
    }
    await until(() => js(`document.querySelector('.wd-page [data-block="0"]')?.textContent === 'Rota for December'`), 'the word typed', 3000).catch(() => false);
    await js(`document.querySelector('.rw-appmark').click(), 1`);
    await until(async () => (await items()).some((s) => s.startsWith('Save as…')), 'the Save as item', 4000).catch(() => false);
    const listed = await items();
    const where = (label) => listed.findIndex((s) => s.startsWith(label) && !s.startsWith(label + ' '));
    const at = where('Save as…');
    check('word: the app menu offers Save as… beside Open… and Save as Template…',
      at > where('Open…') && where('Open…') >= 0 && listed[at + 1]?.startsWith('Save as Template…'),
      JSON.stringify(listed.slice(0, 8)));

    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: copy });
    await js(`[...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim().startsWith('Save as…') && !b.textContent.includes('Template'))?.click(), 1`);
    let words = '';
    const written = await until(() => {
      try {
        words = OoxmlPackage.read(fs.readFileSync(copy)).text('word/document.xml');
        return words.includes('Rota for December') || (words.includes('Rota for ') && words.includes('December'));
      } catch { return false; }
    }, 'the copy written', 6000).then(() => true).catch(() => false);
    const untouched = Buffer.compare(fs.readFileSync(file), original) === 0;
    const title = await until(() => js(`document.title.includes('rota-december') || document.body.textContent.includes('rota-december')`), 'the window on the copy', 4000).then(() => true).catch(() => false);
    check('word: Save as… writes the document under the new name, leaves the first file as it was, and the window goes on with the copy',
      written && untouched && title, JSON.stringify({ written, untouched, title }));

    const complaints = await errorsIn(win);
    check('word: the Save As checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the Save As checks ran', false, err?.message || String(err));
  }
}
