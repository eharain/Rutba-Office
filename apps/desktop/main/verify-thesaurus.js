// Review → Thesaurus in Documents and Worksheets.
//
// With the caret in "good", Review → Thesaurus opens the pane on its
// meanings; Insert beside "excellent" puts that in its place; a word clicked
// is looked up in turn. In a workbook, the active cell's word is looked up
// and Insert writes the cell. Run alone with RUTBA_VERIFY_ONLY=thesaurus.
import fs from 'node:fs';
import path from 'node:path';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';

export async function verifyThesaurus({ open, check, until, wait, errorsIn, doc, sessionFor }, { dir }) {
  const tab = (js, name) => js(`(() => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(); return 1; })()`);
  const press = (js, label) => js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); if (!b) return 'none'; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); b.click(); return 'pressed'; })()`);
  const words = (js) => js(`[...document.querySelectorAll('.pf-th-word')].map((b) => b.textContent.trim())`);
  const insert = (js, word) => js(`(() => { const row = [...document.querySelectorAll('.pf-th-row')].find((r) => r.querySelector('.pf-th-word')?.textContent.trim() === ${JSON.stringify(word)}); row?.querySelector('.pf-th-insert')?.click(); return Boolean(row); })()`);

  const docFile = path.join(dir, 'thesaurus.docx');
  try {
    fs.writeFileSync(docFile, buildDocx({ styles: true, paragraphs: [{ text: 'The plan was good.' }] }));
    const win = await open('word', docFile);
    const js = (code) => win.webContents.executeJavaScript(code);
    const model = () => doc.model({ id: sessionFor('doc').id });
    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="0"]'))`), 'the paragraph', 8000);
    // The caret into "good".
    await js(`(() => {
      const p = document.querySelector('.wd-page [data-block="0"]');
      p.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }));
      const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
      let node; let at = 15;
      while ((node = walker.nextNode())) { if (at <= node.length) break; at -= node.length; }
      const r = document.createRange(); r.setStart(node, at); r.collapse(true);
      const s = getSelection(); s.removeAllRanges(); s.addRange(r);
      document.querySelector('.wd-page').dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
      return 1;
    })()`);
    await until(() => model().selection?.focus?.offset === 15, 'the caret in good', 4000).catch(() => {});
    await tab(js, 'Review');
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Thesaurus'))`), 'Review → Thesaurus', 4000);
    await press(js, 'Thesaurus');
    await until(async () => (await words(js)).includes('excellent'), 'the meanings of good', 6000).catch(() => {});
    const listed = await words(js);
    check('word: Review → Thesaurus opens on the word at the caret, its meanings listed', listed.includes('excellent') && listed.includes('fine'), listed.slice(0, 12).join(', '));
    await insert(js, 'excellent');
    await until(() => model().blocks[0].text === 'The plan was excellent.', 'the word put in', 5000).catch(() => {});
    check('word: Insert puts the word in place of the one at the caret', model().blocks[0].text === 'The plan was excellent.', model().blocks[0].text);
    await js(`[...document.querySelectorAll('.pf-th-word')].find((b) => b.textContent.trim() === 'superb')?.click(), 'look up'`);
    await until(() => js(`document.querySelector('.pf-th-input')?.value === 'superb'`), 'superb looked up', 4000).catch(() => {});
    const back = await js(`Boolean(document.querySelector('.pf-th-back'))`);
    check('word: a word clicked in the pane is looked up in turn, with Back to the one before', back && (await js(`document.querySelector('.pf-th-input')?.value`)) === 'superb', `back ${back}`);
    const complaints = await errorsIn(win);
    check('word: the thesaurus reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the thesaurus checks ran', false, err.message);
  }

  const bookFile = path.join(dir, 'thesaurus.xlsx');
  try {
    fs.writeFileSync(bookFile, buildXlsx({ sheets: [{ name: 'Words', rows: [['happy', 'a plain day']] }] }));
    const win = await open('sheets', bookFile);
    const js = (code) => win.webContents.executeJavaScript(code);
    const model = () => doc.model({ id: sessionFor('sheet').id });
    await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="A1"]'))`), 'the grid', 8000);
    await js(`(() => { document.querySelector('.sh-cell[data-ref="A1"]').dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 })); return 1; })()`);
    await until(() => (model().cells || []).find((c) => c.ref === 'A1')?.active === true, 'A1 active', 4000).catch(() => {});
    await tab(js, 'Review');
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Thesaurus'))`), 'Review → Thesaurus', 4000);
    await press(js, 'Thesaurus');
    await until(async () => (await words(js)).includes('glad'), 'the meanings of happy', 6000).catch(() => {});
    await insert(js, 'glad');
    await until(() => (model().cells || []).find((c) => c.ref === 'A1')?.text === 'glad', 'the cell written', 5000).catch(() => {});
    check('sheets: Review → Thesaurus looks up the active cell\'s word, and Insert writes the cell', (model().cells || []).find((c) => c.ref === 'A1')?.text === 'glad', String((model().cells || []).find((c) => c.ref === 'A1')?.text));
    const complaints = await errorsIn(win);
    check('sheets: the thesaurus reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('sheets: the thesaurus checks ran', false, err.message);
  }
}
