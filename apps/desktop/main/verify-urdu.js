// The windows in Urdu: the language chosen, a window opened after it reads
// right to left in its frame — the ribbon's tabs in Urdu, the page lang and
// dir — while a document's page, a worksheet's grid and a slide's stage
// keep their own direction, typing still reaches the page, and a dialog
// speaks Urdu too. English again at the end, for the checks after it. Run
// alone with RUTBA_VERIFY_ONLY=urdu.
import fs from 'node:fs';
import path from 'node:path';
import { CATALOGUES } from '@rutba/office-ui/catalogues';
import { buildDocx } from '@rutba/ooxml/build';

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string, files: object }} args
 */
export async function verifyUrdu({ open, check, until, wait, errorsIn }, { dir, files }) {
  const ur = CATALOGUES.ur || {};
  const capture = process.env.RUTBA_VERIFY_CAPTURE;
  const snap = async (win, name) => { if (capture) { win.webContents.invalidate(); await wait(400); fs.writeFileSync(path.join(capture, name), (await win.webContents.capturePage()).toPNG()); } };
  let home = null;
  try {
    // The language chosen where a person chooses it: the app menu's store.
    home = await open('home');
    const hjs = (code) => home.webContents.executeJavaScript(code);
    await hjs(`window.rutbaOffice.store.set({ key: 'language', value: 'ur' }).then(() => 1)`);

    // Documents.
    const file = path.join(dir, 'urdu-words.docx');
    fs.writeFileSync(file, buildDocx({ styles: true, paragraphs: [{ text: 'An English line.' }, { text: 'Another line.' }] }));
    const word = await open('word', file);
    const js = (code) => word.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="0"]'))`), 'the page', 10000);
    const frame = await js(`({ lang: document.documentElement.lang, dir: document.documentElement.dir, tabs: [...document.querySelectorAll('.rw-tab')].map((t) => t.textContent.trim()), page: getComputedStyle(document.querySelector('.wd-scroll')).direction, ribbon: getComputedStyle(document.querySelector('.rw-ribbon')).direction })`);
    const homeTab = ur.Home || 'Home';
    const insertTab = ur.Insert || 'Insert';
    check('urdu: a Documents window opened after Urdu is chosen is in Urdu, its frame right to left and its page left to right as the document has it',
      frame.lang === 'ur' && frame.dir === 'rtl' && frame.ribbon === 'rtl' && frame.page === 'ltr' && frame.tabs.includes(homeTab) && frame.tabs.includes(insertTab) && !frame.tabs.includes('Home'),
      JSON.stringify(frame));

    // Typing still reaches the page.
    const at = await js(`(() => { const b = document.querySelector('.wd-page [data-block="1"]'); const r = b.getBoundingClientRect(); return { x: Math.round(r.right - 4), y: Math.round(r.top + r.height / 2) }; })()`);
    word.webContents.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
    word.webContents.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1 });
    await wait(300);
    for (const ch of ' اردو') word.webContents.sendInputEvent({ type: 'char', keyCode: ch });
    const typed = await until(() => js(`document.querySelector('.wd-page [data-block="1"]')?.textContent.includes('اردو')`), 'the Urdu words on the page', 5000).then(() => true).catch(() => false);
    check('urdu: typing in an Urdu window still reaches the page', typed, String(await js(`document.querySelector('.wd-page [data-block="1"]')?.textContent`)));

    // A dialog speaks Urdu: Insert → Table.
    await js(`(() => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(insertTab)})?.click(); return 1; })()`);
    await wait(250);
    await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(ur.Table || 'Table')}); b?.click(); return Boolean(b); })()`);
    // until() says only that it came; the dialog's words are read after.
    await until(() => js(`Boolean(document.querySelector('.rw-dialog')?.textContent)`), 'the table dialog', 4000).catch(() => {});
    const dialog = await js(`document.querySelector('.rw-dialog')?.textContent || ''`);
    await snap(word, 'urdu-documents.png');
    check('urdu: a dialog in an Urdu window speaks Urdu', /[\u0600-\u06FF]/.test(dialog) && !/Insert table/.test(dialog), String(dialog).slice(0, 120));
    await js(`(() => { [...document.querySelectorAll('.rw-dialog .rw-btn')].find((b) => b.textContent.trim() === ${JSON.stringify(ur.Cancel || 'Cancel')})?.click(); return 1; })()`);

    // Worksheets: the grid's columns run A, B, C from the left; the formula bar left to right.
    const sheet = await open('sheets', files.xlsx);
    const sjs = (code) => sheet.webContents.executeJavaScript(code);
    await until(() => sjs(`Boolean(document.querySelector('.sh-grid'))`), 'the grid', 10000);
    const grid = await sjs(`({ dir: document.documentElement.dir, grid: getComputedStyle(document.querySelector('.sh-grid')).direction, formula: getComputedStyle(document.querySelector('.sh-formula') || document.body).direction, tabs: [...document.querySelectorAll('.rw-tab')].map((t) => t.textContent.trim()) })`);
    await snap(sheet, 'urdu-worksheets.png');
    check('urdu: a Worksheets window in Urdu keeps its grid and formula bar left to right, its ribbon in Urdu',
      grid.dir === 'rtl' && grid.grid === 'ltr' && grid.formula === 'ltr' && grid.tabs.includes(homeTab), JSON.stringify(grid));

    // Presentations: the stage as the slide has it.
    const deck = await open('slides', files.pptx);
    const djs = (code) => deck.webContents.executeJavaScript(code);
    await until(() => djs(`Boolean(document.querySelector('.sl-stage'))`), 'the stage', 10000);
    const stage = await djs(`({ dir: document.documentElement.dir, stage: getComputedStyle(document.querySelector('.sl-stage')).direction })`);
    await snap(deck, 'urdu-presentations.png');
    check('urdu: a Presentations window in Urdu keeps its stage as the slide has it', stage.dir === 'rtl' && stage.stage === 'ltr', JSON.stringify(stage));

    const complaints = [...await errorsIn(word), ...await errorsIn(sheet), ...await errorsIn(deck)];
    check('urdu: the Urdu windows report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('urdu: the Urdu checks ran', false, err?.message || String(err));
  } finally {
    // English again, for every check after these.
    try {
      await home?.webContents.executeJavaScript(`window.rutbaOffice.store.set({ key: 'language', value: 'en' }).then(() => 1)`);
    } catch {
      /* the home window gone: the run's own settings say English at the next start */
    }
  }
}
