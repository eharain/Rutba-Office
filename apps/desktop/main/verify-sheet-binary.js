// Worksheets: an Excel 97-2003 workbook, opened in full.
//
// A .xls Excel itself wrote (tests/fixtures/binary/showcase.xls) opens as
// its sheets — the values and the formulas' results, the header row bold
// on its fill, the frozen first row and column, the other sheets in their
// tabs — where it used to open as one cell saying it could not be read.
// Run alone with RUTBA_VERIFY_ONLY=xls.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const FIXTURE = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..', '..', 'tests', 'fixtures', 'binary', 'showcase.xls');

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where the copy is opened from
 */
export async function verifySheetBinary(h, { dir }) {
  const { open, check, until, errorsIn, doc, sessionFor } = h;
  const file = path.join(dir, 'showcase.xls');
  try {
    fs.copyFileSync(FIXTURE, file);
    const win = await open('sheets', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = sessionFor('sheet');
    await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="F2"]'))`), 'the cells to be drawn', 8000).catch(() => {});
    if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'sheet-binary.png'), (await win.webContents.capturePage()).toPNG());
    const shown = await js(`(() => {
      const cell = (ref) => document.querySelector('.sh-cell[data-ref="' + ref + '"]');
      const a1 = cell('A1');
      return {
        a1: a1?.innerText.trim() ?? null,
        f2: cell('F2')?.innerText.trim() ?? null,
        bold: a1 ? Number(getComputedStyle(a1).fontWeight) >= 600 : false,
        fill: a1 ? getComputedStyle(a1).backgroundColor : null,
        tabs: [...document.querySelectorAll('.sh-tab')].map((t) => t.innerText.trim()),
        pinned: document.querySelectorAll('.sh-pin-rows .sh-cell, .sh-pin-corner .sh-cell').length,
      };
    })()`);
    check('sheets: an Excel 97-2003 workbook opens with its values and its formulas\' results, not a "could not be converted" cell',
      session.converted?.from === 'xls' && !session.converted?.partial && shown.a1 === 'Month' && shown.f2 === '7,518',
      `from ${session.converted?.from}${session.converted?.partial ? ' (partial)' : ''}; A1 ${JSON.stringify(shown.a1)}, F2 ${JSON.stringify(shown.f2)}`);
    check('sheets: the .xls\'s header row keeps its bold and its fill, and its first row stays frozen',
      shown.bold && shown.fill && !/rgba\(0, 0, 0, 0\)|transparent/.test(shown.fill) && shown.pinned > 0,
      `bold ${shown.bold}; fill ${shown.fill}; ${shown.pinned} frozen cells`);
    check('sheets: every sheet of the .xls has its tab', ['Sales', 'Summary', 'Data types', 'Table'].every((t) => shown.tabs.some((x) => x.startsWith(t))), shown.tabs.join(' | '));
    // The link and the note on Sales, and the picture on Summary.
    const marks = await js(`(() => ({ link: Boolean(document.querySelector('.sh-cell.link[data-ref="A19"]')), note: Boolean(document.querySelector('.sh-cell.noted[data-ref="A1"]')) }))()`);
    await js(`(() => { [...document.querySelectorAll('.sh-tab')].find((t) => t.innerText.trim().startsWith('Summary'))?.click(); return 1; })()`);
    const pictured = await until(() => js(`document.querySelectorAll('.sh-drawing').length > 0`), 'the Summary sheet\'s picture', 6000).catch(() => false);
    check('sheets: the .xls\'s hyperlink and note are on their cells, and its picture on its sheet',
      marks.link && marks.note && pictured === true, `link ${marks.link}; note ${marks.note}; picture ${pictured}`);

    const complaints = await errorsIn(win);
    check('sheets: opening an .xls reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('sheets: the .xls checks ran', false, err.message);
  }

  // Excel 95 kept its pictures another way (a DIB after the object, not Office Art): drawn all the same.
  const file95 = path.join(dir, 'showcase-95.xls');
  try {
    fs.copyFileSync(FIXTURE.replace(/showcase\.xls$/, 'showcase-95.xls'), file95);
    const win = await open('sheets', file95);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.sh-tab').length > 1`), 'the sheet tabs', 8000).catch(() => {});
    await js(`(() => { [...document.querySelectorAll('.sh-tab')].find((t) => t.innerText.trim().startsWith('Summary'))?.click(); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.sh-drawing image'))`), 'the picture', 6000).catch(() => {});
    const width = await js(`new Promise((done) => {
      const node = document.querySelector('.sh-drawing image');
      const href = node?.getAttribute('href') || node?.getAttribute('xlink:href');
      if (!href) return done(0);
      const img = new Image();
      img.onload = () => done(img.naturalWidth);
      img.onerror = () => done(-1);
      img.src = href;
    })`);
    check('sheets: an Excel 95 workbook\'s picture is drawn on its sheet', width === 240, `the picture decodes ${width} pixels wide`);
  } catch (err) {
    check('sheets: the Excel 95 check ran', false, err.message);
  }
}
