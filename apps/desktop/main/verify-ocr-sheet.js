// Worksheets: Data from Picture, pressed in the window. Data → From Picture →
// Picture From File reads a picture of a ruled table on the main process's
// threads and shows it beside a preview grid of its cells; a cell corrected
// in the preview goes in as corrected; Insert Data puts the table in at the
// selected cell, numbers as numbers, as one step Undo takes back and Redo
// puts back; Picture From Clipboard reads a table set with space alone,
// with no rules, from the clipboard into another place. The clipboard's text
// and picture are kept and put back.
// Run alone with RUTBA_VERIFY_ONLY=ocr-sheet.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { clipboard, nativeImage } from 'electron';
import { buildXlsx } from '@rutba/ooxml/build';

const fixture = (name) => fileURLToPath(new URL(`../test/fixtures/${name}`, import.meta.resolve('@rutba/ocr')));

const answer = (kind, value) => {
  (globalThis.__rutbaCheckDialogAnswers ||= []).push({ kind, answer: value });
};

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where the files are made
 */
export async function verifyOcrSheet(h, { dir }) {
  const { open, check, until, wait, errorsIn, doc, sessionFor } = h;
  const root = path.join(dir, 'ocr-sheet');
  fs.mkdirSync(root, { recursive: true });
  const book = path.join(root, 'from-picture.xlsx');
  fs.writeFileSync(book, buildXlsx({ sheets: [{ name: 'Data', rows: [[]] }] }));
  const ruled = path.join(root, 'invoice-table.png');
  fs.copyFileSync(fixture('table-ruled.png'), ruled);
  const truth = JSON.parse(fs.readFileSync(fixture('table.json'), 'utf8'));
  const keptText = clipboard.readText();
  const keptImage = clipboard.readImage();

  const win = await open('sheets', book);
  const wc = win.webContents;
  const js = (code) => wc.executeJavaScript(code);
  const capture = async (name) => {
    if (!process.env.RUTBA_VERIFY_CAPTURE) return;
    wc.invalidate();
    await wait(700);
    fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG());
  };
  const model = () => doc.model({ id: sessionFor('sheet').id });
  const cell = (ref) => (model().cells || []).find((c) => c.ref === ref) || null;
  const text = (ref) => String(cell(ref)?.text ?? '');
  const fromPicture = async (item) => {
    await js(`[...document.querySelectorAll('.rw-tab')].find((n) => n.textContent.trim() === 'Data')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('[data-role="data-from-picture"]'))`), 'the From Picture button', 5000).catch(() => {});
    await js(`document.querySelector('[data-role="data-from-picture"]')?.click(), 1`);
    await until(() => js(`[...document.querySelectorAll('.rw-menu button')].some((b) => b.textContent.includes(${JSON.stringify(item)}))`), 'the menu', 4000).catch(() => {});
    return js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.includes(${JSON.stringify(item)})); if (!b) return 'no ${item}'; b.click(); return 'clicked'; })()`);
  };
  const preview = () => js(`[...document.querySelectorAll('[data-role="pd-grid"] tr')].map((tr) => [...tr.querySelectorAll('input')].map((i) => i.value))`);
  const select = async (ref) => {
    // A press on the cell where the headings say it is, in the window only.
    const at = await js(`(() => {
      const m = /^([A-Z]+)(\\d+)$/.exec(${JSON.stringify(ref)});
      const col = [...document.querySelectorAll('.sh-colheads .sh-head')].find((n) => n.textContent.trim() === m[1])?.getBoundingClientRect();
      const row = [...document.querySelectorAll('.sh-rowheads .sh-head')].find((n) => n.textContent.trim() === m[2])?.getBoundingClientRect();
      return col && row ? { x: Math.round(col.left + col.width / 2), y: Math.round(row.top + row.height / 2) } : null;
    })()`);
    if (!at) return false;
    wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
    wc.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1 });
    return until(() => model().selection?.ref === ref || model().selection?.ref?.startsWith?.(`${ref}:`), `${ref} selected`, 4000).then(() => true).catch(() => false);
  };

  try {
    await until(() => js(`Boolean(document.querySelector('.sh-colheads .sh-head'))`), 'the grid', 10000).catch(() => {});
    await js(`localStorage.setItem('rutba-ocr-languages', '["eng"]'), 1`);
    await select('B2');

    // From a file: the preview, the picture beside it.
    answer('open', [ruled]);
    const pressed = await fromPicture('Picture From File');
    const shown = await until(() => js(`Boolean(document.querySelector('[data-role="pd-grid"]'))`), 'the preview grid', 30000).then(() => true).catch(() => false);
    const grid = shown ? await preview() : [];
    const look = await js(`(() => { const img = document.querySelector('.pd-picture img')?.getBoundingClientRect(); const g = document.querySelector('.pd-grid-wrap')?.getBoundingClientRect(); return { img: img ? Math.round(img.width) : 0, grid: g ? Math.round(g.width) : 0, beside: Boolean(img && g && img.right <= g.left + 2), where: document.querySelector('.pd-where')?.textContent || '' }; })()`);
    await capture('ocr-sheet-01-preview.png');
    check('ocr-sheet: Data → From Picture → Picture From File shows the picture beside a preview of its cells, read right', pressed === 'clicked' && shown && JSON.stringify(grid) === JSON.stringify(truth) && look.beside && /B2/.test(look.where),
      JSON.stringify({ pressed, shown, grid, look }));

    // A cell corrected in the preview goes in as corrected.
    await js(`(() => {
      const input = document.querySelectorAll('[data-role="pd-grid"] tr')[2]?.querySelectorAll('input')[0];
      if (!input) return false;
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(input, 'Ink cartridges');
      input.dispatchEvent(new Event('input', { bubbles: true }));
      return true;
    })()`);
    await until(async () => (await preview())[2]?.[0] === 'Ink cartridges', 'the corrected cell', 3000).catch(() => {});
    await js(`document.querySelector('[data-role="pd-insert"]')?.click(), 1`);
    const landed = await until(() => text('B2') === 'Item' && /^6\.20?$/.test(text('D6')), 'the data in the sheet', 8000).then(() => true).catch(() => false);
    const got = { B2: text('B2'), C3: text('C3'), B4: text('B4'), D6: text('D6'), numberRight: cell('C3')?.align === 'right' && cell('D4')?.align === 'right', textLeft: cell('B3')?.align !== 'right' };
    await capture('ocr-sheet-02-inserted.png');
    check('ocr-sheet: Insert Data puts the table in at the selected cell, the corrected cell as corrected and numbers as numbers', landed && got.B2 === 'Item' && got.C3 === '12' && got.B4 === 'Ink cartridges' && /^6\.20?$/.test(got.D6) && got.numberRight && got.textLeft,
      JSON.stringify(got));

    // One step: Undo takes it all back, Redo puts it back.
    await js(`[...document.querySelectorAll('.rw-btn')].find((n) => (n.dataset.tip || n.title || '').startsWith('Undo'))?.click(), 1`);
    const undone = await until(() => !text('B2') && !text('D6'), 'the data gone after Undo', 5000).then(() => true).catch(() => false);
    const redoButton = `[...document.querySelectorAll('.rw-btn')].find((n) => (n.dataset.tip || n.title || '').startsWith('Redo'))`;
    await until(() => js(`Boolean(${redoButton} && !${redoButton}.disabled)`), 'Redo ready', 4000).catch(() => {});
    await wait(300);
    await js(`${redoButton}?.click(), 1`);
    const redone = await until(() => text('B2') === 'Item' && text('C3') === '12', 'the data back after Redo', 5000).then(() => true).catch(() => false);
    check('ocr-sheet: the data from the picture is one step that Undo takes back and Redo puts back', undone && redone, JSON.stringify({ undone, redone }));

    // From the clipboard: a table with no rules at all.
    clipboard.writeImage(nativeImage.createFromPath(fixture('table-plain.png')));
    await select('G1');
    const cpressed = await fromPicture('Picture From Clipboard');
    const cshown = await until(() => js(`Boolean(document.querySelector('[data-role="pd-grid"]'))`), 'the preview from the clipboard', 30000).then(() => true).catch(() => false);
    const cgrid = cshown ? await preview() : [];
    await capture('ocr-sheet-03-clipboard.png');
    await js(`document.querySelector('[data-role="pd-insert"]')?.click(), 1`);
    const clanded = await until(() => text('G1') === 'Item' && text('H5') === '1' && /^31(\.00)?$/.test(text('I3')), 'the clipboard table in', 8000).then(() => true).catch(() => false);
    check('ocr-sheet: Picture From Clipboard reads a table set with space alone, no rules, and puts it in at the cell chosen', cpressed === 'clicked' && cshown && JSON.stringify(cgrid) === JSON.stringify(truth) && clanded,
      JSON.stringify({ cpressed, cshown, cgrid, G1: text('G1'), H5: text('H5'), I3: text('I3') }));
    const errors = await errorsIn(win);
    check('ocr-sheet: no error was shown', !errors.length, errors.join(' | '));
    await js(`localStorage.removeItem('rutba-ocr-languages'), 1`).catch(() => {});
  } finally {
    if (keptImage && !keptImage.isEmpty()) clipboard.write({ text: keptText || '', image: keptImage });
    else clipboard.writeText(keptText || '');
  }
}
