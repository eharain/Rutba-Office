// Documents: Table Layout.
//
// With the caret in a table the Table Layout tab is there, as in Word, and
// does what its buttons say: a row inserted below, a column to the right, the
// selected cells merged into one and split back, a column set to a width, the
// table turned right to left, a row deleted; a right-click in the table offers
// the same; and the tab goes when the caret leaves the table. Run alone with
// RUTBA_VERIFY_ONLY=wordtable.
import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { openDocx } from '@rutba/doc-view/backends/ooxml';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where the fixture is written
 */
export async function verifyWordTable(h, { dir }) {
  const { open, check, until, wait, press, errorsIn, doc, sessionFor } = h;
  const file = path.join(dir, 'table-layout.docx');
  const made = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'A table to lay out.' }, { text: 'After the table.' }] }));
  made.setSelection({ block: 0, offset: 0 });
  made.insertTable({ rows: 2, cols: 2 });
  fs.writeFileSync(file, made.save());
  try {
    const win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const session = sessionFor('doc');
    const model = () => doc.model({ id: session.id });
    // The table's shape as the model has it: rows, and cells in the first row.
    const shape = () => {
      const cells = (model().blocks || []).map((b) => /^t\d+:r(\d+):c(\d+)$/.exec(b.container || '')).filter(Boolean);
      const rows = new Set(cells.map((m) => m[1]));
      return { rows: rows.size, cols: cells.filter((m) => m[1] === '0').length };
    };
    const cellBlock = (r, c) => (model().blocks || []).findIndex((b) => new RegExp(`^t\\d+:r${r}:c${c}$`).test(b.container || ''));
    const clickCell = async (r, c, shift = false) => {
      const at = await js(`(() => { const td = document.querySelectorAll('.wd-table tr')[${r}]?.children[${c}]; if (!td) return null; const b = td.getBoundingClientRect(); return { x: Math.round(b.left + b.width / 2), y: Math.round(b.top + b.height / 2) }; })()`);
      if (!at) return false;
      const mods = shift ? ['shift'] : [];
      wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1, modifiers: mods });
      wc.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1, modifiers: mods });
      return until(() => model().selection?.focus?.block === cellBlock(r, c), `the caret in cell ${r},${c}`, 4000).then(() => true, () => false);
    };
    const tabNames = () => js(`[...document.querySelectorAll('.rw-tab')].map((t) => t.textContent.trim())`);
    const openTab = (name) => js(`(() => { const t = [...document.querySelectorAll('.rw-tab')].find((n) => n.textContent.trim() === ${JSON.stringify(name)}); if (!t) return false; t.click(); return true; })()`);
    const press_ = (label) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.textContent || '').trim() === ${JSON.stringify(label)});
      if (!b) return 'missing';
      if (b.disabled) return 'disabled';
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      b.click();
      return 'clicked';
    })()`);
    const pick = async (label) => {
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}))`), `"${label}"`, 3000).catch(() => {});
      return js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}); if (!b) return 'missing'; b.click(); return 'picked'; })()`);
    };
    await until(() => js(`Boolean(document.querySelector('.wd-table td'))`), 'the table', 8000);

    // In a cell, the tab is there.
    const inCell = await clickCell(0, 0);
    const names = await tabNames();
    check('word: with the caret in a table the Table Layout tab is there', inCell && names.includes('Table Layout'), `caret ${inCell}; tabs ${names.join(', ')}`);
    await openTab('Table Layout');
    await wait(250);

    // A row below, a column to the right.
    const below = await press_('Insert Below');
    const threeRows = await until(() => shape().rows === 3, 'three rows', 4000).then(() => true, () => false);
    const right = await press_('Insert Right');
    const threeCols = await until(() => shape().cols === 3, 'three columns', 4000).then(() => true, () => false);
    check('word: Insert Below and Insert Right add a row and a column', below === 'clicked' && right === 'clicked' && threeRows && threeCols, `${below}/${right}; ${JSON.stringify(shape())}`);

    // Two cells selected, merged, and split back.
    await clickCell(0, 0);
    await clickCell(0, 1, true);
    await wait(200);
    const merged = await press_('Merge Cells');
    const two = await until(() => shape().cols === 2, 'the first row two cells', 4000).then(() => true, () => false);
    await clickCell(0, 0);
    await openTab('Table Layout');
    await wait(200);
    const split = await press_('Split Cells');
    const backToThree = await until(() => shape().cols === 3, 'three cells again', 4000).then(() => true, () => false);
    check('word: Merge Cells joins the selected cells and Split Cells parts them again', merged === 'clicked' && two && split === 'clicked' && backToThree, `merge ${merged} ${two}; split ${split} ${backToThree}`);

    // A width for the caret's column.
    await clickCell(0, 0);
    await openTab('Table Layout');
    await wait(200);
    const typed = await js(`(() => {
      const box = [...document.querySelectorAll('.rw-ribbon .wd-fields')].find((f) => f.textContent.includes('Width'))?.querySelector('input');
      if (!box) return false;
      box.focus();
      const set = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
      set.call(box, '4');
      box.dispatchEvent(new Event('input', { bubbles: true }));
      box.blur();
      return true;
    })()`);
    await wait(600);
    if (process.env.RUTBA_VERIFY_CAPTURE) { wc.invalidate(); await wait(400); fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'word-table-layout.png'), (await wc.capturePage()).toPNG()); }

    // Right to left, then a row deleted.
    const rtl = await press_('Right to Left');
    const turned = await until(() => js(`getComputedStyle(document.querySelector('.wd-table')).direction === 'rtl'`), 'the table right to left', 4000).then(() => true, () => false);
    check('word: Right to Left turns the table to run from the right', rtl === 'clicked' && turned, `${rtl}; ${turned}`);
    const del = await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.textContent || '').trim() === 'Delete'); if (!b) return 'missing'; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); b.click(); return 'clicked'; })()`);
    const pickedRows = await pick('Delete rows');
    const twoRows = await until(() => shape().rows === 2, 'two rows', 4000).then(() => true, () => false);
    check('word: Delete → Delete rows takes the caret\'s row out', del === 'clicked' && pickedRows === 'picked' && twoRows, `${del}/${pickedRows}; ${JSON.stringify(shape())}`);

    // The right-click in a cell offers the same.
    await clickCell(0, 0);
    const at = await js(`(() => { const b = document.querySelector('.wd-table td').getBoundingClientRect(); return { x: Math.round(b.left + 10), y: Math.round(b.top + 8) }; })()`);
    wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'right', clickCount: 1 });
    wc.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'right', clickCount: 1 });
    wc.sendInputEvent({ type: 'contextMenu', x: at.x, y: at.y });
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === 'Insert row above'))`), 'the table items', 3000).catch(() => {});
    const offered = await js(`[...document.querySelectorAll('.rw-menu button')].map((b) => b.textContent.trim()).filter((t) => /row|column|cells/i.test(t))`);
    await press(wc, 'Escape');
    check('word: a right-click in a table offers its rows, columns and cells', ['Insert row above', 'Insert column right', 'Delete row', 'Merge cells', 'Split cells'].every((t) => offered.includes(t)), offered.join(', '));

    // Saved: the width and the direction as Word keeps them.
    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    await press(wc, 's', { modifiers: ['control'] });
    await wait(1400);
    const xml = openDocx(fs.readFileSync(file)).doc.doc.xml;
    const firstCol = Number(/<w:gridCol w:w="(\d+)"/.exec(xml)?.[1] || 0);
    check('word: a column set to 4 cm and the right-to-left table are saved as Word keeps them',
      typed === true && Math.abs(firstCol - 2268) <= 30 && /<w:tblPr>[\s\S]*?<w:bidiVisual\/>/.test(xml), `typed ${typed}; first column ${firstCol} twips; bidiVisual ${/<w:bidiVisual\/>/.test(xml)}`);

    // Out of the table, the tab goes.
    const outside = await js(`(() => { const p = [...document.querySelectorAll('.wd-page > .wd-block')].find((n) => n.textContent.includes('After the table')); if (!p) return null; const b = p.getBoundingClientRect(); return { x: Math.round(b.left + 20), y: Math.round(b.top + 6) }; })()`);
    if (outside) {
      wc.sendInputEvent({ type: 'mouseDown', x: outside.x, y: outside.y, button: 'left', clickCount: 1 });
      wc.sendInputEvent({ type: 'mouseUp', x: outside.x, y: outside.y, button: 'left', clickCount: 1 });
    }
    const gone = await until(async () => !(await tabNames()).includes('Table Layout'), 'the tab gone', 4000).then(() => true, () => false);
    check('word: the Table Layout tab goes when the caret leaves the table', gone, (await tabNames()).join(', '));

    const complaints = await errorsIn(win);
    check('word: Table Layout reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the Table Layout checks ran', false, err.message);
  }
}
