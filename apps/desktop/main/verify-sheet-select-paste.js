// Worksheets: a range chosen with the mouse, the clipboard, Paste Special,
// the fill handle's series and its Auto Fill Options, and Fill → Series.
//
// With the real mouse and keyboard: a drag from A1 to B3 selects the range,
// a drag across the column headings selects whole columns; Ctrl+C puts the
// cells on the system clipboard as words and as a table, Ctrl+V pastes them,
// Ctrl+X and Ctrl+V move them; Paste Special pastes them turned rows to
// columns; two numbers dragged by the fill handle carry on as a series, and
// Auto Fill Options turns that into a copy; Fill → Series runs a step to a
// stop. Run alone with RUTBA_VERIFY_ONLY=sheetpaste.
import fs from 'node:fs';
import path from 'node:path';
import { clipboard } from 'electron';
import { buildXlsx } from '@rutba/ooxml/build';

export async function verifySheetSelectPaste({ open, check, until, wait, press, errorsIn, doc, sessionFor }, { dir }) {
  const file = path.join(dir, 'select-paste.xlsx');
  try {
    fs.writeFileSync(file, buildXlsx({ sheets: [{ name: 'Data', rows: [['Item', 'Qty'], ['Ink', 4], ['Paper', 6], [], [1], [2], [], [null, null, 10]] }] }));
    const win = await open('sheets', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const model = () => doc.model({ id: sessionFor('sheet').id });
    const text = (ref) => String((model().cells || []).find((c) => c.ref === ref)?.text ?? '');
    const range = () => { const s = model().selection; return `${s.top},${s.left},${s.bottom},${s.right}`; };
    await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="B3"]'))`), 'the grid', 8000);
    win.focus();
    const centre = (selector) => js(`(() => { const r = document.querySelector(${JSON.stringify(selector)})?.getBoundingClientRect(); return r ? { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) } : null; })()`);
    // The empty grid has no drawn cell: its place from the headings.
    const spot = (ref) => js(`(() => {
      const m = /^([A-Z]+)(\\d+)$/.exec(${JSON.stringify(ref)});
      const col = [...document.querySelectorAll('.sh-colheads .sh-head')].find((h) => h.textContent.trim() === m[1])?.getBoundingClientRect();
      const row = [...document.querySelectorAll('.sh-rowheads .sh-head')].find((h) => h.textContent.trim() === m[2])?.getBoundingClientRect();
      return col && row ? { x: Math.round(col.left + col.width / 2), y: Math.round(row.top + row.height / 2) } : null;
    })()`);
    const drag = async (from, to, steps = 6) => {
      wc.sendInputEvent({ type: 'mouseMove', x: from.x, y: from.y });
      wc.sendInputEvent({ type: 'mouseDown', x: from.x, y: from.y, button: 'left', clickCount: 1 });
      await wait(60);
      for (let i = 1; i <= steps; i++) {
        wc.sendInputEvent({ type: 'mouseMove', x: Math.round(from.x + ((to.x - from.x) * i) / steps), y: Math.round(from.y + ((to.y - from.y) * i) / steps), button: 'left' });
        await wait(40);
      }
      wc.sendInputEvent({ type: 'mouseUp', x: to.x, y: to.y, button: 'left', clickCount: 1 });
      await wait(150);
    };
    const click = async (at) => drag(at, at, 1);

    // A drag from A1 to B3 selects the range.
    await drag(await spot('A1'), await spot('B3'));
    await until(() => range() === '0,0,2,1', 'A1:B3 selected', 4000).catch(() => {});
    check('sheets: a drag with the mouse from A1 to B3 selects the range', range() === '0,0,2,1', range());

    // Ctrl+C: words and a table on the system clipboard.
    clipboard.clear();
    await press(wc, 'C', { modifiers: ['control'] });
    await until(() => clipboard.readText() !== '', 'the clipboard', 4000).catch(() => {});
    const words = clipboard.readText();
    const table = clipboard.readHTML();
    check('sheets: Ctrl+C puts the cells on the system clipboard, as words and as a table', words.replace(/\r/g, '') === 'Item\tQty\nInk\t4\nPaper\t6' && /<table>[\s\S]*Paper[\s\S]*<\/table>/.test(table), JSON.stringify(words) + ' ' + table.slice(0, 80));

    // Ctrl+V at E1.
    await click(await spot('E1'));
    await press(wc, 'V', { modifiers: ['control'] });
    await until(() => text('F3') === '6', 'the paste', 5000).catch(() => {});
    check('sheets: Ctrl+V pastes the copied cells at E1', text('E1') === 'Item' && text('E2') === 'Ink' && text('F3') === '6', ['E1', 'E2', 'F3'].map(text).join(' '));

    // Paste Special, transposed, at H1.
    await click(await spot('H1'));
    await press(wc, 'V', { modifiers: ['control', 'alt'] });
    await until(() => js(`Boolean(document.querySelector('.sh-ps-ok'))`), 'the Paste Special dialog', 4000);
    await js(`(() => { document.querySelector('.sh-ps-transpose').click(); document.querySelector('.sh-ps-ok').click(); return 1; })()`);
    await until(() => text('J2') === '6', 'the transposed paste', 5000).catch(() => {});
    check('sheets: Paste Special → Transpose pastes the rows as columns', text('H1') === 'Item' && text('I1') === 'Ink' && text('J1') === 'Paper' && text('H2') === 'Qty' && text('J2') === '6', ['H1', 'I1', 'J1', 'H2', 'J2'].map(text).join(' '));

    // Two numbers dragged by the fill handle carry on as a series.
    await drag(await spot('A5'), await spot('A6'), 3);
    await until(() => range() === '4,0,5,0', 'A5:A6 selected', 4000).catch(() => {});
    const handle = await centre('.sh-fill');
    const corner = await js(`(() => { const r = document.querySelector('.sh-cell[data-ref="A6"]').getBoundingClientRect(); return { x: Math.round(r.right), y: Math.round(r.bottom) }; })()`);
    const atCorner = handle && Math.abs(handle.x - corner.x) <= 6 && Math.abs(handle.y - corner.y) <= 6;
    await drag(handle, { x: handle.x, y: (await spot('A9')).y }, 8);
    await until(() => text('A9') !== '', 'the fill', 5000).catch(() => {});
    check('sheets: the fill handle sits at the selection\'s corner, and two numbers dragged down carry on as a series', atCorner && ['A7', 'A8', 'A9'].map(text).join(',') === '3,4,5', `handle ${JSON.stringify(handle)} corner ${JSON.stringify(corner)}; ${['A7', 'A8', 'A9'].map(text).join(',')}`);

    // Auto Fill Options → Copy Cells.
    await until(() => js(`Boolean(document.querySelector('.sh-autofill'))`), 'Auto Fill Options', 4000).catch(() => {});
    const options = await js(`Boolean(document.querySelector('.sh-autofill'))`);
    if (options) {
      await js(`(() => { document.querySelector('.sh-autofill').click(); return 1; })()`);
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === 'Copy Cells'))`), 'the options', 3000).catch(() => {});
      await js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === 'Copy Cells')?.click(); return 1; })()`);
      await until(() => text('A7') === '1', 'the copy', 5000).catch(() => {});
    }
    check('sheets: Auto Fill Options → Copy Cells turns the series into a copy of the two cells', options && ['A7', 'A8', 'A9'].map(text).join(',') === '1,2,1', `${options}; ${['A7', 'A8', 'A9'].map(text).join(',')}`);

    // A drag across the column headings selects whole columns.
    const headB = await js(`(() => { const r = [...document.querySelectorAll('.sh-colheads .sh-head')].find((h) => h.textContent.trim() === 'B').getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) }; })()`);
    const headD = await js(`(() => { const r = [...document.querySelectorAll('.sh-colheads .sh-head')].find((h) => h.textContent.trim() === 'D').getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) }; })()`);
    await drag(headB, headD);
    await until(() => model().selection.left === 1 && model().selection.right === 3 && model().selection.top === 0, 'columns B to D', 4000).catch(() => {});
    const s = model().selection;
    check('sheets: a drag across the column headings from B to D selects the three columns', s.left === 1 && s.right === 3 && s.top === 0 && s.bottom >= 8, range());

    // Home → Fill → Series: from 10 in C8, along the row in fives to 30.
    await click(await spot('C8'));
    await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Fill'); b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); b.click(); return 1; })()`);
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === 'Series…'))`), 'Fill → Series', 3000);
    await js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === 'Series…').click(); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.sh-series-ok'))`), 'the Series dialog', 3000);
    await js(`(() => {
      const set = (input, value) => { Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(input, value); input.dispatchEvent(new Event('input', { bubbles: true })); };
      document.querySelector('.sh-series-rows').click();
      set(document.querySelector('input.sh-series-step'), '5');
      set(document.querySelector('input.sh-series-stop'), '30');
      return 1;
    })()`);
    await wait(150);
    await js(`(() => { document.querySelector('.sh-series-ok').click(); return 1; })()`);
    await until(() => text('G8') === '30', 'the series', 5000).catch(() => {});
    check('sheets: Fill → Series runs from 10 along the row in fives to the stop at 30', ['D8', 'E8', 'F8', 'G8', 'H8'].map(text).join(',') === '15,20,25,30,', ['D8', 'E8', 'F8', 'G8', 'H8'].map(text).join(','));

    // Ctrl+X, then Ctrl+V: the row moves.
    await drag(await spot('A2'), await spot('B2'), 3);
    await until(() => range() === '1,0,1,1', 'A2:B2 selected', 4000).catch(() => {});
    await press(wc, 'X', { modifiers: ['control'] });
    await wait(200);
    await click(await spot('A12'));
    await press(wc, 'V', { modifiers: ['control'] });
    await until(() => text('A12') === 'Ink', 'the move', 5000).catch(() => {});
    check('sheets: Ctrl+X then Ctrl+V moves the cells and leaves their old place empty', text('A12') === 'Ink' && text('B12') === '4' && text('A2') === '' && text('B2') === '', ['A12', 'B12', 'A2', 'B2'].map(text).join(' '));

    const complaints = await errorsIn(win);
    check('sheets: selecting, copying, pasting and filling report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('sheets: the select-and-paste checks ran', false, err.message);
  }
}
