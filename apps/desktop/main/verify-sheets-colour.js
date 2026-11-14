// Worksheets: Font colour and Fill colour, as Excel's: split buttons whose
// arrow opens the palette (the workbook theme's colours with their steps,
// the standard colours, Automatic or No fill, More colours) and whose face
// puts on the colour last chosen. Checked on one cell, a range, a whole row
// and a whole column pressed on their headers (where a cell typed later
// takes the colour too: before, only the cells in use were coloured), and on
// words chosen in a cell being typed in, pressed with the real mouse so the
// editor keeps the keyboard; then the saved file carries the row's and the
// column's style as Excel writes them and opens with them.
// Run alone with RUTBA_VERIFY_ONLY=sheets-colour.

import fs from 'node:fs';
import path from 'node:path';
import { OoxmlPackage } from '@rutba/ooxml';
import { buildXlsx } from '@rutba/ooxml/build';
import { labelKeyed } from './verify-keys.js';

/**
 * @param {object} h the harness: open, check, until, wait, capture, doc
 */
export async function verifySheetsColour(h, { dir }) {
  const { open, check, until, wait, capture, doc } = h;
  const own = path.join(dir, 'sheets-colour');
  const file = path.join(own, 'colours.xlsx');
  try {
    fs.mkdirSync(own, { recursive: true });
    fs.writeFileSync(file, buildXlsx({ sheets: [{ name: 'Sheet1', rows: [['Name', 'Qty', 'Price'], ['Tea', 2, 3], ['Milk', 1, 4], ['Rice', 5, 2]] }] }));
  } catch (err) {
    return check('sheets-colour: the fixture was made', false, err.message);
  }

  let win = null;
  try {
    win = await open('sheets', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.sh-cells .sh-cell'))`), 'the grid', 8000);
    const session = doc.sessions().filter((s) => s.kind === 'sheet' && s.path === file).pop();
    const model = () => doc.model({ id: session.id });
    const cellOf = (ref) => (model().cells || []).find((c) => c.ref === ref) || null;
    const colourOf = (ref) => (cellOf(ref)?.style?.font?.colour || '').toUpperCase().replace(/^#/, '') || null;
    const painted = (ref) => js(`(() => {
      const c = document.querySelector('.sh-cell[data-ref="${ref}"]');
      if (!c) return null;
      const all = [c, ...c.querySelectorAll('*')].filter((n) => [...n.childNodes].some((x) => x.nodeType === 3 && x.textContent.trim()));
      return all.length ? getComputedStyle(all[all.length - 1]).color : null;
    })()`);
    const rgb = (hex) => `rgb(${parseInt(hex.slice(0, 2), 16)}, ${parseInt(hex.slice(2, 4), 16)}, ${parseInt(hex.slice(4, 6), 16)})`;

    // The real mouse, as a person presses: at an element's centre, or at a cell's place on the grid.
    const at = (selector) => js(`(() => { const n = document.querySelector(${JSON.stringify(selector)}); if (!n) return null; const r = n.getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) }; })()`);
    const press = async (point, modifiers = []) => {
      if (!point) return 'nowhere';
      wc.sendInputEvent({ type: 'mouseDown', x: point.x, y: point.y, button: 'left', clickCount: 1, modifiers });
      wc.sendInputEvent({ type: 'mouseUp', x: point.x, y: point.y, button: 'left', clickCount: 1, modifiers });
      await wait(150);
      return 'pressed';
    };
    const cellPoint = async (ref) => {
      const m = /^([A-Z]+)([0-9]+)$/.exec(ref);
      const col = m[1].split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
      const row = Number(m[2]) - 1;
      const c = (model().columns || []).find((x) => x.index === col);
      const r = (model().rows || []).find((x) => x.index === row);
      const layer = await js(`(() => { const b = document.querySelector('.sh-cells')?.getBoundingClientRect(); return b ? { left: b.left, top: b.top } : null; })()`);
      return c && r && layer ? { x: Math.round(layer.left + c.x + c.width / 2), y: Math.round(layer.top + r.y + r.height / 2) } : null;
    };
    const clickCell = async (ref, modifiers = []) => press(await cellPoint(ref), modifiers);
    const head = (axis, label) => js(`(() => {
      const h = ((list) => list.find((n) => ${labelKeyed(label)}(n)) || list.find((n) => n.textContent.trim() === ${JSON.stringify(label)}))([...document.querySelectorAll('.sh-${axis}heads .sh-head')]);
      if (!h) return null; const r = h.getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) };
    })()`);
    const type = async (text) => {
      for (const ch of text) {
        wc.sendInputEvent({ type: 'keyDown', keyCode: ch });
        wc.sendInputEvent({ type: 'char', keyCode: ch });
        wc.sendInputEvent({ type: 'keyUp', keyCode: ch });
        await wait(50);
      }
    };
    const enter = async () => { wc.sendInputEvent({ type: 'keyDown', keyCode: 'Return' }); wc.sendInputEvent({ type: 'keyUp', keyCode: 'Return' }); await wait(250); };
    const palette = () => js(`(() => {
      const box = document.querySelector('.rw-colours');
      if (!box) return null;
      const cells = (sel) => [...box.querySelectorAll(sel)].map((b) => b.dataset.colour);
      return { base: cells('.rw-cell.base'), steps: cells('.rw-cell[class*="step"]'), std: cells('.rw-cell.std'), rows: [...box.querySelectorAll('.rw-colours-row')].map((b) => b.textContent.trim()) };
    })()`);
    const arrow = async (name) => {
      const r = await press(await at(`.rw-split[data-kind="${name}"] .rw-split-arrow`));
      await until(async () => Boolean(await palette()), 'the palette', 3000).catch(() => {});
      return r;
    };
    const pick = async (hex) => press(await at(`.rw-colours .rw-cell[data-value="${hex}"]`));

    // 1. The palettes.
    await clickCell('A2');
    await arrow('sheets.font');
    const font = await palette();
    await capture(win, 'sheets-colour-palette.png');
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'Escape' });
    wc.sendInputEvent({ type: 'keyUp', keyCode: 'Escape' });
    await until(async () => !(await palette()), 'the palette to close', 2000).catch(() => {});
    await arrow('sheets.fill');
    const fill = await palette();
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'Escape' });
    wc.sendInputEvent({ type: 'keyUp', keyCode: 'Escape' });
    await until(async () => !(await palette()), 'the palette to close', 2000).catch(() => {});
    check('sheets-colour: Font colour and Fill colour open the palette: the theme\'s ten colours with five steps each, ten standard colours, Automatic or No fill, and More colours',
      font?.base.length === 10 && font.steps.length === 50 && font.std.length === 10 && font.rows[0] === 'Automatic' && font.rows.includes('More colours…')
        && fill?.base.length === 10 && fill.rows[0] === 'No fill',
      JSON.stringify({ font, fill }));

    // 2. One cell.
    await clickCell('A2');
    await arrow('sheets.font');
    await pick('C00000');
    await until(() => colourOf('A2') === 'C00000', 'A2 in dark red', 4000).catch(() => {});
    check('sheets-colour: Font colour colours one cell, and the grid paints it', colourOf('A2') === 'C00000' && (await painted('A2')) === rgb('C00000'),
      JSON.stringify({ engine: colourOf('A2'), painted: await painted('A2') }));

    // 3. A range, A3 to C3.
    await clickCell('A3');
    await clickCell('C3', ['shift']);
    await arrow('sheets.font');
    await pick('0070C0');
    await until(() => ['A3', 'B3', 'C3'].every((r) => colourOf(r) === '0070C0'), 'the range in blue', 4000).catch(() => {});
    check('sheets-colour: Font colour colours every cell of a range', ['A3', 'B3', 'C3'].every((r) => colourOf(r) === '0070C0') && colourOf('A4') !== '0070C0',
      JSON.stringify(['A3', 'B3', 'C3', 'A4'].map((r) => [r, colourOf(r)])));

    // 4. A whole row, pressed on its header; then a cell typed past the cells in use.
    await press(await head('row', '4'));
    await arrow('sheets.font');
    await pick('00B050');
    await until(() => colourOf('A4') === '00B050', 'row 4 in green', 4000).catch(() => {});
    await clickCell('F4');
    await type('later');
    await enter();
    await until(() => cellOf('F4')?.text === 'later', 'the words typed in F4', 4000).catch(() => {});
    check('sheets-colour: a whole row coloured from its header colours its cells, and words typed past them later take the colour too',
      ['A4', 'B4', 'C4'].every((r) => colourOf(r) === '00B050') && colourOf('F4') === '00B050' && (await painted('F4')) === rgb('00B050'),
      JSON.stringify({ row: ['A4', 'B4', 'C4', 'F4'].map((r) => [r, colourOf(r)]), painted: await painted('F4') }));

    // 5. A whole column, below the rows in use.
    await press(await head('col', 'E'));
    await arrow('sheets.font');
    await pick('7030A0');
    await wait(300);
    await clickCell('E9');
    await type('below');
    await enter();
    await until(() => cellOf('E9')?.text === 'below', 'the words typed in E9', 4000).catch(() => {});
    check('sheets-colour: a whole column coloured from its header gives the colour to words typed in it below the rows in use',
      colourOf('E9') === '7030A0' && (await painted('E9')) === rgb('7030A0'), JSON.stringify({ e9: colourOf('E9'), painted: await painted('E9') }));

    // 6. Words chosen in a cell being typed in, the palette pressed with the real mouse.
    const a1 = await cellPoint('A1');
    wc.sendInputEvent({ type: 'mouseDown', x: a1.x, y: a1.y, button: 'left', clickCount: 2 });
    wc.sendInputEvent({ type: 'mouseUp', x: a1.x, y: a1.y, button: 'left', clickCount: 2 });
    await until(() => js(`Boolean(document.querySelector('input.sh-editor'))`), 'the cell editor', 4000).catch(() => {});
    await js(`(() => { const i = document.querySelector('input.sh-editor'); i.focus(); i.setSelectionRange(0, 2); return 1; })()`);
    await arrow('sheets.font');
    const stillEditing = await js(`Boolean(document.querySelector('input.sh-editor, .sh-rich'))`);
    await pick('FF0000');
    await until(() => js(`Boolean(document.querySelector('.sh-rich'))`), 'the rich editor', 3000).catch(() => {});
    const shown = await js(`[...document.querySelectorAll('.sh-rich span')].map((s) => [s.textContent, s.style.color])`);
    await enter();
    // The model's runs name it `colour` (the file's runs, `color`).
    const runs = () => (cellOf('A1')?.runs || []).map((r) => [r.text, String(r.colour || r.color || '').toUpperCase().replace(/^#/, '').slice(-6) || null]);
    await until(() => runs().some(([t, c]) => t === 'Na' && c === 'FF0000'), 'the runs in A1', 4000).catch(() => {});
    check('sheets-colour: words chosen in a cell being typed in take the colour from the palette, the editor keeping the keyboard, and Enter puts them in coloured',
      stillEditing === true && JSON.stringify(runs()) === JSON.stringify([['Na', 'FF0000'], ['me', null]]),
      JSON.stringify({ stillEditing, shown, runs: runs() }));

    // 7. The face puts the last colour on again.
    await clickCell('B2');
    await press(await at('.rw-split[data-kind="sheets.font"] .rw-split-face'));
    await until(() => colourOf('B2') === 'FF0000', 'B2 in red', 4000).catch(() => {});
    check('sheets-colour: the button\'s face puts on the colour last chosen', colourOf('B2') === 'FF0000', JSON.stringify({ b2: colourOf('B2') }));

    // 8. Saved and opened again.
    await doc.save({ id: session.id, path: file });
    const xml = OoxmlPackage.read(fs.readFileSync(file)).text('xl/worksheets/sheet1.xml');
    const rowTag = /<row\b[^>]*\br="4"[^>]*>/.exec(xml)?.[0] || '';
    const colTag = /<col\b[^>]*\bmin="5"[^>]*\/>/.exec(xml)?.[0] || '';
    const again = await doc.open({ path: file });
    const back = doc.model({ id: again.id });
    const cell2 = (ref) => (back.cells || []).find((c) => c.ref === ref);
    const backColour = (ref) => (cell2(ref)?.style?.font?.colour || '').toUpperCase().replace(/^#/, '') || null;
    const reopened = { A2: backColour('A2'), C3: backColour('C3'), F4: backColour('F4'), E9: backColour('E9') };
    doc.close({ id: again.id });
    check('sheets-colour: the saved file keeps the row\'s and the column\'s colour as Excel writes them, and opens with every colour',
      /\bcustomFormat="1"/.test(rowTag) && /\bs="\d+"/.test(rowTag) && /\bstyle="\d+"/.test(colTag) && /\bwidth="/.test(colTag)
        && reopened.A2 === 'C00000' && reopened.C3 === '0070C0' && reopened.F4 === '00B050' && reopened.E9 === '7030A0',
      JSON.stringify({ rowTag, colTag, reopened }));
  } catch (err) {
    check('sheets-colour: the checks ran', false, err.stack || err.message);
  } finally {
    try { win?.close(); } catch { /* gone */ }
  }
}
