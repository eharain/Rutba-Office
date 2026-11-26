// Worksheets: the faults put right for 2.28.0, each one seen in a window.
//
// The rows were proved on the engine by unit tests and handed over with no
// window started. Here a person's steps are taken for each: an entry typed
// into a cell, a press on the ribbon, a pick from a menu, a drag over cells,
// a dialog filled in. What the row promised is then read from the document
// and from the page.
//
//   B40  a date, a time, a percentage, an amount and a grouped number typed
//        into a cell are numbers in those formats; a Text cell keeps 0123
//   B26  a press on an empty cell that a neighbour's words lie over selects it
//   B49  Format Painter paints on the release of the mouse and goes back up;
//        pressed twice it keeps painting until Escape
//   B57  Cell Styles gives a cell a named style, and offers to update the style
//   B27  a table just made is known to a formula at once
//   B33  Outside Borders draws the box alone
//   B51  a Data Table fills from the dialog
//   B52  Goal Seek says the value it found
//   B45  Ctrl+D repeats the first row of the selection
//   B54  a formula that gives a date shows a date
//   B60  IPMT, PPMT, CUMIPMT and CUMPRINC
//   B30  a list rule shows its arrow, opens on Alt+Down, fills the cell from a
//        pick, and refuses an entry that is not on the list (B41, B47)
//   B34  Wrap Text grows the row
//   B43  the Watch Window sits under the grid and reads values as the cells do
//
// Run alone with RUTBA_VERIFY_ONLY=sheet228.
import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';
import { readsDayFirst } from '@rutba/formula';
import { keySelector, labelKeyed } from './verify-keys.js';
import { validationSpec } from '../renderer/apps/sheets/rule-specs.js';

/** How many styles the Cell Styles menu names: Normal, the headings, Title, Total and the rest. */
const CELL_STYLES = 17;

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where the workbook is written
 */
export async function verifySheet228(h, { dir }) {
  const { open, check, until, wait, press, errorsIn, doc, sessionFor } = h;
  const file = path.join(dir, 'sheet-228.xlsx');
  const LONG = 'These words are far too long for one narrow cell';
  // Rows as the sheet numbers them, from 1. A null leaves a cell empty.
  const rows = [];
  const put = (ref, value) => {
    const m = /^([A-Z])([0-9]+)$/.exec(ref);
    const r = Number(m[2]) - 1;
    const c = m[1].charCodeAt(0) - 65;
    while (rows.length <= r) rows.push([]);
    while (rows[r].length <= c) rows[r].push(null);
    rows[r][c] = value;
  };
  put('E1', 'Heading'); put('E2', 'Second'); put('F1', LONG); put('H1', 'Look');
  put('A7', 'A sentence long enough to need three or four lines in a narrow column');
  put('A9', 'Item'); put('B9', 'Qty'); put('C9', 'Price'); put('D9', 'Total');
  put('A10', 'Tea'); put('B10', 2); put('C10', 3);
  put('A11', 'Coffee'); put('B11', 5); put('C11', 4);
  put('A12', 'Milk'); put('B12', 1); put('C12', 2);
  put('B14', 1000); put('B15', 0.1); put('D15', 0.05); put('D16', 0.2); put('D17', 0.5);
  put('H14', 'Mon'); put('I14', 'Item 1');
  fs.writeFileSync(file, buildXlsx({
    sheets: [
      { name: 'Sheet1', rows },
      { name: 'Products', rows: [['Product'], ['Nails'], ['Screws'], ['Bolts']] },
    ],
  }));
  let win = null;
  try {
    win = await open('sheets', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const session = sessionFor('sheet');
    const model = () => doc.model({ id: session.id });
    const cell = (ref) => model().cells?.find((c) => c.ref === ref) ?? null;
    const textOf = (ref) => cell(ref)?.text ?? null;
    const activeRef = () => model().selection?.active?.ref ?? null;
    const place = (ref) => {
      const m = /^([A-Z]+)([0-9]+)$/.exec(ref);
      return { col: m[1].split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1, row: Number(m[2]) - 1 };
    };
    // Where a cell is on the screen, from the grid's own columns and rows: an empty cell is not drawn as an element.
    const centre = async (ref) => {
      const { row, col } = place(ref);
      const now = model();
      const c = (now.columns || []).find((x) => x.index === col);
      const r = (now.rows || []).find((x) => x.index === row);
      if (!c || !r) return null;
      const layer = await js(`(() => { const l = document.querySelector('.sh-cells'); if (!l) return null; const b = l.getBoundingClientRect(); return { left: b.left, top: b.top }; })()`);
      return layer ? { x: Math.round(layer.left + c.x + c.width / 2), y: Math.round(layer.top + r.y + r.height / 2) } : null;
    };
    const mouse = async (type, ref, extra = {}) => {
      const at = await centre(ref);
      if (!at) return false;
      wc.sendInputEvent({ type, x: at.x, y: at.y, button: 'left', clickCount: 1, ...extra });
      return true;
    };
    const click = async (ref, modifiers = []) => {
      const down = await mouse('mouseDown', ref, { modifiers });
      await mouse('mouseUp', ref, { modifiers });
      await wait(180);
      return down;
    };
    const drag = async (from, to) => {
      await mouse('mouseDown', from);
      await wait(80);
      await mouse('mouseMove', to);
      await wait(120);
      await mouse('mouseUp', to);
      await wait(300);
    };
    const type = async (text) => {
      for (const ch of text) {
        const modifiers = /[A-Z]/.test(ch) ? ['shift'] : [];
        wc.sendInputEvent({ type: 'keyDown', keyCode: ch, modifiers });
        wc.sendInputEvent({ type: 'char', keyCode: ch, modifiers });
        wc.sendInputEvent({ type: 'keyUp', keyCode: ch, modifiers });
        await wait(60);
      }
      await wait(150);
    };
    const editing = () => js(`Boolean(document.querySelector('.sh-editor'))`);
    // An entry made in the cell's own editor: the cell picked, F2, the words put in the editor as typing puts them, Enter.
    const enter = async (ref, text) => {
      await click(ref);
      await press(wc, 'F2');
      await until(editing, `the editor on ${ref}`, 5000).catch(() => {});
      const set = await js(`(() => {
        const e = document.querySelector('input.sh-editor');
        if (!e) return false;
        e.focus();
        Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(e, ${JSON.stringify(text)});
        e.dispatchEvent(new Event('input', { bubbles: true }));
        return true;
      })()`);
      await wait(150);
      await press(wc, 'Return');
      await until(async () => !(await editing()), `the entry in ${ref} to be taken`, 5000).catch(() => {});
      await wait(150);
      return set;
    };
    // A control by its label's key, whatever words the window draws for it, else by its words.
    const found = (english, { tip = false, within = '' } = {}) => {
      const selector = keySelector(english, { tip });
      const words = JSON.stringify(english);
      return `((root) => {
        if (!root) return null;
        const ok = (n) => n && !n.disabled && n.getClientRects().length > 0;
        const keyed = ${selector ? `[...root.querySelectorAll(${JSON.stringify(selector)})].map((n) => (n.matches('button') ? n : n.closest('button'))).find(ok)` : 'null'};
        return keyed || [...root.querySelectorAll('button')].find((n) => ok(n) && (n.textContent.trim() === ${words}${tip ? ` || n.title === ${words} || n.dataset.tip === ${words}` : ''})) || null;
      })(${within ? `document.querySelector(${JSON.stringify(within)})` : 'document'})`;
    };
    const pressOn = async (english, o = {}) => {
      await until(() => js(`Boolean(${found(english, o)})`), `"${english}" to be there`, 5000).catch(() => {});
      const did = await js(`(() => { const b = ${found(english, o)}; if (!b) return false; b.click(); return true; })()`);
      await wait(250);
      return did;
    };
    const tab = async (name) => {
      await js(`(() => { ((list) => list.find((t) => ${labelKeyed(name)}(t)) || list.find((t) => t.textContent.trim() === ${JSON.stringify(name)}))([...document.querySelectorAll('.rw-tab')])?.click(); return 1; })()`);
      await wait(300);
    };
    const menuWords = () => js(`[...document.querySelectorAll('.rw-menu button')].map((b) => b.textContent.trim())`);
    const pickWords = async (words) => {
      await until(async () => (await menuWords()).includes(words), `"${words}" in the menu`, 5000).catch(() => {});
      const did = await js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(words)}); if (!b) return false; b.click(); return true; })()`);
      await wait(250);
      return did;
    };
    const fields = (values) => js(`(() => {
      const inputs = [...document.querySelectorAll('.rw-dialog input')];
      const set = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
      const values = ${JSON.stringify(values)};
      for (const [i, v] of Object.entries(values)) {
        const e = inputs[Number(i)];
        if (!e) return false;
        set.call(e, v);
        e.dispatchEvent(new Event('input', { bubbles: true }));
      }
      return inputs.length;
    })()`);
    const apply = (...ops) => doc.apply({ id: session.id, ops });
    const select = (ref) => ({ op: 'select', ...place(ref) });
    await until(() => js(`Boolean(document.querySelector('.sh-cells .sh-cell'))`), 'the grid', 8000);

    // B40: what is typed is read as a person means it.
    const dayFirst = readsDayFirst();
    const DATE = dayFirst ? '30/11/2026' : '11/30/2026';
    await enter('A1', DATE);
    await enter('A2', '14:30');
    await enter('A3', '12.5%');
    await enter('A4', '£1,250.50');
    await enter('A5', '1,250');
    await apply(select('A6'), { op: 'setFormat', delta: { numberFormat: '@' } });
    await enter('A6', '0123');
    await enter('C1', '=AND(ISNUMBER(A1),ISNUMBER(A2),ISNUMBER(A3),ISNUMBER(A4),ISNUMBER(A5))');
    await enter('C2', '=ISTEXT(A6)');
    const typed = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6'].map((ref) => ({ ref, text: textOf(ref), align: cell(ref)?.align ?? null }));
    const [d, t, p, m, n, x] = typed;
    check('sheets: a date, a time, a percentage, an amount of money and a grouped number typed into cells are numbers shown in those formats, on the right',
      /^\d{1,2}\/\d{1,2}\/2026$/.test(d.text || '') && t.text === '14:30' && /^12\.50?%$/.test(p.text || '') && m.text === '£1,250.50' && n.text === '1,250'
        && [d, t, p, m, n].every((c) => c.align === 'right') && textOf('C1') === 'TRUE',
      `${JSON.stringify(typed.slice(0, 5))}; all five numbers ${JSON.stringify(textOf('C1'))}; dates read day first ${dayFirst}`);
    check('sheets: a cell set to Text keeps 0123 as typed', x.text === '0123' && x.align !== 'right' && textOf('C2') === 'TRUE',
      `${JSON.stringify(x)}; is text ${JSON.stringify(textOf('C2'))}`);

    // B26: G1 is empty and F1's words lie over it. A press there is a press on G1.
    await click('A1');
    const pressedG1 = await click('G1');
    const after = activeRef();
    await type('ok');
    await press(wc, 'Return');
    await until(() => textOf('G1') === 'ok', 'G1 to take the words', 4000).catch(() => {});
    check('sheets: a press on an empty cell that the next cell\'s long words lie over selects that cell, and typing goes into it',
      pressedG1 && after === 'G1' && textOf('G1') === 'ok' && textOf('F1') === LONG,
      `pressed ${pressedG1}; selected ${JSON.stringify(after)}; G1 ${JSON.stringify(textOf('G1'))}; F1 ${JSON.stringify(textOf('F1'))}`);

    // B49: Format Painter.
    await apply(select('H1'), { op: 'setFormat', delta: { bold: true, fill: '#FFE699' } });
    await click('A1');
    await click('H1');
    await tab('Home');
    const looks = (ref) => { const s = cell(ref)?.style; return { bold: Boolean(s?.font?.bold), fill: s?.fill ?? null }; };
    const same = (a, b) => a.bold && b.bold && a.fill != null && JSON.stringify(a.fill) === JSON.stringify(b.fill);
    const source = looks('H1');
    const brushOn = await pressOn('Format Painter');
    await until(() => Boolean(model().brush), 'the brush to be picked up', 4000).catch(() => {});
    const held = { brush: model().brush ?? null, pressed: await js(`Boolean(${found('Format Painter')}?.matches('.pressed, [aria-pressed="true"]'))`), cursor: await js(`Boolean(document.querySelector('.sh-cells.brushing'))`) };
    await mouse('mouseDown', 'H3');
    await wait(80);
    await mouse('mouseMove', 'I4');
    await wait(200);
    const before = looks('H3');
    await mouse('mouseUp', 'I4');
    await until(() => !model().brush, 'the brush to be put down', 4000).catch(() => {});
    await wait(200);
    const painted = ['H3', 'I3', 'H4', 'I4'].map(looks);
    check('sheets: Format Painter paints the cells dragged over when the mouse is let go, and the button goes back up',
      brushOn && Boolean(held.brush) && held.cursor && !before.bold && painted.every((c) => same(source, c)) && !model().brush,
      `source ${JSON.stringify(source)}; held ${JSON.stringify(held)}; before release ${JSON.stringify(before)}; painted ${JSON.stringify(painted)}; brush after ${JSON.stringify(model().brush ?? null)}`);
    // Pressed twice it keeps painting. The button is pressed as a person presses it twice: two presses, the second a double one.
    await click('H1');
    const spot = await js(`(() => { const r = ${found('Format Painter')}?.getBoundingClientRect(); return r ? { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) } : null; })()`);
    if (spot) {
      for (const clickCount of [1, 2]) {
        wc.sendInputEvent({ type: 'mouseDown', x: spot.x, y: spot.y, button: 'left', clickCount });
        wc.sendInputEvent({ type: 'mouseUp', x: spot.x, y: spot.y, button: 'left', clickCount });
        await wait(120);
      }
    }
    await until(() => Boolean(model().brush?.keep), 'the brush to be kept', 4000).catch(() => {});
    const kept = model().brush ?? null;
    await click('J3');
    await click('J5');
    const stillHeld = model().brush ?? null;
    const twice = ['J3', 'J5'].map(looks);
    await press(wc, 'Escape');
    await until(() => !model().brush, 'Escape to put the brush down', 4000).catch(() => {});
    check('sheets: Format Painter pressed twice paints each selection in turn until Escape',
      Boolean(kept?.keep) && Boolean(stillHeld) && twice.every((c) => same(source, c)) && !model().brush,
      `kept ${JSON.stringify(kept)}; after two paints ${JSON.stringify(stillHeld)}; painted ${JSON.stringify(twice)}; after Escape ${JSON.stringify(model().brush ?? null)}`);

    // B57: Cell Styles.
    await click('E1');
    await pressOn('Cell Styles');
    const styleNamed = await pressOn('Heading 1', { within: '.rw-menu' });
    await until(() => model().format?.style === 'Heading 1', 'the style to be on the cell', 4000).catch(() => {});
    const styled = { style: model().format?.style ?? null, look: cell('E1')?.style?.font ?? null };
    await click('E2');
    await pressOn('Cell Styles');
    await pressOn('Heading 1', { within: '.rw-menu' });
    await until(() => model().format?.style === 'Heading 1', 'the style to be on the second cell', 4000).catch(() => {});
    await apply(select('E2'), { op: 'setFormat', delta: { fill: '#C6E0B4' } });
    await click('E1');
    await click('E2');
    await pressOn('Cell Styles');
    const wears = model().format?.style ?? null;
    // The offer comes after the names of the styles. It is found by its place: its words are the label's own once the English copy has them.
    const update = `((list) => list.find((n) => n.matches('[data-key="update-style-match"]')) || list[${CELL_STYLES}] || null)([...document.querySelectorAll('.rw-menu button')])`;
    await until(() => js(`Boolean(${update})`), 'the offer to update the style', 4000).catch(() => {});
    const inMenu = await menuWords();
    const offered = await js(`(() => { const b = ${update}; if (!b) return null; const words = b.textContent.trim(); b.click(); return words; })()`);
    await until(() => cell('E1')?.style?.fill != null, 'the style to be updated', 4000).catch(() => {});
    const fills = { E1: cell('E1')?.style?.fill ?? null, E2: cell('E2')?.style?.fill ?? null };
    check('sheets: Cell Styles gives a cell the named style, and a cell changed by hand can update the style for every cell in it',
      styleNamed && styled.style === 'Heading 1' && wears === 'Heading 1' && inMenu.length === CELL_STYLES + 1 && Boolean(offered) && fills.E1 != null && JSON.stringify(fills.E1) === JSON.stringify(fills.E2),
      `picked ${styleNamed}; E1 ${JSON.stringify(styled)}; E2 with a fill still wears ${JSON.stringify(wears)}; the menu had ${inMenu.length} choices; offered ${JSON.stringify(offered)}; fills ${JSON.stringify(fills)}`);

    // B27: a table made a moment ago is known to a formula typed next.
    await click('A9');
    await click('D12', ['shift']);
    await pressOn('Format as Table');
    const tableMade = await js(`(() => { const b = document.querySelectorAll('.rw-menu button')[1]; if (!b) return false; b.click(); return true; })()`);
    await wait(500);
    await enter('D10', '=[@Qty]*[@Price]');
    await enter('F10', '=SUM(Table1[Qty])');
    check('sheets: a table just made is known at once to a formula in it and to one outside it',
      tableMade && textOf('D10') === '6' && textOf('F10') === '8',
      `table made ${tableMade}; D10 ${JSON.stringify(textOf('D10'))}; F10 ${JSON.stringify(textOf('F10'))}`);

    // B33: Outside Borders is a box and nothing inside it.
    const edges = (ref) => { const b = cell(ref)?.style?.border || {}; return ['top', 'right', 'bottom', 'left'].filter((k) => Boolean(b[k])).join(' ') || 'none'; };
    await click('H9');
    await click('J11', ['shift']);
    await pressOn('Borders', { tip: true });
    const boxed = await pressOn('Outside borders', { within: '.rw-menu' });
    await until(() => edges('H9') !== 'none', 'the box', 4000).catch(() => {});
    const box = Object.fromEntries(['H9', 'I9', 'J9', 'H10', 'I10', 'J10', 'H11', 'I11', 'J11'].map((ref) => [ref, edges(ref)]));
    check('sheets: Outside Borders draws a box round the selection and no lines inside it',
      boxed && box.H9 === 'top left' && box.I9 === 'top' && box.J9 === 'top right' && box.H10 === 'left' && box.I10 === 'none' && box.J10 === 'right'
        && box.H11 === 'bottom left' && box.I11 === 'bottom' && box.J11 === 'right bottom',
      `picked ${boxed}; ${JSON.stringify(box)}`);
    // The line's width as the grid draws it: a thin line is 1, a medium one 2.
    const widthOf = (ref, edge) => cell(ref)?.style?.border?.[edge]?.widthPx ?? null;
    const thin = widthOf('H9', 'top');
    await pressOn('Borders', { tip: true });
    const thick = await pressOn('Thick outside borders', { within: '.rw-menu' });
    await until(() => widthOf('H9', 'top') > thin, 'the thicker box', 4000).catch(() => {});
    const thicker = { H9: [widthOf('H9', 'top'), widthOf('H9', 'left')], J11: [widthOf('J11', 'bottom'), widthOf('J11', 'right')], inside: edges('I10') };
    check('sheets: Thick outside borders draws the same box with a thicker line',
      thick && thin > 0 && [...thicker.H9, ...thicker.J11].every((width) => width > thin) && thicker.inside === 'none',
      `picked ${thick}; the thin line ${thin}; ${JSON.stringify(thicker)}`);

    // B51: a Data Table from the dialog. B16 is B14 times B15; E14 shows it; D15 to D17 are the rates to try.
    await enter('B16', '=B14*B15');
    await enter('E14', '=B16');
    await tab('Data');
    await click('D14');
    await click('E17', ['shift']);
    await pressOn('What-If Analysis');
    const tableAsked = await pressOn('Data Table…', { within: '.rw-menu' });
    await until(() => js(`document.querySelectorAll('.rw-dialog input').length >= 3`), 'the Data Table dialog', 5000).catch(() => {});
    await fields({ 2: 'B15' });
    await wait(200);
    const filled = await pressOn('Fill', { within: '.rw-dialog' });
    await until(() => textOf('E17') === '500', 'the table to fill', 6000).catch(() => {});
    check('sheets: Data Table fills the table from the dialog',
      tableAsked && filled && textOf('E15') === '50' && textOf('E16') === '200' && textOf('E17') === '500',
      `asked ${tableAsked}; Fill pressed ${filled}; E15 to E17 ${JSON.stringify([textOf('E15'), textOf('E16'), textOf('E17')])}`);
    if (await js(`Boolean(document.querySelector('.rw-dialog'))`)) { await press(wc, 'Escape'); await wait(300); }

    // B52: Goal Seek says what it found.
    await click('B16');
    await pressOn('What-If Analysis');
    const seekAsked = await pressOn('Goal Seek…', { within: '.rw-menu' });
    await until(() => js(`document.querySelectorAll('.rw-dialog input').length >= 3`), 'the Goal Seek dialog', 5000).catch(() => {});
    await fields({ 0: 'B16', 1: '42', 2: 'B14' });
    await wait(200);
    await pressOn('Find', { within: '.rw-dialog' });
    await until(() => js(`Boolean(document.querySelector('.rw-dialog .ml-result'))`), 'what Goal Seek says', 8000).catch(() => {});
    const said = await js(`(() => { const r = document.querySelector('.rw-dialog .ml-result'); return r ? { good: r.classList.contains('good'), words: r.textContent.trim() } : null; })()`);
    check('sheets: Goal Seek says the value it found, and the cell holds it',
      seekAsked && Boolean(said?.good) && /420/.test(said.words) && textOf('B14') === '420' && textOf('B16') === '42',
      `asked ${seekAsked}; said ${JSON.stringify(said)}; B14 ${JSON.stringify(textOf('B14'))}; B16 ${JSON.stringify(textOf('B16'))}`);
    if (await js(`Boolean(document.querySelector('.rw-dialog'))`)) { await press(wc, 'Escape'); await wait(300); }
    if (await js(`Boolean(document.querySelector('.rw-dialog'))`)) { await pressOn('Close', { within: '.rw-dialog' }); }

    // B45: Ctrl+D repeats the top row. A date, a day's name and a numbered item would each run on as a series if it filled.
    await enter('G14', DATE);
    await click('G14');
    await click('I17', ['shift']);
    await press(wc, 'D', { modifiers: ['control'] });
    await until(() => textOf('I17') === 'Item 1', 'the rows to be repeated', 4000).catch(() => {});
    const top = [textOf('G14'), 'Mon', 'Item 1'];
    const down = [15, 16, 17].map((r) => [textOf(`G${r}`), textOf(`H${r}`), textOf(`I${r}`)]);
    check('sheets: Ctrl+D repeats the first row of the selection in the rows under it, and makes no series of it',
      Boolean(top[0]) && down.every((row) => JSON.stringify(row) === JSON.stringify(top)),
      `first row ${JSON.stringify(top)}; under it ${JSON.stringify(down)}`);

    // B54: a formula that gives a date shows a date.
    await enter('A19', DATE);
    await enter('A20', dayFirst ? '01/12/2026' : '12/01/2026');
    await enter('A21', '=DATE(2026,11,30)');
    check('sheets: a formula that gives a date shows the date, not its number',
      Boolean(textOf('A19')) && textOf('A21') === textOf('A19') && textOf('A21') !== '46356',
      `typed ${JSON.stringify(textOf('A19'))}; the formula ${JSON.stringify(textOf('A21'))}`);

    // B60: the four loan functions.
    const loans = [
      ['C19', '=IPMT(0.1/12,1,36,8000)', '66.67'],
      ['C20', '=PPMT(0.1/12,1,24,2000)', '75.62'],
      ['C21', '=CUMIPMT(0.09/12,360,125000,13,24,0)', '11135.23'],
      ['C22', '=CUMPRINC(0.09/12,360,125000,13,24,0)', '934.11'],
    ];
    for (const [ref, formula] of loans) await enter(ref, formula);
    const lent = loans.map(([ref, formula, want]) => ({ formula, want, text: (textOf(ref) || '').replace(/,/g, '') }));
    check('sheets: IPMT, PPMT, CUMIPMT and CUMPRINC give the interest and the principal of a loan, as money paid out',
      lent.every((l) => Number(l.text) < 0 && Math.abs(Number(l.text)).toFixed(2) === l.want),
      JSON.stringify(lent));

    // B30, B41, B47: a list rule acts in the window.
    await apply(select('F19'), { op: 'validation', spec: validationSpec({ kind: 'list', list: 'Tea, Coffee, Milk', prompt: 'Pick a drink' }) });
    await apply(select('F20'), { op: 'validation', spec: validationSpec({ kind: 'list', list: 'Products!A2:A4' }) });
    await click('E19');
    await click('F19');
    await until(() => js(`Boolean(document.querySelector('.sh-vlist'))`), 'the list arrow', 4000).catch(() => {});
    const arrow = await js(`(() => {
      const a = document.querySelector('.sh-vlist');
      const p = document.querySelector('.sh-vprompt');
      const layer = document.querySelector('.sh-cells').getBoundingClientRect();
      return { arrow: a ? { left: Math.round(a.getBoundingClientRect().left - layer.left), top: Math.round(a.getBoundingClientRect().top - layer.top) } : null, prompt: p ? p.textContent.trim() : null };
    })()`);
    const f19 = (() => { const now = model(); const c = now.columns.find((k) => k.index === 5); const r = now.rows.find((k) => k.index === 18); return c && r ? { right: Math.round(c.x + c.width), top: Math.round(r.y) } : null; })();
    await press(wc, 'Down', { modifiers: ['alt'] });
    await until(async () => (await menuWords()).includes('Coffee'), 'the list to open', 4000).catch(() => {});
    const offer = await menuWords();
    await pickWords('Coffee');
    await until(() => textOf('F19') === 'Coffee', 'the pick to fill the cell', 4000).catch(() => {});
    check('sheets: a cell with a list rule shows an arrow beside it and the rule\'s message, Alt+Down opens the list, and a pick fills the cell',
      Boolean(arrow.arrow) && Boolean(f19) && Math.abs(arrow.arrow.left - f19.right) <= 3 && Math.abs(arrow.arrow.top - f19.top) <= 3 && /Pick a drink/.test(arrow.prompt || '')
        && JSON.stringify(offer) === JSON.stringify(['Tea', 'Coffee', 'Milk']) && textOf('F19') === 'Coffee',
      `arrow and message ${JSON.stringify(arrow)}; the cell's right and top ${JSON.stringify(f19)}; offered ${JSON.stringify(offer)}; F19 ${JSON.stringify(textOf('F19'))}`);
    // An entry that is not on the list: the cell is as it was, and Office says why. The toast is the one error this check asks for.
    await click('E19');
    await click('F19');
    const badBefore = await errorsIn(win);
    await type('Bananas');
    await press(wc, 'Return');
    await until(async () => (await errorsIn(win)).length > badBefore.length, 'Office to say why', 5000).catch(() => {});
    const refusals = (await errorsIn(win)).filter((words) => !badBefore.includes(words));
    if (await editing()) { await press(wc, 'Escape'); await wait(200); }
    check('sheets: an entry that is not on the list is refused: the cell and the selection are as they were, and Office says so once',
      textOf('F19') === 'Coffee' && activeRef() === 'F19' && refusals.length === 1,
      `F19 ${JSON.stringify(textOf('F19'))}; selected ${JSON.stringify(activeRef())}; said ${JSON.stringify(refusals)}`);
    // The words are a label new in 2.28.0: they read so once the English copy carries it.
    check('sheets: what Office says of a refused entry names the cell and the choices of the list',
      /F19/.test(refusals[0] || '') && /Tea, Coffee, Milk/.test(refusals[0] || ''),
      `said ${JSON.stringify(refusals)}`);
    await click('F20');
    await until(() => Boolean(model().validation?.list), 'the list from the other sheet', 4000).catch(() => {});
    const other = model().validation?.list ?? null;
    check('sheets: a list rule that names cells on another sheet offers that sheet\'s values',
      JSON.stringify(other) === JSON.stringify(['Nails', 'Screws', 'Bolts']) && await js(`Boolean(document.querySelector('.sh-vlist'))`),
      `offered ${JSON.stringify(other)}`);

    // B34: Wrap Text grows the row to show every line, and off again it is one line.
    const heightOf = (ref) => model().rows?.find((r) => r.index === place(ref).row)?.height ?? null;
    await tab('Home');
    await click('A7');
    const flat = heightOf('A7');
    await pressOn('Wrap Text');
    await until(() => heightOf('A7') > flat * 1.9, 'the row to grow', 4000).catch(() => {});
    const tall = heightOf('A7');
    const wrapped = Boolean(cell('A7')?.wrap);
    await pressOn('Wrap Text');
    await until(() => heightOf('A7') === flat, 'the row to be one line again', 4000).catch(() => {});
    check('sheets: Wrap Text grows the row to show every line of a long sentence, and turned off the row is one line again',
      wrapped && flat > 0 && tall >= flat * 1.9 && heightOf('A7') === flat && !cell('A7')?.wrap,
      `one line ${flat}; wrapped ${tall} (wrap ${wrapped}); off again ${heightOf('A7')}`);

    // B43: the Watch Window is under the grid, above the sheet tabs, and reads values as the cells show them.
    await tab('Formulas');
    await click('A3');
    await click('A4', ['shift']);
    await pressOn('Watch Window');
    await until(() => js(`Boolean(document.querySelector('.sh-watch'))`), 'the Watch Window', 4000).catch(() => {});
    await pressOn('Add Watch', { within: '.sh-watch' });
    await until(() => js(`document.querySelectorAll('.sh-watch-row:not(.sh-watch-cols) .sh-watch-value').length === 2`), 'the two watched cells', 4000).catch(() => {});
    const watch = await js(`(() => {
      const box = (s) => { const r = document.querySelector(s)?.getBoundingClientRect(); return r ? { top: Math.round(r.top), bottom: Math.round(r.bottom) } : null; };
      return {
        values: [...document.querySelectorAll('.sh-watch-row:not(.sh-watch-cols) .sh-watch-value')].map((n) => n.textContent.trim()),
        pane: box('.sh-watch'), grid: box('.sh-grid'), tabs: box('.sh-tabs'),
      };
    })()`);
    check('sheets: the Watch Window sits under the grid and above the sheet tabs, and its values read as the cells do',
      JSON.stringify(watch.values) === JSON.stringify([textOf('A3'), textOf('A4')]) && /%$/.test(watch.values[0] || '') && /^£/.test(watch.values[1] || '')
        && Boolean(watch.pane && watch.grid && watch.tabs) && watch.grid.bottom <= watch.pane.top + 1 && watch.pane.bottom <= watch.tabs.top + 1,
      JSON.stringify(watch));
    await js(`(() => { const b = [...document.querySelectorAll('.sh-watch-head button')].pop(); b?.click(); return 1; })()`);
    await until(() => js(`!document.querySelector('.sh-watch')`), 'the Watch Window to close', 4000).catch(() => {});

    const errors = (await errorsIn(win)).filter((words) => !refusals.includes(words));
    check('sheets: none of that reported an error, but for the entry the list rule refused', errors.length === 0, JSON.stringify(errors));
  } catch (err) {
    const said = win ? await errorsIn(win).catch(() => []) : [];
    check('sheets: the checks of the 2.28.0 faults ran', false, err.message + (said.length ? '. The window said: ' + said.join(' | ') : ''));
  }
}
