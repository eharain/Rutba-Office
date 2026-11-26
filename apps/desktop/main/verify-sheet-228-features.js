// Worksheets: features added for 2.28.0, each one seen in a window.
//
// The second hand-over of the Worksheets step was proved on the engine by unit
// tests with no window started. Here the steps a person takes are taken:
//
//   B44  a Data Table follows an edit to its model, and Ctrl+Z
//   B59  XLOOKUP gives the whole matching row and spills
//   B55  COUNTIF answers once for each cell of a spill range
//   B31  a sum across sheets, followed on an edit and on a sheet renamed
//   B32  the status bar counts, sums and averages a Ctrl selection
//   B53  the Name Box takes an address, names a selection and goes to a name (B58)
//   B29  a table's formula is carried into a row typed under it
//   B28  a table's total row
//   B42  rows hidden and shown again from their headings; a sheet copied; a tab's colour
//   B38  Remove Duplicates asks which columns, and one Ctrl+Z brings the rows back
//   B37  a number filter and a text filter from a heading's arrow
//
// Not here, and so not yet seen in a window: chart formatting (B39), a date
// filter (B37), the conditional formatting rules (B35, B36), PivotTable
// sorting and Show values as (B48, B56) and sparkline options (B46).
//
// Run alone with RUTBA_VERIFY_ONLY=sheet228b.
import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';
import { keySelector, labelKeyed, shownAs } from './verify-keys.js';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where the workbook is written
 */
export async function verifySheet228Features(h, { dir }) {
  const { open, check, until, wait, press, errorsIn, doc, sessionFor } = h;
  const file = path.join(dir, 'sheet-228-features.xlsx');
  const rows = [];
  const put = (ref, value) => {
    const m = /^([A-Z])([0-9]+)$/.exec(ref);
    const r = Number(m[2]) - 1;
    const c = m[1].charCodeAt(0) - 65;
    while (rows.length <= r) rows.push([]);
    while (rows[r].length <= c) rows[r].push(null);
    rows[r][c] = value;
  };
  put('B1', 1000); put('B2', 0.1); put('D2', 0.05); put('D3', 0.2); put('D4', 0.5);
  ['x', 'y', 'x', 'z', 'x'].forEach((word, i) => put(`H${i + 1}`, word));
  put('A6', 'Name'); put('B6', 'One'); put('C6', 'Two'); put('D6', 'Three');
  put('A7', 'Ann'); put('B7', 1); put('C7', 2); put('D7', 3);
  put('A8', 'Bea'); put('B8', 4); put('C8', 5); put('D8', 6);
  put('A9', 'Cy'); put('B9', 7); put('C9', 8); put('D9', 9);
  put('A11', 'Item'); put('B11', 'Qty'); put('C11', 'Price'); put('D11', 'Total');
  put('A12', 'Tea'); put('B12', 2); put('C12', 3);
  put('A13', 'Coffee'); put('B13', 5); put('C13', 4);
  put('A14', 'Milk'); put('B14', 1); put('C14', 2);
  fs.writeFileSync(file, buildXlsx({
    sheets: [
      { name: 'Sheet1', rows },
      { name: 'Week1', rows: [[1]] },
      { name: 'Week2', rows: [[2]] },
      { name: 'Week3', rows: [[3]] },
      { name: 'Dupes', rows: [['Item', 'Shop'], ['Tea', 'North'], ['Tea', 'North'], ['Milk', 'South'], ['Tea', 'North'], ['Milk', 'East']] },
      { name: 'Sales', rows: [['Item', 'Sold'], ['Ink', 400], ['Paper', 1200], ['Pens', 1500], ['Tape', 90], ['Glue', 2000]] },
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
    const texts = (...refs) => refs.map(textOf);
    const selected = () => model().selection?.ref ?? null;
    const place = (ref) => {
      const m = /^([A-Z]+)([0-9]+)$/.exec(ref);
      return { col: m[1].split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1, row: Number(m[2]) - 1 };
    };
    const centre = async (ref) => {
      const { row, col } = place(ref);
      const now = model();
      const c = (now.columns || []).find((x) => x.index === col);
      const r = (now.rows || []).find((x) => x.index === row);
      if (!c || !r) return null;
      const layer = await js(`(() => { const l = document.querySelector('.sh-cells'); if (!l) return null; const b = l.getBoundingClientRect(); return { left: b.left, top: b.top }; })()`);
      return layer ? { x: Math.round(layer.left + c.x + c.width / 2), y: Math.round(layer.top + r.y + r.height / 2) } : null;
    };
    const pressAt = async (at, modifiers = []) => {
      if (!at) return false;
      wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1, modifiers });
      wc.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1, modifiers });
      await wait(180);
      return true;
    };
    const click = async (ref, modifiers = []) => pressAt(await centre(ref), modifiers);
    // The middle of an element picked by a JavaScript expression for the page.
    const middle = (expression) => js(`(() => { const r = (${expression})?.getBoundingClientRect(); return r ? { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) } : null; })()`);
    const editing = () => js(`Boolean(document.querySelector('.sh-editor'))`);
    const enter = async (ref, text) => {
      await click(ref);
      await press(wc, 'F2');
      await until(editing, `the editor on ${ref}`, 5000).catch(() => {});
      await js(`(() => {
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
    };
    const found = (english, { tip = false, within = '' } = {}) => {
      const selector = keySelector(english, { tip });
      const words = JSON.stringify(english);
      return `((root) => {
        if (!root) return null;
        const ok = (n) => n && !n.disabled && n.getClientRects().length > 0;
        const keyed = ${selector ? `[...root.querySelectorAll(${JSON.stringify(selector)})].map((n) => (n.matches('button') ? n : n.closest('button'))).find(ok)` : 'null'};
        return keyed || [...root.querySelectorAll('button')].find((n) => ok(n) && n.textContent.trim() === ${words}) || null;
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
    // A choice of an open menu. A label new in 2.28.0 is drawn in its own words, and marked with its key, only once the
    // English copy has it: until then the window draws the words of the key's name.
    const pick = async (key, words) => {
      const item = `[...document.querySelectorAll('.rw-menu button')].find((n) => n.dataset.key === ${JSON.stringify(key)} || [${JSON.stringify(words)}, ${JSON.stringify(key.replace(/-/g, ' '))}].includes(n.textContent.trim()))`;
      await until(() => js(`Boolean(${item})`), `"${words}" in the menu`, 5000).catch(() => {});
      const did = await js(`(() => { const b = ${item}; if (!b || b.disabled) return false; b.click(); return true; })()`);
      await wait(300);
      return did;
    };
    const rightClick = (expression) => js(`(() => {
      const n = ${expression};
      if (!n) return false;
      const r = n.getBoundingClientRect();
      n.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, button: 2, clientX: Math.round(r.left + r.width / 2), clientY: Math.round(r.top + r.height / 2) }));
      return true;
    })()`);
    const sheetTabs = () => js(`[...document.querySelectorAll('.sh-tab:not(.sh-tab-add)')].map((n) => n.innerText.trim())`);
    const tabOf = (name) => `[...document.querySelectorAll('.sh-tab:not(.sh-tab-add)')].find((n) => n.innerText.trim() === ${JSON.stringify(name)})`;
    const toSheet = async (name) => {
      await js(`(() => { ${tabOf(name)}?.click(); return 1; })()`);
      await until(() => model().activeSheet === name, `the sheet ${name}`, 5000).catch(() => {});
      await wait(250);
    };
    const heads = () => js(`[...document.querySelectorAll('.sh-rowheads .sh-head')].map((n) => n.textContent.trim())`);
    const headOf = (n) => `[...document.querySelectorAll('.sh-rowheads .sh-head')].find((h) => h.textContent.trim() === '${n}')`;
    const undo = async () => { await press(wc, 'Z', { modifiers: ['control'] }); await wait(350); };
    const apply = (...ops) => doc.apply({ id: session.id, ops });
    await until(() => js(`Boolean(document.querySelector('.sh-cells .sh-cell'))`), 'the grid', 8000);

    // B44: a Data Table is worked out again when its model changes.
    await enter('B3', '=B1*B2');
    await enter('E1', '=B3');
    await apply({ op: 'dataTable', range: 'D1:E4', rowInput: '', colInput: 'B2' });
    await click('A1');
    const made = texts('E2', 'E3', 'E4');
    await enter('B1', '2000');
    await until(() => textOf('E4') === '1000', 'the table to follow the edit', 5000).catch(() => {});
    const followed = texts('E2', 'E3', 'E4');
    await undo();
    await until(() => textOf('E4') === '500', 'the table to follow the undo', 5000).catch(() => {});
    const undone = texts('E2', 'E3', 'E4');
    check('sheets: a Data Table follows an edit to the cell it is worked out from, and follows Ctrl+Z back',
      made.join() === '50,200,500' && followed.join() === '100,400,1000' && undone.join() === '50,200,500' && textOf('B1') === '1000',
      `made ${JSON.stringify(made)}; after 2000 in B1 ${JSON.stringify(followed)}; after Ctrl+Z ${JSON.stringify(undone)}; B1 ${JSON.stringify(textOf('B1'))}`);

    // B59: XLOOKUP gives the whole row it found.
    await enter('F7', '=XLOOKUP("Bea",A7:A9,B7:D9)');
    check('sheets: XLOOKUP over three columns gives the whole matching row, spilled across',
      texts('F7', 'G7', 'H7').join() === '4,5,6', JSON.stringify(texts('F7', 'G7', 'H7')));

    // B55: COUNTIF with a spill range for its criterion answers for each cell of it.
    await enter('I1', '=UNIQUE(H1:H5)');
    await enter('J1', '=COUNTIF(H1:H5,I1#)');
    check('sheets: COUNTIF asked about a spill range gives a count beside each of its cells',
      texts('I1', 'I2', 'I3').join() === 'x,y,z' && texts('J1', 'J2', 'J3').join() === '3,1,1',
      `the list ${JSON.stringify(texts('I1', 'I2', 'I3'))}; the counts ${JSON.stringify(texts('J1', 'J2', 'J3'))}`);

    // B31: a sum across sheets.
    await enter('F1', '=SUM(Week1:Week3!A1)');
    const summed = textOf('F1');
    await toSheet('Week2');
    await enter('A1', '10');
    await toSheet('Sheet1');
    const afterEdit = textOf('F1');
    await apply({ op: 'renameSheet', from: 'Week1', to: 'First' });
    await click('A1');
    await click('F1');
    const renamed = model().formulaBar ?? null;
    check('sheets: a sum across sheets gives the total, follows an edit on a sheet in the run, and reads the new name when an end sheet is renamed',
      summed === '6' && afterEdit === '14' && renamed === '=SUM(First:Week3!A1)' && textOf('F1') === '14',
      `at first ${JSON.stringify(summed)}; after 10 on Week2 ${JSON.stringify(afterEdit)}; after the rename ${JSON.stringify(renamed)} giving ${JSON.stringify(textOf('F1'))}`);

    // B32: the status bar for cells picked with Ctrl.
    await click('B7');
    await click('C8', ['control']);
    await click('D9', ['control']);
    await wait(300);
    const figures = model().status ?? null;
    const want = [shownAs('Count {n}', { n: 3 }), shownAs('Sum {value}', { value: 15 }), shownAs('Average {value}', { value: 5 })];
    const drawn = await js(`((words) => words.map((w) => [...document.querySelectorAll('span, div')].some((n) => n.children.length === 0 && n.textContent.trim() === w)))(${JSON.stringify(want)})`);
    check('sheets: the status bar counts, sums and averages three cells picked apart with Ctrl',
      figures?.count === 3 && figures?.sum === 15 && figures?.average === 5 && drawn.every(Boolean),
      `the figures ${JSON.stringify(figures)}; ${JSON.stringify(want)} drawn ${JSON.stringify(drawn)}`);

    // B53, B58: the Name Box takes typing.
    const nameBox = async (text, key = 'Return') => {
      await pressAt(await middle(`document.querySelector('.sh-namebox')`));
      await until(() => js(`Boolean(document.querySelector('.sh-namebox-input'))`), 'the Name Box to take typing', 4000).catch(() => {});
      const set = await js(`(() => {
        const e = document.querySelector('.sh-namebox-input');
        if (!e) return false;
        e.focus();
        Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(e, ${JSON.stringify(text)});
        e.dispatchEvent(new Event('input', { bubbles: true }));
        return true;
      })()`);
      await wait(150);
      await press(wc, key);
      await wait(400);
      return set;
    };
    await click('A1');
    const typedIn = await nameBox('B7:D9');
    const went = selected();
    await nameBox('Marks');
    await click('A1');
    await nameBox('marks');
    const byName = selected();
    await click('A1');
    await nameBox('C3', 'Escape');
    const escaped = { selected: selected(), shows: await js(`document.querySelector('.sh-namebox')?.textContent.trim() ?? null`), typing: await js(`Boolean(document.querySelector('.sh-namebox-input'))`) };
    check('sheets: the Name Box takes an address typed into it, names the selection from a new word, goes to that name again, and Escape leaves the selection alone',
      typedIn && went === 'B7:D9' && byName === 'B7:D9' && escaped.selected === 'A1' && escaped.shows === 'A1' && !escaped.typing,
      `typed ${typedIn}; went to ${JSON.stringify(went)}; by its name ${JSON.stringify(byName)}; after Escape ${JSON.stringify(escaped)}`);

    // B29: a table with a formula column, and a new row typed under it.
    await tab('Home');
    await click('A11');
    await click('D14', ['shift']);
    await pressOn('Format as Table');
    await js(`(() => { document.querySelectorAll('.rw-menu button')[1]?.click(); return 1; })()`);
    await wait(500);
    for (const r of [12, 13, 14]) await enter(`D${r}`, `=B${r}*C${r}`);
    await enter('A15', 'Sugar');
    await enter('B15', '4');
    await enter('C15', '5');
    await until(() => textOf('D15') === '20', 'the formula to be carried into the new row', 4000).catch(() => {});
    check('sheets: a row typed under a table joins it, and the column\'s formula fills its cell',
      texts('D12', 'D13', 'D14').join() === '6,20,2' && textOf('D15') === '20',
      `the column ${JSON.stringify(texts('D12', 'D13', 'D14'))}; the new row ${JSON.stringify(texts('A15', 'B15', 'C15', 'D15'))}`);

    // B28: the table's total row, from Format as Table.
    await click('D13');
    await pressOn('Format as Table');
    const totalOn = await pick('total-row', 'Total row');
    await until(() => textOf('D16') === '48', 'the total row', 4000).catch(() => {});
    const total = texts('A16', 'D16');
    await click('D13');
    await pressOn('Format as Table');
    await pick('total-row', 'Total row');
    await until(() => !textOf('D16'), 'the total row to go', 4000).catch(() => {});
    check('sheets: Format as Table, Total row puts a total under the table with the sum of its last column, and takes it away again',
      totalOn && total[0] === 'Total' && total[1] === '48' && !textOf('D16') && !textOf('A16'),
      `picked ${totalOn}; the row ${JSON.stringify(total)}; after it is turned off ${JSON.stringify(texts('A16', 'D16'))}`);

    // B42: rows hidden and shown again from their headings.
    await pressAt(await middle(headOf(18)));
    await pressAt(await middle(headOf(19)), ['shift']);
    const whole = model().selection?.whole ?? null;
    await rightClick(headOf(18));
    const hid = await pick('hide-rows', 'Hide rows');
    await until(async () => !(await heads()).includes('18'), 'the rows to be hidden', 4000).catch(() => {});
    const hidden = (await heads()).filter((n) => ['17', '18', '19', '20'].includes(n));
    await pressAt(await middle(headOf(17)));
    await pressAt(await middle(headOf(20)), ['shift']);
    await rightClick(headOf(17));
    const shown = await pick('unhide-rows', 'Unhide rows');
    await until(async () => (await heads()).includes('18'), 'the rows to be shown', 4000).catch(() => {});
    const back = (await heads()).filter((n) => ['17', '18', '19', '20'].includes(n));
    check('sheets: rows picked by their headings are hidden from the right-click menu, and shown again from the rows on either side',
      whole === 'row' && hid && hidden.join() === '17,20' && shown && back.join() === '17,18,19,20',
      `whole ${JSON.stringify(whole)}; Hide rows ${hid}: headings ${JSON.stringify(hidden)}; Unhide rows ${shown}: headings ${JSON.stringify(back)}`);

    // B38: Remove Duplicates asks which columns.
    await toSheet('Dupes');
    await tab('Data');
    await click('A2');
    await pressOn('Remove Duplicates');
    await until(() => js(`Boolean(document.querySelector('.rw-dialog .sh-dup-go'))`), 'the Remove Duplicates dialog', 5000).catch(() => {});
    // The list is words all the way down, so the dialog does not guess that its first row is headings: it is told.
    const guessed = await js(`document.querySelector('.rw-dialog .sh-dup-header')?.checked ?? null`);
    if (guessed === false) { await js(`(() => { document.querySelector('.rw-dialog .sh-dup-header')?.click(); return 1; })()`); await wait(300); }
    const asked = await js(`(() => { const d = document.querySelector('.rw-dialog'); return d ? { ticks: d.querySelectorAll('input[type=checkbox]').length, words: d.innerText } : null; })()`);
    await js(`(() => { document.querySelector('.rw-dialog .sh-dup-go')?.click(); return 1; })()`);
    await until(() => !textOf('A5'), 'the repeated rows to go', 5000).catch(() => {});
    const kept = [1, 2, 3, 4, 5, 6].map((r) => texts(`A${r}`, `B${r}`).join(' ')).filter((line) => line.trim());
    await click('A2');
    await undo();
    await until(() => textOf('A6') === 'Milk', 'the rows to come back', 5000).catch(() => {});
    const restored = [1, 2, 3, 4, 5, 6].map((r) => textOf(`A${r}`));
    check('sheets: Remove Duplicates names the list\'s columns to choose from, removes the rows that repeat, and one Ctrl+Z brings them back',
      Boolean(asked) && /Item/.test(asked.words) && /Shop/.test(asked.words) && asked.ticks >= 2
        && JSON.stringify(kept) === JSON.stringify(['Item Shop', 'Tea North', 'Milk South', 'Milk East']) && restored.join() === 'Item,Tea,Tea,Milk,Tea,Milk',
      `asked ${JSON.stringify(asked)}; kept ${JSON.stringify(kept)}; after Ctrl+Z ${JSON.stringify(restored)}`);

    // B42: a sheet copied from its tab.
    await rightClick(tabOf('Week3'));
    const copied = await pick('copy-sheet', 'Copy sheet');
    await until(async () => (await sheetTabs()).includes('Week3 (2)'), 'the copy', 5000).catch(() => {});
    const tabs = await sheetTabs();
    const on = model().activeSheet;
    check('sheets: Copy sheet on a tab makes a copy beside it, named for it, and shows it',
      copied && tabs.indexOf('Week3 (2)') === tabs.indexOf('Week3') + 1 && on === 'Week3 (2)' && textOf('A1') === '3',
      `picked ${copied}; tabs ${JSON.stringify(tabs)}; showing ${JSON.stringify(on)} with A1 ${JSON.stringify(textOf('A1'))}`);

    // B42: a tab's colour, picked from the tab's own menu and taken away again.
    const colourOf = (name) => js(`${tabOf(name)}?.dataset.tabColour ?? null`);
    await rightClick(tabOf('First'));
    const askedColour = await pick('tab-colour', 'Tab colour…');
    await until(() => js(`Boolean(document.querySelector('.rw-dialog .sh-tabcolour'))`), 'the tab colour dialog', 5000).catch(() => {});
    const offered = await js(`document.querySelectorAll('.rw-dialog .sh-tabcolour').length`);
    await js(`(() => { document.querySelector('.rw-dialog .sh-tabcolour[data-colour="00B050"]')?.click(); return 1; })()`);
    await until(async () => Boolean(await colourOf('First')), 'the colour on the tab', 5000).catch(() => {});
    const painted = await colourOf('First');
    const band = await js(`getComputedStyle(${tabOf('First')}).boxShadow`);
    const others = await colourOf('Week3');
    await rightClick(tabOf('First'));
    await pick('tab-colour', 'Tab colour…');
    await until(() => js(`Boolean(document.querySelector('.rw-dialog .sh-tabcolour-none'))`), 'the tab colour dialog again', 5000).catch(() => {});
    const marked = await js(`document.querySelector('.rw-dialog .sh-tabcolour[aria-pressed="true"]')?.dataset.colour ?? null`);
    await js(`(() => { document.querySelector('.rw-dialog .sh-tabcolour-none')?.click(); return 1; })()`);
    await until(async () => !(await colourOf('First')), 'the colour to go', 5000).catch(() => {});
    const cleared = await colourOf('First');
    check('sheets: a tab takes a colour from its right-click menu, shows it as a band, and gives it up to No colour',
      askedColour && offered === 10 && /00B050/i.test(String(painted)) && /rgb\(0, 176, 80\)/.test(String(band)) && !others && marked === '00B050' && !cleared,
      `picked ${askedColour}; ${offered} colours offered; the tab says ${JSON.stringify(painted)} and draws ${JSON.stringify(band)}; another tab ${JSON.stringify(others)}; the dialog marked ${JSON.stringify(marked)}; after No colour ${JSON.stringify(cleared)}`);

    // B37: a number filter and a text filter, from the arrow of a heading.
    await toSheet('Sales');
    await click('A2');
    await tab('Data');
    await pressOn('Filter');
    await until(() => js(`document.querySelectorAll('.sh-filter-arrow').length >= 2`), 'the filter arrows', 5000).catch(() => {});
    const arrowOf = (ref) => `document.querySelector('.sh-cell[data-ref="${ref}"] .sh-filter-arrow')`;
    const openArrow = async (ref) => {
      await js(`(() => { ${arrowOf(ref)}?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 })); return 1; })()`);
      await until(() => js(`Boolean(document.querySelector('.sh-filter-menu .sh-filter-rules'))`), `the menu of ${ref}`, 5000).catch(() => {});
    };
    const openRule = async () => {
      if (!(await js(`Boolean(document.querySelector('.sh-filter-menu .sh-filter-rule'))`))) await js(`(() => { document.querySelector('.sh-filter-menu .sh-filter-rules')?.click(); return 1; })()`);
      await until(() => js(`Boolean(document.querySelector('.sh-filter-menu .sh-filter-rule .sh-filter-op'))`), 'the rule form', 5000).catch(() => {});
    };
    const setField = (selector, value) => js(`(() => {
      const n = document.querySelector(${JSON.stringify('.sh-filter-menu ' + selector)});
      const e = n?.matches('input, select') ? n : n?.querySelector('input, select');
      if (!e) return false;
      const proto = e.tagName === 'SELECT' ? HTMLSelectElement.prototype : HTMLInputElement.prototype;
      Object.getOwnPropertyDescriptor(proto, 'value').set.call(e, ${JSON.stringify(String(value))});
      e.dispatchEvent(new Event(e.tagName === 'SELECT' ? 'change' : 'input', { bubbles: true }));
      return true;
    })()`);
    const fieldOf = (selector) => js(`(() => { const n = document.querySelector(${JSON.stringify('.sh-filter-menu ' + selector)}); const e = n?.matches('input, select') ? n : n?.querySelector('input, select'); return e ? e.value : null; })()`);
    const okRule = async () => { await wait(200); await js(`(() => { const b = document.querySelector('.sh-filter-menu .sh-filter-rule-ok'); if (b && !b.disabled) b.click(); return 1; })()`); await wait(400); };
    const rowsShown = async () => (await heads()).filter((h) => Number(h) <= 6).join();

    await openArrow('B1');
    await openRule();
    const firstOp = await fieldOf('.sh-filter-op');
    await setField('.sh-filter-a', '1000');
    await okRule();
    await until(async () => (await rowsShown()) === '1,3,4,6', 'the rows over 1000', 5000).catch(() => {});
    const overRows = await rowsShown();
    const funnelOn = await js(`Boolean(${arrowOf('B1')}?.classList.contains('on'))`);
    await openArrow('B1');
    await until(() => js(`Boolean(document.querySelector('.sh-filter-menu .sh-filter-rule .sh-filter-a'))`), 'the rule shown again', 5000).catch(() => {});
    const ruleKept = { op: await fieldOf('.sh-filter-op'), value: await fieldOf('.sh-filter-a') };
    await setField('.sh-filter-op', 'top');
    await until(() => js(`Boolean(document.querySelector('.sh-filter-menu .sh-filter-count'))`), 'the count of a top filter', 5000).catch(() => {});
    await setField('.sh-filter-count', '2');
    await okRule();
    await until(async () => (await rowsShown()) === '1,4,6', 'the top two', 5000).catch(() => {});
    const topTwo = await rowsShown();
    await openArrow('B1');
    await js(`(() => { const b = document.querySelector('.sh-filter-menu .sh-filter-clear'); if (b && !b.disabled) b.click(); return 1; })()`);
    await until(async () => (await rowsShown()) === '1,2,3,4,5,6', 'the rows to come back', 5000).catch(() => {});
    const allRows = await rowsShown();
    check('sheets: a number filter from a heading\'s arrow keeps the rows over a figure, shows its rule when opened again, takes the top two, and is cleared',
      firstOp === 'greater' && overRows === '1,3,4,6' && funnelOn && ruleKept.op === 'greater' && ruleKept.value === '1000' && topTwo === '1,4,6' && allRows === '1,2,3,4,5,6',
      `the form opens on ${JSON.stringify(firstOp)}; over 1000 rows ${overRows}, arrow marked ${funnelOn}; opened again ${JSON.stringify(ruleKept)}; top two rows ${topTwo}; cleared rows ${allRows}`);

    await openArrow('A1');
    await openRule();
    const textOp = await fieldOf('.sh-filter-op');
    await setField('.sh-filter-a', 'pe');
    await okRule();
    await until(async () => (await rowsShown()) === '1,3,4,5', 'the rows holding "pe"', 5000).catch(() => {});
    const holdingRows = await rowsShown();
    await undo();
    await until(async () => (await rowsShown()) === '1,2,3,4,5,6', 'the rows after Ctrl+Z', 5000).catch(() => {});
    const afterUndo = await rowsShown();
    check('sheets: a text filter keeps the rows whose words contain what was typed, whatever its capitals, and one Ctrl+Z shows the rest',
      textOp === 'contains' && holdingRows === '1,3,4,5' && afterUndo === '1,2,3,4,5,6',
      `the form opens on ${JSON.stringify(textOp)}; rows holding "pe" ${holdingRows}; after Ctrl+Z ${afterUndo}`);

    const errors = await errorsIn(win);
    check('sheets: none of that reported an error', errors.length === 0, JSON.stringify(errors));
  } catch (err) {
    const said = win ? await errorsIn(win).catch(() => []) : [];
    check('sheets: the checks of the 2.28.0 features ran', false, err.message + (said.length ? '. The window said: ' + said.join(' | ') : ''));
  }
}
