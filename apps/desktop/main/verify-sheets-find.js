// Worksheets: Find, as Excel's, with the real keys and mouse. Ctrl+F puts
// the keyboard in the Find box at the formula bar's end; Enter finds the
// next match and Shift+Enter the one before, and the box says which of how
// many; Esc goes back to the grid. Ctrl+H opens Find and Replace beside the
// grid, not over it: Match case, Match entire cell contents, Within
// Workbook, Find All's list (a row in it goes to its cell, on its sheet),
// and Replace All, which one Ctrl+Z takes back. Find & Select's Replace…
// opens Replace too. Before, Ctrl+F and Ctrl+H did nothing, and Replace…
// opened Find. Run alone with RUTBA_VERIFY_ONLY=sheets-find.

import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';

/**
 * @param {object} h the harness: open, check, until, wait, capture, doc
 */
export async function verifySheetsFind(h, { dir }) {
  const { open, check, until, wait, capture, doc } = h;
  const own = path.join(dir, 'sheets-find');
  const file = path.join(own, 'find.xlsx');
  try {
    fs.mkdirSync(own, { recursive: true });
    fs.writeFileSync(file, buildXlsx({
      sheets: [
        { name: 'Stock', rows: [['Item', 'Note'], ['Tea', 'Green tea'], ['Apple', 'apple pie'], ['tea', 'Tea']] },
        { name: 'Orders', rows: [['Tea', 2], ['Milk', 'tea later']] },
      ],
    }));
  } catch (err) {
    return check('sheets-find: the fixture was made', false, err.message);
  }

  let win = null;
  try {
    win = await open('sheets', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.sh-cells .sh-cell'))`), 'the grid', 8000);
    const session = doc.sessions().filter((s) => s.kind === 'sheet' && s.path === file).pop();
    const model = () => doc.model({ id: session.id });
    const nameBox = () => js(`document.querySelector('.sh-namebox')?.textContent.trim() || ''`);
    const focusIs = () => js(`(() => { const a = document.activeElement; return a ? a.tagName.toLowerCase() + (a.className ? '.' + String(a.className).trim().split(/\\s+/).join('.') : '') : null; })()`);
    const key = async (keyCode, modifiers = []) => {
      wc.sendInputEvent({ type: 'keyDown', keyCode, modifiers });
      wc.sendInputEvent({ type: 'keyUp', keyCode, modifiers });
      await wait(250);
    };
    const type = async (text) => {
      for (const ch of text) {
        wc.sendInputEvent({ type: 'keyDown', keyCode: ch });
        wc.sendInputEvent({ type: 'char', keyCode: ch });
        wc.sendInputEvent({ type: 'keyUp', keyCode: ch });
        await wait(50);
      }
      await wait(150);
    };
    const centre = (selector) => js(`(() => { const n = ${selector}; if (!n) return null; const r = n.getBoundingClientRect(); return r.width ? { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) } : null; })()`);
    const press = async (point) => {
      if (!point) return 'nowhere';
      wc.sendInputEvent({ type: 'mouseDown', x: point.x, y: point.y, button: 'left', clickCount: 1 });
      wc.sendInputEvent({ type: 'mouseUp', x: point.x, y: point.y, button: 'left', clickCount: 1 });
      await wait(250);
      return 'pressed';
    };
    const cellAt = async (refText) => centre(`document.querySelector('.sh-cell[data-ref="${refText}"]')`);
    const count = () => js(`document.querySelector('.sh-findbox-count')?.textContent.trim() || ''`);
    const panel = () => js(`(() => {
      const p = document.querySelector('.sh-find');
      if (!p) return null;
      const tab = p.querySelector('.sh-find-tabs [aria-selected="true"]')?.textContent.trim();
      return { tab, modal: p.getAttribute('aria-modal'), replaceField: Boolean(p.querySelector('.sh-find-with')), rows: p.querySelectorAll('.sh-find-list .sh-find-row:not(.sh-find-heads)').length, note: p.querySelector('.sh-find-note')?.textContent.trim() || '' };
    })()`);

    // 1. Ctrl+F from the grid: the keyboard is in the Find box.
    await press(await cellAt('A1'));
    await key('F', ['control']);
    const inBox = await focusIs();
    check('sheets-find: Ctrl+F from the grid puts the keyboard in the Find box at the formula bar\'s end',
      /sh-findbox-input/.test(inBox || ''), JSON.stringify({ inBox }));

    // 2. Enter finds the next match and says which of how many; Shift+Enter goes back.
    await type('tea');
    await key('Return');
    await until(async () => (await nameBox()) === 'A2', 'the first match', 3000).catch(() => {});
    const first = { at: await nameBox(), said: await count() };
    await key('Return');
    await until(async () => (await nameBox()) === 'B2', 'the second match', 3000).catch(() => {});
    const second = { at: await nameBox(), said: await count() };
    await key('Return', ['shift']);
    await until(async () => (await nameBox()) === 'A2', 'back to the first', 3000).catch(() => {});
    const backAgain = { at: await nameBox(), said: await count(), focus: await focusIs() };
    await capture(win, 'sheets-find-box.png');
    check('sheets-find: Enter in the box goes to the next match and says "1 of 4", "2 of 4"; Shift+Enter goes back, and the keyboard stays in the box',
      first.at === 'A2' && first.said === '1 of 4' && second.at === 'B2' && second.said === '2 of 4' && backAgain.at === 'A2' && backAgain.said === '1 of 4' && /sh-findbox-input/.test(backAgain.focus || ''),
      JSON.stringify({ first, second, backAgain }));

    // 3. Esc goes back to the grid, and the arrows move the selection again.
    await key('Escape');
    const afterEsc = await focusIs();
    await key('Down');
    await until(async () => (await nameBox()) === 'A3', 'the arrow to move', 2000).catch(() => {});
    check('sheets-find: Esc in the box hands the keyboard back to the grid',
      !/input/.test(afterEsc || '') && (await nameBox()) === 'A3', JSON.stringify({ afterEsc, at: await nameBox() }));

    // 4. Ctrl+H: Find and Replace beside the grid, on Replace, not a dialog over it.
    await key('H', ['control']);
    await until(async () => Boolean(await panel()), 'Find and Replace', 3000).catch(() => {});
    const opened = await panel();
    const grid = await js(`(() => { const p = document.querySelector('.sh-find')?.getBoundingClientRect(); const g = document.querySelector('.sh-grid')?.getBoundingClientRect(); return p && g ? { gridWidth: Math.round(g.width), panelWidth: Math.round(p.width), backdrop: Boolean(document.querySelector('.rw-dialog-backdrop, .rw-scrim, [aria-modal="true"]')) } : null; })()`);
    const typedIn = await focusIs();
    await capture(win, 'sheets-find-panel.png');
    check('sheets-find: Ctrl+H opens Find and Replace on Replace, beside the grid with nothing covering the sheet, the keyboard in Find',
      opened?.tab === 'Replace' && opened.modal === 'false' && opened.replaceField && grid && !grid.backdrop && grid.panelWidth < grid.gridWidth / 2 && /sh-find-what/.test(typedIn || ''),
      JSON.stringify({ opened, grid, typedIn }));

    // 5. Match case, Within Workbook, Find All; a row of the list goes to its cell on its sheet.
    const what = await centre(`document.querySelector('.sh-find .sh-find-what')`);
    await press(what);
    await key('A', ['control']);
    await type('Tea');
    await press(await centre(`document.querySelector('.sh-find .sh-find-case')`));
    await js(`(() => { const s = document.querySelector('.sh-find .sh-find-within'); const set = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set; set.call(s, 'workbook'); s.dispatchEvent(new Event('change', { bubbles: true })); })()`);
    await wait(150);
    await press(await centre(`document.querySelector('.sh-find .sh-find-all')`));
    await until(async () => (await panel())?.rows > 0, 'the list', 3000).catch(() => {});
    const listed = await js(`[...document.querySelectorAll('.sh-find-list .sh-find-row:not(.sh-find-heads)')].map((r) => r.dataset.sheet + '!' + r.dataset.ref)`);
    await press(await centre(`document.querySelector('.sh-find-list .sh-find-row[data-sheet="Orders"]')`));
    await until(() => model().activeSheet === 'Orders', 'the other sheet', 3000).catch(() => {});
    const went = { sheet: model().activeSheet, at: await nameBox() };
    check('sheets-find: Find All with Match case and Within Workbook lists Tea on both sheets and not tea, and a row goes to its cell on its sheet',
      JSON.stringify(listed) === JSON.stringify(['Stock!A2', 'Stock!B4', 'Orders!A1']) && went.sheet === 'Orders' && went.at === 'A1',
      JSON.stringify({ listed, went }));

    // 6. Replace All Within Workbook, then one Ctrl+Z takes all of it back.
    await press(await centre(`document.querySelector('.sh-find .sh-find-with')`));
    await type('Chai');
    await press(await centre(`document.querySelector('.sh-find .sh-find-replace-all')`));
    await wait(400);
    const note = (await panel())?.note || '';
    // The list's row left Orders in view: its A1 was Tea.
    const after = { sheet: model().activeSheet, model: (model().cells || []).find((c) => c.ref === 'A1')?.text ?? null };
    // The panel's own Close, not a tab: the tabs sit in a strip of their own inside the head.
    await press(await centre(`document.querySelector('.sh-find .sh-find-head > button.rw-btn')`));
    await until(async () => !(await panel()), 'the panel to close', 2000).catch(() => {});
    const closed = !(await panel());
    await key('Z', ['control']);
    await wait(400);
    const undone = (model().cells || []).find((c) => c.ref === 'A1')?.text ?? null;
    check('sheets-find: Replace All Within Workbook replaces Tea with Chai (matching case), and one Ctrl+Z puts it back',
      after.model === 'Chai' && /3/.test(note) && closed && undone === 'Tea', JSON.stringify({ note, after, closed, undone }));

    // 7. Find & Select's Replace… opens Replace, not Find.
    const findSelect = await centre(`[...document.querySelectorAll('.rw-ribbon button')].find((b) => (b.textContent || '').trim().startsWith('Find & Select'))`);
    await press(findSelect);
    await until(() => js(`Boolean(document.querySelector('.rw-menu'))`), 'the menu', 2000).catch(() => {});
    await press(await centre(`[...document.querySelectorAll('.rw-menu button')].find((b) => b.querySelector('.rw-menu-label')?.textContent.trim() === 'Replace…')`));
    await until(async () => Boolean(await panel()), 'Replace', 3000).catch(() => {});
    const fromRibbon = await panel();
    check('sheets-find: Find & Select > Replace… opens Find and Replace on Replace',
      fromRibbon?.tab === 'Replace' && fromRibbon.replaceField, JSON.stringify({ fromRibbon, button: findSelect }));
    await key('Escape');
  } catch (err) {
    check('sheets-find: the checks ran', false, err.stack || err.message);
  } finally {
    try { win?.close(); } catch { /* gone */ }
  }
}
