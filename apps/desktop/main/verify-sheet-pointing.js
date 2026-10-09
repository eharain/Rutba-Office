// Worksheets: pointing at cells while a formula is typed, as in Excel.
//
// The owner's report: typing "=" in a cell and then picking cells — =A1*A2,
// =SUM(B3:B12) — should go where the formula's caret is, each reference
// shown in its own colour, in the formula and round its cells. A click used
// to commit the half-typed formula and move the selection instead. Here
// the keys are typed and the cells clicked and dragged as a person does.
// Run alone with RUTBA_VERIFY_ONLY=pointing.
import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where the workbook is written
 */
export async function verifySheetPointing(h, { dir }) {
  const { open, check, until, wait, press, errorsIn, doc, sessionFor } = h;
  const file = path.join(dir, 'pointing.xlsx');
  const rows = [[6, 'x'], [7, 'x']];
  for (let r = 3; r <= 12; r++) rows.push([null, r]);
  fs.writeFileSync(file, buildXlsx({ sheets: [{ name: 'Sheet1', rows }] }));
  let win = null;
  try {
    win = await open('sheets', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const session = sessionFor('sheet');
    const valueOf = (ref) => doc.model({ id: session.id }).cells?.find((c) => c.ref === ref)?.text ?? null;
    // Where a cell is on the screen — an empty one is not drawn as an element, so from the
    // grid's own columns and rows, measured from the cells layer as the window measures a press.
    const centre = async (ref) => {
      const m = /^([A-Z]+)([0-9]+)$/.exec(ref);
      const col = m[1].split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
      const row = Number(m[2]) - 1;
      const model = doc.model({ id: session.id });
      const c = (model.columns || []).find((x) => x.index === col);
      const r = (model.rows || []).find((x) => x.index === row);
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
    const click = async (ref, modifiers = []) => { await mouse('mouseDown', ref, { modifiers }); await mouse('mouseUp', ref, { modifiers }); await wait(150); };
    const type = async (text) => {
      for (const ch of text) {
        wc.sendInputEvent({ type: 'keyDown', keyCode: ch });
        wc.sendInputEvent({ type: 'char', keyCode: ch });
        wc.sendInputEvent({ type: 'keyUp', keyCode: ch });
        await wait(60);
      }
      await wait(120);
    };
    const editor = () => js(`document.querySelector('.sh-editor')?.value ?? null`);
    const boxes = () => js(`[...document.querySelectorAll('.sh-refbox')].map((b) => ({ ref: b.dataset.refText, colour: getComputedStyle(b).borderTopColor, on: b.classList.contains('on') }))`);
    const painted = () => js(`[...document.querySelectorAll('.sh-paint-cell .sh-ref-text')].map((s) => ({ text: s.textContent, colour: s.style.color }))`);
    await until(() => js(`Boolean(document.querySelector('.sh-cells .sh-cell'))`), 'the grid', 8000);

    // =A1*A2: "=" typed, A1 clicked, "*" typed, A2 clicked.
    await click('D1');
    await type('=');
    await until(async () => (await editor()) === '=', 'the edit to start', 4000).catch(() => {});
    await click('A1');
    const afterA1 = await editor();
    await type('*');
    await click('A2');
    const formula = await editor();
    const twoBoxes = await boxes();
    const twoPainted = await painted();
    if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'sheet-pointing.png'), (await wc.capturePage()).toPNG());
    check('sheets: typing "=", clicking A1, "*", clicking A2 builds =A1*A2 in the cell, not a committed half formula',
      afterA1 === '=A1' && formula === '=A1*A2', `after A1 ${JSON.stringify(afterA1)}; then ${JSON.stringify(formula)}; D1 holds ${JSON.stringify(valueOf('D1'))}`);
    check('sheets: each reference is boxed round its cells in its own colour, the one at the caret drawn heavier',
      twoBoxes.length === 2 && twoBoxes[0].ref === 'A1' && twoBoxes[1].ref === 'A2' && twoBoxes[0].colour !== twoBoxes[1].colour && twoBoxes[1].on && !twoBoxes[0].on,
      JSON.stringify(twoBoxes));
    check('sheets: and coloured the same in the formula itself',
      twoPainted.length === 2 && twoPainted[0].text === 'A1' && twoPainted[1].text === 'A2' && twoPainted[0].colour !== twoPainted[1].colour,
      JSON.stringify(twoPainted));
    await press(wc, 'Return');
    const product = await until(() => valueOf('D1') === '42', 'D1 to hold 6×7', 4000).catch(() => false);
    check('sheets: Enter commits the pointed formula, which calculates', product === true, `D1 ${JSON.stringify(valueOf('D1'))}`);

    // Back to a formula already entered: selecting D1 again and editing it
    // (F2, then a double-click) boxes its references again, as Excel does.
    await click('D1');
    await wait(150);
    await press(wc, 'F2');
    await until(async () => (await editor()) === '=A1*A2', 'the F2 edit', 3000).catch(() => {});
    await wait(200);
    const againF2 = await boxes();
    await press(wc, 'Escape');
    await wait(200);
    await click('D1');
    const at = await centre('D1');
    wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 2 });
    wc.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 2 });
    await until(async () => (await editor()) === '=A1*A2', 'the double-click edit', 3000).catch(() => {});
    await wait(200);
    const againDbl = await boxes();
    await press(wc, 'Escape');
    await wait(200);
    // And by clicking into the formula bar, before any key: Excel colours the
    // references and boxes their cells the moment the bar is clicked.
    await click('D1');
    await wait(150);
    const bar = await js(`(() => { const r = document.querySelector('.sh-formula input')?.getBoundingClientRect(); return r ? { x: Math.round(r.left + Math.min(r.width - 4, 120)), y: Math.round(r.top + r.height / 2) } : null; })()`);
    wc.sendInputEvent({ type: 'mouseDown', x: bar.x, y: bar.y, button: 'left', clickCount: 1 });
    wc.sendInputEvent({ type: 'mouseUp', x: bar.x, y: bar.y, button: 'left', clickCount: 1 });
    await until(async () => (await boxes()).length === 2, 'the bar\'s boxes', 3000).catch(() => {});
    const inBar = await boxes();
    const barPainted = await js(`[...document.querySelectorAll('.sh-paint-bar .sh-ref-text')].map((s) => s.textContent)`);
    await click('E8');
    await wait(250);
    const afterBar = await boxes();
    const stillD1 = valueOf('D1');
    check('sheets: clicking into the formula bar of a formula colours its references and boxes their cells at once, and a click on a cell moves on as before',
      inBar.map((b) => b.ref).join() === 'A1,A2' && barPainted.join() === 'A1,A2' && afterBar.length === 0 && stillD1 === '42',
      `bar ${JSON.stringify(inBar)}; painted ${JSON.stringify(barPainted)}; after a click on E8 ${afterBar.length} box(es); D1 ${JSON.stringify(stillD1)}`);
    check('sheets: editing a formula already entered boxes its references again, by F2 and by a double-click',
      againF2.map((b) => b.ref).join() === 'A1,A2' && againDbl.map((b) => b.ref).join() === 'A1,A2',
      `F2 ${JSON.stringify(againF2)}; double-click ${JSON.stringify(againDbl)}`);

    // =SUM(B3:B12): a drag from B3 to B12 makes the range.
    await click('D2');
    await type('=SUM(');
    await mouse('mouseDown', 'B3');
    await wait(80);
    for (const ref of ['B5', 'B8', 'B12']) { await mouse('mouseMove', ref); await wait(80); }
    const dashed = await js(`Boolean(document.querySelector('.sh-refbox.on.pointing'))`);
    await mouse('mouseUp', 'B12');
    await wait(150);
    const dragged = await editor();
    await type(')');
    await press(wc, 'Return');
    const sum = await until(() => valueOf('D2') === '75', 'D2 to hold the sum', 4000).catch(() => false);
    check('sheets: dragging from B3 to B12 after "=SUM(" makes the range, drawn dashed while it is dragged, and sums it',
      dragged === '=SUM(B3:B12' && dashed && sum === true, `${JSON.stringify(dragged)}; dashed ${dashed}; D2 ${JSON.stringify(valueOf('D2'))}`);

    // Arrows point in an edit begun by typing; Shift stretches; F4 adds the dollars.
    await click('D3');
    await type('=');
    await press(wc, 'Left');
    const left = await editor();
    await press(wc, 'Left', { modifiers: ['shift'] });
    const stretched = await editor();
    await press(wc, 'F4');
    const absolute = await editor();
    check('sheets: in an edit begun by typing, the arrows point, Shift+arrow stretches, and F4 makes the reference absolute',
      left === '=C3' && stretched === '=B3:C3' && absolute === '=$B$3:$C$3', `${JSON.stringify(left)} → ${JSON.stringify(stretched)} → ${JSON.stringify(absolute)}`);
    await press(wc, 'Escape');
    await wait(200);

    // Where a reference cannot go, a click leaves the cell as before.
    await click('D4');
    await type('=SUM');
    await click('A1');
    await wait(300);
    const left4 = await js(`document.querySelector('.sh-cell.active')?.dataset.ref ?? null`);
    check('sheets: with the caret in a name, a click leaves the edit and selects the cell, as in Excel',
      left4 === 'A1', `active ${left4}; D4 ${JSON.stringify(valueOf('D4'))}`);

    const complaints = await errorsIn(win);
    check('sheets: pointing reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    const said = win ? await errorsIn(win).catch(() => []) : [];
    check('sheets: the pointing checks ran', false, err.message + (said.length ? ' — the window said: ' + said.join(' | ') : ''));
  }
}
