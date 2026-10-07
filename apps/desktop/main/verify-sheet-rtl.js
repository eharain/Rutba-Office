// Worksheets: a sheet right to left, as an Arabic or Hebrew workbook reads.
//
// Page Layout → Sheet Right-to-Left turns the sheet: column A at the right,
// the next columns to its left, the row headings on the right; the words in
// the cells still read; a click lands on the cell under it, the arrow keys
// go the way they point on the screen, and a column's edge dragged to the
// left widens it. Turned back, column A is at the left again. Run alone
// with RUTBA_VERIFY_ONLY=sheetrtl.
import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where the workbook is written
 */
export async function verifySheetRtl(h, { dir }) {
  const { open, check, until, wait, errorsIn, doc, sessionFor } = h;
  const file = path.join(dir, 'rtl-sheet.xlsx');
  fs.writeFileSync(file, buildXlsx({ sheets: [{ name: 'Sheet1', rows: [['الاسم', 'المبلغ', 'Note'], ['أحمد', 120, 'paid'], ['سارة', 75, 'due']] }] }));
  try {
    const win = await open('sheets', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const session = sessionFor('sheet');
    const model = () => doc.model({ id: session.id });
    const capture = async (name) => {
      if (!process.env.RUTBA_VERIFY_CAPTURE) return;
      wc.invalidate();
      await wait(300);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG());
    };
    const rectOf = (ref) => js(`(() => { const b = document.querySelector('.sh-cells [data-ref="${ref}"]')?.getBoundingClientRect(); return b ? { left: Math.round(b.left), right: Math.round(b.right), top: Math.round(b.top) } : null; })()`);
    await until(() => js(`Boolean(document.querySelector('.sh-cells [data-ref="A1"]'))`), 'the grid', 8000);

    // Page Layout → Sheet Right-to-Left.
    await js(`(() => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Page Layout')?.click(); return 1; })()`);
    await wait(250);
    const pressed = await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.textContent || '').trim() === 'Sheet Right-to-Left'); if (!b) return false; b.click(); return true; })()`);
    const turned = await until(() => model().rtl === true, 'the sheet right to left', 4000).then(() => true, () => false);
    await until(() => js(`Boolean(document.querySelector('.sh.rtl'))`), 'the grid mirrored', 4000).catch(() => {});
    await wait(300);
    const a1 = await rectOf('A1');
    const b1 = await rectOf('B1');
    const seen = await js(`(() => {
      const head = [...document.querySelectorAll('.sh-rowheads .sh-head')].find((n) => n.textContent.trim() === '1')?.getBoundingClientRect();
      const words = document.querySelector('.sh-cells [data-ref="A1"] .sh-words');
      const button = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.textContent || '').trim() === 'Sheet Right-to-Left');
      return {
        head: head ? Math.round(head.left) : null,
        words: words ? { text: words.textContent, scale: getComputedStyle(words).scale } : null,
        pressed: button ? button.getAttribute('aria-pressed') === 'true' || button.classList.contains('pressed') || button.classList.contains('on') : null,
      };
    })()`);
    check('sheets: Page Layout → Sheet Right-to-Left turns the sheet', pressed === true && turned === true, `button ${pressed}; sheet right to left ${model().rtl}`);
    check('sheets: right to left, column A is at the right and B to its left, the row headings on the right',
      Boolean(a1 && b1) && a1.left > b1.left && seen.head != null && seen.head >= a1.right - 2,
      `A1 at ${a1?.left}–${a1?.right}, B1 at ${b1?.left}–${b1?.right}, row 1's heading at ${seen.head}`);
    check('sheets: right to left, the words in a cell are turned back to read', seen.words?.text === 'الاسم' && /^-1\b/.test(String(seen.words?.scale || '')),
      JSON.stringify(seen.words));
    await capture('sheet-rtl.png');

    // A click on an empty cell selects it — measured from the right.
    const centreOf = async (row, col) => {
      const m = model();
      const c = (m.columns || []).find((x) => x.index === col);
      const r = (m.rows || []).find((x) => x.index === row);
      const layer = await js(`(() => { const b = document.querySelector('.sh-cells').getBoundingClientRect(); return { right: b.right, top: b.top }; })()`);
      return { x: Math.round(layer.right - (c.x + c.width / 2)), y: Math.round(layer.top + r.y + r.height / 2) };
    };
    const at = await centreOf(4, 3);
    wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
    wc.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1 });
    const clicked = await until(() => { const a = model().selection?.active; return a && a.row === 4 && a.col === 3; }, 'D5 selected', 4000).then(() => true, () => false);
    const active = () => model().selection?.active;
    check('sheets: right to left, a click selects the cell under it', clicked, `active ${JSON.stringify(active())}`);

    // The arrows go the way they point: Left to the next column, Right back.
    const key = async (keyCode) => { wc.sendInputEvent({ type: 'keyDown', keyCode }); wc.sendInputEvent({ type: 'keyUp', keyCode }); await wait(200); };
    await key('Left');
    const afterLeft = { ...active() };
    await key('Right');
    await key('Right');
    const afterRight = { ...active() };
    check('sheets: right to left, Left moves to the next column and Right to the one before, as on the screen',
      afterLeft.col === 4 && afterRight.col === 2, `after Left ${JSON.stringify(afterLeft)}, after Right twice ${JSON.stringify(afterRight)}`);

    // Column B's edge — on its left now — dragged left widens it.
    const widthOf = (col) => (model().columns || []).find((c) => c.index === col)?.width ?? null;
    const before = widthOf(1);
    const grip = await js(`(() => { const head = [...document.querySelectorAll('.sh-colheads .sh-head')].find((n) => n.textContent.trim() === 'B'); const g = head?.querySelector('.sh-grip')?.getBoundingClientRect(); return g ? { x: Math.round(g.left + g.width / 2), y: Math.round(g.top + g.height / 2) } : null; })()`);
    if (grip) {
      wc.sendInputEvent({ type: 'mouseDown', x: grip.x, y: grip.y, button: 'left', clickCount: 1 });
      for (let i = 1; i <= 6; i++) { wc.sendInputEvent({ type: 'mouseMove', x: grip.x - i * 5, y: grip.y, button: 'left' }); await wait(30); }
      wc.sendInputEvent({ type: 'mouseUp', x: grip.x - 30, y: grip.y, button: 'left', clickCount: 1 });
    }
    await until(() => widthOf(1) !== before, 'column B resized', 4000).catch(() => {});
    const after = widthOf(1);
    check('sheets: right to left, a column\'s edge is on its left, and dragged left it widens the column',
      Boolean(grip) && before != null && after != null && Math.abs((after - before) - 30) <= 3, `grip ${JSON.stringify(grip)}; width ${before} → ${after}`);

    // Turned back.
    await js(`(() => { [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.textContent || '').trim() === 'Sheet Right-to-Left')?.click(); return 1; })()`);
    await until(() => model().rtl === false, 'the sheet left to right', 4000).catch(() => {});
    await until(() => js(`!document.querySelector('.sh.rtl')`), 'the grid unmirrored', 4000).catch(() => {});
    await wait(300);
    const a1Back = await rectOf('A1');
    const b1Back = await rectOf('B1');
    check('sheets: Sheet Right-to-Left again puts column A back at the left', model().rtl === false && Boolean(a1Back && b1Back) && a1Back.left < b1Back.left,
      `A1 at ${a1Back?.left}, B1 at ${b1Back?.left}`);

    const complaints = await errorsIn(win);
    check('sheets: a sheet right to left reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('sheets: the right-to-left sheet checks ran', false, err.message);
  }
}
