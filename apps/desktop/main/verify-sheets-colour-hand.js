// Worksheets: Font colour and Fill colour as a person uses them, with the
// real mouse. A cell is clicked, Font colour and a colour are pressed, then
// Fill colour and a fill; then a range is chosen by a click and a
// Shift+click and the same is done to it. The engine and the drawn grid are
// both read, for every cell.
//
// It reads both the older build's two buttons (each opens a list of names)
// and the palette's split buttons, so the same file runs on an older build
// to show what a person saw there. Run alone with
// RUTBA_VERIFY_ONLY=sheets-colour-hand.

import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';
import { labelKeyed } from './verify-keys.js';

/**
 * @param {object} h the harness: open, check, until, wait, capture, doc
 */
export async function verifySheetsColourByHand(h, { dir }) {
  const { open, check, until, wait, capture, doc } = h;
  const own = path.join(dir, 'sheets-colour-hand');
  const file = path.join(own, 'by-hand.xlsx');
  try {
    fs.mkdirSync(own, { recursive: true });
    fs.writeFileSync(file, buildXlsx({
      sheets: [{ name: 'Sheet1', rows: [['Name', 'Qty', 'Price'], ['Tea', 2, 3], ['Milk', 4, 5], ['Bread', 6, 7]] }],
    }));
  } catch (err) {
    return check('sheets-colour-hand: the fixture was made', false, err.message);
  }

  let win = null;
  try {
    win = await open('sheets', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.sh-cells .sh-cell'))`), 'the grid', 8000);
    await wait(400);
    const session = doc.sessions().filter((s) => s.kind === 'sheet' && s.path === file).pop();
    const model = () => doc.model({ id: session.id });
    const cellOf = (ref) => (model().cells || []).find((c) => c.ref === ref) || null;
    const hex6 = (v) => String(v || '').toUpperCase().replace(/^#/, '').slice(-6) || null;
    const fontOf = (ref) => hex6(cellOf(ref)?.style?.font?.colour ?? cellOf(ref)?.style?.font?.color);
    const fillOf = (ref) => hex6(cellOf(ref)?.style?.fill?.colour ?? cellOf(ref)?.style?.fill?.color ?? cellOf(ref)?.style?.fill);
    const rgb = (hex) => `rgb(${parseInt(hex.slice(0, 2), 16)}, ${parseInt(hex.slice(2, 4), 16)}, ${parseInt(hex.slice(4, 6), 16)})`;
    // The grid as drawn: the cell's words' colour and the cell's background.
    const drawn = (ref) => js(`(() => {
      const c = document.querySelector('.sh-cell[data-ref="${ref}"]');
      if (!c) return null;
      const all = [c, ...c.querySelectorAll('*')].filter((n) => [...n.childNodes].some((x) => x.nodeType === 3 && x.textContent.trim()));
      return { text: all.length ? getComputedStyle(all[all.length - 1]).color : null, fill: getComputedStyle(c).backgroundColor };
    })()`);

    const centre = (selector) => js(`(() => { const n = ${selector}; if (!n) return null; const r = n.getBoundingClientRect(); return r.width ? { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) } : { hidden: true }; })()`);
    const press = async (point, modifiers = []) => {
      if (!point || point.hidden) return 'nowhere';
      wc.sendInputEvent({ type: 'mouseDown', x: point.x, y: point.y, button: 'left', clickCount: 1, modifiers });
      wc.sendInputEvent({ type: 'mouseUp', x: point.x, y: point.y, button: 'left', clickCount: 1, modifiers });
      await wait(250);
      return 'pressed';
    };
    const clickCell = (ref, modifiers = []) => centre(`document.querySelector('.sh-cell[data-ref="${ref}"]')`).then((p) => press(p, modifiers));
    const nameBox = () => js(`document.querySelector('.sh-namebox')?.textContent.trim() || ''`);

    // Font colour or Fill colour, then one colour: the palette's arrow and a
    // standard colour, or the older build's button and a name in its list.
    const choose = async (which) => {
      const kind = which === 'font' ? 'sheets.font' : 'sheets.fill';
      const title = which === 'font' ? 'Font colour' : 'Fill colour';
      const palette = await centre(`document.querySelector('.rw-ribbon .rw-split[data-kind="${kind}"] .rw-split-arrow')`);
      const old = palette ? null : await centre(`[...document.querySelectorAll('.rw-ribbon button')].find((b) => (b.dataset.tip || b.title || '') === ${JSON.stringify(title)})`);
      const button = palette || old;
      if (!button || button.hidden) return { pressed: `no ${title} button on the ribbon` };
      await press(button);
      await until(() => js(`Boolean(document.querySelector('.rw-colours, .rw-menu'))`), 'the colours', 3000).catch(() => {});
      const want = palette ? (which === 'font' ? 'FF0000' : '92D050') : (which === 'font' ? 'E03131' : 'D3F9D8');
      const swatch = palette
        ? await centre(`document.querySelector('.rw-colours .rw-cell.std[data-value="${want}"]')`)
        : await centre(`((list) => list.find((b) => ${labelKeyed(which === 'font' ? 'Red' : 'Light green')}(b)) || list.find((b) => b.querySelector('.rw-menu-label')?.textContent.trim() === ${JSON.stringify(which === 'font' ? 'Red' : 'Light green')}))([...document.querySelectorAll('.rw-menu button')])`);
      if (!swatch || swatch.hidden) return { pressed: 'no colour in the list', build: palette ? 'palette' : 'menu' };
      await press(swatch);
      return { pressed: 'pressed', want, build: palette ? 'palette' : 'menu' };
    };
    const settle = async (refs, read, want, paint) => {
      await until(() => refs.every((r) => read(r) === want), 'the engine', 4000).catch(() => {});
      await until(async () => (await Promise.all(refs.map(drawn))).every((d) => d?.[paint] === rgb(want)), 'the grid to draw it', 4000).catch(() => {});
      return { engine: Object.fromEntries(refs.map((r) => [r, read(r)])), drawn: Object.fromEntries(await Promise.all(refs.map(async (r) => [r, (await drawn(r))?.[paint] ?? null]))) };
    };
    const ok = (got, refs, want) => refs.every((r) => got.engine[r] === want && got.drawn[r] === rgb(want));

    // 1. Font colour on one cell.
    await clickCell('A2');
    const one = await nameBox();
    const f1 = await choose('font');
    const g1 = await settle(['A2'], fontOf, f1.want, 'text');
    check('sheets-colour-hand: Font colour, then a colour, pressed with the mouse colours the clicked cell\'s words, in the engine and as drawn',
      f1.pressed === 'pressed' && ok(g1, ['A2'], f1.want), JSON.stringify({ ...f1, at: one, ...g1 }));

    // 2. Fill colour on the same cell.
    const b2 = await choose('fill');
    const g2 = await settle(['A2'], fillOf, b2.want, 'fill');
    await capture(win, 'sheets-colour-hand-cell.png');
    check('sheets-colour-hand: Fill colour, then a fill, pressed with the mouse fills the clicked cell, in the engine and as drawn',
      b2.pressed === 'pressed' && ok(g2, ['A2'], b2.want), JSON.stringify({ ...b2, ...g2 }));

    // 3. Font colour on a range chosen by a click and a Shift+click.
    const range = ['B2', 'C2', 'B3', 'C3'];
    await clickCell('B2');
    await clickCell('C3', ['shift']);
    const chosen = await nameBox();
    const f3 = await choose('font');
    const g3 = await settle(range, fontOf, f3.want, 'text');
    check('sheets-colour-hand: Font colour on a range chosen by a click and a Shift+click colours every cell of it, in the engine and as drawn',
      f3.pressed === 'pressed' && ok(g3, range, f3.want), JSON.stringify({ ...f3, chosen, ...g3 }));

    // 4. Fill colour on the same range.
    const b4 = await choose('fill');
    const g4 = await settle(range, fillOf, b4.want, 'fill');
    await capture(win, 'sheets-colour-hand-range.png');
    check('sheets-colour-hand: Fill colour on the range fills every cell of it, in the engine and as drawn, and the cells around it keep none',
      b4.pressed === 'pressed' && ok(g4, range, b4.want) && !fillOf('D2') && !fillOf('B4'), JSON.stringify({ ...b4, ...g4, D2: fillOf('D2'), B4: fillOf('B4') }));
  } catch (err) {
    check('sheets-colour-hand: the checks ran', false, err.stack || err.message);
  } finally {
    try { win?.close(); } catch { /* gone */ }
  }
}
