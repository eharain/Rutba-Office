// Worksheets: the Draw tab, with the mouse as the pen — a Pen stroke and a
// Highlighter stroke over the cells kept as ink; with Ink to Shape a drawn
// rectangle becomes a rectangle; the Eraser takes the pen stroke away; the
// Lasso picks the highlighter stroke and Delete takes it; Hide Ink hides
// what is left; and the saved file keeps the ink and the shape.

import fs from 'node:fs';
import path from 'node:path';
import { OoxmlPackage } from '@rutba/ooxml';
import { buildXlsx } from '@rutba/ooxml/build';

/**
 * @param {object} h the harness: open, check, until, wait, doc, sessionFor, capture
 */
export async function verifySheetInk(h, { dir }) {
  const { open, check, until, wait, doc, sessionFor, capture } = h;
  const file = path.join(dir, 'sheet-ink.xlsx');
  try {
    fs.writeFileSync(file, buildXlsx({ sheets: [{ name: 'Sheet1', rows: [['Region', 'Sales'], ['North', 120], ['South', 95]] }] }));
    const win = await open('sheets', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const drawings = () => doc.model({ id: sessionFor('sheet').id }).drawings || [];
    const inks = () => drawings().filter((d) => /^Ink \d+$/.test(d.name || ''));
    const press = (label) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(label)});
      if (!b || b.disabled) return 'no button';
      b.click();
      return 'clicked';
    })()`);
    // Sheet pixels to the window's, through the pen surface itself.
    const frame = async () => js(`(() => { const s = document.querySelector('.sl-ink-surface'); if (!s) return null; const r = s.getBoundingClientRect(); return { left: r.left, top: r.top, k: r.width / parseFloat(s.style.width) }; })()`);
    const stroke = async (points) => {
      const f = await frame();
      if (!f) return false;
      const toWin = ([x, y]) => ({ x: Math.round(f.left + x * f.k), y: Math.round(f.top + y * f.k) });
      const first = toWin(points[0]);
      wc.sendInputEvent({ type: 'mouseDown', x: first.x, y: first.y, button: 'left', clickCount: 1 });
      for (const p of points.slice(1)) {
        const q = toWin(p);
        wc.sendInputEvent({ type: 'mouseMove', x: q.x, y: q.y, button: 'left', modifiers: ['leftButtonDown'] });
        await wait(8);
      }
      const last = toWin(points[points.length - 1]);
      wc.sendInputEvent({ type: 'mouseUp', x: last.x, y: last.y, button: 'left', clickCount: 1 });
      return true;
    };
    const line = (a, b, n = 20) => Array.from({ length: n + 1 }, (_, i) => [a[0] + ((b[0] - a[0]) * i) / n, a[1] + ((b[1] - a[1]) * i) / n]);
    const loop = (corners, n = 10) => corners.flatMap((c, i, a) => (i ? line(a[i - 1], c, n) : [c]));

    await until(() => js(`Boolean(document.querySelector('.rw-tab'))`), 'the ribbon', 8000);
    await js(`(() => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Draw')?.click(); return 1; })()`);
    await wait(250);
    const pen = await press('Pen');
    await until(() => js(`Boolean(document.querySelector('.sl-ink-surface'))`), 'the pen surface', 3000).catch(() => {});
    const drew = await stroke(line([120, 120], [380, 160]));
    const one = await until(() => inks().length === 1, 'the pen stroke', 6000).then(() => true).catch(() => false);
    await press('Highlighter');
    await stroke(line([100, 260], [500, 260]));
    const two = await until(() => inks().length === 2, 'the highlighter stroke', 6000).then(() => true).catch(() => false);
    const seeThrough = await js(`[...document.querySelectorAll('.sh-drawing[data-name^="Ink"] path')].some((p) => p.getAttribute('opacity') === '0.5')`);
    if (capture) await capture(win, 'sheet-ink.png');
    check('worksheets: Draw → Pen and Highlighter keep each stroke over the cells as ink, the highlighter see-through',
      pen === 'clicked' && drew && one && two && seeThrough, JSON.stringify({ pen, drew, one, two, seeThrough, inks: inks().map((d) => ({ name: d.name, x: d.x, y: d.y })) }));

    // Ink to Shape.
    await press('Pen');
    await press('Ink to Shape');
    const before = drawings().length;
    await stroke(loop([[600, 100], [820, 100], [820, 220], [600, 220], [602, 102]], 14));
    const made = await until(() => drawings().length === before + 1 && drawings().some((d) => /^Shape /.test(d.name || '')), 'the rectangle', 6000).then(() => true).catch(() => false);
    await press('Ink to Shape');
    check('worksheets: with Ink to Shape a rectangle drawn over the cells becomes a rectangle', made, JSON.stringify({ made, drawings: drawings().map((d) => d.name) }));

    // The Eraser, then the Lasso and Delete.
    await press('Eraser');
    await stroke(line([250, 100], [250, 190], 10));
    const erased = await until(() => !inks().some((d) => d.name === 'Ink 1'), 'the pen stroke erased', 6000).then(() => true).catch(() => false);
    await press('Lasso');
    await stroke(loop([[60, 220], [560, 220], [560, 300], [60, 300], [60, 222]]));
    await wait(300);
    const picked = await js(`document.querySelectorAll('.sh-drawing.picked').length`);
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'Delete' });
    wc.sendInputEvent({ type: 'keyUp', keyCode: 'Delete' });
    const lassoed = await until(() => inks().length === 0, 'the highlighter stroke deleted', 6000).then(() => true).catch(() => false);
    check('worksheets: the Eraser takes away the stroke it passes over, and the Lasso picks the strokes inside its loop for Delete',
      erased && picked === 1 && lassoed, JSON.stringify({ erased, picked, lassoed, left: drawings().map((d) => d.name) }));

    // One more stroke to hide, and keep.
    await press('Pen');
    await stroke(line([150, 400], [450, 420]));
    await until(() => inks().length === 1, 'a last stroke', 6000).catch(() => {});
    await press('Select');
    // Hide Ink is on the Review tab, as Excel has it.
    await js(`(() => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Review')?.click(); return 1; })()`);
    await wait(250);
    await press('Hide Ink');
    const hidden = await until(() => js(`(() => { const n = document.querySelector('.sh-drawing[data-name^="Ink"]'); return Boolean(n) && getComputedStyle(n).visibility === 'hidden'; })()`), 'the ink hidden', 3000).then(() => true).catch(() => false);
    await press('Hide Ink');
    check('worksheets: Hide Ink takes the strokes out of sight while they stay on the sheet', hidden && inks().length === 1, JSON.stringify({ hidden }));

    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => {
      try {
        const pkg = OoxmlPackage.read(fs.readFileSync(file));
        const part = pkg.partNames().find((n) => /^xl\/drawings\/drawing\d+\.xml$/.test(n));
        const xml = part ? pkg.text(part) : '';
        return (xml.match(/name="Ink \d+"/g) || []).length === 1 && /<a:custGeom>/.test(xml) && /name="Shape \d+"/.test(xml);
      } catch { return false; }
    }, 'the saved ink', 8000).then(() => true).catch(() => false);
    check('worksheets: the saved file keeps the ink stroke and the rectangle drawn with Ink to Shape', saved, saved ? 'one ink stroke and one shape in the drawing part' : 'not in the file');
  } catch (err) {
    check('worksheets: the ink checks ran', false, err.message);
  }
}
