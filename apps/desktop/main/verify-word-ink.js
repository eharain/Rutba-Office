// Documents: the Draw tab, with the mouse as the pen, in Print Layout — a
// Pen stroke and a Highlighter stroke on page 1 float where they were
// drawn; with Ink to Shape a drawn rectangle becomes a rectangle; the
// Eraser takes the pen stroke away; the Lasso picks the highlighter stroke
// and Delete takes it; and the saved document keeps the ink and the shape.

import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { openDocx } from '@rutba/doc-view/backends/ooxml';

/**
 * @param {object} h the harness: open, check, until, wait, capture
 */
export async function verifyWordInk(h, { dir }) {
  const { open, check, until, wait, capture } = h;
  const file = path.join(dir, 'word-ink.docx');
  try {
    fs.writeFileSync(file, buildDocx({ styles: true, paragraphs: [{ text: 'Notes', style: 'Heading1' }, { text: 'Mark this up with the pen.' }, { text: 'And this.' }, { text: '' }] }));
    const win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const drawn = () => js(`[...document.querySelectorAll('.wd-page .wd-drawing')].map((n) => n.getAttribute('alt') || '')`);
    const inks = async () => (await drawn()).filter((n) => /^Ink \d+$/.test(n));
    const press = (label) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(label)});
      if (!b || b.disabled) return 'no button';
      b.click();
      return 'clicked';
    })()`);
    const frame = () => js(`(() => { const s = document.querySelector('.sl-ink-surface'); if (!s) return null; const r = s.getBoundingClientRect(); return { left: r.left, top: r.top, k: r.width / parseFloat(s.style.width) }; })()`);
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

    await until(() => js(`document.querySelectorAll('.wd-sheet').length >= 1`), 'the page', 8000);
    await js(`(() => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Draw')?.click(); return 1; })()`);
    await wait(250);
    const pen = await press('Pen');
    await until(() => js(`Boolean(document.querySelector('.sl-ink-surface'))`), 'the pen surface', 3000).catch(() => {});
    await stroke(line([150, 300], [450, 340]));
    const one = await until(async () => (await inks()).length === 1, 'the pen stroke', 6000).then(() => true).catch(() => false);
    // It floats where it was drawn: about 150 px in from the page's left edge.
    const placed = await js(`(() => { const n = [...document.querySelectorAll('.wd-page .wd-drawing')].find((x) => x.getAttribute('alt') === 'Ink 1'); const s = document.querySelector('.wd-sheet'); if (!n || !s) return null; const k = s.getBoundingClientRect().width / s.offsetWidth; return Math.round((n.getBoundingClientRect().left - s.getBoundingClientRect().left) / k); })()`);
    await press('Highlighter');
    await stroke(line([120, 450], [520, 450]));
    const two = await until(async () => (await inks()).length === 2, 'the highlighter stroke', 6000).then(() => true).catch(() => false);
    if (capture) await capture(win, 'word-ink.png');
    check('documents: Draw → Pen and Highlighter keep each stroke floating on the page where it was drawn',
      pen === 'clicked' && one && two && Math.abs((placed ?? -99) - 147) < 8, JSON.stringify({ pen, one, two, placed, drawn: await drawn() }));

    // Ink to Shape.
    await press('Pen');
    await press('Ink to Shape');
    await stroke(loop([[500, 150], [700, 150], [700, 260], [500, 260], [502, 152]], 14));
    const made = await until(async () => (await drawn()).some((n) => /^Shape \d+$/.test(n)), 'the rectangle', 6000).then(() => true).catch(() => false);
    await press('Ink to Shape');
    check('documents: with Ink to Shape a rectangle drawn on the page becomes a rectangle floating there', made, JSON.stringify({ made, drawn: await drawn() }));

    // The Eraser, then the Lasso and Delete.
    await press('Eraser');
    await stroke(line([300, 280], [300, 360], 10));
    const erased = await until(async () => !(await inks()).includes('Ink 1'), 'the pen stroke erased', 6000).then(() => true).catch(() => false);
    await press('Lasso');
    await stroke(loop([[90, 410], [560, 410], [560, 490], [90, 490], [90, 412]]));
    await wait(300);
    const picked = await js(`document.querySelectorAll('.wd-page .wd-drawing.picked').length`);
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'Delete' });
    wc.sendInputEvent({ type: 'keyUp', keyCode: 'Delete' });
    const lassoed = await until(async () => (await inks()).length === 0, 'the highlighter stroke deleted', 6000).then(() => true).catch(() => false);
    check('documents: the Eraser takes away the stroke it passes over, and the Lasso picks the strokes inside its loop for Delete',
      erased && picked >= 1 && lassoed, JSON.stringify({ erased, picked, lassoed, drawn: await drawn() }));

    // One last stroke to keep.
    await press('Pen');
    await stroke(line([150, 600], [400, 620]));
    await until(async () => (await inks()).length === 1, 'a last stroke', 6000).catch(() => {});
    await press('Select');
    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => {
      try {
        const names = openDocx(fs.readFileSync(file)).doc.doc.drawings().map((d) => d.name);
        return names.filter((n) => /^Ink \d+$/.test(n)).length === 1 && names.some((n) => /^Shape \d+$/.test(n));
      } catch { return false; }
    }, 'the saved ink', 8000).then(() => true).catch(() => false);
    check('documents: the saved document keeps the ink stroke and the rectangle drawn with Ink to Shape', saved, saved ? 'one ink stroke and one shape' : 'not in the file');
  } catch (err) {
    check('documents: the ink checks ran', false, err.message);
  }
}
