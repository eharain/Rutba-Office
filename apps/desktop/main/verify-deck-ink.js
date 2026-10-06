// Presentations: the Draw tab, with the mouse as the pen — a Pen stroke
// and a Highlighter stroke kept as ink; a wobbly stroke begun along the
// Ruler laid straight on it; with Ink to Shape a drawn rectangle becomes a
// rectangle; the Eraser takes the first stroke away; the Lasso selects the
// highlighter stroke and Delete takes it; Ink Replay draws the ink again;
// Hide Ink hides it; and the saved deck keeps the ink.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait, doc, sessionFor, capture
 */
export async function verifyDeckInk(h, { dir }) {
  const { open, check, until, wait, doc, sessionFor, capture } = h;
  const file = path.join(dir, 'ink.pptx');
  try {
    fs.writeFileSync(file, buildPptx({ title: 'Ink', slides: [{ layout: 'blank' }] }));
    const win = await open('slides', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const shapes = () => doc.model({ id: sessionFor('deck').id }).slide?.shapes || [];
    const inks = () => shapes().filter((s) => /^Ink \d+$/.test(s.name || ''));
    const press = (label, nth = 0) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].filter((n) => n.textContent.trim() === ${JSON.stringify(label)})[${nth}];
      if (!b || b.disabled) return 'no button';
      b.click();
      return 'clicked';
    })()`);
    // Slide pixels to the window's own, for the mouse.
    const frame = await (async () => {
      await until(() => js(`Boolean(document.querySelector('.sl-slide'))`), 'the stage', 8000);
      return js(`(() => { const s = document.querySelector('.sl-slide'); const r = s.getBoundingClientRect(); return { left: r.left, top: r.top, k: r.width / parseFloat(s.style.width) }; })()`);
    })();
    const toWin = ([x, y]) => ({ x: Math.round(frame.left + x * frame.k), y: Math.round(frame.top + y * frame.k) });
    const stroke = async (points) => {
      const first = toWin(points[0]);
      wc.sendInputEvent({ type: 'mouseDown', x: first.x, y: first.y, button: 'left', clickCount: 1 });
      for (const p of points.slice(1)) {
        const q = toWin(p);
        wc.sendInputEvent({ type: 'mouseMove', x: q.x, y: q.y, button: 'left', modifiers: ['leftButtonDown'] });
        await wait(8);
      }
      const last = toWin(points[points.length - 1]);
      wc.sendInputEvent({ type: 'mouseUp', x: last.x, y: last.y, button: 'left', clickCount: 1 });
    };
    const line = (a, b, n = 24, wobble = 0) => Array.from({ length: n + 1 }, (_, i) => [a[0] + ((b[0] - a[0]) * i) / n, a[1] + ((b[1] - a[1]) * i) / n + (wobble ? Math.sin(i) * wobble : 0)]);

    await js(`(() => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Draw')?.click(); return 1; })()`);
    await wait(250);
    // The Pen, then the Highlighter.
    const pen = await press('Pen');
    await stroke(line([200, 200], [500, 260]));
    const one = await until(() => inks().length === 1, 'the pen stroke', 5000).then(() => true).catch(() => false);
    await press('Highlighter');
    await stroke(line([200, 450], [700, 450]));
    const two = await until(() => inks().length === 2, 'the highlighter stroke', 5000).then(() => true).catch(() => false);
    const highlight = inks().find((s) => s.line?.alpha === 0.5);
    const drawn = await js(`[...document.querySelectorAll('.sl-svg [data-shape] path')].some((p) => p.getAttribute('stroke-opacity') === '0.5')`);
    if (capture) await capture(win, 'deck-ink.png');
    check('presentations: Draw → Pen and Highlighter keep each stroke as ink on the slide, the highlighter see-through',
      pen === 'clicked' && one && two && Boolean(highlight) && highlight.line.cap === 'square' && drawn,
      JSON.stringify({ pen, one, two, inks: inks().map((s) => ({ name: s.name, line: s.line })), drawn }));

    // The Ruler: a wobbly stroke begun along its top edge comes out straight.
    await press('Pen');
    await press('Ruler');
    const ruled = await until(() => js(`Boolean(document.querySelector('.sl-ink-ruler'))`), 'the ruler', 3000).then(() => true).catch(() => false);
    // Just above the edge, as a pen is laid against a ruler — on it, the press moves the ruler.
    await stroke(line([300, 318], [800, 318], 30, 4));
    await until(() => inks().length === 3, 'the ruled stroke', 5000).catch(() => {});
    const straight = inks()[2];
    await press('Ruler');
    check('presentations: Draw → Ruler lays a straight edge on the slide, and a wobbly stroke begun along it comes out straight on it',
      ruled && Boolean(straight) && straight.geometry.h < 2 && Math.abs(straight.geometry.y - 328) < 2,
      JSON.stringify({ ruled, straight: straight && straight.geometry }));

    // Ink to Shape: a drawn rectangle becomes one.
    await press('Ink to Shape');
    const before = shapes().length;
    const corners = [[850, 150], [1100, 150], [1100, 300], [850, 300], [852, 152]];
    await stroke(corners.flatMap((c, i) => (i ? line(corners[i - 1], c, 14) : [c])));
    const made = await until(() => shapes().length === before + 1 && shapes().some((s) => !/^Ink/.test(s.name || '') && s.kind !== 'picture'), 'the rectangle', 5000).then(() => true).catch(() => false);
    const rect = shapes()[shapes().length - 1];
    await press('Ink to Shape');
    check('presentations: with Ink to Shape a rectangle drawn on the slide becomes a rectangle in the pen\'s colour',
      made && !/^Ink/.test(rect?.name || '') && rect.line?.color && Math.abs(rect.geometry.w - 250) < 20, JSON.stringify({ made, rect: rect && { name: rect.name, g: rect.geometry, line: rect.line } }));

    // The Eraser takes the first stroke; the Lasso selects the highlighter's and Delete takes it.
    await press('Eraser');
    await stroke(line([350, 180], [350, 280], 12));
    const erased = await until(() => !inks().some((s) => s.name === 'Ink 1'), 'the first stroke erased', 5000).then(() => true).catch(() => false);
    await press('Lasso');
    await stroke([[150, 400], [750, 400], [750, 500], [150, 500], [150, 402]].flatMap((c, i, a) => (i ? line(a[i - 1], c, 10) : [c])));
    await wait(300);
    await js(`document.querySelector('.sl-stage')?.focus(), 1`);
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'Delete' });
    wc.sendInputEvent({ type: 'keyUp', keyCode: 'Delete' });
    const lassoed = await until(() => !inks().some((s) => s.line?.alpha === 0.5), 'the highlighter stroke deleted', 5000).then(() => true).catch(() => false);
    check('presentations: the Eraser takes away the stroke it passes over, and the Lasso selects the strokes inside its loop for Delete',
      erased && lassoed, JSON.stringify({ erased, lassoed, left: inks().map((s) => s.name) }));

    // Ink Replay and Hide Ink on what is left.
    const replayed = await press('Ink Replay');
    await wait(900);
    await js(`(() => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Review')?.click(); return 1; })()`);
    await wait(200);
    await press('Hide Ink');
    const hidden = await until(() => js(`(() => { const id = ${JSON.stringify(String(inks()[0]?.id ?? ''))}; const g = document.querySelector('.sl-svg [data-shape="' + id + '"]'); return Boolean(g) && getComputedStyle(g).visibility === 'hidden'; })()`), 'the ink hidden', 3000).then(() => true).catch(() => false);
    await press('Hide Ink');
    check('presentations: Ink Replay draws the ink again, and Hide Ink takes it out of sight while it stays on the slide', replayed === 'clicked' && hidden && inks().length === 1, JSON.stringify({ replayed, hidden, inks: inks().length }));

    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => {
      try {
        const names = Deck.open(fs.readFileSync(file)).slide(0).shapes.map((s) => s.name);
        return names.filter((n) => /^Ink \d+$/.test(n)).length === 1 && names.length === 2;
      } catch { return false; }
    }, 'the saved ink', 8000).then(() => true).catch(() => false);
    check('presentations: the saved deck keeps the ink stroke and the rectangle drawn with Ink to Shape', saved, saved ? 'one ink stroke and one shape' : 'not in the file');
  } catch (err) {
    check('presentations: the ink checks ran', false, err.message);
  }
}
