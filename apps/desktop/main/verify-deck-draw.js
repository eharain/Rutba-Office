// Presentations: a table drawn in its table style, and the preset shapes
// drawn as their outlines.
//
// The showcase deck's shapes slide — a heart, a smiley, a callout, an up
// arrow, a can among its presets — each drawn as its own outline, not a box
// of its size; and its table slide in the style it names: the header in the
// accent with bold light words and a thicker rule under it, the rows banded
// in its tints, light lines between the cells, the one cell picked out in
// its own fill. Run alone with RUTBA_VERIFY_ONLY=deckdraw.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const FIXTURE = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..', '..', 'tests', 'fixtures', 'rich', 'showcase.pptx');

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args where the copy is opened from
 */
export async function verifyDeckDraw(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const file = path.join(dir, 'draw-showcase.pptx');
  try {
    fs.copyFileSync(FIXTURE, file);
    const win = await open('slides', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const capture = async (name) => { if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await win.webContents.capturePage()).toPNG()); };
    const go = async (n) => {
      await until(() => js(`document.querySelectorAll('.sl-thumb').length > ${n}`), 'the slides', 8000).catch(() => {});
      await js(`(() => { document.querySelectorAll('.sl-thumb')[${n}]?.click(); return 1; })()`);
      await until(() => js(`document.querySelectorAll('.sl-thumb')[${n}]?.classList.contains('active')`), `slide ${n + 1}`, 4000).catch(() => {});
      await wait(400);
    };

    // The shapes slide: the presets the stage once drew as boxes, each a path now.
    await go(2);
    await capture('deck-draw-shapes.png');
    const outlines = await js(`(() => {
      const svg = document.querySelector('.sl-stage svg');
      if (!svg) return null;
      // Each shape's own element: the first drawn element after its words' group start is the outline; count paths and boxes.
      return { paths: svg.querySelectorAll('path').length, words: svg.textContent };
    })()`);
    check('slides: the heart, smiley, callout, arrow and can are drawn as their outlines, not boxes',
      outlines && ['Heart', 'Smile', 'Callout', 'Up arrow', 'Parallelogram'].every((w) => outlines.words.includes(w)) && outlines.paths >= 8,
      outlines ? `${outlines.paths} paths` : 'no stage');

    // The table slide: its style's header, bands and borders.
    await go(3);
    await capture('deck-draw-table.png');
    const table = await js(`(() => {
      const svg = document.querySelector('.sl-stage svg');
      if (!svg) return null;
      const fills = [...svg.querySelectorAll('rect')].map((r) => (r.getAttribute('fill') || '').toLowerCase());
      const lines = [...svg.querySelectorAll('line')].map((l) => (l.getAttribute('stroke') || '').toLowerCase() + '@' + l.getAttribute('stroke-width'));
      const header = [...svg.querySelectorAll('tspan'), ...svg.querySelectorAll('text')].find((t) => t.textContent.trim() === 'Region');
      const style = header ? getComputedStyle(header) : null;
      return { fills, lines, headerFill: header?.getAttribute('fill') || header?.closest('[fill]')?.getAttribute('fill') || style?.fill, headerWeight: header?.getAttribute('font-weight') || header?.closest('[font-weight]')?.getAttribute('font-weight') || style?.fontWeight };
    })()`);
    const has = (c) => table?.fills.includes(c);
    check('slides: a table is drawn in its table style: its header in the accent, its rows banded, light lines between, a cell in its own fill',
      table && has('#156082') && has('#ffe699') && new Set(table.fills.filter((c) => /^#[0-9a-f]{6}$/.test(c))).size >= 4
        && table.lines.some((l) => l.startsWith('#ffffff@4')) && table.lines.filter((l) => l.startsWith('#ffffff')).length >= 10
        && /fff/i.test(String(table.headerFill)) && /700|bold/.test(String(table.headerWeight)),
      table ? `fills ${[...new Set(table.fills)].join(' ')}; ${table.lines.length} lines; header ${table.headerFill} ${table.headerWeight}` : 'no stage');

    const complaints = await errorsIn(win);
    check('slides: drawing tables and shapes reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('slides: the drawing checks ran', false, err.message);
  }
}
