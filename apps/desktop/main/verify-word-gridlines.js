// Documents: View → Gridlines, pressed on the ribbon — a grid over the
// words' area of every page, inside the margins, off the paper's edge, and
// gone again when pressed a second time. Nothing is written to the file.

import fs from 'node:fs';
import path from 'node:path';

/**
 * @param {object} h the harness: open, check, until, wait
 */
export async function verifyWordGridlines(h, { file }) {
  const { open, check, until, wait } = h;
  const copy = path.join(path.dirname(file), 'gridlines.docx');
  try {
    fs.copyFileSync(file, copy);
    const before = fs.readFileSync(copy);
    const win = await open('word', copy);
    const js = (code) => win.webContents.executeJavaScript(code);
    const press = () => js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'View')?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Gridlines');
      if (!b) return 'no button'; if (b.disabled) return 'disabled'; b.click(); return 'clicked';
    })()`);
    // Each sheet's grid, measured against the sheet: it must sit inside it,
    // short of every edge (the margins), and carry the line pattern.
    const grids = () => js(`[...document.querySelectorAll('.wd-sheet')].map((sheet) => {
      const g = sheet.querySelector('.wd-gridlines');
      if (!g) return null;
      const s = sheet.getBoundingClientRect(), r = g.getBoundingClientRect();
      return { inside: r.left > s.left && r.top > s.top && r.right < s.right && r.bottom < s.bottom, lines: getComputedStyle(g).backgroundImage.includes('linear-gradient') };
    })`);

    await until(async () => (await js(`document.querySelectorAll('.wd-sheet').length`)) > 0, 'the pages', 6000);
    const off = await grids();
    const pressed = await press();
    await until(async () => { const g = await grids(); return g.length > 0 && g.every(Boolean); }, 'a grid on every page', 4000).catch(() => {});
    const on = await grids();
    const lit = await js(`[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Gridlines')?.getAttribute('aria-pressed')`);
    check('documents: View → Gridlines lays a grid over every page, inside its margins, and lights the button',
      off.every((g) => g === null) && pressed === 'clicked' && on.length > 0 && on.every((g) => g && g.inside && g.lines) && lit === 'true',
      JSON.stringify({ off: off.length, pressed, on, lit }));

    await press();
    const gone = await until(async () => (await grids()).every((g) => g === null), 'the grid to go', 4000).then(() => true).catch(() => false);
    await wait(200);
    check('documents: pressed again, the gridlines go, and the file is untouched — they are a view, not part of the document',
      gone === true && fs.readFileSync(copy).equals(before), JSON.stringify({ gone }));
  } catch (err) {
    check('documents: the Gridlines checks ran', false, err.message);
  }
}
