// Worksheets: Insert → Equation, pressed on the ribbon — typed in the
// equation editor, set over the sheet as math the page draws, opened again
// by a double-click and changed, and kept in the saved file as Excel keeps
// an equation.

import fs from 'node:fs';
import path from 'node:path';
import { OoxmlPackage } from '@rutba/ooxml';
import { buildXlsx } from '@rutba/ooxml/build';

/**
 * @param {object} h the harness: open, check, until, wait
 */
export async function verifySheetEquation(h, { dir }) {
  const { open, check, until, wait } = h;
  const file = path.join(dir, 'equation.xlsx');
  try {
    fs.writeFileSync(file, buildXlsx({ sheets: [{ name: 'Sheet1', rows: [['Area of a circle']] }] }));
    const win = await open('sheets', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const type = (text) => js(`(() => {
      const el = document.querySelector('.wd-eq-input');
      if (!el) return 'no editor';
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(el, ${JSON.stringify(text)});
      el.dispatchEvent(new Event('input', { bubbles: true }));
      return 'typed';
    })()`);
    const drawn = () => js(`(() => {
      const d = document.querySelector('.sh-drawing[data-kind="equation"]');
      const m = d?.querySelector('math');
      return d ? { math: Boolean(m), wide: m ? m.getBoundingClientRect().width : 0, sqrt: Boolean(d.querySelector('msqrt')) } : null;
    })()`);

    await until(() => js(`Boolean(document.querySelector('.rw-tab'))`), 'the ribbon', 8000);
    const pressed = await js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Equation');
      if (!b) return 'no button'; if (b.disabled) return 'disabled'; b.click(); return 'clicked';
    })()`);
    await until(() => js(`Boolean(document.querySelector('.wd-eq-input'))`), 'the equation editor', 4000).catch(() => {});
    const typed = await type('A=\\pi r^2');
    await until(() => js(`!document.querySelector('.wd-eq-ok')?.disabled`), 'a readable equation', 3000).catch(() => {});
    await js(`document.querySelector('.wd-eq-ok')?.click(), 1`);
    await until(async () => (await drawn())?.math === true, 'the equation on the sheet', 6000).catch(() => {});
    const first = await drawn();
    check('worksheets: Insert → Equation sets the typed equation over the sheet as math the page draws',
      pressed === 'clicked' && typed === 'typed' && first?.math === true && first.wide > 10, JSON.stringify({ pressed, typed, first }));

    // Opened again by a double-click, changed.
    await js(`(() => { document.querySelector('.sh-drawing[data-kind="equation"]')?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true })); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.wd-eq-input'))`), 'the editor again', 4000).catch(() => {});
    const prefilled = await js(`document.querySelector('.wd-eq-input')?.value || ''`);
    await type('r=\\sqrt(A/\\pi)');
    await until(() => js(`!document.querySelector('.wd-eq-ok')?.disabled`), 'a readable equation', 3000).catch(() => {});
    await js(`document.querySelector('.wd-eq-ok')?.click(), 1`);
    const changed = await until(async () => (await drawn())?.sqrt === true, 'the changed equation', 6000).then(() => true).catch(() => false);
    const count = await js(`document.querySelectorAll('.sh-drawing[data-kind="equation"]').length`);
    check('worksheets: a double-click opens the equation again with what it says, and changing it keeps one equation',
      /\\pi|π/.test(prefilled) && /r\^2/.test(prefilled) && changed && count === 1, JSON.stringify({ prefilled, changed, count }));

    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => {
      try {
        const pkg = OoxmlPackage.read(fs.readFileSync(file));
        const part = pkg.partNames().find((n) => /^xl\/drawings\/drawing\d+\.xml$/.test(n));
        return Boolean(part) && /<a14:m><m:oMathPara\b[\s\S]*<m:rad>/.test(pkg.text(part));
      } catch {
        return false;
      }
    }, 'the saved equation', 8000).then(() => true).catch(() => false);
    check('worksheets: the saved workbook keeps the equation as Excel keeps one', saved, saved ? 'in the drawing part' : 'not in the file');
    await wait(100);
  } catch (err) {
    check('worksheets: the Equation checks ran', false, err.message);
  }
}
