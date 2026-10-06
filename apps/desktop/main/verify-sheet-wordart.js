// Worksheets: Insert → WordArt, pressed on the ribbon — the gallery shows
// each style on an "A"; the style asks for the words, showing the look as
// they are typed; the WordArt goes over the sheet with its letters outlined
// and its words glowing; a double-click opens its words again and changing
// them keeps the look; and the saved file keeps it as Excel writes WordArt.

import fs from 'node:fs';
import path from 'node:path';
import { OoxmlPackage } from '@rutba/ooxml';
import { buildXlsx } from '@rutba/ooxml/build';

/**
 * @param {object} h the harness: open, check, until, wait, capture
 */
export async function verifySheetWordArt(h, { dir }) {
  const { open, check, until, wait, capture } = h;
  const file = path.join(dir, 'wordart.xlsx');
  try {
    fs.writeFileSync(file, buildXlsx({ sheets: [{ name: 'Sheet1', rows: [['Quarter', 'Sales'], ['Q1', 120], ['Q2', 140]] }] }));
    const win = await open('sheets', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const type = (text) => js(`(() => {
      const el = document.querySelector('.sh-words-text');
      if (!el) return 'no box';
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(el, ${JSON.stringify(text)});
      el.dispatchEvent(new Event('input', { bubbles: true }));
      return 'typed';
    })()`);
    const drawn = () => js(`(() => {
      const d = [...document.querySelectorAll('.sh-drawing[data-kind="shape"]')].find((n) => n.querySelector('text'));
      if (!d) return null;
      const t = [...d.querySelectorAll('text')];
      return { words: t.map((n) => n.textContent).join(' | '), stroke: t[0].getAttribute('stroke'), filter: t[0].getAttribute('style') || '' };
    })()`);

    await until(() => js(`Boolean(document.querySelector('.rw-tab'))`), 'the ribbon', 8000);
    const opened = await js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'WordArt');
      if (!b || b.disabled) return 'no button';
      b.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: 600, clientY: 110 }));
      await new Promise((r) => setTimeout(r, 250));
      return [...document.querySelectorAll('.rw-menu button')].filter((n) => n.querySelector('.rw-menu-preview span')).length;
    })()`);
    await js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.includes('Fill: White; Outline: Blue')); if (!b) return 'no item'; b.click(); return 'picked'; })()`);
    await until(() => js(`Boolean(document.querySelector('.sh-words-text'))`), 'the words box', 4000).catch(() => {});
    const asked = await js(`document.querySelector('.sh-words-text')?.value || ''`);
    await type('Quarterly sales');
    const preview = await js(`(() => { const s = document.querySelector('.sh-words-preview span'); return s ? { text: s.textContent, stroke: getComputedStyle(s).webkitTextStrokeColor } : null; })()`);
    if (capture) await capture(win, 'sheet-wordart-words.png');
    await js(`document.querySelector('.sh-words-ok')?.click(), 1`);
    await until(async () => (await drawn())?.words === 'Quarterly sales', 'the WordArt on the sheet', 6000).catch(() => {});
    const first = await drawn();
    check('worksheets: Insert → WordArt shows the gallery, asks for the words with the look shown, and sets them over the sheet outlined and glowing',
      opened >= 6 && asked === 'Your text here' && preview?.text === 'Quarterly sales' && first?.words === 'Quarterly sales' && first.stroke === '#4472c4' && /drop-shadow/.test(first.filter),
      JSON.stringify({ opened, asked, preview, first }));
    if (capture) await capture(win, 'sheet-wordart.png');

    // A double-click opens the words again; two lines now, the look kept.
    await js(`(() => { [...document.querySelectorAll('.sh-drawing[data-kind="shape"]')].find((n) => n.querySelector('text'))?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true })); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.sh-words-text'))`), 'the words again', 4000).catch(() => {});
    const prefilled = await js(`document.querySelector('.sh-words-text')?.value || ''`);
    await type('Quarterly sales\n2026');
    await js(`document.querySelector('.sh-words-ok')?.click(), 1`);
    const changed = await until(async () => (await drawn())?.words === 'Quarterly sales | 2026', 'the changed words', 6000).then(() => true).catch(() => false);
    const after = await drawn();
    check('worksheets: a double-click opens the WordArt\'s words again, and new words, two lines of them, keep its look',
      prefilled === 'Quarterly sales' && changed && after?.stroke === '#4472c4', JSON.stringify({ prefilled, changed, after }));

    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => {
      try {
        const pkg = OoxmlPackage.read(fs.readFileSync(file));
        const part = pkg.partNames().find((n) => /^xl\/drawings\/drawing\d+\.xml$/.test(n));
        const xml = part ? pkg.text(part) : '';
        return /<xdr:cNvSpPr txBox="1"\/>/.test(xml) && /<a:rPr lang="en-US" sz="3600" b="0" cap="none" spc="0"><a:ln w="12700">/.test(xml) && /<a:glow rad="63500">/.test(xml) && xml.includes('<a:t>2026</a:t>');
      } catch { return false; }
    }, 'the saved WordArt', 8000).then(() => true).catch(() => false);
    check('worksheets: the saved file keeps the WordArt as Excel writes it — a text box, its words\' outline and glow in the run', saved, saved ? 'txBox with a:ln and a:glow in the run' : 'not in the file');
    await wait(50);
  } catch (err) {
    check('worksheets: the WordArt checks ran', false, err.message);
  }
}
