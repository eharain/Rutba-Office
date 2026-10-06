// Documents: View → Immersive Reader, pressed on the ribbon — its own tab
// opens and the page becomes a column for reading; Column Width, Page Color,
// Syllables, Line Focus and Text Spacing each change what is drawn; Close
// puts the pages back, editable again. The file is not touched.

import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';

/**
 * @param {object} h the harness: open, check, until, wait
 */
export async function verifyWordImmersive(h, { dir }) {
  const { open, check, until, wait } = h;
  const file = path.join(dir, 'immersive.docx');
  try {
    const words = 'Extraordinary communication needs understanding, patience and imagination. ';
    fs.writeFileSync(file, buildDocx({ styles: true, paragraphs: Array.from({ length: 30 }, (_, i) => ({ text: `${i + 1}. ${words.repeat(3)}` })) }));
    const before = fs.readFileSync(file);
    const win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const pressTab = (tab, label) => js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(tab)})?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(label)});
      if (!b) return 'no button'; if (b.disabled) return 'disabled'; b.click(); return 'clicked';
    })()`);
    const menuPick = async (button, item) => {
      await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(button)}); b?.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: 300, clientY: 110 })); return 1; })()`);
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(item)}))`), `the ${item} item`, 3000).catch(() => {});
      return js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(item)}); if (!b) return 'no item'; b.click(); return 'picked'; })()`);
    };
    const page = () => js(`(() => {
      const p = document.querySelector('.wd-page');
      const shy = p.querySelector('.wd-shy');
      return {
        width: p.offsetWidth, editable: p.isContentEditable, sheets: document.querySelectorAll('.wd-sheet').length,
        back: getComputedStyle(p).backgroundColor, letters: getComputedStyle(p).letterSpacing,
        tab: document.querySelector('.rw-tab[aria-selected="true"], .rw-tab.active, .rw-tab.on')?.textContent.trim() || null,
        tabs: [...document.querySelectorAll('.rw-tab')].map((t) => t.textContent.trim()),
        dots: shy ? getComputedStyle(shy, '::after').content : null,
      };
    })()`);

    await until(() => js(`document.querySelectorAll('.wd-sheet').length > 0`), 'the pages', 8000);
    const opened = await pressTab('View', 'Immersive Reader');
    await until(async () => (await page()).tabs.includes('Immersive Reader'), 'the reader\'s tab', 4000).catch(() => {});
    const reader = await page();
    check('documents: View → Immersive Reader opens its own tab and turns the pages into a column for reading, not editing',
      opened === 'clicked' && reader.tabs.includes('Immersive Reader') && reader.editable === false && reader.sheets === 0 && reader.width === 720 && reader.letters !== 'normal',
      JSON.stringify(reader));

    const narrow = await menuPick('Column Width', 'Narrow');
    await until(async () => (await page()).width === 560, 'a narrow column', 3000).catch(() => {});
    const sepia = await menuPick('Page Color', 'Sepia');
    await until(async () => (await page()).back === 'rgb(244, 236, 216)', 'sepia', 3000).catch(() => {});
    await pressTab('Immersive Reader', 'Syllables');
    await until(async () => (await page()).dots === '"·"', 'syllables', 3000).catch(() => {});
    await pressTab('Immersive Reader', 'Text Spacing');
    await until(async () => (await page()).letters === 'normal', 'spacing off', 3000).catch(() => {});
    const looks = await page();
    check('documents: Column Width, Page Color, Syllables and Text Spacing each change the reading column',
      narrow === 'picked' && sepia === 'picked' && looks.width === 560 && looks.back === 'rgb(244, 236, 216)' && looks.dots === '"·"' && looks.letters === 'normal',
      JSON.stringify(looks));

    const focus = await menuPick('Line Focus', 'Three Lines');
    await until(() => js(`Boolean(document.querySelector('.wd-linefocus-band'))`), 'the focus band', 3000).catch(() => {});
    const band = await js(`(() => { const b = document.querySelector('.wd-linefocus-band'); const p = document.querySelector('.wd-page .wd-block'); return { band: b ? b.offsetHeight : null, line: p ? parseFloat(getComputedStyle(p).lineHeight) : null, top: document.querySelector('.wd-scroll').scrollTop }; })()`);
    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'Down' });
    wc.sendInputEvent({ type: 'keyUp', keyCode: 'Down' });
    await wait(400);
    const moved = await js(`document.querySelector('.wd-scroll').scrollTop`);
    check('documents: Line Focus shows three lines at a time, and the arrow keys move the words through them a line at a time',
      focus === 'picked' && band.band != null && Math.abs(band.band - band.line * 3) <= 2 && moved > band.top && Math.abs(moved - band.top - band.line) <= 3,
      JSON.stringify({ focus, ...band, moved }));

    const closed = await pressTab('Immersive Reader', 'Close Immersive Reader');
    await until(async () => (await page()).sheets > 0, 'the pages back', 4000).catch(() => {});
    const back = await page();
    check('documents: Close Immersive Reader puts the pages back, editable, and the file is untouched',
      closed === 'clicked' && back.sheets > 0 && back.editable === true && !back.tabs.includes('Immersive Reader') && fs.readFileSync(file).equals(before),
      JSON.stringify({ closed, sheets: back.sheets, editable: back.editable, tabs: back.tabs }));
  } catch (err) {
    check('documents: the Immersive Reader checks ran', false, err.message);
  }
}
