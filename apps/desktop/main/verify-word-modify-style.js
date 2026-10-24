// Word: Modify Style, and Update Heading 1 to Match Selection (2.18.0, A3.4).
// In 2.17.1 a style could be picked but never changed: restyling every
// heading meant doing it one heading at a time. Now a right-click on a style
// in the gallery offers both, the style itself changes, every paragraph with
// it follows on the page, and the saved file keeps it in styles.xml. Run
// alone with RUTBA_VERIFY_ONLY=wordmodifystyle.
import fs from 'node:fs';
import path from 'node:path';
import { OoxmlPackage } from '@rutba/ooxml';
import { buildDocx } from '@rutba/ooxml/build';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn
 * @param {{ dir: string }} args where the fixture is written
 */
export async function verifyWordModifyStyle({ open, check, until, wait, press, errorsIn }, { dir }) {
  const file = path.join(dir, 'modify-style.docx');
  fs.writeFileSync(file, buildDocx({
    styles: true,
    paragraphs: [
      { text: 'Opening jobs', style: 'Heading1' },
      'Unlock the door and turn the till on.',
      { text: 'Closing jobs', style: 'Heading1' },
      'Count the till and lock the door.',
    ],
  }));
  let win = null;
  try {
    win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const looks = () => js(`[0, 2].map((i) => { const el = document.querySelector('.wd-page [data-block="' + i + '"]'); const run = el?.querySelector('span') || el; const s = run ? getComputedStyle(run) : null; return s ? { size: s.fontSize, weight: s.fontWeight, style: s.fontStyle, family: s.fontFamily } : null; })`);
    const rightClickStyle = (id) => js(`(() => { const b = document.querySelector('.rw-ribbon .wd-style[data-style="${id}"]'); if (!b) return false; const r = b.getBoundingClientRect(); b.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: r.left + 8, clientY: r.top + 8 })); return true; })()`);
    const menuItems = () => js(`[...document.querySelectorAll('.rw-menu button')].map((b) => b.textContent.trim())`);
    const pick = (label) => js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim().startsWith(${JSON.stringify(label)})); b?.click(); return Boolean(b); })()`);
    const setField = (role, value) => js(`(() => {
      const el = document.querySelector('.wd-modify-style [data-role="${role}"]');
      if (!el) return false;
      if (el.type === 'checkbox') { if (el.checked !== ${JSON.stringify(Boolean(value))}) el.click(); return true; }
      const proto = el.tagName === 'SELECT' ? HTMLSelectElement.prototype : HTMLInputElement.prototype;
      Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, ${JSON.stringify(String(value))});
      el.dispatchEvent(new Event(el.tagName === 'SELECT' ? 'change' : 'input', { bubbles: true }));
      return true;
    })()`);
    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="3"]')) && Boolean(document.querySelector('.rw-ribbon .wd-style[data-style="Heading1"]'))`), 'the page and the styles gallery', 8000);
    const before = await looks();

    // Right-click Heading 1: both items, then Modify.
    await rightClickStyle('Heading1');
    await until(async () => (await menuItems()).length >= 2, 'the style menu', 3000).catch(() => false);
    const items = await menuItems();
    await pick('Modify');
    await until(() => js(`Boolean(document.querySelector('.wd-modify-style'))`), 'the Modify Style dialog', 4000);
    const title = await js(`document.querySelector('.wd-modify-style')?.closest('.rw-dialog')?.querySelector('h2, .rw-dialog-title')?.textContent || document.querySelector('.rw-dialog')?.textContent.slice(0, 40)`);
    const started = await js(`({ size: document.querySelector('.wd-modify-style [data-role="style-size"]')?.value, bold: document.querySelector('.wd-modify-style [data-role="style-bold"]')?.checked })`);
    check('word: a right-click on Heading 1 in the gallery offers Update Heading 1 to Match Selection and Modify, and Modify opens on the style as it is',
      items.some((s) => s.startsWith('Update Heading 1 to Match Selection')) && items.some((s) => s.startsWith('Modify')) && started.size === '16' && started.bold === true,
      JSON.stringify({ items, title, started }));

    await setField('style-font', 'Georgia');
    await setField('style-size', '20');
    await setField('style-bold', false);
    await setField('style-before', '18');
    await wait(150);
    await js(`[...document.querySelectorAll('.rw-dialog button')].find((b) => b.textContent.trim() === 'OK')?.click(), 1`);
    await until(() => js(`!document.querySelector('.wd-modify-style')`), 'the dialog to close', 4000).catch(() => false);
    let after = null;
    const followed = await until(async () => {
      after = await looks();
      return after.every((l) => l && Math.abs(parseFloat(l.size) - 26.67) < 0.05 && Number(l.weight) < 600 && /Georgia/.test(l.family));
    }, 'both headings to follow', 5000).then(() => true).catch(() => false);
    check('word: Modify Style changes both Heading 1 paragraphs on the page (Georgia, 20 pt, not bold)', followed, JSON.stringify({ before, after }));

    // Update to Match Selection: the first heading made italic, then the style takes it.
    await js(`(() => {
      const el = document.querySelector('.wd-page [data-block="0"]');
      const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
      const first = walker.nextNode();
      document.querySelector('.wd-page').focus();
      const r = document.createRange(); r.setStart(first, 0); r.setEnd(first, first.nodeValue.length);
      const s = getSelection(); s.removeAllRanges(); s.addRange(r);
      document.querySelector('.wd-page').dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
      return 1;
    })()`);
    await wait(200);
    await press(wc, 'i', { modifiers: ['control'] });
    await wait(300);
    await rightClickStyle('Heading1');
    await until(async () => (await menuItems()).length >= 2, 'the style menu', 3000).catch(() => false);
    await pick('Update Heading 1 to Match Selection');
    let matched = null;
    const italic = await until(async () => {
      matched = await looks();
      return matched[1]?.style === 'italic';
    }, 'the second heading italic too', 5000).then(() => true).catch(() => false);
    check('word: Update Heading 1 to Match Selection makes the other heading italic too, keeping the modified size', italic && Math.abs(parseFloat(matched[1]?.size) - 26.67) < 0.05, JSON.stringify(matched));

    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    await press(wc, 's', { modifiers: ['control'] });
    let detail = '';
    const saved = await until(() => {
      try {
        const xml = OoxmlPackage.read(fs.readFileSync(file)).text('word/styles.xml');
        const style = /<w:style\b[^>]*w:styleId="Heading1"[\s\S]*?<\/w:style>/.exec(xml)?.[0] || '';
        detail = style;
        return /w:ascii="Georgia"/.test(style) && /<w:sz w:val="40"\/>/.test(style) && /<w:i\/>/.test(style) && !/<w:b\/>/.test(style) && /w:before="360"/.test(style);
      } catch { return false; }
    }, 'the save', 6000).then(() => true).catch(() => false);
    check('word: the saved styles.xml carries the changed Heading 1 (Georgia, 40 half points, italic, not bold, 18 pt before)', saved, detail);

    const complaints = await errorsIn(win);
    check('word: the Modify Style checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the Modify Style checks ran', false, err?.message || String(err));
  }
}
