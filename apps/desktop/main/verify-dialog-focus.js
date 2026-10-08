// The shared Dialog keeps focus as a modal dialog should: Documents' Insert
// Caption, opened from the ribbon, takes focus into itself; Tab and
// Shift+Tab go round inside it and never reach the page behind; Escape
// closes it and focus goes back to the button that opened it. Run alone with
// RUTBA_VERIFY_ONLY=dialogfocus.

import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';

export async function verifyDialogFocus({ open, check, until, wait, press }, { dir }) {
  const file = path.join(dir, 'dialog-focus.docx');
  try {
    fs.writeFileSync(file, buildDocx({ paragraphs: [{ text: 'A figure goes here.' }, { text: 'After it.' }] }));
    const win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="1"]'))`), 'the page', 8000);
    await js(`(() => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'References')?.click(); return 1; })()`);
    await wait(200);
    // The button is focused as a click would leave it, then pressed.
    const pressed = await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Insert Caption'); if (!b) return 'no button'; b.focus(); b.click(); return 'clicked'; })()`);
    await until(() => js(`Boolean(document.querySelector('.rw-dialog'))`), 'the dialog', 5000).catch(() => {});
    const inside = () => js(`Boolean(document.activeElement && document.querySelector('.rw-dialog')?.contains(document.activeElement))`);
    const focusedIn = await inside();
    check('dialogs: a dialog takes focus into itself when it opens', pressed === 'clicked' && focusedIn, JSON.stringify({ pressed, focusedIn, at: await js(`document.activeElement?.className || document.activeElement?.tagName`) }));

    // Tab all the way round, and Shift+Tab back from the first.
    const count = await js(`[...document.querySelectorAll('.rw-dialog button:not([disabled]), .rw-dialog input:not([disabled]), .rw-dialog select:not([disabled]), .rw-dialog textarea:not([disabled])')].filter((el) => el.getClientRects().length).length`);
    let stayed = true;
    for (let i = 0; i < count + 2; i++) {
      await press(wc, 'Tab');
      if (!(await inside())) { stayed = false; break; }
    }
    await js(`(() => { const all = [...document.querySelectorAll('.rw-dialog button:not([disabled]), .rw-dialog input:not([disabled]), .rw-dialog select:not([disabled]), .rw-dialog textarea:not([disabled])')].filter((el) => el.getClientRects().length); all[0].focus(); return 1; })()`);
    await press(wc, 'Tab', { modifiers: ['shift'] });
    const wrapped = await js(`(() => { const all = [...document.querySelectorAll('.rw-dialog button:not([disabled]), .rw-dialog input:not([disabled]), .rw-dialog select:not([disabled]), .rw-dialog textarea:not([disabled])')].filter((el) => el.getClientRects().length); return document.activeElement === all[all.length - 1]; })()`);
    check('dialogs: Tab and Shift+Tab go round inside the dialog, never into the page behind', count > 1 && stayed && wrapped, JSON.stringify({ count, stayed, wrapped }));

    // Escape closes it, and focus is back on the button that opened it.
    await press(wc, 'Escape');
    const closed = await until(() => js(`!document.querySelector('.rw-dialog')`), 'the dialog closed', 3000).then(() => true).catch(() => false);
    const back = await js(`document.activeElement?.textContent?.trim() === 'Insert Caption'`);
    check('dialogs: Escape closes the dialog and focus goes back to what opened it', closed && back, JSON.stringify({ closed, back, at: await js(`document.activeElement?.tagName + '.' + document.activeElement?.className + ' tabindex=' + document.activeElement?.getAttribute('tabindex')`) }));
  } catch (err) {
    check('dialogs: the focus checks ran', false, err?.message || String(err));
  }
}
