// The menu behind the app mark, in 2.22's order (D-2026-11-02-09: "menu in
// 2.22"): in every app the same rows in the same order, New and Open first,
// then the document's own (Save, Save As, Export, Print), then the app's own,
// then Window, Settings and Help, and Close Window last; no group empty.
// Settings opens its submenu by the keyboard alone, the arrow back closes it
// with the keyboard on Settings again; in Urdu the submenu opens with the
// left arrow and sits to the left. Open Recent has this app's files first.
// Run alone with RUTBA_VERIFY_ONLY=app-menu.

import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';

const APPS = ['home', 'word', 'sheets', 'slides', 'pdf', 'base', 'mail', 'calendar', 'contacts', 'pictures', 'image', 'video', 'recorder', 'talk'];

/** The rows of the menu open now: its own items and separators, not a submenu's. */
const ROWS = `(() => {
  const m = document.querySelector('.rw-menu:not(.rw-submenu)');
  if (!m) return null;
  return [...m.children].filter((n) => !n.classList.contains('rw-menu')).map((n) => n.tagName === 'HR' ? '-' : (n.querySelector('.rw-menu-label')?.textContent || n.textContent).trim());
})()`;

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args the run's scratch folder
 */
export async function verifyAppMenu(h, { dir }) {
  const { open, check, until, wait } = h;
  const press = (wc, key) => { wc.sendInputEvent({ type: 'keyDown', keyCode: key }); wc.sendInputEvent({ type: 'keyUp', keyCode: key }); };
  const openMenu = (js) => js(`(() => { const b = document.querySelector('.rw-appmark'); if (!b) return false; b.focus(); b.dispatchEvent(new MouseEvent('click', { bubbles: true, detail: 0 })); return true; })()`);
  const docFile = path.join(dir, 'menu-recent.docx');
  fs.writeFileSync(docFile, buildDocx({ styles: true, paragraphs: [{ text: 'A file for the recent list.' }] }));

  // Every app: the fixed order, no empty group.
  const seen = {};
  const wrong = [];
  for (const app of APPS) {
    try {
      const win = await open(app, app === 'word' ? docFile : null);
      const js = (code) => win.webContents.executeJavaScript(code);
      await until(() => js(`Boolean(document.querySelector('.rw-appmark'))`), `${app}'s app mark`, 15000);
      await wait(400);
      await openMenu(js);
      const rows = await until(() => js(ROWS), `${app}'s menu`, 5000).then(() => js(ROWS), () => null);
      seen[app] = rows;
      await js(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })), 1`);
      if (!rows) { wrong.push(`${app}: no menu`); continue; }
      const at = (label) => rows.indexOf(label);
      const tail = ['Window', 'Settings', 'Help', ...(app === 'home' ? [] : ['Close Window'])];
      if (tail.some((l) => at(l) < 0) || tail.some((l, i) => i && at(l) < at(tail[i - 1]))) wrong.push(`${app}: Window, Settings, Help, Close Window not all there in order`);
      if (app !== 'home' && rows[rows.length - 1] !== 'Close Window') wrong.push(`${app}: Close Window is not last`);
      const firstSave = rows.findIndex((r) => /^Save|^Print|^Export|^Merge/.test(r));
      const opened = Math.max(at('Open…'), at('Open Recent'));
      if (firstSave >= 0 && opened >= 0 && firstSave < opened) wrong.push(`${app}: Save or Print before Open`);
      if (firstSave >= 0 && firstSave > at('Window')) wrong.push(`${app}: Save or Print after Window`);
      if (rows[0] === '-' || rows[rows.length - 1] === '-' || rows.some((r, i) => r === '-' && rows[i + 1] === '-')) wrong.push(`${app}: an empty group`);
      if (app !== 'home') win.close();
    } catch (err) {
      wrong.push(`${app}: ${err.message}`);
    }
  }
  check('app menu: in every app the same order, New and Open, the document\'s own, the app\'s own, then Window, Settings, Help and Close Window last, with no empty group',
    wrong.length === 0, wrong.length ? `${wrong.join(' | ')} ${JSON.stringify(seen.word)}` : JSON.stringify({ word: seen.word, mail: seen.mail }));

  // Open Recent in Sheets: the files of other apps (the Word file opened above) in their own submenu.
  try {
    const win = await open('sheets');
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.rw-appmark'))`), 'the app mark', 15000);
    await openMenu(js);
    await until(() => js(ROWS), 'the menu', 5000);
    await js(`[...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === 'Open Recent')?.click(), 1`);
    const listOf = (sel) => `(() => { const s = document.querySelector('${sel}'); return s ? [...s.querySelectorAll(':scope > button')].map((b) => b.textContent.trim()) : null; })()`;
    const recent = await until(() => js(listOf('.rw-submenu')), 'Open Recent', 4000).then(() => js(listOf('.rw-submenu')), () => null);
    await js(`[...document.querySelectorAll('.rw-submenu button')].find((b) => b.textContent.trim() === 'Other apps’ files')?.click(), 1`);
    const others = await until(() => js(listOf('.rw-submenu .rw-submenu')), 'the other apps files', 4000).then(() => js(listOf('.rw-submenu .rw-submenu')), () => null);
    await js(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })), 1`);
    check('app menu: Open Recent puts other apps\' files in a submenu of their own, the Word file among them',
      Boolean(recent) && recent.includes('Other apps’ files') && recent.includes('Clear the List') && !recent.includes('menu-recent.docx') && Boolean(others?.includes('menu-recent.docx')),
      JSON.stringify({ recent, others }));
    win.close();
  } catch (err) {
    check('app menu: the Open Recent pass ran', false, err.message);
  }

  // Help, Keyboard Shortcuts in Word: the keys the window has, Save's among them.
  try {
    const win = await open('word', docFile);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.wd-page'))`), 'the page', 15000);
    await openMenu(js);
    await until(() => js(ROWS), 'the menu', 5000);
    await js(`[...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === 'Help')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.rw-submenu'))`), 'Help', 4000).catch(() => {});
    await js(`[...document.querySelectorAll('.rw-submenu button')].find((b) => b.textContent.trim() === 'Keyboard Shortcuts')?.click(), 1`);
    const sheet = await until(() => js(`document.querySelectorAll('.rw-shortcuts-list tr').length > 5`), 'the shortcuts', 5000)
      .then(() => js(`[...document.querySelectorAll('.rw-shortcuts-list tr')].map((r) => [...r.cells].map((c) => c.textContent.trim()).join(' = '))`), () => null);
    await js(`document.querySelector('.rw-shortcuts .rw-dialog-foot button')?.click(), 1`);
    check('app menu: Help, Keyboard Shortcuts lists the keys the window has, Save as Ctrl+S among them',
      Array.isArray(sheet) && sheet.includes('Save = Ctrl+S') && sheet.length > 10, JSON.stringify(sheet?.slice(0, 12)));
    win.close();
  } catch (err) {
    check('app menu: the Keyboard Shortcuts pass ran', false, err.message);
  }

  // The keyboard alone, left to right and then right to left.
  for (const [lang, toward, away] of [[null, 'Right', 'Left'], ['ur', 'Left', 'Right']]) {
    let host = null;
    let was = null;
    try {
      host = await open('calendar');
      const hjs = (code) => host.webContents.executeJavaScript(code);
      await until(() => hjs(`Boolean(window.rutbaOffice)`), 'the window', 15000);
      was = await hjs(`window.rutbaOffice.store.get({ key: 'language', fallback: null })`);
      if (lang) await hjs(`window.rutbaOffice.store.set({ key: 'language', value: ${JSON.stringify(lang)} }).then(() => 1)`);
      const win = await open('contacts');
      const wc = win.webContents;
      const js = (code) => wc.executeJavaScript(code);
      await until(() => js(`Boolean(document.querySelector('.rw-appmark'))`), 'the app mark', 15000);
      win.focus();
      await wait(300);
      await openMenu(js);
      await until(() => js(`document.activeElement?.closest('.rw-menu') != null`), 'the keyboard in the menu', 5000);
      // Down to Settings, by the keys alone: of Window, Settings and Help (in any language), the middle one.
      for (let i = 0; i < 30; i++) {
        if (await js(`(() => { const subs = [...document.querySelectorAll('.rw-menu:not(.rw-submenu) > button[data-sub]')]; return document.activeElement === subs[subs.length - 2]; })()`)) break;
        press(wc, 'Down');
        await wait(60);
      }
      const onSettings = await js(`document.activeElement?.getAttribute('aria-haspopup') === 'menu'`);
      press(wc, toward);
      await until(() => js(`document.activeElement?.closest('.rw-submenu') != null`), 'the submenu', 3000).catch(() => {});
      const inSub = await js(`(() => { const s = document.querySelector('.rw-submenu'); const p = document.querySelector('.rw-menu:not(.rw-submenu)'); if (!s || !p) return null; const a = s.getBoundingClientRect(); const b = p.getBoundingClientRect(); return { focused: s.contains(document.activeElement), label: s.getAttribute('aria-label'), side: a.left >= b.right - 12 ? 'right' : a.right <= b.left + 12 ? 'left' : 'over', dir: document.documentElement.dir, expanded: document.querySelector('.rw-menu:not(.rw-submenu) > button[aria-expanded="true"]') != null }; })()`);
      press(wc, away);
      await wait(300);
      const backOn = await js(`({ gone: !document.querySelector('.rw-submenu'), on: document.activeElement?.getAttribute('aria-haspopup') === 'menu', open: Boolean(document.querySelector('.rw-menu')) })`);
      press(wc, 'Escape');
      await wait(300);
      const closed = await js(`!document.querySelector('.rw-menu')`);
      const side = lang ? 'left' : 'right';
      check(`app menu: ${lang ? 'right to left, in Urdu' : 'left to right'}, the keyboard alone opens a submenu with the ${toward.toLowerCase()} arrow, beside the menu on the ${side}, and the ${away.toLowerCase()} arrow closes it with the keyboard on its item again`,
        onSettings && inSub?.focused && inSub.side === side && inSub.expanded && (lang ? inSub.dir === 'rtl' : true) && backOn.gone && backOn.on && backOn.open && closed,
        JSON.stringify({ onSettings, inSub, backOn, closed }));
      win.close();
    } catch (err) {
      check(`app menu: the keyboard pass${lang ? ' in Urdu' : ''} ran`, false, err.message);
    } finally {
      if (host && !host.isDestroyed()) {
        await host.webContents.executeJavaScript(`window.rutbaOffice.store.set({ key: 'language', value: ${JSON.stringify(was ?? null)} }).then(() => 1)`).catch(() => {});
        host.close();
      }
    }
  }
}
