// Every app from the keyboard alone, as a person who does not use a mouse
// works it. In each app's window: F6 goes round the ribbon, the panes, the
// document and the status bar and comes back where it started, Shift+F6 the
// other way, as Office's F6 does; Tab and Shift+Tab move through a pane and
// stay in it to its end; the app menu opens from the keyboard with the
// keyboard on its first item, the arrows go down it and Escape puts it away
// with the keyboard back on the button; a dialog opened by its shortcut takes
// the keyboard, Tab goes round inside it, and Escape closes it with the
// keyboard back where it was, never on nothing. Every step reads where the
// keyboard is with until(...), the window's own answer. Keys go to the
// window's page only: nothing touches the desktop's keyboard. Run alone with
// RUTBA_VERIFY_ONLY=keyboard-pass; captures keyboard-<app>.png.

import fs from 'node:fs';
import path from 'node:path';
import { fixtures } from './verify-polish-pass.js';

/** Where the keyboard is, in words a failure can be read by. */
const WHERE = `(() => {
  const a = document.activeElement;
  if (!a || a === document.body) return { part: 'nothing', at: 'body' };
  const pane = a.closest('.rw-body .rw-panel');
  const part = a.closest('.rw-dialog') ? 'dialog' : a.closest('.rw-menu') ? 'menu' : a.closest('.rw-ribbon') ? 'ribbon' : a.closest('.rw-status') ? 'status' : pane && pane.querySelector('[data-f6-home]') ? 'document' : pane ? 'pane:' + (a.closest('.rw-panel').getAttribute('aria-label') || a.closest('.rw-panel').dataset.panel || '?') : a.closest('.rw-body') ? 'document' : a.closest('.rw-titlebar') ? 'title' : 'other';
  const name = (a.getAttribute('aria-label') || a.dataset.tip || a.textContent || a.className || a.tagName).trim().slice(0, 30);
  return { part, at: name, ring: a.matches(':focus-visible') };
})()`;

export async function verifyKeyboardPass({ open, check, until, wait, press }, { dir }) {
  const shots = process.env.RUTBA_VERIFY_CAPTURE;
  const only = (process.env.RUTBA_A11Y_APPS || '').split(',').map((x) => x.trim()).filter(Boolean);
  const wanted = (app) => !only.length || only.includes(app);
  const fx = fixtures(dir, 'keyboard-fixtures');
  const snap = async (win, name) => {
    if (!shots) return;
    win.webContents.invalidate();
    await wait(700);
    fs.writeFileSync(path.join(shots, name), (await win.webContents.capturePage()).toPNG());
  };

  /** What `fn` answers once it answers something, or null when it never does. */
  const value = async (fn, what, ms = 2000) => { let v = null; await until(async () => (v = await fn()), what, ms).catch(() => {}); return v; };

  const pass = async (app, win, { dialogKey = null, dialogSel = '.rw-dialog' } = {}) => {
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const where = () => js(WHERE);
    // Enter presses a button with its character, as a keyboard sends it.
    const key = (k, modifiers = []) => press(wc, k, { modifiers, char: k === 'Return' });
    // Keys reach a page that has the keyboard: Chromium draws no focus in a page without it.
    if (!(await js(`document.hasFocus()`))) wc.focus();
    // Start in the document, as a person does once the file is open.
    await js(`(() => { const d = (window.__f6 || null); const b = document.querySelector('.rw-body [data-f6-home]') || document.querySelector('.rw-body [tabindex="0"], .rw-body [contenteditable="true"]'); b?.focus(); return 1; })()`);
    await until(() => where().then((w) => w.part !== 'nothing'), `the ${app} window to take the keyboard`, 3000).catch(() => {});
    // A window with nothing of its own to start on (Home, Pictures' grid): F6 is the way in.
    if ((await where()).part === 'nothing') await key('F6');

    // F6 round the parts and back to the start; Shift+F6 the other way.
    const start = await where();
    const round = [start.part];
    for (let i = 0; i < 8; i++) {
      const before = round[round.length - 1];
      await key('F6');
      const w = (await value(async () => { const x = await where(); return x.part !== 'nothing' ? x : null; }, 'the keyboard somewhere after F6')) || { part: 'nothing' };
      round.push(w.part);
      if (w.part === start.part && i > 0) break;
      if (w.part === before && w.part !== 'nothing') break;
    }
    const parts = new Set(round);
    const back = [];
    for (let i = 0; i < round.length - 1; i++) {
      await key('F6', ['shift']);
      back.push(((await value(async () => { const x = await where(); return x.part !== 'nothing' ? x : null; }, 'the keyboard after Shift+F6')) || { part: 'nothing' }).part);
    }
    const forward = round.slice(1);
    const reversed = [...forward].reverse().slice(1).concat(start.part);
    // The parts this window has: Home has no ribbon.
    const [hasRibbon, hasStatus] = await js(`[Boolean(document.querySelector('.rw-ribbon')), Boolean(document.querySelector('.rw-status'))]`);
    check(`keyboard-pass: ${app}: F6 goes round ${hasRibbon ? 'the ribbon, ' : ''}the document${hasStatus ? ', the status bar' : ''}${[...parts].some((p) => p.startsWith('pane:')) ? ' and the panes' : ''} and back to the start, and Shift+F6 the same way back`,
      (!hasRibbon || parts.has('ribbon')) && parts.has('document') && (!hasStatus || parts.has('status')) && !parts.has('nothing') && round[round.length - 1] === start.part && back.join(',') === reversed.join(','),
      JSON.stringify({ round, back }));
    await snap(win, `keyboard-${app}.png`);

    // Tab and Shift+Tab through the first pane with things in it: each press moves on, none lands on nothing.
    const pane = await js(`(() => { const p = [...document.querySelectorAll('.rw-body .rw-panel')].find((x) => x.getClientRects().length && x.querySelectorAll('button:not([disabled]), input:not([disabled]), select, [tabindex="0"]').length >= 2); if (!p) return null; const f = [...p.querySelectorAll('button:not([disabled]), input:not([disabled]), select, [tabindex="0"]')].find((n) => n.getClientRects().length); f?.focus(); return p.getAttribute('aria-label') || p.dataset.panel || 'a pane'; })()`);
    if (pane) {
      const seen = [];
      for (let i = 0; i < 3; i++) { await key('Tab'); seen.push(await where()); }
      for (let i = 0; i < 2; i++) { await key('Tab', ['shift']); seen.push(await where()); }
      const lost = seen.filter((w) => w.part === 'nothing');
      const moved = new Set(seen.map((w) => w.part + '|' + w.at)).size > 1;
      check(`keyboard-pass: ${app}: Tab and Shift+Tab move through ${pane} and never leave the keyboard on nothing`, !lost.length && moved, JSON.stringify(seen.map((w) => `${w.part}:${w.at}`)));
    }

    // The app menu from the keyboard.
    await js(`document.querySelector('.rw-appmark')?.focus(), 1`);
    await key('Return');
    const menuOpen = (await value(() => js(`document.activeElement?.closest('.rw-menu') ? document.activeElement.textContent.trim() : ''`), 'the keyboard on the menu', 3000)) || '';
    await key('Down');
    const second = (await value(() => js(`(() => { const a = document.activeElement; return a?.closest('.rw-menu') && a.textContent.trim() !== ${JSON.stringify(menuOpen || '')} ? a.textContent.trim() : ''; })()`), 'the next item')) || '';
    const role = await js(`[document.querySelector('.rw-menu')?.getAttribute('role'), document.activeElement?.getAttribute('role')].join('/')`);
    await key('Escape');
    const menuBack = await until(() => js(`!document.querySelector('.rw-menu') && document.activeElement?.matches('.rw-appmark')`), 'the keyboard back on the app button', 2000).then(() => true).catch(() => false);
    check(`keyboard-pass: ${app}: the app menu opens from the keyboard on its first item, Down moves on, Escape puts it away with the keyboard back on its button`,
      Boolean(menuOpen) && Boolean(second) && role === 'menu/menuitem' && menuBack, JSON.stringify({ menuOpen, second, role, menuBack, now: await where() }));

    // A dialog by its shortcut, from the document.
    if (dialogKey) {
      await js(`(() => { const b = document.querySelector('.rw-body [data-f6-home]') || document.querySelector('.rw-body [tabindex="0"], .rw-body [contenteditable="true"]'); b?.focus(); return 1; })()`);
      const from = await where();
      await key(dialogKey, ['control']);
      const inside = await until(() => js(`Boolean(document.activeElement?.closest(${JSON.stringify(dialogSel)}))`), 'the keyboard in the dialog', 4000).then(() => true).catch(() => false);
      const count = await js(`[...(document.querySelector(${JSON.stringify(dialogSel)})?.querySelectorAll('button:not([disabled]), input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex="0"]') || [])].filter((n) => n.getClientRects().length).length`);
      let stayed = true;
      for (let i = 0; i < Math.min(count + 1, 14); i++) {
        await key('Tab');
        if (!(await js(`Boolean(document.activeElement?.closest(${JSON.stringify(dialogSel)}))`))) { stayed = false; break; }
      }
      await key('Escape');
      const closed = await until(() => js(`!document.querySelector(${JSON.stringify(dialogSel)})`), 'the dialog closed', 3000).then(() => true).catch(() => false);
      const after = (await value(async () => { const x = await where(); return x.part !== 'nothing' ? x : null; }, 'the keyboard somewhere after the dialog')) || { part: 'nothing' };
      check(`keyboard-pass: ${app}: Ctrl+${dialogKey}'s dialog takes the keyboard, Tab stays inside it, Escape closes it and the keyboard goes back to the ${from.part}`,
        inside && stayed && closed && after.part === from.part, JSON.stringify({ from, inside, count, stayed, closed, after }));
    }
  };

  const ready = (win, sel, what) => until(() => win.webContents.executeJavaScript(`Boolean(document.querySelector(${JSON.stringify(sel)}))`), what, 15000).catch(() => {});
  const apps = [
    ['home', undefined, '.home-card', {}],
    ['word', fx.docx, '.wd-page', { dialogKey: 'P' }],
    ['sheets', fx.xlsx, '.sh-cell', { dialogKey: 'K' }],
    ['slides', fx.pptx, '.sl-thumb', {}],
    ['pdf', fx.pdf, '.pdf-page canvas[data-drawn]', {}],
    ['pictures', fx.pics, '.rw-ribbon .rw-tab', {}],
    ['image', fx.png, '.rw-ribbon .rw-tab', {}],
    ['video', fx.png, '.vd-clip', {}],
    ['contacts', fx.vcf, '.ct-item', {}],
    ['calendar', fx.ics, '.cal-month-grid, .cal-time', { dialogKey: 'N' }],
    ['mail', undefined, '.ml-row, .ml-list, .rw-empty, .ml-compose-cta', {}],
  ];
  for (const [app, file, sel, opts] of apps.filter(([a]) => wanted(a))) {
    try {
      const win = await open(app, file);
      await ready(win, sel, `the ${app} window`);
      await wait(500);
      await pass(app, win, opts);
    } catch (err) {
      check(`keyboard-pass: the ${app} pass ran`, false, err.stack || err.message);
    }
  }
}
