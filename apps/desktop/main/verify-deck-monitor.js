// Presentations: Slide Show → Monitor — the list names Automatic, the
// primary screen and every screen there is; the choice is kept for the
// computer; a show started after Primary Monitor takes the primary screen,
// and with Presenter View on Automatic takes another screen than the
// presenter's when there is one. A check run keeps its windows off every
// screen, so the screen is read from what the show chose, not from where
// the window went.

import fs from 'node:fs';
import path from 'node:path';
import { screen, BrowserWindow } from 'electron';
import { buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait
 */
export async function verifyDeckMonitor(h, { dir }) {
  const { open, check, until, wait } = h;
  const file = path.join(dir, 'monitor.pptx');
  try {
    fs.writeFileSync(file, buildPptx({ title: 'Monitor', slides: [1, 2].map((n) => ({ layout: 'title', title: `Slide ${n}` })) }));
    const win = await open('slides', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const button = () => js(`[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim().startsWith('Monitor:'))?.textContent.trim() || null`);
    const tab = () => js(`(async () => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Slide Show')?.click(); await new Promise((r) => setTimeout(r, 200)); return 1; })()`);
    const menuItems = () => js(`[...document.querySelectorAll('.rw-menu button')].map((b) => b.textContent.trim())`);
    const pressLabel = (label) => js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); if (!b) return 'no button'; b.click(); return 'clicked'; })()`);
    const openMenu = () => js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim().startsWith('Monitor:')); b?.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: 300, clientY: 120 })); return 1; })()`);
    const pick = (label) => js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); if (!b) return 'no item'; b.click(); return 'picked'; })()`);
    const presenting = () => js(`Boolean(document.querySelector('.sl-present'))`);
    const endShow = async () => {
      win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'Escape' });
      win.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'Escape' });
      await until(async () => !(await presenting()), 'the show to end', 4000).catch(() => {});
      await until(() => !win.isFullScreen(), 'full screen to end', 4000).catch(() => {});
    };

    await tab();
    const first = await button();
    await openMenu();
    await until(async () => (await menuItems()).length >= 3, 'the Monitor list', 3000).catch(() => {});
    const listed = await menuItems();
    const screens = screen.getAllDisplays().length;
    const picked = await pick('Primary Monitor');
    await until(async () => (await button()) === 'Monitor: Primary Monitor', 'the button to say so', 3000).catch(() => {});
    const now = await button();
    const kept = await js(`window.rutbaOffice.store.get({ key: 'slides.monitor', fallback: null })`);
    check('presentations: Slide Show → Monitor lists Automatic, the primary screen and every screen, and keeps the choice for this computer',
      first === 'Monitor: Automatic' && listed[0] === 'Automatic' && listed[1] === 'Primary Monitor' && listed.length === 2 + screens
        && picked === 'picked' && now === 'Monitor: Primary Monitor' && kept === 'primary',
      JSON.stringify({ first, listed, screens, picked, now, kept }));

    // A show from the beginning takes the primary screen.
    win.rutbaShowDisplay = null;
    await pressLabel('From Beginning');
    await until(() => presenting(), 'the show', 5000).catch(() => {});
    await until(() => win.rutbaShowDisplay != null, 'the show\'s screen', 3000).catch(() => {});
    const onPrimary = win.rutbaShowDisplay === String(screen.getPrimaryDisplay().id);
    const shown = win.rutbaShowDisplay;
    await endShow();
    check('presentations: with Monitor: Primary Monitor the show plays on the primary screen', onPrimary, JSON.stringify({ shown, primary: screen.getPrimaryDisplay().id }));

    // Automatic, with Presenter View: another screen than the presenter's, when there is one.
    await openMenu();
    await until(async () => (await menuItems()).includes('Automatic'), 'the list again', 3000).catch(() => {});
    await pick('Automatic');
    await until(async () => (await button()) === 'Monitor: Automatic', 'Automatic', 3000).catch(() => {});
    win.rutbaShowDisplay = null;
    const before = new Set(BrowserWindow.getAllWindows().map((w) => w.id));
    await pressLabel('Use Presenter View');
    await until(() => win.rutbaShowDisplay != null, 'the show\'s screen', 5000).catch(() => {});
    const presenter = BrowserWindow.getAllWindows().find((w) => !before.has(w.id));
    const presenterScreen = presenter ? String(screen.getDisplayMatching(presenter.getBounds()).id) : null;
    const automatic = win.rutbaShowDisplay;
    const right = screens > 1 ? automatic != null && automatic !== presenterScreen : automatic === String(screen.getDisplayMatching(win.getBounds()).id);
    // The presenter, opened first now, still draws the show and the show runs.
    const drawn = presenter ? await until(() => presenter.webContents.executeJavaScript(`Boolean(document.querySelector('.pv-stage svg'))`), 'the presenter\'s stage', 6000).then(() => true).catch(() => false) : false;
    const running = await presenting();
    await endShow();
    check('presentations: with Monitor: Automatic and Presenter View, the show takes another screen than the presenter\'s when there is one',
      Boolean(presenter) && right && drawn && running, JSON.stringify({ screens, automatic, presenterScreen, drawn, running }));
    await wait(100);
  } catch (err) {
    check('presentations: the Monitor checks ran', false, err.message);
  }
}
