// Home's two sets (plan B1, B2): Office and Studio side by side at a wide
// window, stacked, Office first, at a narrow one, a set alone across the
// width; each column's tiles and starts in the registry's order; Choose apps
// and the Window menu with the same two names in the same order, the macOS
// list with a line between the sets; and Studio's starts each opening its app.
// Run alone with RUTBA_VERIFY_ONLY=home-sets.

import { BrowserWindow, Menu } from 'electron';
import { APPS, APP_GROUPS, OFF_UNTIL_CHOSEN } from '@rutba/office-formats/registry';
import { labelKeyed } from './verify-keys.js';

export async function verifyHomeSets({ open, check, until, wait, errorsIn, capture }) {
  const ok = (p) => p.then(() => true).catch(() => false);
  let home = null;
  let js = null;
  const all = { off: [], on: OFF_UNTIL_CHOSEN, addons: ['ai'] };
  const setChoice = async (choice) => {
    await js(`window.rutbaOffice.store.set({ key: 'apps.choice', value: ${JSON.stringify(choice)} }).then(() => 1)`);
    await home.webContents.reload();
    await until(() => js(`document.querySelectorAll('.home-card').length > 0`), 'the launcher after the choice', 10000);
  };
  const shot = async (name) => {
    await wait(500);
    home.webContents.invalidate();
    await js(`new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(() => r(1))))`).catch(() => {});
    await wait(350);
    await capture?.(home, name);
  };
  // Where each set is, and how many tiles a row of it holds.
  const layout = () => js(`(() => {
    const body = document.querySelector('.home').getBoundingClientRect();
    const sets = [...document.querySelectorAll('.home-set')].map((s) => {
      const r = s.getBoundingClientRect();
      const tiles = [...s.querySelectorAll('.home-card')];
      const top = tiles[0]?.getBoundingClientRect().top;
      return { id: s.dataset.set, left: Math.round(r.left), top: Math.round(r.top), width: Math.round(r.width), height: Math.round(r.height), across: tiles.filter((t) => Math.abs(t.getBoundingClientRect().top - top) < 2).length, tile: Math.round(tiles[0]?.getBoundingClientRect().width || 0) };
    });
    return { width: window.innerWidth, inner: Math.round(body.width), sets, overflow: document.documentElement.scrollWidth > window.innerWidth || document.querySelector('.home').scrollWidth > document.querySelector('.home').clientWidth + 1 };
  })()`);
  const settled = async (width) => {
    // The offscreen frame adds a few pixels to what is asked for, so ask, measure and ask again.
    home.setContentSize(width, 860);
    await wait(300);
    const extra = (await layout()).width - width;
    if (extra) home.setContentSize(width - extra, 860);
    await until(async () => (await layout()).width === width, `the window at ${width}`, 4000);
    await wait(350);
    return layout();
  };
  try {
    home = await open('home');
    js = (code) => home.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.home-set').length === 2`), 'the two sets', 10000);
    await setChoice(all);
    await until(() => js(`document.querySelectorAll('.home-card:not([data-app="locker"])').length === ${APP_GROUPS.flatMap((g) => g.apps).length}`), 'every tile', 8000);

    const wide = await settled(1400);
    const [o, s] = wide.sets;
    check('home sets: at 1400 px Office and Studio stand side by side, Office on the left, three and two tiles across at one width, nothing running off the window',
      wide.sets.length === 2 && o.id === 'office' && s.id === 'studio' && o.left < s.left && Math.abs(o.top - s.top) < 2 && o.across === 3 && s.across === 2 && Math.abs(o.tile - s.tile) <= 8 && !wide.overflow, JSON.stringify(wide));
    await shot('home-sets-1400.png');

    const usual = await settled(1180);
    check('home sets: at 1180 px they are still side by side, with nothing running off the window',
      usual.sets[0].left < usual.sets[1].left && Math.abs(usual.sets[0].top - usual.sets[1].top) < 2 && usual.sets[0].across === 3 && !usual.overflow, JSON.stringify(usual));
    await shot('home-sets-1180.png');

    const narrow = await settled(1000);
    check('home sets: at 1000 px they stack, Office first, each the whole width',
      narrow.sets[0].id === 'office' && narrow.sets[0].top < narrow.sets[1].top && narrow.sets[0].width === narrow.sets[1].width && narrow.sets[0].left === narrow.sets[1].left && !narrow.overflow, JSON.stringify(narrow));
    await shot('home-sets-1000.png');
    await settled(1280);

    const order = await js(`[...document.querySelectorAll('.home-set')].map((s) => [...s.querySelectorAll('.home-card:not([data-app="locker"])')].map((c) => c.dataset.app))`);
    check('home sets: each column has its apps in the registry\'s order', JSON.stringify(order) === JSON.stringify(APP_GROUPS.map((g) => g.apps)), JSON.stringify(order));
    const chips = await js(`[...document.querySelectorAll('.home-set')].map((s) => [...s.querySelectorAll('.home-template')].map((c) => c.textContent.trim()))`);
    check('home sets: each column has its own Start something: documents and templates in Office, record, live, video and picture in Studio',
      JSON.stringify(chips[0].slice(0, 9)) === JSON.stringify(['Blank document', 'Letter', 'Report', 'Blank worksheet', 'Budget', 'Invoice', 'Blank presentation', 'Pitch deck', 'More templates'])
        && JSON.stringify(chips[1]) === JSON.stringify(['Record a screen', 'Go live', 'Introduction video', 'New video', 'Take a picture']), JSON.stringify(chips));
    const words = await js(`({ h3: [...document.querySelectorAll('.home-set h3')].map((h) => h.textContent), hero: document.querySelector('.home-hero p').textContent, tagline: document.querySelector('.home-card[data-app="mail"] .home-card-main').dataset.tip, once: document.querySelectorAll('.home-apps-btn').length + document.querySelectorAll('.home-packs-btn').length })`);
    check('home sets: headed Office and Studio, the new first line with no dash in it, the whole of what an app does on hover, and Choose apps and Add or remove apps once for both',
      JSON.stringify(words.h3) === '["Office","Studio"]' && /in Office; pictures, video and going live in Studio/.test(words.hero) && !/[\u2013\u2014]/.test(words.hero) && words.tagline.endsWith(APPS.mail.tagline) && words.once === 2, JSON.stringify(words));

    // Choose apps: the same two names, in the same order, then the add-ons.
    await js(`document.querySelector('.home-apps-btn')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.apps-chooser .apps-tile'))`), 'Choose apps', 5000);
    const chooser = await js(`({ heads: [...document.querySelectorAll('.apps-chooser .apps-group-head h3')].map((h) => h.textContent), tiles: [...document.querySelectorAll('.apps-chooser .apps-group')].map((g) => [...g.querySelectorAll('.apps-tile')].map((t) => t.dataset.app)) })`);
    check('home sets: Choose apps says Office, Studio and Add-ons, in the same order as Home',
      JSON.stringify(chooser.heads) === '["Office","Studio","Add-ons"]' && JSON.stringify(chooser.tiles.slice(0, 2)) === JSON.stringify(APP_GROUPS.map((g) => g.apps)), JSON.stringify(chooser));
    await shot('home-choose-apps.png');
    await js(`((list) => list.find((b) => ${labelKeyed("Cancel")}(b)) || list.find((b) => b.textContent.trim() === 'Cancel'))([...document.querySelectorAll('.apps-chooser .rw-dialog-foot button')])?.click(), 1`);
    await until(() => js(`!document.querySelector('.apps-chooser')`), 'Choose apps closed', 3000).catch(() => {});

    // The app menu's Window: Office and Studio above All apps, each listing its apps.
    await js(`document.querySelector('.rw-appmark')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.rw-menu'))`), 'the app menu', 4000);
    const openSub = async (label) => {
      await js(`((list) => list.find((b) => ${labelKeyed(label)}(b)) || list.find((b) => b.querySelector('.rw-menu-label')?.textContent.trim() === ${JSON.stringify(label)}))([...document.querySelectorAll('.rw-menu button')])?.click(), 1`);
      await until(() => js(`[...document.querySelectorAll('.rw-submenu')].some((m) => m.getAttribute('aria-label') === ${JSON.stringify(label)})`), `the ${label} submenu`, 3000);
    };
    const labels = (name) => js(`[...[...document.querySelectorAll('.rw-submenu')].find((m) => m.getAttribute('aria-label') === ${JSON.stringify(name)}).querySelectorAll(':scope > button')].map((b) => b.querySelector('.rw-menu-label').textContent.trim())`);
    await openSub('Window');
    const windowItems = (await labels('Window')).slice(-3);
    await openSub('Office');
    const officeItems = await labels('Office');
    await shot('home-window-menu.png');
    await js(`window.dispatchEvent(new MouseEvent('mousedown')), 1`);
    await until(() => js(`!document.querySelector('.rw-menu')`), 'the menu closed', 3000).catch(() => {});
    check('home sets: the Window menu has Office and Studio above All apps, Office listing its apps in order',
      JSON.stringify(windowItems) === '["Office","Studio","All apps"]' && JSON.stringify(officeItems) === JSON.stringify(APP_GROUPS[0].apps.map((k) => APPS[k].short)), JSON.stringify({ windowItems, officeItems }));

    // The macOS menu bar's list: the two sets, a line between.
    const appsMenu = Menu.getApplicationMenu()?.items.find((i) => /Apps/.test(i.label))?.submenu?.items || [];
    const native = appsMenu.map((i) => (i.type === 'separator' ? '-' : i.label)).slice(2);
    const between = native.indexOf('-');
    check('home sets: the application menu lists the two sets in order with a line between',
      JSON.stringify(native.slice(0, between)) === JSON.stringify(APP_GROUPS[0].apps.map((k) => APPS[k].short)) && JSON.stringify(native.slice(between + 1)) === JSON.stringify(APP_GROUPS[1].apps.map((k) => APPS[k].short)), JSON.stringify(native));

    // Studio's starts, each opening its own app.
    const opens = async (cls, wanted) => {
      const known = new Set(BrowserWindow.getAllWindows().map((w) => w.id));
      await js(`document.querySelector('.${cls}')?.click(), 1`);
      const fresh = () => BrowserWindow.getAllWindows().find((w) => !known.has(w.id) && !w.isDestroyed()) || null;
      await ok(until(() => Boolean(fresh()), `the window of ${cls}`, 8000));
      const w = fresh();
      if (!w) return null;
      await new Promise((r) => (w.webContents.isLoading() ? w.webContents.once('did-finish-load', r) : r()));
      const url = w.webContents.getURL();
      await wait(250);
      w.destroy();
      return wanted.test(url) ? true : url;
    };
    const went = {
      live: await opens('home-golive', /[?&]app=broadcast\b/),
      video: await opens('home-newvideo', /[?&]app=video\b/),
      picture: await opens('home-picture', /[?&]app=recorder\b.*[?&]picture=1/),
    };
    check('home sets: Go live opens Live Broadcast, New video opens Video and Take a picture opens the Recorder asked for a picture', went.live === true && went.video === true && went.picture === true, JSON.stringify(went));

    // A set alone across the width.
    await setChoice({ off: APP_GROUPS[1].apps.filter((k) => !OFF_UNTIL_CHOSEN.includes(k)), on: APP_GROUPS[0].apps.filter((k) => OFF_UNTIL_CHOSEN.includes(k)), addons: [] });
    await until(() => js(`document.querySelectorAll('.home-set').length === 1`), 'one set', 6000);
    const alone = await settled(1280);
    check('home sets: with only Office on, its column alone takes the whole width', alone.sets.length === 1 && alone.sets[0].id === 'office' && alone.sets[0].width >= alone.inner - 80 && alone.sets[0].across >= 4, JSON.stringify(alone));
    await shot('home-sets-office-only.png');

    await setChoice({ off: APP_GROUPS[0].apps.filter((k) => !OFF_UNTIL_CHOSEN.includes(k)), on: APP_GROUPS[1].apps.filter((k) => OFF_UNTIL_CHOSEN.includes(k)), addons: [] });
    await until(() => js(`document.querySelectorAll('.home-set').length === 1`), 'one set', 6000);
    const studio = await settled(1280);
    check('home sets: with only Studio on, its column alone takes the whole width, with no More templates in it', studio.sets[0].id === 'studio' && studio.sets[0].width >= studio.inner - 80 && !(await js(`Boolean(document.querySelector('.home-more'))`)), JSON.stringify(studio));
    await shot('home-sets-studio-only.png');

    // First start: the card names the sets.
    await js(`window.rutbaOffice.store.delete({ key: 'apps.choice' }).then(() => 1)`);
    await home.webContents.reload();
    await until(() => js(`Boolean(document.querySelector('.home-choose'))`), 'the first-start card', 8000);
    await wait(250);
    const note = await js(`document.querySelector('.home-choose-text span').textContent`);
    check('home sets: the first-start card names the sets', note === 'The Office apps are on. Add Talk, Data Files and the Studio apps when you want them, and turn off what you do not use, at any time.', note);
    await shot('home-sets-first-start.png');

    const complaints = await errorsIn(home);
    check('home sets: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('home sets: the check ran to the end', false, err?.message || String(err));
  } finally {
    await js?.(`window.rutbaOffice.store.set({ key: 'apps.choice', value: ${JSON.stringify(all)} }).then(() => 1)`).catch(() => {});
    // The window stays: the harness closes a block's windows, and this one may be all that keeps the run alive.
  }
}
