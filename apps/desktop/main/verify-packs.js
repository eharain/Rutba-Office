// App packs in a running core (docs/APP-PACKS.md 5.6), over two starts of
// the same profile by build/verify-packs-run.js: a packs folder of a few
// signed packs, and a folder of the release's pack files to add from.
//
// The first start has Documents, Mail and AI Tools. An app not there has no
// channel in the main process and none in the window's bridge, no tile, and
// no window: a file of its kind opens the launcher, which says it is not
// installed. Add or remove apps adds Calendar from the folder and removes
// Mail. The second start has Calendar and not Mail, and Mail's files are
// gone from the packs folder.
//
// RUTBA_PACKS_PHASE says which start this is (first or next). Run with
// RUTBA_VERIFY_ONLY=packs.

import fs from 'node:fs';
import path from 'node:path';
import { ipcMain } from 'electron';
import { CHANNEL_PREFIX, packIpc } from '@rutba/office-shell/contract';
import { packFiles, pendingRemovals } from '@rutba/office-shell/electron/pack-install';
import { LOADED_KEY, toolOffered } from './pack-tools.js';

/** Whether the main process answers a pack's channels: a second handler for a channel is refused while one is there. */
function answered(id) {
  const ipc = packIpc(id);
  const out = [];
  for (const [ns, names] of Object.entries(ipc)) {
    for (const name of names) {
      const channel = `${CHANNEL_PREFIX}/${ns}:${name}`;
      try {
        ipcMain.handle(channel, () => null);
        ipcMain.removeHandler(channel);
        out.push(false);
      } catch {
        out.push(true);
      }
    }
  }
  return { all: out.length > 0 && out.every(Boolean), none: out.every((x) => !x), count: out.length };
}

export async function verifyPacks({ open, check, until, wait, errorsIn }, { dir }) {
  const phase = process.env.RUTBA_PACKS_PHASE || 'first';
  const packsDir = process.env.RUTBA_PACKS_DIR;
  const has = phase === 'first' ? ['word', 'mail', 'ai'] : ['word', 'calendar', 'ai'];
  const absent = phase === 'first' ? ['calendar', 'sheets', 'talk', 'base'] : ['mail', 'sheets', 'talk', 'base'];
  const name = (s) => `packs (${phase} start): ${s}`;
  let home = null;
  try {
    home = await open('home');
    const js = (code) => home.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.home-card').length > 0`), 'the launcher', 15000);

    // What this start loaded, in the window and in the main process.
    const told = await js(`(window.rutbaOffice.__contract.packs || []).map((p) => p.id).sort()`);
    const bridge = await js(`(${JSON.stringify([...has, ...absent])}).map((id) => ({ id, ns: { word: null, mail: 'mail', ai: 'ai', calendar: 'calendar', sheets: null, talk: 'talk', base: 'base' }[id] })).filter((x) => x.ns).map((x) => [x.id, Boolean(window.rutbaOffice[x.ns])])`);
    const inBridge = Object.fromEntries(bridge);
    check(name(`the windows are told of ${has.join(', ')} alone`), JSON.stringify(told) === JSON.stringify([...has].sort()), JSON.stringify(told));
    check(name('an app not installed has no channel in the window\'s bridge, an installed one has its own'),
      absent.every((id) => inBridge[id] !== true) && has.every((id) => inBridge[id] !== false), JSON.stringify(inBridge));
    const main = Object.fromEntries([...has, ...absent].map((id) => [id, answered(id)]));
    check(name('the main process answers the channels of the packs loaded and none of the others'),
      has.every((id) => main[id].count === 0 || main[id].all) && absent.every((id) => main[id].none) && main[phase === 'first' ? 'mail' : 'calendar'].all,
      JSON.stringify(Object.fromEntries(Object.entries(main).map(([k, v]) => [k, v.count ? (v.all ? 'all' : v.none ? 'none' : 'some') : '-']))));

    // The tiles: the installed apps only (a check run turns every app on in Choose apps).
    const tiles = await js(`[...document.querySelectorAll('.home-card')].map((c) => c.dataset.app).sort()`);
    const apps = has.filter((id) => id !== 'ai').sort();
    check(name(`the launcher's tiles are ${apps.join(' and ')}`), JSON.stringify(tiles) === JSON.stringify(apps), JSON.stringify(tiles));

    // The AI tools read what this start loaded.
    const loaded = await js(`window.rutbaOffice.store.get({ key: ${JSON.stringify(LOADED_KEY)}, fallback: null })`);
    const settings = { [LOADED_KEY]: loaded };
    check(name('the AI tools offer Documents\' tools and not Worksheets\''),
      Array.isArray(loaded) && toolOffered('create_document', settings) && !toolOffered('create_workbook', settings) && toolOffered('list_folders', settings), JSON.stringify(loaded));

    // A file of an app not installed opens the launcher, which says so; one of an installed app opens it.
    const gone = phase === 'first' ? { app: 'calendar', ext: 'ics', said: 'Calendar' } : { app: 'mail', ext: 'eml', said: 'Mail' };
    const file = path.join(dir, `not-here.${gone.ext}`);
    fs.writeFileSync(file, '');
    const asked = await open(gone.app, file);
    const ajs = (code) => asked.webContents.executeJavaScript(code);
    await until(() => ajs(`Boolean(document.querySelector('.home-packs-note'))`), 'the note', 10000).catch(() => {});
    const note = await ajs(`document.querySelector('.home-packs-note .home-packs-text')?.textContent || ''`);
    const url = asked.webContents.getURL();
    check(name(`a .${gone.ext} file opens the launcher, saying ${gone.said} is not installed`),
      /app=home/.test(url) && note.includes(`${gone.said} is not installed on this computer.`), `${url.split('?')[1] || url} | ${note.slice(0, 120)}`);
    asked.destroy();

    if (phase === 'first') {
      // Add or remove apps: Calendar added from the folder, Mail removed.
      await js(`document.querySelector('.home-packs-btn')?.click(), 1`);
      await until(() => js(`document.querySelectorAll('.packs-dialog .packs-row').length > 0`), 'Add or remove apps', 8000).catch(() => {});
      const rows = await js(`document.querySelectorAll('.packs-dialog .packs-row').length`);
      await js(`document.querySelector('.packs-row[data-pack="calendar"] .packs-add')?.click(), 1`);
      const adding = await until(() => js(`document.querySelector('.packs-row[data-pack="calendar"]')?.dataset.state === 'adding'`), 'Calendar added', 15000).then(() => true).catch(() => false);
      await js(`document.querySelector('.packs-row[data-pack="mail"] .packs-remove')?.click(), 1`);
      const removing = await until(() => js(`document.querySelector('.packs-row[data-pack="mail"]')?.dataset.state === 'removing'`), 'Mail removed', 8000).then(() => true).catch(() => false);
      const restart = await js(`Boolean(document.querySelector('.packs-dialog .packs-restart'))`);
      check(name('Add or remove apps lists every pack, adds Calendar from the folder and writes Mail down to go, and offers to restart'),
        rows === 15 && adding && removing && restart && fs.existsSync(path.join(packsDir, packFiles('calendar').manifest)) && JSON.stringify(pendingRemovals(packsDir)) === '["mail"]',
        JSON.stringify({ rows, adding, removing, restart, pending: pendingRemovals(packsDir) }));
      // The last app is kept.
      const last = await js(`window.rutbaOffice.packs.remove({ id: 'word' }).then(() => 'removed', (e) => e.message)`);
      check(name('the last app cannot be removed'), /Keep one app at least/.test(last), last);
    } else {
      const f = packFiles('mail');
      const left = [f.manifest, f.archive, f.unpacked].filter((n) => fs.existsSync(path.join(packsDir, n)));
      check(name('Mail\'s files are gone from the packs folder, and nothing is left to remove'), left.length === 0 && pendingRemovals(packsDir).length === 0, JSON.stringify({ left, pending: pendingRemovals(packsDir) }));
      const calendar = await open('calendar');
      const shown = await until(() => calendar.webContents.executeJavaScript(`/app=calendar/.test(location.search) && document.body.innerText.length > 0`), 'the Calendar window', 15000).then(() => true).catch(() => false);
      check(name('Calendar, added, opens'), shown, calendar.webContents.getURL().split('?')[1] || '');
      calendar.destroy();
    }
    await wait(200);
    const complaints = await errorsIn(home);
    check(name('none of that reported an error'), complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check(name('the check ran to the end'), false, err?.message || String(err));
  } finally {
    home?.destroy();
  }
}
