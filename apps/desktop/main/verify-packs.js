// App packs in a running core (docs/APP-PACKS.md 5.6), over two starts of
// the same profile by build/verify-packs-run.js: a packs folder of a few
// signed packs, and a folder of the release's pack files to add from.
//
// The first start has Documents, Mail, AI Tools, Talk, Text recognition and
// Data Files. An app not there has no
// channel in the main process and none in the window's bridge, no tile, and
// no window: a file of its kind opens the launcher, which says it is not
// installed. Add or remove apps adds Calendar from the folder and removes
// Mail. The second start has Calendar and not Mail, and Mail's files are
// gone from the packs folder. Every pack carries its own code: the core is
// told so, the windows open, Talk's, Documents', Mail's and the AI Tools
// window's code is served from their packs, an AI app's program runs AI
// Tools' code from its pack, and text recognition reads a picture with the
// engine and models in its pack.
//
// RUTBA_PACKS_PHASE says which start this is (first or next). Run with
// RUTBA_VERIFY_ONLY=packs.

import fs from 'node:fs';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { app, ipcMain } from 'electron';
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
  const has = phase === 'first' ? ['word', 'mail', 'ai', 'talk', 'ocr', 'base'] : ['word', 'calendar', 'ai', 'talk', 'ocr', 'base'];
  const absent = phase === 'first' ? ['calendar', 'sheets', 'contacts'] : ['mail', 'sheets', 'contacts'];
  const name = (s) => `packs (${phase} start): ${s}`;
  let home = null;
  try {
    home = await open('home');
    const js = (code) => home.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.home-card').length > 0`), 'the launcher', 15000);

    // What this start loaded, in the window and in the main process.
    const told = await js(`(window.rutbaOffice.__contract.packs || []).map((p) => p.id).sort()`);
    const bridge = await js(`(${JSON.stringify([...has, ...absent])}).map((id) => ({ id, ns: { word: null, mail: 'mail', ai: 'ai', calendar: 'calendar', sheets: null, talk: 'talk', base: 'base', ocr: 'ocr', contacts: 'contacts' }[id] })).filter((x) => x.ns).map((x) => [x.id, Boolean(window.rutbaOffice[x.ns])])`);
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
    const apps = has.filter((id) => !['ai', 'ocr'].includes(id)).sort();
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
      // The last app is kept: with Mail, Talk and Data Files going, Documents is the last (Text recognition is not an app); the two are then kept after all.
      const last = await js(`window.rutbaOffice.packs.remove({ id: 'talk' }).then(() => window.rutbaOffice.packs.remove({ id: 'base' })).then(() => window.rutbaOffice.packs.remove({ id: 'word' })).then(() => 'removed', (e) => e.message)`);
      const kept = await js(`Promise.all(['talk', 'base'].map((id) => window.rutbaOffice.packs.keep({ id }))).then(() => 1, (e) => e.message)`);
      check(name('the last app cannot be removed'), /Keep one app at least/.test(last) && kept === 1 && JSON.stringify(pendingRemovals(packsDir)) === '["mail"]', `${last} | ${JSON.stringify(pendingRemovals(packsDir))}`);
      // Talk, from its own pack: its code, its window, and its window's chunk.
      const code = await js(`(window.rutbaOffice.__contract.packs || []).find((p) => p.id === 'talk')?.code || null`);
      const codes = await js(`Object.fromEntries((window.rutbaOffice.__contract.packs || []).map((p) => [p.id, p.code]))`);
      check(name('the core is told Talk, Text recognition, Data Files, Documents, Mail and AI Tools carry their own code'),
        code === 'own' && ['ocr', 'base', 'word', 'mail', 'ai'].every((id) => codes[id] === 'own'), JSON.stringify(codes));
      // Text recognition from its pack: a small picture of nothing read by its engine, on its threads.
      const read = await js(`window.rutbaOffice.ocr.recognize({ image: { width: 64, height: 32, data: new Uint8Array(64 * 32 * 4).fill(255) }, languages: ['eng'], job: 'packs-check' }).then((r) => ({ ok: true, text: r.text, ms: r.ms }), (e) => ({ ok: false, error: e.message }))`);
      check(name('text recognition reads a picture with the engine and models in its own pack'), read.ok === true, JSON.stringify(read));
      // Data Files' window, from its pack.
      const base = await open('base');
      const baseShown = await until(() => base.webContents.executeJavaScript(`/app=base/.test(location.search) && document.body.innerText.length > 0`), 'the Data Files window', 15000).then(() => true).catch(() => false);
      check(name('Data Files opens from its own pack'), baseShown, base.webContents.getURL().split('?')[1] || '');
      base.destroy();
      // Talk's, Documents' and Mail's windows, each with its window's chunk served from its own pack.
      for (const [id, said] of [['talk', 'Talk'], ['word', 'Documents'], ['mail', 'Mail']]) {
        const win = await open(id);
        const wjs = (c) => win.webContents.executeJavaScript(c);
        const shown = await until(() => wjs(`new URLSearchParams(location.search).get('app') === ${JSON.stringify(id)} && document.body.innerText.length > 0`), `the ${said} window`, 15000).then(() => true).catch(() => false);
        // The chunk's name, from the pack's own archive; fetched from the window at the core's address for it.
        const chunk = fs.readdirSync(path.join(packsDir, packFiles(id).archive, 'renderer', 'chunks')).find((n) => n.startsWith(`${id}-`) && n.endsWith('.js')) || null;
        const served = chunk
          ? await wjs(`fetch(new URL('/chunks/${chunk}', location.href)).then((r) => ({ chunk: ${JSON.stringify(chunk)}, status: r.status, pack: r.headers.get('x-rutba-pack') }))`)
          : { chunk: null };
        check(name(`${said} opens, its window's chunk served from its own pack`), shown && served.status === 200 && served.pack === id, JSON.stringify(served));
        win.destroy();
      }
      // The AI Tools window, opened from the app menu: its code from AI Tools' pack.
      await js(`window.dispatchEvent(new Event('rutba-ai-tools')), 1`);
      const aiShown = await until(() => js(`Boolean(document.querySelector('.ai-tools-dialog .ai-modes'))`), 'the AI Tools window', 10000).then(() => true).catch(() => false);
      const aiChunk = fs.readdirSync(path.join(packsDir, packFiles('ai').archive, 'renderer', 'chunks')).find((n) => n.startsWith('ai-tools-dialog-')) || null;
      const aiServed = aiChunk ? await js(`fetch('/chunks/${aiChunk}').then((r) => ({ chunk: ${JSON.stringify(aiChunk)}, status: r.status, pack: r.headers.get('x-rutba-pack') }))`) : { chunk: null };
      await js(`document.querySelector('.ai-tools-dialog .rw-btn.primary')?.click(), 1`);
      check(name('the AI Tools window opens, its code served from AI Tools\' own pack'), aiShown && aiServed.status === 200 && aiServed.pack === 'ai', JSON.stringify(aiServed));
      // The program an AI app starts, as the window says to start it: the core's start file runs AI Tools' program from its pack.
      const { command, args, env } = (await js(`window.rutbaOffice.ai.state()`)).connect || {};
      const started = await new Promise((resolve) => {
        const child = spawn(command, args, { stdio: ['pipe', 'pipe', 'pipe'], env: { ...process.env, ...env } });
        let out = '';
        let err = '';
        const done = (v) => { clearTimeout(timer); child.kill(); resolve(v); };
        const timer = setTimeout(() => done({ error: `no answer; ${err.slice(0, 200)}` }), 30000);
        child.stderr.on('data', (d) => { err += d; });
        child.stdout.on('data', (d) => {
          out += d;
          const at = out.indexOf('\n');
          if (at < 0) return;
          try { done(JSON.parse(out.slice(0, at)).result?.serverInfo || { error: out.slice(0, 200) }); } catch { done({ error: out.slice(0, 200) }); }
        });
        child.on('error', (e) => done({ error: e.message }));
        child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'Packs check', version: '1.0' } } })}\n`);
      });
      check(name('an AI app\'s program runs AI Tools\' code from its pack, told this copy\'s version by the core'), started.name === 'rutba-office' && started.version === app.getVersion(), JSON.stringify(started));
    } else {
      const f = packFiles('mail');
      const left = [f.manifest, f.archive, f.unpacked].filter((n) => fs.existsSync(path.join(packsDir, n)));
      check(name('Mail\'s files are gone from the packs folder, and nothing is left to remove'), left.length === 0 && pendingRemovals(packsDir).length === 0, JSON.stringify({ left, pending: pendingRemovals(packsDir) }));
      const calendar = await open('calendar');
      const shown = await until(() => calendar.webContents.executeJavaScript(`/app=calendar/.test(location.search) && document.body.innerText.length > 0`), 'the Calendar window', 15000).then(() => true).catch(() => false);
      const calendarCode = await js(`(window.rutbaOffice.__contract.packs || []).find((p) => p.id === 'calendar')?.code || null`);
      check(name('Calendar, added, carries its own code and opens'), shown && calendarCode === 'own', `${calendarCode} | ${calendar.webContents.getURL().split('?')[1] || ''}`);
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
