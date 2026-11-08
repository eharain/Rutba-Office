// The studio's setups and its Set up panel (Rutba Office 2.23, A4 and A5).
//
// The first time the window opens it opens on Set up with the quick scenes. A
// source is selected and dragged, pulled to the middle with a guide; its width
// typed in carries its height with it; Undo and Redo walk that back and on. A
// scene is saved and renamed. A setup is saved to a .rstudio file and opened as
// another. A window recorded before is found again by its name after a
// restart, when its id is new; a camera that is gone says Not found, its scene
// says so too, and Choose puts one in its place. Ctrl+1 and Ctrl+2 change panel.
// Run alone with RUTBA_VERIFY_ONLY=studio-setups.

import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { addItem, emptyStack } from '@rutba/studio/stack';
import { STORE } from '@rutba/studio/setup';
import { withScreenStandIn } from './verify-recorder.js';
import { seedSetup, restoreSetup, readSetup, CHECK_SETUP, PANEL, IN_PANEL } from './verify-studio-seed.js';

const SOUNDS = [{ id: 'mic-1', kind: 'mic', ref: '', name: '', live: true, muted: false, gain: 1 }, { id: 'system-1', kind: 'system', ref: '', name: '', live: false, muted: false, gain: 1 }];

export function verifyStudioSetups(h, args) {
  return withScreenStandIn(() => checks(h, args));
}

async function checks(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  const answer = (kind, value) => globalThis.__rutbaCheckDialogAnswers?.push({ kind, answer: value });
  const folder = path.join(dir, 'studio-setups');
  const docFile = path.join(dir, 'setups-fixture.docx');
  const setupFile = path.join(dir, 'Weekly review.rstudio');
  let js = null;
  let kept = null;
  let win = null;
  const shot = async (name) => {
    if (!process.env.RUTBA_VERIFY_CAPTURE) return;
    win.webContents.invalidate();
    await wait(700);
    fs.mkdirSync(process.env.RUTBA_VERIFY_CAPTURE, { recursive: true });
    fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await win.webContents.capturePage()).toPNG());
  };
  try {
    fs.mkdirSync(folder, { recursive: true });
    fs.writeFileSync(docFile, buildDocx({ styles: true, paragraphs: [{ text: 'A window for the setups check.' }] }));
    await open('word', docFile);
    win = await open('recorder');
    js = (code) => win.webContents.executeJavaScript(code);
    const set = (key, value) => js(`window.rutbaOffice.store.set({ key: ${JSON.stringify(key)}, value: ${JSON.stringify(value)} }).then(() => 1)`);
    const get = (key) => js(`window.rutbaOffice.store.get({ key: ${JSON.stringify(key)}, fallback: null })`);
    const menuPick = (re) => js(`[...document.querySelectorAll('[role^="menuitem"]')].find((n) => ${re}.test(n.textContent))?.click(), 1`);
    const openMenu = async () => {
      await js(`document.querySelector('.st-setup-btn')?.click(), 1`);
      await until(() => js(`document.querySelectorAll('[role^="menuitem"]').length > 0`), 'the setups menu', 4000).catch(() => {});
    };

    await until(() => js(`window.rutbaOffice.capture.sources().then((l) => l.some((x) => x.name.includes('setups-fixture')))`), 'the fixture window to be there', 10000).catch(() => {});
    const source = (await js(`window.rutbaOffice.capture.sources()`)).find((x) => x.name.includes('setups-fixture'));

    /* ── the first time ── */
    kept = await seedSetup(js, { app: 'recorder', stack: { ...emptyStack(), sounds: SOUNDS }, studio: { size: 720, fps: 30, countdown: 0, folder, minimise: false } });
    const oldKeys = await Promise.all(['recorder.stack', 'recorder.studio', 'broadcast.stack', 'broadcast.studio'].map(get));
    for (const k of ['recorder.stack', 'recorder.studio', 'broadcast.stack', 'broadcast.studio']) await set(k, null);
    await set(STORE.index, null);
    await win.webContents.reload();
    await until(() => js(`Boolean(document.querySelector('.st-panels'))`), 'the Recorder window', 15000);
    const firstSetup = await ok(until(() => js(IN_PANEL('setup')), 'Set up the first time', 8000));
    const quick = await js(`[...document.querySelectorAll('.st-scene[data-scene]')].map((n) => n.dataset.scene).join(' ')`);
    const sayName = await js(`document.querySelector('.st-setup-btn')?.textContent || ''`);
    const unchosen = await js(`Boolean([...document.querySelectorAll('.st-src')].find((n) => n.dataset.kind === 'window' && /not chosen/.test(n.textContent)))`);
    check('studio setups: the first time it opens on Set up, named My setup, with Screen, A window, Screen and camera, Camera and Sound only, and the window not chosen yet',
      firstSetup && quick === 'screen window screen-camera camera sound-only' && /My setup/.test(sayName) && unchosen, JSON.stringify({ firstSetup, quick, sayName, unchosen }));

    const index = () => get(STORE.index);
    const currentId = async () => (await index())?.current?.recorder;
    const stack = async () => (await get(STORE.setup(await currentId())))?.stack || null;

    /* ── arranging ── */
    await js(`document.querySelector('.st-scene[data-scene="screen-camera"]')?.click(), 1`);
    await until(() => js(`document.querySelectorAll('.st-box').length >= 2`), 'the frame\'s boxes', 6000).catch(() => {});
    await js(`document.querySelector('.st-card[data-kind="camera"]')?.click(), 1`);
    const props = await ok(until(() => js(`document.querySelector('.st-bar')?.dataset.kind === 'camera'`), 'the camera\'s properties', 4000));
    const cam = async () => (await stack())?.items.find((x) => x.kind === 'camera');
    const box0 = (await cam()).box;
    const dragged = await js(`(async () => {
      const f = document.querySelector('.st-frame');
      const r = f.getBoundingClientRect();
      const at = (x, y) => ({ bubbles: true, cancelable: true, button: 0, clientX: r.left + r.width * x, clientY: r.top + r.height * y, pointerId: 1 });
      const sx = ${box0.x + box0.w / 2}; const sy = ${box0.y + box0.h / 2};
      document.querySelector('.st-box[data-id=${JSON.stringify((await cam()).id)}]').dispatchEvent(new PointerEvent('pointerdown', at(sx, sy)));
      for (let k = 1; k <= 8; k++) { window.dispatchEvent(new PointerEvent('pointermove', at(sx + (0.505 - sx) * k / 8, sy + (0.495 - sy) * k / 8))); await new Promise((res) => setTimeout(res, 25)); }
      await new Promise((res) => setTimeout(res, 120));
      const guides = document.querySelectorAll('.st-guide').length;
      window.dispatchEvent(new PointerEvent('pointerup', at(0.505, 0.495)));
      return guides;
    })()`);
    await wait(300);
    const moved = (await cam()).box;
    const centred = Math.abs(moved.x + moved.w / 2 - 0.5) < 0.002 && Math.abs(moved.y + moved.h / 2 - 0.5) < 0.002;
    check('studio setups: a source dragged near the middle is pulled to it, with guides drawn while it moves', props && dragged >= 1 && centred, JSON.stringify({ props, dragged, moved }));

    const ratio = moved.w / moved.h;
    await js(`(() => { const n = document.querySelector('.st-box-w'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, '20'); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    const typed = await ok(until(async () => Math.abs((await cam()).box.w - 0.2) < 0.002, 'the width typed', 3000));
    const after = (await cam()).box;
    const carried = Math.abs(after.w / after.h - ratio) < 0.02;
    await js(`document.querySelector('.st-undo-btn')?.click(), 1`);
    const undone = await ok(until(async () => Math.abs((await cam()).box.w - moved.w) < 0.002, 'Undo', 3000));
    await js(`document.querySelector('.st-redo-btn')?.click(), 1`);
    const redone = await ok(until(async () => Math.abs((await cam()).box.w - 0.2) < 0.002, 'Redo', 3000));
    await js(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, bubbles: true })), 1`);
    const byKey = await ok(until(async () => Math.abs((await cam()).box.w - moved.w) < 0.002, 'Ctrl+Z', 3000));
    check('studio setups: a width typed in carries its height while the source keeps its shape, and Undo, Redo and Ctrl+Z walk it back and on', typed && carried && undone && redone && byKey, JSON.stringify({ typed, carried, undone, redone, byKey, after }));

    await js(`document.querySelector('.st-scene-save')?.click(), 1`);
    const saved = await ok(until(async () => (await stack())?.scenes.length === 5, 'the new scene', 3000));
    const newKey = (await stack())?.scenes[4]?.key;
    await js(`(() => { const b = document.querySelector('.st-scene[data-scene=${JSON.stringify(newKey || '')}]'); const r = b.getBoundingClientRect(); b.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: r.left + 4, clientY: r.top + 4 })); return 1; })()`);
    await until(() => js(`[...document.querySelectorAll('[role^="menuitem"]')].some((n) => /Rename/.test(n.textContent))`), 'the scene\'s menu', 3000).catch(() => {});
    await menuPick('/Rename/');
    await until(() => js(`Boolean(document.querySelector('.st-name-input'))`), 'the name box', 3000).catch(() => {});
    await js(`(() => { const n = document.querySelector('.st-name-input'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, 'Close up'); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await wait(100);
    await js(`document.querySelector('.st-name-ok')?.click(), 1`);
    const renamed = await ok(until(async () => (await stack())?.scenes.some((sc) => sc.name === 'Close up'), 'the scene renamed', 3000));
    await shot('studio-setup.png');
    check('studio setups: Save as scene keeps what the frame shows as a fifth scene, and right-clicking it renames it', saved && renamed, JSON.stringify({ saved, renamed, newKey }));

    /* ── a file ── */
    await openMenu();
    answer('save', setupFile);
    await menuPick('/Save to a file/');
    const written = await ok(until(() => fs.existsSync(setupFile) && fs.statSync(setupFile).size > 200, 'the setup file', 6000));
    const firstId = await currentId();
    await wait(200);
    await openMenu();
    answer('open', [setupFile]);
    await menuPick('/Open a file/');
    const asked = await ok(until(() => js(`Boolean(document.querySelector('.st-open-setup'))`), 'what was found', 6000));
    const rows = await js(`document.querySelector('.st-open-rows')?.textContent || ''`);
    await shot('studio-open-setup.png');
    await js(`document.querySelector('.st-open-add')?.click(), 1`);
    const added = await ok(until(async () => (await index())?.setups.length === 2 && (await currentId()) !== firstId, 'the second setup', 6000));
    const opened = await stack();
    await openMenu();
    const menu = await js(`[...document.querySelectorAll('[role^="menuitem"]')].map((n) => n.textContent).join(' | ')`);
    await js(`document.querySelector('.rw-menu-backdrop, .rw-overlay')?.click(), 1`).catch(() => {});
    await js(`document.activeElement?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })), 1`);
    await until(() => js(`document.querySelectorAll('[role^="menuitem"]').length === 0`), 'the menu closed', 3000).catch(() => {});
    check('studio setups: a setup saved to a .rstudio file opens as another, saying what was found, with its scenes, and both are in the menu',
      written && asked && added && opened?.scenes.some((sc) => sc.name === 'Close up') && (menu.match(/My setup/g) || []).length >= 2 && /Found/.test(rows), JSON.stringify({ written, asked, added, rows: rows.slice(0, 120), menu: menu.slice(0, 200) }));

    /* ── settings: this setup and every setup ── */
    const globalBefore = await get(STORE.global);
    const secondId = await currentId();
    const onValue = (cls) => js(`document.querySelector('.${cls} [aria-checked="true"]')?.dataset.value || ''`);
    const settingsOpen = async () => {
      await js(`document.querySelector('.st-gear')?.click(), 1`);
      return ok(until(() => js(`Boolean(document.querySelector('.st-quality'))`), 'the settings', 4000));
    };
    const settingsDone = async () => {
      await js(`document.querySelector('.st-settings-done')?.click(), 1`);
      await until(() => js(`!document.querySelector('.st-quality')`), 'the settings closed', 4000).catch(() => {});
    };
    const shown = await settingsOpen();
    const parts = await js(`[...document.querySelectorAll('.st-set-head h2')].map((n) => n.textContent).join(' | ')`);
    await js(`document.querySelector('.st-quality [data-value="high"]')?.click(), 1`);
    await js(`document.querySelector('.st-countdown [data-value="5"]')?.click(), 1`);
    await until(async () => (await get(STORE.setup(secondId)))?.settings?.quality === 'high' && (await get(STORE.global))?.countdown === 5, 'the settings kept', 4000).catch(() => {});
    await wait(200);
    await shot('studio-settings.png');
    await settingsDone();
    const inSetup = (await get(STORE.setup(secondId)))?.settings?.quality;
    const inGlobal = (await get(STORE.global))?.countdown;
    await openMenu();
    await menuPick('/^My setup(?! 2)/');
    const switched = await ok(until(async () => (await currentId()) === firstId, 'the first setup again', 5000));
    await wait(300);
    const shownAgain = await settingsOpen();
    const qualityThere = await onValue('st-quality');
    const countdownThere = await onValue('st-countdown');
    await settingsDone();
    check('studio setups: Settings has This setup and Every setup; quality set in one setup stays with it, and the countdown is the same in the other',
      shown && shownAgain && parts === 'This setup | Every setup' && inSetup === 'high' && inGlobal === 5 && switched && qualityThere !== 'high' && countdownThere === '5', JSON.stringify({ shown, parts, inSetup, inGlobal, switched, qualityThere, countdownThere }));
    await set(STORE.global, globalBefore);

    /* ── found again ── */
    let s = { ...emptyStack(), sounds: SOUNDS };
    s = addItem(s, { kind: 'window', ref: 'window:1:0', name: source.name, on: true }).stack;
    await seedSetup(js, { app: 'recorder', stack: s, studio: { size: 720, fps: 30, countdown: 0, folder, minimise: false } });
    await win.webContents.reload();
    await until(() => js(`Boolean(document.querySelector('.st-panels'))`), 'the Recorder again', 15000);
    await wait(800);
    const doc = await readSetup(js);
    const refound = doc?.stack.items[0]?.ref === source.id;
    const onRecord = await js(IN_PANEL('record'));
    const drawn = await ok(until(() => js(`document.querySelector('.st-canvas')?.dataset.layers === 'window'`), 'the window drawn', 15000));
    check('studio setups: after a restart a window whose id is new is found again by its name, and the window opens on Record', refound && onRecord && drawn, JSON.stringify({ was: 'window:1:0', now: doc?.stack.items[0]?.ref, wanted: source.id, onRecord, drawn }));

    /* ── what cannot be found ── */
    s = { ...emptyStack(), sounds: SOUNDS };
    s = addItem(s, { kind: 'camera', ref: 'gone-camera', name: 'Gone camera', on: true }).stack;
    await seedSetup(js, { app: 'recorder', stack: s, studio: { size: 720, fps: 30, countdown: 0, folder, minimise: false } });
    await win.webContents.reload();
    await until(() => js(`Boolean(document.querySelector('.st-panels'))`), 'the Recorder once more', 15000);
    const toSetup = await ok(until(() => js(`Boolean(document.querySelector('.st-src.miss[data-kind="camera"]'))`), 'Not found in Set up', 12000));
    const says = await js(`document.querySelector('.st-src.miss')?.textContent || ''`);
    await shot('studio-not-found.png');
    await js(`document.querySelector('.st-src.miss .st-src-choose')?.click(), 1`);
    await until(() => js(`document.querySelectorAll('[role^="menuitem"]').length > 0`), 'the cameras', 3000).catch(() => {});
    await js(`[...document.querySelectorAll('[role^="menuitem"]')].find((n) => !/None/.test(n.textContent))?.click(), 1`);
    const chosen = await ok(until(() => js(`!document.querySelector('.st-src.miss')`), 'the camera chosen', 5000));
    const camNow = (await get(STORE.setup(CHECK_SETUP)))?.stack.items[0];
    check('studio setups: a camera that is gone opens Set up, says Not found under its name, and Choose puts one there and keeps it',
      toSetup && /Not found/.test(says) && /Gone camera/.test(says) && chosen && camNow?.ref && camNow.ref !== 'gone-camera', JSON.stringify({ toSetup, says, chosen, ref: camNow?.ref?.slice(0, 12) }));

    await js(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '2', ctrlKey: true, bubbles: true })), 1`);
    const k2 = await ok(until(() => js(IN_PANEL('record')), 'Ctrl+2', 3000));
    await js(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '1', ctrlKey: true, bubbles: true })), 1`);
    const k1 = await ok(until(() => js(IN_PANEL('setup')), 'Ctrl+1', 3000));
    await js(PANEL('main'));
    check('studio setups: Ctrl+2 and Ctrl+1 change panel', k1 && k2, JSON.stringify({ k1, k2 }));

    const complaints = await errorsIn(win);
    check('studio setups: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    for (const [i, k] of ['recorder.stack', 'recorder.studio', 'broadcast.stack', 'broadcast.studio'].entries()) await set(k, oldKeys[i] ?? null);
  } catch (err) {
    check('studio setups: the check ran to the end', false, err?.message || String(err));
  } finally {
    if (js && kept) await restoreSetup(js, kept).catch(() => {});
  }
}
