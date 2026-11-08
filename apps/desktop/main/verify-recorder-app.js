// Rutba Recorder, the recorder as an app of its own (owner, 2026-10-19:
// "split the recording tool out to be able to record video"), rebuilt as one
// studio (owner, 2026-11-02: "the screen and cameras all should work exactly
// same way so no separate features and functions on top").
//
// The launcher has its tile. The window opens with an empty share and
// Chromium's stand-in camera and microphone (RUTBA_FAKE_MEDIA); a Documents
// window this check opens is in the list on the right among the windows not
// shown (a check run lists only the suite's own windows), and recorded as
// Chromium's stand-in screen (withScreenStandIn). Shown from the list with
// the camera, the Screen and camera scene draws them as one picture. The
// window and the camera are clipped the same way from the bar over the share;
// the camera is made a circle and mirrored, reordered from the list, stopped
// and started. Video on Placeholder draws the card; a title is added from the
// list's Add. Take Picture keeps the program, the selected source and one
// after a timer, by the button and by its key, and once more while Record
// runs. Record and Stop are one button that shows the time; the scene is
// changed and the recording paused while it runs, and Stop keeps one WebM
// with picture and sound. With Video off, Record records sound alone. Edit in
// Video opens the recording on Video's timeline. Run alone with
// RUTBA_VERIFY_ONLY=recorder-app.

import fs from 'node:fs';
import path from 'node:path';
import { BrowserWindow } from 'electron';
import { shortcuts } from '@rutba/office-shell/electron/shortcuts';
import { buildDocx } from '@rutba/ooxml/build';
import { probeMedia } from '@rutba/media/probe';
import { addItem, emptyStack } from '@rutba/studio/stack';
import { withScreenStandIn } from './verify-recorder.js';
import { seedSetup, restoreSetup, readSetup, PANEL, IN_PANEL } from './verify-studio-seed.js';

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args the run's scratch folder
 */
export function verifyRecorderApp(h, args) {
  return withScreenStandIn(() => appChecks(h, args));
}

/** The setup the check starts from: the fixture window and the camera, not shown yet, and the scenes they make. */
function startStack(source) {
  let s = { ...emptyStack(), sounds: [{ id: 'mic-1', kind: 'mic', ref: '', name: '', live: true, muted: false, gain: 1 }, { id: 'system-1', kind: 'system', ref: '', name: '', live: false, muted: false, gain: 1 }] };
  s = addItem(s, { kind: 'window', ref: source.id, name: source.name, on: false }).stack;
  return addItem(s, { kind: 'camera', ref: '', name: '', on: false, shape: 'circle' }).stack;
}

async function appChecks(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  const folder = path.join(dir, 'studio-recordings');
  const docFile = path.join(dir, 'studio-fixture.docx');
  let js = null;
  let kept = null;
  try {
    fs.mkdirSync(folder, { recursive: true });
    const home = await open('home');
    // The launcher drawn before it is read: beside other work it was read while still empty.
    await until(() => home.webContents.executeJavaScript(`Boolean(document.querySelector('.home-card[data-app="recorder"]'))`), 'the launcher\'s Recorder tile', 15000).catch(() => {});
    const tile = await home.webContents.executeJavaScript(`(() => { const c = document.querySelector('.home-card[data-app="recorder"]'); return c ? c.textContent : ''; })()`);
    check('recorder app: the launcher has a Recorder tile', /Recorder/.test(tile), tile.slice(0, 80));

    fs.writeFileSync(docFile, buildDocx({ styles: true, paragraphs: [{ text: 'Recorded by the Recorder check.' }, { text: 'A second line.' }] }));
    await open('word', docFile);
    const win = await open('recorder');
    js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`window.rutbaOffice.capture.sources().then((l) => l.some((x) => x.name.includes('studio-fixture')))`), 'the fixture window to be there', 10000).catch(() => {});
    const source = (await js(`window.rutbaOffice.capture.sources()`)).find((x) => x.name.includes('studio-fixture'));
    kept = await seedSetup(js, { app: 'recorder', stack: startStack(source), studio: { size: 720, fps: 30, countdown: 0, folder, minimise: false, shot: 'program', shotTimer: 0 } });
    await win.webContents.reload();
    await until(() => js(`Boolean(document.querySelector('.st-panel[data-panel="record"]'))`), 'the Recorder window', 10000);
    const stack = async () => (await readSetup(js))?.stack;
    const go = async (p) => { await js(PANEL(p)); await until(() => js(IN_PANEL(p)), `the ${p} panel`, 4000).catch(() => {}); };
    const itemOf = async (kind) => ((await stack())?.items || []).find((it) => it.kind === kind) || null;
    const canvas = () => js(`(() => { const c = document.querySelector('.st-canvas'); return { layers: c.dataset.layers, ids: c.dataset.ids, video: c.dataset.video, frames: Number(c.dataset.frames), size: [c.width, c.height] }; })()`);

    // Two panels, the Record panel holding the few controls a recording needs.
    const line = await js(`[...document.querySelectorAll('.st-panels [data-panel]')].map((n) => n.textContent).join(' | ')`);
    const recordSays = await js(`document.querySelector('.st-dock .st-record')?.textContent || ''`);
    const noLive = await js(`!document.querySelector('.lv-go') && !document.querySelector('.st-auto') && Boolean(document.querySelector('.st-dock .st-shoot')) && Boolean(document.querySelector('.st-top .st-setup-btn'))`);
    check('recorder app: two panels, Set up and Record, with the setup at the top and Picture and Record on the bar at the foot, no Go Live and no switching',
      line === 'Set up | Record' && /Record/.test(recordSays) && noLive, JSON.stringify({ line, recordSays, noLive }));

    // The fixture window and the camera, listed in Set up, put on the frame with their eyes.
    await go('setup');
    const listed = await ok(until(() => js(`Boolean(document.querySelector('.st-list .st-card[data-kind="window"][data-on="0"]'))`), 'the fixture window in the list', 10000));
    await js(`document.querySelector('.st-list .st-card[data-kind="window"] .st-eye')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.st-list .st-card[data-kind="window"][data-on="1"]'))`), 'the window on the share', 6000).catch(() => {});
    const camSpare = await ok(until(() => js(`Boolean(document.querySelector('.st-list .st-card[data-kind="camera"]'))`), 'the camera in the list', 8000));
    await js(`document.querySelector('.st-list .st-card[data-kind="camera"] .st-eye')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.st-list .st-card[data-kind="camera"][data-on="1"]'))`), 'the camera on the share', 6000).catch(() => {});
    await js(`document.querySelector('.st-scene[data-scene="screen-camera"]')?.click(), 1`);
    const live = await ok(until(() => js(`(() => { const c = document.querySelector('.st-canvas'); return c && c.dataset.layers === 'window camera' && Number(c.dataset.frames) > 10; })()`), 'the window and the camera in the program', 25000));
    const look = await js(`(() => { const c = document.querySelector('.st-canvas'); const g = c.getContext('2d'); const mid = g.getImageData(c.width / 2, c.height / 2, 1, 1).data; return { layers: c.dataset.layers, size: [c.width, c.height], scenes: document.querySelectorAll('.st-scene[data-scene]').length, mid: mid[0] + mid[1] + mid[2], shown: document.querySelectorAll('.st-list .st-card[data-on="1"]').length }; })()`);
    check('recorder app: the window and the camera shown from the list are drawn live as one picture, 1280 by 720 as set, with the four scenes they make and Sound only',
      listed && camSpare && live && look.size[0] === 1280 && look.size[1] === 720 && look.scenes === 5 && look.mid > 60 && look.shown === 2,
      JSON.stringify({ listed, camSpare, live, ...look }));

    // Clip, the same for the window and the camera: Clip on the bar, then a drag over the source's box.
    const drag = (x0, y0, x1, y1) => js(`(async () => {
      const f = document.querySelector('.st-frame');
      const r = f.getBoundingClientRect();
      const at = (x, y) => ({ bubbles: true, cancelable: true, button: 0, clientX: r.left + r.width * x, clientY: r.top + r.height * y, pointerId: 1 });
      f.dispatchEvent(new PointerEvent('pointerdown', at(${x0}, ${y0})));
      for (let k = 1; k <= 6; k++) { window.dispatchEvent(new PointerEvent('pointermove', at(${x0} + (${x1} - ${x0}) * k / 6, ${y0} + (${y1} - ${y0}) * k / 6))); await new Promise((res) => setTimeout(res, 25)); }
      window.dispatchEvent(new PointerEvent('pointerup', at(${x1}, ${y1})));
      return 'drawn';
    })()`);
    const select = (kind) => js(`document.querySelector('.st-list .st-card[data-kind="${kind}"]')?.click(), 1`);
    const clipOf = async (kind) => {
      await select(kind);
      await until(() => js(`document.querySelector('.st-bar')?.dataset.kind === '${kind}'`), `the ${kind}'s bar`, 4000).catch(() => {});
      await js(`document.querySelector('.st-bar .st-clip')?.click(), 1`);
      await until(() => js(`Boolean(document.querySelector('.st-box.clipping'))`), 'clipping', 3000).catch(() => {});
      const b = (await itemOf(kind)).box;
      await drag(b.x + b.w * 0.2, b.y + b.h * 0.2, b.x + b.w * 0.7, b.y + b.h * 0.8);
      await until(async () => Boolean((await itemOf(kind))?.crop), `the ${kind} clipped`, 4000).catch(() => {});
      return (await itemOf(kind))?.crop || null;
    };
    const winCrop = await clipOf('window');
    const camCrop = await clipOf('camera');
    const winBar = await js(`(() => { const b = document.querySelector('.st-bar'); return ['.st-clip', '.st-place', '.st-mirror-btn', '.st-shapes [data-value="circle"]'].every((s) => b.querySelector(s)); })()`);
    const near = (c) => c && Math.abs(c.x - 0.2) < 0.06 && Math.abs(c.y - 0.2) < 0.06 && Math.abs(c.w - 0.5) < 0.08 && Math.abs(c.h - 0.6) < 0.08;
    // The camera is a mirrored circle in that scene, filled by the middle of its picture: the part
    // dragged over is as tall as on the window, narrower, and to the right of its middle, as drawn.
    const nearCircle = (c) => c && Math.abs(c.y - 0.2) < 0.06 && Math.abs(c.h - 0.6) < 0.08 && c.w > 0.2 && c.w < 0.45 && c.x + c.w / 2 > 0.5;
    await select('window');
    await js(`document.querySelector('.st-bar .st-whole')?.click(), 1`);
    const whole = await ok(until(async () => !(await itemOf('window'))?.crop, 'the window whole', 4000));
    check('recorder app: the window and the camera are clipped the same way, Clip on the bar and a drag over the part to show, and Whole picture shows all of it again',
      near(winCrop) && nearCircle(camCrop) && winBar && whole, JSON.stringify({ winCrop, camCrop, winBar, whole }));

    // Shape and mirror the camera, and reorder: the window brought over it from the list.
    await select('camera');
    const wasMirrored = Boolean((await itemOf('camera'))?.mirror);
    await js(`document.querySelector('.st-bar .st-shapes [data-value="rounded"]')?.click(), 1`);
    await until(async () => (await itemOf('camera'))?.shape === 'rounded', 'rounded', 3000).catch(() => {});
    await js(`document.querySelector('.st-bar .st-shapes [data-value="circle"]')?.click(), 1`);
    await js(`document.querySelector('.st-bar .st-mirror-btn')?.click(), 1`);
    const shaped = await ok(until(async () => { const c = await itemOf('camera'); return c?.shape === 'circle' && c.mirror === !wasMirrored; }, 'a circle, mirrored the other way', 4000));
    await js(`(() => { const n = document.querySelector('.st-list .st-card[data-kind="window"][data-on="1"]'); n.focus(); n.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowUp', altKey: true, bubbles: true })); return 1; })()`);
    const reordered = await ok(until(() => js(`document.querySelector('.st-canvas')?.dataset.layers === 'camera window'`), 'the window on top', 4000));
    const listOrder = await js(`[...document.querySelectorAll('.st-list .st-card[data-on="1"]')].map((n) => n.dataset.kind).join(' ')`);
    check('recorder app: the camera is shaped and mirrored from its bar, and Alt+Up in the list puts the window over it, top first in the list',
      shaped && reordered && listOrder === 'window camera', JSON.stringify({ shaped, reordered, listOrder }));
    await js(`(() => { const n = document.querySelector('.st-list .st-card[data-kind="window"][data-on="1"]'); n.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', altKey: true, bubbles: true })); return 1; })()`);
    await until(() => js(`document.querySelector('.st-canvas')?.dataset.layers === 'window camera'`), 'the camera on top again', 4000).catch(() => {});

    // Stop and Start on the camera's card: stopped, it keeps its last picture; started, it is live again.
    await js(`document.querySelector('.st-list .st-card[data-kind="camera"] .st-run')?.click(), 1`);
    const stopped = await ok(until(() => js(`document.querySelector('.st-list .st-card[data-kind="camera"]')?.dataset.state === 'stopped'`), 'the camera stopped', 6000));
    const stillDrawn = await js(`document.querySelector('.st-canvas')?.dataset.layers === 'window camera'`);
    await js(`document.querySelector('.st-list .st-card[data-kind="camera"] .st-run')?.click(), 1`);
    const restarted = await ok(until(() => js(`document.querySelector('.st-list .st-card[data-kind="camera"]')?.dataset.state === 'live'`), 'the camera live again', 15000));
    check('recorder app: Stop on the camera\'s card stops it and its last picture stays on the share; Start brings it back', stopped && stillDrawn && restarted, JSON.stringify({ stopped, stillDrawn, restarted }));

    // A title from Add is drawn over the sources.
    await js(`document.querySelector('.st-add')?.click(), 1`);
    await until(() => js(`[...document.querySelectorAll('[role="menuitem"]')].some((n) => /Title/.test(n.textContent))`), 'the Add menu', 3000).catch(() => {});
    await js(`[...document.querySelectorAll('[role="menuitem"]')].find((n) => /Title/.test(n.textContent))?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.st-bar .st-title'))`), 'the title fields', 4000).catch(() => {});
    await js(`(() => { const n = document.querySelector('.st-bar .st-title'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, 'Imogen Hartley'); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    const titled = await ok(until(async () => (await canvas()).layers === 'window camera title' && (await itemOf('title'))?.text === 'Imogen Hartley', 'the title drawn', 4000));
    if (process.env.RUTBA_VERIFY_CAPTURE) {
      await wait(600);
      fs.mkdirSync(process.env.RUTBA_VERIFY_CAPTURE, { recursive: true });
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'recorder-app.png'), (await win.webContents.capturePage()).toPNG());
    }
    const titleSeen = { layers: (await canvas()).layers, title: (await itemOf('title'))?.text ?? null };
    check('recorder app: a title added from the list is drawn over the sources', titled, JSON.stringify({ titled, ...titleSeen }));

    // Take Picture (owner, 2026-10-20: "there should be a way to capture pictures on camera and from screens etc."):
    // the program as shown, the selected source at its own size, and one after a timer, each one PNG in the folder.
    const pictures = () => { try { return fs.readdirSync(folder).filter((n) => /^Picture .*\.png$/.test(n)).sort(); } catch { return []; } };
    const pngSize = (name) => { const b = fs.readFileSync(path.join(folder, name)); return b.length > 24 && b.readUInt32BE(0) === 0x89504e47 ? { w: b.readUInt32BE(16), h: b.readUInt32BE(20) } : null; };
    const pick = (cls, value) => js(`document.querySelector('.${cls} [data-value="${value}"]')?.click(), 1`);
    const settings = async (fn) => {
      await js(`document.querySelector('.st-gear')?.click(), 1`);
      await until(() => js(`Boolean(document.querySelector('.st-shot-of'))`), 'the settings', 4000).catch(() => {});
      await fn();
      await js(`document.querySelector('.st-settings-done')?.click(), 1`);
      await until(() => js(`!document.querySelector('.st-shot-of')`), 'the settings closed', 4000).catch(() => {});
    };
    const shoot = async (n) => {
      await until(() => js(`(() => { const b = document.querySelector('.st-shoot'); return Boolean(b) && !b.disabled; })()`), 'Take Picture ready', 6000).catch(() => {});
      await js(`document.querySelector('.st-shoot')?.click(), 1`);
      await until(() => pictures().length >= n, 'picture ' + n, 9000).catch(() => {});
      await until(() => { try { return Boolean(pngSize(pictures()[n - 1])); } catch { return false; } }, 'picture ' + n + ' written', 4000).catch(() => {});
      const name = pictures()[n - 1] || '';
      return { name, size: name ? pngSize(name) : null };
    };
    await go('main');
    const asShown = await shoot(1);
    await settings(() => pick('st-shot-of', 'selected'));
    await go('setup');
    await select('camera');
    await go('main');
    const ofSelected = await shoot(2);
    await settings(() => pick('st-shot-timer', '3'));
    await js(`document.querySelector('.st-shoot')?.click(), 1`);
    const counting = await ok(until(() => js(`Boolean(document.querySelector('.st-shot-count')) && /[23]/.test(document.querySelector('.st-shoot')?.textContent || '')`), 'the picture timer', 2500));
    const early = pictures().length;
    await until(() => pictures().length >= 3, 'picture 3', 9000).catch(() => {});
    await wait(300);
    const timed = { name: pictures()[2] || '', size: pictures()[2] ? pngSize(pictures()[2]) : null };
    const bar = await js(`Boolean(document.querySelector('.st-shot .st-shot-open')) && Boolean(document.querySelector('.st-shot .st-shot-edit')) && Boolean(document.querySelector('.st-shot .st-shot-copy')) && (document.querySelector('.st-shot-name')?.textContent || '') === ${JSON.stringify(timed.name)}`);
    check('recorder app: Take Picture keeps a PNG of the program as shown (1280 by 720), of the selected camera at the size of its clip, and one after a 3 second timer, with Open, Edit in Image and Copy',
      asShown.size?.w === 1280 && asShown.size?.h === 720 && ofSelected.size && ofSelected.size.w > 40 && ofSelected.size.w < 1280 && ofSelected.size.h > 40
        && counting && early === 2 && timed.size?.w === ofSelected.size?.w && timed.size?.h === ofSelected.size?.h && bar,
      JSON.stringify({ asShown, ofSelected, counting, early, timed, bar }));
    await settings(async () => { await pick('st-shot-of', 'program'); await pick('st-shot-timer', '0'); });
    // The window drawn with no timer before it is minimised: a minimised window runs a timer a second at most.
    await until(() => js(`(() => { const b = document.querySelector('.st-shoot'); return !document.querySelector('.st-shot-count') && Boolean(b) && !b.disabled; })()`), 'ready to take', 10000).catch(() => {});
    await wait(300);
    // The key (Ctrl+Shift+F8) takes one too, with the window minimised as well as in front.
    // Held through the shortcuts adapter (a stand-in under the checks), and pressed there.
    const keyHeld = Boolean(shortcuts().isRegistered('CommandOrControl+Shift+F8'));
    win.minimize();
    await until(() => win.isMinimized(), 'the window minimised', 5000).catch(() => {});
    shortcuts().press?.('CommandOrControl+Shift+F8');
    const byKey = await ok(until(() => pictures().length >= 4, 'the picture by the key', 20000));
    win.restore();
    const keyTip = await js(`document.querySelector('.st-shoot')?.dataset.tip || ''`);
    check('recorder app: Take Picture\'s key takes a picture with the window minimised, and the button names the key',
      keyHeld && byKey && /Ctrl\+Shift\+F8|Cmd\+Shift\+F8/.test(keyTip), JSON.stringify({ keyHeld, byKey, pictures: pictures().length, keyTip }));
    await wait(300);

    // Record: one button that turns red with the time, one file growing, the scene changed, a picture taken, a pause.
    await js(`document.querySelector('.st-record')?.click(), 1`);
    const recording = await ok(until(() => js(`(() => { const b = document.querySelector('.st-record'); return Boolean(b) && !b.disabled && b.classList.contains('on') && b.getAttribute('aria-pressed') === 'true' && /\\d:\\d\\d/.test(b.textContent); })()`), 'the recording to start', 15000));
    // The check's own clock of what was recorded: from here to Stop, less the pause.
    const recordedFrom = Date.now();
    const files = (re = /\.webm$/) => { try { return fs.readdirSync(folder).filter((n) => re.test(n)).map((n) => ({ name: n, size: fs.statSync(path.join(folder, n)).size })); } catch { return []; } };
    const grew = await ok(until(() => files().some((f) => f.size > 5000), 'the recording on disk', 8000));
    await wait(1000);
    await js(`document.querySelector('.st-scene[data-scene="camera"]')?.click(), 1`);
    const switched = await ok(until(() => js(`document.querySelector('.st-canvas')?.dataset.layers === 'camera'`), 'the camera scene while recording', 4000));
    const capture = async (name) => { if (!process.env.RUTBA_VERIFY_CAPTURE) return; win.webContents.invalidate(); await wait(700); fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await win.webContents.capturePage()).toPNG()); };
    await capture('record-recording.png');
    await js(`document.querySelector('.st-shoot')?.click(), 1`);
    const pictureWhile = await ok(until(() => pictures().length >= 5, 'a picture while recording', 8000));
    await wait(600);
    const pausedFrom = Date.now();
    await js(`document.querySelector('.st-pause')?.click(), 1`);
    const paused = await ok(until(() => js(`/Paused/.test(document.querySelector('.st-badge')?.textContent || '') && document.querySelector('.st-record')?.classList.contains('held')`), 'paused', 8000));
    await wait(2500);
    await js(`document.querySelector('.st-pause')?.click(), 1`);
    const pausedFor = (Date.now() - pausedFrom) / 1000;
    await wait(900);
    check('recorder app: Record turns red with the time and writes one file as it goes; the scene changes, Take Picture takes one, and Pause holds it, all while it records',
      recording && grew && switched && pictureWhile && paused && files().length === 1, JSON.stringify({ recording, grew, switched, pictureWhile, paused, files: files() }));

    // Record again stops: one WebM, picture and sound, as long as it ran without the pause.
    await js(`document.querySelector('.st-record')?.click(), 1`);
    const recordedFor = (Date.now() - recordedFrom) / 1000 - pausedFor;
    const saved = await ok(until(() => js(`Boolean(document.querySelector('.st-saved'))`), 'the saved bar', 30000));
    await capture('record-saved.png');
    const made = files();
    const probe = made.length === 1 ? probeMedia(new Uint8Array(fs.readFileSync(path.join(folder, made[0].name))), { complete: true }) : null;
    const offAgain = await js(`!document.querySelector('.st-record')?.classList.contains('on')`);
    check('recorder app: Record pressed again stops and keeps one recording, named Recording and the time, with picture and sound, 1280 by 720, as long as it recorded and not the pause',
      saved && offAgain && made.length === 1 && /^Recording .*\.webm$/.test(made[0].name) && probe?.hasVideo && probe?.hasAudio && probe.width === 1280 && probe.height === 720 && probe.duration > 2.2 && Math.abs(probe.duration - recordedFor) < 1 && pausedFor >= 2.5,
      JSON.stringify({ saved, offAgain, made, recorded: Math.round(recordedFor * 100) / 100, paused: Math.round(pausedFor * 100) / 100, probe: probe ? { video: probe.hasVideo, audio: probe.hasAudio, w: probe.width, h: probe.height, duration: probe.duration } : null }));

    // Edit in Video: the recording on Video's timeline.
    const before = BrowserWindow.getAllWindows();
    await js(`document.querySelector('.st-saved .st-edit')?.click(), 1`);
    const opened = () => BrowserWindow.getAllWindows().find((w) => !before.includes(w) && !w.isDestroyed()) || null;
    await until(opened, 'a Video window', 8000).catch(() => {});
    const video = opened();
    let edited = false;
    if (video) {
      edited = await ok(until(() => video.webContents.executeJavaScript(`Boolean(document.querySelector('.vd-editor')) && document.querySelectorAll('.vd-clip[data-kind="video"]').length === 1`), 'the recording on the timeline', 15000));
      video.destroy();
    }
    check('recorder app: Edit in Video opens the recording on Video\'s timeline', edited, video ? 'a Video window opened' : 'no window');

    // The Sound only scene: Record says it records sound, and keeps sound alone.
    await js(`document.querySelector('.st-again')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.st-scene[data-scene="sound-only"]'))`), 'the scenes again', 4000).catch(() => {});
    await capture('record-panel.png');
    await js(`document.querySelector('.st-scene[data-scene="sound-only"]')?.click(), 1`);
    const saysSound = await ok(until(() => js(`/sound/i.test(document.querySelector('.st-record')?.textContent || '')`), 'Record sound', 4000));
    await js(`document.querySelector('.st-record')?.click(), 1`);
    const soundOn = await ok(until(() => js(`document.querySelector('.st-record')?.classList.contains('on')`), 'recording sound', 15000));
    await wait(2200);
    await js(`document.querySelector('.st-record')?.click(), 1`);
    await until(() => files(/^Sound .*\.weba$/).length === 1, 'the sound file', 20000).catch(() => {});
    await wait(500);
    const sound = files(/^Sound .*\.weba$/);
    const soundProbe = sound.length === 1 ? probeMedia(new Uint8Array(fs.readFileSync(path.join(folder, sound[0].name))), { complete: true }) : null;
    check('recorder app: with Sound only, Record says it records sound and keeps sound alone, no picture',
      saysSound && soundOn && sound.length === 1 && soundProbe?.hasAudio && !soundProbe?.hasVideo && soundProbe.duration > 1.5,
      JSON.stringify({ saysSound, soundOn, sound, probe: soundProbe ? { video: soundProbe.hasVideo, audio: soundProbe.hasAudio, duration: soundProbe.duration } : null }));

    const complaints = await errorsIn(win);
    check('recorder app: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('recorder app: the check ran to the end', false, err?.message || String(err));
  } finally {
    if (js && kept) await restoreSetup(js, kept).catch(() => {});
  }
}
