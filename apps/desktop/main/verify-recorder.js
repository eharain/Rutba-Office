// The screen recorder (docs/STUDIO.md Phase 2): Video → Record Screen.
//
// Everything recorded here is the suite's own: the window recorded is a
// Documents window this check opens (a check run lists only the suite's
// windows, never the person's desktop), and the screen, the camera and the
// microphone are Chromium's stand-ins (RUTBA_FAKE_MEDIA; the screen's by
// withScreenStandIn below, so these checks run beside other work and with
// the screen locked, and verifyRecorderScreen records the real screen in
// the machine lane), the global key and the tray are
// handed their press by the main process rather than taken from the
// keyboard, and the recordings go to this run's own folder.
//
// Record Screen on the ribbon opens the Recorder for this project (A7), on a
// setup that shows the window, clipped, and the camera in a corner. Record counts down
// from three, both files grow on disk as they are recorded, Pause holds the
// clock, the global key resumes it and stops it, and the recording is copied
// for the launcher while it runs. On stop the takes land on the timeline:
// the screen on the main lane with the region as its crop, the camera above
// it as a mirrored circle in the corner it was dragged to, the voice on a
// sound lane. The takes and an export are read back with probeMedia for
// their length and their tracks; a copy made while recording, recovered,
// settles to the files' real lengths. getDisplayMedia is answered with the
// source chosen; the launcher offers Record a screen; and in Urdu the Recorder
// speaks Urdu right to left with the picture left to right.
// Run alone with RUTBA_VERIFY_ONLY=recorder.

import fs from 'node:fs';
import { waitForExport } from './verify-export-wait.js';
import path from 'node:path';
import { BrowserWindow } from 'electron';
import { buildDocx } from '@rutba/ooxml/build';
import { probeMedia } from '@rutba/media/probe';
import { EXPORT_EXT } from './verify-video-editor.js';
import { readRvid } from '@rutba/studio/rvid';
import { addItem, emptyStack } from '@rutba/studio/stack';
import { seedSetup, restoreSetup, recorderOpened, windowsNow, CHECK_SOUNDS, RECORDING, PRESS_RECORD } from './verify-studio-seed.js';
import { labelKeyed } from './verify-keys.js';

/**
 * What a Recorder still open after Stop shows, for the check's detail: a
 * question it is asking (its own close held to ask Stop recording?), its
 * phase, and whether it shows the saved card (it gave up waiting for the
 * window it recorded for). Null once it has closed.
 */
async function recorderLeft(rw) {
  if (!rw || rw.isDestroyed()) return null;
  return rw.webContents.executeJavaScript(`({
    dialog: document.querySelector('.rw-dialog')?.textContent || null,
    phase: document.querySelector('.st-status')?.dataset.phase || null,
    saved: Boolean(document.querySelector('.st-saved')),
  })`).catch((err) => ({ unreadable: err.message }));
}
import { soundIn } from './verify-measure.js';
import { SIGNAL_WAIT } from './verify-waits.js';

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn, doc
 * @param {{ dir: string }} args the run's scratch folder
 */
export function verifyRecorder(h, args) {
  return withScreenStandIn(() => recorderChecks(h, args));
}

async function recorderChecks(h, { dir }) {
  const { open, check, until, wait, errorsIn, doc } = h;
  const capDir = process.env.RUTBA_VERIFY_CAPTURE || null;
  const capture = async (win, name) => {
    if (!capDir) return;
    try {
      fs.mkdirSync(capDir, { recursive: true });
      win.webContents.invalidate();
      await wait(700);
      const img = await win.webContents.capturePage();
      fs.writeFileSync(path.join(capDir, name), img.toPNG());
    } catch {
      /* a capture is a picture to look at, not a check */
    }
  };
  const folder = path.join(dir, 'recordings');
  const docFile = path.join(dir, 'rec-fixture.docx');
  const ok = (p) => p.then(() => true).catch(() => false);
  const controls = () => globalThis.__rutbaCaptureControls || { on: false };
  let win = null;
  let keptSetup = null;
  try {
    fs.mkdirSync(folder, { recursive: true });
    fs.writeFileSync(docFile, buildDocx({ styles: true, paragraphs: [{ text: 'Recorded by the recorder check.' }, { text: 'A second line, so the page has something on it.' }] }));
    await open('word', docFile);
    win = await open('video');
    const js = (code, gesture = false) => win.webContents.executeJavaScript(code, gesture);
    // What the Recorder will show: the window with a part of it clipped, the camera a circle in the top left, the microphone on.
    await until(() => js(`window.rutbaOffice.capture.sources().then((l) => l.some((x) => x.name.includes('rec-fixture')))`), 'the fixture window to be there', 10000).catch(() => {});
    const source = (await js(`window.rutbaOffice.capture.sources()`)).find((x) => x.name.includes('rec-fixture'));
    let seeded = { ...emptyStack(), sounds: CHECK_SOUNDS() };
    seeded = addItem(seeded, { kind: 'window', ref: source.id, name: source.name, crop: { x: 0.35, y: 0.3, w: 0.5, h: 0.55 } }).stack;
    seeded = addItem(seeded, { kind: 'camera', ref: '', shape: 'circle', mirror: true, box: { x: 0.02, y: 0.03, w: 0.2, h: 0.3556 } }).stack;
    keptSetup = await seedSetup(js, { app: 'recorder', stack: seeded, studio: { size: 720, fps: 30, quality: 'standard', countdown: 3, folder, minimise: true } });

    // The start screen offers it, and the ribbon's Record Screen opens the Recorder for this project.
    const offered = await js(`Boolean(document.querySelector('.vd-start-record'))`);
    const known = windowsNow();
    const pressed = await js(`(() => { const b = ((list) => list.find((n) => ${labelKeyed("Record Screen", { tip: true, start: true })}(n)) || list.find((n) => (n.dataset.tip || '').startsWith('Record Screen')))([...document.querySelectorAll('.rw-ribbon .rw-btn')]); if (!b || b.disabled) return 'no button'; b.click(); return 'clicked'; })()`);
    const { rw, rjs } = await recorderOpened(until, known);
    const says = await rjs(`document.querySelector('.st-into')?.textContent || ''`);
    check('recorder: Video offers Record Screen on its start screen and its ribbon, and it opens the Recorder, saying what it records into',
      offered && pressed === 'clicked' && /Recording into Untitled video/.test(says), JSON.stringify({ offered, pressed, says }));

    const live = await ok(until(() => rjs(`(() => { const l = document.querySelector('.st-canvas')?.dataset.layers || ''; return l.includes('window') && l.includes('camera'); })()`), 'the window and the camera live', 15000));
    check('recorder: the setup shows live in the Recorder, the window and the camera together', live, await rjs(`document.querySelector('.st-canvas')?.dataset.layers || ''`));
    const disk = await rjs(`document.querySelector('.st-status')?.textContent || ''`);
    const space = await js(`window.rutbaOffice.capture.space({ path: ${JSON.stringify(folder)} })`).catch((err) => ({ error: err.message }));
    check('recorder: the disk is checked before recording and the Recorder says how long there is room for', /Room for about \d/.test(disk) && space.free > 0, `${disk}; ${space.free ? `${Math.round(space.free / 1e9)} GB free` : space.error}`);
    await capture(rw, 'recorder-1-setup.png');

    // Record: a countdown from three, then both files growing on disk.
    const startedAt = Date.now();
    await rjs(PRESS_RECORD);
    const counted = await ok(until(() => rjs(`document.querySelector('.rec-count span')?.textContent === '3'`), 'the countdown', 3000));
    const recording = await ok(until(() => rjs(RECORDING), 'the recording to start', 9000));
    const waited = Date.now() - startedAt;
    check('recorder: Record counts down from three, then records', counted && recording && waited >= 2500, JSON.stringify({ counted, recording, waited }));
    const files = () => { try { return fs.readdirSync(folder).filter((n) => /\.(webm|mp4)$/.test(n)).map((n) => ({ name: n, size: fs.statSync(path.join(folder, n)).size })); } catch { return []; } };
    const grown = await ok(until(() => { const f = files(); return f.length === 2 && f.every((x) => x.size > 1000); }, 'the takes on disk', 6000));
    const ctl = controls();
    check('recorder: the screen and the camera are recorded at once, each to a file of its own that grows on disk as it goes, and a global Stop key is held while it runs',
      grown && ctl.on === true && /F10/.test(ctl.keys?.stop || ''), JSON.stringify({ files: files(), on: ctl.on, keys: ctl.keys }));
    await wait(1800);
    await capture(rw, 'recorder-3-recording.png');

    // Pause holds the clock; the global key resumes.
    await rjs(`document.querySelector('.st-pause')?.click(), 1`);
    const held = await ok(until(() => rjs(`Boolean(document.querySelector('.st-pause.on'))`), 'paused', 3000));
    const clockNow = () => rjs(`document.querySelector('.st-record')?.textContent`);
    const t1 = await clockNow();
    await wait(1500);
    const t2 = await clockNow();
    controls().press?.('pause');
    const resumed = await ok(until(() => rjs(`!document.querySelector('.st-pause.on')`), 'resumed', 3000));
    check('recorder: Pause holds the recording and its clock, and the global key resumes it', held && t1 === t2 && resumed, JSON.stringify({ held, t1, t2, resumed }));

    // While it runs, the recording is the window's recovery copy.
    const copyFile = `window-${await js(`document.querySelector('.vd-editor')?.dataset.copy || ''`)}.rvid`;
    const copied = await ok(until(() => doc.recoverable().some((e) => e.file === copyFile), 'the recording\'s recovery copy', 9000));
    let copyBytes = null;
    let copyDoc = null;
    if (copied) {
      copyBytes = (await doc.readCopy({ file: copyFile })).bytes;
      copyDoc = (await readRvid(copyBytes)).doc;
    }
    check('recorder: while it records, the recording is copied for the launcher to offer back after a crash',
      Boolean(copyDoc?.recording?.unfinished && copyDoc.patches.some((p) => p.take === 'screen') && copyDoc.patches.some((p) => p.take === 'camera')),
      copyDoc ? `${copyFile}: ${copyDoc.patches.length} layers, screen ${copyDoc.patches.find((p) => p.take === 'screen')?.timing?.end} s so far` : 'no copy');
    await wait(1500);

    // Stop from the global key: the takes land on the timeline.
    controls().press?.('stop');
    const landed = await ok(until(async () => rw.isDestroyed() && await js(`Boolean(document.querySelector('.vd-clip[data-take="screen"]') && document.querySelector('.vd-clip[data-take="camera"]') && document.querySelector('.vd-clip[data-take="voice"]'))`), 'the takes on the timeline', SIGNAL_WAIT));
    // Not landed in time: what the Recorder shows says why it did not close.
    const left = landed ? null : await recorderLeft(rw);
    const clips = await js(`[...document.querySelectorAll('.vd-clip')].map((n) => ({ kind: n.dataset.kind, take: n.dataset.take, lane: Number(n.dataset.lane), start: Number(n.dataset.start), end: Number(n.dataset.end), inset: n.dataset.inset, crop: n.dataset.crop }))`);
    const screenClip = clips.find((c) => c.take === 'screen');
    const camClip = clips.find((c) => c.take === 'camera');
    const voice = clips.find((c) => c.take === 'voice');
    const inset = (camClip?.inset || '').split(' ');
    check('recorder: on Stop the takes land on the timeline: the screen on the main lane with the region as its crop, the camera on the lane above as a circle in the corner it was dragged to, the voice on a sound lane, together',
      landed && screenClip?.lane === 0 && screenClip.crop.split(' ').length === 4 && camClip?.lane === 1 && inset[4] === 'circle' && Number(inset[0]) < 0.1 && Number(inset[1]) < 0.15 && voice && screenClip.start === camClip.start && voice.start === screenClip.start,
      JSON.stringify(left ? { clips, recorder: left } : clips));
    const released = controls().on === false;
    check('recorder: the global keys are let go when it stops', released, released ? 'released' : 'still held');

    // The takes read back: their length and their tracks.
    const takes = files().map((f) => ({ ...f, probe: probeMedia(new Uint8Array(fs.readFileSync(path.join(folder, f.name))), { complete: true }) }));
    if (capDir) for (const f of files()) { try { fs.copyFileSync(path.join(folder, f.name), path.join(capDir, `${f.name.split(' ')[0]}${path.extname(f.name)}`)); } catch { /* a copy to look at */ } }
    const screenFile = takes.find((f) => /^Screen/.test(f.name));
    const camFile = takes.find((f) => /^Camera/.test(f.name));
    const length = screenClip ? screenClip.end - screenClip.start : 0;
    check('recorder: the screen\'s take is a video as long as its clip, and with the computer\'s sound off no sound of it lands on the timeline',
      Boolean(screenFile?.probe?.hasVideo && Math.abs(screenFile.probe.duration - length) < 0.6 && length > 3 && !clips.some((c) => c.take === 'computer')),
      screenFile ? `${screenFile.name}: ${screenFile.probe?.duration?.toFixed(2)} s for a ${length.toFixed(2)} s clip, tracks ${(screenFile.probe?.tracks || []).map((x) => x.type).join(', ')}` : 'no file');
    check('recorder: the camera\'s take carries the voice, on the same clock as the screen',
      Boolean(camFile?.probe?.hasVideo && camFile.probe.hasAudio && Math.abs(camFile.probe.duration - (screenFile?.probe?.duration || 0)) < 0.5),
      camFile ? `${camFile.name}: ${camFile.probe?.duration?.toFixed(2)} s, tracks ${(camFile.probe?.tracks || []).map((x) => `${x.type} ${x.codec}`).join(', ')}` : 'no file');

    // The stage draws the camera over the screen: the inset's box on the
    // stage at 1.2 s is kept, and compared with the same box once the camera
    // has been moved away (below), where the screen alone is drawn. Inside
    // the circle the camera changed it; in the box's corners, outside the
    // circle, nothing did. (The screen is a stand-in as the camera is, the
    // same kind of moving picture, so a colour read at one point could not
    // tell the two apart.)
    await js(`(() => { const s = document.querySelector('.vd-scrub'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(s, '1.2'); s.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    const ins = inset.slice(0, 4).map(Number);
    // The box read until two reads a fifth of a second apart agree: the frame at 1.2 s drawn, and drawn once.
    const boxRead = (name) => `(() => {
      const c = document.querySelector('.vd-canvas'); if (!c) return null;
      const ins = ${JSON.stringify(ins)};
      const x = Math.floor(c.width * ins[0]); const y = Math.floor(c.height * ins[1]); const w = Math.floor(c.width * ins[2]); const h = Math.floor(c.height * ins[3]);
      const d = c.getContext('2d').getImageData(x, y, w, h).data;
      let sum = 0; for (let i = 0; i < d.length; i += 97) sum = (sum * 31 + d[i]) % 1000000007;
      (window.__rutbaInsetBox ||= {})[${JSON.stringify(name)}] = { w, h, d };
      return { w, h, sum };
    })()`;
    const steady = async (name, what) => {
      let last = null;
      let got = null;
      await until(async () => {
        await wait(200);
        got = await js(boxRead(name));
        const same = Boolean(got && last && got.sum === last.sum && got.w === last.w);
        last = got;
        return same;
      }, what, 15000).catch(() => {});
      return got;
    };
    const withCamera = await steady('camera', 'the stage drawn at 1.2 s');
    await capture(win, 'recorder-4-landed.png');

    // The camera picked: its own properties, moved to another corner and back.
    const camId = await js(`document.querySelector('.vd-clip[data-take="camera"]')?.dataset.id`);
    await js(`(() => { const n = document.querySelector('.vd-clip[data-take="camera"]'); const r = n.getBoundingClientRect(); const o = { bubbles: true, cancelable: true, button: 0, clientX: r.left + 20, clientY: r.top + r.height / 2, pointerId: 1 }; n.dispatchEvent(new PointerEvent('pointerdown', o)); window.dispatchEvent(new PointerEvent('pointerup', o)); return 1; })()`);
    const props = await ok(until(() => js(`Boolean(document.querySelector('.vd-inset-props'))`), 'the inset\'s properties', 8000));
    await js(`document.querySelector('.vd-corner[data-corner="br"]')?.click(), 1`);
    const moved = await ok(until(() => js(`(() => { const v = (document.querySelector('.vd-clip[data-id="${camId}"]')?.dataset.inset || '').split(' ').map(Number); return v[0] > 0.6 && v[1] > 0.5; })()`), 'the inset in the bottom right', 8000));
    await capture(win, 'recorder-5-inset.png');
    check('recorder: the camera stays an inset that can be moved after: Picture in picture in its properties puts it in another corner', props && moved, JSON.stringify({ props, moved }));

    const screenOnly = moved ? await steady('screen', 'the stage drawn again without the camera there') : null;
    const cut = withCamera && screenOnly ? await js(`(() => {
      const a = window.__rutbaInsetBox?.camera; const b = window.__rutbaInsetBox?.screen;
      if (!a || !b || a.w !== b.w || a.h !== b.h) return null;
      const { w, h } = a;
      const r = Math.min(w, h) / 2;
      let inside = 0, insideChanged = 0, corners = 0, cornersChanged = 0;
      for (let y = 0; y < h; y += 2) {
        for (let x = 0; x < w; x += 2) {
          const dist = Math.hypot(x + 0.5 - w / 2, y + 0.5 - h / 2) / r;
          const at = (y * w + x) * 4;
          const changed = Math.abs(a.d[at] - b.d[at]) + Math.abs(a.d[at + 1] - b.d[at + 1]) + Math.abs(a.d[at + 2] - b.d[at + 2]) > 48;
          if (dist < 0.85) { inside++; if (changed) insideChanged++; } else if (dist > 1.1) { corners++; if (changed) cornersChanged++; }
        }
      }
      delete window.__rutbaInsetBox;
      return { size: [w, h], inside: Math.round((insideChanged / inside) * 1000) / 1000, corners: corners ? Math.round((cornersChanged / corners) * 1000) / 1000 : null };
    })()`) : null;
    // Both takes are the same fake camera, so how much of the circle differs from what lies under it depends on how far apart they began (0.08 to 0.34 seen). The corners are what prove the cut.
    check('recorder: the stage draws the camera cut to its circle over the screen', Boolean(cut && cut.inside > 0.03 && cut.corners != null && cut.corners < 0.02), JSON.stringify(cut));

    // Export with the camera over the screen and the voice under it, read back.
    const ext = await js(EXPORT_EXT);
    const target = path.join(dir, `recorder-out.${ext}`);
    const projectLength = await js(`Number(document.querySelector('.vd-editor')?.dataset.duration || 0)`);
    // An answer an earlier block queued for a Save that never asked (a project
    // that already had its path) would be taken by this dialog instead.
    const queue = globalThis.__rutbaCheckDialogAnswers;
    if (queue) for (let i = queue.length - 1; i >= 0; i--) if (queue[i].kind === 'save') queue.splice(i, 1);
    queue?.push({ kind: 'save', answer: target });
    await js(`(() => { ((list) => list.find((n) => ${labelKeyed("Home")}(n)) || list.find((n) => n.textContent.trim() === 'Home'))([...document.querySelectorAll('.rw-tab')])?.click(); return 1; })()`);
    await wait(200);
    await js(`(() => { const b = ((list) => list.find((n) => ${labelKeyed("Export", { tip: true, start: true })}(n) && !n.disabled) || list.find((n) => (n.dataset.tip || '').startsWith('Export') && !n.disabled))([...document.querySelectorAll('.rw-ribbon .rw-btn')]); b?.click(); return Boolean(b); })()`);
    await ok(until(() => js(`Boolean(document.querySelector('.vd-export-go'))`), 'the Export box', 4000));
    await until(() => js(`document.querySelector('.vd-export-go')?.disabled === false`), 'the export format found', 6000).catch(() => {});
    await js(`document.querySelector('.vd-export-go')?.click(), 1`);
    // While the export moves, however slowly; one that stops is cancelled (verify-export-wait.js).
    const written = (await waitForExport(win, () => { try { return fs.statSync(target).size > 4000; } catch { return false; } })).done;
    await ok(until(() => js(`!document.querySelector('.vd-export-busy')`), 'the progress to close', 15000));
    let probe = null;
    if (written) { await wait(400); probe = probeMedia(new Uint8Array(fs.readFileSync(target)), { complete: true }); }
    check('recorder: the recording exports as one video, as long as the project, with its picture and its sound',
      Boolean(probe?.hasVideo && probe?.hasAudio && Math.abs(probe.duration - projectLength) < 0.75),
      probe ? `${path.basename(target)}: ${probe.duration?.toFixed(2)} s for ${projectLength.toFixed(2)} s; tracks ${(probe.tracks || []).map((x) => `${x.type} ${x.codec}`).join(', ')}` : 'no file');
    // A track can be there and silent: the stand-in microphone beeps, so its export has something to hear.
    const heard = written ? await soundIn(js, target) : { error: 'no file' };
    check('recorder: the sound in the export can be heard', !heard.error && heard.peak > 0.02, JSON.stringify(heard));

    // The copy made while recording, recovered: its takes settle to the files' real lengths.
    if (copyBytes) {
      await doc.keepCopy({ key: 'recordercheck', kind: 'video', name: 'recorder check', path: null, bytes: copyBytes });
      const known = new Set(BrowserWindow.getAllWindows().map((w) => w.id));
      await js(`window.rutbaOffice.win.create({ app: 'video', query: { recover: 'window-recordercheck.rvid' } }).then(() => 1)`);
      const fresh = () => BrowserWindow.getAllWindows().find((w) => !known.has(w.id) && !w.isDestroyed()) || null;
      await ok(until(() => Boolean(fresh()), 'the recovered window', 8000));
      const rescued = fresh();
      let back = null;
      if (rescued) {
        await new Promise((r) => (rescued.webContents.isLoading() ? rescued.webContents.once('did-finish-load', r) : r()));
        const rjs = (code) => rescued.webContents.executeJavaScript(code);
        const want = length;
        await ok(until(() => rjs(`(() => { const n = document.querySelector('.vd-clip[data-take="screen"]'); return Boolean(n) && Math.abs(Number(n.dataset.end) - Number(n.dataset.start) - ${want}) < 0.6; })()`), 'the recovered takes settled', 15000));
        back = await rjs(`[...document.querySelectorAll('.vd-clip')].map((n) => ({ take: n.dataset.take, len: Math.round((Number(n.dataset.end) - Number(n.dataset.start)) * 100) / 100 }))`);
        rescued.destroy();
      }
      await doc.discardRecovery({ file: 'window-recordercheck.rvid' });
      const snap = copyDoc.patches.find((p) => p.take === 'screen').timing;
      const screenBack = back?.find((c) => c.take === 'screen');
      check('recorder: a recording recovered after a crash opens in Video with its takes as long as their files, longer than when it was copied',
        Boolean(screenBack && Math.abs(screenBack.len - length) < 0.1 && screenBack.len > snap.end - snap.start + 0.5 && back.some((c) => c.take === 'camera')),
        JSON.stringify({ copied: snap.end - snap.start, back }));
    }

    // The computer's own sound: getDisplayMedia is answered with the source chosen, and refused without one.
    const display = await js(`(async () => {
      const id = (await window.rutbaOffice.capture.sources()).find((s) => s.name.includes('rec-fixture'))?.id;
      let refused = false;
      try { const s0 = await navigator.mediaDevices.getDisplayMedia({ video: true }); s0.getTracks().forEach((t) => t.stop()); } catch { refused = true; }
      const answer = await window.rutbaOffice.capture.choose({ id, audio: false });
      try {
        const s = await navigator.mediaDevices.getDisplayMedia({ video: true });
        const n = s.getVideoTracks().length;
        s.getTracks().forEach((t) => t.stop());
        return { refused, chosen: n, loopback: answer.loopback };
      } catch (err) { return { refused, error: err.message, loopback: answer.loopback }; }
    })()`, true);
    check('recorder: the main process answers getDisplayMedia with the source the window chose (with the computer\'s sound on Windows), and refuses a request nothing was chosen for',
      display.refused && display.chosen === 1 && display.loopback === (process.platform === 'win32'), JSON.stringify(display));

    const complaints = await errorsIn(win);
    check('recorder: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');

    // Studio's Start something offers Record a screen: the Recorder.
    const home = await open('home');
    const hjs = (code) => home.webContents.executeJavaScript(code);
    // The launcher drawn before it is read: beside other work it was read while still empty.
    const card = await ok(until(() => hjs(`Boolean(document.querySelector('.home-record'))`), 'the launcher\'s Record a screen', 15000));
    const known2 = new Set(BrowserWindow.getAllWindows().map((w) => w.id));
    await hjs(`document.querySelector('.home-record')?.click(), 1`);
    const fresh2 = () => BrowserWindow.getAllWindows().find((w) => !known2.has(w.id) && !w.isDestroyed()) || null;
    await ok(until(() => Boolean(fresh2()), 'the new Recorder window', 8000));
    const nv = fresh2();
    let opened = false;
    if (nv) {
      await new Promise((r) => (nv.webContents.isLoading() ? nv.webContents.once('did-finish-load', r) : r()));
      opened = /[?&]app=recorder\b/.test(nv.webContents.getURL());
      await wait(300);
      nv.destroy();
    }
    check('recorder: the launcher offers Record a screen, which opens the Recorder', card && opened, JSON.stringify({ card, opened }));
  } catch (err) {
    check('recorder: the checks ran', false, err.message);
  }

  // In Urdu: the box speaks Urdu right to left, and the picture stays left to right.
  if (win && !win.isDestroyed()) {
    const js = (code) => win.webContents.executeJavaScript(code);
    const was = await js(`window.rutbaOffice.store.get({ key: 'language', fallback: null })`).catch(() => null);
    try {
      await js(`window.rutbaOffice.store.set({ key: 'language', value: 'ur' }).then(() => 1)`);
      const ur = await open('recorder');
      const ujs = (code) => ur.webContents.executeJavaScript(code);
      await ok(until(() => ujs(`Boolean(document.querySelector('.st-panels') && document.querySelector('.st-frame'))`), 'the Recorder in Urdu', 10000));
      await wait(1200);
      const frame = await ujs(`({ dir: document.documentElement.dir, picture: getComputedStyle(document.querySelector('.st-frame')).direction, english: [...document.querySelectorAll('.st-top button, .st-record, .st-scene-foot')].map((n) => n.textContent.trim()).filter((s) => /^(Record|Set up|Picture|Camera|Microphone|Screen|Window)$/.test(s)) })`);
      await capture(ur, 'recorder-6-urdu.png');
      check('recorder: in Urdu the recorder speaks Urdu, right to left, and the picture stays left to right', frame.dir === 'rtl' && frame.picture === 'ltr' && frame.english.length === 0, JSON.stringify(frame));
      ur.destroy();
    } catch (err) {
      check('recorder: the Urdu checks ran', false, err.message);
    } finally {
      await js(`window.rutbaOffice.store.set({ key: 'language', value: ${JSON.stringify(was ?? null)} }).then(() => 1)`).catch(() => {});
      if (keptSetup) await restoreSetup(js, keptSetup).catch(() => {});
    }
  }
}

/**
 * Run `checks` with the screen stood in for: while it runs, the shell lists
 * this application's own windows and records Chromium's stand-in screen for
 * them (globalThis.__rutbaScreenStandIn, read by capture.sources), so the
 * recorder's checks need nothing of the desktop. Put back as it was after.
 */
export async function withScreenStandIn(checks) {
  const was = globalThis.__rutbaScreenStandIn;
  globalThis.__rutbaScreenStandIn = true;
  try {
    return await checks();
  } finally {
    globalThis.__rutbaScreenStandIn = was;
  }
}

/**
 * The recorder on the real screen, in the machine lane: the checks above
 * record a stand-in for the screen (Chromium's own, as the camera and the
 * microphone are), so they run beside other work and with the screen locked.
 * This one captures a window as the computer itself draws it: a Documents
 * window this check opens is listed from the computer's own list of windows,
 * shows live in the frame with its white page, and two seconds of it are
 * recorded to a take that lands on the timeline. Run alone with
 * RUTBA_VERIFY_ONLY=recorder-screen.
 */
export async function verifyRecorderScreen(h, { dir }) {
  const { open, check, until, wait } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  const folder = path.join(dir, 'real-screen-recordings');
  const docFile = path.join(dir, 'rec-real-screen.docx');
  const standIn = globalThis.__rutbaScreenStandIn;
  globalThis.__rutbaScreenStandIn = false;
  let js = null;
  let settingsWere = null;
  try {
    fs.writeFileSync(docFile, buildDocx({ styles: true, paragraphs: [{ text: 'Recorded from the screen itself.' }] }));
    await open('word', docFile);
    const win = await open('video');
    js = (code) => win.webContents.executeJavaScript(code);
    const listed = await ok(until(() => js(`window.rutbaOffice.capture.sources().then((l) => l.some((x) => x.name.includes('rec-real-screen')))`), 'the window in the computer\'s list', 15000));
    const source = (await js(`window.rutbaOffice.capture.sources()`)).find((x) => x.name.includes('rec-real-screen'));
    const stack = addItem({ ...emptyStack(), sounds: CHECK_SOUNDS(false) }, { kind: 'window', ref: source.id, name: source.name }).stack;
    fs.mkdirSync(folder, { recursive: true });
    settingsWere = await seedSetup(js, { app: 'recorder', stack, studio: { size: 720, fps: 30, quality: 'standard', countdown: 0, folder, minimise: false } });
    const known = windowsNow();
    await js(`(() => { const b = ((list) => list.find((n) => ${labelKeyed("Record Screen", { tip: true, start: true })}(n)) || list.find((n) => (n.dataset.tip || '').startsWith('Record Screen')))([...document.querySelectorAll('.rw-ribbon .rw-btn')]); b?.click(); return Boolean(b); })()`);
    const { rw, rjs } = await recorderOpened(until, known);
    const live = await ok(until(() => rjs(`document.querySelector('.st-canvas')?.dataset.layers === 'window'`), 'the window live', 15000));
    await wait(800);
    // The window's own picture: its page is white, a stand-in's pattern is not.
    const page = live ? await rjs(`(() => {
      const c = document.querySelector('.st-canvas');
      const d = c.getContext('2d').getImageData(Math.floor(c.width * 0.4), Math.floor(c.height * 0.6), 16, 16).data; let s = 0; for (let i = 0; i < d.length; i += 4) s += d[i] + d[i + 1] + d[i + 2];
      return { size: [c.width, c.height], light: Math.round(s / (d.length / 4)) };
    })()`) : null;
    await rjs(PRESS_RECORD);
    const recording = await ok(until(() => rjs(RECORDING), 'the recording to start', 15000));
    await wait(2000);
    await rjs(`document.querySelector('.st-record')?.click(), 1`);
    const landed = await ok(until(async () => rw.isDestroyed() && await js(`Boolean(document.querySelector('.vd-clip[data-take="screen"]'))`), 'the take on the timeline', 30000));
    const left = landed ? null : await recorderLeft(rw);
    const made = (() => { try { return fs.readdirSync(folder).filter((n) => /^Screen .*\.(webm|mp4)$/.test(n)); } catch { return []; } })();
    const probe = made.length === 1 ? probeMedia(new Uint8Array(fs.readFileSync(path.join(folder, made[0]))), { complete: true }) : null;
    check('recorder screen: a window of the computer\'s own screen, picked from its list, shows live with its own picture and records to a take on the timeline',
      listed && live && page?.light > 600 && recording && landed && probe?.hasVideo && probe.duration > 1.5,
      JSON.stringify({ listed, live, page, recording, landed, made, probe: probe ? { video: probe.hasVideo, duration: probe.duration } : null, ...(left ? { recorder: left } : {}) }));
  } catch (err) {
    check('recorder screen: the check ran to the end', false, err?.message || String(err));
  } finally {
    globalThis.__rutbaScreenStandIn = standIn;
    if (js && settingsWere) await restoreSetup(js, settingsWere).catch(() => {});
  }
}
