// The screen recorder (docs/STUDIO.md Phase 2): Video → Record Screen.
//
// Everything recorded here is the suite's own: the window recorded is a
// Documents window this check opens (a check run lists only the suite's
// windows, never the person's desktop), the camera and the microphone are
// Chromium's stand-ins (RUTBA_FAKE_MEDIA), the global key and the tray are
// handed their press by the main process rather than taken from the
// keyboard, and the recordings go to this run's own folder.
//
// The recorder opens from the ribbon, lists the window, shows it live with
// the camera's inset where it will sit and the microphone's level moving;
// the inset is dragged to a corner and a region is drawn. Record counts down
// from three, both files grow on disk as they are recorded, Pause holds the
// clock, the global key resumes it and stops it, and the recording is copied
// for the launcher while it runs. On stop the takes land on the timeline:
// the screen on the main lane with the region as its crop, the camera above
// it as a mirrored circle in the corner it was dragged to, the voice on a
// sound lane. The takes and an export are read back with probeMedia for
// their length and their tracks; a copy made while recording, recovered,
// settles to the files' real lengths. getDisplayMedia is answered with the
// source chosen; the launcher offers Record a screen; and in Urdu the box
// speaks Urdu right to left with the picture left to right.
// Run alone with RUTBA_VERIFY_ONLY=recorder.

import fs from 'node:fs';
import path from 'node:path';
import { BrowserWindow } from 'electron';
import { buildDocx } from '@rutba/ooxml/build';
import { probeMedia } from '@rutba/media/probe';
import { EXPORT_EXT } from './verify-video-editor.js';
import { readRvid } from '@rutba/studio/rvid';

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn, doc
 * @param {{ dir: string }} args the run's scratch folder
 */
export async function verifyRecorder(h, { dir }) {
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
  let settingsWere = null;
  try {
    fs.writeFileSync(docFile, buildDocx({ styles: true, paragraphs: [{ text: 'Recorded by the recorder check.' }, { text: 'A second line, so the page has something on it.' }] }));
    await open('word', docFile);
    win = await open('video');
    const js = (code, gesture = false) => win.webContents.executeJavaScript(code, gesture);
    settingsWere = await js(`window.rutbaOffice.store.get({ key: 'recorder.settings', fallback: null })`);
    await js(`window.rutbaOffice.store.set({ key: 'recorder.settings', value: ${JSON.stringify({ source: 'window', fps: 30, cap: 720, quality: 'standard', countdown: 3, camera: true, mirror: true, corner: 'br', insetSize: 'medium', shape: 'circle', mic: true, system: false, minimise: true, folder })} }).then(() => 1)`);

    // The start screen offers it, and the ribbon's Record Screen opens it.
    const offered = await js(`Boolean(document.querySelector('.vd-start-record'))`);
    const pressed = await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || '').startsWith('Record Screen')); if (!b || b.disabled) return 'no button'; b.click(); return 'clicked'; })()`);
    const shown = await ok(until(() => js(`Boolean(document.querySelector('.rec-dialog'))`), 'the recorder', 6000));
    check('recorder: Video offers Record Screen on its start screen and its ribbon, and it opens the recorder',
      offered && pressed === 'clicked' && shown, JSON.stringify({ offered, pressed, shown }));

    // The window to record, picked from the list: live, with the camera inset and the microphone's level.
    const listed = await ok(until(() => js(`[...document.querySelectorAll('.rec-source')].some((n) => (n.dataset.tip || '').includes('rec-fixture'))`), 'the fixture window in the list', 10000));
    await js(`(() => { const b = [...document.querySelectorAll('.rec-source')].find((n) => (n.dataset.tip || '').includes('rec-fixture')); b?.click(); return Boolean(b); })()`);
    const live = await ok(until(() => js(`(() => { const v = document.querySelector('.rec-screen'); return Boolean(v && v.videoWidth > 0 && document.querySelector('.rec-source.on')?.dataset.tip.includes('rec-fixture')); })()`), 'the window live', 12000));
    const cam = await ok(until(() => js(`(() => { const v = document.querySelector('.rec-inset .rec-cam'); return Boolean(v && v.videoWidth > 0); })()`), 'the camera live', 15000));
    const heard = await ok(until(() => js(`Number(document.querySelector('.rec-meter-mic')?.dataset.peak || 0) > 0`), 'the microphone level', 8000));
    const sized = await js(`(() => { const v = document.querySelector('.rec-screen'); return v ? [v.videoWidth, v.videoHeight] : null; })()`);
    check('recorder: a window picked from the list shows live in the frame, with the camera in its inset and the microphone\'s level moving',
      listed && live && cam && heard, JSON.stringify({ listed, live, cam, heard, sized }));
    const disk = await js(`document.querySelector('.rec-disk')?.textContent || ''`);
    const space = await js(`window.rutbaOffice.capture.space({ path: ${JSON.stringify(folder)} })`).catch((err) => ({ error: err.message }));
    check('recorder: the disk is checked before recording and the box says how long there is room for', /\d/.test(disk) && space.free > 0, `${disk}; ${space.free ? `${Math.round(space.free / 1e9)} GB free` : space.error}`);

    // The camera dragged to the top left corner, where it settles.
    const dragged = await js(`(async () => {
      const n = document.querySelector('.rec-inset');
      const f = document.querySelector('.rec-frame').getBoundingClientRect();
      if (!n) return 'no inset';
      const r = n.getBoundingClientRect();
      const x0 = r.left + r.width / 2; const y0 = r.top + r.height / 2;
      const x1 = f.left + f.width * 0.12; const y1 = f.top + f.height * 0.15;
      n.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, cancelable: true, button: 0, clientX: x0, clientY: y0, pointerId: 1 }));
      for (let k = 1; k <= 8; k++) {
        window.dispatchEvent(new PointerEvent('pointermove', { bubbles: true, clientX: x0 + ((x1 - x0) * k) / 8, clientY: y0 + ((y1 - y0) * k) / 8, pointerId: 1 }));
        await new Promise((res) => setTimeout(res, 25));
      }
      window.dispatchEvent(new PointerEvent('pointerup', { bubbles: true, clientX: x1, clientY: y1, pointerId: 1 }));
      return 'dragged';
    })()`);
    const cornered = await ok(until(() => js(`document.querySelector('.rec-inset')?.dataset.corner === 'tl'`), 'the inset in the top left', 4000));
    const kept = await js(`window.rutbaOffice.store.get({ key: 'recorder.settings', fallback: null }).then((s) => s?.corner)`);
    check('recorder: the camera\'s inset is dragged in the preview and settles in the nearest corner, remembered for next time', dragged === 'dragged' && cornered && kept === 'tl', JSON.stringify({ dragged, cornered, kept }));

    // A region drawn on the window: recorded whole, kept as a crop.
    await js(`(() => { document.querySelector('.rec-kinds button[data-value="region"]')?.click(); return 1; })()`);
    await ok(until(() => js(`document.querySelector('.rec-frame')?.classList.contains('drawing') && document.querySelector('.rec-source.on') && document.querySelector('.rec-screen').videoWidth > 0`), 'region mode', 8000));
    await wait(400);
    const drawn = await js(`(async () => {
      const f = document.querySelector('.rec-frame');
      const r = f.getBoundingClientRect();
      const at = (x, y) => ({ bubbles: true, cancelable: true, button: 0, clientX: r.left + r.width * x, clientY: r.top + r.height * y, pointerId: 1 });
      f.dispatchEvent(new PointerEvent('pointerdown', at(0.35, 0.3)));
      for (let k = 1; k <= 6; k++) { window.dispatchEvent(new PointerEvent('pointermove', at(0.35 + 0.5 * k / 6, 0.3 + 0.55 * k / 6))); await new Promise((res) => setTimeout(res, 25)); }
      window.dispatchEvent(new PointerEvent('pointerup', at(0.85, 0.85)));
      return 'drawn';
    })()`);
    const region = await ok(until(() => js(`Boolean(document.querySelector('.rec-region')) && Boolean(document.querySelector('.rec-clear-region'))`), 'the region', 4000));
    check('recorder: a region is drawn by dragging on the picture, and shown as the part to be kept', drawn === 'drawn' && region, `${drawn}; ${region ? 'region shown' : 'no region'}`);
    await capture(win, 'recorder-1-setup.png');

    // Record: a countdown from three, then both files growing on disk.
    const startedAt = Date.now();
    await js(`document.querySelector('.rec-go')?.click(), 1`);
    const counted = await ok(until(() => js(`document.querySelector('.rec-count span')?.textContent === '3'`), 'the countdown', 3000));
    if (counted) await capture(win, 'recorder-2-countdown.png');
    const recording = await ok(until(() => js(`(() => { const b = document.querySelector('.rec-stop'); return Boolean(b) && !b.disabled; })()`), 'the recording to start', 9000));
    const waited = Date.now() - startedAt;
    check('recorder: Record counts down from three, then records', counted && recording && waited >= 2500, JSON.stringify({ counted, recording, waited }));
    const files = () => { try { return fs.readdirSync(folder).filter((n) => /\.(webm|mp4)$/.test(n)).map((n) => ({ name: n, size: fs.statSync(path.join(folder, n)).size })); } catch { return []; } };
    const grown = await ok(until(() => { const f = files(); return f.length === 2 && f.every((x) => x.size > 1000); }, 'the takes on disk', 6000));
    const ctl = controls();
    check('recorder: the screen and the camera are recorded at once, each to a file of its own that grows on disk as it goes, and a global Stop key is held while it runs',
      grown && ctl.on === true && /F10/.test(ctl.keys?.stop || ''), JSON.stringify({ files: files(), on: ctl.on, keys: ctl.keys }));
    await wait(1800);
    await capture(win, 'recorder-3-recording.png');

    // Pause holds the clock; the global key resumes.
    await js(`document.querySelector('.rec-pause')?.click(), 1`);
    const held = await ok(until(() => js(`Boolean(document.querySelector('.rec-chip')) && document.querySelector('.rec-frame').classList.contains('held')`), 'paused', 3000));
    const t1 = await js(`document.querySelector('.rec-time')?.textContent`);
    await wait(1500);
    const t2 = await js(`document.querySelector('.rec-time')?.textContent`);
    controls().press?.('pause');
    const resumed = await ok(until(() => js(`!document.querySelector('.rec-chip')`), 'resumed', 3000));
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
    const landed = await ok(until(() => js(`Boolean(document.querySelector('.vd-clip[data-take="screen"]') && document.querySelector('.vd-clip[data-take="camera"]') && document.querySelector('.vd-clip[data-take="voice"]')) && !document.querySelector('.rec-dialog')`), 'the takes on the timeline', 20000));
    const clips = await js(`[...document.querySelectorAll('.vd-clip')].map((n) => ({ kind: n.dataset.kind, take: n.dataset.take, lane: Number(n.dataset.lane), start: Number(n.dataset.start), end: Number(n.dataset.end), inset: n.dataset.inset, crop: n.dataset.crop }))`);
    const screenClip = clips.find((c) => c.take === 'screen');
    const camClip = clips.find((c) => c.take === 'camera');
    const voice = clips.find((c) => c.take === 'voice');
    const inset = (camClip?.inset || '').split(' ');
    check('recorder: on Stop the takes land on the timeline: the screen on the main lane with the region as its crop, the camera on the lane above as a circle in the corner it was dragged to, the voice on a sound lane, together',
      landed && screenClip?.lane === 0 && screenClip.crop.split(' ').length === 4 && camClip?.lane === 1 && inset[4] === 'circle' && Number(inset[0]) < 0.1 && Number(inset[1]) < 0.15 && voice && screenClip.start === camClip.start && voice.start === screenClip.start,
      JSON.stringify(clips));
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

    // The stage draws the camera over the screen.
    await js(`(() => { const s = document.querySelector('.vd-scrub'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(s, '1.2'); s.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await wait(900);
    const px = await js(`(() => {
      const c = document.querySelector('.vd-canvas'); if (!c) return null;
      const g = c.getContext('2d');
      const avg = (fx, fy) => { const d = g.getImageData(Math.floor(c.width * fx) - 4, Math.floor(c.height * fy) - 4, 8, 8).data; let r = 0, gg = 0, b = 0; for (let i = 0; i < d.length; i += 4) { r += d[i]; gg += d[i + 1]; b += d[i + 2]; } const n = d.length / 4; return [Math.round(r / n), Math.round(gg / n), Math.round(b / n)]; };
      const ins = ${JSON.stringify(inset.slice(0, 4).map(Number))};
      return { inset: avg(ins[0] + ins[2] / 2, ins[1] + ins[3] / 2), corner: avg(0.004, 0.006), screen: avg(0.6, 0.6) };
    })()`);
    const diff = (a, b) => (a && b ? Math.abs(a[0] - b[0]) + Math.abs(a[1] - b[1]) + Math.abs(a[2] - b[2]) : 0);
    check('recorder: the stage draws the camera cut to its circle over the screen', Boolean(px && diff(px.inset, px.screen) > 30 && diff(px.inset, px.corner) > 30), JSON.stringify(px));
    await capture(win, 'recorder-4-landed.png');

    // The camera picked: its own properties, moved to another corner and back.
    const camId = await js(`document.querySelector('.vd-clip[data-take="camera"]')?.dataset.id`);
    await js(`(() => { const n = document.querySelector('.vd-clip[data-take="camera"]'); const r = n.getBoundingClientRect(); const o = { bubbles: true, cancelable: true, button: 0, clientX: r.left + 20, clientY: r.top + r.height / 2, pointerId: 1 }; n.dispatchEvent(new PointerEvent('pointerdown', o)); window.dispatchEvent(new PointerEvent('pointerup', o)); return 1; })()`);
    const props = await ok(until(() => js(`Boolean(document.querySelector('.vd-inset-props'))`), 'the inset\'s properties', 4000));
    await js(`document.querySelector('.vd-corner[data-corner="br"]')?.click(), 1`);
    const moved = await ok(until(() => js(`(() => { const v = (document.querySelector('.vd-clip[data-id="${camId}"]')?.dataset.inset || '').split(' ').map(Number); return v[0] > 0.6 && v[1] > 0.5; })()`), 'the inset in the bottom right', 4000));
    await capture(win, 'recorder-5-inset.png');
    check('recorder: the camera stays an inset that can be moved after: Picture in picture in its properties puts it in another corner', props && moved, JSON.stringify({ props, moved }));

    // Export with the camera over the screen and the voice under it, read back.
    const ext = await js(EXPORT_EXT);
    const target = path.join(dir, `recorder-out.${ext}`);
    const projectLength = await js(`Number(document.querySelector('.vd-editor')?.dataset.duration || 0)`);
    // An answer an earlier block queued for a Save that never asked (a project
    // that already had its path) would be taken by this dialog instead.
    const queue = globalThis.__rutbaCheckDialogAnswers;
    if (queue) for (let i = queue.length - 1; i >= 0; i--) if (queue[i].kind === 'save') queue.splice(i, 1);
    queue?.push({ kind: 'save', answer: target });
    await js(`(() => { [...document.querySelectorAll('.rw-tab')].find((n) => n.textContent.trim() === 'Home')?.click(); return 1; })()`);
    await wait(200);
    await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || '').startsWith('Export') && !n.disabled); b?.click(); return Boolean(b); })()`);
    await ok(until(() => js(`Boolean(document.querySelector('.vd-export-go'))`), 'the Export box', 4000));
    await until(() => js(`document.querySelector('.vd-export-go')?.disabled === false`), 'the export format found', 6000).catch(() => {});
    await js(`document.querySelector('.vd-export-go')?.click(), 1`);
    const written = await ok(until(() => { try { return fs.statSync(target).size > 4000; } catch { return false; } }, 'the exported file', Math.round(projectLength * 1000) + 30000));
    await ok(until(() => js(`!document.querySelector('.vd-export-busy')`), 'the progress to close', 8000));
    let probe = null;
    if (written) { await wait(400); probe = probeMedia(new Uint8Array(fs.readFileSync(target)), { complete: true }); }
    check('recorder: the recording exports as one video, as long as the project, with its picture and its sound',
      Boolean(probe?.hasVideo && probe?.hasAudio && Math.abs(probe.duration - projectLength) < 0.75),
      probe ? `${path.basename(target)}: ${probe.duration?.toFixed(2)} s for ${projectLength.toFixed(2)} s; tracks ${(probe.tracks || []).map((x) => `${x.type} ${x.codec}`).join(', ')}` : 'no file');

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
      const id = [...document.querySelectorAll('.rec-source')][0]?.dataset.source || (await window.rutbaOffice.capture.sources()).find((s) => s.name.includes('rec-fixture'))?.id;
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

    // The launcher's Start something offers Record a screen: a new video on the recorder.
    const home = await open('home');
    const hjs = (code) => home.webContents.executeJavaScript(code);
    const card = await hjs(`Boolean(document.querySelector('.home-record'))`);
    const known2 = new Set(BrowserWindow.getAllWindows().map((w) => w.id));
    await hjs(`document.querySelector('.home-record')?.click(), 1`);
    const fresh2 = () => BrowserWindow.getAllWindows().find((w) => !known2.has(w.id) && !w.isDestroyed()) || null;
    await ok(until(() => Boolean(fresh2()), 'the new Video window', 8000));
    const nv = fresh2();
    let opened = false;
    if (nv) {
      await new Promise((r) => (nv.webContents.isLoading() ? nv.webContents.once('did-finish-load', r) : r()));
      opened = await ok(until(() => nv.webContents.executeJavaScript(`Boolean(document.querySelector('.rec-dialog'))`), 'the recorder in the new window', 8000));
      await nv.webContents.executeJavaScript(`[...document.querySelectorAll('.rec-dialog .rw-btn')].find((b) => b.textContent.trim() === 'Cancel')?.click(), 1`).catch(() => {});
      await wait(300);
      nv.destroy();
    }
    check('recorder: the launcher offers Record a screen, which opens a new video on the recorder', card && opened, JSON.stringify({ card, opened }));
  } catch (err) {
    check('recorder: the checks ran', false, err.message);
  }

  // In Urdu: the box speaks Urdu right to left, and the picture stays left to right.
  if (win && !win.isDestroyed()) {
    const js = (code) => win.webContents.executeJavaScript(code);
    const was = await js(`window.rutbaOffice.store.get({ key: 'language', fallback: null })`).catch(() => null);
    try {
      await js(`window.rutbaOffice.store.set({ key: 'language', value: 'ur' }).then(() => 1)`);
      const ur = await open('video');
      const ujs = (code) => ur.webContents.executeJavaScript(code);
      await ujs(`document.querySelector('.vd-start-record')?.click(), 1`);
      await ok(until(() => ujs(`Boolean(document.querySelector('.rec-dialog .rec-source'))`), 'the recorder in Urdu', 10000));
      await wait(1200);
      const frame = await ujs(`({ dir: document.documentElement.dir, picture: getComputedStyle(document.querySelector('.rec-frame')).direction, english: [...document.querySelectorAll('.rec-dialog button, .rec-dialog h3, .rec-toggle-label, .rw-dialog-head')].map((n) => n.textContent.trim()).filter((s) => /^(Record|Camera|Microphone|Computer sound|Quality|Cancel|Screen|Window|Region)$/.test(s)) })`);
      await capture(ur, 'recorder-6-urdu.png');
      check('recorder: in Urdu the recorder speaks Urdu, right to left, and the picture stays left to right', frame.dir === 'rtl' && frame.picture === 'ltr' && frame.english.length === 0, JSON.stringify(frame));
      await ujs(`(() => { const b = document.querySelector('.rec-dialog .rw-dialog-foot .rw-btn:not(.primary)'); b?.click(); return 1; })()`);
    } catch (err) {
      check('recorder: the Urdu checks ran', false, err.message);
    } finally {
      await js(`window.rutbaOffice.store.set({ key: 'language', value: ${JSON.stringify(was ?? null)} }).then(() => 1)`).catch(() => {});
      await js(`window.rutbaOffice.store.set({ key: 'recorder.settings', value: ${JSON.stringify(settingsWere ?? null)} }).then(() => 1)`).catch(() => {});
    }
  }
}
