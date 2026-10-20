// Rutba Recorder, the recorder as an app of its own (owner, 2026-10-19:
// "split the recording tool out to be able to record video").
//
// The launcher has its tile. The window opens on the Screen and camera
// scene with Chromium's stand-in camera and microphone (RUTBA_FAKE_MEDIA);
// a Documents window this check opens is picked from the window list (a
// check run lists only the suite's own windows). The program is drawn live
// on its canvas: the window with the camera over it. A title is put across
// the foot, and the Be right back card shows on its own. Recorded, the scene
// is changed while it records and the recording paused and resumed; Stop
// keeps one WebM with picture and sound, 1280 by 720 as asked, as long as it
// ran. Edit in Video opens it on Video's timeline. Run alone with
// RUTBA_VERIFY_ONLY=recorder-app.

import fs from 'node:fs';
import path from 'node:path';
import { BrowserWindow } from 'electron';
import { buildDocx } from '@rutba/ooxml/build';
import { probeMedia } from '@rutba/media/probe';

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args the run's scratch folder
 */
export async function verifyRecorderApp(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  const folder = path.join(dir, 'studio-recordings');
  const docFile = path.join(dir, 'studio-fixture.docx');
  let js = null;
  let kept = null;
  try {
    fs.mkdirSync(folder, { recursive: true });
    const home = await open('home');
    const tile = await home.webContents.executeJavaScript(`(() => { const c = document.querySelector('.home-card[data-app="recorder"]'); return c ? c.textContent : ''; })()`);
    check('recorder app: the launcher has a Recorder tile', /Recorder/.test(tile), tile.slice(0, 80));

    fs.writeFileSync(docFile, buildDocx({ styles: true, paragraphs: [{ text: 'Recorded by the Recorder check.' }, { text: 'A second line.' }] }));
    await open('word', docFile);
    const win = await open('recorder');
    js = (code) => win.webContents.executeJavaScript(code);
    kept = await js(`window.rutbaOffice.store.get({ key: 'recorder.studio', fallback: null })`);
    await js(`window.rutbaOffice.store.set({ key: 'recorder.studio', value: ${JSON.stringify({ scene: 'screen-camera', size: 720, fps: 30, countdown: 0, camera: true, mic: true, system: false, titleOn: false, title: '', subtitle: '', folder, minimise: false })} }).then(() => 1)`);
    await win.webContents.reload();
    await until(() => js(`Boolean(document.querySelector('.st-canvas'))`), 'the Recorder window', 10000);

    // The fixture window, picked from the Window list.
    await js(`(() => { document.querySelector('.st-kinds button[data-value="window"]')?.click(); return 1; })()`);
    const listed = await ok(until(() => js(`[...document.querySelectorAll('.st-sources .rec-source')].some((n) => (n.dataset.tip || '').includes('studio-fixture'))`), 'the fixture window in the list', 10000));
    await js(`(() => { const b = [...document.querySelectorAll('.st-sources .rec-source')].find((n) => (n.dataset.tip || '').includes('studio-fixture')); b?.click(); return 1; })()`);
    const live = await ok(until(() => js(`(() => { const c = document.querySelector('.st-canvas'); return c && c.dataset.layers === 'screen camera' && Number(c.dataset.frames) > 10; })()`), 'the window and the camera in the program', 25000));
    const look = await js(`(() => { const c = document.querySelector('.st-canvas'); const g = c.getContext('2d'); const mid = g.getImageData(c.width / 2, c.height / 2, 1, 1).data; return { screen: c.dataset.screen, wait: document.querySelector('.st-wait')?.textContent || '', on: document.querySelector('.st-sources .rec-source.on')?.dataset.tip || '', layers: c.dataset.layers, size: [c.width, c.height], scenes: document.querySelectorAll('.st-scene').length, mid: mid[0] + mid[1] + mid[2], frames: Number(c.dataset.frames) }; })()`);
    check('recorder app: the window picked and the camera are drawn live as one picture, 1280 by 720 as set, with five scenes to choose from',
      listed && live && look.size[0] === 1280 && look.size[1] === 720 && look.scenes === 5 && look.mid > 60,
      JSON.stringify({ listed, live, ...look }));

    // Region (owner, 2026-10-20: "in recorder the region is not present as in video is"):
    // a part of the window drawn on the picture is what is recorded, and Whole picture undoes it.
    await js(`document.querySelector('.st-kinds button[data-value="region"]')?.click(), 1`);
    await until(() => js(`[...document.querySelectorAll('.st-sources .rec-source')].some((n) => (n.dataset.tip || '').includes('studio-fixture'))`), 'the window among the region sources', 8000).catch(() => {});
    await js(`(() => { const b = [...document.querySelectorAll('.st-sources .rec-source')].find((n) => (n.dataset.tip || '').includes('studio-fixture')); b?.click(); return 1; })()`);
    await until(() => js(`document.querySelector('.st-frame')?.classList.contains('drawing') && (document.querySelector('.st-canvas')?.dataset.layers || '').includes('screen')`), 'region mode', 15000).catch(() => {});
    await wait(600);
    const drawn = await js(`(async () => {
      const f = document.querySelector('.st-frame');
      const r = f.getBoundingClientRect();
      const at = (x, y) => ({ bubbles: true, cancelable: true, button: 0, clientX: r.left + r.width * x, clientY: r.top + r.height * y, pointerId: 1 });
      f.dispatchEvent(new PointerEvent('pointerdown', at(0.3, 0.25)));
      for (let k = 1; k <= 6; k++) { window.dispatchEvent(new PointerEvent('pointermove', at(0.3 + 0.4 * k / 6, 0.25 + 0.5 * k / 6))); await new Promise((res) => setTimeout(res, 25)); }
      window.dispatchEvent(new PointerEvent('pointerup', at(0.7, 0.75)));
      return 'drawn';
    })()`);
    const cropped = await until(() => js(`(document.querySelector('.st-canvas')?.dataset.crop || '') !== '' && Boolean(document.querySelector('.st-region-bar'))`), 'the region applied', 5000).then(() => true).catch(() => false);
    const crop = await js(`document.querySelector('.st-canvas')?.dataset.crop || ''`);
    const [cx, cy, cw, ch] = crop.split(',').map(Number);
    await js(`document.querySelector('.st-whole')?.click(), 1`);
    const whole = await until(() => js(`(document.querySelector('.st-canvas')?.dataset.crop || '') === '' && !document.querySelector('.st-region-bar')`), 'the whole picture', 4000).then(() => true).catch(() => false);
    check('recorder app: Region records the part of the window dragged out on the picture, and Whole picture records all of it again',
      drawn === 'drawn' && cropped && cw > 0.1 && cw < 0.95 && ch > 0.1 && ch < 0.95 && cx > 0 && cy > 0 && whole,
      JSON.stringify({ drawn, cropped, crop, whole }));
    await js(`document.querySelector('.st-kinds button[data-value="window"]')?.click(), 1`);
    await wait(400);

    // A title across the foot, then the Be right back card on its own.
    await js(`(() => { document.querySelector('.st-title-btn')?.click(); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.st-title'))`), 'the title fields', 4000).catch(() => {});
    await js(`(() => { const n = document.querySelector('.st-title'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, 'Imogen Hartley'); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    const titled = await ok(until(() => js(`document.querySelector('.st-canvas')?.dataset.title === '1'`), 'the title drawn', 4000));
    if (process.env.RUTBA_VERIFY_CAPTURE) {
      await wait(600);
      fs.mkdirSync(process.env.RUTBA_VERIFY_CAPTURE, { recursive: true });
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'recorder-app.png'), (await win.webContents.capturePage()).toPNG());
    }
    await js(`document.querySelector('.st-scene[data-scene="brb"]')?.click(), 1`);
    const card = await ok(until(() => js(`(() => { const c = document.querySelector('.st-canvas'); return c.dataset.scene === 'brb' && c.dataset.layers === '' && c.dataset.title === '0'; })()`), 'the card', 4000));
    check('recorder app: a title is drawn across the foot of the picture, and Be right back shows its card alone', titled && card, JSON.stringify({ titled, card }));
    await js(`document.querySelector('.st-scene[data-scene="screen-camera"]')?.click(), 1`);
    await wait(400);

    // Record: one file growing, the scene changed while it records, a pause.
    await js(`document.querySelector('.st-record')?.click(), 1`);
    const recording = await ok(until(() => js(`(() => { const b = document.querySelector('.st-stop'); return Boolean(b) && !b.disabled; })()`), 'the recording to start', 8000));
    const files = () => { try { return fs.readdirSync(folder).filter((n) => /\.webm$/.test(n)).map((n) => ({ name: n, size: fs.statSync(path.join(folder, n)).size })); } catch { return []; } };
    const grew = await ok(until(() => files().some((f) => f.size > 5000), 'the recording on disk', 8000));
    await wait(1200);
    await js(`document.querySelector('.st-scene[data-scene="camera"]')?.click(), 1`);
    const switched = await ok(until(() => js(`document.querySelector('.st-canvas')?.dataset.layers === 'camera'`), 'the camera scene while recording', 4000));
    await wait(900);
    await js(`document.querySelector('.st-pause')?.click(), 1`);
    const paused = await ok(until(() => js(`/Paused/.test(document.querySelector('.st-badge')?.textContent || '')`), 'paused', 3000));
    await wait(1200);
    await js(`document.querySelector('.st-pause')?.click(), 1`);
    await wait(900);
    check('recorder app: Record writes one file as it goes, the scene changes while it records, and Pause holds it',
      recording && grew && switched && paused && files().length === 1, JSON.stringify({ recording, grew, switched, paused, files: files() }));

    // Stop: one WebM, picture and sound, as long as it ran without the pause.
    await js(`document.querySelector('.st-stop')?.click(), 1`);
    const saved = await ok(until(() => js(`Boolean(document.querySelector('.st-saved'))`), 'the saved bar', 10000));
    const made = files();
    const probe = made.length === 1 ? probeMedia(new Uint8Array(fs.readFileSync(path.join(folder, made[0].name))), { complete: true }) : null;
    check('recorder app: Stop keeps one recording, named Recording and the time, with picture and sound, 1280 by 720, as long as it recorded and not the pause',
      saved && made.length === 1 && /^Recording .*\.webm$/.test(made[0].name) && probe?.hasVideo && probe?.hasAudio && probe.width === 1280 && probe.height === 720 && probe.duration > 2.2 && probe.duration < 5,
      JSON.stringify({ saved, made, probe: probe ? { video: probe.hasVideo, audio: probe.hasAudio, w: probe.width, h: probe.height, duration: probe.duration } : null }));

    // Edit in Video: the recording on Video's timeline.
    const before = BrowserWindow.getAllWindows();
    await js(`document.querySelector('.st-edit')?.click(), 1`);
    const opened = () => BrowserWindow.getAllWindows().find((w) => !before.includes(w) && !w.isDestroyed()) || null;
    await until(opened, 'a Video window', 8000).catch(() => {});
    const video = opened();
    let edited = false;
    if (video) {
      edited = await ok(until(() => video.webContents.executeJavaScript(`Boolean(document.querySelector('.vd-editor')) && document.querySelectorAll('.vd-clip[data-kind="video"]').length === 1`), 'the recording on the timeline', 15000));
      video.destroy();
    }
    check('recorder app: Edit in Video opens the recording on Video\'s timeline', edited, video ? 'a Video window opened' : 'no window');

    const complaints = await errorsIn(win);
    check('recorder app: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('recorder app: the check ran to the end', false, err?.message || String(err));
  } finally {
    if (js) await js(`window.rutbaOffice.store.set({ key: 'recorder.studio', value: ${JSON.stringify(kept ?? null)} }).then(() => 1)`).catch(() => {});
  }
}
