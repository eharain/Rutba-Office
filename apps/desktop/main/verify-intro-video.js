// Video: introductory videos (docs/STUDIO.md Phase 3).
//
// Everything here is made at run time or is the suite's own: the caption
// files are written into this run's folder, the camera and the microphone
// are Chromium's stand-ins (RUTBA_FAKE_MEDIA), the screen recorded is a
// Documents window this check opens (a check run lists only the suite's own
// windows), the pointer is handed its positions and clicks by the main
// process's feed rather than read from the person's mouse, and every file
// goes to this run's folder. Run alone with RUTBA_VERIFY_ONLY=intro-video.
//
// Captions (X8): a SubRip file with English, Urdu, Arabic and Hindi cues is
// imported from the ribbon; each cue lands as a clip on a track of its own
// above the pictures, the Urdu and Arabic ones right to left, and the stage
// draws the one under the playhead; a cue's words are edited in its
// properties, and the track is written back as WebVTT and read again.
//
// Export presets (X11): the Export box names each destination's frame, and
// the captions alone export for email at 854 by 480.
//
// The pointer (R7) and the automatic zoom (X5): see verifyIntroPointer.
//
// The introduction itself (X1 to X4, X6, and the rest together, the plan's
// "Done when"): see verifyIntroScenes.

import fs from 'node:fs';
import path from 'node:path';
import { parseVtt } from '@rutba/studio/captions';
import { probeMedia } from '@rutba/media/probe';
import { EXPORT_EXT } from './verify-video-editor.js';
import { buildDocx } from '@rutba/ooxml/build';
import { BrowserWindow } from 'electron';
import { gradientPng } from './sample-picture.js';

/** The run's capture: a window's picture, saved when RUTBA_VERIFY_CAPTURE names a folder. */
function capturer(wait) {
  const capDir = process.env.RUTBA_VERIFY_CAPTURE || null;
  return async (win, name) => {
    if (!capDir || !win || win.isDestroyed()) return;
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
}

const ok = (p) => p.then(() => true).catch(() => false);

/**
 * A block's window closed when another is open: the last window destroyed
 * quits the application (window-all-closed), and the next block, or the
 * rest of the full run, still needs it.
 */
function letGo(win) {
  if (!win || win.isDestroyed()) return;
  if (BrowserWindow.getAllWindows().some((w) => w !== win && !w.isDestroyed())) win.destroy();
}

/** Answers queued for the system's Open and Save dialogs, an earlier block's leftovers of that kind cleared first. */
function answer(kind, value) {
  const queue = globalThis.__rutbaCheckDialogAnswers;
  if (!queue) return;
  for (let i = queue.length - 1; i >= 0; i--) if (queue[i].kind === kind) queue.splice(i, 1);
  queue.push({ kind, answer: value });
}

/** In the window: press a ribbon button by the start of its tip, on a tab first when named. */
const pressJs = (tip, tab = null) => `(async () => {
  ${tab ? `[...document.querySelectorAll('.rw-tab')].find((n) => n.textContent.trim() === ${JSON.stringify(tab)})?.click(); await new Promise((r) => setTimeout(r, 250));` : ''}
  const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || n.title || '').startsWith(${JSON.stringify(tip)}));
  if (!b) return 'no button';
  if (b.disabled) return 'disabled';
  b.click();
  return 'clicked';
})()`;

/** In the window: the average colour of a small square of the stage, at fractions of the frame. */
const stagePixelJs = (fx, fy) => `(() => {
  const c = document.querySelector('.vd-canvas'); if (!c) return null;
  const g = c.getContext('2d');
  const d = g.getImageData(Math.floor(c.width * ${fx}) - 6, Math.floor(c.height * ${fy}) - 6, 12, 12).data;
  let r = 0, gg = 0, b = 0; for (let i = 0; i < d.length; i += 4) { r += d[i]; gg += d[i + 1]; b += d[i + 2]; }
  const n = d.length / 4; return [Math.round(r / n), Math.round(gg / n), Math.round(b / n)];
})()`;

/** In the window: move the playhead to `t` seconds by the transport's scrubber. */
const seekJs = (t) => `(() => { const s = document.querySelector('.vd-scrub'); if (!s) return 0; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(s, String(${t})); s.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`;

/** In the window: pick a clip on the timeline as a click would. */
const pickClipJs = (selector) => `(() => { const n = document.querySelector(${JSON.stringify(selector)}); if (!n) return 0; const r = n.getBoundingClientRect(); const o = { bubbles: true, cancelable: true, button: 0, clientX: r.left + Math.min(20, r.width / 2), clientY: r.top + r.height / 2, pointerId: 1 }; n.dispatchEvent(new PointerEvent('pointerdown', o)); window.dispatchEvent(new PointerEvent('pointerup', o)); return 1; })()`;

const SRT = [
  '1', '00:00:00,300 --> 00:00:02,000', 'Hello, and welcome.', '',
  '2', '00:00:02,000 --> 00:00:04,000', 'میں آپ کو دکھاتا ہوں', 'یہ کیسے کام کرتا ہے', '',
  '3', '00:00:04,000 --> 00:00:06,000', 'مرحبا بكم في العرض', '',
  '4', '00:00:06,000 --> 00:00:08,000', 'नमस्ते, यह परिचय है', '',
].join('\r\n');

/**
 * Captions (X8): imported as a track, drawn, edited and written back.
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args the run's scratch folder
 */
export async function verifyIntroCaptions(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const capture = capturer(wait);
  try {
    const srt = path.join(dir, 'intro-captions.srt');
    fs.writeFileSync(srt, '\ufeff' + SRT);
    const win = await open('video');
    const js = (code) => win.webContents.executeJavaScript(code);
    answer('open', [srt]);
    const pressed = await js(pressJs('Import Captions', 'Insert'));
    const landed = await ok(until(() => js(`document.querySelectorAll('.vd-clip[data-caption="1"]').length === 4`), 'the captions on the timeline', 6000));
    const clips = await js(`[...document.querySelectorAll('.vd-clip[data-caption="1"]')].map((n) => ({ lane: Number(n.dataset.lane), start: Number(n.dataset.start), end: Number(n.dataset.end), dir: n.dataset.dir, text: n.querySelector('.vd-clip-name')?.textContent }))`);
    const lanes = new Set(clips.map((c) => c.lane));
    check('intro-video: Import Captions reads a SubRip file and puts each cue on the timeline as a clip of its own, on one track, at its times',
      pressed === 'clicked' && landed && lanes.size === 1 && clips[0]?.start === 0.3 && clips[3]?.end === 8,
      JSON.stringify({ pressed, clips: clips.map((c) => `${c.lane}:${c.start}-${c.end}`) }));
    check('intro-video: the Urdu and Arabic captions run right to left, the English and Hindi ones left to right',
      clips.map((c) => c.dir).join(' ') === 'ltr rtl rtl ltr', clips.map((c) => `${c.dir} ${c.text}`).join(' | '));

    // Drawn on the stage: the caption's pill and its white words over the dark frame.
    const looks = [];
    for (const [i, at] of [[0, 1], [1, 3], [2, 5], [3, 7]]) {
      await js(seekJs(at));
      await wait(350);
      const lit = await js(`(() => {
        const c = document.querySelector('.vd-canvas'); if (!c) return 0;
        const g = c.getContext('2d');
        const d = g.getImageData(Math.floor(c.width * 0.3), Math.floor(c.height * 0.72), Math.floor(c.width * 0.4), Math.floor(c.height * 0.2)).data;
        let n = 0; for (let k = 0; k < d.length; k += 4) if (d[k] > 200 && d[k + 1] > 200 && d[k + 2] > 200) n++;
        return n;
      })()`);
      looks.push(lit);
      if (i === 1) await capture(win, 'intro-captions-urdu.png');
      if (i === 3) await capture(win, 'intro-captions-hindi.png');
    }
    check('intro-video: the stage draws the caption under the playhead in white on its pill, in each of the four languages', looks.every((n) => n > 150), JSON.stringify(looks));

    // A caption's words edited in its properties.
    await js(pickClipJs('.vd-clip[data-caption="1"]'));
    const props = await ok(until(() => js(`Boolean(document.querySelector('.vd-text'))`), 'the caption\'s words', 4000));
    await js(`(() => { const a = document.querySelector('.vd-text'); Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(a, 'Hello, and welcome to the tour.'); a.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    const edited = await ok(until(() => js(`[...document.querySelectorAll('.vd-clip[data-caption="1"] .vd-clip-name')].some((n) => n.textContent === 'Hello, and welcome to the tour.')`), 'the edited caption', 4000));
    const kind = await js(`document.querySelector('.vd-insp-title small')?.textContent || ''`);
    await capture(win, 'intro-captions-edit.png');
    check('intro-video: a caption is picked on the timeline and its words are edited in its properties', props && edited && /^Caption/.test(kind), JSON.stringify({ props, edited, kind }));

    // Written back as WebVTT and read again.
    const vtt = path.join(dir, 'intro-captions.vtt');
    answer('save', vtt);
    const exported = await js(pressJs('Export Captions', 'Insert'));
    const written = await ok(until(() => { try { return fs.statSync(vtt).size > 40; } catch { return false; } }, 'the .vtt', 5000));
    const back = written ? parseVtt(fs.readFileSync(vtt, 'utf8')) : [];
    check('intro-video: Export Captions writes the track as WebVTT, which reads back as the same four cues with the edit',
      exported === 'clicked' && back.length === 4 && back[0].text === 'Hello, and welcome to the tour.' && back[1].text.includes('\n') && back[3].start === 6,
      JSON.stringify(back.map((c) => `${c.start}-${c.end} ${c.text.replace('\n', ' / ')}`)));

    // Export presets (X11): the box names each destination's frame, and the email one writes a small file.
    await js(pressJs('Export', 'Home'));
    const boxed = await ok(until(() => js(`Boolean(document.querySelector('.vd-export-preset'))`), 'the Export box', 4000));
    const frames = {};
    for (const id of ['youtube', 'social-wide', 'social-square', 'vertical', 'email', 'project']) {
      await js(`(() => { const s = document.querySelector('.vd-export-preset'); Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(s, '${id}'); s.dispatchEvent(new Event('change', { bubbles: true })); return 1; })()`);
      await ok(until(() => js(`document.querySelector('.vd-export-preset')?.value === '${id}'`), id, 2000));
      frames[id] = await js(`(() => { const f = document.querySelector('.vd-export-frame'); return f ? \`\${f.dataset.width}x\${f.dataset.height}\` : ''; })()`);
    }
    check('intro-video: Export offers each destination with its frame: YouTube 1920x1080, LinkedIn and X 16:9 or square, Reels, Shorts and TikTok 1080x1920, email 854x480',
      boxed && frames.youtube === '1920x1080' && frames['social-wide'] === '1920x1080' && frames['social-square'] === '1080x1080' && frames.vertical === '1080x1920' && frames.email === '854x480' && frames.project === '1920x1080',
      JSON.stringify(frames));
    await js(`(() => { const s = document.querySelector('.vd-export-preset'); Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(s, 'email'); s.dispatchEvent(new Event('change', { bubbles: true })); return 1; })()`);
    await wait(200);
    await capture(win, 'intro-presets-box.png');
    const ext = await js(EXPORT_EXT);
    const small = path.join(dir, `intro-captions-email.${ext}`);
    answer('save', small);
    await until(() => js(`document.querySelector('.vd-export-go')?.disabled === false`), 'the export format found', 6000).catch(() => {});
    await js(`document.querySelector('.vd-export-go')?.click(), 1`);
    const made = await ok(until(() => { try { return fs.statSync(small).size > 2000; } catch { return false; } }, 'the email export', 30000));
    await ok(until(() => js(`!document.querySelector('.vd-export-busy')`), 'the progress to close', 8000));
    let probe = null;
    if (made) { await wait(400); probe = probeMedia(new Uint8Array(fs.readFileSync(small)), { complete: true }); }
    const remembered = await js(`window.rutbaOffice.store.get({ key: 'video.exportPreset', fallback: null })`);
    check('intro-video: exported for email the captions make a small video of 854 by 480, as long as the project, and the choice is remembered',
      Boolean(probe?.hasVideo && probe.width === 854 && probe.height === 480 && Math.abs(probe.duration - 8) < 0.75 && remembered === 'email'),
      probe ? `${path.basename(small)}: ${probe.width}x${probe.height}, ${probe.duration?.toFixed(2)} s, ${fs.statSync(small).size} bytes, remembered ${remembered}` : 'no file');
    await js(`window.rutbaOffice.store.set({ key: 'video.exportPreset', value: null }).then(() => 1)`);

    const complaints = await errorsIn(win);
    check('intro-video: the captions reported no error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    letGo(win);
  } catch (err) {
    check('intro-video: the caption checks ran', false, err.message);
  }
}

/**
 * Every introduction check, in order: what RUTBA_VERIFY_ONLY=intro-video and
 * the full run call.
 */
export async function verifyIntroVideo(h, args) {
  await verifyIntroCaptions(h, args);
  await verifyIntroPointer(h, args);
  await verifyIntroScenes(h, args);
}

/** In the window: how many pixels of the stage are the pointer's yellow. */
const yellowJs = `(() => {
  const c = document.querySelector('.vd-canvas'); if (!c) return -1;
  const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data;
  let n = 0; for (let k = 0; k < d.length; k += 16) { const r = d[k], g = d[k + 1], b = d[k + 2]; if (r > 170 && g > 140 && b < 150 && r - b > 70 && g - b > 40) n++; }
  return n;
})()`;

/** In the window: the stage as a small grey picture, to tell one frame from another. */
const thumbJs = `(() => {
  const c = document.querySelector('.vd-canvas'); if (!c) return null;
  const s = document.createElement('canvas'); s.width = 48; s.height = 27;
  const g = s.getContext('2d'); g.drawImage(c, 0, 0, 48, 27);
  const d = g.getImageData(0, 0, 48, 27).data; const out = [];
  for (let k = 0; k < d.length; k += 4) out.push(Math.round((d[k] + d[k + 1] + d[k + 2]) / 3));
  return out;
})()`;
const differ = (a, b) => (a && b && a.length === b.length ? a.reduce((s, v, i) => s + Math.abs(v - b[i]), 0) / a.length : 0);

/**
 * The pointer (R7) and the automatic zoom (X5): a suite window recorded with
 * the pointer handed its moves and three clicks; the pointer lands as a
 * layer over the screen, its highlight and a ring drawn on the stage and
 * switched off in its properties; Zoom on Clicks zooms the screen where the
 * clicks were, the zoom made closer and taken away again.
 */
export async function verifyIntroPointer(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const capture = capturer(wait);
  const folder = path.join(dir, 'intro-pointer');
  let win = null;
  let settingsWere = null;
  try {
    const docFile = path.join(dir, 'intro-pointer-fixture.docx');
    fs.writeFileSync(docFile, buildDocx({ styles: true, paragraphs: [{ text: 'A window for the pointer to work in.' }, { text: 'Its second line, and a third below.' }, { text: 'Click here, then here, then here.' }] }));
    await open('word', docFile);
    win = await open('video');
    const js = (code) => win.webContents.executeJavaScript(code);
    settingsWere = await js(`window.rutbaOffice.store.get({ key: 'recorder.settings', fallback: null })`);
    await js(`window.rutbaOffice.store.set({ key: 'recorder.settings', value: ${JSON.stringify({ source: 'window', fps: 30, cap: 720, quality: 'standard', countdown: 0, camera: false, mic: false, system: false, minimise: true, pointer: true, folder })} }).then(() => 1)`);
    await js(pressJs('Record Screen', 'Home'));
    await ok(until(() => js(`[...document.querySelectorAll('.rec-source')].some((n) => (n.dataset.tip || '').includes('intro-pointer-fixture'))`), 'the fixture window in the list', 10000));
    await js(`(() => { const b = [...document.querySelectorAll('.rec-source')].find((n) => (n.dataset.tip || '').includes('intro-pointer-fixture')); b?.click(); return Boolean(b); })()`);
    await ok(until(() => js(`(() => { const v = document.querySelector('.rec-screen'); return Boolean(v && v.videoWidth > 0 && document.querySelector('.rec-source.on')?.dataset.tip.includes('intro-pointer-fixture')); })()`), 'the window live', 12000));
    const toggle = await js(`(() => { const n = document.querySelector('.rec-pointer-toggle input'); return n ? n.checked : null; })()`);
    await js(`document.querySelector('.rec-go')?.click(), 1`);
    const recording = await ok(until(() => js(`(() => { const b = document.querySelector('.rec-stop'); return Boolean(b) && !b.disabled; })()`), 'the recording', 9000));
    const feed = await ok(until(() => Boolean(globalThis.__rutbaPointerFeed), 'the pointer feed', 4000));
    // The pointer: across the window, three clicks on one spot, a rest.
    const f = globalThis.__rutbaPointerFeed;
    for (let i = 0; i <= 20; i++) { f?.move(0.15 + i * 0.022, 0.3 + i * 0.008); await wait(50); }
    for (const [x, y] of [[0.6, 0.45], [0.62, 0.47], [0.58, 0.44]]) { f?.move(x, y); await wait(250); f?.click(x, y); await wait(450); }
    await wait(1500);
    await js(`document.querySelector('.rec-stop')?.click(), 1`);
    const landed = await ok(until(() => js(`Boolean(document.querySelector('.vd-clip[data-take="pointer"]') && document.querySelector('.vd-clip[data-take="screen"]')) && !document.querySelector('.rec-dialog')`), 'the takes on the timeline', 20000));
    const clips = await js(`[...document.querySelectorAll('.vd-clip')].map((n) => ({ take: n.dataset.take, lane: Number(n.dataset.lane), start: Number(n.dataset.start), end: Number(n.dataset.end), clicks: n.dataset.clicks }))`);
    const scr = clips.find((c) => c.take === 'screen');
    const ptr = clips.find((c) => c.take === 'pointer');
    check('intro-video: the recorder writes the pointer down beside the screen, and it lands as a layer of its own on the lane above, as long as the screen, with its three clicks',
      toggle === true && recording && feed && landed && ptr?.lane === scr?.lane + 1 && ptr.start === scr.start && ptr.end === scr.end && ptr.clicks === '3',
      JSON.stringify({ toggle, recording, feed, clips }));
    const released = !globalThis.__rutbaPointerFeed;
    check('intro-video: the pointer is let go when the recording stops', released, released ? 'released' : 'still followed');

    // The stage: the highlight round the pointer and a ring just after a click.
    const clickAt = await js(`(() => { const n = document.querySelector('.vd-clip[data-take="pointer"]'); return n ? Number(n.dataset.start) : 0; })()`);
    // The first click's time: the moves took about 1 s, then 0.25 s before the click.
    let lit = 0;
    let at = 0;
    for (const t of [1.35, 1.45, 1.6, 2.1, 2.3]) {
      await js(seekJs(clickAt + t));
      await wait(450);
      const n = await js(yellowJs);
      if (n > lit) { lit = n; at = t; }
    }
    await js(seekJs(clickAt + at));
    await wait(450);
    await capture(win, 'intro-pointer-highlight.png');
    await js(pickClipJs('.vd-clip[data-take="pointer"]'));
    const props = await ok(until(() => js(`Boolean(document.querySelector('.vd-pointer-props'))`), 'the pointer\'s properties', 4000));
    await js(`(() => { for (const c of ['.vd-pointer-highlight', '.vd-pointer-rings']) { const n = document.querySelector(c); if (n && n.checked) n.click(); } return 1; })()`);
    await ok(until(() => js(`!document.querySelector('.vd-pointer-highlight')?.checked && !document.querySelector('.vd-pointer-rings')?.checked`), 'switched off', 3000));
    await wait(450);
    const dark = await js(yellowJs);
    check('intro-video: the stage draws a highlight round the pointer and a ring where it clicked, and both are switched off after in the pointer\'s properties',
      props && lit > 60 && dark < lit / 4, JSON.stringify({ props, lit, at, dark }));
    await js(`(() => { for (const c of ['.vd-pointer-highlight', '.vd-pointer-rings']) { const n = document.querySelector(c); if (n && !n.checked) n.click(); } return 1; })()`);

    // Zoom on Clicks: one zoom where the three clicks are, the screen drawn closer there.
    await js(pickClipJs('.vd-clip[data-take="screen"]'));
    await ok(until(() => js(`Boolean(document.querySelector('.vd-zoom-on'))`), 'Zoom on Clicks', 4000));
    await js(seekJs(clickAt + 2.4));
    await wait(500);
    const before = await js(thumbJs);
    await js(`document.querySelector('.vd-zoom-on')?.click(), 1`);
    const zoomed = await ok(until(() => js(`document.querySelector('.vd-clip[data-take="screen"]')?.dataset.zooms === '1'`), 'one zoom', 4000));
    await wait(600);
    const during = await js(thumbJs);
    await capture(win, 'intro-pointer-zoom.png');
    const rows = await js(`[...document.querySelectorAll('.vd-zoom-row')].map((n) => n.dataset.level)`);
    check('intro-video: Zoom on Clicks zooms the screen in where the clicks were, as one zoom listed in its properties',
      zoomed && rows.length === 1 && differ(before, during) > 2.5, JSON.stringify({ zoomed, rows, change: Math.round(differ(before, during) * 10) / 10 }));
    // Closer, then taken away.
    await js(`(() => { const s = document.querySelector('.vd-zoom-row .rw-select'); Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(s, '2.6'); s.dispatchEvent(new Event('change', { bubbles: true })); return 1; })()`);
    const closer = await ok(until(() => js(`document.querySelector('.vd-zoom-row')?.dataset.level === '2.6'`), 'closer', 3000));
    await wait(500);
    const tighter = await js(thumbJs);
    await js(`document.querySelector('.vd-zoom-row .rw-btn')?.click(), 1`);
    const gone = await ok(until(() => js(`!document.querySelector('.vd-clip[data-take="screen"]')?.dataset.zooms && Boolean(document.querySelector('.vd-zoom-on'))`), 'the zoom gone', 3000));
    await wait(600);
    const after = await js(thumbJs);
    check('intro-video: a zoom is made closer in the properties and taken away again, the screen back as it was',
      closer && gone && differ(during, tighter) > 2 && differ(before, after) < 2, JSON.stringify({ closer, gone, closerChange: Math.round(differ(during, tighter) * 10) / 10, back: Math.round(differ(before, after) * 10) / 10 }));

    const complaints = await errorsIn(win);
    check('intro-video: the pointer and the zoom reported no error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('intro-video: the pointer checks ran', false, err.message);
  } finally {
    if (win && !win.isDestroyed()) {
      await win.webContents.executeJavaScript(`window.rutbaOffice.store.set({ key: 'recorder.settings', value: ${JSON.stringify(settingsWere ?? null)} }).then(() => 1)`).catch(() => {});
      letGo(win);
    }
  }
}

/** In the window: set a form field's value as typing would, and say so to React. */
const fillJs = (selector, value, event = 'input') => `(() => {
  const n = document.querySelector(${JSON.stringify(selector)}); if (!n) return 'no field';
  const proto = n.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : n.tagName === 'SELECT' ? HTMLSelectElement.prototype : HTMLInputElement.prototype;
  Object.getOwnPropertyDescriptor(proto, 'value').set.call(n, ${JSON.stringify(value)});
  n.dispatchEvent(new Event(${JSON.stringify(event)}, { bubbles: true }));
  return 'filled';
})()`;

/** In the window: the stage's pixels within a box of the frame that are near a colour. */
const nearColourJs = (rgb, box = [0, 0, 1, 1], tol = 40) => `(() => {
  const c = document.querySelector('.vd-canvas'); if (!c) return -1;
  const [x0, y0, x1, y1] = ${JSON.stringify(box)};
  const d = c.getContext('2d').getImageData(Math.floor(c.width * x0), Math.floor(c.height * y0), Math.max(1, Math.floor(c.width * (x1 - x0))), Math.max(1, Math.floor(c.height * (y1 - y0)))).data;
  const [r, g, b] = ${JSON.stringify(rgb)};
  let n = 0; for (let k = 0; k < d.length; k += 16) if (Math.abs(d[k] - r) + Math.abs(d[k + 1] - g) + Math.abs(d[k + 2] - b) < ${tol}) n++;
  return n;
})()`;

/**
 * The introduction from start to finish (STUDIO.md Phase 3, "Done when"):
 * from the launcher's card, a brand kit set (an Urdu name, colours, a
 * font and a logo), the end card taken out to leave three scenes, a script
 * written and shown on the teleprompter while the talking head records with
 * the stand-in camera, the screen scene recorded from a suite window with
 * the pointer's clicks (zoomed on), the talking head taken again, the scenes
 * reordered and the screen scene switched to side by side, a caption file
 * imported, and the whole exported for YouTube and for Reels, both files
 * read back with probeMedia for their size, length and tracks.
 */
export async function verifyIntroScenes(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const capture = capturer(wait);
  const folder = path.join(dir, 'intro-scenes');
  let win = null;
  let home = null;
  let kept = { settings: null, brand: null };
  try {
    const docFile = path.join(dir, 'intro-scenes-fixture.docx');
    fs.writeFileSync(docFile, buildDocx({ styles: true, paragraphs: [{ text: 'How the report is made, in three steps.' }, { text: 'First, open the data. Then press Make Report.' }, { text: 'Last, send it to the team.' }] }));
    await open('word', docFile);
    home = await open('home');
    const hjs = (code) => home.webContents.executeJavaScript(code);
    kept = {
      settings: await hjs(`window.rutbaOffice.store.get({ key: 'recorder.settings', fallback: null })`),
      brand: await hjs(`window.rutbaOffice.store.get({ key: 'video.brand', fallback: null })`),
    };
    await hjs(`window.rutbaOffice.store.set({ key: 'recorder.settings', value: ${JSON.stringify({ source: 'window', fps: 30, cap: 720, quality: 'standard', countdown: 0, camera: true, mirror: true, corner: 'br', insetSize: 'medium', shape: 'circle', mic: true, system: false, minimise: true, pointer: true, folder })} }).then(() => 1)`);
    await hjs(`window.rutbaOffice.store.set({ key: 'video.brand', value: null }).then(() => 1)`);

    // The launcher's card: a new Video window on an introduction of four scenes.
    const card = await hjs(`Boolean(document.querySelector('.home-intro'))`);
    const known = new Set(BrowserWindow.getAllWindows().map((w) => w.id));
    await hjs(`document.querySelector('.home-intro')?.click(), 1`);
    await ok(until(() => Boolean(BrowserWindow.getAllWindows().find((w) => !known.has(w.id) && !w.isDestroyed())), 'the new Video window', 8000));
    win = BrowserWindow.getAllWindows().find((w) => !known.has(w.id) && !w.isDestroyed()) || null;
    if (!win) throw new Error('no introduction window');
    await new Promise((r) => (win.webContents.isLoading() ? win.webContents.once('did-finish-load', r) : r()));
    const js = (code) => win.webContents.executeJavaScript(code);
    const scenes = () => js(`[...document.querySelectorAll('.vi-scene')].map((n) => ({ id: n.dataset.scene, layout: n.dataset.layout, takes: Number(n.dataset.takes), take: n.dataset.take }))`);
    await ok(until(() => js(`document.querySelectorAll('.vi-scene').length === 4`), 'the scenes', 10000));
    const first = await scenes();
    const tabOn = await js(`[...document.querySelectorAll('.rw-tab')].find((n) => n.classList.contains('on') || n.getAttribute('aria-selected') === 'true')?.textContent.trim() || ''`);
    await capture(win, 'intro-1-new.png');
    check('intro-video: the launcher\'s Introduction video opens Video on a template of scenes: a title card, a talking head, the screen with the camera in a corner and an end card',
      card && first.map((s) => s.layout).join(' ') === 'title talking screenCamera end', JSON.stringify({ card, layouts: first.map((s) => s.layout), tab: tabOn }));
    const [titleId, talkId, screenId, endId] = first.map((s) => s.id);
    const pick = async (id) => { await js(`document.querySelector('.vi-scene[data-scene="${id}"] .vi-scene-head')?.click(), 1`); await ok(until(() => js(`document.querySelector('.vi-scene[data-scene="${id}"]')?.classList.contains('open')`), `scene ${id} open`, 3000)); };

    // The brand kit: an Urdu name and an English role, colours, a font and a logo, kept for this person.
    const logo = path.join(dir, 'intro-logo.png');
    fs.writeFileSync(logo, gradientPng(240, 120, [255, 120, 40], [250, 200, 60]));
    await js(pressJs('Brand Kit', 'Introduction'));
    await ok(until(() => js(`Boolean(document.querySelector('.vi-brand-dialog'))`), 'the brand kit', 4000));
    await js(fillJs('.vi-brand-name', 'عائشہ خان'));
    await js(fillJs('.vi-brand-role', 'Product designer'));
    await js(fillJs('.vi-colour-background', '#0a2540'));
    await js(fillJs('.vi-colour-accent', '#00d4aa'));
    await js(fillJs('.vi-brand-font', 'Georgia', 'change'));
    answer('open', [logo]);
    await js(`document.querySelector('.vi-logo-pick')?.click(), 1`);
    await ok(until(() => js(`(document.querySelector('.vi-logo-name')?.textContent || '').includes('intro-logo')`), 'the logo chosen', 4000));
    await wait(300);
    await capture(win, 'intro-2-brand.png');
    await js(`document.querySelector('.vi-brand-save')?.click(), 1`);
    await ok(until(() => js(`!document.querySelector('.vi-brand-dialog')`), 'the kit applied', 4000));
    await js(seekJs(1));
    await wait(800);
    const ground = await js(nearColourJs([10, 37, 64], [0.02, 0.02, 0.2, 0.2], 24));
    const accentOnCard = await js(nearColourJs([0, 212, 170], [0.3, 0.55, 0.7, 0.75], 60));
    const logoPx = await js(nearColourJs([252, 160, 50], [0.35, 0.1, 0.65, 0.45], 70));
    const storedBrand = await js(`window.rutbaOffice.store.get({ key: 'video.brand', fallback: null })`);
    const titleWords = await js(`[...document.querySelectorAll('.vd-clip[data-kind="text"]')].map((n) => ({ text: n.querySelector('.vd-clip-name')?.textContent, dir: n.dataset.dir }))`);
    await capture(win, 'intro-3-title-card.png');
    check('intro-video: the brand kit is drawn into every scene: the title card in its background colour with its logo, the Urdu name right to left and the role in its accent, and it is kept for this person',
      ground > 200 && accentOnCard > 20 && logoPx > 40 && storedBrand?.name === 'عائشہ خان' && storedBrand?.font === 'Georgia' && titleWords.some((w) => w.text === 'عائشہ خان' && w.dir === 'rtl'),
      JSON.stringify({ ground, accentOnCard, logoPx, stored: storedBrand && { name: storedBrand.name, font: storedBrand.font, colors: storedBrand.colors, logo: Boolean(storedBrand.logo) }, titleWords: titleWords.slice(0, 3) }));

    // Three scenes: the end card taken out, the title card two seconds.
    await pick(endId);
    await js(`document.querySelector('.vi-scene.open .vi-remove')?.click(), 1`);
    await ok(until(() => js(`document.querySelectorAll('.vi-scene').length === 3`), 'three scenes', 3000));
    await pick(titleId);
    await js(fillJs('.vi-scene.open .vi-seconds', '2'));
    await ok(until(() => js(`Number(document.querySelector('.vd-clip[data-kind="shape"]')?.dataset.end) === 2`), 'the title card two seconds', 3000));

    // The talking head: a script, shown on the teleprompter above the camera while it records.
    const script = 'Hello, I am Ayesha. I design the reports your team reads every Monday, and in the next minute I will show you how one is made from start to finish.';
    await pick(talkId);
    await js(fillJs('.vi-scene.open .vi-script', script));
    // Typed, then left: the script is the scene's when the box loses the focus.
    await js(`(() => { const n = document.querySelector('.vi-scene.open .vi-script'); n?.dispatchEvent(new FocusEvent('focusout', { bubbles: true })); return 1; })()`);
    const scriptKept = await ok(until(() => js(`Boolean(document.querySelector('.vi-scene.open .vi-note'))`), 'the script kept', 3000));
    void scriptKept;
    await wait(200);
    let shotPrompter = false;
    const starts = [];
    const recordTake = async (sceneId, seconds, onRecording = null) => {
      await pick(sceneId);
      const had = Number(await js(`document.querySelector('.vi-scene[data-scene="${sceneId}"]')?.dataset.takes || 0`));
      await js(`document.querySelector('.vi-scene.open .vi-record')?.click(), 1`);
      await ok(until(() => js(`Boolean(document.querySelector('.rec-dialog'))`), 'the recorder', 6000));
      await ok(until(() => js(`[...document.querySelectorAll('.rec-source')].some((n) => (n.dataset.tip || '').includes('intro-scenes-fixture'))`), 'the fixture window', 10000));
      await js(`(() => { const b = [...document.querySelectorAll('.rec-source')].find((n) => (n.dataset.tip || '').includes('intro-scenes-fixture')); b?.click(); return 1; })()`);
      await ok(until(() => js(`(() => { const v = document.querySelector('.rec-screen'); const c = document.querySelector('.rec-inset .rec-cam'); return Boolean(v && v.videoWidth > 0 && c && c.videoWidth > 0 && document.querySelector('.rec-source.on')?.dataset.tip.includes('intro-scenes-fixture')); })()`), 'the window and the camera live', 15000));
      const before = await js(`({ prompter: document.querySelector('.rec-prompter-text')?.textContent || '', progress: Number(document.querySelector('.rec-prompter')?.dataset.progress || -1) })`);
      // Record once the box can: the first scene opens the camera and the microphone for the first time, which takes a while.
      for (let attempt = 0; attempt < 3; attempt++) {
        await ok(until(() => js(`(() => { const b = document.querySelector('.rec-go'); return Boolean(b) && !b.disabled; })()`), 'Record ready', 10000));
        await js(`document.querySelector('.rec-go')?.click(), 1`);
        if (await ok(until(() => js(`(() => { const b = document.querySelector('.rec-stop'); return Boolean(b) && !b.disabled; })()`), 'recording', 9000))) break;
        starts.push(await js(`[...document.querySelectorAll('[class*="toast"]')].map((n) => n.textContent.trim()).join(' | ') || 'no word'`));
        await wait(1000);
      }
      if (onRecording) await onRecording();
      else await wait(seconds * 1000);
      const during = await js(`Number(document.querySelector('.rec-prompter')?.dataset.progress || -1)`);
      if (sceneId === talkId && !shotPrompter) { shotPrompter = true; await capture(win, 'intro-4-teleprompter.png'); }
      await js(`document.querySelector('.rec-stop')?.click(), 1`);
      await ok(until(() => js(`!document.querySelector('.rec-dialog') && Number(document.querySelector('.vi-scene[data-scene="${sceneId}"]')?.dataset.takes) > ${had}`), 'the take in its scene', 20000));
      await wait(500);
      return { before, during };
    };
    const talk1 = await recordTake(talkId, 3.2);
    const afterTalk = await scenes();
    const talkClips = await js(`[...document.querySelectorAll('.vd-clip')].filter((n) => Number(n.dataset.start) >= 1.99).map((n) => ({ kind: n.dataset.kind, take: n.dataset.take, lane: Number(n.dataset.lane), inset: n.dataset.inset, text: n.querySelector('.vd-clip-name')?.textContent }))`);
    check('intro-video: the talking head records with its script on the teleprompter above the camera, scrolling as it records, and the take lands in its scene: the camera filling the frame, the voice under it',
      talk1.before.prompter === script && talk1.before.progress === 0 && talk1.during > 0.05 && afterTalk.find((s) => s.id === talkId)?.takes === 1
        && talkClips.some((c) => c.take === 'camera' && !c.inset) && talkClips.some((c) => c.take === 'voice'),
      JSON.stringify({ prompter: talk1.before.prompter.slice(0, 30), progress: [talk1.before.progress, talk1.during], takes: afterTalk.map((s) => s.takes), clips: talkClips.map((c) => `${c.take || c.kind}:${c.lane}`), retried: starts }));
    // The lower third: the Urdu name on the right, in the accent.
    const talkStart = await js(`Number(document.querySelector('.vd-clip[data-take="camera"]')?.dataset.start || 0)`);
    await js(seekJs(talkStart + 1.4));
    await wait(900);
    const thirdRight = await js(nearColourJs([0, 212, 170], [0.55, 0.7, 0.98, 0.95], 50));
    const thirdLeft = await js(nearColourJs([0, 212, 170], [0.02, 0.7, 0.45, 0.95], 50));
    await capture(win, 'intro-5-lower-third.png');
    check('intro-video: over the talking head the lower third comes in with the brand\'s name and role, on the right for an Urdu name, in the accent colour',
      thirdRight > 100 && thirdLeft < thirdRight / 4, JSON.stringify({ thirdRight, thirdLeft }));

    // The screen with the camera: recorded from the suite's window with the pointer handed three clicks.
    const screen1 = await recordTake(screenId, 0, async () => {
      await ok(until(() => Boolean(globalThis.__rutbaPointerFeed), 'the pointer feed', 4000));
      const f = globalThis.__rutbaPointerFeed;
      for (let i = 0; i <= 15; i++) { f?.move(0.2 + i * 0.025, 0.3 + i * 0.01); await wait(60); }
      for (const [x, y] of [[0.58, 0.42], [0.6, 0.45], [0.57, 0.44]]) { f?.move(x, y); await wait(250); f?.click(x, y); await wait(400); }
      await wait(1300);
    });
    void screen1;
    const screenClips = await js(`[...document.querySelectorAll('.vd-clip')].map((n) => ({ take: n.dataset.take, start: Number(n.dataset.start), inset: n.dataset.inset, zooms: n.dataset.zooms, clicks: n.dataset.clicks })).filter((c) => c.take)`);
    const scr = screenClips.filter((c) => c.take === 'screen').sort((a, b) => b.start - a.start)[0];
    const cam = screenClips.filter((c) => c.take === 'camera').sort((a, b) => b.start - a.start)[0];
    const ptr = screenClips.find((c) => c.take === 'pointer');
    check('intro-video: the screen scene records the window with the camera in a corner and the pointer\'s clicks, and zooms in on them by itself',
      Boolean(scr && cam && ptr?.clicks === '3' && Number(scr.zooms) >= 1 && cam.inset.endsWith('circle') && Number(cam.inset.split(' ')[0]) > 0.6),
      JSON.stringify({ scr, cam, ptr }));
    await js(seekJs(scr.start + 2.2));
    await wait(900);
    await capture(win, 'intro-6-screen-zoom.png');

    // A retake of the talking head: two takes, the second on the timeline, the first kept.
    const talk2 = await recordTake(talkId, 2.4);
    void talk2;
    const retaken = (await scenes()).find((s) => s.id === talkId);
    const useButtons = await js(`document.querySelectorAll('.vi-scene[data-scene="${talkId}"] .vi-use').length`);
    check('intro-video: a scene is recorded again: the new take goes on the timeline and the first is kept, to be put back with Use',
      retaken?.takes === 2 && retaken.take !== afterTalk.find((s) => s.id === talkId)?.take && useButtons === 1, JSON.stringify({ retaken, useButtons }));
    await capture(win, 'intro-7-retake.png');

    // Reordered: the screen scene before the talking head.
    await pick(screenId);
    await js(`document.querySelector('.vi-scene.open .vi-up')?.click(), 1`);
    await ok(until(() => js(`[...document.querySelectorAll('.vi-scene')].map((n) => n.dataset.scene).join(' ') === '${titleId} ${screenId} ${talkId}'`), 'the new order', 3000));
    const order = (await scenes()).map((s) => s.layout);
    const screenStart = await js(`Math.min(...[...document.querySelectorAll('.vd-clip[data-take="screen"]')].filter((n) => n.dataset.zooms).map((n) => Number(n.dataset.start)))`);
    check('intro-video: the scenes are reordered: the screen scene moved before the talking head, its clips moved with it', order.join(' ') === 'title screenCamera talking' && screenStart === 2, JSON.stringify({ order, screenStart }));

    // The screen scene's layout switched after recording: side by side, the same take.
    await js(`document.querySelector('.vi-scene.open .vi-layout[data-layout="sideBySide"]')?.click(), 1`);
    const switched = await ok(until(() => js(`document.querySelector('.vi-scene[data-scene="${screenId}"]')?.dataset.layout === 'sideBySide'`), 'side by side', 3000));
    const camNow = await js(`(() => { const all = [...document.querySelectorAll('.vd-clip[data-take="camera"]')].filter((n) => Number(n.dataset.start) === 2); return all[0]?.dataset.inset || ''; })()`);
    await js(seekJs(2 + 1));
    await wait(900);
    await capture(win, 'intro-8-side-by-side.png');
    const sb = camNow.split(' ');
    check('intro-video: the screen scene switched to side by side after recording keeps its take: the camera on the left in a rounded box, the screen beside it',
      switched && Number(sb[0]) < 0.1 && Number(sb[2]) > 0.3 && sb[4] === 'rounded', camNow);

    // Captions imported for the whole introduction.
    const srt = path.join(dir, 'intro-scenes.srt');
    fs.writeFileSync(srt, SRT);
    answer('open', [srt]);
    await js(pressJs('Import Captions', 'Insert'));
    const captioned = await ok(until(() => js(`document.querySelectorAll('.vd-clip[data-caption="1"]').length === 4`), 'the captions', 5000));
    const capLane = await js(`Number(document.querySelector('.vd-clip[data-caption="1"]')?.dataset.lane)`);
    check('intro-video: a caption file goes onto the introduction as a track of its own above every scene', captioned && capLane >= 8, JSON.stringify({ captioned, capLane }));

    // Exported for YouTube and for Reels, Shorts and TikTok, both read back.
    const length = await js(`Number(document.querySelector('.vd-editor')?.dataset.duration || 0)`);
    const ext = await js(EXPORT_EXT);
    const exportAs = async (preset, name) => {
      const target = path.join(dir, `${name}.${ext}`);
      answer('save', target);
      await js(pressJs('Export', 'Home'));
      await ok(until(() => js(`Boolean(document.querySelector('.vd-export-preset'))`), 'the Export box', 4000));
      await js(fillJs('.vd-export-preset', preset, 'change'));
      await ok(until(() => js(`document.querySelector('.vd-export-preset')?.value === '${preset}'`), preset, 2000));
      await until(() => js(`document.querySelector('.vd-export-go')?.disabled === false`), 'the export format found', 6000).catch(() => {});
      await js(`document.querySelector('.vd-export-go')?.click(), 1`);
      const made = await ok(until(() => { try { return fs.statSync(target).size > 4000; } catch { return false; } }, `the ${preset} export`, Math.round(length * 1000) + 40000));
      await ok(until(() => js(`!document.querySelector('.vd-export-busy')`), 'the progress to close', 10000));
      if (!made) return null;
      await wait(400);
      return { name: path.basename(target), bytes: fs.statSync(target).size, ...probeMedia(new Uint8Array(fs.readFileSync(target)), { complete: true }) };
    };
    const wide = await exportAs('youtube', 'intro-youtube');
    const tall = await exportAs('vertical', 'intro-vertical');
    const say = (p) => (p ? `${p.name}: ${p.width}x${p.height}, ${p.duration?.toFixed(2)} s of ${length.toFixed(2)}, ${Math.round(p.bytes / 1024)} KB, tracks ${(p.tracks || []).map((x) => `${x.type} ${x.codec}`).join(', ')}` : 'no file');
    check('intro-video: the three-scene introduction exports for YouTube as a 1920 by 1080 video as long as the project, with its picture and its sound',
      Boolean(wide?.hasVideo && wide.hasAudio && wide.width === 1920 && wide.height === 1080 && Math.abs(wide.duration - length) < 0.8), say(wide));
    check('intro-video: and for Reels, Shorts and TikTok as a 1080 by 1920 video of the same length, its scenes laid out upright',
      Boolean(tall?.hasVideo && tall.hasAudio && tall.width === 1080 && tall.height === 1920 && Math.abs(tall.duration - length) < 0.8), say(tall));
    const still = await js(`document.querySelector('.vd-editor')?.dataset.duration`);
    check('intro-video: exporting upright leaves the project itself as it was', Number(still) === length, `${still} s`);

    const complaints = await errorsIn(win);
    check('intro-video: the introduction reported no error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('intro-video: the introduction checks ran', false, err.message);
  } finally {
    const any = (win && !win.isDestroyed() && win) || (home && !home.isDestroyed() && home);
    if (any) {
      await any.webContents.executeJavaScript(`Promise.all([window.rutbaOffice.store.set({ key: 'recorder.settings', value: ${JSON.stringify(kept.settings ?? null)} }), window.rutbaOffice.store.set({ key: 'video.brand', value: ${JSON.stringify(kept.brand ?? null)} })]).then(() => 1)`).catch(() => {});
    }
    letGo(win);
  }
}
