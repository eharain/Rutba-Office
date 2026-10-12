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

import fs from 'node:fs';
import path from 'node:path';
import { parseVtt } from '@rutba/studio/captions';
import { probeMedia } from '@rutba/media/probe';

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
    const ext = await js(`(() => { for (const t of ['video/mp4;codecs=avc1.42E01E', 'video/mp4;codecs=avc1', 'video/mp4']) if (MediaRecorder.isTypeSupported(t)) return 'mp4'; return 'webm'; })()`);
    const small = path.join(dir, `intro-captions-email.${ext}`);
    answer('save', small);
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
    win.destroy();
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
}
