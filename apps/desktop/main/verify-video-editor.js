// Video: the editor on the studio engine (docs/STUDIO.md Phase 1).
//
// Two clips, each a second and a half of colour with a tone under it, and a
// still, made here at run time (the clips encoded in the Video window with
// WebCodecs and written as WebM; nothing comes from the machine). They are
// added through Add Media and land end to end, each clip with its sound in
// one group; the first is split at the playhead (its sound split with it),
// the second trimmed at its end with ripple (the still moves up), a title
// is added from the gallery and a crossfade put into the second clip. The
// project is exported with its sound and the file read back: its length by
// @rutba/media's probe, that it carries a sound track, and a frame that is
// not blank. Autosave leaves a copy in the recovery folder while the work is
// unsaved and drops it on Save; the saved .rvid opens again in a window of
// its own with the same clips. Run alone with RUTBA_VERIFY_ONLY=video-editor.
//
// verifyVideoBasics and verifyVideoSplit are the two Video checks that were
// written for the single-file app (tone.wav opened from the launcher), kept
// here beside the editor's and called from their places in the full run.

import fs from 'node:fs';
import path from 'node:path';
import { BrowserWindow } from 'electron';
import { probeMedia } from '@rutba/media/probe';
import { gradientPng } from './sample-picture.js';

/* ── fixtures: clips with sound, encoded in a window ─────────────────── */

/** Runs in the window: frames as VP8, a tone as Opus, both as base64 chunks with their times (ms). */
async function encodeInWindow({ seconds, hue, freq }) {
  if (typeof VideoEncoder === 'undefined') return null;
  const W = 320;
  const H = 180;
  const fps = 25;
  const c = document.createElement('canvas');
  c.width = W;
  c.height = H;
  const g = c.getContext('2d');
  const b64 = (u8) => {
    let s = '';
    for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode(...u8.subarray(i, i + 0x8000));
    return btoa(s);
  };
  const video = [];
  const venc = new VideoEncoder({
    output: (chunk) => {
      const bytes = new Uint8Array(chunk.byteLength);
      chunk.copyTo(bytes);
      video.push({ key: chunk.type === 'key', ms: Math.round(chunk.timestamp / 1000), b64: b64(bytes) });
    },
    error: () => {},
  });
  venc.configure({ codec: 'vp8', width: W, height: H, bitrate: 500000, framerate: fps });
  const frames = Math.round(seconds * fps);
  for (let i = 0; i < frames; i++) {
    const p = i / frames;
    const grad = g.createLinearGradient(0, 0, W, H);
    grad.addColorStop(0, `hsl(${hue + p * 40}, 75%, 52%)`);
    grad.addColorStop(1, `hsl(${hue + 60 + p * 40}, 70%, 30%)`);
    g.fillStyle = grad;
    g.fillRect(0, 0, W, H);
    g.fillStyle = 'rgba(255,255,255,0.92)';
    g.beginPath();
    g.arc(30 + p * (W - 60), H / 2 + Math.sin(p * Math.PI * 4) * 40, 18, 0, Math.PI * 2);
    g.fill();
    g.fillStyle = 'rgba(0,0,0,0.55)';
    g.fillRect(8, H - 26, 90, 18);
    g.fillStyle = '#fff';
    g.font = 'bold 12px sans-serif';
    g.fillText(`${(i / fps).toFixed(2)} s`, 14, H - 13);
    const frame = new VideoFrame(c, { timestamp: Math.round((i * 1e6) / fps) });
    venc.encode(frame, { keyFrame: i % 25 === 0 });
    frame.close();
  }
  await venc.flush();
  venc.close();

  let audio = null;
  let head = null;
  if (typeof AudioEncoder !== 'undefined') {
    try {
      audio = [];
      const aenc = new AudioEncoder({
        output: (chunk, meta) => {
          const bytes = new Uint8Array(chunk.byteLength);
          chunk.copyTo(bytes);
          audio.push({ ms: Math.round(chunk.timestamp / 1000), b64: b64(bytes) });
          if (meta?.decoderConfig?.description && !head) head = b64(new Uint8Array(meta.decoderConfig.description));
        },
        error: () => {},
      });
      aenc.configure({ codec: 'opus', sampleRate: 48000, numberOfChannels: 1, bitrate: 64000 });
      const per = 960; // 20 ms
      const total = Math.round(seconds * 48000);
      for (let at = 0; at < total; at += per) {
        const data = new Float32Array(per);
        for (let k = 0; k < per; k++) data[k] = Math.sin((2 * Math.PI * freq * (at + k)) / 48000) * 0.35;
        const ad = new AudioData({ format: 'f32-planar', sampleRate: 48000, numberOfFrames: per, numberOfChannels: 1, timestamp: Math.round((at * 1e6) / 48000), data });
        aenc.encode(ad);
        ad.close();
      }
      await aenc.flush();
      aenc.close();
    } catch {
      audio = null;
    }
  }
  return { W, H, video, audio, head };
}

const ebmlSize = (n) => { const out = Buffer.alloc(8); out.writeBigUInt64BE(BigInt(n)); out[0] = 0x01; return out; };
const ebml = (id, ...payload) => { const body = Buffer.concat(payload); return Buffer.concat([Buffer.from(id), ebmlSize(body.length), body]); };
const ebmlUint = (id, n) => { const b = Buffer.alloc(8); b.writeBigUInt64BE(BigInt(n)); let i = 0; while (i < 7 && b[i] === 0) i++; return ebml(id, b.subarray(i)); };
const ebmlFloat = (id, d) => { const b = Buffer.alloc(8); b.writeDoubleBE(d); return ebml(id, b); };
const ebmlStr = (id, str) => ebml(id, Buffer.from(str, 'utf8'));

/** The Opus header a WebM track carries: mono, 48 kHz, the encoder's pre-skip. */
function opusHead() {
  const b = Buffer.alloc(19);
  b.write('OpusHead', 0, 'latin1');
  b[8] = 1;
  b[9] = 1;
  b.writeUInt16LE(312, 10);
  b.writeUInt32LE(48000, 12);
  return b;
}

/** A WebM of a VP8 track and (when there is one) an Opus track, blocks in time order in one cluster. */
function muxClip({ W, H, video, audio, head }) {
  const header = ebml([0x1a, 0x45, 0xdf, 0xa3],
    ebmlUint([0x42, 0x86], 1), ebmlUint([0x42, 0xf7], 1), ebmlUint([0x42, 0xf2], 4), ebmlUint([0x42, 0xf3], 8),
    ebmlStr([0x42, 0x82], 'webm'), ebmlUint([0x42, 0x87], 2), ebmlUint([0x42, 0x85], 2));
  const lastMs = Math.max(video[video.length - 1].ms + 40, audio?.length ? audio[audio.length - 1].ms + 20 : 0);
  const info = ebml([0x15, 0x49, 0xa9, 0x66], ebmlUint([0x2a, 0xd7, 0xb1], 1000000), ebmlFloat([0x44, 0x89], lastMs), ebmlStr([0x4d, 0x80], 'rutba-checks'), ebmlStr([0x57, 0x41], 'rutba-checks'));
  const tracks = [ebml([0xae], ebmlUint([0xd7], 1), ebmlUint([0x73, 0xc5], 1), ebmlUint([0x83], 1), ebmlStr([0x86], 'V_VP8'), ebml([0xe0], ebmlUint([0xb0], W), ebmlUint([0xba], H)))];
  if (audio?.length) {
    tracks.push(ebml([0xae], ebmlUint([0xd7], 2), ebmlUint([0x73, 0xc5], 2), ebmlUint([0x83], 2), ebmlStr([0x86], 'A_OPUS'),
      ebml([0x63, 0xa2], head ? Buffer.from(head, 'base64') : opusHead()),
      ebmlUint([0x56, 0xaa], 6500000), ebmlUint([0x56, 0xbb], 80000000),
      ebml([0xe1], ebmlFloat([0xb5], 48000), ebmlUint([0x9f], 1))));
  }
  const blocks = [
    ...video.map((f) => ({ track: 1, ms: f.ms, key: f.key, bytes: Buffer.from(f.b64, 'base64') })),
    ...(audio || []).map((f) => ({ track: 2, ms: f.ms, key: true, bytes: Buffer.from(f.b64, 'base64') })),
  ].sort((a, b) => a.ms - b.ms || a.track - b.track);
  const simple = blocks.map((b) => ebml([0xa3], Buffer.from([0x80 | b.track, (b.ms >> 8) & 0xff, b.ms & 0xff, b.key ? 0x80 : 0x00]), b.bytes));
  const cluster = ebml([0x1f, 0x43, 0xb6, 0x75], ebmlUint([0xe7], 0), ...simple);
  return Buffer.concat([header, ebml([0x18, 0x53, 0x80, 0x67], info, ebml([0x16, 0x54, 0xae, 0x6b], ...tracks), cluster)]);
}

/** Two clips with sound and a still, in `dir`: { a, b, still, withSound }. */
async function makeFixtures(webContents, dir) {
  const root = path.join(dir, 'video-editor');
  fs.mkdirSync(root, { recursive: true });
  const out = { still: path.join(root, 'still.png'), withSound: false };
  fs.writeFileSync(out.still, gradientPng(640, 360, [20, 110, 140], [240, 190, 90]));
  for (const [name, hue, freq] of [['a', 10, 440], ['b', 200, 660]]) {
    const enc = await webContents.executeJavaScript(`(${encodeInWindow.toString()})(${JSON.stringify({ seconds: 1.6, hue, freq })})`);
    if (!enc || !enc.video?.length) return null;
    out.withSound = Boolean(enc.audio?.length);
    out[name] = path.join(root, `clip-${name}.webm`);
    fs.writeFileSync(out[name], muxClip(enc));
  }
  return out;
}

/* ── the checks ──────────────────────────────────────────────────────── */

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn, doc
 * @param {{ dir: string }} args the run's scratch folder
 */
export async function verifyVideoEditor(h, { dir }) {
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
  try {
    const win = await open('video');
    const js = (code) => win.webContents.executeJavaScript(code);
    const fixtures = await makeFixtures(win.webContents, dir);
    if (!fixtures) {
      check('video editor: clips could be made for the checks (VP8 through WebCodecs)', false, 'no encoder in this build');
      return;
    }
    check('video editor: the fixture clips carry sound (Opus through WebCodecs)', fixtures.withSound, fixtures.withSound ? 'two WebM clips with a tone each' : 'no Opus encoder: the clips are silent');

    const clickTab = (name) => js(`(() => { const t = [...document.querySelectorAll('.rw-tab')].find((n) => n.textContent.trim() === ${JSON.stringify(name)}); if (!t) return 'no tab'; t.click(); return 'tab'; })()`);
    const press = async (label) => {
      const find = `[...document.querySelectorAll('.rw-ribbon .rw-btn, .rw-dialog .rw-btn')].find((n) => (n.dataset.tip || n.title || n.textContent || '').trim().startsWith(${JSON.stringify(label)}) && !n.disabled)`;
      await until(() => js(`Boolean(${find})`), `the ${label} button`, 4000).catch(() => {});
      return js(`(() => { const b = ${find}; if (!b) return 'no button ' + ${JSON.stringify(label)}; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked'; })()`);
    };
    const clips = (kind) => js(`[...document.querySelectorAll('.vd-clip${kind ? `[data-kind="${kind}"]` : ''}')].map((n) => ({ id: n.dataset.id, start: Number(n.dataset.start), end: Number(n.dataset.end), tr: n.dataset.transition })).sort((a, b) => a.start - b.start)`);
    const count = async (kind) => (await clips(kind)).length;
    const duration = () => js(`Number(document.querySelector('.vd-editor')?.dataset.duration || 0)`);
    const seekTo = async (t) => {
      await js(`(() => {
        const s = document.querySelector('.vd-scrub');
        Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(s, ${JSON.stringify(String(t))});
        s.dispatchEvent(new Event('input', { bubbles: true }));
        return s.value;
      })()`);
      await until(() => js(`Math.abs(Number(document.querySelector('.vd-scrub').value) - ${t}) < 0.02`), 'the playhead to move', 3000).catch(() => {});
      await wait(200);
    };
    const pick = (id) => js(`(() => {
      const n = document.querySelector('.vd-clip[data-id="${id}"]');
      if (!n) return false;
      const r = n.getBoundingClientRect();
      const o = { bubbles: true, cancelable: true, button: 0, clientX: r.left + r.width / 2, clientY: r.top + r.height / 2, pointerId: 1 };
      n.dispatchEvent(new PointerEvent('pointerdown', o));
      window.dispatchEvent(new PointerEvent('pointerup', o));
      return true;
    })()`);

    // Add Media: two clips and a still, from the dialog.
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'open', answer: [fixtures.a, fixtures.b, fixtures.still] });
    const added = await press('Add Media');
    const landed = await until(async () => (await count('video')) === 2 && (await count('image')) === 1 && (!fixtures.withSound || (await count('sound')) === 2), 'the clips on the timeline', 20000).then(() => true).catch(() => false);
    const v0 = await clips('video');
    const i0 = await clips('image');
    const s0 = await clips('sound');
    const endToEnd = v0.length === 2 && i0.length === 1 && v0[0].start === 0 && Math.abs(v0[1].start - v0[0].end) < 0.01 && Math.abs(i0[0].start - v0[1].end) < 0.01;
    check('video editor: Add Media puts two clips and a still on the timeline, end to end, each clip with its sound under it',
      landed && endToEnd && (!fixtures.withSound || (s0.length === 2 && Math.abs(s0[1].start - v0[1].start) < 0.01)),
      `${added}; pictures ${v0.map((c) => `${c.start}-${c.end}`).join(', ')}; still ${i0.map((c) => `${c.start}-${c.end}`).join('')}; sound ${s0.length}`);
    const thumbs = await until(() => js(`(() => { const c = document.querySelector('.vd-clip[data-kind="video"] canvas'); if (!c) return false; const g = c.getContext('2d'); const d = g.getImageData(Math.floor(c.width / 2), Math.floor(c.height / 2), 1, 1).data; return d[3] > 0; })()`), 'the thumbnails', 10000).then(() => true).catch(() => false);
    check('video editor: a clip on the timeline shows its own frames', thumbs, thumbs ? 'frames drawn in the first clip' : 'no frames in the clip');
    const painted = await until(() => js(`(() => { const c = document.querySelector('.vd-canvas'); if (!c || !c.width) return false; const d = c.getContext('2d').getImageData(c.width / 2, c.height / 2, 1, 1).data; return d[0] + d[1] + d[2] > 30; })()`), 'the stage', 8000).then(() => true).catch(() => false);
    check('video editor: the stage shows the frame at the playhead', painted, painted ? 'the first clip at 0:00' : 'the stage stayed black');
    await capture(win, 'video-editor-1-added.png');

    // Split the first clip at the playhead: the picture and its sound.
    await seekTo(0.8);
    await clickTab('Edit');
    const split = await press('Split');
    await until(async () => (await count('video')) === 3, 'the clip to split', 4000).catch(() => {});
    const v1 = await clips('video');
    const s1 = await clips('sound');
    check('video editor: Split cuts the clip at the playhead in two, its sound with it',
      v1.length === 3 && Math.abs(v1[0].end - 0.8) < 0.02 && Math.abs(v1[1].start - 0.8) < 0.02 && (!fixtures.withSound || (s1.length === 3 && Math.abs(s1[0].end - 0.8) < 0.02)),
      `${split}; pictures ${v1.map((c) => `${c.start}-${c.end}`).join(', ')}; sound ${s1.map((c) => `${c.start}-${c.end}`).join(', ')}`);

    // Trim the second clip's end at the playhead, with ripple: the still moves up.
    const second = v1[2];
    const before = await duration();
    const at = Math.round((second.start + 1.0) * 100) / 100;
    await seekTo(at);
    const trim = await press('Trim End');
    await until(async () => Math.abs((await clips('video'))[2].end - at) < 0.02, 'the trim', 4000).catch(() => {});
    const v2 = await clips('video');
    const i2 = await clips('image');
    const after = await duration();
    check('video editor: Trim End stops the clip at the playhead and the still behind it moves up (ripple)',
      Math.abs(v2[2].end - at) < 0.02 && Math.abs(i2[0].start - at) < 0.02 && Math.abs(before - after - (second.end - at)) < 0.03,
      `${trim}; clip ${v2[2].start}-${v2[2].end}, still from ${i2[0].start}; ${before} s to ${after} s`);

    // A title from the gallery, at the playhead, on the track above.
    await seekTo(0.3);
    await clickTab('Insert');
    const titled = await press('Title');
    const hasTitle = await until(async () => (await count('text')) === 1, 'the title', 4000).then(() => true).catch(() => false);
    const title = (await clips('text'))[0];
    const words = await js(`document.querySelector('.vd-text')?.value || ''`);
    check('video editor: Insert, Title puts a title at the playhead on a track above the clips, its words ready to change',
      hasTitle && Math.abs(title.start - 0.3) < 0.02 && words.length > 0,
      `${titled}; ${title ? `${title.start}-${title.end}` : 'none'}; "${words}"`);
    const onStage = await until(() => js(`Boolean(document.querySelector('.vd-outline'))`), 'the title outlined on the stage', 4000).then(() => true).catch(() => false);
    check('video editor: the picked title is outlined on the stage', onStage, onStage ? 'outlined' : 'no outline');

    // A crossfade into the second clip (the one trimmed): it moves up under the first.
    await pick(second.id);
    await wait(250);
    await clickTab('Transitions');
    const faded = await press('Crossfade');
    await until(async () => (await clips('video')).find((c) => c.id === second.id)?.tr === 'fade', 'the crossfade', 4000).catch(() => {});
    const v3 = await clips('video');
    const into = v3.find((c) => c.id === second.id);
    const prior = v3.find((c) => c.id !== second.id && c.end > into.start && c.start < into.start);
    check('video editor: Crossfade overlaps the clip with the one before it, and the timeline marks it',
      into?.tr === 'fade' && Boolean(prior) && (await js(`Boolean(document.querySelector('.vd-clip[data-id="${second.id}"] .vd-tr'))`)),
      `${faded}; ${into ? `${into.start}-${into.end}, ${into.tr}` : 'gone'}; ${prior ? `over ${prior.start}-${prior.end}` : 'no overlap'}`);
    // Undo takes the crossfade off and Redo puts it back, as one step each.
    await clickTab('Edit');
    const undone = await press('Undo');
    const off = await until(async () => !(await clips('video')).some((c) => c.tr === 'fade'), 'the crossfade to come off', 4000).then(() => true).catch(() => false);
    const redone = await press('Redo');
    const on = await until(async () => (await clips('video')).find((c) => c.id === second.id)?.tr === 'fade', 'the crossfade to come back', 4000).then(() => true).catch(() => false);
    const v4 = await clips('video');
    check('video editor: Undo takes the last edit back and Redo makes it again', off && on && Math.abs(v4.find((c) => c.id === second.id).start - into.start) < 0.01, `${undone}, ${redone}; ${off ? 'off' : 'still on'}, then ${on ? 'on again' : 'not back'}`);

    await seekTo(Math.round(((into.start + (prior?.end ?? into.start)) / 2) * 100) / 100);
    await capture(win, 'video-editor-2-edited.png');

    // Autosave: a copy in the recovery folder while the work is unsaved.
    const copyFile = `window-${await js(`document.querySelector('.vd-editor')?.dataset.copy || ''`)}.rvid`;
    const copied = await until(() => doc.recoverable().some((e) => e.file === copyFile && e.kind === 'video'), 'the recovery copy', 8000).then(() => true).catch(() => false);
    check('video editor: unsaved work is copied to the recovery folder for the launcher to offer', copied, copied ? copyFile : 'no copy');

    // The launcher's Recover opens a window on the copy: the same clips, and unsaved.
    const known = new Set(BrowserWindow.getAllWindows().map((w) => w.id));
    await js(`window.rutbaOffice.win.create({ app: 'video', query: { recover: ${JSON.stringify(copyFile)} } }).then(() => 1)`);
    const fresh = () => BrowserWindow.getAllWindows().find((w) => !known.has(w.id) && !w.isDestroyed()) || null;
    await until(() => Boolean(fresh()), 'the recovered window', 8000).catch(() => {});
    const rescued = fresh();
    if (rescued) await new Promise((r) => (rescued.webContents.isLoading() ? rescued.webContents.once('did-finish-load', r) : r()));
    let recovered = null;
    if (rescued) {
      const rjs = (code) => rescued.webContents.executeJavaScript(code);
      const want = await count('video');
      await until(() => rjs(`document.querySelectorAll('.vd-clip[data-kind="video"]').length === ${want}`), 'the recovered clips', 15000).catch(() => {});
      recovered = await rjs(`({ clips: document.querySelectorAll('.vd-clip[data-kind="video"]').length, titles: document.querySelectorAll('.vd-clip[data-kind="text"]').length, dirty: document.querySelector('.vd-editor')?.dataset.dirty, duration: Number(document.querySelector('.vd-editor')?.dataset.duration) })`);
      // Gone the way a crash goes: the copy stays for the window that made it to drop on its save.
      rescued.destroy();
    }
    check('video editor: Recover in the launcher opens the unsaved work again, marked unsaved', Boolean(recovered && recovered.clips === (await count('video')) && recovered.titles === 1 && recovered.dirty === '1' && Math.abs(recovered.duration - (await duration())) < 0.01), JSON.stringify(recovered));

    // Export, with sound, and read the file back.
    const ext = await js(`(() => { for (const t of ['video/mp4;codecs=avc1.42E01E,mp4a.40.2', 'video/mp4;codecs=avc1,mp4a.40.2', 'video/mp4']) if (MediaRecorder.isTypeSupported(t)) return 'mp4'; return 'webm'; })()`);
    const target = path.join(dir, `video-editor-out.${ext}`);
    const projectLength = await duration();
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: target });
    await clickTab('Home');
    const exp = await press('Export');
    await until(() => js(`Boolean(document.querySelector('.vd-export-go'))`), 'the Export box', 4000).catch(() => {});
    await js(`document.querySelector('.vd-export-go')?.click(), 1`);
    const busy = await until(() => js(`Boolean(document.querySelector('.vd-export-busy'))`), 'the progress', 6000).then(() => true).catch(() => false);
    if (busy) await capture(win, 'video-editor-3-exporting.png');
    const written = await until(() => { try { return fs.statSync(target).size > 4000; } catch { return false; } }, 'the exported file', Math.round(projectLength * 1000) + 30000).then(() => true).catch(() => false);
    await until(() => js(`!document.querySelector('.vd-export-busy')`), 'the progress to close', 8000).catch(() => {});
    let probe = null;
    if (written) {
      await wait(400);
      probe = probeMedia(new Uint8Array(fs.readFileSync(target)), { complete: true });
    }
    check('video editor: Export writes the project as a video file and says how far it has got', written && busy, `${exp}; ${written ? `${path.basename(target)}, ${fs.statSync(target).size} bytes` : 'no file'}`);
    check('video editor: the exported file runs as long as the project',
      Boolean(probe?.duration) && Math.abs(probe.duration - projectLength) < 0.75,
      probe ? `${probe.format}, ${probe.duration?.toFixed(2)} s for a ${projectLength.toFixed(2)} s project` : 'not read');
    check('video editor: the exported file carries the sound', Boolean(probe?.hasAudio && probe?.hasVideo), probe ? `tracks: ${(probe.tracks || []).map((t) => `${t.type} ${t.codec}`).join(', ')}` : 'not read');
    let frame = null;
    if (written) {
      frame = await js(`(async () => {
        const r = await window.rutbaOffice.fs.read({ path: ${JSON.stringify(target)} });
        const v = document.createElement('video');
        v.muted = true;
        v.src = URL.createObjectURL(new Blob([r.bytes], { type: ${JSON.stringify(ext === 'mp4' ? 'video/mp4' : 'video/webm')} }));
        await new Promise((res) => { v.onloadeddata = res; v.onerror = res; setTimeout(res, 6000); });
        v.currentTime = 0.5;
        await new Promise((res) => { v.onseeked = res; setTimeout(res, 3000); });
        const c = document.createElement('canvas'); c.width = 64; c.height = 36;
        const g = c.getContext('2d'); g.drawImage(v, 0, 0, 64, 36);
        const px = g.getImageData(0, 0, 64, 36).data;
        let sum = 0; let sq = 0; const n = px.length / 4;
        for (let i = 0; i < px.length; i += 4) { const l = (px[i] + px[i + 1] + px[i + 2]) / 3; sum += l; sq += l * l; }
        const mean = sum / n;
        return { mean: Math.round(mean), spread: Math.round(Math.sqrt(Math.max(0, sq / n - mean * mean))) };
      })()`).catch((err) => ({ error: err.message }));
    }
    check('video editor: a frame of the exported file has a picture in it, not a blank', Boolean(frame && frame.mean > 25 && frame.spread > 6), frame ? JSON.stringify(frame) : 'no frame');

    // Save as .rvid: the recovery copy goes, and the project opens again.
    const project = path.join(dir, 'video-editor.rvid');
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: project });
    const saved = await press('Save');
    const wrote = await until(() => fs.existsSync(project) && js(`document.querySelector('.vd-editor')?.dataset.dirty === '0'`), 'the project file', 6000).then(() => true).catch(() => false);
    const head = wrote ? fs.readFileSync(project).subarray(0, 80).toString('latin1') : '';
    check('video editor: Save writes a .rvid project (a zip that names itself)', wrote && head.startsWith('PK') && head.includes('application/vnd.rutba.video-project+zip'), `${saved}; ${wrote ? `${fs.statSync(project).size} bytes` : 'not written'}`);
    const dropped = await until(() => !doc.recoverable().some((e) => e.file === copyFile), 'the recovery copy to go', 5000).then(() => true).catch(() => false);
    check('video editor: once saved, the recovery copy is gone', dropped, dropped ? 'none left' : 'still there');
    const counts = { video: await count('video'), image: await count('image'), sound: await count('sound'), text: await count('text') };
    const length = await duration();
    const complaints = await errorsIn(win);
    check('video editor: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');

    const again = await open('video', project);
    const js2 = (code) => again.webContents.executeJavaScript(code);
    const back = await until(() => js2(`document.querySelectorAll('.vd-clip[data-kind="video"]').length === ${counts.video} && document.querySelectorAll('.vd-clip[data-kind="text"]').length === ${counts.text}`), 'the project to open', 15000).then(() => true).catch(() => false);
    const length2 = await js2(`Number(document.querySelector('.vd-editor')?.dataset.duration || 0)`);
    const tr = await js2(`[...document.querySelectorAll('.vd-clip')].filter((n) => n.dataset.transition === 'fade').length`);
    check('video editor: the saved project opens again with the same clips, title and crossfade', back && Math.abs(length2 - length) < 0.01 && tr === 1,
      `${back ? 'opened' : 'not opened'}; ${length2} s (was ${length}); ${tr} crossfade`);
    await until(() => js2(`(() => { const c = document.querySelector('.vd-clip[data-kind="video"] canvas'); if (!c) return false; const d = c.getContext('2d').getImageData(Math.floor(c.width / 2), Math.floor(c.height / 2), 1, 1).data; return d[3] > 0; })()`), 'the thumbnails', 8000).catch(() => {});
    await js2(`(() => { const n = document.querySelector('.vd-clip[data-kind="sound"]'); if (!n) return 0; const r = n.getBoundingClientRect(); const o = { bubbles: true, button: 0, clientX: r.left + 10, clientY: r.top + 10 }; n.dispatchEvent(new PointerEvent('pointerdown', o)); window.dispatchEvent(new PointerEvent('pointerup', o)); return 1; })()`);
    await wait(300);
    await capture(again, 'video-editor-4-reopened.png');
    await moreChecks({ win: again, check, until, wait, capture, project, open });
  } catch (err) {
    check('video editor: the checks ran', false, err.message);
  }
}

/**
 * On the reopened project: a title dragged to the empty track above, the
 * preview playing, a marker found again, a still turned and letterboxed,
 * a track hidden, the frame made vertical, and the project saved again.
 */
async function moreChecks({ win, check, until, wait, capture, project, open }) {
  const js = (code) => win.webContents.executeJavaScript(code);
  const clickTab = (name) => js(`(() => { const t = [...document.querySelectorAll('.rw-tab')].find((n) => n.textContent.trim() === ${JSON.stringify(name)}); if (!t) return 'no tab'; t.click(); return 'tab'; })()`);
  const press = async (label) => {
    const find = `[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || n.title || n.textContent || '').trim().startsWith(${JSON.stringify(label)}) && !n.disabled)`;
    await until(() => js(`Boolean(${find})`), `the ${label} button`, 4000).catch(() => {});
    return js(`(() => { const b = ${find}; if (!b) return 'no button ' + ${JSON.stringify(label)}; b.click(); return 'clicked'; })()`);
  };
  const seekTo = async (t) => {
    await js(`(() => { const s = document.querySelector('.vd-scrub'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(s, ${JSON.stringify(String(t))}); s.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await wait(250);
  };
  const pick = (selector) => js(`(() => {
    const n = document.querySelector(${JSON.stringify(selector)});
    if (!n) return false;
    const r = n.getBoundingClientRect();
    const o = { bubbles: true, cancelable: true, button: 0, clientX: r.left + r.width / 2, clientY: r.top + r.height / 2, pointerId: 1 };
    n.dispatchEvent(new PointerEvent('pointerdown', o));
    window.dispatchEvent(new PointerEvent('pointerup', o));
    return true;
  })()`);
  /** The stage's pixel at fractions of the frame, as [r, g, b]. */
  const pixel = (fx, fy) => js(`(() => { const c = document.querySelector('.vd-canvas'); const d = c.getContext('2d').getImageData(Math.floor(c.width * ${fx}), Math.floor(c.height * ${fy}), 1, 1).data; return [d[0], d[1], d[2]]; })()`);
  const dark = (p) => p[0] + p[1] + p[2] < 60;

  // Drag the title up onto the empty track, and a little later.
  const dragged = await js(`(async () => {
    const clip = document.querySelector('.vd-clip[data-kind="text"]');
    const row = document.querySelector('.vd-row[data-track="V3"] .vd-row-track');
    if (!clip || !row) return 'nothing to drag';
    const a = clip.getBoundingClientRect();
    const b = row.getBoundingClientRect();
    const x0 = a.left + a.width / 2; const y0 = a.top + a.height / 2;
    const x1 = x0 + 40; const y1 = b.top + b.height / 2;
    clip.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, cancelable: true, button: 0, clientX: x0, clientY: y0, pointerId: 1 }));
    for (let k = 1; k <= 6; k++) {
      window.dispatchEvent(new PointerEvent('pointermove', { bubbles: true, clientX: x0 + ((x1 - x0) * k) / 6, clientY: y0 + ((y1 - y0) * k) / 6, pointerId: 1 }));
      await new Promise((r) => setTimeout(r, 30));
    }
    window.dispatchEvent(new PointerEvent('pointerup', { bubbles: true, clientX: x1, clientY: y1, pointerId: 1 }));
    return 'dragged';
  })()`);
  const onV3 = await until(() => js(`Boolean(document.querySelector('.vd-row[data-track="V3"] .vd-clip[data-kind="text"]'))`), 'the title on V3', 4000).then(() => true).catch(() => false);
  const titleStart = await js(`Number(document.querySelector('.vd-clip[data-kind="text"]')?.dataset.start)`);
  check('video editor: a title dragged on the timeline moves later and onto the track it is dropped on', onV3 && titleStart > 0.35, `${dragged}; on V3: ${onV3}; starts at ${titleStart}`);

  // Play: the playhead moves on by itself, and Pause stops it.
  await seekTo(0);
  await js(`document.querySelector('.vd-play')?.click(), 1`);
  await wait(1300);
  await js(`document.querySelector('.vd-play')?.click(), 1`);
  await wait(200);
  const played = await js(`Number(document.querySelector('.vd-scrub').value)`);
  await wait(500);
  const held = await js(`Number(document.querySelector('.vd-scrub').value)`);
  check('video editor: Play runs the playhead on in real time and Pause holds it', played > 0.6 && played < 2.5 && Math.abs(held - played) < 0.01, `${played.toFixed(2)} s after 1.3 s, then ${held.toFixed(2)}`);

  // A marker at the playhead, found again from the start.
  await seekTo(1.0);
  await clickTab('Insert');
  await press('Add Marker');
  await until(() => js(`document.querySelectorAll('.vd-marker').length === 1`), 'the marker', 3000).catch(() => {});
  await seekTo(0);
  await press('Next Marker');
  await wait(250);
  const atMarker = await js(`Number(document.querySelector('.vd-scrub').value)`);
  check('video editor: Add Marker marks the playhead and Next Marker goes back to it', (await js(`document.querySelectorAll('.vd-marker').length`)) === 1 && Math.abs(atMarker - 1) < 0.02, `playhead at ${atMarker}`);

  // The still, turned a quarter and fitted inside: black bars either side.
  const stillAt = await js(`Number(document.querySelector('.vd-clip[data-kind="image"]').dataset.start)`);
  await seekTo(stillAt + 1);
  const before = await pixel(0.03, 0.5);
  await pick('.vd-clip[data-kind="image"]');
  await clickTab('Clip');
  await press('Rotate Right');
  await press('Fit');
  await until(() => js(`document.querySelector('.vd-clip[data-kind="image"]').dataset.fit === 'contain'`), 'the fit', 3000).catch(() => {});
  await wait(400);
  const after = await pixel(0.03, 0.5);
  const turn = await js(`document.querySelector('.vd-clip[data-kind="image"]').dataset.turn`);
  check('video editor: Rotate Right and Fit turn a still on its side and letterbox it in the frame', turn === '90' && !dark(before) && dark(after), `turn ${turn}; edge ${before.join(',')} then ${after.join(',')}`);
  await capture(win, 'video-editor-5-still-turned.png');

  // Hide V1: the clips under the title go dark.
  await seekTo(0.5);
  const shown = await pixel(0.05, 0.9);
  await js(`(() => { const b = document.querySelector('.vd-row[data-track="V1"] .vd-row-btn'); b?.click(); return 1; })()`);
  await wait(400);
  const hidden = await pixel(0.05, 0.9);
  await js(`(() => { const b = document.querySelector('.vd-row[data-track="V1"] .vd-row-btn'); b?.click(); return 1; })()`);
  check('video editor: hiding a track takes its clips off the picture', !dark(shown) && dark(hidden), `${shown.join(',')} then ${hidden.join(',')}`);

  // Vertical: the frame becomes 1080 by 1920.
  await clickTab('View');
  await press('Vertical');
  const tall = await until(() => js(`(() => { const c = document.querySelector('.vd-canvas'); return c.width === 1080 && c.height === 1920; })()`), 'the vertical frame', 3000).then(() => true).catch(() => false);
  await capture(win, 'video-editor-6-vertical.png');
  await press('Landscape');
  check('video editor: View, Vertical makes the frame 9:16 for Reels, Shorts and TikTok', tall, tall ? '1080 by 1920' : 'unchanged');

  // A clip's own menu on a right-click, with the clip's commands.
  const menu = await js(`(async () => {
    const n = document.querySelector('.vd-clip[data-kind="video"]');
    const r = n.getBoundingClientRect();
    n.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, button: 2, clientX: r.left + 8, clientY: r.top + 8 }));
    await new Promise((res) => setTimeout(res, 300));
    const items = [...document.querySelectorAll('.rw-menu button')].map((b) => b.textContent.trim());
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
    return items;
  })()`);
  const picked = await js(`document.querySelector('.vd-clip[data-kind="video"]').classList.contains('selected')`);
  check('video editor: right-clicking a clip picks it and offers its commands', picked && menu.some((m) => m.startsWith('Split')) && menu.some((m) => m.startsWith('Detach Audio')), menu.join(' | '));

  // Saved again: nothing left unsaved, no copy left for the launcher.
  globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: project });
  await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || '').startsWith('Save')); b?.click(); return 1; })()`);
  const clean = await until(() => js(`document.querySelector('.vd-editor')?.dataset.dirty === '0'`), 'the save', 5000).then(() => true).catch(() => false);
  check('video editor: the changed project saves again', clean, clean ? 'saved' : 'still unsaved');

  // In Urdu: the window's words in Urdu and its frame right to left, the timeline still running left to right.
  const was = await js(`window.rutbaOffice.store.get({ key: 'language', fallback: null })`);
  try {
    await js(`window.rutbaOffice.store.set({ key: 'language', value: 'ur' }).then(() => 1)`);
    const ur = await open('video', project);
    const ujs = (code) => ur.webContents.executeJavaScript(code);
    await until(() => ujs(`document.querySelectorAll('.vd-clip').length > 3`), 'the project in Urdu', 15000).catch(() => {});
    await until(() => ujs(`(() => { const c = document.querySelector('.vd-clip[data-kind="video"] canvas'); if (!c) return false; return c.getContext('2d').getImageData(Math.floor(c.width / 2), Math.floor(c.height / 2), 1, 1).data[3] > 0; })()`), 'the thumbnails', 8000).catch(() => {});
    const frame = await ujs(`({ dir: document.documentElement.dir, tabs: [...document.querySelectorAll('.rw-tab')].map((t) => t.textContent.trim()), timeline: getComputedStyle(document.querySelector('.vd-timeline')).direction, empty: [...document.querySelectorAll('.rw-ribbon .rw-btn')].some((b) => /^(Split|Add Media|Export)$/.test(b.textContent.trim())) })`);
    await capture(ur, 'video-editor-7-urdu.png');
    check('video editor: in Urdu the window speaks Urdu, right to left, and its timeline still runs left to right', frame.dir === 'rtl' && frame.timeline === 'ltr' && !frame.tabs.includes('Home') && !frame.empty, JSON.stringify(frame));
  } finally {
    await js(`window.rutbaOffice.store.set({ key: 'language', value: ${JSON.stringify(was ?? null)} }).then(() => 1)`);
  }
}

/** The media loads and the timeline is built: tone.wav opened from the launcher. */
export async function verifyVideoBasics(h, { wav }) {
  const { open, check, until } = h;
  try {
    const win = await open('video', wav);
    const js = (code) => win.webContents.executeJavaScript(code);
    // The sound decodes on its own schedule and the clip is drawn from what it reports, so both are waited for.
    await until(() => js(`Number(document.querySelector('.vd-editor')?.dataset.duration || 0) > 0.5 && document.querySelectorAll('.vd-clip').length > 0`), 'the media to report a duration', 20000).catch(() => {});
    const length = await js(`Number(document.querySelector('.vd-editor')?.dataset.duration || 0)`);
    check('video: the media loads and reports a duration', length > 0.9 && length < 1.1, `duration ${length}`);
    const clips = await js(`document.querySelectorAll('.vd-clip').length`);
    check('video: a timeline is built from it', clips === 1, `${clips} clips`);
    const clean = await js(`document.querySelector('.vd-editor')?.dataset.dirty`);
    check('video: a file opened in Video is not unsaved work until it is changed', clean === '0', clean === '0' ? 'nothing to save' : 'marked unsaved');
  } catch (err) {
    check('video: the checks ran', false, err.message);
  }
}

/** Split at the playhead from the Edit tab: tone.wav in two. */
export async function verifyVideoSplit(h, { wav }) {
  const { open, check, until, wait, errorsIn } = h;
  try {
    const win = await open('video', wav);
    const js = (code) => win.webContents.executeJavaScript(code);
    // The media decodes on its own schedule; with a corpus run on the same machine it once took past eight seconds.
    await until(() => js(`document.querySelectorAll('.vd-clip').length > 0`), 'the timeline', 20000);
    // Move the playhead to the middle with the scrubber, then split there.
    await js(`(() => {
      const s = document.querySelector('.vd-scrub');
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(s, String(Number(s.max || 1) / 2));
      s.dispatchEvent(new Event('input', { bubbles: true }));
      s.dispatchEvent(new Event('change', { bubbles: true }));
      return s.value;
    })()`);
    await wait(300);
    const scrubbed = await js(`document.querySelector('.vd-time')?.textContent ?? ''`);
    // Split lives on the Edit tab (and on Home).
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Edit')?.click(), 1`);
    await wait(150);
    const pressed = await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || n.textContent || '').trim().startsWith('Split') && !n.disabled); if (!b) return 'no button Split'; b.click(); return 'clicked'; })()`);
    await until(() => js(`document.querySelectorAll('.vd-clip').length === 2`), 'the clip to split in two', 4000).catch(() => {});
    const clips = await js(`document.querySelectorAll('.vd-clip').length`);
    check('video: Split at the playhead makes two clips', clips === 2, `${pressed}; playhead at ${scrubbed}; ${clips} clips`);
    const complaints = await errorsIn(win);
    check('video: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('video: the button checks ran', false, err.message);
  }
}
