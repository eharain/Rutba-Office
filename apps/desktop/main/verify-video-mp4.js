// Video: our own MP4 export, faster than real time (docs/STUDIO.md D3, P7,
// P8, Phase 4).
//
// A clip of 1920 by 1080 at 30 frames a second with a tone under it is made
// here at run time (encoded in the window with WebCodecs, H.264 and AAC, and
// written as an MP4 by @rutba/media's own muxer), beside the two small WebM
// clips with sound and the still the editor's checks use. The timeline is
// the big clip four times between them, sixty seconds and a fifth. Export
// for YouTube (1920 by 1080, 30 frames a second) offers our own MP4 and says
// it runs faster than the project plays; the export is timed from the press
// to the file in place and must take under sixty seconds; its progress says
// how fast it goes. The file is read back with @rutba/media's reader: ftyp,
// mdat and moov in that order, the picture H.264 at 1920 by 1080 for the
// project's length, a sample a frame, a key frame every two seconds, every
// sample inside the mdat, and the sound AAC for the same length. Pictures,
// the suite's own player, plays it: its length, its size, the clock moving
// and a frame with a picture in it. A second export cancelled part way
// leaves no file behind. Run alone with RUTBA_VERIFY_ONLY=video-mp4.

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Mp4Writer, memorySink } from '@rutba/media/mp4-write';
import { readMp4 } from '@rutba/media/mp4-read';
import { makeFixtures } from './verify-video-editor.js';

/** The share of the processors busy over `ms`, from the operating system's own counters. */
async function busyShare(ms) {
  const sample = () => os.cpus().reduce((t, c) => ({ idle: t.idle + c.times.idle, all: t.all + Object.values(c.times).reduce((a, b) => a + b, 0) }), { idle: 0, all: 0 });
  const a = sample();
  await new Promise((r) => setTimeout(r, ms));
  const b = sample();
  return b.all > a.all ? 1 - (b.idle - a.idle) / (b.all - a.all) : 0;
}

/**
 * Wait, up to two minutes, until the processors have been under a quarter
 * busy for three seconds running (as the gate settles before its budgets).
 * Answers how long it waited and how busy they were at the end.
 */
async function settled(limitMs = 120_000) {
  const started = Date.now();
  let calm = 0;
  let share = 1;
  while (Date.now() - started < limitMs && calm < 3) {
    share = await busyShare(1000);
    calm = share < 0.25 ? calm + 1 : 0;
  }
  return { waited: Math.round((Date.now() - started) / 1000), busy: Math.round(share * 100) };
}


/** Runs in the window: a 1080p clip as H.264 with a tone as AAC, as base64 samples with their times (microseconds). */
async function encodeHd({ seconds, fps, width, height, freq }) {
  if (typeof VideoEncoder === 'undefined') return null;
  const b64 = (u8) => {
    let s = '';
    for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode(...u8.subarray(i, i + 0x8000));
    return btoa(s);
  };
  const c = document.createElement('canvas');
  c.width = width;
  c.height = height;
  const g = c.getContext('2d');
  const video = [];
  let vdesc = null;
  let failed = null;
  const venc = new VideoEncoder({
    output: (chunk, meta) => {
      const bytes = new Uint8Array(chunk.byteLength);
      chunk.copyTo(bytes);
      if (meta?.decoderConfig?.description && !vdesc) vdesc = b64(new Uint8Array(ArrayBuffer.isView(meta.decoderConfig.description) ? meta.decoderConfig.description.buffer : meta.decoderConfig.description));
      video.push({ key: chunk.type === 'key', ts: chunk.timestamp, dur: chunk.duration || Math.round(1e6 / fps), b64: b64(bytes) });
    },
    error: (e) => { failed = e.message; },
  });
  venc.configure({ codec: 'avc1.640028', width, height, bitrate: 2500000, framerate: fps, avc: { format: 'avc' } });
  const frames = Math.round(seconds * fps);
  for (let i = 0; i < frames; i++) {
    const p = i / frames;
    const grad = g.createLinearGradient(0, 0, width, height);
    grad.addColorStop(0, `hsl(${200 + p * 120}, 70%, 45%)`);
    grad.addColorStop(1, `hsl(${260 + p * 120}, 65%, 22%)`);
    g.fillStyle = grad;
    g.fillRect(0, 0, width, height);
    g.fillStyle = 'rgba(255,255,255,0.9)';
    g.beginPath();
    g.arc(120 + p * (width - 240), height / 2 + Math.sin(p * Math.PI * 6) * 260, 90, 0, Math.PI * 2);
    g.fill();
    g.fillStyle = 'rgba(0,0,0,0.5)';
    g.fillRect(40, height - 130, 420, 90);
    g.fillStyle = '#fff';
    g.font = 'bold 64px sans-serif';
    g.fillText(`${(i / fps).toFixed(2)} s`, 70, height - 62);
    const frame = new VideoFrame(c, { timestamp: Math.round((i * 1e6) / fps), duration: Math.round(1e6 / fps) });
    venc.encode(frame, { keyFrame: i % (fps * 2) === 0 });
    frame.close();
    while (venc.encodeQueueSize > 6) await new Promise((r) => setTimeout(r, 2));
  }
  await venc.flush();
  venc.close();
  if (failed || !vdesc) return { failed: failed || 'no avcC' };

  const audio = [];
  let adesc = null;
  if (typeof AudioEncoder !== 'undefined') {
    try {
      const aenc = new AudioEncoder({
        output: (chunk, meta) => {
          const bytes = new Uint8Array(chunk.byteLength);
          chunk.copyTo(bytes);
          if (meta?.decoderConfig?.description && !adesc) adesc = b64(new Uint8Array(ArrayBuffer.isView(meta.decoderConfig.description) ? meta.decoderConfig.description.buffer : meta.decoderConfig.description));
          audio.push({ ts: chunk.timestamp, dur: chunk.duration || Math.round(1024e6 / 48000), b64: b64(bytes) });
        },
        error: () => {},
      });
      aenc.configure({ codec: 'mp4a.40.2', sampleRate: 48000, numberOfChannels: 1, bitrate: 96000 });
      const per = 4800;
      const total = Math.round(seconds * 48000);
      for (let at = 0; at < total; at += per) {
        const n = Math.min(per, total - at);
        const data = new Float32Array(n);
        for (let k = 0; k < n; k++) data[k] = Math.sin((2 * Math.PI * freq * (at + k)) / 48000) * 0.3;
        const ad = new AudioData({ format: 'f32-planar', sampleRate: 48000, numberOfFrames: n, numberOfChannels: 1, timestamp: Math.round((at * 1e6) / 48000), data });
        aenc.encode(ad);
        ad.close();
      }
      await aenc.flush();
      aenc.close();
    } catch {
      audio.length = 0;
    }
  }
  return { video, vdesc, audio, adesc };
}

/** The 1080p clip, muxed by our own writer: its path, or null. */
async function makeHdClip(webContents, root, { seconds = 13, fps = 30 } = {}) {
  const enc = await webContents.executeJavaScript(`(${encodeHd.toString()})(${JSON.stringify({ seconds, fps, width: 1920, height: 1080, freq: 523 })})`);
  if (!enc || enc.failed || !enc.video?.length) return { error: enc?.failed || 'no encoder' };
  const sink = memorySink();
  const w = new Mp4Writer({
    sink,
    video: { codec: 'avc1.640028', width: 1920, height: 1080, description: Buffer.from(enc.vdesc, 'base64') },
    audio: enc.audio.length ? { codec: 'mp4a.40.2', sampleRate: 48000, channels: 1, bitrate: 96000, description: enc.adesc ? Buffer.from(enc.adesc, 'base64') : null } : null,
  });
  let a = 0;
  for (const v of enc.video) {
    while (a < enc.audio.length && enc.audio[a].ts <= v.ts) {
      await w.add('audio', { data: Buffer.from(enc.audio[a].b64, 'base64'), timestamp: enc.audio[a].ts, duration: enc.audio[a].dur });
      a++;
    }
    await w.add('video', { data: Buffer.from(v.b64, 'base64'), timestamp: v.ts, duration: v.dur, key: v.key });
  }
  for (; a < enc.audio.length; a++) await w.add('audio', { data: Buffer.from(enc.audio[a].b64, 'base64'), timestamp: enc.audio[a].ts, duration: enc.audio[a].dur });
  await w.finish();
  const file = path.join(root, 'clip-1080p.mp4');
  fs.writeFileSync(file, sink.bytes());
  return { file, withSound: enc.audio.length > 0 };
}

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args the run's scratch folder
 */
export async function verifyVideoMp4(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
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
  const root = path.join(dir, 'video-mp4');
  fs.mkdirSync(root, { recursive: true });
  // A dialog answer an earlier block left unused is not this block's (the full run clears them between blocks; the fast path does not).
  if (Array.isArray(globalThis.__rutbaCheckDialogAnswers)) globalThis.__rutbaCheckDialogAnswers.length = 0;
  const win = await open('video');
  const js = (code) => win.webContents.executeJavaScript(code);

  const small = await makeFixtures(win.webContents, dir, 'video-mp4');
  const hd = await makeHdClip(win.webContents, root);
  if (!small || !hd.file) {
    check('video mp4: a 1080p H.264 clip and two small clips could be made for the checks', false, hd.error || 'no small clips');
    return;
  }
  check('video mp4: a 1080p H.264 clip with an AAC tone is made and muxed by our own writer', hd.withSound && fs.statSync(hd.file).size > 100000, `${path.basename(hd.file)}, ${fs.statSync(hd.file).size} bytes`);

  const clickTab = (name) => js(`(() => { const t = [...document.querySelectorAll('.rw-tab')].find((n) => n.textContent.trim() === ${JSON.stringify(name)}); if (!t) return 'no tab'; t.click(); return 'tab'; })()`);
  const press = async (label) => {
    const find = `[...document.querySelectorAll('.rw-ribbon .rw-btn, .rw-dialog .rw-btn')].find((n) => (n.dataset.tip || n.title || n.textContent || '').trim().startsWith(${JSON.stringify(label)}) && !n.disabled)`;
    await until(() => js(`Boolean(${find})`), `the ${label} button`, 6000).catch(() => {});
    return js(`(() => { const b = ${find}; if (!b) return 'no button ' + ${JSON.stringify(label)}; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked'; })()`);
  };
  const duration = () => js(`Number(document.querySelector('.vd-editor')?.dataset.duration || 0)`);

  // The timeline: the 1080p clip four times, the small clips and the still between.
  const order = [hd.file, small.a, hd.file, small.b, hd.file, small.still, hd.file];
  globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'open', answer: order });
  const added = await press('Add Media');
  await until(async () => (await js(`document.querySelectorAll('.vd-clip[data-kind="video"]').length`)) === 6 && (await duration()) > 59, 'the clips on the timeline', 30000).catch(() => {});
  const length = await duration();
  const counts = await js(`({ video: document.querySelectorAll('.vd-clip[data-kind="video"]').length, image: document.querySelectorAll('.vd-clip[data-kind="image"]').length, sound: document.querySelectorAll('.vd-clip[data-kind="sound"]').length })`);
  check('video mp4: a timeline of a minute: the 1080p clip four times, two small clips and a still, each clip with its sound', counts.video === 6 && counts.image === 1 && counts.sound >= 6 && length >= 59.5 && length < 61, `${added}; ${JSON.stringify(counts)}, ${length.toFixed(2)} s`);
  await capture(win, 'video-mp4-1-timeline.png');

  // Export for YouTube.
  const target = path.join(root, 'export-youtube.mp4');
  const openBox = async () => {
    await clickTab('Home');
    await press('Export');
    await until(() => js(`Boolean(document.querySelector('.vd-export-go'))`), 'the Export box', 6000).catch(() => {});
    await js(`(() => {
      const s = document.querySelector('.vd-export-preset');
      if (!s || s.value === 'youtube') return 1;
      Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(s, 'youtube');
      s.dispatchEvent(new Event('change', { bubbles: true }));
      return 1;
    })()`);
    await until(() => js(`document.querySelector('.vd-export-preset')?.value === 'youtube' && document.querySelector('.vd-export-format')?.dataset.fast === '1' && !document.querySelector('.vd-export-go')?.disabled`), 'the YouTube preset and our own MP4', 8000).catch(() => {});
  };
  await openBox();
  const box = await js(`({ format: document.querySelector('.vd-export-format')?.textContent.trim(), fast: document.querySelector('.vd-export-format')?.dataset.fast, frame: document.querySelector('.vd-export-frame')?.dataset.width + 'x' + document.querySelector('.vd-export-frame')?.dataset.height, hint: document.querySelector('.vd-export-dialog .rw-hint')?.textContent.trim() })`);
  await capture(win, 'video-mp4-2-export-box.png');
  check('video mp4: Export offers our own MP4 (H.264 and AAC) for YouTube at 1920 by 1080, and says it runs faster than the project plays',
    box.fast === '1' && /H\.264/.test(box.format || '') && /AAC/.test(box.format || '') && box.frame === '1920x1080' && /as fast as this computer/.test(box.hint || ''), JSON.stringify(box));

  globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: target });
  // The budget is the export's, so it starts on a machine that has settled:
  // the clip just encoded and written is read by the virus scanner for a
  // while after, and the lanes before this one end their windows; beside
  // that the clock measured the machine.
  const calm = await settled();
  const t0 = Date.now();
  await js(`document.querySelector('.vd-export-go')?.click(), 1`);
  const busy = await until(() => js(`Boolean(document.querySelector('.vd-export-busy'))`), 'the progress', 8000).then(() => true).catch(() => false);
  // The progress with its speed, once there is one to say.
  const speedShown = await until(() => js(`Number(document.querySelector('.vd-export-speed')?.dataset.speed || 0) > 0`), 'the speed', 30000).then(() => true).catch(() => false);
  let progressText = null;
  if (speedShown) {
    progressText = await js(`({ line: document.querySelector('.vd-export-line')?.textContent.trim(), speed: document.querySelector('.vd-export-speed')?.textContent.trim() })`);
    await capture(win, 'video-mp4-3-exporting.png');
  }
  const finished = await until(() => js(`Boolean(document.querySelector('.vd-editor')?.dataset.export)`), 'the export to finish', 180000).then(() => true).catch(() => false);
  const wall = (Date.now() - t0) / 1000;
  const measure = finished ? JSON.parse(await js(`document.querySelector('.vd-editor').dataset.export`)) : null;
  await until(() => js(`!document.querySelector('.vd-export-busy')`), 'the progress to close', 8000).catch(() => {});
  check('video mp4: the export says how far it has got and how much faster than the project plays', busy && speedShown && /Encoding/.test(progressText?.line || '') && /as fast as it plays/.test(progressText?.speed || ''), JSON.stringify(progressText));
  check('video mp4: a 60-second 1080p timeline exports in under 60 seconds (from the press to the file in place)',
    finished && wall < 60 && fs.existsSync(target),
    measure ? `${wall.toFixed(1)} s by the clock for ${length.toFixed(1)} s of 1080p (the encoder's own count ${measure.seconds} s, ${(length / measure.seconds).toFixed(2)}× real time; ${measure.codec}${measure.hardware ? ' on the graphics hardware' : ' in software'}, sound ${measure.audio}; ${(measure.size / 1e6).toFixed(1)} MB; ms spent ${JSON.stringify(measure.spent)}; begun after ${calm.waited} s for the processors to settle, at ${calm.busy}% busy)` : `not finished after ${wall.toFixed(1)} s`);

  // The file, read back by our own reader.
  let file = null;
  if (fs.existsSync(target)) {
    try {
      file = readMp4(new Uint8Array(fs.readFileSync(target)));
    } catch (err) {
      file = { error: err.message };
    }
  }
  const v = file?.tracks?.find((t) => t.kind === 'video');
  const a = file?.tracks?.find((t) => t.kind === 'audio');
  const frames = Math.round(length * 30);
  check('video mp4: the file is ftyp, mdat and moov, a plain MP4 that is not fragmented',
    Boolean(file && !file.error && file.boxes.map((b) => b.type).join(',') === 'ftyp,mdat,moov' && !file.fragmented), file ? (file.error || file.boxes.map((b) => `${b.type} ${b.size}`).join(', ')) : 'no file');
  check('video mp4: the picture is H.264 at 1920 by 1080 for the project\'s length, a sample a frame',
    Boolean(v && /^avc1\./.test(v.codec) && v.width === 1920 && v.height === 1080 && Math.abs(v.duration - length) < 0.05 && Math.abs(v.samples.length - frames) <= 1 && Math.abs(file.duration - length) < 0.1),
    v ? `${v.codec} ${v.width}x${v.height}, ${v.duration.toFixed(3)} s, ${v.samples.length} samples for ${frames} frames; movie ${file.duration.toFixed(3)} s` : 'no picture track');
  const keys = v ? v.samples.map((s, i) => (s.key ? i : -1)).filter((i) => i >= 0) : [];
  check('video mp4: a key frame every two seconds, the first sample one of them',
    Boolean(v && keys[0] === 0 && keys.length >= 30 && keys.every((k, i) => i === 0 || k - keys[i - 1] === 60)), `${keys.length} key frames at ${keys.slice(0, 5).join(', ')}…`);
  const mdat = file?.boxes?.find((b) => b.type === 'mdat');
  const inside = Boolean(mdat && v && a && [...v.samples, ...a.samples].every((s) => s.offset >= mdat.at + 16 && s.offset + s.size <= mdat.at + mdat.size));
  const total = v && a ? [...v.samples, ...a.samples].reduce((n, s) => n + s.size, 0) : 0;
  check('video mp4: every sample sits inside the mdat, which holds them and nothing else',
    inside && mdat && total === mdat.size - 16, mdat ? `${total} bytes of samples in an mdat of ${mdat.size - 16}; ${v?.tables.chunkOffsets.length} picture chunks, ${a?.tables.chunkOffsets.length} sound chunks` : 'no mdat');
  check('video mp4: the sound is AAC at 48 kHz for the same length',
    Boolean(a && /^mp4a\.40\./.test(a.codec) && a.sampleRate === 48000 && Math.abs(a.duration - length) < 0.1), a ? `${a.codec}, ${a.sampleRate} Hz, ${a.channels} channels, ${a.duration.toFixed(3)} s` : 'no sound track');

  // Frames read back where each kind of source lies: the 1080p MP4 clip, both small WebM
  // clips (decoded frame by frame from their VP8), the still, and the last clip's end.
  if (fs.existsSync(target)) {
    const spots = [[1.25, '1080p clip'], [13.5, 'WebM clip a'], [28.2, 'WebM clip b'], [44, 'still'], [48, 'last 1080p clip'], [52, 'last 1080p clip'], [56, 'last 1080p clip'], [59.8, 'last 1080p clip']];
    const seen = await js(`(async () => {
      const r = await window.rutbaOffice.fs.read({ path: ${JSON.stringify(target)} });
      const v = document.createElement('video');
      v.muted = true;
      v.src = URL.createObjectURL(new Blob([r.bytes], { type: 'video/mp4' }));
      await new Promise((res) => { v.onloadeddata = res; v.onerror = res; setTimeout(res, 6000); });
      const c = document.createElement('canvas'); c.width = 64; c.height = 36;
      const g = c.getContext('2d', { willReadFrequently: true });
      const out = [];
      for (const [t] of ${JSON.stringify(spots)}) {
        v.currentTime = t;
        await new Promise((res) => { v.onseeked = res; setTimeout(res, 5000); });
        // The frame for that time presented, where the element says so.
        if (v.requestVideoFrameCallback) await new Promise((res) => { v.requestVideoFrameCallback(() => res()); setTimeout(res, 800); });
        // A frame can reach the element a moment after 'seeked': read until there is one.
        let mean = 0; let spread = 0; let hue = null;
        for (let tries = 0; tries < 20; tries++) {
          g.clearRect(0, 0, 64, 36);
          g.drawImage(v, 0, 0, 64, 36);
          const px = g.getImageData(0, 0, 64, 36).data;
          let sum = 0; let sq = 0; const rgb = [0, 0, 0]; const n = px.length / 4;
          for (let i = 0; i < px.length; i += 4) { const l = (px[i] + px[i + 1] + px[i + 2]) / 3; sum += l; sq += l * l; rgb[0] += px[i]; rgb[1] += px[i + 1]; rgb[2] += px[i + 2]; }
          mean = sum / n;
          spread = Math.sqrt(Math.max(0, sq / n - mean * mean));
          hue = rgb.map((x) => Math.round(x / n));
          if (px[3] > 0 && mean > 0) break;
          await new Promise((res) => setTimeout(res, 150));
        }
        out.push({ t, at: Number(v.currentTime.toFixed(2)), mean: Math.round(mean), spread: Math.round(spread), rgb: hue });
      }
      return out;
    })()`).catch((err) => [{ error: err.message }]);
    check('video mp4: frames read back with a picture wherever a source lies: the 1080p clip, both WebM clips decoded frame by frame, the still and the end',
      Array.isArray(seen) && seen.length === spots.length && seen.every((f) => f.mean > 20 && f.spread > 6),
      (Array.isArray(seen) ? seen : []).map((f, i) => `${spots[i]?.[1]} at ${f.t} s (${f.at}): ${f.error || `mean ${f.mean}, spread ${f.spread}, rgb ${f.rgb}`}`).join('; '));
  }

  // Played in Pictures, the suite's own player.
  if (fs.existsSync(target)) {
    const player = await open('pictures', target);
    const pjs = (code) => player.webContents.executeJavaScript(code);
    const ready = await until(() => pjs(`(() => { const v = document.querySelector('.pv-video'); return Boolean(v && v.readyState >= 2 && v.duration > 0); })()`), 'the player to load the file', 20000).then(() => true).catch(() => false);
    const meta = ready ? await pjs(`(() => { const v = document.querySelector('.pv-video'); return { width: v.videoWidth, height: v.videoHeight, duration: v.duration, src: v.src.split(':')[0] }; })()`) : null;
    await pjs(`(() => { const v = document.querySelector('.pv-video'); if (!v) return 0; v.muted = true; v.play().catch(() => {}); return 1; })()`);
    const moved = await until(() => pjs(`(document.querySelector('.pv-video')?.currentTime || 0) > 1.2`), 'the clock to move', 15000).then(() => true).catch(() => false);
    // The player's picture is read from the window as it is painted: its
    // element draws a file from rutba://file, which a page's canvas may not read.
    let frame = null;
    if (moved) {
      try {
        const r = await pjs(`(() => { const v = document.querySelector('.pv-video'); const b = v.getBoundingClientRect(); return { x: Math.round(b.left + b.width * 0.1), y: Math.round(b.top + b.height * 0.1), width: Math.round(b.width * 0.8), height: Math.round(b.height * 0.8), at: Number(v.currentTime.toFixed(2)) }; })()`);
        player.webContents.invalidate();
        await wait(500);
        const img = await player.webContents.capturePage({ x: r.x, y: r.y, width: r.width, height: r.height });
        const px = img.toBitmap();
        let sum = 0; let sq = 0; let n = 0;
        for (let i = 0; i < px.length; i += 16) { const l = (px[i] + px[i + 1] + px[i + 2]) / 3; sum += l; sq += l * l; n++; }
        const mean = sum / Math.max(1, n);
        frame = { at: r.at, mean: Math.round(mean), spread: Math.round(Math.sqrt(Math.max(0, sq / Math.max(1, n) - mean * mean))) };
      } catch (err) {
        frame = { error: err.message };
      }
    }
    const named = await until(() => pjs(`/MP4/.test([...document.querySelectorAll('.pv-details dd')].map((d) => d.textContent).join(' '))`), 'the details to name the format', 8000).then(() => true).catch(() => false);
    const details = await pjs(`[...document.querySelectorAll('.pv-details dt')].map((d) => d.textContent + ': ' + (d.nextElementSibling?.textContent || '')).join('; ')`);
    await capture(player, 'video-mp4-4-playing.png');
    await pjs(`(() => { const v = document.querySelector('.pv-video'); if (v) v.pause(); return 1; })()`);
    check('video mp4: Pictures plays the exported file: its size and length, the clock moving, a frame with a picture in it, and names it an MP4',
      Boolean(meta && meta.width === 1920 && meta.height === 1080 && Math.abs(meta.duration - length) < 0.2 && moved && frame && frame.mean > 20 && frame.spread > 6 && named),
      `${JSON.stringify(meta)}; ${JSON.stringify(frame)}; ${details}`);
  }

  // Cancel part way: no file left behind.
  const cancelled = path.join(root, 'export-cancelled.mp4');
  await openBox();
  globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: cancelled });
  await js(`document.querySelector('.vd-export-go')?.click(), 1`);
  await until(() => js(`Number(document.querySelector('.vd-export-speed')?.dataset.speed || 0) > 0`), 'the second export under way', 30000).catch(() => {});
  const stopped = await js(`(() => { const b = [...document.querySelectorAll('.vd-export-busy .rw-btn')].find((n) => n.textContent.trim() === 'Cancel'); if (!b) return 'no Cancel'; b.click(); return 'cancelled'; })()`);
  const closed = await until(() => js(`!document.querySelector('.vd-export-busy')`), 'the progress to close', 10000).then(() => true).catch(() => false);
  await wait(800);
  check('video mp4: Cancel stops an export part way and leaves no file behind', closed && stopped === 'cancelled' && !fs.existsSync(cancelled), `${stopped}; ${fs.existsSync(cancelled) ? 'a file was left' : 'no file'}`);

  const complaints = await errorsIn(win);
  check('video mp4: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
}
