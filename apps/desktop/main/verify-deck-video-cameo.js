// Presentations: Export to Video with a cameo the camera was recorded into
// and a narration (2.18.0, A4.7). The cameo's recording (red here) plays in
// the cameo's place, not its still (grey here); the narration's speaker
// (black here), hidden in the show, is not drawn; and in the editor the
// speaker keeps off the default cameo in the corner. Run alone with
// RUTBA_VERIFY_ONLY=deckvideocameo.

import fs from 'node:fs';
import path from 'node:path';
import { buildPptx } from '@rutba/presentation';
import { pngOf } from '../../../tests/fixtures/cube-glb.js';

const solid = (c) => pngOf(Array.from({ length: 8 }, () => Array.from({ length: 8 }, () => c)));

/** Two and a half seconds of silence, as a WAV. */
function silence(ms = 2500) {
  const rate = 8000;
  const n = Math.round((rate * ms) / 1000);
  const b = Buffer.alloc(44 + n * 2);
  b.write('RIFF', 0); b.writeUInt32LE(36 + n * 2, 4); b.write('WAVEfmt ', 8); b.writeUInt32LE(16, 16); b.writeUInt16LE(1, 20); b.writeUInt16LE(1, 22);
  b.writeUInt32LE(rate, 24); b.writeUInt32LE(rate * 2, 28); b.writeUInt16LE(2, 32); b.writeUInt16LE(16, 34); b.write('data', 36); b.writeUInt32LE(n * 2, 40);
  return b;
}

/**
 * @param {object} h the harness: open, check, until, wait, doc, sessionFor
 */
export async function verifyDeckVideoCameo({ open, check, until, wait, doc, sessionFor }, { dir }) {
  const file = path.join(dir, 'video-cameo.pptx');
  const out = path.join(dir, 'video-cameo-out.mp4');
  try {
    fs.writeFileSync(file, buildPptx({ title: 'Talk', slides: [{ layout: 'title', title: 'Hello' }] }));
    const win = await open('slides', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const id = sessionFor('deck').id;
    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 1`), 'the slide', 8000);

    // A recording of the camera stands in for one: three seconds of red.
    const clip = await js(`(async () => {
      const c = document.createElement('canvas'); c.width = 320; c.height = 180;
      const g = c.getContext('2d');
      const stream = c.captureStream(30);
      const rec = new MediaRecorder(stream, { mimeType: 'video/webm' });
      const parts = [];
      rec.ondataavailable = (e) => { if (e.data?.size) parts.push(e.data); };
      rec.start(100);
      const t0 = performance.now();
      while (performance.now() - t0 < 3000) { g.fillStyle = '#e00000'; g.fillRect(0, 0, 320, 180); stream.getVideoTracks()[0].requestFrame?.(); await new Promise((r) => setTimeout(r, 33)); }
      await new Promise((r) => { rec.onstop = r; rec.stop(); });
      stream.getTracks().forEach((t) => t.stop());
      const bytes = new Uint8Array(await new Blob(parts).arrayBuffer());
      let s = ''; for (let i = 0; i < bytes.length; i += 8192) s += String.fromCharCode(...bytes.subarray(i, i + 8192));
      return btoa(s);
    })()`);
    doc.apply({ id, ops: [
      { op: 'addCameo', slide: 0 },
      { op: 'addCameoRecording', slide: 0, data: clip, contentType: 'video/webm', poster: solid([128, 128, 128, 255]).toString('base64') },
      { op: 'addNarration', slide: 0, data: silence(), contentType: 'audio/wav', durationMs: 2500, poster: solid([0, 0, 0, 255]) },
      { op: 'setTransition', slide: 0, spec: { advanceAfter: 2.5 } },
    ] });
    const model = doc.model({ id, slide: 0 });
    const shapes = model.slide?.shapes || [];
    const cameo = shapes.find((s) => s.cameoRecording)?.geometry;
    // The recording as the window reaches it, for the report.
    const url = shapes.find((s) => s.cameoRecording)?.media?.url || null;
    const reach = url ? await js(`fetch(${JSON.stringify(url)}).then((r) => r.blob()).then((b) => b.size + ' ' + b.type).catch((e) => 'failed: ' + e.message)`) : 'no url';
    const speaker = shapes.find((s) => /^Audio Recording/.test(s.name || ''))?.geometry;
    const size = model.size;
    const apart = Boolean(cameo && speaker) && (speaker.x + speaker.w <= cameo.x || speaker.x >= cameo.x + cameo.w || speaker.y + speaker.h <= cameo.y || speaker.y >= cameo.y + cameo.h);
    check('presentations: a narration recorded on a slide with the default cameo puts its speaker beside the cameo, not on its corner', apart, JSON.stringify({ cameo, speaker }));

    const ext = await js(`(() => { for (const t of ['video/mp4;codecs=avc1', 'video/mp4']) if (MediaRecorder.isTypeSupported(t)) return 'mp4'; return 'webm'; })()`);
    const target = ext === 'mp4' ? out : out.replace(/\.mp4$/, '.webm');
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: target });
    await js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Record')?.click();
      await new Promise((r) => setTimeout(r, 200));
      [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Export to Video')?.click();
      return 1;
    })()`);
    await until(() => js(`Boolean(document.querySelector('.sl-video-go'))`), 'the Export to Video box', 4000).catch(() => {});
    await js(`document.querySelector('.sl-video-go')?.click(), 1`);
    const written = await until(() => { try { return fs.statSync(target).size > 2000; } catch { return false; } }, 'the video file', 30000).then(() => true).catch(() => false);
    await until(() => js(`!document.querySelector('.sl-video-busy')`), 'the box to close', 8000).catch(() => {});
    let seen = null;
    if (written && cameo && speaker) {
      const at = (g) => [(g.x + g.w / 2) / size.width, (g.y + g.h / 2) / size.height];
      seen = await js(`(async () => {
        const r = await window.rutbaOffice.fs.read({ path: ${JSON.stringify(target)} });
        const v = document.createElement('video');
        v.muted = true;
        v.src = URL.createObjectURL(new Blob([r.bytes], { type: ${JSON.stringify(ext === 'mp4' ? 'video/mp4' : 'video/webm')} }));
        await new Promise((res) => { v.onloadeddata = res; v.onerror = res; setTimeout(res, 6000); });
        v.currentTime = 1.2;
        await new Promise((res) => { v.onseeked = res; setTimeout(res, 3000); });
        const c = document.createElement('canvas'); c.width = v.videoWidth || 1280; c.height = v.videoHeight || 720;
        const g = c.getContext('2d'); g.drawImage(v, 0, 0, c.width, c.height);
        const px = ([fx, fy]) => [...g.getImageData(Math.round(fx * c.width), Math.round(fy * c.height), 1, 1).data].slice(0, 3);
        return { cameo: px(${JSON.stringify(at(cameo))}), speaker: px(${JSON.stringify(at(speaker))}) };
      })()`);
    }
    const red = seen && seen.cameo[0] > 170 && seen.cameo[1] < 90 && seen.cameo[2] < 90;
    const noSpeaker = seen && seen.speaker.every((v) => v > 170);
    check('presentations: an exported video plays the cameo\'s recording in its place and leaves out the narration\'s speaker', written && red && noSpeaker, JSON.stringify({ written, seen, ext, url, reach }));
    await wait(50);
  } catch (err) {
    check('presentations: the video with a cameo checks ran', false, err?.message || String(err));
  }
}
