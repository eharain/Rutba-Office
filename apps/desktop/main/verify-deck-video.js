// Presentations: Record → Export to Video, pressed on the ribbon — a deck
// of two slides with a second's timing each (the second fading in) made
// into a video file at HD: the box shows its progress, the file is written
// where it was asked for, and it is a real MP4 (or WebM) that plays for
// about the deck's length.

import fs from 'node:fs';
import path from 'node:path';
import { buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait, doc, sessionFor
 */
export async function verifyDeckVideo(h, { dir }) {
  const { open, check, until, wait, doc, sessionFor } = h;
  const file = path.join(dir, 'video.pptx');
  const out = path.join(dir, 'video-out.mp4');
  try {
    fs.writeFileSync(file, buildPptx({ title: 'Film', slides: [{ layout: 'title', title: 'Opening' }, { layout: 'obj', title: 'Closing', body: ['Thanks'] }] }));
    const win = await open('slides', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const id = sessionFor('deck').id;
    doc.apply({ id, ops: [{ op: 'setTransition', slide: 0, spec: { advanceAfter: 1 } }, { op: 'setTransition', slide: 1, spec: { type: 'fade', advanceAfter: 1 } }] });
    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 2`), 'the slides', 8000);

    // The format this machine records in decides the name asked for.
    const ext = await js(`(() => { for (const t of ['video/mp4;codecs=avc1', 'video/mp4']) if (MediaRecorder.isTypeSupported(t)) return 'mp4'; return 'webm'; })()`);
    const target = ext === 'mp4' ? out : out.replace(/\.mp4$/, '.webm');
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: target });
    const pressed = await js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Record')?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Export to Video');
      if (!b || b.disabled) return 'no button'; b.click(); return 'clicked';
    })()`);
    await until(() => js(`Boolean(document.querySelector('.sl-video-go'))`), 'the Export to Video box', 4000).catch(() => {});
    await js(`document.querySelector('.sl-video-go')?.click(), 1`);
    const progress = await until(() => js(`/slide \\d+ of 2/.test(document.querySelector('.sl-video-busy')?.textContent || '')`), 'the progress', 6000).then(() => true).catch(() => false);
    const written = await until(() => { try { return fs.statSync(target).size > 2000; } catch { return false; } }, 'the video file', 20000).then(() => true).catch(() => false);
    await until(() => js(`!document.querySelector('.sl-video-busy')`), 'the box to close', 6000).catch(() => {});
    let real = false;
    let duration = null;
    if (written) {
      const head = fs.readFileSync(target).subarray(0, 12);
      real = ext === 'mp4' ? head.toString('latin1', 4, 8) === 'ftyp' : head.readUInt32BE(0) === 0x1a45dfa3;
      duration = await js(`(async () => {
        const r = await window.rutbaOffice.fs.read({ path: ${JSON.stringify(target)} });
        const v = document.createElement('video');
        v.muted = true;
        v.src = URL.createObjectURL(new Blob([r.bytes], { type: ${JSON.stringify(ext === 'mp4' ? 'video/mp4' : 'video/webm')} }));
        await new Promise((res) => { v.onloadedmetadata = res; v.onerror = res; setTimeout(res, 6000); });
        if (!Number.isFinite(v.duration)) { v.currentTime = 1e6; await new Promise((res) => { v.ondurationchange = res; setTimeout(res, 3000); }); }
        const length = Number.isFinite(v.duration) ? Math.round(v.duration * 10) / 10 : null;
        // A frame half a second in: the title slide, white, not a black screen.
        v.currentTime = 0.5;
        await new Promise((res) => { v.onseeked = res; setTimeout(res, 3000); });
        const c = document.createElement('canvas'); c.width = 64; c.height = 36;
        const g = c.getContext('2d'); g.drawImage(v, 0, 0, 64, 36);
        const px = g.getImageData(0, 0, 64, 36).data; let sum = 0; for (let i = 0; i < px.length; i += 4) sum += (px[i] + px[i + 1] + px[i + 2]) / 3;
        return { length, light: Math.round(sum / (px.length / 4)) };
      })()`);
    }
    check('presentations: Record → Export to Video makes the deck into a real video file at HD, with its progress shown, about as long as the deck\'s timings',
      pressed === 'clicked' && progress && written && real && duration?.length != null && duration.length >= 1.4 && duration.length <= 4 && duration.light > 150,
      JSON.stringify({ pressed, progress, written, real, duration, ext }));
    await wait(50);
  } catch (err) {
    check('presentations: the Export to Video checks ran', false, err.message);
  }
}
