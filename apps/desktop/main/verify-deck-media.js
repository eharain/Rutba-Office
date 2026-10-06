// Presentations: Insert → Video → This Device and Audio → Audio on My PC,
// pressed on the ribbon — a real one-second video (made in the window from
// a painted canvas) goes in at its own shape with its poster frame and a
// play bar under it that plays it; a sound goes in as a speaker; in the
// show a click on the video plays it without moving the show on; and the
// saved deck keeps both as PowerPoint keeps media.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';

/** Half a second of a 440 Hz tone, as a WAV file. */
function toneWav() {
  const rate = 8000;
  const n = rate / 2;
  const buf = Buffer.alloc(44 + n * 2);
  buf.write('RIFF', 0); buf.writeUInt32LE(36 + n * 2, 4); buf.write('WAVE', 8);
  buf.write('fmt ', 12); buf.writeUInt32LE(16, 16); buf.writeUInt16LE(1, 20); buf.writeUInt16LE(1, 22);
  buf.writeUInt32LE(rate, 24); buf.writeUInt32LE(rate * 2, 28); buf.writeUInt16LE(2, 32); buf.writeUInt16LE(16, 34);
  buf.write('data', 36); buf.writeUInt32LE(n * 2, 40);
  for (let i = 0; i < n; i++) buf.writeInt16LE(Math.round(Math.sin((2 * Math.PI * 440 * i) / rate) * 8000), 44 + i * 2);
  return buf;
}

/**
 * @param {object} h the harness: open, check, until, wait, doc, sessionFor, capture
 */
export async function verifyDeckMedia(h, { dir }) {
  const { open, check, until, wait, doc, sessionFor, capture } = h;
  const file = path.join(dir, 'media.pptx');
  const clip = path.join(dir, 'clip.webm');
  const tone = path.join(dir, 'tone.wav');
  try {
    fs.writeFileSync(file, buildPptx({ title: 'Media', slides: [{ layout: 'blank' }, { layout: 'title', title: 'After' }] }));
    fs.writeFileSync(tone, toneWav());
    const win = await open('slides', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const model = () => doc.model({ id: sessionFor('deck').id });
    const media = () => (model().slide?.shapes || []).filter((s) => s.media);
    const press = (label) => js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(label)});
      if (!b || b.disabled) return 'no button';
      b.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: 700, clientY: 110 }));
      return 'clicked';
    })()`);
    const pick = (label) => js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); if (!b) return 'no item'; b.click(); return 'picked'; })()`);

    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 2`), 'the slides', 8000);
    // A second of video, painted in the window and recorded as WebM.
    const made = await js(`(async () => {
      const c = document.createElement('canvas'); c.width = 320; c.height = 180;
      const g = c.getContext('2d');
      const stream = c.captureStream(0); const track = stream.getVideoTracks()[0];
      const rec = new MediaRecorder(stream, { mimeType: 'video/webm' });
      const parts = []; rec.ondataavailable = (e) => parts.push(e.data);
      const done = new Promise((r) => { rec.onstop = r; });
      // Painting starts once the recorder is recording, and goes on frame
      // by frame — thirty of them, however slowly a window off the screen
      // runs its timers — so the clip always has pictures in it.
      // (A recorder says it has started only once a picture reaches it.)
      const paint = (i) => {
        g.fillStyle = '#2b5fd9'; g.fillRect(0, 0, 320, 180);
        g.fillStyle = '#ffc000'; g.fillRect(20 + i * 7, 60, 60, 60);
        track.requestFrame?.();
      };
      const started = new Promise((r) => { rec.onstart = r; setTimeout(r, 600); });
      rec.start(100);
      paint(0);
      await started;
      for (let i = 1; i <= 30; i++) {
        paint(i);
        await new Promise((r) => setTimeout(r, 40));
      }
      rec.stop(); await done;
      const bytes = new Uint8Array(await new Blob(parts, { type: 'video/webm' }).arrayBuffer());
      let s = ''; for (let i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
      return btoa(s);
    })()`);
    fs.writeFileSync(clip, Buffer.from(made, 'base64'));

    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'open', answer: [clip] });
    const pressed = await press('Video');
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === 'This Device…'))`), 'the menu', 3000).catch(() => {});
    await pick('This Device…');
    const placed = await until(() => media().length === 1, 'the video on the slide', 10000).then(() => true).catch(() => false);
    const video = media()[0];
    const toasted = await js(`document.body.innerText.includes('cannot show its picture')`);
    check('presentations: Insert → Video → This Device puts the video on the slide at its own shape, with a poster frame taken from it',
      pressed === 'clicked' && placed && video?.media?.kind === 'video' && Boolean(video.media.url) && Math.abs(video.geometry.w / video.geometry.h - 16 / 9) < 0.05 && !toasted,
      JSON.stringify({ pressed, placed, video: video && { media: video.media, g: video.geometry }, toasted }));

    // The play bar plays it on the stage.
    const bar = await until(() => js(`Boolean(document.querySelector('.sl-media-bar .sl-media-play'))`), 'the play bar', 4000).then(() => true).catch(() => false);
    await js(`document.querySelector('.sl-media-bar .sl-media-play')?.click(), 1`);
    const played = await until(() => js(`(() => { const v = document.querySelector('video.sl-media-video'); return Boolean(v) && v.currentTime > 0.2; })()`), 'the video playing', 6000).then(() => true).catch(() => false);
    await js(`(() => { const v = document.querySelector('video.sl-media-video'); v?.pause(); return 1; })()`);
    if (capture) await capture(win, 'deck-media-stage.png');
    check('presentations: a play bar under the video plays it on the slide, as PowerPoint\'s does', bar && played, JSON.stringify({ bar, played }));

    // A sound, as a speaker.
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'open', answer: [tone] });
    await press('Audio');
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === 'Audio on My PC…'))`), 'the menu', 3000).catch(() => {});
    await pick('Audio on My PC…');
    const sound = await until(() => media().some((s) => s.media.kind === 'audio'), 'the sound on the slide', 8000).then(() => true).catch(() => false);
    const speaker = media().find((s) => s.media.kind === 'audio');
    check('presentations: Insert → Audio → Audio on My PC puts the sound on the slide as a speaker with its own play bar',
      sound && speaker?.geometry?.w === 48 && (await js(`document.querySelectorAll('.sl-media-bar').length`)) === 2, JSON.stringify({ sound, speaker: speaker && speaker.geometry }));

    // In the show, a click on the video plays it and the show stays on the slide.
    await js(`(() => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Slide Show')?.click(); return 1; })()`);
    await wait(200);
    await js(`(() => { [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'From Current Slide')?.click(); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.sl-show-clip video'))`), 'the show', 6000).catch(() => {});
    await wait(600);
    const at = await js(`(() => { const c = document.querySelector('.sl-show-clip[data-shape] video')?.parentElement; if (!c) return null; const r = c.getBoundingClientRect(); return { x: Math.round(r.left + r.width * 0.2), y: Math.round(r.top + r.height * 0.3) }; })()`);
    const centre = await js(`(() => { const r = document.querySelector('.sl-show-stage')?.getBoundingClientRect(); return r ? { x: Math.round(r.left + r.width * (0.375 + 0.25 * 0.2)), y: Math.round(r.top + r.height * (0.375 + 0.25 * 0.3)) } : null; })()`) || { x: -99, y: -99 };
    if (at) {
      win.webContents.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      win.webContents.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1 });
    }
    const showing = await until(() => js(`(() => { const v = document.querySelector('.sl-show-clip video'); return Boolean(v) && v.currentTime > 0.2; })()`), 'the video playing in the show', 6000).then(() => true).catch(() => false);
    const stayed = await js(`document.querySelector('.sl-present-bar')?.textContent.trim().split(' ')[0] || null`);
    win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'Escape' });
    win.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'Escape' });
    await until(() => js(`!document.querySelector('.sl-present-bar')`), 'the show to end', 4000).catch(() => {});
    check('presentations: in the show a click on the video plays it, and the show stays on its slide', Boolean(at) && showing && stayed === '1' && Math.abs(at.x - centre.x) < 4 && Math.abs(at.y - centre.y) < 4, JSON.stringify({ at, centre, showing, stayed }));

    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => {
      try {
        const deck = Deck.open(fs.readFileSync(file));
        const kinds = deck.slide(0).shapes.filter((s) => s.media).map((s) => s.media.kind).sort();
        return JSON.stringify(kinds) === '["audio","video"]' && deck.pkg.has('ppt/media/media1.webm') && deck.pkg.has('ppt/media/media2.wav')
          && /<a:videoFile r:link=/.test(deck.pkg.text('ppt/slides/slide1.xml'));
      } catch { return false; }
    }, 'the saved media', 8000).then(() => true).catch(() => false);
    check('presentations: the saved deck keeps the video and the sound as PowerPoint keeps media', saved, saved ? 'media1.webm, media2.wav, videoFile and audioFile' : 'not in the file');
  } catch (err) {
    check('presentations: the media checks ran', false, err.message);
  }
}
