// Video's sound and the picture's fades (the 2026-10-16 plan's sound and
// time): a clip's volume in decibels up to +12, the switch that lowers the
// rest under a voice, fades as long as asked with a word when together they
// are longer than the clip, a picture clip's fade from and to black, a
// clip's speed, its pitch, the meter, Freeze Frame, Reverse and a crossfade
// between two sounds.
// A clip and a sound are made here at run time. Run alone with
// RUTBA_VERIFY_ONLY=video-sound.

import fs from 'node:fs';
import path from 'node:path';
import { app, nativeImage } from 'electron';
import { makeFixtures } from './verify-video-editor.js';

/** Three seconds of a tone, as a WAV a sound clip is made from. */
function toneWav(file, seconds = 3, rate = 48000) {
  const n = seconds * rate;
  const data = Buffer.alloc(n * 2);
  for (let i = 0; i < n; i++) data.writeInt16LE(Math.round(Math.sin((2 * Math.PI * 440 * i) / rate) * 12000), i * 2);
  const head = Buffer.alloc(44);
  head.write('RIFF', 0); head.writeUInt32LE(36 + data.length, 4); head.write('WAVE', 8);
  head.write('fmt ', 12); head.writeUInt32LE(16, 16); head.writeUInt16LE(1, 20); head.writeUInt16LE(1, 22);
  head.writeUInt32LE(rate, 24); head.writeUInt32LE(rate * 2, 28); head.writeUInt16LE(2, 32); head.writeUInt16LE(16, 34);
  head.write('data', 36); head.writeUInt32LE(data.length, 40);
  fs.writeFileSync(file, Buffer.concat([head, data]));
  return file;
}

export async function verifyVideoSound(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  const win = await open('video');
  const js = (code) => win.webContents.executeJavaScript(code);
  const fixtures = await makeFixtures(win.webContents, dir, 'video-sound');
  if (!fixtures) {
    check('video sound: clips could be made for the checks (VP8 through WebCodecs)', false, 'no encoder in this build');
    return;
  }
  const tone = toneWav(path.join(dir, 'video-sound-tone.wav'));
  const press = async (label) => {
    const find = `[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || n.title || n.textContent || '').trim().startsWith(${JSON.stringify(label)}) && !n.disabled)`;
    await until(() => js(`Boolean(${find})`), `the ${label} button`, 4000).catch(() => {});
    return js(`(() => { const b = ${find}; if (!b) return 'no button'; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'pressed'; })()`);
  };
  const addWith = async (files) => {
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'open', answer: files });
    return press('Add Media');
  };
  const pickKind = (kind) => js(`(() => {
    const n = document.querySelector('.vd-clip[data-kind="${kind}"]');
    if (!n) return false;
    const r = n.getBoundingClientRect();
    const o = { bubbles: true, cancelable: true, button: 0, clientX: r.left + Math.min(8, r.width / 2), clientY: r.top + r.height / 2, pointerId: 1 };
    n.dispatchEvent(new PointerEvent('pointerdown', o));
    window.dispatchEvent(new PointerEvent('pointerup', o));
    return true;
  })()`);
  const setNumber = (selector, value) => js(`(() => { const n = document.querySelector(${JSON.stringify(selector)}); if (!n) return 'none'; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, ${JSON.stringify(String(value))}); n.dispatchEvent(new Event('input', { bubbles: true })); n.dispatchEvent(new Event('change', { bubbles: true })); return 'set'; })()`);
  try {
    await addWith([fixtures.a]);
    await until(() => js(`Boolean(document.querySelector('.vd-clip[data-kind="video"]'))`), 'the clip on the timeline', 8000).catch(() => {});
    await addWith([tone]);
    const added = await ok(until(() => js(`Boolean(document.querySelector('.vd-clip[data-kind="sound"]'))`), 'the sound on the timeline', 8000));

    // The sound: volume in decibels, turned up past 0 dB, and the switch that lowers the rest.
    await pickKind('sound');
    await until(() => js(`Boolean(document.querySelector('.vd-volume-db input'))`), 'the volume in decibels', 4000).catch(() => {});
    const before = await js(`document.querySelector('.vd-volume-db .vd-slider-value')?.textContent || ''`);
    // Set once the panel has settled, and again if it redrew under the setting, as the switch below is.
    await wait(400);
    let louder = false;
    for (let attempt = 0; attempt < 2 && !louder; attempt++) {
      await setNumber('.vd-volume-db input', 6);
      louder = await ok(until(() => js(`(document.querySelector('.vd-volume-db .vd-slider-value')?.textContent || '') === '+6.0 dB'`), '+6 dB', 3000));
    }
    const duckBefore = await js(`(() => { const n = document.querySelectorAll('.vd-duck'); return { n: n.length, checked: n[0]?.checked, disabled: n[0]?.disabled }; })()`);
    // After the volume's own redraw: a click while the panel redraws can land on the box it replaces.
    await wait(400);
    let ducks = false;
    for (let attempt = 0; attempt < 2 && !ducks; attempt++) {
      await js(`document.querySelector('.vd-duck')?.click(), 1`);
      ducks = await ok(until(() => js(`document.querySelector('.vd-duck')?.checked === true`), 'lower the rest', 1500));
    }
    const duckAfter = await js(`document.querySelector('.vd-duck')?.checked`);
    check('video sound: a sound clip\'s volume reads in decibels and goes past 0 dB, and Lower the rest is a switch on the clip',
      added && before === '0.0 dB' && louder && ducks, JSON.stringify({ added, before, louder, ducks, duckBefore, duckAfter }));

    // Fades longer, together, than the three-second clip: said, and met in the middle.
    const fades = await js(`[...document.querySelectorAll('.vd-insp-section')].find((s) => s.querySelector('.vd-volume-db'))?.querySelectorAll('input.vd-num') ? [...[...document.querySelectorAll('.vd-insp-section')].find((s) => s.querySelector('.vd-volume-db')).querySelectorAll('input.vd-num')].length : 0`);
    await js(`(() => { const s = [...document.querySelectorAll('.vd-insp-section')].find((x) => x.querySelector('.vd-volume-db')); const [fin, fout] = s.querySelectorAll('input.vd-num'); for (const [n, v] of [[fin, '2.5'], [fout, '2.5']]) { Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, v); n.dispatchEvent(new Event('input', { bubbles: true })); } return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.vd-fades-note'))`), 'the fades note', 3000).catch(() => {});
    const note = await js(`document.querySelector('.vd-fades-note')?.textContent || ''`);
    check('video sound: fades as long as asked; when together they are longer than the clip the inspector says so and where they meet',
      fades === 2 && /1\.5 s in, 1\.5 s out/.test(note || ''), JSON.stringify({ fades, note }));

    // The picture clip: fade from black and to black.
    await pickKind('video');
    const shown = await ok(until(() => js(`Boolean(document.querySelector('.vd-picture-fades input.vd-num'))`), 'the picture\'s fades', 4000));
    await js(`(() => { const [a, b] = document.querySelectorAll('.vd-picture-fades input.vd-num'); for (const [n, v] of [[a, '1'], [b, '0.5']]) { Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, v); n.dispatchEvent(new Event('input', { bubbles: true })); } return 1; })()`);
    await wait(300);
    const kept = await js(`[...document.querySelectorAll('.vd-picture-fades input.vd-num')].map((n) => n.value)`);
    check('video sound: a picture clip fades from black at its start and to black at its end, as long as asked',
      shown && kept[0] === '1' && kept[1] === '0.5', JSON.stringify({ shown, kept }));

    // Speed: the picture clip at twice its speed takes half the time on the timeline, and back.
    const span = () => js(`(() => { const n = document.querySelector('.vd-clip[data-kind="video"]'); return n ? Number(n.dataset.end) - Number(n.dataset.start) : 0; })()`);
    const was = await span();
    await js(`(() => { const n = document.querySelector('.vd-speed-select'); if (!n) return 'none'; Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(n, '2'); n.dispatchEvent(new Event('change', { bubbles: true })); return 1; })()`);
    const halved = await ok(until(async () => Math.abs((await span()) - was / 2) < 0.06, 'the clip at twice its speed', 4000));
    const fast = await span();
    await pickKind('video');
    await until(() => js(`Boolean(document.querySelector('.vd-speed-select'))`), 'the speed again', 3000).catch(() => {});
    await js(`(() => { const n = document.querySelector('.vd-speed-select'); Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(n, '1'); n.dispatchEvent(new Event('change', { bubbles: true })); return 1; })()`);
    const restored = await ok(until(async () => Math.abs((await span()) - was) < 0.06, 'the clip at its own speed', 4000));
    check('video sound: a clip at twice its speed takes half the time on the timeline, and set back takes its own',
      was > 0.5 && halved && restored, JSON.stringify({ was, fast, halved, restored }));

    // The meter beside the time: dark while stopped, a level while the tone plays.
    const idle = await js(`document.querySelector('.vd-meter')?.dataset.db ?? 'none'`);
    await js(`document.querySelector('.vd-play')?.click(), 1`);
    const moving = await ok(until(() => js(`(() => { const d = document.querySelector('.vd-meter')?.dataset.db; return d !== undefined && d !== '' && Number(d) > -60; })()`), 'the meter moving', 5000));
    const level = await js(`document.querySelector('.vd-meter')?.dataset.db`);
    await js(`document.querySelector('.vd-play')?.click(), 1`);
    const stilled = await ok(until(() => js(`(document.querySelector('.vd-meter')?.dataset.db ?? 'x') === ''`), 'the meter at rest', 4000));
    check('video sound: while the preview plays the meter beside the time shows the sound\'s level, and it rests when stopped',
      idle === '' && moving && stilled, JSON.stringify({ idle, moving, level, stilled }));

    // Pitch: a sound at twice its speed keeps its pitch unless asked not to, a pitch moved in semitones, and it plays.
    const selectTo = (selector, value) => js(`(() => { const n = document.querySelector(${JSON.stringify(selector)}); if (!n) return 'none'; Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(n, ${JSON.stringify(String(value))}); n.dispatchEvent(new Event('change', { bubbles: true })); return 'set'; })()`);
    await pickKind('sound');
    await until(() => js(`Boolean(document.querySelector('.vd-pitch-select'))`), 'the pitch', 4000).catch(() => {});
    const noKeepAt1 = await js(`!document.querySelector('.vd-keep-pitch')`);
    await selectTo('.vd-speed-select', 2);
    const keepShown = await ok(until(() => js(`document.querySelector('.vd-keep-pitch')?.checked === true`), 'keep the pitch, on', 4000));
    await selectTo('.vd-pitch-select', 12);
    const pitched = await ok(until(() => js(`document.querySelector('.vd-pitch-select')?.value === '12'`), 'an octave up', 3000));
    await js(`document.querySelector('.vd-play')?.click(), 1`);
    const plays = await ok(until(() => js(`Number(document.querySelector('.vd-meter')?.dataset.db) > -60`), 'the stretched sound playing', 8000));
    await js(`document.querySelector('.vd-play')?.click(), 1`);
    await wait(300);
    await pickKind('sound');
    await until(() => js(`Boolean(document.querySelector('.vd-pitch-select'))`), 'the pitch again', 3000).catch(() => {});
    await selectTo('.vd-pitch-select', 0);
    await wait(200);
    await selectTo('.vd-speed-select', 1);
    const reset = await ok(until(() => js(`document.querySelector('.vd-pitch-select')?.value === '0' && document.querySelector('.vd-speed-select')?.value === '1'`), 'the sound as recorded', 3000));
    check('video sound: at another speed a sound keeps its pitch (a switch turns that off), its pitch moves in semitones, and stretched it plays',
      noKeepAt1 && keepShown && pitched && plays && reset, JSON.stringify({ noKeepAt1, keepShown, pitched, plays, reset }));

    // Freeze Frame: the picture at the playhead as a still of two seconds, the clip cut there and going on after it.
    const videosWas = app.getPath('videos');
    app.setPath('videos', dir);
    try {
      await js(`[...document.querySelectorAll('.rw-tab')].find((n) => n.textContent.trim() === 'Edit')?.click(), 1`);
      await js(`(() => { const s = document.querySelector('.vd-scrub'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(s, '1'); s.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
      await until(() => js(`Math.abs(Number(document.querySelector('.vd-scrub').value) - 1) < 0.02`), 'the playhead at 1 s', 3000).catch(() => {});
      await wait(200);
      const pressed = await press('Freeze Frame');
      const held = await ok(until(() => js(`Boolean(document.querySelector('.vd-clip[data-kind="image"]'))`), 'the frozen frame', 8000));
      const shape = await js(`(() => {
        const img = document.querySelector('.vd-clip[data-kind="image"]');
        const vids = [...document.querySelectorAll('.vd-clip[data-kind="video"]')].map((n) => [Number(n.dataset.start), Number(n.dataset.end)]).sort((a, b) => a[0] - b[0]);
        return { still: img ? [Number(img.dataset.start), Number(img.dataset.end)] : null, vids };
      })()`);
      const folder = path.join(dir, 'Rutba Recordings');
      const files = fs.existsSync(folder) ? fs.readdirSync(folder).filter((f) => /^Frame .*\.png$/.test(f)) : [];
      const png = files.length ? fs.readFileSync(path.join(folder, files[0])) : null;
      const isPng = Boolean(png && png.length > 1000 && png.subarray(1, 4).toString() === 'PNG');
      // The clip's own picture, not a black or empty frame.
      const bitmap = png ? nativeImage.createFromBuffer(png).toBitmap() : Buffer.alloc(0);
      let lit = 0;
      for (let i = 0; i < bitmap.length; i += 4 * 97) if (bitmap[i] + bitmap[i + 1] + bitmap[i + 2] > 60) lit++;
      const pictured = lit > (bitmap.length / (4 * 97)) * 0.5;
      const s0 = shape.still || [0, 0];
      check('video sound: Freeze Frame holds the picture at the playhead for two seconds as a PNG still, the clip cut there and going on after it',
        pressed === 'pressed' && held && Math.abs(s0[0] - 1) < 0.05 && Math.abs(s0[1] - s0[0] - 2) < 0.05 && shape.vids.length === 2 && Math.abs(shape.vids[1][0] - s0[1]) < 0.05 && isPng && pictured,
        JSON.stringify({ pressed, held, shape, files, isPng, pictured }));

      // Reverse: the first part of the clip from a copy written backward (the ball runs right to left), and back.
      const reverseButton = `[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || n.title || n.textContent || '').trim().startsWith('Reverse'))`;
      await js(`(() => {
        const n = [...document.querySelectorAll('.vd-clip[data-kind="video"]')].find((x) => Number(x.dataset.start) === 0);
        if (!n) return false;
        const r = n.getBoundingClientRect();
        const o = { bubbles: true, cancelable: true, button: 0, clientX: r.left + Math.min(8, r.width / 2), clientY: r.top + r.height / 2, pointerId: 1 };
        n.dispatchEvent(new PointerEvent('pointerdown', o));
        window.dispatchEvent(new PointerEvent('pointerup', o));
        return true;
      })()`);
      const reversing = await press('Reverse');
      const reversed = await ok(until(() => js(`${reverseButton}?.getAttribute('aria-pressed') === 'true'`), 'the clip reversed', 60000));
      const copies = fs.existsSync(folder) ? fs.readdirSync(folder).filter((f) => / reversed( \(\d+\))?\.mp4$/.test(f)) : [];
      const copy = copies.length ? path.join(folder, copies[0]) : null;
      // Where the ball is on the preview near the clip's start and near its end: right to left when reversed.
      const ballAt = async (t) => {
        await js(`(() => { const s = document.querySelector('.vd-scrub'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(s, '${t}'); s.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
        await wait(700);
        return js(`(() => {
          const c = document.querySelector('.vd-canvas');
          const d = c.getContext('2d').getImageData(0, 0, c.width, Math.floor(c.height * 0.85)).data;
          // The whitest pixels (the clip fades from and to black, so not a fixed level).
          let top = 0;
          for (let i = 0; i < d.length; i += 4) top = Math.max(top, Math.min(d[i], d[i + 1], d[i + 2]));
          let sx = 0, n = 0;
          for (let i = 0; i < d.length; i += 4) if (Math.min(d[i], d[i + 1], d[i + 2]) >= top * 0.9) { sx += (i / 4) % c.width; n++; }
          return top > 8 && n ? Math.round(sx / n) : null;
        })()`);
      };
      const ball = { first: await ballAt(0.3), last: await ballAt(0.7) };
      const backward = Boolean(ball && ball.first != null && ball.last != null && ball.first > ball.last + 20);
      await wait(300);
      await press('Reverse');
      const forward = await ok(until(() => js(`${reverseButton}?.getAttribute('aria-pressed') === 'false'`), 'the clip forward again', 5000));
      check('video sound: Reverse plays the clip and its sound from a copy written backward, made once, and reversed again it is back on its source',
        reversing === 'pressed' && reversed && backward && forward, JSON.stringify({ reversing, reversed, copies, ball, forward }));
    } finally {
      app.setPath('videos', videosWas);
    }

    // A crossfade between two sounds of their own: the second tone overlaps the first by a second, fading in as it fades out.
    await js(`[...document.querySelectorAll('.rw-tab')].find((n) => n.textContent.trim() === 'Home')?.click(), 1`);
    await addWith([tone]);
    const tones = `[...document.querySelectorAll('.vd-clip[data-kind="sound"]')].filter((n) => n.textContent.includes('video-sound-tone')).sort((a, b) => Number(a.dataset.start) - Number(b.dataset.start))`;
    const two = await ok(until(() => js(`${tones}.length === 2`), 'a second tone', 8000));
    const soundsSeen = await js(`[...document.querySelectorAll('.vd-clip[data-kind="sound"]')].map((n) => [n.dataset.start, n.dataset.lane, n.textContent.slice(0, 40)])`);
    await js(`(() => {
      const n = ${tones}[1];
      if (!n) return 0;
      const r = n.getBoundingClientRect();
      const o = { bubbles: true, cancelable: true, button: 0, clientX: r.left + Math.min(8, r.width / 2), clientY: r.top + r.height / 2, pointerId: 1 };
      n.dispatchEvent(new PointerEvent('pointerdown', o));
      window.dispatchEvent(new PointerEvent('pointerup', o));
      return 1;
    })()`);
    const kindSelect = `[...document.querySelectorAll('.vd-insp-section')].find((x) => x.querySelector('h3')?.textContent === 'Transition in')?.querySelector('select')`;
    const offered = await ok(until(() => js(`Boolean(${kindSelect})`), 'the transition for the sound', 4000));
    const options = await js(`[...(${kindSelect}?.options || [])].map((o) => o.value)`);
    await js(`(() => { const n = ${kindSelect}; if (!n) return 'none'; Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(n, 'fade'); n.dispatchEvent(new Event('change', { bubbles: true })); return 1; })()`);
    const faded = await ok(until(() => js(`${tones}[1]?.dataset.transition === 'fade'`), 'the crossfade', 4000));
    const spans = await js(`${tones}.map((n) => [Number(n.dataset.start), Number(n.dataset.end)])`);
    const overlap = spans.length === 2 ? Math.round((spans[0][1] - spans[1][0]) * 100) / 100 : null;
    check('video sound: two sounds of their own crossfade, the second overlapping the first by its length; a crossfade is the one kind offered',
      two && offered && faded && overlap === 1 && options.join(',') === 'none,fade', JSON.stringify({ two, soundsSeen, offered, options, faded, spans, overlap }));

    const complaints = await errorsIn(win);
    check('video sound: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('video sound: the check ran to the end', false, err?.message || String(err));
  }
}
