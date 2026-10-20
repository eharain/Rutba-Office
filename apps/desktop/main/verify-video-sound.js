// Video's sound and the picture's fades (the 2026-10-16 plan's sound and
// time): a clip's volume in decibels up to +12, the switch that lowers the
// rest under a voice, fades as long as asked with a word when together they
// are longer than the clip, and a picture clip's fade from and to black.
// A clip and a sound are made here at run time. Run alone with
// RUTBA_VERIFY_ONLY=video-sound.

import fs from 'node:fs';
import path from 'node:path';
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
    await setNumber('.vd-volume-db input', 6);
    const louder = await ok(until(() => js(`(document.querySelector('.vd-volume-db .vd-slider-value')?.textContent || '') === '+6.0 dB'`), '+6 dB', 3000));
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

    const complaints = await errorsIn(win);
    check('video sound: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('video sound: the check ran to the end', false, err?.message || String(err));
  }
}
