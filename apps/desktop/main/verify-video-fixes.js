// Video: a QR code that waits for its address, an end card picked on the
// preview, and the audio library's fetch from the web (Office 2.27.0).
//
// A QR code is added with nothing to scan: the frame shows a box where it
// stands that says what it waits for, the box its address is typed in takes
// the caret, and the box on the frame is picked and moved like a drawn one.
// With an address typed the code is drawn and the box is gone.
//
// An end card covers the frame, so once it has come in a click on the frame
// picks it: it wears the frame's own outline and no grips, and is not moved.
//
// An address in the audio library is fetched once, into a file. While it is
// fetched its row says so with Cancel, which lets go of the request itself;
// a fetch stopped is not remembered and says nothing. A sound that cannot be
// had says which of three things is wrong. The addresses here are the run's
// own server on this computer, listed in the kept list before the window
// opens (the window's own box takes https only). The project is saved in the
// run's folder first, so what is fetched is kept beside it and not in the
// person's own Videos folder.
// Run alone with RUTBA_VERIFY_ONLY=video-fixes.

import fs from 'node:fs';
import http from 'node:http';
import path from 'node:path';
import { AUDIO_LIBRARY_KEY } from '@rutba/studio/audio-library';
import { makeFixtures } from './verify-video-editor.js';
import { labelKeyed } from './verify-keys.js';

/** Half a second of a 440 Hz tone as a WAV file: 16 bit, one channel, 8000 a second. */
function toneWav() {
  const rate = 8000;
  const n = rate / 2;
  const buf = Buffer.alloc(44 + n * 2);
  buf.write('RIFF', 0);
  buf.writeUInt32LE(36 + n * 2, 4);
  buf.write('WAVEfmt ', 8);
  buf.writeUInt32LE(16, 16);
  buf.writeUInt16LE(1, 20);
  buf.writeUInt16LE(1, 22);
  buf.writeUInt32LE(rate, 24);
  buf.writeUInt32LE(rate * 2, 28);
  buf.writeUInt16LE(2, 32);
  buf.writeUInt16LE(16, 34);
  buf.write('data', 36);
  buf.writeUInt32LE(n * 2, 40);
  for (let i = 0; i < n; i++) buf.writeInt16LE(Math.round(Math.sin((i / rate) * 2 * Math.PI * 440) * 9000), 44 + i * 2);
  return buf;
}

/** The run's own web: a file that comes whole, one that is held back until let go, one that is not there, and one that is no sound. */
function standInWeb() {
  const wav = toneWav();
  const seen = { good: 0, slow: 0, slowClosed: 0, held: [] };
  let hold = true;
  const server = http.createServer((req, res) => {
    const name = new URL(req.url, 'http://127.0.0.1').pathname.slice(1);
    if (name === 'good.wav') {
      seen.good += 1;
      res.writeHead(200, { 'content-type': 'audio/wav', 'content-length': wav.length });
      res.end(wav);
    } else if (name === 'slow.wav') {
      seen.slow += 1;
      res.writeHead(200, { 'content-type': 'audio/wav', 'content-length': wav.length });
      if (!hold) { res.end(wav); return; }
      // The first part goes, the rest is held back: the fetch is under way until it is stopped.
      res.write(wav.subarray(0, 1000));
      seen.held.push(res);
      res.on('close', () => { if (!res.writableEnded) seen.slowClosed += 1; });
    } else if (name === 'noise.wav') {
      res.writeHead(200, { 'content-type': 'audio/wav' });
      res.end(Buffer.from('This is not a sound at all, only words kept under a sound\'s name.'.repeat(8)));
    } else {
      res.writeHead(404, { 'content-type': 'text/plain' });
      res.end('nothing here');
    }
  });
  return {
    seen,
    release() { hold = false; },
    start: () => new Promise((resolve, reject) => { server.once('error', reject); server.listen(0, '127.0.0.1', () => resolve(server.address().port)); }),
    stop: () => new Promise((resolve) => { for (const r of seen.held) { try { r.destroy(); } catch { /* gone already */ } } server.closeAllConnections?.(); server.close(() => resolve()); }),
  };
}

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args the run's scratch folder
 */
export async function verifyVideoFixes(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  const web = standInWeb();
  let win = null;
  let js = null;
  let keptList;
  try {
    const port = await web.start();
    const at = (name) => `http://127.0.0.1:${port}/${name}`;
    const missing = path.join(dir, 'video-fixes', 'was-here-once.wav');
    win = await open('video');
    js = (code) => win.webContents.executeJavaScript(code);
    const fixtures = await makeFixtures(win.webContents, dir, 'video-fixes');
    if (!fixtures) {
      check('video fixes: a clip could be made for the checks (VP8 through WebCodecs)', false, 'no encoder in this build');
      return;
    }
    // The list the window reads at its start: four addresses of the run's own web, and a file that is not there.
    keptList = await js(`window.rutbaOffice.store.get({ key: ${JSON.stringify(AUDIO_LIBRARY_KEY)}, fallback: null })`);
    const list = [
      { id: 'c1', kind: 'remote', url: at('slow.wav'), name: 'slow.wav' },
      { id: 'c2', kind: 'remote', url: at('good.wav'), name: 'good.wav' },
      { id: 'c3', kind: 'remote', url: at('away.wav'), name: 'away.wav' },
      { id: 'c4', kind: 'remote', url: at('noise.wav'), name: 'noise.wav' },
      { id: 'c5', kind: 'file', path: missing, name: 'was-here-once.wav' },
    ];
    await js(`window.rutbaOffice.store.set({ key: ${JSON.stringify(AUDIO_LIBRARY_KEY)}, value: ${JSON.stringify(list)} }).then(() => 1)`);
    await win.webContents.reload();
    await until(() => js(`Boolean(document.querySelector('.vd-scrub')) && Boolean(document.querySelector('.vd-lib-tabs'))`), 'the Video window', 20000);

    const q = (sel) => `document.querySelector(${JSON.stringify(sel)})`;
    const has = (sel) => js(`Boolean(${q(sel)})`);
    const click = (sel) => js(`(${q(sel)}?.click(), 1)`);
    const count = (kind) => js(`document.querySelectorAll('.vd-clip[data-kind="${kind}"]').length`);
    const type = (sel, value) => js(`(() => { const n = ${q(sel)}; if (!n) return 0; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, ${JSON.stringify(value)}); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    const tab = async (name) => {
      await click(`.vd-lib-tabs [data-tab="${name}"]`);
      await until(() => js(`${q(`.vd-lib-tabs [data-tab="${name}"]`)}?.getAttribute('aria-selected') === 'true'`), `the ${name} tab`, 4000).catch(() => {});
      await wait(200);
    };
    const seekTo = async (t) => {
      await type('.vd-scrub', String(t));
      await until(() => js(`Math.abs(Number(document.querySelector('.vd-scrub').value) - ${t}) < 0.02`), 'the playhead to move', 3000).catch(() => {});
      await wait(300);
    };
    /** The picked frame as the page draws it, in client pixels, with its grips; and the frame itself. */
    const picked = () => js(`(() => {
      const f = document.querySelector('.vd-frame');
      const o = document.querySelector('.vd-outline');
      const fr = f.getBoundingClientRect();
      const grips = [...document.querySelectorAll('.vd-grip')].map((g) => g.dataset.grip);
      const frame = { x: fr.left, y: fr.top, w: fr.width, h: fr.height };
      if (!o) return { id: f.dataset.picked || '', frame, grips };
      const box = { x: fr.left + parseFloat(o.style.left), y: fr.top + parseFloat(o.style.top), w: parseFloat(o.style.width), h: parseFloat(o.style.height) };
      return { id: f.dataset.picked || '', frame, box, grips };
    })()`);
    /** The box on the frame of a QR code that waits for its address, in client pixels, with its words. */
    const waiting = () => js(`(() => { const n = document.querySelector('.vd-qr-empty'); if (!n) return null; const r = n.getBoundingClientRect(); return { id: n.dataset.qr, x: r.left, y: r.top, w: r.width, h: r.height, words: n.textContent.trim() }; })()`);
    const drag = (x0, y0, x1, y1) => js(`(async () => {
      const c = document.querySelector('.vd-canvas');
      const at = (x, y) => ({ bubbles: true, cancelable: true, button: 0, clientX: x, clientY: y, pointerId: 1 });
      c.dispatchEvent(new PointerEvent('pointerdown', at(${x0}, ${y0})));
      for (let k = 1; k <= 8; k++) {
        window.dispatchEvent(new PointerEvent('pointermove', at(${x0} + (${x1 - x0} * k) / 8, ${y0} + (${y1 - y0} * k) / 8)));
        await new Promise((r) => requestAnimationFrame(r));
      }
      window.dispatchEvent(new PointerEvent('pointerup', at(${x1}, ${y1})));
      return 1;
    })()`);
    const clickAt = (x, y) => drag(x, y, x, y);
    const near = (a, b, tol = 3) => Math.abs(a - b) <= tol;
    /** How much of a box on the frame (client pixels) is drawn nearly white and nearly black. */
    const inks = (b) => js(`(() => {
      const c = document.querySelector('.vd-canvas');
      const f = document.querySelector('.vd-frame').getBoundingClientRect();
      const k = c.width / f.width;
      const x = Math.max(0, Math.floor((${b.x} - f.left) * k)), y = Math.max(0, Math.floor((${b.y} - f.top) * k));
      const w = Math.max(1, Math.floor(${b.w} * k)), hgt = Math.max(1, Math.floor(${b.h} * k));
      const d = c.getContext('2d').getImageData(x, y, w, hgt).data;
      let white = 0, black = 0;
      for (let i = 0; i < d.length; i += 4) { const s = d[i] + d[i + 1] + d[i + 2]; if (s > 690) white += 1; else if (s < 75) black += 1; }
      const n = d.length / 4;
      return { white: white / n, black: black / n };
    })()`);

    await tab('media');
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'open', answer: [fixtures.a] });
    await js(`(() => { const b = document.querySelector('.vd-lib-add-files'); b?.click(); return Boolean(b); })()`);
    const landed = await ok(until(async () => (await count('video')) === 1, 'the clip', 20000));
    const clipId = await js(`document.querySelector('.vd-clip[data-kind="video"]')?.dataset.id || ''`);
    if (!landed) {
      check('video fixes: a clip is on the timeline to work over', false, 'Add Media put nothing there');
      return;
    }

    // Saved in the run's folder: what the film fetches is then kept in a folder beside it.
    const project = path.join(dir, 'video-fixes', 'film.rvid');
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: project });
    await js(`(() => { const b = ((list) => list.find((n) => ${labelKeyed('Save', { tip: true, start: true })}(n)) || list.find((n) => (n.dataset.tip || '').startsWith('Save')))([...document.querySelectorAll('.rw-ribbon .rw-btn, .rw-qat .rw-btn')]); b?.click(); return Boolean(b); })()`);
    const savedFirst = await ok(until(() => fs.existsSync(project) && js(`document.querySelector('.vd-editor')?.dataset.dirty === '0'`), 'the project file', 8000));
    if (!savedFirst) {
      check('video fixes: the film is saved in the folder of the run before anything is fetched for it', false, 'Save wrote no project');
      return;
    }
    /** The files fetched for the film, by name: each in a folder beside the project. */
    const fetched = (name) => fs.readdirSync(path.dirname(project), { withFileTypes: true }).filter((e) => e.isDirectory()).some((e) => fs.existsSync(path.join(path.dirname(project), e.name, name)));

    /* ── a QR code that waits for its address ── */
    await seekTo(0.5);
    await tab('overlays');
    await click('.vd-lib-overlay[data-key="overlay-qr"]');
    const added = await ok(until(async () => (await count('qr')) === 1 && Boolean(await waiting()), 'the QR code and its box on the frame', 6000));
    const qrId = await js(`document.querySelector('.vd-clip[data-kind="qr"]')?.dataset.id || ''`);
    const caret = await ok(until(() => js(`document.activeElement?.classList.contains('vd-qr-data') || false`), 'the caret in the address box', 4000));
    const blank = await waiting();
    let f = await picked();
    const emptyBox = await js(`${q('.vd-qr-data')}?.value === '' && (${q('.vd-qr-data')}?.placeholder || '').length > 3`);
    const before = blank ? await inks(blank) : null;
    check('video fixes: a QR code is added with no address: a box on the frame says what it waits for where the code will stand, picked with its grips, and the box its address is typed in is empty and has the caret',
      added && Boolean(blank) && blank.id === qrId && blank.words.length > 5 && near(blank.w, f.frame.w * 0.18, 4) && near(blank.w, blank.h, 3) && f.id === qrId && Boolean(f.box) && near(f.box.x, blank.x, 4) && near(f.box.w, blank.w, 4) && f.grips.length >= 4 && caret && emptyBox === true,
      JSON.stringify({ added, blank, picked: f.id, grips: f.grips, caret, emptyBox }));

    // Something else picked, then a click on the waiting box picks the code again, and a drag moves it.
    await clickAt(f.frame.x + f.frame.w * 0.15, f.frame.y + f.frame.h * 0.8);
    await until(async () => (await picked()).id === clipId, 'the clip picked', 3000).catch(() => {});
    const other = (await picked()).id;
    await clickAt(blank.x + blank.w / 2, blank.y + blank.h / 2);
    await until(async () => (await picked()).id === qrId, 'the QR code picked on the frame', 3000).catch(() => {});
    const again = (await picked()).id;
    await drag(blank.x + blank.w / 2, blank.y + blank.h / 2, blank.x + blank.w / 2 - 60, blank.y + blank.h / 2 + 40);
    await wait(300);
    const moved = await waiting();
    check('video fixes: the waiting box is picked by a click on the frame and moved by a drag, like a code that is drawn',
      other === clipId && again === qrId && Boolean(moved) && near(moved.x - blank.x, -60, 4) && near(moved.y - blank.y, 40, 4),
      JSON.stringify({ other, again, by: moved ? [Math.round(moved.x - blank.x), Math.round(moved.y - blank.y)] : null }));

    // An address typed: the code is drawn and the box is gone.
    await type('.vd-qr-data', 'https://rutba.io/office');
    const drawn = await ok(until(async () => !(await waiting()), 'the box gone once there is something to scan', 5000));
    await wait(500);
    const after = moved ? await inks(moved) : null;
    check('video fixes: with an address typed the code is drawn where the box stood, and the box is gone',
      drawn && Boolean(before) && Boolean(after) && before.white < 0.05 && after.white > 0.25 && after.black > 0.1,
      JSON.stringify({ drawn, before, after }));

    /* ── an end card, picked on the preview once it has come in ── */
    await click('.vd-lib-overlay[data-key="overlay-outro"]');
    await ok(until(async () => (await count('outro')) === 1, 'the end card', 6000));
    const outroId = await js(`document.querySelector('.vd-clip[data-kind="outro"]')?.dataset.id || ''`);
    const outroStart = await js(`Number(document.querySelector('.vd-clip[data-kind="outro"]')?.dataset.start || 0)`);
    // The clip picked, so nothing of the end card is picked beforehand.
    await seekTo(0.2);
    f = await picked();
    await clickAt(f.frame.x + f.frame.w * 0.15, f.frame.y + f.frame.h * 0.8);
    await until(async () => (await picked()).id === clipId, 'the clip picked', 3000).catch(() => {});
    await seekTo(outroStart + 0.1);
    await clickAt(f.frame.x + f.frame.w / 2, f.frame.y + f.frame.h / 2);
    await wait(300);
    const early = (await picked()).id;
    await seekTo(outroStart + 0.7);
    await clickAt(f.frame.x + f.frame.w / 2, f.frame.y + f.frame.h / 2);
    await until(async () => (await picked()).id === outroId, 'the end card picked', 3000).catch(() => {});
    const card = await picked();
    const onTimeline = await js(`document.querySelector('.vd-clip[data-id="${outroId}"]')?.classList.contains('selected') || false`);
    await drag(f.frame.x + f.frame.w / 2, f.frame.y + f.frame.h / 2, f.frame.x + f.frame.w / 2 + 50, f.frame.y + f.frame.h / 2 + 30);
    await wait(300);
    const held = await picked();
    check('video fixes: an end card is added after the film; a click on the frame as it starts to come in picks nothing of it, and once it has come in picks it, on the timeline too, with the frame\'s own outline and no grips, and a drag does not move it',
      Boolean(outroId) && outroStart >= 1.5 && early !== outroId && card.id === outroId && onTimeline && Boolean(card.box) && near(card.box.x, f.frame.x) && near(card.box.y, f.frame.y) && near(card.box.w, f.frame.w) && near(card.box.h, f.frame.h) && card.grips.length === 0
        && held.id === outroId && Boolean(held.box) && near(held.box.x, card.box.x, 1) && near(held.box.y, card.box.y, 1),
      JSON.stringify({ outroStart, early, picked: card.id, onTimeline, box: card.box, frame: card.frame, grips: card.grips }));

    /* ── the audio library: an address fetched, stopped, and what cannot be had ── */
    await tab('sounds');
    const row = (address) => `.vd-audio-entry[title=${JSON.stringify(address)}]`;
    const listed = await ok(until(() => has(row(at('slow.wav'))), 'the list', 6000));
    const rights = await js(`${q('.vd-audio-rights')}?.textContent.trim().length || 0`);
    const sounds = () => count('sound');
    const soundsAtStart = await sounds();
    const bad = () => js(`[...document.querySelectorAll('.rw-toast.bad')].map((n) => n.textContent.trim())`);

    await click(`${row(at('slow.wav'))} .vd-lib-add-sound`);
    const fetching = await ok(until(() => has(`${row(at('slow.wav'))}.fetching .vd-audio-cancel`), 'the row says it is being fetched', 8000));
    const says = await js(`${q(`${row(at('slow.wav'))} .vd-audio-fetching`)}?.textContent.trim().length || 0`);
    const asked = web.seen.slow;
    await click(`${row(at('slow.wav'))} .vd-audio-cancel`);
    const stopped = await ok(until(() => js(`!${q(`${row(at('slow.wav'))}.fetching`)}`), 'the row as it was', 8000));
    const letGo = await ok(until(() => web.seen.slowClosed >= 1 || null, 'the request let go of', 8000));
    await wait(400);
    const saidOfStop = await bad();
    const halfKept = fetched('slow.wav');
    check('video fixes: an address being fetched says so in its row with Cancel; Cancel lets go of the request itself, keeps no half file, adds nothing and says nothing; and a line says whose music may be used',
      listed && rights > 20 && fetching && says > 3 && asked === 1 && stopped && letGo && !halfKept && saidOfStop.length === 0 && (await sounds()) === soundsAtStart,
      JSON.stringify({ listed, rights, fetching, says, asked, stopped, letGo, halfKept, saidOfStop, sounds: (await sounds()) - soundsAtStart }));

    // A fetch that was stopped is not remembered: asked again, the address is fetched again, whole this time.
    web.release();
    await click(`${row(at('slow.wav'))} .vd-lib-add-sound`);
    const afterStop = await ok(until(async () => (await sounds()) === soundsAtStart + 1, 'the sound on the timeline', 15000));
    // An address is fetched once: added twice, asked for once.
    await click(`${row(at('good.wav'))} .vd-lib-add-sound`);
    await ok(until(async () => (await sounds()) === soundsAtStart + 2, 'the second sound', 15000));
    await click(`${row(at('good.wav'))} .vd-lib-add-sound`);
    const twice = await ok(until(async () => (await sounds()) === soundsAtStart + 3, 'the same sound again', 15000));
    const beside = fetched('slow.wav') && fetched('good.wav');
    check('video fixes: the address that was stopped is fetched afresh when asked again and goes on the timeline; an address added twice is fetched once; the files are kept in a folder beside the film',
      afterStop && web.seen.slow === 2 && twice && web.seen.good === 1 && beside,
      JSON.stringify({ afterStop, slowAsked: web.seen.slow, twice, goodAsked: web.seen.good, beside, sounds: (await sounds()) - soundsAtStart }));

    // The three things that can be wrong, each said in its own words with the entry's name.
    const saidFor = async (sel, name) => {
      const was = (await bad()).length;
      await click(`${sel} .vd-lib-add-sound`);
      await ok(until(async () => (await bad()).some((s) => s.includes(name)), `what is wrong with ${name}`, 12000));
      const said = (await bad()).find((s) => s.includes(name)) || '';
      return { said: said.split(name).join('%'), fresh: (await bad()).length > was };
    };
    const away = await saidFor(row(at('away.wav')), 'away.wav');
    const noise = await saidFor(row(at('noise.wav')), 'noise.wav');
    const gone = await saidFor(row(missing), 'was-here-once.wav');
    const three = new Set([away.said, noise.said, gone.said]);
    check('video fixes: a sound that cannot be fetched, one that cannot be played and a file that is gone are each said in different words with the entry\'s name, and none is put on the timeline',
      [away, noise, gone].every((x) => x.said.length > 10 && x.said.includes('%')) && three.size === 3 && (await sounds()) === soundsAtStart + 3,
      JSON.stringify({ away: away.said, noise: noise.said, gone: gone.said, sounds: (await sounds()) - soundsAtStart }));

    // What was said of the three sounds above was asked for; anything else said in alarm is a fault.
    const complaints = (await errorsIn(win)).filter((s) => !['away.wav', 'noise.wav', 'was-here-once.wav'].some((name) => s.includes(name)));
    check('video fixes: none of that reported an error beyond the three sounds that could not be had', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('video fixes: the check ran to the end', false, err?.message || String(err));
  } finally {
    if (js && keptList !== undefined && win && !win.isDestroyed()) {
      await js(`window.rutbaOffice.store.set({ key: ${JSON.stringify(AUDIO_LIBRARY_KEY)}, value: ${JSON.stringify(keptList ?? [])} }).then(() => 1)`).catch(() => {});
    }
    await web.stop().catch(() => {});
    try { fs.rmSync(path.join(dir, 'video-fixes', 'was-here-once.wav'), { force: true }); } catch { /* never made */ }
  }
}
