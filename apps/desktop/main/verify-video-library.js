// Video's library and panels (the 2.7 plan's library and panels): the
// project's media kept to add again, titles, shapes and icons drawn as what
// they add, sound effects made on this computer, a search, and the library
// and the properties each folded to a rail and opened again, and the preview
// in a window of its own for another screen. A clip is made here at run time. Run alone with RUTBA_VERIFY_ONLY=video-library.

import fs from 'node:fs';
import path from 'node:path';
import { app, BrowserWindow } from 'electron';
import { makeFixtures } from './verify-video-editor.js';

export async function verifyVideoLibrary(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  const win = await open('video');
  const js = (code) => win.webContents.executeJavaScript(code);
  const fixtures = await makeFixtures(win.webContents, dir, 'video-library');
  if (!fixtures) {
    check('video library: a clip could be made for the checks (VP8 through WebCodecs)', false, 'no encoder in this build');
    return;
  }
  const videosWas = app.getPath('videos');
  app.setPath('videos', dir);
  const count = (kind) => js(`document.querySelectorAll('.vd-clip[data-kind="${kind}"]').length`);
  const tab = async (k) => {
    await js(`document.querySelector('.vd-lib-tabs [data-tab="${k}"]')?.click(), 1`);
    await until(() => js(`document.querySelector('.vd-library')?.dataset.tab === '${k}'`), `the ${k} tab`, 3000).catch(() => {});
  };
  try {
    // The library beside the preview, on its media, empty to begin with.
    const shown = await ok(until(() => js(`Boolean(document.querySelector('.vd-library'))`), 'the library', 6000));
    const empty = await js(`Boolean(document.querySelector('.vd-lib-empty'))`);

    // Media: what was added is kept, and a double-click adds it again.
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'open', answer: [fixtures.a] });
    await js(`document.querySelector('.vd-lib-add-files')?.click(), 1`);
    const listed = await ok(until(() => js(`document.querySelectorAll('.vd-lib-source').length === 1`), 'the clip in the library', 8000));
    await until(() => js(`document.querySelectorAll('.vd-clip[data-kind="video"]').length === 1`), 'the clip on the timeline', 6000).catch(() => {});
    await js(`document.querySelector('.vd-lib-source')?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true })), 1`);
    const again = await ok(until(async () => (await count('video')) === 2, 'the clip added again', 6000));
    check('video library: the media of the project is kept in the library, and a double-click adds it again',
      shown && empty && listed && again, JSON.stringify({ shown, empty, listed, again }));

    // Shapes: drawn tiles, one added at the playhead on a track of its own.
    await tab('shapes');
    const tiles = await js(`document.querySelectorAll('.vd-lib-shape-tile').length`);
    const drawn = await js(`(() => { const c = document.querySelector('.vd-lib-shape-tile[data-key="star"] canvas'); if (!c) return false; const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data; for (let i = 3; i < d.length; i += 4) if (d[i] > 0) return true; return false; })()`);
    const starName = await js(`document.querySelector('.vd-lib-shape-tile[data-key="star"]')?.getAttribute('aria-label') || ''`);
    await js(`document.querySelector('.vd-lib-shape-tile[data-key="star"]')?.click(), 1`);
    const shape = await ok(until(async () => (await count('shape')) === 1, 'the star on the timeline', 4000));
    const shapeLane = await js(`document.querySelector('.vd-clip[data-kind="shape"]')?.dataset.lane`);
    check('video library: shapes are drawn as what they add, and one is put at the playhead on a track above the pictures',
      tiles >= 20 && drawn && starName === 'Star' && shape && Number(shapeLane) >= 1, JSON.stringify({ tiles, drawn, starName, shape, shapeLane }));

    if (process.env.RUTBA_VERIFY_CAPTURE) {
      try {
        fs.mkdirSync(process.env.RUTBA_VERIFY_CAPTURE, { recursive: true });
        await wait(500);
        fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'video-library-shapes.png'), (await win.webContents.capturePage()).toPNG());
      } catch { /* a capture is a picture to look at, not a check */ }
    }

    // Search: a word narrows the tiles.
    await js(`(() => { const n = document.querySelector('.vd-lib-search'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, 'arrow'); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await wait(200);
    const found = await js(`document.querySelectorAll('.vd-lib-shape-tile').length`);
    await js(`(() => { const n = document.querySelector('.vd-lib-search'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, ''); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    check('video library: a search narrows what is shown', found > 0 && found < tiles, JSON.stringify({ tiles, found }));

    // Icons and a title.
    await tab('icons');
    await js(`document.querySelector('.vd-lib-icon-tile[data-key="heart"]')?.click(), 1`);
    const icon = await ok(until(async () => (await count('icon')) === 1, 'the heart on the timeline', 4000));
    const iconTitle = await ok(until(() => js(`(document.querySelector('.vd-inspector strong')?.textContent || '') === 'Heart'`), 'the heart picked', 3000));
    const colour = await js(`Boolean(document.querySelector('.vd-graphic-colour .vd-swatch'))`);
    await tab('titles');
    const before = await count('text');
    await js(`document.querySelector('.vd-lib-title[data-key="lowerThird"]')?.click(), 1`);
    const title = await ok(until(async () => (await count('text')) > before, 'the lower third', 4000));
    check('video library: an icon is added and picked with its colours, and a title is added from the library',
      icon && iconTitle && colour && title, JSON.stringify({ icon, iconTitle, colour, title }));

    // Word art: styles drawn by the studio's painter, one added and its words typed.
    const arts = await js(`document.querySelectorAll('.vd-lib-wordart-tile').length`);
    const artDrawn = await js(`(() => { const c = document.querySelector('.vd-lib-wordart-tile[data-key="gold"] canvas'); if (!c) return false; const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data; let lit = 0; for (let i = 0; i < d.length; i += 4) if (d[i] + d[i + 1] + d[i + 2] > 300) lit++; return lit > 30; })()`);
    await js(`document.querySelector('.vd-lib-wordart-tile[data-key="gold"]')?.click(), 1`);
    const art = await ok(until(async () => (await count('wordart')) === 1, 'the word art on the timeline', 4000));
    await until(() => js(`Boolean(document.querySelector('.vd-wordart-props textarea'))`), 'its words', 3000).catch(() => {});
    await js(`(() => { const n = document.querySelector('.vd-wordart-props textarea'); Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(n, 'Grand opening'); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    const typed = await ok(until(() => js(`(document.querySelector('.vd-clip[data-kind="wordart"]')?.textContent || '').includes('Grand opening')`), 'the words on the clip', 3000));
    check('video library: word art styles are drawn as they look, and one is added with its words to type',
      arts === 7 && artDrawn && art && typed, JSON.stringify({ arts, artDrawn, art, typed }));

    // A sound effect: written once beside the project (here, where recordings go) and put at the playhead.
    await tab('sounds');
    const effects = await js(`document.querySelectorAll('.vd-lib-sound').length`);
    const sounds = await count('sound');
    await js(`document.querySelector('.vd-lib-sound[data-key="chime"] .vd-lib-add-sound')?.click(), 1`);
    const added = await ok(until(async () => (await count('sound')) > sounds, 'the chime on the timeline', 8000));
    const wav = path.join(dir, 'Rutba Recordings', 'Chime.wav');
    const written = fs.existsSync(wav) && fs.readFileSync(wav).subarray(0, 4).toString() === 'RIFF';
    check('video library: a sound effect made on this computer is written once as a WAV and put on the timeline',
      effects === 10 && added && written, JSON.stringify({ effects, added, written }));

    // The panels fold to a rail and open again, and stay as they were left.
    await js(`document.querySelector('.vd-fold-library')?.click(), 1`);
    const libFolded = await ok(until(() => js(`!document.querySelector('.vd-library') && Boolean(document.querySelector('.vd-open-library'))`), 'the library folded', 3000));
    await js(`document.querySelector('.vd-fold-properties')?.click(), 1`);
    const propsFolded = await ok(until(() => js(`!document.querySelector('.vd-inspector') && Boolean(document.querySelector('.vd-open-properties'))`), 'the properties folded', 3000));
    const wider = await js(`document.querySelector('.vd-monitor')?.getBoundingClientRect().width || 0`);
    await js(`document.querySelector('.vd-open-library')?.click(), 1`);
    await js(`document.querySelector('.vd-open-properties')?.click(), 1`);
    const back = await ok(until(() => js(`Boolean(document.querySelector('.vd-library')) && Boolean(document.querySelector('.vd-inspector'))`), 'both panels back', 3000));
    const kept = await js(`document.querySelector('.vd-library')?.dataset.tab`);
    const narrower = await js(`document.querySelector('.vd-monitor')?.getBoundingClientRect().width || 0`);
    check('video library: the library and the properties fold to a rail, giving the preview the room, and open again where they were',
      libFolded && propsFolded && back && kept === 'sounds' && wider > narrower + 200, JSON.stringify({ libFolded, propsFolded, back, kept, wider, narrower }));

    // The preview on another screen: a window of its own drawing the same frame, following the playhead and playing with it.
    await js(`[...document.querySelectorAll('.rw-tab')].find((n) => n.textContent.trim() === 'View')?.click(), 1`);
    const before2 = new Set(BrowserWindow.getAllWindows().map((w) => w.id));
    const pressMonitor = () => js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || n.title || n.textContent || '').trim().startsWith('Preview on Another Screen')); if (!b) return 'no button'; b.click(); return 'pressed'; })()`);
    const pressed = await pressMonitor();
    let mon = null;
    await until(() => { mon = BrowserWindow.getAllWindows().find((w) => !before2.has(w.id) && /monitor=/.test(w.webContents.getURL())) || null; return Boolean(mon); }, 'the monitor window', 8000).catch(() => {});
    const mjs = (code) => mon.webContents.executeJavaScript(code);
    const clips = await js(`document.querySelectorAll('.vd-clip').length`);
    const same = mon ? await ok(until(async () => Number(await mjs(`document.querySelector('.vd-monitor-window')?.dataset.clips || 0`)) === clips, 'the project on the monitor', 8000)) : false;
    await js(`(() => { const s = document.querySelector('.vd-scrub'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(s, '1.2'); s.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    const follows = mon ? await ok(until(async () => Math.abs(Number(await mjs(`document.querySelector('.vd-monitor-window')?.dataset.time`)) - 1.2) < 0.02, 'the monitor at 1.2 s', 4000)) : false;
    await js(`document.querySelector('.vd-play')?.click(), 1`);
    const plays = mon ? await ok(until(async () => Number(await mjs(`document.querySelector('.vd-monitor-window')?.dataset.time`)) > 1.6, 'the monitor playing', 5000)) : false;
    await js(`document.querySelector('.vd-play')?.click(), 1`);
    await wait(300);
    const here = Number(await js(`document.querySelector('.vd-scrub').value`));
    const there = mon ? Number(await mjs(`document.querySelector('.vd-monitor-window')?.dataset.time`)) : -1;
    const painted = mon ? await mjs(`(() => { const c = document.querySelector('.vd-monitor-canvas'); if (!c || !c.width) return false; const d = c.getContext('2d').getImageData(c.width / 2, c.height / 2, 1, 1).data; return d[0] + d[1] + d[2] > 30; })()`) : false;
    await pressMonitor();
    const closed = mon ? await ok(until(() => mon.isDestroyed(), 'the monitor closed', 5000)) : false;
    check('video library: Preview on Another Screen opens the picture in a window of its own that draws the project, follows the playhead, plays with it and closes with the button',
      pressed === 'pressed' && Boolean(mon) && same && follows && plays && Math.abs(here - there) < 0.1 && painted && closed,
      JSON.stringify({ pressed, opened: Boolean(mon), clips, same, follows, plays, here, there, painted, closed }));

    const complaints = await errorsIn(win);
    check('video library: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('video library: the check ran to the end', false, err?.message || String(err));
  } finally {
    app.setPath('videos', videosWas);
  }
}

/**
 * Video starters: one laid out on its stand-in pictures, drawn into a folder
 * beside where recordings go, a stand-in replaced by a picture of the
 * person's, and a starter asked for in a window with work in it opening a
 * window of its own. Run alone with RUTBA_VERIFY_ONLY=video-starters.
 */
export async function verifyVideoStarters(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  const videosWas = app.getPath('videos');
  const home = path.join(dir, 'starters-videos');
  fs.mkdirSync(home, { recursive: true });
  app.setPath('videos', home);
  try {
    const win = await open('video');
    const js = (code) => win.webContents.executeJavaScript(code);
    const fixtures = await makeFixtures(win.webContents, dir, 'video-starters');
    const capture = async (name) => {
      if (!process.env.RUTBA_VERIFY_CAPTURE) return;
      try {
        fs.mkdirSync(process.env.RUTBA_VERIFY_CAPTURE, { recursive: true });
        await wait(500);
        fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await win.webContents.capturePage()).toPNG());
      } catch { /* a capture is a picture to look at, not a check */ }
    };
    const offered = await ok(until(() => js(`document.querySelectorAll('.vd-starter').length === 4`), 'the starters', 6000));
    await capture('video-starters-offered.png');
    await js(`document.querySelector('.vd-starter[data-key="slideshow"]')?.click(), 1`);
    const laid = await ok(until(() => js(`document.querySelectorAll('.vd-clip[data-kind="image"]').length === 6`), 'the slideshow', 15000));
    await capture('video-starters-slideshow.png');
    const shape = await js(`({ texts: document.querySelectorAll('.vd-clip[data-kind="text"]').length, fades: [...document.querySelectorAll('.vd-clip[data-kind="image"]')].filter((n) => n.dataset.transition === 'fade').length })`);
    const folder = path.join(home, 'Rutba Recordings', 'Slideshow');
    const drawn = fs.existsSync(folder) ? fs.readdirSync(folder).filter((f) => /^Picture \d+\.png$/.test(f)).length : 0;
    const painted = await js(`(() => { const c = document.querySelector('.vd-canvas'); if (!c || !c.width) return false; const d = c.getContext('2d').getImageData(c.width / 2, c.height * 0.2, 1, 1).data; return d[0] + d[1] + d[2] > 60; })()`);
    check('video starters: Slideshow lays out six stand-in pictures drawn on this computer, crossfading, with a title and an end card',
      offered && laid && shape.texts >= 2 && shape.fades === 5 && drawn === 6 && painted, JSON.stringify({ offered, laid, shape, drawn, painted }));

    // Replace Media: the second stand-in becomes a picture of the person's, in the same place.
    const second = await js(`(() => { const n = [...document.querySelectorAll('.vd-clip[data-kind="image"]')].sort((a, b) => Number(a.dataset.start) - Number(b.dataset.start))[1]; if (!n) return null; const r = n.getBoundingClientRect(); const o = { bubbles: true, cancelable: true, button: 0, clientX: r.left + Math.min(12, r.width / 2), clientY: r.top + r.height / 2, pointerId: 1 }; n.dispatchEvent(new PointerEvent('pointerdown', o)); window.dispatchEvent(new PointerEvent('pointerup', o)); return { id: n.dataset.id, start: n.dataset.start, end: n.dataset.end }; })()`);
    await js(`[...document.querySelectorAll('.rw-tab')].find((n) => n.textContent.trim() === 'Clip')?.click(), 1`);
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'open', answer: [fixtures?.still] });
    const pressed = await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || n.title || n.textContent || '').trim().startsWith('Replace Media') && !n.disabled); if (!b) return 'no button'; b.click(); return 'pressed'; })()`);
    const replaced = second ? await ok(until(() => js(`(document.querySelector('.vd-clip[data-id="${second.id}"]')?.textContent || '').includes('still.png')`), 'the picture replaced', 6000)) : false;
    const kept = second ? await js(`(() => { const n = document.querySelector('.vd-clip[data-id="${second.id}"]'); return n ? [n.dataset.start, n.dataset.end] : null; })()`) : null;
    check('video starters: Replace Media puts a picture of the person\'s where a stand-in was, for the same time',
      Boolean(fixtures) && pressed === 'pressed' && replaced && JSON.stringify(kept) === JSON.stringify([second?.start, second?.end]), JSON.stringify({ pressed, replaced, second, kept }));

    // A starter asked for in a window with work in it opens in a window of its own, upright for Vertical Social.
    await js(`[...document.querySelectorAll('.rw-tab')].find((n) => n.textContent.trim() === 'Home')?.click(), 1`);
    const before = new Set(BrowserWindow.getAllWindows().map((w) => w.id));
    await js(`document.querySelector('.rw-appmark')?.click(), 1`);
    await until(() => js(`[...document.querySelectorAll('.rw-menu button')].some((b) => b.textContent.trim() === 'New Vertical Social Video')`), 'the starter in the menu', 4000).catch(() => {});
    await js(`[...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === 'New Vertical Social Video')?.click(), 1`);
    let other = null;
    await until(() => { other = BrowserWindow.getAllWindows().find((w) => !before.has(w.id) && /starter=social/.test(w.webContents.getURL())) || null; return Boolean(other); }, 'the social window', 8000).catch(() => {});
    const upright = other ? await ok(until(() => other.webContents.executeJavaScript(`document.querySelectorAll('.vd-clip[data-kind="image"]').length === 4 && (() => { const c = document.querySelector('.vd-canvas'); return c && c.height > c.width; })()`), 'the upright starter', 15000)) : false;
    check('video starters: a starter asked for in a window with work in it opens in a window of its own; Vertical Social is upright',
      Boolean(other) && upright, JSON.stringify({ opened: Boolean(other), upright }));
    if (other && !other.isDestroyed()) other.destroy();

    const complaints = await errorsIn(win);
    check('video starters: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('video starters: the check ran to the end', false, err?.message || String(err));
  } finally {
    app.setPath('videos', videosWas);
  }
}
