// Video's library and panels (the 2.7 plan's library and panels): the
// project's media kept to add again, titles, shapes and icons drawn as what
// they add, sound effects made on this computer, a search, and the library
// and the properties each folded to a rail and opened again. A clip is made
// here at run time. Run alone with RUTBA_VERIFY_ONLY=video-library.

import fs from 'node:fs';
import path from 'node:path';
import { app } from 'electron';
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

    const complaints = await errorsIn(win);
    check('video library: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('video library: the check ran to the end', false, err?.message || String(err));
  } finally {
    app.setPath('videos', videosWas);
  }
}
