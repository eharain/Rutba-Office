// Presentations: Insert → Screen Recording, pressed on the ribbon, opens
// the Recorder for a slide (A7), with the screen and no camera: it is
// recorded with the time showing, Stop puts the recording on the
// slide as a video, and the saved deck keeps it. The recording goes by a scratch file, written a few
// megabytes a message, and the file is gone once the deck has it. Only the
// suite's own windows are named in what this check reports.

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { Deck, buildPptx } from '@rutba/presentation';
import { emptyStack } from '@rutba/studio/stack';
import { seedSetup, restoreSetup, recorderOpened, windowsNow, CHECK_SOUNDS, RECORDING, PRESS_RECORD } from './verify-studio-seed.js';

/**
 * @param {object} h the harness: open, check, until, wait, doc, sessionFor
 */
export async function verifyDeckScreenRecording(h, { dir }) {
  const { open, check, until, wait, doc } = h;
  const deckFile = path.join(dir, 'rec-host.pptx');
  const docFile = path.join(dir, 'rec-target.docx');
  const began = Date.now();
  // The scratch folders made since the check began and still there.
  const scratchLeft = () => {
    try {
      return fs.readdirSync(os.tmpdir()).filter((n) => n.startsWith('rutba-office-'))
        .filter((n) => { try { return fs.statSync(path.join(os.tmpdir(), n)).birthtimeMs >= began - 1000; } catch { return false; } });
    } catch { return []; }
  };
  try {
    fs.writeFileSync(deckFile, buildPptx({ title: 'Recording', slides: [{ layout: 'blank' }] }));
    fs.writeFileSync(docFile, buildDocx({ styles: true, paragraphs: [{ text: 'This window is recorded.' }] }));
    const deckWin = await open('slides', deckFile);
    await open('word', docFile);
    const js = (code) => deckWin.webContents.executeJavaScript(code);
    // The Recorder's settings for this: no countdown, no sound.
    fs.mkdirSync(path.join(dir, 'deck-recordings'), { recursive: true });
    const keptSetup = await seedSetup(js, { app: 'recorder', stack: { ...emptyStack(), sounds: CHECK_SOUNDS(false) }, studio: { size: 720, fps: 30, countdown: 0, folder: path.join(dir, 'deck-recordings'), minimise: true } });
    const session = doc.sessions().filter((s) => s.kind === 'deck' && s.path === deckFile).pop();
    const videos = () => (doc.model({ id: session.id }).slide?.shapes || []).filter((s) => s.media?.kind === 'video');

    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 1`), 'the slide', 8000);
    const known = windowsNow();
    const pressed = await js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Screen Recording');
      if (!b || b.disabled) return 'no button'; b.click(); return 'clicked';
    })()`);
    const { rw, rjs } = await recorderOpened(until, known);
    const title = await rjs(`document.querySelector('.st-into')?.textContent || ''`);
    const shows = await until(() => rjs(`/^(screen|window)$/.test(document.querySelector('.st-canvas')?.dataset.layers || '') && !document.querySelector('.st-record').disabled`), 'the screen live in the Recorder', 12000).then(() => true).catch(() => false);
    await rjs(PRESS_RECORD);
    const recording = await until(() => rjs(RECORDING), 'the recording to start', 8000).then(() => true).catch(() => false);
    await wait(1600);
    const shown = await rjs(`document.querySelector('.st-record')?.textContent || ''`);
    const writing = scratchLeft().length === 1;
    await rjs(`document.querySelector('.st-record')?.click(), 1`);
    const placed = await until(() => videos().length === 1, 'the recording on the slide', 15000).then(() => true).catch(() => false);
    const gone = await until(() => rw.isDestroyed(), 'the Recorder to close', 6000).then(() => true).catch(() => false);
    await restoreSetup(js, keptSetup);
    const left = scratchLeft().length;
    // Its poster is a picture from the file, not the plain one a window that cannot decode it shows.
    const pictured = !(await js(`document.body.textContent.includes('cannot show its picture')`));
    check('presentations: Insert → Screen Recording opens the Recorder for a slide with the screen and no camera, records with the time showing, and Stop puts the recording on the slide as a video',
      pressed === 'clicked' && /for a slide/.test(title) && shows && recording && /0:0[1-9]/.test(shown) && placed && gone,
      JSON.stringify({ pressed, title, shows, recording, shown, placed, gone }));
    check('presentations: a screen recording goes to the slide by a scratch file, its poster drawn from the file, and the file is gone once the slide has it',
      writing && placed && left === 0 && pictured, JSON.stringify({ writing, left, pictured }));

    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => {
      try {
        const deck = Deck.open(fs.readFileSync(deckFile));
        const v = deck.slide(0).shapes.find((s) => s.media?.kind === 'video');
        return Boolean(v) && /^ppt\/media\/media1\.(mp4|webm)$/.test(v.media.source.part) && deck.media(v.media.source.part).length > 1000;
      } catch { return false; }
    }, 'the saved recording', 8000).then(() => true).catch(() => false);
    const kept = (() => { try { return Deck.open(fs.readFileSync(deckFile)).slide(0).shapes.find((x) => x.media)?.media.source.part || null; } catch { return null; } })();
    check('presentations: the saved deck keeps the screen recording as a video', saved, saved ? `${kept} in the file` : 'not in the file');
  } catch (err) {
    check('presentations: the Screen Recording checks ran', false, err.message);
  }
}
