// Presentations: Insert → Screen Recording, pressed on the ribbon — the
// picker lists the other windows (this check's own Documents window among
// them), the one picked is recorded with a bar along the top showing the
// time, Stop puts the recording on the slide as a video, and the saved
// deck keeps it. The recording goes by a scratch file, written a few
// megabytes a message, and the file is gone once the deck has it. Only the
// suite's own windows are named in what this check reports.

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { Deck, buildPptx } from '@rutba/presentation';

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
    const session = doc.sessions().filter((s) => s.kind === 'deck' && s.path === deckFile).pop();
    const videos = () => (doc.model({ id: session.id }).slide?.shapes || []).filter((s) => s.media?.kind === 'video');

    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 1`), 'the slide', 8000);
    const pressed = await js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Screen Recording');
      if (!b || b.disabled) return 'no button'; b.click(); return 'clicked';
    })()`);
    const listed = await until(() => js(`[...document.querySelectorAll('.ss-source')].some((n) => n.title.includes('rec-target'))`), 'our window in the picker', 8000).then(() => true).catch(() => false);
    const title = await js(`document.querySelector('.rw-dialog-head')?.textContent || ''`);
    await js(`(() => { const b = [...document.querySelectorAll('.ss-source')].find((n) => n.title.includes('rec-target')); b?.click(); return Boolean(b); })()`);
    const recording = await until(() => js(`(() => { const b = document.querySelector('.sl-screenrec-stop'); return Boolean(b) && !b.disabled; })()`), 'the recording to start', 8000).then(() => true).catch(() => false);
    await wait(1600);
    const shown = await js(`document.querySelector('.sl-screenrec-time')?.textContent || ''`);
    const writing = scratchLeft().length === 1;
    await js(`document.querySelector('.sl-screenrec-stop')?.click(), 1`);
    const placed = await until(() => videos().length === 1, 'the recording on the slide', 15000).then(() => true).catch(() => false);
    const bar = await js(`Boolean(document.querySelector('.sl-screenrec'))`);
    const left = scratchLeft().length;
    // Its poster is a picture from the file, not the plain one a window that cannot decode it shows.
    const pictured = !(await js(`document.body.textContent.includes('cannot show its picture')`));
    check('presentations: Insert → Screen Recording lists the other windows, records the one picked with the time showing, and Stop puts the recording on the slide as a video',
      pressed === 'clicked' && listed && /Screen Recording/.test(title) && recording && /^0:0[1-9]$/.test(shown) && placed && !bar,
      JSON.stringify({ pressed, listed, title, recording, shown, placed, bar }));
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
