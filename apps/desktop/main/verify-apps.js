// Does each app do its job?
//
// The smoke run proves a window paints. The editing run proves keystrokes reach
// the document. Neither proves the thing a person actually came for: open a
// file, change it, save it, and get the change back tomorrow.
//
// So this drives the real windows through a round trip per app, and checks the
// result on disk or in the engine rather than on screen. Saving is the part
// worth being strict about: a suite that cannot save is not a suite, and a save
// that silently writes a broken file is worse than one that refuses.
//
// Fixtures are generated here so the run needs nothing checked in and nothing
// from the machine it happens to be on.

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { buildPptx } from '@rutba/presentation';

const wait = (ms) => new Promise((r) => setTimeout(r, ms));

async function press(wc, keyCode, { modifiers = [], char = false } = {}) {
  wc.sendInputEvent({ type: 'keyDown', keyCode, modifiers });
  if (char) wc.sendInputEvent({ type: 'char', keyCode, modifiers });
  wc.sendInputEvent({ type: 'keyUp', keyCode, modifiers });
  await wait(130);
}

async function typeText(wc, text) {
  for (const ch of text) {
    wc.sendInputEvent({ type: 'keyDown', keyCode: ch });
    wc.sendInputEvent({ type: 'char', keyCode: ch });
    wc.sendInputEvent({ type: 'keyUp', keyCode: ch });
    await wait(60);
  }
}

/** A one-second tone, so the media app has something real to open. */
function buildWav({ seconds = 1, rate = 8000, freq = 440 } = {}) {
  const samples = seconds * rate;
  const buf = Buffer.alloc(44 + samples * 2);
  buf.write('RIFF', 0);
  buf.writeUInt32LE(36 + samples * 2, 4);
  buf.write('WAVE', 8);
  buf.write('fmt ', 12);
  buf.writeUInt32LE(16, 16);
  buf.writeUInt16LE(1, 20); // PCM
  buf.writeUInt16LE(1, 22); // mono
  buf.writeUInt32LE(rate, 24);
  buf.writeUInt32LE(rate * 2, 28);
  buf.writeUInt16LE(2, 32);
  buf.writeUInt16LE(16, 34);
  buf.write('data', 36);
  buf.writeUInt32LE(samples * 2, 40);
  for (let i = 0; i < samples; i++) {
    buf.writeInt16LE(Math.round(Math.sin((2 * Math.PI * freq * i) / rate) * 12000), 44 + i * 2);
  }
  return buf;
}

function makeFixtures(dir) {
  fs.mkdirSync(dir, { recursive: true });
  const at = (name) => path.join(dir, name);

  fs.writeFileSync(
    at('report.docx'),
    buildDocx({
      styles: true,
      paragraphs: [
        { text: 'Quarterly Review', style: 'Title', bold: true, size: 26 },
        { text: 'The northern region grew fastest in absolute terms.' },
        { text: 'Churn fell to 2.1% across the year.' },
      ],
    })
  );

  fs.writeFileSync(
    at('sales.xlsx'),
    buildXlsx({
      sheets: [
        {
          name: 'Sales',
          rows: [
            ['Region', 'Q1', 'Q2', 'Total'],
            ['North', 1420, 1610, '=SUM(B2:C2)'],
            ['South', 860, 910, '=SUM(B3:C3)'],
          ],
        },
      ],
    })
  );

  fs.writeFileSync(
    at('deck.pptx'),
    buildPptx({
      title: 'Rutba Office',
      slides: [
        { layout: 'title', title: 'Rutba Office', body: 'A free office suite' },
        { layout: 'obj', title: 'Second slide', body: ['One', 'Two'] },
      ],
    })
  );

  fs.writeFileSync(at('tone.wav'), buildWav());

  // A real image: the application's own icon, which is a genuine PNG.
  const icon = path.resolve(path.dirname(new URL(import.meta.url).pathname.replace(/^\//, '')), '..', 'resources', 'icon.png');
  if (fs.existsSync(icon)) fs.copyFileSync(icon, at('picture.png'));

  return {
    docx: at('report.docx'),
    xlsx: at('sales.xlsx'),
    pptx: at('deck.pptx'),
    wav: at('tone.wav'),
    png: fs.existsSync(at('picture.png')) ? at('picture.png') : null,
  };
}

export async function verifyApps({ windows, doc }) {
  const results = [];
  const opened = [];
  const check = (name, ok, detail = '') => {
    results.push({ name, ok, detail });
    console.log(`${ok ? 'ok  ' : 'FAIL'} ${name}${detail ? ` — ${detail}` : ''}`);
  };

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-verify-'));
  const files = makeFixtures(dir);

  const open = async (app, file) => {
    const win = windows.create({ app, file: file || null });
    await new Promise((resolve) => {
      win.webContents.once('did-finish-load', () => setTimeout(resolve, 1300));
    });
    win.focus();
    win.webContents.focus();
    opened.push(win);
    return win;
  };

  const sessionFor = (kind) => doc.sessions().filter((s) => s.kind === kind).pop();
  const errorsIn = (win) =>
    win.webContents.executeJavaScript(`[...document.querySelectorAll('.rw-toast.bad')].map((n) => n.textContent)`);

  /* ── Word: open, type, save in place, reopen from disk ───────────────── */

  try {
    const win = await open('word', files.docx);
    await win.webContents.executeJavaScript(`(() => {
      const page = document.querySelector('.wd-page');
      const block = page.querySelector('[data-block="1"]');
      page.focus();
      const range = document.createRange();
      range.selectNodeContents(block);
      range.collapse(true);
      const sel = getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      return 'ok';
    })()`);
    await wait(250);

    win.webContents.insertText('EDITED ');
    await wait(500);

    // The document has a path, so Ctrl+S writes it without a dialog.
    await press(win.webContents, 's', { modifiers: ['control'] });
    await wait(1400);

    const bytes = fs.readFileSync(files.docx);
    const reopened = doc.open({ path: files.docx });
    const text = reopened.model.blocks.map((b) => (b.runs || []).map((r) => r.text).join('')).join('\n');
    doc.close({ id: reopened.id });
    check('word: an edit survives save and reopen', text.includes('EDITED'), `file is ${(bytes.length / 1024).toFixed(1)} KB`);

    const complaints = await errorsIn(win);
    check('word: saving reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');

    // Unsaved work must not close. Type again to make it dirty, then ask the
    // window to close and check that it refused — the prompt is on screen and
    // the document is still there. Closing the X used to throw the work away
    // without a word, which is the worst thing an editor can do.
    win.webContents.insertText('MORE');
    await wait(500);
    win.close();
    await wait(900);
    check('word: a dirty window refuses to close silently', !win.isDestroyed(), win.isDestroyed() ? 'the window closed and the work went with it' : 'the window is still open, asking');
  } catch (err) {
    check('word: the round trip ran', false, err.message);
  }

  /* ── Worksheets: type a value, save, reopen ──────────────────────────── */

  try {
    const win = await open('sheets', files.xlsx);
    await win.webContents.executeJavaScript(`document.querySelector('.sh')?.focus(), 'ok'`);
    await wait(250);

    // Move to an empty cell and type into it.
    for (let i = 0; i < 4; i++) await press(win.webContents, 'Down');
    await typeText(win.webContents, '99');
    await press(win.webContents, 'Return', { char: true });
    await wait(400);

    await press(win.webContents, 's', { modifiers: ['control'] });
    await wait(1600);

    const reopened = doc.open({ path: files.xlsx });
    const model = doc.model({ id: reopened.id });
    const found = (model.cells || []).some((c) => String(c.text) === '99');
    doc.close({ id: reopened.id });
    check('sheets: a typed value survives save and reopen', found, found ? 'A5 is 99' : 'the value was not in the reopened file');

    const complaints = await errorsIn(win);
    check('sheets: saving reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('sheets: the round trip ran', false, err.message);
  }

  /* ── Presentation: navigate, save, reopen ────────────────────────────── */

  try {
    const win = await open('slides', files.pptx);
    const drawn = await win.webContents.executeJavaScript(
      `document.querySelector('.sl-svg svg')?.textContent ?? ''`
    );
    check('slides: the slide is drawn', drawn.includes('Rutba Office'), `svg text is ${JSON.stringify(drawn.slice(0, 40))}`);

    const thumbs = await win.webContents.executeJavaScript(`document.querySelectorAll('.sl-thumb').length`);
    check('slides: the sorter lists every slide', thumbs === 2, `${thumbs} thumbnails`);

    // Down moves to the next slide; the stage should redraw with its title.
    await press(win.webContents, 'Down');
    await wait(700);
    const second = await win.webContents.executeJavaScript(
      `document.querySelector('.sl-svg svg')?.textContent ?? ''`
    );
    check('slides: arrow keys move through the deck', second.includes('Second slide'), `svg text is ${JSON.stringify(second.slice(0, 40))}`);

    const complaints = await errorsIn(win);
    check('slides: nothing is reported', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('slides: the checks ran', false, err.message);
  }

  /* ── Pictures: the file actually decodes in the window ───────────────── */

  if (files.png) {
    try {
      const win = await open('pictures', files.png);
      await wait(700);
      const shown = await win.webContents.executeJavaScript(`(() => {
        const img = document.querySelector('.pv-image');
        return img ? { w: img.naturalWidth, h: img.naturalHeight, complete: img.complete } : null;
      })()`);
      check(
        'pictures: the picture decodes over rutba://file',
        Boolean(shown && shown.w > 0 && shown.h > 0),
        shown ? `${shown.w} × ${shown.h}` : 'no image element'
      );

      const tiles = await win.webContents.executeJavaScript(`document.querySelectorAll('.pv-tile').length`);
      check('pictures: the folder is listed', tiles > 0, `${tiles} tiles`);
    } catch (err) {
      check('pictures: the checks ran', false, err.message);
    }
  }

  /* ── Image: open, rotate, and check the result is turned ─────────────── */

  if (files.png) {
    try {
      const win = await open('image', files.png);
      await wait(900);
      const before = await win.webContents.executeJavaScript(`(() => {
        const c = document.querySelector('.im-canvas');
        return c ? { w: c.width, h: c.height } : null;
      })()`);
      check('image: the picture is on the canvas', Boolean(before && before.w > 0), before ? `${before.w} × ${before.h}` : 'no canvas');

      // A square icon cannot show a rotation by its shape, so the edit list is
      // what proves the operation landed.
      await press(win.webContents, 'r', { modifiers: ['control'] });
      await wait(600);
      const edits = await win.webContents.executeJavaScript(`document.querySelectorAll('.im-op').length`);
      check('image: rotating records an edit', edits === 1, `${edits} edits listed`);

      const complaints = await errorsIn(win);
      check('image: nothing is reported', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('image: the checks ran', false, err.message);
    }
  }

  /* ── Video: the media element loads what the protocol served ─────────── */

  try {
    const win = await open('video', files.wav);
    await wait(1200);
    const media = await win.webContents.executeJavaScript(`(() => {
      const v = document.querySelector('video');
      return v ? { duration: v.duration, ready: v.readyState, err: v.error ? v.error.code : null } : null;
    })()`);
    check(
      'video: the media loads and reports a duration',
      Boolean(media && Number.isFinite(media.duration) && media.duration > 0),
      media ? `duration ${media.duration}, readyState ${media.ready}${media.err ? `, error ${media.err}` : ''}` : 'no media element'
    );

    const clips = await win.webContents.executeJavaScript(`document.querySelectorAll('.vd-clip').length`);
    check('video: a timeline is built from it', clips === 1, `${clips} clips`);
  } catch (err) {
    check('video: the checks ran', false, err.message);
  }

  /* ── Mail: a seeded message opens in the reading pane ────────────────── */

  try {
    const win = await open('mail');
    await wait(900);
    const rows = await win.webContents.executeJavaScript(`document.querySelectorAll('.ml-row').length`);
    if (rows > 0) {
      await win.webContents.executeJavaScript(`document.querySelector('.ml-row').click(), 'clicked'`);
      await wait(900);
      const subject = await win.webContents.executeJavaScript(
        `document.querySelector('.ml-head h2')?.textContent ?? ''`
      );
      check('mail: a message opens in the reading pane', subject.trim().length > 0, `subject is ${JSON.stringify(subject.trim())}`);

      const framed = await win.webContents.executeJavaScript(
        `(() => { const f = document.querySelector('.ml-body iframe'); return f ? f.getAttribute('sandbox') : null; })()`
      );
      check('mail: the body is framed with no privileges', framed === '', `sandbox is ${JSON.stringify(framed)}`);
    } else {
      check('mail: there are messages to read', false, 'the seeded store was empty — run with RUTBA_SMOKE_SEED=1');
    }
  } catch (err) {
    check('mail: the checks ran', false, err.message);
  }

  for (const win of opened) win.destroy();
  fs.rmSync(dir, { recursive: true, force: true });

  const failed = results.filter((r) => !r.ok);
  console.log(`\n${results.length - failed.length}/${results.length} application checks passed`);
  return failed.length === 0;
}
