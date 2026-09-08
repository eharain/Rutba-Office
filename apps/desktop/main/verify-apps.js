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

/**
 * Wait for something to become true, rather than for a number of milliseconds.
 *
 * Every flaky check in this file has been a fixed pause that was long enough on
 * the machine it was written on. Polling for the condition is both faster when
 * it happens quickly and honest when it does not: the failure names what never
 * became true instead of describing whatever the state happened to be.
 */
async function until(condition, what, timeout = 8000) {
  const deadline = Date.now() + timeout;
  for (;;) {
    let held = false;
    try {
      held = await condition();
    } catch {
      held = false;
    }
    if (held) return true;
    if (Date.now() > deadline) throw new Error(`waited ${timeout} ms for ${what} and it never happened`);
    await wait(80);
  }
}

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

const FENCE = String.fromCharCode(96, 96, 96);

const README = [
  '---',
  'title: Rutba Office',
  '---',
  '',
  '# Rutba Office',
  '',
  'A free office suite. See the [documentation](https://office.rutba.io).',
  '',
  '## Status',
  '',
  '- [x] Mail, with **tracker blocking**',
  '- [ ] Translations',
  '- Word, Worksheets and *Presentation*',
  '',
  '| Format | Read | Write |',
  '| :----- | :--: | ----: |',
  '| .docx  | yes  | yes   |',
  '| .pst   | yes  | no    |',
  '',
  '> Nothing leaves your computer.',
  '',
  FENCE + 'js',
  "office.open('report.docx');",
  FENCE,
  '',
].join(String.fromCharCode(10));

/** Where two files stop agreeing, with enough either side to see why. */
function firstDifference(a, b) {
  const at = [...a].findIndex((c, i) => c !== b[i]);
  if (at < 0) return 'one file is longer than the other';
  const window = (s) => JSON.stringify(s.slice(Math.max(0, at - 40), at + 40));
  return `at ${at}: ${window(a)} became ${window(b)}`;
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

  // A README with the parts that usually get lost: a task list, a fenced
  // block, a table with alignment, links, and front matter.
  fs.writeFileSync(at('readme.md'), README);

  // A real image: the application's own icon, which is a genuine PNG.
  const icon = path.resolve(path.dirname(new URL(import.meta.url).pathname.replace(/^\//, '')), '..', 'resources', 'icon.png');
  if (fs.existsSync(icon)) fs.copyFileSync(icon, at('picture.png'));

  return {
    docx: at('report.docx'),
    xlsx: at('sales.xlsx'),
    pptx: at('deck.pptx'),
    wav: at('tone.wav'),
    md: at('readme.md'),
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
    // A window that throws during render paints nothing and reports nothing, so
    // every check against it fails with a description of an empty page rather
    // than of the fault. The console is the only place the fault appears.
    win.webContents.on('console-message', (_event, level, text) => {
      if (level >= 2) console.log(`     [${app}] ${text.split('\n')[0].slice(0, 200)}`);
    });
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

    // Focused, and confirmed focused. Asking for focus and then typing 250 ms
    // later is a race: on a busy machine the keystrokes arrive before the grid
    // has it, land nowhere, and the check fails describing a save that was
    // never asked to happen.
    await until(async () => {
      win.focus();
      win.webContents.focus();
      return win.webContents.executeJavaScript(
        `(() => { const g = document.querySelector('.sh'); if (!g) return false; g.focus(); return document.activeElement === g; })()`
      );
    }, 'the grid to take focus');

    // Move to an empty cell and type into it.
    //
    // Synthesised keystrokes are delivered to the window, not to the widget,
    // and a window that loses focus for one frame in the middle of a sequence
    // swallows the rest of it silently. The check is still that typing reaches
    // the document — that is the whole point of it — so it types again rather
    // than waiting longer. Three attempts that all produce nothing is a real
    // break; one that does not is the operating system.
    const typed = () => (doc.model({ id: sessionFor('sheet').id }).cells || []).some((c) => String(c.text) === '99');
    let attempts = 0;
    while (!typed() && attempts < 3) {
      attempts++;
      win.focus();
      win.webContents.focus();
      await win.webContents.executeJavaScript(`document.querySelector('.sh')?.focus(), 'ok'`);
      for (let i = 0; i < 4; i++) await press(win.webContents, 'Down');
      await typeText(win.webContents, '99');
      await press(win.webContents, 'Return', { char: true });
      try {
        await until(typed, 'the value to reach the engine', 2500);
      } catch {
        // Try once more from the top, with focus taken again.
      }
    }
    await until(typed, `the value to reach the engine after ${attempts} attempts`, 1000);

    // Waited for, not guessed at. A fixed pause is a check that passes on a
    // quiet machine and fails on a busy one, which teaches everybody to ignore
    // it — so this waits for the thing it is actually waiting for.
    const wasSaved = fs.statSync(files.xlsx).mtimeMs;
    await press(win.webContents, 's', { modifiers: ['control'] });
    await until(() => fs.statSync(files.xlsx).mtimeMs !== wasSaved, 'the file to be written');

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

    // The media element decodes on its own schedule and the timeline is drawn
    // from what it reports, so both are waited for rather than slept through.
    await until(
      () => win.webContents.executeJavaScript(`(() => { const v = document.querySelector('video'); return Boolean(v && v.readyState >= 1 && isFinite(v.duration)); })()`),
      'the media to report a duration'
    );
    const media = await win.webContents.executeJavaScript(`(() => {
      const v = document.querySelector('video');
      return v ? { duration: v.duration, ready: v.readyState, err: v.error ? v.error.code : null } : null;
    })()`);
    check(
      'video: the media loads and reports a duration',
      Boolean(media && Number.isFinite(media.duration) && media.duration > 0),
      media ? `duration ${media.duration}, readyState ${media.ready}${media.err ? `, error ${media.err}` : ''}` : 'no media element'
    );

    await until(
      () => win.webContents.executeJavaScript(`document.querySelectorAll('.vd-clip').length > 0`),
      'the timeline to be built'
    );
    const clips = await win.webContents.executeJavaScript(`document.querySelectorAll('.vd-clip').length`);
    check('video: a timeline is built from it', clips === 1, `${clips} clips`);
  } catch (err) {
    check('video: the checks ran', false, err.message);
  }

  /* ── Presentation: a deck can gain a slide, and a slide can gain notes ─ */

  try {
    const win = await open('slides', files.pptx);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = doc.open({ path: files.pptx });
    const run = async (...ops) => {
      await doc.apply({ id: session.id, ops });
      return doc.model({ id: session.id });
    };

    const before = doc.model({ id: session.id }).count;
    const after = (await run({ op: 'insertSlide', after: 0, layout: 'obj', title: 'Inserted', body: ['One'] })).count;
    check('slides: a deck can gain a slide', after === before + 1, `${before} slides became ${after}`);

    await run({ op: 'setNotes', slide: 1, text: 'Mention the tracker report.' });
    const withNotes = doc.model({ id: session.id, slide: 1 });
    check('slides: a slide can carry speaker notes', /tracker report/.test(withNotes.slide?.notes || ''), JSON.stringify(withNotes.slide?.notes || ''));

    // And both have to survive the file, since a notes part that PowerPoint
    // refuses is worse than no notes at all.
    const target = path.join(dir, 'deck-out.pptx');
    doc.save({ id: session.id, path: target });
    const again = doc.open({ path: target, slide: 1 });
    check('slides: the new slide survives the file', again.model.count === after, `${again.model.count} slides read back`);
    check('slides: the notes survive the file', /tracker report/.test(again.model.slide?.notes || ''), JSON.stringify(again.model.slide?.notes || ''));
    doc.close({ id: again.id });
    doc.close({ id: session.id });

    const tabs = await js(`[...document.querySelectorAll('.rw-ribbon-tabs button, .rw-tab')].map((b) => b.textContent.trim()).join(', ')`);
    check(
      'slides: the ribbon has the tabs a presentation has',
      ['Home', 'Insert', 'Design', 'Slide Show', 'View'].every((t) => tabs.includes(t)),
      `tabs are ${JSON.stringify(tabs)}`
    );

    // Presenter view: a second window on the same open document, showing the
    // speaker's side. It has to find the notes that were just written.
    const deck = doc.open({ path: files.pptx });
    doc.apply({ id: deck.id, ops: [{ op: 'setNotes', slide: 0, text: 'Open with the tracker report.' }] });

    const presenter = windows.create({ app: 'slides', query: { presenter: deck.id } });
    opened.push(presenter);
    await new Promise((resolve) => presenter.webContents.once('did-finish-load', () => setTimeout(resolve, 1400)));

    const speaker = await presenter.webContents.executeJavaScript(`(() => ({
      notes: document.querySelector('.pv-notes-text')?.textContent ?? '',
      hasNext: Boolean(document.querySelector('.pv-next .pv-thumb')),
      clock: document.querySelector('.pv-clock')?.textContent ?? '',
      stage: Boolean(document.querySelector('.pv-stage svg')),
    }))()`);

    check('slides: presenter view shows the speaker notes', /tracker report/.test(speaker.notes), JSON.stringify(speaker.notes.slice(0, 40)));
    check('slides: presenter view shows the next slide and a clock', speaker.hasNext && speaker.stage && /\d/.test(speaker.clock), JSON.stringify(speaker));
    doc.close({ id: deck.id });
  } catch (err) {
    check('slides: the deck checks ran', false, err.message);
  }

  /* ── The launcher survives a notice board that is not there ──────────── */

  try {
    const win = await open('home');
    // Long enough for the fetch to have failed and the render to have settled.
    await wait(2200);
    const state = await win.webContents.executeJavaScript(`(() => ({
      notices: document.querySelectorAll('.home-notice').length,
      complaints: [...document.querySelectorAll('.rw-toast.bad')].map((n) => n.textContent),
      apps: document.querySelectorAll('.home-card').length,
    }))()`);

    // The endpoint may or may not exist yet. Either way the launcher must be
    // whole, and a missing notice board must be silent — not a toast, not an
    // empty strip, not a gap where one would go.
    check('home: a notice board that is unreachable says nothing', state.complaints.length === 0, state.complaints.join(' | ') || 'nothing reported');
    check('home: the launcher is whole regardless', state.apps === 7, `${state.apps} app cards, ${state.notices} notices`);
  } catch (err) {
    check('home: the launcher checks ran', false, err.message);
  }

  /* ── Worksheets: the ribbon can actually format a cell ───────────────── */

  try {
    const win = await open('sheets', files.xlsx);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = sessionFor('sheet');

    // Every one of these was in the engine and had no button. A spreadsheet
    // whose cells cannot be made bold is not a spreadsheet, and the gap was
    // invisible because nothing checked for it.
    const press = async (title) =>
      js(`(() => {
        const b = [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || '') === ${JSON.stringify(title)});
        if (!b) return 'no button';
        b.click();
        return 'clicked';
      })()`);

    await js(`document.querySelector('.sh')?.focus(), 'ok'`);
    await wait(200);

    // Driven the way the window drives it — through the operation table — so
    // the check covers the wiring as well as the engine.
    // `apply` answers with a patch when the change is small, so the model is
    // read back explicitly rather than fished out of the reply.
    const run = async (...ops) => {
      await doc.apply({ id: session.id, ops });
      return doc.model({ id: session.id });
    };
    const bolded = await press('Bold');
    await wait(600);
    const afterBold = await doc.viewport({ id: session.id });
    check('sheets: the ribbon can make a cell bold', afterBold.format?.bold === true, `${bolded}, the selection reads bold=${afterBold.format?.bold}`);

    // Freeze panes, protection and named ranges: three more that existed only
    // in the engine until this ribbon.
    const frozen = await run({ op: 'freeze', rows: 1, cols: 0 });
    check('sheets: panes freeze', frozen.frozen?.rows === 1, JSON.stringify(frozen.frozen));

    const locked = await run({ op: 'protect' });
    check('sheets: a sheet can be protected', locked.protection?.sheet === true, JSON.stringify(locked.protection));
    await run({ op: 'unprotect' });

    const named = await run({ op: 'defineName', name: 'Revenue', ref: 'Sales!$B$2:$B$3' });
    check('sheets: a range can be named', (named.names || []).some((n) => n.name === 'Revenue'), `${(named.names || []).length} names defined`);

    const tabs = await js(`[...document.querySelectorAll('.rw-ribbon-tabs button, .rw-tab')].map((b) => b.textContent.trim()).join(', ')`);
    check(
      'sheets: the ribbon has the tabs a spreadsheet has',
      ['Home', 'Insert', 'Formulas', 'Data', 'Review', 'View'].every((t) => tabs.includes(t)),
      `tabs are ${JSON.stringify(tabs)}`
    );
  } catch (err) {
    check('sheets: the ribbon checks ran', false, err.message);
  }

  /* ── Every window offers full screen ─────────────────────────────────── */

  try {
    const win = opened[0];
    const found = await win.webContents.executeJavaScript(
      `[...document.querySelectorAll('.rw-wincontrols button, .rw-titlebar button')].some((b) => /full screen/i.test(b.title || ''))`
    );
    check('every window has a full-screen button', found === true, found ? 'in the title bar' : 'none found in the title bar');
  } catch (err) {
    check('the full-screen check ran', false, err.message);
  }

  /* ── Word: a GitHub README opens, edits and saves as Markdown ────────── */

  try {
    const win = await open('word', files.md);
    await wait(400);

    const shown = await win.webContents.executeJavaScript(`(() => {
      const blocks = [...document.querySelectorAll('.wd-page [data-block]')];
      return {
        count: blocks.length,
        text: blocks.map((b) => b.textContent).join('\\n'),
        headings: blocks.filter((b) => /^h[1-6]$/i.test(b.tagName) || /heading/i.test(b.className)).length,
        links: document.querySelectorAll('.wd-page a, .wd-page [data-link]').length,
      };
    })()`);

    // The structure a README is made of has to be visible, not flattened.
    check('word: a task list survives into the document', /☑|☐/.test(shown.text), `${shown.count} blocks drawn`);
    check('word: a fenced code block keeps its fences', shown.text.includes('```'), 'the fence markers are on screen');
    check('word: a table is a table', shown.text.includes('Format') && shown.text.includes('.pst'), 'table cells are present');

    // Save it straight back and compare with what went in.
    const before = fs.readFileSync(files.md, 'utf8');
    const session = sessionFor('doc');
    doc.export({ id: session.id, format: 'md', path: files.md });
    const after = fs.readFileSync(files.md, 'utf8');

    check(
      'word: saving a README does not rewrite it',
      after.trim() === before.trim(),
      after.trim() === before.trim() ? 'byte for byte' : firstDifference(before, after)
    );
  } catch (err) {
    check('word: the Markdown round trip ran', false, err.message);
  }

  /* ── Word: the ribbon reaches what the engine can do ─────────────────── */

  try {
    const win = await open('word', files.docx);
    const js = (code) => win.webContents.executeJavaScript(code);
    // Its own session, opened by path. Picking "the last document session"
    // finds whichever window happened to open most recently, which is a
    // different document every time a check is added above this one — and the
    // checks then quietly measure the wrong file.
    const session = doc.open({ path: files.docx });
    // `apply` may answer with a patch rather than a whole model, so the model
    // is read back explicitly instead of being fished out of the reply.
    const run = async (...ops) => {
      await doc.apply({ id: session.id, ops });
      return doc.model({ id: session.id });
    };

    const tabs = await js(`[...document.querySelectorAll('.rw-ribbon-tabs button, .rw-tab')].map((b) => b.textContent.trim()).join(', ')`);
    check(
      'word: the ribbon has the tabs a word processor has',
      ['Home', 'Insert', 'Layout', 'References', 'Review', 'View'].every((t) => tabs.includes(t)),
      `tabs are ${JSON.stringify(tabs)}`
    );

    // A link, a comment and a footer: three capabilities that were in the
    // engine with nothing calling them, and are now one press each.
    await run({ op: 'setSelection', anchor: { block: 1, offset: 0 }, focus: { block: 1, offset: 8 } });
    const linked = await run({ op: 'setLink', url: 'https://office.rutba.io' });
    const hasLink = (linked.blocks || []).some((b) => (b.runs || []).some((r) => r.link));
    check('word: text can be made a link', hasLink, hasLink ? 'the run carries its target' : 'no run came back with a link');

    const commented = await run({ op: 'addComment', text: 'Check this figure', author: 'Verify' });
    check('word: a comment can be added', (commented.comments || []).length > 0, `${(commented.comments || []).length} comments on the document`);

    const footed = await run({ op: 'setBand', band: 'footer', lines: ['{PAGE} of {PAGES}'] });
    const footer = footed.bands?.footers?.default;
    check('word: a footer can be set', Boolean(footer), footer ? `holds ${JSON.stringify(footer.paragraphs?.[0]?.text ?? '')}` : 'no footer came back');

    // And all three have to survive being written to disk and read again.
    const target = path.join(dir, 'ribbon.docx');
    doc.save({ id: session.id, path: target });
    const again = doc.open({ path: target });
    const back = again.model;
    const keptLink = (back.blocks || []).some((b) => (b.runs || []).some((r) => r.link));
    check('word: a link survives the file', keptLink, keptLink ? 'read back from the .docx' : 'the hyperlink relationship was lost');
    doc.close({ id: again.id });
    doc.close({ id: session.id });
  } catch (err) {
    check('word: the ribbon checks ran', false, err.message);
  }

  /* ── Word: the buttons a person clicks do what they say ──────────────── */
  //
  // Everything above drives the engine through operations. These drive the
  // window the way a hand does: select words on the page, press a button on
  // the ribbon, and read what the page paints. The first ribbon shipped with
  // every font, colour, indent and style control sending a key the engine did
  // not listen for, and nothing here noticed — because nothing here clicked.

  try {
    const win = await open('word', files.docx);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = doc.open({ path: files.docx });
    doc.close({ id: session.id }); // only needed the path to resolve; the window has its own
    const state = () => js(`(async () => {
      const all = await window.rutbaOffice.doc.sessions({});
      const mine = all.filter((s) => s.kind === 'doc').pop();
      const m = await window.rutbaOffice.doc.model({ id: mine.id });
      // A table reaches the frame as its cell paragraphs, each carrying the
      // address of its cell; there is no "table block". Counting distinct
      // table addresses is counting tables.
      const tables = new Set(m.blocks.map((b) => (b.container || '').split(':')[0]).filter(Boolean)).size;
      const drawn = document.querySelectorAll('.wd-page table.wd-table').length;
      return { format: m.format, block1: m.blocks[1], section: m.section, tables, drawn, blocksTotal: m.blocks.length };
    })()`);

    // Select the whole of the second paragraph with a real DOM range, and tell
    // the page about it the way a mouse would: with a mouseup.
    const selectBlock1 = () => js(`(() => {
      const page = document.querySelector('.wd-page');
      const block = page.querySelector('[data-block="1"]');
      page.focus();
      const range = document.createRange();
      range.selectNodeContents(block);
      const sel = getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      page.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
      return sel.toString().length;
    })()`);
    const pressButton = (title) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || '').startsWith(${JSON.stringify(title)}));
      if (!b) return 'no button titled ' + ${JSON.stringify(title)};
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      b.click();
      return 'clicked';
    })()`);
    const pressMenu = (label) => js(`(() => {
      const item = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(label)});
      if (!item) return 'no menu item ' + ${JSON.stringify(label)};
      item.click();
      return 'clicked';
    })()`);
    const choose = (title, value) => js(`(() => {
      const s = [...document.querySelectorAll('.rw-ribbon select')].find((n) => n.title === ${JSON.stringify(title)});
      if (!s) return 'no select titled ' + ${JSON.stringify(title)};
      const setter = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set;
      setter.call(s, ${JSON.stringify(value)});
      s.dispatchEvent(new Event('change', { bubbles: true }));
      return 'chose';
    })()`);
    const painted = () => js(`(() => {
      const span = document.querySelector('.wd-page [data-block="1"] span');
      const cs = span ? getComputedStyle(span) : null;
      const p = document.querySelector('.wd-page [data-block="1"]');
      return cs ? { weight: cs.fontWeight, size: cs.fontSize, colour: cs.color, align: getComputedStyle(p).textAlign } : null;
    })()`);

    const selected = await selectBlock1();
    await wait(300);
    check('word: words can be selected on the page', selected > 10, `${selected} characters selected`);

    // Bold, from the button. Focus must stay on the page for this to land.
    const b = await pressButton('Bold');
    await until(async () => (await state()).block1?.runs?.some((r) => r.bold), 'the engine to hold bold', 4000).catch(() => {});
    const afterBold = await state();
    const paintBold = await painted();
    check('word: the Bold button makes the selection bold', afterBold.block1?.runs?.some((r) => r.bold) && paintBold?.weight === '700', `${b}; engine bold=${afterBold.block1?.runs?.some((r) => r.bold)}, page weight=${paintBold?.weight}`);

    // Font size, from the dropdown, and it has to paint at that size.
    await selectBlock1();
    await wait(200);
    await choose('Size', '18');
    await until(async () => (await state()).format?.fontSize === 18, 'the engine to hold 18 pt', 4000).catch(() => {});
    const afterSize = await state();
    const paintSize = await painted();
    check('word: the size dropdown changes the font size', afterSize.format?.fontSize === 18 && paintSize?.size === '24px', `engine=${afterSize.format?.fontSize} pt, page=${paintSize?.size} (18 pt is 24 px)`);

    // Colour, from the menu.
    await selectBlock1();
    await wait(200);
    await pressButton('Text colour');
    await wait(200);
    const red = await pressMenu('Red');
    await until(async () => (await state()).format?.fontColour === 'E03131', 'the engine to hold the colour', 4000).catch(() => {});
    const paintColour = await painted();
    check('word: the colour menu colours the text', paintColour?.colour === 'rgb(224, 49, 49)', `${red}; page colour=${paintColour?.colour}`);

    // Alignment, list and style act on the paragraph the caret is in.
    const centred = await pressButton('Centre');
    await until(async () => (await state()).format?.paragraphAlign === 'center', 'the paragraph to centre', 4000).catch(() => {});
    const afterCentre = await state();
    const paintAlign = await painted();
    check('word: the Centre button centres the paragraph', paintAlign?.align === 'center', `${centred}; engine paragraphAlign=${afterCentre.format?.paragraphAlign}, block align=${JSON.stringify(afterCentre.block1?.align ?? null)}, page align=${paintAlign?.align}`);

    await pressButton('Bulleted list');
    await until(async () => (await state()).format?.listType === 'bullet', 'the paragraph to become a list item', 4000).catch(() => {});
    const afterList = await state();
    check('word: the list button makes a list', afterList.format?.listType === 'bullet', `listType=${afterList.format?.listType}`);

    await choose('Paragraph style', 'Heading1');
    await until(async () => (await state()).format?.paragraphStyle === 'Heading1', 'the style to apply', 4000).catch(() => {});
    const afterStyle = await state();
    check('word: the style dropdown applies a style', afterStyle.format?.paragraphStyle === 'Heading1', `paragraphStyle=${afterStyle.format?.paragraphStyle}`);

    // Insert a table through the dialog, and turn the page on its side.
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click(), 'insert tab'`);
    await wait(200);
    await pressButton('Table');
    await wait(300);
    await js(`(() => { [...document.querySelectorAll('.rw-dialog .rw-btn')].find((b) => b.textContent.trim() === 'Insert')?.click(); return 'inserted'; })()`);
    await until(async () => (await state()).tables > 0, 'a table to appear', 4000).catch(() => {});
    const afterTable = await state();
    check('word: Insert → Table puts a table in the document, and the page draws it as one', afterTable.tables > 0 && afterTable.drawn > 0, `${afterTable.tables} table(s) in the document, ${afterTable.drawn} drawn on the page`);

    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Layout')?.click(), 'layout tab'`);
    await wait(200);
    await pressButton('Switch orientation');
    await until(async () => (await state()).section?.orientation === 'landscape', 'the page to turn', 4000).catch(() => {});
    const afterTurn = await state();
    check('word: Layout → orientation turns the page', afterTurn.section?.orientation === 'landscape', `orientation=${afterTurn.section?.orientation}`);

    // A picture, with the bytes a file dialog would have handed over. The
    // dialog itself is the operating system's and cannot be pressed from here;
    // everything after it can.
    if (files.png) {
      const before = (await state()).blocksTotal;
      const inserted = await js(`(async () => {
        const all = await window.rutbaOffice.doc.sessions({});
        const mine = all.filter((s) => s.kind === 'doc').pop();
        const { bytes, stat } = await window.rutbaOffice.fs.read({ path: ${JSON.stringify(files.png)} });
        // Out of the table the previous step left the caret in: the engine
        // refuses a picture inside a cell, and says so, which is correct.
        await window.rutbaOffice.doc.apply({ id: mine.id, ops: [
          { op: 'setSelection', anchor: { block: 0, offset: 0 }, focus: { block: 0, offset: 0 } },
          { op: 'insertImage', name: stat.name, contentType: 'image/png', data: bytes, widthPx: 320, heightPx: 320 },
        ] });
        return 'inserted';
      })()`);
      await until(async () => (await state()).blocksTotal > before, 'the picture to land', 4000).catch(() => {});
      // The insert went in behind the window's back, so the page has not heard.
      // Undo and redo, from the keyboard, put it through the window's own path
      // — and prove that a picture survives both.
      await js(`document.querySelector('.wd-page')?.focus(), 'focused'`);
      await press(win.webContents, 'z', { modifiers: ['control'] });
      await wait(400);
      await press(win.webContents, 'y', { modifiers: ['control'] });
      await until(() => js(`document.querySelectorAll('.wd-page img').length > 0`), 'the picture to be drawn', 4000).catch(() => {});
      const pictured = await state();
      const drawnImage = await js(`document.querySelectorAll('.wd-page img').length`);
      check('word: a picture inserts as its own block and is drawn', pictured.blocksTotal > before && drawnImage > 0, `${inserted}; ${before} → ${pictured.blocksTotal} blocks, ${drawnImage} image(s) on the page`);
    }

    const complaints = await errorsIn(win);
    check('word: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the button checks ran', false, err.message);
  }

  /* ── Worksheets: clicking empty grid, typing, and formatting that paints ── */

  try {
    const win = await open('sheets', files.xlsx);
    const js = (code) => win.webContents.executeJavaScript(code);
    const model = () => js(`(async () => {
      const all = await window.rutbaOffice.doc.sessions({});
      const mine = all.filter((s) => s.kind === 'sheet').pop();
      return window.rutbaOffice.doc.model({ id: mine.id });
    })()`);

    // A press on empty grid, at the coordinates of C7 — a cell the engine has
    // not drawn because there is nothing in it.
    const pressEmpty = (row, col) => js(`(() => {
      const layer = document.querySelector('.sh-cells');
      const rect = layer.getBoundingClientRect();
      const grid = ${JSON.stringify({ row, col })};
      const cols = [...document.querySelectorAll('.sh-colheads .sh-head')];
      const rows = [...document.querySelectorAll('.sh-rowheads .sh-head')];
      const c = cols[grid.col]; const r = rows[grid.row];
      if (!c || !r) return 'headers missing';
      const x = rect.left + c.offsetLeft + c.offsetWidth / 2;
      const y = rect.top + r.offsetTop + r.offsetHeight / 2;
      const hit = document.elementFromPoint(x, y);
      hit.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, clientX: x, clientY: y, button: 0 }));
      return hit.className;
    })()`);

    const hit = await pressEmpty(6, 2);
    await until(async () => { const m = await model(); return m.selection?.active?.row === 6 && m.selection?.active?.col === 2; }, 'the click to select C7', 4000).catch(() => {});
    const afterClick = await model();
    check('sheets: clicking empty grid selects that cell', afterClick.selection?.active?.row === 6 && afterClick.selection?.active?.col === 2, `pressed on ${JSON.stringify(hit)}; active is ${afterClick.selection?.ref || JSON.stringify(afterClick.selection?.active)}`);

    // Type into it. The grid has to have focus for keys to land, and a click
    // on the layer is what gives it.
    await js(`document.querySelector('.sh')?.focus(), 'ok'`);
    await typeText(win.webContents, '77');
    await press(win.webContents, 'Return', { char: true });
    await until(async () => (await model()).cells?.some((c) => c.ref === 'C7' && c.text === '77'), 'the value to land in C7', 4000).catch(() => {});
    const typed = await model();
    check('sheets: typing after the click fills that cell', typed.cells?.some((c) => c.ref === 'C7' && c.text === '77'), `C7 is ${JSON.stringify(typed.cells?.find((c) => c.ref === 'C7')?.text ?? null)}`);

    // Back onto C7, then Bold, then a fill, then borders — and each must PAINT.
    await pressEmpty(6, 2);
    await wait(250);
    const clickRibbon = (title) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || '') === ${JSON.stringify(title)});
      if (!b) return 'no button ' + ${JSON.stringify(title)};
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      b.click();
      return 'clicked';
    })()`);
    const activeStyle = () => js(`(() => {
      const el = document.querySelector('.sh-cell.active');
      if (!el) return null;
      const cs = getComputedStyle(el);
      return { weight: cs.fontWeight, background: cs.backgroundColor, borderTop: cs.borderTopWidth, text: el.textContent };
    })()`);

    await clickRibbon('Bold');
    await until(async () => (await activeStyle())?.weight === '700', 'the cell to paint bold', 4000).catch(() => {});
    const bold = await activeStyle();
    check('sheets: the Bold button paints the cell bold', bold?.weight === '700' && bold?.text === '77', `active cell weight=${bold?.weight}, text=${JSON.stringify(bold?.text)}`);

    await clickRibbon('Fill colour');
    await wait(200);
    await js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === 'Light amber')?.click(); return 'chose'; })()`);
    await until(async () => (await activeStyle())?.background === 'rgb(255, 243, 191)', 'the fill to paint', 4000).catch(() => {});
    const filled = await activeStyle();
    check('sheets: the fill menu paints the cell', filled?.background === 'rgb(255, 243, 191)', `background=${filled?.background}`);

    await clickRibbon('Borders');
    await wait(200);
    await js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === 'All borders')?.click(); return 'chose'; })()`);
    await until(async () => (await activeStyle())?.borderTop === '1px', 'the border to paint', 4000).catch(() => {});
    const bordered = await activeStyle();
    check('sheets: the border menu paints a border', bordered?.borderTop === '1px', `border-top=${bordered?.borderTop}`);

    // A dialog from the ribbon, through to its effect: View → Freeze panes →
    // "Freeze the top row", and the frame has to say row 1 is frozen.
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'View')?.click(), 'view tab'`);
    await wait(150);
    await clickRibbon('Freeze panes');
    await wait(300);
    const chosen = await js(`(() => {
      const item = [...document.querySelectorAll('.rw-dialog .ml-found-item')].find((b) => /Freeze the top row/.test(b.textContent));
      if (!item) return 'no such choice';
      item.click();
      return 'chose';
    })()`);
    await until(async () => (await model()).frozen?.rows === 1, 'the top row to freeze', 4000).catch(() => {});
    const frozen = (await model()).frozen;
    check('sheets: the Freeze panes dialog freezes the top row', frozen?.rows === 1, `${chosen}; frozen=${JSON.stringify(frozen)}`);

    const complaints = await errorsIn(win);
    check('sheets: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('sheets: the click checks ran', false, err.message);
  }

  /* ── Presentation, Pictures, Image, Video, Mail: pressed, not driven ──── */

  const clickIn = (win, title) => win.webContents.executeJavaScript(`(() => {
    const b = [...document.querySelectorAll('.rw-btn, .ml-compose-cta button')].find((n) => (n.title || n.textContent || '').trim().startsWith(${JSON.stringify(title)}) && !n.disabled);
    if (!b) return 'no button ' + ${JSON.stringify(title)};
    b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
    b.click();
    return 'clicked';
  })()`);
  const clickMenu = (win, label) => win.webContents.executeJavaScript(`(() => {
    const item = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(label)});
    if (!item) return 'no menu item ' + ${JSON.stringify(label)};
    item.click();
    return 'clicked';
  })()`);
  const clickTab = (win, label) => win.webContents.executeJavaScript(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(label)})?.click(), 'tab'`);
  const setField = (win, selector, value) => win.webContents.executeJavaScript(`(() => {
    const el = document.querySelector(${JSON.stringify(selector)});
    if (!el) return 'no field ' + ${JSON.stringify(selector)};
    const proto = el.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, ${JSON.stringify(value)});
    el.dispatchEvent(new Event('input', { bubbles: true }));
    return 'set';
  })()`);

  try {
    const win = await open('slides', files.pptx);
    const deckModel = (slide) => win.webContents.executeJavaScript(`(async () => {
      const all = await window.rutbaOffice.doc.sessions({});
      const mine = all.filter((s) => s.kind === 'deck').pop();
      const active = [...document.querySelectorAll('.sl-thumb')].findIndex((t) => t.classList.contains('active'));
      return window.rutbaOffice.doc.model({ id: mine.id, slide: ${slide === undefined ? 'Math.max(0, active)' : slide}, width: 640 });
    })()`);

    const before = (await deckModel()).count;
    await clickIn(win, 'New slide');
    await wait(200);
    await clickMenu(win, 'Title and content');
    await until(async () => (await deckModel()).count === before + 1, 'the deck to gain a slide', 4000).catch(() => {});
    const grew = await deckModel();
    const thumbs = await win.webContents.executeJavaScript(`document.querySelectorAll('.sl-thumb').length`);
    check('slides: the New slide menu adds a slide', grew.count === before + 1 && thumbs === before + 1, `${before} → ${grew.count} slides, ${thumbs} thumbnails`);

    await clickTab(win, 'Insert');
    await wait(150);
    const shapesBefore = (await deckModel()).slide?.shapes?.length ?? 0;
    await clickIn(win, 'Text box');
    await until(async () => ((await deckModel()).slide?.shapes?.length ?? 0) > shapesBefore, 'a text box to appear', 4000).catch(() => {});
    const shapesAfter = (await deckModel()).slide?.shapes?.length ?? 0;
    check('slides: Insert → Text box puts a text box on the slide', shapesAfter === shapesBefore + 1, `${shapesBefore} → ${shapesAfter} shapes`);

    await clickTab(win, 'Home');
    await wait(150);
    await clickIn(win, 'Speaker notes');
    await wait(250);
    await setField(win, '.rw-dialog textarea', 'Say hello first.');
    await win.webContents.executeJavaScript(`[...document.querySelectorAll('.rw-dialog .rw-btn')].find((b) => b.textContent.trim() === 'Save')?.click(), 'saved'`);
    await until(async () => /hello first/.test((await deckModel()).slide?.notes || ''), 'the notes to save', 4000).catch(() => {});
    const noted = await deckModel();
    check('slides: the Speaker notes dialog saves notes', /hello first/.test(noted.slide?.notes || ''), JSON.stringify(noted.slide?.notes || ''));

    await clickIn(win, 'Delete');
    await until(async () => (await deckModel()).count === before, 'the slide to go', 4000).catch(() => {});
    const shrank = await deckModel();
    check('slides: the Delete button removes the slide', shrank.count === before, `${grew.count} → ${shrank.count} slides`);

    const complaints = await errorsIn(win);
    check('slides: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('slides: the button checks ran', false, err.message);
  }

  try {
    const win = await open('pictures', files.png);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.pv-image')) && /\\d+ \\/ \\d+/.test(document.querySelector('.pv-counter')?.textContent || '')`), 'the picture and its counter', 6000);
    const counter = await js(`document.querySelector('.pv-counter').textContent`);

    await clickIn(win, 'Rotate');
    await until(() => js(`/rotate\\(90deg\\)/.test(document.querySelector('.pv-image')?.style.transform || '')`), 'the view to rotate', 4000).catch(() => {});
    const turned = await js(`document.querySelector('.pv-image')?.style.transform || ''`);
    check('pictures: the Rotate button turns the view', /rotate\(90deg\)/.test(turned), `transform is ${JSON.stringify(turned)}`);

    // The fixture folder holds two files, so Next has somewhere to go.
    await clickIn(win, 'Next');
    await until(() => js(`document.querySelector('.pv-counter').textContent !== ${JSON.stringify(counter)}`), 'the counter to advance', 4000).catch(() => {});
    const advanced = await js(`document.querySelector('.pv-counter').textContent`);
    check('pictures: the Next button moves to the next file', advanced !== counter, `${counter} → ${advanced}`);

    const complaints = await errorsIn(win);
    check('pictures: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('pictures: the button checks ran', false, err.message);
  }

  try {
    const win = await open('image', files.png);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`(document.querySelector('.im-canvas')?.width || 0) > 0`), 'the picture to reach the canvas', 6000);
    const ops = () => js(`document.querySelectorAll('.im-op').length`);
    const start = await ops();
    await clickIn(win, 'Right');
    await until(async () => (await ops()) === start + 1, 'a rotate to be recorded', 4000).catch(() => {});
    await clickIn(win, 'Flip');
    await until(async () => (await ops()) === start + 2, 'a flip to be recorded', 4000).catch(() => {});
    const two = await ops();
    check('image: Rotate and Flip each record an edit', two === start + 2, `${start} → ${two} edits`);
    await clickIn(win, 'Reset');
    await until(async () => (await ops()) === 0, 'the edits to clear', 4000).catch(() => {});
    check('image: Reset clears them', (await ops()) === 0, `${await ops()} edits left`);
    const complaints = await errorsIn(win);
    check('image: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('image: the button checks ran', false, err.message);
  }

  try {
    const win = await open('video', files.wav);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.vd-clip').length > 0`), 'the timeline', 8000);
    // Move the playhead to the middle with the scrubber, then split there.
    await js(`(() => {
      const s = document.querySelector('.vd-scrub');
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(s, String(Number(s.max || 1) / 2));
      s.dispatchEvent(new Event('input', { bubbles: true }));
      s.dispatchEvent(new Event('change', { bubbles: true }));
      return s.value;
    })()`);
    await wait(300);
    const scrubbed = await js(`document.querySelector('.vd-time')?.textContent ?? ''`);
    // Split lives on the Edit tab. A check that looked for it on Home found
    // nothing and said "1 clips", which is the wrong sentence for "no button".
    await clickTab(win, 'Edit');
    await wait(150);
    const pressed = await clickIn(win, 'Split');
    await until(() => js(`document.querySelectorAll('.vd-clip').length === 2`), 'the clip to split in two', 4000).catch(() => {});
    const clips = await js(`document.querySelectorAll('.vd-clip').length`);
    check('video: Split at the playhead makes two clips', clips === 2, `${pressed}; playhead at ${scrubbed}; ${clips} clips`);
    const complaints = await errorsIn(win);
    check('video: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('video: the button checks ran', false, err.message);
  }

  /* ── Mail: a seeded message opens in the reading pane ────────────────── */

  try {
    const win = await open('mail');
    await wait(1400);
    const js = (code) => win.webContents.executeJavaScript(code);
    const rows = await js(`document.querySelectorAll('.ml-row').length`);

    if (!rows) {
      check('mail: there are messages to read', false, 'the seeded store produced no rows');
    } else {
      // The seed contains one conversation of four messages, three of which
      // are chained and one of which only matches by subject. If the count
      // badge does not say 4, threading has regressed one way or the other.
      const thread = await js(`(() => {
        const badge = [...document.querySelectorAll('.ml-count')].map((b) => Number(b.textContent)).sort((a, b) => b - a)[0];
        return badge ?? 0;
      })()`);
      check('mail: replies collapse into one conversation', thread === 4, `largest conversation holds ${thread} messages`);

      // Open the newsletter, which is where the whole privacy story lives.
      const opened = await js(`(() => {
        const row = [...document.querySelectorAll('.ml-row')].find((r) => /pricing right/i.test(r.textContent));
        if (!row) return 'not found';
        row.click();
        return 'clicked';
      })()`);
      await wait(1100);

      const subject = await js(`document.querySelector('.ml-head h2')?.textContent ?? ''`);
      check('mail: a message opens in the reading pane', subject.trim().length > 0, `subject is ${JSON.stringify(subject.trim().slice(0, 40))} (${opened})`);

      const framed = await js(`(() => { const f = document.querySelector('.ml-body iframe'); return f ? f.getAttribute('sandbox') : null; })()`);
      check('mail: the body is framed with no privileges', framed === '', `sandbox is ${JSON.stringify(framed)}`);

      const strip = await js(`document.querySelector('.ml-strip')?.textContent ?? ''`);
      // Exactly two: the open pixel and the analytics beacon. The sender's own
      // 180-pixel logo is remote too, and is not a tracker — counting it would
      // be the kind of scaremongering that teaches people to ignore the badge.
      check('mail: trackers are counted, and a logo is not one', /\b2 trackers blocked/.test(strip), `strip reads ${JSON.stringify(strip.slice(0, 70))}`);
      check('mail: the tracking networks are named', /Mailchimp|Google Analytics/.test(strip), `strip names ${JSON.stringify(strip.slice(0, 70))}`);
      check('mail: leaving the list is offered', /unsubscribe/i.test(strip), /unsubscribe/i.test(strip) ? 'one-click unsubscribe offered' : 'no offer on a message that carries the header');

      // The forged message must be visibly forged.
      await js(`(() => {
        const row = [...document.querySelectorAll('.ml-row')].find((r) => /account has been limited/i.test(r.textContent));
        row?.click();
        return 'clicked';
      })()`);
      await wait(900);
      const bad = await js(`document.querySelector('.ml-strip')?.textContent ?? ''`);
      check('mail: a failed sender check is shown', /failed/i.test(bad), `strip reads ${JSON.stringify(bad.slice(0, 70))}`);

      // Selecting rows must offer the actions that only make sense in bulk.
      await js(`(() => {
        const box = document.querySelector('.ml-row .ml-check');
        box.click();
        return 'checked';
      })()`);
      await wait(400);
      const bulk = await js(`document.querySelector('.ml-bulk')?.textContent ?? ''`);
      check('mail: selecting rows opens the bulk actions', /selected/.test(bulk), `bar reads ${JSON.stringify(bulk.slice(0, 40))}`);

      // Gmail and Outlook.com no longer take a password. Typing one of those
      // addresses has to offer the browser rather than a password box that
      // will fail at the server — the worst possible order to find out in.
      await js(`(() => {
        const item = [...document.querySelectorAll('.rw-item')].find((i) => /Add account/.test(i.textContent));
        item?.click();
        return 'opened';
      })()`);
      await wait(700);
      await js(`(() => {
        const input = document.querySelector('.rw-dialog input');
        const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
        setter.call(input, 'someone@gmail.com');
        input.dispatchEvent(new Event('input', { bubbles: true }));
        return 'typed';
      })()`);
      await wait(900);
      const signIn = await js(`document.querySelector('.rw-dialog .ml-found')?.textContent ?? ''`);
      check('mail: a Gmail address is offered a sign-in, not a password box', /Sign in with Google/.test(signIn), `the dialog says ${JSON.stringify(signIn.slice(0, 60))}`);

      await js(`(() => { [...document.querySelectorAll('.rw-dialog .rw-btn')].find((b) => /Cancel/.test(b.textContent))?.click(); return 'closed'; })()`);
      await wait(400);

      // Rules. The point of the preview is that it changes nothing, so that is
      // checked before the rule is allowed to run.
      const account = (await win.webContents.executeJavaScript(
        `(async () => (await window.rutbaOffice.mail.accounts({}))[0]?.id)()`
      ));
      const mail = (method, args = {}) =>
        win.webContents.executeJavaScript(`window.rutbaOffice.mail.${method}(${JSON.stringify(args)})`);

      // The seeded account's folder is named after the archive it came from,
      // not "Inbox" — asking the account rather than assuming is also what
      // the window does.
      const inbox = (await mail('folders', { accountId: account }))[0]?.path;

      const rule = {
        name: 'Newsletters',
        enabled: true,
        all: true,
        conditions: [{ field: 'subject', op: 'contains', value: 'pricing right' }],
        actions: [{ type: 'move', value: 'Reading' }],
      };

      const before = (await mail('messages', { accountId: account, folder: inbox, limit: 500 })).total;
      const dry = await mail('testRules', { accountId: account, folder: inbox, rules: [rule] });
      const stillThere = (await mail('messages', { accountId: account, folder: inbox, limit: 500 })).total;
      check(
        'mail: a rule can be tried without moving anything',
        dry.matched === 1 && stillThere === before,
        `${dry.matched} of ${dry.of} would move; the folder still holds ${stillThere}`
      );

      const ran = await mail('runRules', { accountId: account, folder: inbox, rules: [rule] });
      const after = (await mail('messages', { accountId: account, folder: 'Reading', limit: 500 })).total;
      check('mail: a rule files mail when it is run', ran.moved === 1 && after === 1, `moved ${ran.moved}; Reading now holds ${after}`);

      // Search has to find it in its new home, which is the index noticing that
      // a folder changed underneath it.
      const found = await mail('search', { accountId: account, query: 'subject:pricing', limit: 20 });
      check('mail: search follows a message that a rule moved', found.length === 1 && found[0].folder === 'Reading', `${found.length} hits, first in ${JSON.stringify(found[0]?.folder)}`);

      // Compose, from the big button, to the outbox, and back. Send never
      // goes straight out; it waits in the queue, and Undo takes it back and
      // reopens the message — which is checked before the queue could fire.
      await clickIn(win, 'Compose');
      await wait(400);
      await js(`(() => {
        const inputs = [...document.querySelectorAll('.rw-dialog input:not([disabled]):not([type=checkbox])')];
        const set = (el, v) => { Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, v); el.dispatchEvent(new Event('input', { bubbles: true })); };
        set(inputs[0], 'someone@example.com');
        set(inputs[1], 'Hello from the check');
        return inputs.length;
      })()`);
      await wait(200);
      await js(`[...document.querySelectorAll('.rw-dialog .rw-btn')].find((b) => b.textContent.trim() === 'Send')?.click(), 'sent'`);
      await until(async () => (await mail('outbox')).length === 1, 'the message to reach the outbox', 4000).catch(() => {});
      const queued = await mail('outbox');
      const banner = await js(`document.querySelector('.ml-outbox')?.textContent ?? ''`);
      check('mail: Send puts the message in the outbox with an undo', queued.length === 1 && queued[0].draft?.to === 'someone@example.com' && /Undo/.test(banner), `${queued.length} queued; banner reads ${JSON.stringify(banner.slice(0, 50))}`);

      await js(`document.querySelector('.ml-outbox button')?.click(), 'undo'`);
      await until(async () => (await mail('outbox')).length === 0, 'the message to come back', 4000).catch(() => {});
      const reopened = await js(`Boolean(document.querySelector('.rw-dialog')) && (document.querySelector('.rw-dialog input:not([disabled])')?.value || '')`);
      check('mail: Undo takes it back and reopens it', (await mail('outbox')).length === 0 && reopened === 'someone@example.com', `outbox ${(await mail('outbox')).length}; reopened to ${JSON.stringify(reopened)}`);
      await js(`[...document.querySelectorAll('.rw-dialog .rw-btn')].find((b) => b.textContent.trim() === 'Discard')?.click(), 'discarded'`);
      await wait(300);

      // The attachment view is a place, not a search.
      await js(`(() => {
        const item = [...document.querySelectorAll('.rw-item')].find((i) => /^Attachments/.test(i.textContent));
        item?.click();
        return 'opened';
      })()`);
      await wait(1200);
      const files = await js(`(() => ({
        cards: document.querySelectorAll('.ml-card').length,
        first: document.querySelector('.ml-card .name')?.textContent ?? '',
      }))()`);
      check('mail: every attachment is listed in one place', files.cards > 0, `${files.cards} files, first is ${JSON.stringify(files.first)}`);
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
