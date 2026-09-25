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
import { fileURLToPath } from 'node:url';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml/package';
import { buildPptx, Deck } from '@rutba/presentation';
import { consoleMessage } from './console-message.js';
import { gradientPng, joinPictureParagraphs } from './sample-picture.js';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { verifyViewer } from './verify-viewer.js';
import { verifySheetLinks } from './verify-sheet-links.js';
import { verifyDeckArrange } from './verify-deck-arrange.js';
import { verifyDeckFx } from './verify-deck-fx.js';
import { verifyWordToc } from './verify-word-toc.js';
import { verifyOutline } from './verify-outline.js';
import { verifyWordTrack } from './verify-word-track.js';
import { SheetView } from '@rutba/sheet-view';

const wait = (ms) => new Promise((r) => setTimeout(r, ms));


/**
 * Wait for something to become true, rather than for a number of milliseconds.
 *
 * Every flaky check in this file has been a fixed pause that was long enough on
 * the machine it was written on. Polling for the condition is both faster when
 * it happens quickly and honest when it does not: the failure names what never
 * became true instead of describing whatever the state happened to be.
 */
import { APPS } from '@rutba/office-formats/registry';

// Set once the checks start closing their own windows; a close before that is not theirs.
const closingPhase = { value: false };

/** What each app draws first; a window is ready when one of these — or the refusal — is on the page. */
const READY_IN = {
  home: '.home-card',
  word: '.wd-block',
  sheets: '.sh-cells',
  slides: '.sl-svg, .sl-thumb',
  pictures: '.pv-image, .pv-pdf, .pv-video, .pv-audio, .pv-tile, .pv-empty',
  image: '.im-canvas',
  video: '.vd-clip, .vd-empty',
  mail: '.ml-row, .ml-list, .ml-compose-cta',
};

/**
 * Bring a window's page to the front of the keyboard. Off the desktop the
 * window itself is never activated — that would take the keyboard from
 * whoever is working at the machine — and the page needs no activation to
 * take synthetic input.
 */
function raise(win) {
  if (process.env.RUTBA_WINDOW_DISPLAY !== 'offscreen') win.focus();
  win.webContents.focus();
}

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

  // A workbook with a note on a cell and a link on another — what Review
  // shows and Insert → Link makes. The note is a comments part as Excel
  // writes one; the link is a place on the second sheet.
  {
    const view = SheetView.open(buildXlsx({
      sheets: [
        { name: 'Sales', rows: [['Region', 'Q1', 'Q2'], ['North', 1420, 1610], ['South', 860, 910], ['See the ledger', '', '']] },
        { name: 'Ledger', rows: [['Ledger', 'Amount'], ['North', 1420]] },
      ],
    }));
    const sheetPart = view.workbook.partNameFor('Sales');
    view.pkg.addPart('xl/comments1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Kim Lee</author></authors><commentList><comment ref="B2" authorId="0"><text><r><t>Check this figure against the ledger.</t></r></text></comment></commentList></comments>', 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml');
    view.pkg.addRelationshipTo(sheetPart, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments', '../comments1.xml');
    view.setHyperlink({ row: 3, col: 0, location: 'Ledger!A2', tooltip: 'The ledger row' });
    fs.writeFileSync(at('notes.xlsx'), view.save());
  }

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

  // A report of several pages: paragraphs of up to a dozen lines, so some
  // must cross a page's bottom; headings, which keep with what follows; an
  // explicit page break; and a table of forty rows, taller than a page.
  const lorem = 'The northern region grew fastest in absolute terms, and the margin it opened in the spring held through the autumn despite two price changes and a supply interruption that took most of July to clear. ';
  const report = [{ text: 'The Long Report', style: 'Title' }];
  for (let i = 0; i < 30; i++) {
    if (i % 6 === 0) report.push({ text: `Section ${i / 6 + 1}`, style: 'Heading1' });
    report.push({ text: lorem.repeat(1 + ((i * 7) % 5)) + `(${i + 1})`, align: 'justify' });
  }
  report.push({ text: 'Starts a fresh page.', pageBreakBefore: true });
  report.push({ table: { rows: Array.from({ length: 40 }, (_, r) => [`Row ${r + 1}`, `Item ${r + 1}`, String((r + 1) * 12)]), header: true } });
  for (let i = 0; i < 12; i++) report.push({ text: lorem.repeat(1 + (i % 4)) });
  fs.writeFileSync(at('long.docx'), buildDocx({ styles: true, paragraphs: report }));

  fs.writeFileSync(at('tone.wav'), buildWav());

  // A README with the parts that usually get lost: a task list, a fenced
  // block, a table with alignment, links, and front matter.
  fs.writeFileSync(at('readme.md'), README);

  // Two cards and two events, the events placed round today so the month
  // and week views show them whenever the run happens.
  fs.writeFileSync(at('contacts.vcf'), ['BEGIN:VCARD', 'VERSION:3.0', 'FN:Kim Lee', 'N:Lee;Kim;;;', 'ORG:Tech Style Ltd', 'EMAIL;TYPE=WORK,PREF:kim@example.com', 'TEL;TYPE=CELL:+44 7700 900123', 'END:VCARD', 'BEGIN:VCARD', 'VERSION:3.0', 'FN:Sam Patel', 'N:Patel;Sam;;;', 'EMAIL;TYPE=WORK:sam@example.org', 'END:VCARD', ''].join('\r\n'));
  const day = new Date();
  const stamp = (d, h) => `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}T${String(h).padStart(2, '0')}0000`;
  // Both on today: an all-day event dated tomorrow fell into next week's band
  // whenever the run happened on a Sunday, and the week check failed one day in seven.
  fs.writeFileSync(at('events.ics'), ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Rutba//checks//EN', 'X-WR-CALNAME:Checks', 'BEGIN:VEVENT', 'UID:check-1@rutba.io', `DTSTART:${stamp(day, 10)}`, `DTEND:${stamp(day, 11)}`, 'SUMMARY:Pricing review', 'LOCATION:Room 4', 'END:VEVENT', 'BEGIN:VEVENT', 'UID:check-2@rutba.io', `DTSTART;VALUE=DATE:${stamp(day, 0).slice(0, 8)}`, 'SUMMARY:Bank holiday', 'END:VEVENT', 'END:VCALENDAR', ''].join('\r\n'));

  // A real image: the application's own icon, which is a genuine PNG.
  const icon = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', 'resources', 'icon.png');
  if (fs.existsSync(icon)) fs.copyFileSync(icon, at('picture.png'));

  // Two pictures of their own, a different colour each — what the
  // slideshow check needs to see the show move from one to the next. Made
  // the way the sample pictures elsewhere here are, rather than from the
  // application's icon above, so the check does not depend on that file
  // being where the icon is expected to be.
  fs.writeFileSync(at('show-a.png'), gradientPng(200, 160, [40, 120, 200], [230, 200, 60]));
  fs.writeFileSync(at('show-b.png'), gradientPng(200, 160, [180, 40, 90], [250, 220, 120]));

  // A picture floating at the right of a long paragraph, the words wrapping
  // round it — what a logo beside a letter's opening looks like in Word.
  if (fs.existsSync(icon)) {
    const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Floating pictures', style: 'Heading1' }, { text: lorem.repeat(3) }, { text: 'After the picture.' }] }));
    view.setSelection({ block: 0, offset: 0 });
    view.insertImage({ name: 'logo', contentType: 'image/png', data: fs.readFileSync(icon), widthPx: 160, heightPx: 160 });
    view.setImageLayout({ block: 1, image: 0, wrap: 'square', hAlign: 'right' });
    fs.writeFileSync(at('float.docx'), view.save());
  }

  // A picture whose paragraph lands at the foot of a page, and one taller
  // than a page — a scanned card in a letter (owner, 2026-09-20: "word image
  // across two pages").
  {
    const view = openDocx(buildDocx({ styles: true, paragraphs: [
      { text: 'Pictures at the foot', style: 'Heading1' },
      ...Array.from({ length: 26 }, (_, i) => ({ text: `Line ${i + 1} of the letter, short enough to stay on one line.` })),
      { text: 'After the pictures.' },
    ] }));
    view.setSelection({ block: 26, offset: 0 });
    view.insertImage({ name: 'card', contentType: 'image/png', data: gradientPng(200, 300, [60, 160, 90], [200, 50, 50]), widthPx: 400, heightPx: 600 });
    view.setSelection({ block: 28, offset: 0 });
    view.insertImage({ name: 'scan', contentType: 'image/png', data: gradientPng(300, 700, [50, 90, 200], [250, 200, 60]), widthPx: 600, heightPx: 1400 });
    fs.writeFileSync(at('fit.docx'), view.save());
  }

  // One paragraph of six scans of a card, two to a line, as a scanner's
  // software writes them and the owner's cards-print.docx holds them: three
  // picture lines, taller than a page, so the paragraph has to split between
  // two of them (2026-09-20, "it still is the same").
  {
    const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Cards' }] }));
    view.setSelection({ block: 0, offset: 0 });
    for (let i = 0; i < 6; i++) {
      view.insertImage({ name: `card ${i + 1}`, contentType: 'image/png', data: gradientPng(140, 225, [40 + i * 50, 120, 200 - i * 40], [230, 200 - i * 40, 60 + i * 40]), widthPx: 280, heightPx: 450 });
    }
    fs.writeFileSync(at('cards.docx'), joinPictureParagraphs(view.save()));
  }

  // A paragraph with a first-line indent, one with two tab stops, and a
  // table of three fixed columns — what the ruler and the grips move.
  {
    const view = openDocx(buildDocx({ styles: true, paragraphs: [
      { text: 'The Ruler', style: 'Heading1' },
      { text: lorem.repeat(2) },
      { text: 'Item\tPrice\tTotal' },
      { table: { rows: [['a', 'b', 'c'], ['d', 'e', 'f']], columns: [2400, 3000, 3600] } },
      { text: 'After the table.' },
    ] }));
    view.setSelection({ block: 1, offset: 0 });
    view.setParagraphFormat({ firstLineTwips: 720 });
    view.setSelection({ block: 2, offset: 0 });
    view.setParagraphFormat({ tabs: [{ align: 'left', posTwips: 1440 }, { align: 'right', posTwips: 5760 }] });
    fs.writeFileSync(at('ruler.docx'), view.save());
  }

  // A short report with three headings across two pages — References →
  // Table of Contents lists them, each with a page number the window reads
  // off the screen. "Method" is pushed onto a second page by the body text
  // under "Background", so the check has two different page numbers to see.
  {
    const view = openDocx(buildDocx({ styles: true, paragraphs: [
      { text: 'Annual Report', style: 'Title' },
      { text: 'Introduction', style: 'Heading1' },
      { text: lorem.repeat(2) },
      { text: 'Background', style: 'Heading2' },
      { text: lorem.repeat(22) },
      { text: 'Method', style: 'Heading1' },
      { text: lorem.repeat(2) },
    ] }));
    fs.writeFileSync(at('toc.docx'), view.save());
  }

  // Three plain paragraphs — Review → Track Changes records an edit here,
  // recording off at the start so the check turns it on itself.
  {
    const view = openDocx(buildDocx({ styles: true, paragraphs: [
      { text: 'First paragraph text' },
      { text: 'Second paragraph text' },
      { text: 'Third paragraph text' },
    ] }));
    fs.writeFileSync(at('track.docx'), view.save());
  }

  // A file of accounts, the loose way other clients write one.
  fs.writeFileSync(at('accounts.json'), JSON.stringify([
    { email: 'one@checks.example', password: 'not-a-real-password', host: 'mail.checks.example', port: 993, smtpPort: 587, smtpSecure: false, label: 'one' },
    { email: 'two@checks.example', password: 'not-a-real-password', imap: { host: 'imap.checks.example', port: 143, secure: false }, smtp: { host: 'smtp.checks.example', port: 465, secure: true }, name: 'Two' },
    { email: 'nothing@checks.example', host: 'mail.checks.example' },
  ], null, 2));

  return {
    accounts: at('accounts.json'),
    docx: at('report.docx'),
    long: at('long.docx'),
    float: fs.existsSync(at('float.docx')) ? at('float.docx') : null,
    ruler: at('ruler.docx'),
    fit: at('fit.docx'),
    cards: at('cards.docx'),
    toc: at('toc.docx'),
    track: at('track.docx'),
    xlsx: at('sales.xlsx'),
    notes: at('notes.xlsx'),
    pptx: at('deck.pptx'),
    wav: at('tone.wav'),
    vcf: at('contacts.vcf'),
    ics: at('events.ics'),
    md: at('readme.md'),
    png: fs.existsSync(at('picture.png')) ? at('picture.png') : null,
  };
}

export async function verifyApps({ windows, doc, broadcast = null, update = null }) {
  // An error that escapes a block would end the run with no summary and no
  // name. Name it.
  process.on('unhandledRejection', (e) => console.log(`     [unhandled] ${e?.stack || e}`));
  process.on('uncaughtException', (e) => console.log(`     [uncaught] ${e?.stack || e}`));

  const results = [];
  const opened = [];
  // If the application quits under the run, the run ends with no summary and
  // exit code 0 — which looks like success. Say so.
  import('electron').then(({ app }) => app.once('before-quit', () => console.log('     [before-quit] the application is quitting under the checks'))).catch(() => {});
  // Where the time goes. A timer that should fire every second prints when
  // the main process kept it waiting — a busy main process is what makes a
  // window's reply late and a check read too soon — and a clock line every
  // minute shows which block is slow.
  const started = Date.now();
  let lastCheck = '(before the first check)';
  let lastTick = Date.now();
  let lastClock = Date.now();
  const lagTimer = setInterval(() => {
    const now = Date.now();
    const late = now - lastTick - 1000;
    if (late > 400) console.log(`     [busy] the main process was ${late} ms late, after "${lastCheck}"`);
    if (now - lastClock >= 60000) {
      console.log(`     [clock] ${Math.round((now - started) / 1000)} s in, after "${lastCheck}"`);
      lastClock = now;
    }
    lastTick = now;
  }, 1000);

  // RUTBA_VERIFY_PROFILE=1: sample the main process while the checks run and
  // say where its time went. A late main process is what makes a window's
  // reply late and a check read too soon; the lag timer above says that it
  // was late, this says whether it was busy with our own code or simply
  // starved of the machine — the idle share tells the two apart.
  let profiler = null;
  if (process.env.RUTBA_VERIFY_PROFILE) {
    try {
      const { Session } = await import('node:inspector');
      const session = new Session();
      session.connect();
      const post = (method, params) => new Promise((resolve, reject) => session.post(method, params || {}, (err, result) => (err ? reject(err) : resolve(result))));
      await post('Profiler.enable');
      await post('Profiler.setSamplingInterval', { interval: 2000 });
      await post('Profiler.start');
      profiler = { post };
      console.log('     [profile] sampling the main process');
    } catch (err) {
      console.log(`     [profile] not available: ${err.message}`);
    }
  }
  const printProfile = async () => {
    if (!profiler) return;
    try {
      const { profile } = await profiler.post('Profiler.stop');
      const byId = new Map(profile.nodes.map((n) => [n.id, n]));
      const counts = new Map();
      for (const id of profile.samples) counts.set(id, (counts.get(id) || 0) + 1);
      const byFunction = new Map();
      const byFile = new Map();
      for (const [id, c] of counts) {
        const cf = byId.get(id)?.callFrame || {};
        const file = (cf.url || '').split(/[\\/]/).slice(-2).join('/') || '(native)';
        const key = `${cf.functionName || '(anonymous)'}  ${file}:${(cf.lineNumber ?? -1) + 1}`;
        byFunction.set(key, (byFunction.get(key) || 0) + c);
        byFile.set(file, (byFile.get(file) || 0) + c);
      }
      const total = profile.samples.length || 1;
      const seconds = Math.round((profile.endTime - profile.startTime) / 1e6);
      const pct = (c) => `${((100 * c) / total).toFixed(1).padStart(5)}%`;
      console.log(`     [profile] ${total} samples over ${seconds} s; the top functions by own time:`);
      for (const [k, c] of [...byFunction.entries()].sort((a, b) => b[1] - a[1]).slice(0, 24)) console.log(`     [profile] ${pct(c)}  ${k}`);
      console.log('     [profile] by file:');
      for (const [k, c] of [...byFile.entries()].sort((a, b) => b[1] - a[1]).slice(0, 14)) console.log(`     [profile] ${pct(c)}  ${k}`);
    } catch (err) {
      console.log(`     [profile] failed: ${err.message}`);
    }
  };

  const check = (name, ok, detail = '') => {
    results.push({ name, ok, detail });
    lastCheck = name;
    console.log(`${ok ? 'ok  ' : 'FAIL'} ${name}${detail ? ` — ${detail}` : ''}`);
  };

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-verify-'));
  const files = makeFixtures(dir);

  const open = async (app, file) => {
    const win = windows.create({ app, file: file || null });
    // A window that throws during render paints nothing and reports nothing, so
    // every check against it fails with a description of an empty page rather
    // than of the fault. The console is the only place the fault appears.
    win.webContents.on('console-message', (...args) => {
      const m = consoleMessage(args);
      if (m.level >= 2) console.log(`     [${app}] ${m.text.split('\n')[0].slice(0, 200)}`);
    });
    // A window that goes away mid-run takes its checks with it and, once the
    // last one goes, the whole run — say which one went, and when.
    const id = win.id;
    // A window closed before the checks close theirs was closed by someone
    // else — a person at the machine, or a crash — and the check that was
    // using it fails with "Object has been destroyed".
    win.on('closed', () => console.log(`     [closed] the ${app} window (#${id}) at ${new Date().toTimeString().slice(0, 8)}${closingPhase.value ? '' : ' — not by the checks'}`));
    win.webContents.on('render-process-gone', (_e, details) => console.log(`     [gone] the ${app} renderer: ${details.reason}`));
    await new Promise((resolve) => win.webContents.once('did-finish-load', resolve));
    // Ready when the app has drawn what it draws first — or said why it could
    // not. A fixed pause here was 1.3 seconds, which a loaded machine beat:
    // the grid layer was not there when the first press came.
    const ready = `${READY_IN[app] || '.rw-app'}, .rw-empty`;
    await until(() => win.webContents.executeJavaScript(`Boolean(document.querySelector(${JSON.stringify(ready)}))`), `the ${app} window to draw`, 12000).catch(() => {});
    await wait(300);
    raise(win);
    opened.push(win);
    return win;
  };

  const sessionFor = (kind) => doc.sessions().filter((s) => s.kind === kind).pop();
  const errorsIn = (win) =>
    win.webContents.executeJavaScript(`[...document.querySelectorAll('.rw-toast.bad')].map((n) => n.textContent)`);

  const done = async () => {
    clearInterval(lagTimer);
    await printProfile();
    closingPhase.value = true;
    // The windows are left to app.exit, which closes them without events.
    // Destroying ninety at once here, a moment before it, was followed by
    // Electron ending on 7014 whatever verdict the run handed it.
    try {
      fs.rmSync(dir, { recursive: true, force: true });
    } catch {
      // A window still holding a fixture open leaves the folder behind in
      // the temp directory; nothing reads it again.
    }

    const failed = results.filter((r) => !r.ok);
    console.log(`\n${results.length - failed.length}/${results.length} application checks passed`);
    return failed.length === 0;
  };

  /* ── Word: a long document lays out on pages ─────────────────────────── */
  //
  // The page used to be one endless sheet: a report ran off the bottom of the
  // first page and over the desk, and a page break did nothing anyone could
  // see. The flow is now laid onto sheets by measuring the drawn page. This
  // opens a document of several pages and reads the layout back: every block
  // sits inside a page's text area, the explicit break heads a fresh page, a
  // long paragraph is split at a line and a long table at a row, the status
  // bar counts the pages, the caret in a split paragraph's second part still
  // addresses that paragraph's characters, and a page break typed at the
  // front adds a page.
  const wordPages = async () => {
    try {
      const win = await open('word', files.long);
      const js = (code) => win.webContents.executeJavaScript(code);
      await until(() => js(`document.querySelectorAll('.wd-sheet').length >= 3 && !document.querySelector('.wd-mounting')`), 'the pages to be laid out', 15000);
      await wait(700);
      const laid = await js(`(() => {
        const page = document.querySelector('.wd-page');
        const sheets = [...page.querySelectorAll('.wd-sheet')];
        const H = sheets[0].offsetHeight;
        const P = sheets.length > 1 ? sheets[1].offsetTop - sheets[0].offsetTop : H;
        const cs = getComputedStyle(page);
        const mTop = parseFloat(cs.paddingTop);
        const mBottom = parseFloat(cs.paddingBottom);
        const flow = [...page.children].filter((el) => /wd-block|wd-table|wd-notes/.test(el.className) && !el.classList.contains('wd-notes-measure'));
        const chip = [...document.querySelectorAll('.rw-status .chip')].map((c) => c.textContent).find((t) => /^Page \\d+ of \\d+$/.test(t)) || null;
        const out = { sheets: sheets.length, blocks: flow.length, outside: [], tall: 0, parts: page.querySelectorAll('.wd-block[data-from]').length, tableParts: page.querySelectorAll('.wd-table[data-row-from]').length, breakAt: null, chip };
        for (const el of flow) {
          const top = el.offsetTop;
          const bottom = top + el.offsetHeight;
          const p = Math.floor(top / P);
          const areaTop = p * P + mTop;
          const areaBottom = p * P + H - mBottom;
          if (el.dataset.break === '1') out.breakAt = { page: p + 1, top: Math.round(top - areaTop) };
          if (el.offsetHeight > areaBottom - areaTop) { out.tall += 1; continue; }
          if (top < areaTop - 1 || bottom > areaBottom + 1) out.outside.push(el.className + '#' + (el.dataset.block || el.dataset.table) + ' ' + Math.round(top) + '-' + Math.round(bottom) + ' on page ' + (p + 1) + ' [' + Math.round(areaTop) + ',' + Math.round(areaBottom) + ']');
        }
        return out;
      })()`);
      check('word: every block of a long document sits inside a page', laid.outside.length === 0 && laid.sheets >= 3, `${laid.sheets} pages, ${laid.blocks} blocks, ${laid.tall} taller than a page${laid.outside.length ? '; outside: ' + laid.outside.slice(0, 3).join(' | ') : ''}`);
      check('word: an explicit page break heads a fresh page', Boolean(laid.breakAt) && laid.breakAt.page > 1 && Math.abs(laid.breakAt.top) <= 1, JSON.stringify(laid.breakAt));
      check('word: a long paragraph splits at a line and a long table at a row', laid.parts >= 1 && laid.tableParts >= 1, `${laid.parts} paragraph part(s), ${laid.tableParts} table part(s)`);
      check('word: the status bar counts the pages', laid.chip === `Page 1 of ${laid.sheets}`, laid.chip || 'no page chip');
      if (process.env.RUTBA_VERIFY_CAPTURE) {
        // Scrolled to the foot of page 1, so the capture shows the gap between two sheets.
        await js(`document.querySelector('.wd-scroll').scrollTop = document.querySelector('.wd-sheet').offsetHeight - 420`);
        await wait(400);
        fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'word-pages.png'), (await win.webContents.capturePage()).toPNG());
      }

      // The caret at the head of a split paragraph's second part: a character
      // typed there lands at that part's offset in the engine's paragraph.
      const where = await js(`(() => {
        const page = document.querySelector('.wd-page');
        const part = page.querySelector('.wd-block[data-from]');
        if (!part) return null;
        const node = document.createTreeWalker(part, NodeFilter.SHOW_TEXT).nextNode();
        page.focus();
        const range = document.createRange();
        range.setStart(node, 0);
        range.collapse(true);
        const sel = getSelection();
        sel.removeAllRanges();
        sel.addRange(range);
        page.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
        return { block: Number(part.dataset.block), from: Number(part.dataset.from) };
      })()`);
      const name = "word: the caret in a split paragraph's second part addresses its own characters";
      if (where) {
        await wait(300);
        win.webContents.insertText('Ω');
        const session = sessionFor('doc');
        await until(() => ((doc.model({ id: session.id }).blocks[where.block] || {}).text || '').includes('Ω'), 'the character to land', 4000);
        const text = doc.model({ id: session.id }).blocks[where.block].text || '';
        check(name, text.indexOf('Ω') === where.from, `landed at ${text.indexOf('Ω')}; the part starts at ${where.from}`);
      } else check(name, false, 'no split paragraph to type into');

      // A page break inserted at the front makes one more page. The count is
      // read once it has held still: the typing just above re-lays the pages,
      // and a count read mid-pass (6 of 11, once) made the check lie.
      let before = await js(`document.querySelectorAll('.wd-sheet').length`);
      for (let same = 0; same < 3;) {
        await wait(250);
        const now = await js(`document.querySelectorAll('.wd-sheet').length`);
        if (now === before) same += 1; else { before = now; same = 0; }
      }
      await js(`(() => {
        const page = document.querySelector('.wd-page');
        const block = page.querySelector('[data-block="0"]');
        page.focus();
        const r = document.createRange();
        r.selectNodeContents(block);
        r.collapse(true);
        const s = getSelection();
        s.removeAllRanges();
        s.addRange(r);
        page.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
        return 1;
      })()`);
      await wait(250);
      await js(`(() => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click(); return 1; })()`);
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-btn')].find((b) => /^Page Break$/i.test(b.textContent.trim())))`), 'the Insert tab', 4000);
      await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((b) => /^Page Break$/i.test(b.textContent.trim())).click(); return 1; })()`);
      await until(() => js(`document.querySelectorAll('.wd-sheet').length === ${before + 1}`), 'one more page', 6000).catch(() => {});
      await wait(300);
      const after = await js(`({
        sheets: document.querySelectorAll('.wd-sheet').length,
        chip: [...document.querySelectorAll('.rw-status .chip')].map((c) => c.textContent).find((t) => /^Page \\d+ of \\d+$/.test(t)) || null,
        breaks: [...document.querySelectorAll('[data-break]')].map((e) => e.dataset.block + '@' + Math.round(e.getBoundingClientRect().top - document.querySelector('.wd-page').getBoundingClientRect().top)),
        toasts: [...document.querySelectorAll('.rw-toast')].map((t) => t.textContent.slice(0, 80)),
      })`);
      const heads = doc.model({ id: sessionFor('doc').id }).blocks.slice(0, 3).map((b) => `${b.index}:${b.pageBreakBefore ? 'break ' : ''}${JSON.stringify((b.text || '').slice(0, 14))}`);
      check('word: a page break typed at the front adds a page', after.sheets === before + 1 && (after.chip === `Page 2 of ${before + 1}` || after.chip === `Page 1 of ${before + 1}`), `${before} → ${after.sheets} pages, chip ${after.chip}; breaks at ${after.breaks.join(' ')}; blocks ${heads.join(' ')}${after.toasts.length ? '; toasts: ' + after.toasts.join(' | ') : ''}`);
    } catch (err) {
      check('word: the pages check ran', false, err.message);
    }
  };

  /* ── Worksheets: drag a heading's edge to resize ─────────────────────── */
  //
  // "Column resizing did not work for me": there was nothing to drag. The
  // edge of a column or row heading is now a handle. This drags column B's
  // by sixty pixels and row 2's by fifteen, with real pointer events, and
  // reads the sizes back from the engine.
  const sheetGrips = async () => {
    try {
      const win = await open('sheets', files.xlsx);
      const js = (code) => win.webContents.executeJavaScript(code);
      const wc = win.webContents;
      const widths = () => (doc.model({ id: sessionFor('sheet').id }).columns || []).map((c) => c.width);
      const heights = () => (doc.model({ id: sessionFor('sheet').id }).rows || []).map((r) => r.height);
      const drag = async (selector, dx, dy) => {
        const at = await js(`(() => { const r = document.querySelector(${JSON.stringify(selector)})?.getBoundingClientRect(); return r ? { x: r.left + r.width / 2, y: r.top + r.height / 2 } : null; })()`);
        if (!at) throw new Error(`no handle at ${selector}`);
        const x = Math.round(at.x);
        const y = Math.round(at.y);
        wc.sendInputEvent({ type: 'mouseMove', x, y });
        wc.sendInputEvent({ type: 'mouseDown', x, y, button: 'left', clickCount: 1 });
        await wait(60);
        for (const f of [0.3, 0.6, 1]) {
          wc.sendInputEvent({ type: 'mouseMove', x: Math.round(x + dx * f), y: Math.round(y + dy * f), button: 'left' });
          await wait(40);
        }
        wc.sendInputEvent({ type: 'mouseUp', x: x + dx, y: y + dy, button: 'left', clickCount: 1 });
        await wait(400);
      };

      await until(() => js(`Boolean(document.querySelector('.sh-colheads .sh-head:nth-child(2) .sh-grip.col'))`), 'the column handles', 6000);
      const w0 = widths()[1];
      await drag('.sh-colheads .sh-head:nth-child(2) .sh-grip.col', 60, 0);
      const w1 = widths()[1];
      check('sheets: dragging a column heading\'s edge resizes the column', Math.abs(w1 - (w0 + 60)) <= 2, `B went ${w0} → ${w1} px`);

      const h0 = heights()[1];
      await drag('.sh-rowheads .sh-head:nth-child(2) .sh-grip.row', 0, 15);
      const h1 = heights()[1];
      check('sheets: dragging a row heading\'s edge resizes the row', Math.abs(h1 - (h0 + 15)) <= 2, `row 2 went ${h0} → ${h1} px`);

      const guide = await js(`document.querySelectorAll('.sh-guide').length`);
      check('sheets: the guide line goes when the drag ends', guide === 0, `${guide} guide(s) left`);
      const complaints = await errorsIn(win);
      check('sheets: resizing reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('sheets: the resize check ran', false, err.message);
    }
  };

  /* ── Presentation: the Layers and Designs panes ──────────────────────── */
  //
  // Two panes on the right, opened from the View tab: the slide's shapes as
  // layers (select, reorder, hide, rename) and the deck's layouts as a
  // gallery (put this slide on one, or start a new slide from it). Each verb
  // is driven through the pane and read back from the engine.
  const slidePanes = async () => {
    try {
      const win = await open('slides', files.pptx);
      const js = (code) => win.webContents.executeJavaScript(code);
      const model = () => doc.model({ id: sessionFor('deck').id, slide: 0 });
      const tab = async (name) => {
        await js(`(() => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(); return 1; })()`);
        await wait(200);
      };
      const button = async (label) => {
        await until(() => js(`Boolean([...document.querySelectorAll('.rw-btn')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}))`), `the ${label} button`, 4000);
        await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}).click(); return 1; })()`);
        await wait(250);
      };

      await tab('View');
      await button('Layers');
      await until(() => js(`document.querySelectorAll('.sl-layer').length > 0`), 'the layer rows', 5000);
      const shapes = model().slide.shapes;
      const rows = await js(`[...document.querySelectorAll('.sl-layer')].map((r) => r.dataset.shape)`);
      check('slides: the Layers pane lists the slide\'s shapes, top-most first', rows.length === shapes.length && rows[0] === String(shapes[shapes.length - 1].id), `rows ${rows.join(',')}; shapes ${shapes.map((s) => s.id).join(',')}`);
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'slides-layers.png'), (await win.webContents.capturePage()).toPNG());

      // The bottom-most shape, selected in the pane, brought to the front.
      const bottom = String(shapes[0].id);
      await js(`(() => { [...document.querySelectorAll('.sl-layer')].find((r) => r.dataset.shape === ${JSON.stringify(bottom)}).click(); return 1; })()`);
      await until(() => js(`document.querySelector('.sl-layer.active')?.dataset.shape === ${JSON.stringify(bottom)}`), 'the row to select', 3000);
      await js(`(() => { [...document.querySelectorAll('.sl-layers-tools .rw-btn')].find((b) => b.textContent.trim() === 'Front').click(); return 1; })()`);
      await until(() => String(model().slide.shapes.slice(-1)[0].id) === bottom, 'the shape to come to the front', 4000);
      const order = model().slide.shapes.map((s) => String(s.id));
      check('slides: Front brings the selected shape to the top of the drawing order', order[order.length - 1] === bottom, `order is now ${order.join(',')}`);

      // The eye hides it: the engine says hidden, the row says off.
      await js(`(() => { document.querySelector('.sl-layer[data-shape=' + JSON.stringify(${JSON.stringify(bottom)}) + '] .sl-eye').click(); return 1; })()`);
      await until(() => model().slide.shapes.find((s) => String(s.id) === bottom)?.hidden === true, 'the shape to hide', 4000);
      // The engine answers before the window has painted the answer.
      const off = await until(() => js(`document.querySelector('.sl-layer[data-shape=' + JSON.stringify(${JSON.stringify(bottom)}) + ']').classList.contains('off')`), 'the row to say hidden', 4000).catch(() => false);
      check('slides: the eye hides a shape, and the engine and the row agree', off === true, off ? 'hidden' : 'the row does not say so');

      // Designs: the deck's two layouts; the other one, applied.
      await button('Designs');
      await until(() => js(`document.querySelectorAll('.sl-design').length > 0`), 'the layout cards', 5000);
      const cards = await js(`[...document.querySelectorAll('.sl-design')].map((c) => ({ part: c.dataset.layout, active: c.classList.contains('active'), name: c.querySelector('.sl-design-name')?.textContent }))`);
      const current = model().slide.layout;
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'slides-designs.png'), (await win.webContents.capturePage()).toPNG());
      check('slides: the Designs pane shows the deck\'s layouts with this slide\'s marked', cards.length === 2 && cards.filter((c) => c.active).length === 1 && cards.find((c) => c.active)?.part === current, JSON.stringify(cards));
      const other = cards.find((c) => !c.active);
      if (other) {
        await js(`(() => { document.querySelector('.sl-design[data-layout=' + JSON.stringify(${JSON.stringify(other.part)}) + ']').click(); return 1; })()`);
        await until(() => model().slide.layout === other.part, 'the layout to apply', 4000);
        const marked = await until(() => js(`document.querySelector('.sl-design.active')?.dataset.layout === ${JSON.stringify(other.part)}`), 'the card to be marked', 4000).catch(() => false);
        check('slides: clicking a layout puts the slide on it', model().slide.layout === other.part && marked === true, `now on ${other.name}${marked ? '' : ', but the card is not marked'}`);
        const count = model().count;
        await js(`(() => { document.querySelector('.sl-design[data-layout=' + JSON.stringify(${JSON.stringify(other.part)}) + '] .sl-design-new').click(); return 1; })()`);
        await until(() => model().count === count + 1, 'a new slide', 4000);
        check('slides: New starts a slide on that layout', model().count === count + 1 && doc.model({ id: sessionFor('deck').id, slide: 1 }).slide.layout === other.part, `${count} → ${model().count} slides`);
      } else check('slides: clicking a layout puts the slide on it', false, 'no other layout to click');
      const complaints = await errorsIn(win);
      check('slides: the panes report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('slides: the panes check ran', false, err.message);
    }
  };

  /* ── Rutba Word: a picture floats and the words wrap round it ─────────── */
  //
  // Pictures sat in a paragraph of their own whatever the file said. A
  // picture in a wp:anchor now floats where the file puts it and the words
  // run round it; Wrap Text and Position on the Layout tab move a picked
  // picture between the line, the sides and the middle. The words' first
  // line is measured against the picture each time.
  const wordFloat = async () => {
    if (!files.float) return check('word: the floating-picture fixture exists', false, 'no icon to make it from');
    try {
      const win = await open('word', files.float);
      const js = (code) => win.webContents.executeJavaScript(code);
      const measure = () => js(`(() => {
        const page = document.querySelector('.wd-page');
        const img = page.querySelector('.wd-image');
        const words = page.querySelector('[data-block="2"]');
        if (!img || !words) return null;
        const node = document.createTreeWalker(words, NodeFilter.SHOW_TEXT).nextNode();
        const r = document.createRange();
        r.setStart(node, 0);
        r.setEnd(node, 1);
        const t = r.getBoundingClientRect();
        const i = img.getBoundingClientRect();
        const cs = getComputedStyle(page);
        const p = page.getBoundingClientRect();
        return {
          floating: img.classList.contains('wd-float'), picked: img.classList.contains('picked'),
          text: { left: Math.round(t.left), top: Math.round(t.top), right: Math.round(t.right) },
          img: { left: Math.round(i.left), top: Math.round(i.top), right: Math.round(i.right), bottom: Math.round(i.bottom) },
          column: { left: Math.round(p.left + parseFloat(cs.paddingLeft)), right: Math.round(p.right - parseFloat(cs.paddingRight)) },
        };
      })()`);
      await until(() => js(`Boolean(document.querySelector('.wd-image.wd-float'))`), 'the floating picture', 8000);
      await wait(400);
      const right = await measure();
      check('word: a picture anchored at the right floats there and the first line runs beside it',
        right && right.floating && right.text.left <= right.column.left + 2 && right.text.top < right.img.bottom && right.text.top >= right.img.top - 4 && Math.abs(right.img.right - right.column.right) <= 3,
        JSON.stringify(right));
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'word-float.png'), (await win.webContents.capturePage()).toPNG());

      // Pick it, and put it in the line: the words drop below it.
      await js(`(() => { document.querySelector('.wd-image').click(); return 1; })()`);
      await until(() => js(`document.querySelector('.wd-image')?.classList.contains('picked')`), 'the picture to be picked', 4000);
      await js(`(() => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Layout')?.click(); return 1; })()`);
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-btn')].find((b) => b.textContent.trim() === 'Wrap Text' && !b.disabled))`), 'Wrap Text to be enabled', 4000);
      await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((b) => b.textContent.trim() === 'Wrap Text').click(); return 1; })()`);
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === 'In line with text'))`), 'the Wrap Text menu', 4000);
      await js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === 'In line with text').click(); return 1; })()`);
      // Leaving the float is the heaviest step here: the paragraph regains the
      // picture's height under its words, pushing what follows down and
      // sometimes over a page, so the paginator redoes more work than the
      // still-floating toggles round it do. Xvfb's software rendering, with a
      // busy main process behind it, can turn that into real seconds rather
      // than the milliseconds it costs on the owner's machine (see the "busy"
      // clock above) — the same margin the picture's first float already gets.
      await until(() => js(`Boolean(document.querySelector('.wd-image')) && !document.querySelector('.wd-image.wd-float')`), 'the picture to leave the words', 8000);
      await wait(300);
      const inline = await measure();
      check('word: Wrap Text → In line with text puts the picture in its own line, the words below',
        inline && !inline.floating && inline.text.top >= inline.img.bottom - 2,
        JSON.stringify(inline));

      // Position → Left: it floats at the left, the words at its right.
      await js(`(() => { document.querySelector('.wd-image').click(); return 1; })()`);
      await until(() => js(`document.querySelector('.wd-image')?.classList.contains('picked')`), 'the picture to be picked again', 4000);
      await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((b) => b.textContent.trim() === 'Position').click(); return 1; })()`);
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => /^Left/.test(b.textContent.trim())))`), 'the Position menu', 4000);
      await js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((b) => /^Left/.test(b.textContent.trim())).click(); return 1; })()`);
      await until(() => js(`Boolean(document.querySelector('.wd-image.wd-float'))`), 'the picture to float again', 5000);
      await wait(300);
      const left = await measure();
      check('word: Position → Left floats the picture at the left with the words at its right',
        left && left.floating && Math.abs(left.img.left - left.column.left) <= 3 && left.text.left > left.img.right && left.text.top < left.img.bottom,
        JSON.stringify(left));

      // And the file says so.
      await press(win.webContents, 's', { modifiers: ['control'] });
      await wait(1200);
      const saved = openDocx(fs.readFileSync(files.float)).render({ pages: false }).blocks[1].images[0];
      check('word: the picture\'s wrap and position are saved', saved && saved.anchored && saved.wrap === 'square' && saved.hAlign === 'left', JSON.stringify({ anchored: saved?.anchored, wrap: saved?.wrap, hAlign: saved?.hAlign }));
      const complaints = await errorsIn(win);
      check('word: floating a picture reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('word: the floating-picture check ran', false, err.message);
    }
  };

  /* ── The frame: tooltips of our own, and the zoom slider ─────────────── */
  //
  // Every ribbon button carries its tip as data and the frame draws it as a
  // small chip after a rest, in the suite's own look; the status bar of each
  // document app ends in a zoom slider. Both are driven with real pointer
  // events and read back from the window and from Electron's zoom factor.
  const polish = async () => {
    try {
      const win = await open('word', files.docx);
      const js = (code) => win.webContents.executeJavaScript(code);
      const wc = win.webContents;
      const at = await js(`(() => { const r = document.querySelector('.rw-btn[data-tip^="Bold"]')?.getBoundingClientRect(); return r ? { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) } : null; })()`);
      if (!at) throw new Error('no Bold button to hover');
      wc.sendInputEvent({ type: 'mouseMove', x: at.x - 6, y: at.y });
      await wait(60);
      wc.sendInputEvent({ type: 'mouseMove', x: at.x, y: at.y });
      const shown = await until(() => js(`(() => { const t = document.querySelector('.rw-tip.on'); return t ? t.textContent : null; })()`), 'the tooltip to show', 3000).catch(() => null);
      const tipText = await js(`document.querySelector('.rw-tip.on')?.textContent || document.querySelector('.rw-tip')?.textContent || ''`);
      const tips = await js(`document.querySelectorAll('.rw-tip').length`);
      const native = await js(`document.querySelector('.rw-btn[data-tip^="Bold"]').hasAttribute('title')`);
      check('frame: resting on a ribbon button shows the suite\'s own tooltip, not the browser\'s', shown === true && /^Bold/.test(tipText) && native === false && tips === 1, `tip reads ${JSON.stringify(tipText.slice(0, 40))}; title attribute: ${native}; ${tips} tip element(s)`);
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'frame-tooltip.png'), (await win.webContents.capturePage()).toPNG());
      const page = await js(`(() => { const r = document.querySelector('.wd-page').getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + 120) }; })()`);
      wc.sendInputEvent({ type: 'mouseMove', x: page.x, y: page.y });
      const gone = await until(() => js(`!document.querySelector('.rw-tip.on')`), 'the tooltip to go', 3000).catch(() => false);
      check('frame: the tooltip goes when the pointer leaves', gone === true, gone ? 'gone' : 'still showing');

      // The slider: 150% through the range input, 100% through the level button.
      const setRange = (v) => js(`(() => {
        const input = document.querySelector('.rw-zoom input[type="range"]');
        if (!input) return false;
        const set = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
        set.call(input, ${JSON.stringify(String(v))});
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
        return true;
      })()`);
      // The slider scales the page alone since 2026-09-24 — the window's
      // own factor stays at 1 — so what is read is the page's drawn width
      // against its own; the `zoom` block covers the ribbon and the slider.
      const pageZoom = () => js(`(() => { const p = document.querySelector('.wd-page'); return p ? p.getBoundingClientRect().width / p.offsetWidth : 0; })()`);
      const had = await setRange(1.5);
      const zoomed = await until(async () => Math.abs((await pageZoom()) - 1.5) < 0.02, 'the page to zoom', 4000).catch(() => false);
      const pct = (await until(() => js(`document.querySelector('.rw-zoom-pct')?.textContent === '150%'`), 'the level to read 150%', 3000).catch(() => false)) ? '150%' : await js(`document.querySelector('.rw-zoom-pct')?.textContent`);
      check('frame: the status bar\'s zoom slider zooms the page, and the window stays at 100%', had && zoomed === true && pct === '150%' && Math.abs(wc.getZoomFactor() - 1) < 0.001, `page ${(await pageZoom()).toFixed(2)}, window factor ${wc.getZoomFactor()}, level reads ${pct}`);
      await js(`(() => { document.querySelector('.rw-zoom-pct').click(); return 1; })()`);
      const back = await until(async () => Math.abs((await pageZoom()) - 1) < 0.02, 'the page to reset', 4000).catch(() => false);
      const reads100 = await until(() => js(`document.querySelector('.rw-zoom-pct')?.textContent === '100%'`), 'the level to read 100%', 3000).catch(() => false);
      check('frame: the level button puts the zoom back to 100%', back === true && reads100 === true, `page ${(await pageZoom()).toFixed(2)}, level reads ${await js(`document.querySelector('.rw-zoom-pct')?.textContent`)}`);

      // A deck's slider scales the slide; the level button fits it to the window again.
      const deck = await open('slides', files.pptx);
      const djs = (code) => deck.webContents.executeJavaScript(code);
      await until(() => djs(`Boolean(document.querySelector('.sl-slide'))`), 'the slide', 6000);
      await djs(`(() => { const input = document.querySelector('.rw-zoom input[type="range"]'); const set = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set; set.call(input, '0.5'); input.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
      const scaled = await until(() => djs(`/scale\\(0\\.5\\)/.test(document.querySelector('.sl-slide')?.style.transform || '')`), 'the slide to scale', 4000).catch(() => false);
      await djs(`(() => { document.querySelector('.rw-zoom-pct').click(); return 1; })()`);
      const fitted = await until(() => djs(`!/scale\\(0\\.5\\)/.test(document.querySelector('.sl-slide')?.style.transform || '')`), 'the slide to fit again', 4000).catch(() => false);
      check('slides: the zoom slider scales the slide and the level button fits it again', scaled === true && fitted === true, `scaled ${scaled}, fitted ${fitted}`);
    } catch (err) {
      check('frame: the polish check ran', false, err.message);
    }
  };

  /* ── Presentation: shapes move by hand, and the Format pane recolours them ── */
  //
  // Every shape on the slide has a hit area: a click selects it and shows
  // eight handles, a drag moves it, a handle resizes it, the arrows nudge it,
  // and the engine is told on release. The Format pane sets its fill and
  // outline from the theme's swatches. Real pointer events, geometry read
  // back from the engine.
  const slideShapes = async () => {
    try {
      const win = await open('slides', files.pptx);
      const js = (code) => win.webContents.executeJavaScript(code);
      const wc = win.webContents;
      const model = () => doc.model({ id: sessionFor('deck').id, slide: 0 });
      const shapeId = String(model().slide.shapes[0].id);
      const geometry = () => model().slide.shapes.find((s) => String(s.id) === shapeId).geometry;
      const rectOf = (selector) => js(`(() => { const r = document.querySelector(${JSON.stringify(selector)})?.getBoundingClientRect(); return r ? { x: r.left + r.width / 2, y: r.top + r.height / 2, w: r.width, h: r.height } : null; })()`);
      const drag = async (selector, dx, dy) => {
        const at = await rectOf(selector);
        if (!at) throw new Error(`nothing at ${selector}`);
        const x = Math.round(at.x);
        const y = Math.round(at.y);
        wc.sendInputEvent({ type: 'mouseMove', x, y });
        wc.sendInputEvent({ type: 'mouseDown', x, y, button: 'left', clickCount: 1 });
        await wait(60);
        for (const f of [0.25, 0.5, 0.75, 1]) {
          wc.sendInputEvent({ type: 'mouseMove', x: Math.round(x + dx * f), y: Math.round(y + dy * f), button: 'left' });
          await wait(40);
        }
        wc.sendInputEvent({ type: 'mouseUp', x: x + dx, y: y + dy, button: 'left', clickCount: 1 });
        await wait(500);
      };
      const scale = await js(`(() => { const m = /scale\\(([\\d.]+)\\)/.exec(document.querySelector('.sl-slide')?.style.transform || ''); return m ? Number(m[1]) : 1; })()`);
      const hit = `.sl-hit[data-shape="${shapeId}"]`;
      await until(() => js(`Boolean(document.querySelector(${JSON.stringify(hit)}))`), 'the shape\'s hit area', 6000);

      // Select: handles appear.
      await js(`(() => { document.querySelector(${JSON.stringify(hit)}).click(); return 1; })()`);
      const handles = await until(() => js(`document.querySelectorAll('.sl-handle').length === 8`), 'the eight handles', 4000).catch(() => false);
      check('slides: clicking a shape selects it and shows eight handles', handles === true, `${await js(`document.querySelectorAll('.sl-handle').length`)} handle(s)`);

      // Move by dragging the shape.
      const g0 = geometry();
      await drag(hit, 60, 30);
      const g1 = await (async () => { await until(() => Math.abs(geometry().x - (g0.x + 60 / scale)) <= 2, 'the shape to move', 4000).catch(() => {}); return geometry(); })();
      check('slides: dragging a shape moves it', Math.abs(g1.x - (g0.x + 60 / scale)) <= 2 && Math.abs(g1.y - (g0.y + 30 / scale)) <= 2 && Math.abs(g1.w - g0.w) <= 1 && Math.abs(g1.h - g0.h) <= 1, `from ${Math.round(g0.x)},${Math.round(g0.y)} to ${Math.round(g1.x)},${Math.round(g1.y)} at scale ${scale}; size ${Math.round(g0.w)}×${Math.round(g0.h)} → ${Math.round(g1.w)}×${Math.round(g1.h)}`);

      // Resize by the south-east handle.
      await drag('.sl-handle[data-handle="se"]', 40, 20);
      const g2 = await (async () => { await until(() => Math.abs(geometry().w - (g1.w + 40 / scale)) <= 2, 'the shape to grow', 4000).catch(() => {}); return geometry(); })();
      check('slides: dragging a handle resizes the shape', Math.abs(g2.w - (g1.w + 40 / scale)) <= 2 && Math.abs(g2.h - (g1.h + 20 / scale)) <= 2 && Math.abs(g2.x - g1.x) <= 1, `${Math.round(g1.w)}×${Math.round(g1.h)} → ${Math.round(g2.w)}×${Math.round(g2.h)}`);
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'slides-handles.png'), (await win.webContents.capturePage()).toPNG());

      // The keyboard: an arrow nudges by a pixel.
      await js(`(() => { document.querySelector('.sl-stage').focus(); return 1; })()`);
      await press(wc, 'Right');
      const g3 = await (async () => { await until(() => Math.abs(geometry().x - (g2.x + 1)) <= 0.5, 'the nudge', 3000).catch(() => {}); return geometry(); })();
      check('slides: an arrow key nudges the selected shape', Math.abs(g3.x - (g2.x + 1)) <= 0.5, `x ${g2.x} → ${g3.x}`);

      // The Format pane: a theme swatch for the fill, no outline, a weight.
      await js(`(() => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Home')?.click(); return 1; })()`);
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-btn')].find((b) => b.textContent.trim() === 'Shape Fill' && !b.disabled))`), 'Shape Fill to be enabled', 4000);
      await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((b) => b.textContent.trim() === 'Shape Fill').click(); return 1; })()`);
      await until(() => js(`Boolean(document.querySelector('.sl-format .sl-swatch[data-swatch="accent2"]'))`), 'the Format pane', 4000);
      const accent2 = await js(`getComputedStyle(document.querySelector('.sl-format .sl-swatch[data-swatch="accent2"]')).backgroundColor`);
      await js(`(() => { document.querySelector('.sl-format .sl-swatch[data-swatch="accent2"]').click(); return 1; })()`);
      const filled = await until(() => model().slide.shapes.find((s) => String(s.id) === shapeId).fill?.type === 'solid', 'the fill to be written', 4000).catch(() => false);
      const fill = model().slide.shapes.find((s) => String(s.id) === shapeId).fill;
      const rgb = (hex) => { const m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(hex || ''); return m ? `rgb(${parseInt(m[1], 16)}, ${parseInt(m[2], 16)}, ${parseInt(m[3], 16)})` : null; };
      check('slides: a theme swatch in the Format pane fills the shape with that colour', filled === true && rgb(fill?.color) === accent2, `fill ${JSON.stringify(fill)}, swatch ${accent2}`);
      // The engine wrote the fill before the window redrew it: wait for the
      // drawing, not for the model, or a loaded machine fails this by a frame.
      const drawn = await until(() => js(`document.querySelector('.sl-svg')?.innerHTML.includes(${JSON.stringify(String(fill?.color || '').toLowerCase())}) || document.querySelector('.sl-svg')?.innerHTML.includes(${JSON.stringify(String(fill?.color || '').toUpperCase())})`), 'the slide to redraw with the fill', 4000).catch(() => false);
      check('slides: the slide is redrawn with the new fill', drawn === true, drawn ? 'the colour is in the drawing' : 'not in the drawing');
      await js(`(() => { [...document.querySelectorAll('.sl-format .sl-chip')].find((b) => b.textContent.trim() === 'No outline').click(); return 1; })()`);
      const noLine = await until(() => model().slide.shapes.find((s) => String(s.id) === shapeId).line?.type === 'none', 'the outline to go', 4000).catch(() => false);
      check('slides: No outline removes the outline', noLine === true, JSON.stringify(model().slide.shapes.find((s) => String(s.id) === shapeId).line));
      await js(`(() => { document.querySelector('.sl-format-line .sl-swatch[data-swatch="#0070C0"]').click(); return 1; })()`);
      await until(() => model().slide.shapes.find((s) => String(s.id) === shapeId).line?.color, 'the outline colour', 4000).catch(() => {});
      // The pane paints the answer after the engine has it; the next press reads the pane.
      await until(() => js(`document.querySelector('.sl-format-line .sl-swatch[data-swatch="#0070C0"]')?.classList.contains('current')`), 'the outline swatch to be marked', 4000).catch(() => {});
      await js(`(() => { const sel = [...document.querySelectorAll('.sl-format-line select')][0]; const set = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set; set.call(sel, '3'); sel.dispatchEvent(new Event('change', { bubbles: true })); return 1; })()`);
      const weighed = await until(() => model().slide.shapes.find((s) => String(s.id) === shapeId).line?.width === 3, 'the weight', 4000).catch(() => false);
      const line = model().slide.shapes.find((s) => String(s.id) === shapeId).line;
      check('slides: an outline colour and a weight of 3 pt are written together', weighed === true && String(line?.color).toUpperCase() === '#0070C0', JSON.stringify(line));
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'slides-format.png'), (await win.webContents.capturePage()).toPNG());
      const complaints = await errorsIn(win);
      check('slides: moving and formatting report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('slides: the shapes check ran', false, err.message);
    }
  };

  /* ── Presentation: bullets, list level and line spacing ─────────────── */
  //
  // Home → Bullets on a text box gives its paragraphs a bullet the stage
  // draws; Increase list level and Line spacing → 1.5 ride on the paragraph;
  // the saved file carries all three as PowerPoint keeps them. Run alone
  // with RUTBA_VERIFY_ONLY=bullets.
  const slideParagraphs = async () => {
    try {
      const win = await open('slides', files.pptx);
      const js = (code) => win.webContents.executeJavaScript(code);
      const model = () => doc.model({ id: sessionFor('deck').id, slide: 0 });
      const shape = model().slide.shapes.find((s) => s.text && s.placeholder?.type !== 'title') || model().slide.shapes.find((s) => s.text);
      if (!shape) return check('slides: a text box to give bullets', false, 'no text shape on slide 1');
      const shapeId = String(shape.id);
      const hit = `.sl-hit[data-shape="${shapeId}"]`;
      await until(() => js(`Boolean(document.querySelector(${JSON.stringify(hit)}))`), 'the hit area of the text box', 6000);
      await js(`(() => { document.querySelector(${JSON.stringify(hit)}).click(); return 1; })()`);
      await until(() => js(`document.querySelectorAll('.sl-handle').length === 8`), 'the eight handles', 4000).catch(() => {});
      const clickRibbon = (title) => js(`(() => {
        const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}));
        if (!b) return 'no button ' + ${JSON.stringify(title)};
        b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
        b.click();
        return 'clicked';
      })()`);
      const first = () => model().slide.shapes.find((s) => String(s.id) === shapeId)?.text?.paragraphs?.[0] || {};
      const had = first().bullet?.type || 'none';
      const clickedBullets = await clickRibbon('Bullets');
      const toggled = await until(() => (first().bullet?.type || 'none') === (had === 'char' ? 'none' : 'char'), 'the bullet to toggle', 4000).catch(() => false);
      if (had === 'char') {
        await clickRibbon('Bullets');
        await until(() => first().bullet?.type === 'char', 'the bullet back', 4000).catch(() => {});
      }
      const glyph = await until(() => js(`[...document.querySelectorAll('.sl-slide text, .sl-stage text, svg text')].some((t) => t.textContent.trim() === '•')`), 'the bullet glyph drawn', 4000).catch(() => false);
      check('slides: Home → Bullets gives the paragraph a bullet, and the stage draws it',
        clickedBullets === 'clicked' && toggled === true && glyph === true, `${clickedBullets}; had ${had}; now ${JSON.stringify(first().bullet)}; glyph ${glyph}`);

      const clickedLevel = await clickRibbon('Increase list level');
      const levelled = await until(() => first().level === 1, 'the list level', 4000).catch(() => false);
      // The window's own model follows the engine's a tick later: a press on
      // its heels would format the paragraph as it was before the last one.
      await wait(300);
      const clickedSpacing = await clickRibbon('Line spacing');
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === '1.5'))`), 'the spacing menu', 3000).catch(() => {});
      await js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === '1.5')?.click(); return 1; })()`);
      const spaced = await until(() => first().lineHeight === 1.5, 'the line spacing', 4000).catch(() => false);
      check('slides: Increase list level and Line spacing → 1.5 ride on the paragraph',
        clickedLevel === 'clicked' && levelled === true && clickedSpacing === 'clicked' && spaced === true, `level ${first().level}, spacing ${first().lineHeight}`);

      await wait(300);
      const clickedStrike = await clickRibbon('Strikethrough');
      const struck = await until(() => Boolean(first().runs?.[0]?.strike), 'the strikethrough', 4000).catch(() => false);
      const lined = await until(() => js(`[...document.querySelectorAll('svg tspan')].some((t) => /line-through/.test(t.getAttribute('text-decoration') || ''))`), 'the line through the words', 4000).catch(() => false);
      check('slides: Strikethrough rides on the run, and the stage draws the line through the words',
        clickedStrike === 'clicked' && struck === true && lined === true, `${clickedStrike}; strike ${first().runs?.[0]?.strike}; drawn ${lined}`);

      // The box's own: Align text → Middle sits on the body, and Text
      // direction → rotated up turns the drawing.
      await wait(300);
      const bodyOf = () => model().slide.shapes.find((s) => String(s.id) === shapeId)?.text || {};
      const pickBody = async (button, label) => {
        const clicked = await clickRibbon(button);
        await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim().startsWith(${JSON.stringify(label)})))`), `the ${button} menu`, 3000).catch(() => {});
        await js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim().startsWith(${JSON.stringify(label)}))?.click(); return 1; })()`);
        return clicked;
      };
      const clickedAlign = await pickBody('Align text', 'Middle');
      const anchored = await until(() => bodyOf().anchor === 'middle', 'the middle anchor', 4000).catch(() => false);
      await wait(300);
      const clickedDirection = await pickBody('Text direction', 'Rotate all text 270');
      const turned = await until(() => bodyOf().vert === 'vert270', 'the words to run up', 4000).catch(() => false);
      const drawnTurned = await until(() => js(`[...document.querySelectorAll('svg g[transform]')].some((g) => /rotate\\(-90 /.test(g.getAttribute('transform') || ''))`), 'the stage to turn the words', 4000).catch(() => false);
      check('slides: Align text → Middle and Text direction → rotated up sit on the text body, and the stage turns the words',
        clickedAlign === 'clicked' && anchored === true && clickedDirection === 'clicked' && turned === true && drawnTurned === true,
        `${clickedAlign} anchor ${bodyOf().anchor}; ${clickedDirection} vert ${bodyOf().vert}; drawn ${drawnTurned}`);
      await wait(300);
      await pickBody('Text direction', 'Horizontal');
      await until(() => bodyOf().vert === 'horz', 'the words horizontal again', 4000).catch(() => {});

      await wait(300);
      await clickRibbon('Save');
      const inFile = () => Deck.open(fs.readFileSync(files.pptx)).slide(0).shapes.find((s) => String(s.id) === shapeId).text.paragraphs[0];
      await until(() => { try { return inFile().lineHeight === 1.5; } catch { return false; } }, 'the paragraph to land in the file', 8000).catch(() => false);
      const kept = inFile();
      check('slides: the saved file keeps the bullet, the level and the spacing as PowerPoint does',
        kept.bullet?.char === '•' && kept.level === 1 && kept.lineHeight === 1.5, JSON.stringify({ bullet: kept.bullet, level: kept.level, lineHeight: kept.lineHeight }));

      // Home → Layout: the slide goes onto another of the deck's layouts.
      await wait(300);
      const layouts = model().layouts || [];
      const current = model().slide.layout;
      const other = layouts.find((l) => l.part !== current);
      if (other) {
        const clickedLayout = await clickRibbon('Layout');
        const itemName = other.name || other.part;
        await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(itemName)}))`), 'the layouts menu', 3000).catch(() => {});
        await js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(itemName)})?.click(); return 1; })()`);
        const moved = await until(() => model().slide.layout === other.part, 'the slide on the other layout', 5000).catch(() => false);
        check('slides: Home → Layout puts the slide on another layout of the deck', clickedLayout === 'clicked' && moved === true, `${clickedLayout}; now on ${model().slide.layout}`);
      } else {
        check('slides: Home → Layout puts the slide on another layout of the deck', false, 'no other layout in the deck');
      }
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'slides-bullets.png'), (await win.webContents.capturePage()).toPNG());
      const complaints = await errorsIn(win);
      check('slides: the paragraph checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('slides: the paragraph checks ran', false, err.message);
    }
  };

  /* ── Word: paragraph shading and borders ───────────────────────────── */
  //
  // Home → Shading puts a colour behind the paragraph and Borders a line
  // under it; the page paints both at once, and the saved file keeps them
  // where Word does. Run alone with RUTBA_VERIFY_ONLY=look.
  const wordLook = async () => {
    try {
      const win = await open('word', files.docx);
      const js = (code) => win.webContents.executeJavaScript(code);
      await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="0"]'))`), 'the first paragraph', 8000);
      // The caret into the first paragraph, the way a click puts it there.
      await js(`(() => {
        const b = document.querySelector('.wd-page [data-block="0"]');
        b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }));
        const r = document.createRange(); r.selectNodeContents(b); r.collapse(true);
        const s = getSelection(); s.removeAllRanges(); s.addRange(r);
        // The window learns the caret on mouseup, as a click gives it.
        document.querySelector('.wd-page').dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
        return 1;
      })()`);
      await wait(300);
      const clickRibbon = (title) => js(`(() => {
        const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}));
        if (!b) return 'no button ' + ${JSON.stringify(title)};
        b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
        b.click();
        return 'clicked';
      })()`);
      const pick = async (button, label) => {
        const clicked = await clickRibbon(button);
        await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}))`), `the ${button} menu`, 3000).catch(() => {});
        await js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)})?.click(); return 1; })()`);
        return clicked;
      };
      const paint = () => js(`(() => { const b = document.querySelector('.wd-page [data-block="0"]'); if (!b) return null; const cs = getComputedStyle(b); return { background: cs.backgroundColor, bottom: cs.borderBottomWidth + ' ' + cs.borderBottomStyle }; })()`);
      const pickedShade = await pick('Shading', 'Light yellow');
      const shaded = await until(async () => (await paint())?.background === 'rgb(255, 242, 204)', 'the shading painted', 5000).catch(() => false);
      await wait(300);
      const pickedBorder = await pick('Borders', 'Bottom border');
      const ruled = await until(async () => /^[1-9][0-9.]*px solid$/.test((await paint())?.bottom || ''), 'the rule painted', 5000).catch(() => false);
      check('word: Home → Shading and Borders paint a colour behind the paragraph and a rule under it at once',
        pickedShade === 'clicked' && shaded === true && pickedBorder === 'clicked' && ruled === true, `${pickedShade} ${pickedBorder}; ${JSON.stringify(await paint())}`);
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'word-look.png'), (await win.webContents.capturePage()).toPNG());

      await wait(300);
      await clickRibbon('Save');
      const pPrInFile = () => openDocx(fs.readFileSync(files.docx)).doc.doc.editParagraph(0).pPr || '';
      await until(() => { try { return /w:shd/.test(pPrInFile()); } catch { return false; } }, 'the look to land in the file', 8000).catch(() => false);
      const pPr = pPrInFile();
      check('word: the saved file keeps the shading and the border where Word does — pBdr then shd in the paragraph properties',
        /<w:pBdr><w:bottom w:val="single"[^>]*\/><\/w:pBdr>/.test(pPr) && /<w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"\/>/.test(pPr) && pPr.indexOf('<w:pBdr>') < pPr.indexOf('<w:shd'),
        pPr.slice(0, 220));
      // Design → Page Colour: the page (or the sheets behind the flow in
      // print layout) takes the colour, and the file keeps it before the body.
      await wait(300);
      await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Design')?.click(), 'tab'`);
      await wait(200);
      const pickedPage = await pick('Page Colour', 'Light blue');
      const pageBg = () => js(`(() => { const el = document.querySelector('.wd-sheet') || document.querySelector('.wd-page'); return el ? getComputedStyle(el).backgroundColor : null; })()`);
      const coloured = await until(async () => (await pageBg()) === 'rgb(222, 235, 247)', 'the page colour painted', 5000).catch(() => false);
      check('word: Design → Page Colour paints the page', pickedPage === 'clicked' && coloured === true, `${pickedPage}; page ${await pageBg()}`);
      await wait(300);
      await clickRibbon('Save');
      const documentXml = () => openDocx(fs.readFileSync(files.docx)).doc.doc.xml;
      await until(() => { try { return /<w:background w:color="DEEBF7"\/>/.test(documentXml()); } catch { return false; } }, 'the page colour to land in the file', 8000).catch(() => false);
      const xmlHead = documentXml().slice(0, 400);
      check('word: the saved file keeps the page colour before the body, as Word does', /<w:document\b[^>]*><w:background w:color="DEEBF7"\/>/.test(xmlHead), xmlHead.slice(xmlHead.indexOf('<w:document'), xmlHead.indexOf('<w:document') + 160));
      // Design → Page Borders: a frame on every sheet, and w:pgBorders in the
      // section after the margins.
      await wait(300);
      const pickedFrame = await pick('Page Borders', 'Box — thin line');
      const frame = () => js(`(() => { const el = document.querySelector('.wd-pgborders'); if (!el) return null; const cs = getComputedStyle(el); return { top: cs.borderTopWidth + ' ' + cs.borderTopStyle, left: cs.borderLeftWidth + ' ' + cs.borderLeftStyle, inset: el.offsetLeft, count: document.querySelectorAll('.wd-pgborders').length }; })()`);
      const framed = await until(async () => /^[1-9][0-9.]*px solid$/.test((await frame())?.top || ''), 'the frame painted', 5000).catch(() => false);
      const frameNow = await frame();
      check('word: Design → Page Borders draws a box on every page, 24 pt in from the edge', pickedFrame === 'clicked' && framed === true && frameNow?.inset === 32 && frameNow?.count >= 1, `${pickedFrame}; ${JSON.stringify(frameNow)}`);
      await wait(400);
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'word-frame.png'), (await win.webContents.capturePage()).toPNG());
      await clickRibbon('Save');
      await until(() => { try { return /<w:pgBorders/.test(documentXml()); } catch { return false; } }, 'the page border to land in the file', 8000).catch(() => false);
      const sectPrXml = (() => { const m = /<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/.exec(documentXml()); return m ? m[0] : ''; })();
      // The fixture's section is empty, so the borders are its only child;
      // the engine test holds the order against the margins.
      check('word: the saved file keeps the page border in the section, all four sides as Word writes them',
        /<w:sectPr\b[^>]*><w:pgBorders w:offsetFrom="page"><w:top w:val="single" w:sz="6" w:space="24" w:color="auto"\/><w:left [^>]*\/><w:bottom [^>]*\/><w:right [^>]*\/><\/w:pgBorders>/.test(sectPrXml),
        sectPrXml.slice(0, 320));
      // Design → Watermark: DRAFT rises across the page; the header part
      // carries Word's shape and the reopened file reports it.
      await wait(300);
      const pickedMark = await pick('Watermark', 'DRAFT');
      const mark = () => js(`(() => { const el = document.querySelector('.wd-watermark'); if (!el) return null; const cs = getComputedStyle(el); return { text: el.textContent.trim(), colour: cs.color, turned: cs.transform !== 'none' }; })()`);
      const marked = await until(async () => (await mark())?.text === 'DRAFT', 'the watermark drawn', 5000).catch(() => false);
      await wait(400);
      const markNow = await mark();
      check('word: Design → Watermark writes DRAFT across the page, silver and rising', pickedMark === 'clicked' && marked === true && markNow?.colour === 'rgb(192, 192, 192)' && markNow?.turned === true, `${pickedMark}; ${JSON.stringify(markNow)}`);
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'word-watermark.png'), (await win.webContents.capturePage()).toPNG());
      await clickRibbon('Save');
      const markInFile = () => { try { return openDocx(fs.readFileSync(files.docx)).doc.doc.headerFooters().watermark || null; } catch { return null; } };
      await until(() => markInFile()?.text === 'DRAFT', 'the watermark to land in the file', 8000).catch(() => false);
      const savedMark = markInFile();
      const headerXml = (() => { try { const d = openDocx(fs.readFileSync(files.docx)).doc.doc; const b = d.headerFooters().headers.default; return b ? d.pkg.text(b.part) : ''; } catch { return ''; } })();
      check('word: the saved file keeps the watermark as Word does — a VML text path first in the header, with the namespaces it needs',
        savedMark?.text === 'DRAFT' && savedMark?.rotation === 315 && /<v:shape\b[^>]*fillcolor="silver"[^>]*>[\s\S]*?<v:textpath\b[^>]*string="DRAFT"/.test(headerXml) && /<w:hdr\b[^>]*xmlns:v="urn:schemas-microsoft-com:vml"/.test(headerXml),
        `${JSON.stringify(savedMark)}; ${headerXml.slice(0, 200)}`);
      // Layout → Line Numbers: a number beside every line, and lnNumType
      // in the section after the page borders.
      await wait(300);
      await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Layout')?.click(), 'tab'`);
      await wait(200);
      const pickedNumbers = await pick('Line Numbers', 'Continuous');
      const numbered = await until(() => js(`document.querySelectorAll('.wd-linenos span').length >= 3`), 'the numbers down the margin', 5000).catch(() => false);
      await wait(400);
      const nos = await js(`[...document.querySelectorAll('.wd-linenos span')].slice(0, 4).map((s) => s.textContent)`);
      const numberGap = await js(`(() => { const s = document.querySelector('.wd-linenos span'); const b = document.querySelector('.wd-page [data-block="0"]'); if (!s || !b) return null; return Math.round(b.getBoundingClientRect().left - s.getBoundingClientRect().right); })()`);
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'word-linenos.png'), (await win.webContents.capturePage()).toPNG());
      check('word: Layout → Line Numbers numbers every line down the left margin, a gap before the text', pickedNumbers === 'clicked' && numbered === true && JSON.stringify((nos || []).slice(0, 3)) === '["1","2","3"]' && numberGap > 8, `${pickedNumbers}; ${JSON.stringify(nos)}; gap ${numberGap}`);
      await clickRibbon('Save');
      await until(() => { try { return /<w:lnNumType/.test(documentXml()); } catch { return false; } }, 'the numbering to land in the file', 8000).catch(() => false);
      const sectPrNow = (() => { const m = /<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/.exec(documentXml()); return m ? m[0] : ''; })();
      check('word: the saved file keeps the line numbering in the section after the page borders, as Word does', /<\/w:pgBorders><w:lnNumType w:countBy="1" w:restart="continuous"\/>/.test(sectPrNow), sectPrNow.slice(0, 320));
      // Home → Multilevel list on two paragraphs, then Increase indent on
      // the second: 1. and 1.1., and the file says level 1 of a nine-level list.
      await wait(300);
      await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Home')?.click(), 'tab'`);
      await wait(200);
      // A real click on the paragraph's first words: once the page has been
      // edited, a DOM range put in by script no longer moves the caret.
      const caretTo = async (block) => {
        const at = await js(`(() => { const b = document.querySelector('.wd-page [data-block="${block}"]'); if (!b) return null; const q = b.getBoundingClientRect(); return { x: Math.round(q.left + 12), y: Math.round(q.top + q.height / 2) }; })()`);
        if (!at) return false;
        win.webContents.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
        win.webContents.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1 });
        return true;
      };
      const markerOf = (block) => js(`document.querySelector('.wd-page [data-block="${block}"] .wd-marker')?.textContent.trim() || null`);
      await caretTo(1);
      await wait(300);
      const outlined = await clickRibbon('Multilevel list');
      await until(async () => (await markerOf(1)) === '1.', 'the first number', 5000).catch(() => false);
      await wait(300);
      await caretTo(2);
      await wait(300);
      await clickRibbon('Multilevel list');
      await until(async () => (await markerOf(2)) === '2.', 'the second number', 5000).catch(() => false);
      await wait(300);
      const deeper = await clickRibbon('Increase indent');
      const nested = await until(async () => (await markerOf(2)) === '1.1.', 'the nested number', 5000).catch(() => false);
      await wait(400);
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'word-multilevel.png'), (await win.webContents.capturePage()).toPNG());
      check('word: Home → Multilevel list numbers 1., 2., and Increase indent makes the second 1.1.', outlined === 'clicked' && deeper === 'clicked' && nested === true && (await markerOf(1)) === '1.', `${outlined} ${deeper}; markers ${await markerOf(1)} ${await markerOf(2)}`);
      await clickRibbon('Save');
      await until(() => { try { return /<w:ilvl w:val="1"\/>/.test(documentXml()); } catch { return false; } }, 'the level to land in the file', 8000).catch(() => false);
      const saved2 = (() => { try { const d = openDocx(fs.readFileSync(files.docx)).doc.doc; return { pPr: d.editParagraph(2).pPr || '', numbering: d.pkg.has('word/numbering.xml') ? d.pkg.text('word/numbering.xml') : '' }; } catch { return { pPr: '', numbering: '' }; } })();
      const levelsInFile = (saved2.numbering.match(/<w:lvl w:ilvl="/g) || []).length;
      check('word: the saved file keeps the paragraph at level 1 of a nine-level list whose second level says 1.1., as Word writes it',
        /<w:numPr><w:ilvl w:val="1"\/><w:numId w:val="\d+"\/><\/w:numPr>/.test(saved2.pPr) && levelsInFile >= 9 && /<w:lvl w:ilvl="1">[\s\S]*?<w:lvlText w:val="%1\.%2\."\/>/.test(saved2.numbering),
        `pPr ${saved2.pPr.slice(0, 160)}; levels ${levelsInFile}`);
      // The list off again and the file saved as it was: the hand-driven
      // checks later open this same report.docx and measure block 1's first
      // span, which must be its words and not a list marker.
      await js(`(async () => { const all = await window.rutbaOffice.doc.sessions({}); const mine = all.filter((s) => s.kind === 'doc').pop(); await window.rutbaOffice.doc.apply({ id: mine.id, ops: [ { op: 'setSelection', anchor: { block: 1, offset: 0 }, focus: { block: 2, offset: 0 } }, { op: 'setParagraphFormat', delta: { list: null } } ] }); return 1; })()`);
      await wait(300);
      await clickRibbon('Save');
      await until(() => { try { return !/<w:numPr>/.test(documentXml()); } catch { return false; } }, 'the list off again in the file', 8000).catch(() => false);
      check('word: the look checks leave the paragraphs out of the list again', !/<w:numPr>/.test(documentXml()), 'no numPr left in the file');
      // Home → Sort: the two body paragraphs A to Z, the file following, then
      // back again for the checks after this one.
      const selectBody = () => js(`(async () => { const all = await window.rutbaOffice.doc.sessions({}); const mine = all.filter((s) => s.kind === 'doc').pop(); await window.rutbaOffice.doc.apply({ id: mine.id, ops: [{ op: 'setSelection', anchor: { block: 1, offset: 0 }, focus: { block: 2, offset: 0 } }] }); return 1; })()`);
      // The block's own words: a list marker the page still shows from the
      // list taken off through the service (the window's model lags a
      // direct apply) is not part of them.
      const blockText = (i) => js(`(() => { const b = document.querySelector('.wd-page [data-block="${i}"]'); if (!b) return ''; const c = b.cloneNode(true); c.querySelectorAll('.wd-marker').forEach((m) => m.remove()); return c.textContent.trim(); })()`);
      // Earlier checks may have edited the words, so the order is judged
      // against what is on the page now, the way the engine judges it.
      await wait(300);
      const original = [await blockText(1), await blockText(2)];
      const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });
      const ascending = [...original].sort((a, b) => collator.compare(a, b));
      const startsAs = async (i, text) => (await blockText(i)).slice(0, 20) === text.slice(0, 20);
      await selectBody();
      await wait(300);
      const sortedAZ = await pick('Sort paragraphs', 'A to Z');
      const sorted = await until(async () => (await startsAs(1, ascending[0])) && (await startsAs(2, ascending[1])), 'the paragraphs sorted', 5000).catch(() => false);
      check('word: Home → Sort puts the selected paragraphs A to Z, each moving whole', sortedAZ === 'clicked' && sorted === true && original[0] !== ascending[0], `${sortedAZ}; block 1 ${JSON.stringify((await blockText(1)).slice(0, 30))}, block 2 ${JSON.stringify((await blockText(2)).slice(0, 30))}; was ${JSON.stringify(original.map((t) => t.slice(0, 20)))}`);
      await clickRibbon('Save');
      const orderInFile = () => { try { const d = openDocx(fs.readFileSync(files.docx)); return [1, 2].map((i) => d.block(i).text.slice(0, 20)); } catch { return []; } };
      await until(() => orderInFile()[0] === ascending[0].slice(0, 20), 'the order in the file', 8000).catch(() => false);
      check('word: the saved file keeps the sorted order', orderInFile()[0] === ascending[0].slice(0, 20) && orderInFile()[1] === ascending[1].slice(0, 20), JSON.stringify(orderInFile()));
      // And back in the original order, whichever way that is, for the checks after this one.
      await wait(300);
      await selectBody();
      await wait(300);
      await pick('Sort paragraphs', original[0] === ascending[0] ? 'A to Z' : 'Z to A');
      await until(async () => startsAs(1, original[0]), 'the paragraphs back', 5000).catch(() => false);
      await clickRibbon('Save');
      await until(() => orderInFile()[0] === original[0].slice(0, 20), 'the order back in the file', 8000).catch(() => false);
      const complaints = await errorsIn(win);
      check('word: the shading and border checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('word: the shading and border checks ran', false, err.message);
    }
  };

  /* ── Word: drop cap — Insert → Drop Cap frames the first letter ──────── */
  //
  // Word's own trick: the paragraph's first letter becomes its own paragraph,
  // framed (`w:framePr`) to stand three lines tall and sized to match, the
  // body paragraph flowing round it; None merges the letter back in. Run
  // alone with RUTBA_VERIFY_ONLY=dropcap.
  const wordDropCap = async () => {
    try {
      const win = await open('word', files.docx);
      const js = (code) => win.webContents.executeJavaScript(code);
      const session = sessionFor('doc');
      const model = () => doc.model({ id: session.id });

      await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="1"]'))`), 'the second paragraph', 8000);
      // The caret into the second paragraph, the way wordBookmarks puts it there.
      await js(`(() => {
        const b = document.querySelector('.wd-page [data-block="1"]');
        b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }));
        const r = document.createRange(); r.selectNodeContents(b); r.collapse(true);
        const s = getSelection(); s.removeAllRanges(); s.addRange(r);
        // The window learns the caret on mouseup, as a click gives it.
        document.querySelector('.wd-page').dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
        return 1;
      })()`);
      await wait(300);
      await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click(), 'tab'`);
      await wait(200);
      const clickRibbon = (title) => js(`(() => {
        const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}));
        if (!b) return 'no button ' + ${JSON.stringify(title)};
        b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
        b.click();
        return 'clicked';
      })()`);
      const pick = async (button, label) => {
        const clicked = await clickRibbon(button);
        await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}))`), `the ${button} menu`, 3000).catch(() => {});
        await js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)})?.click(); return 1; })()`);
        return clicked;
      };
      const original = (model().blocks[1] || {}).text || '';

      const droppedClicked = await pick('Drop Cap', 'Dropped');
      const inModel = await until(
        () => {
          const b = model().blocks;
          const d = b.findIndex((x) => x.dropCap?.lines === 3);
          return d >= 0 && b[d].text.length === 1 && b[d + 1]?.text === original.slice(1);
        },
        'the drop cap split in the model',
        5000
      ).catch(() => false);
      const blocksNow = model().blocks;
      const dropIndex = blocksNow.findIndex((b) => b.dropCap?.lines === 3);
      check('word: Insert → Drop Cap splits the first letter into its own three-line paragraph, the rest staying put',
        droppedClicked === 'clicked' && inModel === true,
        `${droppedClicked}; ${JSON.stringify(blocksNow.slice(0, 4).map((b) => [b.dropCap || null, (b.text || '').slice(0, 12)]))}`);

      // The page: the letter's own block floats left, well past a normal
      // letter's size, and the body's first line starts clear of it.
      await wait(300);
      const paint = () => js(`(() => {
        const drop = document.querySelector('.wd-dropcap');
        const body = document.querySelector('.wd-page [data-block="${dropIndex + 1}"]');
        if (!drop || !body) return null;
        const cs = getComputedStyle(drop);
        const r = document.createRange();
        r.selectNodeContents(body);
        const rects = [...r.getClientRects()];
        const bodyLeft = rects.length ? rects[0].left : null;
        // The size is on the letter's run, not the paragraph's box.
        const letter = drop.querySelector('span') || drop;
        return { float: cs.float, fontSize: parseFloat(getComputedStyle(letter).fontSize), bodyLeft, dropRight: drop.getBoundingClientRect().right };
      })()`);
      const painted = await paint();
      check('word: the drop cap floats left of the page at a size past an ordinary letter, and the body starts clear of it',
        painted?.float === 'left' && painted?.fontSize > 30 && painted?.bodyLeft != null && painted.bodyLeft >= painted.dropRight - 1,
        JSON.stringify(painted));

      await wait(300);
      await clickRibbon('Save');
      const dropParagraph = () => {
        try {
          return openDocx(fs.readFileSync(files.docx)).doc.doc.editParagraph(dropIndex);
        } catch {
          return null;
        }
      };
      await until(() => /w:framePr/.test(dropParagraph()?.pPr || ''), 'the frame to land in the file', 8000).catch(() => false);
      const dropSaved = dropParagraph();
      const pPr = dropSaved?.pPr || '';
      check('word: the saved file frames the letter\'s paragraph exactly as Word does, its run sized',
        /<w:framePr w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text"\/>/.test(pPr) && /<w:sz w:val="\d+"\/>/.test(dropSaved?.xml || ''),
        `${pPr.slice(0, 200)}`);

      // None: the letter merges back into the body, the original words whole again.
      await wait(300);
      await pick('Drop Cap', 'None');
      const merged = await until(
        () => (model().blocks[1] || {}).text === original,
        'the paragraph merged back',
        5000
      ).catch(() => false);
      check('word: Drop Cap → None merges the letter back into its paragraph, the words unchanged', merged === true, JSON.stringify((model().blocks[1] || {}).text));
      // Saved again, so the fixture the later blocks open is the one they expect.
      await clickRibbon('Save');
      await until(() => { try { return !/<w:framePr/.test(openDocx(fs.readFileSync(files.docx)).doc.doc.xml); } catch { return false; } }, 'the frame gone from the file', 8000).catch(() => {});

      const complaints = await errorsIn(win);
      check('word: the drop cap checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('word: the drop cap checks ran', false, err.message);
    }
  };

  /* ── Word: bookmarks — Insert → Bookmark names a span of paragraphs ──── */
  //
  // Word's older, position-based anchor: a name on a paragraph (or a run of
  // them), so Go To finds the spot again and — once the engine writes REF
  // fields — a cross-reference will too. Run alone with RUTBA_VERIFY_ONLY=bookmarks.
  const wordBookmarks = async () => {
    try {
      const win = await open('word', files.docx);
      const js = (code) => win.webContents.executeJavaScript(code);
      const session = sessionFor('doc');
      const model = () => doc.model({ id: session.id });

      await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="1"]'))`), 'the second paragraph', 8000);
      // The caret into the second paragraph, the way wordLook puts it into the first.
      await js(`(() => {
        const b = document.querySelector('.wd-page [data-block="1"]');
        b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }));
        const r = document.createRange(); r.selectNodeContents(b); r.collapse(true);
        const s = getSelection(); s.removeAllRanges(); s.addRange(r);
        // The window learns the caret on mouseup, as a click gives it.
        document.querySelector('.wd-page').dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
        return 1;
      })()`);
      await wait(300);
      await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click(), 'tab'`);
      await wait(200);
      const clickRibbon = (title) => js(`(() => {
        const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}));
        if (!b) return 'no button ' + ${JSON.stringify(title)};
        b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
        b.click();
        return 'clicked';
      })()`);
      const closeDialog = () => js(`(() => { [...document.querySelectorAll('.rw-dialog button')].find((b) => b.textContent.trim() === 'Close')?.click(); return 1; })()`);
      const pickRow = () => js(`(() => { [...document.querySelectorAll('.wd-bookmark-row')].find((r) => r.textContent.includes('Summary'))?.click(); return 1; })()`);

      const pressed = await clickRibbon('Bookmark');
      await until(() => js(`Boolean(document.querySelector('.wd-bookmark-name'))`), 'the Bookmark dialog', 5000);
      await js(`(() => { const el = document.querySelector('.wd-bookmark-name'); const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set; setter.call(el, 'Summary'); el.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
      await wait(150);
      await js(`(() => { document.querySelector('.wd-bookmark-add')?.click(); return 1; })()`);

      const added = await until(
        () => (model().bookmarks || []).some((b) => b.name === 'Summary' && b.from === 1 && b.to === 1),
        'the bookmark in the model',
        5000
      ).catch(() => false);
      const m1 = model();
      check(
        'word: Insert → Bookmark names the selected paragraph, in the model, and it stays editable',
        pressed === 'clicked' && added === true && m1.blocks[1]?.structural === false,
        `${pressed}; bookmarks ${JSON.stringify(m1.bookmarks)}; block 1 structural=${m1.blocks[1]?.structural}`
      );

      const rowShown = await js(`Boolean([...document.querySelectorAll('.wd-bookmark-row')].find((r) => r.textContent.includes('Summary')))`);
      check('word: the new bookmark shows in the dialog', rowShown === true, `row present: ${rowShown}`);

      // The ribbon sits under the dialog; close it before reaching for Save.
      await closeDialog();
      await wait(200);
      await clickRibbon('Save');
      const paragraphXml = () => {
        try {
          return openDocx(fs.readFileSync(files.docx)).doc.doc.paragraph(1).xml || '';
        } catch {
          return '';
        }
      };
      await until(() => /w:bookmarkStart/.test(paragraphXml()), 'the bookmark to land in the file', 8000).catch(() => false);
      const pXml = paragraphXml();
      const startMatch = /<w:p\b[^>]*>(<w:pPr>[\s\S]*?<\/w:pPr>)?<w:bookmarkStart w:id="(\d+)" w:name="Summary"\/>/.exec(pXml);
      const id = startMatch ? startMatch[2] : null;
      const endOk = id !== null && new RegExp('<w:bookmarkEnd w:id="' + id + '"/></w:p>$').test(pXml);
      check(
        'word: the saved file writes the bookmark as Word does — bookmarkStart right after pPr, bookmarkEnd right before the paragraph closes',
        Boolean(startMatch) && endOk,
        pXml.slice(0, 240)
      );

      // Go To: reopen, select the row, press Go To — the selection lands on
      // the bookmarked paragraph.
      await clickRibbon('Bookmark');
      await until(() => js(`Boolean(document.querySelector('.wd-bookmark-row'))`), 'the bookmark row', 5000);
      await pickRow();
      await wait(150);
      await js(`(() => { document.querySelector('.wd-bookmark-goto')?.click(); return 1; })()`);
      await wait(300);
      const afterGoto = model();
      check('word: Go To selects the bookmarked paragraph', afterGoto.selection?.focus?.block === 1, JSON.stringify(afterGoto.selection));

      // Delete: reopen, select, Delete — the bookmark is gone from the model.
      await clickRibbon('Bookmark');
      await until(() => js(`Boolean(document.querySelector('.wd-bookmark-row'))`), 'the bookmark row again', 5000);
      await pickRow();
      await wait(150);
      await js(`(() => { document.querySelector('.wd-bookmark-delete')?.click(); return 1; })()`);
      const removed = await until(() => (model().bookmarks || []).length === 0, 'the bookmark removed', 5000).catch(() => false);
      check('word: Delete removes the bookmark', removed === true, JSON.stringify(model().bookmarks));
      await closeDialog();

      const complaints = await errorsIn(win);
      check('word: the bookmark checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('word: the bookmark checks ran', false, err.message);
    }
  };

  /* ── Word: cross-reference — Insert → Cross-reference writes a REF field ── */
  //
  // A REF field to a bookmark: the bookmark's own words at the caret, kept as
  // a field in the file, refreshed on demand by Update Fields (F9) — and the
  // paragraph carrying it stays editable, the same promise a bookmark makes.
  // Run alone with RUTBA_VERIFY_ONLY=xref.
  const wordCrossRef = async () => {
    try {
      const win = await open('word', files.docx);
      const js = (code) => win.webContents.executeJavaScript(code);
      const session = sessionFor('doc');
      const model = () => doc.model({ id: session.id });

      const clickRibbon = (title) => js(`(() => {
        const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}));
        if (!b) return 'no button ' + ${JSON.stringify(title)};
        b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
        b.click();
        return 'clicked';
      })()`);
      const closeDialog = () => js(`(() => { [...document.querySelectorAll('.rw-dialog button')].find((b) => b.textContent.trim() === 'Close')?.click(); return 1; })()`);
      const clickTab = (name) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(), 'tab'`);
      // A collapsed caret at the given end of a block, and the mouseup the
      // window learns it from — exactly as wordBookmarks places one.
      const caretIn = (block, atEnd) => js(`(() => {
        const b = document.querySelector('.wd-page [data-block="${block}"]');
        b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }));
        const r = document.createRange(); r.selectNodeContents(b); r.collapse(${atEnd ? 'false' : 'true'});
        const s = getSelection(); s.removeAllRanges(); s.addRange(r);
        document.querySelector('.wd-page').dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
        return 1;
      })()`);

      await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="2"]'))`), 'the third paragraph', 8000);

      // The bookmark on the second paragraph, exactly as wordBookmarks makes one.
      await caretIn(1, false);
      await wait(200);
      await clickTab('Insert');
      await wait(200);
      await clickRibbon('Bookmark');
      await until(() => js(`Boolean(document.querySelector('.wd-bookmark-name'))`), 'the Bookmark dialog', 5000);
      await js(`(() => { const el = document.querySelector('.wd-bookmark-name'); const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set; setter.call(el, 'Summary'); el.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
      await wait(150);
      await js(`(() => { document.querySelector('.wd-bookmark-add')?.click(); return 1; })()`);
      // Waited on the PAGE's own state, not just the engine's: the engine
      // has the bookmark the moment the main process answers, but the page
      // learns it on its own following round trip, and the Cross-reference
      // dialog reads the page's model — opening it before that lands is
      // exactly the "No bookmarks" empty state, wrongly.
      await until(() => js(`Boolean(document.querySelector('.wd-bookmark-row'))`), 'the bookmark to show in the dialog', 5000);
      await closeDialog();
      await wait(200);

      // The caret at the END of the third paragraph, then Insert → Cross-reference.
      await caretIn(2, true);
      await wait(200);
      const pressed = await clickRibbon('Cross-reference');
      await until(() => js(`Boolean(document.querySelector('.wd-xref-row'))`), 'the Cross-reference dialog', 5000);
      await js(`(() => { [...document.querySelectorAll('.wd-xref-row')].find((r) => r.textContent.includes('Summary'))?.click(); return 1; })()`);
      await wait(150);
      await js(`(() => { document.querySelector('.wd-xref-insert')?.click(); return 1; })()`);

      const summaryText = model().blocks[1]?.text || '';
      const inserted = await until(
        () => (model().blocks[2]?.runs || []).some((r) => r.field?.kind === 'ref' && r.field?.name === 'Summary' && r.text === summaryText),
        'the REF field in the model',
        5000
      ).catch(() => false);
      check(
        "word: Insert → Cross-reference writes a REF field with the bookmark's words, at the caret",
        pressed === 'clicked' && inserted === true,
        `${pressed}; block 2 runs ${JSON.stringify(model().blocks[2]?.runs)}`
      );

      const fieldShown = await js(`Boolean([...document.querySelectorAll('[data-block="2"] .wd-field')].find((s) => s.textContent === ${JSON.stringify(summaryText)}))`);
      check('word: the page shows the field, shaded grey', fieldShown === true, `field shown: ${fieldShown}`);

      await clickRibbon('Save');
      const paragraphXml = () => {
        try {
          return openDocx(fs.readFileSync(files.docx)).doc.doc.paragraph(2).xml || '';
        } catch {
          return '';
        }
      };
      await until(() => /w:fldSimple/.test(paragraphXml()), 'the field to land in the file', 8000).catch(() => false);
      const pXml = paragraphXml();
      const hasField = pXml.includes('<w:fldSimple w:instr=" REF Summary \\h "><w:r>') && pXml.includes(summaryText);
      check(
        "word: the saved file writes the REF field with the words, and the paragraph stays editable",
        hasField && model().blocks[2]?.structural === false,
        pXml.slice(-320)
      );

      // Edit the bookmarked paragraph, then Update Fields — the field picks it up.
      await caretIn(1, false);
      await wait(200);
      await win.webContents.insertText('X');
      await wait(200);
      await clickTab('References');
      await wait(200);
      await clickRibbon('Update Fields');
      const refreshed = await until(
        () => (model().blocks[2]?.runs || []).some((r) => r.field?.kind === 'ref' && String(r.text).startsWith('X')),
        'the field refreshed from the edited bookmark',
        5000
      ).catch(() => false);
      check("word: Update Fields refreshes a REF to its bookmark's current words", refreshed === true, JSON.stringify(model().blocks[2]?.runs));

      // Remove the bookmark, Update again — the field reads Word's own error text.
      await clickTab('Insert');
      await wait(200);
      await clickRibbon('Bookmark');
      await until(() => js(`Boolean(document.querySelector('.wd-bookmark-row'))`), 'the bookmark row', 5000);
      await js(`(() => { [...document.querySelectorAll('.wd-bookmark-row')].find((r) => r.textContent.includes('Summary'))?.click(); return 1; })()`);
      await wait(150);
      await js(`(() => { document.querySelector('.wd-bookmark-delete')?.click(); return 1; })()`);
      await until(() => (model().bookmarks || []).length === 0, 'the bookmark removed', 5000).catch(() => false);
      await closeDialog();
      await wait(200);

      await clickTab('References');
      await wait(200);
      await clickRibbon('Update Fields');
      const errored = await until(
        () => (model().blocks[2]?.runs || []).some((r) => r.field?.kind === 'ref' && r.text === 'Error! Reference source not found.'),
        'the field to read the missing-bookmark error',
        5000
      ).catch(() => false);
      check("word: Update Fields writes Word's own error text once the bookmark is gone", errored === true, JSON.stringify(model().blocks[2]?.runs));

      const complaints = await errorsIn(win);
      check('word: the cross-reference checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('word: the cross-reference checks ran', false, err.message);
    }
  };

  /* ── Word: captions — Insert → Caption writes a SEQ field ────────────── */
  //
  // A labelled, numbered paragraph after the caret's: the label, a SEQ
  // field — a COMPLEX field, the shape Word itself writes for a caption —
  // then the words. A second caption of the same label numbers itself from
  // where it lands, not from a counter the dialog remembers. Run alone with
  // RUTBA_VERIFY_ONLY=captions.
  const wordCaptions = async () => {
    try {
      const win = await open('word', files.docx);
      const js = (code) => win.webContents.executeJavaScript(code);
      const session = sessionFor('doc');
      const model = () => doc.model({ id: session.id });

      const clickRibbon = (title) => js(`(() => {
        const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}));
        if (!b) return 'no button ' + ${JSON.stringify(title)};
        b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
        b.click();
        return 'clicked';
      })()`);
      const clickTab = (name) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(), 'tab'`);
      // A collapsed caret at the given end of a block, the way wordBookmarks
      // and wordCrossRef both place one.
      const caretIn = (block, atEnd) => js(`(() => {
        const b = document.querySelector('.wd-page [data-block="${block}"]');
        b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }));
        const r = document.createRange(); r.selectNodeContents(b); r.collapse(${atEnd ? 'false' : 'true'});
        const s = getSelection(); s.removeAllRanges(); s.addRange(r);
        document.querySelector('.wd-page').dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
        return 1;
      })()`);
      const setCaptionText = (text) => js(`(() => { const el = document.querySelector('.wd-caption-text'); const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set; setter.call(el, ${JSON.stringify(text)}); el.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);

      await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="2"]'))`), 'the third paragraph', 8000);

      await caretIn(2, true);
      await wait(200);
      await clickTab('References');
      await wait(200);

      const pressed = await clickRibbon('Insert Caption');
      await until(() => js(`Boolean(document.querySelector('.wd-caption-text'))`), 'the Caption dialog', 5000);
      await setCaptionText('a diagram of the pipeline');
      await wait(150);
      await js(`(() => { document.querySelector('.wd-caption-insert')?.click(); return 1; })()`);

      const firstText = await until(
        () => (model().blocks || []).some((b) => b.text === 'Figure 1: a diagram of the pipeline'),
        'the first caption in the model',
        5000
      ).catch(() => false);
      check(
        'word: Insert → Caption writes a labelled, numbered paragraph at the caret',
        pressed === 'clicked' && firstText === true,
        `${pressed}; blocks ${JSON.stringify((model().blocks || []).map((b) => b.text))}`
      );

      const onPage = await js(`document.querySelector('.wd-page')?.textContent.includes('Figure 1: a diagram of the pipeline')`);
      check('word: the page reads the caption', onPage === true, `on page: ${onPage}`);

      const shaded = await js(`Boolean([...document.querySelectorAll('.wd-field')].find((s) => s.textContent === '1'))`);
      check("word: the caption's SEQ number is shaded like any other field", shaded === true, `shaded: ${shaded}`);

      // A second caption, at the end of the document now — numbered from
      // where it lands.
      const lastBlock = () => (model().blocks || []).length - 1;
      await caretIn(lastBlock(), true);
      await wait(200);
      await clickRibbon('Insert Caption');
      await until(() => js(`Boolean(document.querySelector('.wd-caption-text'))`), 'the Caption dialog again', 5000);
      await setCaptionText('a screenshot of the report');
      await wait(150);
      await js(`(() => { document.querySelector('.wd-caption-insert')?.click(); return 1; })()`);

      const secondText = await until(
        () => (model().blocks || []).some((b) => b.text === 'Figure 2: a screenshot of the report'),
        'the second caption in the model',
        5000
      ).catch(() => false);
      check(
        'word: a second caption of the same label numbers itself 2',
        secondText === true,
        JSON.stringify((model().blocks || []).map((b) => b.text))
      );

      await clickRibbon('Save');
      const bodyXml = () => {
        try {
          const engine = openDocx(fs.readFileSync(files.docx)).doc.doc;
          return engine.paragraphs().map((p) => p.xml).join('');
        } catch {
          return '';
        }
      };
      await until(() => /SEQ Figure/.test(bodyXml()), 'the SEQ field to land in the file', 8000).catch(() => false);
      const xml = bodyXml();
      const hasSeq = xml.includes('<w:instrText xml:space="preserve"> SEQ Figure \\* ARABIC </w:instrText>')
        && /<w:fldChar w:fldCharType="begin"\/>/.test(xml)
        && /<w:fldChar w:fldCharType="separate"\/>/.test(xml)
        && /<w:fldChar w:fldCharType="end"\/>/.test(xml);
      check("word: the saved file writes the SEQ field as a complex field, Word's own shape", hasSeq, xml.slice(-500));

      const complaints = await errorsIn(win);
      check('word: the caption checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('word: the caption checks ran', false, err.message);
    }
  };

  /* ── Word: text effects — outline, shadow, glow ────────────────────── */
  //
  // Home → the "A" button: Outline hollows the selected words, Shadow casts
  // one behind them, Glow rings them in a colour — each at once on the
  // page, and the saved file keeps them where Word does: the toggles after
  // strike, the glow last with its own namespace. Run alone with
  // RUTBA_VERIFY_ONLY=effects.
  const wordEffects = async () => {
    try {
      const win = await open('word', files.docx);
      const js = (code) => win.webContents.executeJavaScript(code);
      await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="1"]'))`), 'the second paragraph', 8000);
      // The first word of the second paragraph, selected the way wordLook
      // puts the caret in — a real mousedown, then a DOM range — but spanning
      // "The" rather than collapsed, so the format lands on the run at once.
      await js(`(() => {
        const b = document.querySelector('.wd-page [data-block="1"]');
        b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }));
        // The first three characters, whatever runs they fall in: a paragraph
        // an earlier block merged back can start with a one-letter run.
        const walker = document.createTreeWalker(b, NodeFilter.SHOW_TEXT);
        const first = walker.nextNode();
        let node = first; let left = 3;
        while (node && node.nodeValue.length < left) { left -= node.nodeValue.length; node = walker.nextNode(); }
        const r = document.createRange();
        r.setStart(first, 0);
        r.setEnd(node || first, node ? left : first.nodeValue.length);
        const s = getSelection(); s.removeAllRanges(); s.addRange(r);
        // The window learns the caret on mouseup, as a click gives it.
        document.querySelector('.wd-page').dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
        return 1;
      })()`);
      await wait(300);
      const clickRibbon = (title) => js(`(() => {
        const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}));
        if (!b) return 'no button ' + ${JSON.stringify(title)};
        b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
        b.click();
        return 'clicked';
      })()`);
      const pick = async (button, label) => {
        const clicked = await clickRibbon(button);
        await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}))`), `the ${button} menu`, 3000).catch(() => {});
        await js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)})?.click(); return 1; })()`);
        return clicked;
      };
      const model = () => doc.model({ id: sessionFor('doc').id });
      const spanStyle = () => js(`(() => { const b = document.querySelector('.wd-page [data-block="1"]'); const s = b?.querySelector('span'); if (!s) return null; const cs = getComputedStyle(s); return { stroke: cs.webkitTextStrokeWidth, shadow: cs.textShadow }; })()`);

      const pickedOutline = await pick('Text effects', 'Outline');
      const outlineOn = await until(() => model().format.outline === true, 'the outline in the model', 5000).catch(() => false);
      const stroked = await until(async () => parseFloat((await spanStyle())?.stroke || '0') > 0, 'the outline painted', 5000).catch(() => false);
      check('word: Home → Text effects → Outline turns the run hollow, in the model and on the page',
        pickedOutline === 'clicked' && outlineOn === true && stroked === true, `${pickedOutline}; outline ${JSON.stringify(model().format.outline)}; ${JSON.stringify(await spanStyle())}`);

      await wait(200);
      const pickedShadow = await pick('Text effects', 'Shadow');
      const shadowOn = await until(() => model().format.shadow === true, 'the shadow in the model', 5000).catch(() => false);
      check('word: Home → Text effects → Shadow marks the run in the model',
        pickedShadow === 'clicked' && shadowOn === true, `${pickedShadow}; shadow ${JSON.stringify(model().format.shadow)}`);

      await wait(200);
      const pickedGlow = await pick('Text effects', 'Glow: gold');
      const glowOn = await until(() => model().format.glow?.colour === 'FFC000', 'the glow in the model', 5000).catch(() => false);
      const glowed = await until(async () => (await spanStyle())?.shadow.includes('rgb(255, 192, 0)'), 'the glow painted', 5000).catch(() => false);
      check('word: Home → Text effects → Glow: gold marks the run in the model and glows gold on the page',
        pickedGlow === 'clicked' && glowOn === true && glowed === true, `${pickedGlow}; glow ${JSON.stringify(model().format.glow)}; ${JSON.stringify(await spanStyle())}`);

      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'word-effects.png'), (await win.webContents.capturePage()).toPNG());

      await wait(300);
      await clickRibbon('Save');
      const pInFile = () => openDocx(fs.readFileSync(files.docx)).doc.doc.editParagraph(1).xml || '';
      await until(() => { try { return /<w:outline\/>/.test(pInFile()); } catch { return false; } }, 'the effects to land in the file', 8000).catch(() => false);
      const pXml = pInFile();
      check('word: the saved file keeps the outline, the shadow and the glow on the run, Word\'s own way',
        /<w:outline\/>/.test(pXml) && /<w:shadow\/>/.test(pXml) && /<w14:glow\b[^>]*w14:rad="50800"[^>]*>/.test(pXml) && /FFC000/.test(pXml),
        pXml.slice(0, 400));

      // Clear effects, and back to nothing — the model agrees at the caret.
      await wait(300);
      await pick('Text effects', 'Clear effects');
      const cleared = await until(() => {
        const f = model().format;
        return f.outline === false && f.shadow === false && f.glow === null;
      }, 'the effects cleared', 5000).catch(() => false);
      check('word: Home → Text effects → Clear effects turns all three off again', cleared === true, JSON.stringify(model().format));

      const complaints = await errorsIn(win);
      check('word: the text effects checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('word: the text effects checks ran', false, err.message);
    }
  };

  /* ── Presentation: the shape clipboard ──────────────────────────────── */
  //
  // Copy on a selected shape, Paste on another slide: the shape appears
  // there with the same words and a fresh id; Cut takes it away; the saved
  // file carries the paste. Run alone with RUTBA_VERIFY_ONLY=clip.
  /* ── Presentation: the footer band ──────────────────────────────────── */
  //
  // Insert → Slide Number opens the Header & Footer dialog with the number
  // ticked; the footer's words are typed and Apply to All puts both on every
  // slide, the second slide saying 2. The dialog then opens with the slide's
  // own settings, and Apply changes that slide alone.
  const slideFooter = async () => {
    try {
      const win = await open('slides', files.pptx);
      const js = (code) => win.webContents.executeJavaScript(code);
      const model = (slide) => doc.model({ id: sessionFor('deck').id, slide });
      await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 2`), 'the slide sorter', 8000);
      const clickRibbon = (title) => js(`(() => {
        const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}));
        if (!b) return 'no button ' + ${JSON.stringify(title)};
        b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
        b.click();
        return 'clicked';
      })()`);
      const ofType = (slide, type) => model(slide).slide.shapes.find((s) => s.placeholder?.type === type) || null;
      const words = (s) => (s?.text?.paragraphs || []).map((p) => (p.runs || []).map((r) => r.text).join('')).join('');
      await js(`(() => { document.querySelectorAll('.sl-thumb')[0]?.click(); return 1; })()`);
      await until(() => js(`document.querySelectorAll('.sl-thumb')[0]?.classList.contains('active')`), 'the first slide', 4000).catch(() => {});
      await wait(300);
      const hitsBefore = await js(`document.querySelectorAll('.sl-hit').length`);
      await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click(), 'tab'`);
      await wait(300);
      const opened = await clickRibbon('Slide Number');
      await until(() => js(`Boolean(document.querySelector('.sl-hf-number'))`), 'the Header & Footer dialog', 4000).catch(() => {});
      const ticked = await js(`document.querySelector('.sl-hf-number')?.checked`);
      await js(`(() => { const on = document.querySelector('.sl-hf-footer-on'); if (on && !on.checked) on.click(); return 1; })()`);
      await wait(150);
      await js(`(() => { const el = document.querySelector('.sl-hf-footer'); if (!el) return false; const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set; setter.call(el, 'Rutba Office beta'); el.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
      await wait(150);
      await js(`(() => { document.querySelector('.sl-hf-all')?.click(); return 1; })()`);
      const landed = await until(() => words(ofType(1, 'ftr')) === 'Rutba Office beta' && words(ofType(1, 'sldNum')) === '2' && words(ofType(0, 'sldNum')) === '1', 'the footer on every slide', 6000).catch(() => false);
      check('slides: Insert → Slide Number opens Header & Footer with the number ticked, and Apply to All puts the number and the words on every slide',
        opened === 'clicked' && ticked === true && landed === true,
        `${opened}; ticked ${ticked}; slide 2 footer ${JSON.stringify(words(ofType(1, 'ftr')))} number ${JSON.stringify(words(ofType(1, 'sldNum')))}; slide 1 number ${JSON.stringify(words(ofType(0, 'sldNum')))}`);

      // Drawn on the stage: two more shapes to click, and the words on the slide.
      const drawn = await until(() => js(`document.querySelectorAll('.sl-hit').length === ${hitsBefore} + 2`), 'the placeholders on the stage', 5000).catch(() => false);
      await wait(400);
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'slides-footer.png'), (await win.webContents.capturePage()).toPNG());
      const onStage = await js(`(document.querySelector('.sl-stage')?.textContent || '').includes('Rutba Office beta')`);
      check('slides: the footer band is drawn on the slide — two more shapes, the words among them', drawn === true && onStage === true, `hits ${hitsBefore} → ${await js(`document.querySelectorAll('.sl-hit').length`)}; on stage ${onStage}`);

      await wait(300);
      await clickRibbon('Save');
      const saved = () => {
        try {
          const d = Deck.open(fs.readFileSync(files.pptx));
          return { f1: words(d.slide(0).shapes.find((s) => s.placeholder?.type === 'ftr')), n2: words(d.slide(1).shapes.find((s) => s.placeholder?.type === 'sldNum')), xml: d.pkg.text(d.slideParts[1].part) };
        } catch { return null; }
      };
      await until(() => saved()?.n2 === '2', 'the footer in the file', 8000).catch(() => false);
      const inFile = saved();
      check('slides: the saved file carries the placeholders as PowerPoint writes them, the number a field',
        inFile?.f1 === 'Rutba Office beta' && inFile?.n2 === '2' && /<p:ph type="sldNum" sz="quarter" idx="12"\/>/.test(inFile?.xml || '') && /<a:fld id="\{[0-9A-F-]{36}\}" type="slidenum">/.test(inFile?.xml || '') && /<p:ph type="ftr" sz="quarter" idx="11"\/>/.test(inFile?.xml || ''),
        `footer ${JSON.stringify(inFile?.f1)} number ${JSON.stringify(inFile?.n2)}; ${(inFile?.xml || '').slice((inFile?.xml || '').indexOf('<p:ph type="sldNum"'), (inFile?.xml || '').indexOf('<p:ph type="sldNum"') + 120)}`);

      // The dialog opens with the slide's own settings; Apply changes this slide alone.
      await js(`(() => { document.querySelectorAll('.sl-thumb')[1]?.click(); return 1; })()`);
      await until(() => js(`document.querySelectorAll('.sl-thumb')[1]?.classList.contains('active')`), 'the second slide', 4000).catch(() => {});
      await wait(300);
      const reopened = await clickRibbon('Header & Footer');
      await until(() => js(`Boolean(document.querySelector('.sl-hf-footer'))`), 'the dialog again', 4000).catch(() => {});
      const prefilled = await js(`(() => ({ footer: document.querySelector('.sl-hf-footer')?.value, on: document.querySelector('.sl-hf-footer-on')?.checked, number: document.querySelector('.sl-hf-number')?.checked }))()`);
      await js(`(() => { document.querySelector('.sl-hf-footer-on')?.click(); return 1; })()`);
      await wait(150);
      await js(`(() => { document.querySelector('.sl-hf-apply')?.click(); return 1; })()`);
      const gone = await until(() => ofType(1, 'ftr') === null && words(ofType(0, 'ftr')) === 'Rutba Office beta' && words(ofType(1, 'sldNum')) === '2', 'the footer off slide 2 alone', 5000).catch(() => false);
      check('slides: Header & Footer opens with the slide’s own settings, and Apply changes this slide alone',
        reopened === 'clicked' && prefilled?.footer === 'Rutba Office beta' && prefilled?.on === true && prefilled?.number === true && gone === true,
        `${reopened}; ${JSON.stringify(prefilled)}; slide 2 footer ${JSON.stringify(words(ofType(1, 'ftr')))}, slide 1 footer ${JSON.stringify(words(ofType(0, 'ftr')))}`);
      const complaints = await errorsIn(win);
      check('slides: the footer checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('slides: the footer checks ran', false, err.message);
    }
  };

  /* ── Presentation: sections ──────────────────────────────────────────── */
  //
  // Home → Section → Add Section on the second slide starts one there and
  // asks for its name; the strip and the saved file carry a Default Section
  // for the first slide and the named one for the second; Remove Section
  // folds it back, and Remove All Sections takes the list out of the file.
  // Run alone with RUTBA_VERIFY_ONLY=sections.
  const slideSections = async () => {
    try {
      const win = await open('slides', files.pptx);
      const js = (code) => win.webContents.executeJavaScript(code);
      const model = (slide) => doc.model({ id: sessionFor('deck').id, slide });
      await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 2`), 'the slide sorter', 8000);
      const clickRibbon = (title) => js(`(() => {
        const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}));
        if (!b) return 'no button ' + ${JSON.stringify(title)};
        b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
        b.click();
        return 'clicked';
      })()`);
      const pickMenu = async (label) => {
        await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}))`), `the ${label} item`, 4000);
        return js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}); if (b.disabled) return 'disabled'; b.click(); return 'picked'; })()`);
      };
      const headings = () => js(`[...document.querySelectorAll('.sl-sorter .sl-section')].map((h) => [h.querySelector('.sl-section-name')?.textContent, h.querySelector('.sl-section-count')?.textContent])`);
      const names = (m) => (m?.sections || []).map((s) => [s.name, s.slides]);
      // Earlier blocks may have left the fixture with more than its two slides: the first is one section, the rest the other.
      const count = model(0).count;
      const rest = Array.from({ length: count - 1 }, (_, i) => i + 1);
      const all = Array.from({ length: count }, (_, i) => i);

      // The second slide selected; Home → Section → Add Section.
      await js(`(() => { document.querySelectorAll('.sl-thumb')[1]?.click(); return 1; })()`);
      await until(() => js(`document.querySelectorAll('.sl-thumb')[1]?.classList.contains('active')`), 'the second slide', 4000).catch(() => {});
      await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Home')?.click(), 'tab'`);
      await wait(300);
      const before = await headings();
      const opened = await clickRibbon('Section');
      const added = await pickMenu('Add Section');
      await until(() => js(`Boolean(document.querySelector('.sl-section-name'))`), 'the section headings', 5000).catch(() => {});
      const asked = await until(() => js(`document.querySelector('.sl-section-input')?.value === 'Untitled Section'`), 'the name asked for', 4000).catch(() => false);
      check('slides: Home → Section → Add Section on the second slide starts a section there, the first slide in a Default Section, and asks for the name',
        opened === 'clicked' && added === 'picked' && before.length === 0 && asked === true && JSON.stringify(names(model(1))) === JSON.stringify([['Default Section', [0]], ['Untitled Section', rest]]),
        `${opened}; ${added}; headings before ${JSON.stringify(before)}; asked ${asked}; sections ${JSON.stringify(names(model(1)))}`);

      // The name typed and Rename pressed: the heading says so, and the file has the list as PowerPoint writes it.
      await js(`(() => { const el = document.querySelector('.sl-section-input'); if (!el) return false; const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set; setter.call(el, 'Closing'); el.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
      await wait(150);
      await js(`(() => { document.querySelector('.sl-section-ok')?.click(); return 1; })()`);
      const renamed = await until(() => names(model(1))[1]?.[0] === 'Closing', 'the section renamed', 5000).catch(() => false);
      await wait(300);
      const drawn = await headings();
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'slides-sections.png'), (await win.webContents.capturePage()).toPNG());
      check('slides: the name typed into the dialog is the heading drawn above the slide in the strip, with the slide count beside it',
        renamed === true && JSON.stringify(drawn) === JSON.stringify([['Default Section', '1'], ['Closing', String(rest.length)]]),
        `renamed ${renamed}; headings ${JSON.stringify(drawn)}`);

      await clickRibbon('Save');
      const saved = () => {
        try {
          const d = Deck.open(fs.readFileSync(files.pptx));
          return { sections: names({ sections: d.sections() }), xml: d.pkg.text('ppt/presentation.xml') };
        } catch { return null; }
      };
      await until(() => saved()?.sections.length === 2, 'the sections in the file', 8000).catch(() => false);
      const inFile = saved();
      const ids = [...(inFile?.xml || '').matchAll(/<p:sldId id="(\d+)"/g)].map((m) => m[1]);
      check('slides: the saved file carries the sections as PowerPoint writes them — one p:ext with the section-list URI, each section a name, a GUID and its slide ids',
        JSON.stringify(inFile?.sections) === JSON.stringify([['Default Section', [0]], ['Closing', rest]]) && ids.length === count &&
          new RegExp('<p:extLst><p:ext uri="\\{521415D9-36F7-43E2-AB2F-B90AF26B5E84\\}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p14:section name="Default Section" id="\\{[0-9A-F-]{36}\\}"><p14:sldIdLst><p14:sldId id="' + ids[0] + '"/></p14:sldIdLst></p14:section><p14:section name="Closing" id="\\{[0-9A-F-]{36}\\}"><p14:sldIdLst>' + ids.slice(1).map((id) => '<p14:sldId id="' + id + '"/>').join('') + '</p14:sldIdLst></p14:section></p14:sectionLst></p:ext></p:extLst></p:presentation>$').test(inFile?.xml || ''),
        `sections ${JSON.stringify(inFile?.sections)}; ${(inFile?.xml || '').slice((inFile?.xml || '').indexOf('<p:extLst>')).slice(0, 200)}`);

      // Remove Section on the second slide folds Closing into the Default Section; Remove All Sections takes the list out.
      const reopened = await clickRibbon('Section');
      const removed = await pickMenu('Remove Section');
      const folded = await until(() => JSON.stringify(names(model(1))) === JSON.stringify([['Default Section', all]]), 'the section folded back', 5000).catch(() => false);
      await wait(200);
      const one = await headings();
      const again = await clickRibbon('Section');
      const cleared = await pickMenu('Remove All Sections');
      const none = await until(() => names(model(1)).length === 0, 'no sections', 5000).catch(() => false);
      await wait(200);
      const gone = await headings();
      await clickRibbon('Save');
      await until(() => saved()?.sections.length === 0, 'no sections in the file', 8000).catch(() => false);
      const cleanXml = saved()?.xml || '';
      check('slides: Remove Section folds the section into the one before it, and Remove All Sections takes the list out of the strip and the file',
        reopened === 'clicked' && removed === 'picked' && folded === true && JSON.stringify(one) === JSON.stringify([['Default Section', String(count)]]) && again === 'clicked' && cleared === 'picked' && none === true && gone.length === 0 && !/sectionLst|<p:extLst>/.test(cleanXml),
        `${removed}; folded ${folded}; headings ${JSON.stringify(one)} → ${JSON.stringify(gone)}; ${cleared}; none ${none}; file ${/sectionLst/.test(cleanXml) ? 'still has the list' : 'clean'}`);
      const complaints = await errorsIn(win);
      check('slides: the section checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('slides: the section checks ran', false, err.message);
    }
  };

  /* ── Presentation: Slide Show → Hide Slide ───────────────────────────── */
  //
  // A hidden slide stays in the file — PowerPoint's show="0" on the slide's
  // own root tag — and everywhere but the show itself (the strip, the
  // sorter, printing) still draws it, dimmed with its number struck
  // through. Only the show steps over it.
  // Run alone with RUTBA_VERIFY_ONLY=hidden.
  const slideHidden = async () => {
    try {
      const win = await open('slides', files.pptx);
      const js = (code) => win.webContents.executeJavaScript(code);
      const model = (slide) => doc.model({ id: sessionFor('deck').id, slide });
      await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 2`), 'the slide sorter', 8000);
      const clickRibbon = (title) => js(`(() => {
        const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}));
        if (!b) return 'no button ' + ${JSON.stringify(title)};
        b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
        b.click();
        return 'clicked';
      })()`);
      const pressedState = (title) => js(`(() => {
        const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}));
        return b ? b.getAttribute('aria-pressed') : null;
      })()`);

      // The second slide selected, then Slide Show → Hide Slide.
      await js(`(() => { document.querySelectorAll('.sl-thumb')[1]?.click(); return 1; })()`);
      await until(() => js(`document.querySelectorAll('.sl-thumb')[1]?.classList.contains('active')`), 'the second slide', 4000).catch(() => {});
      await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Slide Show')?.click(), 'tab'`);
      await wait(200);
      const hidStart = await clickRibbon('Hide Slide');
      const hiddenInModel = await until(() => model(1).slide.hidden === true, 'the slide marked hidden', 5000).catch(() => false);
      await wait(150);
      const thumbHidden = await js(`document.querySelectorAll('.sl-thumb')[1]?.classList.contains('hidden')`);
      const struck = await js(`getComputedStyle(document.querySelectorAll('.sl-thumb')[1]?.querySelector('.sl-thumb-n')).textDecorationLine`);
      const btnPressed = await pressedState('Hide Slide');
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'slides-hidden.png'), (await win.webContents.capturePage()).toPNG());
      check('slides: Slide Show → Hide Slide marks the slide hidden, strikes its number through in the strip, and the button stays pressed',
        hidStart === 'clicked' && hiddenInModel === true && thumbHidden === true && /line-through/.test(struck || '') && btnPressed === 'true',
        `${hidStart}; hidden ${hiddenInModel}; thumb class ${thumbHidden}; decoration ${struck}; pressed ${btnPressed}`);

      await clickRibbon('Save');
      const saved = () => {
        try { return Deck.open(fs.readFileSync(files.pptx)); } catch { return null; }
      };
      await until(() => saved()?.isSlideHidden(1) === true, 'the hidden slide in the file', 8000).catch(() => {});
      const d = saved();
      const xml1 = d ? d.pkg.text(d.slideParts[1].part) : '';
      const xml0 = d ? d.pkg.text(d.slideParts[0].part) : '';
      check('slides: the saved file carries the hidden slide as PowerPoint writes one — show="0" on the slide\'s own root tag, and no other slide touched',
        Boolean(d) && d.isSlideHidden(1) === true && /<p:sld\b[^>]*\sshow="0"/.test(xml1) && !/<p:sld\b[^>]*\sshow="0"/.test(xml0),
        `hidden ${d?.isSlideHidden(1)}; slide 2 root ${xml1.slice(0, 80)}`);

      // Start the show: a hidden slide never opens it, and is stepped over.
      const started = await clickRibbon('Start from the beginning');
      await until(() => js(`Boolean(document.querySelector('.sl-present-bar'))`), 'the show', 5000).catch(() => {});
      const barText = () => js(`document.querySelector('.sl-present-bar')?.textContent || ''`);
      const bar1 = await barText();
      const outline = model(0).outline || [];
      let expectAfterRight = 0;
      for (let i = 1; i < outline.length; i++) {
        if (!outline[i].hidden) { expectAfterRight = i; break; }
      }
      await js(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }))`);
      await wait(300);
      const bar2 = await barText();
      await js(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))`);
      await until(() => js(`!document.querySelector('.sl-present-bar')`), 'the show to close', 5000).catch(() => {});
      check('slides: the show starts on the first shown slide and steps over a hidden one; Escape leaves it',
        started === 'clicked' && bar1.startsWith('1 / ') && bar2.startsWith(`${expectAfterRight + 1} / `),
        `${started}; bar ${JSON.stringify(bar1)} → ${JSON.stringify(bar2)}; expected slide ${expectAfterRight + 1}`);

      // Hide Slide pressed again shows the slide — on the second slide, which
      // the show left: it opened on the first shown slide and stayed there.
      await js(`(() => { document.querySelectorAll('.sl-thumb')[1]?.click(); return 1; })()`);
      await until(() => js(`document.querySelectorAll('.sl-thumb')[1]?.classList.contains('active')`), 'the second slide again', 4000).catch(() => {});
      await wait(200);
      const shownAgain = await clickRibbon('Hide Slide');
      const unhidden = await until(() => model(1).outline?.[1]?.hidden === false && model(1).slide.hidden === false, 'the slide shown again', 5000).catch(() => false);
      await until(() => js(`document.querySelectorAll('.sl-thumb')[1]?.classList.contains('hidden') === false`), 'the strip to show it again', 4000).catch(() => {});
      const thumbShown = await js(`document.querySelectorAll('.sl-thumb')[1]?.classList.contains('hidden')`);
      const stripState = await js(`[...document.querySelectorAll('.sl-thumb')].map((t) => t.className)`);
      const outlineState = model(1).outline?.map((o) => o.hidden);
      check('slides: Hide Slide pressed again shows the slide, in the model and the strip',
        shownAgain === 'clicked' && unhidden === true && thumbShown === false,
        `${shownAgain}; hidden ${unhidden}; thumb class ${thumbShown}; strip ${JSON.stringify(stripState)}; outline ${JSON.stringify(outlineState)}`);

      const complaints = await errorsIn(win);
      check('slides: the hide-slide checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('slides: the hide-slide checks ran', false, err.message);
    }
  };

  /* ── Presentation: Design → Background Styles ────────────────────────── */
  //
  // A slide's own background — PowerPoint's p:bg, the first child of p:cSld
  // — drawn on the stage and the thumbnails, chosen from the ribbon for this
  // slide or every slide, and taken off again to let the layout's show
  // through. Run alone with RUTBA_VERIFY_ONLY=background.
  const slideBackground = async () => {
    try {
      const win = await open('slides', files.pptx);
      const js = (code) => win.webContents.executeJavaScript(code);
      const model = (slide) => doc.model({ id: sessionFor('deck').id, slide });
      await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 2`), 'the slide sorter', 8000);
      const clickRibbon = (title) => js(`(() => {
        const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}));
        if (!b) return 'no button ' + ${JSON.stringify(title)};
        b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
        b.click();
        return 'clicked';
      })()`);
      const pickMenu = async (label) => {
        await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}))`), `the ${label} item`, 4000);
        return js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}); if (b.disabled) return 'disabled'; b.click(); return 'picked'; })()`);
      };
      const stageFill = () => js(`(document.querySelector('.sl-svg svg rect')?.getAttribute('fill') || '').toLowerCase()`);
      const thumbFill = (i) => js(`(document.querySelectorAll('.sl-thumb-pic')[${i}]?.querySelector('rect')?.getAttribute('fill') || '').toLowerCase()`);
      const saved = () => {
        try { return Deck.open(fs.readFileSync(files.pptx)); } catch { return null; }
      };

      await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Design')?.click(), 'tab'`);
      await wait(200);
      const opened = await clickRibbon('Background Styles');
      const picked = await pickMenu('Dark blue');
      const gotOwn = await until(() => JSON.stringify(model(0).slide.ownBackground) === JSON.stringify({ colour: '1F3864' }), "the slide's own background", 5000).catch(() => false);
      await wait(200);
      const svgFill = await stageFill();
      const thumb0 = await thumbFill(0);
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'slides-background.png'), (await win.webContents.capturePage()).toPNG());
      check('slides: Design → Background Styles → Dark blue gives this slide its own background, ticked in the menu and drawn on the stage and the first thumbnail',
        opened === 'clicked' && picked === 'picked' && gotOwn === true && svgFill === '#1f3864' && thumb0 === '#1f3864',
        `${opened}; ${picked}; own ${gotOwn}; stage ${svgFill}; thumb ${thumb0}`);

      await clickRibbon('Save');
      await until(() => /<p:bg>/.test(saved()?.pkg.text(saved().slideParts[0].part) || ''), 'the background in the file', 8000).catch(() => {});
      const xml0 = saved()?.pkg.text(saved().slideParts[0].part) || '';
      const xml1 = saved()?.pkg.text(saved().slideParts[1].part) || '';
      check('slides: the saved file carries the background as PowerPoint writes one — p:bg the first child of cSld — and touches no other slide',
        /<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="1F3864"\/><\/a:solidFill><a:effectLst\/><\/p:bgPr><\/p:bg><p:spTree>/.test(xml0) && !/<p:bg>/.test(xml1),
        `slide 1 ${xml0.slice(0, 120)}; slide 2 has bg: ${/<p:bg>/.test(xml1)}`);

      const opened2 = await clickRibbon('Background Styles');
      const applied = await pickMenu('Apply to all slides');
      const spreadToAll = await until(() => JSON.stringify(model(1).slide.ownBackground) === JSON.stringify({ colour: '1F3864' }), "every slide's own background", 5000).catch(() => false);
      check('slides: Background Styles → Apply to all slides gives every slide this slide’s own background',
        opened2 === 'clicked' && applied === 'picked' && spreadToAll === true,
        `${opened2}; ${applied}; slide 2 own ${JSON.stringify(model(1).slide.ownBackground)}`);

      const opened3 = await clickRibbon('Background Styles');
      const reset = await pickMenu("Reset to the layout's");
      const clearedOwn = await until(() => model(0).slide.ownBackground === null, "the slide's own background cleared", 5000).catch(() => false);
      check('slides: Background Styles → Reset to the layout’s takes this slide’s own background off, so the layout’s or master’s shows through again',
        opened3 === 'clicked' && reset === 'picked' && clearedOwn === true,
        `${opened3}; ${reset}; own ${JSON.stringify(model(0).slide.ownBackground)}`);

      await clickRibbon('Save');
      await until(() => !/<p:bg>/.test(saved()?.pkg.text(saved().slideParts[0].part) || 'still has it'), 'the background gone from the file', 8000).catch(() => {});
      const cleanXml = saved()?.pkg.text(saved().slideParts[0].part) || '';
      check('slides: the reset background is no longer in the saved file', !/<p:bg>/.test(cleanXml), cleanXml.slice(0, 120));

      const complaints = await errorsIn(win);
      check('slides: the background checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('slides: the background checks ran', false, err.message);
    }
  };

  /* ── Presentation: a table ───────────────────────────────────────────── */
  //
  // Insert → Table: a grid of cells written PowerPoint's own way, drawn on
  // the stage with a hit area over every cell, a cell edited in place, a
  // row added and the table deleted. Run alone with RUTBA_VERIFY_ONLY=table.
  const slideTable = async () => {
    try {
      const win = await open('slides', files.pptx);
      const wc = win.webContents;
      const js = (code) => win.webContents.executeJavaScript(code);
      const model = () => doc.model({ id: sessionFor('deck').id, slide: 0 });
      await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 2`), 'the slide sorter', 8000);
      const clickRibbon = (title) => js(`(() => {
        const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}));
        if (!b) return 'no button ' + ${JSON.stringify(title)};
        b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
        b.click();
        return 'clicked';
      })()`);
      const pickMenu = async (label) => {
        await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}))`), `the ${label} item`, 4000);
        return js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}); if (b.disabled) return 'disabled'; b.click(); return 'picked'; })()`);
      };
      const saved = () => {
        try { return Deck.open(fs.readFileSync(files.pptx)); } catch { return null; }
      };
      const tableShape = () => model().slide.shapes.find((s) => s.kind === 'table');

      await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click(), 'tab'`);
      await wait(200);
      const opened = await clickRibbon('Table');
      const picked = await pickMenu('3 × 3');
      const added = await until(() => tableShape()?.table?.rows === 3 && tableShape()?.table?.cols === 3, 'the table on the model', 5000).catch(() => false);
      check('slides: Insert → Table → 3 × 3 puts a 3-by-3 table on the slide',
        opened === 'clicked' && picked === 'picked' && added === true,
        `${opened}; ${picked}; ${JSON.stringify(model().slide.shapes.map((s) => s.kind))}`);

      const shapeId = tableShape()?.id;
      const cellsOnStage = await until(() => js(`document.querySelectorAll('.sl-cell-hit').length === 9`), 'nine cell hit areas', 5000).catch(() => false);
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'slides-table.png'), (await win.webContents.capturePage()).toPNG());
      check('slides: the table is drawn on the stage, with a hit area over every cell',
        cellsOnStage === true, `${await js(`document.querySelectorAll('.sl-cell-hit').length`)} cell hit area(s)`);

      // A double-click opens the editor over the cell; Enter commits its words.
      await js(`(() => { document.querySelector('.sl-cell-hit[data-row="0"][data-col="0"]')?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, cancelable: true })); return 1; })()`);
      await until(() => js(`Boolean(document.querySelector('.sl-editor'))`), 'the cell editor', 4000);
      await js(`(() => { const ta = document.querySelector('.sl-editor'); const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set; setter.call(ta, 'Sales'); ta.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
      await js(`(() => { document.querySelector('.sl-editor')?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true })); return 1; })()`);
      const typed = await until(() => (tableShape()?.table?.cells?.[0]?.[0]?.paragraphs || []).map((p) => p.plain).join('') === 'Sales', 'the cell to read Sales', 5000).catch(() => false);
      check('slides: a double-click opens an editor over the cell, and Enter commits its words',
        typed === true, JSON.stringify(tableShape()?.table?.cells?.[0]?.[0]));

      await clickRibbon('Save');
      await until(() => /<a:t[^>]*>Sales<\/a:t>/.test(saved()?.pkg.text(saved().slideParts[0].part) || ''), 'the cell in the file', 8000).catch(() => {});
      const xml = saved()?.pkg.text(saved().slideParts[0].part) || '';
      const tblAt = xml.indexOf('<a:tbl>');
      check('slides: the saved file carries the table as PowerPoint writes one — tblPr, the style id, three columns, three rows, and the edited word',
        /<a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>\{5C22544A-7EE6-4342-B048-85BDC9FD1C3A\}<\/a:tableStyleId><\/a:tblPr><a:tblGrid>(?:<a:gridCol[^\/]*\/>){3}<\/a:tblGrid>/.test(xml)
          && (xml.match(/<a:tr h="\d+">/g) || []).length === 3
          && /<a:t[^>]*>Sales<\/a:t>/.test(xml),
        xml.slice(tblAt, tblAt + 400));

      // A row added: the cell's own right-click menu, the way PowerPoint's table offers it.
      await js(`(() => { document.querySelector('.sl-cell-hit[data-row="0"][data-col="0"]')?.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true })); return 1; })()`);
      await pickMenu('Insert row below');
      const gotRow = await until(() => tableShape()?.table?.rows === 4, 'the row added', 5000).catch(() => false);
      check('slides: a right-click on a cell offers Insert row, and the table grows a row', gotRow === true, `rows ${tableShape()?.table?.rows}`);

      await clickRibbon('Save');
      await until(() => (saved()?.pkg.text(saved().slideParts[0].part).match(/<a:tr h="\d+">/g) || []).length === 4, 'four rows in the file', 8000).catch(() => {});
      const xml2 = saved()?.pkg.text(saved().slideParts[0].part) || '';
      const tblAt2 = xml2.indexOf('<a:tbl>');
      check('slides: the saved file carries the added row', (xml2.match(/<a:tr h="\d+">/g) || []).length === 4, xml2.slice(tblAt2, tblAt2 + 200));

      // Deleted like any other shape: selected, then the Delete key.
      await js(`(() => { document.querySelector('.sl-hit[data-shape="${shapeId}"]')?.click(); return 1; })()`);
      await until(() => js(`document.querySelector('.sl-hit.selected')?.dataset.shape === ${JSON.stringify(String(shapeId))}`), 'the table selected', 4000).catch(() => {});
      await js(`(() => { document.querySelector('.sl-stage').focus(); return 1; })()`);
      await press(wc, 'Delete');
      const gone = await until(() => !model().slide.shapes.some((s) => s.id === shapeId), 'the table gone', 5000).catch(() => false);
      check('slides: Delete removes the table', gone === true, JSON.stringify(model().slide.shapes.map((s) => s.kind)));

      const complaints = await errorsIn(win);
      check('slides: the table checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('slides: the table checks ran', false, err.message);
    }
  };

  /* ── Presentation: Insert → Chart ─────────────────────────────────────── */
  //
  // Insert → Chart → Column drops a sample chart, drawn by the chart writer
  // Word and Worksheets already use and framed the way PowerPoint frames
  // one; a double-click opens its data, Apply rewrites the part, and Save
  // writes a chart PowerPoint itself reads.
  const slideChart = async () => {
    try {
      const win = await open('slides', files.pptx);
      const wc = win.webContents;
      const js = (code) => win.webContents.executeJavaScript(code);
      const model = () => doc.model({ id: sessionFor('deck').id, slide: 0 });
      await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 2`), 'the slide sorter', 8000);
      const clickRibbon = (title) => js(`(() => {
        const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}));
        if (!b) return 'no button ' + ${JSON.stringify(title)};
        b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
        b.click();
        return 'clicked';
      })()`);
      const pickMenu = async (label) => {
        await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}))`), `the ${label} item`, 4000);
        return js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}); if (b.disabled) return 'disabled'; b.click(); return 'picked'; })()`);
      };
      const setValue = (selector, value) => js(`(() => { const el = document.querySelector(${JSON.stringify(selector)}); if (!el) return false; const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set; setter.call(el, ${JSON.stringify(String(value))}); el.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
      const saved = () => {
        try { return Deck.open(fs.readFileSync(files.pptx)); } catch { return null; }
      };
      const chartShape = () => model().slide.shapes.find((s) => s.kind === 'chart');

      await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click(), 'tab'`);
      await wait(200);
      const opened = await clickRibbon('Chart');
      const picked = await pickMenu('Column');
      const added = await until(() => chartShape()?.chart?.categories?.length === 4, 'the chart on the model', 5000).catch(() => false);
      check('slides: Insert → Chart → Column puts a chart with four categories on the slide',
        opened === 'clicked' && picked === 'picked' && added === true,
        `${opened}; ${picked}; ${JSON.stringify(model().slide.shapes.map((s) => s.kind))}`);

      const shapeId = chartShape()?.id;
      const barsDrawn = await until(() => js(`document.querySelectorAll('.sl-svg .marks path').length >= 8`), 'at least eight bars', 5000).catch(() => false);
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'slides-chart.png'), (await win.webContents.capturePage()).toPNG());
      check('slides: the chart is drawn on the stage as bars — two series of four categories',
        barsDrawn === true, `${await js(`document.querySelectorAll('.sl-svg .marks path').length`)} bar(s)`);

      // A double-click on the chart's own hit area opens its data.
      await js(`(() => { document.querySelector('.sl-hit[data-shape="${shapeId}"]')?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, cancelable: true })); return 1; })()`);
      await until(() => js(`Boolean(document.querySelector('.sl-chart-data'))`), 'the chart data dialog', 4000);
      await setValue('.sl-chart-cell[data-row="0"][data-col="0"]', 30);
      await js(`(() => { document.querySelector('.sl-chart-apply')?.click(); return 1; })()`);
      const edited = await until(() => chartShape()?.chart?.series?.[0]?.values?.[0] === 30, 'the first value edited to 30', 5000).catch(() => false);
      check('slides: a double-click opens the chart data dialog, and Apply rewrites the first value',
        edited === true, JSON.stringify(chartShape()?.chart?.series?.[0]));

      await clickRibbon('Save');
      await until(() => /<c:v>30<\/c:v>/.test(saved()?.pkg.text('ppt/charts/chart1.xml') || ''), 'the edited value in the file', 8000).catch(() => {});
      const slideXml = saved()?.pkg.text(saved()?.slideParts[0]?.part || '') || '';
      const frameMatch = /<a:graphicData uri="http:\/\/schemas\.openxmlformats\.org\/drawingml\/2006\/chart"><c:chart[^>]*\br:id="([^"]+)"/.exec(slideXml);
      const rels = saved()?.pkg.rels(saved().slideParts[0].part) || [];
      const chartRel = rels.find((r) => r.Id === frameMatch?.[1]);
      const chartPart = chartRel && OoxmlPackage.resolveTarget(saved().slideParts[0].part, chartRel.Target);
      const chartXml = chartPart ? saved()?.pkg.text(chartPart) || '' : '';
      check('slides: the saved file carries the chart as PowerPoint writes one — the frame\'s graphicData and c:chart r:id resolve to a chart part with a barChart, a vertical bar direction and the edited value',
        Boolean(frameMatch) && chartPart === 'ppt/charts/chart1.xml'
          && /<c:barChart>/.test(chartXml) && /<c:barDir val="col"\/>/.test(chartXml)
          && /<c:v>30<\/c:v>/.test(chartXml),
        `frame r:id ${frameMatch?.[1]}; chart part ${chartPart}; ${chartXml.slice(0, 200)}`);
      const contentTypesXml = saved()?.pkg.text('[Content_Types].xml') || '';
      check('slides: the saved file declares the chart part\'s content type',
        contentTypesXml.includes('PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"'),
        contentTypesXml.slice(-300));

      // Deleted like any other shape: selected, then the Delete key.
      await js(`(() => { document.querySelector('.sl-hit[data-shape="${shapeId}"]')?.click(); return 1; })()`);
      await until(() => js(`document.querySelector('.sl-hit.selected')?.dataset.shape === ${JSON.stringify(String(shapeId))}`), 'the chart selected', 4000).catch(() => {});
      await js(`(() => { document.querySelector('.sl-stage').focus(); return 1; })()`);
      await press(wc, 'Delete');
      const gone = await until(() => !model().slide.shapes.some((s) => s.id === shapeId), 'the chart gone', 5000).catch(() => false);
      check('slides: Delete removes the chart', gone === true, JSON.stringify(model().slide.shapes.map((s) => s.kind)));

      const complaints = await errorsIn(win);
      check('slides: the chart checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('slides: the chart checks ran', false, err.message);
    }
  };

  /* ── Presentation: find and replace ──────────────────────────────────── */
  //
  // Home → Find: the words typed, Find lists every shape they are on, a
  // match clicked goes to its slide and selects the shape, and Replace all
  // rewrites them across the deck, each run keeping its look.
  /**
   * Home → Find (Ctrl+F) and Replace (Ctrl+H): a small pane pinned to the
   * stage's own top right, over whatever slide is on screen — Word's find
   * pane, not a dialog that covers the slide. Two marker shapes, one on
   * each slide, are added first, so Next has a real slide to move to
   * whatever else the deck holds by the time this runs; both come off
   * again at the end, so the deck is as this check found it.
   */
  const slideFind = async () => {
    try {
      const win = await open('slides', files.pptx);
      const js = (code) => win.webContents.executeJavaScript(code);
      const model = (slide) => doc.model({ id: sessionFor('deck').id, slide });
      await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 2`), 'the slide sorter', 8000);

      const seedOp = (slide) => ({ op: 'addTextBox', slide, x: 40, y: 380, w: 260, h: 40, paragraphs: [{ runs: [{ text: 'Zephyr marker' }] }] });
      await doc.apply({ id: sessionFor('deck').id, ops: [seedOp(0)] });
      await doc.apply({ id: sessionFor('deck').id, ops: [seedOp(1)] });
      const markers = [model(0).slide.shapes.slice(-1)[0].id, model(1).slide.shapes.slice(-1)[0].id];

      const clickRibbon = (title) => js(`(() => {
        const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}));
        if (!b) return 'no button ' + ${JSON.stringify(title)};
        b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
        b.click();
        return 'clicked';
      })()`);
      const type = (selector, value) => js(`(() => { const el = document.querySelector(${JSON.stringify(selector)}); if (!el) return false; const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set; setter.call(el, ${JSON.stringify(value)}); el.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
      const readCount = () => js(`document.querySelector('.sl-find-count')?.textContent.trim() || ''`);
      const activeSlide = () => js(`[...document.querySelectorAll('.sl-thumb')].findIndex((t) => t.classList.contains('active'))`);
      const wordsOf = (s) => (s?.text?.paragraphs || []).map((p) => (p.runs || []).map((r) => r.text).join('')).join(' ');
      const everywhere = (read) => { const all = []; for (let i = 0; i < (model(0).count || 2); i++) all.push(...read(i).slide.shapes.map(wordsOf)); return all; };
      const savedWords = () => { try { const d = Deck.open(fs.readFileSync(files.pptx)); const all = []; for (let i = 0; i < d.slideCount; i++) all.push(...d.slide(i).shapes.map(wordsOf)); return all; } catch { return []; } };

      await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Home')?.click(), 'tab'`);
      await wait(200);
      const opened = await clickRibbon('Replace');
      await until(() => js(`Boolean(document.querySelector('.sl-find-text'))`), 'the find pane', 4000).catch(() => {});
      await type('.sl-find-text', 'Zephyr');
      const gotOne = await until(() => readCount().then((t) => t === '1 of 2'), 'the pane to read "1 of 2"', 5000).catch(() => false);
      const onFirst = await activeSlide();
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'slides-find.png'), (await win.webContents.capturePage()).toPNG());
      check('slides: Home → Replace (Ctrl+H) opens the pane, and typing a word reads "1 of n" on its first hit\'s slide',
        opened === 'clicked' && gotOne === true && onFirst === 0,
        `${opened}; count ${await readCount()}; slide ${onFirst + 1}`);

      await js(`(() => { document.querySelector('.sl-find-next')?.click(); return 1; })()`);
      const movedOn = await until(() => js(`document.querySelectorAll('.sl-thumb')[1]?.classList.contains('active')`), 'Next to move to the second hit\'s slide', 5000).catch(() => false);
      check('slides: Next moves to the next hit\'s slide, the count following it',
        movedOn === true && (await readCount()) === '2 of 2',
        `slide ${(await activeSlide()) + 1}; count ${await readCount()}`);

      // Match case narrows a query that would otherwise still match.
      await js(`(() => { document.querySelector('.sl-find-case-box')?.click(); return 1; })()`);
      await type('.sl-find-text', 'ZEPHYR');
      const nothingWithCase = await until(() => readCount().then((t) => t === '0 of 0'), 'no matches once Match case is on', 4000).catch(() => false);
      check('slides: Match case narrows the search', nothingWithCase === true, await readCount());
      await js(`(() => { document.querySelector('.sl-find-case-box')?.click(); return 1; })()`);
      await type('.sl-find-text', 'Zephyr');
      await until(() => readCount().then((t) => t === '1 of 2'), 'the count back to "1 of 2"', 4000).catch(() => {});

      await type('.sl-find-with', 'Marker');
      await js(`(() => { document.querySelector('.sl-find-replace')?.click(); return 1; })()`);
      const replaced = await until(() => everywhere(model).some((t) => /Marker marker/.test(t)) && !everywhere(model).some((t) => /Zephyr/.test(t)), 'the words replaced', 6000).catch(() => false);
      // The note is drawn a render after the engine answers.
      await until(() => js(`Boolean(document.querySelector('.sl-find-note'))`), 'the note', 4000).catch(() => {});
      const noteText = await js(`document.querySelector('.sl-find-note')?.textContent.trim() || null`);
      check('slides: Replace all rewrites every hit and says how many',
        replaced === true && noteText === 'Replaced 2 across the deck.',
        `${JSON.stringify(noteText)}; ${JSON.stringify(everywhere(model))}`);

      await clickRibbon('Save');
      await until(() => savedWords().some((t) => /Marker marker/.test(t)), 'the replacement in the saved file', 8000).catch(() => false);
      check('slides: the saved file carries the replaced words',
        savedWords().some((t) => /Marker marker/.test(t)) && !savedWords().some((t) => /Zephyr/.test(t)),
        JSON.stringify(savedWords().filter((t) => /marker/i.test(t))));

      // The marker shapes this check added come off again, so the deck is
      // exactly as it was for whatever runs after this.
      await doc.apply({ id: sessionFor('deck').id, ops: [{ op: 'removeShape', slide: 0, shape: markers[0] }] });
      await doc.apply({ id: sessionFor('deck').id, ops: [{ op: 'removeShape', slide: 1, shape: markers[1] }] });
      await js(`(() => { document.querySelector('.sl-find-close')?.click(); return 1; })()`);
      await wait(200);
      await clickRibbon('Save');
      await until(() => !savedWords().some((t) => /marker/i.test(t)), 'the deck as it was', 8000).catch(() => false);
      check('slides: the find checks leave the deck as they found it', !savedWords().some((t) => /marker/i.test(t)), JSON.stringify(savedWords()));

      const complaints = await errorsIn(win);
      check('slides: the find checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('slides: the find checks ran', false, err.message);
    }
  };

  const slideClipboard = async () => {
    try {
      const win = await open('slides', files.pptx);
      const js = (code) => win.webContents.executeJavaScript(code);
      const model = (slide) => doc.model({ id: sessionFor('deck').id, slide });
      const source = model(0).slide.shapes.find((s) => s.text);
      if (!source) return check('slides: a text box to copy', false, 'no text shape on slide 1');
      const words = (s) => (s?.text?.paragraphs || []).map((p) => (p.runs || []).map((r) => r.text).join('')).join('|');
      const hit = `.sl-hit[data-shape="${source.id}"]`;
      await until(() => js(`Boolean(document.querySelector(${JSON.stringify(hit)}))`), 'the hit area of the text box', 6000);
      await js(`(() => { document.querySelector(${JSON.stringify(hit)}).click(); return 1; })()`);
      await until(() => js(`document.querySelectorAll('.sl-handle').length === 8`), 'the eight handles', 4000).catch(() => {});
      const clickRibbon = (title) => js(`(() => {
        const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}));
        if (!b) return 'no button ' + ${JSON.stringify(title)};
        b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
        b.click();
        return 'clicked';
      })()`);
      const copied = await clickRibbon('Copy');
      await wait(300);
      await js(`(() => { document.querySelectorAll('.sl-thumb')[1]?.click(); return 1; })()`);
      await until(() => js(`document.querySelectorAll('.sl-thumb')[1]?.classList.contains('active')`), 'the second slide', 4000).catch(() => {});
      await wait(300);
      const before = model(1).slide.shapes.length;
      const pasted = await clickRibbon('Paste');
      const grew = await until(() => model(1).slide.shapes.length === before + 1, 'the pasted shape', 5000).catch(() => false);
      const last = model(1).slide.shapes[model(1).slide.shapes.length - 1];
      check('slides: Copy on a shape and Paste on another slide puts it there with the same words and a fresh id',
        copied === 'clicked' && pasted === 'clicked' && grew === true && words(last) === words(source) && last.id !== source.id,
        `${copied}, ${pasted}; slide 2 shapes ${before} → ${model(1).slide.shapes.length}; words ${JSON.stringify(words(last))}; ids ${source.id} → ${last.id}`);

      await wait(300);
      const cut = await clickRibbon('Cut');
      const shrank = await until(() => model(1).slide.shapes.length === before, 'the shape to go', 5000).catch(() => false);
      check('slides: Cut takes the selected shape away', cut === 'clicked' && shrank === true, `${cut}; slide 2 shapes now ${model(1).slide.shapes.length}`);

      await wait(300);
      await clickRibbon('Paste');
      await until(() => model(1).slide.shapes.length === before + 1, 'the shape pasted again', 5000).catch(() => false);
      await wait(300);
      await clickRibbon('Save');
      const inFile = () => Deck.open(fs.readFileSync(files.pptx)).slide(1).shapes;
      await until(() => { try { return inFile().length === before + 1; } catch { return false; } }, 'the paste to land in the file', 8000).catch(() => false);
      const kept = inFile();
      check('slides: the saved file carries the pasted shape', kept.length === before + 1 && words(kept[kept.length - 1]) === words(source), `slide 2 in the file: ${kept.length} shape(s), last ${JSON.stringify(words(kept[kept.length - 1]))}`);
      // Format Painter: the look of one shape onto another. The pasted box
      // is given a fill and an outline first, through the service; the
      // window fetches the slide again when it comes back to it.
      await wait(300);
      const painted = model(1).slide.shapes[model(1).slide.shapes.length - 1];
      const title = model(1).slide.shapes.find((s) => s.id !== painted.id && s.text);
      if (!title) {
        check('slides: a second text shape to paint', false, 'slide 2 has no other text shape');
      } else {
        await doc.apply({ id: sessionFor('deck').id, ops: [{ op: 'setShapeStyle', slide: 1, shape: painted.id, fill: '#FF9900', line: { color: '#333333', width: 2 } }] });
        await js(`(() => { document.querySelectorAll('.sl-thumb')[0]?.click(); return 1; })()`);
        await wait(400);
        await js(`(() => { document.querySelectorAll('.sl-thumb')[1]?.click(); return 1; })()`);
        await until(() => js(`document.querySelectorAll('.sl-thumb')[1]?.classList.contains('active')`), 'the second slide again', 4000).catch(() => {});
        await wait(400);
        await js(`(() => { document.querySelector('.sl-hit[data-shape="${painted.id}"]')?.click(); return 1; })()`);
        await wait(300);
        const armed = await clickRibbon('Format Painter');
        const pressed = await until(() => js(`Boolean([...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Format Painter — armed')))`), 'the painter armed', 3000).catch(() => false);
        await js(`(() => { document.querySelector('.sl-hit[data-shape="${title.id}"]')?.click(); return 1; })()`);
        const took = await until(() => { const s = model(1).slide.shapes.find((x) => x.id === title.id); return String(s?.fill?.color || '').toUpperCase() === '#FF9900' && s?.line?.width === 2; }, 'the look painted', 5000).catch(() => false);
        const after = model(1).slide.shapes.find((x) => x.id === title.id);
        await wait(400);
        if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'slides-painter.png'), (await win.webContents.capturePage()).toPNG());
        check('slides: Format Painter carries a shape’s fill and outline onto the next shape clicked', armed === 'clicked' && pressed === true && took === true, `${armed}; armed ${pressed}; fill ${JSON.stringify(after?.fill)} line ${JSON.stringify(after?.line)}`);
        await wait(300);
        await clickRibbon('Save');
        const savedTitle = () => { try { return Deck.open(fs.readFileSync(files.pptx)).slide(1).shapes.find((x) => x.id === title.id) || null; } catch { return null; } };
        await until(() => String(savedTitle()?.fill?.color || '').toUpperCase() === '#FF9900', 'the painted look in the file', 8000).catch(() => false);
        const saved = savedTitle();
        check('slides: the saved file keeps the painted fill and outline in the shape’s own properties', String(saved?.fill?.color || '').toUpperCase() === '#FF9900' && saved?.line?.width === 2, `${JSON.stringify(saved?.fill)} ${JSON.stringify(saved?.line)}`);
      }
      // Shape Effects: a shadow under the pasted box, from the ribbon's menu,
      // drawn on the stage and written after the outline.
      await wait(300);
      await js(`(() => { document.querySelector('.sl-hit[data-shape="${painted.id}"]')?.click(); return 1; })()`);
      await wait(300);
      const shadowed = await clickRibbon('Shape Effects');
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === 'Shadow: bottom right'))`), 'the effects menu', 3000).catch(() => {});
      await js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === 'Shadow: bottom right')?.click(); return 1; })()`);
      const shadowOf = () => model(1).slide.shapes.find((x) => x.id === painted.id)?.effects || null;
      const cast = await until(() => Math.round(shadowOf()?.shadow?.dir ?? -1) === 45, 'the shadow in the model', 5000).catch(() => false);
      const drawnShadow = await until(() => js(`Boolean(document.querySelector('.sl-stage svg filter feDropShadow, .sl-stage svg filter feOffset'))`), 'the shadow drawn', 4000).catch(() => false);
      await wait(400);
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'slides-shadow.png'), (await win.webContents.capturePage()).toPNG());
      check('slides: Shape Effects puts a shadow under the shape, drawn on the stage', shadowed === 'clicked' && cast === true && drawnShadow === true, `${shadowed}; model ${JSON.stringify(shadowOf())}; drawn ${drawnShadow}`);
      await wait(300);
      await clickRibbon('Save');
      const shadowInFile = () => { try { const d = Deck.open(fs.readFileSync(files.pptx)); const x = d.pkg.text(d.slideParts[1].part); const m = new RegExp('<p:cNvPr id="' + painted.id + '"[\\s\\S]*?</p:sp>').exec(x); return m ? m[0] : ''; } catch { return ''; } };
      await until(() => /<a:effectLst><a:outerShdw/.test(shadowInFile()), 'the shadow in the file', 8000).catch(() => false);
      const spXml = shadowInFile();
      check('slides: the saved file keeps the shadow after the outline in the shape’s own properties',
        /<\/a:ln><a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="40000"\/><\/a:srgbClr><\/a:outerShdw><\/a:effectLst>/.test(spXml),
        spXml.slice(Math.max(0, spXml.indexOf('<a:ln')), Math.max(0, spXml.indexOf('<a:ln')) + 260));

      // Insert → Link: an address on the pasted box's words, drawn as a link,
      // offered to follow, and an External relationship in the file. Nothing
      // is followed: a check never touches the network.
      await wait(300);
      await js(`(() => { document.querySelector('.sl-hit[data-shape="${painted.id}"]')?.click(); return 1; })()`);
      await wait(300);
      await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click(), 'tab'`);
      await wait(200);
      const linked = await clickRibbon('Link');
      await until(() => js(`Boolean(document.querySelector('.sl-link-url'))`), 'the link dialog', 4000).catch(() => {});
      await js(`(() => { const el = document.querySelector('.sl-link-url'); if (!el) return false; const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set; setter.call(el, 'office.rutba.io/help'); el.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
      await wait(100);
      await js(`(() => { document.querySelector('.sl-link-ok')?.click(); return 1; })()`);
      const linkOf = (s) => (s?.text?.paragraphs || []).flatMap((p) => p.runs || []).map((r) => (r.link && typeof r.link === 'object' ? r.link.url : null)).find(Boolean) || null;
      const boxNow = () => model(1).slide.shapes.find((x) => x.id === painted.id);
      const gotLink = await until(() => linkOf(boxNow()) === 'https://office.rutba.io/help', 'the link on the words', 5000).catch(() => false);
      // The stage redraws a render after the window hears back; the link is
      // the underline or the link blue on the words.
      const drawnLink = await until(() => js(`/text-decoration="underline"|fill="#0563C1"/i.test(document.querySelector('.sl-stage svg')?.innerHTML || '')`), 'the link drawn', 10000).catch(() => false);
      await wait(300);
      const hitTip = await js(`document.querySelector('.sl-hit[data-shape="${painted.id}"]')?.title || ''`);
      check('slides: Insert → Link puts an address on the shape’s words, drawn underlined, with Ctrl+click offered to follow it',
        linked === 'clicked' && gotLink === true && drawnLink === true && / — Ctrl\+click to follow https:\/\/office\.rutba\.io\/help$/.test(hitTip),
        `${linked}; link ${linkOf(boxNow())}; drawn ${drawnLink}; tip ${JSON.stringify(hitTip)}`);
      await clickRibbon('Save');
      const linkFile = () => {
        try {
          const d = Deck.open(fs.readFileSync(files.pptx));
          const part = d.slideParts[1].part;
          const relsPart = part.replace(/\/([^/]+)$/, '/_rels/$1.rels');
          const x = d.pkg.text(part);
          const m = new RegExp('<p:cNvPr id="' + painted.id + '"[\\s\\S]*?</p:sp>').exec(x);
          return { sp: m ? m[0] : '', rels: d.pkg.has(relsPart) ? d.pkg.text(relsPart) : '' };
        } catch { return { sp: '', rels: '' }; }
      };
      await until(() => /TargetMode="External"/.test(linkFile().rels), 'the link in the file', 8000).catch(() => false);
      const saved = linkFile();
      check('slides: the saved file keeps the link as PowerPoint writes it — hlinkClick on the run, an External relationship on the slide',
        /<a:hlinkClick r:id="rId\d+"/.test(saved.sp) && /Type="[^"]*\/hyperlink" Target="https:\/\/office\.rutba\.io\/help" TargetMode="External"/.test(saved.rels),
        `${saved.sp.slice(Math.max(0, saved.sp.indexOf('<a:rPr')), Math.max(0, saved.sp.indexOf('<a:rPr')) + 200)}; rels ${saved.rels.slice(-220)}`);
      // Home → Clear all formatting: Bold put on the pasted box's words, then
      // taken off with everything else — the link stays.
      await wait(300);
      await js(`(() => { document.querySelector('.sl-hit[data-shape="${painted.id}"]')?.click(); return 1; })()`);
      await wait(300);
      await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Home')?.click(), 'tab'`);
      await wait(200);
      const bolded = await clickRibbon('Bold');
      const boxRuns = () => (boxNow()?.text?.paragraphs || []).flatMap((p) => p.runs || []).filter((r) => r.text && r.text !== '\n');
      const wentBold = await until(() => boxRuns().some((r) => r.bold), 'a bold run', 5000).catch(() => false);
      const cleared = await clickRibbon('Clear all formatting');
      const plainAgain = await until(() => boxRuns().length > 0 && boxRuns().every((r) => !r.bold && !r.italic && !r.size && !r.color && !r.font) && boxRuns().some((r) => r.link?.url === 'https://office.rutba.io/help'), 'the formatting off, the link on', 5000).catch(() => false);
      check('slides: Home → Clear all formatting takes Bold and the rest off the words and keeps their link', bolded === 'clicked' && wentBold === true && cleared === 'clicked' && plainAgain === true, `${bolded} ${cleared}; bold first ${wentBold}; runs now ${JSON.stringify(boxRuns().map((r) => ({ text: r.text.slice(0, 12), bold: r.bold ?? null, size: r.size ?? null, link: r.link?.url ?? null })))}`);
      await wait(300);
      await clickRibbon('Save');
      await until(() => /<a:rPr lang="en-US"><a:hlinkClick/.test(linkFile().sp) && !/ b="1"/.test(linkFile().sp), 'the plain runs in the file', 8000).catch(() => false);
      const plainSp = linkFile().sp;
      check('slides: the saved file keeps the words plain, the link on them', /<a:rPr lang="en-US"><a:hlinkClick r:id="rId\d+"/.test(plainSp) && !/ b="1"| sz="/.test(plainSp), plainSp.slice(Math.max(0, plainSp.indexOf('<a:p>')), Math.max(0, plainSp.indexOf('<a:p>')) + 220));
      const complaints = await errorsIn(win);
      check('slides: the clipboard checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('slides: the clipboard checks ran', false, err.message);
    }
  };

  /* ── Worksheets: the fill handle ─────────────────────────────────────── */
  //
  // The square at the selection's corner, dragged down two rows: the cells
  // it passed over are filled from the selection and become the selection.
  const sheetFill = async () => {
    try {
      const win = await open('sheets', files.xlsx);
      const js = (code) => win.webContents.executeJavaScript(code);
      const wc = win.webContents;
      const model = () => doc.model({ id: sessionFor('sheet').id });
      const cell = (ref) => (model().cells || []).find((c) => c.ref === ref);
      await js(`(() => { const c = document.querySelector('.sh-cell[data-ref="B2"]'); c.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 })); return 1; })()`);
      await until(() => cell('B2')?.active === true, 'B2 to be active', 4000);
      await until(() => js(`Boolean(document.querySelector('.sh-fill'))`), 'the fill handle', 4000);
      const at = await js(`(() => { const r = document.querySelector('.sh-fill').getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) }; })()`);
      wc.sendInputEvent({ type: 'mouseMove', x: at.x, y: at.y });
      wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      await wait(60);
      for (const dy of [10, 22, 36]) {
        wc.sendInputEvent({ type: 'mouseMove', x: at.x, y: at.y + dy, button: 'left' });
        await wait(40);
      }
      // The guide is drawn by a render the drag queued; under load it lands
      // after the last move, so it is waited for, the button still down.
      const guide = await until(() => js(`Boolean(document.querySelector('.sh-fillguide'))`), 'the fill guide', 2500).catch(() => false);
      wc.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y + 36, button: 'left', clickCount: 1 });
      const filled = await until(() => String(cell('B4')?.text || '') !== '', 'B4 to be filled', 5000).catch(() => false);
      const picked = (model().cells || []).filter((c) => c.selected).map((c) => c.ref).sort();
      const range = picked.length ? { top: 1, bottom: 1 + picked.length - 1, left: 1, right: 1, refs: picked.join(',') } : null;
      if (range && picked.join(',') !== 'B2,B3,B4') range.bottom = -1;
      check('sheets: dragging the fill handle down fills the cells it passes and selects them', filled === true && guide === true && String(cell('B3')?.text || '') !== '' && range && range.top === 1 && range.bottom === 3 && range.left === 1 && range.right === 1, `B2 ${JSON.stringify(cell('B2')?.text)} → B3 ${JSON.stringify(cell('B3')?.text)}, B4 ${JSON.stringify(cell('B4')?.text)}; selection ${JSON.stringify(range)}; guide ${guide}`);
      const complaints = await errorsIn(win);
      check('sheets: the fill handle reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('sheets: the fill-handle check ran', false, err.message);
    }
  };

  /* ── Rutba Word: picture handles ──────────────────────────────────────── */
  //
  // Click a picture and it wears four corner handles; drag one and the
  // picture grows, keeping its proportions, and the engine writes the size.
  const wordPictures = async () => {
    if (!files.float) return check('word: the picture-handles fixture exists', false, 'no icon to make it from');
    try {
      const win = await open('word', files.float);
      const js = (code) => win.webContents.executeJavaScript(code);
      const wc = win.webContents;
      const image = () => doc.model({ id: sessionFor('doc').id }).blocks[1].images[0];
      await until(() => js(`Boolean(document.querySelector('.wd-image'))`), 'the picture', 8000);
      await js(`(() => { document.querySelector('.wd-image').click(); return 1; })()`);
      const handles = await until(() => js(`document.querySelectorAll('.wd-handle').length === 4`), 'the four handles', 4000).catch(() => false);
      check('word: clicking a picture shows four corner handles', handles === true, `${await js(`document.querySelectorAll('.wd-handle').length`)} handle(s)`);
      const w0 = image().widthPx;
      const at = await js(`(() => { const r = document.querySelector('.wd-handle[data-handle="se"]').getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) }; })()`);
      wc.sendInputEvent({ type: 'mouseMove', x: at.x, y: at.y });
      wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      await wait(60);
      for (const dx of [10, 25, 40]) {
        wc.sendInputEvent({ type: 'mouseMove', x: at.x + dx, y: at.y + 10, button: 'left' });
        await wait(40);
      }
      wc.sendInputEvent({ type: 'mouseUp', x: at.x + 40, y: at.y + 10, button: 'left', clickCount: 1 });
      const grew = await until(() => Math.abs(image().widthPx - (w0 + 40)) <= 2, 'the picture to grow', 5000).catch(() => false);
      const img = image();
      const drawn = await js(`Math.round(document.querySelector('.wd-image').getBoundingClientRect().width)`);
      const kept = await until(() => js(`document.querySelectorAll('.wd-handle').length === 4 && document.querySelector('.wd-image.picked') !== null`), 'the pick to survive the drag', 3000).catch(() => false);
      check('word: dragging a corner handle resizes the picture, keeping its proportions', grew === true && Math.abs(img.heightPx - img.widthPx) <= 2 && Math.abs(drawn - img.widthPx) <= 2 && kept === true, `${w0} → ${img.widthPx}×${img.heightPx} px; drawn ${drawn} px wide; still picked: ${kept}`);
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'word-handles.png'), (await win.webContents.capturePage()).toPNG());
      await press(wc, 's', { modifiers: ['control'] });
      await wait(1200);
      const saved = openDocx(fs.readFileSync(files.float)).render({ pages: false }).blocks[1].images[0];
      check('word: the picture\'s size is saved', Math.abs(saved.widthPx - img.widthPx) <= 1, `file says ${saved.widthPx}×${saved.heightPx}`);
      const complaints = await errorsIn(win);
      check('word: resizing a picture reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('word: the picture-handles check ran', false, err.message);
    }
  };

  /* ── Rutba Word: a picture at the foot of a page ──────────────────────── */
  //
  // A picture whose paragraph sits near a page's foot goes whole to the
  // next page, and one taller than a page is drawn to fit it: neither is
  // cut by a sheet's edge, whenever its bytes arrive.
  const wordPictureFits = async () => {
    try {
      const win = await open('word', files.fit);
      const js = (code) => win.webContents.executeJavaScript(code);
      await until(() => js(`document.querySelectorAll('.wd-image').length === 2 && [...document.querySelectorAll('.wd-image')].every((i) => i.complete && i.naturalHeight > 0)`), 'both pictures', 10000);
      await wait(900);
      const seen = await js(`(() => {
        const sheets = [...document.querySelectorAll('.wd-sheet')].map((s) => s.getBoundingClientRect());
        const on = (r) => sheets.findIndex((s) => r.top >= s.top - 1 && r.bottom <= s.bottom + 1);
        const pictures = [...document.querySelectorAll('.wd-image')].map((img) => { const r = img.getBoundingClientRect(); return { sheet: on(r), top: Math.round(r.top), bottom: Math.round(r.bottom), width: Math.round(r.width), height: Math.round(r.height) }; });
        return { pictures, sheets: sheets.length, sheetHeight: Math.round(sheets[0]?.height || 0) };
      })()`);
      const [card, scan] = seen.pictures;
      check('word: a picture at the foot of a page goes whole to the next, and one taller than a page is drawn to fit it — neither crosses a sheet\'s edge',
        Boolean(card && scan) && card.sheet >= 0 && scan.sheet >= 0 && card.height >= 590 && card.width >= 390 && scan.height < seen.sheetHeight - 150 && scan.height > 700 && scan.width < 500,
        JSON.stringify(seen));
      await wait(300);
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'word-fit.png'), (await win.webContents.capturePage()).toPNG());
      const complaints = await errorsIn(win);
      check('word: the pictures at the foot report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('word: the picture-fit check ran', false, err.message);
    }
  };

  /* ── Worksheets: a picture at the cell ─────────────────────────────────── */
  //
  // The picture goes in behind the window's back — the file dialog cannot
  // be driven — through the same op the Pictures button dispatches, on a
  // copy of the sales workbook; the window draws it over the cell, and the
  // saved file carries it as Excel keeps a picture.
  const sheetPicture = async () => {
    if (!files.png) return check('sheets: the picture fixture exists', false, 'no png');
    try {
      const file = path.join(path.dirname(files.xlsx), 'picture.xlsx');
      fs.copyFileSync(files.xlsx, file);
      const win = await open('sheets', file);
      const js = (code) => win.webContents.executeJavaScript(code);
      // The grid holds nodes for filled and active cells only: D6 exists once it is selected.
      await until(() => js(`document.querySelectorAll('.sh-cell').length > 4`), 'the grid', 8000);
      const applied = await js(`(async () => {
        try {
          const all = await window.rutbaOffice.doc.sessions({});
          const mine = all.filter((s) => s.kind === 'sheet').pop();
          if (!mine) return 'no sheet session among ' + all.map((s) => s.kind).join(',');
          const { bytes, stat } = await window.rutbaOffice.fs.read({ path: ${JSON.stringify(files.png)} });
          const r = await window.rutbaOffice.doc.apply({ id: mine.id, ops: [
            { op: 'select', row: 5, col: 3 },
            { op: 'insertPicture', name: stat.name, contentType: 'image/png', data: bytes, widthPx: 160, heightPx: 120 },
          ] });
          return 'applied: ' + JSON.stringify(r && (r.error || r.ok || Object.keys(r))).slice(0, 160);
        } catch (e) { return 'apply failed: ' + e.message; }
      })()`);
      // The insert went in behind the window's back, so the page has not
      // heard. Undo and redo, from the keyboard, put it through the window's
      // own path — and prove a picture survives both.
      await js(`document.querySelector('.sh')?.focus(), 'focused'`);
      await press(win.webContents, 'z', { modifiers: ['control'] });
      await wait(400);
      await press(win.webContents, 'y', { modifiers: ['control'] });
      await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="D6"]'))`), 'D6 to be drawn as the active cell', 6000).catch(() => {});
      const cell = await js(`(() => { const c = document.querySelector('.sh-cell[data-ref="D6"]'); return c ? { x: parseFloat(c.style.left), y: parseFloat(c.style.top) } : { x: NaN, y: NaN }; })()`);
      const readDrawn = `(() => { const d = [...document.querySelectorAll('.sh-drawing')].find((el) => el.querySelector('image')); return d ? { left: parseFloat(d.style.left), top: parseFloat(d.style.top), width: parseFloat(d.style.width), height: parseFloat(d.style.height), href: (d.querySelector('image').getAttribute('href') || d.querySelector('image').getAttribute('xlink:href') || '').slice(0, 15) } : false; })()`;
      const drawnOk = await until(() => js(readDrawn), 'the picture over the cells', 6000).catch(() => false);
      const drawn = drawnOk === true ? await js(readDrawn) : false;
      check('sheets: Insert → Pictures puts the picture over the cell, at its size, from its own bytes',
        drawn !== false && Math.abs(drawn.left - cell.x) < 1.5 && Math.abs(drawn.top - cell.y) < 1.5 && Math.round(drawn.width) === 160 && Math.round(drawn.height) === 120 && /^data:image\/png/.test(drawn.href),
        `${applied}; ${JSON.stringify({ cell, drawn })}`);
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'sheets-picture.png'), (await win.webContents.capturePage()).toPNG());

      // Save through the quick bar's button (the harness's own click helper is declared further down this file).
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || n.textContent || '').trim().startsWith('Save') && !n.disabled))`), 'the Save button', 3000).catch(() => {});
      await js(`(() => { const b = [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || n.textContent || '').trim().startsWith('Save') && !n.disabled); if (!b) return 'no Save'; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked'; })()`);
      await until(() => {
        try { return SheetView.open(fs.readFileSync(file)).pkg.partNames().some((n) => n.startsWith('xl/media/')); } catch { return false; }
      }, 'the save to land', 8000).catch(() => false);
      const re = SheetView.open(fs.readFileSync(file));
      const media = re.pkg.partNames().filter((n) => n.startsWith('xl/media/'));
      const sheetPart = re.workbook.partNameFor(re.activeSheet);
      const drawingRel = re.pkg.rels(sheetPart).find((r) => String(r.Type).endsWith('/drawing'));
      const drawingXml = drawingRel ? re.pkg.text('xl/' + drawingRel.Target.replace(/^\.\.\//, '')) : '';
      const pictures = (re.render().drawings || []).filter((d) => d.kind === 'image');
      check('sheets: the saved file carries the picture as Excel keeps one — the media part, the drawing part\'s one-cell anchor and its relationship',
        media.length === 1 && /\.png$/.test(media[0]) && /<xdr:oneCellAnchor><xdr:from><xdr:col>3<\/xdr:col>[\s\S]*?<xdr:row>5<\/xdr:row>[\s\S]*?<xdr:pic>/.test(drawingXml) && pictures.length === 1 && re.pkg.contentTypes().defaults.get('png') === 'image/png',
        `media ${JSON.stringify(media)}; pictures ${pictures.length}; anchor ${/<xdr:pic>/.test(drawingXml)}`);
      const complaints = await errorsIn(win);
      check('sheets: the picture reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('sheets: the picture check ran', false, err.message);
    }
  };

  /* ── Rutba Word: a paragraph of pictures splits between its lines ─────── */
  //
  // Six scans of a card in one paragraph, two to a line: the paragraph is
  // taller than a page, so the pass must split it between two picture lines
  // rather than let a line run over the page's edge — the first line under
  // the heading, the other two on the next page.
  const wordCards = async () => {
    try {
      const win = await open('word', files.cards);
      const js = (code) => win.webContents.executeJavaScript(code);
      await until(() => js(`document.querySelectorAll('.wd-image').length === 6 && [...document.querySelectorAll('.wd-image')].every((i) => i.complete && i.naturalHeight > 0)`), 'all six pictures', 10000);
      await wait(900);
      const seen = await js(`(() => {
        const sheets = [...document.querySelectorAll('.wd-sheet')].map((s) => s.getBoundingClientRect());
        const on = (r) => sheets.findIndex((s) => r.top >= s.top - 1 && r.bottom <= s.bottom + 1);
        const pictures = [...document.querySelectorAll('.wd-image')].map((img) => { const r = img.getBoundingClientRect(); return { image: img.dataset.image, sheet: on(r), top: Math.round(r.top), left: Math.round(r.left), bottom: Math.round(r.bottom), height: Math.round(r.height) }; });
        const parts = [...document.querySelectorAll('.wd-block[data-block="1"]')].map((p) => ({ from: p.dataset.from || '0', pictures: p.querySelectorAll('.wd-image').length }));
        return { pictures, parts, sheets: sheets.length };
      })()`);
      const sheetsOf = seen.pictures.map((p) => p.sheet);
      const sideBySide = seen.pictures.length === 6 && seen.pictures[1].top === seen.pictures[0].top && seen.pictures[1].left > seen.pictures[0].left + 200;
      check('word: a paragraph of six pictures lays them two to a line and splits between the lines — the first two under the heading, the four others on page 2, none crossing an edge',
        seen.sheets === 2 && sideBySide && sheetsOf.join(',') === '0,0,1,1,1,1' && seen.parts.length === 2 && seen.parts[1].from === '2',
        JSON.stringify(seen));
      // The caret can go into the second part — the pictures' page — and lands at the end of the paragraph's words, not past them.
      const caret = await js(`(() => {
        const part = document.querySelectorAll('.wd-block[data-block="1"]')[1];
        if (!part) return null;
        const r = document.createRange(); r.selectNodeContents(part); r.collapse(false);
        const s = getSelection(); s.removeAllRanges(); s.addRange(r);
        document.querySelector('.wd-page').dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
        return 'placed';
      })()`);
      await wait(400);
      const complaints = await errorsIn(win);
      check('word: the caret goes into the pictures\' second page without an error', caret === 'placed' && complaints.length === 0, complaints.join(' | ') || 'nothing reported');
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'word-cards.png'), (await win.webContents.capturePage()).toPNG());
    } catch (err) {
      check('word: the cards check ran', false, err.message);
    }
  };

  /* ── Rutba Word: the ruler and the grips on a table ───────────────────── */
  //
  // The ruler's markers and the grips on a table's borders are dragged with
  // real pointer events; what they wrote is read from the engine, from the
  // drawing (computed style, cell rectangles) and, at the end, from the file.
  const wordRuler = async () => {
    try {
      const win = await open('word', files.ruler);
      const js = (code) => win.webContents.executeJavaScript(code);
      const wc = win.webContents;
      const modelOf = () => doc.model({ id: sessionFor('doc').id });
      const rect = (selector) => js(`(() => { const r = document.querySelector('${selector}')?.getBoundingClientRect(); return r ? { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2), left: r.left, right: r.right, top: r.top, bottom: r.bottom, width: r.width, height: r.height } : null; })()`);
      const drag = async (from, to, steps = 3) => {
        wc.sendInputEvent({ type: 'mouseMove', x: from.x, y: from.y });
        wc.sendInputEvent({ type: 'mouseDown', x: from.x, y: from.y, button: 'left', clickCount: 1 });
        await wait(60);
        for (let i = 1; i <= steps; i++) {
          wc.sendInputEvent({ type: 'mouseMove', x: Math.round(from.x + ((to.x - from.x) * i) / steps), y: Math.round(from.y + ((to.y - from.y) * i) / steps), button: 'left' });
          await wait(40);
        }
        wc.sendInputEvent({ type: 'mouseUp', x: to.x, y: to.y, button: 'left', clickCount: 1 });
      };
      const click = async (p) => {
        wc.sendInputEvent({ type: 'mouseMove', x: p.x, y: p.y });
        wc.sendInputEvent({ type: 'mouseDown', x: p.x, y: p.y, button: 'left', clickCount: 1 });
        await wait(40);
        wc.sendInputEvent({ type: 'mouseUp', x: p.x, y: p.y, button: 'left', clickCount: 1 });
      };
      const caretTo = async (block) => {
        await click(await rect(`[data-block="${block}"]`));
        await until(() => modelOf().selection?.focus?.block === block, 'the caret in block ' + block, 5000);
      };

      await until(() => js(`Boolean(document.querySelector('.wd-ruler'))`), 'the ruler', 8000);
      check('word: the ruler is shown above the page', true);

      // The paragraph's indents, as drawn: the first line 48 px in.
      await caretTo(1);
      await until(() => js(`Boolean(document.querySelector('.wd-ruler-first') && document.querySelector('.wd-ruler-hang'))`), 'the indent markers', 5000);
      const gap = () => js(`Math.round(document.querySelector('.wd-ruler-first').getBoundingClientRect().left - document.querySelector('.wd-ruler-hang').getBoundingClientRect().left)`);
      await until(async () => Math.abs((await gap()) - 48) <= 2, 'the first-line marker 48 px in', 5000).catch(() => {});
      check('word: the ruler shows the paragraph\'s first-line indent 48 px in from its left indent', Math.abs((await gap()) - 48) <= 2, `${await gap()} px`);

      // Drag the first-line marker 30 px to the right: the file and the words follow.
      const first = await rect('.wd-ruler-first');
      await drag(first, { x: first.x + 30, y: first.y });
      const moved = await until(() => Math.abs((modelOf().blocks[1].firstLinePx ?? 0) - 78) <= 1, 'the first-line indent to grow', 5000).catch(() => false);
      const drawn = await js(`getComputedStyle(document.querySelector('[data-block="1"]')).textIndent`);
      check('word: dragging the first-line marker writes the indent and the first line moves with it', moved === true && Math.abs(parseFloat(drawn) - 78) <= 1,
        `engine ${modelOf().blocks[1].firstLinePx} px, drawn ${drawn}`);

      // The tab stops of the tabbed line, where the file puts them.
      await caretTo(2);
      await until(() => js(`document.querySelectorAll('.wd-ruler-tab').length === 2`), 'two tab stops on the ruler', 5000).catch(() => {});
      const stops = await js(`(() => { const ruler = document.querySelector('.wd-ruler').getBoundingClientRect(); return [...document.querySelectorAll('.wd-ruler-tab')].map((t) => ({ pos: Number(t.dataset.pos), align: t.dataset.align, x: Math.round(t.getBoundingClientRect().left + 4 - ruler.left) })); })()`);
      const ML = Math.round(modelOf().section.margins.left);
      check('word: the ruler shows the paragraph\'s two tab stops where the file puts them',
        stops.length === 2 && stops[0].pos === 96 && stops[0].align === 'left' && Math.abs(stops[0].x - (ML + 96)) <= 1 && stops[1].pos === 384 && stops[1].align === 'right' && Math.abs(stops[1].x - (ML + 384)) <= 1,
        JSON.stringify(stops));

      // A click on the bare ruler adds a stop there; dragging it off removes it.
      const ruler = await rect('.wd-ruler');
      await click({ x: Math.round(ruler.left + ML + 200), y: ruler.y });
      const added = await until(() => modelOf().blocks[2].tabs?.length === 3, 'a third tab stop', 5000).catch(() => false);
      const third = (modelOf().blocks[2].tabs || []).find((t) => Math.abs(t.posPx - 200) <= 2);
      check('word: a click on the ruler adds a tab stop where it was clicked', added === true && Boolean(third) && third.align === 'left', JSON.stringify(modelOf().blocks[2].tabs));
      await until(() => js(`Boolean(document.querySelector('.wd-ruler-tab[data-pos="' + ${third ? Math.round(third.posPx) : -1} + '"]'))`), 'the new stop drawn', 4000).catch(() => {});
      const newStop = await rect(`.wd-ruler-tab[data-pos="${third ? Math.round(third.posPx) : -1}"]`);
      if (newStop) await drag(newStop, { x: newStop.x, y: newStop.y + 60 });
      const removed = await until(() => modelOf().blocks[2].tabs?.length === 2, 'the stop to be removed', 5000).catch(() => false);
      check('word: a tab stop dragged off the ruler is removed', removed === true, JSON.stringify(modelOf().blocks[2].tabs));

      // Into the table: grips on its column edges on the page, and on the ruler.
      await caretTo(3);
      const grips = await until(() => js(`document.querySelectorAll('.wd-tgrip.col').length === 3 && document.querySelectorAll('.wd-ruler-col').length === 3 && document.querySelectorAll('.wd-tgrip.row').length === 2`), 'the table grips', 5000).catch(() => false);
      check('word: with the caret in a table, grips sit on its column edges and row bottoms, and the ruler shows the columns', grips === true,
        `${await js(`document.querySelectorAll('.wd-tgrip.col').length`)} column grip(s), ${await js(`document.querySelectorAll('.wd-ruler-col').length`)} on the ruler`);
      const cellBefore = await rect('[data-block="3"]');
      const grip = await rect('.wd-tgrip.col[data-k="1"]');
      // Taken hold of inside the first row: at the table's middle a row grip crosses it.
      await drag({ x: grip.x, y: Math.round(grip.top + 8) }, { x: grip.x + 40, y: Math.round(grip.top + 8) });
      const resized = await until(() => { const g = modelOf().blocks[3].gridPx; return Boolean(g) && Math.abs(g[0] - 200) <= 1 && Math.abs(g[1] - 160) <= 1 && Math.abs(g[2] - 240) <= 1; }, 'the two columns to change', 5000).catch(() => false);
      const cellAfter = await until(async () => { const r = await rect('[data-block="3"]'); return r && Math.abs(r.width - (cellBefore.width + 40)) <= 2 ? r : null; }, 'the cell to be drawn wider', 4000).catch(() => null);
      check('word: dragging a column border on the page resizes the columns either side of it, the table keeping its width',
        resized === true && Boolean(cellAfter), `grid ${JSON.stringify(modelOf().blocks[3].gridPx)}, cell ${cellBefore.width} → ${cellAfter?.width ?? (await rect('[data-block="3"]'))?.width}`);

      // A row's bottom border dragged down 20 px.
      const rowBefore = await js(`(() => { const tr = document.querySelector('.wd-table tbody tr'); return tr.getBoundingClientRect().height; })()`);
      const rowGrip = await rect('.wd-tgrip.row[data-r="0"]');
      await drag(rowGrip, { x: rowGrip.x, y: rowGrip.y + 20 });
      const taller = await until(() => Math.abs((modelOf().blocks[3].rowHeightPx ?? 0) - (rowBefore + 20)) <= 2, 'the row height to be written', 5000).catch(() => false);
      const rowAfter = await until(async () => { const h = await js(`document.querySelector('.wd-table tbody tr').getBoundingClientRect().height`); return Math.abs(h - (rowBefore + 20)) <= 2 ? h : null; }, 'the row drawn taller', 4000).catch(() => null);
      check('word: dragging a row border on the page sets the row\'s height', taller === true && rowAfter != null, `row ${rowBefore} → ${rowAfter}, engine ${modelOf().blocks[3].rowHeightPx}`);
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'word-ruler.png'), (await win.webContents.capturePage()).toPNG());

      // The left margin's edge on the ruler, dragged 24 px in.
      const bound = await rect('.wd-ruler-bound[data-bound="left"]');
      await drag(bound, { x: bound.x + 24, y: bound.y });
      const margin = await until(() => Math.abs(modelOf().section.margins.left - (ML + 24)) <= 1, 'the margin to move', 5000).catch(() => false);
      const padding = await js(`parseFloat(getComputedStyle(document.querySelector('.wd-page')).paddingLeft)`);
      check('word: dragging the margin edge on the ruler moves the page margin', margin === true && Math.abs(padding - (ML + 24)) <= 1, `margin ${modelOf().section.margins.left} px, page padding ${padding} px`);

      // The file has all of it.
      await press(wc, 's', { modifiers: ['control'] });
      await wait(1200);
      const saved = openDocx(fs.readFileSync(files.ruler)).render({ pages: false });
      const cell = saved.blocks[3];
      check('word: the ruler\'s and the grips\' changes are saved',
        Math.abs(saved.blocks[1].firstLinePx - 78) <= 1 && saved.blocks[2].tabs?.length === 2 && cell.gridPx && Math.abs(cell.gridPx[0] - 200) <= 1 && Math.abs(cell.gridPx[1] - 160) <= 1
          && Math.abs((cell.rowHeightPx ?? 0) - (rowBefore + 20)) <= 2 && Math.abs(saved.section.margins.left - (ML + 24)) <= 1,
        JSON.stringify({ firstLine: saved.blocks[1].firstLinePx, tabs: saved.blocks[2].tabs?.length, grid: cell.gridPx, row: cell.rowHeightPx, margin: saved.section.margins.left }));
      const complaints = await errorsIn(win);
      check('word: the ruler and the grips report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('word: the ruler check ran', false, err.message);
    }
  };

  /* ── Word: Layout → Columns ──────────────────────────────────────────── */
  //
  // Columns used to be a Soon button. This picks Two from the ribbon on the
  // long report, checks the section says so and the screen is honest about
  // laying the single flow as wide as the first column only, saves and reads
  // the file's own `w:cols` back, prints it and checks the PDF actually put
  // words at two different x positions on the first page, then turns columns
  // off again and checks the element is gone.
  const wordColumns = async () => {
    try {
      const win = await open('word', files.long);
      const js = (code) => win.webContents.executeJavaScript(code);
      const modelOf = () => doc.model({ id: sessionFor('doc').id });
      const clickRibbon = (title) => js(`(() => {
        const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}));
        if (!b) return 'no button ' + ${JSON.stringify(title)};
        b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
        b.click();
        return 'clicked';
      })()`);
      const pick = async (button, label) => {
        const clicked = await clickRibbon(button);
        await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}))`), `the ${button} menu`, 3000).catch(() => {});
        await js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)})?.click(); return 1; })()`);
        return clicked;
      };

      await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="0"]'))`), 'the long report', 8000);
      await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Layout')?.click(), 'tab'`);
      await wait(200);

      const pickedTwo = await pick('Columns', 'Two');
      const two = await until(() => modelOf().section?.columns?.count === 2, 'the section to say two columns', 5000).catch(() => false);
      await wait(300);
      const boxes = modelOf().section.columns;
      const contentWidthPx = modelOf().section.widthPx - modelOf().section.margins.left - modelOf().section.margins.right - (modelOf().section.margins.gutter || 0);
      const col0WidthPx = boxes.widths ? boxes.widths[0] : (contentWidthPx - (boxes.spacePx || 0)) / 2;
      const screen = await js(`(() => {
        const p = document.querySelector('.wd-page');
        const cs = getComputedStyle(p);
        return { content: p.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight), chip: document.querySelector('.wd-columns-chip')?.textContent || null };
      })()`);
      check('word: Layout → Columns → Two writes two columns on the section', pickedTwo === 'clicked' && two === true, `${pickedTwo}; ${JSON.stringify(modelOf().section?.columns)}`);
      check('word: the flow on screen is laid as wide as the first column, honestly, and the status bar says so',
        Math.abs(screen.content - col0WidthPx) <= 2 && /2 columns.*laid as one on screen.*flowed into columns in print/.test(screen.chip || ''),
        `content ${screen.content} px vs column ${col0WidthPx} px; chip ${JSON.stringify(screen.chip)}`);
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'word-columns.png'), (await win.webContents.capturePage()).toPNG());

      await clickRibbon('Save');
      await until(() => { try { return /<w:cols/.test(openDocx(fs.readFileSync(files.long)).doc.doc.xml); } catch { return false; } }, 'the columns to land in the file', 8000).catch(() => {});
      const sectPr = (() => { try { const m = /<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/.exec(openDocx(fs.readFileSync(files.long)).doc.doc.xml); return m ? m[0] : ''; } catch { return ''; } })();
      check('word: the saved file keeps the columns as w:cols in the section', /<w:cols w:num="2" w:space="720"\/>/.test(sectPr), sectPr.slice(0, 200));

      // Printed: the same document's PDF actually flows the words into two
      // columns, not just one — the same path Ctrl+P and the Print button use.
      const id = sessionFor('doc').id;
      const target = path.join(path.dirname(files.long), 'print-word-columns.pdf');
      await js(`window.rutbaOffice.print.pdf({ id: ${JSON.stringify(id)}, path: ${JSON.stringify(target)}, options: {} })`);
      const pdfBytes = fs.readFileSync(target);
      const pdfText = pdfBytes.toString('latin1');
      const xs = [...pdfText.matchAll(/1 0 0 1 ([\d.]+) [\d.]+ Tm/g)].map((m) => Math.round(parseFloat(m[1])));
      const distinctXs = [...new Set(xs)];
      check('word: the printed PDF draws the report\'s words from two different x positions on the page', distinctXs.length >= 2, `x origins seen: ${distinctXs.slice(0, 6).join(', ')}`);

      // Columns off again: the element is gone from the section and the file.
      const pickedOne = await pick('Columns', 'One');
      const one = await until(() => (modelOf().section?.columns?.count ?? 1) === 1, 'the section to say one column again', 5000).catch(() => false);
      check('word: Layout → Columns → One takes the columns off again', pickedOne === 'clicked' && one === true, `${pickedOne}; ${JSON.stringify(modelOf().section?.columns)}`);
      await clickRibbon('Save');
      await until(() => { try { return !/<w:cols/.test(openDocx(fs.readFileSync(files.long)).doc.doc.xml); } catch { return false; } }, 'the columns to leave the file', 8000).catch(() => {});
      check('word: the saved file drops w:cols again for one column', !/<w:cols/.test(openDocx(fs.readFileSync(files.long)).doc.doc.xml), 'w:cols still in the file');

      const complaints = await errorsIn(win);
      check('word: Layout → Columns reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('word: the columns check ran', false, err.message);
    }
  };

  /* ── The update prompt, in every window ──────────────────────────────── */
  //
  // A check run contacts nothing, so the service never finds a release here;
  // the states it would publish are broadcast by hand and a Word window —
  // not the launcher — is read, because that is where a person is all day.
  const updatePrompt = async () => {
    if (!broadcast) return check('frame: the update prompt check has the shell\'s broadcast', false, 'verifyApps was not given broadcast');
    try {
      const win = await open('word', files.docx);
      const js = (code) => win.webContents.executeJavaScript(code);
      await until(() => js(`Boolean(document.querySelector('.wd-page'))`), 'the Word window', 8000);
      const base = { version: '1.12.0', automatic: true, snoozed: false, snoozedVersion: null, error: null, notes: null, releasedAt: '2026-09-14T11:40:44.266Z', channel: 'eharain/Rutba-Office' };
      const text = () => js(`document.querySelector('.rw-update')?.textContent.slice(0, 200) ?? 'no prompt'`);
      // On screen, not only in the tree: a box inside the window, and the
      // element under its centre is the prompt (nothing covers it).
      const seen = () => js(`(() => {
        const p = document.querySelector('.rw-update');
        if (!p) return { seen: false, why: 'no prompt' };
        const r = p.getBoundingClientRect();
        const cs = getComputedStyle(p);
        const hit = document.elementFromPoint(r.left + r.width / 2, r.top + r.height / 2);
        const inside = r.width > 200 && r.height > 60 && r.right <= innerWidth + 1 && r.bottom <= innerHeight + 1 && r.left >= 0 && r.top >= 0;
        return { seen: inside && cs.visibility !== 'hidden' && Number(cs.opacity) > 0.9 && Boolean(hit && p.contains(hit)), box: [Math.round(r.left), Math.round(r.top), Math.round(r.width), Math.round(r.height)], window: [innerWidth, innerHeight], position: cs.position, opacity: cs.opacity, under: hit ? hit.className || hit.tagName : null };
      })()`);

      broadcast('update:state', { ...base, state: 'downloading', available: '9.9.9', percent: 42, transferred: 42, total: 100 });
      const downloading = await until(() => js(`(() => { const p = document.querySelector('.rw-update'); return Boolean(p) && p.dataset.state === 'downloading' && /9\\.9\\.9/.test(p.textContent) && /42%/.test(p.textContent) && !p.querySelector('.rw-update-restart'); })()`), 'the prompt while downloading', 5000).catch(() => false);
      const shown = await until(async () => (await seen()).seen === true, 'the prompt on screen', 4000).catch(() => false);
      check('frame: a release the check finds is announced in the window, with the download\'s progress', downloading === true && shown === true, `${JSON.stringify(await seen())} ${(await text()).slice(0, 90)}`);

      broadcast('update:state', { ...base, state: 'ready', available: '9.9.9', percent: 100 });
      const ready = await until(() => js(`(() => { const p = document.querySelector('.rw-update'); return Boolean(p) && p.dataset.state === 'ready' && /ready to install/.test(p.textContent) && Boolean(p.querySelector('.rw-update-restart')); })()`), 'the prompt once downloaded', 5000).catch(() => false);
      check('frame: once downloaded, the prompt offers Restart and update', ready === true, await text());
      // A frame after the prompt has painted, not the one before it: an off-screen window paints on demand.
      await wait(500);
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'frame-update.png'), (await win.webContents.capturePage()).toPNG());

      // Not now: the prompt goes, the service keeps the version for a day,
      // and the same state again does not bring it back.
      await js(`(() => { document.querySelector('.rw-update-later').click(); return 1; })()`);
      const gone = await until(() => js(`!document.querySelector('.rw-update')`), 'the prompt to go', 4000).catch(() => false);
      const remembered = update ? await until(() => update.state().snoozedVersion === '9.9.9', 'the snooze to be kept', 4000).catch(() => false) : 'no service';
      broadcast('update:state', { ...base, state: 'ready', available: '9.9.9', percent: 100 });
      await wait(400);
      const stayedAway = await js(`!document.querySelector('.rw-update')`);
      check('frame: Not now puts the prompt away for that version, and the service keeps the snooze for a day',
        gone === true && remembered === true && stayedAway === true, `gone ${gone}, snoozed ${update ? update.state().snoozedVersion : 'n/a'}, stayed away ${stayedAway}`);

      // A newer release is a new question.
      broadcast('update:state', { ...base, state: 'ready', available: '9.9.10', percent: 100 });
      const again = await until(() => js(`(() => { const p = document.querySelector('.rw-update'); return Boolean(p) && /9\\.9\\.10/.test(p.textContent); })()`), 'the prompt for a newer release', 5000).catch(() => false);
      check('frame: a newer release is announced even after an older one was put away', again === true, await text());
      await js(`(() => { document.querySelector('.rw-update-later')?.click(); return 1; })()`);
      if (update) update.snooze({ version: null });

      // The chip in the status bar: while a download runs it counts, and
      // once the download is done it brings a prompt that was put away back.
      broadcast('update:state', { ...base, state: 'downloading', available: '9.9.11', percent: 63, transferred: 63, total: 100 });
      const chipCounts = await until(() => js(`/63%/.test(document.querySelector('.rw-status .rw-update-chip')?.textContent || '')`), 'the status chip while downloading', 4000).catch(() => false);
      broadcast('update:state', { ...base, state: 'ready', available: '9.9.11', percent: 100 });
      await until(() => js(`Boolean(document.querySelector('.rw-update-later'))`), 'the prompt for 9.9.11', 4000).catch(() => false);
      await js(`(() => { document.querySelector('.rw-update-later')?.click(); return 1; })()`);
      await wait(200);
      broadcast('update:state', { ...base, state: 'ready', available: '9.9.11', percent: 100 });
      const putAway = await until(() => js(`!document.querySelector('.rw-update') && /Update ready/.test(document.querySelector('.rw-status .rw-update-chip')?.textContent || '')`), 'the chip once the prompt is put away', 4000).catch(() => false);
      await js(`(() => { document.querySelector('.rw-status .rw-update-chip')?.click(); return 1; })()`);
      const broughtBack = await until(() => js(`Boolean(document.querySelector('.rw-update-restart'))`), 'the prompt brought back by the chip', 4000).catch(() => false);
      check('frame: the status bar counts the download, and its chip brings a prompt that was put away back', chipCounts === true && putAway === true && broughtBack === true, `counts ${chipCounts}, put away ${putAway}, back ${broughtBack}`);
      await js(`(() => { document.querySelector('.rw-update-later')?.click(); return 1; })()`);
      if (update) update.snooze({ version: null });

      // The first window of a new version says so, and What's new opens the
      // release's own notes, bundled with the build.
      const bundled = JSON.parse(fs.readFileSync(path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', 'build', 'out', 'whatsnew.json'), 'utf8'));
      const firstHeading = (/^##\s+(.+)$/m.exec(bundled.notes || '')?.[1] || '').replace(/`/g, '');
      broadcast('update:state', { ...base, state: 'current', available: null, arrived: { from: '1.12.0', to: bundled.version, at: Date.now(), seen: false } });
      const said = await until(() => js(`(() => { const p = document.querySelector('.rw-update'); return Boolean(p) && p.dataset.state === 'arrived' && p.textContent.includes('is now ${bundled.version}') && p.textContent.includes('from 1.12.0') && Boolean(p.querySelector('.rw-update-whatsnew')); })()`), 'the arrival card', 5000).catch(() => false);
      await js(`(() => { document.querySelector('.rw-update-whatsnew')?.click(); return 1; })()`);
      const notesShown = await until(() => js(`(() => { const d = document.querySelector('.rw-whatsnew'); return Boolean(d) && d.textContent.includes(${JSON.stringify(bundled.version)}) && ${JSON.stringify(firstHeading)} !== '' && d.textContent.includes(${JSON.stringify(firstHeading)}); })()`), 'the notes', 6000).catch(() => false);
      await wait(500);
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'frame-whatsnew.png'), (await win.webContents.capturePage()).toPNG());
      await js(`(() => { document.querySelector('.rw-whatsnew-close')?.click(); return 1; })()`);
      const seenNow = update ? await until(() => { const a = update.state().arrived; return a == null || a.seen === true; }, 'the arrival marked seen', 4000).catch(() => false) : 'no service';
      const cardGone = await until(() => js(`!document.querySelector('.rw-update') && !document.querySelector('.rw-whatsnew')`), 'the card to go', 4000).catch(() => false);
      check('frame: the first window of a new version says which version arrived, What\'s new shows the release\'s notes, and Close marks it seen through the service', said === true && notesShown === true && seenNow === true && cardGone === true, `said ${said}, notes ${notesShown} ("${firstHeading.slice(0, 40)}"), seen ${seenNow}, gone ${cardGone}`);

      const complaints = await errorsIn(win);
      check('frame: the update prompt reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('frame: the update prompt check ran', false, err.message);
    }
  };

  /* ── Pictures: a large, mixed folder ─────────────────────────────────── */
  //
  // The viewer on three hundred pictures with clips among them (verify-viewer.js).
  const viewer = async () => {
    const capture = async (win, name) => {
      if (!process.env.RUTBA_VERIFY_CAPTURE) return;
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await win.webContents.capturePage()).toPNG());
    };
    await verifyViewer({ open, check, until, wait, errorsIn, capture }, { dir, wav: files.wav });
  };

  /* ── Pictures: a clip's length on its tile, and the slideshow ────────── */
  const slideshow = async () => {
    const clickIn = async (win, title) => {
      const find = `[...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || n.textContent || '').trim().startsWith(${JSON.stringify(title)}) && !n.disabled)`;
      await until(() => win.webContents.executeJavaScript(`Boolean(${find})`), `the ${title} button`, 3000).catch(() => {});
      return win.webContents.executeJavaScript(`(() => { const b = ${find}; if (!b) return 'no button ' + ${JSON.stringify(title)}; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked'; })()`);
    };
    try {
      // The folder itself, so the picture is browsed rather than opened —
      // the ribbon is never collapsed there, which the slideshow's own
      // controls need to be reachable.
      const win = await open('pictures', dir);
      const js = (code) => win.webContents.executeJavaScript(code);
      await until(() => js(`document.querySelectorAll('.pv-tile').length >= 3`), 'the folder’s tiles (two pictures, one clip)', 6000);

      const badgeShown = await until(
        () => js(`/^\\d+:\\d{2}$/.test(document.querySelector('.pv-tile[data-name="tone.wav"] .pv-length')?.textContent || '')`),
        'the clip’s length badge',
        8000
      ).catch(() => false);
      const badge = await js(`document.querySelector('.pv-tile[data-name="tone.wav"] .pv-length')?.textContent || ''`);
      check('pictures: a clip’s length is on its tile, as m:ss', badgeShown === true, `badge "${badge}"`);

      // The View tab, where Fit/100% and the slideshow live; 2 seconds
      // chosen before the show starts, so the interval check does not wait
      // out the four-second default.
      await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'View')?.click()`);
      await until(() => js(`Boolean(document.querySelector('select[data-role="show-seconds"]'))`), 'the slideshow’s seconds menu', 4000);
      await js(`(() => { const s = document.querySelector('select[data-role="show-seconds"]'); const set = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set; set.call(s, '2'); s.dispatchEvent(new Event('change', { bubbles: true })); })()`);

      const shownSrc = `(document.querySelector('.pv-show-layer.on')?.src || document.querySelector('.pv-show-media')?.src || '')`;
      await clickIn(win, 'Slideshow');
      await until(() => js(`Boolean(document.querySelector('.pv-show')) && ${shownSrc}.length > 0`), 'the show to open, on the first picture', 4000);
      const first = await js(shownSrc);
      check('pictures: the slideshow opens full-window, on the first picture', Boolean(first), first ? first.slice(-28) : 'nothing shown');

      const movedOn = await until(
        () => js(`${shownSrc} !== ${JSON.stringify(first)} && ${shownSrc}.length > 0`),
        'the show to move on past its interval',
        6000
      ).catch(() => false);
      const second = movedOn ? await js(shownSrc) : first;
      check('pictures: it advances to the next picture once its interval has passed', movedOn === true, `${first.slice(-28)} → ${movedOn ? second.slice(-28) : '(unchanged)'}`);

      await press(win.webContents, 'Space');
      const paused = await until(() => js(`Boolean(document.querySelector('.pv-show-chip'))`), 'the Paused chip', 3000).catch(() => false);
      check('pictures: Space pauses the show, with a chip that says so', paused === true, `paused chip shown: ${paused}`);

      await press(win.webContents, 'Escape');
      const left = await until(() => js(`!document.querySelector('.pv-show')`), 'the show to close on Escape', 4000).catch(() => false);
      // The show may have reached the clip, or a PDF another block printed
      // into the folder, by the time Escape lands, so their viewers count
      // as much as a picture's.
      const stoppedOn = await js(`document.querySelector('.pv-image, .pv-video, .pv-audio, .pv-pdf')?.src || ''`);
      check('pictures: Escape leaves the show and selects the picture it stopped on', left === true && Boolean(stoppedOn), `left ${left}, showing ${stoppedOn.slice(-28)}`);

      const complaints = await errorsIn(win);
      check('pictures: the slideshow and the length badge report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('pictures: the slideshow checks ran', false, err.message);
    }
  };

  /* ── Worksheets: a note, a link, and a link put on a cell ────────────── */
  const sheetLinks = async () => {
    const capture = async (win, name) => {
      if (!process.env.RUTBA_VERIFY_CAPTURE) return;
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await win.webContents.capturePage()).toPNG());
    };
    await verifySheetLinks({ open, check, until, wait, press, errorsIn, capture }, { file: files.notes });
  };

  /* ── Word: table of contents — a real field, not text ────────────────── */
  const wordToc = async () => {
    const capture = async (win, name) => {
      if (!process.env.RUTBA_VERIFY_CAPTURE) return;
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await win.webContents.capturePage()).toPNG());
    };
    await verifyWordToc({ open, check, until, wait, press, errorsIn, capture, doc, sessionFor }, { file: files.toc });
  };

  /* ── Worksheets: Group, the outline gutter, Subtotal ─────────────────── */
  const sheetOutline = async () => {
    const capture = async (win, name) => {
      if (!process.env.RUTBA_VERIFY_CAPTURE) return;
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await win.webContents.capturePage()).toPNG());
    };
    await verifyOutline({ open, check, until, wait, press, errorsIn, capture }, { dir });
  };

  /* ── Word: track changes — recording, not just reading ───────────────── */
  const wordTrack = async () => {
    const capture = async (win, name) => {
      if (!process.env.RUTBA_VERIFY_CAPTURE) return;
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await win.webContents.capturePage()).toPNG());
    };
    await verifyWordTrack({ open, check, until, wait, press, errorsIn, capture, doc, sessionFor }, { file: files.track });
  };

  /* ── The launcher's recent list ──────────────────────────────────────── */
  //
  // Thirty recent files in a window of ordinary height: the list scrolls
  // inside its box with the last row whole at the end (it used to be
  // clipped, the last row in reach cut in half and nothing to scroll), and a
  // right-click offers to take a file off the list.
  const launcherRecent = async () => {
    try {
      const copies = [];
      for (let i = 1; i <= 34; i++) {
        const p = path.join(dir, `recent-${String(i).padStart(2, '0')}.docx`);
        fs.copyFileSync(files.docx, p);
        copies.push(p);
      }
      // The list is the shell's; a Word window adds to it, and the launcher
      // opened afterwards reads it.
      const word = await open('word', files.docx);
      await word.webContents.executeJavaScript(`(async () => { for (const p of ${JSON.stringify(copies)}) await window.rutbaOffice.app.addRecent({ path: p, app: 'word' }); return 1; })()`);
      const win = await open('home');
      const js = (code) => win.webContents.executeJavaScript(code);
      await until(() => js(`document.querySelectorAll('.home-recent-row').length >= 30`), 'the recent list', 8000);
      const list = await js(`(() => {
        const box = document.querySelector('.home-recent');
        const rows = box.querySelectorAll('.home-recent-row');
        box.scrollTop = 1e6;
        const last = rows[rows.length - 1].getBoundingClientRect();
        const b = box.getBoundingClientRect();
        return { rows: rows.length, scrolls: box.scrollHeight > box.clientHeight + 2, overflow: getComputedStyle(box).overflowY, box: Math.round(b.height), lastWhole: last.bottom <= b.bottom + 1 && last.top >= b.top - 1, window: innerHeight };
      })()`);
      check('launcher: thirty recent files scroll inside the list, and the last row is whole at the end', list.rows >= 30 && list.scrolls === true && list.overflow === 'auto' && list.lastWhole === true, JSON.stringify(list));
      // An off-screen window paints on demand, and its last frame was the
      // empty list before the recent files arrived: ask for a fresh one.
      win.webContents.invalidate();
      await wait(700);
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'home-recent.png'), (await win.webContents.capturePage()).toPNG());

      // A right-click offers to take a file off the list, and does.
      const first = await js(`document.querySelector('.home-recent-row')?.title`);
      await js(`(() => { const row = document.querySelector('.home-recent-row'); const r = row.getBoundingClientRect(); row.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: r.left + 30, clientY: r.top + 10 })); return 1; })()`);
      const offered = await until(() => js(`[...document.querySelectorAll('.rw-menu button')].map((b) => b.textContent.trim())`).then((names) => (names.includes('Remove from the list') && names.includes('Show in folder') && names.includes('Clear the list') && names.includes('Open') ? names : false)), 'the row menu', 4000).catch(() => false);
      await js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((b) => /Remove from the list/.test(b.textContent))?.click(); return 1; })()`);
      const removed = await until(() => js(`(() => { const rows = [...document.querySelectorAll('.home-recent-row')]; return rows.length >= 29 && !rows.some((r) => r.title === ${JSON.stringify(first)}) && !document.querySelector('.rw-menu'); })()`), 'the row to go', 4000).catch(() => false);
      check('launcher: a right-click on a recent file offers Open, Show in folder, Remove from the list and Clear the list, and Remove takes it off', offered !== false && removed === true, `offered ${JSON.stringify(offered)}, removed ${removed}`);
      const complaints = await errorsIn(win);
      check('launcher: the recent list reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('launcher: the recent list check ran', false, err.message);
    }
  };

  /* ── The launcher's recent list: rename and remove in place ──────────── */
  //
  // Rename… turns a row's own name into a text box rather than a dialog,
  // and the small × a hover reveals removes a row with no menu at all —
  // both act on the file, not only the list, so the check follows each one
  // onto disk.
  const homeRecentEdit = async () => {
    try {
      const a = path.join(dir, 'recent-edit-a.docx');
      const b = path.join(dir, 'recent-edit-b.docx');
      const taken = path.join(dir, 'recent-edit-taken.docx');
      fs.copyFileSync(files.docx, a);
      fs.copyFileSync(files.docx, b);
      fs.copyFileSync(files.docx, taken);
      // The list is the shell's; a Word window adds to it, and the launcher
      // opened afterwards reads it — the same arrangement launcherRecent uses.
      const word = await open('word', files.docx);
      await word.webContents.executeJavaScript(
        `(async () => { await window.rutbaOffice.app.addRecent({ path: ${JSON.stringify(b)}, app: 'word' }); await window.rutbaOffice.app.addRecent({ path: ${JSON.stringify(a)}, app: 'word' }); return 1; })()`
      );
      const win = await open('home');
      const js = (code) => win.webContents.executeJavaScript(code);
      const setValue = (selector, value) =>
        js(`(() => { const el = document.querySelector(${JSON.stringify(selector)}); if (!el) return false; const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set; setter.call(el, ${JSON.stringify(value)}); el.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
      const pressEnterOn = (selector) =>
        js(`(() => { const el = document.querySelector(${JSON.stringify(selector)}); el?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })); return 1; })()`);
      const rightClickFirstRow = () =>
        js(`(() => { const row = document.querySelector('.home-recent-row'); const r = row.getBoundingClientRect(); row.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: r.left + 30, clientY: r.top + 10 })); return 1; })()`);
      const clickMenuItem = (pattern) =>
        js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((b) => ${pattern}.test(b.textContent))?.click(); return 1; })()`);

      await until(() => js(`document.querySelectorAll('.home-recent-row').length >= 2`), 'the recent list', 8000);

      // Rename the row on top — the file added last, `a` — through the menu.
      await rightClickFirstRow();
      await until(() => js(`[...document.querySelectorAll('.rw-menu button')].some((b) => /Rename/.test(b.textContent))`), 'the Rename item', 4000);
      await clickMenuItem('/Rename/');
      await until(() => js(`!!document.querySelector('.home-recent-rename')`), 'the rename box', 4000);
      await setValue('.home-recent-rename', 'renamed-edit-a.docx');
      await pressEnterOn('.home-recent-rename');

      const renamedTo = path.join(dir, 'renamed-edit-a.docx');
      const shown = await until(
        () => js(`[...document.querySelectorAll('.rr-name')].some((n) => n.textContent === 'renamed-edit-a.docx')`),
        'the renamed row',
        4000
      ).catch(() => false);
      check(
        'launcher: Rename… on a recent file renames it on disk and the row shows the new name',
        shown === true && fs.existsSync(renamedTo) === true && fs.existsSync(a) === false,
        `shown ${shown}, new file ${fs.existsSync(renamedTo)}, old file gone ${!fs.existsSync(a)}`
      );

      // Hover — or just press — the × on the other row, and it goes, but the
      // file it named does not. The row count on its own says nothing here —
      // launcherRecent may have left the list over its 30-row window, so a
      // removal can still show 30 rows with a different one at the bottom —
      // so what matters is that this row's own title is gone.
      await js(
        `(() => { const row = [...document.querySelectorAll('.home-recent-row')].find((r) => r.title === ${JSON.stringify(b)}); row?.querySelector('.home-recent-remove')?.click(); return 1; })()`
      );
      const gone = await until(
        () => js(`![...document.querySelectorAll('.home-recent-row')].some((r) => r.title === ${JSON.stringify(b)})`),
        'the × to remove the row',
        4000
      ).catch(() => false);
      check('launcher: the × on a recent row removes it from the list and leaves the file on disk', gone === true && fs.existsSync(b) === true, `gone ${gone}, file still there ${fs.existsSync(b)}`);

      // Renaming to a name already taken in the folder is refused, and says so.
      await rightClickFirstRow();
      await until(() => js(`[...document.querySelectorAll('.rw-menu button')].some((b) => /Rename/.test(b.textContent))`), 'the Rename item again', 4000);
      await clickMenuItem('/Rename/');
      await until(() => js(`!!document.querySelector('.home-recent-rename')`), 'the rename box again', 4000);
      await setValue('.home-recent-rename', 'recent-edit-taken.docx');
      await pressEnterOn('.home-recent-rename');
      await until(() => js(`!!document.querySelector('.rw-toast.bad')`), 'the refusal toast', 4000).catch(() => {});
      const refusal = await js(`document.querySelector('.rw-toast.bad')?.textContent || ''`);
      check(
        'launcher: renaming to a name already taken is refused, and the launcher says so',
        /already exists/i.test(refusal) && fs.existsSync(renamedTo) === true,
        `refusal "${refusal}", file still there ${fs.existsSync(renamedTo)}`
      );
      // The toast the refusal put up is a genuine `.rw-toast.bad` — let it
      // clear on its own before the check below treats any left standing as
      // something gone wrong, rather than the refusal working as meant.
      await until(() => js(`!document.querySelector('.rw-toast.bad')`), 'the toast to clear', 5000).catch(() => {});

      const complaints = await errorsIn(win);
      check('launcher: the recent rename and remove checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('launcher: the recent rename and remove checks ran', false, err.message);
    }
  };

  /* ── Frozen panes, pinned ────────────────────────────────────────────── */
  //
  // Freeze at B2 — the top row and the first column — scroll away, and the
  // heading row is still under the column headings, the first column still
  // beside the row headings, the corner cell at both, their headings with
  // them; a press on a pinned cell selects it.
  /* ── Zoom scales the page, not the window ─────────────────────────────── */
  //
  // The slider in Rutba Word and Worksheets used to zoom the whole window
  // — the ribbon, the status bar and the slider itself grew under the
  // pointer, which made it unusable (owner, 2026-09-24). Now the page (or
  // the grid) alone is scaled with CSS zoom: the ribbon keeps its height,
  // the window's own zoom factor stays 1, Word's pages still fall where they
  // did, and a press on a Worksheets cell at its zoomed place still selects
  // that cell. Run alone with RUTBA_VERIFY_ONLY=zoom.
  const zoomStaysOnThePage = async () => {
    try {
      const win = await open('word', files.docx);
      const js = (code) => win.webContents.executeJavaScript(code);
      await until(() => js(`document.querySelectorAll('.wd-page .wd-block').length > 2`), 'the document', 8000);
      const setSlider = (v) => js(`(() => {
        const el = document.querySelector('.rw-zoom input[type="range"]');
        if (!el) return 'no slider';
        const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
        setter.call(el, ${JSON.stringify(String(v))});
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        return 'set';
      })()`);
      const shape = () => js(`(() => {
        const page = document.querySelector('.wd-page');
        const ribbon = document.querySelector('.rw-ribbon');
        const slider = document.querySelector('.rw-zoom');
        return {
          pageRect: page.getBoundingClientRect().width, pageOwn: page.offsetWidth,
          ribbon: ribbon ? ribbon.getBoundingClientRect().height : 0,
          slider: slider ? slider.getBoundingClientRect().width : 0,
          sheets: document.querySelectorAll('.wd-sheet').length,
          pct: document.querySelector('.rw-zoom-pct')?.textContent,
        };
      })()`);
      const before = await shape();
      const set = await setSlider(1.5);
      await until(async () => (await shape()).pct === '150%', 'the slider to say 150%', 4000).catch(() => {});
      await wait(600);
      const after = await shape();
      const factor = win.webContents.getZoomFactor();
      check('word: the zoom slider scales the page alone — the page draws half as wide again, the ribbon and the slider keep their size, the window stays at 100%',
        set === 'set' && after.pct === '150%' && Math.abs(after.pageRect / after.pageOwn - 1.5) < 0.02 && Math.abs(after.ribbon - before.ribbon) < 1 && Math.abs(after.slider - before.slider) < 1 && Math.abs(factor - 1) < 0.001,
        `${set}; ${after.pct}; page ${after.pageRect.toFixed(0)}/${after.pageOwn} (was ${before.pageRect.toFixed(0)}/${before.pageOwn}); ribbon ${before.ribbon} → ${after.ribbon}; slider ${before.slider} → ${after.slider}; window factor ${factor}`);
      check('word: the pages fall where they did at 100%', after.sheets === before.sheets && after.sheets > 0, `${before.sheets} → ${after.sheets} sheets`);
      await js(`(() => { document.querySelector('.rw-zoom-pct')?.click(); return 1; })()`);
      await until(async () => (await shape()).pct === '100%', 'the slider back to 100%', 4000).catch(() => {});
      const reset = await shape();
      check('word: the percentage puts the page back to 100%', reset.pct === '100%' && Math.abs(reset.pageRect / reset.pageOwn - 1) < 0.02, `${reset.pct}; page ${reset.pageRect.toFixed(0)}/${reset.pageOwn}`);
      const complaints = await errorsIn(win);
      check('word: the zoom checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');

      // Worksheets: the grid scaled, and a press on a cell at its zoomed place.
      const sheet = await open('sheets', files.xlsx);
      const sjs = (code) => sheet.webContents.executeJavaScript(code);
      await until(() => sjs(`document.querySelectorAll('.sh-cell').length > 4`), 'the grid', 8000);
      const gridBefore = await sjs(`(() => { const g = document.querySelector('.sh-grid'); const r = document.querySelector('.rw-ribbon'); return { grid: g.getBoundingClientRect().width / g.offsetWidth, ribbon: r ? r.getBoundingClientRect().height : 0 }; })()`);
      const sset = await sjs(`(() => {
        const el = document.querySelector('.rw-zoom input[type="range"]');
        if (!el) return 'no slider';
        const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
        setter.call(el, '1.5');
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        return 'set';
      })()`);
      await until(() => sjs(`document.querySelector('.rw-zoom-pct')?.textContent === '150%'`), 'the grid slider to say 150%', 4000).catch(() => {});
      await wait(600);
      const gridAfter = await sjs(`(() => { const g = document.querySelector('.sh-grid'); const r = document.querySelector('.rw-ribbon'); return { grid: g.getBoundingClientRect().width / g.offsetWidth, ribbon: r ? r.getBoundingClientRect().height : 0, cells: document.querySelectorAll('.sh-cell').length }; })()`);
      const sfactor = sheet.webContents.getZoomFactor();
      check('sheets: the zoom slider scales the grid alone, the ribbon keeping its height and the window at 100%',
        sset === 'set' && Math.abs(gridAfter.grid - 1.5) < 0.02 && Math.abs(gridAfter.ribbon - gridBefore.ribbon) < 1 && Math.abs(sfactor - 1) < 0.001 && gridAfter.cells > 4,
        `${sset}; grid ${gridBefore.grid.toFixed(2)} → ${gridAfter.grid.toFixed(2)}; ribbon ${gridBefore.ribbon} → ${gridAfter.ribbon}; window factor ${sfactor}; ${gridAfter.cells} cells`);
      // A press in the middle of B2's zoomed rectangle, on the cells layer, selects B2.
      const pressed = await sjs(`(() => {
        const c = document.querySelector('.sh-cell[data-ref="B2"]');
        const layer = document.querySelector('.sh-cells');
        if (!c || !layer) return 'no B2';
        const r = c.getBoundingClientRect();
        layer.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0, clientX: r.left + r.width / 2, clientY: r.top + r.height / 2 }));
        return 'pressed';
      })()`);
      const activeRef = () => { const m = doc.model({ id: sessionFor('sheet').id }); return m?.active?.ref || m?.cells?.find((c) => c.active)?.ref || null; };
      const landed = await until(() => activeRef() === 'B2', 'B2 selected', 4000).catch(() => false);
      check('sheets: a press on a cell at its zoomed place selects that cell', pressed === 'pressed' && landed === true, `${pressed}; active ${activeRef()}`);
      const sComplaints = await errorsIn(sheet);
      check('sheets: the zoom checks report nothing', sComplaints.length === 0, sComplaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('zoom: the checks ran', false, err.message);
    }
  };

  const sheetFreeze = async () => {
    try {
      const rows = [];
      for (let r = 0; r < 80; r++) {
        const row = [];
        for (let c = 0; c < 30; c++) row.push(r === 0 ? `Head ${c + 1}` : c === 0 ? `Row ${r + 1}` : r * 100 + c);
        rows.push(row);
      }
      const file = path.join(dir, 'freeze.xlsx');
      fs.writeFileSync(file, buildXlsx({ sheets: [{ name: 'Wide', rows }] }));
      const win = await open('sheets', file);
      const js = (code) => win.webContents.executeJavaScript(code);
      await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="B2"]'))`), 'the grid', 8000);
      // As a person would: select B2, View, Freeze Panes, at the selection.
      await js(`(() => { document.querySelector('.sh-cell[data-ref="B2"]').dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 })); return 1; })()`);
      await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="B2"].active'))`), 'B2 active', 4000);
      await js(`(() => { [...document.querySelectorAll('button')].find((b) => b.textContent.trim() === 'View')?.click(); return 1; })()`);
      await until(() => js(`Boolean([...document.querySelectorAll('button')].find((b) => /Freeze Panes/.test(b.textContent)))`), 'the View tab', 4000);
      await js(`(() => { [...document.querySelectorAll('button')].find((b) => /Freeze Panes/.test(b.textContent)).click(); return 1; })()`);
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => /at the selection/.test(b.textContent)))`), 'the freeze menu', 4000);
      await js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((b) => /at the selection/.test(b.textContent)).click(); return 1; })()`);
      await until(() => js(`Boolean(document.querySelector('.sh-pin-rows .sh-cell[data-ref="B1"]')) && Boolean(document.querySelector('.sh-pin-cols .sh-cell[data-ref="A2"]'))`), 'the pinned layers', 6000);

      await js(`(() => { const g = document.querySelector('.sh-grid'); g.scrollTop = 500; g.scrollLeft = 400; return 1; })()`);
      await until(() => js(`Boolean(document.querySelector('.sh-pin-cols .sh-cell[data-ref="A30"]')) && document.querySelector('.sh-grid').scrollTop >= 400`), 'the frame after the scroll', 6000);
      await wait(300);
      const seen = await js(`(() => {
        const rect = (sel) => { const n = document.querySelector(sel); if (!n) return null; const b = n.getBoundingClientRect(); return { top: Math.round(b.top), left: Math.round(b.left), bottom: Math.round(b.bottom), right: Math.round(b.right) }; };
        const heads = rect('.sh-colheads'), rowheads = rect('.sh-rowheads'), grid = rect('.sh-grid');
        const g = document.querySelector('.sh-grid');
        // The frame carries only the columns in view, so the pinned row's cell
        // to read is whichever of them is on screen past the frozen column.
        const box = (n) => { const b = n.getBoundingClientRect(); return { top: Math.round(b.top), left: Math.round(b.left), bottom: Math.round(b.bottom), right: Math.round(b.right) }; };
        const cornerRight = rect('.sh-pin-corner')?.right ?? rowheads.right;
        const rowCell = [...document.querySelectorAll('.sh-pin-rows > .sh-cell')].find((n) => box(n).left >= cornerRight && box(n).right <= grid.right);
        const colCell = [...document.querySelectorAll('.sh-pin-cols .sh-cell')].find((n) => box(n).top > heads.bottom + 10 && box(n).bottom <= grid.bottom);
        return { heads, rowheads, grid, scroll: { top: g.scrollTop, left: g.scrollLeft },
          rowRef: rowCell?.dataset.ref || null, b1: rowCell ? box(rowCell) : null,
          colRef: colCell?.dataset.ref || null, a30: colCell ? box(colCell) : null,
          a1: rect('.sh-pin-corner .sh-cell[data-ref="A1"]'),
          head1: rect('.sh-pin-rowheads .sh-head'), headA: rect('.sh-pin-colheads .sh-head'),
          under: (() => { const b = rowCell?.getBoundingClientRect(); if (!b) return null; const hit = document.elementFromPoint(b.left + 8, b.top + 8); return hit?.dataset?.ref || hit?.className || null; })(),
          // What is on top where the headings and the corner cell are: a heading, and the corner cell.
          overHeads: (() => { const hit = document.elementFromPoint(grid.right - 120, heads.top + 8); return hit?.closest?.('.sh-head') ? 'sh-head' : hit?.className || null; })(),
          overCorner: (() => { const b = document.querySelector('.sh-pin-corner .sh-cell[data-ref="A1"]')?.getBoundingClientRect(); if (!b) return null; const hit = document.elementFromPoint(b.left + 8, b.top + 8); return hit?.dataset?.ref || hit?.className || null; })(),
          headings: document.querySelectorAll('.sh-colheads .sh-head').length };
      })()`);
      const near = (a, b) => a != null && b != null && Math.abs(a - b) <= 1;
      const rowPinned = seen.b1 && near(seen.b1.top, seen.heads.bottom) && seen.b1.left > seen.rowheads.right && seen.b1.right <= seen.grid.right;
      const colPinned = seen.a30 && near(seen.a30.left, seen.rowheads.right) && seen.a30.top > seen.heads.bottom + 10 && seen.a30.bottom <= seen.grid.bottom;
      const cornerPinned = seen.a1 && near(seen.a1.top, seen.heads.bottom) && near(seen.a1.left, seen.rowheads.right);
      const headsPinned = seen.head1 && near(seen.head1.top, seen.heads.bottom) && seen.headA && near(seen.headA.left, seen.rowheads.right);
      check('sheets: frozen at B2 and scrolled away, the top row sits under the column headings, the first column beside the row headings, the corner at both, and their headings with them',
        seen.scroll.top >= 400 && seen.scroll.left >= 300 && /^[A-Z]+1$/.test(seen.rowRef || '') && /^A\d+$/.test(seen.colRef || '') && rowPinned && colPinned && cornerPinned && headsPinned
          && seen.under === seen.rowRef && seen.overCorner === 'A1' && /sh-head/.test(seen.overHeads || '') && seen.headings > 5,
        JSON.stringify(seen));
      await wait(500);
      if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'sheet-freeze.png'), (await win.webContents.capturePage()).toPNG());

      // A press on a pinned cell selects it, where it is drawn.
      const target = seen.rowRef || 'B1';
      await js(`(() => { document.querySelector('.sh-pin-rows .sh-cell[data-ref="${target}"]')?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 })); return 1; })()`);
      const picked = await until(() => js(`Boolean(document.querySelector('.sh-pin-rows .sh-cell[data-ref="${target}"].active')) && !document.querySelector('.sh-cell[data-ref="B2"].active')`), `${target} active`, 4000).catch(() => false);
      const stillScrolled = await js(`document.querySelector('.sh-grid').scrollTop >= 400`);
      check('sheets: a press on a pinned cell selects it, and the grid stays where it was scrolled', picked === true && stillScrolled === true, `picked ${picked}, scrolled ${stillScrolled}`);
      const complaints = await errorsIn(win);
      check('sheets: frozen panes report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('sheets: the frozen panes check ran', false, err.message);
    }
  };

  /**
   * Formulas → Error Checking: type two broken formulas into the sample
   * workbook, open the pane from the ribbon, and walk it — click a row,
   * step with Next, watch a fix drop a row off live, then close it.
   */
  const sheetErrors = async () => {
    try {
      const win = await open('sheets', files.xlsx);
      const js = (code) => win.webContents.executeJavaScript(code);
      const model = () => js(`(async () => {
        const all = await window.rutbaOffice.doc.sessions({});
        const mine = all.filter((s) => s.kind === 'sheet').pop();
        return window.rutbaOffice.doc.model({ id: mine.id });
      })()`);
      const tabTo = (label) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(label)})?.click(), 'tab'`);
      const pushLabel = (label) => js(`(() => {
        const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(label)} && !n.disabled);
        if (!b) return 'no live button ' + ${JSON.stringify(label)};
        b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked';
      })()`);
      // Move the selection by reference through the engine, the way Go To does.
      const act_goto = async (ref) => {
        await js(`(async () => { const all = await window.rutbaOffice.doc.sessions({}); const mine = all.filter((s) => s.kind === 'sheet').pop(); const col = ${JSON.stringify(ref)}.charCodeAt(0) - 65; const row = Number(${JSON.stringify(ref)}.slice(1)) - 1; await window.rutbaOffice.doc.apply({ id: mine.id, ops: [{ op: 'select', row, col }] }); return 'moved'; })()`);
        await wait(150);
      };
      const typeInto = async (ref, text) => {
        await act_goto(ref);
        await js(`document.querySelector('.sh')?.focus(), 'ok'`);
        await typeText(win.webContents, text);
        await press(win.webContents, 'Return', { char: true });
      };

      await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="A1"]'))`), 'the grid', 8000);

      // D8 divides by zero; D9 calls a function the engine does not know —
      // ZZZ1 alone parses as a (blank) cell reference rather than a broken
      // name, so an unknown function is what actually forces #NAME?.
      await typeInto('D8', '=1/0');
      await until(async () => (await model()).cells?.some((c) => c.ref === 'D8' && c.text === '#DIV/0!'), 'D8 to calculate to #DIV/0!', 4000);
      await typeInto('D9', '=NOSUCHFN(1)');
      await until(async () => (await model()).cells?.some((c) => c.ref === 'D9' && /^#/.test(c.text || '')), 'D9 to calculate to an error', 4000);

      await tabTo('Formulas');
      await wait(120);
      await pushLabel('Error Checking');
      await until(() => js(`document.querySelectorAll('.sh-error').length >= 2`), 'the error rows', 4000);

      const rows = await js(`[...document.querySelectorAll('.sh-error')].map((r) => ({
        ref: r.dataset.ref,
        value: r.querySelector('.sh-error-value')?.textContent || '',
        formula: r.querySelector('.sh-error-formula')?.textContent || '',
        reason: r.querySelector('.sh-error-reason')?.textContent || '',
      }))`);
      const d8 = rows.find((r) => r.ref === 'D8');
      const d9 = rows.find((r) => r.ref === 'D9');
      check('sheets: Error Checking lists D8 as #DIV/0!, with its formula and reason', Boolean(d8) && d8.value === '#DIV/0!' && d8.formula === '=1/0' && /divides by zero/i.test(d8.reason), JSON.stringify(d8));
      check('sheets: Error Checking lists D9 as its own error, with a formula and a reason', Boolean(d9) && /^#/.test(d9.value) && d9.formula === '=NOSUCHFN(1)' && d9.reason.length > 0, JSON.stringify(d9));

      // A click on a row selects the cell.
      await js(`document.querySelector('.sh-error[data-ref="D8"]')?.dispatchEvent(new MouseEvent('click', { bubbles: true }))`);
      await until(async () => (await model()).selection?.active?.ref === 'D8', 'D8 to become active', 4000);
      const afterClick = (await model()).selection?.active?.ref;
      check('sheets: clicking an error row selects that cell', afterClick === 'D8', `active ${afterClick}`);

      // Next steps from D8 to D9.
      await js(`document.querySelector('.sh-error-next')?.click(), 'next'`);
      await until(async () => (await model()).selection?.active?.ref === 'D9', 'D9 to become active', 4000);
      const afterNext = (await model()).selection?.active?.ref;
      check('sheets: Next walks from D8 to D9', afterNext === 'D9', `active ${afterNext}`);

      // Fixing D8 drops its row the moment the next frame lands.
      await typeInto('D8', '=1');
      await until(() => js(`!document.querySelector('.sh-error[data-ref="D8"]')`), 'D8 to drop off the list once fixed', 4000);
      const stillD9 = await js(`Boolean(document.querySelector('.sh-error[data-ref="D9"]'))`);
      check('sheets: fixing a cell drops it from the list live', stillD9 === true, 'D9 should still be listed once D8 is fixed');

      // Close turns the checking off: the pane goes and the button lifts.
      await js(`[...document.querySelectorAll('.sh-errors .rw-btn')].find((b) => /Close/.test(b.title || b.dataset.tip || ''))?.click(), 'close'`);
      await until(() => js(`!document.querySelector('.sh-errors')`), 'the pane to close', 4000);
      const stillPressed = await js(`Boolean([...document.querySelectorAll('.rw-ribbon .rw-btn')].find((b) => (b.title || b.dataset.tip || '').startsWith('Error Checking') && b.getAttribute('aria-pressed') === 'true'))`);
      check('sheets: Close turns Error Checking off, the pane gone and the button not pressed', !stillPressed, `pressed ${stillPressed}`);

      const complaints = await errorsIn(win);
      check('sheets: error checking reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('sheets: the error checking check ran', false, err.message);
    }
  };

  /**
   * Formulas → Watch Window: select a formula cell already in the sample
   * workbook, open the pane from the ribbon, Add Watch, read the row it
   * shows, edit the precedent the formula reads through the ordinary
   * engine ops and watch the row's value follow it live, then take the
   * watch off again and see the row go.
   */
  const sheetWatch = async () => {
    try {
      const win = await open('sheets', files.xlsx);
      const js = (code) => win.webContents.executeJavaScript(code);
      const model = () => js(`(async () => {
        const all = await window.rutbaOffice.doc.sessions({});
        const mine = all.filter((s) => s.kind === 'sheet').pop();
        return window.rutbaOffice.doc.model({ id: mine.id });
      })()`);
      const tabTo = (label) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(label)})?.click(), 'tab'`);
      const pushLabel = (label) => js(`(() => {
        const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(label)} && !n.disabled);
        if (!b) return 'no live button ' + ${JSON.stringify(label)};
        b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked';
      })()`);
      // The watch pane's own buttons — Add Watch, Delete Watch, Close —
      // live over the grid, not in the ribbon, so they need their own root.
      const pushPane = (label) => js(`(() => {
        const b = [...document.querySelectorAll('.sh-watch .rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(label)} && !n.disabled);
        if (!b) return 'no live button ' + ${JSON.stringify(label)};
        b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked';
      })()`);
      const act_goto = async (ref) => {
        await js(`(async () => { const all = await window.rutbaOffice.doc.sessions({}); const mine = all.filter((s) => s.kind === 'sheet').pop(); const col = ${JSON.stringify(ref)}.charCodeAt(0) - 65; const row = Number(${JSON.stringify(ref)}.slice(1)) - 1; await window.rutbaOffice.doc.apply({ id: mine.id, ops: [{ op: 'select', row, col }] }); return 'moved'; })()`);
        await wait(150);
      };
      const typeInto = async (ref, text) => {
        await act_goto(ref);
        await js(`document.querySelector('.sh')?.focus(), 'ok'`);
        await typeText(win.webContents, text);
        await press(win.webContents, 'Return', { char: true });
      };
      const watchRow = (ref) => js(`(() => {
        const r = document.querySelector('.sh-watch-row[data-ref="${ref}"]');
        if (!r) return null;
        return {
          sheet: r.querySelector('.sh-watch-sheet')?.textContent || '',
          cell: r.querySelector('.sh-watch-cell')?.textContent || '',
          value: r.querySelector('.sh-watch-value')?.textContent || '',
          formula: r.querySelector('.sh-watch-formula')?.textContent || '',
        };
      })()`);

      await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="A1"]'))`), 'the grid', 8000);

      // D2 already holds =SUM(B2:C2) — 1420 + 1610 — in the sample workbook.
      await act_goto('D2');
      await tabTo('Formulas');
      await wait(120);
      await pushLabel('Watch Window');
      await until(() => js(`Boolean(document.querySelector('.sh-watch'))`), 'the watch pane', 4000);
      await pushPane('Add Watch');
      await until(() => js(`Boolean(document.querySelector('.sh-watch-row[data-ref="D2"]'))`), 'the D2 watch row', 4000);

      const row = await watchRow('D2');
      check(
        'sheets: Add Watch lists D2 with its sheet, value and formula',
        Boolean(row) && row.sheet === 'Sales' && row.cell === 'D2' && row.value === '3030' && row.formula === '=SUM(B2:C2)',
        JSON.stringify(row),
      );

      // B2 is a precedent of D2's formula: changing it must move the watch
      // live, with nothing telling the pane to refresh.
      await typeInto('B2', '5000');
      await until(async () => (await watchRow('D2'))?.value === '6610', 'the watched value to follow B2', 4000);
      const after = await watchRow('D2');
      check('sheets: the watched value follows a precedent edited elsewhere', after?.value === '6610', JSON.stringify(after));

      // The row's own × takes the watch off, live.
      await js(`document.querySelector('.sh-watch-row[data-ref="D2"] .sh-watch-x')?.click(), 'remove'`);
      await until(() => js(`!document.querySelector('.sh-watch-row[data-ref="D2"]')`), 'the D2 row to go', 4000);
      const gone = await js(`Boolean(document.querySelector('.sh-watch-row[data-ref="D2"]'))`);
      check("sheets: a row's × removes the watch", gone === false, `still there: ${gone}`);

      // Close turns the window off: the pane goes and the button lifts.
      await js(`[...document.querySelectorAll('.sh-watch .rw-btn')].find((b) => /Close/.test(b.title || b.dataset.tip || ''))?.click(), 'close'`);
      await until(() => js(`!document.querySelector('.sh-watch')`), 'the pane to close', 4000);
      const stillPressed = await js(`Boolean([...document.querySelectorAll('.rw-ribbon .rw-btn')].find((b) => (b.title || b.dataset.tip || '').startsWith('Watch Window') && b.getAttribute('aria-pressed') === 'true'))`);
      check('sheets: Close turns the Watch Window off, the pane gone and the button not pressed', !stillPressed, `pressed ${stillPressed}`);

      const complaints = await errorsIn(win);
      check('sheets: the watch window reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('sheets: the watch window check ran', false, err.message);
    }
  };

  /**
   * Insert → Sparklines: a line, then a column, from the ribbon's own
   * dialog, drawn in the grid at once, kept through a save, and taken off
   * again — the window end of the extension the engine writes.
   */
  const sheetSparklines = async () => {
    try {
      const win = await open('sheets', files.xlsx);
      const js = (code) => win.webContents.executeJavaScript(code);
      const model = () => js(`(async () => {
        const all = await window.rutbaOffice.doc.sessions({});
        const mine = all.filter((s) => s.kind === 'sheet').pop();
        return window.rutbaOffice.doc.model({ id: mine.id });
      })()`);
      // A mousedown on the first cell, a shift-click on the last — as a
      // person selects a range, and the only way that reaches the WINDOW's
      // own selection state (a raw doc.apply moves the engine but leaves
      // the React model nobody told to catch up).
      const selectRange = async (from, to) => {
        await js(`(() => { document.querySelector('.sh-cell[data-ref="${from}"]')?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 })); return 1; })()`);
        await wait(80);
        await js(`(() => { document.querySelector('.sh-cell[data-ref="${to}"]')?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0, shiftKey: true })); return 1; })()`);
        await until(async () => (await model()).selection?.ref === from + ':' + to, `the selection to become ${from}:${to}`, 4000);
      };
      // The Sparklines group's own buttons — scoped past the Charts group,
      // which the same ribbon also labels "Line" for its line CHART.
      const clickSparkline = (label) => js(`(() => {
        const group = [...document.querySelectorAll('.rw-group')].find((g) => g.querySelector('.rw-group-label')?.textContent.trim() === 'Sparklines');
        const b = group && [...group.querySelectorAll('.rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(label)});
        if (!b) return 'no button ' + ${JSON.stringify(label)};
        b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked';
      })()`);
      const setField = (selector, value) => js(`(() => {
        const el = document.querySelector(${JSON.stringify(selector)});
        if (!el) return 'no field ' + ${JSON.stringify(selector)};
        const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
        setter.call(el, ${JSON.stringify(value)});
        el.dispatchEvent(new Event('input', { bubbles: true }));
        return 'set';
      })()`);

      await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="B2"]'))`), 'the grid', 8000);
      await selectRange('B2', 'C2');

      await js(`(() => { [...document.querySelectorAll('button')].find((b) => b.textContent.trim() === 'Insert')?.click(); return 1; })()`);
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-group-label')].find((g) => g.textContent.trim() === 'Sparklines'))`), 'the Insert tab', 4000);

      const clickedLine = await clickSparkline('Line');
      await until(() => js(`Boolean(document.querySelector('.sh-sparkline-data'))`), 'the sparkline dialog', 4000);
      const prefill = await js(`({ data: document.querySelector('.sh-sparkline-data')?.value, at: document.querySelector('.sh-sparkline-at')?.value })`);
      check('sheets: Insert → Sparklines → Line opens prefilled from the selection, to the cell past it',
        clickedLine === 'clicked' && prefill.data === 'B2:C2' && prefill.at === 'D2', `${clickedLine}; ${JSON.stringify(prefill)}`);

      await js(`document.querySelector('.sh-sparkline-ok')?.click(), 'ok'`);
      await until(async () => (await model()).sparklines?.some((s) => s.at.row === 1 && s.at.col === 3 && s.type === 'line'), 'D2 to carry a line sparkline', 4000);
      const atD2 = (await model()).sparklines.find((s) => s.at.row === 1 && s.at.col === 3);
      check('sheets: the model carries a line sparkline at D2 with the row\'s two values',
        atD2?.type === 'line' && Array.isArray(atD2.values) && atD2.values.length === 2, JSON.stringify(atD2));

      await until(() => js(`Boolean(document.querySelector('.sh-spark[data-ref="D2"] polyline'))`), 'D2 drawn as a line sparkline', 4000);
      const drawnLine = await js(`Boolean(document.querySelector('.sh-spark[data-ref="D2"] polyline'))`);
      check('sheets: the cell layer draws D2 as a line sparkline (a polyline)', drawnLine === true, `drawn ${drawnLine}`);

      let wasSaved = fs.statSync(files.xlsx).mtimeMs;
      await press(win.webContents, 's', { modifiers: ['control'] });
      await until(() => fs.statSync(files.xlsx).mtimeMs !== wasSaved, 'the file to be written', 5000);

      let saved = SheetView.open(fs.readFileSync(files.xlsx));
      let groups = saved.sparklineGroups('Sales');
      const savedLine = groups.find((g) => g.sparklines.some((s) => s.at === 'D2'));
      check('sheets: the saved file carries the line sparkline as Excel writes one, Sales!B2:C2 into D2',
        Boolean(savedLine) && savedLine.type === 'line'
          && savedLine.sparklines.some((s) => s.at === 'D2' && s.data === 'Sales!B2:C2'),
        JSON.stringify(groups));

      // A column sparkline, from the same numbers, beside the first — not
      // on top of it, so the location is changed before OK.
      await selectRange('B2', 'C2');
      const clickedColumn = await clickSparkline('Column');
      await until(() => js(`Boolean(document.querySelector('.sh-sparkline-data'))`), 'the sparkline dialog again', 4000);
      await setField('.sh-sparkline-at', 'E2');
      await js(`document.querySelector('.sh-sparkline-ok')?.click(), 'ok'`);
      await until(async () => (await model()).sparklines?.some((s) => s.at.row === 1 && s.at.col === 4 && s.type === 'column'), 'E2 to carry a column sparkline', 4000);
      check('sheets: Insert → Sparklines → Column opens and applies to a chosen cell', clickedColumn === 'clicked', clickedColumn);

      await until(() => js(`Boolean(document.querySelector('.sh-spark[data-ref="E2"] rect'))`), 'E2 drawn as bars', 4000);
      const drawnBars = await js(`document.querySelectorAll('.sh-spark[data-ref="E2"] rect').length`);
      check('sheets: the cell layer draws E2 as a column sparkline (bars)', drawnBars >= 1, `${drawnBars} rects`);

      // Remove the line sparkline at D2; the column one at E2 rides through.
      await js(`(() => { document.querySelector('.sh-cell[data-ref="D2"]').dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, button: 2 })); return 1; })()`);
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => /Remove sparkline/.test(b.textContent)))`), 'the Remove sparkline item', 4000);
      await js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((b) => /Remove sparkline/.test(b.textContent)).click(); return 1; })()`);
      await until(async () => !(await model()).sparklines?.some((s) => s.at.col === 3), 'D2 to drop off the model', 4000);
      const stillE2 = (await model()).sparklines?.some((s) => s.at.col === 4);
      check('sheets: Remove sparkline takes it off the model, and leaves the other one', stillE2 === true, JSON.stringify((await model()).sparklines));
      check('sheets: the cell layer stops drawing the removed one', await js(`!document.querySelector('.sh-spark[data-ref="D2"]')`) === true, 'D2 still drawn');

      wasSaved = fs.statSync(files.xlsx).mtimeMs;
      await press(win.webContents, 's', { modifiers: ['control'] });
      await until(() => fs.statSync(files.xlsx).mtimeMs !== wasSaved, 'the second save to be written', 5000);
      saved = SheetView.open(fs.readFileSync(files.xlsx));
      groups = saved.sparklineGroups('Sales');
      const stillInFile = groups.some((g) => g.sparklines.some((s) => s.at === 'D2'));
      const columnInFile = groups.some((g) => g.type === 'column' && g.sparklines.some((s) => s.at === 'E2'));
      check('sheets: the removal is saved to the file, the survivor kept', !stillInFile && columnInFile, JSON.stringify(groups));

      const complaints = await errorsIn(win);
      check('sheets: sparklines report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('sheets: the sparklines check ran', false, err.message);
    }
  };

  /**
   * Mail: the big providers are a tile away.
   *
   * The Add account dialog offers a row of tiles above the address field —
   * Gmail, Outlook, Yahoo and the rest — each filling its own servers with no
   * network and saying what that provider wants: a browser sign-in, or an
   * app password with a link to make one.
   */
  const mailProviders = async () => {
    try {
      const win = await open('mail');
      const js = (code) => win.webContents.executeJavaScript(code);
      await until(() => js(`Boolean([...document.querySelectorAll('button')].find((b) => /Add account/.test(b.textContent)))`), 'an Add account button', 8000);
      await js(`[...document.querySelectorAll('button')].find((b) => /Add account/.test(b.textContent)).click(), 'clicked'`);
      await until(() => js(`Boolean(document.querySelector('#ml-address'))`), 'the account dialog', 5000);

      const tiles = await js(`(() => ({ count: document.querySelectorAll('.ml-provider').length, yahoo: Boolean(document.querySelector('.ml-provider[data-provider="yahoo"]')) }))()`);
      check('mail: the Add account dialog offers a tile for the big providers', tiles.count >= 8 && tiles.yahoo, JSON.stringify(tiles));

      await js(`document.querySelector('.ml-provider[data-provider="yahoo"]').click(), 'clicked'`);
      await until(() => js(`Boolean(document.querySelector('.ml-provider-change'))`), 'the tiles to collapse once Yahoo is chosen', 4000);
      await js(`[...document.querySelectorAll('.rw-btn, button')].find((b) => /^Advanced/.test(b.textContent.trim()))?.click(), 'advanced'`);
      await until(() => js(`Boolean(document.querySelector('.ml-advanced'))`), 'the advanced fields', 3000);
      const advanced = await js(`(() => {
        const a = document.querySelector('.ml-advanced');
        if (!a) return null;
        const ports = [...a.querySelectorAll('.ml-servers3 input[type="number"]')].map((i) => i.value);
        return { imap: a.querySelector('input[placeholder="imap.example.com"]')?.value, imapPort: ports[0], smtp: a.querySelector('input[placeholder="smtp.example.com"]')?.value, smtpPort: ports[1] };
      })()`);
      check('mail: choosing the Yahoo tile fills the advanced fields with no network', advanced && advanced.imap === 'imap.mail.yahoo.com' && String(advanced.imapPort) === '993' && advanced.smtp === 'smtp.mail.yahoo.com' && String(advanced.smtpPort) === '465', JSON.stringify(advanced));

      const appPw = await js(`(() => { const b = document.querySelector('.ml-app-password'); return { present: Boolean(b), title: b?.title || b?.dataset.tip || '' }; })()`);
      check('mail: the Yahoo tile offers a link to make an app password', appPw.present && /yahoo/i.test(appPw.title), JSON.stringify(appPw));

      const placeholder = await js(`document.querySelector('#ml-address')?.placeholder || ''`);
      check('mail: choosing a provider sets the address placeholder to its domain', /yahoo\.com$/.test(placeholder), placeholder);

      await js(`document.querySelector('.ml-provider-change')?.click(), 'change'`);
      await until(() => js(`Boolean(document.querySelector('.ml-provider[data-provider="google"]'))`), 'the tiles again', 3000);
      await js(`document.querySelector('.ml-provider[data-provider="google"]').click(), 'clicked'`);
      await until(() => js(`Boolean(document.querySelector('.ml-found-item'))`), 'the Google sign-in card', 4000);
      const google = await js(`(() => {
        const item = document.querySelector('.ml-found-item');
        return { signIn: /Sign in with Google/.test(item?.textContent || ''), appPassword: Boolean(document.querySelector('.ml-app-password')) };
      })()`);
      check('mail: choosing the Gmail tile offers its sign-in card and an app-password line', google.signIn && google.appPassword, JSON.stringify(google));

      const complaints = await errorsIn(win);
      check('mail: choosing a provider tile reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('mail: the provider tiles check ran', false, err.message);
    }
  };

  /**
   * Mail: a signature belongs to the account, not the person.
   *
   * Set from that account's own settings — the Folder tab, not a preference
   * shared by every account in the mailbox — it turns up at the end of a new
   * message after a blank line and the "-- " line mail readers use to spot
   * one, and above the quote rather than below it on a reply, where it sits
   * under the cursor instead of under words that were already sent.
   */
  const mailSignature = async () => {
    let win = null;
    let accountId = null;
    const SIGNATURE = 'Verify Signature Line\nSecond line of it';
    try {
      win = await open('mail');
      const js = (code) => win.webContents.executeJavaScript(code);
      await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the seeded account in the sidebar', 8000);
      accountId = await js(`(async () => (await window.rutbaOffice.mail.accounts())[0]?.id || null)()`);

      await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Folder')?.click(), 'tab'`);
      await until(() => js(`Boolean([...document.querySelectorAll('button')].find((b) => /^Signature$/.test(b.textContent.trim()) && !b.disabled))`), 'the account\'s Signature button', 5000);
      await js(`[...document.querySelectorAll('button')].find((b) => /^Signature$/.test(b.textContent.trim()) && !b.disabled)?.click(), 'clicked'`);
      await until(() => js(`Boolean(document.querySelector('.ml-signature-text'))`), 'the signature editor', 5000);
      await js(`(() => {
        const el = document.querySelector('.ml-signature-text');
        const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set;
        setter.call(el, ${JSON.stringify(SIGNATURE)});
        el.dispatchEvent(new Event('input', { bubbles: true }));
        return 'set';
      })()`);
      await js(`[...document.querySelectorAll('.rw-dialog button')].find((b) => /^Save$/.test(b.textContent.trim()))?.click(), 'saved'`);
      await until(() => js(`!document.querySelector('.ml-signature-text')`), 'the signature editor to close', 5000);

      // A new message: the signature at the end, after a blank line and "-- ".
      // The plain-text body is read rather than the rich editor's rendering,
      // which is free to collapse the trailing space "-- " depends on.
      await js(`document.querySelector('.ml-compose-cta button')?.click(), 'clicked'`);
      await until(() => js(`Boolean(document.querySelector('.ml-toolbar button[data-tip="Plain text"]'))`), 'the composer', 5000);
      await js(`document.querySelector('.ml-toolbar button[data-tip="Plain text"]')?.click(), 'plain'`);
      await until(() => js(`Boolean(document.querySelector('.ml-compose-body'))`), 'the plain-text body', 4000);
      // The signature is filled in a beat after the body appears; read it once it is there.
      await until(() => js(`(document.querySelector('.ml-compose-body')?.value || '').includes('-- ')`), 'the signature in the body', 5000).catch(() => {});
      const fresh = await js(`document.querySelector('.ml-compose-body')?.value || ''`);
      check(
        'mail: a new message carries the account\'s signature, after a blank line and the "-- " line',
        fresh === `\n\n-- \n${SIGNATURE}`,
        JSON.stringify(fresh)
      );
      await js(`[...document.querySelectorAll('button')].find((b) => /^Discard$/.test(b.textContent.trim()))?.click(), 'closed'`);
      await until(() => js(`!document.querySelector('.ml-compose-body, .ml-rich')`), 'the composer to close', 4000).catch(() => {});

      // A reply: the signature above the quote, below the cursor.
      await until(() => js(`document.querySelectorAll('.ml-row').length > 0`), 'the seeded messages', 8000);
      await js(`document.querySelector('.ml-row')?.click(), 'opened'`);
      await until(() => js(`Boolean(document.querySelector('.rw-btn[data-tip="Reply"]'))`), 'the reading pane', 8000);
      await js(`document.querySelector('.rw-btn[data-tip="Reply"]')?.click(), 'reply'`);
      await until(() => js(`Boolean(document.querySelector('.ml-toolbar button[data-tip="Plain text"]'))`), 'the reply composer', 5000);
      await js(`document.querySelector('.ml-toolbar button[data-tip="Plain text"]')?.click(), 'plain'`);
      await until(() => js(`Boolean(document.querySelector('.ml-compose-body'))`), 'the plain-text reply body', 4000);
      await until(() => js(`(document.querySelector('.ml-compose-body')?.value || '').includes('-- ')`), 'the signature in the reply', 5000).catch(() => {});
      const replied = await js(`document.querySelector('.ml-compose-body')?.value || ''`);
      const sigAt = replied.indexOf(SIGNATURE);
      const quoteAt = replied.search(/wrote:/);
      check(
        'mail: a reply carries the signature above the quote, below the cursor',
        replied.startsWith(`\n\n-- \n${SIGNATURE}`) && sigAt >= 0 && quoteAt > sigAt,
        JSON.stringify({ replied: replied.slice(0, 160), sigAt, quoteAt })
      );

      const complaints = await errorsIn(win);
      check('mail: the signature checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('mail: the signature check ran', false, err.message);
    } finally {
      // The account is a fixture every other mail check shares; leaving a
      // signature on it is not this check's to decide.
      if (win && accountId) {
        await win.webContents
          .executeJavaScript(`window.rutbaOffice.mail.updateAccount({ id: ${JSON.stringify(accountId)}, patch: { signature: '' } })`)
          .catch(() => {});
      }
    }
  };

  /**
   * Mail: Send later puts a message in the Outbox instead of the wire, the
   * Outbox is a folder in the list rather than only a ribbon button, and
   * from there it can be sent now, taken back into Compose, or cancelled.
   */
  const mailSendLater = async () => {
    let win = null;
    let accountId = null;
    const SUBJECT = `Verify Send Later ${Date.now()}`;
    try {
      win = await open('mail');
      const js = (code) => win.webContents.executeJavaScript(code);
      const setValue = (selector, value) => js(`(() => { const el = document.querySelector(${JSON.stringify(selector)}); if (!el) return false; const proto = el.tagName === 'TEXTAREA' ? window.HTMLTextAreaElement.prototype : window.HTMLInputElement.prototype; const setter = Object.getOwnPropertyDescriptor(proto, 'value').set; setter.call(el, ${JSON.stringify(value)}); el.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
      const setField = (label, value) => js(`(() => {
        const el = [...document.querySelectorAll('.rw-field')].find((f) => f.querySelector('label')?.textContent.trim() === ${JSON.stringify(label)})?.querySelector('input, textarea');
        if (!el) return false;
        const proto = el.tagName === 'TEXTAREA' ? window.HTMLTextAreaElement.prototype : window.HTMLInputElement.prototype;
        Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, ${JSON.stringify(value)});
        el.dispatchEvent(new Event('input', { bubbles: true }));
        return true;
      })()`);

      await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the seeded account in the sidebar', 8000);
      accountId = await js(`(async () => (await window.rutbaOffice.mail.accounts())[0]?.id || null)()`);

      await js(`document.querySelector('.ml-compose-cta button')?.click(), 'clicked'`);
      await until(() => js(`Boolean(document.querySelector('.ml-to-row'))`), 'the composer', 5000);
      await setField('To', 'later@example.com');
      await setField('Subject', SUBJECT);

      await js(`document.querySelector('.rw-dialog button[data-tip="Send later"]')?.click(), 'clicked'`);
      await until(() => js(`Boolean(document.querySelector('.ml-schedule-custom'))`), 'the schedule panel and its custom-time field', 4000);

      // A custom time, not one of the presets — the one nothing else proves.
      // `datetime-local` reads back as local time (packages/mailbox/src/
      // schedule.js's own `parseCustomSchedule`), so the value fed to it has
      // to be built from local getters — `toISOString()` is UTC, and on a
      // machine whose zone is not UTC that reads as a different moment than
      // the one this check means, in one direction or the other.
      const inOneHour = new Date(Date.now() + 3600_000);
      const pad2 = (n) => String(n).padStart(2, '0');
      const customAt = `${inOneHour.getFullYear()}-${pad2(inOneHour.getMonth() + 1)}-${pad2(inOneHour.getDate())}T${pad2(inOneHour.getHours())}:${pad2(inOneHour.getMinutes())}`;
      await setValue('.ml-schedule-custom', customAt);
      await js(`document.querySelector('.rw-dialog button[data-tip="Schedule"]')?.click(), 'clicked'`);
      await until(() => js(`!document.querySelector('.ml-to-row')`), 'the composer to close', 5000);

      const queued = await js(`(async () => (await window.rutbaOffice.mail.outbox()).find((o) => o.draft?.subject === ${JSON.stringify(SUBJECT)}) || null)()`);
      check(
        'mail: Send later queues the message with the chosen time instead of sending it',
        Boolean(queued) && new Date(queued.at).getTime() > Date.now(),
        JSON.stringify(queued)
      );

      // The Outbox is listed in the folder list, not reached only from the ribbon.
      await js(`[...document.querySelectorAll('.rw-item')].find((b) => b.querySelector('.label')?.textContent.trim() === 'Outbox')?.click(), 'clicked'`);
      await until(() => js(`Boolean(document.querySelector('.rw-dialog[aria-label="Outbox"]'))`), 'the Outbox dialog', 5000);
      const listed = await js(`(() => {
        const item = [...document.querySelectorAll('.ml-found-item')].find((el) => el.querySelector('.who')?.textContent.trim() === ${JSON.stringify(SUBJECT)});
        return item ? item.querySelector('.what')?.textContent || '' : null;
      })()`);
      check('mail: the scheduled message reads in the Outbox with its send time', typeof listed === 'string' && /sends/.test(listed), JSON.stringify(listed));

      // Send now, through the fake transport every check run sends through.
      const clicked = await js(`(() => {
        const item = [...document.querySelectorAll('.ml-found-item')].find((el) => el.querySelector('.who')?.textContent.trim() === ${JSON.stringify(SUBJECT)});
        const btn = [...(item?.querySelectorAll('.ml-outbox-actions button') || [])].find((b) => b.textContent.trim() === 'Send now');
        btn?.click();
        return Boolean(btn);
      })()`);
      check('mail: the Outbox offers Send now on the item', clicked === true, String(clicked));
      await until(
        () => js(`![...document.querySelectorAll('.ml-found-item')].some((el) => el.querySelector('.who')?.textContent.trim() === ${JSON.stringify(SUBJECT)})`),
        'the message to leave the Outbox',
        6000
      );

      const stillQueued = await js(`(async () => (await window.rutbaOffice.mail.outbox()).some((o) => o.draft?.subject === ${JSON.stringify(SUBJECT)}))()`);
      const inSent = await js(`(async () => (await window.rutbaOffice.mail.messages({ accountId: ${JSON.stringify(accountId)}, folder: 'Sent', limit: 200 })).rows.some((r) => r.subject === ${JSON.stringify(SUBJECT)}))()`);
      check('mail: Send now takes it off the Outbox and it lands in Sent', stillQueued === false && inSent === true, JSON.stringify({ stillQueued, inSent }));

      await js(`[...document.querySelectorAll('.rw-dialog-foot button')].find((b) => b.textContent.trim() === 'Close')?.click(), 'closed'`);

      const complaints = await errorsIn(win);
      check('mail: the send-later checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('mail: the send-later check ran', false, err.message);
    }
  };

  /**
   * Mail: automatic replies, on from the account's own settings, answering
   * a fresh arrival once and staying silent about a second message from the
   * same sender — the RFC 3834 checks that keep it a good citizen are proved
   * in tests/mail-ooo.test.js; this proves the account setting, the banner,
   * and the reply actually going out through the fake transport.
   */
  const mailOOO = async () => {
    let win = null;
    let accountId = null;
    const SENDER = `ooo-check-${Date.now()}@example.org`;
    const raw = (msgId, subject) =>
      [
        `From: OOO Check <${SENDER}>`,
        'To: You <you@example.com>',
        `Subject: ${subject}`,
        `Date: ${new Date().toUTCString()}`,
        `Message-ID: <${msgId}>`,
        'MIME-Version: 1.0',
        'Content-Type: text/plain; charset=utf-8',
        '',
        'Checking automatic replies.',
        '',
      ].join('\r\n');
    try {
      win = await open('mail');
      const js = (code) => win.webContents.executeJavaScript(code);
      await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the seeded account in the sidebar', 8000);
      accountId = await js(`(async () => (await window.rutbaOffice.mail.accounts())[0]?.id || null)()`);

      await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Folder')?.click(), 'tab'`);
      await until(() => js(`Boolean([...document.querySelectorAll('button')].find((b) => /Out of office/.test(b.textContent) && !b.disabled))`), "the account's Out of office button", 5000);
      await js(`[...document.querySelectorAll('button')].find((b) => /Out of office/.test(b.textContent) && !b.disabled)?.click(), 'clicked'`);
      await until(() => js(`Boolean(document.querySelector('.rw-dialog[aria-label="Automatic replies"]'))`), 'the Automatic replies dialog', 5000);

      const scope = `document.querySelector('.rw-dialog[aria-label="Automatic replies"]')`;
      await js(`${scope}.querySelector('.ml-ooo-enabled input')?.click(), 'on'`);
      await js(`(() => { const el = ${scope}.querySelector('.ml-ooo-message'); const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set; setter.call(el, 'Away until further notice.'); el.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
      await js(`[...${scope}.querySelectorAll('.rw-dialog-foot button')].find((b) => b.textContent.trim() === 'Save')?.click(), 'saved'`);
      await until(() => js(`!document.querySelector('.rw-dialog[aria-label="Automatic replies"]')`), 'the dialog to close', 5000);

      // On for this account is a banner in the window, not a setting nobody can see.
      await until(() => js(`Boolean(document.querySelector('.ml-ooo-banner'))`), 'the automatic-replies banner', 5000);
      const banner = await js(`document.querySelector('.ml-ooo-banner')?.textContent || ''`);
      check('mail: turning automatic replies on shows a banner naming the account', /Automatic replies are on/.test(banner), banner);

      // A message arrives through the fake inbound path — deliverTest refuses
      // outside a check run, the same way the fake SMTP transport does.
      await js(`(async () => window.rutbaOffice.mail.deliverTest({ accountId: ${JSON.stringify(accountId)}, folder: 'Inbox', raw: ${JSON.stringify(raw('ooo-check-1@example.org', 'First message'))} }))()`);

      const readSentToSender = () =>
        js(`(async () => (await window.rutbaOffice.mail.messages({ accountId: ${JSON.stringify(accountId)}, folder: 'Sent', limit: 300 })).rows.filter((r) => (r.to || []).some((t) => t.address === ${JSON.stringify(SENDER)})))()`);

      await until(async () => (await readSentToSender()).length === 1, 'exactly one automatic reply', 6000);
      const rows = await readSentToSender();
      const full = await js(`window.rutbaOffice.mail.message({ accountId: ${JSON.stringify(accountId)}, folder: 'Sent', id: ${JSON.stringify(rows[0].id)} })`);
      const auto = (full.headers || []).find((h) => h.key === 'auto-submitted');
      check(
        'mail: the reply carries Auto-Submitted, In-Reply-To and a subject naming the original',
        auto?.value === 'auto-replied' && full.inReplyTo === 'ooo-check-1@example.org' && full.subject === 'Automatic reply: First message',
        JSON.stringify({ auto, inReplyTo: full.inReplyTo, subject: full.subject })
      );

      // A second message from the same sender: no second reply.
      await js(`(async () => window.rutbaOffice.mail.deliverTest({ accountId: ${JSON.stringify(accountId)}, folder: 'Inbox', raw: ${JSON.stringify(raw('ooo-check-2@example.org', 'Second message'))} }))()`);
      await wait(500);
      const afterSecond = await readSentToSender();
      check('mail: a second message from the same sender gets no second reply', afterSecond.length === 1, String(afterSecond.length));

      await js(`document.querySelector('.ml-ooo-banner button')?.click(), 'off'`);
      await until(() => js(`!document.querySelector('.ml-ooo-banner')`), 'the banner to go once it is turned off', 5000);

      const complaints = await errorsIn(win);
      check('mail: the automatic-reply checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    } catch (err) {
      check('mail: the automatic-reply check ran', false, err.message);
    } finally {
      // The account is a fixture every other mail check shares.
      if (win && accountId) {
        await win.webContents
          .executeJavaScript(`window.rutbaOffice.mail.updateAccount({ id: ${JSON.stringify(accountId)}, patch: { autoReply: { enabled: false } } })`)
          .catch(() => {});
      }
    }
  };

  // RUTBA_VERIFY_ONLY=pages,grips,panes,float,polish,shapes,fill,pics,ruler,columns,update,viewer,slideshow,links,home,recent,freeze,errors,watch,sparklines,fit,sections,hidden,background,effects,bookmarks,xref,captions,providers,signature,deckfind,sendlater,ooo,arrange,deckfx,toc,track,outline: those blocks alone, for working on them.
  const only = (process.env.RUTBA_VERIFY_ONLY || '').split(',').map((s) => s.trim()).filter(Boolean);
  if (only.length) {
    if (only.includes('pages')) await wordPages();
    if (only.includes('grips')) await sheetGrips();
    if (only.includes('panes')) await slidePanes();
    if (only.includes('float')) await wordFloat();
    if (only.includes('polish')) await polish();
    if (only.includes('shapes')) await slideShapes();
    if (only.includes('bullets')) await slideParagraphs();
    if (only.includes('clip')) await slideClipboard();
    if (only.includes('footer')) await slideFooter();
    if (only.includes('find')) await slideFind();
    if (only.includes('deckfind')) await slideFind();
    if (only.includes('sections')) await slideSections();
    if (only.includes('hidden')) await slideHidden();
    if (only.includes('background')) await slideBackground();
    if (only.includes('table')) await slideTable();
    if (only.includes('chart')) await slideChart();
    if (only.includes('fill')) await sheetFill();
    if (only.includes('pics')) await wordPictures();
    if (only.includes('look')) await wordLook();
    if (only.includes('dropcap')) await wordDropCap();
    if (only.includes('bookmarks')) await wordBookmarks();
    if (only.includes('xref')) await wordCrossRef();
    if (only.includes('captions')) await wordCaptions();
    if (only.includes('toc')) await wordToc();
    if (only.includes('track')) await wordTrack();
    if (only.includes('effects')) await wordEffects();
    if (only.includes('ruler')) await wordRuler();
    if (only.includes('columns')) await wordColumns();
    if (only.includes('update')) await updatePrompt();
    if (only.includes('home')) await launcherRecent();
    if (only.includes('recent')) await homeRecentEdit();
    if (only.includes('freeze')) await sheetFreeze();
    if (only.includes('errors')) await sheetErrors();
    if (only.includes('watch')) await sheetWatch();
    if (only.includes('sparklines')) await sheetSparklines();
    if (only.includes('zoom')) await zoomStaysOnThePage();
    if (only.includes('fit')) await wordPictureFits();
    if (only.includes('cards')) await wordCards();
    if (only.includes('sheetpic')) await sheetPicture();
    if (only.includes('viewer')) await viewer();
    if (only.includes('slideshow')) await slideshow();
    if (only.includes('links')) await sheetLinks();
    if (only.includes('outline')) await sheetOutline();
    if (only.includes('arrange')) await verifyDeckArrange({ open, check, until, wait, press, errorsIn, doc, sessionFor }, { file: files.pptx });
    if (only.includes('deckfx')) await verifyDeckFx({ open, check, until, wait, press, errorsIn, doc, sessionFor }, { file: files.pptx });
    if (only.includes('providers')) await mailProviders();
    if (only.includes('signature')) await mailSignature();
    if (only.includes('sendlater')) await mailSendLater();
    if (only.includes('ooo')) await mailOOO();
    return done();
  }

  /* ── A broken file gets a sentence, not a blank window ───────────────── */
  //
  // A truncated deck made the Presentation window throw "rendered fewer
  // hooks than expected" and show nothing for as long as the person waited:
  // the error return sat above a hook. Each document app opens a file of
  // junk here and must show "This file could not be opened" with the reason,
  // within seconds, with no exception in the console.
  for (const [appName, ext] of [['word', 'docx'], ['sheets', 'xlsx'], ['slides', 'pptx']]) {
    try {
      const broken = path.join(path.dirname(files.docx), `broken.${ext}`);
      fs.writeFileSync(broken, Buffer.from('this is not a zip archive, whatever the name says. '.repeat(400)));
      const win = await open(appName, broken);
      const consoleErrors = [];
      win.webContents.on('console-message', (...args) => { const m = consoleMessage(args); if (m.level >= 2) consoleErrors.push(m.text.slice(0, 120)); });
      await until(() => win.webContents.executeJavaScript(`[...document.querySelectorAll('.rw-empty h3')].some((h) => /could not be opened/i.test(h.textContent))`), 'the refusal to show', 6000).catch(() => {});
      const shown = await win.webContents.executeJavaScript(`(() => {
        const h = [...document.querySelectorAll('.rw-empty h3')].find((x) => /could not be opened/i.test(x.textContent));
        return { refused: Boolean(h), reason: h ? (h.parentElement.textContent || '').replace(h.textContent, '').trim().slice(0, 80) : null, spinner: Boolean(document.querySelector('.rw-spinner')) };
      })()`);
      check(`${appName}: a broken file gets "could not be opened" and a reason, not a blank window`, shown.refused && shown.reason && !shown.spinner && consoleErrors.length === 0, `${JSON.stringify(shown)}${consoleErrors.length ? `; console: ${consoleErrors[0]}` : ''}`);
    } catch (err) {
      check(`${appName}: the broken-file check ran`, false, err.message);
    }
  }

  // The owner file Word keeps beside an open document: 162 bytes named after
  // it with "~$" in front. The corpus run over a Downloads folder began with
  // four of them, each called a stopped download.
  try {
    const owner = path.join(path.dirname(files.docx), '~$report.docx');
    fs.writeFileSync(owner, Buffer.alloc(162, 7));
    const win = await open('word', owner);
    await until(() => win.webContents.executeJavaScript(`[...document.querySelectorAll('.rw-empty h3')].some((h) => /could not be opened/i.test(h.textContent))`), 'the refusal to show', 6000).catch(() => {});
    const reason = await win.webContents.executeJavaScript(`(document.querySelector('.rw-empty')?.textContent || '').slice(0, 300)`);
    check('word: an owner file is called what it is, and names the document to open instead', /owner file/.test(reason) && /Open report.docx itself/.test(reason), JSON.stringify(reason.slice(0, 120)));
  } catch (err) {
    check('word: the owner-file check ran', false, err.message);
  }

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

  } catch (err) {
    check('word: the round trip ran', false, err.message);
  }

  await wordPages();
  await wordFloat();
  await sheetGrips();
  await slidePanes();
  await slideShapes();
  await slideParagraphs();
  await slideClipboard();
  await slideFooter();
  await slideFind();
  await slideSections();
  await slideHidden();
  await slideBackground();
  await slideTable();
  await slideChart();
  await verifyDeckArrange({ open, check, until, wait, press, errorsIn, doc, sessionFor }, { file: files.pptx });
  await verifyDeckFx({ open, check, until, wait, press, errorsIn, doc, sessionFor }, { file: files.pptx });
  await sheetFill();
  await wordPictures();
  await wordLook();
  await wordDropCap();
  await wordBookmarks();
  await wordCrossRef();
  await wordCaptions();
  await wordToc();
  await wordTrack();
  await wordEffects();
  await wordPictureFits();
  await wordCards();
  await wordRuler();
  await wordColumns();
  await updatePrompt();
  await launcherRecent();
  await homeRecentEdit();
  await sheetFreeze();
  await sheetErrors();
  await sheetWatch();
  await sheetSparklines();
  await zoomStaysOnThePage();
  await sheetPicture();
  await viewer();
  await slideshow();
  await sheetLinks();
  await sheetOutline();
  await polish();

  /* ── Worksheets: type a value, save, reopen ──────────────────────────── */

  try {
    const win = await open('sheets', files.xlsx);

    // Focused, and confirmed focused. Asking for focus and then typing 250 ms
    // later is a race: on a busy machine the keystrokes arrive before the grid
    // has it, land nowhere, and the check fails describing a save that was
    // never asked to happen.
    await until(async () => {
      raise(win);
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
      raise(win);
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

      // With a picture open the folder is the filmstrip; the grid is a panel a person turns on.
      const tiles = await win.webContents.executeJavaScript(`document.querySelectorAll('.pv-tile, .pv-strip-item').length`);
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
    // One card per app the registry knows; a card missing is an app the launcher lost.
    check('home: the launcher is whole regardless', state.apps === Object.keys(APPS).length, `${state.apps} app cards of ${Object.keys(APPS).length}, ${state.notices} notices`);
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
        const b = [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}));
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
      `[...document.querySelectorAll('.rw-wincontrols button, .rw-titlebar button')].some((b) => /full screen/i.test(b.title || b.dataset.tip || ''))`
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
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}));
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
      return cs ? { weight: cs.fontWeight, size: cs.fontSize, colour: cs.color, align: getComputedStyle(p).textAlign, span: (span.className || span.tagName) + ':' + span.textContent.slice(0, 24) } : null;
    })()`);

    const selected = await selectBlock1();
    await wait(300);
    check('word: words can be selected on the page', selected > 10, `${selected} characters selected`);

    // Bold, from the button. Focus must stay on the page for this to land.
    const b = await pressButton('Bold');
    await until(async () => (await state()).block1?.runs?.some((r) => r.bold), 'the engine to hold bold', 4000).catch(() => {});
    const afterBold = await state();
    const paintBold = await painted();
    check('word: the Bold button makes the selection bold', afterBold.block1?.runs?.some((r) => r.bold) && paintBold?.weight === '700', `${b}; measured ${JSON.stringify(paintBold)}; engine bold=${afterBold.block1?.runs?.some((r) => r.bold)}, page weight=${paintBold?.weight}`);

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
    await pressButton('Font colour');
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

    await pressButton('Bullets');
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
    await pressButton('Orientation');
    await wait(200);
    await pressMenu('Landscape');
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
      // The bytes survived the trip from the file dialog: what the page drew
      // decodes, and is the file's own. A part written from the digits of a
      // byte array — which is what happened until 2026-09-21 — draws nothing.
      const decoded = await until(() => js(`(() => { const i = document.querySelector('.wd-page img'); return Boolean(i && i.complete && i.naturalWidth > 0); })()`), 'the picture to decode', 4000).catch(() => false);
      const srcAttr = await js(`document.querySelector('.wd-page img')?.getAttribute('src') || ''`);
      const sameBytes = (srcAttr.split(',')[1] || '') === fs.readFileSync(files.png).toString('base64');
      check('word: the picture the page drew decodes, and is the file’s own bytes', decoded === true && sameBytes, `decoded ${decoded}; same bytes ${sameBytes} (${srcAttr.length} chars of src, ${fs.readFileSync(files.png).length} bytes of file)`);
      // Picked with a click and deleted from the keyboard: its block goes too.
      await js(`(() => { document.querySelector('.wd-page img')?.click(); return 1; })()`);
      await until(() => js(`Boolean(document.querySelector('.wd-page img.picked'))`), 'the picture picked', 3000).catch(() => {});
      await js(`document.querySelector('.wd-page')?.focus(), 'focused'`);
      await press(win.webContents, 'Delete');
      const gone = await until(async () => (await state()).blocksTotal === before && (await js(`document.querySelectorAll('.wd-page img').length`)) === 0, 'the picture deleted', 5000).catch(() => false);
      check('word: a picked picture goes with the Delete key, its block with it', gone === true, `${(await state()).blocksTotal} blocks (was ${before}), ${await js(`document.querySelectorAll('.wd-page img').length`)} image(s)`);
      check('word: a picture inserts as its own block and is drawn', pictured.blocksTotal > before && drawnImage > 0, `${inserted}; ${before} → ${pictured.blocksTotal} blocks, ${drawnImage} image(s) on the page`);
    }

    const complaints = await errorsIn(win);
    check('word: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the button checks ran', false, err.message);
  }

  /* ── Word: the rest of the ribbon, tab by tab ─────────────────────────── */

  try {
    const win = await open('word', files.docx);
    const js = (code) => win.webContents.executeJavaScript(code);
    const tabTo = (label) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(label)})?.click(), 'tab'`);
    const press = (title) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}) && !n.disabled);
      if (!b) return 'no live button ' + ${JSON.stringify(title)};
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      b.click();
      return 'clicked';
    })()`);
    const menuItem = (label) => js(`(() => { const i = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); if (!i) return 'no item'; i.click(); return 'clicked'; })()`);
    const engine = () => js(`(async () => {
      const all = await window.rutbaOffice.doc.sessions({});
      const mine = all.filter((s) => s.kind === 'doc').pop();
      const m = await window.rutbaOffice.doc.model({ id: mine.id });
      return { blocks: m.blocks.length, texts: m.blocks.map((b) => b.text), styles: m.blocks.map((b) => b.style), format: m.format, footnotes: (m.footnotes || []).length, toc: (m.tableOfContents?.entries || []).map((e) => e.text) };
    })()`);

    // The page sits in the middle of the window, whatever the ribbon is doing.
    const centred = await js(`(() => {
      const page = document.querySelector('.wd-page').getBoundingClientRect();
      const scroll = document.querySelector('.wd-scroll');
      return { left: Math.round(page.left), right: Math.round(page.right), window: innerWidth, scrollClient: scroll.clientWidth, scrollWidth: scroll.scrollWidth, body: document.querySelector('.rw-body')?.clientWidth };
    })()`);
    check('word: the page is centred in the window', Math.abs((centred.left + centred.right) / 2 - centred.window / 2) < 40, JSON.stringify(centred));

    // Every tab is reachable and draws its groups.
    const tabs = ['Home', 'Insert', 'Draw', 'Design', 'Layout', 'References', 'Mailings', 'Review', 'View', 'Help', 'PDF'];
    const groups = {};
    for (const t of tabs) {
      await tabTo(t);
      await wait(120);
      groups[t] = await js(`document.querySelectorAll('.rw-ribbon .rw-group').length`);
    }
    check('word: every Word tab is there and draws its groups', tabs.every((t) => groups[t] > 0), tabs.map((t) => `${t} ${groups[t]}`).join(', '));

    // What is not wired says so, on the button, rather than pretending.
    const honest = await js(`(() => {
      const dead = [...document.querySelectorAll('.rw-ribbon .rw-btn[disabled]')];
      // Undo and Redo are disabled because there is nothing to undo — that is
      // state, not an apology, and needs no explanation.
      return { count: dead.length, unexplained: dead.filter((b) => !/not built yet/.test(b.title || b.dataset.tip || '') && !/^(Undo|Redo) /.test(b.title || b.dataset.tip || '')).map((b) => b.title || b.dataset.tip || b.textContent.trim()).slice(0, 5) };
    })()`);
    check('word: every disabled control explains itself', honest.unexplained.length === 0, `${honest.count} disabled on the PDF tab; unexplained: ${JSON.stringify(honest.unexplained)}`);

    // View: the navigation pane lists the heading, and clicking it moves the caret.
    await tabTo('View');
    await wait(120);
    await press('Headings, to move around');
    await until(() => js(`document.querySelectorAll('.wd-nav-item').length > 0`), 'the navigation pane', 4000).catch(() => {});
    const navItems = await js(`[...document.querySelectorAll('.wd-nav-item')].map((n) => n.textContent.trim())`);
    check('word: the Navigation Pane lists the headings', navItems.includes('Quarterly Review'), JSON.stringify(navItems));

    await press('Web Layout');
    await wait(150);
    const web = await js(`document.querySelector('.wd')?.className || ''`);
    await press('Print Layout');
    check('word: View → Web Layout changes the view', /mode-web/.test(web), `class was ${JSON.stringify(web)}`);

    // Insert: a symbol and the date, through their dialogs.
    await tabTo('Insert');
    await wait(120);
    await js(`(() => { const page = document.querySelector('.wd-page'); const b = page.querySelector('[data-block="2"]'); page.focus(); const r = document.createRange(); r.selectNodeContents(b); r.collapse(false); const s = getSelection(); s.removeAllRanges(); s.addRange(r); page.dispatchEvent(new MouseEvent('mouseup', { bubbles: true })); return 'caret at end of block 2'; })()`);
    await wait(200);
    await press('Symbol');
    await wait(250);
    await js(`(() => { const b = [...document.querySelectorAll('.wd-symbol')].find((n) => n.textContent.trim() === '©'); b?.click(); return 'symbol'; })()`);
    await until(async () => (await engine()).texts.some((t) => /©/.test(t)), 'the symbol to land', 4000).catch(() => {});
    await js(`[...document.querySelectorAll('.rw-dialog .rw-btn')].find((b) => b.textContent.trim() === 'Close')?.click(), 'closed'`);
    const withSymbol = await engine();
    check('word: Insert → Symbol inserts the character', withSymbol.texts.some((t) => /©/.test(t)), JSON.stringify(withSymbol.texts.find((t) => /©/.test(t)) || '').slice(0, 60));

    await press('Date & Time');
    await wait(250);
    await js(`document.querySelector('.rw-dialog .ml-found-item')?.click(), 'first format'`);
    const year = String(new Date().getFullYear());
    await until(async () => (await engine()).texts.some((t) => t.includes(year)), 'the date to land', 4000).catch(() => {});
    const withDate = await engine();
    check('word: Insert → Date & Time inserts today', withDate.texts.some((t) => t.includes(year)), `a paragraph contains ${year}: ${withDate.texts.some((t) => t.includes(year))}`);

    // Home: the pilcrow button, change case on a selection, a heading style.
    await tabTo('Home');
    await wait(120);
    const pressed = await press('Show formatting marks');
    await until(() => js(`document.querySelector('.wd-page').classList.contains('marks')`), 'the marks', 3000).catch(() => {});
    const marks = await js(`document.querySelector('.wd-page').classList.contains('marks')`);
    check('word: formatting marks toggle on', marks === true, `page has marks=${marks} (${pressed})`);
    await press('Show formatting marks');

    await js(`(() => { const page = document.querySelector('.wd-page'); const b = page.querySelector('[data-block="1"]'); const r = document.createRange(); r.selectNodeContents(b); const s = getSelection(); s.removeAllRanges(); s.addRange(r); page.dispatchEvent(new MouseEvent('mouseup', { bubbles: true })); return s.toString().length; })()`);
    await wait(200);
    await press('Change case');
    await wait(200);
    await menuItem('UPPERCASE');
    await until(async () => /^[A-Z0-9 .,%'-]+$/.test((await engine()).texts[1] || 'x'), 'the text to go upper case', 4000).catch(() => {});
    const upper = await engine();
    check('word: Home → Change case makes the selection UPPERCASE', /[A-Z]{4}/.test(upper.texts[1]) && upper.texts[1] === upper.texts[1].toUpperCase(), JSON.stringify(upper.texts[1]).slice(0, 60));

    // Superscript on the (still selected) paragraph — the engine writes the
    // run property and the page paints it raised.
    await js(`(() => { const page = document.querySelector('.wd-page'); const b = page.querySelector('[data-block="1"]'); const r = document.createRange(); r.selectNodeContents(b); const s = getSelection(); s.removeAllRanges(); s.addRange(r); page.dispatchEvent(new MouseEvent('mouseup', { bubbles: true })); return s.toString().length; })()`);
    await wait(200);
    await press('Superscript');
    await until(async () => (await engine()).format?.vertAlign === 'superscript', 'the superscript', 4000).catch(() => {});
    const sup = await engine();
    const raised = await js(`getComputedStyle(document.querySelector('[data-block="1"] span')).verticalAlign`);
    check('word: Home → x² makes the selection superscript, and it is painted raised', sup.format?.vertAlign === 'superscript' && raised === 'super', `engine ${sup.format?.vertAlign}; painted vertical-align ${raised}`);
    await press('Superscript');
    await until(async () => !(await engine()).format?.vertAlign, 'the superscript to clear', 4000).catch(() => {});

    // A heading, through the style box — the table of contents needs one.
    await js(`(() => {
      const s = [...document.querySelectorAll('.rw-ribbon select')].find((n) => n.title === 'Paragraph style');
      if (!s) return 'no style box';
      Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(s, 'Heading1');
      s.dispatchEvent(new Event('change', { bubbles: true }));
      return 'Heading1';
    })()`);
    await until(async () => (await engine()).styles[1] === 'Heading1', 'the heading style', 4000).catch(() => {});
    const styled = await engine();
    check('word: the style box makes a paragraph a Heading 1', styled.styles[1] === 'Heading1', `block 1 style ${JSON.stringify(styled.styles[1])}`);

    // References: a table of contents built from the headings.
    await tabTo('References');
    await wait(120);
    await js(`(() => { const page = document.querySelector('.wd-page'); const b = page.querySelector('[data-block="0"]'); const r = document.createRange(); r.selectNodeContents(b); r.collapse(true); const s = getSelection(); s.removeAllRanges(); s.addRange(r); page.dispatchEvent(new MouseEvent('mouseup', { bubbles: true })); return 'caret at start'; })()`);
    await wait(200);
    const before = (await engine()).blocks;
    // Since 1.27.0 the table is a TOC field: its entries are read off the
    // model's own tableOfContents, not searched for as a "Contents" line.
    await press('Table of Contents');
    await until(async () => (await engine()).toc.length > 0, 'the table of contents to appear', 6000).catch(() => {});
    const toc = await engine();
    check('word: References → Table of Contents lists the headings', toc.toc.includes(styled.texts[1]) && toc.blocks > before, `${before} → ${toc.blocks} blocks; entries ${JSON.stringify(toc.toc)}`);

    // References → Insert Footnote: the dialog takes the words, the page
    // gets a raised number at the caret and the note under the body.
    // The table of contents added blocks at the top; the note goes on the
    // paragraph that was block 2 before it, not on a line inside the table.
    await js(`(() => { const page = document.querySelector('.wd-page'); const b = page.querySelector('[data-block="${2 + (toc.blocks - before)}"]'); const r = document.createRange(); r.selectNodeContents(b); r.collapse(false); const s = getSelection(); s.removeAllRanges(); s.addRange(r); page.dispatchEvent(new MouseEvent('mouseup', { bubbles: true })); return 'caret at end of the paragraph that was block 2'; })()`);
    await wait(200);
    await press('A raised number at the caret, and its words under the body');
    await until(() => js(`Boolean(document.querySelector('.rw-dialog textarea'))`), 'the footnote dialog', 4000).catch(() => {});
    await js(`(() => { const ta = document.querySelector('.rw-dialog textarea'); if (!ta) return 'no textarea'; Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(ta, 'See regulation 57.'); ta.dispatchEvent(new Event('input', { bubbles: true })); return 'typed'; })()`);
    await wait(150);
    await js(`[...document.querySelectorAll('.rw-dialog .rw-btn')].find((b) => b.textContent.trim() === 'Insert')?.click(), 'inserted'`);
    await until(async () => (await engine()).footnotes === 1, 'the footnote', 5000).catch(() => {});
    const noted = await engine();
    await wait(500);
    const painted = await js(`({ refs: document.querySelectorAll('.wd-page .wd-noteref').length, notes: document.querySelectorAll('.wd-pagenotes .wd-note').length, words: document.querySelector('.wd-pagenotes')?.textContent.trim() || '' })`);
    check('word: References → Insert Footnote puts the number in the text and the note at the foot of the page', noted.footnotes === 1 && painted.refs === 1 && painted.notes === 1 && /regulation 57/.test(painted.words), `engine footnotes ${noted.footnotes}; painted ${JSON.stringify(painted)}`);
    // The note sits at the foot of the page its reference is on: above the
    // bottom margin, below the last line of the body, on the same sheet.
    const foot = await js(`(() => {
      const page = document.querySelector('.wd-page');
      const sheets = [...page.querySelectorAll('.wd-sheet')].map((s) => s.getBoundingClientRect());
      const on = (r) => sheets.findIndex((s) => r.top >= s.top - 1 && r.bottom <= s.bottom + 1);
      const ref = page.querySelector('.wd-noteref');
      const notes = page.querySelector('.wd-pagenotes .wd-notes');
      if (!ref || !notes) return { ref: Boolean(ref), notes: Boolean(notes) };
      const r = ref.getBoundingClientRect();
      const nb = notes.getBoundingClientRect();
      const s = sheets[on(nb)];
      const mBottom = parseFloat(getComputedStyle(page).paddingBottom);
      const body = [...page.querySelectorAll('.wd-block, .wd-table')].map((b) => b.getBoundingClientRect()).filter((b) => s && b.height > 0 && b.top >= s.top && b.top < s.bottom);
      const lastBody = body.length ? Math.max(...body.map((b) => b.bottom)) : null;
      return { refPage: on(r), notesPage: on(nb), footGap: s ? Math.round(s.bottom - mBottom - nb.bottom) : null, belowBody: lastBody === null ? null : Math.round(nb.top - lastBody), measure: getComputedStyle(page.querySelector('.wd-notes-measure')).visibility };
    })()`);
    check('word: the footnote sits at the foot of its reference\'s page — under the body, on the bottom margin, the hidden copy hidden', foot.refPage >= 0 && foot.refPage === foot.notesPage && foot.footGap !== null && foot.footGap >= -1 && foot.footGap <= 4 && foot.belowBody !== null && foot.belowBody >= 0 && foot.measure === 'hidden', JSON.stringify(foot));
    if (process.env.RUTBA_VERIFY_CAPTURE) {
      await js(`document.querySelector('.wd-pagenotes')?.scrollIntoView({ block: 'end' }), 'scrolled to the foot'`);
      await wait(300);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'word-footnote.png'), (await win.webContents.capturePage()).toPNG());
    }

    const complaints = await errorsIn(win);
    check('word: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the ribbon-tab checks ran', false, err.message);
  }

  /* ── Worksheets: the rest of the ribbon, tab by tab ───────────────────── */

  try {
    const win = await open('sheets', files.xlsx);
    const js = (code) => win.webContents.executeJavaScript(code);
    const model = () => js(`(async () => {
      const all = await window.rutbaOffice.doc.sessions({});
      const mine = all.filter((s) => s.kind === 'sheet').pop();
      return window.rutbaOffice.doc.model({ id: mine.id });
    })()`);
    const tabTo = (label) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(label)})?.click(), 'tab'`);
    const pushTitle = (title) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}) && !n.disabled);
      if (!b) return 'no live button ' + ${JSON.stringify(title)};
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked';
    })()`);
    const pushLabel = (label) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(label)} && !n.disabled);
      if (!b) return 'no live button ' + ${JSON.stringify(label)};
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked';
    })()`);
    const menuItem = (label) => js(`(() => { const i = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); if (!i) return 'no item ' + ${JSON.stringify(label)}; i.click(); return 'clicked'; })()`);
    // Move the selection by reference through the engine, the way Go To does.
    const act_goto = async (ref) => {
      await js(`(async () => { const all = await window.rutbaOffice.doc.sessions({}); const mine = all.filter((s) => s.kind === 'sheet').pop(); const col = ${JSON.stringify(ref)}.charCodeAt(0) - 65; const row = Number(${JSON.stringify(ref)}.slice(1)) - 1; await window.rutbaOffice.doc.apply({ id: mine.id, ops: [{ op: 'select', row, col }] }); return 'moved'; })()`);
      await wait(150);
    };

    // Every Excel tab is reachable and draws its groups; what is not wired says so.
    const tabs = ['Home', 'Insert', 'Draw', 'Page Layout', 'Formulas', 'Data', 'Review', 'View', 'Automate', 'Help'];
    const groups = {};
    for (const t of tabs) {
      await tabTo(t);
      await wait(120);
      groups[t] = await js(`document.querySelectorAll('.rw-ribbon .rw-group').length`);
    }
    check('sheets: every Excel tab is there and draws its groups', tabs.every((t) => groups[t] > 0), tabs.map((t) => `${t} ${groups[t]}`).join(', '));
    const honest = await js(`(() => {
      const dead = [...document.querySelectorAll('.rw-ribbon .rw-btn[disabled]')];
      return { count: dead.length, unexplained: dead.filter((b) => !/not built yet/.test(b.title || b.dataset.tip || '') && !/^(Undo|Redo) /.test(b.title || b.dataset.tip || '')).map((b) => b.title || b.dataset.tip || b.textContent.trim()).slice(0, 5) };
    })()`);
    check('sheets: every disabled control explains itself', honest.unexplained.length === 0, `${honest.count} disabled on the Help tab; unexplained: ${JSON.stringify(honest.unexplained)}`);

    // View: the three Show toggles change the grid, and freeze works from the menu.
    await tabTo('View');
    await wait(120);
    await pushLabel('Gridlines');
    await until(() => js(`Boolean(document.querySelector('.sh.no-grid'))`), 'gridlines off', 3000).catch(() => {});
    const noGrid = await js(`Boolean(document.querySelector('.sh.no-grid'))`);
    await pushLabel('Gridlines');
    await pushLabel('Headings');
    await until(() => js(`Boolean(document.querySelector('.sh.no-heads'))`), 'headings off', 3000).catch(() => {});
    const noHeads = await js(`({ off: Boolean(document.querySelector('.sh.no-heads')), rails: getComputedStyle(document.querySelector('.sh-colheads')).display })`);
    await pushLabel('Headings');
    await pushLabel('Formula Bar');
    await until(() => js(`document.querySelector('.sh-formula').hidden`), 'the formula bar to hide', 3000).catch(() => {});
    const noBar = await js(`document.querySelector('.sh-formula').hidden`);
    await pushLabel('Formula Bar');
    check('sheets: View → Gridlines, Headings and Formula Bar toggle what they name', noGrid && noHeads.off && noHeads.rails === 'none' && noBar, `gridlines off ${noGrid}; headings ${JSON.stringify(noHeads)}; formula bar hidden ${noBar}`);

    await pushLabel('Freeze Panes');
    await wait(150);
    await menuItem('Freeze top row');
    await until(async () => (await model()).frozen?.rows === 1, 'the top row to freeze', 4000).catch(() => {});
    const frozen = (await model()).frozen;
    check('sheets: View → Freeze Panes → Freeze top row freezes it', frozen?.rows === 1 && !frozen?.cols, JSON.stringify(frozen));
    await pushLabel('Freeze Panes');
    await wait(150);
    await menuItem('Unfreeze panes');

    // Home: a cell style paints, the decimal buttons rewrite the number format, Go To moves.
    await tabTo('Home');
    await wait(120);
    await js(`(() => { const c = document.querySelector('.sh-cell'); c?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 })); return c?.textContent; })()`);
    await wait(200);
    await pushLabel('Cell Styles');
    await wait(150);
    await menuItem('Good');
    await until(async () => /c6efce/i.test(String((await model()).format?.fill || '')), 'the Good style', 4000).catch(() => {});
    const good = await model();
    const painted = await js(`getComputedStyle(document.querySelector('.sh-cell.active')).backgroundColor`);
    check('sheets: Home → Cell Styles → Good fills the cell Excel-green', /c6efce/i.test(String(good.format?.fill || '')) && painted === 'rgb(198, 239, 206)', `engine fill ${good.format?.fill}; painted ${painted}`);
    await pushLabel('Cell Styles');
    await wait(150);
    await menuItem('Normal');

    await pushTitle('Increase decimal');
    await until(async () => /\.0/.test(String((await model()).format?.numberFormat || '')), 'a decimal', 4000).catch(() => {});
    const dec = (await model()).format?.numberFormat;
    await pushTitle('Decrease decimal');
    check('sheets: Home → Increase decimal turns General into 0.0', dec === '0.0', `number format ${JSON.stringify(dec)}`);

    await pushLabel('Find & Select');
    await wait(150);
    await menuItem('Go To…');
    await until(() => js(`Boolean(document.querySelector('.rw-dialog input'))`), 'the Go To dialog', 3000).catch(() => {});
    await js(`(() => { const i = document.querySelector('.rw-dialog input'); if (!i) return 'no input'; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(i, 'B3'); i.dispatchEvent(new Event('input', { bubbles: true })); return 'typed'; })()`);
    await wait(120);
    await js(`[...document.querySelectorAll('.rw-dialog .rw-btn')].find((b) => b.textContent.trim() === 'Go')?.click(), 'go'`);
    await until(async () => (await model()).selection?.ref === 'B3', 'B3', 4000).catch(() => {});
    const went = (await model()).selection?.ref;
    check('sheets: Home → Find & Select → Go To moves the selection', went === 'B3', `selection ${went}`);

    // Text wider than its cell spills over an empty neighbour, as in Excel —
    // and stops spilling the moment the neighbour has something in it.
    // Row 9 is empty in the fixture; row 1 holds the headers, and a header
    // next door is exactly what stops a spill.
    await act_goto('A9');
    await js(`document.querySelector('.sh')?.focus(), 'ok'`);
    await typeText(win.webContents, 'A sentence far too long for one column');
    await press(win.webContents, 'Return', { char: true });
    await until(() => js(`(document.querySelector('.sh-cell[data-ref="A9"]')?.offsetWidth || 0) > (document.querySelector('.sh-colheads .sh-head')?.offsetWidth || 0) + 10`), 'the text to spill', 4000).catch(() => {});
    const spilled = await js(`({ cell: document.querySelector('.sh-cell[data-ref="A9"]')?.offsetWidth || 0, column: document.querySelector('.sh-colheads .sh-head')?.offsetWidth || 0 })`);
    // The commit is asynchronous; moving on before it lands races the next
    // keystrokes into the old edit.
    await until(async () => { const m = await model(); return !m.editing && m.cells?.some((c) => c.ref === 'A9' && /sentence/.test(c.text)); }, 'A9 to commit', 4000).catch(() => {});
    // B9 is filled through the engine (typing has its own checks), and the
    // window learns of it on its next operation of its own — a click on A9.
    await js(`(async () => { const all = await window.rutbaOffice.doc.sessions({}); const mine = all.filter((s) => s.kind === 'sheet').pop(); await window.rutbaOffice.doc.apply({ id: mine.id, ops: [{ op: 'setCell', row: 8, col: 1, value: 'x' }] }); return 'set'; })()`);
    await js(`(() => { const c = document.querySelector('.sh-cell[data-ref="A9"]'); c?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 })); return 'clicked'; })()`);
    await until(async () => (await model()).cells?.some((c) => c.ref === 'B9' && c.text === 'x'), 'the x to land in B9', 4000).catch(() => {});
    const b9 = (await model()).cells?.find((c) => c.ref === 'B9')?.text ?? null;
    await until(() => js(`(document.querySelector('.sh-cell[data-ref="A9"]')?.offsetWidth || 0) <= (document.querySelector('.sh-colheads .sh-head')?.offsetWidth || 0) + 1`), 'the spill to stop', 4000).catch(() => {});
    const stopped = await js(`document.querySelector('.sh-cell[data-ref="A9"]')?.offsetWidth || 0`);
    check('sheets: long text spills over an empty neighbour and stops when it fills', spilled.cell > spilled.column + 10 && stopped <= spilled.column + 1, `A9 drew ${spilled.cell}px over a ${spilled.column}px column; ${stopped}px once B9 held ${JSON.stringify(b9)}`);

    // Formulas → Insert Function → SUM starts the edit with =SUM( typed.
    await tabTo('Formulas');
    await wait(120);
    await pushLabel('Insert Function');
    await until(() => js(`Boolean(document.querySelector('.rw-dialog'))`), 'the function dialog', 3000).catch(() => {});
    await js(`[...document.querySelectorAll('.rw-dialog .ml-found-item')].find((b) => b.textContent.trim() === 'SUM')?.click(), 'picked'`);
    await until(() => js(`(document.querySelector('.sh-editor')?.value || '').startsWith('=SUM(')`), 'the edit to start', 4000).catch(() => {});
    const editor = await js(`document.querySelector('.sh-editor')?.value ?? null`);
    check('sheets: Formulas → Insert Function → SUM starts the edit with =SUM(', editor === '=SUM(', `editor ${JSON.stringify(editor)}`);
    await js(`document.querySelector('.sh-editor')?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })), 'esc'`);

    const complaints = await errorsIn(win);
    check('sheets: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('sheets: the ribbon-tab checks ran', false, err.message);
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
      if (!layer) return 'no grid yet';
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

    // The grid layer mounts a moment after the window says it is ready — later on a loaded machine.
    await until(() => js(`Boolean(document.querySelector('.sh-cells') && document.querySelectorAll('.sh-colheads .sh-head').length > 2)`), 'the grid and its headers', 8000).catch(() => {});

    // An idle window must settle. The grid once stretched to its own canvas,
    // the canvas grew to cover the grid, and the size observer asked the main
    // process for a wider frame eight times a second for as long as the window
    // lived — 1,088 px wider each time. Every open Worksheets window did it,
    // and the main process was never idle again.
    const gridSize = () => js(`(() => { const g = document.querySelector('.sh-grid'); return g ? { width: g.clientWidth, canvas: g.scrollWidth, window: window.innerWidth } : null; })()`);
    // The engine's first frame can still be on its way when the headers show
    // — later on a loaded machine — and it widens the canvas once, which is
    // not the growth this guards against. Two reads 300 ms apart have to
    // agree before the idle spell starts; a canvas that keeps growing never
    // agrees and still fails below.
    let before = await gridSize();
    for (let i = 0; i < 12; i++) {
      await wait(300);
      const now = await gridSize();
      if (now && before && now.canvas === before.canvas && now.width === before.width) break;
      before = now;
    }
    const settled1 = await gridSize();
    await wait(2500);
    const settled2 = await gridSize();
    check(
      'sheets: an idle window settles — the grid stays inside the window and stops growing',
      Boolean(settled1 && settled2) && settled2.width <= settled2.window && settled2.width === settled1.width && settled2.canvas === settled1.canvas,
      `grid ${settled1?.width} → ${settled2?.width} px wide (window ${settled2?.window}); canvas ${settled1?.canvas} → ${settled2?.canvas}`
    );
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
      // By the START of the title: "Bold (Ctrl+B)" is still the Bold button.
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}));
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
    // The menu, not a fixed pause: under load 200 ms was not always enough
    // and the choice below clicked nothing.
    await until(() => js(`Boolean(document.querySelector('.rw-menu button'))`), 'the fill menu', 3000).catch(() => {});
    await js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === 'Light amber')?.click(); return 'chose'; })()`);
    await until(async () => (await activeStyle())?.background === 'rgb(255, 243, 191)', 'the fill to paint', 4000).catch(() => {});
    const filled = await activeStyle();
    check('sheets: the fill menu paints the cell', filled?.background === 'rgb(255, 243, 191)', `background=${filled?.background}`);

    await clickRibbon('Borders');
    await until(() => js(`Boolean(document.querySelector('.rw-menu button'))`), 'the borders menu', 3000).catch(() => {});
    await js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === 'All borders')?.click(); return 'chose'; })()`);
    // Chromium floors a border to whole device pixels, so on a 175% display
    // a 1px border computes as 0.571px: any width at all is the border.
    const hasBorder = (s) => parseFloat(s?.borderTop || '0') > 0;
    await until(async () => hasBorder(await activeStyle()), 'the border to paint', 4000).catch(() => {});
    const bordered = await activeStyle();
    check('sheets: the border menu paints a border', hasBorder(bordered), `border-top=${bordered?.borderTop}`);


    // A dialog from the ribbon, through to its effect: View → Freeze panes →
    // "Freeze the top row", and the frame has to say row 1 is frozen.
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'View')?.click(), 'view tab'`);
    await wait(150);
    // Freeze Panes is a menu now, as in Excel; "Choose…" opens the dialog.
    await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Freeze Panes'); if (!b) return 'no button'; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked'; })()`);
    await wait(200);
    await js(`[...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === 'Choose…')?.click(), 'choose'`);
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

  const clickIn = async (win, title) => {
    const find = `[...document.querySelectorAll('.rw-btn, .ml-compose-cta button')].find((n) => (n.title || n.dataset.tip || n.textContent || '').trim().startsWith(${JSON.stringify(title)}) && !n.disabled)`;
    // A tab's buttons arrive a render after the tab is chosen — later on a
    // loaded machine — and a fixed pause was missing them.
    await until(() => win.webContents.executeJavaScript(`Boolean(${find})`), `the ${title} button`, 3000).catch(() => {});
    return win.webContents.executeJavaScript(`(() => {
    const b = ${find};
    if (!b) return 'no button ' + ${JSON.stringify(title)};
    b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
    b.click();
    return 'clicked';
  })()`);
  };
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
    await clickIn(win, 'New Slide');
    await wait(200);
    await clickMenu(win, 'Title and Content');
    await until(async () => (await deckModel()).count === before + 1, 'the deck to gain a slide', 4000).catch(() => {});
    const grew = await deckModel();
    const thumbs = await win.webContents.executeJavaScript(`document.querySelectorAll('.sl-thumb').length`);
    check('slides: the New slide menu adds a slide', grew.count === before + 1 && thumbs === before + 1, `${before} → ${grew.count} slides, ${thumbs} thumbnails`);

    await clickTab(win, 'Insert');
    await wait(150);
    const shapesBefore = (await deckModel()).slide?.shapes?.length ?? 0;
    await clickIn(win, 'Text Box');
    await until(async () => ((await deckModel()).slide?.shapes?.length ?? 0) > shapesBefore, 'a text box to appear', 4000).catch(() => {});
    const shapesAfter = (await deckModel()).slide?.shapes?.length ?? 0;
    check('slides: Insert → Text box puts a text box on the slide', shapesAfter === shapesBefore + 1, `${shapesBefore} → ${shapesAfter} shapes`);

    // Insert → Pictures opens a native file dialog, which the harness cannot
    // drive (and which would swallow every later keystroke). The button's
    // state is checked, and the operation the dialog would make is sent
    // through the same door — the document service — with a real PNG.
    const picButton = await win.webContents.executeJavaScript(`(() => {
      const b = [...document.querySelectorAll('.rw-btn')].find((x) => (x.title || x.dataset.tip || '').startsWith('Pictures'));
      return b ? { found: true, disabled: b.disabled } : { found: false };
    })()`);
    const pictured = await win.webContents.executeJavaScript(`(async () => {
      const all = await window.rutbaOffice.doc.sessions({});
      const mine = all.filter((s) => s.kind === 'deck').pop();
      const active = Math.max(0, [...document.querySelectorAll('.sl-thumb')].findIndex((t) => t.classList.contains('active')));
      const bytes = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='), (c) => c.charCodeAt(0));
      const next = await window.rutbaOffice.doc.apply({ id: mine.id, ops: [{ op: 'addPicture', slide: active, contentType: 'image/png', name: 'dot.png', data: bytes }], slide: active, width: 640 });
      const shapes = next.model.slide.shapes;
      const pic = shapes[shapes.length - 1];
      return {
        kind: pic?.kind,
        geometry: pic?.geometry,
        drawn: /<image [^>]*href="[^"]+"/.test(next.model.slide.svg),
        inThumb: /<image /.test(next.model.outline[active]?.thumbnail || ''),
        shapes: shapes.length,
      };
    })()`);
    check(
      'slides: Insert → Pictures is live, and a picture lands on the slide, drawn and in its thumbnail',
      picButton.found && !picButton.disabled && pictured.kind === 'picture' && pictured.drawn && pictured.inThumb && pictured.shapes === shapesAfter + 1,
      JSON.stringify({ button: picButton, ...pictured })
    );

    // Home → Shapes → Oval: a preset shape in the theme's accent, selected as
    // soon as it lands, and drawn as an ellipse rather than a rectangle.
    await clickTab(win, 'Home');
    await wait(150);
    const shapesBeforeOval = (await deckModel()).slide?.shapes?.length ?? 0;
    await clickIn(win, 'Shapes');
    await wait(200);
    await clickMenu(win, 'Oval');
    await until(async () => ((await deckModel()).slide?.shapes?.length ?? 0) > shapesBeforeOval, 'an oval to appear', 4000).catch(() => {});
    const ovalModel = await deckModel();
    const oval = ovalModel.slide?.shapes?.slice(-1)[0];
    const ovalDom = await win.webContents.executeJavaScript(`(() => ({
      selected: document.querySelectorAll('.sl-hit.selected').length,
      ellipses: document.querySelectorAll('.sl-stage ellipse, .sl-fit ellipse, .sl-slide ellipse').length,
    }))()`);
    check(
      'slides: Home → Shapes → Oval puts an oval on the slide, filled, selected, drawn as an ellipse',
      oval?.kind === 'shape' && /<ellipse [^>]*fill="#/.test(ovalModel.slide?.svg || '') && ovalDom.selected === 1,
      JSON.stringify({ kind: oval?.kind, name: oval?.name, ...ovalDom, shapes: `${shapesBeforeOval} → ${ovalModel.slide?.shapes?.length}` })
    );



    await clickTab(win, 'Review');
    await wait(150);
    await clickIn(win, 'Speaker Notes');
    await wait(250);
    await setField(win, '.rw-dialog textarea', 'Say hello first.');
    await win.webContents.executeJavaScript(`[...document.querySelectorAll('.rw-dialog .rw-btn')].find((b) => b.textContent.trim() === 'Save')?.click(), 'saved'`);
    await until(async () => /hello first/.test((await deckModel()).slide?.notes || ''), 'the notes to save', 4000).catch(() => {});
    const noted = await deckModel();
    check('slides: the Speaker notes dialog saves notes', /hello first/.test(noted.slide?.notes || ''), JSON.stringify(noted.slide?.notes || ''));

    await clickTab(win, 'Home');
    await wait(150);
    await clickIn(win, 'Delete');
    await until(async () => (await deckModel()).count === before, 'the slide to go', 4000).catch(() => {});
    const shrank = await deckModel();
    check('slides: the Delete button removes the slide', shrank.count === before, `${grew.count} → ${shrank.count} slides`);

    const complaints = await errorsIn(win);
    check('slides: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('slides: the button checks ran', false, err.message);
  }

  /* ── Presentation: the rest of the ribbon, tab by tab ─────────────────── */

  try {
    const win = await open('slides', files.pptx);
    const js = (code) => win.webContents.executeJavaScript(code);
    const deckModel = () => js(`(async () => {
      const all = await window.rutbaOffice.doc.sessions({});
      const mine = all.filter((s) => s.kind === 'deck').pop();
      const active = [...document.querySelectorAll('.sl-thumb')].findIndex((t) => t.classList.contains('active'));
      return window.rutbaOffice.doc.model({ id: mine.id, slide: Math.max(0, active), width: 640 });
    })()`);
    const pushLabel = (label) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(label)} && !n.disabled);
      if (!b) return 'no live button ' + ${JSON.stringify(label)};
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked';
    })()`);
    const pushTitle = (title) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}) && !n.disabled);
      if (!b) return 'no live button ' + ${JSON.stringify(title)};
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked';
    })()`);

    // Every PowerPoint tab is reachable and draws its groups; what is not wired says so.
    const tabs = ['Home', 'Insert', 'Draw', 'Design', 'Transitions', 'Animations', 'Slide Show', 'Record', 'Review', 'View', 'Help', 'PDF'];
    const groups = {};
    for (const t of tabs) {
      await clickTab(win, t);
      await wait(120);
      groups[t] = await js(`document.querySelectorAll('.rw-ribbon .rw-group').length`);
    }
    check('slides: every PowerPoint tab is there and draws its groups', tabs.every((t) => groups[t] > 0), tabs.map((t) => `${t} ${groups[t]}`).join(', '));
    await clickTab(win, 'Home');
    await wait(120);
    const honest = await js(`(() => {
      const dead = [...document.querySelectorAll('.rw-ribbon .rw-btn[disabled]')];
      return { count: dead.length, unexplained: dead.filter((b) => !/not built yet|Select a text box first|copy a shape first|decides its own size|Changing it rescales/.test(b.title || b.dataset.tip || '') && !/^(Undo|Redo) /.test(b.title || b.dataset.tip || '')).map((b) => b.title || b.dataset.tip || b.textContent.trim()).slice(0, 5) };
    })()`);
    check('slides: every disabled control explains itself', honest.unexplained.length === 0, `${honest.count} disabled on Home with nothing selected; unexplained: ${JSON.stringify(honest.unexplained)}`);

    // View: the Slide Sorter shows every slide, Gridlines draw over the slide.
    await clickTab(win, 'View');
    await wait(120);
    await pushLabel('Slide Sorter');
    await until(() => js(`document.querySelectorAll('.sl-sortergrid .sl-sortercard').length > 0`), 'the sorter', 4000).catch(() => {});
    const sorted = await js(`({ cards: document.querySelectorAll('.sl-sortergrid .sl-sortercard').length, pics: document.querySelectorAll('.sl-sortergrid .sl-thumb-pic > svg').length })`);
    const count = (await deckModel()).count;
    await pushLabel('Normal');
    await until(() => js(`Boolean(document.querySelector('.sl-slide'))`), 'the normal view', 4000).catch(() => {});
    check('slides: View → Slide Sorter shows every slide as a picture', sorted.cards === count && sorted.pics === count, `${sorted.cards} cards, ${sorted.pics} pictures, ${count} slides`);
    await pushLabel('Gridlines');
    await until(() => js(`Boolean(document.querySelector('.sl-gridlines'))`), 'gridlines', 3000).catch(() => {});
    const gridded = await js(`Boolean(document.querySelector('.sl-gridlines'))`);
    await pushLabel('Gridlines');
    check('slides: View → Gridlines draws them over the slide', gridded, `gridlines shown ${gridded}`);

    // Home: click a text box, then Bold — the run in the file goes bold.
    await clickTab(win, 'Home');
    await wait(120);
    // The first shape WITH TEXT, by the model's order, which is the DOM's: a
    // picture or a bare shape ahead of it would take the click and have no
    // run to make bold.
    const textual = ((await deckModel()).slide?.shapes || []).findIndex((s) => s.text?.paragraphs?.some((p) => p.runs?.length));
    const picked = await js(`(() => { const hit = document.querySelectorAll('.sl-hit')[${Math.max(0, textual)}]; if (!hit) return 'no text box'; hit.click(); return 'selected'; })()`);

    await until(() => js(`Boolean(document.querySelector('.sl-hit.selected'))`), 'a selected text box', 3000).catch(() => {});
    await pushTitle('Bold');
    await until(async () => { const m = await deckModel(); return m.slide?.shapes?.some((s) => s.text?.paragraphs?.some((p) => p.runs?.some((r) => r.bold))); }, 'a bold run', 4000).catch(() => {});
    const m = await deckModel();
    const boldRuns = (m.slide?.shapes || []).flatMap((s) => (s.text?.paragraphs || []).flatMap((p) => (p.runs || []).filter((r) => r.bold))).length;
    const painted = await js(`/font-weight="bold"|font-weight:\\s*bold|<tspan[^>]*font-weight="700"|font-weight="700"/.test(document.querySelector('.sl-svg')?.innerHTML || '')`);
    check('slides: Home → Bold makes the selected text box bold, in the file and on the slide', picked === 'selected' && boldRuns > 0 && painted, `${picked}; bold runs ${boldRuns}; painted bold ${painted}`);

    const complaints = await errorsIn(win);
    check('slides: none of the ribbon checks reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('slides: the ribbon-tab checks ran', false, err.message);
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

    // With a picture open the ribbon is folded to its tabs and the picture
    // has the room; a tab click peeks the groups over it; the chevron opens
    // the ribbon back out.
    const folded = await js(`({ collapsed: document.querySelector('.rw-ribbon')?.classList.contains('collapsed'), groups: document.querySelectorAll('.rw-groups').length, stage: document.querySelector('.pv-stage')?.getBoundingClientRect().height || 0, window: innerHeight, padding: getComputedStyle(document.querySelector('.pv-stage')).paddingTop })`);
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'View')?.click(), 'peek'`);
    await until(() => js(`Boolean(document.querySelector('.rw-ribbon.peek .rw-groups'))`), 'the peek', 3000).catch(() => {});
    const peeked = await js(`({ peek: Boolean(document.querySelector('.rw-ribbon.peek .rw-groups')), stage: document.querySelector('.pv-stage')?.getBoundingClientRect().height || 0 })`);
    await js(`document.querySelector('.rw-collapse')?.click(), 'expand'`);
    await until(() => js(`!document.querySelector('.rw-ribbon')?.classList.contains('collapsed') && document.querySelectorAll('.rw-groups').length === 1`), 'the ribbon to open', 3000).catch(() => {});
    const opened = await js(`({ collapsed: document.querySelector('.rw-ribbon')?.classList.contains('collapsed'), groups: document.querySelectorAll('.rw-groups').length, stage: document.querySelector('.pv-stage')?.getBoundingClientRect().height || 0 })`);
    await js(`document.querySelector('.rw-collapse')?.click(), 'fold again'`);
    check('pictures: with a picture open the ribbon folds and the picture gets the room', folded.collapsed && folded.groups === 0 && folded.stage > folded.window * 0.7 && folded.padding === '4px' && peeked.peek && peeked.stage === folded.stage && !opened.collapsed && opened.groups === 1 && opened.stage < folded.stage, `folded ${JSON.stringify(folded)}; peek ${JSON.stringify(peeked)}; opened ${JSON.stringify(opened)}`);

    const complaints = await errorsIn(win);
    check('pictures: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('pictures: the button checks ran', false, err.message);
  }

  try {
    const win = await open('image', files.png);
    const js = (code) => win.webContents.executeJavaScript(code);
    // Fifteen seconds, not six: by this point a dozen check windows are open
    // and the compositor is slow to hand the decoded picture to a canvas.
    await until(() => js(`(document.querySelector('.im-canvas')?.width || 0) > 0`), 'the picture to reach the canvas', 15000);
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
    // The media element decodes on its own schedule; with a corpus run on
    // the same machine it once took past eight seconds.
    await until(() => js(`document.querySelectorAll('.vd-clip').length > 0`), 'the timeline', 20000);
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

  /* ── Mail and the calendar: an invitation in a message is answered from it ─ */
  //
  // The seed carries a message with a text/calendar part. Opening it shows
  // the invitation card; Accept keeps the event and opens a message to the
  // organizer with the reply attached. And Compose completes an address
  // from the people mail has seen.
  try {
    const win = await open('mail');
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.ml-row').length > 0`), 'the seeded messages', 8000);
    await js(`[...document.querySelectorAll('.ml-row')].find((r) => /Invitation: Quarterly numbers/.test(r.textContent))?.click(), 'opened'`);
    await until(() => js(`Boolean(document.querySelector('.ml-invite'))`), 'the invitation card', 8000).catch(() => {});
    const card = await js(`(() => { const c = document.querySelector('.ml-invite'); return c ? { title: c.querySelector('.ml-invite-title')?.textContent, who: c.querySelector('.ml-invite-who')?.textContent, buttons: [...c.querySelectorAll('.rw-btn')].map((b) => b.textContent.trim()) } : null; })()`);
    const rowsShown = await js(`[...document.querySelectorAll('.ml-row')].map((r) => r.textContent.slice(0, 48))`);
    check('mail: a message that carries an invitation shows it, with Accept, Tentative and Decline', Boolean(card) && card.title === 'Quarterly numbers' && /Amina/.test(card.who || '') && card.buttons.join(',') === 'Accept,Tentative,Decline', `${JSON.stringify(card)}; rows ${JSON.stringify(rowsShown)}`);
    await js(`[...document.querySelectorAll('.ml-invite .rw-btn')].find((b) => b.textContent.trim() === 'Accept')?.click(), 'accepted'`);
    await until(() => js(`Boolean(document.querySelector('input[placeholder^="someone@"]'))`), 'the reply message to open', 8000).catch(() => {});
    const reply = await js(`(() => { const to = document.querySelector('input[placeholder^="someone@"]'); const subject = [...document.querySelectorAll('input')].map((i) => i.value).find((v) => /^Accepted/.test(v)); return { to: to?.value || '', subject: subject || '', attachment: /reply\\.ics/i.test(document.body.textContent) }; })()`);
    const kept = await js(`(async () => { const c = await window.rutbaOffice.calendar.calendars({}); return c.reduce((n, x) => n + x.count, 0); })()`);
    check('mail: Accept keeps the event and opens the reply to the organizer with the file attached', /amina@northwind\.example/.test(reply.to) && /^Accepted: Quarterly numbers/.test(reply.subject) && reply.attachment && kept >= 1, `${JSON.stringify(reply)}; events kept ${kept}`);
    // The reply opens, then the window settles the attachment a frame later,
    // which once reset a To typed too soon; type after it has settled, and
    // type again if the first try was swallowed.
    await wait(400);
    await setField(win, 'input[placeholder^="someone@"]', 'ami');
    await until(() => js(`document.querySelectorAll('.ml-suggest button').length > 0`), 'a suggestion', 3000).catch(() => {});
    if (!(await js(`document.querySelectorAll('.ml-suggest button').length`))) {
      await setField(win, 'input[placeholder^="someone@"]', 'amin');
      await until(() => js(`document.querySelectorAll('.ml-suggest button').length > 0`), 'a suggestion, second try', 4000).catch(() => {});
    }
    const suggested = await js(`[...document.querySelectorAll('.ml-suggest button')].map((b) => b.textContent)`);
    check('mail: Compose completes an address from the people mail has seen', suggested.some((s) => /amina@northwind\.example/.test(s)), JSON.stringify(suggested));
    await js(`[...document.querySelectorAll('.rw-btn, button')].find((b) => /^(Close|Discard|Cancel)$/.test(b.textContent.trim()))?.click(), 'closed'`);
    await until(() => js(`!document.querySelector('input[placeholder^="someone@"]')`), 'the message to close', 4000).catch(() => {});
    await js(`document.querySelector('.ml-keep')?.click(), 'kept'`);
    await until(async () => (await js(`window.rutbaOffice.contacts.count({})`)) >= 1, 'the sender to be kept', 6000).catch(() => {});
    const senders = await js(`(async () => (await window.rutbaOffice.contacts.list({ query: 'amina' })).map((c) => c.emails[0]?.value))()`);
    check('mail: a sender is kept in Contacts with one click', senders.includes('amina@northwind.example'), JSON.stringify(senders));
  } catch (err) {
    check('mail: the invitation checks ran', false, err.message);
  }

  /* ── Mail: a seeded message opens in the reading pane ────────────────── */

  try {
    const win = await open('mail');
    const js = (code) => win.webContents.executeJavaScript(code);
    // The store is read off disk after the window loads; a fixed pause was
    // enough on an idle machine and not under load, so wait for the rows.
    await until(async () => (await js(`document.querySelectorAll('.ml-row').length`)) > 0, 'the seeded messages to list', 8000).catch(() => {});
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

      // The strip is written once the body has been inspected, a moment after
      // the pane paints — a longer moment on a loaded machine.
      await until(async () => (await js(`document.querySelector('.ml-strip')?.textContent ?? ''`)).trim().length > 0, 'the message strip', 5000).catch(() => {});
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
      await until(() => js(`/Sign in with Google/.test(document.querySelector('.rw-dialog .ml-found')?.textContent || '')`), 'the sign-in to be offered', 6000).catch(() => {});
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
      // The message comes back first and the window reopens it a frame later.
      await until(() => js(`Boolean(document.querySelector('.rw-dialog input:not([disabled])')?.value)`), 'the message to reopen', 3000).catch(() => {});
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

  /* ── Mail: adding an account is an address and a password ────────────── */
  //
  // The dialog finds the server from the address — a check run asks only the
  // provider table, never the network — offers the browser sign-in the
  // provider wants, and keeps the advanced fields one click away, filled
  // with what was found.
  try {
    const win = await open('mail');
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean([...document.querySelectorAll('button')].find((b) => /Add account/.test(b.textContent)))`), 'an Add account button', 8000);
    await js(`[...document.querySelectorAll('button')].find((b) => /Add account/.test(b.textContent)).click(), 'clicked'`);
    await until(() => js(`Boolean(document.querySelector('.ml-form input[placeholder="you@example.com"]'))`), 'the account dialog', 5000);
    await setField(win, '.ml-form input[placeholder="you@example.com"]', 'someone@gmail.com');
    await until(() => js(`document.querySelector('.ml-search')?.dataset.state === 'found'`), 'the search to find Google', 6000).catch(() => {});
    const found = await js(`(() => { const s = document.querySelector('.ml-search'); return { state: s?.dataset.state, text: (s?.textContent || '').slice(0, 160), signIn: Boolean(document.querySelector('.ml-found-item')), advanced: Boolean(document.querySelector('.ml-advanced')) }; })()`);
    check('mail: an address alone finds the mail server, and the provider\'s sign-in is offered', found.state === 'found' && /imap\.gmail\.com/.test(found.text) && /smtp\.gmail\.com/.test(found.text) && found.signIn && !found.advanced, JSON.stringify(found));
    await js(`[...document.querySelectorAll('.rw-btn, button')].find((b) => /^Advanced/.test(b.textContent.trim()))?.click(), 'advanced'`);
    await until(() => js(`Boolean(document.querySelector('.ml-advanced'))`), 'the advanced fields', 3000).catch(() => {});
    const advanced = await js(`(() => { const a = document.querySelector('.ml-advanced'); if (!a) return null; return { imap: a.querySelector('input[placeholder="imap.example.com"]')?.value, smtp: a.querySelector('input[placeholder="smtp.example.com"]')?.value, selects: a.querySelectorAll('select').length }; })()`);
    check('mail: Advanced is always there, and carries what was found', advanced && advanced.imap === 'imap.gmail.com' && advanced.smtp === 'smtp.gmail.com' && advanced.selects === 2, JSON.stringify(advanced));
    await setField(win, '.ml-form input[placeholder="you@example.com"]', 'someone@nowhere.example');
    await until(() => js(`document.querySelector('.ml-search')?.dataset.state === 'nothing'`), 'an unknown domain to be said so', 6000).catch(() => {});
    const unknown = await js(`(() => { const s = document.querySelector('.ml-search'); return { state: s?.dataset.state, text: (s?.textContent || '').slice(0, 160) }; })()`);
    check('mail: a domain nothing knows says so, and leaves the fields to the person', unknown.state === 'nothing' && /nowhere\.example/.test(unknown.text), JSON.stringify(unknown));
    const fromFile = await js(`Boolean([...document.querySelectorAll('.rw-dialog .rw-btn')].find((b) => /^From a file/.test(b.textContent.trim())))`);
    await js(`[...document.querySelectorAll('.rw-btn, button')].find((b) => b.textContent.trim() === 'Cancel')?.click(), 'closed'`);
    // The file itself, through the service: two set up as the file says,
    // one left out for having no password, and none tried on the network.
    const imported = await js(`window.rutbaOffice.mail.importAccounts({ path: ${JSON.stringify(files.accounts)} })`);
    const listed = await js(`(async () => (await window.rutbaOffice.mail.accounts()).filter((a) => /checks\\.example$/.test(a.email)).map((a) => ({ email: a.email, name: a.name, imap: a.imap.host + ':' + a.imap.port + (a.imap.secure ? ' tls' : ' starttls'), smtp: a.smtp.host + ':' + a.smtp.port + (a.smtp.secure ? ' tls' : ' starttls'), hasPassword: a.hasPassword })))()`);
    check('mail: a file of accounts sets them up at once, ports deciding the security, and the dialog offers it', fromFile && imported.added.length === 2 && imported.skipped.length === 1 && imported.skipped[0].reason === 'no password' && listed.length === 2 && listed.every((a) => a.hasPassword) && listed[0].imap === 'mail.checks.example:993 tls' && listed[0].smtp === 'mail.checks.example:587 starttls' && listed[1].name === 'Two' && listed[1].imap === 'imap.checks.example:143 starttls' && listed[1].smtp === 'smtp.checks.example:465 tls', `${JSON.stringify(imported)}; ${JSON.stringify(listed)}; button ${fromFile}`);
    const again = await js(`window.rutbaOffice.mail.importAccounts({ path: ${JSON.stringify(files.accounts)} })`);
    check('mail: the same file again sets up nothing twice', again.added.length === 0 && again.skipped.filter((x) => x.reason === 'already set up').length === 2, JSON.stringify(again));
    for (const a of imported.added) await js(`window.rutbaOffice.mail.removeAccount({ id: ${JSON.stringify(a.id)} })`);
  } catch (err) {
    check('mail: the account dialog checks ran', false, err.message);
  }
  await mailProviders();
  await mailSignature();
  await mailSendLater();
  await mailOOO();

  /* ── OpenDocument goes out as OpenDocument ───────────────────────────── */
  try {
    const odt = path.join(path.dirname(files.docx), 'report.odt');
    const s = doc.open({ path: files.docx });
    const saved = doc.save({ id: s.id, path: odt });
    doc.close({ id: s.id });
    const back = doc.open({ path: odt });
    check('a document saved as .odt is an OpenDocument the suite opens as one, and writes back', saved.format === 'odt' && back.converted?.from === 'odt' && back.converted.writesBack === true && (back.model?.blocks?.length ?? 0) > 0, `${saved.format}; opened from ${back.converted?.from}, writesBack ${back.converted?.writesBack}, ${back.model?.blocks?.length ?? 0} blocks`);
    doc.close({ id: back.id });
  } catch (err) {
    check('the OpenDocument check ran', false, err.message);
  }

  /* ── Contacts: a file is shown and offered, a card is kept and found ─── */
  try {
    const win = await open('contacts', files.vcf);
    const js = (code) => win.webContents.executeJavaScript(code);
    const before = await js('window.rutbaOffice.contacts.count({})');
    await until(() => js(`document.querySelectorAll('.ct-item').length >= 2`), 'the file\'s cards to list', 8000).catch(() => {});
    const shown = await js(`(() => ({ items: document.querySelectorAll('.ct-item').length, names: [...document.querySelectorAll('.ct-item .name')].map((n) => n.textContent), offer: Boolean(document.querySelector('.ct-offer')), card: document.querySelector('.ct-card h2')?.textContent || '' }))()`);
    check('contacts: a .vcf opens as its cards, shown and offered, not yet kept', shown.items === 2 && /Kim Lee/.test(shown.names.join(',')) && shown.offer && /Kim Lee|Sam Patel/.test(shown.card), JSON.stringify(shown));
    await js(`[...document.querySelectorAll('.ct-offer .rw-btn')].pop()?.click(), 'added'`);
    await until(() => js(`!document.querySelector('.ct-offer') && document.querySelectorAll('.ct-item').length >= 2`), 'the cards to be kept', 8000).catch(() => {});
    // Relative to what the book held: the mail block has already kept a sender.
    const count = await js(`window.rutbaOffice.contacts.count({})`);
    check('contacts: Add keeps the file\'s cards in the address book', count === before + 2 && (await js(`!document.querySelector('.ct-offer')`)), `${count} kept, ${before} before`);
    await js(`[...document.querySelectorAll('.rw-btn')].find((b) => /New contact/.test(b.textContent))?.click(), 'new'`);
    await until(() => js(`Boolean(document.querySelector('.ct-editor input.ct-name'))`), 'the editor', 4000).catch(() => {});
    await setField(win, '.ct-editor input.ct-name', 'Alex Morgan');
    await setField(win, '.ct-editor .ct-line input[placeholder="name@example.com"]', 'alex@example.net');
    await js(`[...document.querySelectorAll('.ct-actions .rw-btn')].find((b) => b.textContent.trim() === 'Save')?.click(), 'saved'`);
    await until(() => js(`document.querySelector('.ct-card h2')?.textContent === 'Alex Morgan'`), 'the new card to be shown', 6000).catch(() => {});
    const after = await js(`(() => ({ items: document.querySelectorAll('.ct-item').length, card: document.querySelector('.ct-card h2')?.textContent || '', email: document.querySelector('.ct-card .ct-link')?.textContent || '' }))()`);
    check('contacts: a new card typed in is kept and shown', after.items === count + 1 && after.card === 'Alex Morgan' && after.email === 'alex@example.net', `${JSON.stringify(after)}; ${count} before`);
    const found = await js(`(async () => { const s = await window.rutbaOffice.contacts.suggest({ query: 'ki' }); return s.map((x) => x.email); })()`);
    check('contacts: Compose can complete an address from the book', found.includes('kim@example.com') && !found.includes('sam@example.org'), JSON.stringify(found));
  } catch (err) {
    check('contacts: the checks ran', false, err.message);
  }

  /* ── Calendar: a file is shown beside the person's own, kept, and added to ─ */
  try {
    const win = await open('calendar', files.ics);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.cal-chip').length >= 2`), 'the file\'s events on the month', 8000).catch(() => {});
    const shown = await js(`(() => ({ chips: [...document.querySelectorAll('.cal-chip .s')].map((n) => n.textContent), banner: document.querySelector('.cal-invite')?.textContent || '', today: Boolean(document.querySelector('.cal-day.today')) }))()`);
    check('calendar: a .ics opens on the month it belongs to, shown beside your own and offered', shown.chips.includes('Pricing review') && shown.chips.includes('Bank holiday') && /not yet kept/.test(shown.banner) && shown.today, JSON.stringify(shown));
    await js(`[...document.querySelectorAll('.cal-invite .rw-btn')].pop()?.click(), 'added'`);
    await until(() => js(`!document.querySelector('.cal-invite')`), 'the events to be kept', 8000).catch(() => {});
    const kept = await js(`(async () => { const c = await window.rutbaOffice.calendar.calendars({}); return c.map((x) => x.count); })()`);
    check('calendar: Add to my calendar keeps the file\'s events', kept.some((n) => n >= 2) && (await js(`document.querySelectorAll('.cal-chip').length >= 2`)), `counts ${JSON.stringify(kept)}`);
    await js(`[...document.querySelectorAll('.rw-btn')].find((b) => /New event/.test(b.textContent))?.click(), 'new'`);
    await until(() => js(`Boolean(document.querySelector('.rw-dialog .cal-form'))`), 'the event dialog', 4000).catch(() => {});
    await setField(win, '.rw-dialog input.cal-title', 'Dentist');
    await js(`[...document.querySelectorAll('.rw-dialog .rw-btn')].find((b) => b.textContent.trim() === 'Save')?.click(), 'saved'`);
    await until(() => js(`[...document.querySelectorAll('.cal-chip .s')].some((n) => n.textContent === 'Dentist')`), 'the new event on the month', 6000).catch(() => {});
    const chips = await js(`[...document.querySelectorAll('.cal-chip .s')].map((n) => n.textContent)`);
    check('calendar: a new event typed in is kept and drawn', chips.includes('Dentist') && (await js(`!document.querySelector('.rw-dialog')`)), JSON.stringify(chips));
    await js(`[...document.querySelectorAll('.cal-view')].find((b) => b.textContent === 'Week')?.click(), 'week'`);
    await until(() => js(`document.querySelectorAll('.cal-block').length >= 1`), 'the week view', 4000).catch(() => {});
    const week = await js(`(() => ({ blocks: [...document.querySelectorAll('.cal-block .s')].map((n) => n.textContent), allDay: [...document.querySelectorAll('.cal-allday-cell .cal-chip .s')].map((n) => n.textContent) }))()`);
    check('calendar: the week view places timed events on the grid and all-day ones in the band', week.blocks.includes('Pricing review') && week.allDay.includes('Bank holiday'), JSON.stringify(week));
  } catch (err) {
    check('calendar: the checks ran', false, err.message);
  }

  /* ── Journeys: the things a person does between the buttons ──────────── */
  //
  // Escape leaves a dialog with nothing changed; a window at its smallest
  // size still shows everything without a sideways scroll; the launcher
  // lists what was opened; Ctrl+F1 folds the ribbon and unfolds it.

  try {
    const win = await open('sheets', files.xlsx);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.sh-cell.active'))`), 'an active cell', 6000);
    const before = await js(`document.querySelector('.sh-cell.active')?.dataset.ref || null`);
    await press(win.webContents, 'G', { modifiers: ['control'] });
    await until(() => js(`Boolean(document.querySelector('.rw-dialog input'))`), 'the Go To dialog', 4000).catch(() => {});
    const opened = await js(`Boolean(document.querySelector('.rw-dialog input'))`);
    await js(`(() => { const i = document.querySelector('.rw-dialog input'); if (i) { i.focus(); i.value = 'Z99'; i.dispatchEvent(new Event('input', { bubbles: true })); } return 'typed'; })()`);
    await press(win.webContents, 'Escape');
    await until(() => js(`!document.querySelector('.rw-dialog')`), 'the dialog to close', 4000).catch(() => {});
    const after = await js(`({ dialog: Boolean(document.querySelector('.rw-dialog')), ref: document.querySelector('.sh-cell.active')?.dataset.ref || null })`);
    check('journeys: Escape closes the Go To dialog and moves nothing', opened && !after.dialog && after.ref === before, `dialog ${opened} then ${after.dialog}; selection ${before} then ${after.ref}`);

    await win.setSize(720, 520);
    await wait(500);
    const small = await js(`(() => {
      const d = document.documentElement;
      const status = document.querySelector('.rw-status');
      const tabs = document.querySelector('.rw-tabs');
      return { w: innerWidth, h: innerHeight, sideways: d.scrollWidth - d.clientWidth, status: status ? status.getBoundingClientRect().bottom <= innerHeight + 1 : false, tabs: tabs ? tabs.scrollWidth >= tabs.clientWidth : false };
    })()`);
    check('journeys: a window at its minimum size has no sideways scroll and keeps its status bar', small.sideways <= 1 && small.status && small.tabs, JSON.stringify(small));
    await win.setSize(1280, 860);
  } catch (err) {
    check('journeys: the sheets journey ran', false, err.message);
  }

  try {
    const win = await open('slides', files.pptx);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.sl-svg'))`), 'the slide', 6000);
    await clickTab(win, 'Review');
    await wait(150);
    await clickIn(win, 'Speaker Notes');
    await until(() => js(`Boolean(document.querySelector('.rw-dialog textarea'))`), 'the notes dialog', 4000).catch(() => {});
    const opened = await js(`Boolean(document.querySelector('.rw-dialog textarea'))`);
    await setField(win, '.rw-dialog textarea', 'Not to be kept.');
    await press(win.webContents, 'Escape');
    await until(() => js(`!document.querySelector('.rw-dialog')`), 'the dialog to close', 4000).catch(() => {});
    const closed = await js(`!document.querySelector('.rw-dialog')`);
    const notes = await win.webContents.executeJavaScript(`(async () => { const all = await window.rutbaOffice.doc.sessions({}); const mine = all.filter((s) => s.kind === 'deck').pop(); const m = await window.rutbaOffice.doc.model({ id: mine.id, slide: 0, width: 640 }); return m.slide?.notes || ''; })()`);
    check('journeys: Escape closes the Speaker Notes dialog and keeps the notes as they were', opened && closed && !/Not to be kept/.test(notes), `dialog ${opened}, closed ${closed}; notes ${JSON.stringify(notes.slice(0, 40))}`);

    await win.setSize(720, 520);
    await wait(500);
    const small = await js(`(() => { const d = document.documentElement; return { sideways: d.scrollWidth - d.clientWidth, slide: Boolean(document.querySelector('.sl-svg')), status: Boolean(document.querySelector('.rw-status')) }; })()`);
    check('journeys: the presentation window at its minimum size still shows the slide', small.sideways <= 1 && small.slide && small.status, JSON.stringify(small));
    await win.setSize(1280, 860);
  } catch (err) {
    check('journeys: the slides journey ran', false, err.message);
  }

  try {
    const win = await open('word', files.docx);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.wd-block'))`), 'the page', 6000);
    await press(win.webContents, 'F1', { modifiers: ['control'] });
    await until(() => js(`Boolean(document.querySelector('.rw-ribbon.collapsed'))`), 'the ribbon to fold', 3000).catch(() => {});
    const folded = await js(`({ collapsed: Boolean(document.querySelector('.rw-ribbon.collapsed')), groups: document.querySelectorAll('.rw-groups').length })`);
    await press(win.webContents, 'F1', { modifiers: ['control'] });
    await until(() => js(`!document.querySelector('.rw-ribbon.collapsed')`), 'the ribbon to unfold', 3000).catch(() => {});
    const unfolded = await js(`({ collapsed: Boolean(document.querySelector('.rw-ribbon.collapsed')), groups: document.querySelectorAll('.rw-groups').length })`);
    check('journeys: Ctrl+F1 folds the ribbon and folds it back', folded.collapsed && folded.groups === 0 && !unfolded.collapsed && unfolded.groups === 1, `folded ${JSON.stringify(folded)}; unfolded ${JSON.stringify(unfolded)}`);

    await win.setSize(720, 520);
    await wait(500);
    const small = await js(`(() => { const d = document.documentElement; const page = document.querySelector('.wd-page'); return { sideways: d.scrollWidth - d.clientWidth, page: Boolean(page), status: Boolean(document.querySelector('.rw-status')) }; })()`);
    check('journeys: the Word window at its minimum size still shows the page without a sideways scroll', small.sideways <= 1 && small.page && small.status, JSON.stringify(small));
    await win.setSize(1280, 860);
  } catch (err) {
    check('journeys: the Word journey ran', false, err.message);
  }

  try {
    const win = await open('home');
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.home-recent-row').length >= 3`), 'the recent list', 8000).catch(() => {});
    const recent = await js(`[...document.querySelectorAll('.home-recent-row')].map((r) => r.textContent.trim().slice(0, 60))`);
    const names = recent.join(' | ');
    check('journeys: the launcher lists the files this run opened', recent.length >= 3 && /report\.docx/.test(names) && /sales\.xlsx/.test(names) && /deck\.pptx/.test(names), `${recent.length} rows: ${names.slice(0, 200)}`);
  } catch (err) {
    check('journeys: the launcher journey ran', false, err.message);
  }

  /* ── A presenter window ends with its editor ─────────────────────────── */
  //
  // The presenter shows the deck its editor holds open. Closing the editor
  // used to leave the presenter up, failing every call it made.
  try {
    const editor = await open('slides', files.pptx);
    const deckId = await editor.webContents.executeJavaScript(`(async () => { const all = await window.rutbaOffice.doc.sessions({}); return all.filter((s) => s.kind === 'deck').pop().id; })()`);
    const presenter = windows.create({ app: 'slides', query: { presenter: deckId } });
    opened.push(presenter);
    await new Promise((resolve) => presenter.webContents.once('did-finish-load', () => setTimeout(resolve, 800)));
    closingPhase.value = true;
    editor.close();
    await until(() => presenter.isDestroyed(), 'the presenter window to close with its editor', 5000).catch(() => {});
    closingPhase.value = false;
    check('journeys: closing the editor closes the presenter window on the same deck', presenter.isDestroyed(), presenter.isDestroyed() ? 'closed together' : 'the presenter stayed open on a deck that is no longer there');
  } catch (err) {
    closingPhase.value = false;
    check('journeys: the presenter-close check ran', false, err.message);
  }

  /* ── Real input: the mouse and the keyboard, not element clicks ──────── */
  //
  // Every other check clicks elements and calls dispatch. The owner's report
  // was about what a person does: click a cell and it is not selected, press
  // an arrow and nothing moves, press a shortcut and nothing happens. So
  // these drive the windows with the same events a mouse and a keyboard
  // send, and read the engine to see where they landed.

  const mouse = async (win, selector, { at = 'centre', modifiers = [] } = {}) => {
    // A cell is named by its reference ("D5"): an empty cell has no element
    // of its own, so the point comes from the column and row headers, the
    // way an eye finds it. Anything else is an element to hit in the middle.
    const box = await win.webContents.executeJavaScript(`(() => {
      const sel = ${JSON.stringify(selector)};
      const cell = /^cell:([A-Z]+)([0-9]+)$/.exec(sel);
      if (cell) {
        const col = [...document.querySelectorAll('.sh-colheads .sh-head')].find((h) => h.textContent.trim() === cell[1]);
        const row = [...document.querySelectorAll('.sh-rowheads .sh-head')].find((h) => h.textContent.trim() === cell[2]);
        if (!col || !row) return null;
        const c = col.getBoundingClientRect(); const r = row.getBoundingClientRect();
        return { x: c.left + c.width / 2, y: r.top + r.height / 2, w: c.width, h: r.height, under: (document.elementFromPoint(c.left + c.width / 2, r.top + r.height / 2) || {}).className || null };
      }
      const el = document.querySelector(sel); if (!el) return null; const r = el.getBoundingClientRect();
      return { x: r.left + r.width / 2, y: r.top + r.height / 2, w: r.width, h: r.height };
    })()`);
    if (!box) return null;

    const x = Math.round(at === 'left' ? box.x - box.w / 4 : box.x);
    const y = Math.round(box.y);
    win.webContents.sendInputEvent({ type: 'mouseMove', x, y, modifiers });
    win.webContents.sendInputEvent({ type: 'mouseDown', x, y, button: 'left', clickCount: 1, modifiers });
    win.webContents.sendInputEvent({ type: 'mouseUp', x, y, button: 'left', clickCount: 1, modifiers });
    await wait(120);
    return box;
  };

  try {
    const win = await open('sheets', files.xlsx);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.sh-cell.active'))`), 'an active cell', 6000);
    const active = () => js(`document.querySelector('.sh-cell.active')?.dataset.ref || null`);
    const focused = () => js(`(document.activeElement && (document.activeElement.className || document.activeElement.tagName)) || 'none'`);
    const session = () => win.webContents.executeJavaScript(`(async () => { const all = await window.rutbaOffice.doc.sessions({}); return all.filter((s) => s.kind === 'sheet').pop().id; })()`);
    const cellText = async (ref) => { const id = await session(); return win.webContents.executeJavaScript(`(async () => { const m = await window.rutbaOffice.doc.model({ id: ${JSON.stringify(id)} }); const c = (m.cells || []).find((x) => x.ref === ${JSON.stringify(ref)}); return c ? String(c.text ?? '') : ''; })()`); };

    // A real click on D5 selects D5.
    const clicked = await mouse(win, 'cell:D5');
    await until(async () => (await active()) === 'D5', 'the click to select D5', 3000).catch(() => {});
    const afterClick = await active();
    // Arrow keys move the selection: Right to E5, Down to E6.
    await press(win.webContents, 'Right');
    await press(win.webContents, 'Down');
    await until(async () => (await active()) === 'E6', 'the arrows to reach E6', 3000).catch(() => {});
    const afterArrows = await active();
    // Typing lands in E6 and Enter commits it and moves down.
    await typeText(win.webContents, '4321');
    await press(win.webContents, 'Return');
    await until(async () => (await cellText('E6')) === '4321', 'the value to land in E6', 4000).catch(() => {});
    const landed = await cellText('E6');
    const afterEnter = await active();
    check('real input: a click selects the cell, the arrows move it, typing lands, Enter moves down', Boolean(clicked) && afterClick === 'D5' && afterArrows === 'E6' && landed === '4321' && afterEnter === 'E7', `click → ${afterClick}; arrows → ${afterArrows}; E6 = ${JSON.stringify(landed)}; after Enter ${afterEnter}; focus ${await focused()}`);

    // A real click on the ribbon's Bold, then typing: the keys must still go
    // to the grid. A button that kept focus would swallow them.
    await mouse(win, '.rw-btn[data-tip^="Bold"]');
    await typeText(win.webContents, '77');
    await press(win.webContents, 'Return');
    await until(async () => (await cellText('E7')) === '77', 'the value typed after the ribbon click to land', 4000).catch(() => {});
    const afterRibbon = await cellText('E7');
    check('real input: typing after a ribbon click still goes to the grid', afterRibbon === '77', `E7 = ${JSON.stringify(afterRibbon)}; focus ${await focused()}`);

    // The formula bar: click it, type, Enter — the value lands and the keys
    // go back to the grid, as they do in Excel.
    await mouse(win, 'cell:B9');
    await until(async () => (await active()) === 'B9', 'B9 to be selected', 3000).catch(() => {});
    const bar = await mouse(win, '.sh-formula input, .sh-formula-input');
    if (bar) {
      // Where the keys are going before they are sent, and what each of the
      // three places that can hold the text says afterwards. This check
      // failed once with B9 reading "ello" and there was no way to tell from
      // the line which of them had eaten the letter.
      // Nothing may still be open from the check before this one. A cell
      // editor left standing takes the first letter and commits it where it
      // was, and the rest lands here: B9 read "ello" twice before this wait,
      // and the state below is what said so.
      await until(() => win.webContents.executeJavaScript(`!document.querySelector('.sh-editor')`), 'the previous edit to have finished', 3000).catch(() => {});
      const before = await win.webContents.executeJavaScript(`(() => {
        const b = document.querySelector('.sh-formula input');
        const a = document.activeElement;
        return { focus: a ? (a.className || a.tagName) : 'none', bar: b ? b.value : null, editor: document.querySelector('.sh-editor')?.value ?? '(none)', cell: document.querySelector('.sh-cell.active')?.dataset.ref ?? null };
      })()`);
      await typeText(win.webContents, 'hello');
      await press(win.webContents, 'Return');
      await until(async () => (await cellText('B9')) === 'hello', 'the formula bar entry to land', 4000).catch(() => {});
      await press(win.webContents, 'Down');
      await wait(150);
      check(
        'real input: Enter in the formula bar commits and hands the keys back to the grid',
        (await cellText('B9')) === 'hello' && (await active()) === 'B11',
        `B9 = ${JSON.stringify(await cellText('B9'))}; after Enter, Down: ${await active()}; focus ${await focused()}; before typing ${JSON.stringify(before)}`
      );
    } else {
      check('real input: the formula bar is there to click', false, 'no formula bar input found');
    }

    // Ctrl+click adds a second rectangle; Bold from the ribbon paints both;
    // a plain click drops the extra one. The owner: "multiselect in excel is
    // not possible".
    const namebox = () => js(`document.querySelector('.sh-namebox')?.textContent || ''`);
    const weightOf = (ref) => js(`(() => { const c = document.querySelector('.sh-cell[data-ref="${ref}"]'); return c ? getComputedStyle(c).fontWeight : null; })()`);
    await mouse(win, 'cell:D5');
    await until(async () => (await active()) === 'D5', 'D5 to be selected', 3000).catch(() => {});
    await mouse(win, 'cell:F7', { modifiers: ['control'] });
    await until(async () => (await namebox()) === 'D5,F7', 'the name box to show both rectangles', 3000).catch(() => {});
    const both = await namebox();
    await mouse(win, '.rw-btn[data-tip^="Bold"]');
    await until(async () => (await weightOf('D5')) === '700' && (await weightOf('F7')) === '700', 'Bold to reach both cells', 4000).catch(() => {});
    const weights = { D5: await weightOf('D5'), F7: await weightOf('F7') };
    await mouse(win, 'cell:B2');
    await until(async () => (await namebox()) === 'B2', 'a plain click to drop the extra rectangle', 3000).catch(() => {});
    check('real input: Ctrl+click adds a second selection, Bold paints both, a plain click drops it', both === 'D5,F7' && weights.D5 === '700' && weights.F7 === '700' && (await namebox()) === 'B2', `name box ${JSON.stringify(both)} then ${JSON.stringify(await namebox())}; weights ${JSON.stringify(weights)}`);
  } catch (err) {
    check('real input: the Worksheets journey ran', false, err.message);
  }

  try {
    const win = await open('word', files.docx);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.wd-block'))`), 'the page', 6000);
    const session = () => win.webContents.executeJavaScript(`(async () => { const all = await window.rutbaOffice.doc.sessions({}); return all.filter((s) => s.kind === 'doc').pop().id; })()`);
    const block = async (i) => { const id = await session(); return win.webContents.executeJavaScript(`(async () => { const m = await window.rutbaOffice.doc.model({ id: ${JSON.stringify(id)} }); const b = m.blocks[${i}]; return { text: b?.text || '', bold: (b?.runs || []).some((r) => r.bold), italic: (b?.runs || []).some((r) => r.italic) }; })()`); };
    const before = await block(0);
    // A real click at the end of the first paragraph, then typing.
    await mouse(win, '[data-block="0"]', { at: 'right' });
    await js(`(() => { const p = document.querySelector('[data-block="0"]'); const r = document.createRange(); r.selectNodeContents(p); r.collapse(false); const s = getSelection(); s.removeAllRanges(); s.addRange(r); return 'caret'; })()`);
    await typeText(win.webContents, ' Typed.');
    await until(async () => /Typed\./.test((await block(0)).text), 'the typing to reach the engine', 4000).catch(() => {});
    const typed = await block(0);
    check('real input: a click into the page and typing reach the engine', typed.text.length > before.text.length && /Typed\./.test(typed.text), `text ${JSON.stringify(typed.text.slice(-24))}`);

    // Ctrl+B by keyboard on a selection, then a real click on Italic.
    await js(`(() => { const p = document.querySelector('[data-block="0"]'); const r = document.createRange(); r.selectNodeContents(p); const s = getSelection(); s.removeAllRanges(); s.addRange(r); return 'selected'; })()`);
    await press(win.webContents, 'B', { modifiers: ['control'] });
    await until(async () => (await block(0)).bold, 'Ctrl+B to reach the engine', 4000).catch(() => {});
    const bolded = await block(0);
    // The selection a person makes: a real drag across the second paragraph,
    // then a real click on the ribbon's Italic. The button must not take the
    // selection with it.
    const drag = await js(`(() => { const p = document.querySelector('[data-block="1"]'); if (!p) return null; const r = p.getBoundingClientRect(); return { x1: r.left + 4, y1: r.top + r.height / 2, x2: r.left + Math.min(r.width - 4, 260), y2: r.top + r.height / 2 }; })()`);
    if (drag) {
      const wc = win.webContents;
      wc.sendInputEvent({ type: 'mouseMove', x: Math.round(drag.x1), y: Math.round(drag.y1) });
      wc.sendInputEvent({ type: 'mouseDown', x: Math.round(drag.x1), y: Math.round(drag.y1), button: 'left', clickCount: 1 });
      for (let i = 1; i <= 6; i++) wc.sendInputEvent({ type: 'mouseMove', x: Math.round(drag.x1 + ((drag.x2 - drag.x1) * i) / 6), y: Math.round(drag.y1), button: 'left', buttons: 1 });
      wc.sendInputEvent({ type: 'mouseUp', x: Math.round(drag.x2), y: Math.round(drag.y2), button: 'left', clickCount: 1 });
      await wait(250);
    }
    const selectedBefore = await js(`(() => { const s = getSelection(); return { collapsed: s.isCollapsed, text: s.toString().slice(0, 30) }; })()`);
    await mouse(win, '.rw-btn[data-tip^="Italic"]');
    const selectedAfter = await js(`(() => { const s = getSelection(); return { collapsed: s.isCollapsed, text: s.toString().slice(0, 30), focus: (document.activeElement && (document.activeElement.className || document.activeElement.tagName)) || 'none' }; })()`);
    await until(async () => (await block(1)).italic, 'the Italic button to reach the engine', 4000).catch(() => {});
    const italic = await block(1);
    check('real input: Ctrl+B on a selection, and a real drag then a real click on Italic, both reach the engine', bolded.bold && italic.italic, `bold ${bolded.bold}; italic ${italic.italic}; drag ${JSON.stringify(drag)}; selection before click ${JSON.stringify(selectedBefore)}, after ${JSON.stringify(selectedAfter)}`);
  } catch (err) {
    check('real input: the Word journey ran', false, err.message);
  }

  /* ── Autosave: a crash does not take the work ────────────────────────── */
  //
  // The copies are written on a timer while a document is open and deleted the
  // moment it is saved or closed, so what is left in the profile is what a
  // crash took. This types into a document, asks the service to write its
  // copy, and then opens the launcher — where the person who lost it would be
  // looking — to see that it is offered back by name.
  try {
    const win = await open('word', null);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block]'))`), 'the page', 8000);
    await js(`(() => {
      const page = document.querySelector('.wd-page');
      const block = page.querySelector('[data-block="0"]');
      page.focus();
      const range = document.createRange();
      range.selectNodeContents(block);
      range.collapse(false);
      const sel = getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      page.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
      return 'placed';
    })()`);
    await wait(200);
    // The page takes keys once it is editable and has the focus; a loaded
    // machine gets there later than a fixed pause allows.
    await until(() => js(`(() => { const p = document.querySelector('.wd-page'); return Boolean(p && p.isContentEditable && document.activeElement === p); })()`), 'the page to be editable and focused', 6000).catch(() => {});
    raise(win);
    win.webContents.insertText('UNSAVED WORK');
    await until(() => js(`/UNSAVED WORK/.test(document.querySelector('.wd-page')?.textContent || '')`), 'the text to land', 8000);

    const written = doc.autosave();
    const offered = doc.recoverable();
    check(
      'autosave: unsaved work is copied where a crash cannot take it',
      (written.length >= 1 || offered.some((e) => e.kind === 'doc' && Date.now() - e.at < 60000)) && offered.some((e) => e.kind === 'doc'),
      `${written.length} written now, ${offered.filter((e) => e.kind === 'doc' && Date.now() - e.at < 60000).length} copied within the minute, ${offered.length} recoverable`
    );

    const launcher = await open('home');
    await until(() => launcher.webContents.executeJavaScript(`document.querySelectorAll('.home-recovered-row').length > 0`), 'the recovered strip', 6000).catch(() => {});
    const strip = await launcher.webContents.executeJavaScript(`(() => {
      const rows = [...document.querySelectorAll('.home-recovered-row')];
      return { rows: rows.length, text: rows.map((r) => r.textContent).join(' | ').slice(0, 160), buttons: rows[0] ? [...rows[0].querySelectorAll('button')].map((b) => b.textContent.trim()) : [] };
    })()`);
    check(
      'autosave: the launcher offers the recovered document back',
      strip.rows >= 1 && /unsaved work from/.test(strip.text) && strip.buttons.includes('Recover') && strip.buttons.includes('Discard'),
      JSON.stringify(strip).slice(0, 200)
    );

    // And it is recoverable as the document it was, with the words in it.
    //
    // The newest one, not the first: by this point in the run a dozen windows
    // are open and unsaved, so the list holds every one of them, and the first
    // check that read it recovered somebody else's document and failed on
    // words it had never typed.
    // The newest is not always ours either — the timer copies whichever
    // dirty window changed last — so every document copy is read and the
    // one holding the typed words is the one this check is about.
    const candidates = offered.filter((e) => e.kind === 'doc').sort((a, b) => b.at - a.at);
    let entry = candidates[0];
    let recovered = null;
    let text = '';
    for (const candidate of candidates) {
      const opened = doc.recover({ file: candidate.file });
      const words = (opened.model.blocks || []).map((b) => (b.runs || []).map((r) => r.text).join('')).join(' ');
      if (words.includes('UNSAVED WORK') || !recovered) {
        if (recovered) doc.close({ id: recovered.id });
        entry = candidate;
        recovered = opened;
        text = words;
        if (words.includes('UNSAVED WORK')) break;
      } else doc.close({ id: opened.id });
    }
    check('autosave: what comes back is the work that was lost', text.includes('UNSAVED WORK'), `recovered ${JSON.stringify(text.slice(0, 60))}`);
    doc.close({ id: recovered.id });
    doc.discardRecovery({ file: entry.file });
  } catch (err) {
    check('autosave: the recovery checks ran', false, err.message);
  }

  /* ── Printing: paper, and a PDF of anything ──────────────────────────── */
  //
  // A suite that cannot print is not an office suite. Until this existed the
  // Print button in Word called a command that was never defined, and the one
  // in Worksheets asked for a PDF that was refused. Each window opens the
  // dialog with Ctrl+P, is asked how many pages that would be, and then writes
  // the PDF through the same door the dialog's own button uses — which is what
  // proves the layout reached Chromium and came back as pages.
  for (const [appName, file, kind, options] of [
    ['word', files.docx, 'doc', {}],
    ['sheets', files.xlsx, 'sheet', { headings: true, gridlines: true }],
    ['slides', files.pptx, 'deck', { layout: 'handout', perPage: 4 }],
  ]) {
    try {
      const win = await open(appName, file);
      const js = (code) => win.webContents.executeJavaScript(code);
      const session = () => js(`(async () => { const all = await window.rutbaOffice.doc.sessions({}); return all.filter((s) => s.kind === ${JSON.stringify(kind)}).pop().id; })()`);
      const id = await session();

      // The dialog says how many pages before it prints anything.
      raise(win);
      await press(win.webContents, 'p', { modifiers: ['control'] });
      await until(() => js(`Boolean(document.querySelector('.rw-dialog'))`), 'the print dialog', 4000).catch(() => {});
      // The count arrives when the plan does — later for a sheet than for a document.
      // A busy main process answers late; twenty seconds is the most a person would wait.
      await until(() => js(`/[0-9]+ page/.test(document.querySelector('.rw-dialog')?.textContent || '')`), 'the page count', 20000).catch(() => {});
      const dialog = await js(`(() => {
        const d = document.querySelector('.rw-dialog');
        if (!d) return null;
        return { title: d.querySelector('.rw-dialog-head')?.textContent || '', text: d.textContent.slice(0, 1500), chips: [...d.querySelectorAll('.chip')].map((c) => c.textContent) };
      })()`);
      // The count is in the summary chip; a workbook's dialog is long enough that a slice of its text would miss it.
      const counted = Boolean(dialog) && (dialog.chips.some((c) => /[0-9]+ page/.test(c)) || /[0-9]+ page/.test(dialog.text));
      const focusedBefore = await js(`(document.activeElement && (document.activeElement.className || document.activeElement.tagName)) || 'none'`);
      await press(win.webContents, 'Escape');
      await until(() => js(`!document.querySelector('.rw-dialog')`), 'the dialog to close', 3000).catch(() => {});
      const closed = await js(`!document.querySelector('.rw-dialog')`);
      check(
        `${appName}: Ctrl+P opens a print dialog that says how many pages, and Escape closes it`,
        Boolean(dialog) && dialog.title === 'Print' && counted && closed,
        dialog
          ? `${JSON.stringify((/(\d+ pages?[^"]*)/.exec(dialog.text) || [])[1] || dialog.chips.join(' | ') || dialog.text.slice(0, 60))}; title ${JSON.stringify(dialog.title)}; ${closed ? 'closed on Escape' : `still open after Escape (focus was on ${focusedBefore})`}`
          : 'no dialog'
      );
      if (!closed) await js(`[...document.querySelectorAll('.rw-dialog button')].find((b) => b.textContent.trim() === 'Cancel')?.click(), 'cancelled'`);

      // And the PDF it would write is a PDF, with the pages it promised.
      const target = path.join(path.dirname(files.docx), `print-${appName}.pdf`);
      const promised = await js(`window.rutbaOffice.print.summary({ id: ${JSON.stringify(id)}, options: ${JSON.stringify(options)} })`);
      await js(`window.rutbaOffice.print.pdf({ id: ${JSON.stringify(id)}, path: ${JSON.stringify(target)}, options: ${JSON.stringify(options)} })`);
      const bytes = fs.readFileSync(target);
      const drawn = (bytes.toString('latin1').match(/\/Type\s*\/Page[^s]/g) || []).length;
      check(
        `${appName}: it writes a PDF with the pages it said it would`,
        bytes.subarray(0, 5).toString('latin1') === '%PDF-' && drawn === promised.pages && drawn > 0,
        `${promised.pages} promised, ${drawn} drawn, ${(bytes.length / 1024).toFixed(0)} KB`
      );
    } catch (err) {
      check(`${appName}: the printing checks ran`, false, err.message);
    }
  }

  /* ── Last of all: a dirty window refuses to close ─────────────────────── */


  //
  // Unsaved work must not close. Type to make a fresh document dirty, ask the
  // window to close, and check that it refused — the prompt is on screen and
  // the document is still there. Closing the X used to throw the work away
  // without a word, which is the worst thing an editor can do.
  //
  // LAST, because the refusal is a native, window-modal message box that
  // nothing can dismiss from here; while it stood mid-run, keystrokes meant
  // for later checks reached it, answered it, and the run lost windows.
  try {
    const win = await open('word', null);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block]'))`), 'the page and its first paragraph', 8000);
    // Put the caret in the first paragraph the way the round-trip check does,
    // then insert text through the same path — a keystroke into a page that
    // has not yet taken focus lands nowhere.
    await js(`(() => {
      const page = document.querySelector('.wd-page');
      const block = page.querySelector('[data-block="0"]');
      page.focus();
      const range = document.createRange();
      range.selectNodeContents(block);
      range.collapse(false);
      const sel = getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      page.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
      return 'caret placed';
    })()`);
    await wait(250);
    win.webContents.insertText('MORE');
    await until(() => js(`/MORE/.test(document.querySelector('.wd-page')?.textContent || '')`), 'the text to land', 4000);
    await wait(300);
    win.close();
    await wait(900);
    check('word: a dirty window refuses to close silently', !win.isDestroyed(), win.isDestroyed() ? 'the window closed and the work went with it' : 'the window is still open, asking');
  } catch (err) {
    check('word: the dirty-close check ran', false, err.message);
  }

  return done();
}
