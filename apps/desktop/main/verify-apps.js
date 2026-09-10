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
import { consoleMessage } from './console-message.js';

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
  pictures: '.pv-image, .pv-pdf, .pv-video, .pv-audio, .pv-empty',
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
  const tomorrow = new Date(day.getFullYear(), day.getMonth(), day.getDate() + 1);
  fs.writeFileSync(at('events.ics'), ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Rutba//checks//EN', 'X-WR-CALNAME:Checks', 'BEGIN:VEVENT', 'UID:check-1@rutba.io', `DTSTART:${stamp(day, 10)}`, `DTEND:${stamp(day, 11)}`, 'SUMMARY:Pricing review', 'LOCATION:Room 4', 'END:VEVENT', 'BEGIN:VEVENT', 'UID:check-2@rutba.io', `DTSTART;VALUE=DATE:${stamp(tomorrow, 0).slice(0, 8)}`, 'SUMMARY:Bank holiday', 'END:VEVENT', 'END:VCALENDAR', ''].join('\r\n'));

  // A real image: the application's own icon, which is a genuine PNG.
  const icon = path.resolve(path.dirname(new URL(import.meta.url).pathname.replace(/^\//, '')), '..', 'resources', 'icon.png');
  if (fs.existsSync(icon)) fs.copyFileSync(icon, at('picture.png'));

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
    xlsx: at('sales.xlsx'),
    pptx: at('deck.pptx'),
    wav: at('tone.wav'),
    vcf: at('contacts.vcf'),
    ics: at('events.ics'),
    md: at('readme.md'),
    png: fs.existsSync(at('picture.png')) ? at('picture.png') : null,
  };
}

export async function verifyApps({ windows, doc }) {
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
    for (const win of opened) if (!win.isDestroyed()) win.destroy();
    fs.rmSync(dir, { recursive: true, force: true });

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
        const flow = [...page.children].filter((el) => /wd-block|wd-table|wd-notes/.test(el.className));
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

      // A page break inserted at the front makes one more page.
      const before = await js(`document.querySelectorAll('.wd-sheet').length`);
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

  // RUTBA_VERIFY_ONLY=pages,grips,panes: those blocks alone, for working on them.
  const only = (process.env.RUTBA_VERIFY_ONLY || '').split(',').map((s) => s.trim()).filter(Boolean);
  if (only.length) {
    if (only.includes('pages')) await wordPages();
    if (only.includes('grips')) await sheetGrips();
    if (only.includes('panes')) await slidePanes();
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
  await sheetGrips();
  await slidePanes();

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
        const b = [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || '').startsWith(${JSON.stringify(title)}));
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
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || '').startsWith(${JSON.stringify(title)}) && !n.disabled);
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
      return { blocks: m.blocks.length, texts: m.blocks.map((b) => b.text), styles: m.blocks.map((b) => b.style), format: m.format, footnotes: (m.footnotes || []).length };
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
      return { count: dead.length, unexplained: dead.filter((b) => !/not built yet/.test(b.title || '') && !/^(Undo|Redo) /.test(b.title || '')).map((b) => b.title || b.textContent.trim()).slice(0, 5) };
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
    await press('Built from the headings');
    await until(async () => (await engine()).texts.includes('Contents'), 'the contents to appear', 4000).catch(() => {});
    const toc = await engine();
    check('word: References → Table of Contents lists the headings', toc.texts.includes('Contents') && toc.blocks > before, `${before} → ${toc.blocks} blocks; ${JSON.stringify(toc.texts.slice(0, 3))}`);

    // References → Insert Footnote: the dialog takes the words, the page
    // gets a raised number at the caret and the note under the body.
    await js(`(() => { const page = document.querySelector('.wd-page'); const b = page.querySelector('[data-block="2"]'); const r = document.createRange(); r.selectNodeContents(b); r.collapse(false); const s = getSelection(); s.removeAllRanges(); s.addRange(r); page.dispatchEvent(new MouseEvent('mouseup', { bubbles: true })); return 'caret at end of block 2'; })()`);
    await wait(200);
    await press('A raised number at the caret, and its words under the body');
    await until(() => js(`Boolean(document.querySelector('.rw-dialog textarea'))`), 'the footnote dialog', 4000).catch(() => {});
    await js(`(() => { const ta = document.querySelector('.rw-dialog textarea'); if (!ta) return 'no textarea'; Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(ta, 'See regulation 57.'); ta.dispatchEvent(new Event('input', { bubbles: true })); return 'typed'; })()`);
    await wait(150);
    await js(`[...document.querySelectorAll('.rw-dialog .rw-btn')].find((b) => b.textContent.trim() === 'Insert')?.click(), 'inserted'`);
    await until(async () => (await engine()).footnotes === 1, 'the footnote', 5000).catch(() => {});
    const noted = await engine();
    const painted = await js(`({ refs: document.querySelectorAll('.wd-page .wd-noteref').length, notes: document.querySelectorAll('.wd-notes .wd-note').length, words: document.querySelector('.wd-notes')?.textContent.trim() || '' })`);
    check('word: References → Insert Footnote puts the number in the text and the note under the body', noted.footnotes === 1 && painted.refs === 1 && painted.notes === 1 && /regulation 57/.test(painted.words), `engine footnotes ${noted.footnotes}; painted ${JSON.stringify(painted)}`);

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
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || '').startsWith(${JSON.stringify(title)}) && !n.disabled);
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
      return { count: dead.length, unexplained: dead.filter((b) => !/not built yet/.test(b.title || '') && !/^(Undo|Redo) /.test(b.title || '')).map((b) => b.title || b.textContent.trim()).slice(0, 5) };
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
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || '').startsWith(${JSON.stringify(title)}));
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
    const find = `[...document.querySelectorAll('.rw-btn, .ml-compose-cta button')].find((n) => (n.title || n.textContent || '').trim().startsWith(${JSON.stringify(title)}) && !n.disabled)`;
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
      const b = [...document.querySelectorAll('.rw-btn')].find((x) => (x.title || '').startsWith('Pictures'));
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
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || '').startsWith(${JSON.stringify(title)}) && !n.disabled);
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
      return { count: dead.length, unexplained: dead.filter((b) => !/not built yet|Select a text box first|decides its own size|Changing it rescales/.test(b.title || '') && !/^(Undo|Redo) /.test(b.title || '')).map((b) => b.title || b.textContent.trim()).slice(0, 5) };
    })()`);
    check('slides: every disabled control explains itself', honest.unexplained.length === 0, `${honest.count} disabled on Home with nothing selected; unexplained: ${JSON.stringify(honest.unexplained)}`);

    // View: the Slide Sorter shows every slide, Gridlines draw over the slide.
    await clickTab(win, 'View');
    await wait(120);
    await pushLabel('Slide Sorter');
    await until(() => js(`document.querySelectorAll('.sl-sortergrid .sl-sortercard').length > 0`), 'the sorter', 4000).catch(() => {});
    const sorted = await js(`({ cards: document.querySelectorAll('.sl-sortergrid .sl-sortercard').length, pics: document.querySelectorAll('.sl-sortergrid .sl-thumb-pic svg').length })`);
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
    await mouse(win, '.rw-btn[title^="Bold"]');
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
    await mouse(win, '.rw-btn[title^="Bold"]');
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
    await mouse(win, '.rw-btn[title^="Italic"]');
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
