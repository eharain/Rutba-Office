// An archive indexed on a thread of its own, the windows answering meanwhile.
//
// A 1 GB .pst (20,000 messages, the archive of the performance budget) is
// written for the check into a folder of its own and opened in place. While
// it is indexed: a document in another window is typed into and each key is
// on its page within 200 ms; the main process is timed by a timer that asks
// for a turn every millisecond, and is never held as long as 100 ms; the
// archive's strip counts how far indexing has got; a search over all mail
// answers from what is indexed so far and says so ("4,000 of 20,000 indexed
// so far"), its results growing as it goes on. Stop indexing pauses it (the
// count stands still, what is indexed stays searchable), Resume indexing
// carries on to the end, and the search then finds everything with no word
// of indexing. Closing the archive leaves no thread behind and the file as it
// was written. Run alone with RUTBA_VERIFY_ONLY=mail-index-worker.

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import { Worker } from 'node:worker_threads';
import { PerformanceObserver } from 'node:perf_hooks';
import { buildDocx } from '@rutba/ooxml/build';

const COUNT = 20_000;
/** How long a key may take to reach the page while the archive is indexed. */
const KEY_MS = 200;
/**
 * The longest the main process may be held while it is: 100 ms, the time
 * within which a response still reads as instant. Run alone it stays under
 * 50 ms; in the full gate, with every other window's work on the same
 * process, turns of 59 and 62 ms were seen (merging and sending search
 * results), where the indexing this guards once held it for seconds.
 */
const STALL_MS = 100;

/**
 * The 1 GB fixture, written on a thread so the run's other windows are not
 * held the nine seconds it takes, flushed to the disk there and read
 * through once: a gigabyte just written is still being written out by the
 * system for seconds after, and is scanned by the virus checker the first
 * time it is opened, and anything of this process's own meanwhile waits
 * behind both - the fixture's doing, not the archive's. A person's archive
 * was written long ago.
 */
function writeFixture(file) {
  const writer = new URL('../../../tests/fixtures/pst-writer.js', import.meta.url).href;
  const code = `const { parentPort, workerData } = require('node:worker_threads');
const fs = require('node:fs');
const flushed = (r) => { let fd = fs.openSync(workerData.file, 'r+'); fs.fsyncSync(fd); fs.closeSync(fd); fd = fs.openSync(workerData.file, 'r'); const b = Buffer.allocUnsafe(8 << 20); let at = 0, n; while ((n = fs.readSync(fd, b, 0, b.length, at)) > 0) at += n; fs.closeSync(fd); return r; };
import(workerData.writer).then((m) => m.writeCorpusPst(workerData.file, { count: workerData.count, attachBytes: 190000 })).then(flushed).then((r) => parentPort.postMessage({ ok: true, r }), (e) => parentPort.postMessage({ ok: false, e: String(e && e.stack || e) }));`;
  return new Promise((resolve, reject) => {
    const w = new Worker(code, { eval: true, workerData: { writer, file, count: COUNT } });
    w.once('message', (m) => {
      w.terminate().catch(() => {});
      if (m.ok) resolve(m.r);
      else reject(new Error(m.e));
    });
    w.once('error', reject);
  });
}

const sha = (file) => crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');

/**
 * @param {object} h the harness: open, check, until, wait, capture, errorsIn
 */
export async function verifyMailIndexWorker(h) {
  const { open, check, until, wait, capture, errorsIn } = h;
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-verify-index-'));
  const file = path.join(root, 'Large archive.pst');
  const note = path.join(root, 'index-note.docx');
  let mailWin = null;
  let archiveId = null;
  let tick = null;
  try {
    const written = await writeFixture(file);
    fs.writeFileSync(note, buildDocx({ paragraphs: ['Notes while the archive is indexed', 'Typed while it indexes: '], styles: true }));
    const before = sha(file);

    mailWin = await open('mail');
    const word = await open('word', note);
    const mjs = (code) => mailWin.webContents.executeJavaScript(code);
    const wjs = (code) => word.webContents.executeJavaScript(code);
    const wc = word.webContents;
    const mail = (method, args) => mjs(`window.rutbaOffice.mail.${method}(${JSON.stringify(args ?? {})})`);
    const shot = async (win, name) => {
      win.webContents.invalidate();
      await wait(700);
      const was = measuring;
      measuring = false;
      await capture(win, name);
      await wait(20);
      last = performance.now();
      lastCpu = threadCpu?.() ?? 0;
      measuring = was;
    };
    const waitFor = async (fn, what, ms) => {
      let got = null;
      await until(async () => {
        got = await fn();
        return Boolean(got);
      }, what, ms).catch(() => {});
      return got || null;
    };
    const stateOf = async () => (await mail('archives')).find((a) => a.id === archiveId)?.index || null;
    await until(() => mjs(`document.querySelectorAll('.ml-accounts button, .ml-archives button').length > 0 || Boolean(document.querySelector('.ml-sidebar, .rw-ribbon'))`), 'the mail window', 10000).catch(() => {});
    await until(() => wjs(`Boolean(document.querySelector('.wd-page [data-block="1"]'))`), 'the document', 10000);

    // The main process timed from the moment the archive is opened, by the
    // phase of the check it was in: a timer asks for a turn every
    // millisecond, and between two turns the main thread's own CPU time,
    // less its garbage collector's pauses, is how long it was held running
    // something. The time between the turns also counts a thread the machine
    // did not run (another program busy), and the collector's pauses are
    // this process's whole heap's, the check run's; both are said beside it.
    // A capture is the harness's own work on this process, a PNG encoded
    // here, so it is left out.
    const threadCpu = typeof process.threadCpuUsage === 'function' ? () => { const u = process.threadCpuUsage(); return (u.user + u.system) / 1000; } : null;
    let last = performance.now();
    let lastCpu = threadCpu?.() ?? 0;
    let measuring = true;
    let phase = 'open';
    const worst = { ms: 0, phase: '', gap: 0, gapPhase: '' };
    const gaps = [];
    const heldTicks = [];
    // The garbage collector's pauses on this process, said beside the rest:
    // its heap is the whole check run's, not the archive's.
    const pauses = [];
    const gcSeen = new PerformanceObserver((list) => {
      for (const e of list.getEntries()) pauses.push({ start: e.startTime, duration: e.duration });
    });
    gcSeen.observe({ entryTypes: ['gc'] });
    tick = setInterval(() => {
      const now = performance.now();
      const cpu = threadCpu?.() ?? 0;
      const gap = now - last;
      const held = threadCpu ? Math.min(gap, cpu - lastCpu) : gap;
      if (measuring) {
        if (held > 20) heldTicks.push({ from: last, to: now, held, phase });
        if (gap > worst.gap) Object.assign(worst, { gap, gapPhase: phase });
        if (gap > 30) gaps.push(`${phase}+${Math.round(gap)}${threadCpu ? ` (held ${Math.round(held)})` : ''}`);
      }
      last = now;
      lastCpu = cpu;
    }, 1);

    const opened = await mail('openArchive', { path: file });
    archiveId = opened.id;
    check(`mail index worker: a ${Math.round(written.bytes / 1048576)} MB .pst of ${COUNT.toLocaleString('en-GB')} messages opens in place and starts indexing`, opened.messages === COUNT && ['indexing', 'waiting'].includes(opened.index.state), JSON.stringify({ messages: opened.messages, index: opened.index, bytes: written.bytes }));

    phase = 'typing';
    // ── The document typed into while the archive indexes.
    const caret = async () => {
      const at = await wjs(`(() => { const b = document.querySelector('.wd-page [data-block="1"]'); if (!b) return null; b.scrollIntoView({ block: 'center' }); const q = b.getBoundingClientRect(); const r = document.createRange(); r.selectNodeContents(b); const t = r.getBoundingClientRect(); return { x: Math.round(t.right + 6), y: Math.round(q.top + q.height / 2) }; })()`);
      if (!at) return false;
      wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      wc.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      await wait(350);
      return true;
    };
    // The document window times its own keys: when each "q" went down and when the page first held it.
    const probe = () => wjs(`(() => {
      const count = () => ((document.querySelector('.wd-page')?.textContent || '').match(/q/g) || []).length;
      window.__ix = { base: count(), downs: [], seen: [] };
      if (!window.__ixHooked) {
        window.__ixHooked = true;
        window.addEventListener('keydown', (e) => { if (e.key === 'q' && window.__ix) window.__ix.downs.push(performance.now()); }, true);
        new MutationObserver(() => {
          const p = window.__ix;
          if (!p) return;
          const n = count() - p.base;
          while (p.seen.length < n) p.seen.push(performance.now());
        }).observe(document.body, { subtree: true, childList: true, characterData: true });
      }
      return window.__ix.base;
    })()`);
    const key = () => {
      wc.sendInputEvent({ type: 'keyDown', keyCode: 'q' });
      wc.sendInputEvent({ type: 'char', keyCode: 'q' });
      wc.sendInputEvent({ type: 'keyUp', keyCode: 'q' });
    };
    await caret();
    await probe();
    let typed = 0;
    let during = 0;
    while (typed < 14) {
      const s = await stateOf();
      const busy = s?.state === 'indexing';
      if (!busy && typed >= 2) break;
      key();
      typed++;
      if (busy) during++;
      await wait(90);
    }
    await until(() => wjs(`window.__ix.seen.length >= ${typed}`), 'every key on the page', 8000).catch(() => {});
    const p = await wjs(`({ downs: window.__ix.downs, seen: window.__ix.seen })`);
    const lags = p.downs.map((d, i) => (p.seen[i] == null ? null : Math.round(p.seen[i] - d)));
    const worstKey = lags.some((l) => l == null) ? null : Math.max(...lags);
    check(`mail index worker: while the archive indexes, another window's typed keys reach its page within ${KEY_MS} ms`,
      during >= 4 && lags.length === typed && worstKey != null && worstKey < KEY_MS,
      `${typed} keys, ${during} while indexing, ms to the page ${JSON.stringify(lags)}`);
    await shot(word, 'mail-index-1-document-typed.png');

    phase = 'strip';
    // ── The archive's strip counts how far indexing has got, with Stop.
    await mjs(`(() => { const b = [...document.querySelectorAll('.ml-archives button')].find((n) => n.textContent.trim().startsWith('Large archive')); b?.click(); return Boolean(b); })()`);
    const strip = await waitFor(async () => {
      const s = await mjs(`(() => ({ note: document.querySelector('.ml-archive-note')?.textContent || '', stop: Boolean(document.querySelector('.ml-archive-stop')), rows: document.querySelectorAll('.ml-row').length }))()`);
      return /Indexing for search: [\d,]+ of 20,000 messages, searchable as they are indexed\./.test(s.note) && s.stop && s.rows > 0 ? s : false;
    }, 'the strip counting the indexing', 10000);
    check('mail index worker: the archive opens as a mailbox while it indexes, and its strip counts the messages indexed ("… of 20,000") with Stop indexing beside it', Boolean(strip), JSON.stringify(strip));
    await shot(mailWin, 'mail-index-2-strip-indexing.png');

    phase = 'stop';
    // ── Stop indexing: the count stands still, and what is indexed is searched and said to be so.
    const pressed = await mjs(`(() => { const b = document.querySelector('.ml-archive-stop'); if (!b) return 'no Stop'; b.click(); return 'pressed'; })()`);
    const paused = await waitFor(async () => {
      const s = await mjs(`(() => ({ note: document.querySelector('.ml-archive-note')?.textContent || '', resume: Boolean(document.querySelector('.ml-archive-resume')) }))()`);
      const st = await stateOf();
      return st?.state === 'paused' && s.resume && /Indexing stopped at [\d,]+ of 20,000 messages; what is indexed can be searched\./.test(s.note) ? { ...s, done: st.done } : false;
    }, 'the indexing paused', 8000);
    await shot(mailWin, 'mail-index-3-stopped.png');

    phase = 'search while stopped';
    await mjs(`(() => { const t = [...document.querySelectorAll('.rw-tab')].find((n) => n.textContent.trim() === 'Search'); t?.click(); return 1; })()`);
    await mjs(`(() => { const el = document.querySelector('.rw-search input'); if (!el) return 'no field'; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, 'invoice'); el.dispatchEvent(new Event('input', { bubbles: true })); return 'typed'; })()`);
    await until(() => mjs(`document.querySelectorAll('.ml-scope-seg button').length === 2`), 'the scope bar', 5000).catch(() => {});
    await mjs(`(() => { const b = [...document.querySelectorAll('.ml-scope-seg button')].find((n) => n.textContent.trim() === 'All mail'); b?.click(); return Boolean(b); })()`);
    const readScope = () => mjs(`(() => { const f = document.querySelector('.ml-scope-found'); const i = document.querySelector('.ml-scope-indexing'); return { total: f ? Number(f.dataset.total) : null, query: f?.dataset.query || null, indexing: i?.textContent || '', rows: document.querySelectorAll('.ml-row').length }; })()`);
    const counted = (text, words) => {
      const m = new RegExp(`Large archive: ([\\d,]+) of 20,000 ${words}`).exec(text);
      return m ? Number(m[1].replace(/,/g, '')) : null;
    };
    const atPause = await waitFor(async () => {
      const s = await readScope();
      const done = counted(s.indexing, 'indexed, and indexing is stopped; the rest is not searched until it is resumed\\.');
      return s.query === 'invoice' && s.total > 0 && s.rows > 0 && done != null && done < COUNT ? { ...s, done } : false;
    }, 'a search answered from what was indexed before the stop', 15000);
    await wait(1500);
    const still = await stateOf();
    const again = await readScope();
    await shot(mailWin, 'mail-index-4-search-stopped.png');
    check('mail index worker: Stop indexing pauses it (the strip says where with Resume indexing, the count stands still), and a search over all mail answers from what was indexed and says so ("… of 20,000 indexed, and indexing is stopped")',
      pressed === 'pressed' && Boolean(paused) && Boolean(atPause) && still?.state === 'paused' && still.done === paused.done && again.total === atPause.total,
      JSON.stringify({ pressed, paused, atPause, still, again: again.total }));

    phase = 'resume';
    // ── Resume: the search, asked again as indexing goes on, grows and says how far.
    await mjs(`(() => { const t = [...document.querySelectorAll('.rw-tab')].find((n) => n.textContent.trim() === 'Home'); t?.click(); return 1; })()`);
    const resumed = await waitFor(() => mjs(`(() => { const b = document.querySelector('.ml-archive-resume'); if (!b) return false; b.click(); return 'pressed'; })()`), 'Resume indexing', 6000);
    let early = null;
    const grown = await waitFor(async () => {
      const s = await readScope();
      const done = counted(s.indexing, 'indexed so far; more may be found as indexing goes on\\.');
      if (done != null && !early && s.total > (atPause?.total || 0)) early = { ...s, done };
      // On a quiet machine the rest can be indexed between two looks, so the
      // search goes straight to everything: grown results from a finished
      // index show indexing carried on just as well.
      if (!early && s.total > (atPause?.total || 0)) {
        const now = await stateOf();
        if (now?.state === 'ready') early = { ...s, done: now.done };
      }
      return early ? early : false;
    }, 'the results growing as indexing goes on', 20000);
    if (grown) await shot(mailWin, 'mail-index-5-search-while-indexing.png');
    check('mail index worker: Resume indexing carries on, and the search over all mail, asked again as it goes, grows and says how far ("… of 20,000 indexed so far")',
      resumed === 'pressed' && Boolean(grown) && grown.done > (paused?.done || 0), JSON.stringify({ resumed, grown, paused: paused?.done }));
    const ready = await waitFor(async () => {
      const s = await stateOf();
      return s?.state === 'ready' ? s : false;
    }, 'the archive indexed to the end', 120000);
    measuring = false;
    // Each long turn's time less the collector's pauses inside it.
    await wait(50);
    gcSeen.disconnect();
    for (const t of heldTicks) {
      const inGc = pauses.reduce((n, p) => n + Math.max(0, Math.min(t.to, p.start + p.duration) - Math.max(t.from, p.start)), 0);
      const work = Math.max(0, t.held - inGc);
      if (work > worst.ms) Object.assign(worst, { ms: work, phase: t.phase });
    }
    const longestGc = Math.max(0, ...pauses.map((p) => p.duration));
    check(`mail index worker: indexing goes on to all ${COUNT.toLocaleString('en-GB')} messages`, Boolean(ready) && ready.done === COUNT, JSON.stringify(ready));
    check(`mail index worker: the main process was never held ${STALL_MS} ms by work while the archive was opened, indexed and searched`, worst.ms < STALL_MS, `held at most ${worst.ms.toFixed(1)} ms by work (${threadCpu ? 'its own CPU time between turns' : 'the time between turns'}, less the collector's pauses)${worst.phase ? `, while ${worst.phase}` : ''}; the longest time between turns ${worst.gap.toFixed(1)} ms, while ${worst.gapPhase}; turns over 30 ms apart ${JSON.stringify(gaps)}; the garbage collector's longest pause ${Math.round(longestGc)} ms of ${pauses.length}, the heap ${Math.round(process.memoryUsage().heapUsed / 1048576)} MB`);

    // ── Once indexed, the search finds them all, with no word of indexing.
    await mjs(`(() => { const t = [...document.querySelectorAll('.rw-tab')].find((n) => n.textContent.trim() === 'Search'); t?.click(); return 1; })()`);
    const whole = await mail('searchAll', { query: 'invoice', limit: 5 });
    const settled = await waitFor(async () => {
      const s = await readScope();
      return s.query === 'invoice' && s.total === whole.total && !s.indexing ? s : false;
    }, 'the search over the whole archive', 10000);
    check('mail index worker: once indexed, the search over all mail finds every match, with no word of indexing left', Boolean(settled) && whole.indexing.length === 0 && whole.total > (early?.total || 0), JSON.stringify({ settled, total: whole.total, early: early?.total }));
    await shot(mailWin, 'mail-index-5-search-indexed.png');

    const complaints = await errorsIn(mailWin);
    check('mail index worker: the mail window reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');

    await mail('closeArchive', { id: archiveId });
    archiveId = null;
    check('mail index worker: closing the archive leaves the file byte for byte as it was written', sha(file) === before, path.basename(file));
  } catch (err) {
    check('mail index worker: the checks ran', false, err.stack || err.message);
  } finally {
    clearInterval(tick);
    try {
      if (archiveId && mailWin && !mailWin.isDestroyed()) await mailWin.webContents.executeJavaScript(`window.rutbaOffice.mail.closeArchive({ id: ${JSON.stringify(archiveId)} })`).catch(() => {});
    } finally {
      try {
        fs.rmSync(root, { recursive: true, force: true });
      } catch {
        /* a window may still hold a file */
      }
    }
  }
}
