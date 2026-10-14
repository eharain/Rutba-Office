// An mbox and an Outlook for Mac archive past 2 GB, opened where they lie.
//
// Both are written for the check, into a folder of its own, on a thread (the
// seeded corpus, 9,000 messages with an attachment each, about 2.3 GB apiece,
// the .olm with zip64 records) - more than a Node buffer can hold, so either
// read whole would fail. Opened in place in the Mail window: each is listed
// with every message counted; the mbox's newest message, written past the
// 2 GB mark, opens in the reading pane from its list; both are indexed, and
// one search over all mail finds the message written last in each and opens
// the .olm's. The application's memory, sampled throughout, grows by less
// than a third of one archive's size. The files are unchanged at the end.
// Run alone with RUTBA_VERIFY_ONLY=mail-big-mbox.

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import { Worker } from 'node:worker_threads';

const COUNT = 9_000;
const TWO_GB = 2 ** 31;
const LAST_SUBJECT = 'Harbour lease: the final signed copy';

/** Both archives, written and flushed to the disk on a thread, so the run's windows are not held meanwhile. */
function writeFixtures(files) {
  const writer = new URL('../../../tests/fixtures/big-archives.js', import.meta.url).href;
  const code = `const { parentPort, workerData } = require('node:worker_threads');
const fs = require('node:fs');
const flush = (f) => { const fd = fs.openSync(f, 'r+'); fs.fsyncSync(fd); fs.closeSync(fd); };
import(workerData.writer).then((m) => {
  const a = m.writeCorpusMbox(workerData.files.mbox, { count: workerData.count });
  const b = m.writeCorpusOlm(workerData.files.olm, { count: workerData.count, zip64: true });
  flush(workerData.files.mbox);
  flush(workerData.files.olm);
  return { mbox: a, olm: b };
}).then((r) => parentPort.postMessage({ ok: true, r }), (e) => parentPort.postMessage({ ok: false, e: String(e && e.stack || e) }));`;
  return new Promise((resolve, reject) => {
    const w = new Worker(code, { eval: true, workerData: { writer, files, count: COUNT } });
    w.once('message', (m) => {
      w.terminate().catch(() => {});
      if (m.ok) resolve(m.r);
      else reject(new Error(m.e));
    });
    w.once('error', reject);
  });
}

/** The first and last 64 KB and the size: enough to say the file was not written to, without reading 2 GB twice. */
function stamp(file) {
  const fd = fs.openSync(file, 'r');
  try {
    const size = fs.fstatSync(fd).size;
    const head = Buffer.alloc(65536);
    const tail = Buffer.alloc(65536);
    fs.readSync(fd, head, 0, head.length, 0);
    fs.readSync(fd, tail, 0, tail.length, size - tail.length);
    return `${size}:${crypto.createHash('sha256').update(head).update(tail).digest('hex')}:${fs.fstatSync(fd).mtimeMs}`;
  } finally {
    fs.closeSync(fd);
  }
}

/**
 * @param {object} h the harness: open, check, until, wait, capture, errorsIn
 */
export async function verifyMailBigMbox(h) {
  const { open, check, until, wait, capture, errorsIn } = h;
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-verify-big-'));
  const files = { mbox: path.join(root, 'Gmail export.mbox'), olm: path.join(root, 'Mac archive.olm') };
  const ids = [];
  let win = null;
  let sampler = null;
  try {
    const written = await writeFixtures(files);
    const before = { mbox: stamp(files.mbox), olm: stamp(files.olm) };
    const sizes = { mbox: fs.statSync(files.mbox).size, olm: fs.statSync(files.olm).size };

    win = await open('mail');
    const js = (code) => win.webContents.executeJavaScript(code);
    const mail = (method, args) => js(`window.rutbaOffice.mail.${method}(${JSON.stringify(args ?? {})})`);
    const shot = async (name) => {
      win.webContents.invalidate();
      await wait(700);
      await capture(win, name);
    };
    const waitFor = async (fn, what, ms) => {
      let got = null;
      await until(async () => {
        got = await fn();
        return Boolean(got);
      }, what, ms).catch(() => {});
      return got || null;
    };
    await until(() => js(`Boolean(document.querySelector('.rw-ribbon'))`), 'the mail window', 10000).catch(() => {});

    // The application's memory from here: both archive threads are in this process.
    const base = process.memoryUsage().rss;
    let peak = base;
    sampler = setInterval(() => {
      peak = Math.max(peak, process.memoryUsage().rss);
    }, 100);

    const mb = await mail('openArchive', { path: files.mbox });
    ids.push(mb.id);
    const olm = await mail('openArchive', { path: files.olm });
    ids.push(olm.id);
    check(`mail big archives: an mbox of ${(sizes.mbox / 1048576).toFixed(0)} MB and an .olm of ${(sizes.olm / 1048576).toFixed(0)} MB, both past 2 GB, open in place with every message counted`,
      sizes.mbox > TWO_GB && sizes.olm > TWO_GB && mb.messages === COUNT + 1 && olm.messages === COUNT + 1 && mb.kind === 'mbox' && olm.kind === 'olm',
      JSON.stringify({ mbox: [mb.kind, mb.messages, mb.bytes], olm: [olm.kind, olm.messages, olm.bytes], written }));

    // The mbox in the window: listed from its file, its newest message (past 2 GB) opened.
    await js(`(() => { const b = [...document.querySelectorAll('.ml-archives button')].find((n) => n.textContent.trim().startsWith('Gmail export')); b?.click(); return Boolean(b); })()`);
    const listed = await waitFor(async () => {
      const s = await js(`(() => ({ rows: document.querySelectorAll('.ml-row').length, first: document.querySelector('.ml-row .ml-subject')?.textContent.trim() || '', banner: document.querySelector('.ml-archive-banner')?.textContent || '' }))()`);
      return s.rows > 0 && s.first === LAST_SUBJECT && /Mbox archive/.test(s.banner) ? s : false;
    }, 'the mbox listed', 60000);
    await js(`(() => { const r = [...document.querySelectorAll('.ml-row .ml-subject')].find((n) => n.textContent.trim() === ${JSON.stringify(LAST_SUBJECT)}); r?.click(); return Boolean(r); })()`);
    const reading = await waitFor(async () => {
      const s = await js(`(() => ({ head: document.querySelector('.ml-head h2')?.textContent || '', body: document.querySelector('.ml-body iframe')?.srcdoc || '', attachment: [...document.querySelectorAll('.ml-attachments .ml-attachment')].some((n) => n.textContent.includes('harbour-lease-signed.pdf')) }))()`);
      return s.head.includes(LAST_SUBJECT) && /quartzberry/.test(s.body) ? s : false;
    }, 'the newest message in the reading pane', 20000);
    check('mail big archives: the mbox lists its messages from the file, and its newest, written past the 2 GB mark, opens in the reading pane with its words and attachment',
      Boolean(listed) && Boolean(reading) && reading.attachment, JSON.stringify({ listed: listed && { rows: listed.rows, first: listed.first }, reading: reading && { head: reading.head, attachment: reading.attachment } }));
    await shot('mail-big-1-mbox-open.png');

    // Both indexed, then one search over all mail finds the message written last in each.
    const ready = await waitFor(async () => {
      const list = await mail('archives');
      const ours = list.filter((a) => ids.includes(a.id));
      return ours.length === 2 && ours.every((a) => a.index.state === 'ready') ? ours.map((a) => [a.name, a.index.done, a.index.total]) : false;
    }, 'both archives indexed', 240000);
    check(`mail big archives: both are indexed to their last message`, Boolean(ready) && ready.every(([, done]) => done === COUNT + 1), JSON.stringify(ready || (await mail('archives')).map((a) => [a.name, a.index])));
    await js(`(() => { const t = [...document.querySelectorAll('.rw-tab')].find((n) => n.textContent.trim() === 'Search'); t?.click(); return 1; })()`);
    await js(`(() => { const el = document.querySelector('.rw-search input'); if (!el) return 'no field'; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, 'quartzberry'); el.dispatchEvent(new Event('input', { bubbles: true })); return 'typed'; })()`);
    await until(() => js(`document.querySelectorAll('.ml-scope-seg button').length === 2`), 'the scope bar', 5000).catch(() => {});
    await js(`(() => { const b = [...document.querySelectorAll('.ml-scope-seg button')].find((n) => n.textContent.trim() === 'All mail'); b?.click(); return Boolean(b); })()`);
    const found = await waitFor(async () => {
      const s = await js(`(() => { const f = document.querySelector('.ml-scope-found'); return f ? { total: Number(f.dataset.total), query: f.dataset.query, where: [...document.querySelectorAll('.ml-where span')].map((n) => n.textContent) } : null; })()`);
      return s && s.query === 'quartzberry' && s.total === 2 && s.where.some((w) => w.startsWith('Gmail export')) && s.where.some((w) => w.startsWith('Mac archive')) ? s : false;
    }, 'the search over both', 20000);
    await js(`(() => { const r = [...document.querySelectorAll('.ml-row')].find((n) => n.querySelector('.ml-where')?.textContent.includes('Mac archive')); r?.querySelector('.ml-subject')?.click(); return Boolean(r); })()`);
    const olmMessage = await waitFor(async () => {
      const s = await js(`(() => ({ head: document.querySelector('.ml-head h2')?.textContent || '' }))()`);
      return s.head.includes(LAST_SUBJECT) ? s : false;
    }, 'the .olm result in the reading pane', 20000);
    check('mail big archives: one search over all mail finds the message written last in each archive, and the .olm\'s opens from its file',
      Boolean(found) && Boolean(olmMessage), JSON.stringify({ found, olmMessage }));
    await shot('mail-big-2-search-both.png');

    clearInterval(sampler);
    sampler = null;
    const grew = peak - base;
    const allowed = Math.min(sizes.mbox, sizes.olm) / 3;
    check('mail big archives: the application\'s memory grew by less than a third of one archive\'s size while both were opened, read and indexed',
      grew < allowed, `grew ${(grew / 1048576).toFixed(0)} MB, at most ${(allowed / 1048576).toFixed(0)} MB`);

    const complaints = await errorsIn(win);
    check('mail big archives: the window reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    for (const id of ids.splice(0)) await mail('closeArchive', { id });
    check('mail big archives: both archives are as they were written', stamp(files.mbox) === before.mbox && stamp(files.olm) === before.olm, Object.keys(files).join(', '));
  } catch (err) {
    check('mail big archives: the checks ran', false, err.stack || err.message);
  } finally {
    if (sampler) clearInterval(sampler);
    try {
      if (win && !win.isDestroyed()) for (const id of ids) await win.webContents.executeJavaScript(`window.rutbaOffice.mail.closeArchive({ id: ${JSON.stringify(id)} })`).catch(() => {});
    } finally {
      try {
        fs.rmSync(root, { recursive: true, force: true });
      } catch {
        /* a window may still hold a file */
      }
    }
  }
}
