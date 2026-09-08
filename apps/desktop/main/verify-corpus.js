// Open every file in a corpus in the REAL application, one window at a time,
// and say what happened to each: drawn, refused, hung, or crashed, and how
// long it took.
//
// The engine tests open files in node; `verify:apps` drives a handful of
// fixtures through the windows. Neither is what a person does, which is to
// double-click their own document and wait. This does that, for every
// document in the folders it is pointed at — the same main process, the same
// service, the same window, the same bundle — and records, per file:
//
//   - `ok`       the window drew the document (the app's own "I am showing
//                the file" element appeared), and the time it took;
//   - `refused`  the app said "This file could not be opened", with its reason;
//   - `timeout`  nothing appeared within the limit — a hang, or a file so slow
//                the person would have given up;
//   - `crashed`  the renderer went away;
//   - `skipped`  no app claims the extension.
//
// Console errors the window logged are kept with the file. The report names
// the failures, the slowest files, and the numbers per format.
//
//   npm run verify:corpus                    the default corpus folders
//   RUTBA_CORPUS_DIRS="D:\docs;E:\more"  ...  other folders (semicolon-separated)
//   RUTBA_CORPUS_LIST=files.txt          ...  an explicit list, one path per line
//   RUTBA_CORPUS_LIMIT=50                ...  the first N files only
//   RUTBA_CORPUS_TIMEOUT=30000           ...  ms to wait for each window
//   RUTBA_CORPUS_OUT=<dir>               ...  where report.json / report.txt go

import fs from 'node:fs';
import path from 'node:path';
// One reading of a console-message for both harnesses.
import { consoleMessage } from './console-message.js';

const EXTENSIONS = new Set([
  '.docx', '.doc', '.odt', '.rtf', '.txt', '.md', '.html', '.htm',
  '.xlsx', '.xls', '.ods', '.csv', '.tsv',
  '.pptx', '.ppt', '.odp',
  '.pdf', '.png', '.jpg', '.jpeg', '.gif', '.webp', '.bmp',
  '.mp4', '.webm', '.mkv', '.mov', '.mp3', '.wav', '.ogg', '.m4a',
]);

/** The element each app renders only once it is showing the file. */
const READY = {
  word: '.wd-page .wd-block, .wd-block',
  sheets: '.sh-cells',
  slides: '.sl-svg, .sl-thumb',
  pictures: '.pv-image, .pv-pdf, .pv-video, .pv-audio',
  image: '.im-canvas',
  video: '.vd-clip',
};

/** What each app shows when it gives up on a file. */
const REFUSED = '.rw-empty h3';

const SKIP_DIRS = new Set(['node_modules', '.git', '.next', 'release', '.worktrees']);

function collect(dirs) {
  const out = [];
  const walk = (dir) => {
    let entries = [];
    try {
      entries = fs.readdirSync(dir, { withFileTypes: true });
    } catch {
      return;
    }
    for (const e of entries) {
      const full = path.join(dir, e.name);
      if (e.isDirectory()) {
        if (!SKIP_DIRS.has(e.name)) walk(full);
      } else if (EXTENSIONS.has(path.extname(e.name).toLowerCase())) {
        out.push(full);
      }
    }
  };
  for (const d of dirs) walk(d);
  return out.sort((a, b) => a.localeCompare(b));
}

const wait = (ms) => new Promise((r) => setTimeout(r, ms));

/** executeJavaScript that cannot hang the run when the renderer is wedged. */
function ask(win, code, ms = 4000) {
  return Promise.race([
    win.webContents.executeJavaScript(code).catch((e) => ({ __error: e.message })),
    wait(ms).then(() => ({ __hung: true })),
  ]);
}

const percentile = (values, p) => {
  if (!values.length) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  return sorted[Math.min(sorted.length - 1, Math.floor((sorted.length - 1) * p))];
};

export async function verifyCorpus({ windows, doc = null, appForFile }) {

  const outDir = process.env.RUTBA_CORPUS_OUT || path.join(process.cwd(), 'build', 'corpus');
  fs.mkdirSync(outDir, { recursive: true });
  const timeout = Number(process.env.RUTBA_CORPUS_TIMEOUT || 30000);
  const limit = Number(process.env.RUTBA_CORPUS_LIMIT || 0);

  let files;
  if (process.env.RUTBA_CORPUS_LIST) {
    files = fs.readFileSync(process.env.RUTBA_CORPUS_LIST, 'utf8').split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  } else {
    const dirs = (process.env.RUTBA_CORPUS_DIRS || '').split(';').map((d) => d.trim()).filter(Boolean);
    if (!dirs.length) throw new Error('RUTBA_CORPUS_DIRS or RUTBA_CORPUS_LIST is required');
    files = collect(dirs);
  }
  // A run cut short resumes where it stopped: RUTBA_CORPUS_START=138.
  const start = Number(process.env.RUTBA_CORPUS_START || 0);
  if (start > 0) files = files.slice(start);
  if (limit > 0) files = files.slice(0, limit);

  console.log(`corpus: ${files.length} files, ${timeout} ms each at most`);

  const results = [];
  const log = fs.createWriteStream(path.join(outDir, 'progress.txt'), { flags: 'a' });
  // One window stays open for the whole run: closing the last window quits
  // the application, and each file's window is closed as soon as it has
  // answered. The launcher is the cheapest window there is.
  const keeper = windows.create({ app: 'home' });

  let n = 0;
  for (const file of files) {
    n += 1;
    const ext = path.extname(file).toLowerCase();
    const app = appForFile(file);
    let size = 0;
    try {
      size = fs.statSync(file).size;
    } catch {
      results.push({ file, ext, app, status: 'missing', ms: 0, messages: [] });
      continue;
    }
    if (!READY[app]) {
      results.push({ file, ext, app, size, status: 'skipped', ms: 0, messages: [`no window for ${app}`] });
      continue;
    }

    const started = Date.now();
    const messages = [];
    let gone = false;
    let unresponsive = false;
    const win = windows.create({ app, file });
    win.webContents.on('console-message', (...args) => {
      const m = consoleMessage(args);
      if (m.level >= 2) messages.push(`${m.text.slice(0, 300)} (${m.source.split('/').pop()}:${m.line})`);
    });
    win.webContents.on('render-process-gone', (_e, details) => {
      gone = true;
      messages.push(`renderer gone: ${details.reason}`);
    });
    win.on('unresponsive', () => {
      unresponsive = true;
    });

    let status = 'timeout';
    let facts = null;
    let reason = null;
    let loadedMs = null;
    const loaded = await new Promise((resolve) => {
      const timer = setTimeout(() => resolve(false), Math.min(timeout, 15000));
      win.webContents.once('did-finish-load', () => {
        clearTimeout(timer);
        loadedMs = Date.now() - started;
        resolve(true);
      });

      win.webContents.once('did-fail-load', (_e, code, desc) => {
        clearTimeout(timer);
        messages.push(`did-fail-load ${code} ${desc}`);
        resolve(false);
      });
    });

    if (loaded) {
      const probe = `(() => {
        const ready = document.querySelector(${JSON.stringify(READY[app])});
        const refused = [...document.querySelectorAll(${JSON.stringify(REFUSED)})].find((h) => /could not be opened/i.test(h.textContent));
        if (refused) return { refused: (refused.parentElement?.textContent || '').replace(refused.textContent, '').trim().slice(0, 300) };

        if (!ready) return { waiting: true };
        // An element is not a picture. An <img> whose file never decoded is
        // still an <img>, so a corrupt PNG drew an empty frame and counted as
        // shown; a media element answers for itself only once it has
        // metadata. Waiting here means the deadline decides, as it does for
        // every other kind of file.
        if (ready.tagName === 'IMG' && !ready.naturalWidth) return { waiting: true };
        if (ready.tagName === 'VIDEO' && !(ready.readyState >= 1)) return { waiting: true };
        return {
          ready: true,
          blocks: document.querySelectorAll('.wd-block').length,
          cells: document.querySelectorAll('.sh-cell').length,
          slides: document.querySelectorAll('.sl-thumb').length,
          text: (document.body.innerText || '').length,
        };
      })()`;
      while (Date.now() - started < timeout) {
        if (gone) {
          status = 'crashed';
          break;
        }
        const r = await ask(win, probe);
        if (unresponsive) {
          status = 'hung';
          break;
        }
        // No answer in four seconds is a busy renderer, not yet a hung one:
        // a window drawing a thousand blocks answers late and then answers.
        // Only the deadline decides.
        if (r?.__hung) continue;

        if (r?.refused) {
          status = 'refused';
          reason = r.refused;
          break;
        }
        if (r?.ready) {
          status = 'ok';
          facts = { blocks: r.blocks, cells: r.cells, slides: r.slides, text: r.text };
          break;
        }
        await wait(150);
      }
    } else if (gone) {
      status = 'crashed';
    }
    const ms = Date.now() - started;
    if (status === 'timeout' && unresponsive) status = 'hung';
    // Shown, but only after the person would have given up.
    if (status === 'ok' && ms > 10000) messages.push(`slow: ${ms} ms to first paint`);


    try {
      win.destroy();
    } catch {
      /* already gone */
    }
    await wait(120);
    // What the run is costing the main process, per file: a leak shows here
    // as a number that only ever grows.
    const rssMb = Math.round(process.memoryUsage().rss / 1048576);
    const open = doc?.sessions ? doc.sessions().length : null;

    results.push({ file, ext, app, size, status, ms, loadedMs, reason, facts, messages, rssMb, sessions: open });
    const line = `${String(n).padStart(4)}/${files.length} ${status.padEnd(7)} ${String(ms).padStart(6)} ms ${(size / 1048576).toFixed(1).padStart(6)} MB ${app.padEnd(8)} rss ${String(rssMb).padStart(5)} MB${open != null ? ` open ${open}` : ''} ${path.basename(file)}${reason ? ` — ${reason}` : ''}${messages.length ? ` [${messages.length} console]` : ''}`;
    console.log(line);
    log.write(line + '\n');
  }

  log.end();
  try {
    keeper.destroy();
  } catch {
    /* gone already */
  }

  // The report: failures first, then the slow ones, then the numbers.

  const lines = [];
  const bad = results.filter((r) => ['refused', 'timeout', 'hung', 'crashed', 'missing'].includes(r.status));
  lines.push(`# Corpus: ${results.length} files, ${bad.length} not shown`);
  lines.push('');
  if (bad.length) {
    lines.push('## Not shown');
    for (const r of bad) {
      lines.push(`${r.status.padEnd(7)} ${r.app.padEnd(8)} ${(r.size / 1048576).toFixed(1).padStart(6)} MB ${String(r.ms).padStart(6)} ms  ${r.file}`);
      if (r.reason) lines.push(`         ${r.reason}`);
      for (const m of r.messages.slice(0, 4)) lines.push(`         console: ${m}`);
    }
    lines.push('');
  }
  const noisy = results.filter((r) => r.status === 'ok' && r.messages.length);
  if (noisy.length) {
    lines.push(`## Shown, but the window logged an error (${noisy.length})`);
    for (const r of noisy.slice(0, 60)) {
      lines.push(`${r.app.padEnd(8)} ${r.file}`);
      for (const m of r.messages.slice(0, 3)) lines.push(`         ${m}`);
    }
    lines.push('');
  }
  const shown = results.filter((r) => r.status === 'ok');
  lines.push('## Slowest twenty that did open');
  for (const r of [...shown].sort((a, b) => b.ms - a.ms).slice(0, 20)) {
    lines.push(`${String(r.ms).padStart(6)} ms ${(r.size / 1048576).toFixed(1).padStart(6)} MB ${r.app.padEnd(8)} ${r.file}`);
  }
  lines.push('');
  lines.push('## By format');
  lines.push('ext     files    ok  refused timeout hung crash  skip   median    p90     max');
  const byExt = new Map();
  for (const r of results) {
    const e = byExt.get(r.ext) || { files: 0, ok: 0, refused: 0, timeout: 0, hung: 0, crashed: 0, skipped: 0, times: [] };
    e.files += 1;
    e[r.status] = (e[r.status] || 0) + 1;
    if (r.status === 'ok') e.times.push(r.ms);
    byExt.set(r.ext, e);
  }
  for (const [ext, e] of [...byExt.entries()].sort((a, b) => b[1].files - a[1].files)) {
    lines.push(
      `${ext.padEnd(7)} ${String(e.files).padStart(5)} ${String(e.ok).padStart(5)} ${String(e.refused).padStart(8)} ${String(e.timeout).padStart(7)} ${String(e.hung).padStart(4)} ${String(e.crashed).padStart(5)} ${String(e.skipped).padStart(5)} ` +
      `${String(percentile(e.times, 0.5)).padStart(8)} ${String(percentile(e.times, 0.9)).padStart(6)} ${String(percentile(e.times, 1)).padStart(7)}`
    );
  }
  lines.push('');
  lines.push(`${shown.length}/${results.length} files shown; ${bad.length} not; ${noisy.length} shown with console errors`);

  const report = lines.join('\n');
  fs.writeFileSync(path.join(outDir, 'report.txt'), report + '\n');
  fs.writeFileSync(path.join(outDir, 'report.json'), JSON.stringify(results, null, 2));
  console.log(report);
  return bad.length === 0;
}
