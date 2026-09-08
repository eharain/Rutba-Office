// Boot the real application, photograph it, and quit.
//
// A build that compiles proves nothing about a window that has to paint. This
// opens each app for real — same main process, same preload, same bundle — waits
// for it to finish its first render, captures it, and reports what the renderer
// logged on the way. Any exception in a window shows up here as a failed
// capture rather than as a blank screen a user finds later.
//
//   npm run smoke               every app
//   npm run smoke -- word mail  just those
//   RUTBA_SMOKE_FILE=x.xlsx npm run smoke -- sheets

import fs from 'node:fs';
import path from 'node:path';
import { seedMail } from './seed-mail.js';
import { samplePictures } from './sample-picture.js';
import { consoleMessage } from './console-message.js';

const ALL = ['home', 'word', 'sheets', 'slides', 'pictures', 'image', 'video', 'mail'];

/**
 * Put a small archive through the real import path, so the mail screenshots
 * show the app doing its job rather than its empty state.
 *
 * The messages are generated here on purpose: a screenshot of somebody's actual
 * inbox has no business in a build artefact.
 */
export async function seedFor({ stores, mail }) {
  return seedMail({ stores, mail });
}

export async function runSmoke({ windows, outDir, stores, mail }) {
  const asked = process.argv.slice(2).filter((a) => ALL.includes(a));
  const list = asked.length ? asked : ALL;
  const file = process.env.RUTBA_SMOKE_FILE || null;
  fs.mkdirSync(outDir, { recursive: true });

  if (process.env.RUTBA_SMOKE_SEED && mail && stores) {
    try {
      const seeded = await seedMail({ stores, mail });
      console.log(`seeded ${seeded.messages} messages into ${seeded.folders} folder(s)`);
    } catch (err) {
      console.log(`seeding failed: ${err.message}`);
    }
  }

  const results = [];
  const opened = [];
  let samples = null;
  for (const app of list) {
    const started = Date.now();
    // The Pictures window opened with no file lists the operating system's own
    // pictures folder, so a capture run photographed whoever's machine it was
    // on. It gets pictures drawn for it instead — see sample-picture.js.
    const shown = file || (app === 'pictures' ? (samples || (samples = samplePictures())).first : null);
    const win = windows.create({ app, file: shown && app !== 'home' ? shown : null });
    const messages = [];
    win.webContents.on('console-message', (...args) => {
      const m = consoleMessage(args);
      if (m.level >= 2) messages.push(`${m.text} (${m.source.split('/').pop()}:${m.line})`);
    });
    win.webContents.on('render-process-gone', (_e, details) => messages.push(`renderer gone: ${details.reason}`));

    const ok = await new Promise((resolve) => {
      const timer = setTimeout(() => resolve(false), 12000);
      win.webContents.once('did-finish-load', () => {
        clearTimeout(timer);
        // Time for the window to SETTLE, not just to have loaded.
        //
        // 900ms is enough for a window that draws from data it already has.
        // It is not enough for one still decoding a half-megabyte JPEG or
        // waiting on an async read: the website captures came back with the
        // photo viewer showing blank thumbnails and an image rendered down to
        // the halfway line, and a mail window whose folder had not populated.
        // Both looked like product defects and were neither.
        //
        // RUTBA_SMOKE_SETTLE overrides it, which is what the screenshot
        // build does — a smoke run wants to be quick, and a marketing
        // capture wants to be right.
        setTimeout(() => resolve(true), Number(process.env.RUTBA_SMOKE_SETTLE) || 900);
      });
      win.webContents.once('did-fail-load', (_e, code, desc) => {
        clearTimeout(timer);
        messages.push(`did-fail-load ${code} ${desc}`);
        resolve(false);
      });
    });

    let captured = null;
    if (ok) {
      try {
        const image = await win.webContents.capturePage();
        const png = image.toPNG();
        captured = path.join(outDir, `${app}.png`);
        fs.writeFileSync(captured, png);
        if (png.length < 5000) messages.push(`capture is suspiciously small (${png.length} bytes)`);
      } catch (err) {
        messages.push(`capture failed: ${err.message}`);
        // What the page was when the compositor gave up — its height, its
        // pictures — and one retry with the page's shadow off, the usual
        // suspect when a 181-page contract is one element 200,000 px tall.
        try {
          const facts = await win.webContents.executeJavaScript(`(() => {
            const page = document.querySelector('.wd-page');
            return { docHeight: document.documentElement.scrollHeight, page: page ? page.scrollHeight : null, images: document.images.length, svg: document.querySelectorAll('svg').length };
          })()`);
          messages.push(`page: ${JSON.stringify(facts)}`);
          await win.webContents.executeJavaScript(`document.querySelectorAll('.wd-page').forEach((p) => { p.style.boxShadow = 'none'; }); 'done'`);
          await new Promise((r) => setTimeout(r, 500));
          const image = await win.webContents.capturePage();
          const png = image.toPNG();
          captured = path.join(outDir, `${app}.png`);
          fs.writeFileSync(captured, png);
          messages.push('captured on retry with the page shadow off');
        } catch (err2) {
          messages.push(`retry failed: ${err2.message}`);
        }
      }
    }

    results.push({ app, ok: ok && Boolean(captured), ms: Date.now() - started, capture: captured, messages });
    // Windows are kept until the end: destroying the last one fires
    // window-all-closed, which quits the application mid-run.
    opened.push(win);
  }

  const failed = results.filter((r) => !r.ok || r.messages.length);
  const lines = [];
  for (const r of results) {
    lines.push(`${r.ok ? 'ok  ' : 'FAIL'} ${r.app.padEnd(9)} ${String(r.ms).padStart(5)} ms  ${r.capture ? path.basename(r.capture) : '—'}`);
    for (const m of r.messages) lines.push(`       ${m}`);
  }
  lines.push('');
  lines.push(`${results.length - failed.length}/${results.length} windows rendered clean`);

  const report = lines.join('\n');
  // Electron on Windows is a GUI subsystem binary: its stdout is not attached
  // to the console that launched it, so the report is written to a file too.
  fs.writeFileSync(path.join(outDir, 'report.txt'), `${report}\n`);
  fs.writeFileSync(path.join(outDir, 'report.json'), JSON.stringify(results, null, 2));
  console.log(report);

  for (const win of opened) win.destroy();
  return failed.length === 0;
}
