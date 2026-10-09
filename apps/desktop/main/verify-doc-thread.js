// The document service on its own thread: while it lays a 240-page document
// out for print, seconds of work, the main process goes on answering — the
// launcher's Recent list comes back at once, not after the pages. Before
// the move the same request waited for every page. Run alone with
// RUTBA_VERIFY_ONLY=docthread.
import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';

const WORDS = 'the quick brown fox jumps over a lazy dog while ledgers balance and totals agree'.split(' ');
const sentence = (seed, n) => Array.from({ length: n }, (_, i) => WORDS[(seed * 7 + i * 3) % WORDS.length]).join(' ');

/** What the main process may take to answer while the thread is busy. */
const ANSWER_MS = 400;

/**
 * @param {object} h the harness: open, check, until, errorsIn
 * @param {{ dir: string }} args where the file is written
 */
export async function verifyDocThread({ open, check, until, errorsIn }, { dir }) {
  const file = path.join(dir, 'thread-240-pages.docx');
  fs.writeFileSync(file, buildDocx({ paragraphs: Array.from({ length: 900 }, (_, i) => ({ runs: [{ text: `${i + 1}. ${sentence(i, 120)} ` }, { text: sentence(i + 1, 80), bold: i % 3 === 0 }] })) }));
  try {
    const win = await open('word', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="0"]'))`), 'the document', 20000);
    const measured = await js(`(async () => {
      const all = await window.rutbaOffice.doc.sessions({});
      const id = all.filter((s) => s.kind === 'doc').pop().id;
      let laidOut = 0;
      const t0 = performance.now();
      // Seconds of layout on the thread, not awaited yet.
      const pages = window.rutbaOffice.print.summary({ id, options: {} }).then((s) => { laidOut = performance.now() - t0; return s.pages; });
      const r0 = performance.now();
      await window.rutbaOffice.app.recent();
      const recentMs = performance.now() - r0;
      const whileBusy = laidOut === 0;
      const count = await pages;
      return { recentMs: Math.round(recentMs), layoutMs: Math.round(laidOut), whileBusy, pages: count };
    })()`);
    check(`documents: while the document service lays out 240 pages on its own thread, the main process answers another window's request in under ${ANSWER_MS} ms`,
      measured.whileBusy && measured.recentMs < ANSWER_MS && measured.pages >= 200 && measured.layoutMs > measured.recentMs,
      JSON.stringify(measured));
    const complaints = await errorsIn(win);
    check('documents: the busy thread reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('documents: the document thread checks ran', false, err?.message || String(err));
  }
}
