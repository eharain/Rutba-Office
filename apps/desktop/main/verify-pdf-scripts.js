// Worksheets and Presentations: Arabic and Hebrew in the PDF — a workbook
// and a deck holding Arabic, Hebrew and Greek are written as PDFs the way
// Ctrl+P and Export do (through the page Chromium lays out and prints), and
// the PDF carries those letters as text: the fonts drawn with map each glyph
// back to its letter, so the words can be found, read aloud and copied out.
// Run alone with RUTBA_VERIFY_ONLY=pdfscripts.
import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';
import { buildXlsx } from '@rutba/ooxml/build';
import { buildPptx } from '@rutba/presentation';

const ARABIC = 'مرحبا بالعالم';
const HEBREW = 'שלום עולם';
const GREEK = 'Καλημέρα';

/** The letters a PDF's ToUnicode maps give back: every Unicode code point named in a bfchar or bfrange. */
export function pdfLetters(bytes) {
  const text = bytes.toString('latin1');
  const found = new Set();
  for (const m of text.matchAll(/<<([^]*?)>>\s*stream\r?\n/g)) {
    const start = m.index + m[0].length;
    const end = text.indexOf('endstream', start);
    if (end < 0) continue;
    const raw = bytes.subarray(start, end);
    let body;
    try { body = /FlateDecode/.test(m[1]) ? zlib.inflateSync(raw).toString('latin1') : raw.toString('latin1'); } catch { continue; }
    if (!/begincmap|beginbfchar|beginbfrange/.test(body)) continue;
    for (const c of body.matchAll(/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]{4,})>/g)) {
      const hex = c[2];
      for (let i = 0; i + 4 <= hex.length; i += 4) found.add(parseInt(hex.slice(i, i + 4), 16));
    }
    for (const r of body.matchAll(/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]{4,})>/g)) {
      const from = parseInt(r[1], 16);
      const to = parseInt(r[2], 16);
      const base = parseInt(r[3].slice(0, 4), 16);
      for (let k = 0; k <= Math.min(256, to - from); k++) found.add(base + k);
    }
  }
  return found;
}

/** Every letter of a word among those a PDF maps its glyphs back to (Arabic's joined forms come back as their letters). */
const has = (letters, word) => [...word.replace(/\s/g, '')].every((ch) => letters.has(ch.codePointAt(0)));

export async function verifyPdfScripts({ open, check, until, wait, errorsIn, doc, sessionFor }, { dir }) {
  const book = path.join(dir, 'scripts.xlsx');
  const deck = path.join(dir, 'scripts.pptx');
  try {
    fs.writeFileSync(book, buildXlsx({ sheets: [{ name: 'Words', rows: [['Language', 'Greeting'], ['Arabic', ARABIC], ['Hebrew', HEBREW], ['Greek', GREEK]] }] }));
    fs.writeFileSync(deck, buildPptx({ title: 'Greetings', slides: [{ layout: 'obj', title: ARABIC, body: [HEBREW, GREEK] }] }));
    for (const [app, file, kind, label] of [['sheets', book, 'sheet', 'sheets'], ['slides', deck, 'deck', 'slides']]) {
      const win = await open(app, file);
      const js = (code) => win.webContents.executeJavaScript(code);
      await until(() => js(`Boolean(document.querySelector('.rw-tab'))`), `the ${app} window`, 8000).catch(() => {});
      await wait(500);
      const session = doc.sessions().filter((s) => s.kind === kind && s.path === file).pop();
      const target = path.join(dir, `scripts-${app}.pdf`);
      const said = await js(`window.rutbaOffice.print.pdf({ id: ${JSON.stringify(session.id)}, path: ${JSON.stringify(target)}, options: {} }).then(() => 'written', (e) => 'failed ' + e.message)`);
      const letters = fs.existsSync(target) ? pdfLetters(fs.readFileSync(target)) : new Set();
      const found = { arabic: has(letters, ARABIC), hebrew: has(letters, HEBREW), greek: has(letters, GREEK) };
      check(`${label}: the PDF carries Arabic, Hebrew and Greek as text, each glyph mapped back to its letter`, said === 'written' && found.arabic && found.hebrew && found.greek, `${said}; ${JSON.stringify(found)}; ${letters.size} letters mapped`);
      const complaints = await errorsIn(win);
      check(`${label}: writing the PDF in those scripts reports nothing`, complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    }
  } catch (err) {
    check('the PDF script checks ran', false, err?.message || JSON.stringify(err));
  }
}
