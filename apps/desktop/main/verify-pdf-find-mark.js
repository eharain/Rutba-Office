// PDF: Find's mark sits on the word it found. A four-page PDF made here
// holds an English page that is only a picture of text, a typed page (a line
// in Helvetica and one in Courier), an Urdu page and an Arabic page that are
// only pictures. Recognise Text reads the three scans; then Find looks for
// words at the start, the middle and the end of lines on each page, at 100%,
// 150% and 200%, and each mark is measured against the word's own ink on the
// drawn page: the mark's left and right edges must be where the word's ink
// begins and ends, give or take a fraction of the line's height. Captures go
// to RUTBA_VERIFY_CAPTURE as pdf-find-mark-*.png. The fixture is a file of
// its own, in a folder of its own; the languages Recognise Text remembers
// are put back as they were. Run alone with RUTBA_VERIFY_ONLY=pdf-find-mark.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import pdfWriter from '@rutba/pdf';
import { decodePng } from '@rutba/ocr/png';

const { PdfDocument } = pdfWriter;
const ZOOMS = [1, 1.5, 2];

/** A page the suite drew, from @rutba/ocr's own test pages, as RGB samples. */
function scan(name) {
  const file = fileURLToPath(new URL(`../test/fixtures/${name}.png`, import.meta.resolve('@rutba/ocr')));
  const { width, height, data } = decodePng(fs.readFileSync(file));
  const rgb = Buffer.alloc(width * height * 3);
  for (let i = 0, o = 0; i < data.length; i += 4, o += 3) { rgb[o] = data[i]; rgb[o + 1] = data[i + 1]; rgb[o + 2] = data[i + 2]; }
  return { width, height, data: rgb };
}

function fixture(dir) {
  const own = path.join(dir, 'pdf-find-mark');
  fs.mkdirSync(own, { recursive: true });
  const file = path.join(own, 'find-marks.pdf');
  const doc = new PdfDocument({ title: 'Find marks', created: '2026-10-15T09:00:00Z' });
  const put = (name) => {
    const s = scan(name);
    const ref = doc.addImage({ ...s, colorSpace: 'DeviceRGB' });
    doc.addPage().image(ref, 36, 48, (s.width / 150) * 72, (s.height / 150) * 72);
  };
  put('eng-serif');
  const typed = doc.addPage();
  typed.text('This page was typed, not scanned, in a sans face.', 72, 110, { font: 'Helvetica', size: 18 });
  typed.text('And this one in Courier.', 72, 150, { font: 'Courier', size: 18 });
  put('urd-naskh');
  put('ara-naskh');
  fs.writeFileSync(file, doc.toBuffer());
  return file;
}

/**
 * The words looked for, on which page (1-based) and in what script; `or`,
 * others from the page to look for in their place when the reading did not
 * give the first exactly (Arabic is read less surely than English).
 */
const PROBES = [
  { word: 'Rutba', page: 1, script: 'English' },
  { word: 'invoice', page: 1, script: 'English' },
  { word: 'account', page: 1, script: 'English' },
  { word: 'keeping', page: 1, script: 'English' },
  { word: 'exactly', page: 1, script: 'English' },
  { word: 'sans', page: 2, script: 'typed' },
  { word: 'Courier', page: 2, script: 'typed' },
  { word: 'کمپیوٹر', page: 3, script: 'Urdu' },
  { word: 'تصویر', page: 3, script: 'Urdu' },
  { word: 'الكلمات', page: 4, script: 'Arabic', or: ['نهاية', 'الرسالة', 'مستندا'] },
  { word: 'الفاتورة', page: 4, script: 'Arabic' },
];

/**
 * The current match's mark and the word's ink under it: the mark's box (the
 * range Find highlights), and the run of dark columns on the drawn page in
 * the middle of that line that the mark overlaps, in the window's pixels.
 */
const MEASURE = `(() => {
  const h = globalThis.CSS?.highlights?.get('pdf-find-now');
  if (!h || !h.size) return null;
  const range = [...h][0];
  const rects = [...range.getClientRects()].filter((r) => r.width > 0.5 && r.height > 0.5);
  if (!rects.length) return null;
  const L = Math.min(...rects.map((r) => r.left)), R = Math.max(...rects.map((r) => r.right));
  const T = Math.min(...rects.map((r) => r.top)), B = Math.max(...rects.map((r) => r.bottom));
  const page = (range.startContainer.parentElement || range.startContainer).closest('.pdf-page');
  const canvas = page?.querySelector('canvas');
  if (!canvas) return { L, R, T, B, ink: null };
  const cr = canvas.getBoundingClientRect();
  const kx = canvas.width / cr.width, ky = canvas.height / cr.height;
  const ctx = canvas.getContext('2d');
  const hgt = B - T;
  const x0 = Math.max(0, Math.floor((L - 2 * hgt - cr.left) * kx)), x1 = Math.min(canvas.width, Math.ceil((R + 2 * hgt - cr.left) * kx));
  const y0 = Math.max(0, Math.floor((T + hgt * 0.2 - cr.top) * ky)), y1 = Math.min(canvas.height, Math.ceil((B - hgt * 0.2 - cr.top) * ky));
  if (x1 <= x0 || y1 <= y0) return { L, R, T, B, ink: null };
  const w = x1 - x0, hh = y1 - y0;
  const img = ctx.getImageData(x0, y0, w, hh).data;
  const ink = new Array(w).fill(false);
  for (let x = 0; x < w; x++) for (let y = 0; y < hh; y++) { const i = (y * w + x) * 4; if (img[i] + img[i + 1] + img[i + 2] < 420) { ink[x] = true; break; } }
  // Letters close together are one run; a word's space parts two.
  const gap = Math.max(2, 0.12 * hgt * kx);
  const runs = [];
  let s = -1, last = -1e9;
  for (let x = 0; x < w; x++) {
    if (!ink[x]) continue;
    if (s < 0 || x - last > gap) { if (s >= 0) runs.push([s, last + 1]); s = x; }
    last = x;
  }
  if (s >= 0) runs.push([s, last + 1]);
  const css = runs.map(([a, b]) => [cr.left + (x0 + a) / kx, cr.left + (x0 + b) / kx]);
  const over = css.filter(([a, b]) => Math.min(b, R) - Math.max(a, L) >= 0.3 * (b - a));
  const r1 = (v) => Math.round(v * 10) / 10;
  if (!over.length) return { L: r1(L), R: r1(R), height: r1(hgt), ink: null, text: range.toString() };
  return { L: r1(L), R: r1(R), height: r1(hgt), inkL: r1(Math.min(...over.map((r) => r[0]))), inkR: r1(Math.max(...over.map((r) => r[1]))), text: range.toString(), page: page.dataset.key };
})()`;

export async function verifyPdfFindMark({ open, check, until, wait }, { dir }) {
  const shots = process.env.RUTBA_VERIFY_CAPTURE;
  let js = null;
  let languagesWere;
  try {
    const file = fixture(dir);
    const win = await open('pdf', file);
    const wc = win.webContents;
    js = (code) => wc.executeJavaScript(code);
    const snap = async (name) => {
      if (!shots) return;
      wc.invalidate();
      await wait(700);
      fs.writeFileSync(path.join(shots, name), (await wc.capturePage()).toPNG());
    };
    const tip = (label) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-btn')].find((n) => (n.dataset.tip || n.title || '').startsWith(${JSON.stringify(label)}));
      if (!b) return 'no button ${label}';
      if (b.disabled) return 'disabled ${label}';
      b.click();
      return 'clicked';
    })()`);
    const setInput = (sel, value, proto = 'HTMLInputElement') => js(`(() => {
      const el = document.querySelector(${JSON.stringify(sel)});
      if (!el) return false;
      Object.getOwnPropertyDescriptor(${proto}.prototype, 'value').set.call(el, ${JSON.stringify(String(value))});
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
      return true;
    })()`);
    win.setContentSize(1300, 900);
    await until(() => js(`document.querySelectorAll('.pdf-thumb').length === 4 && document.querySelectorAll('.pdf-page canvas[data-drawn]').length > 0`), 'four pages drawn', 20000);
    languagesWere = await js(`localStorage.getItem('rutba-ocr-languages')`);

    // Recognise Text on the three scans, in English, Urdu and Arabic.
    await tip('Recognise Text');
    await until(() => js(`Boolean(document.querySelector('[data-role="ocr-dialog"]'))`), 'the Recognise Text dialog', 5000);
    await js(`(() => {
      for (const n of document.querySelectorAll('.ocr-lang')) {
        const want = ['eng', 'urd', 'ara'].includes(n.dataset.code);
        const box = n.querySelector('input');
        if (box.checked !== want) box.click();
      }
      return 1;
    })()`);
    await wait(200);
    await js(`document.querySelector('[data-role="ocr-go"]')?.click(), 1`);
    const read = await until(() => js(`!document.querySelector('[data-role="ocr-progress"]') && !document.querySelector('[data-role="ocr-dialog"]') && [...document.querySelectorAll('.rw-toast')].some((n) => /recognised/i.test(n.textContent))`), 'recognition to finish', 120000).then(() => true, () => false);
    check('pdf-find-mark: Recognise Text reads the English, Urdu and Arabic scans', read, read ? 'read' : 'never finished');

    await tip('Find');
    await until(() => js(`Boolean(document.querySelector('[data-role="find-input"]'))`), 'the find box', 4000);
    // Each probe as the reading gave it: the first of its words found on its page.
    for (const p of PROBES) {
      for (const word of [p.word, ...(p.or || [])]) {
        // Cleared first, so the count read is this word's and not the last one's.
        await setInput('[data-role="find-input"]', '');
        await until(() => js(`document.querySelector('.pdf-find-count')?.dataset.query === 'no'`), 'the find box cleared', 4000).catch(() => {});
        await setInput('[data-role="find-input"]', word);
        const COUNT = `(() => { const c = document.querySelector('.pdf-find-count'); return c && c.dataset.query === 'yes' && c.textContent ? Number(c.dataset.count) : null; })()`;
        await until(() => js(COUNT), `find ${word}`, 6000).catch(() => {});
        await wait(500);
        const n = (await js(COUNT)) || 0;
        if (n >= 1) { p.word = word; break; }
      }
    }
    const results = [];
    for (const zoom of ZOOMS) {
      await setInput('.rw-zoom input', zoom);
      await until(() => js(`(document.querySelector('.rw-zoom-pct')?.textContent || '').trim() === '${Math.round(zoom * 100)}%'`), `zoom ${zoom * 100}%`, 5000).catch(() => {});
      for (const p of PROBES) {
        await setInput('[data-role="find-input"]', '');
        await setInput('[data-role="find-input"]', p.word);
        // Settled: one match, current, on its page, with the page's drawing and words there to measure against.
        let seen = null;
        const m = await until(async () => {
          const v = await js(MEASURE);
          const at = await js(`document.querySelector('.pdf-thumb.current')?.dataset.page`);
          seen = v ? { text: v.text, ink: v.inkL != null, page: at } : { page: at };
          // Find folds the letters Arabic and Urdu share (ک and ك, ی and ي): the match is the word, letter for letter.
          return v && v.inkL != null && [...v.text].length === [...p.word].length && at === String(p.page - 1) ? v : null;
        }, `the mark on "${p.word}" at ${zoom * 100}%`, 10000).catch(() => null);
        await wait(250);
        const v = m ? (await js(MEASURE)) || m : null;
        // A recognised word's box is its ink; a typed word's runs from its first
        // letter's side bearing to its last's, a little wider than the ink.
        const tol = v ? Math.max(2.5, (p.script === 'typed' ? 0.2 : 0.12) * v.height) : 0;
        const dl = v ? Math.round((v.L - v.inkL) * 10) / 10 : null;
        const dr = v ? Math.round((v.R - v.inkR) * 10) / 10 : null;
        results.push({ ...p, seen: m ? null : seen, zoom, ok: Boolean(v) && Math.abs(dl) <= tol && Math.abs(dr) <= tol, dl, dr, tol: Math.round(tol * 10) / 10, height: v?.height });
        if (zoom === 1.5 || p.script === 'English') {
          if (p.word === 'invoice' || p.word === 'account' || p.word === 'sans' || p.word === 'تصویر' || p.word === 'الفاتورة') await snap(`pdf-find-mark-${p.script.toLowerCase()}-${p.page}-${zoom * 100}-${results.length}.png`);
        }
      }
    }
    const line = (r) => `${r.word} @${r.zoom * 100}%: ${r.dl == null ? `no mark (${JSON.stringify(r.seen)})` : `left ${r.dl}, right ${r.dr} (±${r.tol})`}`;
    for (const script of ['English', 'typed', 'Urdu', 'Arabic']) {
      const mine = results.filter((r) => r.script === script);
      const bad = mine.filter((r) => !r.ok);
      check(`pdf-find-mark: on the ${script === 'typed' ? 'typed page' : `${script} scan`}, Find's mark sits on the word it found at 100%, 150% and 200% (its edges where the word's ink begins and ends)`,
        mine.length && !bad.length, (bad.length ? bad : mine).slice(0, 6).map(line).join('; '));
    }
    console.log(`     [pdf-find-mark] ${results.map(line).join(' | ')}`);
    await js(`document.querySelector('[data-tip^="Close Find"]')?.click(), 1`);
    await setInput('.rw-zoom input', 1);
  } catch (err) {
    check('pdf-find-mark: the checks ran', false, err.stack || err.message);
  } finally {
    if (js) {
      await js(languagesWere == null ? `localStorage.removeItem('rutba-ocr-languages'), 1` : `localStorage.setItem('rutba-ocr-languages', ${JSON.stringify(languagesWere)}), 1`).catch(() => {});
    }
  }
}
