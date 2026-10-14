// PDF: redaction that removes, pressed in the window and proved from the
// file it writes. A two-page letter made here carries an account code on
// both pages, a phone number and a photograph. The Redact tab says that
// marks remove nothing until they are applied; the phone number is marked
// by selecting it, every match of the account code by Mark Matches,
// and the photograph by dragging out an area; Apply Redactions warns that
// the file is written whole as a new file and why, and writes it; the window
// opens the redacted file, where Find no longer finds the code. Read back,
// the redacted words are found neither by this package's reader nor by a
// plain search of every byte and every inflated stream, the words beside
// them are kept, the photograph is painted out under its mark, and the
// original file is as it was.
// Run alone with RUTBA_VERIFY_ONLY=pdf-redact.

import fs from 'node:fs';
import path from 'node:path';
import pdfWriter from '@rutba/pdf';
import pdfReader from '@rutba/pdf/reader';
import pdfRedact from '@rutba/pdf/redact';

const { PdfDocument } = pdfWriter;
const { openPdf, Stream } = pdfReader;
const { bytesContain } = pdfRedact;

function makeLetter(file) {
  const doc = new PdfDocument({ title: 'Client letter', created: '2026-10-14T09:00:00Z' });
  // A photograph: a soft gradient, 120 by 80, drawn 240 by 160 points.
  const w = 120;
  const h = 80;
  const rgb = Buffer.alloc(w * h * 3);
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) { const o = (y * w + x) * 3; rgb[o] = 90 + x; rgb[o + 1] = 120 + y; rgb[o + 2] = 200; }
  const photo = doc.addImage({ width: w, height: h, colorSpace: 'DeviceRGB', data: rgb });
  const one = doc.addPage();
  one.text('Client letter', 72, 100, { font: 'Helvetica-Bold', size: 26 });
  one.text('Account SECRET-4417 belongs to the client.', 72, 150, { font: 'Helvetica', size: 14 });
  one.text('Call 0300-1234567 about the account.', 72, 180, { font: 'Helvetica', size: 14 });
  one.text('The photograph below is private.', 72, 210, { font: 'Helvetica', size: 14 });
  one.image(photo, 72, 240, 240, 160);
  const two = doc.addPage();
  two.text('Second page', 72, 100, { font: 'Helvetica-Bold', size: 26 });
  two.text('Again SECRET-4417 is named here.', 72, 150, { font: 'Helvetica', size: 14 });
  fs.writeFileSync(file, doc.toBuffer());
}

function driver(win, { until, wait }) {
  const wc = win.webContents;
  const js = (code) => wc.executeJavaScript(code);
  const capture = async (name) => {
    if (!process.env.RUTBA_VERIFY_CAPTURE) return;
    wc.invalidate();
    await wait(700);
    fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG());
  };
  const tab = (name) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(), 1`);
  const role = async (r) => {
    const sel = `[data-role="${r}"]`;
    await until(() => js(`Boolean(document.querySelector(${JSON.stringify(sel)}) && !document.querySelector(${JSON.stringify(sel)}).disabled)`), `the ${r} button`, 5000).catch(() => {});
    return js(`(() => { const b = document.querySelector(${JSON.stringify(sel)}); if (!b) return 'no button ${r}'; if (b.disabled) return 'disabled ${r}'; b.click(); return 'clicked'; })()`);
  };
  const setInput = (sel, value) => js(`(() => {
    const el = document.querySelector(${JSON.stringify(sel)});
    if (!el) return false;
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, ${JSON.stringify(String(value))});
    el.dispatchEvent(new Event('input', { bubbles: true }));
    return true;
  })()`);
  const select = async (index, from, to) => {
    await wait(250);
    await until(() => js(`Boolean(document.querySelector('.pdf-page[data-index="${index}"] .textLayer[data-ready]'))`), 'the words laid over the page', 8000).catch(() => {});
    return js(`(() => {
      const layer = document.querySelector('.pdf-page[data-index="${index}"] .textLayer[data-ready]');
      if (!layer) return 'no text';
      const nodes = [];
      const walker = document.createTreeWalker(layer, NodeFilter.SHOW_TEXT);
      for (let n = walker.nextNode(); n; n = walker.nextNode()) nodes.push(n);
      const a = nodes.find((n) => n.data.includes(${JSON.stringify(from)}));
      const b = nodes.find((n) => n.data.includes(${JSON.stringify(to)}));
      if (!a || !b) return 'no words';
      const r = document.createRange();
      r.setStart(a, a.data.indexOf(${JSON.stringify(from)}));
      r.setEnd(b, b.data.indexOf(${JSON.stringify(to)}) + ${JSON.stringify(to)}.length);
      getSelection().removeAllRanges();
      getSelection().addRange(r);
      return getSelection().toString();
    })()`);
  };
  const dragOn = (index, points) => js(`(async () => {
    const svg = document.querySelector('.pdf-page[data-index="${index}"] .pdf-annots');
    if (!svg) return 'no layer';
    const r = svg.getBoundingClientRect();
    const at = ([fx, fy]) => ({ clientX: r.left + r.width * fx, clientY: r.top + r.height * fy, bubbles: true, cancelable: true, button: 0, buttons: 1 });
    const pts = ${JSON.stringify(points)};
    svg.dispatchEvent(new MouseEvent('mousedown', at(pts[0])));
    for (const p of pts.slice(1)) { window.dispatchEvent(new MouseEvent('mousemove', at(p))); await new Promise((res) => setTimeout(res, 20)); }
    window.dispatchEvent(new MouseEvent('mouseup', at(pts[pts.length - 1])));
    return 'dragged';
  })()`);
  const marks = () => js(`[...document.querySelectorAll('.pdf-page')].map((p) => p.querySelectorAll('.pdf-annot[data-type="Redact"]').length)`);
  return { wc, js, capture, tab, role, setInput, select, dragOn, marks };
}

const answer = (kind, value) => {
  (globalThis.__rutbaCheckDialogAnswers ||= []).push({ kind, answer: value });
};

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args where the files are made (a folder of this block's own is made in it)
 */
export async function verifyPdfRedact(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const root = path.join(dir, 'pdf-redact');
  fs.mkdirSync(root, { recursive: true });
  const file = path.join(root, 'client-letter.pdf');
  makeLetter(file);
  const original = fs.readFileSync(file);

  const win = await open('pdf', file);
  const d = driver(win, { until, wait });
  await until(() => d.js(`document.querySelectorAll('.pdf-page canvas[data-drawn]').length >= 1 && Boolean(document.querySelector('.pdf-page[data-index="0"] .pdf-annots[data-ready]'))`), 'the first page drawn', 20000).catch(() => {});
  const total = async () => (await d.marks()).reduce((a, b) => a + b, 0);

  await d.tab('Redact');
  await until(() => d.js(`Boolean(document.querySelector('[data-role="redact-notice"]'))`), 'the notice', 4000).catch(() => {});
  const notice = await d.js(`document.querySelector('[data-role="redact-notice"]')?.textContent || ''`);
  check('pdf-redact: the Redact tab says marks remove nothing until applied, and that applying saves a new file', /nothing is removed until you apply/.test(notice) && /new file/.test(notice), notice);

  // The phone number, selected and marked.
  const selected = await d.select(0, '0300', '1234567');
  const marked = await d.role('redact-text');
  const one = await until(async () => (await total()) === 1, 'the first mark', 10000).then(() => true).catch(() => false);
  check('pdf-redact: words selected are marked for redaction with Mark Text', selected === '0300-1234567' && marked === 'clicked' && one, JSON.stringify({ selected, marked, marks: await d.marks() }));

  // Every match of the code, on both pages, from Find.
  await d.js(`(document.querySelector('.pdf-view').focus(), 1)`);
  const noFind = await d.role('redact-search');
  await until(() => d.js(`Boolean(document.querySelector('[data-role="find-input"]'))`), 'the find box', 4000).catch(() => {});
  await d.setInput('[data-role="find-input"]', 'SECRET-4417');
  await until(() => d.js(`document.querySelector('.pdf-find-count')?.dataset.count === '2'`), 'two matches', 8000).catch(() => {});
  await d.role('redact-search');
  const searched = await until(async () => (await total()) === 3, 'both matches marked', 12000).then(() => true).catch(() => false);
  const perPage = await d.marks();
  check('pdf-redact: Mark Matches marks every match of the words in Find, on every page', noFind === 'clicked' && searched && perPage[0] === 2 && perPage[1] === 1, JSON.stringify({ noFind, perPage }));
  await d.js(`document.querySelector('[data-tip^="Close Find"]')?.click(), 1`);

  // The photograph, by an area dragged out over its middle.
  await d.role('redact-area');
  await d.dragOn(0, [[0.2, 0.33], [0.3, 0.38], [0.4, 0.42]]);
  const area = await until(async () => (await total()) === 4, 'the area mark', 10000).then(() => true).catch(() => false);
  await d.js(`(document.querySelector('[data-role="redact-area"]').click(), 1)`);
  await d.capture('pdf-redact-01-marked.png');
  check('pdf-redact: an area of the page, over the photograph, is marked with Mark Area', area, JSON.stringify(await d.marks()));

  // Apply: the warning, then a new file, opened here.
  const target = path.join(root, 'client-letter_Redacted.pdf');
  await d.role('redact-apply');
  await until(() => d.js(`Boolean(document.querySelector('[data-role="redact-dialog"]'))`), 'the Apply dialog', 5000).catch(() => {});
  const warning = await d.js(`document.querySelector('[data-role="redact-warning"]')?.textContent || ''`);
  const lead = await d.js(`document.querySelector('.pdf-redact-lead')?.textContent || ''`);
  await d.capture('pdf-redact-02-dialog.png');
  check('pdf-redact: Apply Redactions says how many marks it applies, and warns that the file is written whole as a new one because an update keeps the old bytes', /4 marks on 2 pages/.test(lead) && /new file/.test(warning) && /written whole/.test(warning) && /keeps the original's bytes/.test(warning),
    JSON.stringify({ lead, warning }));
  answer('save', target);
  await d.role('redact-go');
  const wrote = await until(() => fs.existsSync(target), 'the redacted file', 15000).then(() => true).catch(() => false);
  const opened = await until(() => d.js(`[...document.querySelectorAll('.pdf-thumb')].length === 2 && [...document.querySelectorAll('.pdf-thumb')].every((n) => n.dataset.from === 'client-letter_Redacted.pdf')`), 'the redacted file in the window', 15000).then(() => true).catch(() => false);
  await until(() => d.js(`Boolean(document.querySelector('.pdf-page[data-index="0"] .textLayer[data-ready]'))`), 'its words', 10000).catch(() => {});
  const layerText = await d.js(`[...document.querySelectorAll('.pdf-page .textLayer')].map((l) => l.textContent).join(' ')`);
  await d.tab('Home');
  await d.js(`[...document.querySelectorAll('.rw-btn')].find((n) => (n.dataset.tip || '').startsWith('Find'))?.click(), 1`);
  await until(() => d.js(`Boolean(document.querySelector('[data-role="find-input"]'))`), 'the find box', 4000).catch(() => {});
  await d.setInput('[data-role="find-input"]', 'SECRET');
  await until(() => d.js(`(document.querySelector('.pdf-find-count')?.textContent || '') !== ''`), 'find', 6000).catch(() => {});
  await wait(400);
  const found = await d.js(`Number(document.querySelector('.pdf-find-count')?.dataset.count || 0)`);
  await d.js(`document.querySelector('[data-tip^="Close Find"]')?.click(), 1`);
  await d.capture('pdf-redact-03-applied.png');
  check('pdf-redact: the redacted file opens in the window, where pdf.js finds the redacted words nowhere and the words beside them stay', wrote && opened && found === 0 && !/SECRET|1234567/.test(layerText) && /belongs to the client/.test(layerText) && /about the account/.test(layerText),
    JSON.stringify({ wrote, opened, found, layerText: layerText.slice(0, 240) }));

  // Read back: this package's reader, a plain search of the bytes, the photograph's pixels, the original.
  let back = {};
  if (wrote) {
    const bytes = fs.readFileSync(target);
    const pdf = openPdf(bytes);
    const text = pdf.pageText(0) + pdf.pageText(1);
    const res = pdf.resolve(pdf.pages()[0].dict.get('Resources'));
    const xo = pdf.resolve(res?.get('XObject'));
    const pics = xo ? [...xo.keys()].map((k) => pdf.resolve(xo.get(k))).filter((s) => s instanceof Stream && s.dict.get('Subtype')?.name === 'Image') : [];
    const px = pics.length ? pdf.decode(pics[0]) : Buffer.alloc(0);
    let black = 0;
    for (let i = 0; i + 2 < px.length; i += 3) if (px[i] === 0 && px[i + 1] === 0 && px[i + 2] === 0) black++;
    back = {
      readerClean: !/SECRET|1234567/.test(text),
      kept: /belongs to the client/.test(text) && /Again/.test(text),
      bytesClean: !bytesContain(bytes, 'SECRET-4417') && !bytesContain(bytes, '0300-1234567') && !bytesContain(bytes, '1234567'),
      whole: !bytes.subarray(0, original.length).equals(original),
      pictures: pics.length,
      black,
      originalKept: fs.readFileSync(file).equals(original) && bytesContain(original, 'SECRET-4417'),
      marksGone: !/\/Redact/.test(bytes.toString('latin1')),
    };
  }
  check('pdf-redact: read back, the redacted words are found neither by the reader nor in any byte or stream, the photograph is painted out under its mark, the marks are gone and the original is as it was',
    wrote && back.readerClean && back.kept && back.bytesClean && back.whole && back.pictures === 1 && back.black > 50 && back.black < 120 * 80 && back.originalKept && back.marksGone, JSON.stringify(back));
  const errors = await errorsIn(win);
  check('pdf-redact: no error was shown', !errors.length, errors.join(' | '));
}
