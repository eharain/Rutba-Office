// Pictures and Image: Copy Text from Picture, pressed in their windows. A
// picture of an English letter opens in Pictures; Copy Text reads it on the
// main process's threads, puts every word on the clipboard and lays the
// lines over the picture, where one line can be selected and copied alone;
// zoomed in, the lines stay on their words; the next picture leaves the last
// one's words behind; Languages offers the six with Urdu marked beta. In
// Image the same letter is read from the picture as edited, an edit takes
// the words away, and an Urdu letter read in Urdu comes out right to left.
// The clipboard's text is kept and put back.
// Run alone with RUTBA_VERIFY_ONLY=ocr-picture.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { clipboard } from 'electron';

const fixture = (name) => fileURLToPath(new URL(`../test/fixtures/${name}`, import.meta.resolve('@rutba/ocr')));

function driver(win, { until, wait }) {
  const wc = win.webContents;
  const js = (code) => wc.executeJavaScript(code);
  const capture = async (name) => {
    if (!process.env.RUTBA_VERIFY_CAPTURE) return;
    wc.invalidate();
    await wait(700);
    fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG());
  };
  const role = (r) => js(`(() => { const b = document.querySelector('[data-role="${r}"]'); if (!b) return 'no ${r}'; if (b.disabled) return 'disabled ${r}'; b.click(); return 'clicked'; })()`);
  /** The lines laid over the picture, and the picture's box, in window pixels. */
  const layer = (picture) => js(`(() => {
    const pic = document.querySelector(${JSON.stringify(picture)});
    const p = pic?.getBoundingClientRect();
    const lines = [...document.querySelectorAll('.ocr-layer .ocr-line')].map((n) => { const r = n.getBoundingClientRect(); return { text: n.textContent, dir: n.getAttribute('dir'), left: r.left, top: r.top, right: r.right, bottom: r.bottom }; });
    return { pic: p ? { left: p.left, top: p.top, right: p.right, bottom: p.bottom, width: p.width } : null, lines, bar: document.querySelector('[data-role="ocr-bar"]')?.dataset.state || null };
  })()`);
  const inside = (l) => l.pic && l.lines.length && l.lines.every((x) => x.left >= l.pic.left - 2 && x.right <= l.pic.right + 2 && x.top >= l.pic.top - 2 && x.bottom <= l.pic.bottom + 2);
  /** One recognised line selected by a range, as a drag across it selects it, and the text it copies. */
  const selectLine = (i) => js(`(() => {
    const n = document.querySelectorAll('.ocr-layer .ocr-line')[${i}];
    if (!n) return '';
    const r = document.createRange();
    r.selectNodeContents(n);
    getSelection().removeAllRanges();
    getSelection().addRange(r);
    const text = getSelection().toString();
    getSelection().removeAllRanges();
    return text;
  })()`);
  return { wc, js, capture, role, layer, inside, selectLine };
}

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args where the files are made
 */
export async function verifyOcrPicture(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const root = path.join(dir, 'ocr-picture');
  fs.mkdirSync(root, { recursive: true });
  const letter = path.join(root, 'a-letter.png');
  const other = path.join(root, 'b-hindi.png');
  const urdu = path.join(root, 'c-urdu.png');
  fs.copyFileSync(fixture('eng-serif.png'), letter);
  fs.copyFileSync(fixture('hin.png'), other);
  fs.copyFileSync(fixture('urd-naskh.png'), urdu);
  const truth = fs.readFileSync(fixture('eng-serif.txt'), 'utf8');
  const kept = clipboard.readText();

  try {
    // ── Pictures ──
    const win = await open('pictures', letter);
    const d = driver(win, { until, wait });
    await until(() => d.js(`Boolean(document.querySelector('.pv-image')?.complete && document.querySelector('.pv-image').naturalWidth)`), 'the picture shown', 15000).catch(() => {});
    await d.js(`localStorage.setItem('rutba-ocr-languages', '["eng"]'), 1`);
    clipboard.writeText('');
    const pressed = await d.role('copy-text');
    const ready = await until(() => d.js(`document.querySelector('[data-role="ocr-bar"]')?.dataset.state === 'ready'`), 'the words read', 30000).then(() => true).catch(() => false);
    const copied = clipboard.readText();
    const first = await d.layer('.pv-image');
    await d.capture('ocr-picture-01-pictures.png');
    check('ocr-picture: Copy Text in Pictures reads the picture and puts all its words on the clipboard, in order', pressed === 'clicked' && ready
      && copied.replace(/\s+/g, ' ').trim().startsWith('Rutba Office reads the words in a picture on this computer.') && /a total of 59\.20/.test(copied),
      JSON.stringify({ pressed, ready, copied: copied.slice(0, 120), want: truth.slice(0, 40) }));
    check('ocr-picture: the recognised lines lie over the picture, each inside it', first.lines.length >= 6 && d.inside(first), JSON.stringify({ lines: first.lines.length, pic: first.pic, firstLine: first.lines[0] }));
    const one = await d.selectLine(0);
    check('ocr-picture: one line selected over the picture copies that line alone', /^Rutba Office reads the words in a picture/.test(one) && !/invoice/.test(one), one);

    // Zoomed in, the lines follow the picture.
    await d.js(`[...document.querySelectorAll('.pv-nav .rw-btn')].find((n) => (n.dataset.tip || n.title || '').startsWith('Zoom in'))?.click(), 1`);
    await until(async () => { const l = await d.layer('.pv-image'); return l.pic && l.pic.width > first.pic.width * 1.1 && d.inside(l); }, 'the lines on the zoomed picture', 6000).catch(() => {});
    const zoomed = await d.layer('.pv-image');
    await d.capture('ocr-picture-02-zoomed.png');
    check('ocr-picture: zoomed in, the lines grow with the picture and stay on it', zoomed.pic && zoomed.pic.width > first.pic.width * 1.1 && d.inside(zoomed)
      && Math.abs((zoomed.lines[0].left - zoomed.pic.left) / zoomed.pic.width - (first.lines[0].left - first.pic.left) / first.pic.width) < 0.01,
      JSON.stringify({ before: first.pic, after: zoomed.pic }));

    // Languages: six, Urdu marked beta.
    await d.role('text-languages');
    await until(() => d.js(`Boolean(document.querySelector('[data-role="ocr-languages"]'))`), 'the languages dialog', 4000).catch(() => {});
    const langs = await d.js(`[...document.querySelectorAll('[data-role="ocr-languages"] .ocr-lang')].map((n) => n.dataset.code + (n.querySelector('.ocr-beta') ? '*' : '')).join()`);
    await d.capture('ocr-picture-03-languages.png');
    await d.js(`[...document.querySelectorAll('.rw-dialog .rw-btn')].find((n) => n.textContent.trim() === 'Cancel')?.click(), 1`);
    check('ocr-picture: Languages offers the six languages Copy Text reads, Urdu marked beta', langs === 'eng,urd*,ara,hin,chi_sim,jpn', langs);

    // The next picture: the words of the last one are gone.
    await d.js(`[...document.querySelectorAll('.rw-btn')].find((n) => (n.dataset.tip || n.title || '') === 'Next')?.click(), 1`);
    const cleared = await until(() => d.js(`/b-hindi/.test(document.querySelector('.pv-image')?.alt || '') && !document.querySelector('.ocr-layer') && !document.querySelector('[data-role="ocr-bar"]')`), 'the next picture without words', 6000).then(() => true).catch(() => false);
    check('ocr-picture: the next picture leaves the last one\'s words behind', cleared, String(cleared));
    const pErrors = await errorsIn(win);

    // ── Image ──
    const img = await open('image', letter);
    const e = driver(img, { until, wait });
    await until(() => e.js(`(document.querySelector('.im-canvas')?.width || 0) > 100`), 'the picture drawn', 15000).catch(() => {});
    clipboard.writeText('');
    const ipressed = await e.role('copy-text');
    const iready = await until(() => e.js(`document.querySelector('[data-role="ocr-bar"]')?.dataset.state === 'ready'`), 'the words read in Image', 30000).then(() => true).catch(() => false);
    const icopied = clipboard.readText();
    const ilayer = await e.layer('.im-canvas');
    const ione = await e.selectLine(1);
    await e.capture('ocr-picture-04-image.png');
    check('ocr-picture: Copy Text in Image reads the picture as edited, copies its words and lays its lines over it', ipressed === 'clicked' && iready && /Rutba Office reads the words/.test(icopied)
      && ilayer.lines.length >= 6 && e.inside(ilayer) && ione.length > 10, JSON.stringify({ ipressed, iready, lines: ilayer.lines.length, ione }));
    await e.js(`[...document.querySelectorAll('.rw-btn')].find((n) => (n.dataset.tip || n.title || n.textContent || '').trim().startsWith('Right'))?.click(), 1`);
    const gone = await until(() => e.js(`!document.querySelector('.ocr-layer')`), 'the words gone after an edit', 6000).then(() => true).catch(() => false);
    check('ocr-picture: an edit to the picture takes its recognised words away', gone, String(gone));
    const iErrors = await errorsIn(img);

    // Urdu, read in Urdu: right to left, with its own letters.
    const uw = await open('image', urdu);
    const u = driver(uw, { until, wait });
    await until(() => u.js(`(document.querySelector('.im-canvas')?.width || 0) > 100`), 'the Urdu picture drawn', 15000).catch(() => {});
    await u.role('text-languages');
    await until(() => u.js(`Boolean(document.querySelector('[data-role="ocr-languages"]'))`), 'the languages dialog', 4000).catch(() => {});
    await u.js(`(() => { for (const n of document.querySelectorAll('[data-role="ocr-languages"] .ocr-lang')) { const box = n.querySelector('input'); if (box.checked !== (n.dataset.code === 'urd')) box.click(); } return 1; })()`);
    const note = await u.js(`document.querySelector('[data-role="ocr-urdu-note"]')?.textContent || ''`);
    await u.role('ocr-languages-ok');
    clipboard.writeText('');
    await u.role('copy-text');
    const uready = await until(() => u.js(`document.querySelector('[data-role="ocr-bar"]')?.dataset.state === 'ready'`), 'the Urdu words read', 30000).then(() => true).catch(() => false);
    const ucopied = clipboard.readText();
    const ulayer = await u.layer('.im-canvas');
    await u.capture('ocr-picture-05-urdu.png');
    check('ocr-picture: an Urdu picture read in Urdu (marked beta) copies Urdu words, its lines right to left', uready && /کمپیوٹر/.test(ucopied) && /beta/i.test(note)
      && ulayer.lines.length >= 4 && ulayer.lines.every((l) => l.dir === 'rtl') && u.inside(ulayer), JSON.stringify({ uready, ucopied: ucopied.slice(0, 60), lines: ulayer.lines.length }));
    const uErrors = await errorsIn(uw);
    const errors = [...pErrors, ...iErrors, ...uErrors];
    check('ocr-picture: no error was shown', !errors.length, errors.join(' | '));
    await u.js(`localStorage.removeItem('rutba-ocr-languages'), 1`).catch(() => {});
  } finally {
    clipboard.writeText(kept || '');
  }
}
