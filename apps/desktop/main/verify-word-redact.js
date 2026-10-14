// Documents: Review → Redact. The pane marks a name (Find, Mark All), the
// e-mail and telephone patterns and a pattern of one's own; the marks show
// on the page as they are made and are counted in every part of the file
// (the body, headers, footers, notes, comments, properties); an area of the
// picture is marked by dragging over it; Apply saves a copy under a new
// name and says so, the open document left as it was; the copy read back
// from the disk holds none of the marked words in any part, the picture's
// area is black, and the copy opens with black bars where the words were.
// Run alone with RUTBA_VERIFY_ONLY=word-redact.

import fs from 'node:fs';
import path from 'node:path';
import { nativeImage, BrowserWindow } from 'electron';
import { OoxmlPackage } from '@rutba/ooxml';
import { partsHolding } from '@rutba/ooxml/redact';
import { secretLetter } from './verify-review-fixtures.js';

/**
 * @param {object} h the harness: open, check, until, wait, capture, doc
 */
export async function verifyWordRedact(h, { dir }) {
  const { open, check, wait, capture, doc } = h;
  // Waits as the harness's until does, and answers what the condition found once it held.
  const until = async (fn, what, ms) => { let found; await h.until(async () => { found = await fn(); return found; }, what, ms); return found; };
  const own = path.join(dir, 'word-redact');
  const file = path.join(own, 'offer.docx');
  const copy = path.join(own, 'offer (redacted).docx');
  let original;
  try {
    fs.mkdirSync(own, { recursive: true });
    original = secretLetter();
    fs.writeFileSync(file, original);
  } catch (err) {
    return check('word-redact: the fixture was made', false, err.message);
  }
  try {
    const win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const session = doc.sessions().filter((s) => s.kind === 'doc' && s.path === file).pop();
    const fresh = async (w, name) => {
      w.webContents.invalidate();
      await wait(700);
      await capture(w, name);
    };
    const setInput = (sel, value) => js(`(() => { const el = document.querySelector(${JSON.stringify(sel)}); if (!el) return false; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, ${JSON.stringify(value)}); el.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
    const click = (sel) => js(`(() => { const b = document.querySelector(${JSON.stringify(sel)}); if (!b) return 'none'; if (b.disabled) return 'disabled'; b.click(); return 'clicked'; })()`);

    await until(() => js(`Boolean(document.querySelector('.wd-page img.wd-image'))`), 'the letter and its picture', 10000);
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Review')?.click(), 1`);
    await wait(250);
    const pressed = await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || n.title || '').startsWith('Redact')); if (!b) return 'no button'; b.click(); return 'clicked'; })()`);
    const pane = await until(() => js(`Boolean(document.querySelector('.rd-pane .rd-find'))`), 'the Redact pane', 5000).catch(() => false);

    // 1. Marks: the name by Find, the two patterns, a pattern of one's own.
    await setInput('.rd-find', 'Ayesha Khan');
    await click('.rd-find-mark');
    await js(`document.querySelector('.rd-pattern-email')?.click(), 1`);
    await js(`document.querySelector('.rd-pattern-phone')?.click(), 1`);
    await setInput('.rd-custom', 'INV-\\d{4}');
    await click('.rd-custom-add');
    const marked = await until(() => js(`document.querySelectorAll('.rd-mark').length === 4`), 'four marks', 4000).catch(() => false);
    const lit = await until(() => js(`document.querySelectorAll('.rd-hit').length >= 5`), 'the marks on the page', 4000).catch(() => false);
    const litText = await js(`[...new Set([...document.querySelectorAll('.rd-hit')].map((n) => n.dataset.text))]`);
    const where = await until(() => js(`(() => { const k = [...document.querySelectorAll('.rd-where-parts span')].map((s) => s.dataset.kind); return ['body', 'headers', 'footers', 'notes', 'comments', 'properties'].every((x) => k.includes(x)) ? k : null; })()`), 'the parts counted', 6000).catch(() => null);
    check(
      'word-redact: Review → Redact marks a name found by Find, the e-mail and telephone patterns and a pattern of one\'s own, shows each on the page as it is made, and counts them in the body, headers, footers, notes, comments and properties',
      pressed === 'clicked' && pane === true && marked === true && lit === true && litText.some((s) => /Ayesha\s*Khan/.test(s)) && litText.some((s) => s === 'ayesha.khan@example.pk') && litText.includes('INV-2026') && Array.isArray(where),
      JSON.stringify({ pressed, pane, marked, lit, litText: litText.slice(0, 8), where }),
    );

    // 2. An area of the picture, dragged over it.
    await click('.rd-pick-area');
    await wait(200);
    const r = await js(`(() => { const i = document.querySelector('.wd-page img.wd-image'); i.scrollIntoView({ block: 'center' }); const q = i.getBoundingClientRect(); return { left: q.left, top: q.top, width: q.width, height: q.height }; })()`);
    await wait(300);
    const r2 = await js(`(() => { const q = document.querySelector('.wd-page img.wd-image').getBoundingClientRect(); return { left: q.left, top: q.top, width: q.width, height: q.height }; })()`);
    const p0 = { x: Math.round(r2.left + r2.width * 0.6), y: Math.round(r2.top + r2.height * 0.1) };
    const p1 = { x: Math.round(r2.left + r2.width * 0.9), y: Math.round(r2.top + r2.height * 0.45) };
    wc.sendInputEvent({ type: 'mouseDown', x: p0.x, y: p0.y, button: 'left', clickCount: 1 });
    for (let k = 1; k <= 6; k++) {
      wc.sendInputEvent({ type: 'mouseMove', x: Math.round(p0.x + ((p1.x - p0.x) * k) / 6), y: Math.round(p0.y + ((p1.y - p0.y) * k) / 6), button: 'left', modifiers: ['leftButtonDown'] });
      await wait(20);
    }
    wc.sendInputEvent({ type: 'mouseUp', x: p1.x, y: p1.y, button: 'left', clickCount: 1 });
    const area = await until(() => js(`document.querySelectorAll('.rd-mark').length === 5 && document.querySelectorAll('.rd-area:not(.live)').length === 1`), 'the picture area', 4000).catch(() => false);
    await click('.rd-pick-area');
    await fresh(win, 'word-redact-marks.png');
    check('word-redact: dragging over a picture marks an area of it, drawn over the picture', area === true && Boolean(r && r.width > 50), JSON.stringify({ area, r }));

    // 3. Apply: a copy saved under a new name, the open document left as it was.
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: copy });
    await click('.rd-apply-btn');
    const done = await until(() => js(`Boolean(document.querySelector('.rd-done'))`), 'the copy saved', 20000).catch(() => false);
    const said = await js(`[...document.querySelectorAll('.rw-toast')].map((n) => n.textContent).join(' | ')`);
    const untouched = fs.readFileSync(file).equals(original) && doc.meta({ id: session.id }).dirty === false;
    check(
      'word-redact: Apply saves the redacted copy as a new file and says so, leaving the open document and its file as they were',
      done === true && fs.existsSync(copy) && /redacted copy/i.test(said) && untouched,
      JSON.stringify({ done, exists: fs.existsSync(copy), said: said.slice(0, 160), untouched }),
    );

    // 4. The proof: the copy read back from the disk.
    const bytes = fs.readFileSync(copy);
    const left = partsHolding(bytes, ['Ayesha Khan', 'Ayesha', 'ayesha.khan@example.pk', 'ayesha.k@old.example.pk', '300 1234567', 'INV-2026']);
    const pkg = OoxmlPackage.read(bytes);
    const body = pkg.text('word/document.xml');
    const bars = /Dear ███████<\/w:t>[\s\S]*?<w:t>████<\/w:t>/.test(body) && /██████████████████████ and ███████████████ about ████████\./.test(body);
    const img = nativeImage.createFromBuffer(pkg.read('word/media/image1.png'));
    const size = img.getSize();
    const bmp = img.toBitmap();
    const at = (fx, fy) => { const x = Math.floor(size.width * fx); const y = Math.floor(size.height * fy); const o = (y * size.width + x) * 4; return [bmp[o], bmp[o + 1], bmp[o + 2]]; };
    const inside = at(0.75, 0.27);
    const outside = at(0.2, 0.2);
    const black = inside.every((v) => v < 16) && !outside.every((v) => v < 16);
    check(
      'word-redact: the copy read back holds none of the marked words in any part, body, header, footer, footnote, comment and its author, tracked changes, the link\'s address, alt text, properties, black bars of the same length stand in their place, and the picture\'s area is black',
      left.length === 0 && bars && black,
      JSON.stringify({ left, bars, inside, outside, size }),
    );

    // 5. The copy opens with its bars.
    await click('.rd-open-copy');
    const opened = await until(() => doc.sessions().some((s) => s.kind === 'doc' && s.path === copy), 'the copy opened', 10000).then(() => true).catch(() => false);
    const s2 = doc.sessions().filter((s) => s.kind === 'doc' && s.path === copy).pop();
    const text = s2 ? (doc.model({ id: s2.id }).blocks || []).map((b) => b.text).join('\n') : '';
    check('word-redact: Open the Copy opens the redacted document, its words black bars where the name was', opened && /Dear ███████████, we write to ██████████████████████/.test(text), text.slice(0, 160));
    const w2 = BrowserWindow.getAllWindows().filter((w) => !w.isDestroyed()).sort((a, b) => b.id - a.id)[0];
    if (w2) await fresh(w2, 'word-redact-copy.png');
  } catch (err) {
    check('word-redact: the Redact checks ran', false, err.stack || err.message);
  }
}

