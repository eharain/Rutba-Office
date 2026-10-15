// Presentations: Review → Redact and File → Info → Inspect Presentation.
// A name and an address are marked and shown on the slide; an area of the
// picture is marked by dragging over it; Apply saves a copy, and the copy
// read back holds the words in no part (slides, notes, comments, alt text,
// properties) with the picture's area black. Then the inspector finds the
// comment, the notes, the hidden and the parked shapes and the authors,
// and Remove All takes them out of the open deck. Run alone with
// RUTBA_VERIFY_ONLY=deck-redact.

import fs from 'node:fs';
import path from 'node:path';
import { nativeImage } from 'electron';
import { OoxmlPackage } from '@rutba/ooxml';
import { partsHolding } from '@rutba/ooxml/redact';
import { Deck } from '@rutba/presentation';
import { secretDeck } from './verify-review-fixtures.js';

/**
 * @param {object} h the harness: open, check, until, wait, capture, doc
 */
export async function verifyDeckRedact(h, { dir }) {
  const { open, check, wait, capture, doc } = h;
  // Waits as the harness's until does, and answers what the condition found once it held.
  const until = async (fn, what, ms) => { let found; await h.until(async () => { found = await fn(); return found; }, what, ms); return found; };
  const own = path.join(dir, 'deck-redact');
  const file = path.join(own, 'offer.pptx');
  const copy = path.join(own, 'offer (redacted).pptx');
  try {
    fs.mkdirSync(own, { recursive: true });
    fs.writeFileSync(file, await secretDeck());
  } catch (err) {
    return check('deck-redact: the fixture was made', false, err.message);
  }
  try {
    const win = await open('slides', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const session = doc.sessions().filter((s) => s.kind === 'deck' && s.path === file).pop();
    const fresh = async (name) => { wc.invalidate(); await wait(700); await capture(win, name); };
    const setInput = (sel, value) => js(`(() => { const el = document.querySelector(${JSON.stringify(sel)}); if (!el) return false; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, ${JSON.stringify(value)}); el.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
    const click = (sel) => js(`(() => { const b = document.querySelector(${JSON.stringify(sel)}); if (!b) return 'none'; if (b.disabled) return 'disabled'; b.click(); return 'clicked'; })()`);

    await until(() => js(`Boolean(document.querySelector('.sl-slide [data-shape] image'))`), 'the slide and its picture', 10000);
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Review')?.click(), 1`);
    await wait(250);
    const pressed = await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || n.title || '').startsWith('Redact')); if (!b) return 'no button'; b.click(); return 'clicked'; })()`);
    await until(() => js(`Boolean(document.querySelector('.rd-pane .rd-find'))`), 'the Redact pane', 5000).catch(() => {});
    await setInput('.rd-find', 'Ayesha Khan');
    await click('.rd-find-mark');
    await js(`document.querySelector('.rd-pattern-email')?.click(), 1`);
    const lit = await until(() => js(`(() => { const t = [...document.querySelectorAll('.rd-hit')].map((n) => n.dataset.text); return t.some((x) => /Ayesha Khan/.test(x)) && t.some((x) => x === 'ayesha.khan@example.pk') ? t : null; })()`), 'the marks on the slide', 6000).catch(() => null);
    const where = await until(() => js(`(() => { const k = [...document.querySelectorAll('.rd-where-parts span')].map((s) => s.dataset.kind); return ['slides', 'notes', 'comments', 'properties'].every((x) => k.includes(x)) ? k : null; })()`), 'the parts counted', 6000).catch(() => null);
    check(
      'deck-redact: Review → Redact marks a name and the e-mail pattern, outlines each on the slide as it is made, and counts them on the slides, in the notes, the comments and the properties',
      pressed === 'clicked' && Array.isArray(lit) && Array.isArray(where),
      JSON.stringify({ pressed, lit, where }),
    );

    // An area of the picture.
    await click('.rd-pick-area');
    await wait(200);
    const r = await js(`(() => { const i = document.querySelector('.sl-slide [data-shape] image'); const q = i.getBoundingClientRect(); return { left: q.left, top: q.top, width: q.width, height: q.height }; })()`);
    const p0 = { x: Math.round(r.left + r.width * 0.6), y: Math.round(r.top + r.height * 0.1) };
    const p1 = { x: Math.round(r.left + r.width * 0.9), y: Math.round(r.top + r.height * 0.45) };
    wc.sendInputEvent({ type: 'mouseDown', x: p0.x, y: p0.y, button: 'left', clickCount: 1 });
    for (let k = 1; k <= 6; k++) {
      wc.sendInputEvent({ type: 'mouseMove', x: Math.round(p0.x + ((p1.x - p0.x) * k) / 6), y: Math.round(p0.y + ((p1.y - p0.y) * k) / 6), button: 'left', modifiers: ['leftButtonDown'] });
      await wait(20);
    }
    wc.sendInputEvent({ type: 'mouseUp', x: p1.x, y: p1.y, button: 'left', clickCount: 1 });
    const area = await until(() => js(`document.querySelectorAll('.rd-area:not(.live)').length === 1`), 'the picture area', 4000).catch(() => false);
    await click('.rd-pick-area');
    await fresh('deck-redact-marks.png');

    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: copy });
    await click('.rd-apply-btn');
    const done = await until(() => js(`Boolean(document.querySelector('.rd-done'))`), 'the copy saved', 20000).catch(() => false);
    const bytes = fs.existsSync(copy) ? fs.readFileSync(copy) : null;
    const left = bytes ? partsHolding(bytes, ['Ayesha Khan', 'ayesha.khan@example.pk', 'Ayesha']) : ['no copy'];
    let black = false;
    let probe = null;
    if (bytes) {
      const pkg = OoxmlPackage.read(bytes);
      const media = pkg.partNames().find((n) => /^ppt\/media\/.*\.png$/.test(n));
      const img = nativeImage.createFromBuffer(pkg.read(media));
      const size = img.getSize();
      const bmp = img.toBitmap();
      const at = (fx, fy) => { const o = (Math.floor(size.height * fy) * size.width + Math.floor(size.width * fx)) * 4; return [bmp[o], bmp[o + 1], bmp[o + 2]]; };
      probe = { inside: at(0.75, 0.27), outside: at(0.2, 0.2) };
      black = probe.inside.every((v) => v < 16) && !probe.outside.every((v) => v < 16);
    }
    const copyDeck = bytes ? Deck.open(bytes) : null;
    check(
      'deck-redact: Apply saves a copy; read back, no part of it holds the marked words, slides, notes, comment, alt text, properties, black bars stand in their place on the slide, and the picture\'s area is black',
      area === true && done === true && left.length === 0 && black && /For ███████████/.test(JSON.stringify(copyDeck?.slide(0).shapes.map((s) => s.text))),
      JSON.stringify({ area, done, left, probe }),
    );

    // File → Info → Check for Issues → Inspect Presentation.
    await js(`document.querySelector('.rw-appmark')?.click(), 1`);
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim().startsWith('Info')))`), 'the app menu', 4000).catch(() => {});
    await js(`[...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim().startsWith('Info'))?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.pw-card[data-card="inspect"]'))`), 'the Inspect card', 5000).catch(() => {});
    await js(`document.querySelector('.pw-card[data-card="inspect"] .pw-card-tile')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.pw-menu [data-item="inspect"]'))`), 'Check for Issues', 3000).catch(() => {});
    await js(`document.querySelector('.pw-menu [data-item="inspect"]')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.di-inspect'))`), 'the inspector', 6000).catch(() => {});
    await js(`document.querySelector('.di-inspect')?.click(), 1`);
    const found = await until(() => js(`(() => { const rows = [...document.querySelectorAll('.di-dialog .di-row[data-found]')]; return rows.length ? Object.fromEntries(rows.map((r) => [r.dataset.kind, Number(r.dataset.found)])) : null; })()`), 'the results', 8000).catch(() => null);
    await fresh('deck-redact-inspect.png');
    for (const k of ['comments', 'notes', 'invisible', 'offSlide', 'personal']) {
      await js(`document.querySelector('.di-remove-${k}')?.click(), 1`);
      await until(() => js(`document.querySelector('.di-row[data-kind="${k}"]')?.dataset.found === '0'`), `${k} removed`, 8000).catch(() => {});
    }
    const after = await js(`Object.fromEntries([...document.querySelectorAll('.di-dialog .di-row[data-found]')].map((r) => [r.dataset.kind, Number(r.dataset.found)]))`);
    await js(`document.querySelector('.di-close')?.click(), 1`);
    const m1 = doc.model({ id: session.id, slide: 1 });
    check(
      'deck-redact: Inspect Presentation finds the comment, the notes, the hidden and the off-slide shapes and the authors, and Remove All takes each out of the open deck',
      Boolean(found) && found.comments === 1 && found.notes === 1 && found.invisible === 1 && found.offSlide === 1 && found.personal >= 2
        && ['comments', 'notes', 'invisible', 'offSlide', 'personal'].every((k) => after[k] === 0)
        && (m1.comments || []).length === 0 && !(m1.slide?.shapes || []).some((s) => s.name === 'Hidden note' || s.name === 'Parked'),
      JSON.stringify({ found, after }),
    );
  } catch (err) {
    check('deck-redact: the Redact and Inspect checks ran', false, err.stack || err.message);
  }
}
