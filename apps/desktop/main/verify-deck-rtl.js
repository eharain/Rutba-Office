// Presentations: right-to-left paragraphs, pressed on the ribbon.
//
// A text box of Arabic, selected; Home → Paragraph → Right-to-left text
// direction turns its paragraph — written rtl, aligned right as PowerPoint
// does — and the words are drawn from the box's right edge; Left-to-right
// turns it back to the left. Run alone with RUTBA_VERIFY_ONLY=deckrtl.
import fs from 'node:fs';
import path from 'node:path';
import { buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where the deck is written
 */
export async function verifyDeckRtl(h, { dir }) {
  const { open, check, until, wait, errorsIn, doc, sessionFor } = h;
  const file = path.join(dir, 'rtl-deck.pptx');
  const WORDS = 'مرحبا بالعالم';
  fs.writeFileSync(file, buildPptx({ title: 'RTL', slides: [{ layout: 'blank', textBoxes: [{ x: 120, y: 160, w: 560, h: 80, paragraphs: [{ runs: [{ text: WORDS, size: 28 }] }] }] }] }));
  try {
    const win = await open('slides', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const session = sessionFor('deck');
    const model = () => doc.model({ id: session.id });
    const shape = () => (model().slide?.shapes || []).find((s) => s.text);
    await until(() => js(`Boolean(document.querySelector('.sl-stage .sl-hit'))`), 'the slide drawn', 8000);
    const id = shape()?.id;
    // Select the box as a person does: a press on it.
    const at = await js(`(() => { const el = document.querySelector('.sl-hit[data-shape="${id}"]'); if (!el) return null; const r = el.getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) }; })()`);
    if (at) {
      wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      wc.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1 });
    }
    await until(() => js(`document.querySelector('.sl-hit.selected')?.dataset.shape === ${JSON.stringify(String(id))}`), 'the box selected', 4000).catch(() => {});
    const pressTitled = (start) => js(`(() => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Home')?.click();
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.getAttribute('data-tip') || n.getAttribute('aria-label') || '').startsWith(${JSON.stringify(start)}));
      if (!b || b.disabled) return b ? 'disabled' : 'missing';
      b.click();
      return 'clicked';
    })()`);
    // Where the words are drawn, against the box.
    const placed = () => js(`(() => {
      const hit = document.querySelector('.sl-hit[data-shape="${id}"]')?.getBoundingClientRect();
      const text = [...document.querySelectorAll('.sl-stage svg text')].find((t) => t.textContent.includes(${JSON.stringify(WORDS)}));
      if (!hit || !text) return null;
      const r = text.getBoundingClientRect();
      return { dir: text.getAttribute('direction') || 'ltr', fromLeft: Math.round(r.left - hit.left), fromRight: Math.round(hit.right - r.right) };
    })()`);
    const before = await placed();
    const pressed = await pressTitled('Right-to-left text direction');
    const turned = await until(() => shape()?.text?.paragraphs?.[0]?.rtl === true, 'the paragraph right to left', 4000).then(() => true, () => false);
    await wait(300);
    const during = await placed();
    const align = shape()?.text?.paragraphs?.[0]?.align;
    check('presentations: Home → Paragraph → Right-to-left text direction turns the paragraph and aligns it right, as PowerPoint does',
      pressed === 'clicked' && turned && align === 'right', `button ${pressed}; rtl ${shape()?.text?.paragraphs?.[0]?.rtl}; align ${align}`);
    check('presentations: a right-to-left paragraph is drawn from the box\'s right edge',
      Boolean(before && during) && during.dir === 'rtl' && during.fromRight < 20 && during.fromLeft > during.fromRight && before.fromLeft < 20,
      JSON.stringify({ before, during }));
    if (process.env.RUTBA_VERIFY_CAPTURE) {
      wc.invalidate();
      await wait(300);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'deck-rtl.png'), (await wc.capturePage()).toPNG());
    }
    const back = await pressTitled('Left-to-right text direction');
    const undone = await until(() => shape()?.text?.paragraphs?.[0]?.rtl === false, 'the paragraph left to right', 4000).then(() => true, () => false);
    await wait(300);
    const after = await placed();
    check('presentations: Left-to-right text direction turns it back, drawn from the left again',
      back === 'clicked' && undone && shape()?.text?.paragraphs?.[0]?.align === 'left' && after?.dir === 'ltr' && after.fromLeft < 20,
      JSON.stringify({ back, undone, after, align: shape()?.text?.paragraphs?.[0]?.align }));
    const complaints = await errorsIn(win);
    check('presentations: right-to-left paragraphs report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('presentations: the right-to-left paragraph checks ran', false, err.message);
  }
}
