// Presentations: Insert → Zoom, pressed on the ribbon — a Slide Zoom to
// slide 3 put on slide 1 as its picture, linked to it; in the show a click
// on it goes to slide 3 and the next step comes back to slide 1; a Summary
// Zoom of slides 2 and 4 is a new Summary slide before slide 2 with a
// picture of each; and the saved deck keeps the zooms as pictures linked to
// their slides, which any PowerPoint follows.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait, doc, sessionFor
 */
export async function verifyDeckZoom(h, { dir }) {
  const { open, check, until, wait, doc, sessionFor } = h;
  const file = path.join(dir, 'zoom.pptx');
  try {
    fs.writeFileSync(file, buildPptx({ title: 'Zoom', slides: ['Start', 'Second', 'Third', 'Fourth'].map((t, i) => ({ layout: i ? 'obj' : 'title', title: t })) }));
    const win = await open('slides', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const model = (slide) => doc.model({ id: sessionFor('deck').id, slide });
    const zoomMenu = (item) => js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Zoom');
      if (!b || b.disabled) return 'no button';
      b.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: 700, clientY: 110 }));
      await new Promise((r) => setTimeout(r, 200));
      const m = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(item)});
      if (!m) return 'no item'; m.click(); return 'picked';
    })()`);
    const choose = async (slides) => {
      await until(() => js(`Boolean(document.querySelector('.sl-zoom-ok'))`), 'the Zoom box', 4000).catch(() => {});
      for (const s of slides) await js(`(() => { document.querySelector('.sl-zoom-item[data-slide="${s}"] input')?.click(); return 1; })()`);
      await js(`document.querySelector('.sl-zoom-ok')?.click(), 1`);
    };
    const zooms = (slide) => (model(slide).slide?.shapes || []).filter((s) => /Zoom \d+$/.test(s.name || ''));
    const bar = () => js(`document.querySelector('.sl-present-bar')?.textContent.trim().split(' ')[0] || null`);

    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 4`), 'the slides', 8000);
    const opened = await zoomMenu('Slide Zoom…');
    await choose([2]);
    const placed = await until(() => zooms(0).length === 1 && zooms(0)[0].action?.kind === 'slide' && zooms(0)[0].action.slide === 2, 'the slide zoom', 10000).then(() => true).catch(() => false);
    check('presentations: Insert → Zoom → Slide Zoom puts a picture of slide 3 on this slide, linked to it',
      opened === 'picked' && placed, JSON.stringify({ opened, zooms: zooms(0).map((s) => ({ name: s.name, action: s.action, kind: s.kind })) }));

    // The show: the zoom goes there, the next step comes back.
    await js(`(() => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Slide Show')?.click(); return 1; })()`);
    await wait(200);
    await js(`(() => { [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'From Beginning')?.click(); return 1; })()`);
    await until(async () => (await bar()) === '1', 'the show on slide 1', 6000).catch(() => {});
    await wait(500);
    const at = await js(`(() => { const id = ${JSON.stringify(String(zooms(0)[0]?.id))}; const g = document.querySelector('.sl-show-stage [data-shape="' + id + '"]'); if (!g) return null; const r = g.getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) }; })()`);
    if (at) {
      wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      wc.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1 });
    }
    const went = await until(async () => (await bar()) === '3', 'slide 3', 5000).then(() => true).catch(() => false);
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'Right' });
    wc.sendInputEvent({ type: 'keyUp', keyCode: 'Right' });
    const back = await until(async () => (await bar()) === '1', 'back to slide 1', 5000).then(() => true).catch(() => false);
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'Escape' });
    wc.sendInputEvent({ type: 'keyUp', keyCode: 'Escape' });
    await until(async () => (await bar()) === null, 'the show to end', 4000).catch(() => {});
    check('presentations: in the show a click on the zoom goes to its slide, and the next step comes back to the zoom', Boolean(at) && went && back, JSON.stringify({ at, went, back }));

    // Summary Zoom of slides 2 and 4.
    await zoomMenu('Summary Zoom…');
    await choose([1, 3]);
    const summary = await until(() => model(0).count === 5 && zooms(1).length === 2, 'the summary slide', 12000).then(() => true).catch(() => false);
    const title = (model(1).slide?.shapes || []).find((s) => s.placeholder?.type === 'title')?.text?.paragraphs?.[0]?.plain;
    const targets = zooms(1).map((s) => s.action?.slide);
    check('presentations: Insert → Zoom → Summary Zoom makes a Summary slide before the first slide picked, a picture of each linked to it',
      summary && title === 'Summary' && JSON.stringify(targets) === '[2,4]', JSON.stringify({ summary, count: model(0).count, title, targets }));

    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => {
      try {
        const deck = Deck.open(fs.readFileSync(file));
        return (deck.pkg.text('ppt/slides/slide1.xml').match(/action="ppaction:\/\/hlinksldjump"/g) || []).length === 1 && deck.slideCount === 5;
      } catch { return false; }
    }, 'the saved zooms', 8000).then(() => true).catch(() => false);
    check('presentations: the saved deck keeps the zooms as pictures linked to their slides', saved, saved ? 'hlinksldjump on the zoom picture' : 'not in the file');
  } catch (err) {
    check('presentations: the Zoom checks ran', false, err.message);
  }
}
