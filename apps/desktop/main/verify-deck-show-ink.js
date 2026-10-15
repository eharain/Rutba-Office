// Presentations: ink as PowerPoint keeps it, and the show's pen. A stroke
// drawn with Draw → Pen in edit view is saved as PowerPoint saves ink, a
// p14 content part of InkML with the freeform its fallback, and drawn
// when the deck is opened again. In the show, Ctrl+P takes the pen and a
// press draws instead of moving on; the bar's Highlighter draws too; on
// leaving the show the editor asks "Do you want to keep your ink
// annotations?": Keep puts the strokes on their slide as ink, and a second
// show's ink, Discarded, leaves nothing. Run alone with RUTBA_VERIFY_ONLY=deck-ink.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait, capture, doc
 */
export async function verifyDeckShowInk(h, { dir }) {
  const { open, check, wait, capture, doc } = h;
  // Waits as the harness's until does, and answers what the condition found once it held.
  const until = async (fn, what, ms) => { let found; await h.until(async () => { found = await fn(); return found; }, what, ms); return found; };
  const own = path.join(dir, 'deck-ink');
  const file = path.join(own, 'talk.pptx');
  try {
    fs.mkdirSync(own, { recursive: true });
    fs.writeFileSync(file, buildPptx({ title: 'Talk', slides: [{ layout: 'title', title: 'Quarter in review', body: 'Draft' }, { layout: 'obj', title: 'Numbers', body: ['Up 12%'] }] }));
  } catch (err) {
    return check('deck-ink: the fixture was made', false, err.message);
  }
  try {
    const win = await open('slides', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const session = doc.sessions().filter((s) => s.kind === 'deck' && s.path === file).pop();
    const inksOn = (i) => (doc.model({ id: session.id, slide: i }).slide?.shapes || []).filter((s) => /^Ink \d+$/.test(s.name || ''));
    const ribbon = (tip) => js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || n.title || '').startsWith(${JSON.stringify(tip)}) || n.textContent.trim() === ${JSON.stringify(tip)}); if (!b) return 'no button'; if (b.disabled) return 'disabled'; b.click(); return 'clicked'; })()`);
    const drag = async (from, to, steps = 12) => {
      wc.sendInputEvent({ type: 'mouseDown', x: from.x, y: from.y, button: 'left', clickCount: 1 });
      for (let k = 1; k <= steps; k++) { wc.sendInputEvent({ type: 'mouseMove', x: Math.round(from.x + ((to.x - from.x) * k) / steps), y: Math.round(from.y + ((to.y - from.y) * k) / steps), button: 'left', modifiers: ['leftButtonDown'] }); await wait(12); }
      wc.sendInputEvent({ type: 'mouseUp', x: to.x, y: to.y, button: 'left', clickCount: 1 });
    };

    await until(() => js(`Boolean(document.querySelector('.sl-slide'))`), 'the deck', 10000);

    // 1. Draw → Pen in edit view, saved as PowerPoint saves ink.
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Draw')?.click(), 1`);
    await wait(250);
    await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Pen'); b?.click(); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.sl-ink-surface'))`), 'the pen surface', 3000).catch(() => {});
    const s = await js(`(() => { const r = document.querySelector('.sl-ink-surface').getBoundingClientRect(); return { left: r.left, top: r.top, width: r.width, height: r.height }; })()`);
    await drag({ x: Math.round(s.left + s.width * 0.2), y: Math.round(s.top + s.height * 0.75) }, { x: Math.round(s.left + s.width * 0.6), y: Math.round(s.top + s.height * 0.8) });
    const drawn = await until(() => inksOn(0).length === 1, 'the stroke on the slide', 6000).then(() => true).catch(() => false);
    await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Select'); b?.click(); return 1; })()`);
    await js(`(() => { const b = [...document.querySelectorAll('.rw-btn')].find((n) => (n.dataset.tip || n.title || '').startsWith('Save')); b?.click(); return 1; })()`);
    await until(() => doc.meta({ id: session.id }).dirty === false, 'the save', 15000).catch(() => {});
    const saved = Deck.open(fs.readFileSync(file));
    const slideXml = saved.pkg.text(saved.slideParts[0].part);
    const inkPart = saved.pkg.partNames().find((n) => /^ppt\/ink\/ink\d+\.xml$/.test(n));
    const asPowerPoint = /<mc:Choice xmlns:p14="[^"]+" Requires="p14"><p:contentPart p14:bwMode="auto" r:id="rId\d+"><p14:nvContentPartPr><p14:cNvPr id="\d+" name="Ink 1"\/>/.test(slideXml)
      && /<mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="\d+" name="Ink 1"\/>/.test(slideXml)
      && Boolean(inkPart) && /<inkml:trace\b/.test(saved.pkg.text(inkPart)) && /application\/inkml\+xml/.test(saved.pkg.text('[Content_Types].xml'));
    const reopened = saved.slide(0).shapes.filter((x) => /^Ink \d+$/.test(x.name || '')).length === 1;
    check('deck-ink: a stroke drawn with Draw → Pen is saved as PowerPoint saves ink, a p14 content part of InkML, the freeform its fallback, and is there when the deck is opened again', drawn && asPowerPoint && reopened, JSON.stringify({ drawn, asPowerPoint, inkPart, reopened }));

    // 2. The show's pen: Ctrl+P, a stroke; the Highlighter from the bar, a stroke; Escape twice; Keep.
    const started = await ribbon('Start from the beginning');
    await until(() => js(`Boolean(document.querySelector('.sl-present .sl-show-stage'))`), 'the show', 6000).catch(() => {});
    await wait(500);
    // Pressed as a person presses them, so every listener hears them in its own order.
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'P', modifiers: ['control'] }); wc.sendInputEvent({ type: 'keyUp', keyCode: 'P', modifiers: ['control'] }); await wait(150);
    const penOn = await until(() => js(`Boolean(document.querySelector('.sl-show-ink.on')) && Boolean(document.querySelector('.sl-show-pen.on[data-tool="pen"]'))`), 'the pen in the show', 3000).catch(() => false);
    const st = await js(`(() => { const r = document.querySelector('.sl-present .sl-show-stage').getBoundingClientRect(); return { left: r.left, top: r.top, width: r.width, height: r.height }; })()`);
    const bar1 = await js(`document.querySelector('.sl-present-bar')?.textContent || ''`);
    await drag({ x: Math.round(st.left + st.width * 0.3), y: Math.round(st.top + st.height * 0.6) }, { x: Math.round(st.left + st.width * 0.7), y: Math.round(st.top + st.height * 0.62) });
    await js(`document.querySelector('.sl-show-pen[data-tool="highlighter"]')?.click(), 1`);
    await drag({ x: Math.round(st.left + st.width * 0.2), y: Math.round(st.top + st.height * 0.4) }, { x: Math.round(st.left + st.width * 0.6), y: Math.round(st.top + st.height * 0.4) });
    const strokes = await until(() => js(`document.querySelectorAll('.sl-show-ink .sl-show-stroke').length === 2`), 'two strokes in the show', 4000).catch(() => false);
    const bar2 = await js(`document.querySelector('.sl-present-bar')?.textContent || ''`);
    wc.invalidate();
    await wait(600);
    await capture(win, 'deck-ink-show.png');
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'Escape' }); wc.sendInputEvent({ type: 'keyUp', keyCode: 'Escape' }); await wait(150);
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'Escape' }); wc.sendInputEvent({ type: 'keyUp', keyCode: 'Escape' }); await wait(150);
    const asked = await until(() => js(`Boolean(document.querySelector('.sl-keep-ink .sl-keep-ink-keep'))`), 'the keep question', 6000).catch(() => false);
    const question = await js(`document.querySelector('.sl-keep-ink')?.textContent || ''`);
    wc.invalidate();
    await wait(600);
    await capture(win, 'deck-ink-keep.png');
    await js(`document.querySelector('.sl-keep-ink .sl-keep-ink-keep')?.click(), 1`);
    const kept = await until(() => inksOn(0).length === 3, 'the strokes kept on the slide', 8000).then(() => true).catch(() => false);
    check(
      'deck-ink: in the show Ctrl+P takes the pen and a press draws instead of moving on, the bar\'s Highlighter draws too, and on leaving the show "Do you want to keep your ink annotations?", Keep puts the strokes on the slide as ink',
      started === 'clicked' && penOn === true && strokes === true && bar1 === bar2 && asked === true && /keep your ink annotations/.test(question) && kept,
      JSON.stringify({ started, penOn, strokes, bar1, bar2, asked, kept, inks: inksOn(0).length }),
    );

    // 3. Another show, its ink discarded.
    await ribbon('Start from the beginning');
    await until(() => js(`Boolean(document.querySelector('.sl-present .sl-show-stage'))`), 'the show again', 6000).catch(() => {});
    await wait(400);
    // Pressed as a person presses them, so every listener hears them in its own order.
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'P', modifiers: ['control'] }); wc.sendInputEvent({ type: 'keyUp', keyCode: 'P', modifiers: ['control'] }); await wait(150);
    await drag({ x: Math.round(st.left + st.width * 0.3), y: Math.round(st.top + st.height * 0.3) }, { x: Math.round(st.left + st.width * 0.5), y: Math.round(st.top + st.height * 0.35) });
    await until(() => js(`document.querySelectorAll('.sl-show-ink .sl-show-stroke').length === 1`), 'a stroke', 4000).catch(() => {});
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'Escape' }); wc.sendInputEvent({ type: 'keyUp', keyCode: 'Escape' }); await wait(150);
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'Escape' }); wc.sendInputEvent({ type: 'keyUp', keyCode: 'Escape' }); await wait(150);
    await until(() => js(`Boolean(document.querySelector('.sl-keep-ink .sl-keep-ink-discard'))`), 'the keep question again', 6000).catch(() => {});
    await js(`document.querySelector('.sl-keep-ink .sl-keep-ink-discard')?.click(), 1`);
    await wait(600);
    check('deck-ink: Discard lets a show\'s ink go, the slide left as it was', inksOn(0).length === 3 && !(await js(`Boolean(document.querySelector('.sl-keep-ink'))`)), String(inksOn(0).length));
  } catch (err) {
    check('deck-ink: the ink checks ran', false, err.stack || err.message);
  }
}
