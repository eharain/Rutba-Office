// Presentations: Animations → Motion Paths and Effect Options → By
// Paragraph, pressed on the ribbon — a rectangle given the Down path, drawn
// on the stage from its centre; the body's Fade split into one effect per
// paragraph; and in the show the first click moves the rectangle down the
// path and leaves it there, the next brings in the first paragraph alone.
// Run alone with RUTBA_VERIFY_ONLY=motion.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn, doc
 */
export async function verifyDeckMotion(h, { dir }) {
  const { open, check, until, wait, errorsIn, doc } = h;
  const file = path.join(dir, 'motion.pptx');
  try {
    const deck = Deck.open(buildPptx({ title: 'Motion', slides: [{ layout: 'obj', title: 'Three steps', body: ['Plan', 'Build', 'Ship'] }, { layout: 'obj', title: 'Next' }] }));
    deck.addShape(0, { preset: 'rect', x: 80, y: 120, w: 160, h: 90, name: 'Box' });
    fs.writeFileSync(file, deck.save());
    const win = await open('slides', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = doc.sessions().filter((s) => s.kind === 'deck' && s.path === file).pop();
    const model = () => doc.model({ id: session.id, slide: 0 }).slide;
    const shapes = model().shapes;
    const box = String(shapes.find((s) => s.name === 'Box').id);
    const body = String(shapes.find((s) => (s.text?.paragraphs || []).some((p) => p.plain === 'Build')).id);
    const pickShape = (id) => js(`(() => { const el = document.querySelector('.sl-hit[data-shape="${id}"]'); el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true })); el.click(); return 1; })()`);
    const ribbonButton = (label) => js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); if (!b) return false; b.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: 400, clientY: 110 })); return true; })()`);
    const menuItem = async (label) => {
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}))`), `the menu item ${label}`, 3000).catch(() => {});
      return js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); if (!b) return false; b.click(); return true; })()`);
    };

    // Motion Paths → Down on the rectangle.
    await until(() => js(`document.querySelectorAll('.sl-hit[data-shape]').length >= 3`), 'the shapes', 8000);
    await pickShape(box);
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Animations')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.sl-an-paths'))`), 'the Motion Paths button', 4000).catch(() => {});
    await wait(200);
    const opened = await ribbonButton('Motion Paths');
    const picked = await menuItem('Down');
    await until(() => (model().animations || []).some((e) => e.kind === 'path'), 'the path', 5000).catch(() => {});
    const pathFx = (model().animations || []).find((e) => e.kind === 'path');
    await until(() => js(`Boolean(document.querySelector('.sl-motion-paths path'))`), 'the path on the stage', 3000).catch(() => {});
    const drawn = await js(`document.querySelector('.sl-motion-paths path')?.getAttribute('d') || null`);
    check('presentations: Animations → Motion Paths → Down gives the shape a path, drawn on the stage from its centre',
      opened && picked && pathFx?.path === 'M 0 0 L 0 0.25 E' && String(pathFx.shapeId) === box && /^M160\.0 165\.0 /.test(drawn || ''),
      JSON.stringify({ opened, picked, path: pathFx?.path, drawn: drawn?.slice(0, 40) }));

    // Fade on the body, then Effect Options → By Paragraph.
    await pickShape(body);
    await wait(200);
    await js(`(() => { const b = document.querySelector('.sl-an-pick[data-effect="fade"]'); b?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); b?.click(); return 1; })()`);
    await until(() => (model().animations || []).length === 2, 'the Fade', 5000).catch(() => {});
    await wait(300);
    const options = await ribbonButton('Effect Options');
    const byPara = await menuItem('By Paragraph');
    await until(() => (model().animations || []).filter((e) => e.paragraph != null).length === 3, 'the paragraphs', 5000).catch(() => {});
    const list = (model().animations || []).map((e) => [e.kind, e.paragraph, e.trigger]);
    await ribbonButton('Animation Pane');
    await until(() => js(`document.querySelectorAll('.sl-animrow').length === 4`), 'the Animation Pane', 3000).catch(() => {});
    const rows = await js(`[...document.querySelectorAll('.sl-animrow .sl-layer-title')].map((n) => n.textContent.trim())`).catch(() => []);
    check('presentations: Effect Options → By Paragraph plays the body\'s Fade one paragraph at a time, each on a click, listed by its words',
      options && byPara && JSON.stringify(list) === JSON.stringify([['path', null, 'onClick'], ['entr', 0, 'onClick'], ['entr', 1, 'onClick'], ['entr', 2, 'onClick']]) && ['Plan', 'Build', 'Ship'].every((w, i) => String(rows[i + 1] || '').endsWith(': ' + w)),
      JSON.stringify({ options, byPara, list, rows }));

    // The show: the box moves down the path and stays; then the first paragraph comes in alone.
    await js(`(async () => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Slide Show')?.click(); await new Promise((r) => setTimeout(r, 200)); [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'From Beginning')?.click(); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.sl-present [data-shape="${body}"] [data-para="2"]'))`), 'the show', 6000).catch(() => {});
    const paraVis = (n) => js(`(() => { const g = document.querySelector('.sl-present [data-shape="${body}"] [data-para="${n}"]'); return g ? getComputedStyle(g).visibility : null; })()`);
    const startVis = [await paraVis(0), await paraVis(1), await paraVis(2)];
    await js(`document.querySelector('.sl-present')?.click(), 1`);
    const H = model().size?.height || doc.model({ id: session.id }).size?.height || 720;
    const moved = await until(() => js(`(() => { const t = document.querySelector('.sl-present [data-shape="${box}"]')?.style.transform || ''; const m = /translate\\(([-\\d.]+)px,\\s*([-\\d.]+)px\\)/.exec(t); return m ? Math.abs(Number(m[2]) - ${H} * 0.25) < 1 && Math.abs(Number(m[1])) < 1 : false; })()`), 'the box moved', 6000).then(() => true).catch(() => false);
    const where = await js(`document.querySelector('.sl-present [data-shape="${box}"]')?.style.transform || null`);
    await js(`document.querySelector('.sl-present')?.click(), 1`);
    await until(async () => (await paraVis(0)) === 'visible', 'the first paragraph', 4000).catch(() => {});
    await wait(700);
    const afterVis = [await paraVis(0), await paraVis(1), await paraVis(2)];
    win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'Escape' });
    win.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'Escape' });
    check('presentations: in the show the box moves down its path and stays there, and the next click brings in the first paragraph alone',
      moved && JSON.stringify(startVis) === '["hidden","hidden","hidden"]' && JSON.stringify(afterVis) === '["visible","hidden","hidden"]',
      JSON.stringify({ moved, where, startVis, afterVis }));

    // Saved: the path and the paragraph build in the file as PowerPoint writes them.
    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    await until(() => { try { return /animMotion/.test(Deck.open(fs.readFileSync(file)).pkg.text('ppt/slides/slide1.xml')); } catch { return false; } }, 'the save', 6000).catch(() => {});
    const saved = Deck.open(fs.readFileSync(file)).pkg.text('ppt/slides/slide1.xml');
    check('presentations: the saved deck keeps the motion path and the paragraph build as PowerPoint writes them',
      /<p:animMotion origin="layout" path="M 0 0 L 0 0\.25 E"/.test(saved) && /<p:bldP build="p" spid="\d+"/.test(saved) && /<p:pRg st="2" end="2"\/>/.test(saved),
      saved.includes('animMotion') ? 'animMotion in the file' : 'not saved');

    const complaints = await errorsIn(win);
    check('presentations: motion paths and effects by paragraph report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('presentations: the motion path checks ran', false, err.message);
  }
}
