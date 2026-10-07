// Presentations: Edit Points.
//
// A rectangle's right-click offers Edit Points: its outline and four points
// are drawn over it, its handles put away. Dragging a corner moves it — the
// shape becomes a path of its own and its box grows to hold it — a
// right-click deletes a point or adds one on the outline, Escape leaves, and
// the file keeps the path as <a:custGeom>. Run alone with RUTBA_VERIFY_ONLY=points.
import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where the fixture is written
 */
export async function verifyDeckPoints(h, { dir }) {
  const { open, check, until, wait, press, errorsIn, doc, sessionFor } = h;
  const file = path.join(dir, 'points.pptx');
  const deck = Deck.open(buildPptx({ title: 'Points', slides: [{ layout: 'blank' }] }));
  const id = deck.addShape(0, { preset: 'rect', x: 300, y: 200, w: 320, h: 180, name: 'Box' });
  fs.writeFileSync(file, deck.save());
  try {
    const win = await open('slides', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const model = () => doc.model({ id: sessionFor('deck').id, slide: 0 });
    const box = () => model().slide.shapes.find((s) => s.name === 'Box');
    const pointCount = () => js(`document.querySelectorAll('.sl-points .sl-point').length`);
    const centre = (selector) => js(`(() => { const r = document.querySelector(${JSON.stringify(selector)})?.getBoundingClientRect(); return r ? { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) } : null; })()`);
    const rightClick = (selector) => js(`(() => {
      const el = document.querySelector(${JSON.stringify(selector)});
      if (!el) return 'no ' + ${JSON.stringify(selector)};
      const r = el.getBoundingClientRect();
      el.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: r.left + r.width / 2, clientY: r.top + r.height / 2 }));
      return 'opened';
    })()`);
    const pick = async (label) => {
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)} && !b.disabled))`), `"${label}"`, 3000);
      return js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}); b.click(); return 'picked'; })()`);
    };
    const hit = `.sl-hit[data-shape="${id}"]`;
    await until(() => js(`Boolean(document.querySelector(${JSON.stringify(hit)}))`), 'the rectangle on the stage', 8000);

    // Right-click → Edit Points: the outline and four corners, the handles put away.
    await js(`(() => { const el = document.querySelector(${JSON.stringify(hit)}); el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); el.click(); return 'selected'; })()`);
    await rightClick(hit);
    const entered = await pick('Edit Points');
    await until(async () => (await pointCount()) === 4, 'four points', 4000).catch(() => false);
    const handles = await js(`document.querySelectorAll('.sl-handle').length`);
    check('slides: Edit Points draws the shape\'s outline and its four corners, its resize handles put away',
      entered === 'picked' && (await pointCount()) === 4 && handles === 0, `${entered}; ${await pointCount()} points, ${handles} handles`);

    // Drag the top-right corner out and up.
    const before = { ...box().geometry };
    const corner = await centre('.sl-points .sl-point[data-point="1"]');
    wc.sendInputEvent({ type: 'mouseDown', x: corner.x, y: corner.y, button: 'left', clickCount: 1 });
    for (let i = 1; i <= 6; i++) {
      wc.sendInputEvent({ type: 'mouseMove', x: corner.x + i * 10, y: corner.y - i * 5, button: 'left' });
      await wait(30);
    }
    wc.sendInputEvent({ type: 'mouseUp', x: corner.x + 60, y: corner.y - 30, button: 'left', clickCount: 1 });
    const reshaped = await until(() => box()?.preset === 'custom', 'the shape a path of its own', 5000).catch(() => false);
    const after = box().geometry;
    check('slides: dragging a corner moves it — the shape becomes a path of its own and its box grows to hold it',
      reshaped === true && after.w > before.w + 20 && after.h > before.h + 10 && after.y < before.y - 5 && Math.abs(after.x - before.x) < 1,
      `${JSON.stringify(before)} → ${JSON.stringify(after)}`);
    await wait(300);
    wc.invalidate();
    await wait(400);
    if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'slides-points.png'), (await wc.capturePage()).toPNG());

    // Right-click a point → Delete Point; the outline → Add Point.
    await rightClick('.sl-points .sl-point[data-point="2"]');
    await pick('Delete Point');
    const three = await until(async () => (await pointCount()) === 3 && (box()?.path?.d.match(/[ML]/g) || []).length === 3, 'three points', 4000).catch(() => false);
    await rightClick('.sl-points .sl-points-path');
    await pick('Add Point');
    const four = await until(async () => (await pointCount()) === 4 && (box()?.path?.d.match(/[ML]/g) || []).length === 4, 'four points again', 4000).catch(() => false);
    check('slides: a right-click deletes a point, or adds one on the outline', three === true && four === true, `three ${three}, four ${four}`);

    // Escape leaves Edit Points; the handles come back.
    await press(wc, 'Escape');
    const left = await until(async () => (await pointCount()) === 0 && (await js(`document.querySelectorAll('.sl-handle').length`)) > 0, 'Edit Points left', 3000).catch(() => false);
    check('slides: Escape leaves Edit Points and the handles come back', left === true, String(left));

    // Saved: the path in the file.
    await js(`document.querySelector('.sl-slide')?.focus(), 1`);
    await press(wc, 's', { modifiers: ['control'] });
    await wait(1500);
    const saved = Deck.open(fs.readFileSync(file));
    const xml = saved.pkg.text('ppt/slides/slide1.xml');
    const kept = saved.slide(0).shapes.find((s) => s.name === 'Box');
    check('slides: the file keeps the edited outline as the shape\'s own <a:custGeom>',
      kept?.preset === 'custom' && /<a:custGeom>[\s\S]*<a:moveTo>/.test(xml) && !/prst="rect"/.test(xml) && (kept.path.d.match(/[ML]/g) || []).length === 4,
      `${kept?.preset}; ${kept?.path?.d}`);

    const complaints = await errorsIn(win);
    check('slides: Edit Points reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('slides: the Edit Points checks ran', false, err.message);
  }
}
