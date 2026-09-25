// Presentation: Home → Arrange — multi-select, Align, Group/Ungroup, Rotate
// and Flip.
//
// Two rectangles are added through the ribbon (never through a raw engine
// call the window would not know about — the window's own model only
// changes when it makes the apply itself), one dragged clear of the other so
// each has a hit area of its own. Shift+click builds the selection, Align
// Left from the Arrange menu moves them, Group folds them into one p:grpSp
// the window then treats as a single shape, and Ungroup, Rotate and Flip are
// each read back from the engine's own model and, once saved, from the file.
// Everything this check adds comes off again before it returns, so the
// fixture is exactly as later checks expect it. Run alone with
// RUTBA_VERIFY_ONLY=arrange.

import fs from 'node:fs';
import path from 'node:path';
import { Deck } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, capture, doc, sessionFor
 * @param {{ file: string }} args the deck fixture (files.pptx)
 */
export async function verifyDeckArrange(h, { file }) {
  const { open, check, until, wait, press, errorsIn, doc, sessionFor } = h;
  if (!file) return check('slides: the arrange checks have a fixture', false, 'no files.pptx');
  try {
    const win = await open('slides', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const model = () => doc.model({ id: sessionFor('deck').id, slide: 0 });
    const shapeOf = (id) => model().slide.shapes.find((s) => String(s.id) === String(id));

    const clickRibbon = (title) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}) && !n.disabled);
      if (!b) return 'no button ' + ${JSON.stringify(title)};
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      b.click();
      return 'clicked';
    })()`);
    const clickMenuItem = async (label) => {
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)} && !b.disabled))`), `the "${label}" menu item`, 3000);
      return js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}); if (!b) return 'missing'; b.click(); return 'clicked'; })()`);
    };
    const shiftClick = (selector) => js(`(() => {
      const el = document.querySelector(${JSON.stringify(selector)});
      if (!el) return false;
      el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, shiftKey: true }));
      el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, shiftKey: true }));
      return true;
    })()`);
    const plainClick = (selector) => js(`(() => {
      const el = document.querySelector(${JSON.stringify(selector)});
      if (!el) return false;
      el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
      return true;
    })()`);
    const rectOf = (selector) => js(`(() => { const r = document.querySelector(${JSON.stringify(selector)})?.getBoundingClientRect(); return r ? { x: r.left + r.width / 2, y: r.top + r.height / 2 } : null; })()`);
    const drag = async (selector, dx, dy) => {
      const at = await rectOf(selector);
      if (!at) throw new Error(`nothing at ${selector}`);
      const x = Math.round(at.x);
      const y = Math.round(at.y);
      wc.sendInputEvent({ type: 'mouseMove', x, y });
      wc.sendInputEvent({ type: 'mouseDown', x, y, button: 'left', clickCount: 1 });
      await wait(60);
      for (const f of [0.25, 0.5, 0.75, 1]) {
        wc.sendInputEvent({ type: 'mouseMove', x: Math.round(x + dx * f), y: Math.round(y + dy * f), button: 'left' });
        await wait(40);
      }
      wc.sendInputEvent({ type: 'mouseUp', x: x + dx, y: y + dy, button: 'left', clickCount: 1 });
      await wait(400);
    };
    const selectedCount = () => js(`document.querySelectorAll('.sl-hit.selected').length`);

    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 1`), 'the slide sorter', 8000);
    const before = model().slide.shapes.length;

    // Two rectangles, through the ribbon — Home → Drawing → Shapes — so the
    // window's own model has them from the start, drawn with hit areas.
    await clickRibbon('Shapes');
    await clickMenuItem('Rectangle');
    await until(() => model().slide.shapes.length === before + 1, 'the first rectangle', 5000);
    const a = model().slide.shapes[model().slide.shapes.length - 1].id;
    await until(() => js(`Boolean(document.querySelector('.sl-hit[data-shape="${a}"]'))`), 'its hit area', 4000);

    await clickRibbon('Shapes');
    await clickMenuItem('Rectangle');
    await until(() => model().slide.shapes.length === before + 2, 'the second rectangle', 5000);
    const b = model().slide.shapes[model().slide.shapes.length - 1].id;
    await until(() => js(`Boolean(document.querySelector('.sl-hit[data-shape="${b}"]'))`), 'its hit area', 4000);

    // Both land centred, one on the other; b (added last, on top) is dragged
    // clear so each has a hit area of its own to click.
    await drag(`.sl-hit[data-shape="${b}"]`, 260, 0);
    const dragged = await until(() => shapeOf(b).geometry.x > shapeOf(a).geometry.x + 50, 'b to clear a', 4000).catch(() => false);
    if (!dragged) {
      // In the full run an earlier check can leave something over the
      // stage's centre, and the drag lands on that instead. Moving b is not
      // what this block tests, so move it through the engine and let the
      // window fetch the slide again (away to slide 2 and back).
      const g = shapeOf(b).geometry;
      await doc.apply({ id: sessionFor('deck').id, ops: [{ op: 'setGeometry', slide: 0, shape: b, x: g.x + 260, y: g.y, w: g.w, h: g.h }] });
      await js(`document.querySelectorAll('.sl-thumb')[1]?.click(), 1`);
      await wait(400);
      await js(`document.querySelectorAll('.sl-thumb')[0]?.click(), 1`);
      await until(() => shapeOf(b).geometry.x > shapeOf(a).geometry.x + 50, 'b to clear a', 4000);
      await until(() => js(`Boolean(document.querySelector('.sl-hit[data-shape="${b}"]'))`), 'its hit area again', 4000);
    }

    // Shift+click builds the selection: a first, then b added to it.
    await plainClick(`.sl-hit[data-shape="${a}"]`);
    await until(() => js(`document.querySelector('.sl-hit[data-shape="${a}"]')?.classList.contains('selected')`), 'a selected', 3000);
    await shiftClick(`.sl-hit[data-shape="${b}"]`);
    const two = await until(() => selectedCount().then((n) => n === 2), 'both selected', 3000).catch(() => false);
    const frame = await js(`Boolean(document.querySelector('.sl-selection-frame'))`);
    check('slides: Shift+click adds a second shape to the selection, with one dashed frame around both', two === true && frame === true, `selected ${await selectedCount()}; frame ${frame}`);

    // Align Left, from the Arrange menu — the leftmost of the two, since
    // nothing said "to the slide" and there are two of them.
    const leftmost = Math.min(shapeOf(a).geometry.x, shapeOf(b).geometry.x);
    const arranged = await clickRibbon('Arrange');
    await clickMenuItem('Align Left');
    const aligned = await until(() => Math.abs(shapeOf(a).geometry.x - leftmost) < 2 && Math.abs(shapeOf(b).geometry.x - leftmost) < 2, 'both at the left edge', 4000).catch(() => false);
    check('slides: Align Left, from the ribbon, lines both shapes up on their combined left edge', arranged === 'clicked' && aligned === true, `${arranged}; a ${Math.round(shapeOf(a).geometry.x)} b ${Math.round(shapeOf(b).geometry.x)} (target ${Math.round(leftmost)})`);

    // Group: one grpSp, read back from the model.
    await clickRibbon('Arrange');
    await clickMenuItem('Group');
    const grouped = await until(() => model().slide.shapes.some((s) => s.kind === 'group' && shapeOf(a).groupId === s.id), 'the group', 4000).catch(() => false);
    const group = model().slide.shapes.find((s) => s.kind === 'group');
    const oneHit = await until(() => js(`document.querySelectorAll('.sl-hit').length`).then((n) => n === before + 1), 'one hit area for the group', 4000).catch(() => false);
    check('slides: Group folds the two shapes into one p:grpSp, which the stage then shows as a single shape',
      grouped === true && Boolean(group) && shapeOf(b).groupId === group.id && oneHit === true,
      `group ${JSON.stringify(group?.geometry)}; a.groupId ${shapeOf(a).groupId}; hits ${await js(`document.querySelectorAll('.sl-hit').length`)}`);
    await wait(300);
    if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'slides-group-selected.png'), (await win.webContents.capturePage()).toPNG());

    // Move the group: the keyboard nudges the selected shape, which is now
    // the group's own id — both members move by the same amount.
    const a0 = { ...shapeOf(a).geometry };
    const b0 = { ...shapeOf(b).geometry };
    await js(`(() => { document.querySelector('.sl-stage')?.focus(); return 1; })()`);
    for (let i = 0; i < 4; i++) await press(wc, 'Right', { modifiers: ['shift'] });
    const moved = await until(() => shapeOf(a).geometry.x > a0.x + 20, 'the group to move', 4000).catch(() => false);
    const dx = shapeOf(a).geometry.x - a0.x;
    const movedTogether = Math.abs((shapeOf(b).geometry.x - b0.x) - dx) < 1;
    check('slides: moving the selected group carries both members by the same amount', moved === true && movedTogether === true, `a moved ${Math.round(dx)}, b moved ${Math.round(shapeOf(b).geometry.x - b0.x)}`);

    // Ungroup: the members are back at the top level, at their true position.
    const beforeUngroup = { a: { ...shapeOf(a).geometry }, b: { ...shapeOf(b).geometry } };
    await clickRibbon('Arrange');
    await clickMenuItem('Ungroup');
    const ungrouped = await until(() => !model().slide.shapes.some((s) => s.kind === 'group'), 'the group to go', 4000).catch(() => false);
    const positionsKept = Math.abs(shapeOf(a).geometry.x - beforeUngroup.a.x) < 2 && Math.abs(shapeOf(b).geometry.x - beforeUngroup.b.x) < 2;
    const bothSelectedAgain = await until(() => selectedCount().then((n) => n === 2), 'both members selected again after Ungroup', 3000).catch(() => false);
    check('slides: Ungroup puts the members back at the top level, at the position the group left them, and selects them both',
      ungrouped === true && positionsKept === true && bothSelectedAgain === true,
      `groupId now ${shapeOf(a).groupId}; a ${JSON.stringify(beforeUngroup.a)} → ${JSON.stringify(shapeOf(a).geometry)}`);

    // Rotate and Flip — one shape, from here on.
    await plainClick(`.sl-hit[data-shape="${a}"]`);
    await until(() => js(`document.querySelector('.sl-hit[data-shape="${a}"]')?.classList.contains('selected')`), 'a selected on its own', 3000);
    await clickRibbon('Arrange');
    await clickMenuItem('Rotate Right 90°');
    const rotated = await until(() => shapeOf(a).geometry.rot === 90, 'the rotation', 4000).catch(() => false);
    check('slides: Rotate Right 90°, from the ribbon, turns the shape and the engine records it', rotated === true, `rot ${shapeOf(a).geometry.rot}`);

    await clickRibbon('Arrange');
    await clickMenuItem('Flip Horizontal');
    const flipped = await until(() => shapeOf(a).geometry.flipH === true, 'the flip', 4000).catch(() => false);
    const stillRotated = shapeOf(a).geometry.rot === 90;
    check('slides: Flip Horizontal writes flipH and leaves the rotation exactly as it was', flipped === true && stillRotated === true, `flipH ${shapeOf(a).geometry.flipH}, rot ${shapeOf(a).geometry.rot}`);
    const drawnFlip = await until(() => js(`/scale\\(-1 1\\)/.test(document.querySelector('.sl-svg')?.innerHTML || '')`), 'the flip drawn on the stage', 4000).catch(() => false);
    check('slides: the stage draws the flip', drawnFlip === true, drawnFlip ? 'a mirrored transform is in the SVG' : 'not found');

    // Save, and read the rotation and the flip back off the file's own XML.
    await clickRibbon('Save');
    const inFile = () => { try { return Deck.open(fs.readFileSync(file)).slide(0).shapes.find((s) => String(s.id) === String(a)); } catch { return null; } };
    const saved = await until(() => inFile()?.geometry?.flipH === true && inFile()?.geometry?.rot === 90, 'the rotation and flip in the file', 8000).catch(() => false);
    check('slides: the saved file carries the rotation and the flip on the shape’s own a:xfrm', saved === true, JSON.stringify(inFile()?.geometry));

    // Clean up: the two rectangles come off, and the deck is saved as it was
    // found. The removal is the engine's own op, on the same session the
    // window is showing, so the window's next Save writes the file exactly
    // as this leaves it — the window's own cached model catching up is not
    // needed, since nothing after this reads it.
    await doc.apply({ id: sessionFor('deck').id, ops: [{ op: 'removeShape', slide: 0, shape: a }] });
    await doc.apply({ id: sessionFor('deck').id, ops: [{ op: 'removeShape', slide: 0, shape: b }] });
    await until(() => model().slide.shapes.length === before, 'the extra shapes to go', 4000).catch(() => {});
    await clickRibbon('Save');
    await until(() => { try { return Deck.open(fs.readFileSync(file)).slide(0).shapes.length === before; } catch { return false; } }, 'the deck back to how it was found', 8000).catch(() => false);
    check('slides: the arrange checks leave the deck as they found it', (() => { try { return Deck.open(fs.readFileSync(file)).slide(0).shapes.length === before; } catch { return false; } })(), `now ${(() => { try { return Deck.open(fs.readFileSync(file)).slide(0).shapes.length; } catch { return '?'; } })()} shapes, started with ${before}`);

    const complaints = await errorsIn(win);
    check('slides: the arrange checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('slides: the arrange checks ran', false, err.message);
  }
}
