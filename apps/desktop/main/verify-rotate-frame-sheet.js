// Worksheets: a turned shape's or picture's frame turns with it.
//
// A sheet with a long shape with words, a picture and a chart. The shape is
// turned to 30 degrees by its rotate handle (Shift snapping to 15), to 90 by
// Page Layout → Rotate → More Rotation Options…, and to 135 by its rotate
// handle again. At each angle, at 100% and 150%, the ring's four corner
// handles stand on the shape's turned corners and round what the sheet
// draws, a click at its centre picks it and a click in the empty corner of
// its old upright box does not. Undo and Redo bring the frame back with the
// angle, the arrows nudge it with its frame, and a corner drag resizes it
// along its own sides with the opposite corner staying put. The picture,
// turned by Rotate Right 90°, keeps its frame; a chart does not turn, as in
// Excel. On a sheet right to left the frame stays on the turned shape.
// Saved and opened again, the angles are in the drawing part (a:xfrm rot)
// and the frame is still on the shape. Run alone with
// RUTBA_VERIFY_ONLY=rotate-frame-sheet.

import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { gradientPng } from './sample-picture.js';
import { pointer, emptyCorner, shoot, centreOf, frameCorners, frameBounds } from './verify-rotate-frame.js';

/** The fixture: a small table, a long shape with words, a picture and a chart of the table. */
function fixture() {
  const view = new SheetView(buildXlsx({ sheets: [{ name: 'Plan', rows: [['Quarter', 'Sales'], ['Q1', 120], ['Q2', 140], ['Q3', 90]] }] }));
  view.select(5, 3);
  view.select(7, 7, { extend: true });
  view.insertShape({ geometry: 'rect', text: 'Turn me' });
  view.select(12, 2);
  view.insertPicture({ name: 'Logo', contentType: 'image/png', data: gradientPng(80, 30, [43, 95, 217], [230, 200, 60]), widthPx: 160, heightPx: 60 });
  view.select(0, 0);
  view.select(3, 1, { extend: true });
  view.insertChart({ title: 'Sales', kind: 'column' });
  // The chart off to the right, clear of the shape and the picture.
  const chart = view.render().drawings.find((x) => x.kind === 'chart');
  if (chart) view.setDrawingBox({ id: chart.id, x: 900, y: 40, width: 300, height: 200 });
  return view.save();
}

/** How far apart each expected point is from the nearest of the points got, at worst. */
function nearestGap(got, want) {
  if (!got || got.some((p) => !p)) return Infinity;
  return Math.max(...want.map((w) => Math.min(...got.map((g) => Math.hypot(g.x - w.x, g.y - w.y)))));
}

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args the run's scratch folder; the block writes in a folder of its own
 */
export async function verifyRotateFrameSheet(h, { dir }) {
  const { open, check, until, wait, press, errorsIn, doc, sessionFor } = h;
  const root = path.join(dir, 'rotate-frame-sheet');
  fs.mkdirSync(root, { recursive: true });
  const file = path.join(root, 'turned.xlsx');
  try {
    fs.writeFileSync(file, fixture());
  } catch (err) {
    return check('rotate-frame-sheet: the fixture is written', false, err.stack || err.message);
  }
  let win = await open('sheets', file);
  try {
    let wc = win.webContents;
    let js = (code) => wc.executeJavaScript(code);
    let mouse = pointer(win, wait);
    const session = () => sessionFor('sheet');
    const model = () => doc.model({ id: session().id });
    const byKind = (kind) => (model().drawings || []).find((d) => d.kind === kind);
    const SHAPE = 'shape';
    const PIC = 'image';
    const d = (kind) => byKind(kind);
    const sel = (kind) => `.sh-drawing[data-id="${d(kind)?.id}"]`;
    let zoom = 1;
    const rtl = () => js(`Boolean(document.querySelector('.sh.rtl'))`);
    /**
     * The drawing's upright box on the screen: its middle where the sheet
     * draws it (the middle of a turned box is where it was), its size the
     * model's at the grid's zoom.
     */
    const boxOf = async (kind) => {
      const g = d(kind);
      const c = await js(`(() => { const r = document.querySelector(${JSON.stringify(sel(kind))})?.getBoundingClientRect(); return r ? { x: r.left + r.width / 2, y: r.top + r.height / 2 } : null; })()`);
      if (!g || !c) return null;
      return { x: c.x - (g.width * zoom) / 2, y: c.y - (g.height * zoom) / 2, w: g.width * zoom, h: g.height * zoom };
    };
    const picked = () => js(`[...document.querySelectorAll('.sh-drawing.picked')].map((n) => n.dataset.id)`);
    const handles = () => js(`(() => ['nw', 'ne', 'se', 'sw'].map((n) => { const r = document.querySelector('.sh-obj-handle[data-handle="' + n + '"]')?.getBoundingClientRect(); return r && r.width ? { x: r.left + r.width / 2, y: r.top + r.height / 2 } : null; }))()`);
    /** What the sheet draws for the drawing: the box round its SVG's painted parts. */
    const drawn = (kind) => js(`(() => {
      const svg = document.querySelector(${JSON.stringify(sel(kind))} + ' > svg');
      if (!svg) return null;
      const rs = [...svg.children].filter((n) => n.tagName !== 'title' && n.tagName !== 'defs').map((n) => n.getBoundingClientRect()).filter((r) => r.width || r.height);
      if (!rs.length) return null;
      const x = Math.min(...rs.map((r) => r.left)); const y = Math.min(...rs.map((r) => r.top));
      return { x, y, w: Math.max(...rs.map((r) => r.right)) - x, h: Math.max(...rs.map((r) => r.bottom)) - y };
    })()`);
    const clickRibbon = (title) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || n.textContent || '').trim().startsWith(${JSON.stringify(title)}) && !n.disabled);
      if (!b) return 'no button ' + ${JSON.stringify(title)};
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      b.click();
      return 'clicked';
    })()`);
    const menuPick = async (label) => {
      await js(`(() => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Page Layout')?.click(); return 1; })()`);
      await wait(200);
      const pressed = await clickRibbon('Rotate');
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)} && !b.disabled))`), `"${label}"`, 3000).catch(() => false);
      const chosen = await js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}); if (!b || b.disabled) return 'missing'; b.click(); return 'clicked'; })()`);
      return pressed === 'clicked' && chosen === 'clicked' ? 'clicked' : `${pressed}/${chosen}`;
    };
    const setZoom = async (z) => {
      await js(`(() => { const input = document.querySelector('.rw-zoom input[type="range"]'); if (!input) return 0; const set = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set; set.call(input, ${JSON.stringify(String(z))}); input.dispatchEvent(new Event('input', { bubbles: true })); input.dispatchEvent(new Event('change', { bubbles: true })); return 1; })()`);
      const ok = await until(() => js(`(document.querySelector('.rw-zoom-pct')?.textContent || '') === '${Math.round(z * 100)}%'`), `the sheet at ${z * 100}%`, 4000).then(() => true).catch(() => false);
      zoom = z;
      await wait(300);
      return ok;
    };
    const bringIntoView = (kind) => js(`(() => { document.querySelector(${JSON.stringify(sel(kind))})?.scrollIntoView({ block: 'center', inline: 'center' }); return 1; })()`);
    const centreOn = async (kind) => { const b = await boxOf(kind); return b ? { x: b.x + b.w / 2, y: b.y + b.h / 2 } : null; };
    const letGo = async () => {
      await js(`document.querySelector('.sh')?.focus(), 1`);
      await press(wc, 'Escape');
      await until(async () => (await picked()).length === 0, 'nothing picked', 2000).catch(() => {});
    };
    const select = async (kind) => {
      await bringIntoView(kind);
      await wait(150);
      const at = await centreOn(kind);
      if (at) await mouse.click(at);
      return until(async () => JSON.stringify(await picked()) === JSON.stringify([d(kind)?.id]), `the ${kind} picked`, 3000).then(() => true).catch(() => false);
    };

    /** The ring's corner handles against the drawing's turned corners, and what is drawn within them. */
    const frameOn = async (kind) => {
      let last = null;
      const ok = await until(async () => {
        const b = await boxOf(kind);
        const rot = d(kind)?.rot || 0;
        if (!b) return false;
        const want = frameCorners(b, rot);
        const got = await handles();
        const bounds = frameBounds(b, rot);
        const art = await drawn(kind);
        // The drawing's painted parts lie within the turned box's bounds (a stroke's width over).
        const over = art ? Math.max(bounds.x - art.x, bounds.y - art.y, art.x + art.w - (bounds.x + bounds.w), art.y + art.h - (bounds.y + bounds.h)) : Infinity;
        // And they fill it: a rectangle or a picture reaches every side.
        const short = art ? Math.max(art.x - bounds.x, art.y - bounds.y, bounds.x + bounds.w - (art.x + art.w), bounds.y + bounds.h - (art.y + art.h)) : Infinity;
        last = { rot, z: zoom, corners: +nearestGap(got, want).toFixed(1), drawnOver: +over.toFixed(1), drawnShort: +short.toFixed(1) };
        return last.corners <= 3 && last.drawnOver <= 4 && last.drawnShort <= 4;
      }, `the frame on the ${kind}`, 4000).then(() => true).catch(() => false);
      return { ok, ...last };
    };

    /** A click at the centre picks it; one in the old box's empty corner does not. */
    const clicks = async (kind) => {
      await letGo();
      await bringIntoView(kind);
      await wait(150);
      const b = await boxOf(kind);
      const rot = d(kind)?.rot || 0;
      await mouse.click({ x: b.x + b.w / 2, y: b.y + b.h / 2 });
      const centreTakes = await until(async () => (await picked()).includes(d(kind)?.id), 'the centre to pick', 3000).then(() => true).catch(() => false);
      await letGo();
      const corner = emptyCorner(b, rot, 5);
      const under = corner ? await js(`(() => { const n = document.elementFromPoint(${corner.x}, ${corner.y}); return n?.closest('.sh-drawing')?.dataset.id || (n?.closest('.sh-cells') ? 'cells' : n?.className || 'nothing'); })()`) : 'no empty corner';
      if (corner) await mouse.click(corner);
      await wait(250);
      const cornerTakes = (await picked()).includes(d(kind)?.id);
      return { ok: centreTakes && Boolean(corner) && !cornerTakes && under !== d(kind)?.id, centreTakes, cornerTakes, under };
    };

    const atAngle = async (kind, label) => {
      const results = [];
      for (const z of [1, 1.5]) {
        const zoomed = await setZoom(z);
        const held = await select(kind);
        const frame = await frameOn(kind);
        await shoot(win, wait, `rotate-frame-sheet-${label}${z === 1 ? '' : '-150'}.png`);
        const click = await clicks(kind);
        results.push({ z, zoomed, held, frame, click });
      }
      await setZoom(1);
      return { ok: results.every((r) => r.zoomed && r.held && r.frame.ok && r.click.ok), results };
    };

    /** Drag the rotate handle round the drawing's centre to `deg`, Shift snapping to 15. */
    const turnTo = async (kind, deg) => {
      await select(kind);
      const centre = await centreOn(kind);
      await until(() => js(centreOf('.sh-obj-rotate')), 'the rotate handle', 3000).catch(() => null);
      const grip = await js(centreOf('.sh-obj-rotate'));
      if (!grip || !centre) return false;
      const r = Math.hypot(grip.x - centre.x, grip.y - centre.y);
      const a = (deg * Math.PI) / 180;
      await mouse.drag(grip, { x: centre.x + r * Math.sin(a), y: centre.y - r * Math.cos(a) }, { modifiers: ['shift'], steps: 12 });
      return until(() => Math.round(d(kind)?.rot || 0) === deg, `the ${kind} at ${deg}`, 5000).then(() => true).catch(() => false);
    };

    await until(() => js(`document.querySelectorAll('.sh-drawing').length >= 3`), 'the drawings', 10000);
    await setZoom(1);

    // 30, by the rotate handle.
    const turned30 = await turnTo(SHAPE, 30);
    const at30 = await atAngle(SHAPE, '30');
    check('rotate-frame-sheet: a shape turned to 30 by its rotate handle wears its ring, eight handles and rotate handle turned with it at 100% and 150%; a click at its centre picks it, one in the empty corner of its old upright box does not',
      turned30 && at30.ok, JSON.stringify({ rot: d(SHAPE)?.rot, ...at30 }));

    // A corner drag at 30: along its own sides, the opposite corner staying put.
    await select(SHAPE);
    const b0 = await boxOf(SHAPE);
    const g0 = { w: d(SHAPE).width, h: d(SHAPE).height };
    const fixed = frameCorners(b0, 30)[0];
    const se = (await handles())[2];
    const a30 = Math.PI / 6;
    const along = { x: 60 * Math.cos(a30) - 30 * Math.sin(a30), y: 60 * Math.sin(a30) + 30 * Math.cos(a30) };
    if (se) await mouse.drag(se, { x: se.x + along.x, y: se.y + along.y }, { steps: 8 });
    const resized = await until(() => Math.abs((d(SHAPE)?.width || 0) - (g0.w + 60)) <= 3, 'the shape resized', 5000).then(() => true).catch(() => false);
    let moved = Infinity;
    await until(async () => {
      const b1 = await boxOf(SHAPE);
      const kept = frameCorners(b1, d(SHAPE)?.rot || 0)[0];
      moved = +Math.hypot(kept.x - fixed.x, kept.y - fixed.y).toFixed(1);
      return moved <= 3;
    }, 'the opposite corner to stay put', 4000).catch(() => false);
    const reframed = await frameOn(SHAPE);
    await shoot(win, wait, 'rotate-frame-sheet-resized.png');
    check('rotate-frame-sheet: a corner drag on the turned shape resizes it along its own sides, the opposite corner staying put on the sheet, and the ring follows',
      resized && Math.abs(d(SHAPE).height - (g0.h + 30)) <= 3 && moved <= 3 && d(SHAPE)?.rot === 30 && reframed.ok,
      JSON.stringify({ before: g0, after: { w: d(SHAPE)?.width, h: d(SHAPE)?.height }, moved, reframed }));

    // 90, by More Rotation Options…
    await select(SHAPE);
    const opened = await menuPick('More Rotation Options…');
    await until(() => js(`Boolean(document.querySelector('.rotation-angle'))`), 'the Rotation dialog', 3000).catch(() => false);
    const shown = await js(`document.querySelector('.rotation-angle')?.value`);
    await js(`(() => { const i = document.querySelector('.rotation-angle'); if (!i) return 0; const set = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set; set.call(i, '90'); i.dispatchEvent(new Event('input', { bubbles: true })); document.querySelector('.rotation-ok')?.click(); return 1; })()`);
    const turned90 = await until(() => Math.round(d(SHAPE)?.rot || 0) === 90, 'the shape at 90', 5000).then(() => true).catch(() => false);
    const at90 = await atAngle(SHAPE, '90');
    check('rotate-frame-sheet: Page Layout → Rotate → More Rotation Options… shows the angle and turns the shape to the one typed; at 90 the ring stays on it at 100% and 150% and the clicks land by its turned outline',
      opened === 'clicked' && shown === '30' && turned90 && at90.ok, JSON.stringify({ opened, shown, rot: d(SHAPE)?.rot, ...at90 }));

    // 135, by the rotate handle; Undo and Redo.
    const turned135 = await turnTo(SHAPE, 135);
    const at135 = await atAngle(SHAPE, '135');
    check('rotate-frame-sheet: turned on to 135, the ring stays on the shape at 100% and 150% and the clicks land by its turned outline',
      turned135 && at135.ok, JSON.stringify({ rot: d(SHAPE)?.rot, ...at135 }));
    await select(SHAPE);
    await js(`document.querySelector('.sh')?.focus(), 1`);
    await press(wc, 'z', { modifiers: ['control'] });
    const undone = await until(() => Math.round(d(SHAPE)?.rot || 0) === 90, 'undo to 90', 4000).then(() => true).catch(() => false);
    await select(SHAPE);
    const undoFrame = await frameOn(SHAPE);
    await js(`document.querySelector('.sh')?.focus(), 1`);
    await press(wc, 'y', { modifiers: ['control'] });
    const redone = await until(() => Math.round(d(SHAPE)?.rot || 0) === 135, 'redo to 135', 4000).then(() => true).catch(() => false);
    await select(SHAPE);
    const redoFrame = await frameOn(SHAPE);
    check('rotate-frame-sheet: Undo takes the last turn back and Redo puts it again, the ring following each',
      undone && redone && undoFrame.ok && redoFrame.ok, JSON.stringify({ undone, undoFrame, redone, redoFrame }));

    // The arrows nudge it, and the ring comes along.
    const x0 = d(SHAPE).x;
    await js(`document.querySelector('.sh')?.focus(), 1`);
    for (let i = 0; i < 4; i++) await press(wc, 'Right');
    const nudged = await until(() => d(SHAPE).x > x0, 'the nudge', 4000).then(() => true).catch(() => false);
    const nudgeFrame = await frameOn(SHAPE);
    check('rotate-frame-sheet: the arrows nudge the turned shape and its ring comes along', nudged && nudgeFrame.ok, JSON.stringify({ moved: d(SHAPE).x - x0, nudgeFrame }));

    // The picture, by Rotate Right 90°; a chart does not turn.
    await select(PIC);
    const r90 = await menuPick('Rotate Right 90°');
    const picTurned = await until(() => d(PIC)?.rot === 90, 'the picture at 90', 4000).then(() => true).catch(() => false);
    await select(PIC);
    const picFrame = await frameOn(PIC);
    await shoot(win, wait, 'rotate-frame-sheet-picture.png');
    const picClicks = await clicks(PIC);
    check('rotate-frame-sheet: a picture turned by Rotate Right 90° keeps its ring on it and takes clicks by its turned outline',
      r90 === 'clicked' && picTurned && picFrame.ok && picClicks.ok, JSON.stringify({ r90, picFrame, picClicks }));
    await select('chart');
    await js(`(() => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Page Layout')?.click(); return 1; })()`);
    await wait(200);
    const chartRotate = await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Rotate')); return b ? { disabled: b.disabled, tip: b.dataset.tip || b.title } : null; })()`);
    const chartHandle = await js(`Boolean(document.querySelector('.sh-obj-rotate'))`);
    check('rotate-frame-sheet: a chart does not turn, as in Excel: Rotate is off for it and it wears no rotate handle',
      chartRotate?.disabled === true && chartHandle === false, JSON.stringify({ chartRotate, chartHandle }));

    // Right to left.
    await letGo();
    const toRtl = await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Sheet Right-to-Left')); return b ? (b.click(), 'clicked') : 'none'; })()`);
    let rtlFrame = { ok: false };
    let rtlClicks = { ok: false };
    if (toRtl === 'clicked' && await until(() => rtl(), 'the sheet right to left', 4000).then(() => true).catch(() => false)) {
      await select(SHAPE);
      rtlFrame = await frameOn(SHAPE);
      await shoot(win, wait, 'rotate-frame-sheet-rtl.png');
      rtlClicks = await clicks(SHAPE);
      await js(`(() => { [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Sheet Right-to-Left'))?.click(); return 1; })()`);
      await until(async () => !(await rtl()), 'the sheet left to right again', 4000).catch(() => {});
    }
    check('rotate-frame-sheet: on a sheet right to left, the turned shape’s ring stays on it and the clicks land by its turned outline',
      rtlFrame.ok && rtlClicks.ok, JSON.stringify({ toRtl, rtlFrame, rtlClicks }));

    // Saved and opened again.
    await letGo();
    await js(`document.querySelector('.sh')?.focus(), 1`);
    await press(wc, 's', { modifiers: ['control'] });
    const rotsInFile = () => {
      try {
        const v = new SheetView(fs.readFileSync(file));
        const ds = v.render().drawings || [];
        return ds.map((x) => ({ kind: x.kind, rot: x.rot || 0 }));
      } catch { return []; }
    };
    const saved = await until(() => { const r = rotsInFile(); return r.some((x) => x.kind === 'shape' && x.rot === 135) && r.some((x) => x.kind === 'image' && x.rot === 90); }, 'the angles in the file', 8000).then(() => true).catch(() => false);
    const sid = session()?.id;
    const keep = path.join(root, 'keep.xlsx');
    fs.writeFileSync(keep, buildXlsx({ sheets: [{ name: 'Keep', rows: [['Keep']] }] }));
    await open('sheets', keep);
    win.close();
    await until(() => win.isDestroyed(), 'the window closed', 4000).catch(() => {});
    if (sid && doc.sessions().some((s) => s.id === sid)) await doc.close({ id: sid }).catch(() => {});
    win = await open('sheets', file);
    wc = win.webContents;
    js = (code) => wc.executeJavaScript(code);
    mouse = pointer(win, wait);
    await until(() => js(`document.querySelectorAll('.sh-drawing').length >= 3`), 'the drawings again', 10000).catch(() => {});
    zoom = 1;
    await setZoom(1);
    const reopenedRot = d(SHAPE)?.rot;
    await select(SHAPE);
    const reopenedFrame = await frameOn(SHAPE);
    await shoot(win, wait, 'rotate-frame-sheet-reopened.png');
    const reopenedClicks = await clicks(SHAPE);
    check('rotate-frame-sheet: saved and opened again, the shape and the picture keep their angles and the ring is still on the turned shape',
      saved && reopenedRot === 135 && reopenedFrame.ok && reopenedClicks.ok, JSON.stringify({ saved, file: rotsInFile(), reopenedRot, reopenedFrame, reopenedClicks }));

    const complaints = await errorsIn(win);
    check('rotate-frame-sheet: the turned-frame checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('rotate-frame-sheet: the turned-frame checks ran', false, err.stack || err.message);
  }
}
