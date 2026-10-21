// Presentations: a turned shape's frame turns with it.
//
// A long box with words, a picture and a group of two on a blank slide. The
// box is turned to 30 degrees by its rotate handle (Shift snapping to 15),
// to 90 by Arrange → Rotate → More Rotation Options…, and to 135 by
// Alt+Right three times. At each angle, at 100% and 150%, the frame's four
// corner handles stand on the box's turned corners and round what the slide
// draws, a click at its centre takes hold of it and a click in the empty
// corner of its old upright box does not. Undo and Redo bring the frame
// back with the angle; a corner drag resizes along the box's own sides with
// the opposite corner staying put; a double-click opens its words, the
// editor turned with it. The picture turned by Rotate Right 90° and
// flipped, the group turned by the keyboard and the dashed frame of a
// multiple selection all follow. Saved and opened again, every angle is in
// the file and the frame still on the box. Run alone with
// RUTBA_VERIFY_ONLY=rotate-frame-deck.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';
import { gradientPng } from './sample-picture.js';
import { pointer, worstGap, emptyCorner, shoot, centreOf, frameCorners, frameBounds } from './verify-rotate-frame.js';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args the run's scratch folder; the block writes in a folder of its own
 */
export async function verifyRotateFrameDeck(h, { dir }) {
  const { open, check, until, wait, press, errorsIn, doc, sessionFor } = h;
  const root = path.join(dir, 'rotate-frame-deck');
  fs.mkdirSync(root, { recursive: true });
  const file = path.join(root, 'turned.pptx');
  const deck = Deck.open(buildPptx({ title: 'Turned', slides: [{ layout: 'blank' }, { layout: 'blank' }] }));
  const boxId = deck.addShape(0, { preset: 'rect', x: 460, y: 300, w: 340, h: 90, name: 'Box', text: 'Turn me' });
  const picId = deck.addPicture(0, { data: gradientPng(160, 100, [40, 120, 200], [230, 200, 60]), contentType: 'image/png', name: 'Pic', x: 120, y: 470, w: 160, h: 100 });
  const g1 = deck.addShape(0, { preset: 'rect', x: 940, y: 110, w: 120, h: 60, name: 'G1' });
  const g2 = deck.addShape(0, { preset: 'ellipse', x: 1080, y: 110, w: 110, h: 60, name: 'G2' });
  deck.groupShapes(0, [g1, g2]);
  fs.writeFileSync(file, deck.save());
  const ids = Object.fromEntries(Deck.open(fs.readFileSync(file)).slide(0).shapes.map((s) => [s.name, String(s.id)]));
  void boxId; void picId;

  let win = await open('slides', file);
  try {
    let wc = win.webContents;
    let js = (code) => wc.executeJavaScript(code);
    let mouse = pointer(win, wait);
    const session = () => sessionFor('deck');
    const model = () => doc.model({ id: session().id, slide: 0 });
    const shape = (name) => model().slide.shapes.find((s) => s.name === name);
    const geo = (name) => shape(name)?.geometry;

    // The slide on screen: where its top left is and how much it is scaled.
    const stage = () => js(`(() => { const s = document.querySelector('.sl-slide'); if (!s) return null; const r = s.getBoundingClientRect(); return { left: r.left, top: r.top, k: r.width / s.offsetWidth }; })()`);
    const toScreen = (st, p) => ({ x: st.left + p.x * st.k, y: st.top + p.y * st.k });
    const handles = () => js(`(() => ['nw', 'ne', 'se', 'sw'].map((n) => { const r = document.querySelector('.sl-handle[data-handle="' + n + '"]')?.getBoundingClientRect(); return r && r.width ? { x: r.left + r.width / 2, y: r.top + r.height / 2 } : null; }))()`);
    // What the slide draws for a shape: its own group in the SVG, or, for a group, its members' together.
    const drawn = (list) => js(`(() => {
      const rs = ${JSON.stringify(list)}.map((id) => document.querySelector('.sl-svg g[data-shape="' + id + '"]')?.getBoundingClientRect()).filter(Boolean);
      if (!rs.length) return null;
      const x = Math.min(...rs.map((r) => r.left)); const y = Math.min(...rs.map((r) => r.top));
      return { x, y, w: Math.max(...rs.map((r) => r.right)) - x, h: Math.max(...rs.map((r) => r.bottom)) - y };
    })()`);
    const selectedIds = () => js(`[...document.querySelectorAll('.sl-hit.selected')].map((n) => n.dataset.shape)`);
    const hitAt = (p) => js(`(() => { const n = document.elementFromPoint(${p.x}, ${p.y}); return n?.closest('.sl-hit')?.dataset.shape || (n ? n.className?.baseVal ?? n.className : null); })()`);
    const clickRibbon = (title) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}) && !n.disabled);
      if (!b) return 'no button ' + ${JSON.stringify(title)};
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      b.click();
      return 'clicked';
    })()`);
    const menuPick = async (button, label) => {
      const pressed = await clickRibbon(button);
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)} && !b.disabled))`), `"${label}"`, 3000).catch(() => false);
      const picked = await js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}); if (!b || b.disabled) return 'missing'; b.click(); return 'clicked'; })()`);
      return pressed === 'clicked' && picked === 'clicked' ? 'clicked' : `${pressed}/${picked}`;
    };
    const setZoom = async (z) => {
      await js(`(() => { const input = document.querySelector('.rw-zoom input[type="range"]'); if (!input) return 0; const set = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set; set.call(input, ${JSON.stringify(String(z))}); input.dispatchEvent(new Event('input', { bubbles: true })); input.dispatchEvent(new Event('change', { bubbles: true })); return 1; })()`);
      return until(async () => Math.abs(((await stage())?.k || 0) - z) < 0.01, `the slide at ${z * 100}%`, 4000).then(() => true).catch(() => false);
    };
    // The named shape in the middle of the stage, so a click can reach it at any zoom.
    const bringIntoView = (name) => js(`(() => { document.querySelector('.sl-hit[data-shape="${ids[name]}"]')?.scrollIntoView({ block: 'center', inline: 'center' }); return 1; })()`);
    const centreOnScreen = async (name) => {
      const st = await stage();
      const g = geo(name);
      return toScreen(st, { x: g.x + g.w / 2, y: g.y + g.h / 2 });
    };
    const select = async (name) => {
      await bringIntoView(name);
      await wait(150);
      const at = await centreOnScreen(name);
      await mouse.click(at);
      return until(async () => JSON.stringify(await selectedIds()) === JSON.stringify([ids[name]]), `${name} selected`, 3000).then(() => true).catch(() => false);
    };
    const letGo = async () => {
      await js(`document.querySelector('.sl-stage')?.focus(), 1`);
      await press(wc, 'Escape');
      await until(async () => (await selectedIds()).length === 0, 'nothing selected', 2000).catch(() => {});
    };

    /**
     * The frame on the named shape, as drawn now: its corner handles against
     * the shape's turned corners and the drawing's own box, waited for, since
     * the window draws a beat after the engine answers.
     */
    const frameOn = async (name) => {
      let last = null;
      const ok = await until(async () => {
        const st = await stage();
        const g = geo(name);
        if (!st || !g) return false;
        const want = frameCorners(g, g.rot || 0).map((p) => toScreen(st, p));
        const got = await handles();
        const members = model().slide.shapes.filter((s) => String(s.groupId) === ids[name]).map((s) => String(s.id));
        const box = await drawn(members.length ? members : [ids[name]]);
        const b = frameBounds(g, g.rot || 0);
        const sides = box ? [box.x - (st.left + b.x * st.k), box.y - (st.top + b.y * st.k), (st.left + (b.x + b.w) * st.k) - (box.x + box.w), (st.top + (b.y + b.h) * st.k) - (box.y + box.h)] : null;
        // A shape's drawing fills its turned box's bounds; a group's members lie within them.
        const around = !sides ? Infinity : members.length ? Math.max(0, ...sides.map((d) => -d)) : Math.max(...sides.map(Math.abs));
        last = { rot: g.rot || 0, k: +st.k.toFixed(2), corners: +worstGap(got, want).toFixed(1), drawn: +around.toFixed(1) };
        return last.corners <= 3 && last.drawn <= 4;
      }, `the frame on ${name}`, 4000).then(() => true).catch(() => false);
      return { ok, ...last };
    };

    /** Centre takes hold; the old box's empty corner does not. */
    const clicks = async (name) => {
      await letGo();
      const st = await stage();
      const g = geo(name);
      const centre = toScreen(st, { x: g.x + g.w / 2, y: g.y + g.h / 2 });
      await mouse.click(centre);
      const centreTakes = await until(async () => (await selectedIds()).includes(ids[name]), 'the centre to select', 3000).then(() => true).catch(() => false);
      await letGo();
      const corner = emptyCorner(g, g.rot || 0);
      const cornerAt = corner ? toScreen(st, corner) : null;
      const under = cornerAt ? await hitAt(cornerAt) : 'no empty corner';
      if (cornerAt) await mouse.click(cornerAt);
      await wait(200);
      const cornerTakes = (await selectedIds()).includes(ids[name]);
      return { ok: centreTakes && Boolean(cornerAt) && !cornerTakes && under !== ids[name], centreTakes, cornerTakes, under };
    };

    /** The whole spell at one angle: the frame and the clicks at 100% and 150%. */
    const atAngle = async (name, label) => {
      const results = [];
      for (const z of [1, 1.5]) {
        const zoomed = await setZoom(z);
        await bringIntoView(name);
        const held = await select(name);
        const frame = await frameOn(name);
        if (z === 1.5) await shoot(win, wait, `rotate-frame-deck-${label}-150.png`);
        else await shoot(win, wait, `rotate-frame-deck-${label}.png`);
        const click = await clicks(name);
        results.push({ z, zoomed, held, frame, click });
      }
      await setZoom(1);
      return { ok: results.every((r) => r.zoomed && r.held && r.frame.ok && r.click.ok), results };
    };

    await until(() => js(`Boolean(document.querySelector('.sl-hit[data-shape="${ids.Box}"]'))`), 'the box on the stage', 10000);
    await setZoom(1);

    // 30 degrees, by the rotate handle, Shift snapping to 15.
    await select('Box');
    const st = await stage();
    const g0 = geo('Box');
    const centre = toScreen(st, { x: g0.x + g0.w / 2, y: g0.y + g0.h / 2 });
    await until(() => js(centreOf('.sl-rotate-handle')), 'the rotate handle', 3000).catch(() => null);
    const grip = await js(centreOf('.sl-rotate-handle'));
    const radius = grip ? Math.hypot(grip.x - centre.x, grip.y - centre.y) : 80;
    const a30 = (30 * Math.PI) / 180;
    if (grip) await mouse.drag(grip, { x: centre.x + radius * Math.sin(a30), y: centre.y - radius * Math.cos(a30) }, { modifiers: ['shift'], steps: 10 });
    const turned30 = await until(() => Math.round(geo('Box')?.rot || 0) === 30, 'the box at 30', 5000).then(() => true).catch(() => false);
    const at30 = await atAngle('Box', '30');
    check('rotate-frame-deck: turned to 30 by its rotate handle, the frame, its handles and its rotate handle turn with the box at 100% and 150%; a click at its centre takes hold of it, one in the empty corner of its old upright box does not',
      turned30 && at30.ok, JSON.stringify({ rot: geo('Box')?.rot, ...at30 }));

    // A corner drag at 30: along the box's own sides, the opposite corner staying put.
    await select('Box');
    const before = { ...geo('Box') };
    const fixed = frameCorners(before, 30)[0];
    const se = (await handles())[2];
    const stNow = await stage();
    const along = { x: 60 * Math.cos(a30) - 30 * Math.sin(a30), y: 60 * Math.sin(a30) + 30 * Math.cos(a30) };
    if (se) await mouse.drag(se, { x: se.x + along.x * stNow.k, y: se.y + along.y * stNow.k }, { steps: 8 });
    const resized = await until(() => Math.abs((geo('Box')?.w || 0) - (before.w + 60)) <= 3, 'the box resized', 5000).then(() => true).catch(() => false);
    const after = geo('Box');
    const kept = frameCorners(after, after.rot || 0)[0];
    const sized = { w: Math.round(after.w), h: Math.round(after.h), rot: after.rot, moved: +Math.hypot(kept.x - fixed.x, kept.y - fixed.y).toFixed(1) };
    const reframed = await frameOn('Box');
    await shoot(win, wait, 'rotate-frame-deck-resized.png');
    check('rotate-frame-deck: a corner drag on the turned box resizes it along its own sides, the opposite corner staying put on the slide, and the frame follows',
      resized && Math.abs(after.w - (before.w + 60)) <= 3 && Math.abs(after.h - (before.h + 30)) <= 3 && sized.moved <= 3 && after.rot === 30 && reframed.ok,
      JSON.stringify({ before: { w: before.w, h: before.h }, ...sized, reframed }));

    // A double-click on its words opens them, the editor turned with the box.
    await letGo();
    const mid = await centreOnScreen('Box');
    await mouse.click(mid, { count: 2 });
    await until(() => js(`Boolean(document.querySelector('.sl-editor'))`), 'the editor', 3000).catch(() => null);
    const editor = await js(`(() => { const e = document.querySelector('.sl-editor'); return e ? { value: e.innerText.trim(), turn: e.style.transform } : null; })()`);
    await shoot(win, wait, 'rotate-frame-deck-editing.png');
    check('rotate-frame-deck: a double-click on the turned box\'s words opens them for typing, the editor turned with the box',
      editor?.value === 'Turn me' && /rotate\(30deg\)/.test(editor?.turn || ''), JSON.stringify(editor));
    await press(wc, 'Escape');
    await until(() => js(`!document.querySelector('.sl-editor')`), 'the editor to close', 2000).catch(() => {});

    // 90, by More Rotation Options…
    await select('Box');
    const opened = await menuPick('Arrange', 'More Rotation Options…');
    await until(() => js(`Boolean(document.querySelector('.rotation-angle'))`), 'the Rotation dialog', 3000).catch(() => false);
    const shown = await js(`document.querySelector('.rotation-angle')?.value`);
    await js(`(() => { const i = document.querySelector('.rotation-angle'); if (!i) return 0; const set = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set; set.call(i, '90'); i.dispatchEvent(new Event('input', { bubbles: true })); document.querySelector('.rotation-ok')?.click(); return 1; })()`);
    const turned90 = await until(() => Math.round(geo('Box')?.rot || 0) === 90, 'the box at 90', 5000).then(() => true).catch(() => false);
    const at90 = await atAngle('Box', '90');
    check('rotate-frame-deck: Arrange → Rotate → More Rotation Options… shows the angle and turns the box to the one typed; at 90 the frame stays on it at 100% and 150% and the clicks land by its turned outline',
      opened === 'clicked' && shown === '30' && turned90 && at90.ok, JSON.stringify({ opened, shown, rot: geo('Box')?.rot, ...at90 }));

    // 135, by Alt+Right three times; Undo and Redo bring the frame back with the angle.
    await select('Box');
    for (let i = 0; i < 3; i++) { await press(wc, 'Right', { modifiers: ['alt'] }); await wait(250); }
    const turned135 = await until(() => Math.round(geo('Box')?.rot || 0) === 135, 'the box at 135', 5000).then(() => true).catch(() => false);
    const at135 = await atAngle('Box', '135');
    check('rotate-frame-deck: Alt+Right three times turns the box to 135; the frame stays on it at 100% and 150% and the clicks land by its turned outline',
      turned135 && at135.ok, JSON.stringify({ rot: geo('Box')?.rot, ...at135 }));
    await select('Box');
    await press(wc, 'z', { modifiers: ['control'] });
    const undone = await until(() => Math.round(geo('Box')?.rot || 0) === 120, 'undo to 120', 4000).then(() => true).catch(() => false);
    await select('Box');
    const undoFrame = await frameOn('Box');
    await press(wc, 'y', { modifiers: ['control'] });
    const redone = await until(() => Math.round(geo('Box')?.rot || 0) === 135, 'redo to 135', 4000).then(() => true).catch(() => false);
    await select('Box');
    const redoFrame = await frameOn('Box');
    check('rotate-frame-deck: Undo takes the last turn back and Redo puts it again, the frame following each',
      undone && redone && undoFrame.ok && redoFrame.ok, JSON.stringify({ undone, undoFrame, redone, redoFrame }));

    // Nudged with the arrows, the frame comes along.
    const x0 = geo('Box').x;
    for (let i = 0; i < 2; i++) await press(wc, 'Right', { modifiers: ['shift'] });
    const nudged = await until(() => geo('Box').x >= x0 + 19, 'the nudge', 4000).then(() => true).catch(() => false);
    const nudgeFrame = await frameOn('Box');
    check('rotate-frame-deck: Shift+arrow nudges the turned box and its frame comes along', nudged && nudgeFrame.ok, JSON.stringify({ moved: geo('Box').x - x0, nudgeFrame }));

    // The picture: Rotate Right 90°, then Flip Horizontal.
    await select('Pic');
    const r90 = await menuPick('Arrange', 'Rotate Right 90°');
    const picTurned = await until(() => geo('Pic')?.rot === 90, 'the picture at 90', 4000).then(() => true).catch(() => false);
    const picFrame = await frameOn('Pic');
    await shoot(win, wait, 'rotate-frame-deck-picture.png');
    const picClicks = await clicks('Pic');
    await select('Pic');
    const flip = await menuPick('Arrange', 'Flip Horizontal');
    const flipped = await until(() => geo('Pic')?.flipH === true, 'the flip', 4000).then(() => true).catch(() => false);
    const flipFrame = await frameOn('Pic');
    check('rotate-frame-deck: a picture turned by Rotate Right 90° and then flipped keeps its frame on it, and takes clicks by its turned outline',
      r90 === 'clicked' && picTurned && picFrame.ok && picClicks.ok && flip === 'clicked' && flipped && flipFrame.ok,
      JSON.stringify({ r90, picFrame, picClicks, flip, flipFrame }));

    // The group, turned 30 by the keyboard: one frame on the whole group.
    const groupName = Object.keys(ids).find((n) => /^Group /.test(n)) || null;
    if (groupName) await select(groupName);
    if (groupName) for (let i = 0; i < 2; i++) { await press(wc, 'Right', { modifiers: ['alt'] }); await wait(250); }
    const groupTurned = groupName ? await until(() => Math.round(geo(groupName)?.rot || 0) === 30, 'the group at 30', 4000).then(() => true).catch(() => false) : false;
    const groupFrame = groupName ? await frameOn(groupName) : { ok: false };
    await shoot(win, wait, 'rotate-frame-deck-group.png');
    const groupClicks = groupName ? await clicks(groupName) : { ok: false };
    check('rotate-frame-deck: a group turned with Alt+Right wears one frame turned with it, and takes clicks by its turned outline',
      groupTurned && groupFrame.ok && groupClicks.ok, JSON.stringify({ groupName, rot: geo(groupName)?.rot, groupFrame, groupClicks }));

    // Several selected: one dashed frame round the turned outlines of all of them.
    await select('Box');
    const pic = await centreOnScreen('Pic');
    await mouse.click(pic, { modifiers: ['shift'] });
    const two = await until(async () => (await selectedIds()).length === 2, 'two selected', 3000).then(() => true).catch(() => false);
    let multi = false;
    await until(async () => {
      const r = await js(`(() => { const f = document.querySelector('.sl-selection-frame')?.getBoundingClientRect(); return f ? { x: f.left, y: f.top, w: f.width, h: f.height } : null; })()`);
      if (!r) return false;
      const s = await stage();
      const bs = ['Box', 'Pic'].map((n) => frameBounds(geo(n), geo(n).rot || 0));
      const x = Math.min(...bs.map((b) => b.x));
      const y = Math.min(...bs.map((b) => b.y));
      const right = Math.max(...bs.map((b) => b.x + b.w));
      const bottom = Math.max(...bs.map((b) => b.y + b.h));
      const gap = Math.max(Math.abs(r.x - (s.left + x * s.k)), Math.abs(r.y - (s.top + y * s.k)), Math.abs(r.x + r.w - (s.left + right * s.k)), Math.abs(r.y + r.h - (s.top + bottom * s.k)));
      multi = +gap.toFixed(1);
      return gap <= 3;
    }, 'the dashed frame round both', 3000).catch(() => false);
    await shoot(win, wait, 'rotate-frame-deck-multiple.png');
    check('rotate-frame-deck: with two turned shapes selected, the dashed frame goes round both turned outlines', two && multi !== false && multi <= 3, JSON.stringify({ two, gap: multi }));

    // Right to left, as the window is in Urdu or Arabic: the frame stays on the box.
    await letGo();
    await js(`document.documentElement.dir = 'rtl', 1`);
    await wait(400);
    await select('Box');
    const rtlFrame = await frameOn('Box');
    await shoot(win, wait, 'rotate-frame-deck-rtl.png');
    const rtlClicks = await clicks('Box');
    await js(`document.documentElement.dir = 'ltr', 1`);
    await wait(300);
    check('rotate-frame-deck: with the window right to left, the turned box’s frame stays on it and the clicks land by its turned outline',
      rtlFrame.ok && rtlClicks.ok, JSON.stringify({ rtlFrame, rtlClicks }));

    // Saved and opened again: the angles in the file, the frame still on the box.
    await letGo();
    await js(`document.querySelector('.sl-stage')?.focus(), 1`);
    await press(wc, 's', { modifiers: ['control'] });
    const inFile = () => { try { return Object.fromEntries(Deck.open(fs.readFileSync(file)).slide(0).shapes.map((s) => [s.name, s.geometry])); } catch { return {}; } };
    const saved = await until(() => { const f = inFile(); return f.Box?.rot === 135 && f.Pic?.rot === 90 && f.Pic?.flipH === true && f[groupName]?.rot === 30; }, 'the angles in the file', 8000).then(() => true).catch(() => false);
    const xml = (() => { try { return Deck.open(fs.readFileSync(file)).pkg.text('ppt/slides/slide1.xml'); } catch { return ''; } })();
    const rotAttrs = (xml.match(/<a:xfrm\b[^>]*\brot="(\d+)"/g) || []).length;
    const sid = session()?.id;
    // A second deck stays open while this one is closed and opened again:
    // the last window closing would end the run.
    const keep = path.join(root, 'keep.pptx');
    fs.writeFileSync(keep, Deck.open(buildPptx({ title: 'Keep', slides: [{ layout: 'blank' }] })).save());
    await open('slides', keep);
    win.close();
    await until(() => win.isDestroyed(), 'the window closed', 4000).catch(() => {});
    if (sid && doc.sessions().some((s) => s.id === sid)) await doc.close({ id: sid }).catch(() => {});
    win = await open('slides', file);
    wc = win.webContents;
    js = (code) => wc.executeJavaScript(code);
    mouse = pointer(win, wait);
    await until(() => js(`Boolean(document.querySelector('.sl-hit[data-shape="${ids.Box}"]'))`), 'the box on the stage again', 10000).catch(() => {});
    await setZoom(1);
    const reopenedRot = geo('Box')?.rot;
    await select('Box');
    const reopenedFrame = await frameOn('Box');
    await shoot(win, wait, 'rotate-frame-deck-reopened.png');
    const reopenedClicks = await clicks('Box');
    check('rotate-frame-deck: saved and opened again, each shape keeps its angle as a:xfrm rot and the frame is still on the turned box',
      saved && rotAttrs >= 3 && reopenedRot === 135 && reopenedFrame.ok && reopenedClicks.ok,
      JSON.stringify({ saved, rotAttrs, file: { Box: inFile().Box?.rot, Pic: inFile().Pic?.rot, Group: inFile()[groupName]?.rot }, reopenedRot, reopenedFrame, reopenedClicks }));

    const complaints = await errorsIn(win);
    check('rotate-frame-deck: the turned-frame checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('rotate-frame-deck: the turned-frame checks ran', false, err.stack || err.message);
  }
}
