// Video: what shows on the preview can be picked, moved, resized and turned
// there (owner, 2026-10-16: "visible object sizeable in view with ability to
// rotate adjust size and position", and "transparency settings").
//
// A clip and a still, made here at run time, are added through Add Media. A
// click on the preview picks the clip that fills the frame, with corner
// grips and no turning grip; a corner drag makes it a picture in picture
// anchored at the opposite corner, which then has a turning grip; the grip
// turns it to any angle; a corner drag on the turned clip keeps the opposite
// corner where it stands; a drag inside moves it; the Angle field sets it;
// Opacity shows through; Undo puts it back. The still is made a picture in
// picture from the inspector and filled again; a locked track's clips cannot
// be picked. Run alone with RUTBA_VERIFY_ONLY=video-stage.

import { makeFixtures } from './verify-video-editor.js';

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args the run's scratch folder
 */
export async function verifyVideoStage(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const win = await open('video');
  const js = (code) => win.webContents.executeJavaScript(code);
  const fixtures = await makeFixtures(win.webContents, dir, 'video-stage');
  if (!fixtures) {
    check('video stage: clips could be made for the checks (VP8 through WebCodecs)', false, 'no encoder in this build');
    return;
  }

  const press = async (label) => {
    const find = `[...document.querySelectorAll('.rw-ribbon .rw-btn, .rw-qat .rw-btn, .rw-dialog .rw-btn')].find((n) => (n.dataset.tip || n.title || n.textContent || '').trim().startsWith(${JSON.stringify(label)}) && !n.disabled)`;
    await until(() => js(`Boolean(${find})`), `the ${label} button`, 4000).catch(() => {});
    return js(`(() => { const b = ${find}; if (!b) return 'no button ' + ${JSON.stringify(label)}; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked'; })()`);
  };
  const count = (kind) => js(`document.querySelectorAll('.vd-clip[data-kind="${kind}"]').length`);
  const seekTo = async (t) => {
    await js(`(() => { const s = document.querySelector('.vd-scrub'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(s, ${JSON.stringify(String(t))}); s.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await until(() => js(`Math.abs(Number(document.querySelector('.vd-scrub').value) - ${t}) < 0.02`), 'the playhead to move', 3000).catch(() => {});
    await wait(300);
  };
  /** The picked frame as the page draws it: its box in client pixels, angle, grips and the clip it belongs to. */
  const picked = () => js(`(() => {
    const f = document.querySelector('.vd-frame');
    const o = document.querySelector('.vd-outline');
    const fr = f.getBoundingClientRect();
    const grips = Object.fromEntries([...document.querySelectorAll('.vd-grip')].map((g) => { const r = g.getBoundingClientRect(); return [g.dataset.grip, { x: r.left + r.width / 2, y: r.top + r.height / 2 }]; }));
    if (!o) return { id: f.dataset.picked || '', frame: { x: fr.left, y: fr.top, w: fr.width, h: fr.height }, grips };
    const box = { x: fr.left + parseFloat(o.style.left), y: fr.top + parseFloat(o.style.top), w: parseFloat(o.style.width), h: parseFloat(o.style.height) };
    return { id: f.dataset.picked || '', frame: { x: fr.left, y: fr.top, w: fr.width, h: fr.height }, box, rot: Number(o.dataset.rot || 0), grips, centre: { x: box.x + box.w / 2, y: box.y + box.h / 2 } };
  })()`);
  /** A press at (x0, y0) on the preview, moved in steps to (x1, y1) and let go there. */
  const drag = (x0, y0, x1, y1) => js(`(async () => {
    const c = document.querySelector('.vd-canvas');
    const at = (x, y) => ({ bubbles: true, cancelable: true, button: 0, clientX: x, clientY: y, pointerId: 1 });
    c.dispatchEvent(new PointerEvent('pointerdown', at(${x0}, ${y0})));
    for (let k = 1; k <= 8; k++) {
      window.dispatchEvent(new PointerEvent('pointermove', at(${x0} + (${x1 - x0} * k) / 8, ${y0} + (${y1 - y0} * k) / 8)));
      await new Promise((r) => requestAnimationFrame(r));
    }
    window.dispatchEvent(new PointerEvent('pointerup', at(${x1}, ${y1})));
    return 1;
  })()`);
  const clickAt = (x, y) => drag(x, y, x, y);
  const near = (a, b, tol = 2.5) => Math.abs(a - b) <= tol;
  const centrePixel = () => js(`(() => { const c = document.querySelector('.vd-canvas'); const o = document.querySelector('.vd-outline'); const f = document.querySelector('.vd-frame').getBoundingClientRect(); const k = c.width / f.width; const x = (parseFloat(o.style.left) + parseFloat(o.style.width) / 2) * k; const y = (parseFloat(o.style.top) + parseFloat(o.style.height) / 2) * k; const d = c.getContext('2d').getImageData(Math.floor(x), Math.floor(y), 1, 1).data; return d[0] + d[1] + d[2]; })()`);

  globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'open', answer: [fixtures.a, fixtures.still] });
  await press('Add Media');
  const landed = await until(async () => (await count('video')) === 1 && (await count('image')) === 1, 'the clip and the still', 20000).then(() => true).catch(() => false);
  const clipId = await js(`document.querySelector('.vd-clip[data-kind="video"]')?.dataset.id || ''`);
  const stillId = await js(`document.querySelector('.vd-clip[data-kind="image"]')?.dataset.id || ''`);
  const stillStart = await js(`Number(document.querySelector('.vd-clip[data-kind="image"]')?.dataset.start || 0)`);
  check('video stage: Add Media puts a clip and a still on the timeline', landed && Boolean(clipId) && Boolean(stillId), `clip ${clipId}, still ${stillId} from ${stillStart} s`);

  // A clip that fills the frame is picked by a click on the preview: corner grips, no turning grip.
  await seekTo(0.5);
  let f = await picked();
  await clickAt(f.frame.x + f.frame.w / 2, f.frame.y + f.frame.h / 2);
  await until(async () => (await picked()).id === clipId, 'the clip picked', 3000).catch(() => {});
  f = await picked();
  const timelinePicked = await js(`document.querySelector('.vd-clip[data-id="${clipId}"]')?.classList.contains('selected') || false`);
  check('video stage: a click on the preview picks the clip that fills the frame, on the timeline too, with a grip at each corner and none to turn it',
    f.id === clipId && timelinePicked && ['ne', 'nw', 'se', 'sw'].every((k) => f.grips[k]) && !f.grips.rotate,
    `picked ${f.id || 'nothing'}; grips ${Object.keys(f.grips).join(',')}`);

  // The corner opposite stays: a drag of the bottom right corner makes it a picture in picture from the top left.
  const se = f.grips.se;
  await drag(se.x, se.y, se.x - f.frame.w * 0.45, se.y - f.frame.h * 0.45);
  await wait(300);
  const inset = await picked();
  const turnable = Boolean(inset.grips.rotate);
  const fields = await js(`Boolean(document.querySelector('.vd-place-props .vd-fx'))`);
  check('video stage: dragging a corner of a clip that fills the frame makes it a picture in picture, the opposite corner held, which can then be turned',
    inset.box && near(inset.box.x, f.frame.x) && near(inset.box.y, f.frame.y) && near(inset.box.w, f.frame.w * 0.55, 6) && turnable && fields,
    inset.box ? `box ${Math.round(inset.box.x - f.frame.x)},${Math.round(inset.box.y - f.frame.y)} ${Math.round(inset.box.w)}x${Math.round(inset.box.h)} of ${Math.round(f.frame.w)}; turning grip ${turnable}; fields ${fields}` : 'no frame');

  // The round grip turns it to any angle: held to the right of the centre, a quarter turn.
  const r0 = inset.grips.rotate;
  const c0 = inset.centre;
  if (r0) await drag(r0.x, r0.y, c0.x + 120, c0.y);
  await wait(300);
  const turned = await picked();
  const angleField = await js(`Number(document.querySelector('.vd-rot')?.value)`);
  check('video stage: the round grip turns the clip to the angle it is held at, the frame turned with it and the Angle field saying so',
    near(turned.rot, 90, 1) && near(angleField, 90, 1) && near(turned.centre.x, c0.x) && near(turned.centre.y, c0.y),
    `${turned.rot} degrees, field ${angleField}; centre moved ${Math.round(turned.centre.x - c0.x)},${Math.round(turned.centre.y - c0.y)}`);

  // A corner drag on the turned clip: the opposite corner stays where it stands turned.
  const before = turned.grips;
  if (before.se) await drag(before.se.x, before.se.y, before.se.x - 30, before.se.y + 30);
  await wait(300);
  const grown = await picked();
  check('video stage: a corner drag on a turned clip resizes it with the opposite corner staying where it stands',
    grown.grips.nw && before.nw && near(grown.grips.nw.x, before.nw.x, 3) && near(grown.grips.nw.y, before.nw.y, 3) && Math.abs(grown.box.w - turned.box.w) > 10,
    grown.grips.nw ? `nw moved ${Math.round(grown.grips.nw.x - before.nw.x)},${Math.round(grown.grips.nw.y - before.nw.y)}; width ${Math.round(turned.box.w)} to ${Math.round(grown.box.w)}` : 'no grips');

  // A drag inside moves it.
  await drag(grown.centre.x, grown.centre.y, grown.centre.x - 40, grown.centre.y + 20);
  await wait(300);
  const moved = await picked();
  check('video stage: a drag inside the clip moves it by as much as the pointer moved',
    near(moved.centre.x - grown.centre.x, -40, 3) && near(moved.centre.y - grown.centre.y, 20, 3) && moved.id === clipId,
    `moved ${Math.round(moved.centre.x - grown.centre.x)},${Math.round(moved.centre.y - grown.centre.y)}`);

  // The Angle field sets the angle.
  await js(`(() => { const n = document.querySelector('.vd-rot'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, '30'); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
  const typed = await until(async () => near((await picked()).rot, 30, 0.5), 'the typed angle', 3000).then(() => true).catch(() => false);
  check('video stage: a number in the Angle field turns the clip to it', typed, `${(await picked()).rot} degrees`);

  // Opacity shows the black under it through the clip, and Undo puts it back.
  const solid = await centrePixel();
  await js(`(() => { const n = document.querySelector('.vd-place-props input[type="range"]'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, '0.3'); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
  await wait(400);
  const faint = await centrePixel();
  await press('Undo');
  await wait(400);
  const back = await centrePixel();
  check('video stage: Opacity lets the frame under the clip show through, and Undo puts it back',
    faint < solid * 0.6 && near(back, solid, 12),
    `centre ${solid}, at 30% ${faint}, after Undo ${back}`);

  // A still: made a picture in picture from the inspector, then filling the frame again.
  await seekTo(stillStart + 0.4);
  f = await picked();
  await clickAt(f.frame.x + f.frame.w / 2, f.frame.y + f.frame.h / 2);
  await until(async () => (await picked()).id === stillId, 'the still picked', 3000).catch(() => {});
  const fillsStill = await picked();
  await js(`(() => { document.querySelector('.vd-to-inset')?.click(); return 1; })()`);
  await wait(400);
  const boxedStill = await picked();
  const width = await js(`Number(document.querySelector('.vd-fw')?.value)`);
  await js(`(() => { document.querySelector('.vd-still-full')?.click(); return 1; })()`);
  await wait(400);
  const fullAgain = await picked();
  check('video stage: a still that fills the frame is made a picture in picture from the inspector, which can be turned, and fills the frame again',
    fillsStill.id === stillId && !fillsStill.grips.rotate && boxedStill.grips.rotate && near(width, 40, 0.5) && !fullAgain.grips.rotate && fullAgain.box && near(fullAgain.box.w, f.frame.w, 3),
    `picked ${fillsStill.id}; turning grip ${Boolean(fillsStill.grips.rotate)} then ${Boolean(boxedStill.grips.rotate)} then ${Boolean(fullAgain.grips.rotate)}; width ${width}%`);

  // A locked track's clips are not picked on the preview.
  await js(`(() => { document.querySelector('.vd-row[data-track="V1"] .vd-row-btn[data-tip="Lock this track"]')?.click(); return 1; })()`);
  await wait(300);
  await clickAt(f.frame.x + 6, f.frame.y + 6);
  await clickAt(f.frame.x + f.frame.w / 2, f.frame.y + f.frame.h / 2);
  await wait(300);
  const lockedPick = await picked();
  check('video stage: a clip on a locked track is not picked on the preview', lockedPick.id === '', `picked ${lockedPick.id || 'nothing'}`);

  const complaints = await errorsIn(win);
  check('video stage: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
}
