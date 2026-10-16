// Video: where Add Media puts files, the tracks, and the In, Out and marker
// points dragged (owner, 2026-10-16: "the insert media should have choice to
// add layer or next to existing layer", "ability to split a layer into tow so
// multiple clips can be brought next to each other", "video start and end
// markers").
//
// Two clips and a still, made here at run time. Add Media puts a file after
// the clips on the picked track; Add on a New Track puts the still on a track
// of its own from the playhead; Insert at Playhead splits the clip under the
// playhead and moves what follows on. Split Track at the playhead moves the
// clips after it onto a new track; Move Track Down takes a track below the
// one under it; solo leaves the other picture tracks out of the preview; a
// locked track refuses Delete. The In and Out points are dragged on the ruler
// and on the scrub bar, a marker is dragged, and Undo puts a drag back. Run
// alone with RUTBA_VERIFY_ONLY=video-tracks.

import { makeFixtures } from './verify-video-editor.js';

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args the run's scratch folder
 */
export async function verifyVideoTracks(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const win = await open('video');
  const js = (code) => win.webContents.executeJavaScript(code);
  const fixtures = await makeFixtures(win.webContents, dir, 'video-tracks');
  if (!fixtures) {
    check('video tracks: clips could be made for the checks (VP8 through WebCodecs)', false, 'no encoder in this build');
    return;
  }
  const clickTab = (name) => js(`(() => { const t = [...document.querySelectorAll('.rw-tab')].find((n) => n.textContent.trim() === ${JSON.stringify(name)}); if (!t) return 'no tab'; t.click(); return 'tab'; })()`);
  const press = async (label) => {
    const find = `[...document.querySelectorAll('.rw-ribbon .rw-btn, .rw-dialog .rw-btn')].find((n) => (n.dataset.tip || n.title || n.textContent || '').trim().startsWith(${JSON.stringify(label)}) && !n.disabled)`;
    await until(() => js(`Boolean(${find})`), `the ${label} button`, 4000).catch(() => {});
    return js(`(() => { const b = ${find}; if (!b) return 'no button ' + ${JSON.stringify(label)}; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked'; })()`);
  };
  const addWith = async (label, files) => {
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'open', answer: files });
    return press(label);
  };
  /** The clips by track, each as [kind, start, end], in time order. */
  const tracks = () => js(`(() => {
    const out = {};
    for (const n of document.querySelectorAll('.vd-clip')) {
      const track = n.closest('.vd-row')?.dataset.track || '?';
      (out[track] ||= []).push([n.dataset.kind, Number(n.dataset.start), Number(n.dataset.end), n.dataset.id]);
    }
    for (const k of Object.keys(out)) out[k].sort((a, b) => a[1] - b[1]);
    return out;
  })()`);
  const shape = (rows) => (rows || []).map(([kind, s, e]) => `${kind} ${s}-${e}`).join(', ');
  const seekTo = async (t) => {
    await js(`(() => { const s = document.querySelector('.vd-scrub'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(s, ${JSON.stringify(String(t))}); s.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await until(() => js(`Math.abs(Number(document.querySelector('.vd-scrub').value) - ${t}) < 0.02`), 'the playhead to move', 3000).catch(() => {});
    await wait(250);
  };
  const pick = (id) => js(`(() => {
    const n = document.querySelector('.vd-clip[data-id="${id}"]');
    if (!n) return false;
    const r = n.getBoundingClientRect();
    const o = { bubbles: true, cancelable: true, button: 0, clientX: r.left + Math.min(8, r.width / 2), clientY: r.top + r.height / 2, pointerId: 1 };
    n.dispatchEvent(new PointerEvent('pointerdown', o));
    window.dispatchEvent(new PointerEvent('pointerup', o));
    return true;
  })()`);
  /** A press on `selector` dragged `dx` pixels along and let go. */
  const dragBy = (selector, dx) => js(`(async () => {
    const n = document.querySelector(${JSON.stringify(selector)});
    if (!n) return 'none';
    const r = n.getBoundingClientRect();
    const x0 = r.left + r.width / 2;
    const y = r.top + r.height / 2;
    const at = (x) => ({ bubbles: true, cancelable: true, button: 0, clientX: x, clientY: y, pointerId: 1, altKey: true });
    n.dispatchEvent(new PointerEvent('pointerdown', at(x0)));
    for (let k = 1; k <= 6; k++) { window.dispatchEvent(new PointerEvent('pointermove', at(x0 + (${dx} * k) / 6))); await new Promise((res) => requestAnimationFrame(res)); }
    window.dispatchEvent(new PointerEvent('pointerup', at(x0 + ${dx})));
    return 'dragged';
  })()`);
  const near = (a, b, tol = 0.02) => Math.abs(a - b) <= tol;

  // After the clips on the picked track.
  await addWith('Add Media', [fixtures.a]);
  await until(async () => (await tracks()).V1?.length === 1, 'the first clip', 15000).catch(() => {});
  const first = (await tracks()).V1?.[0];
  await pick(first?.[3]);
  await addWith('Add Media', [fixtures.b]);
  await until(async () => (await tracks()).V1?.length === 2, 'the second clip', 15000).catch(() => {});
  let tr = await tracks();
  check('video tracks: Add Media puts a file after the clips on the picked track',
    tr.V1?.length === 2 && near(tr.V1[1][1], tr.V1[0][2]), `V1: ${shape(tr.V1)}`);

  // On a new track, from the playhead.
  await seekTo(0.5);
  await clickTab('Insert');
  await addWith('Add on a New Track', [fixtures.still]);
  await until(async () => (await tracks()).V2?.length === 1, 'the still on a new track', 15000).catch(() => {});
  tr = await tracks();
  check('video tracks: Add on a New Track puts the still on a track of its own above the clips, from the playhead',
    tr.V2?.length === 1 && tr.V2[0][0] === 'image' && near(tr.V2[0][1], 0.5) && tr.V1.length === 2, `V1: ${shape(tr.V1)}; V2: ${shape(tr.V2)}`);

  // At the playhead on the picked track: the clip under it split, what follows moved on.
  const b = tr.V1[1];
  await pick(b[3]);
  await seekTo(2);
  await addWith('Insert at Playhead', [fixtures.still]);
  await until(async () => (await tracks()).V1?.length === 4, 'the still inserted', 15000).catch(() => {});
  tr = await tracks();
  const v1 = tr.V1 || [];
  check('video tracks: Insert at Playhead splits the clip under the playhead and moves what follows on to make room',
    v1.length === 4 && v1[2][0] === 'image' && near(v1[2][1], 2) && near(v1[1][2], 2) && near(v1[3][1], v1[2][2]) && near(v1[3][2] - v1[3][1], b[2] - 2, 0.03),
    `V1: ${shape(v1)}`);
  const soundInStep = await js(`(() => { const s = [...document.querySelectorAll('.vd-clip[data-kind="sound"]')].map((n) => Number(n.dataset.start)); return s; })()`);
  check('video tracks: the moved clip keeps its sound in step', soundInStep.some((s) => near(s, v1[3]?.[1] ?? -1)), `sound starts ${soundInStep.join(', ')}; the moved clip at ${v1[3]?.[1]}`);

  // Split Track at the playhead: the clips after it move onto a new track.
  await pick(v1[0][3]);
  await seekTo(1);
  await clickTab('Edit');
  await press('Split Track');
  await until(async () => (await tracks()).V3?.length === 1, 'the tracks renumbered', 4000).catch(() => {});
  tr = await tracks();
  check('video tracks: Split Track moves the picked track’s clips from the playhead on onto a new track just above, the clip under it split there',
    tr.V1?.length === 1 && near(tr.V1[0][2], 1) && tr.V2?.length === 4 && near(tr.V2[0][1], 1) && tr.V3?.length === 1 && tr.V3[0][0] === 'image',
    `V1: ${shape(tr.V1)}; V2: ${shape(tr.V2)}; V3: ${shape(tr.V3)}`);

  // Move Track Down: the still's track goes below the one under it.
  await pick(tr.V3[0][3]);
  await press('Move Track Down');
  await until(async () => (await tracks()).V2?.length === 1, 'the track moved', 4000).catch(() => {});
  tr = await tracks();
  check('video tracks: Move Track Down takes the picked track below the one under it, its clip with it',
    tr.V2?.length === 1 && tr.V2[0][0] === 'image' && tr.V3?.length === 4, `V2: ${shape(tr.V2)}; V3: ${shape(tr.V3)}`);

  // Solo V1: the other picture tracks are left out of the preview.
  await seekTo(0.6);
  const canvasCentre = await js(`(() => { const r = document.querySelector('.vd-canvas').getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; })()`);
  const clickStage = () => js(`(() => { const c = document.querySelector('.vd-canvas'); const o = { bubbles: true, cancelable: true, button: 0, clientX: ${canvasCentre.x}, clientY: ${canvasCentre.y}, pointerId: 1 }; c.dispatchEvent(new PointerEvent('pointerdown', o)); window.dispatchEvent(new PointerEvent('pointerup', o)); return document.querySelector('.vd-frame').dataset.picked || ''; })()`);
  await clickStage();
  await wait(300);
  const topBefore = await js(`document.querySelector('.vd-frame').dataset.picked || ''`);
  await js(`(() => { document.querySelector('.vd-row[data-track="V1"] .vd-solo')?.click(); return 1; })()`);
  await wait(300);
  await clickStage();
  await wait(300);
  const topSolo = await js(`document.querySelector('.vd-frame').dataset.picked || ''`);
  const pressed = await js(`document.querySelector('.vd-row[data-track="V1"] .vd-solo')?.getAttribute('aria-pressed')`);
  check('video tracks: soloing V1 leaves the tracks above out of the preview, so its clip is the one picked there',
    topBefore === tr.V2[0][3] && topSolo === tr.V1[0][3] && pressed === 'true', `picked ${topBefore} then ${topSolo}; pressed ${pressed}`);
  await js(`(() => { document.querySelector('.vd-row[data-track="V1"] .vd-solo')?.click(); return 1; })()`);

  // A locked track refuses Delete.
  await js(`(() => { document.querySelector('.vd-row[data-track="V1"] .vd-row-btn[data-tip="Lock this track"]')?.click(); return 1; })()`);
  await pick(tr.V1[0][3]);
  await press('Delete');
  await wait(400);
  const kept = (await tracks()).V1?.length === 1;
  const said = await js(`[...document.querySelectorAll('.rw-toast, [role="status"]')].some((n) => /locked track/.test(n.textContent))`);
  check('video tracks: Delete on a clip of a locked track leaves it, and says why', kept && said, `kept ${kept}; said ${said}`);
  await js(`(() => { document.querySelector('.vd-row[data-track="V1"] .vd-row-btn[data-tip="Unlock this track"]')?.click(); return 1; })()`);

  // In and Out: set, then dragged on the ruler and on the scrub bar; Undo puts a drag back.
  await seekTo(0.5);
  await press('Mark In');
  await seekTo(2.5);
  await press('Mark Out');
  await until(() => js(`Boolean(document.querySelector('.vd-range-edge.e-out'))`), 'the In and Out flags', 3000).catch(() => {});
  const out0 = await js(`Number(document.querySelector('.vd-range-edge.e-out')?.getAttribute('aria-valuenow'))`);
  await dragBy('.vd-range-edge.e-out', 60);
  await wait(300);
  const out1 = await js(`Number(document.querySelector('.vd-range-edge.e-out')?.getAttribute('aria-valuenow'))`);
  await press('Undo');
  await wait(300);
  const out2 = await js(`Number(document.querySelector('.vd-range-edge.e-out')?.getAttribute('aria-valuenow'))`);
  check('video tracks: the Out point drags along the ruler, and Undo puts it back',
    near(out0, 2.5) && out1 > out0 + 0.1 && near(out2, out0), `Out ${out0} s, dragged to ${out1} s, after Undo ${out2} s`);
  const in0 = await js(`Number(document.querySelector('.vd-range-edge.e-in')?.getAttribute('aria-valuenow'))`);
  await dragBy('.vd-scrub-edge.e-in', 40);
  await wait(300);
  const in1 = await js(`Number(document.querySelector('.vd-range-edge.e-in')?.getAttribute('aria-valuenow'))`);
  check('video tracks: the In point drags along the preview’s scrub bar, the ruler’s flag moving with it',
    near(in0, 0.5) && in1 > in0 + 0.05, `In ${in0} s, dragged on the scrub bar to ${in1} s`);

  // A marker dragged along the ruler.
  await clickTab('Insert');
  await seekTo(1.5);
  await press('Add Marker');
  await until(() => js(`Boolean(document.querySelector('.vd-marker'))`), 'the marker', 3000).catch(() => {});
  const m0 = await js(`parseFloat(document.querySelector('.vd-marker')?.style.left)`);
  await dragBy('.vd-marker', 50);
  await wait(300);
  const m1 = await js(`parseFloat(document.querySelector('.vd-marker')?.style.left)`);
  check('video tracks: a marker drags along the ruler', m1 > m0 + 30, `marker at ${m0}px, dragged to ${m1}px`);

  const complaints = await errorsIn(win);
  check('video tracks: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
}
