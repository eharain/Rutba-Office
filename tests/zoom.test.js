// Zoom by Ctrl and the wheel, and by Ctrl+Plus, Ctrl+Minus and Ctrl+0: the
// arithmetic every app's page, sheet, slide and message zooms with. A
// notch or a key press is 10% on Word's grid, a touchpad pinch scales
// smoothly without losing a slow pinch to rounding, every level stays in
// the app's limits, and the point under the pointer is put back where it
// was. The window checks (RUTBA_VERIFY_ONLY=zoom) prove it on screen.
import test from 'node:test';
import assert from 'node:assert/strict';
import { stepZoom, wheelZoom, anchorAt, anchorShift, scrollRatio } from '@rutba/office-ui/zoom';
import { keyMatches } from '@rutba/office-ui/commands';

test('a step is 10% and lands on the 10% grid, as Word does', () => {
  assert.equal(stepZoom(1, 1), 1.1);
  assert.equal(stepZoom(1, -1), 0.9);
  assert.equal(stepZoom(1.37, 1), 1.4);
  assert.equal(stepZoom(1.37, -1), 1.3);
  assert.equal(stepZoom(0.457, 1), 0.5);
  assert.equal(stepZoom(0.457, -1), 0.4);
  // 1.1 is not taken for 1.0999...: floating point does not skip a step.
  assert.equal(stepZoom(stepZoom(1, 1), 1), 1.2);
  assert.equal(stepZoom(stepZoom(1, -1), -1), 0.8);
});

test('a step stays in the app\'s own limits', () => {
  assert.equal(stepZoom(2.95, 1, { min: 0.3, max: 3 }), 3);
  assert.equal(stepZoom(3, 1, { min: 0.3, max: 3 }), 3);
  assert.equal(stepZoom(0.3, -1, { min: 0.3, max: 3 }), 0.3);
  assert.equal(stepZoom(0.5, -1, { min: 0.5, max: 3 }), 0.5);
});

test('a mouse wheel notch is one step; a notch towards the person zooms in', () => {
  assert.equal(wheelZoom(1, null, { deltaY: -100 }).level, 1.1);
  assert.equal(wheelZoom(1, null, { deltaY: 100 }).level, 0.9);
  // Two notches in one event (a fast flick), and a wheel that reports 120.
  assert.equal(wheelZoom(1, null, { deltaY: -200 }).level, 1.2);
  assert.equal(wheelZoom(1, null, { deltaY: -120 }).level, 1.1);
  // Lines and pages, as some wheels report them.
  assert.equal(wheelZoom(1, null, { deltaY: -3, deltaMode: 1 }).level, 1.1);
  assert.equal(wheelZoom(1, null, { deltaY: 1, deltaMode: 2 }).level, 0.9);
  assert.equal(wheelZoom(2.95, null, { deltaY: -100 }, { min: 0.3, max: 3 }).level, 3);
  assert.equal(wheelZoom(1, null, { deltaY: 0 }).level, 1);
});

test('a pinch scales smoothly, and a slow pinch adds up rather than rounding away', () => {
  // A pinch of 1.5 arrives as -100 ln 1.5.
  const once = wheelZoom(1, null, { deltaY: -100 * Math.log(1.5) });
  assert.equal(once.level, 1.5);
  // A hundred tiny events, each well under a percent.
  let level = 1;
  let exact = null;
  for (let i = 0; i < 100; i++) ({ level, exact } = wheelZoom(level, exact, { deltaY: -0.2953 }));
  assert.ok(Math.abs(level - 1.34) <= 0.01, `a slow pinch reached ${level}`);
  // Out again, to within a percent of where it began.
  for (let i = 0; i < 100; i++) ({ level, exact } = wheelZoom(level, exact, { deltaY: 0.2953 }));
  assert.ok(Math.abs(level - 1) <= 0.01, `back to ${level}`);
  // In its limits.
  assert.equal(wheelZoom(2.9, null, { deltaY: -50.5 }, { min: 0.3, max: 3 }).level, 3);
});

test('a pinch starts from the level on screen once something else has moved it', () => {
  // The carried level was 1.234 but the slider has since set 2.
  const next = wheelZoom(2, 1.234, { deltaY: -100 * Math.log(1.1) });
  assert.equal(next.level, 2.2);
});

test('the point under the pointer is put back under it', () => {
  // A page 800 wide at 100 from the area's left, the pointer at 300 across.
  const before = { left: 100, top: -300, width: 800, height: 1100 };
  const anchor = anchorAt(before, 300, 200);
  assert.deepEqual([anchor.fx, anchor.fy], [0.25, 500 / 1100]);
  // At 150% the page, still centred, is 1200 wide: the point is now 100
  // to the left of the pointer and 250 below it.
  const after = { left: -100, top: -300, width: 1200, height: 1650 };
  const shift = anchorShift(anchor, after);
  assert.ok(Math.abs(shift.dx + 100) < 1e-9 && Math.abs(shift.dy - 250) < 1e-9, JSON.stringify(shift));
  // Scrolled by that much, the point is back.
  const scrolled = { ...after, left: after.left - shift.dx, top: after.top - shift.dy };
  const back = anchorShift(anchor, scrolled);
  assert.ok(Math.abs(back.dx) < 1e-9 && Math.abs(back.dy) < 1e-9);
  // Nothing to keep, nothing to scroll.
  assert.deepEqual(anchorShift(null, after), { dx: 0, dy: 0 });
  assert.deepEqual(anchorAt({ left: 0, top: 0, width: 0, height: 0 }, 5, 5).fx, 0);
});

test('an area that scrolls in its own zoomed pixels is found out from one scroll', () => {
  // The sheet's grid at 150%: 100 scrolled pixels move the point 150 on screen.
  assert.equal(scrollRatio(300, 150, 100), 1.5);
  // An area at the end of its scroll did not move: nothing can be told.
  assert.equal(scrollRatio(300, 300, 0), 0);
  // A move the wrong way is not believed.
  assert.equal(scrollRatio(300, 450, 100), 0);
});

test('Ctrl+Plus is Ctrl and + or =, with Shift or without; Ctrl+Minus and Ctrl+0 are as pressed', () => {
  const key = (k, extra = {}) => ({ key: k, ctrlKey: true, metaKey: false, shiftKey: false, altKey: false, ...extra });
  const mod = /Mac|iPhone|iPad/.test(globalThis.navigator?.platform || '') ? { metaKey: true, ctrlKey: false } : {};
  assert.equal(keyMatches('Mod+Plus', key('=', mod)), true);
  assert.equal(keyMatches('Mod+Plus', key('+', { shiftKey: true, ...mod })), true);
  assert.equal(keyMatches('Mod+Plus', key('+', mod)), true);
  assert.equal(keyMatches('Mod+Plus', key('=', { altKey: true, ...mod })), false);
  assert.equal(keyMatches('Mod+Plus', { key: '=', ctrlKey: false, metaKey: false, shiftKey: false, altKey: false }), false);
  assert.equal(keyMatches('Mod+-', key('-', mod)), true);
  assert.equal(keyMatches('Mod+0', key('0', mod)), true);
  assert.equal(keyMatches('Mod+0', key('0', { shiftKey: true, ...mod })), false);
});
