/**
 * Zoom and pan on the stage — `apps/studio/lib/zoom.js`, the arithmetic
 * `components/editor/Stage.js` draws with.
 *
 *   node --test tests/zoom.test.mjs
 */
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

import {
    ZOOM_STEPS, MIN_ZOOM, MAX_ZOOM, clampZoom, fitScale, stepZoom, zoomLabel, scrollToKeep, zoomKey,
} from '../packages/studio/src/zoom.js';

test('fit shows the whole frame and never enlarges it', () => {
    assert.equal(fitScale({ width: 540, height: 960 }, { width: 1080, height: 1920 }), 0.5);
    assert.equal(fitScale({ width: 2000, height: 2000 }, { width: 1080, height: 1080 }), 1, 'a roomy stage stays at 100%');
    assert.equal(fitScale({ width: 400, height: 900 }, { width: 800, height: 800 }), 0.5, 'the tighter side decides');
    assert.equal(fitScale({ width: 0, height: 0 }, { width: 800, height: 800 }), 1, 'an unmeasured box is harmless');
});

test('the buttons step through the list, from anywhere, and stop at the ends', () => {
    assert.equal(stepZoom(1, 1), 1.25);
    assert.equal(stepZoom(1, -1), 0.75);
    assert.equal(stepZoom(0.42, 1), 0.5, 'from an in-between fit, the next step up');
    assert.equal(stepZoom(0.42, -1), 0.33, 'and the next step down');
    assert.equal(stepZoom(MAX_ZOOM, 1), MAX_ZOOM);
    assert.equal(stepZoom(MIN_ZOOM, -1), MIN_ZOOM);
    assert.deepEqual([...ZOOM_STEPS].sort((a, b) => a - b), ZOOM_STEPS, 'the list is in order');
    assert.equal(clampZoom(99), MAX_ZOOM);
    assert.equal(clampZoom('x'), 1);
    assert.equal(zoomLabel(0.667), '67%');
});

test('zooming keeps the point under the pointer where it was', () => {
    // A 1000px frame at 50% with 16px round it; the pointer 300px into the
    // visible box, nothing scrolled: content point (300 - 16) / 0.5 = 568.
    const next = scrollToKeep({ scrollLeft: 0, scrollTop: 0 }, { x: 300, y: 116 }, 0.5, 1, 16);
    // At 100% that point is at 16 + 568 = 584 from the scroller's origin; the
    // pointer is 300 in, so the scroller must be 284 along.
    assert.deepEqual(next, { left: 284, top: 100 });
    // Zooming back out about the same pointer returns to where it started.
    assert.deepEqual(scrollToKeep({ scrollLeft: 284, scrollTop: 100 }, { x: 300, y: 116 }, 1, 0.5, 16), { left: 0, top: 0 });
    // Never a negative scroll.
    assert.deepEqual(scrollToKeep({ scrollLeft: 0, scrollTop: 0 }, { x: 100, y: 100 }, 1, 0.5, 16), { left: 0, top: 0 });
});

test('the keys are Ctrl or Cmd with plus, minus and nought', () => {
    assert.equal(zoomKey({ ctrlKey: true, key: '=' }), 'in');
    assert.equal(zoomKey({ metaKey: true, key: '+' }), 'in');
    assert.equal(zoomKey({ ctrlKey: true, key: '-' }), 'out');
    assert.equal(zoomKey({ ctrlKey: true, key: '0' }), 'fit');
    assert.equal(zoomKey({ key: '-' }), null, 'a bare minus is typing');
    assert.equal(zoomKey({ ctrlKey: true, altKey: true, key: '0' }), null);
    assert.equal(zoomKey({ ctrlKey: true, key: 'z' }), null, 'undo is not ours');
});

test('the stage draws at the zoom and offers the controls', () => {
    const stage = readFileSync(new URL('../packages/studio/port/editor/Stage.js', import.meta.url), 'utf8');
    assert.ok(stage.includes('from "../../lib/zoom"'), 'Stage uses the zoom arithmetic');
    assert.ok(/const scale = zoom \?\? fitted;/.test(stage), 'a chosen zoom overrides fit');
    assert.ok(stage.includes('aria-label="Zoom"'), 'and shows the zoom bar');
});
