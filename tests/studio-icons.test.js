/**
 * The icon set — that every glyph is well formed and inside its box.
 *
 * WHY THIS FILE EXISTS. Path data is unreviewable by eye: a glyph two units off
 * centre, or one with a stray comma, looks exactly like a correct one in the
 * source and shows up as a mark sitting wrong in somebody's finished video.
 * These are the checks a reader cannot make.
 *
 * What they do NOT check is whether an icon looks like the thing it is named
 * after — no test can, and the composition helpers (`circle`, `rect`, `poly`,
 * `bar`, `spoke`) are the real defence there: an icon assembled from named
 * primitives is reviewable in a way a `d` string never is.
 *
 *   node --test tests/icons.test.mjs
 */
import { test } from 'node:test';
import assert from 'node:assert/strict';

import { ICONS, ICON_NAMES, ICON_GROUPS, ICON_LABELS } from '../packages/studio/src/icons.js';
import { registeredLayerTypes, applyLayerPatches } from '../packages/studio/src/index.js';

/** Every number in a path, wherever it appears. */
const numbersIn = (d) => (d.match(/-?\d*\.?\d+/g) || []).map(Number);

/**
 * The real bounding box of a path.
 *
 * A FIRST ATTEMPT AT THIS just looked at the biggest number in the string, and
 * it was wrong in the way that matters: a circle is `M2.6 12 a9.4 9.4 …`, whose
 * only absolute coordinates are two points on the left and right of the ring —
 * so `ring` measured as 12 units wide and was reported as never reaching the
 * right half of its box. The extent of a round glyph lives in its radii, not in
 * its movetos, and nothing short of walking the path finds it.
 *
 * Curves are bounded by their control hull, which is conservative — the true
 * curve is inside it — so a glyph that passes is definitely inside the box.
 * Arcs are converted to centre form and measured properly, including the axis
 * extremes the sweep actually crosses; every arc in this set has no x-rotation,
 * which is the one case this skips.
 */
function bbox(d) {
    let minX = Infinity; let minY = Infinity; let maxX = -Infinity; let maxY = -Infinity;
    const see = (x, y) => {
        if (!Number.isFinite(x) || !Number.isFinite(y)) return;
        minX = Math.min(minX, x); maxX = Math.max(maxX, x);
        minY = Math.min(minY, y); maxY = Math.max(maxY, y);
    };

    let x = 0; let y = 0; let sx = 0; let sy = 0;
    for (const [, cmd, rest] of d.matchAll(/([MmLlHhVvCcSsQqTtAaZz])([^MmLlHhVvCcSsQqTtAaZz]*)/g)) {
        const a = (rest.match(/-?\d*\.?\d+(?:[eE][+-]?\d+)?/g) || []).map(Number);
        const rel = cmd === cmd.toLowerCase();
        const up = cmd.toUpperCase();
        let i = 0;
        const take = (k) => a.slice(i, i + k);

        if (up === 'Z') { x = sx; y = sy; see(x, y); continue; }

        while (i < a.length || (i === 0 && a.length === 0)) {
            if (up === 'M' || up === 'L' || up === 'T') {
                const [px, py] = take(2); i += 2;
                x = rel ? x + px : px; y = rel ? y + py : py;
                if (up === 'M' && i === 2) { sx = x; sy = y; }
                see(x, y);
            } else if (up === 'H') {
                const [px] = take(1); i += 1;
                x = rel ? x + px : px; see(x, y);
            } else if (up === 'V') {
                const [py] = take(1); i += 1;
                y = rel ? y + py : py; see(x, y);
            } else if (up === 'C') {
                const [x1, y1, x2, y2, px, py] = take(6); i += 6;
                see(rel ? x + x1 : x1, rel ? y + y1 : y1);
                see(rel ? x + x2 : x2, rel ? y + y2 : y2);
                x = rel ? x + px : px; y = rel ? y + py : py; see(x, y);
            } else if (up === 'S' || up === 'Q') {
                const [x1, y1, px, py] = take(4); i += 4;
                see(rel ? x + x1 : x1, rel ? y + y1 : y1);
                x = rel ? x + px : px; y = rel ? y + py : py; see(x, y);
            } else if (up === 'A') {
                const [rx, ry, , large, sweep, px, py] = take(7); i += 7;
                const nx = rel ? x + px : px;
                const ny = rel ? y + py : py;
                see(nx, ny);
                arcExtremes(x, y, nx, ny, Math.abs(rx), Math.abs(ry), large, sweep, see);
                x = nx; y = ny;
            } else { break; }
            if (a.length === 0) break;
        }
    }
    return { minX, minY, maxX, maxY, w: maxX - minX, h: maxY - minY };
}

/** The axis-extreme points an elliptical arc actually passes through. */
function arcExtremes(x1, y1, x2, y2, rx, ry, large, sweep, see) {
    if (!rx || !ry) return;
    // Centre parameterisation, with no x-axis rotation (none in this set).
    const dx = (x1 - x2) / 2;
    const dy = (y1 - y2) / 2;
    let lam = (dx * dx) / (rx * rx) + (dy * dy) / (ry * ry);
    let RX = rx; let RY = ry;
    if (lam > 1) { const s = Math.sqrt(lam); RX *= s; RY *= s; lam = 1; }
    const num = Math.max(0, RX * RX * RY * RY - RX * RX * dy * dy - RY * RY * dx * dx);
    const den = RX * RX * dy * dy + RY * RY * dx * dx;
    const co = (den === 0 ? 0 : Math.sqrt(num / den)) * (large === sweep ? -1 : 1);
    const cx = co * ((RX * dy) / RY) + (x1 + x2) / 2;
    const cy = co * (-(RY * dx) / RX) + (y1 + y2) / 2;

    const ang = (px, py) => Math.atan2((py - cy) / RY, (px - cx) / RX);
    let a0 = ang(x1, y1);
    let a1 = ang(x2, y2);
    if (sweep) { if (a1 < a0) a1 += Math.PI * 2; } else if (a1 > a0) { a0 += Math.PI * 2; }
    const [lo, hi] = a0 <= a1 ? [a0, a1] : [a1, a0];
    for (let k = -2; k <= 4; k++) {
        for (const base of [0, Math.PI / 2, Math.PI, -Math.PI / 2]) {
            const t = base + k * Math.PI * 2;
            if (t >= lo && t <= hi) see(cx + RX * Math.cos(t), cy + RY * Math.sin(t));
        }
    }
}

// ── the data ────────────────────────────────────────────────────────────────

test('the set is not empty and every name has a path', () => {
    assert.ok(ICON_NAMES.length >= 40, `only ${ICON_NAMES.length} icons`);
    for (const name of ICON_NAMES) {
        assert.equal(typeof ICONS[name], 'string', `${name} has no path`);
        assert.ok(ICONS[name].length > 8, `${name} is too short to be a glyph`);
    }
});

test('every path is syntactically well formed', () => {
    // Commands, numbers, and separators — nothing else. This is what catches a
    // typo'd letter or a stray character in a composed string.
    const GRAMMAR = /^[MmLlHhVvCcSsQqTtAaZz0-9eE .,+-]+$/;
    for (const name of ICON_NAMES) {
        const d = ICONS[name];
        assert.match(d, GRAMMAR, `${name} has a character no path command uses`);
        assert.ok(/^[Mm]/.test(d.trim()), `${name} does not start with a moveto`);
        assert.ok(/[Zz]\s*$/.test(d.trim()), `${name} does not close its last subpath`);
        assert.ok(!/NaN|undefined|Infinity/.test(d), `${name} has a bad number in it`);
    }
});

test('no path carries a NaN, which a helper would produce silently', () => {
    for (const name of ICON_NAMES) {
        for (const v of numbersIn(ICONS[name])) {
            assert.ok(Number.isFinite(v), `${name} contains a non-finite coordinate`);
        }
    }
});

test('every glyph stays inside its 24x24 box', () => {
    // A hair of slack: curve bounds come from the control hull, which sits
    // slightly outside the curve it describes.
    for (const name of ICON_NAMES) {
        const b = bbox(ICONS[name]);
        assert.ok(b.minX >= -1.2 && b.minY >= -1.2, `${name} starts at ${b.minX},${b.minY}`);
        assert.ok(b.maxX <= 25.2 && b.maxY <= 25.2, `${name} reaches ${b.maxX},${b.maxY}`);
    }
});

test('a glyph fills a reasonable amount of the box', () => {
    // A glyph drawn at half scale reads as a rendering bug next to its
    // neighbours — every icon in a set has to carry the same weight.
    //
    // The rule is about the LONGEST axis, not both. A first version demanded
    // eight units in each direction and `minus` failed it, correctly: a
    // horizontal bar is supposed to be thin, and a test that calls that a
    // defect is a test that would push someone to fatten a minus sign.
    for (const name of ICON_NAMES) {
        const b = bbox(ICONS[name]);
        assert.ok(Math.max(b.w, b.h) >= 15,
            `${name} spans only ${Math.max(b.w, b.h).toFixed(1)} units at its longest`);
        assert.ok(Math.min(b.w, b.h) >= 2.5, `${name} is degenerate in one axis`);
    }
});

test('a glyph is roughly centred in its box', () => {
    // Icons are placed by their CENTRE, so one drawn off-centre lands off-centre
    // wherever it is used — and the error is invisible in the source.
    for (const name of ICON_NAMES) {
        const b = bbox(ICONS[name]);
        const cx = (b.minX + b.maxX) / 2;
        const cy = (b.minY + b.maxY) / 2;
        assert.ok(Math.abs(cx - 12) <= 2.5, `${name} centres at x=${cx.toFixed(1)}, not 12`);
        assert.ok(Math.abs(cy - 12) <= 2.5, `${name} centres at y=${cy.toFixed(1)}, not 12`);
    }
});

// ── the palette's view of it ────────────────────────────────────────────────

test('every icon is in exactly one group, and every group names a real icon', () => {
    // The silent failure this guards: an icon that exists and is in no group is
    // an icon nobody can ever reach.
    const seen = new Set();
    for (const g of ICON_GROUPS) {
        for (const name of g.names) {
            assert.ok(ICONS[name], `group "${g.label}" names ${name}, which does not exist`);
            assert.ok(!seen.has(name), `${name} appears in two groups`);
            seen.add(name);
        }
    }
    for (const name of ICON_NAMES) {
        assert.ok(seen.has(name), `${name} exists but is in no group`);
    }
});

test('every icon has a human label', () => {
    for (const name of ICON_NAMES) {
        assert.ok(ICON_LABELS[name], `${name} has no label`);
    }
});

// ── the layer type ──────────────────────────────────────────────────────────

test('the icon type registers itself on import', () => {
    assert.ok(registeredLayerTypes().includes('icon'));
});

const stubPlan = () => ({
    W: 1080, H: 1920, duration: 5, fps: 30,
    theme: { bg: '#000', scrim: 'rgba(0,0,0,.7)', text: '#fff', dim: '#aaa', accent: '#ffc107' },
    layers: [], images: [], videos: {}, assets: {}, context: {},
});

const compile = (patches) => {
    const plan = stubPlan();
    applyLayerPatches(plan, patches);
    return plan;
};

test('an icon compiles square in pixels, so it is round on a tall frame', () => {
    const plan = compile([{ id: 'i', type: 'icon', name: 'heart', fw: 0.2 }]);
    const layer = plan.layers[0];
    // fw is a fraction of W and fh of H, so equal fractions would be an oval.
    assert.ok(Math.abs(layer.fw * plan.W - layer.fh * plan.H) < 1, 'not square in pixels');
});

test('an unknown icon name compiles to a known one rather than to nothing', () => {
    // Same forward-compatibility posture as compileShape's rect fallback: a
    // recipe from a newer editor degrades rather than failing to compile.
    const plan = compile([{ id: 'i', type: 'icon', name: 'no-such-glyph' }]);
    assert.ok(ICONS[plan.layers[0].name], 'fell back to a real glyph');
});

test('an icon layer carries the common controls every layer type has', () => {
    const plan = compile([{
        id: 'i', type: 'icon', name: 'star',
        opacity: 0.5, rot: 30, blend: 'screen', mask: { kind: 'ellipse' },
    }]);
    const l = plan.layers[0];
    assert.equal(l.opacity, 0.5);
    assert.equal(l.rot, 30);
    assert.equal(l.blend, 'screen');
    assert.ok(l.mask, 'masking works on an icon like anything else');
});

test('fill: null survives, so an outline-only icon is possible', () => {
    // `||` would swallow it — the same trap compileShape documents.
    const plan = compile([{ id: 'i', type: 'icon', name: 'star', fill: null, stroke: { width: 0.004 } }]);
    assert.equal(plan.layers[0].fill, null);
});
