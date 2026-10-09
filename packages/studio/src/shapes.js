/**
 * Vector shape layers — the second half of "make it by hand".
 *
 * Word art gives a headline; this gives everything a headline sits on: badges,
 * price bursts, banners, speech bubbles, rules, arrows. Together they are what
 * replaces the product data-binding Studio deliberately does not have. A media
 * tool that cannot draw a rectangle sends people back to their old editor for
 * one shape and loses the whole creative with it.
 *
 * WHY NOT @rutba/drawing. The consumer shelf carries a scene graph with exactly
 * these geometries (`@rutba/drawing`, whose own description names Studio as an
 * intended consumer), and reusing it was the first instinct. It renders SVG.
 * This renderer is canvas 2D, in a package whose stated contract is "browser
 * engine only, zero dependencies" — the property that lets the render worker
 * run it unchanged. Importing an SVG scene graph to then reimplement its output
 * as canvas paths buys a shared vocabulary and pays a dependency plus a
 * conversion for it. The GEOMETRY is the cheap part and lives here; when the
 * slide designer needs real SVG export, @rutba/drawing is what it exports TO.
 *
 * Everything is fractional (fx/fy/fw/fh of the frame) so one shape is correct
 * at 9:16 and at 1:1 — the same contract every other layer keeps.
 */

import { registerLayerType } from './index.js';

/**
 * Geometries a shape layer can take.
 *
 * ORDER IS THE INSERT PALETTE'S ORDER, grouped by what someone is reaching for
 * rather than alphabetically — boxes, then rounds, then pointers, then the
 * decorative ones. `SHAPE_GROUPS` below carries the headings; this stays the
 * flat list because `compileShape` validates against it and a recipe naming an
 * unknown geometry must degrade to a rectangle rather than to nothing.
 */
export const GEOMETRIES = [
    // boxes
    'rect', 'pill', 'parallelogram', 'trapezoid', 'diamond', 'frame',
    // rounds
    'ellipse', 'semicircle', 'ring', 'pie', 'teardrop', 'blob', 'cloud',
    // polygons and points
    'triangle', 'polygon', 'star', 'burst', 'gear', 'plus', 'bolt',
    // pointers and bands
    'arrow', 'doubleArrow', 'line', 'wave', 'banner', 'chevron',
    // labels
    'bubble', 'tag', 'bookmark', 'shield',
];

/**
 * The palette's headings. Grouping is presentation, so it lives beside the
 * geometries rather than in the app — an editor should not have to know which
 * of thirty names is a "round" to lay out a panel, and a geometry added here
 * appears in the palette without the app being touched at all.
 */
export const SHAPE_GROUPS = [
    { label: 'Boxes', geometries: ['rect', 'pill', 'parallelogram', 'trapezoid', 'diamond', 'frame'] },
    { label: 'Rounds', geometries: ['ellipse', 'semicircle', 'ring', 'pie', 'teardrop', 'blob', 'cloud'] },
    { label: 'Points', geometries: ['triangle', 'polygon', 'star', 'burst', 'gear', 'plus', 'bolt'] },
    { label: 'Pointers', geometries: ['arrow', 'doubleArrow', 'line', 'wave', 'banner', 'chevron'] },
    { label: 'Labels', geometries: ['bubble', 'tag', 'bookmark', 'shield'] },
];

/** What each geometry is called, and what to search for it by. */
export const GEOMETRY_LABELS = {
    rect: 'Rectangle', pill: 'Pill', parallelogram: 'Parallelogram', trapezoid: 'Trapezoid',
    diamond: 'Diamond', frame: 'Frame', ellipse: 'Ellipse', semicircle: 'Semicircle',
    ring: 'Ring', pie: 'Pie slice', teardrop: 'Teardrop', blob: 'Blob', cloud: 'Cloud',
    triangle: 'Triangle', polygon: 'Polygon', star: 'Star', burst: 'Burst', gear: 'Gear',
    plus: 'Cross', bolt: 'Lightning', arrow: 'Arrow', doubleArrow: 'Double arrow',
    line: 'Rule', wave: 'Wave', banner: 'Banner', chevron: 'Chevron', bubble: 'Speech bubble',
    tag: 'Price tag', bookmark: 'Bookmark', shield: 'Shield',
};

const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));
const num = (v, fallback) => (Number.isFinite(Number(v)) ? Number(v) : fallback);
const TAU = Math.PI * 2;

// ── geometry ─────────────────────────────────────────────────────────────────

/**
 * Build the shape's path into `ctx`, in a box centred on the origin.
 *
 * Every geometry is written against a `[-w/2, -h/2] .. [w/2, h/2]` box, so
 * rotation (which the frame wrapper applies about the layer's centre) needs no
 * per-shape correction, and a shape can be swapped for another without moving.
 */
function buildPath(ctx, layer, w, h) {
    const g = layer.geometry || 'rect';
    const hw = w / 2;
    const hh = h / 2;
    const points = Math.max(3, Math.round(num(layer.points, 5)));

    ctx.beginPath();

    if (g === 'ellipse') {
        ctx.ellipse(0, 0, hw, hh, 0, 0, TAU);
        return;
    }

    if (g === 'triangle') {
        ctx.moveTo(0, -hh);
        ctx.lineTo(hw, hh);
        ctx.lineTo(-hw, hh);
        ctx.closePath();
        return;
    }

    if (g === 'polygon') {
        // Flat-topped is what a badge wants; -PI/2 puts a VERTEX at the top,
        // which is what a pentagon wants. Vertex-up reads as deliberate.
        for (let i = 0; i < points; i++) {
            const a = -Math.PI / 2 + (i / points) * TAU;
            const x = Math.cos(a) * hw;
            const y = Math.sin(a) * hh;
            if (i === 0) ctx.moveTo(x, y); else ctx.lineTo(x, y);
        }
        ctx.closePath();
        return;
    }

    if (g === 'star' || g === 'burst') {
        // A burst is a star with many shallow points — the price-splat shape.
        const spikes = g === 'burst' ? Math.max(8, points * 2) : points;
        const inner = g === 'burst'
            ? clamp(num(layer.innerRatio, 0.82), 0.05, 0.98)
            : clamp(num(layer.innerRatio, 0.45), 0.05, 0.98);
        for (let i = 0; i < spikes * 2; i++) {
            const a = -Math.PI / 2 + (i / (spikes * 2)) * TAU;
            const r = i % 2 === 0 ? 1 : inner;
            const x = Math.cos(a) * hw * r;
            const y = Math.sin(a) * hh * r;
            if (i === 0) ctx.moveTo(x, y); else ctx.lineTo(x, y);
        }
        ctx.closePath();
        return;
    }

    if (g === 'arrow') {
        // Shaft thickness and head length as fractions of the box, so an arrow
        // stays an arrow when it is stretched long and thin.
        const shaft = clamp(num(layer.thickness, 0.42), 0.05, 0.95) * hh;
        const head = clamp(num(layer.headLength, 0.36), 0.05, 0.9) * w;
        const xTip = hw;
        const xNeck = hw - head;
        ctx.moveTo(-hw, -shaft);
        ctx.lineTo(xNeck, -shaft);
        ctx.lineTo(xNeck, -hh);
        ctx.lineTo(xTip, 0);
        ctx.lineTo(xNeck, hh);
        ctx.lineTo(xNeck, shaft);
        ctx.lineTo(-hw, shaft);
        ctx.closePath();
        return;
    }

    if (g === 'bubble') {
        // Rounded body plus a tail. The tail hangs off the bottom edge at
        // `tailAt` (0..1 across the width) so it can point at what it is about.
        const r = clamp(num(layer.radius, 0.12), 0, 0.5) * Math.min(w, h) * 2;
        const tailW = clamp(num(layer.tailWidth, 0.16), 0.02, 0.6) * w;
        const tailH = clamp(num(layer.tailHeight, 0.22), 0.02, 0.9) * h;
        const bodyH = h - tailH;
        const bt = -hh;
        const bb = bt + bodyH;
        const tailX = -hw + clamp(num(layer.tailAt, 0.28), 0, 1) * w;
        const rr = Math.min(r, bodyH / 2, w / 2);

        ctx.moveTo(-hw + rr, bt);
        ctx.lineTo(hw - rr, bt);
        ctx.quadraticCurveTo(hw, bt, hw, bt + rr);
        ctx.lineTo(hw, bb - rr);
        ctx.quadraticCurveTo(hw, bb, hw - rr, bb);
        // Down the tail and back up, clamped inside the body so a tail near an
        // edge does not invert the outline.
        const t0 = clamp(tailX - tailW / 2, -hw + rr, hw - rr - tailW);
        ctx.lineTo(t0 + tailW, bb);
        ctx.lineTo(clamp(tailX, -hw + rr, hw - rr), bb + tailH);
        ctx.lineTo(t0, bb);
        ctx.lineTo(-hw + rr, bb);
        ctx.quadraticCurveTo(-hw, bb, -hw, bb - rr);
        ctx.lineTo(-hw, bt + rr);
        ctx.quadraticCurveTo(-hw, bt, -hw + rr, bt);
        ctx.closePath();
        return;
    }

    if (g === 'line') {
        // A rule: a rectangle of the stroke's thickness. Kept as a filled path
        // rather than a stroked segment so caps, dashes and a gradient fill all
        // behave the way they do for every other shape.
        const t = Math.max(1, clamp(num(layer.thickness, 0.08), 0.001, 1) * h);
        ctx.rect(-hw, -t / 2, w, t);
        return;
    }

    if (g === 'banner') {
        // A ribbon: a rectangle with notched ends — the "SALE" strip.
        const notch = clamp(num(layer.notch, 0.12), 0, 0.45) * w;
        ctx.moveTo(-hw, -hh);
        ctx.lineTo(hw, -hh);
        ctx.lineTo(hw - notch, 0);
        ctx.lineTo(hw, hh);
        ctx.lineTo(-hw, hh);
        ctx.lineTo(-hw + notch, 0);
        ctx.closePath();
        return;
    }

    if (g === 'blob') {
        // An organic sticker background. Deterministic: the wobble comes from
        // the point index, not from a random source, because a shape that
        // reshapes itself on every frame of a video is not a shape.
        const lobes = Math.max(4, points);
        const wobble = clamp(num(layer.innerRatio, 0.18), 0, 0.6);
        for (let i = 0; i <= lobes; i++) {
            const a = (i / lobes) * TAU;
            const r = 1 - wobble * Math.sin(a * 3 + 0.6) * Math.cos(a * 2);
            const x = Math.cos(a) * hw * r;
            const y = Math.sin(a) * hh * r;
            if (i === 0) ctx.moveTo(x, y);
            else {
                const pa = ((i - 0.5) / lobes) * TAU;
                const pr = 1 - wobble * Math.sin(pa * 3 + 0.6) * Math.cos(pa * 2);
                ctx.quadraticCurveTo(Math.cos(pa) * hw * pr * 1.06, Math.sin(pa) * hh * pr * 1.06, x, y);
            }
        }
        ctx.closePath();
        return;
    }

    if (g === 'diamond') {
        ctx.moveTo(0, -hh); ctx.lineTo(hw, 0); ctx.lineTo(0, hh); ctx.lineTo(-hw, 0);
        ctx.closePath();
        return;
    }

    if (g === 'pill') {
        // Fully rounded on the short axis — the chip a label sits in. Not a
        // `rect` with radius 0.5, because that clamps against BOTH axes and a
        // wide pill would come out as a stadium with square-ish ends.
        const r = Math.min(w, h) / 2;
        if (typeof ctx.roundRect === 'function') ctx.roundRect(-hw, -hh, w, h, r);
        else ctx.rect(-hw, -hh, w, h);
        return;
    }

    if (g === 'parallelogram' || g === 'trapezoid') {
        const s = clamp(num(layer.slant, 0.22), 0, 0.9) * w;
        if (g === 'parallelogram') {
            ctx.moveTo(-hw + s, -hh); ctx.lineTo(hw, -hh);
            ctx.lineTo(hw - s, hh); ctx.lineTo(-hw, hh);
        } else {
            // Symmetric taper: a trapezoid that leans is a parallelogram, and
            // the two are different requests.
            ctx.moveTo(-hw + s, -hh); ctx.lineTo(hw - s, -hh);
            ctx.lineTo(hw, hh); ctx.lineTo(-hw, hh);
        }
        ctx.closePath();
        return;
    }

    if (g === 'frame') {
        // A border band: the box with a smaller box punched out of it. Both
        // subpaths wind the same way, so the hole needs the even-odd rule —
        // buildPath says so by returning it.
        const t = clamp(num(layer.thickness, 0.12), 0.01, 0.49);
        ctx.rect(-hw, -hh, w, h);
        ctx.rect(-hw + w * t, -hh + h * t, w * (1 - t * 2), h * (1 - t * 2));
        return 'evenodd';
    }

    if (g === 'semicircle') {
        // Centred on the BOTTOM edge with a full-height radius, so the dome
        // fills the box rather than sitting in the top half of it.
        ctx.ellipse(0, hh, hw, h, 0, Math.PI, TAU);
        ctx.closePath();
        return;
    }

    if (g === 'ring') {
        const inner = clamp(num(layer.innerRatio, 0.6), 0.05, 0.95);
        ctx.ellipse(0, 0, hw, hh, 0, 0, TAU);
        ctx.ellipse(0, 0, hw * inner, hh * inner, 0, 0, TAU);
        return 'evenodd';
    }

    if (g === 'pie') {
        // A sector from twelve o'clock, clockwise. The default three quarters
        // rather than a half, because a half-circle is `semicircle` and this
        // control is for "most of a circle, with a bite out".
        const sweep = clamp(num(layer.sweep, 0.75), 0.01, 1) * TAU;
        const from = -Math.PI / 2;
        ctx.moveTo(0, 0);
        ctx.ellipse(0, 0, hw, hh, 0, from, from + sweep);
        ctx.closePath();
        return;
    }

    if (g === 'teardrop') {
        // Point up, so it reads as a drop rather than as a map pin. A pin is
        // this rotated 180°, which the layer's own rotation already offers.
        ctx.moveTo(0, -hh);
        ctx.bezierCurveTo(hw * 0.95, -hh * 0.25, hw, hh * 0.45, 0, hh);
        ctx.bezierCurveTo(-hw, hh * 0.45, -hw * 0.95, -hh * 0.25, 0, -hh);
        ctx.closePath();
        return;
    }

    if (g === 'cloud') {
        // Flat bottom, three bumps. Quadratics rather than arcs because the box
        // is not square in general, and an arc would go oval-shaped per bump at
        // different rates than the box does.
        ctx.moveTo(-hw, hh);
        ctx.quadraticCurveTo(-hw, -hh * 0.15, -hw * 0.46, -hh * 0.2);
        ctx.quadraticCurveTo(-hw * 0.34, -hh, hw * 0.04, -hh * 0.56);
        ctx.quadraticCurveTo(hw * 0.46, -hh, hw * 0.62, -hh * 0.14);
        ctx.quadraticCurveTo(hw, -hh * 0.08, hw, hh);
        ctx.closePath();
        return;
    }

    if (g === 'gear') {
        // Teeth as flat-topped blocks, plus a bore. `points` is the tooth
        // count, `thickness` how deep they cut, `innerRatio` the bore — all
        // three reusing fields other geometries already carry rather than
        // inventing a gear-only vocabulary.
        const teeth = Math.max(5, Math.round(num(layer.points, 8)));
        const depth = clamp(num(layer.thickness, 0.24), 0.02, 0.6);
        const root = 1 - depth;
        const step = TAU / teeth;
        const half = step * 0.26;
        const at = (a, r) => [Math.cos(a) * hw * r, Math.sin(a) * hh * r];
        let started = false;
        for (let i = 0; i < teeth; i++) {
            const a = -Math.PI / 2 + i * step;
            const pts = [
                at(a - half * 1.7, root), at(a - half, 1),
                at(a + half, 1), at(a + half * 1.7, root),
            ];
            for (const [x, y] of pts) {
                if (started) ctx.lineTo(x, y);
                else { ctx.moveTo(x, y); started = true; }
            }
        }
        ctx.closePath();
        const bore = clamp(num(layer.innerRatio, 0.34), 0, 0.8);
        if (bore > 0.02) ctx.ellipse(0, 0, hw * bore, hh * bore, 0, 0, TAU);
        return 'evenodd';
    }

    if (g === 'plus') {
        const a = clamp(num(layer.thickness, 0.34), 0.05, 0.95);
        const ax = hw * a;
        const ay = hh * a;
        ctx.moveTo(-ax, -hh); ctx.lineTo(ax, -hh); ctx.lineTo(ax, -ay);
        ctx.lineTo(hw, -ay); ctx.lineTo(hw, ay); ctx.lineTo(ax, ay);
        ctx.lineTo(ax, hh); ctx.lineTo(-ax, hh); ctx.lineTo(-ax, ay);
        ctx.lineTo(-hw, ay); ctx.lineTo(-hw, -ay); ctx.lineTo(-ax, -ay);
        ctx.closePath();
        return;
    }

    if (g === 'bolt') {
        ctx.moveTo(hw * 0.18, -hh);
        ctx.lineTo(-hw * 0.62, hh * 0.14);
        ctx.lineTo(-hw * 0.06, hh * 0.14);
        ctx.lineTo(-hw * 0.26, hh);
        ctx.lineTo(hw * 0.62, -hh * 0.2);
        ctx.lineTo(hw * 0.04, -hh * 0.2);
        ctx.closePath();
        return;
    }

    if (g === 'doubleArrow') {
        const shaft = clamp(num(layer.thickness, 0.42), 0.05, 0.95) * hh;
        const head = clamp(num(layer.headLength, 0.28), 0.05, 0.49) * w;
        ctx.moveTo(-hw, 0);
        ctx.lineTo(-hw + head, -hh); ctx.lineTo(-hw + head, -shaft);
        ctx.lineTo(hw - head, -shaft); ctx.lineTo(hw - head, -hh);
        ctx.lineTo(hw, 0);
        ctx.lineTo(hw - head, hh); ctx.lineTo(hw - head, shaft);
        ctx.lineTo(-hw + head, shaft); ctx.lineTo(-hw + head, hh);
        ctx.closePath();
        return;
    }

    if (g === 'wave') {
        // A sampled band rather than two quadratics: the crest count is a
        // control, and a curve per crest would be a loop of curves anyway.
        const amp = clamp(num(layer.thickness, 0.3), 0.02, 0.5) * h;
        const cycles = Math.max(0.5, num(layer.points, 2));
        const band = h - amp * 2;
        const N = 48;
        const yTop = (t) => -hh + amp + Math.sin(t * TAU * cycles) * amp;
        for (let i = 0; i <= N; i++) {
            const t = i / N;
            const x = -hw + t * w;
            if (i === 0) ctx.moveTo(x, yTop(t)); else ctx.lineTo(x, yTop(t));
        }
        for (let i = N; i >= 0; i--) {
            const t = i / N;
            ctx.lineTo(-hw + t * w, yTop(t) + band);
        }
        ctx.closePath();
        return;
    }

    if (g === 'chevron') {
        // The band, not the outline: a banner whose left edge notches the same
        // way its right edge points, so a row of them reads as a flow.
        const n = clamp(num(layer.notch, 0.24), 0.01, 0.9) * w;
        ctx.moveTo(-hw, -hh);
        ctx.lineTo(hw - n, -hh); ctx.lineTo(hw, 0); ctx.lineTo(hw - n, hh);
        ctx.lineTo(-hw, hh); ctx.lineTo(-hw + n, 0);
        ctx.closePath();
        return;
    }

    if (g === 'tag') {
        // The cut corner is on the LEFT, where the string goes, and the eyelet
        // is a real hole so a coloured background shows through it.
        const cut = clamp(num(layer.notch, 0.18), 0.02, 0.6) * w;
        ctx.moveTo(-hw + cut, -hh);
        ctx.lineTo(hw, -hh); ctx.lineTo(hw, hh); ctx.lineTo(-hw + cut, hh);
        ctx.lineTo(-hw, 0);
        ctx.closePath();
        const eye = clamp(num(layer.innerRatio, 0.12), 0, 0.4) * Math.min(w, h);
        if (eye > 0.5) ctx.ellipse(-hw + cut * 0.75, 0, eye, eye, 0, 0, TAU);
        return 'evenodd';
    }

    if (g === 'bookmark') {
        const n = clamp(num(layer.notch, 0.24), 0.01, 0.9) * h;
        ctx.moveTo(-hw, -hh); ctx.lineTo(hw, -hh); ctx.lineTo(hw, hh);
        ctx.lineTo(0, hh - n); ctx.lineTo(-hw, hh);
        ctx.closePath();
        return;
    }

    if (g === 'shield') {
        const shoulder = -hh + h * clamp(num(layer.slant, 0.45), 0.05, 0.9);
        ctx.moveTo(-hw, -hh);
        ctx.lineTo(hw, -hh);
        ctx.lineTo(hw, shoulder);
        ctx.quadraticCurveTo(hw, hh, 0, hh);
        ctx.quadraticCurveTo(-hw, hh, -hw, shoulder);
        ctx.closePath();
        return;
    }

    // rect, and the fallback for an unknown geometry — a recipe from a newer
    // editor degrades to a plain box rather than to nothing at all.
    const r = clamp(num(layer.radius, 0), 0, 0.5) * Math.min(w, h);
    if (r > 0 && typeof ctx.roundRect === 'function') ctx.roundRect(-hw, -hh, w, h, r);
    else ctx.rect(-hw, -hh, w, h);
    return undefined;
}

// ── painting ─────────────────────────────────────────────────────────────────

/** A fill spec → a fillStyle, in the shape's own box. Shared shape with wordart. */
function fillStyle(ctx, fill, w, h) {
    if (!fill || !fill.kind || fill.kind === 'solid') return (fill && fill.color) || '#ffc107';
    const stops = Array.isArray(fill.stops) && fill.stops.length
        ? fill.stops
        : [[0, '#ffffff'], [1, '#888888']];
    let g;
    if (fill.kind === 'radial') {
        const r = Math.max(w, h) / 2;
        g = ctx.createRadialGradient(0, 0, r * 0.05, 0, 0, r);
    } else {
        const a = (num(fill.angle, 90) * Math.PI) / 180;
        const half = (Math.abs(Math.cos(a)) * w + Math.abs(Math.sin(a)) * h) / 2;
        g = ctx.createLinearGradient(-Math.cos(a) * half, -Math.sin(a) * half, Math.cos(a) * half, Math.sin(a) * half);
    }
    for (const [at, color] of stops) g.addColorStop(clamp(num(at, 0), 0, 1), color);
    return g;
}

function paintShape(ctx, plan, layer) {
    // ONE sizing rule, shared with shapeBounds below. If the painter and the
    // bounds ever disagree, the handles stop matching the shape and every drag
    // is off by the difference — so they read the same helper, not two copies.
    const { w, h } = shapeSize(plan, layer);
    if (w <= 0 || h <= 0) return;

    ctx.save();
    ctx.translate(layer.x, layer.y);

    const sh = layer.shadow;
    if (sh) {
        ctx.shadowColor = sh.color || 'rgba(0,0,0,0.35)';
        ctx.shadowBlur = Math.max(0, num(sh.blur, 0.02) * plan.W);
        ctx.shadowOffsetX = num(sh.dx, 0) * plan.W;
        ctx.shadowOffsetY = num(sh.dy, 0.01) * plan.W;
    }

    // Geometries with a hole in them — the ring, the frame, the tag's eyelet,
    // the gear's bore — wind their inner subpath the same way as the outer one,
    // so they need the even-odd rule to punch rather than to overpaint.
    // buildPath returns it; everything else returns nothing and fills the way
    // it always has, which keeps the baseline geometries' calls unchanged.
    const rule = buildPath(ctx, layer, w, h);

    if (layer.fill !== null) {
        ctx.fillStyle = fillStyle(ctx, layer.fill, w, h);
        if (rule) ctx.fill(rule); else ctx.fill();
    }

    // The outline is drawn after the fill and without the shadow: a shadow cast
    // by both puts a second, offset copy of the rim under the shape.
    if (layer.stroke && num(layer.stroke.width, 0) > 0) {
        ctx.shadowColor = 'transparent';
        ctx.shadowBlur = 0;
        ctx.strokeStyle = layer.stroke.color || '#000000';
        ctx.lineWidth = Math.max(1, num(layer.stroke.width, 0.006) * plan.W);
        ctx.lineJoin = 'round';
        if (Array.isArray(layer.stroke.dash) && layer.stroke.dash.length) {
            ctx.setLineDash(layer.stroke.dash.map((d) => Math.max(1, num(d, 4) * plan.W)));
        }
        ctx.stroke();
    }

    ctx.restore();
}

/** The shape's pixel size. A missing `fh` means square in PIXELS, not in fractions. */
function shapeSize(plan, layer) {
    const w = plan.W * (layer.fw == null ? 0.3 : layer.fw);
    return { w, h: layer.fh ? plan.H * layer.fh : w };
}

/** Centred fractional box — the same convention `fractionalBounds` assumes. */
function shapeBounds(ctx, plan, layer) {
    const { w, h } = shapeSize(plan, layer);
    return { x: layer.x - w / 2, y: layer.y - h / 2, w, h };
}

// ── the stored form ──────────────────────────────────────────────────────────

function compileShape(patch, cx) {
    const { W, H } = cx;
    const fw = num(patch.fw, 0.3);
    const fh = patch.fh == null ? +((W * fw) / H).toFixed(4) : num(patch.fh, 0.3);
    const env = cx.envelope(patch.anim);

    return {
        name: patch.name || 'Shape',
        geometry: GEOMETRIES.includes(patch.geometry) ? patch.geometry : 'rect',
        visible: patch.visible !== false,
        timing: patch.timing || null,
        enter: patch.enter || env.enter,
        exit: patch.exit || env.exit,
        z: patch.z == null ? cx.nextZ() : patch.z,
        x: W * (patch.fx == null ? 0.5 : patch.fx),
        y: H * (patch.fy == null ? 0.5 : patch.fy),
        fx: patch.fx == null ? 0.5 : patch.fx,
        fy: patch.fy == null ? 0.5 : patch.fy,
        fw, fh,
        // `fill: null` is meaningful — an outline-only shape — so it survives
        // the default, which a `||` would swallow.
        fill: patch.fill === null ? null : (patch.fill || { kind: 'solid', color: cx.color(patch.color, cx.theme.accent) }),
        stroke: patch.stroke || null,
        shadow: patch.shadow || null,
        radius: patch.radius,
        points: patch.points,
        innerRatio: patch.innerRatio,
        thickness: patch.thickness,
        headLength: patch.headLength,
        notch: patch.notch,
        // The lean of a parallelogram / trapezoid, and a shield's shoulder.
        slant: patch.slant,
        // How much of a circle a pie slice takes, as a fraction of the turn.
        sweep: patch.sweep,
        tailAt: patch.tailAt,
        tailWidth: patch.tailWidth,
        tailHeight: patch.tailHeight,
        rot: patch.rot || 0,
        opacity: patch.opacity == null ? 1 : patch.opacity,
        blend: patch.blend || null,
        mask: patch.mask || null,
        anim: patch.anim || 'none',
        keys: cx.keys(patch.keys),
    };
}

export const SHAPE = registerLayerType('shape', {
    paint: paintShape,
    compile: compileShape,
    bounds: shapeBounds,
});

/**
 * Draw one geometry into a box, for a palette tile.
 *
 * THE TILE IS THE REAL SHAPE, painted by the code that will paint the layer.
 * The alternative — an icon font glyph that approximates each geometry — is
 * how a palette comes to lie: the tile for `blob` would be a circle forever,
 * and a geometry added here would need a second artefact drawn somewhere else
 * before anyone could find it.
 *
 * Exported from the geometry module rather than reimplemented in the app for
 * the same reason `paintShape` and `shapeBounds` share `shapeSize`: two copies
 * of a shape's definition disagree, and the disagreement shows up as a tile
 * that does not match what it inserts.
 *
 * The caller sets its own fillStyle/strokeStyle beforehand; this only builds
 * the path and fills it, centred in `(0,0)`-origin coordinates the caller has
 * translated to.
 */
export function paintShapeTile(ctx, geometry, w, h, params = {}) {
    const layer = { geometry, ...params };
    const rule = buildPath(ctx, layer, w, h);
    if (rule) ctx.fill(rule); else ctx.fill();
    return rule;
}
