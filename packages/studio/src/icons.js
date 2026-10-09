/**
 * Icon layers — vector glyphs, drawn from path data.
 *
 * WHY NOT AN ICON FONT, which is the obvious answer and the wrong one. This
 * renderer has two hosts: the editor's browser and the render worker's headless
 * page. A webfont that has loaded in one and not the other produces tofu, or —
 * worse — a *different* glyph from a fallback family, and the difference shows
 * up only in the finished video. "Renders identically" is the property this
 * whole package is organised around, and a font is a runtime dependency that
 * cannot be made to hold it. Path data is just numbers: it draws the same on
 * any host, needs nothing installed, and stays inside the zero-dependency
 * contract that lets the worker run this package unchanged.
 *
 * WHY NOT @rutba/drawing OR AN ICON LIBRARY, for the same reason shapes.js
 * gives: those render SVG, this renders canvas 2D, and importing a scene graph
 * to reimplement its output as canvas paths costs the dependency and pays for
 * it with a conversion. The path strings below are the cheap part.
 *
 * THE GLYPHS ARE COMPOSED, NOT TYPED OUT. Every one is built from `circle`,
 * `rect`, `rrect`, `poly` and `spoke` below rather than written as a literal
 * `d` string. Hand-typing path data is how an icon set acquires a glyph that is
 * two units off centre and stays that way for a year: the coordinates are
 * unreviewable. Composed, each icon reads as what it is — a circle here, a bar
 * there — and `tests/icons.test.mjs` checks every result parses and stays in
 * the box.
 *
 * EVERY ICON IS DRAWN EVEN-ODD, so an inner subpath punches a hole rather than
 * painting over its parent. That is what makes a ring a ring and a clock face
 * show its hands. The consequence to design around: two subpaths that OVERLAP
 * without meaning to punch will cancel where they meet, so a composed icon
 * keeps its parts either nested or apart.
 *
 * All coordinates live in a 24×24 box, the same convention every icon set
 * uses, and the painter scales that box into the layer.
 */

import { registerLayerType } from './index.js';

const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));
const num = (v, fallback) => (Number.isFinite(Number(v)) ? Number(v) : fallback);

// ── path builders ────────────────────────────────────────────────────────────

const n = (v) => {
    const r = Math.round(v * 100) / 100;
    return Object.is(r, -0) ? 0 : r;
};

/** A full circle, as the two-arc idiom every path renderer agrees on. */
const circle = (cx, cy, r) =>
    `M${n(cx - r)} ${n(cy)}a${n(r)} ${n(r)} 0 1 0 ${n(r * 2)} 0a${n(r)} ${n(r)} 0 1 0 ${n(-r * 2)} 0Z`;

const rect = (x, y, w, h) => `M${n(x)} ${n(y)}h${n(w)}v${n(h)}h${n(-w)}Z`;

/** A rounded rectangle. `r` is clamped so a fat radius cannot invert the path. */
const rrect = (x, y, w, h, r) => {
    const k = clamp(r, 0, Math.min(w, h) / 2);
    return `M${n(x + k)} ${n(y)}h${n(w - k * 2)}a${n(k)} ${n(k)} 0 0 1 ${n(k)} ${n(k)}`
        + `v${n(h - k * 2)}a${n(k)} ${n(k)} 0 0 1 ${n(-k)} ${n(k)}`
        + `h${n(-(w - k * 2))}a${n(k)} ${n(k)} 0 0 1 ${n(-k)} ${n(-k)}`
        + `v${n(-(h - k * 2))}a${n(k)} ${n(k)} 0 0 1 ${n(k)} ${n(-k)}Z`;
};

const poly = (...pts) => `M${pts.map(([x, y]) => `${n(x)} ${n(y)}`).join('L')}Z`;

/** A quad from (x1,y1) to (x2,y2) of a given half-thickness — a thick line. */
function bar(x1, y1, x2, y2, half) {
    const dx = x2 - x1;
    const dy = y2 - y1;
    const len = Math.hypot(dx, dy) || 1;
    const px = (-dy / len) * half;
    const py = (dx / len) * half;
    return poly([x1 + px, y1 + py], [x2 + px, y2 + py], [x2 - px, y2 - py], [x1 - px, y1 - py]);
}

/** One ray of a sun, from radius r0 to r1 at an angle, as a quad. */
function spoke(cx, cy, deg, r0, r1, half) {
    const a = (deg * Math.PI) / 180;
    return bar(cx + Math.cos(a) * r0, cy + Math.sin(a) * r0,
        cx + Math.cos(a) * r1, cy + Math.sin(a) * r1, half);
}

/** An n-pointed star, so the geometry is computed rather than typed. */
function starPath(cx, cy, points, R, r) {
    const pts = [];
    for (let i = 0; i < points * 2; i++) {
        const a = -Math.PI / 2 + (i / (points * 2)) * Math.PI * 2;
        const rad = i % 2 === 0 ? R : r;
        pts.push([cx + Math.cos(a) * rad, cy + Math.sin(a) * rad]);
    }
    return poly(...pts);
}

// ── the set ──────────────────────────────────────────────────────────────────

/**
 * A curated set, not a library.
 *
 * Chosen for what a social or retail creative actually reaches for — a price
 * badge, a call to action, a delivery promise — rather than for coverage. The
 * table is the extension point: an icon is a name and a path, so an org's own
 * marks can be added here (or, later, uploaded) without touching the painter.
 */
export const ICONS = {
    // ── marks ────────────────────────────────────────────────────────────────
    check: poly([3.2, 13.4], [9, 19.2], [20.8, 7.4], [18.6, 5.2], [9, 14.8], [5.4, 11.2]),
    // Sized to span the same ~18 units the rest of the set does. The first cut
    // spanned 14.8 and looked lighter than everything beside it — caught by the
    // bounding-box test rather than by eye, which is the point of having it.
    close: poly([5.11, 2.9], [12, 9.79], [18.89, 2.9], [21.1, 5.11], [14.21, 12], [21.1, 18.89],
        [18.89, 21.1], [12, 14.21], [5.11, 21.1], [2.9, 18.89], [9.79, 12], [2.9, 5.11]),
    plus: poly([10.4, 3], [13.6, 3], [13.6, 10.4], [21, 10.4], [21, 13.6], [13.6, 13.6],
        [13.6, 21], [10.4, 21], [10.4, 13.6], [3, 13.6], [3, 10.4], [10.4, 10.4]),
    minus: rect(3, 10.4, 18, 3.2),
    star: starPath(12, 12.3, 5, 9.7, 4.3),
    sparkle: poly([12, 2.5], [13.9, 10.1], [21.5, 12], [13.9, 13.9], [12, 21.5],
        [10.1, 13.9], [2.5, 12], [10.1, 10.1]),
    heart: 'M12 21.2C12 21.2 2.8 15 2.8 8.9 2.8 5.6 5.4 3 8.7 3 10.5 3 12 4.3 12 4.3'
        + ' 12 4.3 13.5 3 15.3 3 18.6 3 21.2 5.6 21.2 8.9 21.2 15 12 21.2 12 21.2Z',
    circle: circle(12, 12, 9.4),
    square: rrect(3, 3, 18, 18, 2.4),
    // A ring, and the reason even-odd matters: the inner circle punches.
    ring: circle(12, 12, 9.4) + circle(12, 12, 5.6),

    // ── arrows ───────────────────────────────────────────────────────────────
    arrowRight: poly([3, 10.2], [14, 10.2], [14, 5.4], [21.4, 12], [14, 18.6], [14, 13.8], [3, 13.8]),
    arrowLeft: poly([21, 10.2], [10, 10.2], [10, 5.4], [2.6, 12], [10, 18.6], [10, 13.8], [21, 13.8]),
    arrowUp: poly([10.2, 21], [10.2, 10], [5.4, 10], [12, 2.6], [18.6, 10], [13.8, 10], [13.8, 21]),
    arrowDown: poly([10.2, 3], [10.2, 14], [5.4, 14], [12, 21.4], [18.6, 14], [13.8, 14], [13.8, 3]),
    chevronRight: poly([8.4, 3.4], [17.6, 12], [8.4, 20.6], [5.9, 18.1], [12.6, 12], [5.9, 5.9]),
    chevronDown: poly([3.4, 8.4], [12, 17.6], [20.6, 8.4], [18.1, 5.9], [12, 12.6], [5.9, 5.9]),

    // ── media ────────────────────────────────────────────────────────────────
    play: poly([6.5, 3.5], [20, 12], [6.5, 20.5]),
    pause: rect(6, 3.5, 4, 17) + rect(14, 3.5, 4, 17),
    stop: rrect(3.6, 3.6, 16.8, 16.8, 2.2),
    camera: rrect(2.5, 6.5, 19, 14, 2.4)
        + poly([8, 6.6], [9.4, 3.6], [14.6, 3.6], [16, 6.6])
        + circle(12, 13.4, 4.4),
    image: rrect(2.5, 4.5, 19, 15, 2.2)
        + circle(8, 9.4, 1.9)
        + poly([4.4, 17.5], [10, 11.4], [14, 15.6], [16.4, 13.2], [19.6, 17.5]),
    music: poly([11, 4], [19, 2.2], [19, 5.4], [11, 7.2])
        + rect(11, 7.2, 1.9, 9.6)
        + circle(9, 17.2, 3.1) + circle(17.1, 15.4, 3.1) + rect(17.1, 5.4, 1.9, 10),
    mic: rrect(9.2, 2.5, 5.6, 11.6, 2.8)
        + poly([6, 11.4], [7.7, 11.4], [7.7, 13], [16.3, 13], [16.3, 11.4], [18, 11.4],
            [18, 13.6], [12.9, 17.6], [12.9, 20], [15.4, 20], [15.4, 21.6], [8.6, 21.6],
            [8.6, 20], [11.1, 20], [11.1, 17.6], [6, 13.6]),

    // ── people and talking ───────────────────────────────────────────────────
    user: circle(12, 7.6, 4.6) + 'M12 14.2c-4.5 0-8.2 2.9-8.2 6.6V21.8h16.4V20.8c0-3.7-3.7-6.6-8.2-6.6Z',
    users: circle(8.6, 7.4, 4)
        + 'M8.6 13.4c-3.9 0-6.9 2.5-6.9 5.6V20.6h13.8V19c0-3.1-3-5.6-6.9-5.6Z'
        + circle(17.6, 8.2, 3.1)
        + 'M17.9 14.2c-.6 0-1.2.1-1.7.3 1.1 1.3 1.8 3 1.8 4.5V20.6h4.3V19c0-2.6-1.9-4.8-4.4-4.8Z',
    chat: 'M2.6 4h18.8v12.4h-8.2l-5.1 4.6v-4.6H2.6Z',
    bell: 'M12 2.4a2 2 0 0 1 2 2v.9c2.9 1 4.9 3.6 4.9 6.7v4.4l2 2.5H3.1l2-2.5V12c0-3.1 2-5.7 4.9-6.7v-.9a2 2 0 0 1 2-2Z'
        + 'M9.4 20.1h5.2a2.6 2.6 0 0 1-5.2 0Z',
    share: circle(18.4, 5.4, 3) + circle(5.6, 12, 3) + circle(18.4, 18.6, 3)
        + bar(7.9, 10.8, 16.2, 6.6, 0.85) + bar(7.9, 13.2, 16.2, 17.4, 0.85),
    thumbsUp: 'M7.4 10.6 12 2.8c1.6 0 2.7 1.2 2.7 2.8v3.6h4.6c1.3 0 2.3 1.2 2 2.5l-1.6 7.2c-.2 1-1.1 1.7-2.1 1.7H7.4Z'
        + rect(2.6, 10.6, 3.4, 10.4),

    // ── places and time ──────────────────────────────────────────────────────
    home: poly([12, 2.4], [22, 11.2], [19, 11.2], [19, 21.2], [14.2, 21.2], [14.2, 15],
        [9.8, 15], [9.8, 21.2], [5, 21.2], [5, 11.2], [2, 11.2]),
    pin: 'M12 2.2c-4 0-7.2 3.2-7.2 7.2 0 5.4 7.2 12.4 7.2 12.4s7.2-7 7.2-12.4c0-4-3.2-7.2-7.2-7.2Z'
        + circle(12, 9.4, 2.7),
    clock: circle(12, 12, 9.4)
        + poly([11.1, 5.6], [12.9, 5.6], [12.9, 12.2], [17.4, 14.8], [16.5, 16.3], [11.1, 13.1]),
    calendar: rrect(2.8, 4.4, 18.4, 16.8, 2)
        + rect(2.8, 4.4, 18.4, 4.4)
        + rect(6.6, 11.4, 3, 3) + rect(10.5, 11.4, 3, 3) + rect(14.4, 11.4, 3, 3)
        + rect(6.6, 16, 3, 3) + rect(10.5, 16, 3, 3),

    // ── commerce ─────────────────────────────────────────────────────────────
    cart: 'M2.4 3h3.2l.9 3.2h15.1l-2.5 8.8H8l-.5 2.2h12v2.2H5.6L3.9 5.2H2.4Z'
        + circle(9, 20.6, 1.7) + circle(17.4, 20.6, 1.7),
    tag: poly([2.6, 12], [11.4, 3.2], [21.4, 3.2], [21.4, 13.2], [12.6, 22])
        + circle(17.2, 7.4, 1.9),
    gift: rrect(2.6, 9.6, 18.8, 11.8, 1.4) + rect(2.6, 7, 18.8, 3.6) + rect(10.6, 7, 2.8, 14.4)
        + circle(8.4, 4.6, 2.4) + circle(15.6, 4.6, 2.4),
    percent: circle(7.2, 7.2, 3.4) + circle(16.8, 16.8, 3.4) + bar(18.2, 4.4, 5.8, 19.6, 1.4),
    card: rrect(2, 4.6, 20, 14.8, 2.2) + rect(2, 8, 20, 3.4) + rect(4.8, 14.4, 5.6, 2.2),
    truck: rect(1.6, 6, 12.4, 9.6)
        + poly([14, 8.8], [18.4, 8.8], [22.4, 12.4], [22.4, 15.6], [14, 15.6])
        + circle(6.6, 18.2, 2.4) + circle(17.6, 18.2, 2.4),
    store: poly([2, 8.6], [4.2, 3.4], [19.8, 3.4], [22, 8.6])
        + rect(3.4, 10, 17.2, 11) + rect(8.4, 14, 4.4, 7),

    // ── attention ────────────────────────────────────────────────────────────
    fire: 'M12 1.8c3.4 3.6 5.4 6.4 5.4 8.8 0 1.4-.6 2.4-1.6 3 .3-1.9-.5-3.6-2.2-5.1'
        + ' .5 3.4-1.2 4.9-2.6 6.3-1 1-1.6 2.1-1.6 3.4 0 .9.3 1.7.8 2.4'
        + '-2.6-.9-4.6-3.4-4.6-6.4 0-4.4 3.6-6.6 6.4-12.4Z',
    bolt: poly([13.4, 1.8], [4.6, 13.4], [10.4, 13.4], [8.6, 22.2], [19.4, 9.6], [13, 9.6]),
    crown: poly([2.2, 7.4], [7.2, 11.4], [12, 3.6], [16.8, 11.4], [21.8, 7.4], [19.8, 18.4], [4.2, 18.4])
        + rect(4.2, 19.4, 15.6, 2.2),
    trophy: 'M7 3h10v6.6a5 5 0 0 1-10 0Z'
        + rect(10.6, 14.4, 2.8, 4) + rect(7, 18.4, 10, 2.6)
        + 'M7 4.4H3.6v2.2A4 4 0 0 0 7 10.5Z'
        + 'M17 4.4h3.4v2.2A4 4 0 0 1 17 10.5Z',
    rocket: 'M12 1.6c3.4 2.8 5.2 6.6 5.2 10.8l-1.8 3.4H8.6L6.8 12.4C6.8 8.2 8.6 4.4 12 1.6Z'
        + circle(12, 8.6, 2.2)
        + poly([6.6, 12.6], [4.2, 17], [4.2, 20.2], [8.2, 17.6])
        + poly([17.4, 12.6], [19.8, 17], [19.8, 20.2], [15.8, 17.6])
        + poly([10.2, 17.4], [13.8, 17.4], [12, 22.4]),
    bulb: circle(12, 9.4, 6.4)
        + rect(9.4, 15.4, 5.2, 3.2) + rect(10.2, 19.4, 3.6, 2.2),
    leaf: 'M20.4 3.4C20.4 3.4 8.4 2.4 5 9.4c-2.4 5 .6 9.4 2.6 11.2L9.2 19'
        + 'C7.4 17.2 6 14.6 7.4 11.4 9.4 7 15.4 6.6 15.4 6.6S9.6 9 8.4 15.6c3.8 2.6 8.8.4 10.4-3.6 1.4-3.4 1.6-8.6 1.6-8.6Z',

    // ── things ───────────────────────────────────────────────────────────────
    search: circle(10.4, 10.4, 7.2) + circle(10.4, 10.4, 4.9)
        + bar(15.2, 16.4, 20.8, 21.2, 1.4),
    lock: rrect(4.6, 10.2, 14.8, 11.2, 2)
        + 'M8.2 10.2V7.6a3.8 3.8 0 0 1 7.6 0v2.6h-2.3V7.6a1.5 1.5 0 0 0-3 0v2.6Z'
        + circle(12, 15.4, 1.8),
    key: circle(7.6, 16.2, 5) + circle(7.6, 16.2, 1.9)
        + bar(10.6, 12.6, 20.4, 3.4, 1.4)
        + bar(17.2, 6.8, 19.4, 9, 1.2)
        + bar(14.6, 9.4, 16.4, 11.2, 1.2),
    eye: 'M12 4.6C6.6 4.6 2.4 8.6 1 12c1.4 3.4 5.6 7.4 11 7.4S21.6 15.4 23 12c-1.4-3.4-5.6-7.4-11-7.4Z'
        + circle(12, 12, 3.8),
    shield: 'M12 2 21 5.2v6.2c0 5.2-3.6 9.4-9 10.6-5.4-1.2-9-5.4-9-10.6V5.2Z',
    download: poly([10.2, 2.6], [13.8, 2.6], [13.8, 11.4], [18, 11.4], [12, 18.2], [6, 11.4], [10.2, 11.4])
        + rect(3.6, 19.4, 16.8, 2.4),
    upload: poly([10.2, 20.4], [13.8, 20.4], [13.8, 11.6], [18, 11.6], [12, 4.8], [6, 11.6], [10.2, 11.6])
        + rect(3.6, 2.2, 16.8, 2.4),
    flag: poly([3.4, 2.4], [5.8, 2.4], [5.8, 21.6], [3.4, 21.6])
        + poly([6.6, 3.2], [20.4, 3.2], [17.4, 8], [20.4, 12.8], [6.6, 12.8]),
    sun: circle(12, 12, 5.2)
        + spoke(12, 12, -90, 7.4, 10.8, 1.1) + spoke(12, 12, -45, 7.4, 10.8, 1.1)
        + spoke(12, 12, 0, 7.4, 10.8, 1.1) + spoke(12, 12, 45, 7.4, 10.8, 1.1)
        + spoke(12, 12, 90, 7.4, 10.8, 1.1) + spoke(12, 12, 135, 7.4, 10.8, 1.1)
        + spoke(12, 12, 180, 7.4, 10.8, 1.1) + spoke(12, 12, -135, 7.4, 10.8, 1.1),
    // A crescent: one disc with a second punched out of it, offset. The whole
    // glyph is one even-odd fill — there is no subtraction operator in a path.
    moon: circle(12, 12, 9.4) + circle(16.6, 8.2, 8),
    droplet: 'M12 2.2c4.2 5 6.6 8.4 6.6 11.4a6.6 6.6 0 0 1-13.2 0c0-3 2.4-6.4 6.6-11.4Z',
};

/** Every icon name, in the order the palette should offer them. */
export const ICON_NAMES = Object.keys(ICONS);

/**
 * The palette's headings.
 *
 * Same arrangement as `SHAPE_GROUPS` and for the same reason: grouping is
 * presentation, and an editor should not have to know which of forty names is a
 * "mark" in order to lay out a panel.
 */
export const ICON_GROUPS = [
    { label: 'Marks', names: ['check', 'close', 'plus', 'minus', 'star', 'sparkle', 'heart', 'circle', 'square', 'ring'] },
    { label: 'Arrows', names: ['arrowRight', 'arrowLeft', 'arrowUp', 'arrowDown', 'chevronRight', 'chevronDown'] },
    { label: 'Media', names: ['play', 'pause', 'stop', 'camera', 'image', 'music', 'mic'] },
    { label: 'People', names: ['user', 'users', 'chat', 'bell', 'share', 'thumbsUp'] },
    { label: 'Place & time', names: ['home', 'pin', 'clock', 'calendar'] },
    { label: 'Commerce', names: ['cart', 'tag', 'gift', 'percent', 'card', 'truck', 'store'] },
    { label: 'Attention', names: ['fire', 'bolt', 'crown', 'trophy', 'rocket', 'bulb', 'leaf'] },
    { label: 'Things', names: ['search', 'lock', 'key', 'eye', 'shield', 'download', 'upload', 'flag', 'sun', 'moon', 'droplet'] },
];

/** What each icon is called, for a tile caption and for search. */
export const ICON_LABELS = {
    check: 'Check', close: 'Close', plus: 'Plus', minus: 'Minus', star: 'Star',
    sparkle: 'Sparkle', heart: 'Heart', circle: 'Circle', square: 'Square', ring: 'Ring',
    arrowRight: 'Arrow right', arrowLeft: 'Arrow left', arrowUp: 'Arrow up',
    arrowDown: 'Arrow down', chevronRight: 'Chevron', chevronDown: 'Chevron down',
    play: 'Play', pause: 'Pause', stop: 'Stop', camera: 'Camera', image: 'Picture',
    music: 'Music', mic: 'Microphone', user: 'Person', users: 'People', chat: 'Chat',
    bell: 'Bell', share: 'Share', thumbsUp: 'Thumbs up', home: 'Home', pin: 'Location',
    clock: 'Clock', calendar: 'Calendar', cart: 'Cart', tag: 'Tag', gift: 'Gift',
    percent: 'Percent', card: 'Card', truck: 'Delivery', store: 'Store', fire: 'Fire',
    bolt: 'Lightning', crown: 'Crown', trophy: 'Trophy', rocket: 'Rocket', bulb: 'Idea',
    leaf: 'Leaf', search: 'Search', lock: 'Lock', key: 'Key', eye: 'Eye', shield: 'Shield',
    download: 'Download', upload: 'Upload', flag: 'Flag', sun: 'Sun', moon: 'Moon',
    droplet: 'Droplet',
};

// ── painting ─────────────────────────────────────────────────────────────────

const BOX = 24;

/**
 * `Path2D` objects, parsed once per icon name.
 *
 * Parsing a path string is not free and a video asks for the same glyph thirty
 * times a second. Cached at module scope rather than per plan because the paths
 * are constants — the same memoisation `wordart.js` does for its layout, for
 * the same reason.
 *
 * Built lazily: `Path2D` is a browser API, and this module is imported by tests
 * that never paint.
 */
const parsed = new Map();
function pathFor(name) {
    if (parsed.has(name)) return parsed.get(name);
    const d = ICONS[name];
    let p = null;
    if (d && typeof Path2D !== 'undefined') {
        try { p = new Path2D(d); } catch { p = null; }
    }
    parsed.set(name, p);
    return p;
}

/** A fill spec → a fillStyle, in the icon's own box. Shared shape with shapes.js. */
function fillStyle(ctx, fill, w, h) {
    if (!fill || !fill.kind || fill.kind === 'solid') return (fill && fill.color) || '#ffffff';
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

/** An icon is square: one size, scaled uniformly so a glyph never distorts. */
function iconSize(plan, layer) {
    const s = plan.W * (layer.fw == null ? 0.16 : layer.fw);
    return { w: s, h: s };
}

function iconBounds(ctx, plan, layer) {
    const { w, h } = iconSize(plan, layer);
    return { x: layer.x - w / 2, y: layer.y - h / 2, w, h };
}

function paintIcon(ctx, plan, layer) {
    const { w, h } = iconSize(plan, layer);
    if (w <= 0) return;
    const path = pathFor(layer.name);
    // An unknown name paints NOTHING rather than a placeholder glyph. A
    // creative from a newer editor should lose one mark, not gain a wrong one —
    // the same forward-compatibility rule paintFrame keeps for unknown types.
    if (!path) return;

    ctx.save();
    ctx.translate(layer.x, layer.y);

    const sh = layer.shadow;
    if (sh) {
        ctx.shadowColor = sh.color || 'rgba(0,0,0,0.35)';
        ctx.shadowBlur = Math.max(0, num(sh.blur, 0.02) * plan.W);
        ctx.shadowOffsetX = num(sh.dx, 0) * plan.W;
        ctx.shadowOffsetY = num(sh.dy, 0.01) * plan.W;
    }

    // Scale the CONTEXT rather than the path, so the fill, the stroke width and
    // the shadow all stay measured in frame pixels — the same units every other
    // layer type uses. The gradient is built in layer space before the scale
    // for exactly that reason.
    const style = fillStyle(ctx, layer.fill, w, h);
    const s = w / BOX;
    ctx.scale(s, s);
    ctx.translate(-BOX / 2, -BOX / 2);

    if (layer.fill !== null) {
        ctx.fillStyle = style;
        // Even-odd throughout: an inner subpath punches a hole. See the note at
        // the top — it is what makes the ring a ring and the clock show hands.
        ctx.fill(path, 'evenodd');
    }

    if (layer.stroke && num(layer.stroke.width, 0) > 0) {
        ctx.shadowColor = 'transparent';
        ctx.shadowBlur = 0;
        ctx.strokeStyle = layer.stroke.color || '#000000';
        // Divided by the scale so a stroke width stated as a fraction of the
        // frame comes out that wide, not that wide times the glyph scale.
        ctx.lineWidth = Math.max(0.2, (num(layer.stroke.width, 0.006) * plan.W) / s);
        ctx.lineJoin = 'round';
        ctx.stroke(path);
    }

    ctx.restore();
}

function compileIcon(patch, cx) {
    const { W, H } = cx;
    const fw = num(patch.fw, 0.16);
    return {
        name: ICONS[patch.name] ? patch.name : 'star',
        label: patch.label || ICON_LABELS[patch.name] || 'Icon',
        visible: patch.visible !== false,
        timing: patch.timing || null,
        enter: patch.enter || cx.envelope(patch.anim).enter,
        exit: patch.exit || cx.envelope(patch.anim).exit,
        z: patch.z == null ? cx.nextZ() : patch.z,
        x: W * (patch.fx == null ? 0.5 : patch.fx),
        y: H * (patch.fy == null ? 0.5 : patch.fy),
        fx: patch.fx == null ? 0.5 : patch.fx,
        fy: patch.fy == null ? 0.5 : patch.fy,
        fw,
        // Square in PIXELS, so an icon is round on a 9:16 frame — the same rule
        // shapeSize keeps for a shape with no explicit height.
        fh: +((W * fw) / H).toFixed(4),
        fill: patch.fill === null ? null : (patch.fill || { kind: 'solid', color: cx.color(patch.color, cx.theme.text) }),
        stroke: patch.stroke || null,
        shadow: patch.shadow || null,
        rot: patch.rot || 0,
        opacity: patch.opacity == null ? 1 : patch.opacity,
        blend: patch.blend || null,
        mask: patch.mask || null,
        anim: patch.anim || 'none',
        keys: cx.keys(patch.keys),
    };
}

export const ICON = registerLayerType('icon', {
    paint: paintIcon,
    compile: compileIcon,
    bounds: iconBounds,
});
