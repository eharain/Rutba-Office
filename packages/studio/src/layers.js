/**
 * An image project: the studio document as a stack of layers on a canvas of
 * pixels (the Image app's advanced editor, docs/STUDIO.md 7.2).
 *
 * There is no second model, as for video (D1): an image project IS a studio
 * document (doc.js) of `kind: 'image'`, its `size` the canvas in pixels and
 * its patches the layers, painted bottom first in the order of the list. A
 * picture is an `image` layer, text a `textblock`, word art, shapes, icons and
 * charts are the toolset's own types, and an adjustment layer (`adjust`)
 * changes the pixels already painted beneath it. Two things are added here
 * and nowhere else:
 *
 * - **Groups.** A group is a patch of its own, `{ type: 'group' }`, with a
 *   name, visibility, opacity and a lock; its members name it as `parent`
 *   and sit in one run directly beneath it in the list. The renderer never
 *   sees a group: `imagePatches` folds a group's visibility, opacity and lock
 *   into its members and leaves the group out. (`parent`, not the envelope's
 *   `group`, which links a clip's picture and sound in a video project.)
 * - **The canvas.** Every layer's geometry is a fraction of the canvas, so a
 *   new image size is only a new `size`; a new canvas size (with an anchor)
 *   or a trim keeps every layer where it was in pixels, so its fractions are
 *   worked out again (`reframe`).
 *
 * Everything here is pure arithmetic over the document: no DOM, no canvas,
 * and every function answers a NEW document, so undo is a snapshot. Masks,
 * selections and brushes (I10 to I14) will add fields to a layer's patch;
 * nothing here assumes a layer has only the fields it knows, and `reframe`
 * carries a mask's fractions along with the layer's own.
 */

import { emptyDoc, normaliseDoc, nextId } from './doc.js';

/** The canvas compositing modes a layer can blend with, in the order a menu offers them. */
export const BLEND_MODES = [
    'normal', 'multiply', 'screen', 'overlay', 'darken', 'lighten',
    'color-dodge', 'color-burn', 'hard-light', 'soft-light',
    'difference', 'exclusion', 'hue', 'saturation', 'color', 'luminosity',
];

/** The nine places a canvas can grow or shrink from. */
export const ANCHORS = ['nw', 'n', 'ne', 'w', 'c', 'e', 'sw', 's', 'se'];

/** The largest canvas side the editor makes (a browser canvas stops drawing past about 16k). */
export const MAX_SIDE = 16384;

const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));
const round4 = (n) => +Number(n).toFixed(4);
const side = (n) => clamp(Math.round(Number(n) || 1), 1, MAX_SIDE);

// ── the document ─────────────────────────────────────────────────────────────

/** A blank image project of `width` × `height` pixels, transparent. */
export function emptyImage({ width = 1200, height = 800 } = {}) {
    const doc = emptyDoc('image');
    return {
        ...doc,
        size: { width: side(width), height: side(height) },
        options: { ...doc.options, bare: true, background: 'transparent' },
    };
}

/** Anything handed to us, as an image project. */
export function normaliseImage(raw) {
    const doc = normaliseDoc(raw, 'image');
    const size = raw && raw.size && typeof raw.size === 'object' ? raw.size : { width: 1200, height: 800 };
    return {
        ...doc,
        kind: 'image',
        size: { width: side(size.width), height: side(size.height) },
        options: { ...doc.options, bare: true, background: doc.options?.background || 'transparent' },
    };
}

/** The canvas size, as numbers. */
export const canvasOf = (doc) => ({ width: side(doc?.size?.width), height: side(doc?.size?.height) });

/**
 * A picture as a layer, at one picture pixel to one canvas pixel: placed at
 * the top left by default, or centred. `width` and `height` are the
 * picture's own pixels.
 */
export function pictureLayer(doc, { url, width, height, label = 'Picture', centre = false, id = null }) {
    const { width: W, height: H } = canvasOf(doc);
    const fw = width / W;
    const x = centre ? (W - width) / 2 : 0;
    const y = centre ? (H - height) / 2 : 0;
    return {
        id: id || nextId(doc, 'image'),
        type: 'image',
        label,
        url,
        fx: round4(x / W),
        fy: round4(y / H),
        fw: round4(fw),
        shadow: false,
    };
}

// ── the stack: what is a unit, and who is whose ──────────────────────────────

export const isGroup = (p) => p?.type === 'group';
const indexOf = (doc, id) => doc.patches.findIndex((p) => p.id === id);

/**
 * The run of the list one layer stands for: itself, or a group with its
 * members (which sit directly beneath it). Answers [first, last] indices.
 */
export function unitOf(doc, id) {
    const at = indexOf(doc, id);
    if (at < 0) return null;
    if (!isGroup(doc.patches[at])) return [at, at];
    let first = at;
    while (first > 0 && doc.patches[first - 1]?.parent === id) first -= 1;
    return [first, at];
}

/** The layers as the panel lists them: top first, a group's members under it, one deeper. */
export function layerRows(doc) {
    const rows = [];
    const list = doc.patches || [];
    for (let i = list.length - 1; i >= 0; i--) {
        const p = list[i];
        rows.push({ id: p.id, patch: p, depth: p.parent ? 1 : 0, parent: p.parent || null, group: isGroup(p) });
    }
    return rows;
}

/** A layer's name as a person sees it. */
export function layerLabel(p) {
    return String(p?.label || (p?.type === 'textblock' || p?.type === 'text' || p?.type === 'wordart' ? String(p.text || '').split('\n')[0].slice(0, 40) : '') || p?.id || '');
}

/** Pull a run out of the list: [rest, run]. */
function cut(patches, [first, last]) {
    return [[...patches.slice(0, first), ...patches.slice(last + 1)], patches.slice(first, last + 1)];
}

/**
 * Move a layer (with its members, for a group) so it lands above or below
 * another. Dropped beside a member of a group, a layer joins that group; a
 * group is never put inside another (it lands beside the group instead).
 * `where`: 'above', 'below', or 'into' (a group: its top member).
 */
export function dropLayer(doc, id, targetId, where = 'above') {
    if (id === targetId) return doc;
    const unit = unitOf(doc, id);
    const target = doc.patches.find((p) => p.id === targetId);
    if (!unit || !target) return doc;
    const moving = doc.patches[unit[1]];
    // A group dropped next to a member lands next to that member's group.
    let anchorId = targetId;
    let mode = where;
    let parent = null;
    if (mode === 'into' && !isGroup(target)) mode = 'above';
    if (mode === 'into') {
        if (isGroup(moving)) mode = 'above';
        else parent = target.id;
    } else if (target.parent) {
        if (isGroup(moving)) anchorId = target.parent;
        else parent = target.parent;
    }
    if (isGroup(moving) && doc.patches.find((p) => p.id === anchorId)?.parent === id) return doc;
    const [rest, run] = cut(doc.patches, unit);
    const placed = run.map((p) => (p.id === id && !isGroup(p) ? withParent(p, parent) : p));
    let at;
    if (mode === 'into') {
        at = rest.findIndex((p) => p.id === anchorId); // directly beneath the group patch = its top member
    } else {
        const anchor = unitOf({ patches: rest }, anchorId);
        if (!anchor) return doc;
        // Above the anchor's unit goes after its last index; below goes before its first.
        at = mode === 'below' ? anchor[0] : anchor[1] + 1;
    }
    const patches = [...rest.slice(0, at), ...placed, ...rest.slice(at)];
    return { ...doc, patches: stripZ(patches) };
}

function withParent(p, parent) {
    const next = { ...p };
    if (parent) next.parent = parent; else delete next.parent;
    return next;
}

const stripZ = (patches) => patches.map((p) => {
    if (p.z === undefined) return p;
    const { z, ...rest } = p;
    return rest;
});

/** The units beside a layer at its own level, bottom first: [first, last] each. */
function siblingUnits(doc, id) {
    const p = doc.patches.find((x) => x.id === id);
    const parent = p?.parent || null;
    const out = [];
    for (let i = 0; i < doc.patches.length; i++) {
        const q = doc.patches[i];
        if ((q.parent || null) !== parent) continue;
        out.push(unitOf(doc, q.id));
    }
    return out;
}

/**
 * Bring forward (+1) or send backward (-1) one place among the layers at
 * its own level. A group's top member brought forward leaves the group,
 * above it; its bottom member sent backward leaves it, below.
 */
export function stepLayer(doc, id, delta) {
    const p = doc.patches.find((x) => x.id === id);
    if (!p) return doc;
    const units = siblingUnits(doc, id);
    const mine = units.findIndex((u) => doc.patches[u[1]].id === id);
    const next = mine + (delta > 0 ? 1 : -1);
    if (next < 0 || next >= units.length) {
        if (!p.parent) return doc;
        return dropLayer(doc, id, p.parent, delta > 0 ? 'above' : 'below');
    }
    const other = doc.patches[units[next][1]].id;
    return dropLayer(doc, id, other, delta > 0 ? 'above' : 'below');
}

/** Bring to front (+1) or send to back (-1) of the layers at its own level. */
export function layerToEnd(doc, id, end) {
    const units = siblingUnits(doc, id);
    if (!units.length) return doc;
    const target = doc.patches[(end > 0 ? units[units.length - 1] : units[0])[1]].id;
    if (target === id) return doc;
    return dropLayer(doc, id, target, end > 0 ? 'above' : 'below');
}

/**
 * Group layers: they move together into one run under a new group, where
 * the topmost of them was. Groups among them are taken apart first (a group
 * holds layers, not groups). Answers { doc, id }.
 */
export function groupLayers(doc, ids, { label = 'Group' } = {}) {
    let next = doc;
    const wanted = [];
    for (const id of ids) {
        const p = next.patches.find((x) => x.id === id);
        if (!p) continue;
        if (isGroup(p)) {
            const members = next.patches.filter((x) => x.parent === id).map((x) => x.id);
            next = ungroupLayers(next, id);
            wanted.push(...members);
        } else wanted.push(id);
    }
    const set = new Set(wanted);
    if (!set.size) return { doc, id: null };
    const gid = nextId(next, 'group');
    const n = next.patches.filter(isGroup).length + 1;
    const members = next.patches.filter((p) => set.has(p.id)).map((p) => ({ ...withParent(p, gid) }));
    const top = Math.max(...next.patches.map((p, i) => (set.has(p.id) ? i : -1)));
    // Where the topmost member was, counted in what is left once they are taken out.
    let before = next.patches.slice(0, top + 1).filter((p) => !set.has(p.id)).length;
    const rest = next.patches.filter((p) => !set.has(p.id));
    // Never inside another group's run: past that group, where its members end.
    const prev = rest[before - 1];
    if (prev?.parent) {
        const owner = rest.findIndex((p) => p.id === prev.parent);
        if (owner >= before) before = owner + 1;
    }
    const group = { id: gid, type: 'group', label: `${label} ${n}`, visible: true, opacity: 1 };
    const patches = [...rest.slice(0, before), ...members, group, ...rest.slice(before)];
    return { doc: { ...next, patches: stripZ(dropEmptyGroups(patches)) }, id: gid };
}

/** Take a group apart: its members stay where they are, at the top level. */
export function ungroupLayers(doc, id) {
    const g = doc.patches.find((p) => p.id === id);
    if (!isGroup(g)) return doc;
    return {
        ...doc,
        patches: doc.patches.filter((p) => p.id !== id).map((p) => (p.parent === id ? withParent(p, null) : p)),
    };
}

/** A group left with no members has nothing to show and goes. */
function dropEmptyGroups(patches) {
    const used = new Set(patches.map((p) => p.parent).filter(Boolean));
    return patches.filter((p) => !isGroup(p) || used.has(p.id));
}

/** Delete layers; a group goes with its members. */
export function removeLayers(doc, ids) {
    const gone = new Set(ids);
    for (const p of doc.patches) if (p.parent && gone.has(p.parent)) gone.add(p.id);
    return { ...doc, patches: dropEmptyGroups(doc.patches.filter((p) => !gone.has(p.id))) };
}

/**
 * Duplicate a layer (a group with its members) directly above itself,
 * nudged a little so the copy can be seen. Answers { doc, id }.
 */
export function duplicateLayers(doc, id, { nudge = 0.02 } = {}) {
    const unit = unitOf(doc, id);
    if (!unit) return { doc, id: null };
    let work = doc;
    const run = doc.patches.slice(unit[0], unit[1] + 1);
    const map = new Map();
    for (const p of run) {
        const nid = nextId(work, p.type || 'layer');
        map.set(p.id, nid);
        work = { ...work, patches: [...work.patches, { id: nid }] };
    }
    const copies = run.map((p) => {
        const c = { ...p, id: map.get(p.id) };
        if (p.parent && map.has(p.parent)) c.parent = map.get(p.parent);
        if (!isGroup(p) && p.fx !== undefined) c.fx = round4(p.fx + nudge);
        if (!isGroup(p) && p.fy !== undefined) c.fy = round4(p.fy + nudge);
        if (p.label) c.label = `${p.label} 2`.replace(/ 2 2$/, ' 3');
        delete c.z;
        return c;
    });
    const patches = [...doc.patches.slice(0, unit[1] + 1), ...copies, ...doc.patches.slice(unit[1] + 1)];
    return { doc: { ...doc, patches }, id: map.get(id) };
}

/** Change one layer's own settings (visible, opacity, blend, locked, label, …). */
export function setLayer(doc, id, change) {
    return {
        ...doc,
        patches: doc.patches.map((p) => {
            if (p.id !== id) return p;
            const next = { ...p, ...change };
            for (const key of ['style', 'adjust', 'mask', 'fill', 'stroke']) {
                if (change[key] && typeof change[key] === 'object' && p[key] && typeof p[key] === 'object' && !Array.isArray(change[key])) {
                    next[key] = { ...p[key], ...change[key] };
                }
            }
            if (next.blend === 'normal') delete next.blend;
            return next;
        }),
    };
}

/** Whether a layer, or the group it is in, is locked. */
export function isLocked(doc, id) {
    const p = doc.patches.find((x) => x.id === id);
    if (!p) return false;
    if (p.locked) return true;
    const g = p.parent ? doc.patches.find((x) => x.id === p.parent) : null;
    return Boolean(g?.locked);
}

/**
 * The patches the renderer paints, in order: groups folded into their
 * members (a hidden group hides them, its opacity multiplies theirs, its
 * lock locks them) and left out themselves; a blend of `normal` is no blend.
 */
export function imagePatches(doc) {
    const groups = new Map(doc.patches.filter(isGroup).map((g) => [g.id, g]));
    const out = [];
    for (const p of doc.patches) {
        if (isGroup(p)) continue;
        const g = p.parent ? groups.get(p.parent) : null;
        const q = { ...p };
        delete q.z;
        if (g) {
            q.visible = p.visible !== false && g.visible !== false;
            const go = g.opacity == null ? 1 : Number(g.opacity);
            const po = p.opacity == null ? 1 : Number(p.opacity);
            if (go !== 1 || p.opacity != null) q.opacity = clamp(go * po, 0, 1);
            if (g.locked) q.locked = true;
        }
        if (!BLEND_MODES.includes(q.blend) || q.blend === 'normal') delete q.blend;
        out.push(q);
    }
    return out;
}

// ── the canvas ───────────────────────────────────────────────────────────────

/**
 * Where the old canvas's top-left corner lands in a new one of `next` size,
 * grown or shrunk from `anchor` ('nw', 'n', … 'se', 'c'). Answers the new
 * canvas's origin in the OLD canvas's pixels, which is what `reframe` takes.
 */
export function anchorOffset(old, next, anchor = 'c') {
    const a = ANCHORS.includes(anchor) ? anchor : 'c';
    const col = a.includes('w') ? 0 : a.includes('e') ? 1 : 0.5;
    const row = a.startsWith('n') ? 0 : a.startsWith('s') ? 1 : 0.5;
    return {
        x: Math.round((old.width - next.width) * col) || 0,
        y: Math.round((old.height - next.height) * row) || 0,
    };
}

/** Where a layer type puts itself when its patch names no position or size. */
const DEFAULT_AT = { text: [0.5, 0.1], qr: [0.72, 0.06] };
const DEFAULT_FW = { image: 0.16, shape: 0.3, qr: 0.22, icon: 0.16, chart: 0.7, textblock: 0.7 };
const DEFAULT_SIZE = { text: 0.035, wordart: 0.09, textblock: 0.035, chart: 0.022 };

/**
 * A new canvas: `rect` is its origin and size in the OLD canvas's pixels
 * (a crop, a trim, or a canvas grown or shrunk from an anchor). Every layer
 * keeps its place and its size in pixels; only the fractions change. An
 * adjustment with no region of its own still covers the whole canvas.
 */
export function reframe(doc, rect) {
    const { width: W, height: H } = canvasOf(doc);
    const W2 = side(rect.width);
    const H2 = side(rect.height);
    const dx = Number(rect.x) || 0;
    const dy = Number(rect.y) || 0;
    const fx = (v) => round4((v * W - dx) / W2);
    const fy = (v) => round4((v * H - dy) / H2);
    const sw = (v) => round4((v * W) / W2);
    const sh = (v) => round4((v * H) / H2);
    const patches = doc.patches.map((p) => {
        if (isGroup(p)) return p;
        const q = { ...p };
        const [ax, ay] = DEFAULT_AT[p.type] || [0.5, 0.5];
        const whole = p.type === 'adjust' && p.fw == null && p.fh == null;
        if (!whole) {
            q.fx = fx(p.fx ?? ax);
            q.fy = fy(p.fy ?? ay);
            const fw = p.fw ?? DEFAULT_FW[p.type];
            if (fw != null) q.fw = sw(fw);
            if (p.fh != null) q.fh = sh(p.fh);
            const size = p.sizeFrac ?? DEFAULT_SIZE[p.type];
            if (size != null) q.sizeFrac = sw(size);
        }
        if (p.mask && typeof p.mask === 'object') {
            const m = { ...p.mask };
            if (m.fx != null || m.shape) m.fx = fx(m.fx ?? 0.5);
            if (m.fy != null || m.shape) m.fy = fy(m.fy ?? 0.5);
            if (m.fw != null || m.shape) m.fw = sw(m.fw ?? 0.5);
            if (m.fh != null) m.fh = sh(m.fh);
            if (m.radius != null) m.radius = sw(m.radius);
            if (Array.isArray(m.points)) m.points = m.points.map(([px, py]) => [fx(px), fy(py)]);
            q.mask = m;
        }
        // A painted mask (I10) is a bitmap laid over the canvas from its top
        // left corner; it stays over the same pixels, its feather as long.
        if (p.paintMask && typeof p.paintMask === 'object') {
            const m = { ...p.paintMask };
            m.fx = fx(m.fx ?? 0);
            m.fy = fy(m.fy ?? 0);
            m.fw = sw(m.fw ?? 1);
            m.fh = sh(m.fh ?? 1);
            if (m.feather != null) m.feather = sw(m.feather);
            q.paintMask = m;
        }
        // A picture on four corners (I14) keeps each corner where it was.
        if (Array.isArray(p.corners)) q.corners = p.corners.map(([cx, cy]) => [fx(cx), fy(cy)]);
        return q;
    });
    return { ...doc, size: { width: W2, height: H2 }, patches };
}

/** Canvas size: the canvas made `width` × `height`, grown or cut from `anchor`, the layers kept where they are. */
export function canvasSize(doc, { width, height, anchor = 'c' }) {
    const old = canvasOf(doc);
    const next = { width: side(width), height: side(height) };
    const { x, y } = anchorOffset(old, next, anchor);
    return reframe(doc, { x, y, ...next });
}

/**
 * Image size: everything scaled to `width` × `height`. The geometry is
 * fractional, so this is only the size; text grows with the width.
 */
export function imageSize(doc, { width, height }) {
    return { ...doc, size: { width: side(width), height: side(height) } };
}

/**
 * The smallest rectangle holding every pixel more opaque than `alpha`, in an
 * RGBA buffer of `width` × `height`: { x, y, width, height }, or null when
 * every pixel is clear.
 */
export function trimBounds(data, width, height, { alpha = 0 } = {}) {
    let top = -1;
    let bottom = -1;
    let left = width;
    let right = -1;
    for (let y = 0; y < height; y++) {
        const row = y * width * 4;
        let first = -1;
        let last = -1;
        for (let x = 0; x < width; x++) {
            if (data[row + x * 4 + 3] > alpha) { first = x; break; }
        }
        if (first < 0) continue;
        for (let x = width - 1; x >= first; x--) {
            if (data[row + x * 4 + 3] > alpha) { last = x; break; }
        }
        if (top < 0) top = y;
        bottom = y;
        if (first < left) left = first;
        if (last > right) right = last;
    }
    if (top < 0) return null;
    return { x: left, y: top, width: right - left + 1, height: bottom - top + 1 };
}

/** Trim to `bounds` (from `trimBounds` on the flattened picture); nothing to trim answers the same document. */
export function trimTo(doc, bounds) {
    if (!bounds) return doc;
    const { width, height } = canvasOf(doc);
    if (bounds.x === 0 && bounds.y === 0 && bounds.width === width && bounds.height === height) return doc;
    return reframe(doc, bounds);
}

// ── a box dragged by a corner ────────────────────────────────────────────────

/**
 * A box `b` ({ x, y, w, h }, pixels) scaled by `k` from the corner opposite
 * the one held (`kind`: 'nw', 'ne', 'se', 'sw'): the new box, the far corner
 * where it was.
 */
export function scaleBox(b, kind, k) {
    const w = b.w * k;
    const h = b.h * k;
    const x = kind === 'se' || kind === 'ne' ? b.x : b.x + b.w - w;
    const y = kind === 'se' || kind === 'sw' ? b.y : b.y + b.h - h;
    return { x, y, w, h };
}

/**
 * The patch a corner drag writes on a layer drawn in a box centred on its
 * point (shapes, icons, charts, paragraphs, adjustments with a region): the
 * size multiplied by `k`, the opposite corner held. `layer` carries its
 * stored fractions; `b` is its box in canvas pixels.
 */
export function resizeCentred(canvas, layer, b, kind, k) {
    const box = scaleBox(b, kind, k);
    const patch = {
        id: layer.id,
        fx: round4((box.x + box.w / 2) / canvas.width),
        fy: round4((box.y + box.h / 2) / canvas.height),
    };
    if (layer.fw != null) patch.fw = round4(layer.fw * k);
    if (layer.fh != null) patch.fh = round4(layer.fh * k);
    if (layer.sizeFrac != null && (layer.type === 'textblock' || layer.type === 'chart')) patch.sizeFrac = round4(layer.sizeFrac * k);
    return patch;
}

/**
 * The angle, in degrees, a rotation handle held at `point` gives a layer
 * centred on `centre` (0 = upright, the handle above the centre); `snap`
 * rounds to 15 degrees.
 */
export function rotationAt(centre, point, { snap = false } = {}) {
    let deg = (Math.atan2(point.y - centre.y, point.x - centre.x) * 180) / Math.PI + 90;
    if (snap) deg = Math.round(deg / 15) * 15;
    deg = ((deg % 360) + 540) % 360 - 180;
    return Math.round(deg * 10) / 10;
}

// ── pixels: paint layers, painted masks, corners (I10 to I14) ────────────────

/**
 * A paint layer: a clear bitmap the size of the canvas, laid over all of it,
 * for the brush, the fill, the gradient, the clone stamp and the healing
 * brush to put pixels on. It is a picture layer like any other (`image`), so
 * it moves, fades, blends and masks the same; `kind: 'paint'` only names it.
 */
export function paintLayer(doc, { id = null, label = null, kind = 'paint' } = {}) {
    const lid = id || nextId(doc, kind === 'paint' ? 'paint' : kind);
    return { id: lid, type: 'image', kind, ...(label ? { label } : {}), url: `rimg:bitmaps/${lid}.png`, fx: 0, fy: 0, fw: 1, shadow: false };
}

/** Whether a layer is a picture whose own pixels can be painted (a picture or a paint layer). */
export const isPixelLayer = (p) => Boolean(p && p.type === 'image' && typeof p.url === 'string' && p.url);

/**
 * Put a new layer directly above another (in the same group as it), or on
 * top of everything when there is no such layer.
 */
export function insertAbove(doc, targetId, patch) {
    const at = doc.patches.findIndex((p) => p.id === targetId);
    if (at < 0) return { ...doc, patches: [...doc.patches, patch] };
    const target = doc.patches[at];
    const placed = { ...patch };
    if (target.parent && !isGroup(target)) placed.parent = target.parent; else delete placed.parent;
    const unit = unitOf(doc, targetId);
    const end = unit ? unit[1] : at;
    // Above a group means above the group's own row, outside it.
    const index = isGroup(target) ? end + 1 : at + 1;
    return { ...doc, patches: [...doc.patches.slice(0, index), placed, ...doc.patches.slice(index)] };
}

/**
 * A painted mask on a layer: its bitmap `url` laid over the whole canvas,
 * what lies beyond its edges shown (`outside: 1`) or hidden (0).
 */
export function addMask(doc, id, { url, outside = 1 } = {}) {
    return setLayer(doc, id, { paintMask: { url, fx: 0, fy: 0, fw: 1, fh: 1, outside: outside ? 1 : 0, feather: 0, invert: false, off: false } });
}

/** Change a layer's mask (feather, invert, off, …); a layer with none is unchanged. */
export function setMask(doc, id, change) {
    const p = doc.patches.find((x) => x.id === id);
    if (!p?.paintMask) return doc;
    return setLayer(doc, id, { paintMask: { ...p.paintMask, ...change } });
}

/** Take a layer's mask away. */
export function removeMask(doc, id) {
    return {
        ...doc,
        patches: doc.patches.map((p) => {
            if (p.id !== id || !p.paintMask) return p;
            const { paintMask, ...rest } = p;
            return rest;
        }),
    };
}

/**
 * A layer moved by (dx, dy), fractions of the canvas: its place, its
 * corners and its painted mask together, as the mask is linked to it.
 */
export function moveLayerBy(doc, id, dx, dy) {
    return {
        ...doc,
        patches: doc.patches.map((p) => {
            if (p.id !== id) return p;
            const q = { ...p };
            if (p.fx != null) q.fx = round4(p.fx + dx);
            if (p.fy != null) q.fy = round4(p.fy + dy);
            if (Array.isArray(p.corners)) q.corners = p.corners.map(([x, y]) => [round4(x + dx), round4(y + dy)]);
            if (p.paintMask) q.paintMask = { ...p.paintMask, fx: round4((p.paintMask.fx ?? 0) + dx), fy: round4((p.paintMask.fy ?? 0) + dy) };
            return q;
        }),
    };
}

/**
 * Where a picture layer lies on the canvas, in canvas pixels: its box
 * `{ x, y, w, h }` (top left, at one bitmap's shape), its turn, and its
 * corners when it has been transformed. `bitmap` is { width, height }.
 */
export function pictureBox(doc, patch, bitmap) {
    const { width: W, height: H } = canvasOf(doc);
    const fw = patch.fw ?? 0.16;
    const w = W * fw;
    const h = bitmap && bitmap.width ? (w * bitmap.height) / bitmap.width : w;
    const box = { x: W * (patch.fx ?? 0.5), y: H * (patch.fy ?? 0.5), w, h };
    const corners = Array.isArray(patch.corners) ? patch.corners.map(([x, y]) => [x * W, y * H]) : null;
    return { box, rot: corners ? 0 : patch.rot || 0, corners };
}

/**
 * The map from canvas pixels to a picture layer's own bitmap pixels, as a
 * 3 × 3 matrix (row major): what a brush stroke on the canvas goes through
 * to land on the layer. A turned layer is turned back; a layer on corners is
 * mapped back through them.
 */
export function canvasToBitmap(doc, patch, bitmap) {
    const { box, rot, corners } = pictureBox(doc, patch, bitmap);
    const bw = bitmap.width;
    const bh = bitmap.height;
    if (corners) {
        return homographyOf(corners, [[0, 0], [bw, 0], [bw, bh], [0, bh]]);
    }
    const cx = box.x + box.w / 2;
    const cy = box.y + box.h / 2;
    const a = (-(rot || 0) * Math.PI) / 180;
    const cos = Math.cos(a);
    const sin = Math.sin(a);
    const sx = bw / box.w;
    const sy = bh / box.h;
    // translate(-c) → rotate(-rot) → translate(box/2) → scale(bitmap/box)
    return [
        sx * cos, -sx * sin, sx * (-cx * cos + cy * sin + box.w / 2),
        sy * sin, sy * cos, sy * (-cx * sin - cy * cos + box.h / 2),
        0, 0, 1,
    ];
}

/** The 3 × 3 projective map from four points to four points (warp.js has the full set). */
function homographyOf(src, dst) {
    const A = [];
    const b = [];
    for (let i = 0; i < 4; i++) {
        const [x, y] = src[i];
        const [u, v] = dst[i];
        A.push([x, y, 1, 0, 0, 0, -u * x, -u * y]);
        b.push(u);
        A.push([0, 0, 0, x, y, 1, -v * x, -v * y]);
        b.push(v);
    }
    const n = 8;
    const M = A.map((row, i) => [...row, b[i]]);
    for (let c = 0; c < n; c++) {
        let p = c;
        for (let r = c + 1; r < n; r++) if (Math.abs(M[r][c]) > Math.abs(M[p][c])) p = r;
        if (Math.abs(M[p][c]) < 1e-12) return [1, 0, 0, 0, 1, 0, 0, 0, 1];
        [M[p], M[c]] = [M[c], M[p]];
        for (let r = 0; r < n; r++) {
            if (r === c) continue;
            const f = M[r][c] / M[c][c];
            for (let k = c; k <= n; k++) M[r][k] -= f * M[c][k];
        }
    }
    const h = M.map((row, i) => row[n] / row[i]);
    return [...h, 1];
}

/** A point through a 3 × 3 map. */
export function mapPoint(m, x, y) {
    const w = m[6] * x + m[7] * y + m[8] || 1e-12;
    return { x: (m[0] * x + m[1] * y + m[2]) / w, y: (m[3] * x + m[4] * y + m[5]) / w };
}

/**
 * The flat (affine) map nearest a 3 × 3 map about one point, as a canvas's
 * setTransform takes it: [a, b, c, d, e, f]. Exact for a map that has no
 * perspective; about a point of a warped layer, close enough for one dab.
 */
export function affineAt(m, x, y) {
    const p = mapPoint(m, x, y);
    const e = 0.5;
    const px = mapPoint(m, x + e, y);
    const py = mapPoint(m, x, y + e);
    const a = (px.x - p.x) / e;
    const b = (px.y - p.y) / e;
    const c = (py.x - p.x) / e;
    const d = (py.y - p.y) / e;
    return [a, b, c, d, p.x - a * x - c * y, p.y - b * x - d * y];
}
