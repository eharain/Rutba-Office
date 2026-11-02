// The Recorder's studio as a stack of items (owner, 2026-11-02: "the screen
// and cameras all should work exactly same way so no separate features and
// functions on top. i can clip screen i should be able to clip camera").
//
// Every picture the studio can show is an item: a screen, a window, a camera,
// a title or a picture. Each has the same parts: shown or not, started or
// stopped, a box on the frame, a clip of its source, a shape and a mirror.
// The order of the items is the order of the list, top first, and the order
// they are stacked on the frame. The sounds are items of their own beside
// them: the microphones and the computer's sound, each muted or not, with a
// level. A scene is the shown items' places, kept under a name.
//
// Boxes and clips are fractions (0 to 1) of the frame and of the source, so a
// stack holds at any output size. Nothing here touches a canvas or a stream,
// so the whole of it is tested in Node.

import { normaliseStudio } from './compositor.js';

/** The kinds of picture an item can be, in the order the list groups what is not shown. */
export const PICTURE_KINDS = ['screen', 'window', 'camera', 'title', 'picture'];

/** The kinds of sound. */
export const SOUND_KINDS = ['mic', 'system'];

/** Where Place puts an item. */
export const PLACES = ['fill', 'centre', 'left', 'right', 'tl', 'tr', 'bl', 'br'];

export const SHAPES = ['rect', 'rounded', 'circle'];

/** The first line's two switches. */
export const VIDEO_MODES = ['on', 'off', 'placeholder'];
export const SOUND_MODES = ['on', 'off', 'music'];

const MARGIN = 0.03;
const CORNER_W = 0.24;
const FRAME_ASPECT = 16 / 9;
const MAX_SCENES = 9;

const r4 = (n) => Math.round(n * 10000) / 10000;
const num = (v, lo, hi, fallback) => (Number.isFinite(Number(v)) ? Math.max(lo, Math.min(hi, Number(v))) : fallback);
const str = (v, fallback = '', max = 200) => (typeof v === 'string' ? v.slice(0, max) : fallback);
const pick = (v, list, fallback) => (list.includes(v) ? v : fallback);

/** A box (fractions of the frame), whole and inside it. */
export function clampBox(b) {
    const w = num(b?.w, 0.02, 1, 1);
    const h = num(b?.h, 0.02, 1, 1);
    return { x: r4(num(b?.x, 0, 1 - w, 0)), y: r4(num(b?.y, 0, 1 - h, 0)), w: r4(w), h: r4(h) };
}

/** A clip (fractions of the source), or null for the whole of it. */
export function clampCrop(c) {
    if (!c || typeof c !== 'object') return null;
    const w = num(c.w, 0.02, 1, 1);
    const h = num(c.h, 0.02, 1, 1);
    const crop = { x: r4(num(c.x, 0, 1 - w, 0)), y: r4(num(c.y, 0, 1 - h, 0)), w: r4(w), h: r4(h) };
    return crop.w >= 0.999 && crop.h >= 0.999 ? null : crop;
}

/**
 * The box Place gives an item whose picture is `aspect` wide to one high
 * (after its clip). A circle is as tall as it is wide on the frame. The
 * frame is 16 by 9, so a box's height in frame fractions is its width times
 * 16/9 over the picture's aspect.
 */
export function placeBox(place, { aspect = FRAME_ASPECT, shape = 'rect' } = {}) {
    const a = shape === 'circle' ? 1 : (aspect > 0 ? aspect : FRAME_ASPECT);
    const hOf = (w) => (w * FRAME_ASPECT) / a;
    const fit = (maxW, maxH) => {
        let w = maxW;
        let h = hOf(w);
        if (h > maxH) { h = maxH; w = (h * a) / FRAME_ASPECT; }
        return { w, h };
    };
    const my = MARGIN * FRAME_ASPECT;
    switch (place) {
        // A circle fills the frame's height, in the middle.
        case 'fill': { if (shape !== 'circle') return clampBox({ x: 0, y: 0, w: 1, h: 1 }); const w = 1 / FRAME_ASPECT; return clampBox({ x: (1 - w) / 2, y: 0, w, h: 1 }); }
        case 'centre': { const { w, h } = fit(0.6, 0.8); return clampBox({ x: (1 - w) / 2, y: (1 - h) / 2, w, h }); }
        case 'left': { const { w, h } = fit(0.5 - MARGIN * 1.5, 1 - 2 * my); return clampBox({ x: MARGIN, y: (1 - h) / 2, w, h }); }
        case 'right': { const { w, h } = fit(0.5 - MARGIN * 1.5, 1 - 2 * my); return clampBox({ x: 1 - MARGIN - w, y: (1 - h) / 2, w, h }); }
        case 'tl': case 'tr': case 'bl': case 'br': {
            const { w, h } = fit(CORNER_W, 0.45);
            const x = place[1] === 'l' ? MARGIN : 1 - MARGIN - w;
            const y = place[0] === 't' ? my : 1 - my - h;
            return clampBox({ x, y, w, h });
        }
        default: return clampBox({ x: 0, y: 0, w: 1, h: 1 });
    }
}

/** The band a title sits in, across the foot of the frame on the left. */
export const TITLE_BOX = Object.freeze({ x: 0.04, y: 0.81, w: 0.46, h: 0.13 });

/** What a new item of a kind starts as. */
function base(kind) {
    return {
        id: '',
        kind,
        ref: '',
        name: '',
        on: false,
        live: true,
        box: kind === 'title' ? { ...TITLE_BOX } : { x: 0, y: 0, w: 1, h: 1 },
        crop: null,
        shape: 'rect',
        mirror: kind === 'camera',
        text: '',
        subtitle: '',
        path: '',
    };
}

/** An item as stored, or anything, as an item; null when it cannot be one. */
export function normaliseItem(raw) {
    if (!raw || typeof raw !== 'object' || !PICTURE_KINDS.includes(raw.kind)) return null;
    const b = base(raw.kind);
    return {
        ...b,
        id: str(raw.id, '', 60),
        ref: str(raw.ref, '', 400),
        name: str(raw.name, '', 200),
        on: Boolean(raw.on),
        live: raw.live === undefined ? true : Boolean(raw.live),
        box: clampBox(raw.box || b.box),
        crop: clampCrop(raw.crop),
        shape: pick(raw.shape, SHAPES, b.shape),
        mirror: raw.mirror === undefined ? b.mirror : Boolean(raw.mirror),
        text: str(raw.text, '', 120),
        subtitle: str(raw.subtitle, '', 120),
        path: str(raw.path, '', 2000),
    };
}

/** A sound as stored, or anything, as a sound; null when it cannot be one. */
export function normaliseSound(raw) {
    if (!raw || typeof raw !== 'object' || !SOUND_KINDS.includes(raw.kind)) return null;
    return {
        id: str(raw.id, '', 60),
        kind: raw.kind,
        ref: str(raw.ref, '', 400),
        name: str(raw.name, '', 200),
        live: Boolean(raw.live),
        muted: Boolean(raw.muted),
        gain: r4(num(raw.gain, 0, 2, 1)),
    };
}

/** A new id, not among `taken`. */
export function newId(kind, taken = []) {
    const used = new Set(taken);
    for (let n = 1; ; n++) {
        const id = `${kind}-${n}`;
        if (!used.has(id)) return id;
    }
}

/** An empty studio. */
export function emptyStack() {
    return { version: 2, items: [], sounds: [], scenes: [], video: 'on', sound: 'on', placeholder: '', placeholderText: '', music: '' };
}

/** A stored stack, or anything, as a stack the studio can use. */
export function normaliseStack(raw) {
    const r = raw && typeof raw === 'object' ? raw : {};
    const items = [];
    for (const x of Array.isArray(r.items) ? r.items : []) {
        const it = normaliseItem(x);
        if (!it) continue;
        if (!it.id || items.some((o) => o.id === it.id)) it.id = newId(it.kind, items.map((o) => o.id));
        items.push(it);
    }
    const sounds = [];
    for (const x of Array.isArray(r.sounds) ? r.sounds : []) {
        const s = normaliseSound(x);
        if (!s) continue;
        if (!s.id || sounds.some((o) => o.id === s.id)) s.id = newId(s.kind, sounds.map((o) => o.id));
        sounds.push(s);
    }
    const scenes = (Array.isArray(r.scenes) ? r.scenes : []).map(normaliseScene).filter(Boolean).slice(0, MAX_SCENES);
    return {
        version: 2,
        items,
        sounds,
        scenes,
        video: pick(r.video, VIDEO_MODES, 'on'),
        sound: pick(r.sound, SOUND_MODES, 'on'),
        placeholder: str(r.placeholder, '', 2000),
        placeholderText: str(r.placeholderText, '', 120),
        music: str(r.music, '', 2000),
    };
}

/**
 * A stack from the settings the Recorder kept before it was a stack
 * (scene, camera corner and size, title, logo, microphone, computer sound),
 * so a person's Recorder opens as they left it. The screen or window is
 * not known from them: `screen` is the source to show, when there is one.
 */
export function stackFromSettings(settings, { screen = null } = {}) {
    const s = normaliseStudio(settings);
    const out = emptyStack();
    const add = (it) => { const item = { ...base(it.kind), ...it }; item.id = newId(item.kind, out.items.map((o) => o.id)); out.items.push(normaliseItem(item)); return item; };
    // Top first: logo, title, camera, then the screen at the bottom.
    if (s.logo) add({ kind: 'picture', path: s.logo, name: '', on: true, box: placeBox(s.logoCorner, { aspect: 1 }) });
    if (s.titleOn || s.title || s.subtitle) add({ kind: 'title', text: s.title, subtitle: s.subtitle, on: Boolean(s.titleOn) });
    const wantScreen = s.scene !== 'camera' && s.scene !== 'brb';
    if (s.camera) {
        let box;
        let shape = 'rect';
        if (s.scene === 'camera' || !screen) box = placeBox('fill');
        else if (s.scene === 'split') { box = placeBox('right', { aspect: 9 / 16 }); shape = 'rounded'; }
        else if (s.scene === 'screen') box = placeBox(s.corner, { shape: s.shape, aspect: 16 / 9 });
        else {
            shape = s.shape;
            const w = { small: 0.16, medium: 0.24, large: 0.34 }[s.insetSize] || CORNER_W;
            const corner = placeBox(s.corner, { shape, aspect: 16 / 9 });
            const k = w / corner.w;
            box = clampBox({ x: s.corner[1] === 'l' ? corner.x : corner.x + corner.w - corner.w * k, y: s.corner[0] === 't' ? corner.y : corner.y + corner.h - corner.h * k, w: corner.w * k, h: corner.h * k });
        }
        add({ kind: 'camera', ref: s.cameraId, on: s.scene !== 'screen', mirror: s.mirror, shape, box });
    }
    if (screen) add({ kind: screen.kind === 'window' ? 'window' : 'screen', ref: screen.id, name: screen.name || '', on: wantScreen, box: s.scene === 'split' ? placeBox('left') : placeBox('fill') });
    out.sounds.push(normaliseSound({ id: 'mic-1', kind: 'mic', ref: s.micId, live: s.mic }));
    out.sounds.push(normaliseSound({ id: 'system-1', kind: 'system', live: s.system }));
    if (s.scene === 'brb') { out.video = 'placeholder'; out.placeholderText = s.brb; }
    return out;
}

/** Shown items, bottom first: the order they are drawn in. */
export function drawOrder(stack) {
    return stack.items.filter((it) => it.on).reverse();
}

/** The stack with one item changed (a patch, or a function of the item). */
export function updateItem(stack, id, patch) {
    return {
        ...stack,
        items: stack.items.map((it) => {
            if (it.id !== id) return it;
            const next = typeof patch === 'function' ? patch(it) : { ...it, ...patch };
            return normaliseItem({ ...next, id }) || it;
        }),
    };
}

/** The stack with one sound changed. */
export function updateSound(stack, id, patch) {
    return { ...stack, sounds: stack.sounds.map((s) => (s.id === id ? normaliseSound({ ...s, ...patch, id }) || s : s)) };
}

/**
 * An item added: shown, on top, placed by its kind (a screen or window to
 * fill the frame, a camera in a corner, a title in its band, a picture in
 * the top right corner). Answers { stack, id }.
 */
export function addItem(stack, raw, { aspect = FRAME_ASPECT } = {}) {
    const kind = raw.kind;
    const b = base(kind);
    const id = newId(kind, stack.items.map((o) => o.id));
    const placed = raw.box ? raw.box
        : kind === 'camera' ? placeBox(stack.items.some((o) => o.on) ? 'br' : 'fill', { aspect, shape: raw.shape || 'rect' })
            : kind === 'picture' ? placeBox('tr', { aspect })
                : kind === 'title' ? b.box
                    : placeBox(stack.items.some((o) => o.on && (o.kind === 'screen' || o.kind === 'window')) ? 'centre' : 'fill', { aspect });
    const item = normaliseItem({ ...b, ...raw, id, on: raw.on ?? true, box: placed });
    // Shown, on top; not shown, at the top of the ones not shown.
    const items = item.on ? [item, ...stack.items] : [...stack.items.filter((o) => o.on), item, ...stack.items.filter((o) => !o.on)];
    return { stack: { ...stack, items }, id };
}

/** An item taken out of the studio altogether. */
export function removeItem(stack, id) {
    return { ...stack, items: stack.items.filter((it) => it.id !== id), scenes: stack.scenes.map((sc) => ({ ...sc, layers: sc.layers.filter((l) => l.id !== id) })) };
}

/** A second copy of an item, on top of it, to clip or shape another way. */
export function duplicateItem(stack, id) {
    const it = stack.items.find((o) => o.id === id);
    if (!it) return { stack, id: null };
    const copy = { ...it, id: newId(it.kind, stack.items.map((o) => o.id)), on: true, box: clampBox({ ...it.box, x: it.box.x + 0.03, y: it.box.y + 0.03 }) };
    const at = stack.items.indexOf(it);
    const items = [...stack.items];
    items.splice(at, 0, normaliseItem(copy));
    return { stack: { ...stack, items }, id: copy.id };
}

/**
 * An item moved in the list to `to` (an index of the list as it is now,
 * top first). Shown items above the not-shown ones keep the list honest:
 * an item moved among the shown ones is shown, among the others hidden.
 */
export function moveItem(stack, id, to) {
    const items = [...stack.items];
    const from = items.findIndex((it) => it.id === id);
    if (from < 0) return stack;
    const [it] = items.splice(from, 1);
    const at = Math.max(0, Math.min(items.length, to));
    const shownCount = items.filter((o) => o.on).length;
    items.splice(at, 0, { ...it, on: at <= shownCount && (it.on || at < shownCount) });
    return { ...stack, items: [...items.filter((o) => o.on), ...items.filter((o) => !o.on)] };
}

/** One step up (toward the top) or down. */
export function nudgeItem(stack, id, by) {
    const from = stack.items.findIndex((it) => it.id === id);
    if (from < 0) return stack;
    return moveItem(stack, id, from + by);
}

/** Shown or hidden: shown goes to the top of the shown ones, hidden to the top of the rest. */
export function showItem(stack, id, on) {
    const it = stack.items.find((o) => o.id === id);
    if (!it) return stack;
    const rest = stack.items.filter((o) => o.id !== id);
    const next = { ...it, on: Boolean(on) };
    const shown = rest.filter((o) => o.on);
    const hidden = rest.filter((o) => !o.on);
    return { ...stack, items: on ? [next, ...shown, ...hidden] : [...shown, next, ...hidden] };
}

/** Place: an item put by name, its box worked out from its picture's aspect after its clip. */
export function placeItem(stack, id, place, { aspect = FRAME_ASPECT } = {}) {
    return updateItem(stack, id, (it) => ({ ...it, box: placeBox(place, { aspect: clippedAspect(aspect, it.crop), shape: it.shape }) }));
}

/** A picture `aspect` wide to one high, clipped. */
export function clippedAspect(aspect, crop) {
    const a = aspect > 0 ? aspect : FRAME_ASPECT;
    return crop ? (a * crop.w) / crop.h : a;
}

/**
 * The shape changed. A circle's box is made square on the frame about its
 * middle; from a circle, the box takes the picture's aspect again.
 */
export function shapeItem(stack, id, shape, { aspect = FRAME_ASPECT } = {}) {
    return updateItem(stack, id, (it) => {
        if (!SHAPES.includes(shape)) return it;
        const b = it.box;
        const full = b.w >= 0.999 && b.h >= 0.999;
        if (full) return { ...it, shape };
        const a = shape === 'circle' ? 1 : clippedAspect(aspect, it.crop);
        const h = Math.min(1, (b.w * FRAME_ASPECT) / a);
        return { ...it, shape, box: clampBox({ x: b.x, y: b.y + (b.h - h) / 2, w: b.w, h }) };
    });
}

/**
 * The clip changed. The box keeps its width and takes the clipped picture's
 * aspect (a circle stays round), so nothing is squeezed.
 */
export function clipItem(stack, id, crop, { aspect = FRAME_ASPECT } = {}) {
    return updateItem(stack, id, (it) => {
        const c = clampCrop(crop);
        const b = it.box;
        const full = b.w >= 0.999 && b.h >= 0.999;
        if (full || it.shape === 'circle' || it.kind === 'title') return { ...it, crop: c };
        const h = Math.min(1, (b.w * FRAME_ASPECT) / clippedAspect(aspect, c));
        return { ...it, crop: c, box: clampBox({ x: b.x, y: b.y, w: b.w, h }) };
    });
}

/**
 * A point on the frame (fractions) inside an item's box, as a point on its
 * source (fractions of the whole source), through its clip. It follows how
 * the program draws the item: a rectangle shows the whole of its clip inside
 * the box, a circle or a rounded box is filled by the middle of it. `aspect`
 * is the source's width over its height; without it the clip is taken to
 * fill the box exactly.
 */
export function frameToSource(item, fx, fy, { aspect = null } = {}) {
    const b = item.box;
    const c = item.crop || { x: 0, y: 0, w: 1, h: 1 };
    // The box and the clip in one unit: the frame's height, and the source's.
    const bw = b.w * FRAME_ASPECT;
    const bh = b.h;
    const pw = aspect ? c.w * aspect : bw;
    const ph = aspect ? c.h : bh;
    const px = (fx - b.x) * FRAME_ASPECT;
    const py = fy - b.y;
    let u;
    let v;
    if (item.shape === 'rect') {
        const k = Math.min(bw / pw, bh / ph);
        const dw = pw * k;
        const dh = ph * k;
        u = (px - (bw - dw) / 2) / dw;
        v = (py - (bh - dh) / 2) / dh;
    } else {
        const k = Math.max(bw / pw, bh / ph);
        const vw = bw / k;
        const vh = bh / k;
        u = ((pw - vw) / 2 + (px / bw) * vw) / pw;
        v = ((ph - vh) / 2 + (py / bh) * vh) / ph;
    }
    u = Math.max(0, Math.min(1, u));
    v = Math.max(0, Math.min(1, v));
    const su = item.mirror ? 1 - u : u;
    return { x: c.x + su * c.w, y: c.y + v * c.h };
}

/** The shown item on top at a point of the frame, or null. */
export function itemAt(stack, fx, fy) {
    for (const it of stack.items) {
        if (!it.on) continue;
        const b = it.box;
        if (fx < b.x || fy < b.y || fx > b.x + b.w || fy > b.y + b.h) continue;
        if (it.shape === 'circle') {
            const dx = (fx - (b.x + b.w / 2)) / (b.w / 2);
            const dy = (fy - (b.y + b.h / 2)) / (b.h / 2);
            if (dx * dx + dy * dy > 1) continue;
        }
        return it;
    }
    return null;
}

/* ── scenes ────────────────────────────────────────────────────────────── */

function normaliseScene(raw) {
    if (!raw || typeof raw !== 'object' || !Array.isArray(raw.layers)) return null;
    const layers = raw.layers
        .filter((l) => l && typeof l === 'object' && typeof l.id === 'string')
        .map((l) => ({ id: l.id.slice(0, 60), box: clampBox(l.box), crop: clampCrop(l.crop), shape: pick(l.shape, SHAPES, 'rect') }));
    return { name: str(raw.name, '', 60), layers };
}

/** The shown items' places kept as a scene, top first. */
export function sceneOf(stack, name) {
    return { name: str(name, '', 60), layers: stack.items.filter((it) => it.on).map((it) => ({ id: it.id, box: it.box, crop: it.crop, shape: it.shape })) };
}

/** A scene kept, in its own place among the person's scenes; at most nine. */
export function saveScene(stack, name) {
    if (stack.scenes.length >= MAX_SCENES) return stack;
    return { ...stack, scenes: [...stack.scenes, sceneOf(stack, name)] };
}

export function removeScene(stack, index) {
    return { ...stack, scenes: stack.scenes.filter((_, i) => i !== index) };
}

/**
 * A scene put on the share: its items shown in its order and places, every
 * other item hidden. An item the scene names that is gone is left out.
 */
export function applyScene(stack, scene) {
    if (!scene) return stack;
    const byId = new Map(stack.items.map((it) => [it.id, it]));
    const shown = [];
    for (const l of scene.layers) {
        const it = byId.get(l.id);
        if (!it) continue;
        shown.push(normaliseItem({ ...it, on: true, box: l.box, crop: l.crop, shape: l.shape }));
        byId.delete(l.id);
    }
    const hidden = stack.items.filter((it) => byId.has(it.id)).map((it) => ({ ...it, on: false }));
    return { ...stack, items: [...shown, ...hidden] };
}

/**
 * The ready-made scenes, made from the first screen or window and the first
 * camera there are: screen and camera, screen alone, camera alone, side by
 * side. A scene that needs what there is not is left out.
 */
export function builtInScenes(stack, names = {}) {
    const screen = stack.items.find((it) => it.kind === 'screen' || it.kind === 'window');
    const camera = stack.items.find((it) => it.kind === 'camera');
    const overlays = stack.items.filter((it) => (it.kind === 'title' || it.kind === 'picture') && it.on).map((it) => ({ id: it.id, box: it.box, crop: it.crop, shape: it.shape }));
    const out = [];
    const layer = (it, box, shape = it.shape) => ({ id: it.id, box, crop: it.crop, shape });
    if (screen && camera) out.push({ key: 'screen-camera', name: names['screen-camera'] || 'Screen and camera', layers: [...overlays, layer(camera, placeBox('br', { shape: 'circle' }), 'circle'), layer(screen, placeBox('fill'), 'rect')] });
    if (screen) out.push({ key: 'screen', name: names.screen || 'Screen', layers: [...overlays, layer(screen, placeBox('fill'), 'rect')] });
    if (camera) out.push({ key: 'camera', name: names.camera || 'Camera', layers: [...overlays, layer(camera, placeBox('fill'), 'rect')] });
    if (screen && camera) out.push({ key: 'split', name: names.split || 'Side by side', layers: [...overlays, layer(camera, placeBox('right', { aspect: 9 / 16 }), 'rounded'), layer(screen, placeBox('left'), 'rounded')] });
    return out;
}

/* ── what Record makes ─────────────────────────────────────────────────── */

/**
 * What Record records with the first line's switches as they are: 'video'
 * (picture and sound, or picture and silence), 'sound' (sound alone, when
 * the picture is off), or null when both are off.
 */
export function recordKind(stack) {
    if (stack.video === 'off') return stack.sound === 'off' ? null : 'sound';
    return 'video';
}
