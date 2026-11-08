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
import { normaliseTallLayouts } from './tall.js';

/** The kinds of picture an item can be, in the order the list groups what is not shown. */
export const PICTURE_KINDS = ['screen', 'window', 'camera', 'title', 'picture', 'web'];

/** The kinds of sound. */
export const SOUND_KINDS = ['mic', 'system'];

/** Where Place puts an item. */
export const PLACES = ['fill', 'centre', 'left', 'right', 'tl', 'tr', 'bl', 'br'];

export const SHAPES = ['rect', 'rounded', 'circle'];

/** The first line's two switches. */
export const VIDEO_MODES = ['on', 'off', 'placeholder'];
export const SOUND_MODES = ['on', 'off', 'music'];
/** How one shot gives way to the next: at once, or the old one fading out over the new. */
export const TRANSITIONS = ['cut', 'fade'];
/** How long a fade takes, in milliseconds. */
export const FADE_MS = 400;
/** The seconds a timed rotation may hold each shot for. */
export const ROTATE_EVERY = [3, 5, 8, 10, 15, 30];
/** Auto: nothing, Follow the speaker, or a timed rotation (the one menu in place of two buttons, A3). */
export const AUTO_MODES = ['off', 'follow', 'rotate'];
/** The key of the Back soon scene, the last of the scenes on Live: the placeholder and its music. */
export const BACK_SOON = 'back-soon';
/** The key of the Recorder's Sound only scene: the picture off, the sound recorded alone. */
export const SOUND_ONLY = 'sound-only';

const MARGIN = 0.03;
const CORNER_W = 0.24;
const FRAME_ASPECT = 16 / 9;
// Scenes are as many as a person wants; this only guards against a damaged file.
const MAX_SCENES = 500;

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
        label: '',
        on: false,
        live: true,
        keep: true,
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
        // The person's own name for it; `name` is what the source calls itself, which finds it again (A4.1).
        label: str(raw.label, '', 200),
        on: Boolean(raw.on),
        live: raw.live === undefined ? true : Boolean(raw.live),
        // Keep its shape: sizing keeps its proportions unless this is off.
        keep: raw.keep === undefined ? true : Boolean(raw.keep),
        box: clampBox(raw.box || b.box),
        crop: clampCrop(raw.crop),
        shape: pick(raw.shape, SHAPES, b.shape),
        mirror: raw.mirror === undefined ? b.mirror : Boolean(raw.mirror),
        text: str(raw.text, '', 120),
        subtitle: str(raw.subtitle, '', 120),
        path: str(raw.path, '', 2000),
        // A screen's size, so it is found again by what it is on another computer (A4.1).
        size: normaliseSize(raw.size),
        // Standing in for another screen until one is chosen: which it was ("Was Screen 2 (2560 × 1440)").
        was: str(raw.was, '', 200),
        ...(raw.kind === 'web' ? { sound: Boolean(raw.sound) } : {}),
    };
}

function normaliseSize(raw) {
    const w = Math.round(Number(raw?.w));
    const h = Math.round(Number(raw?.h));
    return w > 0 && h > 0 ? { w, h } : null;
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
    return { version: 3, items: [], sounds: [], scenes: [], video: 'on', sound: 'on', placeholder: '', placeholderText: '', music: '', follow: {}, rotate: { every: 8, keys: [] }, auto: 'off', transition: 'fade', previewFirst: false, cameraOff: false, tall: {}, scene: '' };
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
    const scenes = normaliseScenes(r.scenes);
    // A stack from before scenes had keys named a person's own scene by its place: own-0, own-1.
    const ownKey = (k) => { const m = /^own-(\d+)$/.exec(k); return m && scenes[Number(m[1])] ? scenes[Number(m[1])].key : k; };
    const follow = Object.fromEntries(Object.entries(normaliseFollow(r.follow)).map(([k, v]) => [ownKey(k), v]));
    return {
        version: 3,
        items,
        sounds,
        scenes,
        // The scene on the frame, when one is.
        scene: str(r.scene, '', 80),
        video: pick(r.video, VIDEO_MODES, 'on'),
        sound: pick(r.sound, SOUND_MODES, 'on'),
        placeholder: str(r.placeholder, '', 2000),
        placeholderText: str(r.placeholderText, '', 120),
        music: str(r.music, '', 2000),
        // Live Broadcast's switching: which microphone each shot follows, the
        // timed rotation's pace and shots, and how one shot gives way to the next.
        follow,
        rotate: {
            every: pick(Number(r.rotate?.every), ROTATE_EVERY, 8),
            keys: (Array.isArray(r.rotate?.keys) ? r.rotate.keys : []).filter((k) => typeof k === 'string' && k).slice(0, 20).map((k) => ownKey(k.slice(0, 80))),
        },
        auto: pick(r.auto, AUTO_MODES, 'off'),
        transition: pick(r.transition, TRANSITIONS, 'fade'),
        // Preview first: a scene clicked waits in Next until Take (A3); straight on air is the default.
        previewFirst: Boolean(r.previewFirst),
        // The Camera button on the bar: off, no camera is opened or drawn, in any scene.
        cameraOff: Boolean(r.cameraOff),
        // Each scene's tall layout, when it was changed by hand (tall.js).
        tall: normaliseTallLayouts(r.tall),
    };
}

/** Each shot's microphone, by the shot's key and the sound's id; at most twenty. */
function normaliseFollow(raw) {
    const out = {};
    if (!raw || typeof raw !== 'object') return out;
    for (const [k, v] of Object.entries(raw).slice(0, 20)) if (typeof v === 'string' && v) out[k.slice(0, 80)] = v.slice(0, 40);
    return out;
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
    return { key: str(raw.key, '', 80), name: str(raw.name, '', 60), layers };
}

/** Stored scenes as scenes, each with a key of its own. */
function normaliseScenes(raw) {
    const scenes = (Array.isArray(raw) ? raw : []).map(normaliseScene).filter(Boolean).slice(0, MAX_SCENES);
    const taken = new Set();
    for (const sc of scenes) {
        if (sc.key && sc.key !== BACK_SOON && sc.key !== SOUND_ONLY && !taken.has(sc.key)) { taken.add(sc.key); continue; }
        sc.key = '';
    }
    for (const sc of scenes) {
        if (sc.key) continue;
        sc.key = newSceneKey(taken);
        taken.add(sc.key);
    }
    return scenes;
}

/** A new scene key, not among `taken`. */
export function newSceneKey(taken = []) {
    const used = taken instanceof Set ? taken : new Set(taken);
    for (let n = 1; ; n++) {
        const key = `scene-${n}`;
        if (!used.has(key)) return key;
    }
}

/** The shown items' places kept as a scene, top first. */
export function sceneOf(stack, name) {
    return { key: '', name: str(name, '', 60), layers: stack.items.filter((it) => it.on).map((it) => ({ id: it.id, box: it.box, crop: it.crop, shape: it.shape })) };
}

/** A scene kept: what the frame shows, under a name, last among the scenes. */
export function saveScene(stack, name) {
    if (stack.scenes.length >= MAX_SCENES) return stack;
    const key = newSceneKey(stack.scenes.map((sc) => sc.key));
    return { ...stack, scenes: [...stack.scenes, { ...sceneOf(stack, name), key }], scene: key };
}

/** A scene by its key (or its place, from before scenes had keys) taken out. */
export function removeScene(stack, which) {
    return { ...stack, scenes: stack.scenes.filter((sc, i) => (typeof which === 'number' ? i !== which : sc.key !== which)) };
}

export function renameScene(stack, key, name) {
    return { ...stack, scenes: stack.scenes.map((sc) => (sc.key === key ? { ...sc, name: str(name, sc.name, 60) } : sc)) };
}

/** A copy of a scene, just after it, under `name`. */
export function duplicateScene(stack, key, name) {
    const at = stack.scenes.findIndex((sc) => sc.key === key);
    if (at < 0 || stack.scenes.length >= MAX_SCENES) return stack;
    const copy = { ...stack.scenes[at], key: newSceneKey(stack.scenes.map((sc) => sc.key)), name: str(name, stack.scenes[at].name, 60), layers: stack.scenes[at].layers.map((l) => ({ ...l })) };
    const scenes = [...stack.scenes];
    scenes.splice(at + 1, 0, copy);
    return { ...stack, scenes };
}

/** A scene moved to the gap `to` (0 to the number of scenes), as a drop gives it. */
export function moveScene(stack, key, to) {
    const from = stack.scenes.findIndex((sc) => sc.key === key);
    if (from < 0) return stack;
    const scenes = [...stack.scenes];
    const [sc] = scenes.splice(from, 1);
    scenes.splice(Math.max(0, Math.min(scenes.length, from < to ? to - 1 : to)), 0, sc);
    return { ...stack, scenes };
}

/** A scene made again from what the frame shows now: Update from the frame. */
export function updateSceneFromFrame(stack, key) {
    return { ...stack, scenes: stack.scenes.map((sc) => (sc.key === key ? { ...sc, layers: sceneOf(stack, sc.name).layers } : sc)) };
}

/**
 * A scene put on the share: its items shown in its order and places, every
 * other item hidden. An item the scene names that is gone is left out.
 */
export function applyScene(stack, scene) {
    if (!scene) return stack;
    if (scene.key === BACK_SOON) return { ...stack, video: 'placeholder', sound: 'music', scene: scene.key };
    if (scene.key === SOUND_ONLY) return { ...stack, video: 'off', sound: 'on', scene: scene.key };
    const byId = new Map(stack.items.map((it) => [it.id, it]));
    const shown = [];
    for (const l of scene.layers) {
        const it = byId.get(l.id);
        if (!it) continue;
        shown.push(normaliseItem({ ...it, on: true, box: l.box, crop: l.crop, shape: l.shape }));
        byId.delete(l.id);
    }
    const hidden = stack.items.filter((it) => byId.has(it.id)).map((it) => ({ ...it, on: false }));
    // Any scene but Back soon or Sound only brings the picture and the sounds back.
    return { ...stack, items: [...shown, ...hidden], video: 'on', sound: 'on', scene: scene.key || stack.scene };
}

/** What the frame shows written into the scene on it, so a change in Set up is a change to the scene being made. */
export function followFrame(stack) {
    if (!stack.scene || !stack.scenes.some((sc) => sc.key === stack.scene)) return stack;
    return updateSceneFromFrame(stack, stack.scene);
}

/**
 * A box being dragged, pulled to the frame's edges and centre lines and the
 * other sources' edges and centres when it is close to them. Answers the box
 * and the guides to draw: [{ axis: 'x'|'y', at }].
 */
export function snapMove(box, others = [], { threshold = 0.012 } = {}) {
    const xs = [0, 0.5, 1];
    const ys = [0, 0.5, 1];
    for (const o of others) { xs.push(o.x, o.x + o.w / 2, o.x + o.w); ys.push(o.y, o.y + o.h / 2, o.y + o.h); }
    const best = (edges, lines) => {
        let hit = null;
        for (const e of edges) for (const l of lines) { const d = l - e; if (Math.abs(d) <= threshold && (!hit || Math.abs(d) < Math.abs(hit.d))) hit = { d, at: l }; }
        return hit;
    };
    const hx = best([box.x, box.x + box.w / 2, box.x + box.w], xs);
    const hy = best([box.y, box.y + box.h / 2, box.y + box.h], ys);
    const guides = [];
    if (hx) guides.push({ axis: 'x', at: hx.at });
    if (hy) guides.push({ axis: 'y', at: hy.at });
    return { box: clampBox({ ...box, x: box.x + (hx?.d || 0), y: box.y + (hy?.d || 0) }), guides };
}

/** A corner being dragged, pulled to the same lines. Answers { x, y, guides }. */
export function snapPoint(x, y, others = [], { threshold = 0.012 } = {}) {
    const xs = [0, 0.5, 1];
    const ys = [0, 0.5, 1];
    for (const o of others) { xs.push(o.x, o.x + o.w / 2, o.x + o.w); ys.push(o.y, o.y + o.h / 2, o.y + o.h); }
    const near = (v, lines) => lines.reduce((b, l) => (Math.abs(l - v) <= threshold && (b == null || Math.abs(l - v) < Math.abs(b - v)) ? l : b), null);
    const sx = near(x, xs);
    const sy = near(y, ys);
    const guides = [];
    if (sx != null) guides.push({ axis: 'x', at: sx });
    if (sy != null) guides.push({ axis: 'y', at: sy });
    return { x: sx ?? x, y: sy ?? y, guides };
}

/** The Back soon scene, the last on Live: the placeholder's picture or card, and its music. */
export function backSoonScene(names = {}) {
    return { key: BACK_SOON, name: names.backSoon || 'Back soon', layers: [], back: true };
}

/** The Recorder's Sound only scene: no picture, the sounds recorded alone. */
export function soundOnlyScene(names = {}) {
    return { key: SOUND_ONLY, name: names.soundOnly || 'Sound only', layers: [], soundOnly: true };
}

/**
 * The quick scenes of a new setup, ordinary scenes from then on: Screen, A
 * window, Screen and camera, and Camera, from the first screen, window and
 * camera the stack has. One that needs what there is not is left out.
 */
export function quickScenes(stack, names = {}) {
    const screen = stack.items.find((it) => it.kind === 'screen');
    const win = stack.items.find((it) => it.kind === 'window');
    const camera = stack.items.find((it) => it.kind === 'camera');
    const layer = (it, box, shape) => ({ id: it.id, box, crop: null, shape });
    const out = [];
    if (screen) out.push({ key: 'screen', name: names.screen || 'Screen', layers: [layer(screen, placeBox('fill'), 'rect')] });
    if (win) out.push({ key: 'window', name: names.window || 'A window', layers: [layer(win, placeBox('fill'), 'rect')] });
    if (screen && camera) out.push({ key: 'screen-camera', name: names['screen-camera'] || 'Screen and camera', layers: [layer(camera, placeBox('br', { shape: 'circle' }), 'circle'), layer(screen, placeBox('fill'), 'rect')] });
    if (camera) out.push({ key: 'camera', name: names.camera || 'Camera', layers: [layer(camera, placeBox('fill'), 'rect')] });
    return out;
}

/**
 * The stack of a new setup: the main screen, a window yet to be chosen, the
 * first camera, the microphone and the computer's sound, with the four quick
 * scenes and Screen and camera up (or Screen, with no camera).
 */
export function newSetupStack({ screen = null, hasCamera = true, names = {} } = {}) {
    let s = emptyStack();
    ({ stack: s } = addItem(s, { kind: 'screen', ref: screen?.id || '', name: screen?.name || '', on: false, size: screen?.display?.width ? { w: screen.display.width, h: screen.display.height } : null, box: placeBox('fill') }));
    ({ stack: s } = addItem(s, { kind: 'window', ref: '', name: '', on: false, box: placeBox('fill') }));
    if (hasCamera) ({ stack: s } = addItem(s, { kind: 'camera', ref: '', name: '', on: false, box: placeBox('br', { shape: 'circle' }), shape: 'circle' }));
    s.sounds = [normaliseSound({ id: 'mic-1', kind: 'mic', live: true }), normaliseSound({ id: 'system-1', kind: 'system', live: false })];
    s.scenes = quickScenes(s, names);
    return normaliseStack(applyScene(s, s.scenes.find((sc) => sc.key === 'screen-camera') || s.scenes[0]));
}

/**
 * The scenes that used to be made up on the spot (Screen and camera, Side by
 * side, Camera 1, Everyone), turned into ordinary scenes in front of the
 * person's own, for a stack carried over from before setups (A4.2).
 */
export function materialiseScenes(stack, names = {}) {
    const made = builtInScenes(stack, names).map((sc) => ({ key: sc.key, name: sc.name, layers: sc.layers }));
    const taken = new Set(made.map((sc) => sc.key));
    return normaliseStack({ ...stack, scenes: [...made, ...stack.scenes.filter((sc) => !taken.has(sc.key))] });
}

/**
 * The ready-made scenes, made from the first screen or window and the
 * cameras there are: screen and camera, screen alone, side by side, and the
 * camera alone. With two cameras or more (a front and a side, or one on
 * each speaker), each camera has a shot of its own, keyed `cam:<id>`, and
 * the wide shot puts them all side by side. A scene that needs what there
 * is not is left out. `names.cameraN(index, item)` names a camera's shot.
 */
export function builtInScenes(stack, names = {}) {
    const screen = stack.items.find((it) => it.kind === 'screen' || it.kind === 'window');
    // The cameras in the order they were added (camera-1 first), whatever their order on the share.
    const cameras = stack.items.filter((it) => it.kind === 'camera').sort((x, y) => x.id.localeCompare(y.id, undefined, { numeric: true }));
    const camera = cameras[0];
    const overlays = stack.items.filter((it) => (it.kind === 'title' || it.kind === 'picture') && it.on).map((it) => ({ id: it.id, box: it.box, crop: it.crop, shape: it.shape }));
    const out = [];
    const layer = (it, box, shape = it.shape) => ({ id: it.id, box, crop: it.crop, shape });
    if (screen && camera) out.push({ key: 'screen-camera', name: names['screen-camera'] || 'Screen and camera', layers: [...overlays, layer(camera, placeBox('br', { shape: 'circle' }), 'circle'), layer(screen, placeBox('fill'), 'rect')] });
    if (screen) out.push({ key: 'screen', name: names.screen || 'Screen', layers: [...overlays, layer(screen, placeBox('fill'), 'rect')] });
    if (screen && camera) out.push({ key: 'split', name: names.split || 'Side by side', layers: [...overlays, layer(camera, placeBox('right', { aspect: 9 / 16 }), 'rounded'), layer(screen, placeBox('left'), 'rounded')] });
    if (cameras.length === 1) out.push({ key: 'camera', name: names.camera || 'Camera', layers: [...overlays, layer(camera, placeBox('fill'), 'rect')] });
    if (cameras.length > 1) {
        cameras.forEach((cam, i) => out.push({ key: `cam:${cam.id}`, name: names.cameraN ? names.cameraN(i, cam) : `Camera ${i + 1}`, layers: [...overlays, layer(cam, placeBox('fill'), 'rect')] }));
        out.push({ key: 'wide', name: names.wide || 'Everyone', layers: [...overlays, ...wideLayers(cameras).map((b, i) => layer(cameras[i], b, 'rounded'))] });
    }
    return out;
}

/** The wide shot: up to four cameras side by side across the frame, each filled by its picture. */
function wideLayers(cameras) {
    const n = Math.min(4, cameras.length);
    const gap = MARGIN / 2;
    const w = (1 - gap * (n + 1)) / n;
    // As tall as a 9:16 picture would be at that width, within the frame's height.
    const h = Math.min(1 - 2 * gap * FRAME_ASPECT, (w * FRAME_ASPECT) / (9 / 16) > 1 ? 1 - 2 * gap * FRAME_ASPECT : (w * FRAME_ASPECT) / (9 / 16));
    return cameras.slice(0, n).map((_, i) => clampBox({ x: gap + i * (w + gap), y: (1 - h) / 2, w, h }));
}

/* ── following the speaker ─────────────────────────────────────────────── */

/**
 * The timings Follow the speaker keeps (D-2026-11-03-02): a voice must go on
 * for `speakMs` before its shot is cut to, a shot stays on for `holdMs` at
 * least, and after `quietMs` with nobody speaking (or with more than one)
 * the wide shot is cut to.
 */
export const FOLLOW = Object.freeze({ speakMs: 1500, holdMs: 3000, quietMs: 2500 });

/**
 * One step of Follow the speaker. `state` is what the last step returned
 * (or nothing); `speaking` is the ids of the sounds heard just now; `pairs`
 * is each shot's microphone (stack.follow); `wide` is the shot to go to when
 * nobody, or more than one, speaks (null to stay where it is). Answers the
 * state to keep and the shot to cut to now, or null.
 */
export function followStep(state, { now, speaking = [], pairs = {}, wide = null, on = null, timing = FOLLOW } = {}) {
    const s = { on: state?.on ?? on, cand: state?.cand ?? null, candSince: state?.candSince ?? now, switchedAt: state?.switchedAt ?? -Infinity, quietSince: state?.quietSince ?? null };
    const heard = new Set(speaking);
    const keys = Object.keys(pairs).filter((k) => heard.has(pairs[k]));
    let target = null;
    let wait = timing.speakMs;
    if (keys.length === 1) {
        target = keys[0];
        s.quietSince = null;
    } else if (keys.length > 1) {
        target = wide;
        s.quietSince = null;
    } else {
        if (s.quietSince == null) s.quietSince = now;
        if (now - s.quietSince >= timing.quietMs) { target = wide; wait = 0; }
    }
    // A pause between words, not yet long enough to go wide, keeps whoever was about to be shown waiting.
    if (!target) return { state: s, cut: null };
    if (target === s.on) { s.cand = null; return { state: s, cut: null }; }
    if (s.cand !== target) { s.cand = target; s.candSince = now; }
    if (now - s.candSince < wait || now - s.switchedAt < timing.holdMs) return { state: s, cut: null };
    s.on = target;
    s.cand = null;
    s.switchedAt = now;
    return { state: s, cut: target };
}

/**
 * The shot after `current` in a timed rotation: through `keys` in order, back
 * to the first after the last; the first when `current` is not among them.
 */
export function nextInRotation(keys, current) {
    if (!keys.length) return null;
    const i = keys.indexOf(current);
    return keys[(i + 1) % keys.length];
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

/**
 * Each source as a file of its own (A7): what the frame shows now, screens and
 * windows first, each with its role on Video's timeline. One take carries the
 * sound: the camera's when there is one, as a voice belongs with a face; with
 * no picture at all, the sound is the one take.
 * @returns {{ role: 'screen'|'camera'|'voice', itemId: string|null, sound: boolean }[]}
 */
export function takesPlan(stack) {
  const shown = stack.items.filter((it) => it.on && it.live && (it.kind === 'camera' || ((it.kind === 'screen' || it.kind === 'window') && it.ref)));
  const cameras = shown.filter((it) => it.kind === 'camera');
  const screens = shown.filter((it) => it.kind !== 'camera');
  if (!cameras.length && !screens.length) return [{ role: 'voice', itemId: null, sound: true }];
  const carrier = cameras[0] || screens[0];
  return [...screens, ...cameras].map((it) => ({ role: it.kind === 'camera' ? 'camera' : 'screen', itemId: it.id, sound: it === carrier }));
}
