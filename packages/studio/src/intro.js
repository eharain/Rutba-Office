/**
 * Introductory videos (docs/STUDIO.md section 9, Phase 3): a project made of
 * scenes, recorded one at a time, in a brand of one's own.
 *
 * An introduction is a video project whose `intro` holds its scenes, in
 * order, and its brand kit:
 *
 *   intro: {
 *     scenes: [{ id, layout, script, title, subtitle, seconds, corner, shape,
 *                takes: [{ id, files, duration }], take, trim, zooms, look }],
 *     brand: { name, role, logo, colors, font, lowerThird, watermark },
 *     words: { placeholder, ... },   // the template's words, in the language it was made in
 *     prompter: { speed },          // words a minute
 *   }
 *
 * A scene is a fragment of the timeline: its layers carry `scene` (its id)
 * and `role` (ground, screen, pointer, camera, logo, band, title, subtitle,
 * the take's sounds), each role on a lane of its own, so that the scenes,
 * one after another, make one timeline of the usual kind that every editing
 * tool works on. A layout (X1) says where each role sits in the frame for
 * each of the frame's shapes; the screen and the camera are separate layers
 * (D6), so a layout can be switched after recording (X2), and an export to
 * another shape lays every scene out again for it.
 *
 * Every change to the scenes (a take recorded, a layout switched, a scene
 * moved, the brand changed) is `composeIntro`: what the timeline says about
 * each scene is read back first (its words, how its take was trimmed, its
 * zooms, its pointer's look, the camera's corner), the scenes' layers are
 * made again from the scenes, one after another from the start, and
 * everything that is not a scene's (captions, anything added) is left where
 * it is. Pure: documents in, documents out.
 */

import { nextId } from './doc.js';
import { emptyProject, clipsOf, clipById } from './project.js';
import { insetRect, cornerOf } from './recorder.js';
import { compactPointer, findZooms, zoomKeys, POINTER_LOOK } from './pointer.js';
import { layoutCaptions } from './captions.js';

const round = (t) => Math.round(t * 1000) / 1000;

// ── the template (X1) ───────────────────────────────────────────────────────

/** The layouts a scene may have. */
export const LAYOUTS = ['title', 'talking', 'screenCamera', 'sideBySide', 'end'];

/** The layouts that are recorded; the title and end cards are made of words. */
export const RECORDED = new Set(['talking', 'screenCamera', 'sideBySide']);

/** A new introduction's scenes. */
export const INTRO_TEMPLATE = ['title', 'talking', 'screenCamera', 'end'];

/** How long a scene runs before it has a take (or always, for a card). */
export const SCENE_SECONDS = { title: 4, talking: 6, screenCamera: 8, sideBySide: 8, end: 5 };

/** The lanes each role is drawn on, from the ground up. */
export const LANES = { ground: 0, screen: 1, pointer: 2, camera: 3, logo: 4, band: 5, title: 6, subtitle: 7, placeholder: 6 };
export const SOUND_LANES = { voice: 0, computer: 1 };

/** How long a scene fades in and out. */
export const SCENE_FADE = 0.35;

/** What a brand kit starts as. */
export const BRAND_DEFAULTS = Object.freeze({
    name: '',
    role: '',
    logo: null, // { url, name }
    colors: Object.freeze({ background: '#15171d', text: '#ffffff', accent: '#ffc107' }),
    font: 'Segoe UI',
    lowerThird: true,
    watermark: true,
});

/** The fonts a brand may use: ones every Windows has, and the stack after them for any script they lack. */
export const BRAND_FONTS = ['Segoe UI', 'Arial', 'Georgia', 'Verdana', 'Tahoma', 'Trebuchet MS', 'Times New Roman', 'Nirmala UI'];

const hex = (v, fallback) => (/^#[0-9a-f]{6}$/i.test(String(v || '')) ? String(v).toLowerCase() : fallback);

/** Anything stored as a brand kit, as one the scenes can use. */
export function normaliseBrand(raw) {
    const r = raw && typeof raw === 'object' ? raw : {};
    const d = BRAND_DEFAULTS;
    const c = r.colors && typeof r.colors === 'object' ? r.colors : {};
    return {
        name: typeof r.name === 'string' ? r.name.slice(0, 80) : '',
        role: typeof r.role === 'string' ? r.role.slice(0, 120) : '',
        logo: r.logo && typeof r.logo.url === 'string' && r.logo.url ? { url: r.logo.url, name: String(r.logo.name || '') } : null,
        colors: { background: hex(c.background, d.colors.background), text: hex(c.text, d.colors.text), accent: hex(c.accent, d.colors.accent) },
        font: typeof r.font === 'string' && r.font.trim() ? r.font.trim().slice(0, 60) : d.font,
        lowerThird: r.lowerThird === undefined ? d.lowerThird : Boolean(r.lowerThird),
        watermark: r.watermark === undefined ? d.watermark : Boolean(r.watermark),
    };
}

/** A fresh id for a scene or a take, unused in the list. */
function freshId(list, prefix) {
    const used = new Set((list || []).map((x) => x.id));
    let n = 1;
    while (used.has(`${prefix}${n}`)) n += 1;
    return `${prefix}${n}`;
}

/** A scene of a layout, with the template's words for a card. */
export function newScene(layout, scenes = [], words = {}) {
    const kind = LAYOUTS.includes(layout) ? layout : 'talking';
    return {
        id: freshId(scenes, 's'),
        layout: kind,
        script: '',
        title: kind === 'title' ? (words.title || '') : kind === 'end' ? (words.endTitle || '') : '',
        subtitle: kind === 'title' ? (words.subtitle || '') : kind === 'end' ? (words.endSubtitle || '') : '',
        seconds: SCENE_SECONDS[kind],
        corner: 'br',
        shape: 'circle',
        takes: [],
        take: null,
        trim: null,
        zooms: null,
        look: null,
    };
}

/**
 * A new introduction: the template's scenes in a landscape project, in the
 * brand given. `words` are the template's words in the window's language:
 * { title, subtitle, endTitle, endSubtitle, placeholder }.
 */
export function newIntro({ brand = null, words = {}, layouts = INTRO_TEMPLATE, aspect = 'landscape' } = {}) {
    const doc = emptyProject();
    const scenes = [];
    for (const layout of layouts) scenes.push(newScene(layout, scenes, words));
    const b = normaliseBrand(brand);
    // A title card says who it is when the brand knows.
    const first = scenes.find((s) => s.layout === 'title');
    if (first && b.name) first.title = b.name;
    if (first && b.role) first.subtitle = b.role;
    return composeIntro({ ...doc, options: { ...doc.options, aspect }, intro: { scenes, brand: b, words: { ...words }, prompter: { speed: 140 } } });
}

export const isIntro = (doc) => Boolean(doc?.intro && Array.isArray(doc.intro.scenes));

/** A scene by id. */
export const sceneById = (doc, id) => (doc?.intro?.scenes || []).find((s) => s.id === id) || null;

/** The take on the timeline for a scene, or null. */
export const activeTake = (scene) => (scene?.takes || []).find((k) => k.id === scene.take) || null;

/** Where each scene runs on the clock: [{ id, start, end }], one after another. */
export function sceneSpans(doc) {
    const out = [];
    let at = 0;
    for (const s of doc?.intro?.scenes || []) {
        const len = sceneLength(s);
        out.push({ id: s.id, start: round(at), end: round(at + len) });
        at += len;
    }
    return out;
}

/** How long a scene runs: its take as trimmed, or the card's own length. */
export function sceneLength(scene) {
    const take = activeTake(scene);
    if (take && RECORDED.has(scene.layout)) {
        const room = Math.max(0.5, Number(take.duration) || 0.5);
        if (scene.trim && scene.trim.length > 0) return round(Math.min(scene.trim.length, room - (scene.trim.offset || 0)));
        return round(room);
    }
    return round(Math.max(0.5, Number(scene.seconds) || SCENE_SECONDS[scene.layout] || 4));
}

/** The scene that holds `t` on the clock, or null. */
export function sceneAt(doc, t) {
    const span = sceneSpans(doc).find((s) => t >= s.start - 0.001 && t < s.end - 0.001) || null;
    return span ? sceneById(doc, span.id) : null;
}

// ── where things sit (X1, X2) ───────────────────────────────────────────────

/** The frame in pixels for a shape, as the engine draws it. */
const FRAMES = { landscape: { W: 1920, H: 1080 }, square: { W: 1080, H: 1080 }, vertical: { W: 1080, H: 1920 } };

/** Words grow in a narrower frame: text sizes are fractions of the frame's width. */
const TEXT_SCALE = { landscape: 1, square: 1.45, vertical: 1.75 };

const RTL = /[֐-ࣿﭐ-﷿ﹰ-ﻼ]/;

/**
 * Where each role sits for a layout in a frame of `aspect`: fractions of
 * the frame (fx, fy the top left corner for pictures), `fit` for a picture
 * inside its box, or `full` for a whole-frame picture. `corner` and `shape`
 * place the camera's inset; `camAspect` is the camera's height over width.
 */
export function layoutGeometry(layout, aspect = 'landscape', { corner = 'br', shape = 'circle', camAspect = 9 / 16 } = {}) {
    const { W, H } = FRAMES[aspect] || FRAMES.landscape;
    const tall = aspect === 'vertical';
    const square = aspect === 'square';
    if (layout === 'talking') return { camera: { full: true }, screen: null };
    if (layout === 'screenCamera') {
        if (tall) {
            const fw = 0.56;
            return { screen: { fx: 0, fy: 0.07, fw: 1, fh: 0.5, fit: 'contain' }, camera: { fx: (1 - fw) / 2, fy: 0.6, fw, fh: (fw * W) / H, shape: 'circle' } };
        }
        const inset = insetRect({ corner, size: square ? 'medium' : 'medium', shape, W, H, aspect: camAspect });
        return { screen: { fx: 0, fy: 0, fw: 1, fh: 1, fit: 'contain' }, camera: { ...inset, shape } };
    }
    if (layout === 'sideBySide') {
        if (tall) return { camera: { fx: 0.1, fy: 0.06, fw: 0.8, fh: 0.4, shape: 'rounded' }, screen: { fx: 0.04, fy: 0.5, fw: 0.92, fh: 0.42, fit: 'contain' } };
        if (square) return { camera: { fx: 0.25, fy: 0.05, fw: 0.5, fh: 0.42, shape: 'rounded' }, screen: { fx: 0.04, fy: 0.5, fw: 0.92, fh: 0.45, fit: 'contain' } };
        return { camera: { fx: 0.04, fy: 0.17, fw: 0.36, fh: 0.66, shape: 'rounded' }, screen: { fx: 0.43, fy: 0.17, fw: 0.53, fh: 0.66, fit: 'contain' } };
    }
    // The title and end cards: no pictures, words and the logo.
    return { camera: null, screen: null };
}

/** The words of a card and the logo above them, for a frame of `aspect`. */
function cardGeometry(aspect, withLogo) {
    const k = TEXT_SCALE[aspect] || 1;
    const logoW = aspect === 'vertical' ? 0.3 : aspect === 'square' ? 0.22 : 0.14;
    if (withLogo) return { logo: { fx: 0.5 - logoW / 2, fy: aspect === 'vertical' ? 0.3 : 0.18, fw: logoW }, title: { fy: aspect === 'vertical' ? 0.5 : 0.52, size: 0.058 * k }, subtitle: { fy: aspect === 'vertical' ? 0.56 : 0.64, size: 0.028 * k } };
    return { logo: null, title: { fy: aspect === 'vertical' ? 0.42 : 0.4, size: 0.064 * k }, subtitle: { fy: aspect === 'vertical' ? 0.49 : 0.54, size: 0.03 * k } };
}

/** The lower third for a name and a role: a band behind the two, right to left when the name is. */
function lowerThirdGeometry(aspect, name, role) {
    const k = TEXT_SCALE[aspect] || 1;
    const nameSize = 0.032 * k;
    const roleSize = 0.021 * k;
    const tall = aspect === 'vertical';
    const top = tall ? 0.66 : aspect === 'square' ? 0.74 : 0.76;
    const H = (FRAMES[aspect] || FRAMES.landscape).H;
    const W = (FRAMES[aspect] || FRAMES.landscape).W;
    // The band as wide as the longer line, roughly: half a size a letter.
    const est = Math.max(String(name).length * nameSize * 0.56, String(role).length * roleSize * 0.52);
    const fw = Math.max(0.2, Math.min(0.88, est + 0.05));
    const margin = 0.045;
    const rtl = RTL.test(`${name}${role}`);
    const nameY = top + 0.012;
    const roleY = nameY + (nameSize * W * 1.32) / H;
    const bandH = (roleY - top) + (roleSize * W * 1.5) / H + 0.012;
    const left = rtl ? 1 - margin - fw : margin;
    return {
        band: { fx: left + fw / 2, fy: top + bandH / 2, fw, fh: bandH },
        name: { fx: rtl ? 1 - margin - 0.02 : margin + 0.02, fy: nameY, size: nameSize, align: rtl ? 'right' : 'left' },
        role: { fx: rtl ? 1 - margin - 0.02 : margin + 0.02, fy: roleY, size: roleSize, align: rtl ? 'right' : 'left' },
    };
}

/** The small logo in a corner of a recorded scene: the corner away from the camera. */
function watermarkGeometry(aspect, cameraCorner) {
    const fw = aspect === 'vertical' ? 0.16 : aspect === 'square' ? 0.11 : 0.07;
    const corner = { tl: 'tr', tr: 'tl', bl: 'tr', br: 'tr' }[cameraCorner] || 'tr';
    const mx = 0.025;
    const my = 0.03;
    return { fx: corner === 'tl' ? mx : 1 - mx - fw, fy: my, fw };
}

// ── reading the timeline back ───────────────────────────────────────────────

/** A scene's layers on the timeline, by role (the first of each). */
function layersOf(doc, sceneId) {
    const out = {};
    for (const p of doc.patches || []) if (p.scene === sceneId && p.role && !out[p.role]) out[p.role] = p;
    return out;
}

/**
 * What the timeline says about a scene, written back into it: the words of
 * its card and lower third, how its take was trimmed, its zooms, its
 * pointer's look and where its camera sits.
 */
function harvest(doc, scene) {
    const L = layersOf(doc, scene.id);
    const next = { ...scene };
    if (L.title && (scene.layout === 'title' || scene.layout === 'end')) next.title = String(L.title.text ?? next.title);
    if (L.subtitle && (scene.layout === 'title' || scene.layout === 'end')) next.subtitle = String(L.subtitle.text ?? next.subtitle);
    if (L.ltName) next.ltName = String(L.ltName.text ?? '');
    if (L.ltRole) next.ltRole = String(L.ltRole.text ?? '');
    // Read from the take that is on the timeline only: a layer of another take says nothing about this one.
    const take = activeTake(scene);
    const urls = new Set((take?.files || []).map((f) => f.url));
    const main = [L.screen, L.camera].find((p) => p && urls.has(p.url)) || null;
    if (main && take) next.trim = { offset: round(Number(main.offset) || 0), length: round(main.timing.end - main.timing.start) };
    else if (L.ground && !activeTake(scene)) next.seconds = round(L.ground.timing.end - L.ground.timing.start);
    // A screen hidden by its layout (a talking head) keeps its zooms for when it shows again.
    if (L.screen && urls.has(L.screen.url) && L.screen.visible !== false) next.zooms = L.screen.autoZoom ? L.screen.autoZoom : null;
    if (L.pointer) next.look = { highlight: L.pointer.highlight !== false, rings: L.pointer.rings !== false, arrow: Boolean(L.pointer.arrow), size: L.pointer.size || 1, color: L.pointer.color || POINTER_LOOK.color };
    if (L.camera && L.camera.fw && scene.layout === 'screenCamera') {
        next.corner = cornerOf(L.camera);
        if (L.camera.insetShape) next.shape = L.camera.insetShape;
    }
    return next;
}

// ── making the scenes' layers ───────────────────────────────────────────────

const fade = (seconds = SCENE_FADE) => ({ kind: 'fade', seconds });
const none = () => ({ kind: 'none', seconds: 0 });

/**
 * The layers of one scene from `start`, in a frame of `aspect`, in `brand`:
 * patches without ids. `lowerThird` puts the name and role over it.
 */
function sceneLayers(scene, { start, aspect, brand, words, lowerThird }) {
    const length = sceneLength(scene);
    const timing = { start: round(start), end: round(start + length) };
    const base = (role, extra) => ({ scene: scene.id, role, timing, enter: fade(), exit: fade(), ...extra });
    const out = [];
    const k = TEXT_SCALE[aspect] || 1;
    const font = brand.font;
    out.push(base('ground', { type: 'shape', name: 'ground', geometry: 'rect', fx: 0.5, fy: 0.5, fw: 1, fh: 1, color: brand.colors.background, lane: LANES.ground, locked: true }));

    if (scene.layout === 'title' || scene.layout === 'end') {
        const g = cardGeometry(aspect, Boolean(brand.logo));
        if (brand.logo) out.push(base('logo', { type: 'image', name: brand.logo.name || 'logo', url: brand.logo.url, fx: g.logo.fx, fy: g.logo.fy, fw: g.logo.fw, shadow: false, lane: LANES.logo }));
        out.push(base('title', { type: 'text', name: 'title', text: scene.title || '', fx: 0.5, fy: g.title.fy, sizeFrac: round(g.title.size), weight: 800, align: 'center', color: brand.colors.text, fontFamily: font, lane: LANES.title, enter: { kind: 'slide-up', seconds: 0.6, dist: 60 } }));
        out.push(base('subtitle', { type: 'text', name: 'subtitle', text: scene.subtitle || '', fx: 0.5, fy: g.subtitle.fy, sizeFrac: round(g.subtitle.size), weight: 600, align: 'center', color: brand.colors.accent, fontFamily: font, lane: LANES.subtitle }));
        return out;
    }

    const take = activeTake(scene);
    if (!take) {
        // Not recorded yet: the scene holds its place, saying so.
        out.push(base('placeholder', { type: 'text', name: 'placeholder', text: words.placeholder || '', fx: 0.5, fy: 0.46, sizeFrac: round(0.034 * k), weight: 600, align: 'center', color: brand.colors.text, opacity: 0.55, fontFamily: font, lane: LANES.placeholder }));
        return out;
    }

    const files = take.files || [];
    const screenFile = files.find((f) => f.role === 'screen') || null;
    const cameraFile = files.find((f) => f.role === 'camera') || null;
    const camAspect = cameraFile?.width && cameraFile?.height ? cameraFile.height / cameraFile.width : 9 / 16;
    let geo = layoutGeometry(scene.layout, aspect, { corner: scene.corner || 'br', shape: scene.shape || 'circle', camAspect });
    // A layout that wants a picture the take has not got shows the one it has, whole.
    if (!cameraFile && geo.screen === null) geo = { screen: { fx: 0, fy: 0, fw: 1, fh: 1, fit: 'contain' }, camera: null };
    if (!screenFile && geo.camera && !geo.camera.full) geo = { camera: { full: true }, screen: null };
    const offset = round(scene.trim?.offset || 0);
    const group = `scene-${scene.id}`;
    const media = (role, f, extra) => base(role, { type: 'video', name: f.name || role, url: f.url, offset, group, take: role, lane: LANES[role], ...extra });

    if (screenFile) {
        const g = geo.screen;
        // Locked: the screen is placed by its layout, never dragged on the stage by a stray press.
        const extra = g ? { fx: g.fx, fy: g.fy, fw: g.fw, fh: g.fh, fit: g.fit, locked: true } : { visible: false };
        if (screenFile.crop) extra.crop = screenFile.crop;
        const zooms = scene.zooms && scene.zooms.length ? scene.zooms : null;
        if (zooms && g) {
            extra.autoZoom = zooms;
            const keys = zoomKeys(zooms, { crop: screenFile.crop || null });
            if (keys) extra.keys = keys;
        }
        out.push(media('screen', screenFile, extra));
        if (screenFile.pointer && g) {
            const look = scene.look || POINTER_LOOK;
            out.push(base('pointer', { type: 'pointer', name: 'pointer', take: 'pointer', screenUrl: screenFile.url, offset, group, lane: LANES.pointer, enter: none(), exit: none(), highlight: look.highlight, rings: look.rings, arrow: look.arrow, size: look.size, color: look.color, samples: screenFile.pointer.samples || [], clicks: screenFile.pointer.clicks || [] }));
        }
    }
    if (cameraFile) {
        const g = geo.camera;
        const extra = !g ? { visible: false } : g.full ? { fit: 'cover', flipX: cameraFile.mirror !== false } : { fx: g.fx, fy: g.fy, fw: g.fw, fh: g.fh, insetShape: g.shape || 'rounded', flipX: cameraFile.mirror !== false };
        out.push(media('camera', cameraFile, extra));
    }
    for (const f of files) {
        if (!f.hasAudio) continue;
        const kind = f.role === 'camera' ? 'voice' : (cameraFile ? 'computer' : 'voice');
        out.push(base(kind, { type: 'sound', name: f.name || kind, url: f.url, offset, volume: 1, group, take: kind, lane: SOUND_LANES[kind], enter: fade(0.15), exit: fade(0.2) }));
    }

    if (brand.logo && brand.watermark) {
        const g = watermarkGeometry(aspect, scene.layout === 'screenCamera' ? scene.corner || 'br' : 'br');
        out.push(base('logo', { type: 'image', name: brand.logo.name || 'logo', url: brand.logo.url, fx: g.fx, fy: g.fy, fw: g.fw, shadow: false, opacity: 0.9, lane: LANES.logo }));
    }
    if (lowerThird && cameraFile) {
        const name = scene.ltName ?? brand.name;
        const role = scene.ltRole ?? brand.role;
        if (String(name || '').trim()) {
            const g = lowerThirdGeometry(aspect, name, role);
            const lt = { start: round(start + Math.min(0.8, length * 0.15)), end: round(start + Math.max(1, Math.min(length - 0.4, 5.5))) };
            const third = (role2, extra) => ({ ...base(role2, extra), timing: lt });
            out.push(third('band', { type: 'shape', name: 'lower third', geometry: 'rect', fx: g.band.fx, fy: g.band.fy, fw: g.band.fw, fh: g.band.fh, color: brand.colors.accent, radius: 0.012, opacity: 0.94, lane: LANES.band, enter: { kind: 'slide-up', seconds: 0.5, dist: 40 }, exit: fade(0.4) }));
            out.push(third('ltName', { type: 'text', name: 'name', text: name, fx: g.name.fx, fy: g.name.fy, sizeFrac: round(g.name.size), weight: 800, align: g.name.align, color: brand.colors.background, fontFamily: font, lane: LANES.title, enter: { kind: 'slide-up', seconds: 0.5, dist: 40 }, exit: fade(0.4) }));
            if (String(role || '').trim()) out.push(third('ltRole', { type: 'text', name: 'role', text: role, fx: g.role.fx, fy: g.role.fy, sizeFrac: round(g.role.size), weight: 600, align: g.role.align, color: brand.colors.background, fontFamily: font, lane: LANES.subtitle, enter: { kind: 'slide-up', seconds: 0.5, dist: 40 }, exit: fade(0.4) }));
        }
    }
    return out;
}

/**
 * The introduction's timeline made again from its scenes: each scene's
 * layers rebuilt one after another from the start (after the timeline's own
 * edits to them are read back), in the brand, laid out for `aspect` (the
 * project's own shape unless an export asks for another). Layers that are
 * no scene's are left as they are.
 */
export function composeIntro(doc, { aspect = null } = {}) {
    if (!isIntro(doc)) return doc;
    return build(doc, doc.intro.scenes.map((s) => harvest(doc, s)), aspect);
}

/** The scenes' layers made from `scenes` as they are (already read back from the timeline). */
function build(doc, scenes, aspect = null) {
    const shape = aspect || doc.options?.aspect || 'landscape';
    const brand = normaliseBrand(doc.intro.brand);
    const rest = (doc.patches || []).filter((p) => !p.scene);
    const made = [];
    let at = 0;
    // The lower third goes on the first recorded scene that shows the camera.
    const third = brand.lowerThird ? scenes.find((s) => RECORDED.has(s.layout) && (activeTake(s)?.files || []).some((f) => f.role === 'camera')) : null;
    for (const scene of scenes) {
        made.push(...sceneLayers(scene, { start: at, aspect: shape, brand, words: doc.intro.words || {}, lowerThird: third?.id === scene.id }));
        at += sceneLength(scene);
    }
    const scratch = { patches: [...rest] };
    const patches = [...rest];
    for (const p of made) {
        const id = nextId(scratch, p.type);
        scratch.patches.push({ id });
        patches.push({ id, ...p });
    }
    return { ...doc, patches, intro: { ...doc.intro, scenes, brand } };
}

/** An introduction laid out for another shape (what an export preset hands its scenes and captions). */
export function relayoutIntro(doc, aspect) {
    const laid = composeIntro(doc, { aspect });
    return isIntro(doc) ? laid : layoutCaptions(doc, aspect);
}

// ── changing the scenes (X1, X2, X4) ────────────────────────────────────────

/** The scenes changed by `fn` (handed them as the timeline has them) and the timeline made again. */
function withScenes(doc, fn) {
    if (!isIntro(doc)) return doc;
    const current = doc.intro.scenes.map((s) => harvest(doc, s));
    const scenes = fn(current);
    if (!scenes) return doc;
    return build({ ...doc, intro: { ...doc.intro, scenes } }, scenes);
}

/** A scene added after `after` (or at the end), of a layout. Answers { doc, id }. */
export function addScene(doc, layout, { after = null } = {}) {
    let id = null;
    const next = withScenes(doc, (scenes) => {
        const scene = newScene(layout, scenes, doc.intro.words || {});
        id = scene.id;
        const at = after ? scenes.findIndex((s) => s.id === after) + 1 : scenes.length;
        const out = [...scenes];
        out.splice(at > 0 ? at : scenes.length, 0, scene);
        return out;
    });
    return { doc: next, id };
}

/** A scene taken out, its takes with it (their files stay where they are). */
export function removeScene(doc, id) {
    return withScenes(doc, (scenes) => (scenes.length > 1 ? scenes.filter((s) => s.id !== id) : null));
}

/** A scene moved `by` places (-1 earlier, 1 later). */
export function moveScene(doc, id, by) {
    return withScenes(doc, (scenes) => {
        const i = scenes.findIndex((s) => s.id === id);
        const j = i + by;
        if (i < 0 || j < 0 || j >= scenes.length) return null;
        const out = [...scenes];
        const [s] = out.splice(i, 1);
        out.splice(j, 0, s);
        return out;
    });
}

/** A scene's layout switched (X2): its take stays, laid out the new way. */
export function setLayout(doc, id, layout) {
    if (!LAYOUTS.includes(layout)) return doc;
    return withScenes(doc, (scenes) => scenes.map((s) => (s.id === id ? { ...s, layout, seconds: activeTake(s) ? s.seconds : SCENE_SECONDS[layout] } : s)));
}

/** A scene's own fields changed: its script, a card's words, how long a card runs. */
export function updateScene(doc, id, change) {
    return withScenes(doc, (scenes) => scenes.map((s) => (s.id === id ? { ...s, ...change } : s)));
}

/**
 * A take recorded for a scene (X4): its files ({ role, url, name, duration,
 * width, height, hasAudio, crop?, mirror?, pointer? }) kept as a new take,
 * which goes on the timeline in place of the one before; the earlier takes
 * are kept with the scene until the project is cleaned up. A screen with
 * recorded clicks is zoomed where they are (X5). Answers { doc, take }.
 */
export function recordScene(doc, id, files, { autoZoom = true } = {}) {
    let made = null;
    const next = withScenes(doc, (scenes) => scenes.map((s) => {
        if (s.id !== id) return s;
        const kept = (files || []).filter((f) => f && f.url && (f.role === 'screen' || f.role === 'camera')).map((f) => ({
            role: f.role, url: f.url, name: f.name || '', duration: Number(f.duration) || 0, width: f.width || null, height: f.height || null,
            hasAudio: Boolean(f.hasAudio), ...(f.crop ? { crop: f.crop } : {}), ...(f.role === 'camera' ? { mirror: f.mirror !== false } : {}),
            ...(f.pointer ? { pointer: compactPointer(f.pointer) } : {}),
        }));
        if (!kept.length) return s;
        const duration = round(Math.max(...kept.map((f) => f.duration || 0), 0.5));
        const take = { id: freshId(s.takes, 'k'), files: kept, duration };
        made = take;
        return { ...s, takes: [...s.takes, take], take: take.id, trim: null, zooms: autoZoom ? takeZooms(take) : null, ltName: undefined, ltRole: undefined };
    }));
    return { doc: next, take: made };
}

/** The automatic zooms a take's screen makes on its clicks (X5), or null. */
function takeZooms(take) {
    const screen = (take?.files || []).find((f) => f.role === 'screen');
    if (!screen?.pointer?.clicks?.length) return null;
    const zooms = findZooms(screen.pointer.clicks, { length: take.duration });
    return zooms && zooms.length ? zooms : null;
}

/** An earlier take put back on the timeline for a scene, zoomed on its own clicks. */
export function useTake(doc, id, takeId) {
    return withScenes(doc, (scenes) => scenes.map((s) => {
        const take = s.id === id ? s.takes.find((k) => k.id === takeId) : null;
        return take ? { ...s, take: takeId, trim: null, zooms: takeZooms(take) } : s;
    }));
}

/** The brand kit applied to every scene (X6). */
export function applyBrand(doc, brand) {
    if (!isIntro(doc)) return doc;
    const b = normaliseBrand(brand);
    const old = normaliseBrand(doc.intro.brand);
    const words = doc.intro.words || {};
    // A title card still saying the template's words, or the old brand's, says the new brand's.
    const follows = (text, placeholder, was) => !String(text || '').trim() || text === placeholder || text === was;
    // The lower third's words follow the brand again.
    const scenes = doc.intro.scenes.map((s) => harvest(doc, s)).map((s) => ({
        ...s,
        ltName: undefined,
        ltRole: undefined,
        ...(s.layout === 'title' && b.name && follows(s.title, words.title, old.name) ? { title: b.name } : {}),
        ...(s.layout === 'title' && b.role && follows(s.subtitle, words.subtitle, old.role) ? { subtitle: b.role } : {}),
    }));
    return build({ ...doc, intro: { ...doc.intro, scenes, brand: b } }, scenes);
}

/** The teleprompter's speed, in words a minute. */
export function setPrompterSpeed(doc, speed) {
    if (!isIntro(doc)) return doc;
    return { ...doc, intro: { ...doc.intro, prompter: { ...(doc.intro.prompter || {}), speed: Math.max(60, Math.min(260, Math.round(Number(speed) || 140))) } } };
}

/** The files every take of the project uses, the ones on the timeline and the ones kept: what a clean-up may offer to delete. */
export function takeFiles(doc) {
    const onLine = new Set(clipsOf(doc).map((p) => p.url).filter(Boolean));
    const all = [];
    for (const s of doc?.intro?.scenes || []) for (const k of s.takes || []) for (const f of k.files || []) all.push({ scene: s.id, take: k.id, url: f.url, used: onLine.has(f.url) });
    return all;
}

/** The scene a timeline layer belongs to, or null. */
export const sceneOfLayer = (doc, layerId) => {
    const p = clipById(doc, layerId);
    return p?.scene ? sceneById(doc, p.scene) : null;
};

// ── the teleprompter (X3) ───────────────────────────────────────────────────

/** The words a minute the teleprompter offers. */
export const PROMPTER_SPEEDS = { slow: 110, normal: 140, fast: 170 };

/** How many words a script has. */
export const wordCount = (script) => (String(script || '').trim().match(/\S+/g) || []).length;

/** About how long a script takes to read at `wpm`, in seconds. */
export const scriptSeconds = (script, wpm = 140) => round((wordCount(script) / Math.max(1, wpm)) * 60);

/**
 * How far through its script the teleprompter has scrolled, 0 to 1, after
 * `elapsed` seconds of recording at `wpm`: a second's pause first to start
 * speaking, then the words at the speed set.
 */
export function prompterProgress(script, elapsed, wpm = 140, lead = 1) {
    const words = wordCount(script);
    if (!words) return 0;
    const read = Math.max(0, (Number(elapsed) || 0) - lead) * (Math.max(1, wpm) / 60);
    return Math.max(0, Math.min(1, read / words));
}
