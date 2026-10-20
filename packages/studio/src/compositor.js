// The Recorder's scenes, as boxes on the frame (docs/STUDIO.md, Recorder).
//
// A recording or a stream is one picture made live from its sources: the
// screen, the camera, a title, a logo. Each scene says where each goes, and
// this works out the boxes; the window draws them on a canvas a frame at a
// time, and the same picture goes to the file and to the stream. Nothing
// here touches a canvas, so the layouts are tested in Node.
//
// Boxes are in pixels of the output frame, W by H.

import { insetRect } from './recorder.js';

/** The scenes there are, in the order the window offers them. */
export const SCENES = ['screen-camera', 'screen', 'camera', 'split', 'brb'];

/** The sizes a recording or a stream is made at, by height. */
export const OUTPUT_SIZES = {
    720: { width: 1280, height: 720 },
    1080: { width: 1920, height: 1080 },
};

/** What a Recorder starts with. */
export const STUDIO_DEFAULTS = Object.freeze({
    scene: 'screen-camera',
    size: 1080,
    fps: 30,
    quality: 'standard',
    countdown: 3,
    camera: true,
    cameraId: '',
    mirror: true,
    corner: 'br',
    insetSize: 'medium',
    shape: 'circle',
    mic: true,
    micId: '',
    system: false,
    titleOn: false,
    title: '',
    subtitle: '',
    brb: 'Be right back',
    logo: '',
    logoCorner: 'tr',
    folder: '',
    minimise: false,
    // Go Live: a copy recorded on this computer while it streams.
    liveCopy: true,
    // Take Picture: of the program as shown, the screen or the camera; after a few seconds or at once.
    shot: 'program',
    shotTimer: 0,
});

const pick = (v, list, fallback) => (list.includes(v) ? v : fallback);
const text = (v, fallback, max = 120) => (typeof v === 'string' ? v.slice(0, max) : fallback);

/** Stored settings, or anything, as settings the Recorder can use. */
export function normaliseStudio(raw) {
    const r = raw && typeof raw === 'object' ? raw : {};
    const d = STUDIO_DEFAULTS;
    const bool = (k) => (r[k] === undefined ? d[k] : Boolean(r[k]));
    return {
        scene: pick(r.scene, SCENES, d.scene),
        size: pick(Number(r.size), [720, 1080], d.size),
        fps: pick(Number(r.fps), [24, 30, 60], d.fps),
        quality: pick(r.quality, ['small', 'standard', 'high'], d.quality),
        countdown: pick(Number(r.countdown), [0, 3, 5, 10], d.countdown),
        camera: bool('camera'),
        cameraId: text(r.cameraId, '', 400),
        mirror: bool('mirror'),
        corner: pick(r.corner, ['tl', 'tr', 'bl', 'br'], d.corner),
        insetSize: pick(r.insetSize, ['small', 'medium', 'large'], d.insetSize),
        shape: pick(r.shape, ['circle', 'rounded', 'rect'], d.shape),
        mic: bool('mic'),
        micId: text(r.micId, '', 400),
        system: bool('system'),
        titleOn: bool('titleOn'),
        title: text(r.title, d.title),
        subtitle: text(r.subtitle, d.subtitle),
        brb: text(r.brb, d.brb),
        logo: text(r.logo, '', 2000),
        logoCorner: pick(r.logoCorner, ['tl', 'tr', 'bl', 'br'], d.logoCorner),
        folder: text(r.folder, '', 2000),
        minimise: bool('minimise'),
        liveCopy: bool('liveCopy'),
        shot: pick(r.shot, ['program', 'screen', 'camera'], d.shot),
        shotTimer: pick(Number(r.shotTimer), [0, 3, 5, 10], d.shotTimer),
    };
}

/** Where a picture of w by h sits inside a box, whole, letterboxed. */
export function containRect(w, h, box) {
    if (!(w > 0) || !(h > 0)) return { ...box };
    const k = Math.min(box.w / w, box.h / h);
    const cw = w * k;
    const ch = h * k;
    return { x: box.x + (box.w - cw) / 2, y: box.y + (box.h - ch) / 2, w: cw, h: ch };
}

/** The part of a picture of w by h that fills a box, cropped from its middle: { sx, sy, sw, sh }. */
export function coverCrop(w, h, box) {
    if (!(w > 0) || !(h > 0) || !(box.w > 0) || !(box.h > 0)) return { sx: 0, sy: 0, sw: w || 1, sh: h || 1 };
    const want = box.w / box.h;
    const have = w / h;
    if (have > want) {
        const sw = h * want;
        return { sx: (w - sw) / 2, sy: 0, sw, sh: h };
    }
    const sh = w / want;
    return { sx: 0, sy: (h - sh) / 2, sw: w, sh };
}

/**
 * A scene as layers to draw, bottom first: { background, layers, card }.
 * A layer is { source: 'screen' | 'camera', box, fit: 'contain' | 'cover',
 * shape: 'rect' | 'rounded' | 'circle', mirror }. `card` is the words of a
 * scene that is a card (Be right back). A source that is not there is left
 * out, and a scene that needs it shows what there is.
 */
export function sceneLayout(scene, { W = 1920, H = 1080, screen = null, camera = null, settings = {} } = {}) {
    const s = normaliseStudio(settings);
    const layers = [];
    const cam = (box, shape, fit = 'cover') => ({ source: 'camera', box, fit, shape, mirror: s.mirror });
    const scr = (box) => ({ source: 'screen', box, fit: 'contain', shape: 'rect', mirror: false });
    const full = { x: 0, y: 0, w: W, h: H };
    switch (scene) {
        case 'screen':
            if (screen) layers.push(scr(full));
            else if (camera) layers.push(cam(full, 'rect'));
            break;
        case 'camera':
            if (camera) layers.push(cam(full, 'rect'));
            else if (screen) layers.push(scr(full));
            break;
        case 'split': {
            // The screen on the left two thirds, the camera beside it, both with a margin.
            const m = Math.round(H * 0.04);
            const left = { x: m, y: m, w: Math.round(W * 0.68) - m * 1.5, h: H - 2 * m };
            const right = { x: left.x + left.w + m, y: m, w: W - (left.x + left.w + m) - m, h: H - 2 * m };
            if (screen) layers.push({ ...scr(left), shape: 'rounded' });
            if (camera) layers.push(cam(screen ? right : full, screen ? 'rounded' : 'rect'));
            break;
        }
        case 'brb':
            return { background: 'card', layers: [], card: { text: s.brb || STUDIO_DEFAULTS.brb } };
        case 'screen-camera':
        default:
            if (screen) layers.push(scr(full));
            if (camera) {
                if (!screen) layers.push(cam(full, 'rect'));
                else {
                    const aspect = camera.w > 0 && camera.h > 0 ? camera.h / camera.w : 9 / 16;
                    const r = insetRect({ corner: s.corner, size: s.insetSize, shape: s.shape, W, H, aspect });
                    layers.push(cam({ x: r.fx * W, y: r.fy * H, w: r.fw * W, h: r.fh * H }, s.shape));
                }
            }
    }
    return { background: scene === 'split' ? 'studio' : 'black', layers, card: null };
}

/** The title band across the foot of the frame, its words' sizes and places. */
export function titleBand({ W = 1920, H = 1080 } = {}) {
    const h = Math.round(H * 0.13);
    const x = Math.round(W * 0.04);
    const y = Math.round(H - h - H * 0.06);
    return {
        band: { x, y, w: Math.round(W * 0.46), h },
        accent: { x, y, w: Math.round(W * 0.008), h },
        title: { x: x + Math.round(W * 0.022), y: y + Math.round(h * 0.42), size: Math.round(h * 0.34) },
        subtitle: { x: x + Math.round(W * 0.022), y: y + Math.round(h * 0.78), size: Math.round(h * 0.22) },
    };
}

/** A logo of w by h in a corner, at most a tenth of the frame's width, inside the margin. */
export function logoBox({ W = 1920, H = 1080, w = 1, h = 1, corner = 'tr' } = {}) {
    const m = Math.round(H * 0.035);
    const maxW = W * 0.1;
    const maxH = H * 0.1;
    const k = Math.min(maxW / Math.max(1, w), maxH / Math.max(1, h));
    const bw = Math.round(w * k);
    const bh = Math.round(h * k);
    const x = corner === 'tl' || corner === 'bl' ? m : W - m - bw;
    const y = corner === 'tl' || corner === 'tr' ? m : H - m - bh;
    return { x, y, w: bw, h: bh };
}
