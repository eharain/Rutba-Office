// Animations, played.
//
// The engine reads a slide's main sequence into a list — each effect with
// its shape, kind, effect, direction, trigger, duration and delay — and the
// slide's SVG wraps every shape's drawing in `<g data-shape="id">` (a
// group's members also carry `data-groups`). This turns the list into click
// groups with start times, and plays one group at a time with the Web
// Animations API on those wrappers: transform, opacity, clip-path and
// visibility, each effect holding its end state, so a shape that has come
// in stays in and one that has gone out stays out. An effect on one
// paragraph plays on that paragraph's lines (`data-para`), and a motion
// path moves the shape along its path, sampled into keyframes, and leaves
// it where the path ends.

import { MOTION_PATHS } from '@rutba/presentation/timing';
import { t } from '@rutba/office-ui';

/** The Animation gallery's entrances, in PowerPoint's order, and what each does. */
export const ANIMATION_GALLERY = [
  ['appear', t('Appear'), t('the shape is simply there on the click')],
  ['fade', t('Fade'), t('the shape fades in')],
  ['fly', t('Fly In'), t('the shape flies in from off the slide')],
  ['float', t('Float In'), t('the shape drifts up into place as it fades in')],
  ['split', t('Split'), t('the shape opens out from its middle')],
  ['wipe', t('Wipe'), t('the shape is wiped on from one side')],
  ['zoom', t('Zoom'), t('the shape grows from its centre')],
];
/** More Effects and Add Animation: every effect, under PowerPoint's three headings. */
export const EFFECT_MENU = [
  ['entr', t('Entrance'), ANIMATION_GALLERY.map(([effect, label]) => [effect, label])],
  ['emph', t('Emphasis'), [['pulse', t('Pulse')], ['spin', t('Spin')], ['grow', t('Grow/Shrink')]]],
  ['exit', t('Exit'), [['appear', t('Disappear')], ['fade', t('Fade')], ['fly', t('Fly Out')], ['float', t('Float Out')], ['split', t('Split')], ['wipe', t('Wipe')], ['zoom', t('Zoom')]]],
  // Motion Paths: each a path from the shape's centre (`motionPath`).
  ['path', t('Motion Paths'), MOTION_PATHS.map(([path, label]) => [path, t(label)])],
];
/** Effect Options for each effect that has a direction, as `[value, label]` — the engine's own values. */
export const ANIMATION_OPTIONS = {
  fly: [['bottom', t('From Bottom')], ['bottom-left', t('From Bottom-Left')], ['left', t('From Left')], ['top-left', t('From Top-Left')], ['top', t('From Top')], ['top-right', t('From Top-Right')], ['right', t('From Right')], ['bottom-right', t('From Bottom-Right')]],
  wipe: [['bottom', t('From Bottom')], ['left', t('From Left')], ['right', t('From Right')], ['top', t('From Top')]],
  split: [['vertical-out', t('Vertical Out')], ['vertical-in', t('Vertical In')], ['horizontal-out', t('Horizontal Out')], ['horizontal-in', t('Horizontal In')]],
  float: [['up', t('Float Up')], ['down', t('Float Down')]],
  spin: [['clockwise', t('Clockwise')], ['counterclockwise', t('Counterclockwise')]],
};

/** Every element drawing a shape, or the members of a group — or, given a paragraph, that paragraph's lines in the shape. */
export function shapeNodes(root, id, paragraph = null) {
  if (!root || id == null) return [];
  const safe = String(id).replace(/"/g, '');
  const nodes = [...root.querySelectorAll(`[data-shape="${safe}"], [data-groups~="${safe}"]`)];
  if (paragraph == null) return nodes;
  return nodes.flatMap((el) => [...el.querySelectorAll(`[data-para="${Number(paragraph)}"]`)]);
}

/** What an effect acts on, as the show keeps it: the shape, or one paragraph of it. */
const targetKey = (e) => (e.paragraph != null ? `${e.shapeId}#${e.paragraph}` : String(e.shapeId));
const targetNodes = (root, key) => {
  const [id, para] = String(key).split('#');
  return shapeNodes(root, id, para === undefined ? null : para);
};

/**
 * A motion path (PowerPoint's form: M, L, C, Z in fractions of the slide,
 * ending in E) as points along it in pixels, each with how far along the
 * whole path it lies — the keyframes' offsets.
 */
export function pathPoints(path, size) {
  const W = size?.width || 960;
  const H = size?.height || 540;
  const tokens = String(path || '').replace(/([MmLlCcZzEe])/g, ' $1 ').trim().split(/[\s,]+/).filter(Boolean);
  const pts = [];
  let i = 0;
  let cmd = null;
  let x = 0, y = 0, sx = 0, sy = 0;
  const num = () => Number(tokens[i++]);
  while (i < tokens.length) {
    if (/^[A-Za-z]$/.test(tokens[i])) cmd = tokens[i++];
    if (!cmd || cmd === 'E' || cmd === 'e') break;
    const rel = cmd === cmd.toLowerCase();
    const C = cmd.toUpperCase();
    if (C === 'Z') { pts.push([sx, sy]); x = sx; y = sy; cmd = null; continue; }
    if (C === 'M' || C === 'L') {
      const nx = num(), ny = num();
      if (!Number.isFinite(nx) || !Number.isFinite(ny)) break;
      x = rel ? x + nx : nx; y = rel ? y + ny : ny;
      if (C === 'M') { sx = x; sy = y; }
      pts.push([x, y]);
      if (C === 'M') cmd = rel ? 'l' : 'L';
    } else if (C === 'C') {
      const v = [num(), num(), num(), num(), num(), num()];
      if (v.some((n) => !Number.isFinite(n))) break;
      const [x1, y1, x2, y2, x3, y3] = rel ? [x + v[0], y + v[1], x + v[2], y + v[3], x + v[4], y + v[5]] : v;
      for (let k = 1; k <= 16; k++) {
        const t = k / 16, u = 1 - t;
        pts.push([u * u * u * x + 3 * u * u * t * x1 + 3 * u * t * t * x2 + t * t * t * x3, u * u * u * y + 3 * u * u * t * y1 + 3 * u * t * t * y2 + t * t * t * y3]);
      }
      x = x3; y = y3;
    } else {
      break;
    }
  }
  if (!pts.length) return [];
  const px = pts.map(([a, b]) => [a * W, b * H]);
  let total = 0;
  const along = [0];
  for (let k = 1; k < px.length; k++) { total += Math.hypot(px[k][0] - px[k - 1][0], px[k][1] - px[k - 1][1]); along.push(total); }
  return px.map(([a, b], k) => ({ x: a, y: b, offset: total ? along[k] / total : k / Math.max(1, px.length - 1) }));
}

/** Where a shape stands after its motion paths so far: the sum of each path's end, in pixels. */
function pathEnd(path, size) {
  const pts = pathPoints(path, size);
  const end = pts[pts.length - 1];
  return end ? { x: end.x - pts[0].x, y: end.y - pts[0].y } : { x: 0, y: 0 };
}

/**
 * The list as the show plays it: the group that starts with the slide (or
 * null) and one group per click, each effect with its start in seconds
 * from the group's start — With Previous shares the start of the effect
 * before it, After Previous starts when everything before it has ended.
 */
export function sequence(animations = [], { interactive = false } = {}) {
  const groups = [];
  let current = null;
  let subStart = 0;
  let subEnd = 0;
  // The slide's own sequence leaves out what a click on a trigger starts.
  for (const e of interactive ? animations : animations.filter((a) => !a.triggerShape)) {
    if (!current || e.trigger === 'onClick') {
      current = { auto: !interactive && !groups.length && e.trigger !== 'onClick', effects: [], length: 0 };
      groups.push(current);
      subStart = 0;
      subEnd = 0;
    } else if (e.trigger === 'afterPrevious') {
      subStart = subEnd;
    }
    const start = subStart + (Number(e.delay) || 0);
    subEnd = Math.max(subEnd, start + (Number(e.duration) || 0));
    current.effects.push({ ...e, start });
    current.length = Math.max(current.length, subEnd);
  }
  const auto = groups[0]?.auto ? groups[0] : null;
  return { auto, clicks: auto ? groups.slice(1) : groups };
}

/** How many clicks a slide's animations take before the next click moves the show on. */
export const clickCount = (animations) => sequence(animations).clicks.length;

/**
 * Animations → Trigger: each trigger shape's own groups, played one per
 * click on that shape — `Map(shapeId → groups)`.
 */
export function triggered(animations = []) {
  const by = new Map();
  for (const e of animations) {
    if (!e.triggerShape) continue;
    if (!by.has(e.triggerShape)) by.set(e.triggerShape, []);
    by.get(e.triggerShape).push(e);
  }
  return new Map([...by].map(([id, list]) => [id, sequence(list, { interactive: true }).clicks]));
}

/**
 * Whether each animated shape shows once `step` clicks have played (the
 * group that starts with the slide counted in when `auto` is true), and
 * each trigger's groups that `fired` says have played: a shape whose first
 * effect is an entrance starts hidden, an entrance shows it and an exit
 * hides it.
 */
export function visibilityAt(animations = [], step = 0, auto = true, fired = null) {
  const seq = sequence(animations);
  const shown = new Map();
  for (const e of animations) if (e.shapeId != null && e.kind !== 'path' && !shown.has(targetKey(e))) shown.set(targetKey(e), e.kind !== 'entr');
  const byTrigger = fired?.size ? triggered(animations) : null;
  const firedGroups = byTrigger ? [...fired].flatMap(([id, n]) => (byTrigger.get(id) || []).slice(0, n)) : [];
  const played = [...(auto && seq.auto ? [seq.auto] : []), ...seq.clicks.slice(0, Math.max(0, step)), ...firedGroups];
  for (const g of played) {
    for (const e of g.effects) {
      if (e.kind === 'entr') shown.set(targetKey(e), true);
      else if (e.kind === 'exit') shown.set(targetKey(e), false);
    }
  }
  return shown;
}

/** How far each shape has travelled by its motion paths once `step` clicks have played: `Map(shapeId → { x, y })`. */
export function travelAt(animations = [], step = 0, auto = true, size = null) {
  const seq = sequence(animations);
  const moved = new Map();
  const played = [...(auto && seq.auto ? [seq.auto] : []), ...seq.clicks.slice(0, Math.max(0, step))];
  for (const g of played) {
    for (const e of g.effects) {
      if (e.kind !== 'path' || !e.path) continue;
      const d = pathEnd(e.path, size);
      const was = moved.get(String(e.shapeId)) || { x: 0, y: 0 };
      moved.set(String(e.shapeId), { x: was.x + d.x, y: was.y + d.y });
    }
  }
  return moved;
}

/**
 * Put a slide's shapes where the sequence has them after `step` clicks, at
 * once — no animation: entering a slide, stepping back, or the presenter's
 * picture of where the show is.
 */
export function applyState(root, animations = [], step = 0, auto = true, fired = null, size = null) {
  const shown = visibilityAt(animations, step, auto, fired);
  for (const [key, visible] of shown) {
    for (const el of targetNodes(root, key)) {
      for (const a of el.getAnimations?.() || []) a.cancel();
      el.style.visibility = visible ? '' : 'hidden';
    }
  }
  // A shape a motion path has moved stands where the path left it.
  if (animations.some((e) => e.kind === 'path')) {
    const moved = travelAt(animations, step, auto, size || (root?.querySelector?.('svg')?.viewBox?.baseVal?.width ? { width: root.querySelector('svg').viewBox.baseVal.width, height: root.querySelector('svg').viewBox.baseVal.height } : null));
    for (const e of animations) {
      if (e.kind !== 'path') continue;
      const d = moved.get(String(e.shapeId));
      for (const el of shapeNodes(root, e.shapeId)) {
        for (const a of el.getAnimations?.() || []) a.cancel();
        el.style.transform = d && (d.x || d.y) ? `translate(${d.x}px, ${d.y}px)` : '';
      }
    }
  }
}

const EASE_OUT = 'cubic-bezier(.2, .7, .3, 1)';
const EASE = 'cubic-bezier(.45, 0, .25, 1)';

/** Two clip-path bands closing on (or opening from) the middle: `a` is each band's width in percent. */
function bands(vertical, a) {
  const b = 100 - a;
  return vertical
    ? `polygon(0% 0%, ${a}% 0%, ${a}% 100%, 0% 100%, 0% 0%, ${b}% 0%, 100% 0%, 100% 100%, ${b}% 100%, ${b}% 0%)`
    : `polygon(0% 0%, 100% 0%, 100% ${a}%, 0% ${a}%, 0% 0%, 0% ${b}%, 100% ${b}%, 100% 100%, 0% 100%, 0% ${b}%)`;
}

/** Keyframes for an entrance, in the direction it comes from. Exits play these backwards. */
function entranceFrames(effect, direction, g, size) {
  const W = size?.width || 960;
  const H = size?.height || 540;
  switch (effect) {
    case 'appear':
      return { frames: [{ opacity: 1 }, { opacity: 1 }], easing: 'linear' };
    case 'fade':
      return { frames: [{ opacity: 0 }, { opacity: 1 }], easing: 'linear' };
    case 'fly': {
      const side = direction || 'bottom';
      const dx = side.includes('left') ? -(g.x + g.w) : side.includes('right') ? W - g.x : 0;
      const dy = side.includes('top') ? -(g.y + g.h) : side.includes('bottom') ? H - g.y : 0;
      return { frames: [{ transform: `translate(${dx}px, ${dy}px)` }, { transform: 'translate(0px, 0px)' }], easing: EASE_OUT };
    }
    case 'float': {
      const dy = (direction === 'down' ? -0.1 : 0.1) * H;
      return { frames: [{ transform: `translate(0px, ${dy}px)`, opacity: 0 }, { transform: 'translate(0px, 0px)', opacity: 1 }], easing: EASE_OUT };
    }
    case 'wipe': {
      const start = { bottom: 'inset(100% 0% 0% 0%)', top: 'inset(0% 0% 100% 0%)', left: 'inset(0% 100% 0% 0%)', right: 'inset(0% 0% 0% 100%)' }[direction || 'bottom'] || 'inset(100% 0% 0% 0%)';
      return { frames: [{ clipPath: start }, { clipPath: 'inset(0% 0% 0% 0%)' }], easing: 'linear' };
    }
    case 'split': {
      const [orient, way] = String(direction || 'vertical-out').split('-');
      const vertical = orient === 'vertical';
      if (way === 'in') return { frames: [{ clipPath: bands(vertical, 0) }, { clipPath: bands(vertical, 50) }], easing: 'linear' };
      const closed = vertical ? 'inset(0% 50% 0% 50%)' : 'inset(50% 0% 50% 0%)';
      return { frames: [{ clipPath: closed }, { clipPath: 'inset(0% 0% 0% 0%)' }], easing: 'linear' };
    }
    case 'zoom':
      return { frames: [{ transform: 'scale(0)', opacity: 0 }, { transform: 'scale(1)', opacity: 1 }], easing: EASE_OUT };
    default:
      return { frames: [{ opacity: 0 }, { opacity: 1 }], easing: 'linear' };
  }
}

/** One effect's keyframes and easing, visibility included: an entrance shows its shape, an exit hides it at the end. */
function effectFrames(e, g, size, from = { x: 0, y: 0 }) {
  if (e.kind === 'path') {
    // Along the path from where the shape stands, smooth at the start and the end.
    const pts = pathPoints(e.path, size);
    if (!pts.length) return { frames: [{}, {}], easing: 'linear' };
    const x0 = pts[0].x, y0 = pts[0].y;
    const frames = pts.map((p) => ({ transform: `translate(${from.x + p.x - x0}px, ${from.y + p.y - y0}px)`, offset: p.offset }));
    frames[0].offset = 0;
    frames[frames.length - 1].offset = 1;
    for (let k = 1; k < frames.length; k++) if (frames[k].offset < frames[k - 1].offset) frames[k].offset = frames[k - 1].offset;
    return { frames, easing: EASE };
  }
  if (e.kind === 'emph') {
    switch (e.effect) {
      case 'spin':
        return { frames: [{ transform: 'rotate(0deg)' }, { transform: `rotate(${e.direction === 'counterclockwise' ? -360 : 360}deg)` }], easing: EASE };
      case 'grow':
        return { frames: [{ transform: 'scale(1)' }, { transform: 'scale(1.5)' }], easing: EASE };
      case 'pulse':
      default:
        return { frames: [{ transform: 'scale(1)', opacity: 1 }, { transform: 'scale(1.05)', opacity: 0.75, offset: 0.5 }, { transform: 'scale(1)', opacity: 1 }], easing: EASE };
    }
  }
  const effect = e.effect || 'fade';
  if (e.kind === 'exit') {
    // An exit is its entrance backwards, and ends hidden.
    const side = e.direction;
    const { frames, easing } = entranceFrames(effect, effect === 'float' ? (side === 'down' ? 'up' : 'down') : effect === 'split' ? String(side || 'vertical-in').replace(/-(in|out)$/, (m, w) => (w === 'in' ? '-out' : '-in')) : side, g, size);
    const back = frames.slice().reverse().map((f, i, all) => ({ ...f, offset: undefined, visibility: i === all.length - 1 ? 'hidden' : 'visible' }));
    return { frames: back, easing: easing === EASE_OUT ? 'cubic-bezier(.7, 0, .8, .3)' : easing };
  }
  // An entrance (or anything else) shows the shape from its first frame.
  const { frames, easing } = entranceFrames(effect, e.direction, g, size);
  return { frames: frames.map((f) => ({ ...f, visibility: 'visible' })), easing };
}

/**
 * Play one group of effects on a slide's layer. Returns `{ finished,
 * finish() }`: the promise settles when every effect has ended, and
 * `finish()` jumps them all to their end at once — a click during an
 * animation completes it, as in PowerPoint.
 */
export function playGroup(root, group, { shapes = [], size = null } = {}) {
  const geometry = new Map(shapes.map((s) => [String(s.id), s.geometry]));
  const animations = [];
  for (const e of group?.effects || []) {
    const g = geometry.get(String(e.shapeId)) || { x: 0, y: 0, w: 0, h: 0 };
    // A path starts where the paths before it left the shape.
    const at = e.kind === 'path' ? (shapeNodes(root, e.shapeId)[0]?.style.transform.match(/translate\(([-\d.]+)px,\s*([-\d.]+)px\)/) || null) : null;
    const { frames, easing } = effectFrames(e, g, size, at ? { x: Number(at[1]), y: Number(at[2]) } : { x: 0, y: 0 });
    const duration = Math.max(1, Math.round((Number(e.duration) || 0) * 1000));
    for (const el of shapeNodes(root, e.shapeId, e.kind === 'path' ? null : e.paragraph)) {
      // An emphasis starts from where the shape is; an entrance or an exit
      // replaces what the shape held, from the state it had: hidden until
      // an entrance's own start (its delay included), shown until an exit's.
      if (e.kind === 'entr' || e.kind === 'exit') {
        for (const a of el.getAnimations()) a.cancel();
        el.style.visibility = e.kind === 'entr' ? 'hidden' : '';
      }
      const anim = el.animate(frames, { duration, delay: Math.round(e.start * 1000), easing, fill: 'forwards' });
      // A path's end is kept in the element's own style, so the next path starts there and stepping back can undo it.
      if (e.kind === 'path') anim.finished.then(() => { const last = frames[frames.length - 1]; el.style.transform = last.transform || ''; }).catch(() => {});
      animations.push(anim);
    }
  }
  const finished = Promise.all(animations.map((a) => a.finished.catch(() => null))).then(() => undefined);
  return {
    finished,
    running: () => animations.some((a) => a.playState === 'running' || a.playState === 'pending'),
    finish() {
      for (const a of animations) { try { a.finish(); } catch { /* cancelled */ } }
    },
  };
}
