// Animations, played.
//
// The engine reads a slide's main sequence into a list — each effect with
// its shape, kind, effect, direction, trigger, duration and delay — and the
// slide's SVG wraps every shape's drawing in `<g data-shape="id">` (a
// group's members also carry `data-groups`). This turns the list into click
// groups with start times, and plays one group at a time with the Web
// Animations API on those wrappers: transform, opacity, clip-path and
// visibility, each effect holding its end state, so a shape that has come
// in stays in and one that has gone out stays out.

/** The Animation gallery's entrances, in PowerPoint's order, and what each does. */
export const ANIMATION_GALLERY = [
  ['appear', 'Appear', 'the shape is simply there on the click'],
  ['fade', 'Fade', 'the shape fades in'],
  ['fly', 'Fly In', 'the shape flies in from off the slide'],
  ['float', 'Float In', 'the shape drifts up into place as it fades in'],
  ['split', 'Split', 'the shape opens out from its middle'],
  ['wipe', 'Wipe', 'the shape is wiped on from one side'],
  ['zoom', 'Zoom', 'the shape grows from its centre'],
];
/** More Effects and Add Animation: every effect, under PowerPoint's three headings. */
export const EFFECT_MENU = [
  ['entr', 'Entrance', ANIMATION_GALLERY.map(([effect, label]) => [effect, label])],
  ['emph', 'Emphasis', [['pulse', 'Pulse'], ['spin', 'Spin'], ['grow', 'Grow/Shrink']]],
  ['exit', 'Exit', [['appear', 'Disappear'], ['fade', 'Fade'], ['fly', 'Fly Out'], ['float', 'Float Out'], ['split', 'Split'], ['wipe', 'Wipe'], ['zoom', 'Zoom']]],
];
/** Effect Options for each effect that has a direction, as `[value, label]` — the engine's own values. */
export const ANIMATION_OPTIONS = {
  fly: [['bottom', 'From Bottom'], ['bottom-left', 'From Bottom-Left'], ['left', 'From Left'], ['top-left', 'From Top-Left'], ['top', 'From Top'], ['top-right', 'From Top-Right'], ['right', 'From Right'], ['bottom-right', 'From Bottom-Right']],
  wipe: [['bottom', 'From Bottom'], ['left', 'From Left'], ['right', 'From Right'], ['top', 'From Top']],
  split: [['vertical-out', 'Vertical Out'], ['vertical-in', 'Vertical In'], ['horizontal-out', 'Horizontal Out'], ['horizontal-in', 'Horizontal In']],
  float: [['up', 'Float Up'], ['down', 'Float Down']],
  spin: [['clockwise', 'Clockwise'], ['counterclockwise', 'Counterclockwise']],
};

/** Every element drawing a shape, or the members of a group. */
export function shapeNodes(root, id) {
  if (!root || id == null) return [];
  const safe = String(id).replace(/"/g, '');
  return [...root.querySelectorAll(`[data-shape="${safe}"], [data-groups~="${safe}"]`)];
}

/**
 * The list as the show plays it: the group that starts with the slide (or
 * null) and one group per click, each effect with its start in seconds
 * from the group's start — With Previous shares the start of the effect
 * before it, After Previous starts when everything before it has ended.
 */
export function sequence(animations = []) {
  const groups = [];
  let current = null;
  let subStart = 0;
  let subEnd = 0;
  for (const e of animations) {
    if (!current || e.trigger === 'onClick') {
      current = { auto: !groups.length && e.trigger !== 'onClick', effects: [], length: 0 };
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
 * Whether each animated shape shows once `step` clicks have played (the
 * group that starts with the slide counted in when `auto` is true): a
 * shape whose first effect is an entrance starts hidden, an entrance shows
 * it and an exit hides it.
 */
export function visibilityAt(animations = [], step = 0, auto = true) {
  const seq = sequence(animations);
  const shown = new Map();
  for (const e of animations) if (e.shapeId != null && !shown.has(e.shapeId)) shown.set(e.shapeId, e.kind !== 'entr');
  const played = [...(auto && seq.auto ? [seq.auto] : []), ...seq.clicks.slice(0, Math.max(0, step))];
  for (const g of played) {
    for (const e of g.effects) {
      if (e.kind === 'entr') shown.set(e.shapeId, true);
      else if (e.kind === 'exit') shown.set(e.shapeId, false);
    }
  }
  return shown;
}

/**
 * Put a slide's shapes where the sequence has them after `step` clicks, at
 * once — no animation: entering a slide, stepping back, or the presenter's
 * picture of where the show is.
 */
export function applyState(root, animations = [], step = 0, auto = true) {
  const shown = visibilityAt(animations, step, auto);
  for (const [id, visible] of shown) {
    for (const el of shapeNodes(root, id)) {
      for (const a of el.getAnimations?.() || []) a.cancel();
      el.style.visibility = visible ? '' : 'hidden';
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
function effectFrames(e, g, size) {
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
    const { frames, easing } = effectFrames(e, g, size);
    const duration = Math.max(1, Math.round((Number(e.duration) || 0) * 1000));
    for (const el of shapeNodes(root, e.shapeId)) {
      // An emphasis starts from where the shape is; an entrance or an exit
      // replaces what the shape held, from the state it had: hidden until
      // an entrance's own start (its delay included), shown until an exit's.
      if (e.kind === 'entr' || e.kind === 'exit') {
        for (const a of el.getAnimations()) a.cancel();
        el.style.visibility = e.kind === 'entr' ? 'hidden' : '';
      }
      animations.push(el.animate(frames, { duration, delay: Math.round(e.start * 1000), easing, fill: 'forwards' }));
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
