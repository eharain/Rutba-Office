// Transitions, played.
//
// The engine reads and writes `<p:transition>`; this plays one between two
// layers — the slide going and the slide coming — with the Web Animations
// API, so the browser runs it on the compositor and the show stays smooth
// however heavy the slide's drawing is. Each effect is keyframes on the two
// layers' transform, opacity, clip-path or mask, honouring the transition's
// own duration and direction, and the same player drives the slideshow and
// the ribbon's Preview on the stage.

/** The Transition to This Slide gallery, in PowerPoint's order and names. `shape` stands for circle, diamond, plus and zoom — PowerPoint's own "Shape". */
export const TRANSITION_GALLERY = [
  ['none', 'None', 'no effect: the slide simply replaces the last'],
  ['cut', 'Cut', 'the slide replaces the last at once'],
  ['fade', 'Fade', 'the slide fades in over the last'],
  ['push', 'Push', 'the slide pushes the last one off'],
  ['wipe', 'Wipe', 'the slide is wiped on across the last'],
  ['split', 'Split', 'the slide opens from the middle, or closes in from the edges'],
  ['pull', 'Uncover', 'the last slide slides away to show this one'],
  ['cover', 'Cover', 'the slide slides in over the last'],
  ['randomBar', 'Random Bars', 'the slide appears in random bars'],
  ['shape', 'Shape', 'the slide grows out of a circle, a diamond or a plus'],
  ['dissolve', 'Dissolve', 'the slide appears in small random squares'],
];

const SHAPE_TYPES = ['circle', 'diamond', 'plus', 'zoom'];
/** The gallery button a transition type presses. */
export const galleryKeyOf = (type) => (!type || type === 'none' ? 'none' : SHAPE_TYPES.includes(type) ? 'shape' : type);

const SIDES = [['u', 'From Bottom'], ['r', 'From Left'], ['l', 'From Right'], ['d', 'From Top']];
const EIGHT = [...SIDES, ['lu', 'From Bottom-Right'], ['ru', 'From Bottom-Left'], ['ld', 'From Top-Right'], ['rd', 'From Top-Left']];
/**
 * Effect Options, per gallery button: each `[value, label]`. A value for
 * Shape names the type too (`circle`, `zoom:in`), since PowerPoint's Shape
 * options switch between elements rather than set a direction.
 */
export const TRANSITION_OPTIONS = {
  cut: [['smooth', 'Cut'], ['black', 'Through Black']],
  fade: [['smooth', 'Smoothly'], ['black', 'Through Black']],
  push: SIDES,
  wipe: SIDES,
  pull: EIGHT,
  cover: EIGHT,
  split: [['vert-out', 'Vertical Out'], ['vert-in', 'Vertical In'], ['horz-out', 'Horizontal Out'], ['horz-in', 'Horizontal In']],
  randomBar: [['vert', 'Vertical'], ['horz', 'Horizontal']],
  shape: [['circle', 'Circle'], ['diamond', 'Diamond'], ['plus', 'Plus'], ['zoom:in', 'In'], ['zoom:out', 'Out']],
};

/** The option value a transition currently stands at, in TRANSITION_OPTIONS' terms. */
export function optionOf(transition) {
  if (!transition) return null;
  if (SHAPE_TYPES.includes(transition.type)) return transition.type === 'zoom' ? `zoom:${transition.direction || 'out'}` : transition.type;
  return transition.direction;
}

/** A transition as a sentence, for the ribbon's tips and the strip's star. */
export function describeTransition(t) {
  if (!t) return 'No transition';
  const key = galleryKeyOf(t.type);
  const name = TRANSITION_GALLERY.find(([k]) => k === key)?.[1] || t.type;
  const option = (TRANSITION_OPTIONS[key] || []).find(([v]) => v === optionOf(t))?.[1];
  const effect = t.known === false ? `${t.type} (played as ${t.fallback || 'fade'})` : key === 'none' ? 'No effect' : option && option !== name ? `${name}, ${option}` : name;
  const timing = key === 'none' || t.duration == null ? '' : `, ${Number(t.duration || 0).toFixed(2)} s`;
  const after = t.advanceAfter != null ? `; moves on after ${Number(t.advanceAfter).toFixed(2)} s` : '';
  return `${effect}${timing}${after}`;
}

const EASE = 'cubic-bezier(.45, 0, .25, 1)';
const VECTORS = { l: [-1, 0], u: [0, -1], r: [1, 0], d: [0, 1], lu: [-1, -1], ru: [1, -1], ld: [-1, 1], rd: [1, 1] };

/** A mask image of the cells of a grid whose own random threshold is below `p`. */
function cellMask(cols, rows, thresholds, p) {
  let d = '';
  for (let i = 0; i < thresholds.length; i++) {
    if (thresholds[i] >= p) continue;
    d += `M${i % cols} ${Math.floor(i / cols)}h1v1h-1z`;
  }
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${cols} ${rows}" preserveAspectRatio="none"><path d="${d || 'M0 0h0'}" fill="#000"/></svg>`;
  return `url("data:image/svg+xml,${encodeURIComponent(svg)}")`;
}

/** Stepped mask keyframes: the incoming slide appears cell by cell, in random order. */
function cellKeyframes(cols, rows, steps = 24) {
  const thresholds = Array.from({ length: cols * rows }, () => Math.random());
  const frames = [];
  for (let k = 0; k <= steps; k++) frames.push({ maskImage: cellMask(cols, rows, thresholds, k / steps), offset: k / steps });
  return frames;
}

/**
 * Keyframes for one transition: `{ from, to, top, black }` — the outgoing
 * and incoming layers' keyframes (either may be null for "stays as it
 * is"), which layer is drawn on top, and whether the stage shows black
 * behind them. Unknown effects play as their fallback.
 */
export function transitionFrames(transition) {
  const known = transition?.known !== false;
  const type = known ? transition?.type : transition?.fallback || 'fade';
  const dir = known ? transition?.direction : transition?.fallbackDirection ?? null;
  switch (type) {
    case 'fade':
      return dir === 'black'
        ? { from: [{ opacity: 1 }, { opacity: 0, offset: 0.5 }, { opacity: 0 }], to: [{ opacity: 0 }, { opacity: 0, offset: 0.5 }, { opacity: 1 }], top: 'to', black: true, easing: 'linear' }
        : { from: null, to: [{ opacity: 0 }, { opacity: 1 }], top: 'to', easing: 'linear' };
    case 'cut':
      return dir === 'black'
        ? { from: [{ opacity: 0 }, { opacity: 0 }], to: [{ opacity: 0 }, { opacity: 0, offset: 0.98 }, { opacity: 1 }], top: 'to', black: true, easing: 'linear' }
        : { from: null, to: [{ opacity: 1 }, { opacity: 1 }], top: 'to', easing: 'linear' };
    case 'push': {
      const [dx, dy] = VECTORS[dir] || VECTORS.l;
      const at = (k) => `translate(${dx * k * 100}%, ${dy * k * 100}%)`;
      return { from: [{ transform: at(0) }, { transform: at(1) }], to: [{ transform: at(-1) }, { transform: at(0) }], top: 'to' };
    }
    case 'cover': {
      const [dx, dy] = VECTORS[dir] || VECTORS.l;
      return { from: null, to: [{ transform: `translate(${-dx * 100}%, ${-dy * 100}%)` }, { transform: 'translate(0, 0)' }], top: 'to' };
    }
    case 'pull': {
      const [dx, dy] = VECTORS[dir] || VECTORS.l;
      return { from: [{ transform: 'translate(0, 0)' }, { transform: `translate(${dx * 100}%, ${dy * 100}%)` }], to: null, top: 'from' };
    }
    case 'wipe': {
      const start = { l: 'inset(0 0 0 100%)', r: 'inset(0 100% 0 0)', u: 'inset(100% 0 0 0)', d: 'inset(0 0 100% 0)' }[dir] || 'inset(0 0 0 100%)';
      return { from: null, to: [{ clipPath: start }, { clipPath: 'inset(0 0 0 0)' }], top: 'to' };
    }
    case 'split': {
      const [orient, way] = String(dir || 'horz-out').split('-');
      const closed = orient === 'vert' ? 'inset(0 50% 0 50%)' : 'inset(50% 0 50% 0)';
      return way === 'in'
        ? { from: [{ clipPath: 'inset(0 0 0 0)' }, { clipPath: closed }], to: null, top: 'from' }
        : { from: null, to: [{ clipPath: closed }, { clipPath: 'inset(0 0 0 0)' }], top: 'to' };
    }
    case 'randomBar':
      return { from: null, to: dir === 'vert' ? cellKeyframes(48, 1) : cellKeyframes(1, 36), top: 'to', easing: 'linear', mask: true };
    case 'dissolve':
      return { from: null, to: cellKeyframes(32, 18, 20), top: 'to', easing: 'linear', mask: true };
    case 'circle':
      return { from: null, to: [{ clipPath: 'circle(0% at 50% 50%)' }, { clipPath: 'circle(72% at 50% 50%)' }], top: 'to' };
    case 'diamond':
      return { from: null, to: [{ clipPath: 'polygon(50% 50%, 50% 50%, 50% 50%, 50% 50%)' }, { clipPath: 'polygon(50% -50%, 150% 50%, 50% 150%, -50% 50%)' }], top: 'to' };
    case 'plus': {
      const plus = (t, l) => `polygon(${[
        [50 - t, 50 - l], [50 + t, 50 - l], [50 + t, 50 - t], [50 + l, 50 - t], [50 + l, 50 + t], [50 + t, 50 + t],
        [50 + t, 50 + l], [50 - t, 50 + l], [50 - t, 50 + t], [50 - l, 50 + t], [50 - l, 50 - t], [50 - t, 50 - t],
      ].map(([x, y]) => `${x}% ${y}%`).join(', ')})`;
      return { from: null, to: [{ clipPath: plus(0, 0) }, { clipPath: plus(14, 60), offset: 0.45 }, { clipPath: plus(50, 100) }], top: 'to' };
    }
    case 'zoom':
      return dir === 'in'
        ? { from: null, to: [{ transform: 'scale(.08)', opacity: 0 }, { transform: 'scale(1)', opacity: 1 }], top: 'to' }
        : { from: [{ transform: 'scale(1)', opacity: 1 }, { transform: 'scale(2.4)', opacity: 0 }], to: null, top: 'from' };
    default:
      return null;
  }
}

/**
 * Play a transition between two layers stacked in one box. The layers
 * must already be in place (the incoming one showing); the promise
 * resolves when the effect is over — at once for None — and every style
 * the effect put on the layers is taken off again, whether it finished
 * or was cut short by `cancel()`.
 */
export function playTransition({ stage, fromEl, toEl, transition }) {
  const frames = transition ? transitionFrames(transition) : null;
  const ms = Math.max(0, Math.round(Number(transition?.duration || 0) * 1000));
  if (!frames || !ms || !toEl) return { finished: Promise.resolve(), cancel() {} };
  const animations = [];
  const touched = [];
  const hold = (el, style) => {
    if (!el) return;
    touched.push(el);
    Object.assign(el.style, style);
  };
  if (frames.black && stage) hold(stage, { background: '#000' });
  // The layer on top is drawn over the other; "from on top" is how Uncover and the closing splits work.
  hold(fromEl, { zIndex: frames.top === 'from' ? 2 : 1 });
  hold(toEl, { zIndex: frames.top === 'from' ? 1 : 2 });
  if (frames.mask) hold(toEl, { maskSize: '100% 100%', maskRepeat: 'no-repeat', webkitMaskSize: '100% 100%', webkitMaskRepeat: 'no-repeat' });
  const options = { duration: ms, easing: frames.easing || EASE, fill: 'both' };
  if (frames.from && fromEl) animations.push(fromEl.animate(frames.from, options));
  if (frames.to) animations.push(toEl.animate(frames.to, options));
  // Once, whichever comes first: the effect's own end, or a cancel — which
  // must put the layers straight at once, since the next transition is
  // about to style the same elements.
  let done = false;
  const cleanup = () => {
    if (done) return;
    done = true;
    for (const a of animations) { try { a.cancel(); } catch { /* gone */ } }
    for (const el of touched) {
      el.style.zIndex = '';
      el.style.background = '';
      el.style.maskSize = el.style.maskRepeat = el.style.webkitMaskSize = el.style.webkitMaskRepeat = '';
    }
  };
  const finished = Promise.all(animations.map((a) => a.finished.catch(() => null))).then(cleanup);
  return {
    finished,
    cancel: cleanup,
  };
}
