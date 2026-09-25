// Slide transitions — read from a slide's XML and written back into it.
//
// A transition is the `<p:transition>` element PresentationML puts after
// the slide's `<p:clrMapOvr>` and before its `<p:timing>`: one child naming
// the effect (`<p:push dir="u"/>`), the speed on the element itself, and
// how the slide advances (`advClick`, `advTm`). PowerPoint 2010 and later
// state an exact duration in `p14:dur`, which a 2007 reader would not
// understand, so it writes the element twice inside `mc:AlternateContent`:
// the `p14` choice with the duration, and a fallback with only the nearest
// of the three speeds. Both forms are read here; the second is written
// whenever the duration is not exactly one of the three speeds.
//
// Only the classic set — the effects a plain `<p:transition>` child can
// express — is authored. A PowerPoint 2010 effect (`<p14:prism>`,
// `<p14:vortex>`, …) is read as what it is, played as its own fallback,
// and kept byte for byte until somebody picks another transition.

const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const P14_NS = 'http://schemas.microsoft.com/office/powerpoint/2010/main';

/**
 * The classic transitions this writes, each with the direction PowerPoint's
 * gallery gives it on a fresh pick and its gallery duration in seconds.
 * The direction values are the XML's own, per effect: a side (`l`, `u`, `r`,
 * `d`) for push and wipe — the way the new slide travels, so "From Bottom"
 * is `u` — eight ways for cover and uncover (`pull`), `orient-dir` for
 * split, `horz`/`vert` for random bars, `in`/`out` for zoom, and
 * `smooth`/`black` for fade and cut (`thruBlk`).
 */
export const TRANSITIONS = {
  cut: { direction: 'smooth', duration: 0.1 },
  fade: { direction: 'smooth', duration: 0.7 },
  push: { direction: 'u', duration: 1 },
  wipe: { direction: 'l', duration: 1 },
  split: { direction: 'vert-out', duration: 1.5 },
  pull: { direction: 'l', duration: 1 },
  cover: { direction: 'l', duration: 1 },
  randomBar: { direction: 'vert', duration: 1 },
  circle: { direction: null, duration: 2 },
  diamond: { direction: null, duration: 2 },
  plus: { direction: null, duration: 2 },
  zoom: { direction: 'in', duration: 2 },
  dissolve: { direction: null, duration: 1.2 },
};

/** What an attribute left out means, per ECMA-376 — not what the gallery picks. */
const SCHEMA_DEFAULT_DIRECTION = {
  cut: 'smooth', fade: 'smooth', push: 'l', wipe: 'l', split: 'horz-out', pull: 'l', cover: 'l', randomBar: 'horz', zoom: 'out',
};

const SIDES = ['l', 'u', 'r', 'd'];
const EIGHT = ['l', 'u', 'r', 'd', 'lu', 'ru', 'ld', 'rd'];

/** Seconds ↔ the three speeds a 2007 reader knows. */
const SPEED_SECONDS = { fast: 0.5, med: 0.75, slow: 1 };
function speedFor(seconds) {
  if (seconds <= 0.5) return 'fast';
  if (seconds <= 0.75) return 'med';
  return 'slow';
}

/** The slide's own root-level tail: everything after `</p:cSld>`. */
function tailStart(xml) {
  const at = xml.indexOf('</p:cSld>');
  return at < 0 ? -1 : at + '</p:cSld>'.length;
}

/**
 * Where the slide's transition sits, as a range of its XML: the whole
 * `mc:AlternateContent` when PowerPoint wrapped it in one, the bare element
 * otherwise. Null when the slide has none.
 */
export function transitionRange(xml) {
  const from = tailStart(xml);
  if (from < 0) return null;
  const tail = xml.slice(from);
  // A wrapped one first: an AlternateContent in the tail whose content is a transition.
  const acRe = /<mc:AlternateContent\b[^>]*>([\s\S]*?)<\/mc:AlternateContent>/g;
  let m;
  while ((m = acRe.exec(tail))) {
    if (/<p:transition\b/.test(m[1])) return { start: from + m.index, end: from + m.index + m[0].length, wrapped: true };
  }
  const bare = /<p:transition\b[^>]*?\/>|<p:transition\b[^>]*>[\s\S]*?<\/p:transition>/.exec(tail);
  if (bare) return { start: from + bare.index, end: from + bare.index + bare[0].length, wrapped: false };
  return null;
}

/** One `<p:transition>` element's attributes and effect child. */
function readElement(el) {
  const open = /<p:transition\b([^>]*?)\/?>/.exec(el);
  const attrs = {};
  for (const a of (open?.[1] || '').matchAll(/([\w:]+)="([^"]*)"/g)) attrs[a[1]] = a[2];
  // The effect is the element's first child, whatever its namespace.
  const body = el.slice(open ? open.index + open[0].length : 0);
  // A sound (`p:sndAc`) or an extension list is not the effect.
  const inner = body.replace(/<\/p:transition>[\s\S]*$/, '').replace(/<p:sndAc\b[\s\S]*?<\/p:sndAc>|<p:extLst\b[\s\S]*?<\/p:extLst>/g, '');
  const child = /<([\w]+):([\w]+)\b([^>]*?)\/?>/.exec(inner);
  const sound = /<p:sndAc\b[\s\S]*?<\/p:sndAc>/.exec(body)?.[0] || null;
  const childAttrs = {};
  if (child) for (const a of child[3].matchAll(/([\w:]+)="([^"]*)"/g)) childAttrs[a[1]] = a[2];
  return { attrs, prefix: child?.[1] || null, name: child?.[2] || null, childAttrs, sound };
}

/** The direction in the model's terms, from an effect element's own attributes. */
function directionOf(name, a) {
  switch (name) {
    case 'fade':
    case 'cut':
      return a.thruBlk === '1' || a.thruBlk === 'true' ? 'black' : 'smooth';
    case 'push':
    case 'wipe':
    case 'pull':
    case 'cover':
      return a.dir || SCHEMA_DEFAULT_DIRECTION[name];
    case 'split':
      return `${a.orient || 'horz'}-${a.dir || 'out'}`;
    case 'randomBar':
      return a.dir || 'horz';
    case 'zoom':
      return a.dir || 'out';
    default:
      return a.dir || null;
  }
}

/**
 * The slide's transition, or null when it has none:
 * `{ type, direction, duration, advanceOnClick, advanceAfter, known, fallback }`.
 * `type` is the effect element's own name (`push`, `fade`, `prism`, …), or
 * `none` for a transition that only says how the slide advances. `known`
 * is false for an effect this does not author (a PowerPoint 2010 one);
 * `fallback` then names the classic effect PowerPoint wrote beside it,
 * which is what the show plays. Durations are seconds; `advanceAfter` is
 * null unless the slide moves on by itself.
 */
export function readTransition(xml) {
  const range = transitionRange(xml);
  if (!range) return null;
  const block = xml.slice(range.start, range.end);
  let chosen = block;
  let fallback = null;
  if (range.wrapped) {
    const choice = /<mc:Choice\b[^>]*>([\s\S]*?)<\/mc:Choice>/.exec(block);
    const fb = /<mc:Fallback\b[^>]*>([\s\S]*?)<\/mc:Fallback>/.exec(block);
    chosen = choice && /<p:transition\b/.test(choice[1]) ? choice[1] : fb?.[1] || block;
    if (fb) fallback = readElement(fb[1]);
  }
  const el = readElement(chosen);
  const known = !el.name || (el.prefix === 'p' && el.name in TRANSITIONS);
  const type = el.name ? el.name : 'none';
  const durMs = el.attrs['p14:dur'] != null ? Number(el.attrs['p14:dur']) : null;
  const duration = Number.isFinite(durMs) && durMs != null ? durMs / 1000 : SPEED_SECONDS[el.attrs.spd || 'fast'] ?? 0.5;
  const advTm = el.attrs.advTm != null ? Number(el.attrs.advTm) : null;
  const fallbackType = fallback?.name && fallback.prefix === 'p' && fallback.name in TRANSITIONS ? fallback.name : null;
  return {
    type,
    direction: el.name ? directionOf(el.name, el.childAttrs) : null,
    duration: Math.round(duration * 1000) / 1000,
    advanceOnClick: !(el.attrs.advClick === '0' || el.attrs.advClick === 'false'),
    advanceAfter: Number.isFinite(advTm) ? advTm / 1000 : null,
    known,
    fallback: known ? null : fallbackType || 'fade',
    fallbackDirection: known || !fallback?.name ? null : directionOf(fallback.name, fallback.childAttrs),
    // A transition sound is not authored here, but one already in the file
    // rides along whenever the transition is rewritten.
    sound: el.sound,
  };
}

/** The effect element for a classic transition, from the model's direction. */
function effectXml(type, direction) {
  const d = direction ?? TRANSITIONS[type]?.direction ?? null;
  switch (type) {
    case 'fade':
    case 'cut':
      return d === 'black' ? `<p:${type} thruBlk="1"/>` : `<p:${type}/>`;
    case 'push':
    case 'wipe': {
      const side = SIDES.includes(d) ? d : 'l';
      return side === 'l' ? `<p:${type}/>` : `<p:${type} dir="${side}"/>`;
    }
    case 'pull':
    case 'cover': {
      const way = EIGHT.includes(d) ? d : 'l';
      return way === 'l' ? `<p:${type}/>` : `<p:${type} dir="${way}"/>`;
    }
    case 'split': {
      const [orient, dir] = String(d || 'horz-out').split('-');
      const bits = [];
      if (orient === 'vert') bits.push('orient="vert"');
      if (dir === 'in') bits.push('dir="in"');
      return `<p:split${bits.length ? ' ' + bits.join(' ') : ''}/>`;
    }
    case 'randomBar':
      return d === 'vert' ? '<p:randomBar dir="vert"/>' : '<p:randomBar/>';
    case 'zoom':
      return d === 'in' ? '<p:zoom dir="in"/>' : '<p:zoom/>';
    case 'circle':
    case 'diamond':
    case 'plus':
    case 'dissolve':
      return `<p:${type}/>`;
    default:
      throw new Error(`"${type}" is not a transition this writes`);
  }
}

/**
 * A transition as XML, the way PowerPoint writes it: the bare element when
 * the duration is one of the three speeds, the `mc:AlternateContent` pair
 * (a `p14:dur` choice and a speed-only fallback) when it is not.
 */
export function transitionXml({ type = 'none', direction = null, duration = null, advanceOnClick = true, advanceAfter = null, sound = null } = {}) {
  const effect = (type && type !== 'none' ? effectXml(type, direction) : '') + (sound || '');
  const seconds = duration != null && Number.isFinite(Number(duration)) ? Math.max(0.01, Number(duration)) : TRANSITIONS[type]?.duration ?? 0.75;
  const advance = `${advanceOnClick === false ? ' advClick="0"' : ''}${advanceAfter != null && Number.isFinite(Number(advanceAfter)) ? ` advTm="${Math.max(0, Math.round(Number(advanceAfter) * 1000))}"` : ''}`;
  if (!effect) return `<p:transition${advance}/>`;
  const spd = speedFor(seconds);
  const exact = Math.abs(SPEED_SECONDS[spd] - seconds) < 0.0005;
  if (exact) return `<p:transition spd="${spd}"${advance}>${effect}</p:transition>`;
  return (
    `<mc:AlternateContent xmlns:mc="${MC_NS}">` +
    `<mc:Choice xmlns:p14="${P14_NS}" Requires="p14"><p:transition spd="${spd}" p14:dur="${Math.round(seconds * 1000)}"${advance}>${effect}</p:transition></mc:Choice>` +
    `<mc:Fallback><p:transition spd="${spd}"${advance}>${effect}</p:transition></mc:Fallback>` +
    `</mc:AlternateContent>`
  );
}

/**
 * A transition this does not author (a PowerPoint 2010 effect) given new
 * timing in place: every `<p:transition>` tag in the block — the choice's
 * and the fallback's — gets the new speed, duration and advance settings,
 * and the effect elements themselves are left exactly as they were.
 */
function retimed(block, { duration, advanceOnClick, advanceAfter }) {
  return block.replace(/<p:transition\b([^>]*?)(\/?)>/g, (m, attrs, selfClose) => {
    let next = attrs;
    const set = (name, value) => {
      const re = new RegExp(`\\s${name.replace(':', '\\:')}="[^"]*"`);
      next = next.replace(re, '');
      if (value != null) next += ` ${name}="${value}"`;
    };
    if (duration != null) {
      set('spd', speedFor(duration));
      if (/p14:dur=/.test(attrs)) set('p14:dur', String(Math.round(duration * 1000)));
    }
    if (advanceOnClick !== undefined) set('advClick', advanceOnClick === false ? '0' : null);
    if (advanceAfter !== undefined) set('advTm', advanceAfter == null ? null : String(Math.round(Number(advanceAfter) * 1000)));
    return `<p:transition${next}${selfClose}>`;
  });
}

/**
 * The slide's XML with its transition set, changed or taken away.
 *
 * `spec` null takes the transition off entirely. Otherwise it is merged
 * over what the slide has: a new `type` brings that effect's gallery
 * direction and duration unless they are given too; leaving `type` out
 * keeps the effect and changes only what is named. `type: 'none'` keeps
 * the advance settings with no effect — PowerPoint's "None" — and a
 * transition left with nothing to say is removed rather than written empty.
 */
export function withTransition(xml, spec) {
  const range = transitionRange(xml);
  const current = range ? readTransition(xml) : null;
  const without = range ? xml.slice(0, range.start) + xml.slice(range.end) : xml;
  if (spec == null) return without;

  let block;
  if (current && !current.known && spec.type === undefined) {
    // Timing only, on an effect this does not write: kept, re-timed in place.
    block = retimed(xml.slice(range.start, range.end), spec);
  } else {
    const typeChanged = spec.type !== undefined && spec.type !== current?.type;
    const type = spec.type ?? (current?.known ? current.type : 'none');
    if (type !== 'none' && !(type in TRANSITIONS)) throw new Error(`"${type}" is not a transition this writes`);
    const base = typeChanged || !current?.known
      ? { direction: TRANSITIONS[type]?.direction ?? null, duration: TRANSITIONS[type]?.duration ?? null }
      : { direction: current.direction, duration: current.duration };
    const next = {
      type,
      direction: spec.direction !== undefined ? spec.direction : base.direction,
      duration: spec.duration !== undefined && spec.duration !== null ? spec.duration : base.duration,
      advanceOnClick: spec.advanceOnClick !== undefined ? spec.advanceOnClick : current?.advanceOnClick ?? true,
      advanceAfter: spec.advanceAfter !== undefined ? spec.advanceAfter : current?.advanceAfter ?? null,
      sound: current?.sound || null,
    };
    if (type === 'none' && next.advanceOnClick !== false && next.advanceAfter == null && !next.sound) return without;
    block = transitionXml(next);
  }
  return insertTransition(without, block);
}

/** A transition block put where the schema wants it: after clrMapOvr (or cSld), before timing. */
export function insertTransition(xml, block) {
  const ovr = xml.indexOf('</p:clrMapOvr>');
  const tail = tailStart(xml);
  const at = ovr >= 0 ? ovr + '</p:clrMapOvr>'.length : tail;
  if (at < 0) throw new Error('the slide has no cSld to put a transition after');
  return xml.slice(0, at) + block + xml.slice(at);
}

/** The transition block itself, verbatim — what Apply To All copies. */
export function transitionBlock(xml) {
  const range = transitionRange(xml);
  return range ? xml.slice(range.start, range.end) : null;
}
