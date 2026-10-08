// Animations — a slide's main sequence, read from its `<p:timing>` and
// written back the way PowerPoint writes it.
//
// PowerPoint keeps a slide's animations as a time-node tree:
//
//   p:timing / p:tnLst / p:par / p:cTn nodeType="tmRoot"
//     p:seq / p:cTn nodeType="mainSeq"           the click-by-click sequence
//       p:par / p:cTn  (delay="indefinite")      one per click
//         p:par / p:cTn  (delay = start)         effects that start together
//           p:par / p:cTn presetClass presetID   one effect: its behaviours
//                                                (p:set, p:animEffect, p:anim…)
//   p:bldLst / p:bldP                            which shapes build
//
// "With Previous" puts an effect in the same inner group as the one before
// it; "After Previous" starts a new inner group whose delay is the time the
// previous group ends; "On Click" starts a new click group. So the list the
// Animation Pane shows is the tree read in order, and writing a changed
// list back is rebuilding the three levels from it.
//
// A motion path is written as PowerPoint writes one drawn by hand — a
// custom path, `p:animMotion` moving the shape from where it stands along
// a path in fractions of the slide — and a text shape's effect can be
// split into one per paragraph (Effect Options → By Paragraph), each
// aimed at its paragraph (`p:txEl/p:pRg`) with the build list saying so.
//
// Everything this does not model is kept: an effect it does not know (a
// media call, anything from a newer PowerPoint) stays byte
// for byte and moves with the list; interactive sequences (triggers), media
// nodes and the rest of the tree are never touched; and the ids every time
// node carries are renumbered once, at the end, with the references to them
// (`p:tn val`) following.

const ATTR_RE = /([\w:.-]+)\s*=\s*"([^"]*)"/g;

/* ── a small tree over the XML, keeping every node's exact range ────────── */

/**
 * The elements of `xml` as a tree, each node with its name, attributes and
 * the offsets of its opening tag, contents and end — so a node can be read
 * as a tree and cut out, kept or replaced as text.
 */
export function rangeTree(xml) {
  const root = { name: '#root', attrs: {}, start: 0, end: xml.length, innerStart: 0, innerEnd: xml.length, children: [] };
  const stack = [root];
  const re = /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[[\s\S]*?\]\]>|<(\/?)([\w:.-]+)((?:\s+[\w:.-]+\s*=\s*"[^"]*")*)\s*(\/?)>/g;
  let m;
  while ((m = re.exec(xml))) {
    if (!m[2]) continue;
    const top = stack[stack.length - 1];
    if (m[1] === '/') {
      if (top.name === m[2]) {
        top.innerEnd = m.index;
        top.end = m.index + m[0].length;
        stack.pop();
      }
      continue;
    }
    const attrs = {};
    for (const a of m[3].matchAll(ATTR_RE)) attrs[a[1]] = a[2];
    const node = { name: m[2], attrs, start: m.index, openEnd: m.index + m[0].length, innerStart: m.index + m[0].length, innerEnd: m.index + m[0].length, end: m.index + m[0].length, children: [], parent: top };
    top.children.push(node);
    if (m[4] !== '/') stack.push(node);
  }
  return root;
}

/** Whether the time-node tags in `xml` open and close in equal number — a tree this can rebuild. */
function balanced(xml) {
  return ['p:seq', 'p:par', 'p:cTn', 'p:childTnLst'].every((tag) => {
    const opens = (xml.match(new RegExp(`<${tag}\\b(?:[^>"]|"[^"]*")*?(?<!/)>`, 'g')) || []).length;
    const closes = (xml.match(new RegExp(`</${tag}>`, 'g')) || []).length;
    return opens === closes;
  });
}

const kid = (node, name) => node?.children.find((c) => c.name === name) || null;
const kidsOf = (node, name) => (node ? node.children.filter((c) => c.name === name) : []);
function* walk(node) {
  for (const c of node.children) {
    yield c;
    yield* walk(c);
  }
}

/* ── the effects PowerPoint's gallery offers, and their preset numbers ──── */

/**
 * Every effect this writes: its preset class and id, the direction it gets
 * on a fresh pick, and its gallery duration in seconds. Directions are the
 * side an entrance comes from (or an exit goes to), the way the Effect
 * Options menu names them.
 */
export const EFFECTS = {
  entr: {
    appear: { id: 1, duration: 0, name: 'Appear' },
    fade: { id: 10, duration: 0.5, name: 'Fade' },
    fly: { id: 2, duration: 0.5, direction: 'bottom', name: 'Fly In' },
    float: { id: 42, duration: 1, direction: 'up', name: 'Float In' },
    split: { id: 16, duration: 0.5, direction: 'vertical-out', name: 'Split' },
    wipe: { id: 22, duration: 0.5, direction: 'bottom', name: 'Wipe' },
    zoom: { id: 53, duration: 0.5, name: 'Zoom' },
  },
  emph: {
    pulse: { id: 26, duration: 0.5, name: 'Pulse' },
    spin: { id: 8, duration: 2, direction: 'clockwise', name: 'Spin' },
    grow: { id: 6, duration: 2, name: 'Grow/Shrink' },
  },
  exit: {
    appear: { id: 1, duration: 0, name: 'Disappear' },
    fade: { id: 10, duration: 0.5, name: 'Fade' },
    fly: { id: 2, duration: 0.5, direction: 'bottom', name: 'Fly Out' },
    float: { id: 42, duration: 1, direction: 'down', name: 'Float Out' },
    split: { id: 16, duration: 0.5, direction: 'vertical-in', name: 'Split' },
    wipe: { id: 22, duration: 0.5, direction: 'bottom', name: 'Wipe' },
    zoom: { id: 53, duration: 0.5, name: 'Zoom' },
  },
  // Animations → Motion Paths: every path is written as a custom one, its
  // shape in `path` (motionPath below gives the gallery's).
  path: {
    custom: { id: 0, duration: 2, name: 'Custom Path' },
  },
};

/**
 * The Motion Paths gallery: each path from the shape's centre, in PowerPoint's
 * own form — `M`, `L`, `C` and `Z` in fractions of the slide's width (x)
 * and height (y), y downwards, ending in `E` — `aspect` being the slide's
 * height over its width, so a circle is round on the slide.
 */
export const MOTION_PATHS = [
  ['down', 'Down'], ['up', 'Up'], ['left', 'Left'], ['right', 'Right'],
  ['arcDown', 'Arc Down'], ['arcUp', 'Arc Up'], ['turnDown', 'Turn Down'], ['turnUp', 'Turn Up'],
  ['circle', 'Circle'], ['square', 'Square'], ['loop', 'Loop de Loop'],
];
export function motionPath(name, aspect = 9 / 16) {
  const n = (v) => String(Math.round(v * 100000) / 100000);
  // Points given in heights; x scaled to the width.
  const pt = (x, y) => `${n(x * aspect)} ${n(y)}`;
  const k = 0.5523;
  const r = 0.125;
  switch (name) {
    case 'down': return `M 0 0 L ${pt(0, 0.25)} E`;
    case 'up': return `M 0 0 L ${pt(0, -0.25)} E`;
    case 'left': return `M 0 0 L ${pt(-0.3, 0)} E`;
    case 'right': return `M 0 0 L ${pt(0.3, 0)} E`;
    case 'arcDown': return `M 0 0 C ${pt(0, 0.16)} ${pt(0.3, 0.16)} ${pt(0.3, 0)} E`;
    case 'arcUp': return `M 0 0 C ${pt(0, -0.16)} ${pt(0.3, -0.16)} ${pt(0.3, 0)} E`;
    case 'turnDown': return `M 0 0 L ${pt(0.15, 0)} C ${pt(0.205, 0)} ${pt(0.25, 0.045)} ${pt(0.25, 0.1)} L ${pt(0.25, 0.25)} E`;
    case 'turnUp': return `M 0 0 L ${pt(0.15, 0)} C ${pt(0.205, 0)} ${pt(0.25, -0.045)} ${pt(0.25, -0.1)} L ${pt(0.25, -0.25)} E`;
    case 'circle': return `M 0 0 C ${pt(k * r, 0)} ${pt(r, r - k * r)} ${pt(r, r)} C ${pt(r, r + k * r)} ${pt(k * r, 2 * r)} ${pt(0, 2 * r)} C ${pt(-k * r, 2 * r)} ${pt(-r, r + k * r)} ${pt(-r, r)} C ${pt(-r, r - k * r)} ${pt(-k * r, 0)} ${pt(0, 0)} Z E`;
    case 'square': return `M 0 0 L ${pt(0.2, 0)} L ${pt(0.2, 0.2)} L ${pt(0, 0.2)} Z E`;
    case 'loop': return `M 0 0 C ${pt(0.15, 0)} ${pt(0.25, -0.05)} ${pt(0.25, -0.15)} C ${pt(0.25, -0.25)} ${pt(0.1, -0.25)} ${pt(0.1, -0.15)} C ${pt(0.1, -0.05)} ${pt(0.2, 0)} ${pt(0.35, 0)} E`;
    default: throw new Error(`"${name}" is not a motion path in the gallery`);
  }
}

/** Fly and Wipe: presetSubtype bits — 1 top, 2 right, 4 bottom, 8 left. */
const SIDE_BITS = { top: 1, right: 2, bottom: 4, left: 8, 'top-right': 3, 'bottom-right': 6, 'top-left': 9, 'bottom-left': 12 };
const BITS_SIDE = Object.fromEntries(Object.entries(SIDE_BITS).map(([k, v]) => [v, k]));
/** Split: presetSubtype and the barn filter it draws with. */
const SPLIT = {
  'vertical-out': { subtype: 37, filter: 'barn(outVertical)' },
  'horizontal-out': { subtype: 42, filter: 'barn(outHorizontal)' },
  'vertical-in': { subtype: 21, filter: 'barn(inVertical)' },
  'horizontal-in': { subtype: 26, filter: 'barn(inHorizontal)' },
};
/** Wipe: the side it comes from, and the filter's own word — which way the edge moves. */
const WIPE_FILTER = { top: 'wipe(down)', right: 'wipe(left)', bottom: 'wipe(up)', left: 'wipe(right)' };

/** The Effect Options a written effect offers, as `[value, label]`. */
export const EFFECT_DIRECTIONS = {
  fly: [['bottom', 'From Bottom'], ['bottom-left', 'From Bottom-Left'], ['left', 'From Left'], ['top-left', 'From Top-Left'], ['top', 'From Top'], ['top-right', 'From Top-Right'], ['right', 'From Right'], ['bottom-right', 'From Bottom-Right']],
  wipe: [['bottom', 'From Bottom'], ['left', 'From Left'], ['right', 'From Right'], ['top', 'From Top']],
  split: [['vertical-out', 'Vertical Out'], ['vertical-in', 'Vertical In'], ['horizontal-out', 'Horizontal Out'], ['horizontal-in', 'Horizontal In']],
  float: [['up', 'Float Up'], ['down', 'Float Down']],
  spin: [['clockwise', 'Clockwise'], ['counterclockwise', 'Counterclockwise']],
};

/** What an effect par is: its class, effect and direction, from its preset numbers and behaviours. */
function identify(ctnAttrs, xml) {
  const cls = ctnAttrs.presetClass || null;
  const id = Number(ctnAttrs.presetID);
  const sub = Number(ctnAttrs.presetSubtype || 0);
  const kind = cls === 'entr' || cls === 'exit' || cls === 'emph' ? cls : cls === 'path' ? 'path' : cls === 'mediacall' ? 'media' : cls || 'other';
  const out = { kind, effect: null, direction: null };
  if (kind === 'entr' || kind === 'exit') {
    if (id === 1) out.effect = 'appear';
    else if (id === 10) out.effect = 'fade';
    else if (id === 2) { out.effect = 'fly'; out.direction = BITS_SIDE[sub] || 'bottom'; }
    else if (id === 42 || id === 47) { out.effect = 'float'; out.direction = (id === 42) === (kind === 'entr') ? 'up' : 'down'; }
    else if (id === 16) {
      out.effect = 'split';
      const filter = /filter="(barn\([^)]*\))"/.exec(xml)?.[1];
      out.direction = Object.entries(SPLIT).find(([, v]) => v.filter === filter || v.subtype === sub)?.[0] || 'vertical-out';
    } else if (id === 22) { out.effect = 'wipe'; out.direction = BITS_SIDE[sub] || 'bottom'; }
    else if (id === 53 || id === 23) out.effect = 'zoom';
  } else if (kind === 'path') {
    out.effect = 'custom';
    out.path = /<p:animMotion\b[^>]*\spath="([^"]*)"/.exec(xml)?.[1] ?? null;
  } else if (kind === 'emph') {
    if (id === 26) out.effect = 'pulse';
    else if (id === 8) {
      out.effect = 'spin';
      const by = Number(/<p:animRot\b[^>]*\bby="(-?\d+)"/.exec(xml)?.[1] || 21600000);
      out.direction = by < 0 ? 'counterclockwise' : 'clockwise';
    } else if (id === 6) out.effect = 'grow';
  }
  return out;
}

/* ── reading ─────────────────────────────────────────────────────────────── */

/** The slide's `<p:timing>` element's range, or null. */
export function timingRange(xml) {
  const at = xml.indexOf('<p:timing');
  if (at < 0) return null;
  const close = xml.indexOf('</p:timing>', at);
  if (close < 0) {
    const self = /<p:timing\b[^>]*\/>/.exec(xml.slice(at));
    return self ? { start: at, end: at + self[0].length } : null;
  }
  return { start: at, end: close + '</p:timing>'.length };
}

/**
 * The parts of a timing tree this edits: the tmRoot, the main sequence and
 * its click groups, each effect with its trigger, delay and duration.
 */
function readTree(timingXml) {
  const tree = rangeTree(timingXml);
  const timing = kid(tree, 'p:timing');
  const tnLst = kid(timing, 'p:tnLst');
  const rootPar = kid(tnLst, 'p:par');
  const rootCtn = kid(rootPar, 'p:cTn');
  const rootKids = kid(rootCtn, 'p:childTnLst');
  const seqs = kidsOf(rootKids, 'p:seq');
  const mainSeq = seqs.find((s) => kid(s, 'p:cTn')?.attrs.nodeType === 'mainSeq') || null;
  const mainCtn = kid(mainSeq, 'p:cTn');
  const mainKids = kid(mainCtn, 'p:childTnLst');
  // Animations → Trigger: an interactive sequence started by a click on a
  // shape. One started any other way (a media bookmark) is kept, not read.
  const interactive = [];
  for (const seq of seqs) {
    const ctn = kid(seq, 'p:cTn');
    if (ctn?.attrs.nodeType !== 'interactiveSeq') continue;
    const cond = kidsOf(kid(ctn, 'p:stCondLst'), 'p:cond').find((c) => c.attrs.evt === 'onClick');
    const spTgt = cond ? [...walk(cond)].find((n) => n.name === 'p:spTgt') : null;
    // Only one whose tags balance is read; anything else stays as it is.
    if (spTgt?.attrs.spid != null && balanced(timingXml.slice(seq.start, seq.end))) interactive.push({ seq, kids: kid(ctn, 'p:childTnLst'), spid: String(spTgt.attrs.spid) });
  }
  const effects = [];
  let editable = true;
  collect(mainKids, null);
  for (const s of interactive) collect(s.kids, s.spid);
  return { tree, timing, rootCtn, rootKids, mainSeq, mainCtn, mainKids, interactive, effects, editable };

  // The effects of one sequence's click groups, in order. In a trigger's
  // sequence every group is started by a click on the trigger.
  function collect(seqKids, triggerShape) {
  let click = 0;
  kidsOf(seqKids, 'p:par').forEach((groupPar, gi) => {
    const groupCtn = kid(groupPar, 'p:cTn');
    const conds = kidsOf(kid(groupCtn, 'p:stCondLst'), 'p:cond');
    // A group that starts with the slide (an After/With Previous first
    // effect) carries an onBegin condition beside "indefinite".
    const auto = !triggerShape && gi === 0 && (conds.some((c) => c.attrs.evt === 'onBegin') || conds.every((c) => c.attrs.delay !== 'indefinite'));
    if (!auto) click += 1;
    const subs = kidsOf(kid(groupCtn, 'p:childTnLst'), 'p:par');
    if (!subs.length) editable = false;
    subs.forEach((subPar, si) => {
      const subCtn = kid(subPar, 'p:cTn');
      const pars = kidsOf(kid(subCtn, 'p:childTnLst'), 'p:par');
      if (!pars.length) editable = false;
      pars.forEach((effectPar, ei) => {
        const ctn = kid(effectPar, 'p:cTn');
        if (!ctn) { editable = false; return; }
        const xml = timingXml.slice(effectPar.start, effectPar.end);
        const firstCond = kid(kid(ctn, 'p:stCondLst'), 'p:cond');
        const delay = Number(firstCond?.attrs.delay);
        let longest = 0;
        let anyTimed = false;
        for (const n of walk(ctn)) {
          if (n.name !== 'p:cTn' || n.attrs.dur == null) continue;
          const d = Number(n.attrs.dur);
          if (!Number.isFinite(d)) continue;
          // A p:set lasts a millisecond; it is not what Duration shows.
          if (d > 1) anyTimed = true;
          longest = Math.max(longest, d);
        }
        const spTgt = [...walk(ctn)].find((n) => n.name === 'p:spTgt');
        const pRg = spTgt ? [...walk(spTgt)].find((n) => n.name === 'p:pRg') : null;
        const nodeType = ctn.attrs.nodeType || null;
        const trigger = nodeType === 'clickEffect' ? 'onClick'
          : nodeType === 'withEffect' ? 'withPrevious'
          : nodeType === 'afterEffect' ? 'afterPrevious'
          : si === 0 && ei === 0 ? (auto ? 'withPrevious' : 'onClick')
          : ei === 0 ? 'afterPrevious' : 'withPrevious';
        const kind = identify(ctn.attrs, xml);
        effects.push({
          xml,
          group: auto ? 0 : click,
          trigger,
          delay: Number.isFinite(delay) ? delay : 0,
          duration: anyTimed ? longest : 0,
          shapeId: spTgt?.attrs.spid ?? null,
          paragraph: pRg ? Number(pRg.attrs.st) : null,
          grpId: ctn.attrs.grpId != null ? Number(ctn.attrs.grpId) : null,
          presetId: ctn.attrs.presetID != null ? Number(ctn.attrs.presetID) : null,
          presetSubtype: ctn.attrs.presetSubtype != null ? Number(ctn.attrs.presetSubtype) : null,
          triggerShape,
          ...kind,
        });
      });
    });
  });
  }
}

/**
 * The slide's main sequence, in the order the Animation Pane lists it:
 * `[{ index, shapeId, kind, effect, direction, trigger, duration, delay,
 * group, paragraph, known }]` — `kind` is entr, emph, exit, path, media or
 * other; `effect` one of the EFFECTS names (null for one this does not
 * know); `group` the click that starts it (0 for one that starts with the
 * slide); durations and delays in seconds; `known` false for an effect this
 * keeps but would not write.
 */
export function readAnimations(xml) {
  const range = timingRange(xml);
  if (!range) return [];
  const { effects } = readTree(xml.slice(range.start, range.end));
  return effects.map((e, index) => publicEntry(e, index));
}

function publicEntry(e, index) {
  const known = writes(e);
  return {
    index,
    shapeId: e.shapeId,
    kind: e.kind,
    effect: e.effect,
    name: e.kind === 'path' ? 'Motion Path' : e.effect ? EFFECTS[e.kind]?.[e.effect]?.name || e.effect : e.kind === 'media' ? 'Play' : 'Custom',
    path: e.path ?? null,
    direction: e.direction,
    trigger: e.trigger,
    duration: Math.round(e.duration) / 1000,
    delay: Math.round(e.delay) / 1000,
    group: e.group,
    paragraph: e.paragraph,
    presetId: e.presetId,
    // Animations → Trigger: the shape a click on which starts it, or null for the slide's own sequence.
    triggerShape: e.triggerShape ?? null,
    known,
  };
}
/**
 * Whether this would write the effect exactly as it is: one of its own
 * presets under that preset's own number — not a legacy variant (Basic
 * Zoom, 23) it reads and plays but would turn into its modern cousin.
 */
function writes(e) {
  if (e.kind === 'path') return false;
  const spec = e.effect ? EFFECTS[e.kind]?.[e.effect] : null;
  if (!spec || e.presetId == null) return false;
  return e.effect === 'float' ? e.presetId === 42 || e.presetId === 47 : spec.id === e.presetId;
}

/* ── writing ─────────────────────────────────────────────────────────────── */

let tokenSeq = 0;
/** A placeholder time-node id, unique within one rewrite; renumbered at the end. */
const token = () => `@${++tokenSeq}`;

function target(spid, paragraph) {
  return paragraph == null
    ? `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>`
    : `<p:tgtEl><p:spTgt spid="${spid}"><p:txEl><p:pRg st="${paragraph}" end="${paragraph}"/></p:txEl></p:spTgt></p:tgtEl>`;
}

function setVisibility(tgt, value, delay = 0) {
  return `<p:set><p:cBhvr><p:cTn id="${token()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst></p:cTn>${tgt}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="${value}"/></p:to></p:set>`;
}
function animEffect(tgt, transition, filter, dur) {
  return `<p:animEffect transition="${transition}" filter="${filter}"><p:cBhvr><p:cTn id="${token()}" dur="${dur}"/>${tgt}</p:cBhvr></p:animEffect>`;
}
function animProp(tgt, attr, from, to, dur, { additive = false, fromFloat = false } = {}) {
  const val = (v, float) => (float ? `<p:fltVal val="${v}"/>` : `<p:strVal val="${v}"/>`);
  return `<p:anim calcmode="lin" valueType="num"><p:cBhvr${additive ? ' additive="base"' : ''}><p:cTn id="${token()}" dur="${dur}" fill="hold"/>${tgt}<p:attrNameLst><p:attrName>${attr}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val>${val(from, fromFloat)}</p:val></p:tav><p:tav tm="100000"><p:val>${val(to, false)}</p:val></p:tav></p:tavLst></p:anim>`;
}

/** Where Fly In starts (or Fly Out ends): the shape's centre just off the slide on that side. */
function offSlide(side) {
  const x = side.includes('left') ? '0-#ppt_w/2' : side.includes('right') ? '1+#ppt_w/2' : '#ppt_x';
  const y = side.includes('top') ? '0-#ppt_h/2' : side.includes('bottom') ? '1+#ppt_h/2' : '#ppt_y';
  return { x, y };
}

/**
 * One effect as PowerPoint writes it: the effect's own `p:par` and `p:cTn`
 * (preset class, id and subtype, grpId, node type and delay) and its
 * behaviours — a visibility set, an animEffect filter, ppt_x/ppt_y/ppt_w/
 * ppt_h animations, a rotation or a scale.
 */
export function effectXml({ kind, effect, direction = null, duration, delay = 0, spid, paragraph = null, grpId = 0, nodeType = 'clickEffect', path = null }) {
  const spec = EFFECTS[kind]?.[effect];
  if (!spec) throw new Error(`"${kind} ${effect}" is not an animation this writes`);
  const dir = direction ?? spec.direction ?? null;
  const dur = Math.max(1, Math.round((duration ?? spec.duration) * 1000));
  const tgt = target(spid, paragraph);
  if (kind === 'path') {
    // A custom path, smooth at its start and end, the shape held where it ends.
    const d = String(path || '').trim();
    if (!/^[Mm]\s/.test(d)) throw new Error('A motion path needs a path to follow');
    return (
      `<p:par><p:cTn id="${token()}" presetID="0" presetClass="path" presetSubtype="0" accel="50000" decel="50000" fill="hold" grpId="${grpId}" nodeType="${nodeType}">` +
      `<p:stCondLst><p:cond delay="${Math.max(0, Math.round(delay))}"/></p:stCondLst><p:childTnLst>` +
      `<p:animMotion origin="layout" path="${d.replace(/"/g, '')}" pathEditMode="relative"><p:cBhvr><p:cTn id="${token()}" dur="${dur}" fill="hold"/>${tgt}<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr></p:animMotion>` +
      `</p:childTnLst></p:cTn></p:par>`
    );
  }
  const entering = kind === 'entr';
  let subtype = 0;
  let presetId = spec.id;
  let body = '';
  switch (effect) {
    case 'appear':
      body = setVisibility(tgt, entering ? 'visible' : 'hidden');
      break;
    case 'fade':
      body = entering
        ? setVisibility(tgt, 'visible') + animEffect(tgt, 'in', 'fade', dur)
        : animEffect(tgt, 'out', 'fade', dur) + setVisibility(tgt, 'hidden', dur - 1);
      break;
    case 'fly': {
      const side = SIDE_BITS[dir] ? dir : 'bottom';
      subtype = SIDE_BITS[side];
      const off = offSlide(side);
      body = entering
        ? setVisibility(tgt, 'visible') + animProp(tgt, 'ppt_x', off.x, '#ppt_x', dur, { additive: true }) + animProp(tgt, 'ppt_y', off.y, '#ppt_y', dur, { additive: true })
        : animProp(tgt, 'ppt_x', '#ppt_x', off.x, dur, { additive: true }) + animProp(tgt, 'ppt_y', '#ppt_y', off.y, dur, { additive: true }) + setVisibility(tgt, 'hidden', dur - 1);
      break;
    }
    case 'float': {
      const up = dir !== 'down';
      presetId = entering ? (up ? 42 : 47) : (up ? 47 : 42);
      const offset = up ? '+.1' : '-.1';
      body = entering
        ? setVisibility(tgt, 'visible') + animEffect(tgt, 'in', 'fade', dur) + animProp(tgt, 'ppt_x', '#ppt_x', '#ppt_x', dur) + animProp(tgt, 'ppt_y', `#ppt_y${offset}`, '#ppt_y', dur)
        : animEffect(tgt, 'out', 'fade', dur) + animProp(tgt, 'ppt_x', '#ppt_x', '#ppt_x', dur) + animProp(tgt, 'ppt_y', '#ppt_y', `#ppt_y${up ? '-.1' : '+.1'}`, dur) + setVisibility(tgt, 'hidden', dur - 1);
      break;
    }
    case 'split': {
      const s = SPLIT[dir] || SPLIT['vertical-out'];
      subtype = s.subtype;
      body = entering
        ? setVisibility(tgt, 'visible') + animEffect(tgt, 'in', s.filter, dur)
        : animEffect(tgt, 'out', s.filter, dur) + setVisibility(tgt, 'hidden', dur - 1);
      break;
    }
    case 'wipe': {
      const side = WIPE_FILTER[dir] ? dir : 'bottom';
      subtype = SIDE_BITS[side];
      body = entering
        ? setVisibility(tgt, 'visible') + animEffect(tgt, 'in', WIPE_FILTER[side], dur)
        : animEffect(tgt, 'out', WIPE_FILTER[side], dur) + setVisibility(tgt, 'hidden', dur - 1);
      break;
    }
    case 'zoom':
      subtype = 16;
      body = entering
        ? setVisibility(tgt, 'visible') + animProp(tgt, 'ppt_w', '0', '#ppt_w', dur, { fromFloat: true }) + animProp(tgt, 'ppt_h', '0', '#ppt_h', dur, { fromFloat: true }) + animEffect(tgt, 'in', 'fade', dur)
        : animProp(tgt, 'ppt_w', '#ppt_w', '0', dur) + animProp(tgt, 'ppt_h', '#ppt_h', '0', dur) + animEffect(tgt, 'out', 'fade', dur) + setVisibility(tgt, 'hidden', dur - 1);
      break;
    case 'pulse':
      body = `<p:animEffect transition="out" filter="fade"><p:cBhvr><p:cTn id="${token()}" dur="${dur}" tmFilter="0, 0; .2, .5; .8, .5; 1, 0"/>${tgt}</p:cBhvr><p:progress><p:fltVal val="0.25"/></p:progress></p:animEffect>` +
        `<p:animScale><p:cBhvr><p:cTn id="${token()}" dur="${Math.max(1, Math.round(dur / 2))}" autoRev="1" fill="hold"/>${tgt}</p:cBhvr><p:by x="105000" y="105000"/></p:animScale>`;
      break;
    case 'spin':
      body = `<p:animRot by="${dir === 'counterclockwise' ? -21600000 : 21600000}"><p:cBhvr><p:cTn id="${token()}" dur="${dur}" fill="hold"/>${tgt}<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>`;
      break;
    case 'grow':
      body = `<p:animScale><p:cBhvr><p:cTn id="${token()}" dur="${dur}" fill="hold"/>${tgt}</p:cBhvr><p:by x="150000" y="150000"/></p:animScale>`;
      break;
    default:
      throw new Error(`"${effect}" is not an animation this writes`);
  }
  return (
    `<p:par><p:cTn id="${token()}" presetID="${presetId}" presetClass="${kind}" presetSubtype="${subtype}" fill="hold" grpId="${grpId}" nodeType="${nodeType}">` +
    `<p:stCondLst><p:cond delay="${Math.max(0, Math.round(delay))}"/></p:stCondLst><p:childTnLst>${body}</p:childTnLst></p:cTn></p:par>`
  );
}

const NODE_TYPE = { onClick: 'clickEffect', withPrevious: 'withEffect', afterPrevious: 'afterEffect' };

/**
 * An effect par already in the file, given its place in the rebuilt tree:
 * the effect cTn's node type and its own delay set, and — when asked — its
 * behaviours' durations scaled to a new length. Everything else is as read.
 */
function placed(e) {
  const tree = rangeTree(e.xml);
  const par = kid(tree, 'p:par');
  const ctn = kid(par, 'p:cTn');
  if (!ctn) return e.xml;
  const edits = [];
  // The effect cTn's own opening tag: node type.
  const open = e.xml.slice(ctn.start, ctn.openEnd);
  const nodeType = NODE_TYPE[e.trigger];
  let nextOpen = /\snodeType="[^"]*"/.test(open) ? open.replace(/\snodeType="[^"]*"/, () => ` nodeType="${nodeType}"`) : open.replace(/(\/?>)$/, (m, p1) => ` nodeType="${nodeType}"${p1 ?? ''}`);
  edits.push([ctn.start, ctn.openEnd, nextOpen]);
  // Its own start condition: the delay.
  const cond = kid(kid(ctn, 'p:stCondLst'), 'p:cond');
  if (cond) {
    const tag = e.xml.slice(cond.start, cond.openEnd);
    const delay = String(Math.max(0, Math.round(e.delay)));
    edits.push([cond.start, cond.openEnd, /\sdelay="[^"]*"/.test(tag) ? tag.replace(/\sdelay="[^"]*"/, () => ` delay="${delay}"`) : tag.replace(/(\/?>)$/, (m, p1) => ` delay="${delay}"${p1 ?? ''}`)]);
  }
  // A new length: every timed behaviour stretched by the same factor.
  if (e.scaleTo != null && e.duration > 1) {
    const factor = e.scaleTo / e.duration;
    for (const n of walk(ctn)) {
      if (n === ctn || n.name !== 'p:cTn' || n.attrs.dur == null) continue;
      const d = Number(n.attrs.dur);
      if (!Number.isFinite(d) || d <= 1) continue;
      const tag = e.xml.slice(n.start, n.openEnd);
      edits.push([n.start, n.openEnd, tag.replace(/\sdur="[^"]*"/, () => ` dur="${Math.max(2, Math.round(d * factor))}"`)]);
    }
  }
  edits.sort((a, b) => b[0] - a[0]);
  let out = e.xml;
  for (const [s, t, text] of edits) out = out.slice(0, s) + text + out.slice(t);
  return out;
}

/**
 * The click groups of a sequence, as XML, from the list in order. In a
 * trigger's sequence (`interactive`) the first group starts with the click
 * on the trigger itself, and none starts with the slide.
 */
function groupsXml(effects, mainSeqId, { interactive = false } = {}) {
  const groups = [];
  for (const e of effects) {
    if (!groups.length || e.trigger === 'onClick') groups.push({ auto: !interactive && !groups.length && e.trigger !== 'onClick', first: !groups.length, subs: [] });
    const g = groups[groups.length - 1];
    if (!g.subs.length || e.trigger === 'afterPrevious') g.subs.push({ effects: [] });
    g.subs[g.subs.length - 1].effects.push(e);
  }
  return groups.map((g) => {
    let start = 0;
    const subs = g.subs.map((s) => {
      const at = start;
      const length = Math.max(0, ...s.effects.map((e) => (e.delay || 0) + (e.scaleTo ?? e.duration ?? 0)));
      start = at + length;
      return `<p:par><p:cTn id="${token()}" fill="hold"><p:stCondLst><p:cond delay="${Math.round(at)}"/></p:stCondLst><p:childTnLst>${s.effects.map(placed).join('')}</p:childTnLst></p:cTn></p:par>`;
    });
    const cond = interactive && g.first ? '<p:cond delay="0"/>'
      : g.auto ? `<p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="${mainSeqId}"/></p:cond>` : '<p:cond delay="indefinite"/>';
    return `<p:par><p:cTn id="${token()}" fill="hold"><p:stCondLst>${cond}</p:stCondLst><p:childTnLst>${subs.join('')}</p:childTnLst></p:cTn></p:par>`;
  }).join('');
}

/** A trigger's interactive sequence — a click on shape `spid` plays its groups in turn — as PowerPoint writes one. */
function interactiveSeqXml(spid, effects) {
  const id = token();
  const click = `<p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cond>`;
  return `<p:seq concurrent="1" nextAc="seek"><p:cTn id="${id}" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq">` +
    `<p:stCondLst>${click}</p:stCondLst><p:endSync evt="end" delay="0"><p:rtn val="all"/></p:endSync>` +
    `<p:childTnLst>${groupsXml(effects, id, { interactive: true })}</p:childTnLst></p:cTn><p:nextCondLst>${click}</p:nextCondLst></p:seq>`;
}

const MAIN_SEQ_CONDS = '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>';

/**
 * Every time node numbered 1, 2, 3… in document order, as PowerPoint
 * numbers them, and every `p:tn val` reference following its node.
 */
function renumber(xml) {
  const map = new Map();
  let n = 0;
  const numbered = xml.replace(/(<p:cTn\b[^>]*?\sid=")([^"]*)(")/g, (m, a, id, b) => {
    n += 1;
    if (!map.has(id)) map.set(id, String(n));
    return `${a}${n}${b}`;
  });
  return numbered.replace(/(<p:tn\b[^>]*?\sval=")([^"]*)(")/g, (m, a, id, b) => (map.has(id) ? `${a}${map.get(id)}${b}` : m));
}

/** The shapes of a slide this can tell apart for the build list. */
function shapeInfo(slideXml) {
  const info = new Map();
  const re = /<p:(sp|pic|graphicFrame|grpSp|cxnSp)>[\s\S]*?<p:cNvPr\b[^>]*\bid="(\d+)"/g;
  let m;
  while ((m = re.exec(slideXml))) {
    const tag = m[1];
    const id = m[2];
    if (info.has(id)) continue;
    // Only a p:sp builds; whether it is a placeholder with words decides animBg.
    const body = tag === 'sp' ? slideXml.slice(m.index, slideXml.indexOf('</p:sp>', m.index)) : '';
    info.set(id, { tag, placeholder: /<p:ph\b/.test(body), text: /<a:t>|<a:t\s/.test(body) });
  }
  return info;
}

/**
 * The build list: every bldP a remaining effect still refers to (by spid
 * and grpId), plus one for each new effect on a p:sp, the way PowerPoint
 * writes them — `animBg="1"` for a shape that is not a text placeholder.
 * Anything else in the list (a chart's or a diagram's build) stays.
 */
function buildList(oldList, allEffectRefs, fresh, slideXml) {
  // build="p" on a shape whose effects go paragraph by paragraph, and off one whose do not.
  const byPara = (tag, key) => {
    const want = allEffectRefs.get(key) === true;
    const has = /\sbuild="p"/.test(tag);
    if (want === has) return tag;
    return want ? tag.replace(/^<p:bldP\b/, '<p:bldP build="p"') : tag.replace(/\sbuild="p"/, '');
  };
  const info = shapeInfo(slideXml);
  const keep = [];
  const have = new Set();
  const others = [];
  if (oldList) {
    for (const m of oldList.matchAll(/<p:bldP\b[^>]*?(?:\/>|>[\s\S]*?<\/p:bldP>)|<p:(bldGraphic|bldDgm|bldOleChart)\b[^>]*?(?:\/>|>[\s\S]*?<\/p:\1>)/g)) {
      if (m[0].startsWith('<p:bldP')) {
        const spid = /\sspid="([^"]*)"/.exec(m[0])?.[1];
        const grp = /\sgrpId="([^"]*)"/.exec(m[0])?.[1] ?? '0';
        const key = `${spid}:${grp}`;
        if (allEffectRefs.has(key) && !have.has(key)) { keep.push(byPara(m[0], key)); have.add(key); }
      } else {
        others.push(m[0]);
      }
    }
  }
  for (const f of fresh) {
    const key = `${f.spid}:${f.grpId}`;
    if (have.has(key)) continue;
    const s = info.get(String(f.spid));
    if (!s || s.tag !== 'sp') continue;
    keep.push(byPara(`<p:bldP spid="${f.spid}" grpId="${f.grpId}"${s.placeholder && s.text ? '' : ' animBg="1"'}/>`, key));
    have.add(key);
  }
  const all = [...keep, ...others];
  return all.length ? `<p:bldLst>${all.join('')}</p:bldLst>` : '';
}

/** Every (spid, grpId) pair an effect anywhere in the tree refers to, each with whether it aims at a paragraph. */
function effectRefs(timingXml) {
  const refs = new Map();
  const tree = rangeTree(timingXml);
  for (const n of walk(tree)) {
    if (n.name !== 'p:cTn' || n.attrs.grpId == null) continue;
    const tgt = [...walk(n)].find((c) => c.name === 'p:spTgt');
    if (!tgt) continue;
    const key = `${tgt.attrs.spid}:${n.attrs.grpId}`;
    const para = tgt.children.some((c) => c.name === 'p:txEl');
    refs.set(key, Boolean(refs.get(key)) || para);
  }
  return refs;
}

/**
 * The slide's XML with its main sequence replaced by `list` — entries as
 * `readTree` gives them (kept XML) or fresh ones (`fresh: true`, with the
 * spec to build from). Returns the slide XML; a list left empty takes the
 * main sequence out, and a timing with nothing else in it goes too.
 */
function writeList(slideXml, list) {
  const range = timingRange(slideXml);
  const original = range ? slideXml.slice(range.start, range.end) : null;
  const read = original ? readTree(original) : null;
  if (read && !read.editable) throw new Error('This slide\'s animations are laid out in a way this cannot rewrite; they are kept as they are.');

  // The effects in their new order, each as XML ready to place.
  const fresh = [];
  const all = list.map((e) => {
    if (!e.fresh) return e;
    const xml = effectXml({ ...e.spec, grpId: e.grpId, nodeType: NODE_TYPE[e.trigger], delay: e.delay, spid: e.shapeId, paragraph: e.paragraph ?? null });
    fresh.push({ spid: e.shapeId, grpId: e.grpId });
    return { ...e, xml, fresh: false, scaleTo: null };
  });
  // The main sequence's effects; a trigger's go in a sequence of their own.
  const effects = all.filter((e) => !e.triggerShape);
  const byTrigger = new Map();
  for (const e of all) {
    if (!e.triggerShape) continue;
    if (!byTrigger.has(e.triggerShape)) byTrigger.set(e.triggerShape, []);
    byTrigger.get(e.triggerShape).push(e);
  }

  // The trigger sequences read from the file come out here, and those the
  // list has go back in after the main sequence below.
  let before = original;
  if (read?.interactive.length) {
    for (const s of [...read.interactive].sort((a, b) => b.seq.start - a.seq.start)) before = before.slice(0, s.seq.start) + before.slice(s.seq.end);
  }
  const parts = before ? readTree(before) : null;

  const mainId = parts?.mainCtn?.attrs.id ?? token();
  let timing;
  if (!parts) {
    if (!effects.length) timing = '';
    else {
    const root = token();
    timing = `<p:timing><p:tnLst><p:par><p:cTn id="${root}" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>` +
      `<p:seq concurrent="1" nextAc="seek"><p:cTn id="${mainId}" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${groupsXml(effects, mainId)}</p:childTnLst></p:cTn>${MAIN_SEQ_CONDS}</p:seq>` +
      `</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
    }
  } else {
    const t = before;
    if (parts.mainSeq) {
      if (effects.length) {
        // Only the main sequence's click groups change; its own cTn and conditions stay.
        const inner = groupsXml(effects, mainId);
        timing = parts.mainKids
          ? t.slice(0, parts.mainKids.innerStart) + inner + t.slice(parts.mainKids.innerEnd)
          : t.slice(0, parts.mainCtn.innerEnd) + `<p:childTnLst>${inner}</p:childTnLst>` + t.slice(parts.mainCtn.innerEnd);
      } else {
        timing = t.slice(0, parts.mainSeq.start) + t.slice(parts.mainSeq.end);
      }
    } else if (effects.length) {
      const seq = `<p:seq concurrent="1" nextAc="seek"><p:cTn id="${mainId}" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${groupsXml(effects, mainId)}</p:childTnLst></p:cTn>${MAIN_SEQ_CONDS}</p:seq>`;
      timing = parts.rootKids
        ? t.slice(0, parts.rootKids.innerStart) + seq + t.slice(parts.rootKids.innerStart)
        : t.slice(0, parts.rootCtn.innerEnd) + `<p:childTnLst>${seq}</p:childTnLst>` + t.slice(parts.rootCtn.innerEnd);
    } else {
      timing = t;
    }
    // A root with nothing left under it: the whole timing goes.
    if (/<p:childTnLst>\s*<\/p:childTnLst>/.test(timing)) {
      const check = readTree(timing);
      if (!check.rootKids || !check.rootKids.children.length) timing = '';
    }
  }

  // The triggers' sequences, after the main sequence (or first under the
  // root when there is none), as PowerPoint orders them.
  if (byTrigger.size) {
    const seqs = [...byTrigger].map(([spid, es]) => interactiveSeqXml(spid, es.map((e, i) => (i === 0 ? { ...e, trigger: 'onClick' } : e)))).join('');
    if (!timing) {
      timing = `<p:timing><p:tnLst><p:par><p:cTn id="${token()}" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>${seqs}</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
    } else {
      const now = readTree(timing);
      const at = now.mainSeq ? now.mainSeq.end : now.rootKids ? now.rootKids.innerStart : null;
      timing = at != null
        ? timing.slice(0, at) + seqs + timing.slice(at)
        : timing.slice(0, now.rootCtn.innerEnd) + `<p:childTnLst>${seqs}</p:childTnLst>` + timing.slice(now.rootCtn.innerEnd);
    }
  }

  if (timing) {
    // The build list, rebuilt against the effects that are left — from the
    // list the slide had, which a timing rebuilt from nothing has lost.
    const oldList = /<p:bldLst\b[^>]*?(?:\/>|>[\s\S]*?<\/p:bldLst>)/.exec(original || timing)?.[0] || null;
    const ownList = /<p:bldLst\b[^>]*?(?:\/>|>[\s\S]*?<\/p:bldLst>)/.exec(timing)?.[0] || null;
    const withoutList = ownList ? timing.replace(ownList, '') : timing;
    const refs = effectRefs(withoutList);
    const bld = buildList(oldList, refs, fresh, slideXml);
    timing = bld ? withoutList.replace(/<\/p:tnLst>/, () => `</p:tnLst>${bld}`) : withoutList;
    timing = renumber(timing);
  }

  const bare = range ? slideXml.slice(0, range.start) + slideXml.slice(range.end) : slideXml;
  return timing ? insertTiming(bare, timing) : bare;
}

/** A timing block put where the schema wants it: after the transition (or clrMapOvr, or cSld), before extLst. */
export function insertTiming(xml, block) {
  const tail = xml.indexOf('</p:cSld>');
  if (tail < 0) throw new Error('the slide has no cSld to put its animations after');
  // The root's own extension list is the last thing before its closing tag.
  const ext = /<p:extLst\b(?:(?!<p:extLst\b)[\s\S])*<\/p:extLst>\s*<\/p:sld>\s*$/.exec(xml);
  const end = /<\/p:sld>\s*$/.exec(xml);
  const at = ext && ext.index > tail ? ext.index : end ? end.index : xml.length;
  return xml.slice(0, at) + block + xml.slice(at);
}

/* ── the editing verbs ───────────────────────────────────────────────────── */

function currentList(slideXml) {
  const range = timingRange(slideXml);
  if (!range) return [];
  const parts = readTree(slideXml.slice(range.start, range.end));
  if (!parts.editable) throw new Error('This slide\'s animations are laid out in a way this cannot rewrite; they are kept as they are.');
  return parts.effects;
}

/** The next free build group for a shape: one more than any effect on it already uses. */
function nextGrpId(slideXml, spid) {
  const range = timingRange(slideXml);
  if (!range) return 0;
  let max = -1;
  for (const key of effectRefs(slideXml.slice(range.start, range.end)).keys()) {
    const [s, g] = key.split(':');
    if (s === String(spid)) max = Math.max(max, Number(g));
  }
  return max + 1;
}

/**
 * Add an effect to a shape — Add Animation, or the gallery on a shape with
 * none. `spec`: `{ kind, effect, direction?, duration? (s), delay? (s),
 * trigger? }`, at the end of the sequence unless `at` says where.
 * @returns {{ xml: string, index: number }}
 */
export function addAnimation(slideXml, shapeId, spec = {}, at = null) {
  const kind = spec.kind || 'entr';
  const effect = spec.effect || (kind === 'path' ? 'custom' : 'fade');
  if (!EFFECTS[kind]?.[effect]) throw new Error(`"${kind} ${effect}" is not an animation this writes`);
  if (kind === 'path' && !/^[Mm]\s/.test(String(spec.path || '').trim())) throw new Error('A motion path needs a path to follow');
  if (!new RegExp(`<p:cNvPr\\b[^>]*\\bid="${shapeId}"`).test(slideXml)) throw new Error(`shape ${shapeId} is not on this slide`);
  const list = currentList(slideXml);
  const entry = {
    fresh: true,
    spec: { kind, effect, direction: spec.direction ?? EFFECTS[kind][effect].direction ?? null, duration: spec.duration ?? EFFECTS[kind][effect].duration, path: kind === 'path' ? spec.path : null },
    shapeId: String(shapeId),
    grpId: nextGrpId(slideXml, shapeId),
    trigger: spec.trigger || 'onClick',
    delay: Math.round((spec.delay || 0) * 1000),
    duration: Math.round((spec.duration ?? EFFECTS[kind][effect].duration) * 1000),
    paragraph: spec.paragraph ?? null,
  };
  const index = at == null ? list.length : Math.max(0, Math.min(list.length, at));
  list.splice(index, 0, entry);
  return { xml: writeList(slideXml, list), index };
}

/**
 * Change one effect: its trigger, duration or delay — kept XML is edited in
 * place — or its effect and direction, which rebuilds it (same shape, same
 * paragraph, same build group).
 */
export function setAnimation(slideXml, index, patch = {}) {
  const list = currentList(slideXml);
  const e = list[index];
  if (!e) throw new RangeError(`no animation at ${index}`);
  const next = { ...e };
  if (patch.trigger) next.trigger = patch.trigger;
  // Animations → Trigger: on a click on a shape of the slide, or (null) back
  // in the slide's own sequence. Moved onto a trigger it starts on that click.
  if (patch.triggerShape !== undefined) {
    const spid = patch.triggerShape == null ? null : String(patch.triggerShape);
    if (spid && !new RegExp(`<p:cNvPr\\b[^>]*\\bid="${spid}"`).test(slideXml)) throw new Error(`shape ${spid} is not on this slide to start the effect`);
    next.triggerShape = spid;
    if (spid && !patch.trigger) next.trigger = 'onClick';
  }
  if (patch.delay != null) next.delay = Math.max(0, Math.round(Number(patch.delay) * 1000));
  const rebuild = patch.effect !== undefined || patch.kind !== undefined || patch.direction !== undefined || patch.path !== undefined;
  if (rebuild) {
    const kind = patch.kind ?? e.kind;
    const effect = patch.effect ?? (patch.kind === 'path' ? 'custom' : e.effect);
    const spec = EFFECTS[kind]?.[effect];
    if (!spec) throw new Error(`"${kind} ${effect}" is not an animation this writes`);
    const sameEffect = kind === e.kind && effect === e.effect;
    const duration = patch.duration != null ? Number(patch.duration) : sameEffect && e.duration ? e.duration / 1000 : spec.duration;
    const direction = patch.direction !== undefined ? patch.direction : sameEffect ? e.direction : spec.direction ?? null;
    const path = kind === 'path' ? patch.path ?? e.path ?? null : null;
    if (kind === 'path' && !path) throw new Error('A motion path needs a path to follow');
    list[index] = { ...next, fresh: true, spec: { kind, effect, direction, duration, path }, grpId: e.grpId ?? nextGrpId(slideXml, e.shapeId), duration: Math.round(duration * 1000) };
  } else {
    if (patch.duration != null) {
      const ms = Math.max(1, Math.round(Number(patch.duration) * 1000));
      if (writes(e)) {
        // One this writes: rebuilt at the new length, exactly as it would be new.
        list[index] = { ...next, fresh: true, spec: { kind: e.kind, effect: e.effect, direction: e.direction, duration: ms / 1000 }, grpId: e.grpId ?? 0, duration: ms };
      } else {
        list[index] = { ...next, scaleTo: ms };
      }
    } else {
      list[index] = next;
    }
  }
  // A fresh entry carries what writeList needs.
  if (list[index].fresh) list[index].paragraph = e.paragraph;
  return writeList(slideXml, list);
}

/** The paragraphs of a shape's text that have words, by their place among the shape's paragraphs. */
export function shapeParagraphs(slideXml, shapeId) {
  const re = new RegExp(`<p:sp>(?:(?!</p:sp>)[\\s\\S])*?<p:cNvPr\\b[^>]*\\bid="${String(shapeId).replace(/\D/g, '')}"[\\s\\S]*?</p:sp>`);
  const sp = re.exec(slideXml)?.[0];
  const body = sp ? /<p:txBody>([\s\S]*?)<\/p:txBody>/.exec(sp)?.[1] : null;
  if (!body) return [];
  const out = [];
  [...body.matchAll(/<a:p\b(?:\/>|>([\s\S]*?)<\/a:p>)/g)].forEach((m, i) => {
    const words = [...String(m[1] || '').matchAll(/<a:t(?:\s[^>]*)?>([^<]*)<\/a:t>/g)].map((t) => t[1]).join('');
    if (words.trim()) out.push(i);
  });
  return out;
}

/**
 * Effect Options → Sequence: the effect at `index` played as one object
 * (`object`), or split into one per paragraph of its shape's words — each
 * on a click of its own (`paragraph`), or all together (`together`). A
 * split effect is gathered back into one from any of its paragraphs.
 * @returns {{ xml: string, index: number }}
 */
export function setSequence(slideXml, index, how = 'paragraph') {
  const list = currentList(slideXml);
  const e = list[index];
  if (!e) throw new RangeError(`no animation at ${index}`);
  // The run of this shape's effects that one split made: same shape, same build group, a paragraph each.
  let first = index;
  let last = index;
  if (e.paragraph != null) {
    const same = (x) => x && x.paragraph != null && String(x.shapeId) === String(e.shapeId) && x.grpId === e.grpId && !x.triggerShape === !e.triggerShape;
    while (same(list[first - 1])) first -= 1;
    while (same(list[last + 1])) last += 1;
  }
  const head = list[first];
  const kind = head.kind;
  const effect = head.effect;
  if (!EFFECTS[kind]?.[effect] || kind === 'path' || (!writes(head) && !head.fresh)) throw new Error('Only an entrance, emphasis or exit this writes can be played by paragraph');
  const spec = { kind, effect, direction: head.direction ?? null, duration: (head.duration || 0) / 1000 || EFFECTS[kind][effect].duration };
  const base = { fresh: true, spec, shapeId: String(head.shapeId), grpId: head.grpId ?? nextGrpId(slideXml, head.shapeId), delay: head.delay || 0, duration: Math.round(spec.duration * 1000), triggerShape: head.triggerShape ?? null };
  let entries;
  if (how === 'object') {
    entries = [{ ...base, trigger: head.trigger, paragraph: null }];
  } else {
    const paras = shapeParagraphs(slideXml, head.shapeId);
    if (!paras.length) throw new Error('The shape has no words to play paragraph by paragraph');
    entries = paras.map((para, i) => ({ ...base, delay: i ? 0 : base.delay, trigger: i === 0 ? head.trigger : how === 'together' ? 'withPrevious' : 'onClick', paragraph: para }));
  }
  list.splice(first, last - first + 1, ...entries);
  return { xml: writeList(slideXml, list), index: first };
}

/** Take one effect out; an emptied click group goes with it. */
export function removeAnimation(slideXml, index) {
  const list = currentList(slideXml);
  if (!list[index]) throw new RangeError(`no animation at ${index}`);
  list.splice(index, 1);
  return writeList(slideXml, list);
}

/** Every effect on a shape taken out — the gallery's None, or the shape deleted. */
export function removeShapeAnimations(slideXml, shapeId) {
  if (!timingRange(slideXml)) return slideXml;
  const list = currentList(slideXml);
  // Its effects go, and so do the effects a click on it started.
  const kept = list.filter((e) => String(e.shapeId) !== String(shapeId) && String(e.triggerShape) !== String(shapeId));
  return kept.length === list.length ? slideXml : writeList(slideXml, kept);
}

/** Effects on shapes no longer on the slide taken out, so the file never points at nothing. */
export function pruneAnimations(slideXml) {
  const range = timingRange(slideXml);
  if (!range) return slideXml;
  const parts = readTree(slideXml.slice(range.start, range.end));
  if (!parts.editable) return slideXml;
  const ids = new Set([...slideXml.matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"/g)].map((m) => m[1]));
  const kept = parts.effects.filter((e) => (e.shapeId == null || ids.has(String(e.shapeId))) && (!e.triggerShape || ids.has(String(e.triggerShape))));
  return kept.length === parts.effects.length ? slideXml : writeList(slideXml, kept);
}

/**
 * Move Earlier / Move Later (or a drag in the pane): the effect at `index`
 * to `to` — 'earlier', 'later' or a position. It keeps its own trigger.
 * @returns {{ xml: string, index: number }}
 */
export function moveAnimation(slideXml, index, to) {
  const list = currentList(slideXml);
  if (!list[index]) throw new RangeError(`no animation at ${index}`);
  const target = to === 'earlier' ? index - 1 : to === 'later' ? index + 1 : Number(to);
  const dest = Math.max(0, Math.min(list.length - 1, target));
  if (dest === index) return { xml: slideXml, index };
  const [e] = list.splice(index, 1);
  list.splice(dest, 0, e);
  return { xml: writeList(slideXml, list), index: dest };
}
