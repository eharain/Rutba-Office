/**
 * Office Math (OMML) — read into MathML and into Word's linear form.
 *
 * An equation in a .docx is an `m:oMathPara` (a display equation, on its own
 * line) or an `m:oMath` (inline, in the run of the text) sitting among a
 * paragraph's runs. The document engine keeps that XML untouched — a saved
 * file writes back exactly the bytes it read — and asks this module for two
 * readings of it:
 *
 *   - MathML, which the page draws with Chromium's own MathML Core layout:
 *     a fraction with a real bar, a radical with its overbar, stacked limits,
 *     italic variables and upright function names, as Word draws them;
 *   - the LINEAR form, Word's UnicodeMath (`x=(-b±√(b^2-4ac))/2a`), which
 *     the equation editor shows when an equation is opened again and which a
 *     printout without a MathML renderer falls back to.
 *
 * Pure: no DOM, no Node API — the Word window imports it for the editor's
 * live preview, the engine for the model. Anything this does not know is
 * read as its text, never dropped: an equation that draws wrong is visible,
 * one that vanished is not.
 */

export const MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
export const MATHML_NS = 'http://www.w3.org/1998/Math/MathML';

/* ── a small XML reader for one equation's subtree ─────────────────────── */

const ENTITIES = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
function unescapeXml(s) {
  return String(s).replace(/&(#x[0-9a-fA-F]+|#\d+|[a-zA-Z]+);/g, (m, e) => {
    if (e[0] === '#') {
      const code = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return Number.isFinite(code) ? String.fromCodePoint(code) : m;
    }
    return ENTITIES[e] ?? m;
  });
}
const escapeText = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const escapeAttr = (s) => escapeText(s).replace(/"/g, '&quot;');

/**
 * Parse an XML fragment into `{ name, attrs, children }` elements and
 * `{ text }` nodes. Namespace prefixes are kept as written (`m:f`, `w:rPr`):
 * Word always writes Office Math as `m:` and its run properties as `w:`.
 */
export function parseXml(xml) {
  const root = { name: '#root', attrs: {}, children: [] };
  const stack = [root];
  const re = /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[([\s\S]*?)\]\]>|<\/([^\s>]+)\s*>|<([^\s/>!?]+)((?:\s+[^\s=/>]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>|([^<]+)/g;
  let m;
  while ((m = re.exec(String(xml)))) {
    const top = stack[stack.length - 1];
    if (m[1] !== undefined) top.children.push({ text: m[1] });
    else if (m[2] !== undefined) {
      for (let i = stack.length - 1; i > 0; i--) {
        if (stack[i].name === m[2]) { stack.length = i; break; }
      }
    } else if (m[3] !== undefined) {
      const attrs = {};
      for (const a of m[4].matchAll(/([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) attrs[a[1]] = unescapeXml(a[2] ?? a[3] ?? '');
      const el = { name: m[3], attrs, children: [] };
      top.children.push(el);
      if (m[5] !== '/') stack.push(el);
    } else if (m[6] !== undefined) {
      top.children.push({ text: unescapeXml(m[6]) });
    }
  }
  return root;
}

const kids = (node, name = null) => (node?.children || []).filter((c) => c.name && (name === null || c.name === name));
const kid = (node, name) => (node?.children || []).find((c) => c.name === name) || null;
/** A property element's `m:val`, or undefined when the element is absent. */
function prop(pr, name) {
  const el = kid(pr, name);
  if (!el) return undefined;
  return el.attrs['m:val'] ?? el.attrs.val ?? '';
}
/** An on/off property: present with no value, or 1/on/true, is on. */
function flag(pr, name) {
  const v = prop(pr, name);
  if (v === undefined) return false;
  return v === '' || !/^(0|off|false)$/i.test(v);
}
/** Every `m:t` (and a stray `w:t`) under a node, in order. */
function textUnder(node) {
  let out = '';
  for (const c of node?.children || []) {
    if (c.text !== undefined) continue;
    if (c.name === 'm:t' || c.name === 'w:t') out += (c.children || []).map((t) => t.text ?? '').join('');
    else if (c.name === 'w:del' || /Pr$/.test(c.name)) continue;
    else out += textUnder(c);
  }
  return out;
}

/** The plain text of an equation — what a search, a count or an accessibility reader sees. */
export function ommlText(xml) {
  return textUnder(parseXml(xml));
}

/** Top-level facts about an equation fragment: display or inline, its justification. */
export function ommlInfo(xml) {
  const root = parseXml(xml);
  const para = kid(root, 'm:oMathPara');
  if (para) {
    const pr = kid(para, 'm:oMathParaPr');
    return { display: true, jc: prop(pr, 'm:jc') || 'centerGroup', count: kids(para, 'm:oMath').length };
  }
  return { display: false, jc: null, count: kids(root, 'm:oMath').length };
}

/* ── characters ────────────────────────────────────────────────────────── */

const INTEGRALS = new Set(['∫', '∬', '∭', '∮', '∯', '∰', '∱', '∲', '∳', '⨌']);
/** Identifiers that are not letters: drawn upright, as Word draws them. */
const SYMBOL_IDENTIFIERS = new Set(['∞', '∂', '∇', 'ℏ', '∅', 'ℓ', '℘', 'ℜ', 'ℑ', 'ℵ', '…', '⋯', '⋮', '⋱', '°']);
/** A combining accent Word stores in m:acc, and the spacing character MathML draws above the base. */
const ACCENT_GLYPH = {
  '\u0300': '`', '\u0301': '´', '\u0302': 'ˆ', '\u0303': '˜', '\u0304': '¯', '\u0305': '‾',
  '\u0306': '˘', '\u0307': '˙', '\u0308': '¨', '\u030C': 'ˇ', '\u030A': '˚',
  '\u20D6': '←', '\u20D7': '→', '\u20E1': '↔', '\u20D0': '↼', '\u20D1': '⇀', '\u20DB': '\u20DB',
};
const isLetter = (c) => /\p{L}/u.test(c);
const isDigit = (c) => c >= '0' && c <= '9';
const isUpperGreek = (c) => c >= 'Α' && c <= 'Ω';

/** Word's m:scr (script, fraktur, double-struck…) as Unicode's mathematical alphabets. */
const SCRIPT_HOLES = {
  script: { B: 'ℬ', E: 'ℰ', F: 'ℱ', H: 'ℋ', I: 'ℐ', L: 'ℒ', M: 'ℳ', R: 'ℛ', e: 'ℯ', g: 'ℊ', o: 'ℴ' },
  fraktur: { C: 'ℭ', H: 'ℌ', I: 'ℑ', R: 'ℜ', Z: 'ℨ' },
  'double-struck': { C: 'ℂ', H: 'ℍ', N: 'ℕ', P: 'ℙ', Q: 'ℚ', R: 'ℝ', Z: 'ℤ' },
};
const SCRIPT_BASE = { script: 0x1D49C, fraktur: 0x1D504, 'double-struck': 0x1D538, 'sans-serif': 0x1D5A0, monospace: 0x1D670 };
function scriptLetter(c, scr) {
  const base = SCRIPT_BASE[scr];
  if (!base || !/[A-Za-z]/.test(c)) return c;
  const hole = SCRIPT_HOLES[scr]?.[c];
  if (hole) return hole;
  const i = c >= 'a' ? 26 + c.charCodeAt(0) - 97 : c.charCodeAt(0) - 65;
  return String.fromCodePoint(base + i);
}
const SCR_NAMES = { script: 'script', fraktur: 'fraktur', 'double-struck': 'double-struck', 'sans-serif': 'sans-serif', monospace: 'monospace', roman: null };

const HIGHLIGHT = {
  yellow: '#FFFF00', green: '#00FF00', cyan: '#00FFFF', magenta: '#FF00FF', blue: '#0000FF', red: '#FF0000',
  darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000',
  darkYellow: '#808000', darkGray: '#808080', lightGray: '#C0C0C0', black: '#000000', white: '#FFFFFF',
};

/* ── OMML → MathML ─────────────────────────────────────────────────────── */

/** An empty slot Word shows as a dotted box while editing; the page shows it too, faintly. */
const PLACEHOLDER = '<mi mathvariant="normal" style="color:#9aa0a6">⬚</mi>';
/** The mark `&` leaves in an equation array's text: split into columns there, dropped elsewhere. */
const AMP = '\u0000&\u0000';

/**
 * One `m:r`: its words tokenised the way MathML wants them — identifiers
 * (one letter each, so each is italic as Word draws a variable), numbers,
 * operators — styled by the run's own `m:sty`/`m:scr`/`m:nor` and its
 * colour and highlight.
 */
function runToMathml(r, ctx) {
  const mpr = kid(r, 'm:rPr');
  const wpr = kid(r, 'w:rPr');
  const text = textUnder(r);
  if (!text) return '';
  const sty = prop(mpr, 'm:sty') ?? (ctx.upright ? 'p' : undefined);
  const scr = SCR_NAMES[prop(mpr, 'm:scr')] ?? null;
  const bold = sty === 'b' || sty === 'bi' || flag(wpr, 'w:b');
  const colour = kid(wpr, 'w:color')?.attrs['w:val'];
  const highlight = kid(wpr, 'w:highlight')?.attrs['w:val'];
  const css = [
    bold ? 'font-weight:bold' : '',
    colour && /^[0-9A-Fa-f]{6}$/.test(colour) ? `color:#${colour}` : '',
    highlight && HIGHLIGHT[highlight] ? `background:${HIGHLIGHT[highlight]}` : '',
  ].filter(Boolean).join(';');
  const style = css ? ` style="${css}"` : '';
  if (flag(mpr, 'm:nor')) return `<mtext${style}>${escapeText(text)}</mtext>`;

  const upright = sty === 'p' || sty === 'b';
  const out = [];
  const chars = [...text];
  for (let i = 0; i < chars.length; i++) {
    const c = chars[i];
    if (c === '&') { out.push(ctx.amp ? AMP : ''); continue; }
    if (c === ' ') { out.push('<mspace width="0.25em"/>'); continue; }
    if (isDigit(c)) {
      let n = c;
      while (i + 1 < chars.length && (isDigit(chars[i + 1]) || ((chars[i + 1] === '.' || chars[i + 1] === ',') && isDigit(chars[i + 2] ?? '')))) n += chars[++i];
      out.push(`<mn${style}>${escapeText(n)}</mn>`);
      continue;
    }
    if (isLetter(c)) {
      if (upright && !scr) {
        // An upright run keeps its letters together — `sin`, a unit, a
        // label — and a multi-letter identifier is upright by nature.
        let w = c;
        while (i + 1 < chars.length && isLetter(chars[i + 1])) w += chars[++i];
        out.push(`<mi${w.length === 1 ? ' mathvariant="normal"' : ''}${style}>${escapeText(w)}</mi>`);
        continue;
      }
      const glyph = scr ? scriptLetter(c, scr) : c;
      const normal = scr || isUpperGreek(c) ? ' mathvariant="normal"' : '';
      out.push(`<mi${normal}${style}>${escapeText(glyph)}</mi>`);
      continue;
    }
    if (SYMBOL_IDENTIFIERS.has(c)) { out.push(`<mi mathvariant="normal"${style}>${escapeText(c)}</mi>`); continue; }
    const op = c === '-' ? '−' : c === "'" ? '′' : c === '*' ? '∗' : c;
    // A prime and a comma hug what they follow; everything else takes the
    // dictionary's spacing, which is what gives `a+b` its air.
    out.push(`<mo${ctx.script ? ' lspace="0.05em" rspace="0.05em"' : ''}${style}>${escapeText(op)}</mo>`);
  }
  return out.join('');
}

/** A container (m:e, m:num, m:sub…) as one `<mrow>` — or a placeholder when it is empty. */
function rowOf(node, ctx, { placeholder = false } = {}) {
  const inner = node ? childrenToMathml(node, ctx) : '';
  if (!inner) return placeholder ? PLACEHOLDER : '<mrow></mrow>';
  return `<mrow>${inner}</mrow>`;
}

function childrenToMathml(node, ctx) {
  let out = '';
  for (const c of node.children || []) {
    if (c.text !== undefined) continue;
    out += elementToMathml(c, ctx);
  }
  return out;
}

/** A script's context: its operators hug their operands, as they do under a sum in Word. */
const sc = (ctx) => (ctx.script ? ctx : { ...ctx, script: true });

const mo = (c, attrs = '') => `<mo${attrs}>${escapeText(c)}</mo>`;
const isEmpty = (node) => !node || !textUnder(node) && !kids(node).some((c) => !/Pr$/.test(c.name));

/** A base with a rule drawn over or under it. */
const barOver = (row, top) => (row.startsWith("<mrow>") ? row : "<mrow>" + row + "</mrow>").replace(/^<mrow>/, `<mrow style="border-${top ? 'top' : 'bottom'}:0.06em solid currentColor;padding-${top ? 'top' : 'bottom'}:0.08em">`);

function elementToMathml(el, ctx) {
  const pr = kid(el, el.name + 'Pr');
  switch (el.name) {
    case 'm:r': return runToMathml(el, ctx);
    case 'm:oMath': return childrenToMathml(el, ctx);
    case 'm:e': case 'm:num': case 'm:den': case 'm:sub': case 'm:sup': case 'm:deg': case 'm:lim': case 'm:fName':
      return rowOf(el, ctx);
    case 'w:ins': case 'w:smartTag': case 'w:customXml': case 'w:hyperlink':
      return childrenToMathml(el, ctx);
    case 'm:f': {
      const type = prop(pr, 'm:type') || 'bar';
      const num = rowOf(kid(el, 'm:num'), ctx, { placeholder: true });
      const den = rowOf(kid(el, 'm:den'), ctx, { placeholder: true });
      if (type === 'noBar') return `<mfrac linethickness="0">${num}${den}</mfrac>`;
      if (type === 'lin' || type === 'skw') return `<mrow>${num}${mo(type === 'skw' ? '∕' : '/')}${den}</mrow>`;
      return `<mfrac>${num}${den}</mfrac>`;
    }
    case 'm:rad': {
      const e = rowOf(kid(el, 'm:e'), ctx, { placeholder: true });
      const deg = kid(el, 'm:deg');
      if (flag(pr, 'm:degHide') || isEmpty(deg)) return `<msqrt>${e}</msqrt>`;
      return `<mroot>${e}${rowOf(deg, sc(ctx))}</mroot>`;
    }
    case 'm:sSup':
      return `<msup>${rowOf(kid(el, 'm:e'), ctx, { placeholder: true })}${rowOf(kid(el, 'm:sup'), sc(ctx), { placeholder: true })}</msup>`;
    case 'm:sSub':
      return `<msub>${rowOf(kid(el, 'm:e'), ctx, { placeholder: true })}${rowOf(kid(el, 'm:sub'), sc(ctx), { placeholder: true })}</msub>`;
    case 'm:sSubSup':
      return `<msubsup>${rowOf(kid(el, 'm:e'), ctx, { placeholder: true })}${rowOf(kid(el, 'm:sub'), sc(ctx), { placeholder: true })}${rowOf(kid(el, 'm:sup'), sc(ctx), { placeholder: true })}</msubsup>`;
    case 'm:sPre':
      return `<mmultiscripts>${rowOf(kid(el, 'm:e'), ctx, { placeholder: true })}<mprescripts/>${rowOf(kid(el, 'm:sub'), sc(ctx))}${rowOf(kid(el, 'm:sup'), sc(ctx))}</mmultiscripts>`;
    case 'm:nary': {
      const chr = prop(pr, 'm:chr') || '∫';
      const loc = prop(pr, 'm:limLoc') || (INTEGRALS.has(chr) ? 'subSup' : 'undOvr');
      const sub = flag(pr, 'm:subHide') || isEmpty(kid(el, 'm:sub')) ? null : rowOf(kid(el, 'm:sub'), sc(ctx));
      const sup = flag(pr, 'm:supHide') || isEmpty(kid(el, 'm:sup')) ? null : rowOf(kid(el, 'm:sup'), sc(ctx));
      // Limits under and over a sum in a display; beside it inline — the
      // way Word sets a sum in running text, which `movablelimits` is.
      const op = mo(chr, ` largeop="true" movablelimits="${loc === 'undOvr' ? 'true' : 'false'}"`);
      let head;
      if (loc === 'undOvr') {
        head = sub && sup ? `<munderover>${op}${sub}${sup}</munderover>` : sub ? `<munder>${op}${sub}</munder>` : sup ? `<mover>${op}${sup}</mover>` : op;
      } else {
        head = sub && sup ? `<msubsup>${op}${sub}${sup}</msubsup>` : sub ? `<msub>${op}${sub}</msub>` : sup ? `<msup>${op}${sup}</msup>` : op;
      }
      return `<mrow>${head}${rowOf(kid(el, 'm:e'), ctx, { placeholder: true })}</mrow>`;
    }
    case 'm:d': {
      const beg = prop(pr, 'm:begChr') ?? '(';
      const end = prop(pr, 'm:endChr') ?? ')';
      const sep = prop(pr, 'm:sepChr') ?? '|';
      const fence = ' fence="true" stretchy="true" symmetric="true"';
      const parts = kids(el, 'm:e').map((e) => rowOf(e, ctx));
      const inner = parts.join(mo(sep, ' separator="true" stretchy="true"'));
      return `<mrow>${beg ? mo(beg, fence + ' form="prefix"') : ''}${inner || '<mrow></mrow>'}${end ? mo(end, fence + ' form="postfix"') : ''}</mrow>`;
    }
    case 'm:func': {
      // A function's name is upright — the runs that spell it, not what a
      // structure round it holds: the `n→∞` under `lim` stays italic.
      const fName = kid(el, 'm:fName');
      const spelt = (fName?.children || []).map((c) => (c.name === 'm:r' ? runToMathml(c, { ...ctx, upright: true }) : c.name ? elementToMathml(c, ctx) : '')).join('');
      const name = spelt ? `<mrow>${spelt}</mrow>` : PLACEHOLDER;
      const arg = kid(el, 'm:e');
      // A thin space between the name and a bare argument (sin θ), none
      // before a bracket (sin(x)) — where Word puts it.
      const fenced = kids(arg)[0]?.name === 'm:d';
      return `<mrow>${name}${fenced ? '' : '<mspace width="0.1667em"/>'}${mo('\u2061')}${rowOf(arg, ctx, { placeholder: true })}</mrow>`;
    }
    case 'm:acc': {
      const chr = prop(pr, 'm:chr') || '\u0302';
      if (chr === '\u0304' || chr === '\u0305' || chr === '¯' || chr === '‾') return barOver(rowOf(kid(el, 'm:e'), ctx, { placeholder: true }), true);
      const glyph = ACCENT_GLYPH[chr] ?? chr;
      return `<mover accent="true">${rowOf(kid(el, 'm:e'), ctx, { placeholder: true })}${mo(glyph, ' stretchy="true"')}</mover>`;
    }
    case 'm:bar': {
      const top = (prop(pr, 'm:pos') || 'bot') === 'top';
      const e = rowOf(kid(el, 'm:e'), ctx, { placeholder: true });
      // A rule over (or under) the base, drawn as the base's own border:
      // Chromium does not stretch a macron or an overline across a base,
      // and a bar that covers one letter of two is not Word's bar.
      return barOver(e, top);
    }
    case 'm:groupChr': {
      const chr = prop(pr, 'm:chr') || '⏟';
      const top = (prop(pr, 'm:pos') || 'bot') === 'top';
      const e = rowOf(kid(el, 'm:e'), ctx, { placeholder: true });
      return top ? `<mover>${e}${mo(chr, ' stretchy="true"')}</mover>` : `<munder>${e}${mo(chr, ' stretchy="true"')}</munder>`;
    }
    case 'm:limLow':
      return `<munder>${rowOf(kid(el, 'm:e'), ctx, { placeholder: true })}${rowOf(kid(el, 'm:lim'), sc(ctx))}</munder>`;
    case 'm:limUpp':
      return `<mover>${rowOf(kid(el, 'm:e'), ctx, { placeholder: true })}${rowOf(kid(el, 'm:lim'), sc(ctx))}</mover>`;
    case 'm:m': {
      const rows = kids(el, 'm:mr').map((mr) => `<mtr>${kids(mr, 'm:e').map((e) => `<mtd>${rowOf(e, ctx, { placeholder: true })}</mtd>`).join('')}</mtr>`);
      return `<mtable>${rows.join('')}</mtable>`;
    }
    case 'm:eqArr': {
      // `&` marks where the rows line up: the columns alternate right and
      // left of it, the way Word (and TeX's align) sets an equation array.
      const rows = kids(el, 'm:e').map((e) => {
        const cells = childrenToMathml(e, { ...ctx, amp: true }).split(AMP);
        return `<mtr>${cells.map((c, i) => `<mtd style="text-align:${cells.length === 1 ? 'center' : i % 2 ? 'left' : 'right'};padding:0.1em 0">${c ? `<mrow>${c}</mrow>` : ''}</mtd>`).join('')}</mtr>`;
      });
      return `<mtable>${rows.join('')}</mtable>`;
    }
    case 'm:box':
      return rowOf(kid(el, 'm:e'), ctx);
    case 'm:borderBox': {
      const side = (s) => (flag(pr, 'm:hide' + s) ? 'none' : '0.06em solid currentColor');
      return `<mrow style="border-top:${side('Top')};border-bottom:${side('Bot')};border-left:${side('Left')};border-right:${side('Right')};padding:0.12em 0.18em">${childrenToMathml(kid(el, 'm:e') || { children: [] }, ctx)}</mrow>`;
    }
    case 'm:phant': {
      const e = rowOf(kid(el, 'm:e'), ctx);
      return prop(pr, 'm:show') !== undefined && !flag(pr, 'm:show') ? `<mphantom>${e}</mphantom>` : e;
    }
    case 'w:del': case 'm:ctrlPr': case 'm:argPr':
      return '';
    default:
      if (/Pr$/.test(el.name)) return '';
      // Something this reader does not know: its words, never nothing.
      return tokensOfText(textUnder(el), ctx);
  }
}

function tokensOfText(text, ctx) {
  if (!text) return '';
  return `<mrow>${runToMathml({ name: 'm:r', attrs: {}, children: [{ name: 'm:t', attrs: {}, children: [{ text }] }] }, ctx)}</mrow>`;
}

/**
 * An equation fragment — `<m:oMathPara>…` or `<m:oMath>…` as it sits in the
 * paragraph — as MathML: one `<math>` per `m:oMath`, `display="block"` for
 * a display equation. `jc` (left, right, centred) rides a data attribute the
 * page aligns by.
 */
export function ommlToMathml(xml) {
  const root = parseXml(xml);
  const para = kid(root, 'm:oMathPara');
  const maths = para ? kids(para, 'm:oMath') : kids(root, 'm:oMath');
  const display = Boolean(para);
  const ctx = {};
  if (!maths.length) {
    // Not an equation at all — its text, so it is at least seen.
    return `<math display="${display ? 'block' : 'inline'}">${tokensOfText(textUnder(root), ctx)}</math>`;
  }
  return maths.map((m) => `<math display="${display ? 'block' : 'inline'}">${childrenToMathml(m, ctx) || PLACEHOLDER}</math>`).join('');
}

/* ── OMML → linear (UnicodeMath) ───────────────────────────────────────── */

/**
 * Word's linear format for an equation — what its equation editor shows in
 * "Linear" mode and what this suite's editor opens with. Each converter
 * answers `{ s, atom }`: `atom` means the text stands as one operand, so a
 * fraction or a script around it needs no brackets.
 */
const ATOMIC = /^[\p{L}\p{N}.]+$/u;
const lin = (s, atom = false) => ({ s, atom: atom || s.length <= 1 || ATOMIC.test(s) });
const wrap = (x) => (x.atom ? x.s : '(' + x.s + ')');
const wrapBody = (x) => (x.atom ? x.s : '〖' + x.s + '〗');

function seqLinear(node) {
  const parts = [];
  for (const c of node?.children || []) {
    if (c.text !== undefined) continue;
    const x = elementToLinear(c);
    if (x && x.s !== '') parts.push(x);
  }
  if (!parts.length) return lin('');
  if (parts.length === 1) return parts[0];
  // A space where a structure's last operand would otherwise run into the
  // next letters: `x^k a^(n-k)`, not `x^ka^(n-k)`.
  let s = parts[0].s;
  for (let i = 1; i < parts.length; i++) {
    if (!parts[i - 1].run && /[\p{L}\p{N}]$/u.test(parts[i - 1].s) && /^[\p{L}\p{N}]/u.test(parts[i].s)) s += ' ';
    s += parts[i].s;
  }
  return lin(s);
}
const partLinear = (node) => (node ? seqLinear(node) : lin(''));

function elementToLinear(el) {
  const pr = kid(el, el.name + 'Pr');
  switch (el.name) {
    case 'm:r': {
      const text = textUnder(el);
      if (flag(kid(el, 'm:rPr'), 'm:nor')) return lin('"' + text + '"', true);
      return { ...lin(text), run: true };
    }
    case 'm:oMath': case 'm:e': case 'm:num': case 'm:den': case 'm:sub': case 'm:sup': case 'm:deg': case 'm:lim': case 'm:fName':
    case 'w:ins': case 'w:smartTag': case 'w:customXml': case 'w:hyperlink':
      return seqLinear(el);
    case 'm:f': {
      const type = prop(pr, 'm:type') || 'bar';
      const op = type === 'noBar' ? '¦' : type === 'lin' ? '⊘' : type === 'skw' ? '∕' : '/';
      // Not an atom: `(a/b)^2` needs its brackets back.
      return lin(wrap(partLinear(kid(el, 'm:num'))) + op + wrap(partLinear(kid(el, 'm:den'))), false);
    }
    case 'm:rad': {
      const e = partLinear(kid(el, 'm:e'));
      const deg = kid(el, 'm:deg');
      if (flag(pr, 'm:degHide') || isEmpty(deg)) return lin('√' + wrap(e), true);
      return lin('√(' + partLinear(deg).s + '&' + e.s + ')', true);
    }
    case 'm:sSup': return lin(wrap(partLinear(kid(el, 'm:e'))) + '^' + wrap(partLinear(kid(el, 'm:sup'))), true);
    case 'm:sSub': return lin(wrap(partLinear(kid(el, 'm:e'))) + '_' + wrap(partLinear(kid(el, 'm:sub'))), true);
    case 'm:sSubSup':
      return lin(wrap(partLinear(kid(el, 'm:e'))) + '_' + wrap(partLinear(kid(el, 'm:sub'))) + '^' + wrap(partLinear(kid(el, 'm:sup'))), true);
    case 'm:sPre':
      return lin('_' + wrap(partLinear(kid(el, 'm:sub'))) + '^' + wrap(partLinear(kid(el, 'm:sup'))) + ' ' + wrap(partLinear(kid(el, 'm:e'))));
    case 'm:nary': {
      const chr = prop(pr, 'm:chr') || '∫';
      const sub = flag(pr, 'm:subHide') ? lin('') : partLinear(kid(el, 'm:sub'));
      const sup = flag(pr, 'm:supHide') ? lin('') : partLinear(kid(el, 'm:sup'));
      const body = partLinear(kid(el, 'm:e'));
      return lin(chr + (sub.s ? '_' + wrap(sub) : '') + (sup.s ? '^' + wrap(sup) : '') + '▒' + wrapBody(body));
    }
    case 'm:d': {
      const beg = prop(pr, 'm:begChr') ?? '(';
      const end = prop(pr, 'm:endChr') ?? ')';
      const sep = prop(pr, 'm:sepChr') ?? '|';
      const parts = kids(el, 'm:e').map((e) => seqLinear(e).s);
      return lin((beg || '├') + parts.join(sep) + (end || '┤'), true);
    }
    case 'm:func': {
      const name = partLinear(kid(el, 'm:fName'));
      const e = partLinear(kid(el, 'm:e'));
      // A function Word knows by name reads back without the invisible
      // function-application mark; any other name keeps it, so it parses
      // back as the function it was.
      const known = FUNCTION_NAMES.has(/^\p{L}+/u.exec(name.s)?.[0] ?? '');
      // An argument of more than one operand is grouped invisibly, \u3016\u2026\u3017, as
      // Word writes `sin\u2061\u3016n\u03c0x/L\u3017`.
      const arg = e.atom ? e.s : '\u3016' + e.s + '\u3017';
      return lin(name.s + (known ? (/^[(\[{|\u3016]/.test(arg) ? '' : ' ') : '\u2061') + arg);
    }
    case 'm:acc': return lin(wrap(partLinear(kid(el, 'm:e'))) + (prop(pr, 'm:chr') || '\u0302'), true);
    case 'm:bar': return lin(((prop(pr, 'm:pos') || 'bot') === 'top' ? '¯' : '▁') + wrap(partLinear(kid(el, 'm:e'))), true);
    case 'm:groupChr': return lin((prop(pr, 'm:chr') || '⏟') + wrap(partLinear(kid(el, 'm:e'))), true);
    case 'm:limLow': return lin(wrap(partLinear(kid(el, 'm:e'))) + '┬' + wrap(partLinear(kid(el, 'm:lim'))), true);
    case 'm:limUpp': return lin(wrap(partLinear(kid(el, 'm:e'))) + '┴' + wrap(partLinear(kid(el, 'm:lim'))), true);
    case 'm:m':
      return lin('■(' + kids(el, 'm:mr').map((mr) => kids(mr, 'm:e').map((e) => seqLinear(e).s).join('&')).join('@') + ')', true);
    case 'm:eqArr':
      return lin('█(' + kids(el, 'm:e').map((e) => seqLinear(e).s).join('@') + ')', true);
    case 'm:box': return lin('□' + wrap(partLinear(kid(el, 'm:e'))), true);
    case 'm:borderBox': return lin('▭' + wrap(partLinear(kid(el, 'm:e'))), true);
    case 'm:phant': return lin('⟡' + wrap(partLinear(kid(el, 'm:e'))), true);
    case 'w:del': case 'm:ctrlPr': case 'm:argPr': return lin('');
    default:
      if (/Pr$/.test(el.name)) return lin('');
      return lin(textUnder(el));
  }
}

/** Names Word sets upright as functions when typed in the linear form. */
export const FUNCTION_NAMES = new Set([
  'sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'sinh', 'cosh', 'tanh', 'coth', 'sech', 'csch',
  'arcsin', 'arccos', 'arctan', 'arccot', 'arcsec', 'arccsc',
  'log', 'ln', 'lg', 'exp', 'lim', 'max', 'min', 'sup', 'inf', 'det', 'dim', 'gcd', 'lcm', 'arg', 'deg', 'hom', 'ker', 'Pr', 'mod',
]);

/** An equation fragment in Word's linear format. Several `m:oMath` in one paragraph are joined by a space. */
export function ommlToLinear(xml) {
  const root = parseXml(xml);
  const para = kid(root, 'm:oMathPara');
  const maths = para ? kids(para, 'm:oMath') : kids(root, 'm:oMath');
  if (!maths.length) return textUnder(root);
  return maths.map((m) => seqLinear(m).s).join(' ');
}

/**
 * The linear form spelt in characters a base-14 PDF font can draw — the
 * printout's last resort when no MathML layout is to hand. Square roots,
 * sums, Greek letters become their names; everything else stays.
 */
const ASCII_MATH = {
  '√': 'sqrt', '∛': 'cbrt', '∑': 'sum', '∏': 'prod', '∫': 'int', '∬': 'iint', '∮': 'oint', '▒': ' ', '〖': '(', '〗': ')',
  '−': '-', '≤': '<=', '≥': '>=', '≠': '!=', '≈': '~', '→': '->', '←': '<-', '⇒': '=>', '∞': 'inf', '∂': 'd', '∇': 'nabla',
  '┬': '_', '┴': '^', '■': 'matrix', '█': '', '□': '', '▭': '', '¯': 'bar', '▁': 'underbar', '⏟': '', '⏞': '', '⟡': '',
  '├': '', '┤': '', '¦': '/', '⊘': '/', '∕': '/', '\u2061': ' ', '⋅': '*', '∈': ' in ', '∀': 'for all ', '∃': 'exists ',
  'α': 'alpha', 'β': 'beta', 'γ': 'gamma', 'δ': 'delta', 'ε': 'epsilon', 'ζ': 'zeta', 'η': 'eta', 'θ': 'theta', 'ι': 'iota',
  'κ': 'kappa', 'λ': 'lambda', 'μ': 'mu', 'ν': 'nu', 'ξ': 'xi', 'π': 'pi', 'ρ': 'rho', 'σ': 'sigma', 'τ': 'tau', 'υ': 'upsilon',
  'φ': 'phi', 'χ': 'chi', 'ψ': 'psi', 'ω': 'omega', 'Γ': 'Gamma', 'Δ': 'Delta', 'Θ': 'Theta', 'Λ': 'Lambda', 'Ξ': 'Xi',
  'Π': 'Pi', 'Σ': 'Sigma', 'Φ': 'Phi', 'Ψ': 'Psi', 'Ω': 'Omega',
};
export function asciiLinear(linear) {
  return [...String(linear)].map((c) => {
    if (ASCII_MATH[c] !== undefined) return ASCII_MATH[c];
    if (/[\u0300-\u036F\u20D0-\u20FF]/.test(c)) return '';
    return c;
  }).join('').replace(/ {2,}/g, ' ').trim();
}

export { escapeText as escapeMathText, escapeAttr as escapeMathAttr, ACCENT_GLYPH };
