// A deck.
//
// Opens a .pptx over the preserving OOXML package, resolves the theme, master
// and layout chain, and hands out slide scenes. Edits rewrite only the XML of
// the part they touch — a text edit rewrites one slide, not the presentation —
// so everything we do not model (animations, transitions, embedded media,
// custom XML, macros, signatures) survives a round trip untouched.
//
// That is the same bargain the spreadsheet and document engines make, and it is
// the only honest way to ship a presentation editor without twenty years of
// feature work behind it.

import { OoxmlPackage } from '@rutba/ooxml/package';
import { parse, kids, first, all, escapeXml } from '@rutba/office-formats/xml';
import { emuToPx, pxToEmu, ptToSz } from './units.js';
import { readSlideScene, readXfrm, readTextBody, placeholderOf, sceneText, composeGroupChild, REFLECTION_PRESETS, readLevels, mergeLevels, readFill } from './slide.js';
import { THEMES, PALETTES, FONT_PAIRS, EFFECT_PRESETS, COLOUR_SLOTS, themePartXml, clrSchemeXml, fontSchemeXml, fmtSchemeXml, masterBackgroundXml, clrMapAttrs, variantsOf, themeById } from './themes.js';
import { slideXml } from './build.js';
import { chartPartXml } from '@rutba/ooxml/build';
import { readTransition, withTransition, transitionBlock, insertTransition, transitionRange } from './motion.js';
import { readAnimations, addAnimation, setAnimation, removeAnimation, moveAnimation, removeShapeAnimations, pruneAnimations } from './timing.js';
import { parseChartXml } from '@rutba/drawing';

const A = (n) => `a:${n}`;
const P = (n) => `p:${n}`;

const CT = {
  slide: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
  layout: 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
  master: 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
  notes: 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
  chart: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
};
const REL = {
  slide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
  layout: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
  master: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  notes: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  chart: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
};
/** The chart kinds `chartPartXml` can write — Insert → Chart and its data editor may only ask for one of these. */
const CHART_KINDS = ['column', 'bar', 'line', 'area', 'pie', 'doughnut'];/** The picture types PowerPoint itself embeds; anything else is converted first. */
const IMAGE_EXTENSIONS = {
  'image/png': 'png',
  'image/jpeg': 'jpeg',
  'image/jpg': 'jpeg',
  'image/gif': 'gif',
  'image/bmp': 'bmp',
};


/** Resolve a relationship target against the part that declares it. */
function resolveTarget(fromPart, target) {
  if (!target) return null;
  if (target.startsWith('/')) return target.slice(1);
  const dir = fromPart.split('/').slice(0, -1).join('/');
  const parts = `${dir}/${target}`.split('/');
  const out = [];
  for (const seg of parts) {
    if (seg === '.' || seg === '') continue;
    if (seg === '..') out.pop();
    else out.push(seg);
  }
  return out.join('/');
}

/**
 * The spTree's top-level children in drawing order, each with its XML range
 * and its id. Bottom first: the tree draws in order, so the last child is on
 * top. A group counts as one child; the shapes inside it are its own affair.
 */
function topLevelShapes(xml, bounds = null) {
  // The whole slide's own top level, unless the caller names a narrower
  // range — a group's own children, one level deep, are found the same way
  // between the end of its `grpSpPr` and its own closing tag, since a group
  // has no `<p:spTree>` of its own to anchor the default range.
  let start;
  let end;
  if (bounds) {
    ({ start, end } = bounds);
  } else {
    start = xml.indexOf('<p:spTree');
    end = xml.lastIndexOf('</p:spTree>');
    if (start < 0 || end < 0) return [];
  }
  const re = /<(\/?)p:(sp|pic|graphicFrame|cxnSp|grpSp)\b[^>]*?(\/?)>/g;
  re.lastIndex = start;
  const out = [];
  const stack = [];
  let m;
  while ((m = re.exec(xml)) && m.index < end) {
    const closing = m[1] === '/';
    if (!closing) {
      if (m[3] === '/') {
        if (!stack.length) out.push({ tag: m[2], start: m.index, end: m.index + m[0].length, id: null });
        continue;
      }
      stack.push({ tag: m[2], start: m.index });
      continue;
    }
    const open = stack.pop();
    if (!open) continue;
    if (!stack.length) {
      const id = /<p:cNvPr\b[^>]*\bid="([^"]+)"/.exec(xml.slice(open.start, m.index))?.[1] ?? null;
      out.push({ tag: open.tag, start: open.start, end: m.index + m[0].length, id });
    }
  }
  return out;
}

class Theme {
  constructor(xml, clrMap) {
    this.colors = {};
    this.fonts = { major: 'Calibri Light', minor: 'Calibri' };
    this.clrMap = clrMap || {};
    /** The names the gallery ticks by: the theme's, its colours', its fonts' and its effects'. */
    this.name = null;
    this.colorName = null;
    this.fontName = null;
    this.effectName = null;
    /** The format scheme's four lists, as nodes a style reference picks from by index. */
    this.fmt = { fills: [], lines: [], effects: [], bgFills: [] };
    /** Set on a copy by `withPh`: what `phClr` means while a format-scheme entry is read. */
    this.ph = null;
    if (!xml) return;
    const root = parse(xml);
    this.name = first(root, A('theme'))?.attrs.name ?? null;
    const scheme = first(root, A('clrScheme'));
    this.colorName = scheme?.attrs.name ?? null;
    for (const node of kids(scheme || { children: [] })) {
      const name = node.name.replace(/^a:/, '');
      const child = kids(node)[0];
      if (!child) continue;
      if (child.name === A('srgbClr')) this.colors[name] = `#${child.attrs.val}`;
      else if (child.name === A('sysClr')) this.colors[name] = `#${child.attrs.lastClr || '000000'}`;
    }
    const fontScheme = first(root, A('fontScheme'));
    const major = first(first(fontScheme || { children: [] }, A('majorFont')) || { children: [] }, A('latin'));
    const minor = first(first(fontScheme || { children: [] }, A('minorFont')) || { children: [] }, A('latin'));
    if (major?.attrs.typeface) this.fonts.major = major.attrs.typeface;
    if (minor?.attrs.typeface) this.fonts.minor = minor.attrs.typeface;
    this.fontName = fontScheme?.attrs.name ?? null;
    const fmt = first(root, A('fmtScheme'));
    if (fmt) {
      this.effectName = fmt.attrs.name ?? null;
      this.fmt = {
        fills: kids(first(fmt, A('fillStyleLst')) || { children: [] }),
        lines: kids(first(fmt, A('lnStyleLst')) || { children: [] }),
        effects: kids(first(fmt, A('effectStyleLst')) || { children: [] }),
        bgFills: kids(first(fmt, A('bgFillStyleLst')) || { children: [] }),
      };
    }
  }

  /** This theme reading `phClr` as the given colour — how a style reference colours the entry it picks. */
  withPh(hex) {
    const t = Object.create(this);
    t.ph = hex;
    return t;
  }

  /** A fill style by a `fillRef`/`bgRef` index: 1–999 the shape fills, 1001– the backgrounds. */
  fillStyle(idx) {
    const i = Number(idx) || 0;
    if (i >= 1001) return this.fmt.bgFills[i - 1001] || null;
    return i >= 1 ? this.fmt.fills[i - 1] || null : null;
  }

  lineStyle(idx) {
    const i = Number(idx) || 0;
    return i >= 1 ? this.fmt.lines[i - 1] || null : null;
  }

  effectStyle(idx) {
    const i = Number(idx) || 0;
    return i >= 1 ? this.fmt.effects[i - 1] || null : null;
  }

  /** `schemeClr val="tx1"` → the theme colour, through the master's map. */
  color(name) {
    if (name === 'phClr') return this.ph || null;
    const mapped = this.clrMap[name] || name;
    return (
      this.colors[mapped] ||
      this.colors[name] ||
      { tx1: '#000000', tx2: '#44546a', bg1: '#ffffff', bg2: '#e7e6e6', dk1: '#000000', lt1: '#ffffff' }[mapped] ||
      null
    );
  }

  font(typeface) {
    if (/^\+mj-/.test(typeface)) return this.fonts.major;
    if (/^\+mn-/.test(typeface)) return this.fonts.minor;
    return typeface;
  }
}

export class Deck {
  /** @param {OoxmlPackage} pkg */
  constructor(pkg) {
    this.pkg = pkg;
    /** Parsed slides by part, keyed on their XML — see `slide()`. */
    this._scenes = new Map();
    this.dirty = false;
    /**
     * Bumped whenever a master, a layout or a theme changes: every slide
     * draws from them, so a scene or a thumbnail kept against a slide's own
     * XML alone would go on showing the old design.
     */
    this.designStamp = 0;
    /** Undo and redo: whole-package snapshots, cheap because a part's bytes are replaced, never mutated. */
    this._undo = [];
    this._redo = [];
    this.#load();
  }

  // ---- undo ----------------------------------------------------------------
  //
  // A snapshot is each entry's own bytes and flags as they stand. Every edit
  // replaces a part's buffer rather than writing into it, so holding the old
  // buffer is holding the old part — a snapshot of a fifteen-megabyte deck
  // is a list of references, not fifteen megabytes.

  /** The package as it stands, to hand back to `pushUndo` once an edit has gone through. */
  snapshot() {
    return {
      entries: this.pkg.entries.map((e) => ({ e, raw: e._raw, modified: e.modified, size: e.uncompressedSize, crc: e.crc })),
      dirty: this.dirty,
    };
  }

  /** One undo step: the state before an edit. A new edit forgets what was undone. */
  pushUndo(snap) {
    if (!snap) return;
    this._undo.push(snap);
    if (this._undo.length > 100) this._undo.shift();
    this._redo = [];
  }

  get canUndo() {
    return this._undo.length > 0;
  }

  get canRedo() {
    return this._redo.length > 0;
  }

  undo() {
    const snap = this._undo.pop();
    if (!snap) return false;
    this._redo.push(this.snapshot());
    this.#restore(snap);
    return true;
  }

  redo() {
    const snap = this._redo.pop();
    if (!snap) return false;
    this._undo.push(this.snapshot());
    this.#restore(snap);
    return true;
  }

  #restore(snap) {
    this.pkg.entries = snap.entries.map(({ e }) => e);
    for (const { e, raw, modified, size, crc } of snap.entries) {
      e._raw = raw;
      e.modified = modified;
      e.uncompressedSize = size;
      e.crc = crc;
    }
    this.pkg.byName = new Map(this.pkg.entries.map((e) => [e.name, e]));
    this.pkg._contentTypes = null;
    this.dirty = true;
    this._scenes.clear();
    this.#load();
    this.designStamp++;
  }

  /** A slide index to its part; a master's or a layout's part name to itself, so Slide Master view edits them with the same verbs. */
  #partOf(target) {
    if (typeof target === 'string') {
      return /^ppt\/(slideMasters\/slideMaster|slideLayouts\/slideLayout)\d+\.xml$/.test(target) && this.pkg.has(target) ? target : null;
    }
    return this.slideParts[target]?.part || null;
  }

  /** Whether a part is a master's or a layout's — an edit to one restyles every slide on it. */
  static isDesignPart(part) {
    return /^ppt\/(slideMasters|slideLayouts|theme)\//.test(String(part || ''));
  }

  /** Every cache that read a master, a layout or a theme, dropped. */
  #designChanged() {
    this.designStamp++;
    this.themes.clear();
    this.layouts.clear();
    this._scenes.clear();
    this.dirty = true;
  }

  static open(bytes) {
    const buf = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes);
    return new Deck(OoxmlPackage.read(buf));
  }

  static async openAsync(bytes) {
    return new Deck(await OoxmlPackage.readAsync(bytes));
  }

  // `rels()` hands back the raw attributes — `Id`, `Type`, `Target`,
  // `TargetMode` — so normalise once here rather than at every call site.
  #relMap(partName) {
    const map = new Map();
    try {
      for (const r of this.pkg.rels(partName) || []) {
        const external = r.TargetMode === 'External';
        map.set(r.Id, {
          id: r.Id,
          type: r.Type,
          target: r.Target,
          mode: r.TargetMode || 'Internal',
          resolved: external ? r.Target : resolveTarget(partName, r.Target),
        });
      }
    } catch {
      /* a part with no rels is normal */
    }
    return map;
  }

  #load() {
    const main = 'ppt/presentation.xml';
    if (!this.pkg.has(main)) {
      throw new Error('not a presentation: ppt/presentation.xml is missing');
    }
    const xml = this.pkg.text(main);
    const root = parse(xml);
    const pres = first(root, P('presentation')) || root;
    const sz = first(pres, P('sldSz'));
    this.size = {
      cx: Number(sz?.attrs.cx || 12192000),
      cy: Number(sz?.attrs.cy || 6858000),
      width: emuToPx(sz?.attrs.cx || 12192000),
      height: emuToPx(sz?.attrs.cy || 6858000),
    };
    this.rels = this.#relMap(main);

    const ids = kids(first(pres, P('sldIdLst')) || { children: [] }, P('sldId'));
    this.slideParts = ids
      .map((s) => ({ id: s.attrs.id, rId: s.attrs['r:id'], part: this.rels.get(s.attrs['r:id'])?.resolved }))
      .filter((s) => s.part && this.pkg.has(s.part));

    this.themes = new Map();
    this.layouts = new Map();
    this.masters = new Map();

    // Sections — PowerPoint 2010's p14:sectionLst, an extension on the
    // presentation part — as lists of sldId ids. A section naming a slide
    // that is not in the show is tolerated on the way in and dropped on the
    // way out, which is what PowerPoint does with the same file.
    this._sections = readSectionList(xml);
  }

  get slideCount() {
    return this.slideParts.length;
  }

  // ---- sections ----------------------------------------------------------

  /**
   * The deck's sections in order, each with the indexes of the slides it
   * holds. Empty when the deck has none: every slide is then in no section,
   * which the window draws as a plain strip.
   */
  sections() {
    const pos = new Map(this.slideParts.map((s, i) => [String(s.id), i]));
    return this._sections.map((s, index) => ({
      index,
      name: s.name,
      id: s.id,
      slides: s.slideIds.map((id) => pos.get(String(id))).filter((i) => i != null),
    }));
  }

  /** The index of the section a slide is in, or -1. */
  sectionOf(slideIndex) {
    return this.sections().findIndex((s) => s.slides.includes(slideIndex));
  }

  /**
   * A section starting at the slide at `index`, as Home → Section → Add
   * Section makes one before the selected slide. A deck with no sections
   * gets a "Default Section" for the slides before it first, as PowerPoint
   * does, since every slide of a sectioned deck is in some section. The
   * section the slide was in keeps the slides before it. A slide that already
   * starts a section starts that one still: nothing changes and its index is
   * returned.
   * @returns {number} the index of the section that starts at the slide
   */
  addSection(index, name = 'Untitled Section') {
    const entry = this.slideParts[index];
    if (!entry) throw new RangeError(`no slide at index ${index}`);
    const ids = this.slideParts.map((s) => String(s.id));
    if (!this._sections.length) {
      const before = ids.slice(0, index);
      const next = [];
      if (before.length) next.push({ name: 'Default Section', id: sectionGuid(), slideIds: before });
      next.push({ name, id: sectionGuid(), slideIds: ids.slice(index) });
      this.#writeSections(next);
      return next.length - 1;
    }
    const sections = this.sections();
    const at = sections.findIndex((s) => s.slides.includes(index));
    if (at < 0) {
      // A slide no section lists — a file another program wrote — joins
      // the sections the way an added slide does, then the split is retried.
      this.#syncSections();
      return this.addSection(index, name);
    }
    if (sections[at].slides[0] === index) return at;
    const own = this._sections[at];
    const keep = own.slideIds.filter((id) => (this.slideParts.findIndex((s) => String(s.id) === String(id))) < index);
    const moved = own.slideIds.filter((id) => !keep.includes(id));
    const next = this._sections.map((s) => ({ ...s, slideIds: [...s.slideIds] }));
    next[at] = { ...own, slideIds: keep };
    next.splice(at + 1, 0, { name, id: sectionGuid(), slideIds: moved });
    this.#writeSections(next);
    return at + 1;
  }

  /** A section's new name. */
  renameSection(index, name) {
    const own = this._sections[index];
    if (!own) throw new RangeError(`no section at index ${index}`);
    const clean = String(name ?? '').trim() || 'Untitled Section';
    if (clean === own.name) return false;
    const next = this._sections.map((s, i) => (i === index ? { ...s, name: clean } : { ...s }));
    this.#writeSections(next);
    return true;
  }

  /**
   * A section taken away, its slides staying in the show: they join the
   * section before it, or — for the first — the one after. The last section
   * standing takes the list with it, so the deck is back to having none.
   */
  removeSection(index) {
    const own = this._sections[index];
    if (!own) throw new RangeError(`no section at index ${index}`);
    if (this._sections.length === 1) return this.removeAllSections();
    const next = this._sections.map((s) => ({ ...s, slideIds: [...s.slideIds] }));
    const into = index > 0 ? index - 1 : 1;
    next[into].slideIds = index > 0 ? [...next[into].slideIds, ...own.slideIds] : [...own.slideIds, ...next[into].slideIds];
    next.splice(index, 1);
    this.#writeSections(next);
    return true;
  }

  /** Every section gone; the slides stay as they are. */
  removeAllSections() {
    if (!this._sections.length) return false;
    this.#writeSections([]);
    return true;
  }

  /**
   * The section list written into presentation.xml as PowerPoint writes it:
   * one p:ext with PowerPoint's section-list URI, the p14 namespace declared
   * on the list, each section a name, a GUID and its slide ids in show order.
   * An empty list takes the extension out, and the extLst with it when
   * nothing else is in there.
   */
  #writeSections(sections) {
    const presXml = this.pkg.text('ppt/presentation.xml');
    const order = new Map(this.slideParts.map((s, i) => [String(s.id), i]));
    const kept = sections.map((s) => ({
      name: s.name,
      id: s.id || sectionGuid(),
      slideIds: [...new Set(s.slideIds.map(String))].filter((id) => order.has(id)).sort((a, b) => order.get(a) - order.get(b)),
    }));
    const extRe = /<p:ext\b[^>]*uri="\{521415D9-36F7-43E2-AB2F-B90AF26B5E84\}"[^>]*>[\s\S]*?<\/p:ext>/;
    let next;
    if (!kept.length) {
      next = presXml.replace(extRe, '').replace(/<p:extLst>\s*<\/p:extLst>/, '');
    } else {
      const ext =
        `<p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">` +
        `<p14:sectionLst xmlns:p14="${P14_NS}">` +
        kept.map((s) => `<p14:section name="${escapeXml(s.name)}" id="${s.id}">${s.slideIds.length ? `<p14:sldIdLst>${s.slideIds.map((id) => `<p14:sldId id="${id}"/>`).join('')}</p14:sldIdLst>` : '<p14:sldIdLst/>'}</p14:section>`).join('') +
        `</p14:sectionLst></p:ext>`;
      if (extRe.test(presXml)) next = presXml.replace(extRe, ext);
      else if (/<\/p:extLst>\s*<\/p:presentation>/.test(presXml)) next = presXml.replace(/<\/p:extLst>(\s*<\/p:presentation>)/, `${ext}</p:extLst>$1`);
      else next = presXml.replace(/<\/p:presentation>\s*$/, `<p:extLst>${ext}</p:extLst></p:presentation>`);
    }
    if (next === presXml) return;
    this.pkg.write_('ppt/presentation.xml', Buffer.from(next, 'utf8'));
    this.dirty = true;
    this._sections = kept;
  }

  /**
   * The sections brought into step with the slide list after a slide is
   * added, removed or moved: a slide no section lists joins the section of
   * the slide before it (the first section when it is first), a slide gone
   * from the show goes from its section, and each section lists its slides
   * in show order. `loose` names slide ids to place afresh — a moved slide
   * belongs where it landed, not where it came from.
   */
  #syncSections(loose = []) {
    if (!this._sections.length) return;
    const order = this.slideParts.map((s) => String(s.id));
    const known = new Set(order);
    const free = new Set(loose.map(String));
    const lists = this._sections.map((s) => s.slideIds.map(String).filter((id) => known.has(id) && !free.has(id)));
    const where = new Map();
    lists.forEach((ids, si) => ids.forEach((id) => where.set(id, si)));
    let previous = 0;
    for (const id of order) {
      if (where.has(id)) { previous = where.get(id); continue; }
      lists[previous].push(id);
      where.set(id, previous);
    }
    const next = this._sections.map((s, i) => ({ ...s, slideIds: lists[i] }));
    const same = next.every((s, i) => s.slideIds.join() === this._sections[i].slideIds.map(String).join());
    if (!same) this.#writeSections(next);
  }

  #masterFor(layoutPart) {
    const rels = this.#relMap(layoutPart);
    for (const r of rels.values()) if (r.type === REL.master) return r.resolved;
    return null;
  }

  #layoutFor(slidePart) {
    const rels = this.#relMap(slidePart);
    for (const r of rels.values()) if (r.type === REL.layout) return r.resolved;
    return null;
  }

  #themePartOf(masterPart) {
    for (const r of this.#relMap(masterPart).values()) if (r.type === REL.theme) return r.resolved;
    return null;
  }

  #themeFor(masterPart) {
    if (!masterPart) return new Theme(null, {});
    if (this.themes.has(masterPart)) return this.themes.get(masterPart);
    const rels = this.#relMap(masterPart);
    let themePart = null;
    for (const r of rels.values()) if (r.type === REL.theme) themePart = r.resolved;
    // The colour map lives on the master, and decides what tx1/bg1 point at.
    let clrMap = {};
    if (this.pkg.has(masterPart)) {
      const m = parse(this.pkg.text(masterPart));
      const cm = first(m, P('clrMap'));
      if (cm) clrMap = { ...cm.attrs };
    }
    const theme = new Theme(themePart && this.pkg.has(themePart) ? this.pkg.text(themePart) : null, clrMap);
    this.themes.set(masterPart, theme);
    return theme;
  }

  /**
   * A master's text styles — title, body and other — and the presentation's
   * default text style, as nine levels each: the bottom of the cascade a
   * run without a size, colour or face falls through to.
   */
  #textStyles(masterPart, theme, cache = this.layouts) {
    const key = `styles:${masterPart}`;
    if (cache.has(key)) return cache.get(key);
    const out = { title: null, body: null, other: null, default: null };
    if (masterPart && this.pkg.has(masterPart)) {
      const tx = first(parse(this.pkg.text(masterPart)), P('txStyles'));
      if (tx) {
        out.title = readLevels(first(tx, P('titleStyle')), theme);
        out.body = readLevels(first(tx, P('bodyStyle')), theme);
        out.other = readLevels(first(tx, P('otherStyle')), theme);
      }
    }
    const def = first(parse(this.pkg.text('ppt/presentation.xml')), P('defaultTextStyle'));
    if (def) out.default = readLevels(def, theme);
    cache.set(key, out);
    return out;
  }

  /**
   * One shape's text style, level by level, lowest layer first: the theme's
   * own text colour and face, the master's title, body or other style (the
   * presentation's default for a shape that fills no placeholder), the
   * master's placeholder's list style, the layout's, a shape style's font
   * reference, and the shape's own list style. What a run states for itself
   * is laid over this when it is drawn, never baked into it — so a new
   * theme or a master edit reaches every run that did not choose otherwise.
   */
  #cascade(shape, hits, styles, theme) {
    const ph = shape.placeholder;
    const t = ph?.type;
    const kind = !ph ? 'default' : t === 'title' || t === 'ctrTitle' ? 'title' : ['body', 'subTitle', 'obj'].includes(t) ? 'body' : 'other';
    const face = kind === 'title' ? theme.fonts.major : theme.fonts.minor;
    const themeBase = Array.from({ length: 9 }, () => ({ color: theme.color('tx1') || undefined, font: face, linkColor: theme.color('hlink') || undefined }));
    const base = kind === 'default' ? mergeLevels(styles.other, styles.default) : styles[kind];
    const fontRef = shape.styleText ? Array.from({ length: 9 }, () => shape.styleText) : null;
    const above = mergeLevels(hits.master?.text?.levels, hits.layout?.text?.levels, fontRef, shape.text?.levels);
    const merged = mergeLevels(themeBase, base, above);
    // A subtitle or a centred title keeps the body style's size and colour
    // but not its bullets and hanging indents, unless a layer of its own
    // asks for them — every title-slide layout PowerPoint ships says so, and
    // a deck whose layout forgot should not grow a bullet on its subtitle.
    // The footer band — date, footer, number — is small unless a placeholder
    // of its own says otherwise; a master without one should not hand it
    // the other style's eighteen points.
    if (t === 'dt' || t === 'ftr' || t === 'sldNum') {
      merged.forEach((lv, i) => { if (above[i].size === undefined) delete lv.size; });
    }
    if (t === 'subTitle' || t === 'ctrTitle') {
      merged.forEach((lv, i) => {
        if (above[i].bullet === undefined) delete lv.bullet;
        if (above[i].indent === undefined) delete lv.indent;
        if (above[i].hanging === undefined) delete lv.hanging;
      });
    }
    return merged;
  }

  /** Placeholder shapes of a layout or master, indexed for inheritance. */
  #placeholders(partName, theme, cache = this.layouts) {
    if (!partName || !this.pkg.has(partName)) return new Map();
    const key = `${partName}`;
    if (cache.has(key)) return cache.get(key);
    const scene = readSlideScene(this.pkg.text(partName), { theme });
    const map = new Map();
    for (const s of scene.shapes) {
      if (!s.placeholder) continue;
      const { type, idx } = s.placeholder;
      if (idx != null) map.set(`idx:${idx}`, s);
      if (type) map.set(`type:${type}`, s);
    }
    map.set('#background', scene.background);
    cache.set(key, map);
    return map;
  }

  /**
   * One slide, fully resolved: background, shapes, inherited geometry and
   * text style, and picture sources as part names the caller can read.
   */
  slide(index) {
    if (typeof index === 'string') return this.partScene(index);
    const entry = this.slideParts[index];
    if (!entry) throw new RangeError(`no slide at index ${index}`);
    const slidePart = entry.part;
    // Parsed once per VERSION of the slide's XML (and its notes', and the
    // design it draws from): the outline, the thumbnails and the model all
    // ask for every slide, and a fifteen-megabyte deck of nineteen slides
    // parsed nineteen slides three times per keystroke. The key is the XML
    // itself, so a stale entry is not expressible; an edit writes new XML
    // and misses, and a master or theme edit moves the design stamp.
    const slideXml = this.pkg.text(slidePart);
    const rels = this.#relMap(slidePart);
    const notesPart = [...rels.values()].find((r) => r.type === REL.notes && this.pkg.has(r.resolved))?.resolved || null;
    const notesXml = notesPart ? this.pkg.text(notesPart) : '';
    const cacheKey = index + ':' + this.designStamp + ':' + slideXml.length + ':' + notesXml.length + ':' + slideXml + notesXml;
    const cached = this._scenes.get(slidePart);
    if (cached && cached.key === cacheKey) return cached.scene;
    const result = this.#scene(index);
    this._scenes.set(slidePart, { key: cacheKey, scene: result });
    return result;
  }

  /**
   * A slide's scene, built. `design` is for a preview — Design → Themes
   * drawing this slide as it would look under another theme — and carries
   * the theme to read it with, the master background to show and whether
   * the layout's own background goes (a new theme takes it away); a
   * preview reads the master and layout afresh rather than from the caches
   * the real design fills.
   */
  #scene(index, design = null) {
    const entry = this.slideParts[index];
    const slidePart = entry.part;
    const layoutPart = this.#layoutFor(slidePart);
    const masterPart = layoutPart ? this.#masterFor(layoutPart) : null;
    const theme = design?.theme || this.#themeFor(masterPart);
    const cache = design ? new Map() : this.layouts;
    const layoutPh = this.#placeholders(layoutPart, theme, cache);
    const masterPh = this.#placeholders(masterPart, theme, cache);
    const styles = this.#textStyles(masterPart, theme, cache);
    const rels = this.#relMap(slidePart);

    // The layout's and the master's placeholder a slide's placeholder
    // matches, each looked up by index first and then by type.
    const keysOf = placeholderKeys;
    const hitsFor = (ph) => {
      if (!ph) return { layout: null, master: null };
      const keys = keysOf(ph);
      const find = (source) => { for (const k of keys) { const hit = source.get(k); if (hit) return hit; } return null; };
      return { layout: find(layoutPh), master: find(masterPh) };
    };

    // Inheritance is a fall-through, not a lookup. A layout can name a
    // placeholder and still say nothing about where it goes, leaving the
    // position to the master — so each property is taken from the first
    // ancestor that actually states it, rather than from the first ancestor
    // that mentions the placeholder at all.
    const inherit = (ph) => {
      const keys = keysOf(ph);
      const out = { geometry: null, fill: null, line: null, text: null };
      for (const source of [layoutPh, masterPh]) {
        for (const k of keys) {
          const hit = source.get(k);
          if (!hit) continue;
          out.geometry = out.geometry || hit.geometry;
          out.fill = out.fill || hit.fill;
          out.line = out.line || hit.line;
          out.text = out.text || hit.text;
          if (!out.anchor && hit.text?.anchorStated) out.anchor = hit.text.anchor;
        }
      }
      return out.geometry || out.fill || out.line || out.text ? out : null;
    };

    const rel = (id) => {
      const r = rels.get(id);
      if (!r) return null;
      return { part: r.resolved, external: r.mode === 'External', target: r.target };
    };

    const slideXml = this.pkg.text(slidePart);
    const notesPart = [...rels.values()].find((r) => r.type === REL.notes && this.pkg.has(r.resolved))?.resolved || null;
    const notesXml = notesPart ? this.pkg.text(notesPart) : '';

    // `readPart` lets a chart frame read its chart part; only a slide gets
    // it, since a layout or master never carries a chart of its own.
    const readPart = (part) => (part && this.pkg.has(part) ? this.pkg.text(part) : null);
    const scene = readSlideScene(slideXml, { theme, inherit, rel, readPart });
    if (!scene.background) {
      const layoutBg = design?.dropLayoutBackground ? null : layoutPh.get('#background');
      const masterBg = design && 'masterBackground' in design ? design.masterBackground : masterPh.get('#background');
      scene.background = layoutBg || masterBg || null;
    }
    for (const s of scene.shapes) {
      if (s.text || s.inheritedText) s.textStyle = this.#cascade(s, hitsFor(s.placeholder), styles, theme);
    }
    // What the master and the layout draw under every slide on them — a
    // band, a logo, a line — unless the slide hides background graphics.
    const underlay = /<p:sld\b[^>]*\bshowMasterSp="0"/.test(slideXml) ? [] : this.#underlay(layoutPart, masterPart, theme, styles, cache, index + 1);

    let notes = '';
    if (notesPart) notes = sceneText(readSlideScene(notesXml, { theme }));

    return {
      index,
      part: slidePart,
      layout: layoutPart,
      master: masterPart,
      size: this.size,
      background: scene.background,
      shapes: withLinks(withSlideNumber(scene.shapes, index + 1), rel),
      underlay,
      notes,
      hidden: slideHiddenFrom(slideXml),
      // Transitions → the effect this slide comes in with, and how it moves on.
      transition: readTransition(slideXml),
      // Animations → the main sequence, in the Animation Pane's order.
      animations: safeAnimations(slideXml),
      theme: { colors: theme.colors, fonts: theme.fonts, name: theme.name, clrMap: theme.clrMap },
    };
  }

  /**
   * The shapes a layout and its master draw that are not placeholders, in
   * drawing order — the master's first (unless the layout hides them with
   * `showMasterSp="0"`), then the layout's own — each resolved against its
   * own part's relationships, so a logo on the master is found.
   */
  #underlay(layoutPart, masterPart, theme, styles, cache = this.layouts, slideNumber = null) {
    const key = `underlay:${layoutPart}:${masterPart}`;
    let list = cache.get(key);
    if (!list) {
      list = [];
      const layoutXml = layoutPart && this.pkg.has(layoutPart) ? this.pkg.text(layoutPart) : '';
      const hideMaster = /<p:sldLayout\b[^>]*\bshowMasterSp="0"/.test(layoutXml);
      const parts = [hideMaster ? null : masterPart, layoutPart].filter((p) => p && this.pkg.has(p));
      for (const part of parts) {
        for (const s of this.#ownShapes(part, theme, styles)) list.push(s);
      }
      cache.set(key, list);
    }
    return slideNumber == null ? list : withSlideNumber(list, slideNumber);
  }

  /** A master's or a layout's shapes that are not placeholders, read with the part's own relationships. */
  #ownShapes(part, theme, styles) {
    const rels = this.#relMap(part);
    const rel = (id) => {
      const r = rels.get(id);
      return r ? { part: r.resolved, external: r.mode === 'External', target: r.target } : null;
    };
    const scene = readSlideScene(this.pkg.text(part), { theme, rel });
    const placed = new Set(scene.shapes.filter((s) => s.placeholder).map((s) => String(s.id)));
    const out = [];
    for (const s of scene.shapes) {
      if (s.placeholder || (s.groupId != null && placed.has(String(s.groupId)))) continue;
      if (s.text) s.textStyle = this.#cascade(s, { layout: null, master: null }, styles, theme);
      out.push({ ...s, underlay: true, from: part });
    }
    return out;
  }

  /**
   * A master's or a layout's own drawing — what Slide Master view shows and
   * edits: its shapes, its placeholders styled by the cascade (a layout's
   * from the master's), a prompt in each placeholder that holds no words of
   * its own ("Click to edit Master title style", the five levels of the
   * body), its background or the master's, and, under a layout, the
   * master's own shapes. Scene-shaped, so the stage draws it as a slide.
   */
  partScene(part, { prompts = true } = {}) {
    if (!this.#partOf(part) || typeof part !== 'string') throw new RangeError(`no master or layout ${part}`);
    const isLayout = /slideLayouts\//.test(part);
    const masterPart = isLayout ? this.#masterFor(part) : part;
    const key = `part:${part}:${this.designStamp}:${prompts}`;
    const hit = this._scenes.get(key);
    if (hit) return hit;
    const theme = this.#themeFor(masterPart);
    const masterPh = isLayout ? this.#placeholders(masterPart, theme) : new Map();
    const styles = this.#textStyles(masterPart, theme);
    const rels = this.#relMap(part);
    const rel = (id) => {
      const r = rels.get(id);
      return r ? { part: r.resolved, external: r.mode === 'External', target: r.target } : null;
    };
    const xml = this.pkg.text(part);
    const find = (ph) => { for (const k of placeholderKeys(ph)) { const h = masterPh.get(k); if (h) return h; } return null; };
    const inherit = isLayout
      ? (ph) => {
          const h = find(ph);
          return h ? { geometry: h.geometry, fill: h.fill, line: h.line, text: h.text, anchor: h.text?.anchorStated ? h.text.anchor : null } : null;
        }
      : null;
    const scene = readSlideScene(xml, { theme, inherit, rel });
    const shapes = scene.shapes.map((s) => {
      const out = { ...s };
      const master = isLayout && s.placeholder ? find(s.placeholder) : null;
      if (s.text || s.inheritedText || s.placeholder) out.textStyle = this.#cascade(s, { layout: null, master }, styles, theme);
      if (prompts && s.placeholder && !hasWords(s.text)) {
        out.text = promptBody(s.placeholder.type, isLayout, s.text || s.inheritedText);
        out.prompt = true;
      }
      return out;
    });
    const background = scene.background || (isLayout ? masterPh.get('#background') || null : null);
    const hidesMaster = isLayout && /<p:sldLayout\b[^>]*\bshowMasterSp="0"/.test(xml);
    const underlay = isLayout && !hidesMaster ? this.#ownShapes(masterPart, theme, styles) : [];
    const name = unescapeXml(/<p:cSld\b[^>]*\bname="([^"]*)"/.exec(xml)?.[1] || (isLayout ? 'Layout' : 'Slide Master'));
    const result = {
      index: -1,
      kind: isLayout ? 'layout' : 'master',
      name,
      part,
      layout: isLayout ? part : null,
      master: masterPart,
      size: this.size,
      background,
      shapes,
      underlay,
      hidesBackgroundGraphics: hidesMaster,
      notes: '',
      hidden: false,
      transition: null,
      animations: [],
      theme: { colors: theme.colors, fonts: theme.fonts, name: theme.name, clrMap: theme.clrMap },
    };
    this._scenes.set(key, result);
    return result;
  }

  /**
   * Every slide, thumbnail-shaped: enough to draw a sorter without the parts.
   *
   * Read from the slide XML directly, not from the scene. The outline is
   * asked for at open, and building nineteen scenes with their layouts and
   * masters cost 2.6 seconds on a fifteen-megabyte deck before the window
   * could paint anything. A title placeholder's text, a count of shapes and
   * the notes text need none of that: a few milliseconds a slide, whatever
   * the deck weighs.
   */
  outline() {
    // The section each slide is in, by index into `sections()`, or null.
    const sectionAt = new Map();
    for (const s of this.sections()) for (const i of s.slides) sectionAt.set(i, s.index);
    return this.slideParts.map((entry, i) => {
      const xml = this.pkg.text(entry.part);
      const rels = this.#relMap(entry.part);
      const notesPart = [...rels.values()].find((r) => r.type === REL.notes && this.pkg.has(r.resolved))?.resolved || null;
      return {
        index: i,
        part: entry.part,
        title: outlineTitle(xml),
        shapes: (xml.match(/<p:(sp|pic|graphicFrame|grpSp|cxnSp)\b/g) || []).length,
        notes: notesPart ? plainTextOf(this.pkg.text(notesPart)) : '',
        section: sectionAt.has(i) ? sectionAt.get(i) : null,
        hidden: slideHiddenFrom(xml),
        // The strip marks a slide that has a transition, the way PowerPoint's does.
        transition: transitionRange(xml) ? (readTransition(xml)?.type ?? null) : null,
        // …or animations: the strip's star stands for either, as in PowerPoint.
        animated: xml.includes('<p:timing') ? safeAnimations(xml).length > 0 : false,
      };
    });
  }


  /** Bytes of an image part, for the renderer to show. */
  media(partName) {
    return this.pkg.has(partName) ? this.pkg.read(partName) : null;
  }

  // ---- editing -----------------------------------------------------------
  //
  // Every edit is a splice into one part's XML. Finding the shape by its
  // `<p:cNvPr id="…">` is reliable because that id is unique within a slide and
  // PowerPoint never renumbers on open.

  #shapeRange(xml, shapeId) {
    const marker = new RegExp(`<p:cNvPr[^>]*\\bid="${shapeId}"`, 'g');
    // The shape tree's own `<p:cNvPr>` comes before any shape, and a master
    // written by an older build of this suite gave its title the same id —
    // so the first match that follows a shape's opening tag is the shape.
    const firstShape = Math.min(...['<p:sp>', '<p:pic>', '<p:graphicFrame>', '<p:cxnSp>', '<p:grpSp>'].map((o) => { const i = xml.indexOf(o); return i < 0 ? Infinity : i; }));
    let at = -1;
    for (const m of xml.matchAll(marker)) if (m.index > firstShape) { at = m.index; break; }
    if (at < 0) return null;
    // Walk back to the opening tag of the containing shape. `<p:grpSp>` is
    // here too: a group's own `<p:cNvPr>` sits in its `nvGrpSpPr`, first
    // among its children, so this finds the group's own opening tag rather
    // than diving into whichever member happens to match some other search
    // — which is exactly what lets setGeometry, removeShape, shapeClip and
    // pasteShape treat a whole group as one shape, the way PowerPoint does.
    const openers = ['<p:sp>', '<p:pic>', '<p:graphicFrame>', '<p:cxnSp>', '<p:grpSp>'];
    let start = -1;
    let tag = null;
    for (const o of openers) {
      const idx = xml.lastIndexOf(o, at);
      if (idx > start) {
        start = idx;
        tag = o;
      }
    }
    if (start < 0) return null;
    const closeTag = tag.replace('<', '</');
    let end;
    if (tag === '<p:grpSp>') {
      // A group can hold another group; the first `</p:grpSp>` after `at`
      // may belong to a nested one, not to this one, so the close is found
      // by counting depth rather than by the first match.
      const pairRe = /<p:grpSp>|<\/p:grpSp>/g;
      pairRe.lastIndex = start;
      let depth = 0;
      let m;
      end = -1;
      while ((m = pairRe.exec(xml))) {
        if (m[0] === '<p:grpSp>') depth += 1;
        else if (--depth === 0) { end = m.index; break; }
      }
    } else {
      end = xml.indexOf(closeTag, at);
    }
    if (end < 0) return null;
    return { start, end: end + closeTag.length, tag };
  }

  #writeSlide(part, xml) {
    this.pkg.write_(part, Buffer.from(xml, 'utf8'));
    this.dirty = true;
    // A master or a layout is under every slide on it.
    if (Deck.isDesignPart(part)) this.#designChanged();
  }

  /** Replace a shape's text with paragraphs of runs. */
  setText(slideIndex, shapeId, paragraphs) {
    const part = this.#partOf(slideIndex);
    if (!part) throw new RangeError(`no slide at index ${slideIndex}`);
    const xml = this.pkg.text(part);
    const range = this.#shapeRange(xml, shapeId);
    if (!range) throw new Error(`shape ${shapeId} not found on slide ${slideIndex + 1}`);
    const shapeXml = xml.slice(range.start, range.end);
    const bodyStart = shapeXml.indexOf('<p:txBody>');
    const bodyEnd = shapeXml.indexOf('</p:txBody>');
    const newBody = buildTextBody(paragraphs, shapeXml, bodyStart, bodyEnd);
    let nextShape;
    if (bodyStart < 0) {
      // A shape with no text body yet — add one before the closing tag.
      const closeAt = shapeXml.lastIndexOf(range.tag.replace('<', '</'));
      nextShape = shapeXml.slice(0, closeAt) + newBody + shapeXml.slice(closeAt);
    } else {
      nextShape = shapeXml.slice(0, bodyStart) + newBody + shapeXml.slice(bodyEnd + '</p:txBody>'.length);
    }
    this.#writeSlide(part, xml.slice(0, range.start) + nextShape + xml.slice(range.end));
    return true;
  }

  /**
   * Move or resize a shape (or a group, whose own `<p:grpSpPr><a:xfrm>` this
   * finds exactly as it finds a plain shape's `<p:spPr><a:xfrm>` — writing
   * only its own off/ext, never its chOff/chExt, which is exactly how a
   * group resize is meant to scale its members). Values are pixels; EMU
   * conversion happens here. `rot`, `flipH` and `flipV` are each left alone
   * when not given at all — undefined, not falsy — so a plain move never
   * disturbs a rotation or a flip the shape already had; a given value
   * (including 0 or false) is written, which is how one is taken off.
   */
  setGeometry(slideIndex, shapeId, { x, y, w, h, rot, flipH, flipV }) {
    const part = this.#partOf(slideIndex);
    if (!part) throw new RangeError(`no slide at index ${slideIndex}`);
    const xml = this.pkg.text(part);
    const range = this.#shapeRange(xml, shapeId);
    if (!range) throw new Error(`shape ${shapeId} not found`);
    let shapeXml = xml.slice(range.start, range.end);

    const offRe = /<a:off\b[^>]*\/>/;
    const extRe = /<a:ext\b[^>]*\/>/;
    const hasXfrm = /<a:xfrm\b/.test(shapeXml);

    const offXml = `<a:off x="${pxToEmu(x)}" y="${pxToEmu(y)}"/>`;
    const extXml = `<a:ext cx="${Math.max(1, pxToEmu(w))}" cy="${Math.max(1, pxToEmu(h))}"/>`;
    const touchesTransform = rot !== undefined || flipH !== undefined || flipV !== undefined;

    if (hasXfrm) {
      shapeXml = offRe.test(shapeXml) ? shapeXml.replace(offRe, offXml) : shapeXml.replace(/<a:xfrm\b[^>]*>/, (m) => m + offXml);
      shapeXml = extRe.test(shapeXml) ? shapeXml.replace(extRe, extXml) : shapeXml.replace(offXml, offXml + extXml);
      if (touchesTransform) {
        shapeXml = shapeXml.replace(/<a:xfrm\b([^>]*)>/, (m, attrs) => {
          let next = attrs;
          if (rot !== undefined) { next = next.replace(/\s*rot="[^"]*"/, ''); if (rot) next += ` rot="${Math.round(rot * 60000)}"`; }
          if (flipH !== undefined) { next = next.replace(/\s*flipH="[^"]*"/, ''); if (flipH) next += ' flipH="1"'; }
          if (flipV !== undefined) { next = next.replace(/\s*flipV="[^"]*"/, ''); if (flipV) next += ' flipV="1"'; }
          return `<a:xfrm${next}>`;
        });
      }
    } else {
      // The shape inherited its geometry; state it explicitly now that the user
      // has moved it, inserting the xfrm as the first child of spPr.
      const rotAttr = rot ? ` rot="${Math.round(rot * 60000)}"` : '';
      const flipAttr = `${flipH ? ' flipH="1"' : ''}${flipV ? ' flipV="1"' : ''}`;
      shapeXml = shapeXml.replace(
        /<p:spPr\s*\/>|<p:spPr\b[^>]*>/,
        (m) => (m.endsWith('/>')
          ? `<p:spPr><a:xfrm${rotAttr}${flipAttr}>${offXml}${extXml}</a:xfrm></p:spPr>`
          : `${m}<a:xfrm${rotAttr}${flipAttr}>${offXml}${extXml}</a:xfrm>`)
      );
    }
    this.#writeSlide(part, xml.slice(0, range.start) + shapeXml + xml.slice(range.end));
    return true;
  }

  /**
   * A shape as something to paste elsewhere: its XML and the relationships
   * it points at (a picture's media), so a paste on another slide can point
   * that slide at the same parts.
   */
  shapeClip(slideIndex, shapeId) {
    const part = this.#partOf(slideIndex);
    if (!part) throw new RangeError(`no slide at index ${slideIndex}`);
    const xml = this.pkg.text(part);
    const range = this.#shapeRange(xml, shapeId);
    if (!range) return null;
    const shapeXml = xml.slice(range.start, range.end);
    const rels = this.#relMap(part);
    const used = [];
    for (const m of shapeXml.matchAll(/\br:(?:embed|link|id|pict)="([^"]+)"/g)) {
      const rel = rels.get(m[1]);
      if (rel && !used.some((u) => u.id === rel.id)) used.push({ id: rel.id, type: rel.type, target: rel.target, mode: rel.mode });
    }
    return { xml: shapeXml, tag: range.tag, rels: used };
  }

  /**
   * Paste a shape copied with `shapeClip` onto a slide: a fresh id, the
   * relationships it needs added to this slide (or reused where the slide
   * already points at the same part), a placeholder kept only where the
   * slide has none of that kind, and the geometry the caller asks for.
   *
   * @returns {number} the new shape's id
   */
  pasteShape(slideIndex, clip, geometry = null) {
    const part = this.#partOf(slideIndex);
    if (!part) throw new RangeError(`no slide at index ${slideIndex}`);
    if (!clip || !clip.xml) throw new Error('nothing to paste');
    let xml = this.pkg.text(part);
    let shapeXml = String(clip.xml);

    const have = this.#relMap(part);
    for (const rel of clip.rels || []) {
      const same = [...have.values()].find((r) => r.type === rel.type && r.target === rel.target && r.mode === rel.mode);
      const rId = same ? same.id : this.pkg.addRelationshipTo(part, rel.type, rel.target, { external: rel.mode === 'External' });
      if (rId !== rel.id) shapeXml = shapeXml.replace(new RegExp('(\\br:(?:embed|link|id|pict)=")' + rel.id + '"', 'g'), '$1' + rId + '"');
    }
    // A slide that has never had a relationship may not declare the prefix.
    const head = xml.slice(0, Math.max(0, xml.indexOf('<p:cSld')));
    if ((clip.rels || []).length && !/xmlns:r=/.test(head)) {
      xml = xml.replace(/<p:sld\b/, '<p:sld xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
    }

    const id = nextShapeId(xml);
    shapeXml = shapeXml.replace(/<p:cNvPr\b([^>]*?)\bid="\d+"/, (m, before) => `<p:cNvPr${before}id="${id}"`);
    // A placeholder pasted where the slide already has one of that kind
    // becomes a plain shape; otherwise it keeps its place in the layout.
    const ph = /<p:ph\b([^>]*)\/>/.exec(shapeXml);
    if (ph) {
      const type = /\btype="([^"]+)"/.exec(ph[1])?.[1] || 'body';
      const taken = new RegExp('<p:ph\\b[^>]*\\btype="' + type + '"').test(xml) || (type === 'body' && /<p:ph\b(?![^>]*\btype=)[^>]*\/>/.test(xml));
      if (taken) shapeXml = shapeXml.replace(ph[0], '');
    }
    const at = xml.lastIndexOf('</p:spTree>');
    if (at < 0) throw new Error('slide has no shape tree');
    this.#writeSlide(part, xml.slice(0, at) + shapeXml + xml.slice(at));
    if (geometry) this.setGeometry(slideIndex, id, geometry);
    return id;
  }

  /**
   * Home → Reset: the slide's placeholders go back where its layout puts
   * them. A placeholder's own position and size are dropped so the layout's
   * apply again; the words stay.
   *
   * @returns {number} how many placeholders were reset
   */
  resetSlide(slideIndex) {
    const part = this.#partOf(slideIndex);
    if (!part) throw new RangeError(`no slide at index ${slideIndex}`);
    const xml = this.pkg.text(part);
    let count = 0;
    const next = xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, (sp) => {
      if (!/<p:nvPr>\s*<p:ph\b/.test(sp)) return sp;
      const out = sp.replace(/<a:xfrm\b[^>]*>[\s\S]*?<\/a:xfrm>/, '');
      if (out !== sp) count += 1;
      return out;
    });
    if (count) this.#writeSlide(part, next);
    return count;
  }

  /**
   * The text body's own properties — where the words sit in the box
   * (anchor: top, middle, bottom), which way they run (vert: horz, vert,
   * vert270) and how many columns — written on a:bodyPr as PowerPoint keeps
   * them. null, 'horz' and one column take the attribute off.
   */
  setBodyProps(slideIndex, shapeId, { anchor, vert, columns } = {}) {
    const part = this.#partOf(slideIndex);
    if (!part) throw new RangeError(`no slide at index ${slideIndex}`);
    const xml = this.pkg.text(part);
    const range = this.#shapeRange(xml, shapeId);
    if (!range) throw new Error(`shape ${shapeId} not found on slide ${slideIndex + 1}`);
    let shapeXml = xml.slice(range.start, range.end);
    if (!/<a:bodyPr\b/.test(shapeXml)) {
      if (!/<p:txBody>/.test(shapeXml)) throw new Error('the shape has no text body');
      shapeXml = shapeXml.replace('<p:txBody>', '<p:txBody><a:bodyPr/>');
    }
    const set = (attrs, name, value) => {
      const without = attrs.replace(new RegExp(' ?' + name + '="[^"]*"'), '');
      return value == null ? without : `${without} ${name}="${value}"`;
    };
    shapeXml = shapeXml.replace(/<a:bodyPr\b([^>]*?)(\/?)>/, (m, attrs, close) => {
      let next = attrs;
      if (anchor !== undefined) next = set(next, 'anchor', anchor === null ? null : ({ top: 't', middle: 'ctr', bottom: 'b' }[anchor] || 't'));
      if (vert !== undefined) next = set(next, 'vert', vert === null || vert === 'horz' ? null : vert);
      if (columns !== undefined) next = set(next, 'numCol', columns === null || Number(columns) <= 1 ? null : Math.min(16, Math.round(Number(columns))));
      return `<a:bodyPr${next}${close}>`;
    });
    this.#writeSlide(part, xml.slice(0, range.start) + shapeXml + xml.slice(range.end));
    return true;
  }

  /** Remove a shape from a slide. */
  removeShape(slideIndex, shapeId) {
    const part = this.#partOf(slideIndex);
    if (!part) throw new RangeError(`no slide at index ${slideIndex}`);
    const xml = this.pkg.text(part);
    const range = this.#shapeRange(xml, shapeId);
    if (!range) return false;
    this.#writeSlide(part, xml.slice(0, range.start) + xml.slice(range.end));
    // Its animations go with it, as they do in PowerPoint.
    this.#pruneTiming(part);
    return true;
  }

  /** A named top-level shape's (or group's) current box, off the scene — the same absolute, rotation-aware geometry the stage draws. */
  #boxOf(scene, id) {
    const s = scene.shapes.find((x) => String(x.id) === String(id));
    if (!s?.geometry) throw new Error(`shape ${id} not found, or has no geometry`);
    if (s.groupId != null) throw new Error(`shape ${id} is inside a group — align, distribute, rotate or flip the group itself, or ungroup it first`);
    return s;
  }

  /**
   * Home → Arrange → Align: the given edge of every named shape (or group)
   * lined up. One shape aligns to the slide, since PowerPoint has nothing
   * else to line it up against; several align to their own combined bounds
   * by default — "Align Selected Objects" — or to the slide when `to` is
   * `'slide'` — "Align to Slide", each shape moving to that same edge of
   * the whole page. A rotated shape aligns by its visual (rotated) bounds,
   * the way it looks on the stage, not by its unrotated box.
   * @param {number} slideIndex
   * @param {Array<number|string>} ids
   * @param {'left'|'center'|'right'|'top'|'middle'|'bottom'} edge
   * @param {{ to?: 'slide'|'selection' }} [opts]
   * @returns {boolean} true when a shape moved
   */
  alignShapes(slideIndex, ids, edge, { to } = {}) {
    if (!ids || !ids.length) throw new Error('align needs at least one shape');
    const scene = this.slide(slideIndex);
    const items = ids.map((id) => this.#boxOf(scene, id));
    const boxes = items.map((s) => shapeAabb(s.geometry));
    // One shape has nothing else to align to but the slide — PowerPoint
    // offers no other choice for it either — so a single shape always goes
    // to the slide, whatever the ribbon's Align to Slide/Selected Objects
    // toggle happens to be set to; the toggle only means something once
    // there are two or more.
    const mode = items.length > 1 ? (to || 'selection') : 'slide';
    const ref = mode === 'slide' ? { x: 0, y: 0, w: this.size.width, h: this.size.height } : unionBox(boxes);
    let changed = false;
    items.forEach((s, i) => {
      const box = boxes[i];
      let nx = box.x;
      let ny = box.y;
      if (edge === 'left') nx = ref.x;
      else if (edge === 'center') nx = ref.x + ref.w / 2 - box.w / 2;
      else if (edge === 'right') nx = ref.x + ref.w - box.w;
      else if (edge === 'top') ny = ref.y;
      else if (edge === 'middle') ny = ref.y + ref.h / 2 - box.h / 2;
      else if (edge === 'bottom') ny = ref.y + ref.h - box.h;
      else throw new Error(`unknown edge: ${edge}`);
      const dx = nx - box.x;
      const dy = ny - box.y;
      if (Math.abs(dx) < 0.01 && Math.abs(dy) < 0.01) return;
      const g = s.geometry;
      this.setGeometry(slideIndex, s.id, { x: g.x + dx, y: g.y + dy, w: g.w, h: g.h });
      changed = true;
    });
    return changed;
  }

  /**
   * Home → Arrange → Distribute Horizontally/Vertically: three or more
   * shapes spaced with equal gaps between them, along the axis given —
   * between the outermost two shapes' own edges by default ("Align
   * Selected Objects"), or across the whole slide when `to` is `'slide'`
   * ("Align to Slide"), the same toggle `alignShapes` reads.
   * @param {number} slideIndex
   * @param {Array<number|string>} ids
   * @param {'horizontal'|'vertical'} axis
   * @param {{ to?: 'slide'|'selection' }} [opts]
   * @returns {boolean} true when a shape moved
   */
  distributeShapes(slideIndex, ids, axis, { to } = {}) {
    if (!ids || ids.length < 3) throw new Error('distribute needs three or more shapes');
    const scene = this.slide(slideIndex);
    const items = ids.map((id) => {
      const s = this.#boxOf(scene, id);
      return { id: s.id, g: s.geometry, box: shapeAabb(s.geometry) };
    });
    const horizontal = axis === 'horizontal';
    const key = horizontal ? 'x' : 'y';
    const size = horizontal ? 'w' : 'h';
    const sorted = [...items].sort((a, b) => a.box[key] - b.box[key]);
    const mode = to || 'selection';
    const start = mode === 'slide' ? 0 : sorted[0].box[key];
    const end = mode === 'slide'
      ? (horizontal ? this.size.width : this.size.height)
      : sorted[sorted.length - 1].box[key] + sorted[sorted.length - 1].box[size];
    const totalSize = sorted.reduce((n, it) => n + it.box[size], 0);
    const gap = (end - start - totalSize) / (sorted.length - 1);
    let cursor = start;
    let changed = false;
    for (const it of sorted) {
      const delta = cursor - it.box[key];
      if (Math.abs(delta) > 0.01) {
        this.setGeometry(slideIndex, it.id, {
          x: it.g.x + (horizontal ? delta : 0),
          y: it.g.y + (horizontal ? 0 : delta),
          w: it.g.w,
          h: it.g.h,
        });
        changed = true;
      }
      cursor += it.box[size] + gap;
    }
    return changed;
  }

  /**
   * Home → Arrange → Rotate: each named shape turned by `deltaDeg` about
   * its own centre — Right 90°/Left 90° from the ribbon, or Alt+Left/
   * Alt+Right's 15° step. Several shapes each turn about their own centre,
   * not the selection's, exactly as PowerPoint's own does for shapes that
   * are not grouped.
   * @returns {boolean} true when a shape rotated
   */
  rotateShapes(slideIndex, ids, deltaDeg) {
    if (!ids || !ids.length) throw new Error('rotate needs at least one shape');
    const scene = this.slide(slideIndex);
    let changed = false;
    for (const id of ids) {
      const s = this.#boxOf(scene, id);
      const g = s.geometry;
      const rot = (((g.rot || 0) + deltaDeg) % 360 + 360) % 360;
      this.setGeometry(slideIndex, id, { x: g.x, y: g.y, w: g.w, h: g.h, rot });
      changed = true;
    }
    return changed;
  }

  /**
   * Home → Arrange → Flip Horizontal/Flip Vertical: each named shape
   * mirrored about its own centre. Rotation is left exactly as it was —
   * writing `flipH`/`flipV` on `a:xfrm` beside whatever `rot` is already
   * there, never touching it.
   * @param {'horizontal'|'vertical'} axis
   * @returns {boolean} true when a shape flipped
   */
  flipShapes(slideIndex, ids, axis) {
    if (!ids || !ids.length) throw new Error('flip needs at least one shape');
    const scene = this.slide(slideIndex);
    let changed = false;
    for (const id of ids) {
      const s = this.#boxOf(scene, id);
      const g = s.geometry;
      const flip = axis === 'horizontal' ? { flipH: !g.flipH } : { flipV: !g.flipV };
      this.setGeometry(slideIndex, id, { x: g.x, y: g.y, w: g.w, h: g.h, ...flip });
      changed = true;
    }
    return changed;
  }

  /**
   * Home → Arrange → Group: the named top-level shapes (or groups) gathered
   * into one `p:grpSp`, in their current drawing order, at the position the
   * topmost of them held — the way PowerPoint's own Group does. The child
   * coordinate window (`chOff`/`chExt`) is set equal to the new group's own
   * bounding box, so grouping never moves or resizes a member by itself;
   * only a later resize of the group (`setGeometry` on its own id) scales
   * them, because `chOff`/`chExt` stay while `off`/`ext` change.
   * @param {number} slideIndex
   * @param {Array<number|string>} ids
   * @returns {number} the new group's id
   */
  groupShapes(slideIndex, ids) {
    const part = this.#partOf(slideIndex);
    if (!part) throw new RangeError(`no slide at index ${slideIndex}`);
    const wanted = [...new Set((ids || []).map(String))];
    if (wanted.length < 2) throw new Error('grouping needs two or more shapes');
    const scene = this.slide(slideIndex);
    const boxes = wanted.map((id) => shapeAabb(this.#boxOf(scene, id).geometry));
    const bbox = unionBox(boxes);

    const xml = this.pkg.text(part);
    const entries = topLevelShapes(xml);
    const memberSet = new Set(wanted);
    const members = entries.filter((e) => memberSet.has(String(e.id)));
    if (members.length !== wanted.length) throw new Error('one or more shapes were not found at the top level of this slide');

    const groupId = nextShapeId(xml);
    const offX = pxToEmu(bbox.x);
    const offY = pxToEmu(bbox.y);
    const extCx = Math.max(1, pxToEmu(bbox.w));
    const extCy = Math.max(1, pxToEmu(bbox.h));
    const grpXml =
      `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${groupId}" name="Group ${groupId}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      `<p:grpSpPr><a:xfrm><a:off x="${offX}" y="${offY}"/><a:ext cx="${extCx}" cy="${extCy}"/>` +
      `<a:chOff x="${offX}" y="${offY}"/><a:chExt cx="${extCx}" cy="${extCy}"/></a:xfrm></p:grpSpPr>` +
      members.map((m) => xml.slice(m.start, m.end)).join('') +
      `</p:grpSp>`;

    // Members keep their relative order (the group draws them exactly as
    // they drew before), and the group itself takes the topmost one's spot
    // — the last member in document order, since the spTree draws bottom
    // first. Everything else stays exactly where it was.
    const head = xml.slice(0, entries[0].start);
    const tail = xml.slice(entries[entries.length - 1].end);
    const topmost = members[members.length - 1];
    let body = '';
    for (const e of entries) {
      if (memberSet.has(String(e.id))) {
        if (e === topmost) body += grpXml;
        continue;
      }
      body += xml.slice(e.start, e.end);
    }
    this.#writeSlide(part, head + body + tail);
    return groupId;
  }

  /**
   * Home → Arrange → Ungroup: a group's members put back on the slide at
   * their true coordinates, mapped out of the group's child coordinate
   * space through its own off/ext/chOff/chExt/rot/flip — the inverse of
   * what reading a group through the scene already does, so a group that
   * was moved, resized, rotated or flipped before being ungrouped leaves
   * its members exactly where they visibly were, not where they were drawn
   * before the group was touched.
   * @returns {Array<number|string>} the members' own ids, now top-level
   */
  ungroupShape(slideIndex, groupId) {
    const part = this.#partOf(slideIndex);
    if (!part) throw new RangeError(`no slide at index ${slideIndex}`);
    const xml = this.pkg.text(part);
    const range = this.#shapeRange(xml, groupId);
    if (!range) throw new Error(`shape ${groupId} not found`);
    if (range.tag !== '<p:grpSp>') throw new Error(`shape ${groupId} is not a group`);
    const grpXml = xml.slice(range.start, range.end);

    const grpSpPr = /<p:grpSpPr\b[^>]*>[\s\S]*?<\/p:grpSpPr>|<p:grpSpPr\b[^>]*\/>/.exec(grpXml)?.[0] || '';
    const gxfrm = /<a:xfrm\b[^>]*>[\s\S]*?<\/a:xfrm>|<a:xfrm\b[^>]*\/>/.exec(grpSpPr)?.[0] || '';
    const num = (scope, name) => Number(new RegExp(`\\b${name}="(-?\\d+)"`).exec(scope || '')?.[1] || 0);
    const offXml = /<a:off\b[^>]*\/>/.exec(gxfrm)?.[0];
    const extXml = /<a:ext\b[^>]*\/>/.exec(gxfrm)?.[0];
    const chOffXml = /<a:chOff\b[^>]*\/>/.exec(gxfrm)?.[0];
    const chExtXml = /<a:chExt\b[^>]*\/>/.exec(gxfrm)?.[0];
    const containerPx = {
      offX: emuToPx(num(offXml, 'x')), offY: emuToPx(num(offXml, 'y')),
      extX: emuToPx(num(extXml, 'cx')) || 1, extY: emuToPx(num(extXml, 'cy')) || 1,
      chOffX: emuToPx(num(chOffXml, 'x')), chOffY: emuToPx(num(chOffXml, 'y')),
      chExtX: chExtXml ? emuToPx(num(chExtXml, 'cx')) : emuToPx(num(extXml, 'cx')) || 1,
      chExtY: chExtXml ? emuToPx(num(chExtXml, 'cy')) : emuToPx(num(extXml, 'cy')) || 1,
      rot: Number(/\brot="(-?\d+)"/.exec(gxfrm)?.[1] || 0) / 60000,
      flipH: /\bflipH="1"/.test(gxfrm),
      flipV: /\bflipV="1"/.test(gxfrm),
    };

    const grpSpPrM = /<p:grpSpPr\b[^>]*\/>|<p:grpSpPr\b[^>]*>[\s\S]*?<\/p:grpSpPr>/.exec(grpXml);
    const childStart = grpSpPrM ? grpSpPrM.index + grpSpPrM[0].length : 0;
    const childEnd = grpXml.lastIndexOf('</p:grpSp>');
    const children = topLevelShapes(grpXml, { start: childStart, end: childEnd });
    if (!children.length) throw new Error('this group has no members');
    const memberXml = children.map((c) => xml.slice(range.start + c.start, range.start + c.end)).join('');
    const head = xml.slice(0, range.start);
    const tail = xml.slice(range.end);
    this.#writeSlide(part, head + memberXml + tail);

    // Newly top-level, the members still carry their group-relative box;
    // the scene now reads it straight (nothing wraps them any more), which
    // is exactly the `local` box composeGroupChild needs to place absolutely.
    const memberIds = children.map((c) => c.id);
    const scene = this.slide(slideIndex);
    for (const id of memberIds) {
      const s = scene.shapes.find((x) => String(x.id) === String(id));
      if (!s?.geometry) continue;
      const g = s.geometry;
      const local = { offX: g.x, offY: g.y, extX: g.w, extY: g.h, rot: g.rot || 0, flipH: Boolean(g.flipH), flipV: Boolean(g.flipV) };
      const abs = composeGroupChild(containerPx, local);
      this.setGeometry(slideIndex, id, { x: abs.offX, y: abs.offY, w: abs.extX, h: abs.extY, rot: abs.rot, flipH: abs.flipH, flipV: abs.flipV });
    }
    // An animation on the group itself has nothing left to animate.
    this.#pruneTiming(this.slideParts[slideIndex].part);
    return memberIds;
  }

  /**
   * Move a shape in the drawing order — bring forward, send backward, to the
   * front, to the back. The spTree draws its children in order, so the order
   * IS the layering: the shape's XML moves among its siblings, nothing else.
   */
  reorderShape(slideIndex, shapeId, to) {
    const part = this.#partOf(slideIndex);
    if (!part) throw new RangeError(`no slide at index ${slideIndex}`);
    const xml = this.pkg.text(part);
    const shapes = topLevelShapes(xml);
    const i = shapes.findIndex((s) => String(s.id) === String(shapeId));
    if (i < 0) throw new Error(`shape ${shapeId} not found`);
    const j = to === 'front' ? shapes.length - 1
      : to === 'back' ? 0
        : to === 'forward' ? Math.min(shapes.length - 1, i + 1)
          : to === 'backward' ? Math.max(0, i - 1)
            : i;
    if (j === i) return false;
    const order = shapes.slice();
    const [moved] = order.splice(i, 1);
    order.splice(j, 0, moved);
    const head = xml.slice(0, shapes[0].start);
    const tail = xml.slice(shapes[shapes.length - 1].end);
    this.#writeSlide(part, head + order.map((s) => xml.slice(s.start, s.end)).join('') + tail);
    return true;
  }

  /**
   * A shape's fill, outline and effects — the Format pane. Written into the
   * shape's own spPr, where it beats the style reference the shape may
   * carry, in the order the schema wants: geometry, fill, line, effects. A
   * picture takes an outline as a frame; a table or chart frame has no
   * spPr, and a group has none of its own to speak of, so both say so.
   *
   * `fill` — null leaves it; `'none'` (or `{ type: 'none' }`) clears it;
   * `{ gradient: { stops?, preset?, color?, angle? } }` writes a linear
   * gradient — `stops` is `{ pos: 0–1, ...colour }[]`, `colour` anything
   * below reads; without `stops`, `preset` ('light', the default, or
   * 'dark') builds one from `color`, PowerPoint's own gallery swatches;
   * `{ picture: { data, contentType, name?, tile? } }` embeds a picture and
   * fills the shape with it, clipped to its outline, the bytes normalised
   * the way `documents.js` already normalises an inserted picture's; a
   * picture fill with no `data` (just `{ picture: { tile } }`) keeps the
   * blip the shape already had and only changes stretch/tile — the Format
   * pane's own Tile toggle, which does not want to hold a picture's bytes
   * just to flip a switch;
   * anything else is a solid colour — a hex string, or `{ scheme } | { color }`
   * with an optional `lumMod`/`lumOff`/`alpha` (0–1, transparency).
   *
   * `line` — null leaves it; `'none'` clears it; `{ color, width (points),
   * dash }` sets it, `color` the same colour shape `fill` reads.
   *
   * `effects` — null leaves the whole list; `'none'` clears it; otherwise
   * `{ shadow?, glow?, softEdge?, reflection? }` is the *complete* list from
   * here on — a key left out is a key turned off, since the whole
   * `a:effectLst` is rebuilt each time in the order the schema wants (glow,
   * the shadow, the reflection, the soft edge) rather than merged with
   * whatever was already there. `shadow` is `{ dist, dir, blur, color,
   * alpha }` in points, degrees and 0–1; `glow` is `{ radius, color, alpha }`,
   * radius in points (5/8/11/18 pt is PowerPoint's own gallery); `softEdge`
   * is `{ radius }`, points; `reflection` is `'tight' | 'half' | 'full'`.
   */
  setShapeStyle(slideIndex, shapeId, { fill = null, line = null, effects = null } = {}) {
    const part = this.#partOf(slideIndex);
    if (!part) throw new RangeError(`no slide at index ${slideIndex}`);
    // The picture is embedded first, when there is one: its own write (a
    // media part and a relationship) touches the slide's rels, not its
    // XML, so it is safe to do before the XML below is even read. A
    // picture fill that replaces an earlier one leaves the earlier media
    // part and relationship in the package, unreferenced, the way an
    // undone edit leaves one — not deleted, just no longer pointed at.
    const pictureRel = fill?.picture?.data ? this.#embedImage(part, fill.picture) : null;
    let xml = this.pkg.text(part);
    if (pictureRel) {
      const head = xml.slice(0, Math.max(0, xml.indexOf('<p:cSld')));
      if (!/xmlns:r=/.test(head)) {
        xml = xml.replace(/<p:sld\b/, '<p:sld xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
      }
    }
    const range = this.#shapeRange(xml, shapeId);
    if (!range) throw new Error(`shape ${shapeId} not found`);
    if (range.tag === '<p:graphicFrame>') throw new Error('A table or chart frame has no fill or outline of its own.');
    if (range.tag === '<p:grpSp>') throw new Error('A group has no fill or outline of its own — format a shape inside it.');
    let shapeXml = xml.slice(range.start, range.end);
    const spPrRe = /<p:spPr\b[^>]*\/>|<p:spPr\b[^>]*>[\s\S]*?<\/p:spPr>/;
    const m = spPrRe.exec(shapeXml);
    if (!m) throw new Error(`shape ${shapeId} has no properties to write`);
    const spPr = m[0].endsWith('/>') ? m[0].replace(/\/>$/, '></p:spPr>') : m[0];
    const open = /^<p:spPr\b[^>]*>/.exec(spPr)[0];
    let inner = spPr.slice(open.length, -'</p:spPr>'.length);
    // The line first: it holds a fill of its own, which must not be mistaken for the shape's.
    const lnRe = /<a:ln\b[^>]*\/>|<a:ln\b[^>]*>[\s\S]*?<\/a:ln>/;
    const hadLine = lnRe.exec(inner)?.[0] ?? '';
    inner = inner.replace(lnRe, '');
    // The effects after both. Order in the schema: blur?, fillOverlay?,
    // glow?, innerShdw?, outerShdw?, prstShdw?, reflection?, softEdge? — of
    // which this writes glow, the shadow, the reflection and the soft edge.
    const effectRe = /<a:effectLst\b[^>]*\/>|<a:effectLst\b[^>]*>[\s\S]*?<\/a:effectLst>/;
    const hadEffects = effectRe.exec(inner)?.[0] ?? '';
    inner = inner.replace(effectRe, '');
    const effectXml = (() => {
      if (effects === null) return hadEffects;
      if (effects === 'none') return '<a:effectLst/>';
      const kids = [];
      if (effects.glow) kids.push(glowXml(effects.glow));
      if (effects.shadow) kids.push(outerShadowXml(effects.shadow));
      if (effects.reflection) kids.push(reflectionXml(effects.reflection));
      if (effects.softEdge) kids.push(softEdgeXml(effects.softEdge));
      return kids.length ? `<a:effectLst>${kids.join('')}</a:effectLst>` : '<a:effectLst/>';
    })();
    const fillRe = /<a:(noFill|solidFill|gradFill|blipFill|pattFill|grpFill)\b[^>]*\/>|<a:(noFill|solidFill|gradFill|blipFill|pattFill|grpFill)\b[^>]*>[\s\S]*?<\/a:\2>/;
    const hadFill = fillRe.exec(inner)?.[0] ?? '';
    inner = inner.replace(fillRe, '');
    // A picture fill with no fresh bytes (`fill.picture` names no `data`)
    // keeps the blip it already had — the Format pane's own Tile toggle
    // does exactly this, changing only stretch/tile without asking the
    // window to hold onto the picture's bytes after they are embedded.
    const pictureRId = pictureRel?.rId ?? (fill?.picture ? /<a:blip\b[^>]*\br:embed="([^"]+)"/.exec(hadFill)?.[1] : null);
    if (fill?.picture && !pictureRId) throw new Error('this shape has no picture fill to change — give it one with `data`');
    const fillXml = fill === null
      ? hadFill
      : fill === 'none' || fill?.type === 'none'
        ? '<a:noFill/>'
        : fill.gradient
          ? gradientFillXml(fill.gradient)
          : fill.picture
            ? `<a:blipFill><a:blip r:embed="${pictureRId}"/>${fill.picture.tile ? '<a:tile/>' : '<a:stretch><a:fillRect/></a:stretch>'}</a:blipFill>`
            : `<a:solidFill>${colourXml(fill)}</a:solidFill>`;
    const lineXml = line === null
      ? hadLine
      : line === 'none' || line?.type === 'none'
        ? '<a:ln><a:noFill/></a:ln>'
        : `<a:ln w="${Math.round((line.width ?? 1) * 12700)}"><a:solidFill>${colourXml(line.color ?? { scheme: 'accent1' })}</a:solidFill>${line.dash && line.dash !== 'solid' ? `<a:prstDash val="${escapeXml(String(line.dash))}"/>` : ''}</a:ln>`;
    // After the geometry, else after the transform, else first.
    const geom = /<a:(prstGeom|custGeom)\b[^>]*\/>|<a:(prstGeom|custGeom)\b[^>]*>[\s\S]*?<\/a:\2>/.exec(inner);
    const xfrm = geom ? null : /<a:xfrm\b[^>]*\/>|<a:xfrm\b[^>]*>[\s\S]*?<\/a:xfrm>/.exec(inner);
    const at = geom ? geom.index + geom[0].length : xfrm ? xfrm.index + xfrm[0].length : 0;
    inner = inner.slice(0, at) + fillXml + lineXml + effectXml + inner.slice(at);
    shapeXml = shapeXml.slice(0, m.index) + open + inner + '</p:spPr>' + shapeXml.slice(m.index + m[0].length);
    this.#writeSlide(part, xml.slice(0, range.start) + shapeXml + xml.slice(range.end));
    return true;
  }

  /** Hide a shape, or show it again — the selection pane's eye. It stays in the file, undrawn. */
  setShapeHidden(slideIndex, shapeId, hidden) {
    return this.#editShapeProps(slideIndex, shapeId, (open) => {
      const cleaned = open.replace(/\s+hidden="[^"]*"/, '');
      return hidden ? cleaned.replace(/\/?>$/, (m) => ` hidden="1"${m}`) : cleaned;
    });
  }

  /** Give a shape the name the selection pane shows — and PowerPoint keeps. */
  renameShape(slideIndex, shapeId, name) {
    const safe = escapeXml(String(name ?? '').trim() || `Shape ${shapeId}`);
    return this.#editShapeProps(slideIndex, shapeId, (open) => (
      /\sname="/.test(open) ? open.replace(/\sname="[^"]*"/, ` name="${safe}"`) : open.replace(/\/?>$/, (m) => ` name="${safe}"${m}`)
    ));
  }

  /** Rewrite the opening tag of a shape's own p:cNvPr. */
  #editShapeProps(slideIndex, shapeId, edit) {
    const part = this.#partOf(slideIndex);
    if (!part) throw new RangeError(`no slide at index ${slideIndex}`);
    const xml = this.pkg.text(part);
    const re = new RegExp(`<p:cNvPr\\b[^>]*\\bid="${String(shapeId).replace(/[^\w-]/g, '')}"[^>]*>`);
    const m = re.exec(xml);
    if (!m) throw new Error(`shape ${shapeId} not found`);
    const next = edit(m[0]);
    if (next === m[0]) return false;
    this.#writeSlide(part, xml.slice(0, m.index) + next + xml.slice(m.index + m[0].length));
    return true;
  }

  /** The layout part a slide uses. */
  layoutOf(slideIndex) {
    const part = this.#partOf(slideIndex);
    return part ? this.#layoutFor(part) : null;
  }

  /**
   * The deck's layouts, in part order: the name PowerPoint shows, the type,
   * and each placeholder as a box — a placeholder that states no geometry
   * of its own takes the master's — which is enough to draw the gallery.
   */
  layoutList() {
    const parts = (this.pkg.partNames() || [])
      .filter((p) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(p))
      .sort((a, b) => Number(/(\d+)\.xml$/.exec(a)[1]) - Number(/(\d+)\.xml$/.exec(b)[1]));
    return parts.map((part) => {
      const xml = this.pkg.text(part);
      const name = unescapeXml(/<p:cSld\b[^>]*\bname="([^"]*)"/.exec(xml)?.[1] || part.split('/').pop());
      const type = /<p:sldLayout\b[^>]*\btype="([^"]*)"/.exec(xml)?.[1] || null;
      const masterPart = this.#masterFor(part);
      const theme = this.#themeFor(masterPart);
      const masterPh = this.#placeholders(masterPart, theme);
      const placeholders = [];
      for (const s of readSlideScene(xml, { theme }).shapes) {
        if (!s.placeholder) continue;
        const kind = s.placeholder.type || 'body';
        const inherited = (s.placeholder.idx != null ? masterPh.get(`idx:${s.placeholder.idx}`) : null) || masterPh.get(`type:${kind}`);
        const geometry = s.geometry || inherited?.geometry || null;
        if (geometry) placeholders.push({ type: kind, geometry });
      }
      return { part, name, type, placeholders };
    });
  }

  /**
   * Put a slide on another of the deck's layouts. Its placeholders then
   * inherit that layout's geometry and text style wherever they state none
   * of their own — which is what PowerPoint's Layout gallery does.
   */
  applyLayout(slideIndex, layoutPart) {
    const part = this.#partOf(slideIndex);
    if (!part) throw new RangeError(`no slide at index ${slideIndex}`);
    if (!/^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(layoutPart) || !this.pkg.has(layoutPart)) throw new Error(`no layout ${layoutPart} in this deck`);
    const relsPath = part.replace(/([^/]+)$/, '_rels/$1.rels');
    const rels = this.pkg.text(relsPath);
    const target = `../slideLayouts/${layoutPart.split('/').pop()}`;
    const next = rels.replace(/<Relationship\b[^>]*>/g, (el) => (/\bType="[^"]*\/slideLayout"/.test(el) ? el.replace(/\bTarget="[^"]*"/, `Target="${target}"`) : el));
    if (next === rels) throw new Error('this slide has no layout to change');
    this.pkg.write_(relsPath, Buffer.from(next, 'utf8'));
    // The scene cache is keyed on the slide's XML, which this leaves as it
    // was; the scene it holds was built on the old layout.
    this._scenes.delete(part);
    this.dirty = true;
    return true;
  }

  // ---- Slide Master view -----------------------------------------------------
  //
  // A master and its layouts are parts like a slide, and the same verbs edit
  // them (every verb takes a master's or a layout's part name where it takes
  // a slide index). What only a master has — its layouts, the text styles
  // its title and body placeholders hand every slide, the placeholders a
  // layout offers — is here.

  /** How many slides sit on a layout. */
  layoutUsage(layoutPart) {
    return this.slideParts.filter((s) => this.#layoutFor(s.part) === layoutPart).length;
  }

  /** The deck's masters, each with its name and its layouts in its own order — the Slide Master view's strip. */
  masterList() {
    return this.masterParts().map((master) => {
      const xml = this.pkg.text(master);
      const rels = this.#relMap(master);
      const layouts = [...xml.matchAll(/<p:sldLayoutId\b[^>]*\br:id="([^"]+)"/g)]
        .map((m) => rels.get(m[1])?.resolved)
        .filter((p) => p && this.pkg.has(p))
        .map((part) => {
          const lx = this.pkg.text(part);
          return {
            part,
            name: unescapeXml(/<p:cSld\b[^>]*\bname="([^"]*)"/.exec(lx)?.[1] || 'Layout'),
            type: /<p:sldLayout\b[^>]*\btype="([^"]*)"/.exec(lx)?.[1] || null,
            used: this.layoutUsage(part),
            hidesBackgroundGraphics: /<p:sldLayout\b[^>]*\bshowMasterSp="0"/.test(lx),
            hasTitle: /<p:ph\b[^>]*\btype="(title|ctrTitle)"/.test(lx),
            hasFooters: /<p:ph\b[^>]*\btype="(dt|ftr|sldNum)"/.test(lx),
          };
        });
      return {
        part: master,
        name: unescapeXml(/<p:cSld\b[^>]*\bname="([^"]*)"/.exec(xml)?.[1] || `${this.#themeFor(master).name || 'Office'} Slide Master`),
        preserve: /<p:sldMaster\b[^>]*\bpreserve="1"/.test(xml),
        hasTitle: /<p:ph\b[^>]*\btype="title"/.test(xml),
        hasFooters: /<p:ph\b[^>]*\btype="(dt|ftr|sldNum)"/.test(xml),
        layouts,
      };
    });
  }

  /** Rename a master or a layout — the name PowerPoint shows in the gallery. */
  renamePart(part, name) {
    const p = this.#partOf(part);
    if (!p || typeof part !== 'string') throw new RangeError(`no master or layout ${part}`);
    const clean = String(name ?? '').trim();
    if (!clean) throw new Error('a layout needs a name');
    const xml = this.pkg.text(p);
    const next = /<p:cSld\b[^>]*\bname="[^"]*"/.test(xml)
      ? xml.replace(/(<p:cSld\b[^>]*\bname=")[^"]*"/, (m, a) => `${a}${escapeXml(clean)}"`)
      : xml.replace(/<p:cSld\b/, `<p:cSld name="${escapeXml(clean)}"`);
    if (next === xml) return false;
    this.#writeSlide(p, next);
    return true;
  }

  /** Slide Master → Preserve: a master PowerPoint keeps even when no slide uses it. */
  setMasterPreserve(master, on) {
    if (!/slideMasters\//.test(String(master)) || !this.pkg.has(master)) throw new RangeError(`no master ${master}`);
    const xml = this.pkg.text(master);
    const bare = xml.replace(/(<p:sldMaster\b[^>]*?)\s+preserve="[^"]*"/, '$1');
    const next = on ? bare.replace(/<p:sldMaster\b/, '<p:sldMaster preserve="1"') : bare;
    if (next === xml) return false;
    this.#writeSlide(master, next);
    return true;
  }

  /** Background → Hide Background Graphics: a layout (or a slide) drawn without its master's shapes. */
  setHideBackgroundGraphics(target, hide) {
    const part = this.#partOf(target);
    if (!part || /slideMasters\//.test(part)) throw new RangeError('background graphics are hidden on a layout or a slide');
    const tag = /slideLayouts\//.test(part) ? 'p:sldLayout' : 'p:sld';
    const xml = this.pkg.text(part);
    const open = new RegExp(`<${tag}\\b[^>]*>`).exec(xml);
    if (!open) return false;
    const bare = open[0].replace(/\s+showMasterSp="[^"]*"/, '');
    const nextOpen = hide ? bare.replace(new RegExp(`^<${tag}`), `<${tag} showMasterSp="0"`) : bare;
    if (nextOpen === open[0]) return false;
    this.#writeSlide(part, xml.slice(0, open.index) + nextOpen + xml.slice(open.index + open[0].length));
    if (tag === 'p:sld') this._scenes.delete(part);
    return true;
  }

  /**
   * Slide Master → Insert Layout: a new layout at the end of a master's
   * list, PowerPoint's own "Custom Layout" — a title and the footers, each
   * taking its place and look from the master.
   * @returns {string} the new layout's part
   */
  insertLayout(master = null, { name = 'Custom Layout' } = {}) {
    const m = master || this.masterParts()[0];
    if (!m || !this.pkg.has(m)) throw new RangeError('no master to add a layout to');
    const n = this.pkg.nextPartNumber('ppt/slideLayouts/', 'slideLayout');
    const part = `ppt/slideLayouts/slideLayout${n}.xml`;
    const masterXml = this.pkg.text(m);
    const ph = (id, type, label, idx = null) =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${label} ${id - 1}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="${type}"${type === 'title' ? '' : ' sz="quarter"'}${idx != null ? ` idx="${idx}"` : ''}/></p:nvPr></p:nvSpPr><p:spPr/>` +
      `<p:txBody><a:bodyPr/><a:lstStyle/><a:p>${type === 'title' ? '<a:r><a:rPr lang="en-US"/><a:t>Click to edit Master title style</a:t></a:r>' : ''}<a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
    const shapes = [ph(2, 'title', 'Title')];
    let id = 3;
    for (const [type, label, idx] of [['dt', 'Date Placeholder', 10], ['ftr', 'Footer Placeholder', 11], ['sldNum', 'Slide Number Placeholder', 12]]) {
      if (new RegExp(`<p:ph\\b[^>]*\\btype="${type}"`).test(masterXml)) shapes.push(ph(id++, type, label, idx));
    }
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" preserve="1" userDrawn="1">` +
      `<p:cSld name="${escapeXml(name)}"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
      shapes.join('') + `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
    this.pkg.addPart(part, Buffer.from(xml, 'utf8'), CT.layout);
    // Its own relationships part, typed by the package's rels default as PowerPoint writes it.
    this.pkg.addPart(`ppt/slideLayouts/_rels/slideLayout${n}.xml.rels`, Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL.master}" Target="../slideMasters/${m.split('/').pop()}"/></Relationships>`, 'utf8'));
    const rId = this.pkg.addRelationshipTo(m, REL.layout, `../slideLayouts/slideLayout${n}.xml`);
    // Layout ids share one number space with the masters' — above 2^31, unique across the deck.
    const pres = this.pkg.text('ppt/presentation.xml');
    const ids = [...pres.matchAll(/<p:sldMasterId\b[^>]*\bid="(\d+)"/g)].map((x) => Number(x[1]));
    for (const mp of this.masterParts()) for (const x of this.pkg.text(mp).matchAll(/<p:sldLayoutId\b[^>]*\bid="(\d+)"/g)) ids.push(Number(x[1]));
    const next = Math.max(2147483648, ...ids) + 1;
    const entry = `<p:sldLayoutId id="${next}" r:id="${rId}"/>`;
    const mx = this.pkg.text(m);
    const withEntry = /<\/p:sldLayoutIdLst>/.test(mx)
      ? mx.replace('</p:sldLayoutIdLst>', `${entry}</p:sldLayoutIdLst>`)
      : mx.replace(/(<p:clrMap\b[^>]*\/>)/, `$1<p:sldLayoutIdLst>${entry}</p:sldLayoutIdLst>`);
    this.#writeSlide(m, withEntry);
    return part;
  }

  /** Slide Master → Delete: a layout no slide uses, taken out of its master's list and the package. */
  removeLayout(part) {
    if (!/slideLayouts\//.test(String(part)) || !this.pkg.has(part)) throw new RangeError(`no layout ${part}`);
    const used = this.layoutUsage(part);
    if (used) throw new Error(`${used === 1 ? 'A slide uses' : `${used} slides use`} this layout; put ${used === 1 ? 'it' : 'them'} on another layout first.`);
    const master = this.#masterFor(part);
    if (master) {
      const rel = [...this.#relMap(master).values()].find((r) => r.type === REL.layout && r.resolved === part);
      if (rel) {
        const mx = this.pkg.text(master).replace(new RegExp(`<p:sldLayoutId\\b[^>]*\\br:id="${rel.id}"[^>]*/>`), '');
        this.pkg.write_(master, Buffer.from(mx, 'utf8'));
        const relsPath = master.replace(/([^/]+)$/, '_rels/$1.rels');
        const rx = this.pkg.text(relsPath).replace(new RegExp(`<Relationship\\b[^>]*\\bId="${rel.id}"[^>]*/>`), '');
        this.pkg.write_(relsPath, Buffer.from(rx, 'utf8'));
      }
    }
    this.pkg.removePart(part);
    this.#designChanged();
    return true;
  }

  /**
   * Slide Master → Title and Footers: a layout (or the master) with its
   * title placeholder, or without; with the date, footer and number
   * placeholders, or without. An added placeholder states no geometry of
   * its own, so it takes the master's place and look; on the master itself
   * it goes where PowerPoint's own masters put it.
   */
  setMasterPlaceholders(target, { title, footers } = {}) {
    const part = this.#partOf(target);
    if (!part || typeof target !== 'string') throw new RangeError(`no master or layout ${target}`);
    const isMaster = /slideMasters\//.test(part);
    let xml = this.pkg.text(part);
    const blocks = (re) => {
      const out = [];
      const sp = /<p:sp>[\s\S]*?<\/p:sp>/g;
      let m;
      while ((m = sp.exec(xml))) if (re.test(m[0])) out.push({ start: m.index, end: m.index + m[0].length });
      return out;
    };
    const drop = (re) => { for (const b of blocks(re).reverse()) xml = xml.slice(0, b.start) + xml.slice(b.end); };
    const { cx, cy } = this.size;
    const box = (x, y, w, h) => `<p:spPr><a:xfrm><a:off x="${Math.round(cx * x)}" y="${Math.round(cy * y)}"/><a:ext cx="${Math.round(cx * w)}" cy="${Math.round(cy * h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>`;
    const add = (type, label, idx, geometry, words = '', algn = null) => {
      const id = nextShapeId(xml);
      const sp = `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${label} ${id - 1}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="${type}"${type === 'title' ? '' : ' sz="quarter"'}${idx != null ? ` idx="${idx}"` : ''}/></p:nvPr></p:nvSpPr>` +
        `${isMaster || !masterPlaces(type) ? geometry : '<p:spPr/>'}<p:txBody><a:bodyPr/><a:lstStyle/><a:p>${algn ? `<a:pPr algn="${algn}"/>` : ''}${words ? `<a:r><a:rPr lang="en-US"/><a:t>${words}</a:t></a:r>` : ''}<a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
      const at = type === 'title' ? xml.indexOf('</p:grpSpPr>') + '</p:grpSpPr>'.length : xml.lastIndexOf('</p:spTree>');
      xml = xml.slice(0, at) + sp + xml.slice(at);
    };
    // A layout's new placeholder takes its place from the master's of that
    // kind; where the master has none, it is placed where PowerPoint's own
    // masters put it, so it is not a box with nowhere to be.
    const masterXml = isMaster ? '' : this.pkg.text(this.#masterFor(part) || part);
    const masterPlaces = (type) => new RegExp(`<p:ph\\b[^>]*\\btype="${type}"`).test(masterXml) && /<a:off\b/.test(masterXml);
    const hasTitle = () => /<p:ph\b[^>]*\btype="(title|ctrTitle)"/.test(xml);
    if (title === false) drop(/<p:ph\b[^>]*\btype="(title|ctrTitle)"/);
    if (title === true && !hasTitle()) add('title', 'Title', null, box(0.0688, 0.0533, 0.8625, 0.1933), 'Click to edit Master title style');
    if (footers === false) drop(/<p:ph\b[^>]*\btype="(dt|ftr|sldNum)"/);
    if (footers === true) {
      const band = [['dt', 'Date Placeholder', 10, box(0.0688, 0.9267, 0.225, 0.0533), 'l'], ['ftr', 'Footer Placeholder', 11, box(0.3313, 0.9267, 0.3375, 0.0533), 'ctr'], ['sldNum', 'Slide Number Placeholder', 12, box(0.7063, 0.9267, 0.225, 0.0533), 'r']];
      for (const [type, label, idx, geometry, algn] of band) {
        if (!new RegExp(`<p:ph\\b[^>]*\\btype="${type}"`).test(xml)) add(type, label, idx, geometry, '', isMaster || !masterPlaces(type) ? algn : null);
      }
    }
    if (xml === this.pkg.text(part)) return false;
    this.#writeSlide(part, xml);
    return true;
  }

  /**
   * Slide Master → Insert Placeholder: a content, text or picture
   * placeholder on a layout, in the middle of it, numbered past the
   * layout's own. Slides on the layout offer it to be filled.
   * @returns {number} the new shape's id
   */
  insertPlaceholder(layout, kind = 'content', geometry = null) {
    if (!/slideLayouts\//.test(String(layout)) || !this.pkg.has(layout)) throw new RangeError('placeholders are inserted on a layout');
    let xml = this.pkg.text(layout);
    const used = [...xml.matchAll(/<p:ph\b[^>]*\bidx="(\d+)"/g)].map((m) => Number(m[1]));
    const idx = Math.max(12, ...used) + 1;
    const id = nextShapeId(xml);
    const g = geometry || { x: this.size.width * 0.2, y: this.size.height * 0.3, w: this.size.width * 0.6, h: this.size.height * 0.45 };
    const type = kind === 'picture' ? ' type="pic"' : kind === 'text' ? ' type="body"' : '';
    const label = kind === 'picture' ? 'Picture Placeholder' : kind === 'text' ? 'Text Placeholder' : 'Content Placeholder';
    const words = kind === 'picture' ? '' : '<a:r><a:rPr lang="en-US"/><a:t>Click to edit Master text styles</a:t></a:r>';
    const sp = `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${label} ${id - 1}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph${type} sz="quarter" idx="${idx}"/></p:nvPr></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="${pxToEmu(g.x)}" y="${pxToEmu(g.y)}"/><a:ext cx="${pxToEmu(g.w)}" cy="${pxToEmu(g.h)}"/></a:xfrm></p:spPr>` +
      `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr lvl="0"/>${words}<a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
    const at = xml.lastIndexOf('</p:spTree>');
    xml = xml.slice(0, at) + sp + xml.slice(at);
    this.#writeSlide(layout, xml);
    return id;
  }

  /**
   * Formatting a placeholder in Slide Master view: on the master, a title
   * or body placeholder's look is the master's title or body text style —
   * what every slide's title and body inherit — and any other placeholder's
   * (and every layout placeholder's) is its own list style. `props` is any
   * of size (the first level's, the others keeping their proportion), bold,
   * italic, underline, color ('RRGGBB' or '#RRGGBB'), font and align.
   */
  setTextStyle(target, shapeId, props = {}) {
    const part = this.#partOf(target);
    if (!part || typeof target !== 'string') throw new RangeError(`no master or layout ${target}`);
    let xml = this.pkg.text(part);
    const range = this.#shapeRange(xml, shapeId);
    if (!range) throw new Error(`shape ${shapeId} not found`);
    const shapeXml = xml.slice(range.start, range.end);
    const type = /<p:ph\b([^>]*)\/?>/.exec(shapeXml) ? (/\btype="([^"]+)"/.exec(/<p:ph\b([^>]*)\/?>/.exec(shapeXml)[1])?.[1] || 'body') : null;
    const isMaster = /slideMasters\//.test(part);
    const kind = type === 'title' || type === 'ctrTitle' ? 'title' : type === 'body' || type === 'obj' || type === 'subTitle' ? 'body' : 'other';
    const levels = kind === 'body' ? 5 : 1;
    if (isMaster && type && kind !== 'other') {
      const tag = kind === 'title' ? 'p:titleStyle' : 'p:bodyStyle';
      const re = new RegExp(`<${tag}>[\\s\\S]*?</${tag}>|<${tag}/>`);
      const style = re.exec(xml);
      const edited = editListStyle(style ? style[0].replace(`<${tag}/>`, `<${tag}></${tag}>`) : `<${tag}></${tag}>`, tag, props, levels);
      if (style) xml = xml.slice(0, style.index) + edited + xml.slice(style.index + style[0].length);
      else if (/<p:txStyles>/.test(xml)) xml = xml.replace('<p:txStyles>', `<p:txStyles>${edited}`);
      else xml = xml.replace(/<\/p:sldMaster>\s*$/, `<p:txStyles>${edited}</p:txStyles></p:sldMaster>`);
    } else {
      const lst = /<a:lstStyle>[\s\S]*?<\/a:lstStyle>|<a:lstStyle\/>/.exec(shapeXml);
      let nextShape;
      if (lst) {
        const edited = editListStyle(lst[0] === '<a:lstStyle/>' ? '<a:lstStyle></a:lstStyle>' : lst[0], 'a:lstStyle', props, levels);
        nextShape = shapeXml.slice(0, lst.index) + edited + shapeXml.slice(lst.index + lst[0].length);
      } else if (/<a:bodyPr\b[^>]*\/>|<a:bodyPr\b[^>]*>[\s\S]*?<\/a:bodyPr>/.test(shapeXml)) {
        const edited = editListStyle('<a:lstStyle></a:lstStyle>', 'a:lstStyle', props, levels);
        nextShape = shapeXml.replace(/<a:bodyPr\b[^>]*\/>|<a:bodyPr\b[^>]*>[\s\S]*?<\/a:bodyPr>/, (m) => m + edited);
      } else {
        throw new Error('this shape holds no words to style');
      }
      xml = xml.slice(0, range.start) + nextShape + xml.slice(range.end);
    }
    if (xml === this.pkg.text(part)) return false;
    this.#writeSlide(part, xml);
    return true;
  }

  // ---- design: themes, variants, colours, fonts, effects -------------------
  //
  // Design → Themes and its neighbours rewrite the theme part every master
  // points at — its colour scheme, its font scheme, its format scheme, or
  // all three — and, for a whole theme, the master's background and colour
  // map. No slide is touched: whatever a slide states for itself it keeps,
  // and everything it inherits changes with the design.

  /** The deck's masters, in the presentation's own order. */
  masterParts() {
    const pres = this.pkg.text('ppt/presentation.xml');
    const listed = [...pres.matchAll(/<p:sldMasterId\b[^>]*\br:id="([^"]+)"/g)]
      .map((m) => this.rels.get(m[1])?.resolved)
      .filter((p) => p && this.pkg.has(p));
    if (listed.length) return listed;
    return (this.pkg.partNames() || []).filter((p) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(p));
  }

  /** The master a slide (by index), a layout or a master (by part) draws from. */
  masterOf(target = 0) {
    const part = this.#partOf(target);
    if (!part) return this.masterParts()[0] || null;
    if (/slideMasters\//.test(part)) return part;
    if (/slideLayouts\//.test(part)) return this.#masterFor(part);
    const layout = this.#layoutFor(part);
    return layout ? this.#masterFor(layout) : this.masterParts()[0] || null;
  }

  /**
   * What the Design tab shows as current: the theme's name and which
   * built-in theme it is (by name), its colours (the twelve slots, as hex),
   * the colour scheme's name, the fonts and their scheme's name, the format
   * scheme's name, and whether the master maps its backgrounds dark.
   */
  designInfo(target = 0) {
    const master = this.masterOf(target);
    const theme = this.#themeFor(master);
    const colors = {};
    for (const k of COLOUR_SLOTS) colors[k] = String(theme.colors[k] || '').replace('#', '').toUpperCase() || null;
    const builtIn = THEMES.find((t) => t.name === theme.name) || null;
    const dark = theme.clrMap?.bg1 === 'dk1';
    const variants = this.variants(target);
    const variant = variants.findIndex((v) => v.dark === dark && COLOUR_SLOTS.every((k) => String(v.colors[k]).toUpperCase() === colors[k]));
    return {
      master,
      themePart: master ? this.#themePartOf(master) : null,
      name: theme.name,
      builtIn: builtIn?.id || null,
      colors,
      colorName: theme.colorName,
      fonts: { ...theme.fonts },
      fontName: theme.fontName,
      effectName: theme.effectName,
      effects: EFFECT_PRESETS.find((p) => p.name === theme.effectName)?.id || null,
      dark,
      variant,
    };
  }

  /**
   * Design → Variants: four colourings of the current theme. A built-in
   * theme's are its own; a theme from elsewhere gets four drawn from the
   * colours it has now.
   */
  variants(target = 0) {
    const master = this.masterOf(target);
    const theme = this.#themeFor(master);
    const builtIn = THEMES.find((t) => t.name === theme.name);
    if (builtIn) return variantsOf(builtIn.palette, Boolean(builtIn.dark));
    const colors = {};
    for (const k of COLOUR_SLOTS) colors[k] = String(theme.colors[k] || '#000000').replace('#', '').toUpperCase();
    return variantsOf(colors, theme.clrMap?.bg1 === 'dk1');
  }

  /** Each distinct theme part the deck's masters use, rewritten by `edit`; answers whether anything changed. */
  #rewriteThemes(edit) {
    const done = new Set();
    let changed = false;
    for (const master of this.masterParts()) {
      const part = this.#themePartOf(master);
      if (!part || done.has(part) || !this.pkg.has(part)) continue;
      done.add(part);
      const xml = this.pkg.text(part);
      const next = edit(xml, part, master);
      if (next != null && next !== xml) {
        this.pkg.write_(part, Buffer.from(next, 'utf8'));
        changed = true;
      }
    }
    if (changed) this.#designChanged();
    return changed;
  }

  /** One element of a theme's `themeElements` replaced (or put in its place when the part lacks it). */
  static #withThemeElement(xml, tag, element) {
    const re = new RegExp('<a:' + tag + '\\b[^>]*?(?:/>|>[\\s\\S]*?</a:' + tag + '>)');
    if (re.test(xml)) return xml.replace(re, () => element);
    const order = ['clrScheme', 'fontScheme', 'fmtScheme'];
    const after = order.slice(0, order.indexOf(tag)).reverse().map((t) => new RegExp('</a:' + t + '>')).find((r) => r.test(xml));
    if (after) return xml.replace(after, (m) => m + element);
    return xml.replace(/<a:themeElements>/, (m) => m + element);
  }

  /** A master's colour map written light or dark, and its background, in place. */
  #styleMaster(master, { dark, background = null } = {}) {
    let xml = this.pkg.text(master);
    if (dark != null) {
      const map = `<p:clrMap ${clrMapAttrs(dark)}/>`;
      xml = /<p:clrMap\b[^>]*\/>/.test(xml) ? xml.replace(/<p:clrMap\b[^>]*\/>/, map) : xml.replace(/<\/p:cSld>/, (m) => m + map);
    }
    if (background) {
      if (/<p:bg>[\s\S]*?<\/p:bg>/.test(xml)) xml = xml.replace(/<p:bg>[\s\S]*?<\/p:bg>/, () => background);
      else xml = xml.replace(/<p:cSld\b[^>]*>/, (m) => m + background);
    }
    this.pkg.write_(master, Buffer.from(xml, 'utf8'));
  }

  /** The layouts a master lists. */
  layoutsOf(master) {
    return [...this.#relMap(master).values()].filter((r) => r.type === REL.layout && this.pkg.has(r.resolved)).map((r) => r.resolved);
  }

  /**
   * Design → Themes: one of the suite's themes, in one of its variants,
   * over the whole deck — its colours, fonts and effects as a new theme
   * part, the master's background treatment and colour map, and the
   * layouts' own backgrounds taken away so the new one shows through, as
   * PowerPoint's own swap of layouts does. Slides keep what they state.
   */
  applyTheme(id, { variant = 0 } = {}) {
    const t = themeById(id);
    if (!t) throw new Error(`no theme "${id}"`);
    const v = variantsOf(t.palette, Boolean(t.dark))[variant] || variantsOf(t.palette, Boolean(t.dark))[0];
    const xml = themePartXml({ name: t.name, colors: v.colors, colorName: variant ? `${t.name} ${v.name}` : t.name, fonts: t.fonts, fontName: t.name, effects: t.effects });
    const masters = this.masterParts();
    const written = new Set();
    for (const master of masters) {
      const part = this.#themePartOf(master);
      if (part && !written.has(part) && this.pkg.has(part)) {
        this.pkg.write_(part, Buffer.from(xml, 'utf8'));
        written.add(part);
      }
      this.#styleMaster(master, { dark: v.dark, background: masterBackgroundXml(t.background) });
      for (const layout of this.layoutsOf(master)) {
        const lx = this.pkg.text(layout);
        const nx = lx.replace(/<p:bg>[\s\S]*?<\/p:bg>/, '');
        if (nx !== lx) this.pkg.write_(layout, Buffer.from(nx, 'utf8'));
      }
    }
    this.#designChanged();
    return true;
  }

  /** Design → Variants: the current theme in another of its four colourings. */
  applyVariant(index, target = 0) {
    const v = this.variants(target)[index];
    if (!v) throw new RangeError(`no variant ${index}`);
    const name = this.#themeFor(this.masterOf(target)).name || 'Custom';
    this.#rewriteThemes((xml) => Deck.#withThemeElement(xml, 'clrScheme', clrSchemeXml(index ? `${name} ${v.name}` : name, v.colors)));
    for (const master of this.masterParts()) this.#styleMaster(master, { dark: v.dark });
    this.#designChanged();
    return true;
  }

  /** Design → Colours: a palette over the deck — a built-in one by id, or the twelve slots of Customise Colours with a name. */
  setThemeColors(spec, name = null) {
    const builtIn = typeof spec === 'string' ? PALETTES.find((p) => p.id === spec) : null;
    if (typeof spec === 'string' && !builtIn) throw new Error(`no palette "${spec}"`);
    const colors = builtIn ? builtIn.colors : spec;
    for (const k of COLOUR_SLOTS) {
      if (!/^#?[0-9a-f]{6}$/i.test(String(colors?.[k] || ''))) throw new Error(`${k} needs a colour like 1F6FB2`);
    }
    const label = String(name || builtIn?.name || 'Custom').trim() || 'Custom';
    return this.#rewriteThemes((xml) => Deck.#withThemeElement(xml, 'clrScheme', clrSchemeXml(label, colors)));
  }

  /** Design → Fonts: a heading and a body face — a built-in pair by id, or Customise Fonts' two with a name. */
  setThemeFonts(spec, name = null) {
    const pair = typeof spec === 'string' ? FONT_PAIRS.find((p) => p.id === spec) : null;
    if (typeof spec === 'string' && !pair) throw new Error(`no font pair "${spec}"`);
    const major = String(pair ? pair.major : spec?.major || '').trim();
    const minor = String(pair ? pair.minor : spec?.minor || '').trim();
    if (!major || !minor) throw new Error('a heading font and a body font are both needed');
    const label = String(name || pair?.name || 'Custom').trim() || 'Custom';
    return this.#rewriteThemes((xml) => Deck.#withThemeElement(xml, 'fontScheme', fontSchemeXml(label, { major, minor })));
  }

  /** Design → Effects: one of the format schemes. */
  setThemeEffects(id) {
    if (!EFFECT_PRESETS.some((p) => p.id === id)) throw new Error(`no effects "${id}"`);
    return this.#rewriteThemes((xml) => Deck.#withThemeElement(xml, 'fmtScheme', fmtSchemeXml(id)));
  }

  /**
   * A slide as it would look under another design, without writing a
   * byte: `{ theme, variant }` a built-in theme, `{ variant }` one of
   * this theme's variants, `{ colors }` a palette id or twelve slots,
   * `{ fonts }` a pair id or { major, minor }, `{ effects }` a format
   * scheme. The gallery's live thumbnails are these, drawn by the same
   * renderer as the slide itself.
   */
  previewSlide(index, spec = {}) {
    if (!this.slideParts[index]) throw new RangeError(`no slide at index ${index}`);
    const master = this.masterOf(index);
    const themePart = master ? this.#themePartOf(master) : null;
    let xml = themePart && this.pkg.has(themePart) ? this.pkg.text(themePart) : themePartXml({ name: 'Rutba', colors: THEMES[0].palette, fonts: THEMES[0].fonts });
    const current = this.#themeFor(master);
    let clrMap = { ...(current.clrMap || {}) };
    const design = {};
    const mapOf = (dark) => Object.fromEntries([...clrMapAttrs(dark).matchAll(/(\w+)="(\w+)"/g)].map((m) => [m[1], m[2]]));
    if (spec.theme) {
      const t = themeById(spec.theme);
      if (!t) throw new Error(`no theme "${spec.theme}"`);
      const v = variantsOf(t.palette, Boolean(t.dark))[spec.variant || 0];
      xml = themePartXml({ name: t.name, colors: v.colors, fonts: t.fonts, effects: t.effects });
      clrMap = mapOf(v.dark);
      design.dropLayoutBackground = true;
      design.masterBackgroundXml = masterBackgroundXml(t.background);
    } else if (spec.variant != null) {
      const v = this.variants(index)[spec.variant];
      if (!v) throw new RangeError(`no variant ${spec.variant}`);
      xml = Deck.#withThemeElement(xml, 'clrScheme', clrSchemeXml('Preview', v.colors));
      clrMap = mapOf(v.dark);
    } else if (spec.colors) {
      const colors = typeof spec.colors === 'string' ? PALETTES.find((p) => p.id === spec.colors)?.colors : spec.colors;
      if (!colors) throw new Error(`no palette "${spec.colors}"`);
      xml = Deck.#withThemeElement(xml, 'clrScheme', clrSchemeXml('Preview', colors));
    } else if (spec.fonts) {
      const pair = typeof spec.fonts === 'string' ? FONT_PAIRS.find((p) => p.id === spec.fonts) : spec.fonts;
      if (!pair) throw new Error(`no font pair "${spec.fonts}"`);
      xml = Deck.#withThemeElement(xml, 'fontScheme', fontSchemeXml('Preview', pair));
    } else if (spec.effects) {
      xml = Deck.#withThemeElement(xml, 'fmtScheme', fmtSchemeXml(spec.effects));
    }
    const theme = new Theme(xml, clrMap);
    design.theme = theme;
    if (design.masterBackgroundXml) {
      const holder = `<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld>${design.masterBackgroundXml}<p:spTree/></p:cSld></p:sldMaster>`;
      design.masterBackground = readSlideScene(holder, { theme }).background;
    }
    return this.#scene(index, design);
  }

  /**
   * Design → Effects' own little picture of a format scheme: three shapes
   * in the first three accents, subtle, moderate and intense, drawn the way
   * a shape that takes its look from the theme is — the same trio
   * PowerPoint's gallery shows.
   */
  effectsSample(id, target = 0) {
    const master = this.masterOf(target);
    const themePart = master ? this.#themePartOf(master) : null;
    const base = themePart && this.pkg.has(themePart) ? this.pkg.text(themePart) : themePartXml({ name: 'Rutba', colors: THEMES[0].palette, fonts: THEMES[0].fonts });
    const theme = new Theme(Deck.#withThemeElement(base, 'fmtScheme', fmtSchemeXml(id)), { ...(this.#themeFor(master).clrMap || {}) });
    const shape = (n, x, y) => `<p:sp><p:nvSpPr><p:cNvPr id="${n + 1}" name="s${n}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="1600000" cy="1000000"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom></p:spPr>` +
      `<p:style><a:lnRef idx="${n}"><a:schemeClr val="accent${n}"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="${n}"><a:schemeClr val="accent${n}"/></a:fillRef><a:effectRef idx="${n}"><a:schemeClr val="accent${n}"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style></p:sp>`;
    const xml = `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree>${shape(1, 250000, 300000)}${shape(2, 2150000, 800000)}${shape(3, 4050000, 1300000)}</p:spTree></p:cSld></p:sld>`;
    const scene = readSlideScene(xml, { theme });
    // On white, whatever the deck's own background: an outline or a shadow is judged against paper.
    return { size: { width: 6000000 / 9525, height: 2700000 / 9525 }, background: { type: 'solid', color: '#ffffff' }, shapes: scene.shapes };
  }

  /** Append a text box, which is how the editor adds new content. */
  /**
   * The footer band — Insert → Header & Footer: the footer's words, the
   * slide number and the date, each a placeholder as PowerPoint writes
   * them (`ftr`, `sldNum`, `dt`; the number and an automatic date are
   * fields). A layout or master that places the placeholder places it
   * here too; failing both, it goes along the bottom where PowerPoint's
   * own templates put it. A string sets the footer, '' or null takes it
   * off; `slideNumber` true or false; `date` { auto: true }, { text }
   * or false. Anything left undefined is left alone.
   */
  setFooter(slideIndex, { footer, slideNumber, date } = {}) {
    const part = this.#partOf(slideIndex);
    if (!part) throw new RangeError(`no slide at index ${slideIndex}`);
    let xml = this.pkg.text(part);
    const layoutPart = this.#layoutFor(part);
    const masterPart = layoutPart ? this.#masterFor(layoutPart) : null;
    const theme = this.#themeFor(masterPart);
    const placed = (type) => this.#placeholders(layoutPart, theme).get(`type:${type}`)?.geometry
      || this.#placeholders(masterPart, theme).get(`type:${type}`)?.geometry || null;
    const { cx, cy } = this.size;
    // PowerPoint's own band: the date left, the footer centred, the number right.
    const BAND = {
      dt: { x: 0.0693, w: 0.2354, algn: 'l' },
      ftr: { x: 0.3403, w: 0.3194, algn: 'ctr' },
      sldNum: { x: 0.7104, w: 0.2203, algn: 'r' },
    };
    const IDX = { dt: 10, ftr: 11, sldNum: 12 };
    const NAME = { dt: 'Date Placeholder', ftr: 'Footer Placeholder', sldNum: 'Slide Number Placeholder' };
    const blockOf = (type) => {
      const re = /<p:sp>[\s\S]*?<\/p:sp>/g;
      let m;
      while ((m = re.exec(xml))) {
        if (new RegExp('<p:ph\\b[^>]*\\btype="' + type + '"').test(m[0])) return { start: m.index, end: m.index + m[0].length, xml: m[0] };
      }
      return null;
    };
    const guid = () => '{' + [8, 4, 4, 4, 12].map((n) => Array.from({ length: n }, () => '0123456789ABCDEF'[Math.floor(Math.random() * 16)]).join('')).join('-') + '}';
    const runXml = (type, text) => {
      const safe = escapeXml(String(text));
      if (type === 'sldNum') return `<a:fld id="${guid()}" type="slidenum"><a:rPr lang="en-US"/><a:t>${safe}</a:t></a:fld>`;
      if (type === 'dt' && text === true) {
        const now = new Date();
        const shown = String(now.getDate()).padStart(2, '0') + '/' + String(now.getMonth() + 1).padStart(2, '0') + '/' + now.getFullYear();
        return `<a:fld id="${guid()}" type="datetime1"><a:rPr lang="en-US"/><a:t>${shown}</a:t></a:fld>`;
      }
      return `<a:r><a:rPr lang="en-US" dirty="0"/><a:t>${safe}</a:t></a:r>`;
    };
    const bodyXml = (type, text, inherited) => {
      const pPr = inherited ? '' : `<a:pPr algn="${BAND[type].algn}"/>`;
      return `<p:txBody><a:bodyPr/><a:lstStyle/><a:p>${pPr}${runXml(type, text)}</a:p></p:txBody>`;
    };
    const put = (type, text) => {
      const had = blockOf(type);
      if (text === false || text === null || text === '') {
        if (had) xml = xml.slice(0, had.start) + xml.slice(had.end);
        return;
      }
      const inherited = Boolean(placed(type));
      if (had) {
        const body = /<p:txBody>[\s\S]*?<\/p:txBody>/.exec(had.xml);
        const next = body ? had.xml.slice(0, body.index) + bodyXml(type, text, inherited) + had.xml.slice(body.index + body[0].length) : had.xml;
        xml = xml.slice(0, had.start) + next + xml.slice(had.end);
        return;
      }
      const id = nextShapeId(xml);
      const spPr = inherited
        ? '<p:spPr/>'
        : `<p:spPr><a:xfrm><a:off x="${Math.round(cx * BAND[type].x)}" y="${Math.round(cy * 0.9046)}"/><a:ext cx="${Math.round(cx * BAND[type].w)}" cy="${Math.round(cy * 0.0533)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>`;
      const sp = `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${NAME[type]} ${id}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>`
        + `<p:nvPr><p:ph type="${type}" sz="quarter" idx="${IDX[type]}"/></p:nvPr></p:nvSpPr>${spPr}${bodyXml(type, text, inherited)}</p:sp>`;
      const at = xml.lastIndexOf('</p:spTree>');
      if (at < 0) throw new Error('slide has no shape tree');
      xml = xml.slice(0, at) + sp + xml.slice(at);
    };
    if (footer !== undefined) put('ftr', footer == null ? '' : String(footer).trim());
    if (slideNumber !== undefined) put('sldNum', slideNumber ? String(slideIndex + 1) : '');
    if (date !== undefined) put('dt', !date ? '' : date.auto ? true : String(date.text ?? '').trim());
    this.#writeSlide(part, xml);
    return true;
  }

  /**
   * Design → Background Styles: this slide's own background, or null to take
   * it off and let the layout's or master's show through again. `spec` is
   * `{ colour: 'RRGGBB' }` (a solid colour), `{ scheme, lumMod?, lumOff? }`
   * (a theme colour, as `colourXml` writes one) or
   * `{ gradient: { from, to, angle? } }` — a two-stop linear gradient, `from`
   * and `to` each a colour spec of their own, `angle` in degrees (default 90,
   * top to bottom).
   *
   * Written as PowerPoint writes it: `<p:bg><p:bgPr>FILL<a:effectLst/>
   * </p:bgPr></p:bg>` as the first child of `<p:cSld>`, before `<p:spTree>`,
   * replacing whatever `<p:bg>` the slide already had — a `<p:bgPr>` or a
   * `<p:bgRef>` — and null takes the element out altogether.
   *
   * @returns {boolean} true when the slide's own background changed
   */
  setBackground(slideIndex, spec) {
    const part = this.#partOf(slideIndex);
    if (!part) throw new RangeError(`no slide at index ${slideIndex}`);
    const xml = this.pkg.text(part);
    const openCSld = /<p:cSld\b[^>]*>/.exec(xml);
    if (!openCSld) throw new Error(`slide ${slideIndex + 1} has no cSld`);
    const had = /<p:bg>[\s\S]*?<\/p:bg>/.exec(xml);
    const bgXml = spec ? backgroundXml(spec) : '';
    let next;
    if (had) next = xml.slice(0, had.index) + bgXml + xml.slice(had.index + had[0].length);
    else if (bgXml) {
      const at = openCSld.index + openCSld[0].length;
      next = xml.slice(0, at) + bgXml + xml.slice(at);
    } else return false;
    if (next === xml) return false;
    this.#writeSlide(part, next);
    return true;
  }

  /**
   * The slide's own background — what `setBackground` would need to write to
   * reproduce it — distinct from `slide(index).background`, which is this
   * when the slide states one and the layout's or master's otherwise. null
   * when the slide states none of its own.
   */
  background(slideIndex) {
    const part = this.#partOf(slideIndex);
    if (!part) throw new RangeError(`no slide at index ${slideIndex}`);
    return readOwnBackground(this.pkg.text(part));
  }

  /**
   * The formatting off a shape's words — Home → Clear all formatting: every
   * run keeps its text, its link, a field or a break, and nothing else, so
   * the words fall back to what the placeholder and the theme give them.
   * A paragraph's own properties — level, bullet, alignment — stay.
   */
  clearTextFormat(slideIndex, shapeId) {
    const shape = this.slide(slideIndex).shapes.find((s) => String(s.id) === String(shapeId));
    if (!shape) throw new Error(`shape ${shapeId} not found`);
    if (!shape.text?.paragraphs) throw new Error('Clear formatting needs a shape with words.');
    const paragraphs = shape.text.paragraphs.map((p) => {
      const { plain, runs, ...props } = p;
      return {
        ...props,
        runs: (runs || []).map((r) => ({
          text: r.text,
          ...(r.link ? { link: r.link } : {}),
          ...(r.field ? { field: r.field } : {}),
          ...(r.break ? { break: true } : {}),
        })),
      };
    });
    this.setText(slideIndex, shapeId, paragraphs);
    return true;
  }

  /**
   * A link on a shape's words — Insert → Link: every run of the shape points
   * at an External relationship the slide carries, as PowerPoint writes one;
   * null takes the link off every run (the relationship stays, unreferenced,
   * as an undone edit leaves it). Returns the relationship id, or null.
   */
  setLink(slideIndex, shapeId, url) {
    const part = this.#partOf(slideIndex);
    if (!part) throw new RangeError(`no slide at index ${slideIndex}`);
    const shape = this.slide(slideIndex).shapes.find((s) => String(s.id) === String(shapeId));
    if (!shape) throw new Error(`shape ${shapeId} not found`);
    if (!shape.text?.paragraphs) throw new Error('A link needs a shape with words.');
    const address = url == null ? '' : String(url).trim();
    const rId = address ? this.pkg.addRelationshipTo(part, REL.hyperlink, address, { external: true }) : null;
    const paragraphs = shape.text.paragraphs.map((p) => {
      const { plain, runs, ...props } = p;
      return {
        ...props,
        runs: (runs || []).map((r) => {
          const { link, ...rest } = r;
          return rId && r.text && !r.break ? { ...rest, link: rId } : rest;
        }),
      };
    });
    this.setText(slideIndex, shapeId, paragraphs);
    return rId;
  }

  /**
   * Every place the words appear across the deck — per slide and shape, how
   * many times, with the shape's words to show. Case-insensitive unless
   * asked. Only shapes with a text body are searched: a picture has none,
   * a table's cells are the table's.
   */
  findText(query, { matchCase = false } = {}) {
    const needle = String(query ?? '');
    if (!needle) return [];
    const target = matchCase ? needle : needle.toLowerCase();
    const hits = [];
    for (let i = 0; i < this.slideCount; i++) {
      for (const s of this.slide(i).shapes) {
        const paragraphs = s.text?.paragraphs;
        if (!paragraphs) continue;
        let count = 0;
        const words = [];
        for (const p of paragraphs) {
          const text = (p.runs || []).map((r) => r.text).join('');
          words.push(text);
          const hay = matchCase ? text : text.toLowerCase();
          for (let at = hay.indexOf(target); at !== -1; at = hay.indexOf(target, at + target.length)) count += 1;
        }
        if (count) hits.push({ slide: i, shape: s.id, name: s.name || s.kind, count, text: words.join(' ').replace(/\s+/g, ' ').trim().slice(0, 80) });
      }
    }
    return hits;
  }

  /**
   * The words replaced everywhere they appear, run by run so each keeps its
   * look; a field's text and a match straddling two runs are left alone.
   * Returns how many were replaced.
   */
  replaceText(query, replacement, { matchCase = false } = {}) {
    const needle = String(query ?? '');
    if (!needle) return 0;
    const target = matchCase ? needle : needle.toLowerCase();
    const re = new RegExp(needle.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), matchCase ? 'g' : 'gi');
    const after = String(replacement ?? '');
    let total = 0;
    for (let i = 0; i < this.slideCount; i++) {
      for (const s of this.slide(i).shapes) {
        const paragraphs = s.text?.paragraphs;
        if (!paragraphs || s.kind !== 'shape') continue;
        let count = 0;
        const next = paragraphs.map((p) => {
          const { plain, runs, ...props } = p;
          return {
            ...props,
            runs: (runs || []).map((r) => {
              if (!r.text || r.field || r.break) return r;
              const hay = matchCase ? r.text : r.text.toLowerCase();
              if (!hay.includes(target)) return r;
              count += (r.text.match(re) || []).length;
              return { ...r, text: r.text.replace(re, () => after) };
            }),
          };
        });
        if (count) {
          this.setText(i, s.id, next);
          total += count;
        }
      }
    }
    return total;
  }

  /**
   * Every occurrence across the deck, one entry per match rather than one
   * per shape — the find pane's own list, in slide/shape/paragraph/run
   * order, precise enough for `replace` to rewrite exactly one of them. A
   * table's cells are walked too, each hit carrying the row and column it
   * is in; a shape's own words carry neither. A field's cached text and a
   * line break are never searched, the way `findText` skips them.
   *
   * A match that would straddle two runs is not reported: the boundary
   * between two runs is where their look can change, and rewriting across
   * it would either lose the second run's formatting or have to invent a
   * blend of the two — so `find` only reports what `replace` can rewrite
   * cleanly, a match that sits inside a single run.
   */
  find(text, { matchCase = false } = {}) {
    const needle = String(text ?? '');
    if (!needle) return [];
    const target = matchCase ? needle : needle.toLowerCase();
    const hits = [];
    const walk = (paragraphs, extra) => {
      paragraphs.forEach((p, pi) => {
        (p.runs || []).forEach((r, ri) => {
          if (!r.text || r.field || r.break) return;
          const hay = matchCase ? r.text : r.text.toLowerCase();
          for (let at = hay.indexOf(target); at !== -1; at = hay.indexOf(target, at + target.length)) {
            hits.push({ ...extra, paragraph: pi, run: ri, offset: at, length: needle.length, text: r.text.slice(at, at + needle.length) });
          }
        });
      });
    };
    for (let i = 0; i < this.slideCount; i++) {
      for (const s of this.slide(i).shapes) {
        if (s.text?.paragraphs) walk(s.text.paragraphs, { slide: i, shape: s.id, row: null, col: null });
        if (s.table) {
          s.table.rows.forEach((row, ri) => {
            row.cells.forEach((cell, ci) => {
              if (cell.text?.paragraphs) walk(cell.text.paragraphs, { slide: i, shape: s.id, row: ri, col: ci });
            });
          });
        }
      }
    }
    return hits;
  }

  /**
   * One hit, exactly as `find` returned it, rewritten where it sits — the
   * run it is in keeps every other thing about its look, since only the
   * slice of its words that matched is replaced. Returns whether the shape
   * (or cell) and run were still there to rewrite: a hit kept past an edit
   * that removed its shape is simply skipped rather than thrown at.
   */
  replace(hit, replacement) {
    if (!hit) return false;
    const after = String(replacement ?? '');
    const scene = this.slideParts[hit.slide] ? this.slide(hit.slide) : null;
    const shape = scene?.shapes.find((s) => String(s.id) === String(hit.shape));
    const paragraphs = hit.row != null ? shape?.table?.rows?.[hit.row]?.cells?.[hit.col]?.text?.paragraphs : shape?.text?.paragraphs;
    const run = paragraphs?.[hit.paragraph]?.runs?.[hit.run];
    if (!run || run.text == null) return false;
    const text = run.text.slice(0, hit.offset) + after + run.text.slice(hit.offset + hit.length);
    const nextParagraphs = paragraphs.map((p, pi) => {
      const { plain, runs, ...props } = p;
      if (pi !== hit.paragraph) return { ...props, runs };
      return { ...props, runs: runs.map((r, ri) => (ri === hit.run ? { ...r, text } : r)) };
    });
    if (hit.row != null) this.setTableCell(hit.slide, hit.shape, hit.row, hit.col, nextParagraphs);
    else this.setText(hit.slide, hit.shape, nextParagraphs);
    return true;
  }

  /**
   * Every occurrence of `text` rewritten to `replacement`, across the whole
   * deck. Within one shape (or cell) the hits go last first — the last
   * paragraph, the last run, the rightmost offset — so replacing one never
   * shifts the offset of a hit still waiting its turn in the same run.
   * Returns how many were replaced.
   */
  replaceAll(text, replacement, { matchCase = false } = {}) {
    const groups = new Map();
    for (const hit of this.find(text, { matchCase })) {
      const key = `${hit.slide}:${hit.shape}:${hit.row ?? ''}:${hit.col ?? ''}`;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(hit);
    }
    let total = 0;
    for (const group of groups.values()) {
      group.sort((a, b) => b.paragraph - a.paragraph || b.run - a.run || b.offset - a.offset);
      for (const hit of group) if (this.replace(hit, replacement)) total += 1;
    }
    return total;
  }

  addTextBox(slideIndex, { x, y, w, h, paragraphs, name = 'TextBox' }) {
    const part = this.#partOf(slideIndex);
    if (!part) throw new RangeError(`no slide at index ${slideIndex}`);
    const xml = this.pkg.text(part);
    const id = nextShapeId(xml);
    const body = buildTextBody(paragraphs, '', -1, -1);
    const sp =
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escapeXml(name)} ${id}"/>` +
      `<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="${pxToEmu(x)}" y="${pxToEmu(y)}"/>` +
      `<a:ext cx="${Math.max(1, pxToEmu(w))}" cy="${Math.max(1, pxToEmu(h))}"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>${body}</p:sp>`;
    const at = xml.lastIndexOf('</p:spTree>');
    if (at < 0) throw new Error('slide has no shape tree');
    this.#writeSlide(part, xml.slice(0, at) + sp + xml.slice(at));
    return id;
  }

  /**
   * A picture on a slide.
   *
   * Three writes, and PowerPoint insists on all of them: the bytes as a media
   * part (`ppt/media/imageN.ext`), a relationship from the slide to it, and a
   * `<p:pic>` whose blip names the relationship. The extension's content type
   * is a package-wide default, as Office writes it. The drawn size is the
   * caller's, in pixels — the engine does not decode pictures — so a caller
   * that has read the picture's own size keeps its aspect.
   *
   * @param {number} slideIndex
   * @param {{ data: Buffer|Uint8Array|string, contentType: string, name?: string, x?: number, y?: number, w: number, h: number }} spec
   * @returns {{ id: number, part: string }} the shape id and the media part
   */
  /**
   * A picture's bytes as a media part plus a relationship from a slide to
   * it — numbered across every extension, as PowerPoint numbers them — the
   * one write `addPicture` and a picture fill (`setShapeStyle`) both need.
   * Bytes come in as a `Buffer`, a `Uint8Array` or base64 text, the way
   * they reach `documents.js` from the window's own file dialog.
   */
  #embedImage(part, { data, contentType }) {
    const ext = IMAGE_EXTENSIONS[String(contentType || '').toLowerCase()];
    if (!ext) throw new Error(`unsupported picture type: ${contentType} (png, jpeg, gif or bmp)`);
    const bytes = Buffer.isBuffer(data) ? data : data instanceof Uint8Array ? Buffer.from(data) : Buffer.from(String(data), 'base64');
    if (!bytes.length) throw new Error('the picture has no bytes');
    const names = this.pkg.partNames() || [];
    let n = 1;
    while (names.some((p) => p.startsWith(`ppt/media/image${n}.`))) n += 1;
    const media = `ppt/media/image${n}.${ext}`;
    this.pkg.ensureDefault(ext, ext === 'jpeg' ? 'image/jpeg' : contentType);
    this.pkg.addPart(media, bytes);
    const rId = this.pkg.addRelationshipTo(part, REL.image, `../media/image${n}.${ext}`);
    return { rId, media };
  }

  addPicture(slideIndex, { data, contentType, name = 'Picture', x = 0, y = 0, w, h }) {
    const part = this.#partOf(slideIndex);
    if (!part) throw new RangeError(`no slide at index ${slideIndex}`);
    if (!(w > 0) || !(h > 0)) throw new Error('a picture needs a positive width and height');
    const { rId, media } = this.#embedImage(part, { data, contentType });

    let xml = this.pkg.text(part);
    // A slide that has never had a relationship may not declare the prefix.
    const head = xml.slice(0, Math.max(0, xml.indexOf('<p:cSld')));
    if (!/xmlns:r=/.test(head)) {
      xml = xml.replace(/<p:sld\b/, '<p:sld xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
    }
    const id = nextShapeId(xml);
    const label = escapeXml(name);
    const pic =
      `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${label}" descr="${label}"/>` +
      `<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
      `<p:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
      `<p:spPr><a:xfrm><a:off x="${pxToEmu(x)}" y="${pxToEmu(y)}"/>` +
      `<a:ext cx="${Math.max(1, pxToEmu(w))}" cy="${Math.max(1, pxToEmu(h))}"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
    const at = xml.lastIndexOf('</p:spTree>');
    if (at < 0) throw new Error('slide has no shape tree');
    this.#writeSlide(part, xml.slice(0, at) + pic + xml.slice(at));
    return { id, part: media };
  }

  /**
   * A preset shape — a rectangle, an oval, an arrow, a star — with a fill, a
   * line and, when asked, centred text.
   *
   * Written the way PowerPoint writes one it has just drawn: preset geometry,
   * a solid fill in the theme's first accent, a line in the same accent
   * darkened by half, and a text body anchored to the middle, so the shape
   * takes the deck's colours and a theme change recolours it. A hex colour is
   * accepted anywhere a scheme colour is.
   *
   * @param {number} slideIndex
   * @param {{ preset?: string, x?: number, y?: number, w: number, h: number,
   *   fill?: string|{ scheme: string, lumMod?: number }|'none',
   *   line?: { color?: string|{ scheme: string, lumMod?: number }, width?: number }|'none',
   *   text?: string|Array, name?: string }} spec pixels and points
   * @returns {number} the shape id
   */
  addShape(slideIndex, {
    preset = 'rect', x = 0, y = 0, w, h,
    fill = { scheme: 'accent1' },
    line = { color: { scheme: 'accent1', lumMod: 50 }, width: 1 },
    text = null, name = null,
  } = {}) {
    const part = this.#partOf(slideIndex);
    if (!part) throw new RangeError(`no slide at index ${slideIndex}`);
    if (!(w > 0) || !(h >= 0)) throw new Error('a shape needs a positive width and a height');
    if (!/^[A-Za-z][A-Za-z0-9]*$/.test(String(preset))) throw new Error(`not a preset geometry: ${preset}`);

    const xml = this.pkg.text(part);
    const id = nextShapeId(xml);
    const label = escapeXml(name || `${PRESET_NAMES[preset] || 'Shape'} ${id}`);
    const fillXml = fill === 'none' || fill?.type === 'none' ? '<a:noFill/>' : `<a:solidFill>${colourXml(fill)}</a:solidFill>`;
    const lineXml = line === 'none' || line?.type === 'none'
      ? '<a:ln><a:noFill/></a:ln>'
      : line
        ? `<a:ln w="${Math.round((line.width ?? 1) * 12700)}"><a:solidFill>${colourXml(line.color ?? { scheme: 'accent1' })}</a:solidFill></a:ln>`
        : '';
    const paragraphs = text == null
      ? []
      : typeof text === 'string'
        ? text.split('\n').map((t) => ({ align: 'center', runs: [{ text: t, color: '#FFFFFF' }] }))
        : text;
    const body = paragraphs.length
      ? buildTextBody(paragraphs, '', -1, -1).replace('<a:bodyPr/>', '<a:bodyPr rtlCol="0" anchor="ctr"/>')
      : '<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/></a:p></p:txBody>';
    const sp =
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${label}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="${pxToEmu(x)}" y="${pxToEmu(y)}"/>` +
      `<a:ext cx="${Math.max(1, pxToEmu(w))}" cy="${Math.max(0, pxToEmu(h))}"/></a:xfrm>` +
      `<a:prstGeom prst="${preset}"><a:avLst/></a:prstGeom>${fillXml}${lineXml}</p:spPr>${body}</p:sp>`;
    const at = xml.lastIndexOf('</p:spTree>');
    if (at < 0) throw new Error('slide has no shape tree');
    this.#writeSlide(part, xml.slice(0, at) + sp + xml.slice(at));
    return id;
  }

  /**
   * A table on a slide — Insert → Table.
   *
   * Written exactly as PowerPoint writes one it has just drawn: a graphic
   * frame holding `a:tbl`, `firstRow`/`bandRow` banding and Medium Style 2 —
   * Accent 1 (the gallery's own first pick), equal columns, and a row per
   * entry whose cell is a run of words or, empty, just an `endParaRPr` so it
   * still measures a line's height.
   *
   * @param {number} slideIndex
   * @param {{ rows?: number, cols?: number, x?: number, y?: number, w?: number, h?: number, cells?: string[][] }} [spec] pixels; cells[row][col] is that cell's words
   * @returns {number} the frame's id
   */
  addTable(slideIndex, { rows = 3, cols = 3, x, y, w, h, cells = null } = {}) {
    const part = this.#partOf(slideIndex);
    if (!part) throw new RangeError(`no slide at index ${slideIndex}`);
    const r = Math.max(1, Math.round(rows));
    const c = Math.max(1, Math.round(cols));
    const { cx, cy } = this.size;
    const frameW = w != null ? pxToEmu(w) : Math.round(cx * 0.6);
    const frameH = h != null ? pxToEmu(h) : TABLE_ROW_H * r;
    const frameX = x != null ? pxToEmu(x) : Math.round((cx - frameW) / 2);
    const frameY = y != null ? pxToEmu(y) : Math.round((cy - frameH) / 2);
    const colW = Math.round(frameW / c);

    const xml = this.pkg.text(part);
    const id = nextShapeId(xml);
    const grid = `<a:tblGrid>${Array.from({ length: c }, () => `<a:gridCol w="${colW}"/>`).join('')}</a:tblGrid>`;
    const rowsXml = Array.from({ length: r }, (_, ri) =>
      `<a:tr h="${TABLE_ROW_H}">${Array.from({ length: c }, (_, ci) => tableCellXml(cells?.[ri]?.[ci] ?? null)).join('')}</a:tr>`
    ).join('');
    const tbl = `<a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>${TABLE_STYLE_ID}</a:tableStyleId></a:tblPr>${grid}${rowsXml}</a:tbl>`;
    const frame =
      `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id}"/>` +
      `<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>` +
      `<p:xfrm><a:off x="${frameX}" y="${frameY}"/><a:ext cx="${frameW}" cy="${frameH}"/></p:xfrm>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">${tbl}</a:graphicData></a:graphic></p:graphicFrame>`;
    const at = xml.lastIndexOf('</p:spTree>');
    if (at < 0) throw new Error('slide has no shape tree');
    this.#writeSlide(part, xml.slice(0, at) + frame + xml.slice(at));
    return id;
  }

  /**
   * A chart on a slide — Insert → Chart.
   *
   * Written the way PowerPoint writes a chart it has just drawn: a part of
   * its own (`chartPartXml`, the writer Word and Worksheets already use), a
   * relationship from the slide to it, and a graphic frame whose
   * `a:graphicData` names the DrawingML chart namespace and points at the
   * part through `c:chart r:id`. The one thing PowerPoint's own gallery adds
   * that this does not is an embedded workbook — `chartPartXml` bakes the
   * values into the part as a cache instead, which is exactly what a chart
   * with nothing to reference needs, and the same shape PowerPoint itself
   * reads happily whenever a chart's embedding has gone missing.
   *
   * @param {number} slideIndex
   * @param {{ type?: 'column'|'bar'|'line'|'area'|'pie'|'doughnut', title?: string,
   *   categories?: string[], series: Array<{ name?: string, values: Array<number|null> }>,
   *   x?: number, y?: number, w?: number, h?: number }} spec pixels; the frame
   *   defaults to centred, six tenths of the slide wide and just over half tall.
   * @returns {number} the frame's id
   */
  addChart(slideIndex, { type = 'column', title = null, categories = [], series = [], x, y, w, h } = {}) {
    const part = this.#partOf(slideIndex);
    if (!part) throw new RangeError(`no slide at index ${slideIndex}`);
    if (!CHART_KINDS.includes(type)) throw new Error(`unknown chart type: ${type}`);
    if (!series.length) throw new Error('a chart needs at least one series');

    const { cx, cy } = this.size;
    const frameW = w != null ? pxToEmu(w) : Math.round(cx * 0.6);
    const frameH = h != null ? pxToEmu(h) : Math.round(cy * 0.55);
    const frameX = x != null ? pxToEmu(x) : Math.round((cx - frameW) / 2);
    const frameY = y != null ? pxToEmu(y) : Math.round((cy - frameH) / 2);

    const n = this.pkg.nextPartNumber('ppt/charts/', 'chart');
    const chartPart = `ppt/charts/chart${n}.xml`;
    this.pkg.addPart(chartPart, chartPartXml({
      kind: type,
      title: title || undefined,
      categories: { values: categories },
      series: series.map((s) => ({ name: s.name, values: s.values })),
    }), CT.chart);
    const rId = this.pkg.addRelationshipTo(part, REL.chart, `../charts/chart${n}.xml`);

    let xml = this.pkg.text(part);
    // A slide that has never had a relationship may not declare the prefix.
    const head = xml.slice(0, Math.max(0, xml.indexOf('<p:cSld')));
    if (!/xmlns:r=/.test(head)) {
      xml = xml.replace(/<p:sld\b/, '<p:sld xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
    }
    const id = nextShapeId(xml);
    const frame =
      `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Chart ${id}"/>` +
      `<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
      `<p:xfrm><a:off x="${frameX}" y="${frameY}"/><a:ext cx="${frameW}" cy="${frameH}"/></p:xfrm>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
      `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${rId}"/>` +
      `</a:graphicData></a:graphic></p:graphicFrame>`;
    const at = xml.lastIndexOf('</p:spTree>');
    if (at < 0) throw new Error('slide has no shape tree');
    this.#writeSlide(part, xml.slice(0, at) + frame + xml.slice(at));
    return id;
  }

  /** The chart part a chart frame's `c:chart r:id` resolves to. */
  #chartPartFor(slideIndex, shapeId) {
    const part = this.#partOf(slideIndex);
    if (!part) throw new RangeError(`no slide at index ${slideIndex}`);
    const xml = this.pkg.text(part);
    const range = this.#shapeRange(xml, shapeId);
    if (!range) throw new Error(`shape ${shapeId} not found on slide ${slideIndex + 1}`);
    const frameXml = xml.slice(range.start, range.end);
    const m = /<c:chart\b[^>]*\br:id="([^"]+)"/.exec(frameXml);
    if (!m) throw new Error(`shape ${shapeId} is not a chart`);
    const rel = this.#relMap(part).get(m[1]);
    const chartPart = rel?.resolved;
    if (!chartPart || !this.pkg.has(chartPart)) throw new Error(`shape ${shapeId}'s chart part is missing`);
    return chartPart;
  }

  /**
   * A chart's data, read back from its own part — the same reader that
   * paints it. Plain data: a type, a title (or null), the categories and
   * each series with its name and values.
   */
  chartData(slideIndex, shapeId) {
    const chartPart = this.#chartPartFor(slideIndex, shapeId);
    const spec = parseChartXml(this.pkg.text(chartPart));
    if (!spec) throw new Error(`shape ${shapeId}'s chart part could not be read`);
    return {
      type: spec.type,
      title: spec.title,
      categories: spec.categories,
      series: spec.series.map((s) => ({ name: s.name, values: s.values })),
    };
  }

  /**
   * The chart data dialog's Apply: the chart part rewritten from new values
   * with `chartPartXml`, the way `addChart` first wrote it. The frame itself
   * — its id, its name, its position — is untouched.
   */
  setChartData(slideIndex, shapeId, { type, title, categories = [], series = [] } = {}) {
    const chartPart = this.#chartPartFor(slideIndex, shapeId);
    if (!series.length) throw new Error('a chart needs at least one series');
    const kind = type ?? parseChartXml(this.pkg.text(chartPart))?.type ?? 'column';
    if (!CHART_KINDS.includes(kind)) throw new Error(`unknown chart type: ${kind}`);
    const xml = chartPartXml({
      kind,
      title: title || undefined,
      categories: { values: categories },
      series: series.map((s) => ({ name: s.name, values: s.values })),
    });
    this.pkg.write_(chartPart, Buffer.from(xml, 'utf8'));
    this.dirty = true;
    return true;
  }

  /** The range of a table's `a:tbl` within its graphic frame, on the slide's own XML. */
  #tableRange(slideIndex, shapeId) {
    const part = this.#partOf(slideIndex);
    if (!part) throw new RangeError(`no slide at index ${slideIndex}`);
    const xml = this.pkg.text(part);
    const range = this.#shapeRange(xml, shapeId);
    if (!range) throw new Error(`shape ${shapeId} not found on slide ${slideIndex + 1}`);
    const frameXml = xml.slice(range.start, range.end);
    const tbl = /<a:tbl>[\s\S]*<\/a:tbl>/.exec(frameXml);
    if (!tbl) throw new Error(`shape ${shapeId} is not a table`);
    return { part, xml, range, frameXml, tbl };
  }

  /**
   * One cell's words rewritten through `buildTextBody`, the way a text box's
   * are — runs and their looks kept, only the words themselves replaced.
   */
  setTableCell(slideIndex, shapeId, row, col, paragraphs) {
    const { part, xml, range, frameXml, tbl } = this.#tableRange(slideIndex, shapeId);
    const trs = [...tbl[0].matchAll(TR_RE)];
    const tr = trs[row];
    if (!tr) throw new Error(`no row ${row} in table`);
    const tcs = [...tr[0].matchAll(TC_RE)];
    const tc = tcs[col];
    if (!tc) throw new Error(`no column ${col} in table`);
    const cellXml = tc[0];
    const bodyStart = cellXml.indexOf('<a:txBody>');
    const bodyEnd = cellXml.indexOf('</a:txBody>');
    const body = buildTextBody(paragraphs, cellXml, bodyStart, bodyEnd)
      .replace(/^<p:txBody>/, '<a:txBody>')
      .replace(/<\/p:txBody>$/, '</a:txBody>');
    const newCell = bodyStart < 0
      ? cellXml.slice(0, cellXml.lastIndexOf('</a:tc>')) + body + cellXml.slice(cellXml.lastIndexOf('</a:tc>'))
      : cellXml.slice(0, bodyStart) + body + cellXml.slice(bodyEnd + '</a:txBody>'.length);
    const newRow = tr[0].slice(0, tc.index) + newCell + tr[0].slice(tc.index + tc[0].length);
    const newTbl = tbl[0].slice(0, tr.index) + newRow + tbl[0].slice(tr.index + tr[0].length);
    const newFrame = frameXml.slice(0, tbl.index) + newTbl + frameXml.slice(tbl.index + tbl[0].length);
    this.#writeSlide(part, xml.slice(0, range.start) + newFrame + xml.slice(range.end));
    return true;
  }

  /**
   * A row added to a table, at `at` (append when omitted): a cell per grid
   * column, empty, in the neighbouring row's height — the frame's own
   * `p:xfrm` height grows by that same amount, so the table does not
   * overlap what is drawn beneath it.
   */
  insertTableRow(slideIndex, shapeId, at) {
    const { part, xml, range, frameXml, tbl } = this.#tableRange(slideIndex, shapeId);
    const trs = [...tbl[0].matchAll(TR_RE)];
    const cols = (tbl[0].match(/<a:gridCol\b/g) || []).length || 1;
    const insertAt = Math.max(0, Math.min(trs.length, at == null ? trs.length : at));
    const near = trs[Math.min(insertAt, trs.length - 1)]?.[0] || trs[0]?.[0] || '';
    const rowH = Number(/\bh="(\d+)"/.exec(near)?.[1]) || TABLE_ROW_H;
    const newRow = `<a:tr h="${rowH}">${Array.from({ length: cols }, () => tableCellXml(null)).join('')}</a:tr>`;
    const at2 = insertAt < trs.length ? trs[insertAt].index : tbl[0].lastIndexOf('</a:tbl>');
    const newTbl = tbl[0].slice(0, at2) + newRow + tbl[0].slice(at2);
    const newFrame = growFrameHeight(frameXml.slice(0, tbl.index) + newTbl + frameXml.slice(tbl.index + tbl[0].length), rowH);
    this.#writeSlide(part, xml.slice(0, range.start) + newFrame + xml.slice(range.end));
    return true;
  }

  /** A row taken out of a table; the frame's own height shrinks by its own. */
  removeTableRow(slideIndex, shapeId, at) {
    const { part, xml, range, frameXml, tbl } = this.#tableRange(slideIndex, shapeId);
    const trs = [...tbl[0].matchAll(TR_RE)];
    if (trs.length <= 1) throw new Error('a table needs at least one row');
    const tr = trs[at];
    if (!tr) throw new Error(`no row ${at} in table`);
    const rowH = Number(/\bh="(\d+)"/.exec(tr[0])?.[1]) || TABLE_ROW_H;
    const newTbl = tbl[0].slice(0, tr.index) + tbl[0].slice(tr.index + tr[0].length);
    const newFrame = growFrameHeight(frameXml.slice(0, tbl.index) + newTbl + frameXml.slice(tbl.index + tbl[0].length), -rowH);
    this.#writeSlide(part, xml.slice(0, range.start) + newFrame + xml.slice(range.end));
    return true;
  }

  /**
   * A column added to a table, at `at` (append when omitted): a `gridCol`
   * and an empty cell in every row. Widths share the frame's own width —
   * the frame itself does not grow, the way inserting a row grows its
   * height, because PowerPoint keeps a table's own width and only splits it
   * differently among more columns.
   */
  insertTableColumn(slideIndex, shapeId, at) {
    const { part, xml, range, frameXml, tbl } = this.#tableRange(slideIndex, shapeId);
    const grid = /<a:tblGrid>[\s\S]*?<\/a:tblGrid>/.exec(tbl[0]);
    if (!grid) throw new Error('the table has no grid');
    const cols = [...grid[0].matchAll(GRIDCOL_RE)];
    const insertAt = Math.max(0, Math.min(cols.length, at == null ? cols.length : at));
    const frameW = frameExtCx(frameXml);
    const nextCount = cols.length + 1;
    const colW = Math.round(frameW / nextCount);
    const newGrid = `<a:tblGrid>${Array.from({ length: nextCount }, () => `<a:gridCol w="${colW}"/>`).join('')}</a:tblGrid>`;
    let newTbl = tbl[0].slice(0, grid.index) + newGrid + tbl[0].slice(grid.index + grid[0].length);
    newTbl = newTbl.replace(TR_RE, (rowXml) => {
      const tcs = [...rowXml.matchAll(TC_RE)];
      const pos = insertAt < tcs.length ? tcs[insertAt].index : rowXml.lastIndexOf('</a:tr>');
      return rowXml.slice(0, pos) + tableCellXml(null) + rowXml.slice(pos);
    });
    const newFrame = frameXml.slice(0, tbl.index) + newTbl + frameXml.slice(tbl.index + tbl[0].length);
    this.#writeSlide(part, xml.slice(0, range.start) + newFrame + xml.slice(range.end));
    return true;
  }

  /** A column taken out of a table — its `gridCol` and the cell it gave every row. */
  removeTableColumn(slideIndex, shapeId, at) {
    const { part, xml, range, frameXml, tbl } = this.#tableRange(slideIndex, shapeId);
    const grid = /<a:tblGrid>[\s\S]*?<\/a:tblGrid>/.exec(tbl[0]);
    if (!grid) throw new Error('the table has no grid');
    const cols = [...grid[0].matchAll(GRIDCOL_RE)];
    if (cols.length <= 1) throw new Error('a table needs at least one column');
    const col = cols[at];
    if (!col) throw new Error(`no column ${at} in table`);
    const frameW = frameExtCx(frameXml);
    const nextCount = cols.length - 1;
    const colW = Math.round(frameW / nextCount);
    const newGrid = `<a:tblGrid>${Array.from({ length: nextCount }, () => `<a:gridCol w="${colW}"/>`).join('')}</a:tblGrid>`;
    let newTbl = tbl[0].slice(0, grid.index) + newGrid + tbl[0].slice(grid.index + grid[0].length);
    newTbl = newTbl.replace(TR_RE, (rowXml) => {
      const tcs = [...rowXml.matchAll(TC_RE)];
      const tc = tcs[at];
      if (!tc) return rowXml;
      return rowXml.slice(0, tc.index) + rowXml.slice(tc.index + tc[0].length);
    });
    const newFrame = frameXml.slice(0, tbl.index) + newTbl + frameXml.slice(tbl.index + tbl[0].length);
    this.#writeSlide(part, xml.slice(0, range.start) + newFrame + xml.slice(range.end));
    return true;
  }

  /** Duplicate a slide, which is the cheapest way to add one that matches. */
  duplicateSlide(index) {
    const src = this.slideParts[index];
    if (!src) throw new RangeError(`no slide at index ${index}`);
    const n = this.pkg.nextPartNumber('ppt/slides/', 'slide');
    const newPart = `ppt/slides/slide${n}.xml`;
    this.pkg.addPart(newPart, this.pkg.read(src.part), CT.slide);

    // Copy the slide's relationships so its layout and images resolve.
    const relsPart = `ppt/slides/_rels/slide${n}.xml.rels`;
    const srcRels = `ppt/slides/_rels/${src.part.split('/').pop()}.rels`;
    if (this.pkg.has(srcRels)) {
      this.pkg.addPart(relsPart, this.pkg.read(srcRels), 'application/vnd.openxmlformats-package.relationships+xml');
    }

    const rId = this.pkg.addRelationshipTo('ppt/presentation.xml', REL.slide, `slides/slide${n}.xml`);
    const presXml = this.pkg.text('ppt/presentation.xml');
    const maxId = Math.max(255, ...[...presXml.matchAll(/<p:sldId\s+id="(\d+)"/g)].map((m) => Number(m[1])));
    const entry = `<p:sldId id="${maxId + 1}" r:id="${rId}"/>`;
    const anchor = this.slideParts[index].rId;
    const marker = new RegExp(`<p:sldId\\b[^>]*r:id="${anchor}"[^>]*/>`);
    const next = marker.test(presXml)
      ? presXml.replace(marker, (m) => m + entry)
      : presXml.replace('</p:sldIdLst>', `${entry}</p:sldIdLst>`);
    this.pkg.write_('ppt/presentation.xml', Buffer.from(next, 'utf8'));
    this.dirty = true;
    this.#load();
    this.#syncSections();
    return this.slideParts.findIndex((s) => s.part === newPart);
  }

  /**
   * A new slide, after the one at `after`.
   *
   * Until this existed a deck could only grow by duplicating a slide it already
   * had — which meant a new presentation could not gain a second slide at all,
   * and is the single most-used button in every presentation program there is.
   *
   * The layout is borrowed from a neighbouring slide rather than invented: a
   * deck's layouts carry its theme, its placeholder geometry and its fonts, and
   * a slide that points at a layout the deck does not have is a slide that
   * opens blank in PowerPoint.
   *
   * @param {number} after index to insert after; -1 puts it first
   * @param {{ layout?: 'title'|'obj'|'blank', title?: string, body?: string|string[] }} [spec]
   * @returns {number} the index of the new slide
   */
  insertSlide(after = this.slideCount - 1, spec = {}) {
    const n = this.pkg.nextPartNumber('ppt/slides/', 'slide');
    const newPart = `ppt/slides/slide${n}.xml`;

    const { layout = 'obj', title = '', body = '' } = spec;
    const content = { layout, ...(layout === 'blank' ? {} : { title, body }) };
    this.pkg.addPart(newPart, Buffer.from(slideXml(content), 'utf8'), CT.slide);

    // The layout of the slide we are inserting after, or of the first slide, or
    // — for a deck with no slides at all — whatever layout part exists.
    const neighbour = this.slideParts[after] || this.slideParts[0] || null;
    // The gallery names a layout; otherwise the neighbour's is borrowed.
    let layoutTarget = spec.layoutPart && this.pkg.has(spec.layoutPart) ? `../slideLayouts/${spec.layoutPart.split('/').pop()}` : null;
    if (!layoutTarget && neighbour) {
      for (const r of this.pkg.rels(neighbour.part) || []) {
        if (r.type === REL.layout) layoutTarget = r.target;
      }
    }
    if (!layoutTarget) {
      const any = (this.pkg.partNames() || []).find((p) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(p));
      layoutTarget = any ? `../slideLayouts/${any.split('/').pop()}` : '../slideLayouts/slideLayout2.xml';
    }

    this.pkg.addPart(
      `ppt/slides/_rels/slide${n}.xml.rels`,
      Buffer.from(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
          `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rId1" Type="${REL.layout}" Target="${layoutTarget}"/>` +
          `</Relationships>`,
        'utf8'
      ),
      'application/vnd.openxmlformats-package.relationships+xml'
    );

    const rId = this.pkg.addRelationshipTo('ppt/presentation.xml', REL.slide, `slides/slide${n}.xml`);
    const presXml = this.pkg.text('ppt/presentation.xml');
    const maxId = Math.max(255, ...[...presXml.matchAll(/<p:sldId\s+id="(\d+)"/g)].map((m) => Number(m[1])));
    const entry = `<p:sldId id="${maxId + 1}" r:id="${rId}"/>`;

    const anchor = this.slideParts[after]?.rId;
    const marker = anchor ? new RegExp(`<p:sldId\\b[^>]*r:id="${anchor}"[^>]*/>`) : null;
    const next = marker && marker.test(presXml)
      ? presXml.replace(marker, (m) => m + entry)
      : after < 0 && /<p:sldIdLst>/.test(presXml)
        ? presXml.replace('<p:sldIdLst>', `<p:sldIdLst>${entry}`)
        : presXml.replace('</p:sldIdLst>', `${entry}</p:sldIdLst>`);

    this.pkg.write_('ppt/presentation.xml', Buffer.from(next, 'utf8'));
    this.dirty = true;
    this.#load();
    this.#syncSections();
    return this.slideParts.findIndex((s) => s.part === newPart);
  }

  /**
   * Speaker notes for a slide.
   *
   * The deck has always been able to read them; this writes them. A notesSlide
   * is a part like any other, related from the slide — so an existing one is
   * rewritten in place and a missing one is created, and the presentation's
   * notes size (already declared by every deck this package builds) governs how
   * PowerPoint lays it out.
   *
   * @param {number} index
   * @param {string} text one paragraph per line
   */
  setNotes(index, text) {
    const entry = this.slideParts[index];
    if (!entry) throw new RangeError(`no slide at index ${index}`);

    const paragraphs = String(text ?? '')
      .split('\n')
      .map((line) =>
        line.trim()
          ? `<a:p><a:r><a:rPr lang="en-GB" dirty="0"/><a:t>${escapeXml(line)}</a:t></a:r></a:p>`
          : '<a:p/>'
      )
      .join('') || '<a:p/>';

    const xml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
      `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
      `<p:cSld><p:spTree>` +
      `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      `<p:grpSpPr/>` +
      `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder 1"/>` +
      `<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
      `<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>` +
      `<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</p:txBody></p:sp>` +
      `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;

    // An existing notes part is rewritten; a missing one is added and related.
    let target = null;
    for (const r of this.pkg.rels(entry.part) || []) {
      if (r.type === REL.notes) target = r.resolved;
    }

    if (target && this.pkg.has(target)) {
      this.pkg.write_(target, Buffer.from(xml, 'utf8'));
    } else {
      const n = this.pkg.nextPartNumber('ppt/notesSlides/', 'notesSlide');
      const part = `ppt/notesSlides/notesSlide${n}.xml`;
      this.pkg.addPart(part, Buffer.from(xml, 'utf8'), CT.notes);
      this.pkg.addRelationshipTo(entry.part, REL.notes, `../notesSlides/notesSlide${n}.xml`);
    }

    this.dirty = true;
    this.#load();
    return true;
  }

  /** Remove a slide from the show. The part stays, unreferenced, as PowerPoint does. */
  removeSlide(index) {
    const entry = this.slideParts[index];
    if (!entry) throw new RangeError(`no slide at index ${index}`);
    const presXml = this.pkg.text('ppt/presentation.xml');
    const next = presXml.replace(new RegExp(`<p:sldId\\b[^>]*r:id="${entry.rId}"[^>]*/>`), '');
    this.pkg.write_('ppt/presentation.xml', Buffer.from(next, 'utf8'));
    this.dirty = true;
    this.#load();
    this.#syncSections();
    return true;
  }

  /** Reorder: move the slide at `from` to `to`. */
  moveSlide(from, to) {
    const entry = this.slideParts[from];
    const target = this.slideParts[to];
    if (!entry || !target || from === to) return false;
    const presXml = this.pkg.text('ppt/presentation.xml');
    const re = new RegExp(`<p:sldId\\b[^>]*r:id="${entry.rId}"[^>]*/>`);
    const m = re.exec(presXml);
    if (!m) return false;
    const without = presXml.replace(re, '');
    const targetRe = new RegExp(`<p:sldId\\b[^>]*r:id="${target.rId}"[^>]*/>`);
    const next = without.replace(targetRe, (t) => (to > from ? t + m[0] : m[0] + t));
    this.pkg.write_('ppt/presentation.xml', Buffer.from(next, 'utf8'));
    this.dirty = true;
    this.#load();
    // The moved slide belongs to the section it landed in, not the one it left.
    this.#syncSections([entry.id]);
    return true;
  }

  /** Whether a slide is left out of the show — Slide Show → Hide Slide. */
  isSlideHidden(index) {
    const entry = this.slideParts[index];
    if (!entry) throw new RangeError(`no slide at index ${index}`);
    return slideHiddenFrom(this.pkg.text(entry.part));
  }

  /**
   * Hide a slide, or bring it back — Slide Show → Hide Slide. PowerPoint
   * marks a hidden slide `show="0"` on the slide's own root tag, and writes
   * no such attribute at all once the slide is shown again, so un-hiding
   * takes the attribute off rather than writing `show="1"`.
   *
   * @returns {boolean} true when the slide's state changed
   */
  setSlideHidden(index, hidden) {
    const entry = this.slideParts[index];
    if (!entry) throw new RangeError(`no slide at index ${index}`);
    const xml = this.pkg.text(entry.part);
    const m = /<p:sld\b[^>]*>/.exec(xml);
    if (!m) throw new Error(`slide ${index + 1} has no root element`);
    const cleaned = m[0].replace(/\s+show="[^"]*"/, '');
    const next = hidden ? cleaned.replace(/>$/, ' show="0">') : cleaned;
    if (next === m[0]) return false;
    this.#writeSlide(entry.part, xml.slice(0, m.index) + next + xml.slice(m.index + m[0].length));
    return true;
  }

  // ---- transitions --------------------------------------------------------

  /** The slide's transition, as `readTransition` gives it, or null. */
  transition(index) {
    const entry = this.slideParts[index];
    if (!entry) throw new RangeError(`no slide at index ${index}`);
    return readTransition(this.pkg.text(entry.part));
  }

  /**
   * Transitions → Transition to This Slide, Effect Options, Duration and
   * Advance Slide: `spec` merged over the slide's own transition (see
   * `withTransition`), or null to take it off.
   * @returns {boolean} true when the slide's XML changed
   */
  setTransition(index, spec) {
    const entry = this.slideParts[index];
    if (!entry) throw new RangeError(`no slide at index ${index}`);
    const xml = this.pkg.text(entry.part);
    const next = withTransition(xml, spec);
    if (next === xml) return false;
    this.#writeSlide(entry.part, next);
    return true;
  }

  /**
   * Transitions → Apply To All: this slide's transition — effect, duration
   * and advance settings, exactly as written — on every slide; a slide with
   * none takes every other slide's off.
   * @returns {number} how many slides changed
   */
  applyTransitionToAll(from) {
    const entry = this.slideParts[from];
    if (!entry) throw new RangeError(`no slide at index ${from}`);
    const block = transitionBlock(this.pkg.text(entry.part));
    let changed = 0;
    this.slideParts.forEach((s, i) => {
      if (i === from) return;
      const xml = this.pkg.text(s.part);
      const bare = withTransition(xml, null);
      const next = block ? insertTransition(bare, block) : bare;
      if (next === xml) return;
      this.#writeSlide(s.part, next);
      changed += 1;
    });
    return changed;
  }

  // ---- animations ---------------------------------------------------------

  #editTiming(index, edit) {
    const entry = this.slideParts[index];
    if (!entry) throw new RangeError(`no slide at index ${index}`);
    const xml = this.pkg.text(entry.part);
    const out = edit(xml);
    const next = typeof out === 'string' ? out : out.xml;
    if (next !== xml) this.#writeSlide(entry.part, next);
    return out;
  }

  /** The slide's main sequence, as `readAnimations` gives it. */
  animations(index) {
    const entry = this.slideParts[index];
    if (!entry) throw new RangeError(`no slide at index ${index}`);
    return readAnimations(this.pkg.text(entry.part));
  }

  /**
   * Animations → the gallery (on a shape with none) and Add Animation: an
   * effect on a shape, at the end of the sequence unless `at` says where.
   * @returns {number} its index in the sequence
   */
  addAnimation(index, shapeId, spec = {}, at = null) {
    return this.#editTiming(index, (xml) => addAnimation(xml, shapeId, spec, at)).index;
  }

  /** Effect Options, Start, Duration, Delay, or another effect from the gallery, for the effect at `at`. */
  setAnimation(index, at, patch = {}) {
    this.#editTiming(index, (xml) => setAnimation(xml, at, patch));
    return true;
  }

  /** The effect at `at` taken out of the sequence. */
  removeAnimation(index, at) {
    this.#editTiming(index, (xml) => removeAnimation(xml, at));
    return true;
  }

  /**
   * Move Earlier / Move Later: the effect at `at` one place either way, or
   * to a position.
   * @returns {number} where it is now
   */
  moveAnimation(index, at, to) {
    return this.#editTiming(index, (xml) => moveAnimation(xml, at, to)).index;
  }

  /** Every effect on a shape taken out — the gallery's None. */
  removeShapeAnimations(index, shapeId) {
    const entry = this.slideParts[index];
    if (!entry) throw new RangeError(`no slide at index ${index}`);
    const xml = this.pkg.text(entry.part);
    const next = removeShapeAnimations(xml, shapeId);
    if (next === xml) return false;
    this.#writeSlide(entry.part, next);
    return true;
  }

  /** After a shape goes: the effects that pointed at it go too, as PowerPoint drops them. */
  #pruneTiming(part) {
    const xml = this.pkg.text(part);
    if (!xml.includes('<p:timing')) return;
    let next = xml;
    try { next = pruneAnimations(xml); } catch { /* a sequence this cannot rewrite is left alone */ }
    if (next !== xml) this.#writeSlide(part, next);
  }

  save() {
    return this.pkg.write();
  }
}

/**
 * The keys a placeholder is looked up by on its layout and master: its
 * index first, then its type, then the types PowerPoint lets stand in for
 * it — a slide's body can fill a layout's subtitle, a centred title a title.
 */
function placeholderKeys(ph) {
  const keys = [];
  if (ph.idx != null) keys.push(`idx:${ph.idx}`);
  if (ph.type) keys.push(`type:${ph.type}`);
  // A slide's "body" placeholder can match a layout's "subTitle" or "ctrTitle".
  if (ph.type === 'body') keys.push('type:subTitle', 'type:ctrTitle', 'type:title');
  if (ph.type === 'ctrTitle' || ph.type === 'subTitle') keys.push('type:title', 'type:body');
  if (ph.type === 'title') keys.push('type:ctrTitle');
  return keys;
}

/** Whether a text body holds any words. */
function hasWords(body) {
  return Boolean(body?.paragraphs?.some((p) => (p.runs || []).some((r) => r.text && r.text.trim())));
}

/** The words Slide Master view shows in a placeholder that holds none, PowerPoint's own. */
const PROMPTS = {
  title: 'Click to edit Master title style',
  ctrTitle: 'Click to edit Master title style',
  subTitle: 'Click to edit Master subtitle style',
  dt: 'Date',
  ftr: 'Footer',
  sldNum: '‹#›',
  pic: 'Picture',
  chart: 'Chart',
  tbl: 'Table',
  media: 'Media',
};
const BODY_PROMPTS = ['Click to edit Master text styles', 'Second level', 'Third level', 'Fourth level', 'Fifth level'];

/** A placeholder's prompt as a text body, keeping the body's own anchor, insets and list style. */
function promptBody(type, isLayout, body) {
  const base = body ? { ...body } : { anchor: 'top', vert: 'horz', columns: 1, wrap: true, autofit: false, insets: { l: 7.2, t: 3.6, r: 7.2, b: 3.6 }, levels: null };
  const words = PROMPTS[type];
  const align = body?.paragraphs?.[0]?.align;
  const paragraphs = words
    ? [{ ...(align ? { align } : {}), runs: [{ text: words }] }]
    : BODY_PROMPTS.map((text, level) => ({ ...(level ? { level } : {}), runs: [{ text }] }));
  return { ...base, paragraphs };
}

/** An element's top-level children, each as its own XML string. */
function topElements(inner) {
  const out = [];
  const re = /<(\/?)([A-Za-z0-9_:]+)\b[^>]*?(\/?)>/g;
  let depth = 0;
  let start = -1;
  let m;
  while ((m = re.exec(inner))) {
    if (m[1]) {
      depth -= 1;
      if (depth === 0 && start >= 0) { out.push(inner.slice(start, m.index + m[0].length)); start = -1; }
    } else if (m[3]) {
      if (depth === 0) out.push(m[0]);
    } else {
      if (depth === 0) start = m.index;
      depth += 1;
    }
  }
  return out;
}

/** Where a run property's children go, in the schema's order (CT_TextCharacterProperties). */
const RPR_ORDER = ['a:ln', 'a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill', 'a:effectLst', 'a:effectDag', 'a:highlight', 'a:uLnTx', 'a:uLn', 'a:uFillTx', 'a:uFill', 'a:latin', 'a:ea', 'a:cs', 'a:sym', 'a:hlinkClick', 'a:hlinkMouseOver', 'a:rtl', 'a:extLst'];
/** A list level's children, in the schema's order (CT_TextParagraphProperties). */
const PPR_ORDER = ['a:lnSpc', 'a:spcBef', 'a:spcAft', 'a:buClrTx', 'a:buClr', 'a:buSzTx', 'a:buSzPct', 'a:buSzPts', 'a:buFontTx', 'a:buFont', 'a:buNone', 'a:buAutoNum', 'a:buChar', 'a:buBlip', 'a:tabLst', 'a:defRPr', 'a:extLst'];
const nameOfEl = (el) => /^<([A-Za-z0-9_:]+)/.exec(el)?.[1] || '';
const sortBy = (order, els) => [...els].sort((a, b) => order.indexOf(nameOfEl(a)) - order.indexOf(nameOfEl(b)));

/**
 * A list style (`a:lstStyle`, `p:titleStyle`, `p:bodyStyle`) with the
 * given look laid on its levels — the first `levels` of them, created
 * where missing — each level's `a:defRPr` and alignment rewritten, every
 * other thing it says kept, children in the schema's order.
 */
function editListStyle(xml, tag, props, levels) {
  const open = new RegExp(`^<${tag}\\b[^>]*>`).exec(xml)[0];
  const inner = xml.slice(open.length, xml.length - `</${tag}>`.length);
  const els = topElements(inner);
  const byLevel = new Map();
  for (const el of els) {
    const m = /^<a:lvl(\d)pPr\b/.exec(el);
    if (m) byLevel.set(Number(m[1]), el);
  }
  const firstSize = (() => {
    const l1 = byLevel.get(1);
    const sz = l1 && /<a:defRPr\b[^>]*\bsz="(\d+)"/.exec(l1)?.[1];
    return sz ? Number(sz) : null;
  })();
  const ratio = props.size != null && firstSize ? (Number(props.size) * 100) / firstSize : null;
  const colour = props.color ? String(props.color).replace('#', '').toUpperCase() : null;
  const setAttr = (openTag, name, value) => {
    const bare = openTag.replace(new RegExp(`\\s+${name}="[^"]*"`), '');
    return value == null ? bare : bare.replace(/^(<[A-Za-z0-9_:]+)/, `$1 ${name}="${value}"`);
  };
  const editLevel = (n, el) => {
    let lvl = el || `<a:lvl${n}pPr></a:lvl${n}pPr>`;
    if (/^<a:lvl\dpPr\b[^>]*\/>$/.test(lvl)) lvl = lvl.replace(/\/>$/, `></a:lvl${n}pPr>`);
    const lvlOpen = /^<a:lvl\dpPr\b[^>]*>/.exec(lvl)[0];
    let head = lvlOpen;
    if (props.align) head = setAttr(head, 'algn', { left: 'l', center: 'ctr', right: 'r', justify: 'just' }[props.align] || 'l');
    const kidsOf = topElements(lvl.slice(lvlOpen.length, lvl.length - `</a:lvl${n}pPr>`.length));
    let def = kidsOf.find((k) => nameOfEl(k) === 'a:defRPr') || '<a:defRPr/>';
    const rest = kidsOf.filter((k) => nameOfEl(k) !== 'a:defRPr');
    let defOpen = /^<a:defRPr\b[^>]*?\/?>/.exec(def)[0];
    const defKids = defOpen.endsWith('/>') ? [] : topElements(def.slice(defOpen.length, def.length - '</a:defRPr>'.length));
    defOpen = defOpen.replace(/\/>$/, '>');
    if (props.size != null) {
      const own = Number(/\bsz="(\d+)"/.exec(defOpen)?.[1] || 0);
      const sz = n === 1 || !own || !ratio ? (n === 1 ? Math.round(Number(props.size) * 100) : null) : Math.round(own * ratio);
      if (sz) defOpen = setAttr(defOpen, 'sz', sz);
    }
    for (const [key, attr, on] of [['bold', 'b', '1'], ['italic', 'i', '1'], ['underline', 'u', 'sng']]) {
      if (props[key] === undefined) continue;
      defOpen = setAttr(defOpen, attr, props[key] ? on : attr === 'u' ? 'none' : '0');
    }
    let children = defKids;
    if (colour) children = [...children.filter((k) => !/^<a:(noFill|solidFill|gradFill|blipFill|pattFill|grpFill)\b/.test(k)), `<a:solidFill><a:srgbClr val="${colour}"/></a:solidFill>`];
    if (props.font !== undefined) {
      children = children.filter((k) => nameOfEl(k) !== 'a:latin');
      if (props.font) children.push(`<a:latin typeface="${escapeXml(props.font)}"/>`);
    }
    children = sortBy(RPR_ORDER, children);
    def = children.length ? `${defOpen}${children.join('')}</a:defRPr>` : defOpen.replace(/>$/, '/>');
    return `${head}${sortBy(PPR_ORDER, [...rest, def]).join('')}</a:lvl${n}pPr>`;
  };
  const kept = [];
  const done = new Set();
  for (const el of els) {
    const m = /^<a:lvl(\d)pPr\b/.exec(el);
    if (!m) { kept.push(el); continue; }
    const n = Number(m[1]);
    kept.push(n <= levels || byLevel.has(n) ? editLevel(n, el) : el);
    done.add(n);
  }
  for (let n = 1; n <= levels; n++) if (!done.has(n)) kept.push(editLevel(n, null));
  // defPPr first, then the levels in order, then extLst.
  const rank = (el) => (nameOfEl(el) === 'a:defPPr' ? 0 : nameOfEl(el) === 'a:extLst' ? 99 : Number(/^<a:lvl(\d)pPr/.exec(el)?.[1] || 50));
  kept.sort((a, b) => rank(a) - rank(b));
  return `${open}${kept.join('')}</${tag}>`;
}

/** A slide's animations, or none when its timing is past reading — the slide still opens. */
function safeAnimations(xml) {
  try {
    return readAnimations(xml);
  } catch {
    return [];
  }
}

const unescapeXml = (s) =>
  String(s)
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n)))
    .replace(/&amp;/g, '&');

const P14_NS = 'http://schemas.microsoft.com/office/powerpoint/2010/main';

/** A GUID in braces, upper case, the way PowerPoint names a section. */
function sectionGuid() {
  const hex = '0123456789ABCDEF';
  let out = '';
  for (let i = 0; i < 32; i++) {
    let n = Math.floor(Math.random() * 16);
    if (i === 12) n = 4;
    if (i === 16) n = 8 + (n & 3);
    out += hex[n];
    if (i === 7 || i === 11 || i === 15 || i === 19) out += '-';
  }
  return `{${out}}`;
}

/**
 * The sections of a presentation part: PowerPoint's p14:sectionLst, each
 * section its name, its GUID and the ids of the slides it lists. A deck with
 * no list has no sections.
 */
function readSectionList(xml) {
  const list = /<p14:sectionLst\b[^>]*>([\s\S]*?)<\/p14:sectionLst>/.exec(String(xml));
  if (!list) return [];
  const out = [];
  for (const m of list[1].matchAll(/<p14:section\b([^>]*?)(?:\/>|>([\s\S]*?)<\/p14:section>)/g)) {
    const name = /\bname="([^"]*)"/.exec(m[1])?.[1];
    const id = /\bid="([^"]*)"/.exec(m[1])?.[1];
    const slideIds = [...String(m[2] || '').matchAll(/<p14:sldId\b[^>]*\bid="(\d+)"/g)].map((s) => s[1]);
    out.push({ name: name != null ? unescapeXml(name) : 'Untitled Section', id: id || sectionGuid(), slideIds });
  }
  return out;
}

/** Every run of text in a fragment, as one line. */
function plainTextOf(xml) {
  return [...String(xml).matchAll(/<a:t(?:\s[^>]*)?>([^<]*)<\/a:t>/g)]
    .map((m) => unescapeXml(m[1]))
    .join(' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/** Whether a slide part's root tag marks it left out of the show. */
function slideHiddenFrom(xml) {
  return /<p:sld\b[^>]*\sshow="0"/.test(xml);
}

/** The text of a slide's title placeholder, straight from its XML. */
function outlineTitle(xml) {
  for (const m of String(xml).matchAll(/<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g)) {
    if (!/<p:ph\b[^>]*\btype="(?:title|ctrTitle)"/.test(m[0])) continue;
    const text = plainTextOf(m[0]);
    if (text) return text;
  }
  return '';
}

/** What PowerPoint calls each preset, for the shape's name. */

const PRESET_NAMES = {
  rect: 'Rectangle', roundRect: 'Rectangle: Rounded Corners', ellipse: 'Oval', triangle: 'Isosceles Triangle',
  rtTriangle: 'Right Triangle', diamond: 'Diamond', parallelogram: 'Parallelogram', trapezoid: 'Trapezoid',
  pentagon: 'Pentagon', hexagon: 'Hexagon', octagon: 'Octagon', star5: 'Star: 5 Points',
  rightArrow: 'Arrow: Right', chevron: 'Chevron', line: 'Straight Connector',
};

/**
 * A colour as DrawingML writes it: a hex string, or a spec — `{ scheme }` or
 * `{ color }` (a hex string of its own, for a spec that carries transforms a
 * bare string could not) — with an optional luminance modifier, offset and
 * alpha (0–1; only written under 1, since a colour with no transparency of
 * its own writes none). The transforms are legal on a scheme colour or a
 * literal one alike, which is exactly how PowerPoint writes a gradient
 * stop's "80% transparent, half as light" without naming a whole new colour.
 */
function colourXml(c) {
  if (typeof c === 'string') return `<a:srgbClr val="${escapeXml(c.replace('#', '').toUpperCase())}"/>`;
  const mods = [];
  if (c.lumMod != null) mods.push(`<a:lumMod val="${Math.round(c.lumMod * 1000)}"/>`);
  if (c.lumOff != null) mods.push(`<a:lumOff val="${Math.round(c.lumOff * 1000)}"/>`);
  if (c.alpha != null && c.alpha < 1) mods.push(`<a:alpha val="${Math.round(Math.max(0, Math.min(1, c.alpha)) * 100000)}"/>`);
  if (c.scheme) return mods.length ? `<a:schemeClr val="${escapeXml(c.scheme)}">${mods.join('')}</a:schemeClr>` : `<a:schemeClr val="${escapeXml(c.scheme)}"/>`;
  const hex = escapeXml(String(c.color ?? c.hex ?? '#000000').replace('#', '').toUpperCase());
  return mods.length ? `<a:srgbClr val="${hex}">${mods.join('')}</a:srgbClr>` : `<a:srgbClr val="${hex}"/>`;
}

/**
 * A two-stop gradient in one colour's own family — PowerPoint's "Light
 * Variation" and "Dark Variation" swatches in the gradient gallery: the
 * chosen colour at full strength fading to a paler (light) or a deeper
 * (dark) shade of itself, so recolouring the shape later is one colour
 * change, not a pair of unrelated ones.
 */
function gradientPresetStops(preset, colour) {
  const spec = typeof colour === 'string' ? { color: colour } : { ...colour };
  const at = (mods) => ({ ...spec, ...mods });
  if (preset === 'dark') return [{ pos: 0, ...at({ lumMod: 100 }) }, { pos: 1, ...at({ lumMod: 50 }) }];
  return [{ pos: 0, ...at({ lumMod: 100 }) }, { pos: 1, ...at({ lumMod: 20, lumOff: 80 }) }];
}

/**
 * `fill.gradient` written as `<a:gradFill>`: `stops` (two or three,
 * `{ pos: 0–1, ...colour }`, `colour` anything `colourXml` reads) if given,
 * else one of the light/dark presets in `color`; `angle` in degrees, 90
 * (top to bottom) unless stated — PowerPoint's own default for a fill
 * applied from the gallery.
 */
function gradientFillXml(g) {
  const stops = g.stops && g.stops.length ? g.stops : gradientPresetStops(g.preset === 'dark' ? 'dark' : 'light', g.color ?? { scheme: 'accent1' });
  const gsLst = stops.map((s) => `<a:gs pos="${Math.round(Math.max(0, Math.min(1, s.pos)) * 100000)}">${colourXml(s)}</a:gs>`).join('');
  const angle = g.angle != null ? g.angle : 90;
  return `<a:gradFill rotWithShape="1"><a:gsLst>${gsLst}</a:gsLst><a:lin ang="${Math.round(((angle % 360) + 360) % 360 * 60000)}" scaled="1"/></a:gradFill>`;
}

/** An outer shadow's own element — points and degrees in, DrawingML's EMU and 60,000ths out. */
function outerShadowXml(sh) {
  return `<a:outerShdw blurRad="${Math.round((sh.blur ?? 4) * 12700)}" dist="${Math.round((sh.dist ?? 3) * 12700)}" dir="${Math.round((sh.dir ?? 45) * 60000)}" algn="ctr" rotWithShape="0">`
    + `${colourXml({ color: sh.color || '#000000', alpha: sh.alpha ?? 0.4 })}</a:outerShdw>`;
}

/** Shape Effects → Glow: a radius in points (5/8/11/18 pt is PowerPoint's own gallery) and a colour, which usually carries its own transparency. */
function glowXml(g) {
  const base = typeof g.color === 'string' ? { color: g.color } : g.color || { scheme: 'accent1' };
  const spec = { ...base, alpha: g.alpha ?? base.alpha ?? 0.6 };
  return `<a:glow rad="${Math.round((g.radius ?? 8) * 12700)}">${colourXml(spec)}</a:glow>`;
}

/** Shape Effects → Soft Edges: a radius in points, the blurred band's own width. */
function softEdgeXml(s) {
  return `<a:softEdge rad="${Math.round((s.radius ?? 2.5) * 12700)}"/>`;
}

/**
 * Shape Effects → Reflection: PowerPoint's three gallery presets by name —
 * "tight" (touching, fading fast), "half" (a gap, fading by the middle) and
 * "full" (touching, fading only at the very end) — each a mirrored, faded
 * copy read straight off the shape itself, the way DrawingML always draws a
 * reflection rather than storing a second picture of one. The numbers
 * themselves live in `slide.js`'s `REFLECTION_PRESETS`, so what this writes
 * is exactly what `reflectionKindOf` reads back.
 */
function reflectionXml(kind) {
  const p = REFLECTION_PRESETS[kind] || REFLECTION_PRESETS.half;
  return `<a:reflection blurRad="${p.blurRad}" stA="${p.stA}" stPos="${p.stPos}" endA="${p.endA}" endPos="${p.endPos}" dist="${p.dist}" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/>`;
}

/** A background colour spec — a hex string, `{ colour }` or `{ scheme, lumMod?, lumOff? }` — through `colourXml`. */
function bgColourXml(c) {
  if (typeof c === 'string') return colourXml(c);
  if (c && c.colour) return colourXml(c.colour);
  if (c && c.scheme) return colourXml(c);
  throw new Error('a background colour needs a colour or a scheme');
}

/** A background spec, as `setBackground` writes the fill inside `<p:bgPr>`. */
function backgroundXml(spec) {
  const fillXml = spec.gradient
    ? `<a:gradFill><a:gsLst><a:gs pos="0">${bgColourXml(spec.gradient.from)}</a:gs>` +
      `<a:gs pos="100000">${bgColourXml(spec.gradient.to)}</a:gs></a:gsLst>` +
      `<a:lin ang="${Math.round((spec.gradient.angle != null ? spec.gradient.angle : 90) * 60000)}"/></a:gradFill>`
    : `<a:solidFill>${bgColourXml(spec)}</a:solidFill>`;
  return `<p:bg><p:bgPr>${fillXml}<a:effectLst/></p:bgPr></p:bg>`;
}

/** A colour node (`a:srgbClr` or `a:schemeClr`) read back into a background colour spec. */
function colourSpecOf(node) {
  if (!node) return null;
  if (node.name === A('srgbClr')) return { colour: String(node.attrs.val || '').toUpperCase() };
  if (node.name === A('schemeClr')) {
    const out = { scheme: node.attrs.val };
    for (const mod of kids(node)) {
      if (mod.name === A('lumMod')) out.lumMod = Number(mod.attrs.val) / 1000;
      if (mod.name === A('lumOff')) out.lumOff = Number(mod.attrs.val) / 1000;
    }
    return out;
  }
  return null;
}

/**
 * A slide's own `<p:bg>`, read back into what `setBackground` would need to
 * reproduce it. null when the slide has none of its own — the common case,
 * since most slides take their background from the layout or master.
 */
function readOwnBackground(xml) {
  const root = parse(xml);
  const sld = first(root, P('sld')) || root;
  const cSld = first(sld, P('cSld'));
  const bg = cSld && first(cSld, P('bg'));
  if (!bg) return null;
  const bgPr = first(bg, P('bgPr'));
  if (bgPr) {
    const solid = first(bgPr, A('solidFill'));
    if (solid) return colourSpecOf(kids(solid)[0]);
    const grad = first(bgPr, A('gradFill'));
    if (grad) {
      const gsLst = first(grad, A('gsLst'));
      const stops = kids(gsLst || { children: [] }, A('gs'));
      const at = (pos) => {
        const gs = stops.find((s) => s.attrs.pos === pos);
        return gs ? colourSpecOf(kids(gs)[0]) : null;
      };
      const lin = first(grad, A('lin'));
      return { gradient: { from: at('0'), to: at('100000'), angle: lin?.attrs.ang != null ? Number(lin.attrs.ang) / 60000 : 90 } };
    }
    return null;
  }
  const ref = first(bg, P('bgRef'));
  return ref ? colourSpecOf(kids(ref)[0]) : null;
}

/** A run's link, the relationship id the reader gives, resolved to { id, url } through the slide's relationships. */
function withLinks(shapes, rel) {
  if (typeof rel !== 'function') return shapes;
  return shapes.map((s) => {
    const paragraphs = s.text?.paragraphs;
    if (!paragraphs || !paragraphs.some((p) => (p.runs || []).some((r) => typeof r.link === 'string'))) return s;
    return {
      ...s,
      text: {
        ...s.text,
        paragraphs: paragraphs.map((p) => ({
          ...p,
          runs: (p.runs || []).map((r) => {
            if (typeof r.link !== 'string') return r;
            const target = rel(r.link);
            return { ...r, link: { id: r.link, url: target?.external ? target.target : target?.target || null } };
          }),
        })),
      },
    };
  });
}

/** The slide-number fields say the slide's own number, wherever the slide has moved to. */
function withSlideNumber(shapes, number) {
  const text = String(number);
  return shapes.map((s) => {
    const paragraphs = s.text?.paragraphs;
    if (!paragraphs || !paragraphs.some((p) => (p.runs || []).some((r) => r.field === 'slidenum' && r.text !== text))) return s;
    return {
      ...s,
      text: {
        ...s.text,
        paragraphs: paragraphs.map((p) => ({
          ...p,
          runs: (p.runs || []).map((r) => (r.field === 'slidenum' ? { ...r, text } : r)),
          plain: (p.runs || []).map((r) => (r.field === 'slidenum' ? text : r.text)).join(''),
        })),
      },
    };
  });
}

/**
 * A shape's own axis-aligned footprint on the slide — its unrotated box when
 * it has no rotation, otherwise the box that exactly contains it turned by
 * `rot` about its own centre. This is the box Align and Distribute work
 * from, since PowerPoint lines shapes up by how they look, not by the
 * unrotated numbers underneath a turned one.
 */
function shapeAabb(g) {
  if (!g.rot) return { x: g.x, y: g.y, w: g.w, h: g.h };
  const cx = g.x + g.w / 2;
  const cy = g.y + g.h / 2;
  const rad = (g.rot * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const corners = [
    [g.x, g.y], [g.x + g.w, g.y], [g.x, g.y + g.h], [g.x + g.w, g.y + g.h],
  ].map(([px, py]) => {
    const dx = px - cx;
    const dy = py - cy;
    return [cx + dx * cos - dy * sin, cy + dx * sin + dy * cos];
  });
  const xs = corners.map((c) => c[0]);
  const ys = corners.map((c) => c[1]);
  const x = Math.min(...xs);
  const y = Math.min(...ys);
  return { x, y, w: Math.max(...xs) - x, h: Math.max(...ys) - y };
}

/** The smallest box that holds every one of a list of boxes. */
function unionBox(boxes) {
  const x = Math.min(...boxes.map((b) => b.x));
  const y = Math.min(...boxes.map((b) => b.y));
  const r = Math.max(...boxes.map((b) => b.x + b.w));
  const b = Math.max(...boxes.map((b) => b.y + b.h));
  return { x, y, w: r - x, h: b - y };
}

function nextShapeId(xml) {

  const ids = [...xml.matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"/g)].map((m) => Number(m[1]));
  return (ids.length ? Math.max(...ids) : 1) + 1;
}

/**
 * Build a `<p:txBody>`, reusing the body properties the shape already had so an
 * edit does not silently reset autofit, insets or vertical anchoring.
 */
function buildTextBody(paragraphs, shapeXml, bodyStart, bodyEnd) {
  let bodyPr = '<a:bodyPr/>';
  let lstStyle = '<a:lstStyle/>';
  if (bodyStart >= 0) {
    const inner = shapeXml.slice(bodyStart, bodyEnd);
    const bp = /<a:bodyPr\b[^>]*(\/>|>[\s\S]*?<\/a:bodyPr>)/.exec(inner);
    if (bp) bodyPr = bp[0];
    const ls = /<a:lstStyle\b[^>]*(\/>|>[\s\S]*?<\/a:lstStyle>)/.exec(inner);
    if (ls) lstStyle = ls[0];
  }
  const ps = (paragraphs || []).map((p) => {
    const pPrBits = [];
    if (p.indent != null) pPrBits.push(`marL="${pxToEmu(p.indent)}"`);
    if (p.level) pPrBits.push(`lvl="${p.level}"`);
    if (p.hanging != null) pPrBits.push(`indent="${pxToEmu(p.hanging)}"`);
    if (p.align) pPrBits.push(`algn="${{ left: 'l', center: 'ctr', right: 'r', justify: 'just' }[p.align] || 'l'}"`);
    // The paragraph's own spacing and bullet, in schema order (lnSpc, spcBef,
    // spcAft, buClr, then one of buNone / buAutoNum / buChar) — the shape the
    // reader gives back, so a format press keeps a bullet the slide already
    // had instead of silently dropping it.
    const kids = [];
    if (p.lineHeight) kids.push(`<a:lnSpc><a:spcPct val="${Math.round(Number(p.lineHeight) * 100000)}"/></a:lnSpc>`);
    else if (p.lineHeightPt) kids.push(`<a:lnSpc><a:spcPts val="${Math.round(Number(p.lineHeightPt) * 100)}"/></a:lnSpc>`);
    if (p.spaceBefore != null) kids.push(`<a:spcBef><a:spcPts val="${Math.round(Number(p.spaceBefore) * 100)}"/></a:spcBef>`);
    if (p.spaceAfter != null) kids.push(`<a:spcAft><a:spcPts val="${Math.round(Number(p.spaceAfter) * 100)}"/></a:spcAft>`);
    const bullet = p.bullet === false ? { type: 'none' } : p.bullet;
    if (bullet && bullet.type) {
      if (bullet.color) kids.push(`<a:buClr><a:srgbClr val="${String(bullet.color).replace('#', '')}"/></a:buClr>`);
      if (bullet.type === 'none') kids.push('<a:buNone/>');
      else if (bullet.type === 'number') kids.push(`<a:buAutoNum type="${escapeXml(bullet.scheme || 'arabicPeriod')}"${bullet.start && Number(bullet.start) !== 1 ? ` startAt="${Number(bullet.start)}"` : ''}/>`);
      else kids.push(`<a:buChar char="${escapeXml(bullet.char || '•')}"/>`);
    }
    const pPr = kids.length
      ? `<a:pPr${pPrBits.length ? ' ' + pPrBits.join(' ') : ''}>${kids.join('')}</a:pPr>`
      : pPrBits.length ? `<a:pPr ${pPrBits.join(' ')}/>` : '';
    const runs = (p.runs || [])
      .filter((r) => r.text != null)
      .map((r) => {
        if (r.text === '\n') return '<a:br/>';
        const bits = ['lang="en-US"'];
        if (r.size) bits.push(`sz="${ptToSz(r.size)}"`);
        if (r.bold != null) bits.push(`b="${r.bold ? 1 : 0}"`);
        if (r.italic != null) bits.push(`i="${r.italic ? 1 : 0}"`);
        if (r.underline) bits.push('u="sng"');
        // The rest of what the reader reads, so an edit keeps it: strike,
        // character spacing (points, spc is hundredths), caps, super/sub.
        if (r.strike) bits.push('strike="sngStrike"');
        if (r.spacing) bits.push(`spc="${Math.round(Number(r.spacing) * 100)}"`);
        if (r.caps) bits.push(`cap="${r.caps === 'small' ? 'small' : 'all'}"`);
        if (r.baseline) bits.push(`baseline="${r.baseline === 'super' ? 30000 : -25000}"`);
        const fill = r.color ? `<a:solidFill><a:srgbClr val="${String(r.color).replace('#', '')}"/></a:solidFill>` : '';
        // A highlight sits after the fill and before the font in rPr's order.
        const highlight = r.highlight ? `<a:highlight><a:srgbClr val="${String(r.highlight).replace('#', '')}"/></a:highlight>` : '';
        const font = r.font ? `<a:latin typeface="${escapeXml(r.font)}"/>` : '';
        // A link is the run's relationship id — as the reader gives it, or
        // inside the { id, url } the scene resolves it to — last in rPr.
        const rId = typeof r.link === 'string' ? r.link : r.link?.id;
        const link = rId && /^rId\d+$/.test(rId) ? `<a:hlinkClick r:id="${rId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>` : '';
        const rPr = fill || highlight || font || link
          ? `<a:rPr ${bits.join(' ')}>${fill}${highlight}${font}${link}</a:rPr>`
          : `<a:rPr ${bits.join(' ')}/>`;
        // `xml:space` keeps leading and trailing spaces, which a title often has.
        return `<a:r>${rPr}<a:t xml:space="preserve">${escapeXml(r.text)}</a:t></a:r>`;
      })
      .join('');
    return `<a:p>${pPr}${runs}</a:p>`;
  });
  if (!ps.length) ps.push('<a:p/>');
  return `<p:txBody>${bodyPr}${lstStyle}${ps.join('')}</p:txBody>`;
}

/** PowerPoint's own "Medium Style 2 — Accent 1" — the table gallery's first pick, and what a table gets when nobody chooses otherwise. */
const TABLE_STYLE_ID = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';
/** A new row's height in EMU — PowerPoint's own default for a freshly inserted table. */
const TABLE_ROW_H = 370840;
const TR_RE = /<a:tr\b[^>]*>[\s\S]*?<\/a:tr>/g;
const TC_RE = /<a:tc\b[^>]*>[\s\S]*?<\/a:tc>/g;
const GRIDCOL_RE = /<a:gridCol\b[^>]*\/>|<a:gridCol\b[^>]*>[\s\S]*?<\/a:gridCol>/g;

/** One table cell: a run of words, or — empty — just enough to measure a line's height. */
function tableCellXml(text) {
  const body = text
    ? `<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>${escapeXml(String(text))}</a:t></a:r></a:p></a:txBody>`
    : `<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p></a:txBody>`;
  return `<a:tc>${body}<a:tcPr/></a:tc>`;
}

/** A graphic frame's own width, in EMU, off its `p:xfrm`. */
function frameExtCx(frameXml) {
  return Number(/<p:xfrm\b[^>]*>[\s\S]*?<a:ext\b[^>]*\bcx="(\d+)"/.exec(frameXml)?.[1]) || 0;
}

/** A graphic frame's `p:xfrm` height grown (or, negative, shrunk) by a row going in or out. */
function growFrameHeight(frameXml, deltaEmu) {
  return frameXml.replace(
    /(<p:xfrm\b[^>]*>[\s\S]*?<a:ext\b[^>]*\bcy=")(\d+)(")/,
    (m, pre, cy, post) => `${pre}${Math.max(1, Number(cy) + deltaEmu)}${post}`
  );
}

export { Theme, buildTextBody, resolveTarget };
