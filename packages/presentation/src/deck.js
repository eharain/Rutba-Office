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
import { readSlideScene, readXfrm, readTextBody, placeholderOf, sceneText } from './slide.js';
import { slideXml } from './build.js';

const A = (n) => `a:${n}`;
const P = (n) => `p:${n}`;

const CT = {
  slide: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
  layout: 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
  master: 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
  notes: 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
};
const REL = {
  slide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
  layout: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
  master: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  notes: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
};/** The picture types PowerPoint itself embeds; anything else is converted first. */
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
function topLevelShapes(xml) {
  const start = xml.indexOf('<p:spTree');
  const end = xml.lastIndexOf('</p:spTree>');
  if (start < 0 || end < 0) return [];
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
    if (!xml) return;
    const root = parse(xml);
    const scheme = first(root, A('clrScheme'));
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
  }

  /** `schemeClr val="tx1"` → the theme colour, through the master's map. */
  color(name) {
    const mapped = this.clrMap[name] || name;
    return (
      this.colors[mapped] ||
      this.colors[name] ||
      { tx1: '#000000', tx2: '#44546a', bg1: '#ffffff', bg2: '#e7e6e6', dk1: '#000000', lt1: '#ffffff' }[mapped] ||
      null
    );
  }

  font(typeface) {
    if (typeface === '+mj-lt') return this.fonts.major;
    if (typeface === '+mn-lt') return this.fonts.minor;
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
    this.#load();
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
  }

  get slideCount() {
    return this.slideParts.length;
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

  /** Placeholder shapes of a layout or master, indexed for inheritance. */
  #placeholders(partName, theme) {
    if (!partName || !this.pkg.has(partName)) return new Map();
    const key = `${partName}`;
    if (this.layouts.has(key)) return this.layouts.get(key);
    const scene = readSlideScene(this.pkg.text(partName), { theme });
    const map = new Map();
    for (const s of scene.shapes) {
      if (!s.placeholder) continue;
      const { type, idx } = s.placeholder;
      if (idx != null) map.set(`idx:${idx}`, s);
      if (type) map.set(`type:${type}`, s);
    }
    map.set('#background', scene.background);
    this.layouts.set(key, map);
    return map;
  }

  /**
   * One slide, fully resolved: background, shapes, inherited geometry and
   * text style, and picture sources as part names the caller can read.
   */
  slide(index) {
    const entry = this.slideParts[index];
    if (!entry) throw new RangeError(`no slide at index ${index}`);
    const slidePart = entry.part;
    const layoutPart = this.#layoutFor(slidePart);
    const masterPart = layoutPart ? this.#masterFor(layoutPart) : null;
    const theme = this.#themeFor(masterPart);
    const layoutPh = this.#placeholders(layoutPart, theme);
    const masterPh = this.#placeholders(masterPart, theme);
    const rels = this.#relMap(slidePart);

    // Inheritance is a fall-through, not a lookup. A layout can name a
    // placeholder and still say nothing about where it goes, leaving the
    // position to the master — so each property is taken from the first
    // ancestor that actually states it, rather than from the first ancestor
    // that mentions the placeholder at all.
    const inherit = (ph) => {
      const keys = [];
      if (ph.idx != null) keys.push(`idx:${ph.idx}`);
      if (ph.type) keys.push(`type:${ph.type}`);
      // A slide's "body" placeholder can match a layout's "subTitle" or "ctrTitle".
      if (ph.type === 'body') keys.push('type:subTitle', 'type:ctrTitle', 'type:title');
      if (ph.type === 'ctrTitle' || ph.type === 'subTitle') keys.push('type:title', 'type:body');
      if (ph.type === 'title') keys.push('type:ctrTitle');

      const out = { geometry: null, fill: null, line: null, text: null };
      for (const source of [layoutPh, masterPh]) {
        for (const k of keys) {
          const hit = source.get(k);
          if (!hit) continue;
          out.geometry = out.geometry || hit.geometry;
          out.fill = out.fill || hit.fill;
          out.line = out.line || hit.line;
          out.text = out.text || hit.text;
        }
      }
      return out.geometry || out.fill || out.line || out.text ? out : null;
    };

    const rel = (id) => {
      const r = rels.get(id);
      if (!r) return null;
      return { part: r.resolved, external: r.mode === 'External', target: r.target };
    };

    // Parsed once per VERSION of the slide's XML (and its notes'): the
    // outline, the thumbnails and the model all ask for every slide, and a
    // fifteen-megabyte deck of nineteen slides parsed nineteen slides three
    // times per keystroke. The key is the XML itself, so a stale entry is
    // not expressible; an edit writes new XML and misses.
    const slideXml = this.pkg.text(slidePart);
    const notesPart = [...rels.values()].find((r) => r.type === REL.notes && this.pkg.has(r.resolved))?.resolved || null;
    const notesXml = notesPart ? this.pkg.text(notesPart) : '';
    const cacheKey = index + ':' + slideXml.length + ':' + notesXml.length + ':' + slideXml + notesXml;
    const cached = this._scenes.get(slidePart);
    if (cached && cached.key === cacheKey) return cached.scene;

    // `readPart` lets a chart frame read its chart part; only a slide gets
    // it, since a layout or master never carries a chart of its own.
    const readPart = (part) => (part && this.pkg.has(part) ? this.pkg.text(part) : null);
    const scene = readSlideScene(slideXml, { theme, inherit, rel, readPart });
    if (!scene.background) {
      scene.background = layoutPh.get('#background') || masterPh.get('#background') || null;
    }

    let notes = '';
    if (notesPart) notes = sceneText(readSlideScene(notesXml, { theme }));

    const result = {
      index,
      part: slidePart,
      layout: layoutPart,
      master: masterPart,
      size: this.size,
      background: scene.background,
      shapes: withLinks(withSlideNumber(scene.shapes, index + 1), rel),
      notes,
      theme: { colors: theme.colors, fonts: theme.fonts },
    };
    this._scenes.set(slidePart, { key: cacheKey, scene: result });
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
    const marker = new RegExp(`<p:cNvPr[^>]*\\bid="${shapeId}"`);
    const at = xml.search(marker);
    if (at < 0) return null;
    // Walk back to the opening tag of the containing shape.
    const openers = ['<p:sp>', '<p:pic>', '<p:graphicFrame>', '<p:cxnSp>'];
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
    const end = xml.indexOf(closeTag, at);
    if (end < 0) return null;
    return { start, end: end + closeTag.length, tag };
  }

  #writeSlide(part, xml) {
    this.pkg.write_(part, Buffer.from(xml, 'utf8'));
    this.dirty = true;
  }

  /** Replace a shape's text with paragraphs of runs. */
  setText(slideIndex, shapeId, paragraphs) {
    const part = this.slideParts[slideIndex]?.part;
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

  /** Move or resize a shape. Values are pixels; EMU conversion happens here. */
  setGeometry(slideIndex, shapeId, { x, y, w, h, rot }) {
    const part = this.slideParts[slideIndex]?.part;
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

    if (hasXfrm) {
      shapeXml = offRe.test(shapeXml) ? shapeXml.replace(offRe, offXml) : shapeXml.replace(/<a:xfrm\b[^>]*>/, (m) => m + offXml);
      shapeXml = extRe.test(shapeXml) ? shapeXml.replace(extRe, extXml) : shapeXml.replace(offXml, offXml + extXml);
      if (rot != null) {
        shapeXml = shapeXml.replace(/<a:xfrm\b([^>]*)>/, (m, attrs) => {
          const cleaned = attrs.replace(/\s*rot="[^"]*"/, '');
          return `<a:xfrm${cleaned}${rot ? ` rot="${Math.round(rot * 60000)}"` : ''}>`;
        });
      }
    } else {
      // The shape inherited its geometry; state it explicitly now that the user
      // has moved it, inserting the xfrm as the first child of spPr.
      const rotAttr = rot ? ` rot="${Math.round(rot * 60000)}"` : '';
      shapeXml = shapeXml.replace(
        /<p:spPr\s*\/>|<p:spPr\b[^>]*>/,
        (m) => (m.endsWith('/>')
          ? `<p:spPr><a:xfrm${rotAttr}>${offXml}${extXml}</a:xfrm></p:spPr>`
          : `${m}<a:xfrm${rotAttr}>${offXml}${extXml}</a:xfrm>`)
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
    const part = this.slideParts[slideIndex]?.part;
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
    const part = this.slideParts[slideIndex]?.part;
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
    const part = this.slideParts[slideIndex]?.part;
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
    const part = this.slideParts[slideIndex]?.part;
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
    const part = this.slideParts[slideIndex]?.part;
    if (!part) throw new RangeError(`no slide at index ${slideIndex}`);
    const xml = this.pkg.text(part);
    const range = this.#shapeRange(xml, shapeId);
    if (!range) return false;
    this.#writeSlide(part, xml.slice(0, range.start) + xml.slice(range.end));
    return true;
  }

  /**
   * Move a shape in the drawing order — bring forward, send backward, to the
   * front, to the back. The spTree draws its children in order, so the order
   * IS the layering: the shape's XML moves among its siblings, nothing else.
   */
  reorderShape(slideIndex, shapeId, to) {
    const part = this.slideParts[slideIndex]?.part;
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
   * A shape's fill and outline — the Format pane. `fill` is 'none', a
   * colour ('#RRGGBB' or { scheme, lumMod }) or null to leave it as it is;
   * `line` is 'none', { color, width (points), dash } or null. Written into
   * the shape's own spPr, where it beats the style reference the shape may
   * carry, in the order the schema wants: geometry, fill, line. A picture
   * takes an outline as a frame; a table or chart frame has no spPr and
   * says so.
   */
  setShapeStyle(slideIndex, shapeId, { fill = null, line = null, effects = null } = {}) {
    const part = this.slideParts[slideIndex]?.part;
    if (!part) throw new RangeError(`no slide at index ${slideIndex}`);
    const xml = this.pkg.text(part);
    const range = this.#shapeRange(xml, shapeId);
    if (!range) throw new Error(`shape ${shapeId} not found`);
    if (range.tag === '<p:graphicFrame>') throw new Error('A table or chart frame has no fill or outline of its own.');
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
    // The effects after both: `effects` is 'none' (an empty list, which turns a
    // style's shadow off too), { shadow: { dist, dir, blur, color, alpha } }
    // in points, degrees and a 0–1 alpha, or null to leave them as they are.
    const effectRe = /<a:effectLst\b[^>]*\/>|<a:effectLst\b[^>]*>[\s\S]*?<\/a:effectLst>/;
    const hadEffects = effectRe.exec(inner)?.[0] ?? '';
    inner = inner.replace(effectRe, '');
    const shadowXml = (sh) => `<a:effectLst><a:outerShdw blurRad="${Math.round((sh.blur ?? 4) * 12700)}" dist="${Math.round((sh.dist ?? 3) * 12700)}" dir="${Math.round((sh.dir ?? 45) * 60000)}" algn="ctr" rotWithShape="0">`
      + `<a:srgbClr val="${String(sh.color || '#000000').replace('#', '').toUpperCase()}"><a:alpha val="${Math.round(Math.min(1, Math.max(0, sh.alpha ?? 0.4)) * 100000)}"/></a:srgbClr></a:outerShdw></a:effectLst>`;
    const effectXml = effects === null ? hadEffects : effects === 'none' || !effects.shadow ? '<a:effectLst/>' : shadowXml(effects.shadow);
    const fillRe = /<a:(noFill|solidFill|gradFill|blipFill|pattFill|grpFill)\b[^>]*\/>|<a:(noFill|solidFill|gradFill|blipFill|pattFill|grpFill)\b[^>]*>[\s\S]*?<\/a:\2>/;
    const hadFill = fillRe.exec(inner)?.[0] ?? '';
    inner = inner.replace(fillRe, '');
    const fillXml = fill === null ? hadFill : fill === 'none' || fill?.type === 'none' ? '<a:noFill/>' : `<a:solidFill>${colourXml(fill)}</a:solidFill>`;
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
    const part = this.slideParts[slideIndex]?.part;
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
    const part = this.slideParts[slideIndex]?.part;
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
    const part = this.slideParts[slideIndex]?.part;
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
    const part = this.slideParts[slideIndex]?.part;
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
   * A link on a shape's words — Insert → Link: every run of the shape points
   * at an External relationship the slide carries, as PowerPoint writes one;
   * null takes the link off every run (the relationship stays, unreferenced,
   * as an undone edit leaves it). Returns the relationship id, or null.
   */
  setLink(slideIndex, shapeId, url) {
    const part = this.slideParts[slideIndex]?.part;
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

  addTextBox(slideIndex, { x, y, w, h, paragraphs, name = 'TextBox' }) {
    const part = this.slideParts[slideIndex]?.part;
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
  addPicture(slideIndex, { data, contentType, name = 'Picture', x = 0, y = 0, w, h }) {
    const part = this.slideParts[slideIndex]?.part;
    if (!part) throw new RangeError(`no slide at index ${slideIndex}`);
    const ext = IMAGE_EXTENSIONS[String(contentType || '').toLowerCase()];
    if (!ext) throw new Error(`unsupported picture type: ${contentType} (png, jpeg, gif or bmp)`);
    const bytes = Buffer.isBuffer(data) ? data : data instanceof Uint8Array ? Buffer.from(data) : Buffer.from(String(data), 'base64');
    if (!bytes.length) throw new Error('the picture has no bytes');
    if (!(w > 0) || !(h > 0)) throw new Error('a picture needs a positive width and height');

    // Numbered across every extension, as PowerPoint numbers them.
    const names = this.pkg.partNames() || [];
    let n = 1;
    while (names.some((p) => p.startsWith(`ppt/media/image${n}.`))) n += 1;
    const media = `ppt/media/image${n}.${ext}`;
    this.pkg.ensureDefault(ext, ext === 'jpeg' ? 'image/jpeg' : contentType);
    this.pkg.addPart(media, bytes);
    const rId = this.pkg.addRelationshipTo(part, REL.image, `../media/image${n}.${ext}`);

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
    const part = this.slideParts[slideIndex]?.part;
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
    return true;
  }

  save() {
    return this.pkg.write();
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

/** Every run of text in a fragment, as one line. */
function plainTextOf(xml) {
  return [...String(xml).matchAll(/<a:t(?:\s[^>]*)?>([^<]*)<\/a:t>/g)]
    .map((m) => unescapeXml(m[1]))
    .join(' ')
    .replace(/\s+/g, ' ')
    .trim();
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

/** A colour as DrawingML writes it: a hex string, or a scheme colour with an optional luminance modifier. */
function colourXml(c) {
  if (typeof c === 'string') return `<a:srgbClr val="${escapeXml(c.replace('#', '').toUpperCase())}"/>`;
  const mods = c.lumMod != null ? `<a:lumMod val="${Math.round(c.lumMod * 1000)}"/>` : '';
  return mods ? `<a:schemeClr val="${escapeXml(c.scheme)}">${mods}</a:schemeClr>` : `<a:schemeClr val="${escapeXml(c.scheme)}"/>`;
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

export { Theme, buildTextBody, resolveTarget };
