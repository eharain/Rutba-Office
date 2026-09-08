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
    const cacheKey = slideXml.length + ':' + notesXml.length + ':' + slideXml + notesXml;
    const cached = this._scenes.get(slidePart);
    if (cached && cached.key === cacheKey) return cached.scene;

    const scene = readSlideScene(slideXml, { theme, inherit, rel });
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
      shapes: scene.shapes,
      notes,
      theme: { colors: theme.colors, fonts: theme.fonts },
    };
    this._scenes.set(slidePart, { key: cacheKey, scene: result });
    return result;
  }

  /** Every slide, thumbnail-shaped: enough to draw a sorter without the parts. */
  outline() {
    return this.slideParts.map((entry, i) => {
      const scene = this.slide(i);
      const title = scene.shapes.find((s) => s.placeholder?.type === 'title' || s.placeholder?.type === 'ctrTitle');
      const body = title?.text || title?.inheritedText;
      return {
        index: i,
        part: entry.part,
        title: body ? body.paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join(' ').trim() : '',
        shapes: scene.shapes.length,
        notes: scene.notes,
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

  /** Append a text box, which is how the editor adds new content. */
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
    let layoutTarget = null;
    if (neighbour) {
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
    if (p.level) pPrBits.push(`lvl="${p.level}"`);
    if (p.align) pPrBits.push(`algn="${{ left: 'l', center: 'ctr', right: 'r', justify: 'just' }[p.align] || 'l'}"`);
    const pPr = pPrBits.length ? `<a:pPr ${pPrBits.join(' ')}/>` : '';
    const runs = (p.runs || [])
      .filter((r) => r.text != null)
      .map((r) => {
        if (r.text === '\n') return '<a:br/>';
        const bits = ['lang="en-US"'];
        if (r.size) bits.push(`sz="${ptToSz(r.size)}"`);
        if (r.bold != null) bits.push(`b="${r.bold ? 1 : 0}"`);
        if (r.italic != null) bits.push(`i="${r.italic ? 1 : 0}"`);
        if (r.underline) bits.push('u="sng"');
        const fill = r.color ? `<a:solidFill><a:srgbClr val="${String(r.color).replace('#', '')}"/></a:solidFill>` : '';
        const font = r.font ? `<a:latin typeface="${escapeXml(r.font)}"/>` : '';
        const rPr = fill || font
          ? `<a:rPr ${bits.join(' ')}>${fill}${font}</a:rPr>`
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
