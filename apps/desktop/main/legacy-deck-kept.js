// A PowerPoint 97-2003 presentation that PowerPoint 2007 or later saved,
// rebuilt as they made it. Those keep, beside the 97-2003 description, the
// deck as they drew it: each master's theme, colour map and text styles,
// each layout whole, the table styles, and every drawing the older records
// cannot say all of — a shape's preset, effects and words in their fonts, a
// table, a background — as DrawingML. From those come the theme, the master
// with its placeholders, the layouts and each slide's drawings, each slide on
// the layout it was on. A drawing the older records say all of (a plain
// shape, a picture, a title's words) PowerPoint kept no DrawingML for: a
// placeholder's words are written here as a placeholder of its layout, in
// only the looks they have of their own, and the rest is handed back for the
// deck builder to draw from the older records as before, each in its place in
// the drawing order. So is a drawing an older PowerPoint changed after it was
// kept — its words or its place no longer those kept.

import zlib from 'node:zlib';
import { readZip } from '@rutba/ooxml/zip';
import { OoxmlPackage } from '@rutba/ooxml/package';

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const CT = 'application/vnd.openxmlformats-officedocument.';
const EMU = 9525; // to a slide pixel
const MEDIA = { png: 'image/png', jpeg: 'image/jpeg', jpg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', emf: 'image/x-emf', wmf: 'image/x-wmf', tif: 'image/tiff', tiff: 'image/tiff' };
const EXT = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/x-emf': 'emf', 'image/x-wmf': 'wmf', 'image/tiff': 'tiff' };
const GROUP = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const unesc = (s) => s.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (all, e) => (e[0] === '#'
  ? String.fromCodePoint(e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : Number(e.slice(1)))
  : { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }[e.toLowerCase()]));
const attrsOf = (tag) => Object.fromEntries([...tag.matchAll(/([\w:.-]+)="([^"]*)"/g)].map((m) => [m[1], m[2]]));

/** A kept package's parts by name, or null. */
function unzip(bytes) {
  if (!bytes) return null;
  try { return new Map(readZip(Buffer.from(bytes)).entries.map((e) => [e.name, Buffer.from(e.data)])); } catch { return null; }
}
/** A part's XML, without its declaration. */
const xmlOf = (parts, name) => {
  const b = parts?.get(name);
  return b ? b.toString('utf8').replace(/^﻿/, '').replace(/^<\?xml[^>]*\?>\s*/, '') : null;
};
/** A relationship's target, from the folder of the part that names it. */
function resolve(dir, target) {
  const out = dir ? dir.split('/') : [];
  for (const seg of target.split('/')) {
    if (seg === '..') out.pop();
    else if (seg && seg !== '.') out.push(seg);
  }
  return out.join('/');
}

/** One part of the deck being written, and the relationships it names. */
class Part {
  constructor(name) { this.name = name; this.rels = []; }
  rel(type, target, external = false) {
    const id = 'rId' + (this.rels.length + 1);
    this.rels.push({ id, type, target, external });
    return id;
  }
  relsXml() {
    return DECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + this.rels.map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${esc(r.target)}"${r.external ? ' TargetMode="External"' : ''}/>`).join('')
      + '</Relationships>';
  }
}

/**
 * The presentation as PowerPoint 2007 or later kept it in the .ppt — or null
 * when it kept too little to rebuild from. Returns the package and, for each
 * slide, the drawings left for the deck builder (each with its place in the
 * drawing order), the ids the kept ones have now by the ids the older records
 * give them (for their effects), and whether the slide's own background is
 * still to be set.
 */
export function keptDeck(model) {
  const kept = model.kept;
  const main = kept?.masters.find((m) => m.theme && !m.layout && m.items);
  if (!main) return null;
  const themeParts = unzip(main.theme);
  const theme = xmlOf(themeParts, 'theme/theme/theme1.xml');
  const textStyles = xmlOf(unzip(main.textStyles), 'drs/slideMasters/slideMaster1.xml');
  const colourMap = main.colourMap ? Buffer.from(main.colourMap).toString('utf8') : '';
  const map = /<a:clrMap\b([^>]*?)\/?>/.exec(colourMap);
  if (!theme || !textStyles || !map) return null;
  const layouts = kept.masters.filter((m) => m.layout).map((m) => ({ id: m.id, parts: unzip(m.layout) })).filter((l) => l.parts);
  if (!layouts.length) return null;

  const files = [];
  const media = new Map();
  /** A picture in the deck's media, once for each picture; its path from the ppt folder. */
  const mediaOf = (bytes, ext) => {
    const key = Buffer.from(bytes).toString('base64');
    if (!media.has(key)) media.set(key, { name: `media/kept${media.size + 1}.${ext}`, bytes: Buffer.from(bytes) });
    return media.get(key).name;
  };
  const ctx = { mediaOf, images: model.images || [] };

  // The theme.
  const themePart = new Part('ppt/theme/theme1.xml');
  const themeXml = adopt(theme, themeParts, 'theme/theme/_rels/theme1.xml.rels', 'theme/theme', themePart, ctx) ?? theme.replace(/\sr:embed="[^"]*"/g, '');
  files.push({ part: themePart, xml: themeXml });

  // The master: its background and drawings, its colour map, its layouts and its text styles.
  const masterPart = new Part('ppt/slideMasters/slideMaster1.xml');
  const masterSp = [];
  const behind = []; // drawings on the master with no kept DrawingML that are not placeholders: drawn on each slide, as before
  let nextId = 2;
  for (const item of main.items) {
    const sh = item.shapes[0];
    const xml = item.drawingML ? keptShape(item, masterPart, ctx) : null;
    if (xml) masterSp.push(withId(xml, nextId++));
    else if (item.shapes.length === 1 && sh.placeholder != null && sh.type === 'text') masterSp.push(placeholderXml(sh, nextId++, null, true));
    else behind.push(...item.shapes);
  }
  const masterBg = backgroundXml(main.backgroundML) || '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>';
  const layoutIds = layouts.map((l, i) => `<p:sldLayoutId id="${2147483649 + i}" r:id="${masterPart.rel(REL + 'slideLayout', `../slideLayouts/slideLayout${i + 1}.xml`)}"/>`);
  masterPart.rel(REL + 'theme', '../theme/theme1.xml');
  const txStyles = textStyles.replace(/^<p:txStyles\b[^>]*>/, '<p:txStyles>');
  files.push({
    part: masterPart,
    xml: `${DECL}<p:sldMaster ${NS}><p:cSld>${masterBg}<p:spTree>${GROUP}${masterSp.join('')}</p:spTree></p:cSld>`
      + `<p:clrMap${map[1].replace(/\s+xmlns(:\w+)?="[^"]*"/g, '')}/><p:sldLayoutIdLst>${layoutIds.join('')}</p:sldLayoutIdLst>${txStyles}</p:sldMaster>`,
  });

  // The layouts, whole.
  const layoutIndex = new Map();
  layouts.forEach((l, i) => {
    const part = new Part(`ppt/slideLayouts/slideLayout${i + 1}.xml`);
    part.rel(REL + 'slideMaster', '../slideMasters/slideMaster1.xml');
    const name = [...l.parts.keys()].find((n) => /^drs\/slideLayouts\/[^/]+\.xml$/.test(n));
    let xml = xmlOf(l.parts, name);
    if (!xml || !/^<p:sldLayout\b/.test(xml)) throw new Error('a layout kept in a form not known');
    xml = adopt(xml, l.parts, name.replace(/([^/]+)$/, '_rels/$1.rels'), 'drs/slideLayouts', part, ctx) ?? xml;
    files.push({ part, xml: DECL + xml });
    layoutIndex.set(l.id, i);
    // Its placeholders, which a slide's are matched to.
    l.placeholders = [...xml.matchAll(/<p:ph\b([^>]*?)\/?>/g)].map((m) => attrsOf(m[1]));
  });

  // The slides: each on its layout, its kept drawings as kept, its placeholders' words, and the rest left in their places.
  const plans = [];
  const slideNames = [];
  model.slides.forEach((slide, i) => {
    const part = new Part(`ppt/slides/slide${i + 1}.xml`);
    part.rel(REL + 'slideLayout', `../slideLayouts/slideLayout${(layoutIndex.get(slide.kept?.layout) ?? 0) + 1}.xml`);
    const plan = { pending: [], ids: new Map(), background: false };
    const sp = [];
    let id = 2;
    const layout = layouts[layoutIndex.get(slide.kept?.layout) ?? 0];
    const taken = new Set();
    // What the master draws that it kept no DrawingML for, under the slide's own, as before.
    if (slide.kept?.masterObjects) for (const sh of behind) plan.pending.push({ sh, at: sp.length + plan.pending.length });
    for (const item of slide.kept?.items || []) {
      const at = sp.length + plan.pending.length;
      const sh = item.shapes[0];
      const xml = item.drawingML ? keptShape(item, part, ctx, i + 1) : null;
      if (xml) {
        if (item.spid != null) plan.ids.set(item.spid, { id, animation: sh?.animation ?? null });
        sp.push(withId(xml, id++));
      } else if (item.shapes.length === 1 && sh.placeholder != null && sh.type === 'text' && sh.ownParagraphs) {
        if (item.spid != null) plan.ids.set(item.spid, { id, animation: sh.animation ?? null });
        sp.push(placeholderXml(sh, id++, matchPlaceholder(sh, layout?.placeholders || [], taken)));
      } else {
        item.shapes.forEach((s, k) => plan.pending.push({ sh: s, at: at + k }));
      }
    }
    // Pending drawings take their places among the written ones as they are added, so their places count both.
    let bg = backgroundXml(slide.kept?.backgroundML);
    if (!bg && slide.background && slide.background !== 'none' && slide.kept && !slide.kept.backgroundML && slide.kept.ownBackground) plan.background = true;
    files.push({ part, xml: `${DECL}<p:sld ${NS}><p:cSld>${bg || ''}<p:spTree>${GROUP}${sp.join('')}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>` });
    slideNames.push(part.name);
    plans.push(plan);
  });

  // The table styles PowerPoint kept, or none of the deck's own.
  const tableStyles = xmlOf(unzip(kept.tableStyles), 'tableStyles.xml')
    ?? '<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>';

  // The presentation: its master and slides at their size; its default text the master's other text.
  const size = model.size || { width: 960, height: 720 };
  const pres = new Part('ppt/presentation.xml');
  const masterRel = pres.rel(REL + 'slideMaster', 'slideMasters/slideMaster1.xml');
  const slideRels = slideNames.map((n) => pres.rel(REL + 'slide', n.replace('ppt/', '')));
  pres.rel(REL + 'theme', 'theme/theme1.xml');
  pres.rel(REL + 'tableStyles', 'tableStyles.xml');
  const other = /<p:otherStyle>([\s\S]*?)<\/p:otherStyle>/.exec(txStyles)?.[1] ?? '<a:defPPr><a:defRPr lang="en-US"/></a:defPPr>';
  const cx = Math.round(size.width * EMU);
  const cy = Math.round(size.height * EMU);
  const presXml = `${DECL}<p:presentation ${NS} saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="${masterRel}"/></p:sldMasterIdLst>`
    + `<p:sldIdLst>${slideRels.map((r, i) => `<p:sldId id="${256 + i}" r:id="${r}"/>`).join('')}</p:sldIdLst>`
    + `<p:sldSz cx="${cx}" cy="${cy}"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle>${other}</p:defaultTextStyle></p:presentation>`;

  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const parts = [];
  const overrides = [
    ['/ppt/presentation.xml', CT + 'presentationml.presentation.main+xml'],
    ['/ppt/slideMasters/slideMaster1.xml', CT + 'presentationml.slideMaster+xml'],
    ['/ppt/theme/theme1.xml', CT + 'theme+xml'],
    ['/ppt/tableStyles.xml', CT + 'presentationml.tableStyles+xml'],
    ['/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'],
    ['/docProps/app.xml', CT + 'extended-properties+xml'],
    ...layouts.map((_, i) => [`/ppt/slideLayouts/slideLayout${i + 1}.xml`, CT + 'presentationml.slideLayout+xml']),
    ...slideNames.map((n) => ['/' + n, CT + 'presentationml.slide+xml']),
  ];
  parts.push({
    name: '[Content_Types].xml',
    data: DECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      + Object.entries(MEDIA).map(([ext, type]) => `<Default Extension="${ext}" ContentType="${type}"/>`).join('')
      + overrides.map(([name, type]) => `<Override PartName="${name}" ContentType="${type}"/>`).join('') + '</Types>',
  });
  parts.push({ name: '_rels/.rels', data: DECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + REL + 'officeDocument" Target="ppt/presentation.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="' + REL + 'extended-properties" Target="docProps/app.xml"/></Relationships>' });
  parts.push({ name: 'docProps/core.xml', data: `${DECL}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Presentation</dc:title><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>` });
  parts.push({ name: 'docProps/app.xml', data: `${DECL}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Rutba Office</Application><Slides>${slideNames.length}</Slides></Properties>` });
  parts.push({ name: 'ppt/presentation.xml', data: presXml });
  parts.push({ name: 'ppt/_rels/presentation.xml.rels', data: pres.relsXml() });
  parts.push({ name: 'ppt/tableStyles.xml', data: DECL + tableStyles });
  for (const { part, xml } of files) {
    parts.push({ name: part.name, data: xml.startsWith('<?xml') ? xml : DECL + xml });
    if (part.rels.length) parts.push({ name: part.name.replace(/([^/]+)$/, '_rels/$1.rels'), data: part.relsXml() });
  }
  for (const m of media.values()) parts.push({ name: 'ppt/' + m.name, data: m.bytes });
  return { bytes: OoxmlPackage.fromParts(parts).write(), slides: plans };
}

/**
 * A kept drawing's XML, its relationships made the part's own — or null when
 * it is not one this knows, names what its package does not hold, or is no
 * longer the drawing the older records describe.
 */
function keptShape(item, part, ctx, slideNumber = null) {
  const parts = unzip(item.drawingML);
  if (!parts) return null;
  const name = [...parts.keys()].find((n) => /^drs\/[^/]+\.xml$/.test(n) && !/downrev/i.test(n));
  if (!name) return null;
  let xml = xmlOf(parts, name);
  // A table PowerPoint keeps as a frame of another name, whole.
  if (/^<p:E2oFrame\b/.test(xml || '')) xml = xml.replace(/^<p:E2oFrame\b/, '<p:graphicFrame').replace(/<\/p:E2oFrame>\s*$/, '</p:graphicFrame>');
  // A connector it keeps under a shape's name.
  if (/^<p:sp\b/.test(xml || '') && /^<p:sp\b[^>]*><p:nvCxnSpPr>/.test(xml)) xml = xml.replace(/^<p:sp\b/, '<p:cxnSp').replace(/<\/p:sp>\s*$/, '</p:cxnSp>');
  if (!xml || !/^<p:(sp|cxnSp|pic|grpSp|graphicFrame)\b/.test(xml) || !inPlace(xml, item)) return null;
  xml = withWords(xml, item, slideNumber);
  if (xml == null) return null;
  // A picture's own bytes, from the older records, should its package not hold them.
  const picture = item.shapes.find((s) => s.type === 'picture');
  const fallback = picture ? ctx.images[picture.image] : null;
  return adopt(xml, parts, name.replace(/([^/]+)$/, '_rels/$1.rels'), 'drs', part, ctx, fallback);
}

/**
 * Kept XML moved into a part: each relationship it names the part's own — a
 * link as it was, a picture copied into the deck's media (or, missing from
 * the package, the picture the older records hold). Null when it names one
 * that is neither.
 */
function adopt(xml, parts, relsName, dir, part, ctx, fallback = null) {
  const rels = new Map();
  const relsXml = xmlOf(parts, relsName);
  if (relsXml) for (const m of relsXml.matchAll(/<Relationship\b([^>]*?)\/?>/g)) { const a = attrsOf(m[1]); rels.set(a.Id, a); }
  let failed = false;
  const out = xml.replace(/\br:(id|embed|link|pict)="([^"]*)"/g, (all, kind, rid) => {
    const r = rels.get(rid);
    if (r && r.TargetMode === 'External') return `r:${kind}="${part.rel(r.Type, r.Target, true)}"`;
    const bytes = r ? parts.get(resolve(dir, r.Target)) : null;
    if (bytes && /\/image$/.test(r.Type)) return `r:${kind}="${part.rel(r.Type, '../' + ctx.mediaOf(bytes, r.Target.split('.').pop().toLowerCase()))}"`;
    if (kind === 'embed' && fallback?.bytes) {
      let data = fallback.bytes;
      if (fallback.deflated) { try { data = zlib.inflateSync(Buffer.from(data)); } catch { failed = true; return all; } }
      return `r:${kind}="${part.rel(REL + 'image', '../' + ctx.mediaOf(data, EXT[fallback.contentType] || 'png'))}"`;
    }
    failed = true;
    return all;
  });
  return failed ? null : out;
}

/** A text's words as the older records hold them, its fields left out: paragraphs split at \\r, a line break a \\v. */
const wordsOf = (shape) => shape?.words ?? shape?.paragraphs?.words ?? '';

/**
 * Kept XML with its words put back: PowerPoint keeps each character masked
 * (a letter "_", the rest a space), the words themselves being in the older
 * records — each paragraph takes the next of theirs, each run as many of its
 * characters as its mask has, a line break theirs, and a field (whose value
 * moves) its value. A table's cells each take their cell's. Null when the
 * words do not fit the masks: an older PowerPoint changed them since.
 */
function withWords(xml, item, slideNumber) {
  if (/^<p:graphicFrame\b/.test(xml) && /<a:tbl>/.test(xml)) {
    const table = item.shapes[0];
    if (table?.type !== 'table') return /<a:t(\s[^>]*)?>[^<]+<\/a:t>/.test(xml) ? null : xml;
    let r = -1;
    let failed = false;
    const out = xml.replace(/<a:tr\b[\s\S]*?<\/a:tr>/g, (row) => {
      r += 1;
      let c = -1;
      return row.replace(/<a:tc\b[\s\S]*?<\/a:tc>/g, (cell) => {
        c += 1;
        const own = table.cells[r]?.[c];
        const done = own ? fillWords(cell, own.covered ? '' : wordsOf(own), slideNumber) : null;
        if (done == null) failed = true;
        return done ?? cell;
      });
    });
    return failed ? null : out;
  }
  if (/<p:grpSp\b/.test(xml.slice(1))) return /<a:t(\s[^>]*)?>[^<]+<\/a:t>/.test(xml) ? null : xml;
  if (!/<p:txBody\b/.test(xml)) return xml;
  return fillWords(xml, item.shapes.length === 1 ? wordsOf(item.shapes[0]) : '', slideNumber);
}

/** One text body's masks filled from its words, or null when they do not fit. */
function fillWords(xml, words, slideNumber) {
  const paragraphs = words.split('\r');
  let pi = 0;
  let failed = false;
  const out = xml.replace(/<a:p\b[^>]*?(?:\/>|>([\s\S]*?)<\/a:p>)/g, (whole, body) => {
    const text = [...(paragraphs[pi++] ?? '')];
    if (pi > paragraphs.length) failed = true;
    if (body == null) { if (text.length) failed = true; return whole; }
    let pos = 0;
    const filled = body.replace(/<a:r\b[\s\S]*?<\/a:r>|<a:br\b[^>]*?(?:\/>|>[\s\S]*?<\/a:br>)|<a:fld\b[\s\S]*?<\/a:fld>/g, (el) => {
      if (el.startsWith('<a:br')) {
        if (text[pos] === '\u000b' || text[pos] === '\n') pos += 1;
        else failed = true;
        return el;
      }
      if (el.startsWith('<a:fld')) {
        // A slide number its slide's; any other field as kept, its value its own.
        if (slideNumber != null && /\btype="slidenum"/.test(el)) return el.replace(/(<a:t(?:\s[^>]*)?>)[^<]*(<\/a:t>)/, (_, a, b) => a + slideNumber + b);
        return el;
      }
      return el.replace(/(<a:t(?:\s[^>]*)?>)([^<]*)(<\/a:t>)/, (_, a, mask, b) => {
        const n = [...unesc(mask)].length;
        const real = text.slice(pos, pos + n).join('');
        pos += n;
        if (real.length !== n || [...real].some((ch, i) => (ch === ' ') !== ([...unesc(mask)][i] === ' ') && /[a-z]/.test(ch))) failed = true;
        return a + esc(real.replace(/\u000b/g, '\n')) + b;
      });
    });
    if (pos !== text.length) failed = true;
    return whole.slice(0, whole.length - body.length - '</a:p>'.length) + filled + '</a:p>';
  });
  if (pi < paragraphs.length && !(paragraphs.length === 1 && paragraphs[0] === '')) failed = true;
  return failed ? null : out;
}

/**
 * Whether kept XML still stands where the older records put the drawing — an
 * older PowerPoint that moved it moved only those records.
 */
function inPlace(xml, item) {
  const sh = item.shapes[0];
  const box = /<a:off x="(-?\d+)" y="(-?\d+)"\/>\s*<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(xml);
  if (!box || !sh || item.group || item.shapes.length !== 1) return true;
  const [x, y, w, h] = box.slice(1).map(Number);
  const near = (a, b) => Math.abs(a - b) <= 3200; // two of the older records' units
  const cx = x + w / 2;
  const cy = y + h / 2;
  const scx = (sh.x + sh.w / 2) * EMU;
  const scy = (sh.y + sh.h / 2) * EMU;
  const sw = sh.w * EMU;
  const shh = sh.h * EMU;
  return near(cx, scx) && near(cy, scy) && ((near(w, sw) && near(h, shh)) || (near(w, shh) && near(h, sw)));
}

/** Kept XML given this part's id for it. */
const withId = (xml, id) => xml.replace(/(<p:cNvPr\b[^>]*?\bid=")\d+"/, (_, head) => `${head}${id}"`);

/** A kept background as the slide's or master's background, or ''. */
function backgroundXml(bytes) {
  const xml = xmlOf(unzip(bytes), 'drs/shapexml.xml');
  const m = xml && /^<p:background\b[^>]*>([\s\S]*)<\/p:background>$/.exec(xml.trim());
  return m && /<p:(bgPr|bgRef)\b/.test(m[1]) ? `<p:bg>${m[1]}</p:bg>` : '';
}

/** The older records' placeholder kinds, as a layout's: the type, and whether it stands upright. */
const PLACEHOLDERS = {
  1: 'title', 2: 'body', 3: 'ctrTitle', 4: 'subTitle', 7: 'dt', 8: 'sldNum', 9: 'ftr', 10: 'hdr', 13: 'title', 14: 'body', 15: 'ctrTitle',
  16: 'subTitle', 17: 'title', 18: 'body', 19: null, 20: 'chart', 21: 'tbl', 22: 'clipArt', 23: 'dgm', 24: 'media', 25: null, 26: 'pic',
};

/** A placeholder's kind as a layout names it: a content one names none. */
const kindOf = (ph) => ph.type || 'obj';
const OWN_KINDS = new Set(['title', 'ctrTitle', 'dt', 'ftr', 'sldNum', 'hdr', 'sldImg']);

/**
 * The layout's placeholder a slide's is: the older records give a slide's
 * placeholder no number of its own, so it is matched as PowerPoint matches
 * one — a title to the layout's title, a date, footer or number to its, and
 * the rest to the layout's content placeholders in order, one of the same
 * kind first.
 */
function matchPlaceholder(sh, placeholders, taken) {
  const type = PLACEHOLDERS[sh.placeholder] ?? null;
  const titled = type === 'title' || type === 'ctrTitle';
  const fits = (p) => (titled ? kindOf(p) === 'title' || kindOf(p) === 'ctrTitle' : OWN_KINDS.has(type) ? kindOf(p) === type : !OWN_KINDS.has(kindOf(p)));
  const open = placeholders.map((p, i) => ({ p, i })).filter(({ p, i }) => !taken.has(i) && fits(p));
  const pick = open.find(({ p }) => kindOf(p) === (type ?? 'obj')) ?? open[0];
  if (!pick) return null;
  taken.add(pick.i);
  return pick.p;
}

/**
 * A placeholder from the older records: of its layout's kind and number, in
 * its place, its words in only the looks they have of their own — the rest
 * its layout's and master's, as in PowerPoint. On the master, with the body
 * properties PowerPoint gives its title and body.
 */
function placeholderXml(sh, id, layoutPh = null, master = false) {
  const type = layoutPh ? layoutPh.type ?? null : PLACEHOLDERS[sh.placeholder] ?? null;
  const titled = type === 'title' || type === 'ctrTitle';
  const number = layoutPh ? layoutPh.idx : !titled && sh.placeholderIdx > 0 ? sh.placeholderIdx : null;
  const idx = number != null ? ` idx="${number}"` : '';
  const upright = (layoutPh ? layoutPh.orient === 'vert' : [17, 18, 25].includes(sh.placeholder)) ? ' orient="vert"' : '';
  const size = layoutPh?.sz ? ` sz="${layoutPh.sz}"` : '';
  const name = sh.name || `${titled ? 'Title' : 'Text Placeholder'} ${id - 1}`;
  const xfrm = `<a:xfrm${sh.rotation ? ` rot="${Math.round(sh.rotation * 60000)}"` : ''}><a:off x="${Math.round(sh.x * EMU)}" y="${Math.round(sh.y * EMU)}"/><a:ext cx="${Math.round(sh.w * EMU)}" cy="${Math.round(sh.h * EMU)}"/></a:xfrm>`;
  const paragraphs = (sh.ownParagraphs || []).map(ownParagraphXml).join('') || '<a:p><a:endParaRPr lang="en-US"/></a:p>';
  const anchor = { top: 't', middle: 'ctr', bottom: 'b' }[sh.anchor] ?? (titled ? 'ctr' : null);
  const body = master
    ? `<a:bodyPr vert="${upright ? 'eaVert' : 'horz'}" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"${anchor && anchor !== 't' ? ` anchor="${anchor}"` : ''}>${titled || type === 'body' ? '<a:normAutofit/>' : ''}</a:bodyPr>`
    : '<a:bodyPr/>';
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${esc(name)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph${type ? ` type="${type}"` : ''}${upright}${size}${idx}/></p:nvPr></p:nvSpPr>`
    + `<p:spPr>${xfrm}</p:spPr><p:txBody>${body}<a:lstStyle/>${paragraphs}</p:txBody></p:sp>`;
}

/** A paragraph with only its own looks: its level, and what its exception and its runs say. */
function ownParagraphXml(p) {
  const own = p.own || {};
  const attrs = [];
  if (p.level) attrs.push(`lvl="${p.level}"`);
  if (own.leftMargin != null) attrs.push(`marL="${Math.round(own.leftMargin * 1587.5)}"`);
  if (own.indent != null && own.leftMargin != null) attrs.push(`indent="${Math.round((own.indent - own.leftMargin) * 1587.5)}"`);
  if (own.align != null) attrs.push(`algn="${['l', 'ctr', 'r', 'just'][own.align] || 'l'}"`);
  let inner = '';
  const spacing = (v) => (v > 0 ? `<a:spcPct val="${v * 1000}"/>` : `<a:spcPts val="${Math.round((-v / 8) * 100)}"/>`);
  if (own.lineSpacing != null && own.lineSpacing !== 0) inner += `<a:lnSpc>${spacing(own.lineSpacing)}</a:lnSpc>`;
  if (own.spaceBefore != null) inner += `<a:spcBef>${spacing(own.spaceBefore || -0.0001)}</a:spcBef>`;
  if (own.spaceAfter != null) inner += `<a:spcAft>${spacing(own.spaceAfter || -0.0001)}</a:spcAft>`;
  if (own.bulletOn === false) inner += '<a:buNone/>';
  else if (own.bulletOn && p.bullet?.char) inner += `${p.bullet.color ? `<a:buClr><a:srgbClr val="${p.bullet.color.replace('#', '')}"/></a:buClr>` : ''}<a:buChar char="${esc(p.bullet.char)}"/>`;
  const pPr = attrs.length || inner ? `<a:pPr${attrs.length ? ' ' + attrs.join(' ') : ''}>${inner}</a:pPr>` : '';
  const runs = (p.runs || []).filter((r) => r.text).map((r) => {
    if (r.text === '\n') return '<a:br><a:rPr lang="en-US"/></a:br>';
    const a = [`lang="${esc(r.lang || 'en-US')}"`];
    if (r.altLang) a.push(`altLang="${esc(r.altLang)}"`);
    if (r.size) a.push(`sz="${Math.round(r.size * 100)}"`);
    if (r.bold) a.push('b="1"');
    if (r.italic) a.push('i="1"');
    if (r.underline) a.push('u="sng"');
    if (r.baseline === 'super') a.push('baseline="30000"');
    else if (r.baseline === 'sub') a.push('baseline="-25000"');
    const fill = r.color ? `<a:solidFill><a:srgbClr val="${r.color.replace('#', '')}"/></a:solidFill>` : '';
    const font = r.font ? `<a:latin typeface="${esc(r.font)}"/>` : '';
    return `<a:r><a:rPr ${a.join(' ')}${fill || font ? `>${fill}${font}</a:rPr>` : '/>'}<a:t>${esc(r.text)}</a:t></a:r>`;
  }).join('');
  return `<a:p>${pPr}${runs || '<a:endParaRPr lang="en-US"/>'}</a:p>`;
}
