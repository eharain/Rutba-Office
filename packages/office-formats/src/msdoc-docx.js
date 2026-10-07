// A Word binary document's model, as a .docx.
//
// msdoc.js (and msdoc-old.js for the older versions) read a .doc into
// paragraphs, runs, tables, styles, lists, sections, notes and pictures;
// this writes them out as WordprocessingML — the styles as styles, the
// lists as numbering, each paragraph and run with only what it says beyond
// its style, the tables with their grid and merged cells, the fields as
// fields (a hyperlink as a link), each section's page and its headers and
// footers, the notes, and the pictures as parts of their own. What a .doc
// holds that has no place here is left out rather than guessed at.

import zlib from 'node:zlib';
import { buildDocx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml/package';
import { anchorXml, textBoxRun, DRAWING_NS, Z_BASE } from '@rutba/ooxml/drawings';
import { ICO } from './msdoc.js';

// WordprocessingML, and every namespace a drawing in it may use — floating
// text boxes and shapes (wps, inside mc:AlternateContent, with a VML twin).
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
  + Object.entries(DRAWING_NS).map(([p, uri]) => `xmlns:${p}="${uri}"`).join(' ');
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
// Characters XML 1.0 cannot carry: control codes a .doc keeps that have no meaning here.
const clean = (s) => String(s).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]/g, '');

const JC = ['left', 'center', 'right', 'both', 'distribute'];
const UNDERLINE = { 1: 'single', 2: 'words', 3: 'double', 4: 'dotted', 6: 'thick', 7: 'dash', 9: 'dotDash', 10: 'dotDotDash', 11: 'wave', 20: 'dottedHeavy', 23: 'dashedHeavy', 25: 'dashDotHeavy', 26: 'dashDotDotHeavy', 27: 'wavyHeavy', 39: 'dashLong', 43: 'wavyDouble', 55: 'dashLongHeavy' };
const HIGHLIGHT = { '000000': 'black', '0000FF': 'blue', '00FFFF': 'cyan', '00FF00': 'green', 'FF00FF': 'magenta', 'FF0000': 'red', 'FFFF00': 'yellow', 'FFFFFF': 'white', '000080': 'darkBlue', '008080': 'darkCyan', '008000': 'darkGreen', '800080': 'darkMagenta', '800000': 'darkRed', '808000': 'darkYellow', '808080': 'darkGray', 'C0C0C0': 'lightGray' };
const NUMFMT = { 0: 'decimal', 1: 'upperRoman', 2: 'lowerRoman', 3: 'upperLetter', 4: 'lowerLetter', 5: 'ordinal', 6: 'cardinalText', 7: 'ordinalText', 22: 'decimalZero', 23: 'bullet', 255: 'none' };
const BORDER = { 1: 'single', 2: 'thick', 3: 'double', 5: 'single', 6: 'dotted', 7: 'dashed', 8: 'dotDash', 9: 'dotDotDash', 10: 'triple', 11: 'thinThickSmallGap', 12: 'thickThinSmallGap', 13: 'thinThickThinSmallGap', 18: 'threeDEmboss', 19: 'threeDEngrave', 20: 'wave', 21: 'doubleWave', 22: 'dashSmallGap', 23: 'dashDotStroked' };
const SECTION_BREAK = ['continuous', 'nextColumn', 'nextPage', 'evenPage', 'oddPage'];

/** A style name as a style id: its words run together, as Word makes them. */
function styleIdOf(name, used) {
  let id = String(name || 'Style').split(/[^A-Za-z0-9]+/).filter(Boolean).map((w) => w[0].toUpperCase() + w.slice(1)).join('') || 'Style';
  let n = 1;
  const baseId = id;
  while (used.has(id)) id = baseId + ++n;
  used.add(id);
  return id;
}

/**
 * Write the model as a .docx: returns the file's bytes.
 * @param {object} model what msdoc.js's readDoc (or msdoc-old.js) gives
 */
export function docModelToDocx(model) {
  const pkg = OoxmlPackage.read(buildDocx({ styles: true, paragraphs: [] }));
  const fonts = model.fonts || [];
  const fontName = (i) => fonts[i]?.name || null;

  // Styles: ids from names, unique; the resolved CHP and PAP of each for "what the style already says".
  const used = new Set();
  const styleById = new Map();
  for (const st of model.styles) {
    st.id = st.istd === 0 ? 'Normal' : styleIdOf(st.name, used);
    if (st.istd === 0) used.add('Normal');
    styleById.set(st.istd, st);
  }
  const styleOf = (istd) => styleById.get(istd) || styleById.get(0) || { pap: {}, chp: {} };

  // Numbering: an abstract list per list, a numbering instance per override.
  const listIndex = new Map((model.lists?.lists || []).map((l, i) => [l.lsid, i]));
  const numIdOf = (ilfo) => {
    if (!ilfo || ilfo >= 0x7ff) return null;
    const lfo = model.lists?.lfos?.[ilfo - 1];
    return lfo && listIndex.has(lfo.lsid) ? ilfo : null;
  };

  /* ── properties ─────────────────────────────────────────────────────── */

  const borderXml = (name, b) => (b ? `<w:${name} w:val="${BORDER[b.type] || 'single'}" w:sz="${Math.max(2, b.width || 4)}" w:space="${b.space || 0}" w:color="${b.colour || 'auto'}"/>` : '');

  /** A run's properties: only what differs from `base` (its paragraph style's CHP). */
  const rPrXml = (chp, base = {}) => {
    const out = [];
    const diff = (k) => chp[k] !== undefined && chp[k] !== base[k];
    if (chp.istd != null && styleById.get(chp.istd)?.kind === 'character' && styleById.get(chp.istd).sti !== 65) out.push(`<w:rStyle w:val="${styleById.get(chp.istd).id}"/>`);
    if (diff('font') || diff('fontFE') || diff('fontOther')) {
      const a = fontName(chp.font);
      const ea = fontName(chp.fontFE);
      const cs = fontName(chp.fontOther);
      out.push(`<w:rFonts${a ? ` w:ascii="${esc(a)}" w:hAnsi="${esc(a)}"` : ''}${ea ? ` w:eastAsia="${esc(ea)}"` : ''}${cs ? ` w:cs="${esc(cs)}"` : ''}/>`);
    }
    const tog = (k, el) => { if (diff(k)) out.push(chp[k] ? `<w:${el}/>` : `<w:${el} w:val="0"/>`); };
    tog('bold', 'b'); tog('boldBi', 'bCs'); tog('italic', 'i'); tog('italicBi', 'iCs'); tog('caps', 'caps'); tog('smallCaps', 'smallCaps');
    tog('strike', 'strike'); tog('dstrike', 'dstrike'); tog('outline', 'outline'); tog('shadow', 'shadow'); tog('vanish', 'vanish');
    if (diff('colour') || (chp.colourSet && !base.colourSet && chp.colour)) out.push(`<w:color w:val="${chp.colour || 'auto'}"/>`);
    if (diff('spacing')) out.push(`<w:spacing w:val="${chp.spacing}"/>`);
    if (diff('size')) out.push(`<w:sz w:val="${chp.size}"/>`);
    if (diff('sizeBi')) out.push(`<w:szCs w:val="${chp.sizeBi}"/>`);
    if (diff('highlight') && chp.highlight && HIGHLIGHT[chp.highlight]) out.push(`<w:highlight w:val="${HIGHLIGHT[chp.highlight]}"/>`);
    if (diff('underline')) out.push(`<w:u w:val="${chp.underline ? UNDERLINE[chp.underline] || 'single' : 'none'}"/>`);
    if (diff('shading') && chp.shading) out.push(`<w:shd w:val="clear" w:color="auto" w:fill="${chp.shading}"/>`);
    if (diff('iss')) out.push(`<w:vertAlign w:val="${chp.iss === 1 ? 'superscript' : chp.iss === 2 ? 'subscript' : 'baseline'}"/>`);
    if (diff('rtl') && chp.rtl) out.push('<w:rtl/>');
    return out.length ? `<w:rPr>${out.join('')}</w:rPr>` : '';
  };

  /** A paragraph's properties: its style, and what it says beyond the style. */
  const pPrXml = (pap, extra = '', markChp = null) => {
    const st = styleOf(pap.istd ?? 0);
    const base = st.pap || {};
    const out = [];
    if ((pap.istd ?? 0) !== 0 && st.kind === 'paragraph') out.push(`<w:pStyle w:val="${st.id}"/>`);
    const d = (k) => pap[k] !== undefined && pap[k] !== base[k];
    if (d('keepNext')) out.push(pap.keepNext ? '<w:keepNext/>' : '<w:keepNext w:val="0"/>');
    if (d('keep')) out.push(pap.keep ? '<w:keepLines/>' : '<w:keepLines w:val="0"/>');
    if (d('pageBreakBefore')) out.push(pap.pageBreakBefore ? '<w:pageBreakBefore/>' : '<w:pageBreakBefore w:val="0"/>');
    if (d('widowControl')) out.push(pap.widowControl ? '<w:widowControl/>' : '<w:widowControl w:val="0"/>');
    const numId = numIdOf(pap.ilfo);
    if (numId && (pap.ilfo !== base.ilfo || pap.ilvl !== base.ilvl)) out.push(`<w:numPr><w:ilvl w:val="${pap.ilvl || 0}"/><w:numId w:val="${numId}"/></w:numPr>`);
    else if (base.ilfo && !pap.ilfo) out.push('<w:numPr><w:numId w:val="0"/></w:numPr>');
    if (pap.borders && JSON.stringify(pap.borders) !== JSON.stringify(base.borders)) {
      const b = pap.borders;
      out.push(`<w:pBdr>${borderXml('top', b.top)}${borderXml('left', b.left)}${borderXml('bottom', b.bottom)}${borderXml('right', b.right)}</w:pBdr>`);
    }
    if (d('shading') && pap.shading) out.push(`<w:shd w:val="clear" w:color="auto" w:fill="${pap.shading}"/>`);
    if (pap.tabs && JSON.stringify(pap.tabs) !== JSON.stringify(base.tabs)) {
      const kinds = ['left', 'center', 'right', 'decimal', 'bar'];
      const leaders = ['none', 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot'];
      out.push(`<w:tabs>${pap.tabs.map((t) => `<w:tab w:val="${kinds[t.jc] || 'left'}" w:pos="${t.pos}"${t.leader ? ` w:leader="${leaders[t.leader] || 'dot'}"` : ''}/>`).join('')}</w:tabs>`);
    }
    if (d('bidi')) out.push(pap.bidi ? '<w:bidi/>' : '<w:bidi w:val="0"/>');
    const sp = [];
    if (d('before')) sp.push(`w:before="${pap.before}"`);
    if (d('after')) sp.push(`w:after="${pap.after}"`);
    if (pap.line && JSON.stringify(pap.line) !== JSON.stringify(base.line)) {
      if (pap.line.mult) sp.push(`w:line="${pap.line.dya}" w:lineRule="auto"`);
      else sp.push(`w:line="${Math.abs(pap.line.dya)}" w:lineRule="${pap.line.dya < 0 ? 'exact' : 'atLeast'}"`);
    }
    if (sp.length) out.push(`<w:spacing ${sp.join(' ')}/>`);
    const ind = [];
    if (d('left')) ind.push(`w:left="${pap.left}"`);
    if (d('right')) ind.push(`w:right="${pap.right}"`);
    if (d('firstLine')) ind.push(pap.firstLine < 0 ? `w:hanging="${-pap.firstLine}"` : `w:firstLine="${pap.firstLine}"`);
    if (ind.length) out.push(`<w:ind ${ind.join(' ')}/>`);
    if (d('jc')) out.push(`<w:jc w:val="${JC[pap.jc] || 'left'}"/>`);
    if (d('outLvl') && pap.outLvl < 9) out.push(`<w:outlineLvl w:val="${pap.outLvl}"/>`);
    if (markChp) {
      const r = rPrXml(markChp, st.chp || {});
      if (r) out.push(r);
    }
    if (extra) out.push(extra);
    return out.length ? `<w:pPr>${out.join('')}</w:pPr>` : '';
  };

  /* ── runs, fields, pictures ─────────────────────────────────────────── */

  let drawingId = 0;
  const imageRel = new Map(); // `${part}|${index}` -> rId
  let mediaN = 0;
  const imageRelFor = (part, index) => {
    const key = part + '|' + index;
    if (imageRel.has(key)) return imageRel.get(key);
    const img = model.images[index];
    let bytes = img.bytes;
    if (img.deflated) {
      try { bytes = zlib.inflateSync(Buffer.from(bytes)); } catch { bytes = Buffer.from(bytes); }
    }
    const name = `word/media/image${++mediaN}.${img.ext}`;
    pkg.addPart(name, Buffer.from(bytes), null);
    ensureDefaultType(pkg, img.ext, img.contentType);
    const rId = pkg.addRelationshipTo(part, REL + 'image', 'media/' + name.split('/').pop());
    imageRel.set(key, rId);
    return rId;
  };

  /** A picture's graphic: its bytes by relationship, stretched to its box. */
  const pictureGraphic = (rId, cx, cy, id, name) => '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
    + `<pic:nvPicPr><pic:cNvPr id="${id}" name="${esc(name)}"/><pic:cNvPicPr/></pic:nvPicPr>`
    + `<pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
    + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`
    + '</pic:pic></a:graphicData></a:graphic>';

  const pictureXml = (index, part) => {
    const img = model.images[index];
    if (!img) return '';
    const rId = imageRelFor(part, index);
    const cx = Math.max(1, Math.round((img.widthTwips || 1440) * 635));
    const cy = Math.max(1, Math.round((img.heightTwips || 1440) * 635));
    const id = ++drawingId;
    return `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${id}" name="Picture ${id}"/>`
      + pictureGraphic(rId, cx, cy, id, `Picture ${id}`) + '</wp:inline></w:drawing>';
  };

  /**
   * A floating drawing, where its anchor stood: each of its pictures, text
   * boxes and shapes (a group's members one by one) at its place on the
   * page or by the paragraph, wrapped and stacked as it was — written as the
   * editor writes its own floating drawings.
   */
  let stack = Z_BASE;
  const floatXml = (float, ctx) => float.items.map((item) => {
    const id = ++drawingId;
    stack += 1024;
    const name = item.name || `${item.kind === 'picture' ? 'Picture' : item.kind === 'textbox' ? 'Text Box' : 'Shape'} ${id}`;
    const place = { wrap: float.wrap, side: float.side, behind: float.behind, relativeHeight: stack, h: { rel: float.relH, offsetPx: item.x }, v: { rel: float.relV, offsetPx: item.y } };
    const cx = Math.max(1, Math.round(item.w * 9525));
    const cy = Math.max(1, Math.round(item.h * 9525));
    const pieces = (graphic, frame = '<wp:cNvGraphicFramePr/>') => ({
      open: '<w:drawing>', extent: `<wp:extent cx="${cx}" cy="${cy}"/>`, effectExtent: '<wp:effectExtent l="0" t="0" r="0" b="0"/>',
      docPr: `<wp:docPr id="${id}" name="${esc(name)}"/>`, frame, graphic,
    });
    if (item.kind === 'picture') {
      if (!model.images[item.image]) return '';
      const graphic = pictureGraphic(imageRelFor(ctx.part, item.image), cx, cy, id, name);
      return `<w:r>${anchorXml(pieces(graphic, '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'), place)}</w:r>`;
    }
    if (item.kind === 'textbox') {
      const inner = blocksXml(item.blocks, { part: ctx.part });
      return textBoxRun({ id, name, widthPx: item.w, heightPx: item.h, ...place, fill: item.fill, line: item.line, lineWidthPx: item.lineWidthPx, paragraphs: /<w:p[ >]/.test(inner) ? inner : '<w:p/>' });
    }
    const e = (v) => Math.round(v * 9525);
    const geom = item.path
      ? `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="${cx}" h="${cy}"${item.path.filled ? '' : ' fill="none"'}>`
        + item.path.commands.map((c) => (c.op === 'M' ? `<a:moveTo><a:pt x="${e(c.pts[0][0])}" y="${e(c.pts[0][1])}"/></a:moveTo>`
          : c.op === 'L' ? `<a:lnTo><a:pt x="${e(c.pts[0][0])}" y="${e(c.pts[0][1])}"/></a:lnTo>`
            : c.op === 'C' ? `<a:cubicBezTo>${c.pts.map(([x, y]) => `<a:pt x="${e(x)}" y="${e(y)}"/>`).join('')}</a:cubicBezTo>` : '<a:close/>')).join('')
        + '</a:path></a:pathLst></a:custGeom>'
      : `<a:prstGeom prst="${esc(item.preset)}"><a:avLst/></a:prstGeom>`;
    const fill = item.fill ? `<a:solidFill><a:srgbClr val="${item.fill}"/></a:solidFill>` : '<a:noFill/>';
    const line = item.line ? `<a:ln w="${Math.max(1, Math.round((item.lineWidthPx || 1) * 9525))}"><a:solidFill><a:srgbClr val="${item.line}"/></a:solidFill></a:ln>` : '<a:ln><a:noFill/></a:ln>';
    const graphic = `<a:graphic><a:graphicData uri="${DRAWING_NS.wps}"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm${item.flipH ? ' flipH="1"' : ''}${item.flipV ? ' flipV="1"' : ''}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>${geom}${fill}${line}</wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic>`;
    return `<w:r><mc:AlternateContent><mc:Choice Requires="wps">${anchorXml(pieces(graphic), place)}</mc:Choice></mc:AlternateContent></w:r>`;
  }).join('');

  /** One run, of any kind, as w:r (or w:r runs). */
  const runXml = (run, base, ctx, instr = false) => {
    const rPr = rPrXml(run.chp || {}, base);
    if (run.text != null) {
      const t = clean(run.text);
      if (!t) return '';
      return instr ? `<w:r>${rPr}<w:instrText xml:space="preserve">${esc(t)}</w:instrText></w:r>` : `<w:r>${rPr}<w:t xml:space="preserve">${esc(t)}</w:t></w:r>`;
    }
    switch (run.kind) {
      case 'tab': return `<w:r>${rPr}<w:tab/></w:r>`;
      case 'break': return `<w:r>${rPr}${run.type === 'page' ? '<w:br w:type="page"/>' : run.type === 'column' ? '<w:br w:type="column"/>' : '<w:br/>'}</w:r>`;
      case 'noBreakHyphen': return `<w:r>${rPr}<w:noBreakHyphen/></w:r>`;
      case 'softHyphen': return `<w:r>${rPr}<w:softHyphen/></w:r>`;
      case 'symbol': return `<w:r>${rPr}<w:sym w:font="${esc(fontName(run.font) || 'Symbol')}" w:char="${run.char.toString(16).toUpperCase().padStart(4, '0')}"/></w:r>`;
      case 'picture': return `<w:r>${rPr}${pictureXml(run.image, ctx.part)}</w:r>`;
      case 'float': return floatXml(run.float, ctx);
      case 'noteRef': return `<w:r>${rPr}<w:${run.note.kind}Reference w:id="${run.note.id}"/></w:r>`;
      case 'noteMark': return ctx.noteKind ? `<w:r>${rPr}<w:${ctx.noteKind}Ref/></w:r>` : '';
      default: return '';
    }
  };

  /**
   * A paragraph's runs, its fields among them: a field is begin, its
   * instruction, separator, its result, end — nested fields inside. A
   * HYPERLINK becomes a link round its result.
   */
  const runsXml = (runs, base, ctx) => {
    let out = '';
    let i = 0;
    const renderUntil = (stop) => {
      let s = '';
      while (i < runs.length) {
        const r = runs[i];
        if (r.kind === 'fieldSep' || r.kind === 'fieldEnd') {
          if (stop) return s;
          i += 1;
          continue;
        }
        if (r.kind === 'fieldBegin') { s += fieldXml(); continue; }
        s += runXml(r, base, ctx);
        i += 1;
      }
      return s;
    };
    const fieldXml = () => {
      i += 1; // the begin
      let instr = '';
      const instrRuns = [];
      while (i < runs.length && runs[i].kind !== 'fieldSep' && runs[i].kind !== 'fieldEnd') {
        if (runs[i].kind === 'fieldBegin') { // a field in the instruction: kept as its text
          const inner = fieldXml();
          instrRuns.push({ xml: inner });
          continue;
        }
        if (runs[i].text != null) { instr += runs[i].text; instrRuns.push({ run: runs[i] }); }
        i += 1;
      }
      let result = '';
      const hasSep = runs[i]?.kind === 'fieldSep';
      if (hasSep) { i += 1; result = renderUntil(true); }
      if (runs[i]?.kind === 'fieldEnd') i += 1;
      const link = /^\s*HYPERLINK\s+(?:\\l\s+)?"([^"]*)"/i.exec(instr);
      if (link && result && !/^\s*HYPERLINK\s+\\l/i.test(instr)) {
        const rId = pkg.addRelationshipTo(ctx.part, REL + 'hyperlink', link[1], { external: true });
        return `<w:hyperlink r:id="${rId}">${result}</w:hyperlink>`;
      }
      const first = instrRuns.find((x) => x.run)?.run;
      const rPr = rPrXml(first?.chp || {}, base);
      return `<w:r>${rPr}<w:fldChar w:fldCharType="begin"/></w:r>`
        + instrRuns.map((x) => (x.run ? runXml(x.run, base, ctx, true) : x.xml)).join('')
        + (hasSep ? `<w:r>${rPr}<w:fldChar w:fldCharType="separate"/></w:r>${result}` : '')
        + `<w:r>${rPr}<w:fldChar w:fldCharType="end"/></w:r>`;
    };
    while (i < runs.length) out += renderUntil(false);
    return out;
  };

  /* ── blocks ─────────────────────────────────────────────────────────── */

  const paragraphXml = (p, ctx, extraPPr = '') => {
    const st = styleOf(p.pap.istd ?? 0);
    return `<w:p>${pPrXml(p.pap, extraPPr, p.markChp)}${runsXml(p.runs, st.chp || {}, ctx)}</w:p>`;
  };

  const tableXml = (t, ctx) => {
    // The grid: every cell edge of every row, from the left.
    const edges = new Set([0]);
    for (const row of t.rows) {
      let x = 0;
      for (const c of row.tap.cells || []) { x += Math.max(0, c.width); edges.add(x); }
    }
    const grid = [...edges].sort((a, b) => a - b);
    const cols = grid.slice(1).map((x, i) => x - grid[i]).filter((w) => w > 0);
    const tblBorders = t.rows.find((r) => r.tap.borders)?.tap.borders;
    let xml = '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>';
    if (tblBorders) xml += `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((k) => borderXml(k, tblBorders[k])).join('')}</w:tblBorders>`;
    xml += '<w:tblLayout w:type="fixed"/><w:tblLook w:val="0000"/></w:tblPr>';
    if (cols.length) xml += `<w:tblGrid>${cols.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;
    for (const row of t.rows) {
      const tapCells = row.tap.cells || [];
      let trPr = '';
      if (row.tap.rowHeight) trPr += `<w:trHeight w:val="${Math.abs(row.tap.rowHeight)}"${row.tap.rowHeight < 0 ? ' w:hRule="exact"' : ''}/>`;
      if (row.tap.header) trPr += '<w:tblHeader/>';
      xml += `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ''}`;
      let x = 0;
      for (let c = 0; c < row.cells.length; c++) {
        const tc = tapCells[c] || { width: cols[c] || 2000 };
        if (tc.hMerge >= 2) { x += Math.max(0, tc.width); continue; }
        // A cell that starts a horizontal merge takes the merged cells' widths.
        let width = Math.max(0, tc.width);
        let k = c + 1;
        while (tc.hMerge === 1 && tapCells[k]?.hMerge >= 2) { width += Math.max(0, tapCells[k].width); k += 1; }
        const startCol = grid.indexOf(x);
        const endCol = grid.indexOf(x + width);
        const span = startCol >= 0 && endCol > startCol ? endCol - startCol : 1;
        x += width;
        let tcPr = `<w:tcW w:w="${width}" w:type="dxa"/>`;
        if (span > 1) tcPr += `<w:gridSpan w:val="${span}"/>`;
        if (tc.vMerge === 3) tcPr += '<w:vMerge w:val="restart"/>';
        else if (tc.vMerge === 2 || tc.vMerge === 1) tcPr += '<w:vMerge/>';
        if (tc.borders && Object.values(tc.borders).some(Boolean)) tcPr += `<w:tcBorders>${borderXml('top', tc.borders.top)}${borderXml('left', tc.borders.left)}${borderXml('bottom', tc.borders.bottom)}${borderXml('right', tc.borders.right)}</w:tcBorders>`;
        const fill = row.tap.cellShading?.[c];
        if (fill) tcPr += `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>`;
        if (tc.vAlign) tcPr += `<w:vAlign w:val="${['top', 'center', 'bottom'][tc.vAlign] || 'top'}"/>`;
        const inner = blocksXml(row.cells[c], ctx);
        xml += `<w:tc><w:tcPr>${tcPr}</w:tcPr>${/<w:p[ >]/.test(inner) ? inner : inner + '<w:p/>'}</w:tc>`;
      }
      xml += '</w:tr>';
    }
    return xml + '</w:tbl>';
  };

  const blocksXml = (blocks, ctx) => blocks.map((b) => {
    if (b.type === 'table') return tableXml(b, ctx) + (b.sectionEnd != null && ctx.sectPr ? `<w:p><w:pPr>${ctx.sectPr(b.sectionEnd)}</w:pPr></w:p>` : '');
    const extra = b.sectionEnd != null && ctx.sectPr ? ctx.sectPr(b.sectionEnd) : '';
    return paragraphXml(b, ctx, extra);
  }).join('');

  /* ── headers, footers, sections ─────────────────────────────────────── */

  const storyPart = new Map(); // the story's model object -> rId
  let headerN = 0;
  let footerN = 0;
  const storyRel = (blocks, kind) => {
    if (storyPart.has(blocks)) return storyPart.get(blocks);
    const n = kind === 'header' ? ++headerN : ++footerN;
    const name = `word/${kind}${n}.xml`;
    pkg.addPart(name, '', CT + kind + '+xml');
    const tag = kind === 'header' ? 'hdr' : 'ftr';
    const inner = blocksXml(blocks, { part: name });
    pkg.write_(name, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${tag} ${W}>${/<w:p[ >]/.test(inner) ? inner : '<w:p/>'}</w:${tag}>`);
    const rId = pkg.addRelationshipTo('word/document.xml', REL + kind, `${kind}${n}.xml`);
    storyPart.set(blocks, rId);
    return rId;
  };

  const sectPrXml = (s) => {
    const sec = model.sections[s];
    const sep = sec.sep;
    const hf = model.headers?.[s] || { headers: {}, footers: {} };
    let refs = '';
    const has = (blocks) => blocks && blocks.some((b) => b.type === 'table' || b.runs?.length);
    for (const [type, blocks] of Object.entries(hf.headers)) if (has(blocks) && (type !== 'first' || sep.titlePage) && (type !== 'even' || model.evenAndOdd)) refs += `<w:headerReference w:type="${type}" r:id="${storyRel(blocks, 'header')}"/>`;
    for (const [type, blocks] of Object.entries(hf.footers)) if (has(blocks) && (type !== 'first' || sep.titlePage) && (type !== 'even' || model.evenAndOdd)) refs += `<w:footerReference w:type="${type}" r:id="${storyRel(blocks, 'footer')}"/>`;
    const type = s > 0 && SECTION_BREAK[sep.break] && sep.break !== 2 ? `<w:type w:val="${SECTION_BREAK[sep.break]}"/>` : '';
    return `<w:sectPr>${refs}${type}<w:pgSz w:w="${sep.width}" w:h="${sep.height}"${sep.landscape ? ' w:orient="landscape"' : ''}/>`
      + `<w:pgMar w:top="${sep.top}" w:right="${sep.right}" w:bottom="${sep.bottom}" w:left="${sep.left}" w:header="${sep.header}" w:footer="${sep.footer}" w:gutter="0"/>`
      + (sep.columns > 1 ? `<w:cols w:num="${sep.columns}"/>` : '<w:cols w:space="720"/>')
      + (sep.titlePage ? '<w:titlePg/>' : '') + '</w:sectPr>';
  };

  /* ── the parts ──────────────────────────────────────────────────────── */

  // Notes first, so the body's references point at notes that exist.
  for (const kind of ['footnote', 'endnote']) {
    const list = model.notes?.[kind] || [];
    if (!list.length) continue;
    const part = `word/${kind}s.xml`;
    pkg.addPart(part, '', CT + kind + 's+xml');
    const sep = (type, id, el) => `<w:${kind} w:type="${type}" w:id="${id}"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:${el}/></w:r></w:p></w:${kind}>`;
    const body = list.map((n) => `<w:${kind} w:id="${n.id}">${blocksXml(n.blocks, { part, noteKind: kind }) || '<w:p/>'}</w:${kind}>`).join('');
    pkg.write_(part, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${kind}s ${W}>${sep('separator', -1, 'separator')}${sep('continuationSeparator', 0, 'continuationSeparator')}${body}</w:${kind}s>`);
    pkg.addRelationshipTo('word/document.xml', REL + kind + 's', `${kind}s.xml`);
  }

  const lastSection = model.sections.length - 1;
  const body = blocksXml(model.body, { part: 'word/document.xml', sectPr: (s) => (s < lastSection ? sectPrXml(s) : '') });
  pkg.write_('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body || '<w:p/>'}${sectPrXml(lastSection)}</w:body></w:document>`);

  pkg.write_('word/styles.xml', stylesXml(model, styleById, fontName));
  if (model.lists?.lists?.length) {
    pkg.addPart('word/numbering.xml', numberingXml(model, fontName), CT + 'numbering+xml');
    pkg.addRelationshipTo('word/document.xml', REL + 'numbering', 'numbering.xml');
  }
  return pkg.write();
}

/** Register a file extension's content type, once. */
function ensureDefaultType(pkg, ext, type) {
  const ctName = '[Content_Types].xml';
  const xml = pkg.text(ctName);
  if (new RegExp(`Extension="${ext}"`, 'i').test(xml)) return;
  pkg.write_(ctName, xml.replace('<Override', () => `<Default Extension="${ext}" ContentType="${type}"/><Override`));
}

/** The style sheet: the document defaults, then each style with its own properties over its base. */
function stylesXml(model, styleById, fontName) {
  const font = fontName(model.defaultFonts?.[0] ?? 0);
  let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W}>`
    + `<w:docDefaults><w:rPrDefault><w:rPr>${font ? `<w:rFonts w:ascii="${esc(font)}" w:hAnsi="${esc(font)}" w:cs="${esc(font)}"/>` : ''}<w:sz w:val="20"/><w:szCs w:val="20"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>`;
  for (const st of model.styles) {
    if (!['paragraph', 'character', 'table', 'numbering'].includes(st.kind)) continue;
    const base = st.base != null ? styleById.get(st.base) : null;
    const next = st.next != null ? styleById.get(st.next) : null;
    const isDefault = st.istd === 0 || (st.kind === 'character' && st.sti === 65) || (st.kind === 'table' && st.sti === 105) || (st.kind === 'numbering' && st.sti === 107);
    xml += `<w:style w:type="${st.kind}" w:styleId="${st.id}"${isDefault ? ' w:default="1"' : ''}><w:name w:val="${esc(st.name || st.id)}"/>`;
    if (base) xml += `<w:basedOn w:val="${base.id}"/>`;
    if (next && st.kind === 'paragraph') xml += `<w:next w:val="${next.id}"/>`;
    if (st.kind === 'paragraph' || st.kind === 'character') xml += '<w:qFormat/>';
    if (st.kind === 'paragraph') {
      const p = ownPap(st, base);
      if (p) xml += p;
    }
    if (st.kind === 'paragraph' || st.kind === 'character') {
      const r = ownRpr(st, base, fontName);
      if (r) xml += r;
    }
    xml += '</w:style>';
  }
  return xml + '</w:styles>';
}

/** What a paragraph style says beyond its base. */
function ownPap(st, base) {
  const p = st.pap || {};
  const b = base?.pap || {};
  const out = [];
  const d = (k) => p[k] !== undefined && p[k] !== b[k];
  if (d('keepNext') && p.keepNext) out.push('<w:keepNext/>');
  if (d('keep') && p.keep) out.push('<w:keepLines/>');
  if (d('pageBreakBefore') && p.pageBreakBefore) out.push('<w:pageBreakBefore/>');
  const sp = [];
  if (d('before')) sp.push(`w:before="${p.before}"`);
  if (d('after')) sp.push(`w:after="${p.after}"`);
  if (p.line && JSON.stringify(p.line) !== JSON.stringify(b.line)) {
    if (p.line.mult) sp.push(`w:line="${p.line.dya}" w:lineRule="auto"`);
    else sp.push(`w:line="${Math.abs(p.line.dya)}" w:lineRule="${p.line.dya < 0 ? 'exact' : 'atLeast'}"`);
  }
  if (sp.length) out.push(`<w:spacing ${sp.join(' ')}/>`);
  const ind = [];
  if (d('left')) ind.push(`w:left="${p.left}"`);
  if (d('right')) ind.push(`w:right="${p.right}"`);
  if (d('firstLine')) ind.push(p.firstLine < 0 ? `w:hanging="${-p.firstLine}"` : `w:firstLine="${p.firstLine}"`);
  if (ind.length) out.push(`<w:ind ${ind.join(' ')}/>`);
  if (d('jc')) out.push(`<w:jc w:val="${JC[p.jc] || 'left'}"/>`);
  if (d('outLvl') && p.outLvl < 9) out.push(`<w:outlineLvl w:val="${p.outLvl}"/>`);
  return out.length ? `<w:pPr>${out.join('')}</w:pPr>` : '';
}

/** What a style's characters say beyond its base. */
function ownRpr(st, base, fontName) {
  const c = st.chp || {};
  const b = base?.chp || {};
  const out = [];
  const d = (k) => c[k] !== undefined && c[k] !== b[k];
  if (d('font')) { const n = fontName(c.font); if (n) out.push(`<w:rFonts w:ascii="${esc(n)}" w:hAnsi="${esc(n)}"/>`); }
  const tog = (k, el) => { if (d(k)) out.push(c[k] ? `<w:${el}/>` : `<w:${el} w:val="0"/>`); };
  tog('bold', 'b'); tog('italic', 'i'); tog('caps', 'caps'); tog('smallCaps', 'smallCaps'); tog('strike', 'strike'); tog('vanish', 'vanish');
  if (d('colour')) out.push(`<w:color w:val="${c.colour || 'auto'}"/>`);
  if (d('size')) out.push(`<w:sz w:val="${c.size}"/><w:szCs w:val="${c.size}"/>`);
  if (d('underline')) out.push(`<w:u w:val="${c.underline ? UNDERLINE[c.underline] || 'single' : 'none'}"/>`);
  if (d('iss')) out.push(`<w:vertAlign w:val="${c.iss === 1 ? 'superscript' : c.iss === 2 ? 'subscript' : 'baseline'}"/>`);
  return out.length ? `<w:rPr>${out.join('')}</w:rPr>` : '';
}

/** The lists as numbering: an abstract definition per list, an instance per override. */
function numberingXml(model, fontName) {
  const { lists, lfos } = model.lists;
  let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${W}>`;
  lists.forEach((l, i) => {
    xml += `<w:abstractNum w:abstractNumId="${i}"><w:multiLevelType w:val="${l.levels.length === 1 ? 'singleLevel' : 'hybridMultilevel'}"/>`;
    l.levels.forEach((lv, n) => {
      const text = lv.codes.map((c) => (c < 9 ? '%' + (c + 1) : String.fromCharCode(c))).join('');
      const font = lv.chp.font != null ? fontName(lv.chp.font) : null;
      const ind = [];
      if (lv.pap.left != null) ind.push(`w:left="${lv.pap.left}"`);
      if (lv.pap.firstLine != null) ind.push(lv.pap.firstLine < 0 ? `w:hanging="${-lv.pap.firstLine}"` : `w:firstLine="${lv.pap.firstLine}"`);
      xml += `<w:lvl w:ilvl="${n}"><w:start w:val="${lv.start || 1}"/><w:numFmt w:val="${NUMFMT[lv.nfc] || 'decimal'}"/>`
        + `<w:suff w:val="${lv.follow === 1 ? 'space' : lv.follow === 2 ? 'nothing' : 'tab'}"/>`
        + `<w:lvlText w:val="${esc(clean(text))}"/><w:lvlJc w:val="${JC[lv.jc] || 'left'}"/>`
        + (ind.length ? `<w:pPr><w:ind ${ind.join(' ')}/></w:pPr>` : '')
        + (font ? `<w:rPr><w:rFonts w:ascii="${esc(font)}" w:hAnsi="${esc(font)}" w:hint="default"/></w:rPr>` : '')
        + '</w:lvl>';
    });
    xml += '</w:abstractNum>';
  });
  const index = new Map(lists.map((l, i) => [l.lsid, i]));
  lfos.forEach((lfo, i) => {
    if (index.has(lfo.lsid)) xml += `<w:num w:numId="${i + 1}"><w:abstractNumId w:val="${index.get(lfo.lsid)}"/></w:num>`;
  });
  return xml + '</w:numbering>';
}
