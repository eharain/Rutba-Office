/**
 * Creating OOXML files from nothing.
 *
 * The rest of this package OPENS a package and edits it while preserving
 * everything it does not understand. That discipline is what makes it safe to
 * touch a customer's spreadsheet, and it has no answer at all for a file that
 * has never existed. This module is that answer: the parts, the relationships,
 * the content types and the style table for a workbook or a document built from
 * a specification rather than from bytes.
 *
 * WHERE IT CAME FROM, AND WHY THAT MATTERED. It lived in
 * `workspace/packages/erp-data-bindings` — a package named for the `erp.*`
 * data bindings — because it was written for the M0 spike, when its own comment
 * said "not intended to become one" and "production writes go through the
 * engine's Document Builder". Both premises died: we own the engine, and this
 * became the production creator. Every template in the gallery, every seeded
 * document and the workbook a `.csv` becomes when somebody opens it comes from
 * here.
 *
 * Leaving it there meant OOXML part knowledge lived in two packages, and the
 * roadmap had it recorded as "a duplicate writer to retire" — which was the
 * wrong name and implied the wrong fix. Deleting the "duplicate" would have
 * deleted the only code that can produce a valid file from nothing. The right
 * move was the one made here: bring it to where writing OOXML already lives, so
 * the two halves of the subject sit together and can stop drifting apart.
 *
 * The zip layer, the CRC table, the escaper and the column-letter function are
 * this package's own — see `OoxmlPackage.fromParts`. What remains below is
 * genuinely about generating parts.
 *
 * SCOPE, honestly. Exactly enough OOXML to carry values, styles, defined names,
 * a couple of drawings and content controls. It is not a spreadsheet library.
 * A file that needs more than this should be built by editing a real template
 * through `Workbook`, which is the path that preserves everything.
 */
import { Buffer } from 'node:buffer';
import { OoxmlPackage, esc } from './package.js';
import { indexToCol, unesc } from './workbook.js';

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/**
 * Parts in, bytes out.
 *
 * One line, and it is the whole point of this change: the zip layer, the CRC
 * table and the timestamp policy are somebody else's problem now.
 */
const packParts = (entries) => OoxmlPackage.fromParts(entries).write();

/**
 * Bytes in, parts out — the read side of the same trade.
 *
 * A `Map` of name to decompressed bytes, because that is the shape the readers
 * below already worked against. Going through `OoxmlPackage` rather than a
 * private zip reader means a file this cannot open now fails with the same
 * sentence the editor would give ("empty file — zero bytes, nothing to open"),
 * instead of a second opinion phrased differently.
 */
const partsOf = (buf) => {
  const pkg = OoxmlPackage.read(buf);
  return new Map(pkg.partNames().map((name) => [name, pkg.read(name)]));
};

/**
 * Column index -> letters, zero-based: 0 is "A".
 *
 * Re-exported rather than reimplemented. Both copies were already character-
 * for-character the same function under two names, which is the cheapest kind
 * of duplication to remove and the easiest to get wrong: writing this as
 * `indexToCol(n + 1)` on the assumption that the shared one counted from one
 * shifted every cell in every generated workbook one column to the right, and
 * nothing about the file looked broken — it opened, it just had an empty
 * column A. The round-trip comparison against the previous implementation is
 * what caught it.
 */
export const colName = indexToCol;

export const cellRef = (row, col) => colName(col) + (row + 1);

/**
 * Build an .xlsx.
 *
 * @param {object} spec
 * @param {Array<{name: string, rows: Array<Array<string|number|null>>}>} spec.sheets
 * @param {Array<{name: string, ref: string, comment?: string}>} [spec.definedNames]
 *        ref is a workbook-scoped A1 reference, e.g. "Data!$A$1:$C$12"
 *
 * A sheet may also carry `validations` — data-validation rules written into
 * its `<dataValidations>` block; the shape is documented at the emitter below.
 */

/**
 * Optional sheet drawings — a callout shape and a logo are ordinary furniture in
 * a real report, so a template must be able to ship them. Deliberately small:
 * a preset geometry with a fill and some words, or an embedded image. Anything
 * richer belongs in the editor, not in a template definition.
 *
 * @typedef {{kind:'shape', geometry?:string, fill?:string, text?:string, bold?:boolean,
 *            from:{col:number,row:number}, to:{col:number,row:number}}} TemplateShape
 * @typedef {{kind:'picture', bytes:Buffer, extension?:string,
 *            from:{col:number,row:number}, widthPx?:number, heightPx?:number}} TemplatePicture
 * @typedef {{kind:'chart', title?:string, name?:string,
 *            categories?:{ref?:string, values:Array<string>},
 *            series:Array<{name?:string, nameRef?:string, ref?:string, values:Array<number|null>}>,
 *            from:{col:number,row:number}, to:{col:number,row:number}}} TemplateChart
 */
const EMU_PX = 9525;

/**
 * One drawing anchor — chart frame, picture or shape — as the editor and the
 * creator both write it. Exported so the sheet editor's Insert surface and
 * the template creator share ONE writer; the shapes here are the ones the
 * Office gate proved real Excel accepts.
 */
export function drawingAnchorXml(d, relIdOf) {
  return drawingPartXml([d], relIdOf)
    .replace(/^[\s\S]*?<xdr:wsDr[^>]*>/, '')
    .replace(/<\/xdr:wsDr>$/, '');
}

const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const A14_NS = 'http://schemas.microsoft.com/office/drawing/2010/main';
const MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const CAMBRIA_MATH = '<a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/>';

/**
 * Insert → Equation on a sheet, as Excel writes one: a text box whose
 * paragraph holds the equation (`a14:m` around the OMML, its runs in
 * DrawingML's own run properties, in Cambria Math), inside
 * `mc:AlternateContent` — a reader without math gets the same box with the
 * equation's linear form as words. `d.omml` is OMML as `linearToOmml` makes
 * it; `d.linear` the linear form; `d.size` the type size in points.
 */
export function equationContentXml(d) {
  const sz = Math.round((d.size ?? 11) * 100);
  const math = String(d.omml || '')
    .replace(/<w:rPr>([\s\S]*?)<\/w:rPr>/g, (m, inner) => '<a:rPr lang="en-GB"' + (/<w:i\s*\/>|<w:i\b[^>]*w:val="(1|true|on)"/.test(inner) ? ' i="1"' : '') + '>' + CAMBRIA_MATH + '</a:rPr>')
    .replace(/<w:rPr\s*\/>/g, () => '<a:rPr lang="en-GB">' + CAMBRIA_MATH + '</a:rPr>')
    .replace(/^\s*<(m:oMathPara|m:oMath)\b(?![^>]*\bxmlns:m=)/, (m, tag) => '<' + tag + ' xmlns:m="' + MATH_NS + '"');
  const label = esc(d.name ?? 'Equation ' + d.id);
  const nv = '<xdr:nvSpPr><xdr:cNvPr id="' + d.id + '" name="' + label + '"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>';
  const spPr = '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></xdr:spPr>';
  const bodyPr = '<a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="none" rtlCol="0" anchor="ctr"><a:spAutoFit/></a:bodyPr><a:lstStyle/>';
  const choice = '<xdr:sp macro="" textlink="">' + nv + spPr + '<xdr:txBody>' + bodyPr
    + '<a:p><a:pPr algn="ctr"/><a14:m>' + math + '</a14:m><a:endParaRPr lang="en-GB" sz="' + sz + '"/></a:p></xdr:txBody></xdr:sp>';
  const fallback = '<xdr:sp macro="" textlink="">' + nv + spPr + '<xdr:txBody>' + bodyPr
    + '<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-GB" sz="' + sz + '"/><a:t>' + esc(d.linear ?? '') + '</a:t></a:r></a:p></xdr:txBody></xdr:sp>';
  return '<mc:AlternateContent xmlns:mc="' + MC_NS + '"><mc:Choice xmlns:a14="' + A14_NS + '" Requires="a14">' + choice
    + '</mc:Choice><mc:Fallback>' + fallback + '</mc:Fallback></mc:AlternateContent>';
}

/** A colour as DrawingML writes it, from `#RRGGBB` (or `RRGGBB`), with its alpha. */
const srgbXml = (hex, alpha = 1) => {
  const val = String(hex || '000000').replace('#', '').toUpperCase();
  return alpha < 1 ? '<a:srgbClr val="' + val + '"><a:alpha val="' + Math.round(alpha * 100000) + '"/></a:srgbClr>' : '<a:srgbClr val="' + val + '"/>';
};

/**
 * A run's WordArt look as run-property children, in the order
 * CT_TextCharacterProperties wants: the outline round the letters (`a:ln`),
 * the fill (or `a:noFill` for an outline alone), then a glow and an outer
 * shadow (`a:effectLst`). `run` is `{ color, noFill, outline: { width,
 * color }, textEffects: { glow: { radiusPt, color, alpha }, shadow: {
 * blurPx, distPx, dir, color, alpha } } }`, as Presentations reads a run.
 */
export function textLookXml(run = {}) {
  let out = '';
  if (run.outline?.color) out += '<a:ln w="' + Math.round(Math.max(0.25, Number(run.outline.width) || 1) * 12700) + '"><a:solidFill>' + srgbXml(run.outline.color) + '</a:solidFill></a:ln>';
  if (run.noFill) out += '<a:noFill/>';
  else if (run.color) out += '<a:solidFill>' + srgbXml(run.color) + '</a:solidFill>';
  const fx = run.textEffects || {};
  if (fx.glow || fx.shadow) {
    out += '<a:effectLst>';
    if (fx.glow) out += '<a:glow rad="' + Math.round((Number(fx.glow.radiusPt) || 5) * 12700) + '">' + srgbXml(fx.glow.color, fx.glow.alpha ?? 0.6) + '</a:glow>';
    if (fx.shadow) {
      const emu = (px) => Math.round((Number(px) || 0) * 0.75 * 12700);
      out += '<a:outerShdw blurRad="' + emu(fx.shadow.blurPx) + '" dist="' + emu(fx.shadow.distPx) + '" dir="' + Math.round((Number(fx.shadow.dir) || 0) * 60000) + '" algn="tl" rotWithShape="0">' + srgbXml(fx.shadow.color, fx.shadow.alpha ?? 0.4) + '</a:outerShdw>';
    }
    out += '</a:effectLst>';
  }
  return out;
}

/**
 * Insert → WordArt on a sheet, as Excel writes it: a text box with no fill
 * and no line, sized to its words, the words centred, big, and carrying
 * their WordArt look in the run. `d.text` may hold line breaks; `d.size` is
 * the type size in points; `d.run` the look (see `textLookXml`).
 */
export function wordArtContentXml(d) {
  const label = esc(d.name ?? 'TextBox ' + d.id);
  const sz = Math.round((d.size ?? 36) * 100);
  const rPr = '<a:rPr lang="en-US" sz="' + sz + '" b="0" cap="none" spc="0">' + textLookXml(d.run) + '</a:rPr>';
  const paras = String(d.text ?? '').split(/\r?\n/).map((line) => '<a:p><a:pPr algn="ctr"/>'
    + (line ? '<a:r>' + rPr + '<a:t>' + esc(line) + '</a:t></a:r>' : '<a:endParaRPr lang="en-US" sz="' + sz + '"/>') + '</a:p>').join('');
  return '<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="' + d.id + '" name="' + label + '"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>'
    + '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></xdr:spPr>'
    + '<xdr:txBody><a:bodyPr wrap="none" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="ctr"><a:spAutoFit/></a:bodyPr><a:lstStyle/>' + paras + '</xdr:txBody></xdr:sp>';
}

/**
 * Draw → Pen and Highlighter on a sheet: one stroke as a freeform shape
 * named "Ink N" — `d.points` (in sheet pixels, relative to the drawing's
 * box, which is `d.box`) joined by a smooth path in the pen's colour and
 * width (points), see-through at `d.alpha`, with round or square ends.
 */
export function inkContentXml(d) {
  const EMU = 9525;
  const cx = Math.max(1, Math.round(d.box.width * EMU));
  const cy = Math.max(1, Math.round(d.box.height * EMU));
  const pts = d.points.length > 1 ? d.points : [d.points[0], [d.points[0][0] + 0.5, d.points[0][1]]];
  const at = (p) => '<a:pt x="' + Math.round(p[0] * EMU) + '" y="' + Math.round(p[1] * EMU) + '"/>';
  let path = '<a:moveTo>' + at(pts[0]) + '</a:moveTo>';
  for (let i = 1; i < pts.length - 1; i++) path += '<a:quadBezTo>' + at(pts[i]) + at([(pts[i][0] + pts[i + 1][0]) / 2, (pts[i][1] + pts[i + 1][1]) / 2]) + '</a:quadBezTo>';
  path += '<a:lnTo>' + at(pts[pts.length - 1]) + '</a:lnTo>';
  return '<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="' + d.id + '" name="' + esc(d.name) + '"/><xdr:cNvSpPr/></xdr:nvSpPr>'
    + '<xdr:spPr><a:xfrm><a:off x="' + Math.round(d.box.x * EMU) + '" y="' + Math.round(d.box.y * EMU) + '"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm>'
    + '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="' + cx + '" h="' + cy + '" fill="none" extrusionOk="0">' + path + '</a:path></a:pathLst></a:custGeom>'
    + '<a:noFill/><a:ln w="' + Math.round(Math.max(0.25, Number(d.width) || 1) * 12700) + '" cap="' + (d.cap === 'sq' ? 'sq' : 'rnd') + '"><a:solidFill>' + srgbXml(d.color, d.alpha ?? 1) + '</a:solidFill><a:round/></a:ln></xdr:spPr></xdr:sp>';
}

const DIAGRAM_NAMES = { roundRect: 'Rectangle: Rounded Corners', rect: 'Rectangle', ellipse: 'Oval', chevron: 'Arrow: Chevron', rightArrow: 'Arrow: Right', line: 'Straight Connector' };

/**
 * Insert → SmartArt on a sheet: one group (`xdr:grpSp`) of preset shapes —
 * `d.shapes` in sheet pixels, in drawing order: a node filled (a theme
 * colour when it names one) with its words centred in white at their size,
 * an arrow filled, a ring or a line edged — the group named `d.name`, its
 * members numbered on from its id. Excel opens it as a group of shapes.
 */
export function diagramContentXml(d) {
  const EMU = 9525;
  const xs = d.shapes.map((s) => s.x);
  const ys = d.shapes.map((s) => s.y);
  const x0 = Math.min(...xs);
  const y0 = Math.min(...ys);
  const cx = Math.max(1, Math.round((Math.max(...d.shapes.map((s) => s.x + s.w)) - x0) * EMU));
  const cy = Math.max(1, Math.round((Math.max(...d.shapes.map((s) => s.y + s.h)) - y0) * EMU));
  const place = '<a:off x="' + Math.round(x0 * EMU) + '" y="' + Math.round(y0 * EMU) + '"/><a:ext cx="' + cx + '" cy="' + cy + '"/>';
  const kids = d.shapes.map((s, i) => {
    const id = d.id + 1 + i;
    const fill = s.fill ? '<a:solidFill>' + (s.scheme ? '<a:schemeClr val="' + esc(s.scheme) + '"/>' : srgbXml(s.fill)) + '</a:solidFill>' : '<a:noFill/>';
    const line = s.line ? '<a:ln w="' + Math.round((Number(s.lineWidth) || 1) * 12700) + '"><a:solidFill>' + srgbXml(s.line) + '</a:solidFill></a:ln>' : '<a:ln><a:noFill/></a:ln>';
    const sz = Math.round((Number(s.size) || 14) * 100);
    const words = s.text ? '<a:r><a:rPr lang="en-US" sz="' + sz + '"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>' + esc(s.text) + '</a:t></a:r>' : '';
    return '<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="' + id + '" name="' + esc((DIAGRAM_NAMES[s.preset] || 'Shape') + ' ' + id) + '"/><xdr:cNvSpPr/></xdr:nvSpPr>'
      + '<xdr:spPr><a:xfrm><a:off x="' + Math.round(s.x * EMU) + '" y="' + Math.round(s.y * EMU) + '"/><a:ext cx="' + Math.round(s.w * EMU) + '" cy="' + Math.round(s.h * EMU) + '"/></a:xfrm>'
      + '<a:prstGeom prst="' + esc(s.preset) + '"><a:avLst/></a:prstGeom>' + fill + line + '</xdr:spPr>'
      + '<xdr:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/>' + words + '</a:p></xdr:txBody></xdr:sp>';
  }).join('');
  return '<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="' + d.id + '" name="' + esc(d.name ?? 'Group ' + d.id) + '"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>'
    + '<xdr:grpSpPr><a:xfrm>' + place + place.replace('<a:off', '<a:chOff').replace('<a:ext', '<a:chExt') + '</a:xfrm></xdr:grpSpPr>' + kids + '</xdr:grpSp>';
}

function drawingPartXml(drawings, relIdOf) {
  const anchors = drawings.map((d) => {
    // A marker's offsets into its cell (EMU), when the drawing does not start at the cell's corner.
    const from = '<xdr:from><xdr:col>' + d.from.col + '</xdr:col><xdr:colOff>' + Math.round(d.from.colOff || 0) + '</xdr:colOff>'
      + '<xdr:row>' + d.from.row + '</xdr:row><xdr:rowOff>' + Math.round(d.from.rowOff || 0) + '</xdr:rowOff></xdr:from>';
    if (d.kind === 'chart') {
      const to = '<xdr:to><xdr:col>' + d.to.col + '</xdr:col><xdr:colOff>0</xdr:colOff>'
        + '<xdr:row>' + d.to.row + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>';
      return '<xdr:twoCellAnchor>' + from + to
        + '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>'
        + '<xdr:cNvPr id="' + d.id + '" name="' + esc(d.name ?? 'Chart') + '"/>'
        + '<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'
        + '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
        + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
        + '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="' + relIdOf(d) + '"/>'
        + '</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>';
    }
    if (d.kind === 'picture') {
      const ext = '<xdr:ext cx="' + Math.round((d.widthPx ?? 96) * EMU_PX)
        + '" cy="' + Math.round((d.heightPx ?? 48) * EMU_PX) + '"/>';
      return '<xdr:oneCellAnchor>' + from + ext
        + '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="' + d.id + '" name="' + esc(d.name ?? 'Picture') + '"/>'
        + '<xdr:cNvPicPr/></xdr:nvPicPr>'
        + '<xdr:blipFill><a:blip r:embed="' + relIdOf(d) + '"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
        + '<xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>';
    }
    const to = '<xdr:to><xdr:col>' + d.to.col + '</xdr:col><xdr:colOff>' + Math.round(d.to.colOff || 0) + '</xdr:colOff>'
      + '<xdr:row>' + d.to.row + '</xdr:row><xdr:rowOff>' + Math.round(d.to.rowOff || 0) + '</xdr:rowOff></xdr:to>';
    if (d.kind === 'equation') return '<xdr:twoCellAnchor>' + from + to + equationContentXml(d) + '<xdr:clientData/></xdr:twoCellAnchor>';
    if (d.kind === 'wordart') return '<xdr:twoCellAnchor>' + from + to + wordArtContentXml(d) + '<xdr:clientData/></xdr:twoCellAnchor>';
    if (d.kind === 'ink') return '<xdr:twoCellAnchor editAs="oneCell">' + from + to + inkContentXml(d) + '<xdr:clientData/></xdr:twoCellAnchor>';
    if (d.kind === 'diagram') return '<xdr:twoCellAnchor editAs="oneCell">' + from + to + diagramContentXml(d) + '<xdr:clientData/></xdr:twoCellAnchor>';
    // Ink to Shape: a shape with no fill and a line in the pen's colour.
    const line = d.line?.color ? '<a:ln w="' + Math.round((Number(d.line.width) || 1) * 12700) + '"><a:solidFill>' + srgbXml(d.line.color) + '</a:solidFill></a:ln>' : '';
    const fill = d.fill === 'none' ? '<a:noFill/>' + line : (d.fill
      ? '<a:solidFill>' + (/^[0-9A-Fa-f]{6}$/.test(d.fill)
        ? '<a:srgbClr val="' + d.fill.toUpperCase() + '"/>'
        : '<a:schemeClr val="' + esc(d.fill) + '"/>') + '</a:solidFill>'
      : '') + (d.fill === 'none' ? '' : line);
    const body = d.text
      ? '<xdr:txBody><a:bodyPr/><a:p><a:r>'
        + '<a:rPr lang="en-GB"' + (d.bold ? ' b="1"' : '') + (d.textSize ? ' sz="' + Math.round(d.textSize * 100) + '"' : '') + '/>'
        + '<a:t>' + esc(d.text) + '</a:t></a:r></a:p></xdr:txBody>'
      : '<xdr:txBody><a:bodyPr/><a:p/></xdr:txBody>';
    return '<xdr:twoCellAnchor>' + from + to
      + '<xdr:sp macro="" textlink=""><xdr:nvSpPr>'
      + '<xdr:cNvPr id="' + d.id + '" name="' + esc(d.name ?? 'Shape') + '"/><xdr:cNvSpPr/></xdr:nvSpPr>'
      + '<xdr:spPr><a:prstGeom prst="' + esc(d.geometry ?? 'rect') + '"><a:avLst/></a:prstGeom>'
      + fill + '</xdr:spPr>' + body
      + '</xdr:sp><xdr:clientData/></xdr:twoCellAnchor>';
  }).join('');

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
    + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    + anchors + '</xdr:wsDr>';
}

/**
 * A chart part — one clustered column chart, per ECMA-376 §21.2.
 *
 * The one chart type a template has needed. The `<c:f>` references keep the
 * chart honest in Excel (edit a figure and it re-plots); the caches are what
 * every reader paints before recalculating, ours included, so the template
 * author supplies cached values that MATCH the sheet's figures — the tests
 * hold that agreement in place.
 */
export function chartPartXml(chart) {
  const CAT_AX = 100010001;
  const VAL_AX = 100010002;
  // `literal: true` writes c:strLit/c:numLit — data that LIVES in the chart,
  // for a chart in a document with no workbook behind it. The default stays
  // the ref-plus-cache shape a template's sheet chart wants, byte for byte.
  const literal = Boolean(chart.literal);
  const kind = chart.kind ?? 'column';
  const pts = (list) => list.map((v, i) => (v === null || v === undefined || v === ''
    ? '' : '<c:pt idx="' + i + '"><c:v>' + esc(String(v)) + '</c:v></c:pt>')).join('');
  const strRef = (ref, list) => (literal
    ? '<c:strLit><c:ptCount val="' + list.length + '"/>' + pts(list) + '</c:strLit>'
    : '<c:strRef>' + (ref ? '<c:f>' + esc(ref) + '</c:f>' : '')
      + '<c:strCache><c:ptCount val="' + list.length + '"/>' + pts(list) + '</c:strCache></c:strRef>');
  const numRef = (ref, list) => (literal
    ? '<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="' + list.length + '"/>'
      + pts(list) + '</c:numLit>'
    : '<c:numRef>' + (ref ? '<c:f>' + esc(ref) + '</c:f>' : '')
      + '<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="' + list.length + '"/>'
      + pts(list) + '</c:numCache></c:numRef>');

  // A scatter (XY) chart plots numbers against numbers: the categories are
  // its X values (xVal) and each series its Y values (yVal). Excel's three
  // looks — markers only, straight lines, smooth lines, each with markers —
  // are scatterStyle and each series' own line: "markers only" is the line
  // style with the series line switched off, which is how Excel writes it.
  const scatter = kind === 'scatter';
  const look = chart.scatterStyle === 'lines' || chart.scatterStyle === 'smooth' ? chart.scatterStyle : 'markers';
  const tx = (s, i) => '<c:tx>' + (s.nameRef
    ? strRef(s.nameRef, [s.name ?? 'Series ' + (i + 1)])
    : '<c:v>' + esc(s.name ?? 'Series ' + (i + 1)) + '</c:v>') + '</c:tx>';
  const series = chart.series.map((s, i) => (scatter
    ? '<c:ser>'
      + '<c:idx val="' + i + '"/><c:order val="' + i + '"/>' + tx(s, i)
      + (look === 'markers' ? '<c:spPr><a:ln w="19050" cap="rnd"><a:noFill/><a:round/></a:ln></c:spPr>' : '')
      + '<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>'
      + (chart.categories ? '<c:xVal>' + numRef(chart.categories.ref, chart.categories.values) + '</c:xVal>' : '')
      + '<c:yVal>' + numRef(s.ref, s.values ?? []) + '</c:yVal>'
      + '<c:smooth val="' + (look === 'smooth' ? 1 : 0) + '"/>'
      + '</c:ser>'
    : '<c:ser>'
      + '<c:idx val="' + i + '"/><c:order val="' + i + '"/>' + tx(s, i)
      + (s.line ? '<c:spPr><a:ln w="' + Math.round(s.line.width ?? 28575) + '" cap="rnd"><a:solidFill><a:srgbClr val="' + esc(s.line.colour) + '"/></a:solidFill>'
        + (s.line.dash ? '<a:prstDash val="' + esc(s.line.dash) + '"/>' : '') + '<a:round/></a:ln></c:spPr>' : '')
      + (s.marker === false ? '<c:marker><c:symbol val="none"/></c:marker>' : '')
      + (chart.categories ? '<c:cat>' + strRef(chart.categories.ref, chart.categories.values) + '</c:cat>' : '')
      + '<c:val>' + numRef(s.ref, s.values ?? []) + '</c:val>'
      + '</c:ser>')).join('');

  const title = chart.title
    ? '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>' + esc(chart.title)
      + '</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>'
    : '';

  const axRefs = '<c:axId val="' + CAT_AX + '"/><c:axId val="' + VAL_AX + '"/>';
  // Pie and doughnut have no axes; everything else shares the two-axis frame.
  // The kind list matches Studio's vocabulary, which is the convergence the
  // roadmap asks for: column, bar, line, area, pie, doughnut.
  const plot = scatter
    ? '<c:scatterChart><c:scatterStyle val="' + (look === 'smooth' ? 'smoothMarker' : 'lineMarker') + '"/><c:varyColors val="0"/>' + series + axRefs + '</c:scatterChart>'
    : kind === 'pie'
    ? '<c:pieChart><c:varyColors val="1"/>' + series + '<c:firstSliceAng val="0"/></c:pieChart>'
    : kind === 'doughnut'
      ? '<c:doughnutChart><c:varyColors val="1"/>' + series + '<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart>'
      : kind === 'line'
        ? '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>' + series + '<c:marker val="1"/>' + axRefs + '</c:lineChart>'
        : kind === 'area'
          ? '<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>' + series + axRefs + '</c:areaChart>'
          : '<c:barChart><c:barDir val="' + (kind === 'bar' ? 'bar' : 'col') + '"/><c:grouping val="clustered"/><c:varyColors val="0"/>' + series + axRefs + '</c:barChart>';
  // A scatter's X axis is a value axis too: two valAx, the X one at the foot.
  const axes = scatter ? (
    '<c:valAx><c:axId val="' + CAT_AX + '"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
    + '<c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:crossAx val="' + VAL_AX + '"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>'
    + '<c:valAx><c:axId val="' + VAL_AX + '"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
    + '<c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:crossAx val="' + CAT_AX + '"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>'
  ) : (kind === 'pie' || kind === 'doughnut') ? '' : (
    '<c:catAx><c:axId val="' + CAT_AX + '"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
    + '<c:delete val="0"/><c:axPos val="b"/><c:crossAx val="' + VAL_AX + '"/></c:catAx>'
    + '<c:valAx><c:axId val="' + VAL_AX + '"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
    + '<c:delete val="0"/><c:axPos val="l"/><c:crossAx val="' + CAT_AX + '"/></c:valAx>');

  return XML_DECL
    + '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
    + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    + '<c:chart>' + title + '<c:plotArea><c:layout/>'
    + plot
    + axes
    + '</c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>';
}


/**
 * A tiny style table, so a generated workbook does not arrive looking damaged.
 *
 * Deliberately not a style engine. A template says which ranges are a header,
 * a total or a note; this turns that into the four cellXfs entries a workbook
 * needs and nothing more. Anything richer belongs in the editor, where a user
 * can see what they are doing.
 *
 * @typedef {{bold?:boolean, fill?:string, align?:string, border?:boolean,
 *            wrap?:boolean, size?:number, colour?:string, numFmt?:string}} TemplateStyle
 */
const STYLE_KEY = (st) => JSON.stringify([
  Boolean(st.bold), st.fill ?? null, st.align ?? null, Boolean(st.border),
  Boolean(st.wrap), st.size ?? null, st.colour ?? null, st.numFmt ?? null,
]);

function buildStyleTable(styleList) {
  // Index 0 must stay the default: a cell with no s= attribute means "entry 0",
  // so shifting it would restyle every unstyled cell in the file.
  const fonts = ['<font><sz val="11"/><color theme="1"/><name val="Calibri"/></font>'];
  const fills = [
    '<fill><patternFill patternType="none"/></fill>',
    '<fill><patternFill patternType="gray125"/></fill>',
  ];
  const borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
  const numFmts = [];
  const xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
  const indexByKey = new Map();

  const BORDER = '<border>'
    + '<left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right>'
    + '<top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom>'
    + '</border>';

  for (const st of styleList) {
    const key = STYLE_KEY(st);
    if (indexByKey.has(key)) continue;

    let fontId = 0;
    if (st.bold || st.size || st.colour) {
      const colour = st.colour ? '<color rgb="FF' + String(st.colour).replace(/^#/, '').toUpperCase() + '"/>' : '<color theme="1"/>';
      fonts.push('<font>' + (st.bold ? '<b/>' : '') + '<sz val="' + (st.size ?? 11) + '"/>'
        + colour + '<name val="Calibri"/></font>');
      fontId = fonts.length - 1;
    }

    let fillId = 0;
    if (st.fill) {
      fills.push('<fill><patternFill patternType="solid"><fgColor rgb="FF'
        + String(st.fill).replace(/^#/, '').toUpperCase() + '"/><bgColor indexed="64"/></patternFill></fill>');
      fillId = fills.length - 1;
    }

    let borderId = 0;
    if (st.border) {
      borders.push(BORDER);
      borderId = borders.length - 1;
    }

    let numFmtId = 0;
    if (st.numFmt) {
      numFmtId = 164 + numFmts.length;
      numFmts.push('<numFmt numFmtId="' + numFmtId + '" formatCode="' + esc(st.numFmt) + '"/>');
    }

    const alignment = (st.align || st.wrap)
      ? '<alignment' + (st.align ? ' horizontal="' + esc(st.align) + '"' : '')
        + ' vertical="center"' + (st.wrap ? ' wrapText="1"' : '') + '/>'
      : '';

    xfs.push('<xf numFmtId="' + numFmtId + '" fontId="' + fontId + '" fillId="' + fillId
      + '" borderId="' + borderId + '" xfId="0"'
      + (numFmtId ? ' applyNumberFormat="1"' : '')
      + (fontId ? ' applyFont="1"' : '') + (fillId ? ' applyFill="1"' : '')
      + (borderId ? ' applyBorder="1"' : '') + (alignment ? ' applyAlignment="1"' : '')
      + (alignment ? '>' + alignment + '</xf>' : '/>'));
    indexByKey.set(key, xfs.length - 1);
  }

  const xml = XML_DECL
    + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    + (numFmts.length ? '<numFmts count="' + numFmts.length + '">' + numFmts.join('') + '</numFmts>' : '')
    + '<fonts count="' + fonts.length + '">' + fonts.join('') + '</fonts>'
    + '<fills count="' + fills.length + '">' + fills.join('') + '</fills>'
    + '<borders count="' + borders.length + '">' + borders.join('') + '</borders>'
    + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    + '<cellXfs count="' + xfs.length + '">' + xfs.join('') + '</cellXfs>'
    + '</styleSheet>';

  return { xml, indexByKey, used: xfs.length > 1 };
}

/** "A1:C1" or "A1" -> a predicate over row/col indices. */
function rangeMatcher(ref) {
  const parse = (r) => {
    const m = /^([A-Z]+)(\d+)$/.exec(String(r).trim());
    if (!m) return null;
    let col = 0;
    for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
    return { row: Number(m[2]) - 1, col: col - 1 };
  };
  const [a, b] = String(ref).split(':');
  const from = parse(a);
  const to = parse(b ?? a);
  if (!from || !to) return () => false;
  const top = Math.min(from.row, to.row);
  const bottom = Math.max(from.row, to.row);
  const left = Math.min(from.col, to.col);
  const right = Math.max(from.col, to.col);
  return (row, col) => row >= top && row <= bottom && col >= left && col <= right;
}

export function buildXlsx({ sheets, definedNames = [] }) {
  if (!sheets?.length) throw new Error('at least one sheet required');

  // Every distinct style across every sheet, deduplicated into one table. A
  // header repeated on three sheets is one cellXfs entry, not three.
  const declared = [];
  for (const sheet of sheets) {
    for (const [, style] of Object.entries(sheet.styles ?? {})) declared.push(style);
  }
  const styleTable = buildStyleTable(declared);
  const stylesFor = (sheet) => Object.entries(sheet.styles ?? {}).map(([ref, style]) => ({
    matches: rangeMatcher(ref),
    index: styleTable.indexByKey.get(STYLE_KEY(style)) ?? 0,
  }));

  const sheetXml = (sheet) => {
    const ranges = stylesFor(sheet);
    // Last declaration wins, so a total row can override the body style it
    // sits inside without the caller having to carve the range up.
    const styleAt = (row, col) => {
      let found = 0;
      for (const r of ranges) if (r.matches(row, col)) found = r.index;
      return found;
    };
    const rows = sheet.rows
      .map((cells, r) => {
        const body = cells
          .map((v, c) => {
            const styleIndex = styleAt(r, c);
            const sAttr = styleIndex ? ' s="' + styleIndex + '"' : '';
            // A STYLED but empty cell is still written: a shaded header row with
            // a blank column in it must keep its band unbroken.
            if (v === null || v === undefined || v === '') {
              return styleIndex ? '<c r="' + cellRef(r, c) + '"' + sAttr + '/>' : '';
            }
            const ref = cellRef(r, c);
            if (typeof v === 'number' && Number.isFinite(v)) {
              return '<c r="' + ref + '"' + sAttr + '><v>' + v + '</v></c>';
            }
            const s = String(v);
            if (s.startsWith('=')) {
              return '<c r="' + ref + '"' + sAttr + '><f>' + esc(s.slice(1)) + '</f></c>';
            }
            return '<c r="' + ref + '"' + sAttr + ' t="inlineStr"><is><t xml:space="preserve">' + esc(s) + '</t></is></c>';
          })
          .join('');
        return body ? '<row r="' + (r + 1) + '">' + body + '</row>' : '';
      })
      .join('');
    const drawingRef = sheet.drawings?.length
      ? '<drawing r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'
      : '';
    const merges = sheet.merges?.length
      ? '<mergeCells count="' + sheet.merges.length + '">'
        + sheet.merges.map((m) => '<mergeCell ref="' + esc(m) + '"/>').join('')
        + '</mergeCells>'
      : '';
    // Data validation rules: {type, sqref, formula1?, formula2?, operator?,
    // allowBlank?, errorStyle?, errorTitle?, error?, promptTitle?, prompt?}.
    // A list's inline members keep their quotes: formula1: '"Yes,No"'.
    const validations = sheet.validations?.length
      ? '<dataValidations count="' + sheet.validations.length + '">'
        + sheet.validations.map((v) => {
          let attrs = ' type="' + esc(v.type) + '"';
          if (v.operator) attrs += ' operator="' + esc(v.operator) + '"';
          if (v.errorStyle) attrs += ' errorStyle="' + esc(v.errorStyle) + '"';
          if (v.allowBlank !== false) attrs += ' allowBlank="1"';
          attrs += ' showInputMessage="1" showErrorMessage="1"';
          for (const key of ['errorTitle', 'error', 'promptTitle', 'prompt']) {
            if (v[key]) attrs += ' ' + key + '="' + esc(v[key]) + '"';
          }
          attrs += ' sqref="' + esc(v.sqref) + '"';
          return '<dataValidation' + attrs + '>'
            + (v.formula1 !== undefined ? '<formula1>' + esc(v.formula1) + '</formula1>' : '')
            + (v.formula2 !== undefined ? '<formula2>' + esc(v.formula2) + '</formula2>' : '')
            + '</dataValidation>';
        }).join('')
        + '</dataValidations>'
      : '';
    // Schema order is fixed: sheetData, then mergeCells, then dataValidations,
    // then drawing. Out of order and Excel rejects the file rather than
    // ignoring the element.
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetData>' + rows + '</sheetData>' + merges + validations + drawingRef + '</worksheet>'
    );
  };

  const definedNamesXml = definedNames.length
    ? '<definedNames>' +
      definedNames
        .map((d) => '<definedName name="' + esc(d.name) + '">' + esc(d.ref) + '</definedName>')
        .join('') +
      '</definedNames>'
    : '';

  const workbookXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheets>' +
    sheets
      .map((s, i) => '<sheet name="' + esc(s.name) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>')
      .join('') +
    '</sheets>' +
    definedNamesXml +
    '</workbook>';

  const workbookRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    sheets
      .map(
        (s, i) =>
          '<Relationship Id="rId' + (i + 1) +
          '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"' +
          ' Target="worksheets/sheet' + (i + 1) + '.xml"/>',
      )
      .join('') +
    (styleTable.used
      ? '<Relationship Id="rId' + (sheets.length + 1) + '" Type="http://schemas.openxmlformats.org/'
        + 'officeDocument/2006/relationships/styles" Target="styles.xml"/>'
      : '') +
    '</Relationships>';

  const rootRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
    '</Relationships>';

  // Drawings, their relationships and any media they embed. Ids are assigned
  // here rather than by the caller so a template cannot collide with itself.
  const drawingEntries = [];
  const chartEntries = [];
  const mediaEntries = [];
  const sheetRelEntries = [];
  const mediaExtensions = new Set();
  let shapeId = 1;
  let mediaIndex = 0;
  let chartIndex = 0;

  sheets.forEach((sheetDef, i) => {
    const drawings = sheetDef.drawings ?? [];
    if (!drawings.length) return;
    const n = i + 1;
    const rels = [];
    const withIds = drawings.map((d) => {
      const item = { ...d, id: (shapeId += 1) };
      if (d.kind === 'picture' && d.bytes) {
        mediaIndex += 1;
        const extension = (d.extension ?? 'png').toLowerCase();
        mediaExtensions.add(extension);
        const target = 'xl/media/image' + mediaIndex + '.' + extension;
        mediaEntries.push({ name: target, data: d.bytes, method: 0 });
        item.relId = 'rId' + (rels.length + 1);
        rels.push('<Relationship Id="' + item.relId + '" Type="http://schemas.openxmlformats.org/'
          + 'officeDocument/2006/relationships/image" Target="../media/image' + mediaIndex + '.' + extension + '"/>');
      }
      if (d.kind === 'chart') {
        chartIndex += 1;
        // `kind` on a drawing says which KIND OF DRAWING it is; `kind` on a
        // chart says which plot to draw. They collide on this one object, and
        // for a sheet-anchored chart the drawing wins - which quietly made
        // every kind but `column` unreachable from `buildXlsx`, even though
        // chartPartXml has drawn six since it was written. `chartKind` is the
        // plot type, and omitting it keeps the previous bytes exactly.
        chartEntries.push({
          name: 'xl/charts/chart' + chartIndex + '.xml',
          data: chartPartXml({ ...d, kind: d.chartKind ?? 'column' }),
        });
        item.relId = 'rId' + (rels.length + 1);
        rels.push('<Relationship Id="' + item.relId + '" Type="http://schemas.openxmlformats.org/'
          + 'officeDocument/2006/relationships/chart" Target="../charts/chart' + chartIndex + '.xml"/>');
      }
      return item;
    });

    drawingEntries.push({
      name: 'xl/drawings/drawing' + n + '.xml',
      data: drawingPartXml(withIds, (d) => d.relId),
    });
    if (rels.length) {
      drawingEntries.push({
        name: 'xl/drawings/_rels/drawing' + n + '.xml.rels',
        data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
          + rels.join('') + '</Relationships>',
      });
    }
    sheetRelEntries.push({
      name: 'xl/worksheets/_rels/sheet' + n + '.xml.rels',
      data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        + 'relationships/drawing" Target="../drawings/drawing' + n + '.xml"/></Relationships>',
    });
  });

  // Content types are declared last, because a drawing or a media file adds to
  // them and a part that is not declared is a corrupt package, not a missing
  // feature — Excel refuses the whole file.
  const MEDIA_CONTENT_TYPES = {
    png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif',
    bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp',
    emf: 'image/x-emf', wmf: 'image/x-wmf',
  };
  const contentTypes =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    [...mediaExtensions]
      .map((ext) => '<Default Extension="' + ext + '" ContentType="'
        + (MEDIA_CONTENT_TYPES[ext] ?? 'application/octet-stream') + '"/>')
      .join('') +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    sheets
      .map(
        (s, i) =>
          '<Override PartName="/xl/worksheets/sheet' + (i + 1) +
          '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
      )
      .join('') +
    drawingEntries
      .filter((e) => e.name.endsWith('.xml') && !e.name.includes('_rels'))
      .map((e) => '<Override PartName="/' + e.name
        + '" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>')
      .join('') +
    chartEntries
      .map((e) => '<Override PartName="/' + e.name
        + '" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>')
      .join('') +
    (styleTable.used
      ? '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-'
        + 'officedocument.spreadsheetml.styles+xml"/>'
      : '') +
    '</Types>';

  const entries = [
    { name: '[Content_Types].xml', data: contentTypes },
    { name: '_rels/.rels', data: rootRels },
    { name: 'xl/workbook.xml', data: workbookXml },
    { name: 'xl/_rels/workbook.xml.rels', data: workbookRels },
    ...sheets.map((s, i) => ({ name: 'xl/worksheets/sheet' + (i + 1) + '.xml', data: sheetXml(s) })),
    ...sheetRelEntries,
    ...drawingEntries,
    ...chartEntries,
    ...mediaEntries,
    ...(styleTable.used ? [{ name: 'xl/styles.xml', data: styleTable.xml }] : []),
  ];
  return packParts(entries);
}

// The standard style catalogue lives in docstyles.js — one definition serving
// both this creator and the editor's ensureParagraphStyles.
import { STANDARD_STYLES_XML as STANDARD_DOC_STYLES } from './docstyles.js';

/**
 * One paragraph.
 *
 * A plain string produces exactly the XML this function has always produced,
 * byte for byte — the goldens pin it, and a template that wants formatting
 * should not change what a plain one writes. An object adds only what it names:
 *
 *   { text, runs, style, align, bold, italic, underline, size, colour, font }
 */
const HYPERLINK_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

/**
 * A table in a generated document.
 *
 * `rows` is the whole grid, first row treated as the header when `header` is
 * set — which is the shape a Markdown table, a CSV and a query result all
 * already have, so nothing has to be rearranged to get here.
 *
 * The letter builder has its own copy of this with a different signature
 * (`columns`, a separate `header` row) whose output is byte-pinned by the
 * confirmation-letter goldens. Merging them would be a change to those files
 * for no gain, so they stay apart.
 */
function tableXml({ rows = [], header = false, align = [], columns: widths = null }) {
  if (!rows.length) return '';
  const columns = Math.max(...rows.map((r) => r.length));
  // Column widths in twips when the caller gives them (the table is then a
  // fixed width, their sum); otherwise equal shares of the text width.
  const widthOf = (i) => (widths && widths[i] > 0 ? Math.round(widths[i]) : Math.round(9360 / columns));
  const grid = '<w:tblGrid>' + Array.from({ length: columns }, (_, i) => '<w:gridCol w:w="' + widthOf(i) + '"/>').join('') + '</w:tblGrid>';
  const tblW = widths
    ? '<w:tblW w:w="' + Array.from({ length: columns }, (_, i) => widthOf(i)).reduce((a, b) => a + b, 0) + '" w:type="dxa"/>'
    : '<w:tblW w:w="5000" w:type="pct"/>';

  const cell = (text, i, isHeader) => {
    const shading = isHeader ? '<w:shd w:val="clear" w:fill="D9E2F3"/>' : '';
    const jc = align[i] ? '<w:pPr><w:jc w:val="' + esc(align[i] === 'center' ? 'center' : align[i]) + '"/></w:pPr>' : '';
    const rPr = isHeader ? '<w:rPr><w:b/></w:rPr>' : '';
    return '<w:tc><w:tcPr><w:tcW w:w="' + widthOf(i) + '" w:type="dxa"/>' + shading + '</w:tcPr>'
      + '<w:p>' + jc + '<w:r>' + rPr + '<w:t xml:space="preserve">' + esc(text ?? '') + '</w:t></w:r></w:p></w:tc>';
  };

  const row = (cells, isHeader) => '<w:tr>'
    + (isHeader ? '<w:trPr><w:tblHeader/></w:trPr>' : '')
    + Array.from({ length: columns }, (_, i) => cell(cells[i], i, isHeader)).join('')
    + '</w:tr>';

  return '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/>' + tblW
    + '<w:tblBorders><w:top w:val="single" w:sz="8" w:color="7F7F7F"/>'
    + '<w:bottom w:val="single" w:sz="8" w:color="7F7F7F"/>'
    + '<w:insideH w:val="single" w:sz="4" w:color="D9D9D9"/>'
    + '<w:insideV w:val="single" w:sz="4" w:color="D9D9D9"/></w:tblBorders></w:tblPr>'
    + grid
    + rows.map((r, n) => row(r, header && n === 0)).join('')
    + '</w:tbl>';
}

/**
 * @param {object|string} p the paragraph
 * @param {Map<string,string>} [links] collects url → relationship id. Passing
 *   it is what turns a run's `link` into a real hyperlink; without it a link is
 *   ignored, which is how the header and footer builders keep their exact
 *   byte-pinned output.
 */
function paragraphXml(p, links = null) {
  if (typeof p === 'string' || p == null) {
    return '<w:p><w:r><w:t xml:space="preserve">' + esc(p ?? '') + '</w:t></w:r></w:p>';
  }
  if (p.table) return tableXml(p.table);

  const pPrBits = [];
  if (p.style) pPrBits.push('<w:pStyle w:val="' + esc(p.style) + '"/>');
  // In the order the schema wants them: keepNext, keepLines, pageBreakBefore, then jc.
  if (p.keepNext) pPrBits.push('<w:keepNext/>');
  if (p.keepLines) pPrBits.push('<w:keepLines/>');
  if (p.pageBreakBefore) pPrBits.push('<w:pageBreakBefore/>');
  if (p.align) {
    const jc = { left: 'left', center: 'center', centre: 'center', right: 'right', justify: 'both' }[p.align];
    if (jc) pPrBits.push('<w:jc w:val="' + jc + '"/>');
  }
  const pPr = pPrBits.length ? '<w:pPr>' + pPrBits.join('') + '</w:pPr>' : '';

  const runXml = (r) => {
    const bits = [];
    if (r.bold) bits.push('<w:b/>');
    if (r.italic) bits.push('<w:i/>');
    if (r.underline) bits.push('<w:u w:val="single"/>');
    if (r.strike) bits.push('<w:strike/>');
    if (r.font) bits.push('<w:rFonts w:ascii="' + esc(r.font) + '" w:hAnsi="' + esc(r.font) + '"/>');
    // Word measures text in half-points.
    if (r.size) bits.push('<w:sz w:val="' + Math.round(r.size * 2) + '"/><w:szCs w:val="' + Math.round(r.size * 2) + '"/>');
    if (r.colour || r.color) bits.push('<w:color w:val="' + String(r.colour || r.color).replace('#', '').toUpperCase() + '"/>');
    const rPr = bits.length ? '<w:rPr>' + bits.join('') + '</w:rPr>' : '';
    return '<w:r>' + rPr + '<w:t xml:space="preserve">' + esc(r.text ?? '') + '</w:t></w:r>';
  };

  const runs = Array.isArray(p.runs) && p.runs.length
    ? runsWithLinks(p.runs, runXml, links)
    : runsWithLinks([{ ...p, text: p.text ?? '' }], runXml, links);

  return '<w:p>' + pPr + runs + '</w:p>';
}

/**
 * Consecutive runs that share a link go inside one `<w:hyperlink>`, which is
 * how Word groups them and how this package's reader expects to find them.
 * Link runs also take the blue underline, because a link nobody can see is not
 * a link — Word's own Hyperlink character style does the same thing.
 */
function runsWithLinks(list, runXml, links) {
  const out = [];
  let open = null;
  for (const run of list) {
    const url = links && run.link ? String(run.link) : null;
    if (url !== open) {
      if (open !== null) out.push('</w:hyperlink>');
      if (url !== null) {
        if (!links.has(url)) links.set(url, 'rIdLink' + (links.size + 1));
        out.push('<w:hyperlink r:id="' + links.get(url) + '">');
      }
      open = url;
    }
    out.push(runXml(url ? { underline: true, colour: '0563C1', ...run } : run));
  }
  if (open !== null) out.push('</w:hyperlink>');
  return out.join('');
}

/**
 * A .docx from nothing: the template gallery, the document a .odt or .rtf
 * becomes on import, and the fixtures the tests open.
 *
 * `paragraphs` takes plain strings or the object above. `styles: true` adds the
 * standard style catalogue — what the template gallery passes, so a person in a
 * fresh document can pick Heading 1. Off by default: the minimal shape is
 * byte-pinned by the goldens and plenty of tests want a document with no styles
 * part at all.
 */
export function buildDocx({ paragraphs = [], styles = false }) {
  // Links are collected while the body is written, because each one needs a
  // relationship id and the ids have to be minted in the order they appear.
  // The map stays empty for a document with no links, and everything below is
  // then byte-for-byte what it was before hyperlinks existed — which the
  // goldens depend on.
  const links = new Map();
  const body = paragraphs.map((p) => paragraphXml(p, links)).join('');
  const linked = links.size > 0;

  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
    (linked ? ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' : '') +
    '>' +
    '<w:body>' + body + '<w:sectPr/></w:body></w:document>';
  const contentTypes =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    (styles ? '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' : '') +
    '</Types>';
  const rootRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>';
  const documentRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    (styles ? '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' : '') +
    [...links].map(([url, id]) =>
      '<Relationship Id="' + id + '" Type="' + HYPERLINK_REL + '" Target="' + esc(url) + '" TargetMode="External"/>'
    ).join('') +
    '</Relationships>';

  // A document with links needs its relationship part whether or not it has a
  // style table: an r:id with nothing behind it is a file Word refuses.
  const needsRels = styles || linked;
  return packParts([
    { name: '[Content_Types].xml', data: contentTypes },
    { name: '_rels/.rels', data: rootRels },
    { name: 'word/document.xml', data: documentXml },
    ...(needsRels ? [{ name: 'word/_rels/document.xml.rels', data: documentRels }] : []),
    ...(styles ? [{ name: 'word/styles.xml', data: STANDARD_DOC_STYLES }] : []),
  ]);
}

/** Read back the <definedNames> table. This is the assertion that matters. */
export function readDefinedNames(xlsxBuf) {
  const parts = partsOf(xlsxBuf);
  const wb = parts.get('xl/workbook.xml');
  if (!wb) throw new Error('xl/workbook.xml missing - not an xlsx');
  const xml = wb.toString('utf8');
  const out = [];
  const re = /<definedName\b[^>]*name="([^"]+)"[^>]*>([^<]*)<\/definedName>/g;
  let m;
  while ((m = re.exec(xml))) out.push({ name: m[1], ref: m[2] });
  return out;
}

/** Read cell values of sheet N as a row-major array of strings. */
export function readSheetValues(xlsxBuf, sheetIndex = 0) {
  const parts = partsOf(xlsxBuf);
  const sheet = parts.get('xl/worksheets/sheet' + (sheetIndex + 1) + '.xml');
  if (!sheet) throw new Error('sheet ' + (sheetIndex + 1) + ' missing');
  const xml = sheet.toString('utf8');
  const rows = [];
  const rowRe = /<row\b[^>]*r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g;
  let rm;
  while ((rm = rowRe.exec(xml))) {
    const rowIdx = Number(rm[1]) - 1;
    const cells = [];
    const cellRe = /<c\b[^>]*r="([A-Z]+)(\d+)"[^>]*?(?:\/>|>([\s\S]*?)<\/c>)/g;
    let cm;
    while ((cm = cellRe.exec(rm[2]))) {
      let col = 0;
      for (const ch of cm[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
      col -= 1;
      const inner = cm[3] ?? '';
      const t = /<t[^>]*>([\s\S]*?)<\/t>/.exec(inner);
      const v = /<v>([\s\S]*?)<\/v>/.exec(inner);
      const f = /<f>([\s\S]*?)<\/f>/.exec(inner);
      let value = null;
      if (t) value = t[1];
      else if (v) value = v[1];
      else if (f) value = '=' + f[1];
      cells[col] = value === null ? null : unesc(value);
    }
    rows[rowIdx] = cells;
  }
  return rows;
}

/** Sheet names in workbook order. */
export function readSheetNames(xlsxBuf) {
  const parts = partsOf(xlsxBuf);
  const wb = parts.get('xl/workbook.xml');
  if (!wb) throw new Error('xl/workbook.xml missing - not an xlsx');
  const xml = wb.toString('utf8');
  const names = [];
  const re = /<sheet\b[^>]*name="([^"]+)"/g;
  let m;
  while ((m = re.exec(xml))) names.push(unesc(m[1]));
  return names;
}

/** Whole workbook as a plain object - the shape the tests and the fake builder use. */
export function readWorkbook(xlsxBuf) {
  const names = readSheetNames(xlsxBuf);
  return {
    sheets: names.map((name, i) => ({ name, rows: readSheetValues(xlsxBuf, i) })),
    definedNames: readDefinedNames(xlsxBuf),
  };
}

// `unesc` comes from @rutba/ooxml now. The copy that used to live here handled
// four entities; that one also handles &apos; and numeric character
// references, so this is a bug fix wearing a deduplication's clothes — a cell
// containing &#233; used to read back with the entity still in it.

/**
 * A letter with content controls — the document counterpart of a bound sheet.
 *
 * Content controls (`w:sdt` with a `w:tag`) are the anchors business data
 * lands in, and `w:sectPr` carries the letterhead geometry. Both are the things
 * a careless editor destroys, which is why the document view refuses to type
 * into paragraphs that hold them.
 *
 * @param {object} spec
 * @param {Array<string|{text?: string, style?: string, control?: {tag: string, alias: string, placeholder?: string}, field?: {instr: string, value: string}}>} spec.blocks
 */
export function buildLetterDocx({ blocks = [], title = '', letterhead = null, pageNumbers = false }) {
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  let controlId = 100000;

  /**
   * A table block. A confirmation letter states its figures in a table, not in a
   * sentence — the bank's own form has one and a letter without it reads as a
   * draft. Kept deliberately plain: a header row, rows of cells, right-aligned
   * where the caller says so.
   */
  const renderTable = (block) => {
    const widths = block.columns ?? [];
    const grid = '<w:tblGrid>' + widths.map((w) => '<w:gridCol w:w="' + Math.round(w) + '"/>').join('') + '</w:tblGrid>';
    const cell = (text, i, { header = false, align = null } = {}) => {
      const width = widths[i] ? '<w:tcW w:w="' + Math.round(widths[i]) + '" w:type="dxa"/>' : '';
      const shading = header ? '<w:shd w:val="clear" w:fill="D9E2F3"/>' : '';
      const jc = align ? '<w:pPr><w:jc w:val="' + esc(align) + '"/></w:pPr>' : '';
      const rPr = header ? '<w:rPr><w:b/></w:rPr>' : '';
      return '<w:tc><w:tcPr>' + width + shading + '</w:tcPr>'
        + '<w:p>' + jc + '<w:r>' + rPr + '<w:t xml:space="preserve">' + esc(text) + '</w:t></w:r></w:p></w:tc>';
    };
    const row = (cells, opts) => '<w:tr>'
      + (opts.header ? '<w:trPr><w:tblHeader/></w:trPr>' : '')
      + cells.map((t, i) => cell(t, i, { header: opts.header, align: block.align?.[i] ?? null })).join('')
      + '</w:tr>';

    return '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="5000" w:type="pct"/>'
      + '<w:tblBorders><w:top w:val="single" w:sz="8" w:color="7F7F7F"/>'
      + '<w:bottom w:val="single" w:sz="8" w:color="7F7F7F"/>'
      + '<w:insideH w:val="single" w:sz="4" w:color="D9D9D9"/>'
      + '<w:insideV w:val="single" w:sz="4" w:color="D9D9D9"/></w:tblBorders></w:tblPr>'
      + grid
      + (block.header ? row(block.header, { header: true }) : '')
      + (block.rows ?? []).map((r) => row(r, { header: false })).join('')
      + '</w:tbl>';
  };

  const renderBlock = (block) => {
    if (typeof block === 'string') {
      return '<w:p><w:r><w:t xml:space="preserve">' + esc(block) + '</w:t></w:r></w:p>';
    }
    if (block.table) return renderTable(block.table);
    const pPr = block.style ? '<w:pPr><w:pStyle w:val="' + esc(block.style) + '"/></w:pPr>' : '';
    let inner = '';
    if (block.text) inner += '<w:r><w:t xml:space="preserve">' + esc(block.text) + '</w:t></w:r>';
    if (block.field) {
      inner += '<w:fldSimple w:instr="' + esc(block.field.instr) + '">'
        + '<w:r><w:t>' + esc(block.field.value) + '</w:t></w:r></w:fldSimple>';
    }
    if (block.control) {
      controlId += 1;
      inner += '<w:sdt><w:sdtPr>'
        + '<w:alias w:val="' + esc(block.control.alias) + '"/>'
        + '<w:tag w:val="' + esc(block.control.tag) + '"/>'
        + '<w:id w:val="' + controlId + '"/><w:text/></w:sdtPr>'
        + '<w:sdtContent><w:r><w:rPr><w:b/></w:rPr><w:t>'
        + esc(block.control.placeholder ?? '—') + '</w:t></w:r></w:sdtContent></w:sdt>';
    }
    if (block.after) inner += '<w:r><w:t xml:space="preserve">' + esc(block.after) + '</w:t></w:r>';
    return '<w:p>' + pPr + inner + '</w:p>';
  };

  // A letter that leaves the organisation has a letterhead and a page number;
  // without them a two-page confirmation arrives looking like a loose draft.
  const R_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

  const documentXml = XML_DECL +
    '<w:document ' + W + ' ' + R_NS + '><w:body>' +
    blocks.map(renderBlock).join('') +
    '<w:sectPr>' +
    (letterhead ? '<w:headerReference w:type="default" r:id="rIdHdr"/>' : '') +
    (pageNumbers ? '<w:footerReference w:type="default" r:id="rIdFtr"/>' : '') +
    '<w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708"/>' +
    '</w:sectPr></w:body></w:document>';

  const headerXml = XML_DECL + '<w:hdr ' + W + '><w:p><w:pPr><w:jc w:val="right"/></w:pPr>' +
    '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">' + esc(letterhead ?? '') + '</w:t></w:r></w:p></w:hdr>';

  // "Page 1 of 3" — two fields, because NUMPAGES is the half that tells a reader
  // whether they are holding all of it.
  const footerXml = XML_DECL + '<w:ftr ' + W + '><w:p><w:pPr><w:jc w:val="center"/></w:pPr>' +
    '<w:r><w:t xml:space="preserve">Page </w:t></w:r>' +
    '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>' +
    '<w:r><w:t xml:space="preserve"> of </w:t></w:r>' +
    '<w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>1</w:t></w:r></w:fldSimple>' +
    '</w:p></w:ftr>';

  const styles = XML_DECL + '<w:styles ' + W + '>' +
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>' +
    '<w:rPr><w:b/><w:sz w:val="32"/><w:color w:val="1F5F8B"/></w:rPr></w:style>' +
    '</w:styles>';

  const contentTypes = XML_DECL +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    (letterhead ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' : '') +
    (pageNumbers ? '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' : '') +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    '</Types>';

  const rootRels = XML_DECL +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
    '</Relationships>';

  const documentRels = XML_DECL +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    (letterhead ? '<Relationship Id="rIdHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' : '') +
    (pageNumbers ? '<Relationship Id="rIdFtr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>' : '') +
    '</Relationships>';

  const core = XML_DECL +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
    'xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>' + esc(title) +
    '</dc:title><dc:creator>Rutba</dc:creator></cp:coreProperties>';

  return packParts([
    { name: '[Content_Types].xml', data: contentTypes },
    { name: '_rels/.rels', data: rootRels },
    { name: 'word/document.xml', data: documentXml },
    { name: 'word/_rels/document.xml.rels', data: documentRels },
    { name: 'word/styles.xml', data: styles },
    ...(letterhead ? [{ name: 'word/header1.xml', data: headerXml }] : []),
    ...(pageNumbers ? [{ name: 'word/footer1.xml', data: footerXml }] : []),
    { name: 'docProps/core.xml', data: core },
  ]);
}
