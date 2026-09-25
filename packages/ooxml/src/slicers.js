/**
 * Slicers — read and written as Excel 2010 and later write them.
 *
 * A slicer is four things in a package, and a reader has to find all four:
 *
 *   - a SLICER CACHE (`xl/slicerCaches/slicerCacheN.xml`), one per field
 *     being sliced, listed from the workbook's extension list — `x14:slicerCaches`
 *     for a pivot's, `x15:slicerCaches` for a table's. A pivot's cache names
 *     the pivot tables it drives and keeps each item's state (`<i x s nd>`);
 *     a table's names the table and the column (`x15:tableSlicerCache`) and
 *     keeps no state of its own — the table's autoFilter is the state.
 *   - a SLICERS part (`xl/slicers/slicerN.xml`) per sheet, listed from the
 *     sheet's extension list (`x14:slicerList`), with one `<slicer>` per panel:
 *     its name, its caption, its cache, how many columns of buttons.
 *   - a GRAPHIC FRAME in the sheet's drawing part (`sle:slicer`), inside
 *     `mc:AlternateContent` with a plain rectangle as the fallback older
 *     readers draw — this is WHERE the panel sits.
 *   - a defined name for the cache, `Slicer_Region`, as Excel adds one.
 *
 * Written from ECMA-376 and [MS-XLSX], not from another implementation.
 */
import { OoxmlPackage, esc } from './package.js';
import { unesc } from './workbook.js';

const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const X15 = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main';
const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL_SLICER = 'http://schemas.microsoft.com/office/2007/relationships/slicer';
const REL_SLICER_CACHE = 'http://schemas.microsoft.com/office/2007/relationships/slicerCache';
const CT_SLICER = 'application/vnd.ms-excel.slicer+xml';
const CT_SLICER_CACHE = 'application/vnd.ms-excel.slicerCache+xml';

/** The extension URIs Excel files these lists under. */
export const EXT = {
  workbookPivotCaches: '{BBE1A952-AA13-448e-AADC-164F8A28A991}',
  workbookTableCaches: '{46BE6895-7355-4a93-B00E-2C351335B9C9}',
  sheetPivotSlicers: '{A8765BA9-456A-4dab-B4F3-ACF838C121DE}',
  sheetTableSlicers: '{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}',
  tableSlicerCache: '{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}',
  pivotCacheId: '{725AE2AE-9491-48be-B2B4-4EB974FC3084}',
};

/** A slicer button's height when the file does not say — Excel's 0.26 in. */
export const DEFAULT_ROW_HEIGHT = 241300;

const attrsOf = (tag) => {
  const out = {};
  const re = /([\w:.-]+)\s*=\s*"([^"]*)"/g;
  let m;
  while ((m = re.exec(tag ?? ''))) out[m[1]] = m[2];
  return out;
};

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

// ---- reading ---------------------------------------------------------------

/** The slicer-cache parts the workbook lists, with which list each came from. */
function cacheParts(wb) {
  const xml = wb.pkg.text(wb.mainPart);
  const rels = new Map(wb.pkg.rels(wb.mainPart).map((r) => [r.Id, r]));
  const out = [];
  for (const ext of xml.matchAll(/<(?:\w+:)?ext\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?ext>/g)) {
    const uri = attrsOf(ext[1]).uri;
    const kind = uri === EXT.workbookPivotCaches ? 'pivot' : uri === EXT.workbookTableCaches ? 'table' : null;
    if (!kind) continue;
    for (const c of ext[2].matchAll(/<x14:slicerCache\b([^>]*?)\/?>/g)) {
      const rel = rels.get(attrsOf(c[1])['r:id']);
      if (!rel) continue;
      const part = OoxmlPackage.resolveTarget(wb.mainPart, rel.Target);
      if (wb.pkg.has(part)) out.push({ part, relId: rel.Id, list: kind });
    }
  }
  return out;
}

/** One slicer cache, read: its name, the field it slices, and what it drives. */
export function readSlicerCache(xml, part) {
  const head = attrsOf((/<slicerCacheDefinition\b([^>]*?)>/.exec(xml) ?? [])[1] ?? '');
  const pivots = [...xml.matchAll(/<pivotTable\b([^>]*?)\/?>/g)].map((m) => {
    const a = attrsOf(m[1]);
    return { tabId: Number(a.tabId), name: unesc(a.name ?? '') };
  });
  const tabular = attrsOf((/<tabular\b([^>]*?)>/.exec(xml) ?? [])[1] ?? '');
  const items = [...xml.matchAll(/<i\b([^>]*?)\/?>/g)].map((m) => {
    const a = attrsOf(m[1]);
    return { x: Number(a.x), selected: a.s === '1' || a.s === 'true', noData: a.nd === '1' || a.nd === 'true' };
  });
  const table = attrsOf((/<x15:tableSlicerCache\b([^>]*?)\/?>/.exec(xml) ?? [])[1] ?? '');
  return {
    part,
    name: unesc(head.name ?? ''),
    sourceName: unesc(head.sourceName ?? ''),
    kind: table.tableId ? 'table' : 'pivot',
    pivots,
    pivotCacheId: tabular.pivotCacheId ? Number(tabular.pivotCacheId) : null,
    items,
    table: table.tableId ? { id: Number(table.tableId), column: Number(table.column) } : null,
    sortOrder: tabular.sortOrder ?? table.sortOrder ?? 'ascending',
    crossFilter: tabular.crossFilter ?? table.crossFilter ?? 'showItemsWithDataAtTop',
  };
}

/** The slicers parts a sheet lists, from its relationships. */
function sheetSlicerParts(wb, sheetPart) {
  return wb.pkg.rels(sheetPart)
    .filter((r) => String(r.Type) === REL_SLICER || String(r.Type).endsWith('/slicer'))
    .map((r) => ({ relId: r.Id, part: OoxmlPackage.resolveTarget(sheetPart, r.Target) }))
    .filter((p) => wb.pkg.has(p.part));
}

/**
 * Every slicer in the workbook: each panel with its sheet, its cache and
 * what the cache drives. Position is the drawing's business, found by name.
 */
export function readSlicers(wb) {
  const caches = new Map();
  for (const c of cacheParts(wb)) {
    const cache = readSlicerCache(wb.pkg.text(c.part), c.part);
    caches.set(cache.name, { ...cache, relId: c.relId, list: c.list });
  }
  const out = [];
  for (const { name: sheet, part: sheetPart } of wb.sheets()) {
    for (const sp of sheetSlicerParts(wb, sheetPart)) {
      const xml = wb.pkg.text(sp.part);
      for (const m of xml.matchAll(/<slicer\b([^>]*?)\/?>/g)) {
        const a = attrsOf(m[1]);
        const cache = caches.get(unesc(a.cache ?? '')) ?? null;
        out.push({
          sheet,
          part: sp.part,
          name: unesc(a.name ?? ''),
          caption: a.caption !== undefined ? unesc(a.caption) : unesc(a.name ?? ''),
          cacheName: unesc(a.cache ?? ''),
          columns: Math.max(1, Number(a.columnCount ?? 1) || 1),
          rowHeight: Number(a.rowHeight ?? DEFAULT_ROW_HEIGHT) || DEFAULT_ROW_HEIGHT,
          style: a.style ?? 'SlicerStyleLight1',
          showCaption: a.showCaption !== '0',
          cache,
        });
      }
    }
  }
  return out;
}

// ---- writing ---------------------------------------------------------------

/** A name nothing in the list holds yet: `Region`, `Region 1`, `Region 2`… */
function freshName(base, taken, sep = ' ') {
  const used = new Set([...taken].map((s) => String(s).toLowerCase()));
  if (!used.has(base.toLowerCase())) return base;
  for (let n = 1; ; n++) if (!used.has((base + sep + n).toLowerCase())) return base + sep + n;
}

/** A defined-name-safe form of a field's name: `Slicer_Sales_Rep`. */
function cacheNameFor(field) {
  const body = String(field).replace(/[^A-Za-z0-9_.]+/g, '_').replace(/^([^A-Za-z_])/, '_$1');
  return 'Slicer_' + (body || 'Field');
}

/** Put one `<ext>` child into an extension list, making the list or the ext when missing. */
function addToExt(extLst, uri, nsDecl, listOpen, entry) {
  const listTag = listOpen.split(/\s/)[0];
  const re = new RegExp('<(?:\\w+:)?ext\\b[^>]*uri="' + uri.replace(/[{}]/g, '\\$&') + '"[^>]*>([\\s\\S]*?)</(?:\\w+:)?ext>');
  const found = re.exec(extLst ?? '');
  if (found) {
    const inner = found[1].replace('</' + listTag + '>', entry + '</' + listTag + '>');
    return extLst.replace(found[0], () => found[0].replace(found[1], () => inner));
  }
  const ext = '<ext uri="' + uri + '" ' + nsDecl + '><' + listOpen + '>' + entry + '</' + listTag + '></ext>';
  if (extLst) return extLst.replace(/<\/extLst>\s*$/, ext + '</extLst>');
  return '<extLst>' + ext + '</extLst>';
}

/** Take one entry (by r:id) out of an extension list, dropping emptied lists. */
function removeFromExt(extLst, relId) {
  if (!extLst) return extLst;
  let out = extLst.replace(new RegExp('<x14:(slicerCache|slicer)\\b[^>]*r:id="' + relId + '"[^>]*/>', 'g'), '');
  out = out.replace(/<(?:\w+:)?ext\b[^>]*>\s*<(x14|x15):(slicerCaches|slicerList)\b[^>]*>\s*<\/\1:\2>\s*<\/(?:\w+:)?ext>/g, '');
  return /<(?:\w+:)?ext\b/.test(out) ? out : null;
}

/** The workbook's own extList, as text, and a writer for it — `extLst` is always last. */
function workbookExt(wb) {
  const xml = wb.pkg.text(wb.mainPart);
  const m = /<extLst>[\s\S]*<\/extLst>/.exec(xml);
  return {
    xml: m ? m[0] : null,
    write(next) {
      const cur = wb.pkg.text(wb.mainPart);
      const has = /<extLst>[\s\S]*<\/extLst>/.exec(cur);
      let out;
      if (has) out = next ? cur.replace(has[0], () => next) : cur.replace(has[0], '');
      else out = next ? cur.replace(/<\/workbook>\s*$/, next + '</workbook>') : cur;
      if (!/xmlns:r=/.test(out.slice(0, out.indexOf('>', out.indexOf('<workbook'))))) {
        out = out.replace('<workbook ', '<workbook xmlns:r="' + R_NS + '" ');
      }
      wb.pkg.write_(wb.mainPart, out);
    },
  };
}

/**
 * The x14 id a pivot cache is known to slicers by, written into the cache
 * definition's own extension list when it carries none yet.
 */
function ensurePivotCacheId(wb, pivot) {
  const xml = wb.pkg.text(pivot.cachePart);
  const had = /<x14:pivotCacheDefinition\b([^>]*?)\/?>/.exec(xml);
  if (had && attrsOf(had[1]).pivotCacheId) return Number(attrsOf(had[1]).pivotCacheId);
  // Unique in the workbook: every other cache's id, and the workbook's cacheIds.
  const used = new Set();
  for (const p of wb.pkg.partNames()) {
    if (!/^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/.test(p)) continue;
    const m = /<x14:pivotCacheDefinition\b[^>]*pivotCacheId="(\d+)"/.exec(wb.pkg.text(p));
    if (m) used.add(Number(m[1]));
  }
  let id = (pivot.cacheId ?? 1) + 1000;
  while (used.has(id)) id += 1;
  const ext = '<ext uri="' + EXT.pivotCacheId + '" xmlns:x14="' + X14 + '"><x14:pivotCacheDefinition pivotCacheId="' + id + '"/></ext>';
  let next;
  if (had) next = xml.replace(had[0], '<x14:pivotCacheDefinition' + had[1].replace(/\/$/, '') + ' pivotCacheId="' + id + '"/>');
  else if (/<extLst>/.test(xml)) next = xml.replace('</extLst>', ext + '</extLst>');
  else next = xml.replace(/<\/pivotCacheDefinition>\s*$/, '<extLst>' + ext + '</extLst></pivotCacheDefinition>');
  wb.pkg.write_(pivot.cachePart, next);
  return id;
}

/** A pivot slicer cache's items: each shared item, selected or not, with data or not. */
export function slicerItemsXml(items) {
  return '<items count="' + items.length + '">'
    + items.map((it) => '<i x="' + it.x + '"' + (it.selected ? ' s="1"' : '') + (it.noData ? ' nd="1"' : '') + '/>').join('')
    + '</items>';
}

/** Rewrite a pivot slicer cache's item states. */
export function writeSlicerCacheItems(wb, cachePart, items) {
  const xml = wb.pkg.text(cachePart);
  const block = /<items\b[^>]*>[\s\S]*?<\/items>|<items\b[^>]*\/>/;
  const next = block.test(xml)
    ? xml.replace(block, () => slicerItemsXml(items))
    : xml.replace(/(<tabular\b[^>]*>)/, (m) => m + slicerItemsXml(items));
  if (next !== xml) wb.pkg.write_(cachePart, next);
}

/**
 * The drawing anchor for a slicer panel: a graphic frame naming the slicer,
 * inside the markup-compatibility wrapper Excel writes, with the rectangle
 * older readers draw in its place.
 */
export function slicerAnchorXml({ id, name, kind, from, to, offset = { x: 0, y: 0 }, size = { cx: 1828800, cy: 2524125 } }) {
  const marker = (tag, p) => '<xdr:' + tag + '><xdr:col>' + p.col + '</xdr:col><xdr:colOff>' + (p.colOff ?? 0) + '</xdr:colOff>'
    + '<xdr:row>' + p.row + '</xdr:row><xdr:rowOff>' + (p.rowOff ?? 0) + '</xdr:rowOff></xdr:' + tag + '>';
  const choice = kind === 'table'
    ? '<mc:Choice xmlns:sle15="http://schemas.microsoft.com/office/drawing/2012/slicer" Requires="sle15">'
    : '<mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">';
  const what = kind === 'table' ? 'a table slicer' : 'a slicer';
  return '<xdr:twoCellAnchor editAs="oneCell">' + marker('from', from) + marker('to', to)
    + '<mc:AlternateContent xmlns:mc="' + MC + '">' + choice
    + '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="' + id + '" name="' + esc(name) + '"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'
    + '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
    + '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2010/slicer">'
    + '<sle:slicer xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer" name="' + esc(name) + '"/>'
    + '</a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice>'
    + '<mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="0" name=""/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr>'
    + '<xdr:spPr><a:xfrm><a:off x="' + offset.x + '" y="' + offset.y + '"/><a:ext cx="' + size.cx + '" cy="' + size.cy + '"/></a:xfrm>'
    + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val="white"/></a:solidFill><a:ln w="1"><a:solidFill><a:prstClr val="green"/></a:solidFill></a:ln></xdr:spPr>'
    + '<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/><a:p><a:r><a:rPr lang="en-GB" sz="1100"/>'
    + '<a:t>This shape represents ' + what + '. Slicers are supported in Excel 2010 and later, and in Rutba Worksheets.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback>'
    + '</mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>';
}

/**
 * Add one slicer: its cache (with the workbook's relationship, extension
 * entry and defined name), its `<slicer>` in the sheet's slicers part (made,
 * related and listed in the sheet's extension list when the sheet has none
 * of that kind yet). The drawing anchor is the caller's to append, since it
 * knows where the panel goes; this answers the name to give it.
 *
 * @param {object} wb
 * @param {object} spec
 * @param {string} spec.sheet     the sheet the panel sits on
 * @param {'pivot'|'table'} spec.kind
 * @param {string} spec.field     the field or column being sliced
 * @param {object} [spec.pivot]   the pivot descriptor (readPivots)
 * @param {number} [spec.pivotTabId] the sheetId of the pivot's sheet
 * @param {Array<{x:number, selected:boolean, noData:boolean}>} [spec.items]
 * @param {{id:number, column:number}} [spec.table] the table's id, and the column's id (1-based)
 * @param {number} [spec.columns]
 * @returns {{ name: string, cacheName: string, cachePart: string, slicerPart: string }}
 */
export function addSlicer(wb, spec) {
  const existing = readSlicers(wb);
  const taken = [...existing.map((s) => s.name)];
  const cachesTaken = [...existing.map((s) => s.cacheName), ...wb.definedNames().map((d) => d.name)];
  const name = freshName(String(spec.field), taken);
  const cacheName = freshName(cacheNameFor(spec.field), cachesTaken, '');

  // ---- the cache
  const n = wb.pkg.nextPartNumber('xl/slicerCaches/', 'slicerCache');
  const cachePart = 'xl/slicerCaches/slicerCache' + n + '.xml';
  let body;
  if (spec.kind === 'pivot') {
    const pivotCacheId = ensurePivotCacheId(wb, spec.pivot);
    body = '<pivotTables><pivotTable tabId="' + spec.pivotTabId + '" name="' + esc(spec.pivot.name) + '"/></pivotTables>'
      + '<data><tabular pivotCacheId="' + pivotCacheId + '">' + slicerItemsXml(spec.items ?? []) + '</tabular></data>';
  } else {
    body = '<extLst><x:ext uri="' + EXT.tableSlicerCache + '" xmlns:x15="' + X15 + '">'
      + '<x15:tableSlicerCache tableId="' + spec.table.id + '" column="' + spec.table.column + '"/></x:ext></extLst>';
  }
  wb.pkg.addPart(cachePart, DECL
    + '<slicerCacheDefinition xmlns="' + X14 + '" xmlns:mc="' + MC + '" mc:Ignorable="x" xmlns:x="' + MAIN + '"'
    + ' name="' + esc(cacheName) + '" sourceName="' + esc(spec.field) + '">' + body + '</slicerCacheDefinition>',
    CT_SLICER_CACHE);
  const cacheRel = wb.pkg.addRelationshipTo(wb.mainPart, REL_SLICER_CACHE, 'slicerCaches/slicerCache' + n + '.xml');
  const ext = workbookExt(wb);
  const entry = '<x14:slicerCache r:id="' + cacheRel + '"/>';
  ext.write(spec.kind === 'pivot'
    ? addToExt(ext.xml, EXT.workbookPivotCaches, 'xmlns:x14="' + X14 + '"', 'x14:slicerCaches', entry)
    : addToExt(ext.xml, EXT.workbookTableCaches, 'xmlns:x15="' + X15 + '"', 'x15:slicerCaches xmlns:x14="' + X14 + '"', entry));
  wb.setDefinedName(cacheName, '#N/A');

  // ---- the panel, in the sheet's slicers part of this kind
  const sheetPart = wb.partNameFor(spec.sheet);
  const list = spec.kind === 'pivot' ? EXT.sheetPivotSlicers : EXT.sheetTableSlicers;
  const sheetExt = wb._sheetPart(spec.sheet).part.tailElement('extLst');
  const listed = new Set();
  const listRe = new RegExp('<(?:\\w+:)?ext\\b[^>]*uri="' + list.replace(/[{}]/g, '\\$&') + '"[^>]*>([\\s\\S]*?)</(?:\\w+:)?ext>');
  const listBlock = listRe.exec(sheetExt ?? '');
  if (listBlock) for (const m of listBlock[1].matchAll(/r:id="([^"]+)"/g)) listed.add(m[1]);
  const own = sheetSlicerParts(wb, sheetPart).find((p) => listed.has(p.relId));
  const slicerXml = '<slicer name="' + esc(name) + '" cache="' + esc(cacheName) + '" caption="' + esc(spec.caption ?? spec.field) + '"'
    + (spec.columns > 1 ? ' columnCount="' + spec.columns + '"' : '') + ' rowHeight="' + DEFAULT_ROW_HEIGHT + '"/>';
  let slicerPart;
  if (own) {
    slicerPart = own.part;
    wb.pkg.write_(slicerPart, wb.pkg.text(slicerPart).replace(/<\/slicers>\s*$/, slicerXml + '</slicers>'));
  } else {
    const sn = wb.pkg.nextPartNumber('xl/slicers/', 'slicer');
    slicerPart = 'xl/slicers/slicer' + sn + '.xml';
    wb.pkg.addPart(slicerPart, DECL
      + '<slicers xmlns="' + X14 + '" xmlns:mc="' + MC + '" mc:Ignorable="x" xmlns:x="' + MAIN + '">' + slicerXml + '</slicers>',
      CT_SLICER);
    const rel = wb.pkg.addRelationshipTo(sheetPart, REL_SLICER, '../slicers/slicer' + sn + '.xml');
    const item = '<x14:slicer r:id="' + rel + '"/>';
    // The r: prefix the list's entries use is declared on the ext itself, so
    // a sheet whose root never declared it still reads.
    const next = spec.kind === 'pivot'
      ? addToExt(sheetExt, list, 'xmlns:x14="' + X14 + '" xmlns:r="' + R_NS + '"', 'x14:slicerList', item)
      : addToExt(sheetExt, list, 'xmlns:x15="' + X15 + '" xmlns:r="' + R_NS + '"', 'x14:slicerList xmlns:x14="' + X14 + '"', item);
    wb._sheetPart(spec.sheet).part.setTailElement('extLst', next);
  }
  return { name, cacheName, cachePart, slicerPart };
}

/**
 * Take a slicer away: its `<slicer>`, its cache when no other slicer uses it
 * (with the relationship, the extension entry and the defined name), and a
 * slicers part left empty (with its relationship and list entry). The
 * drawing anchor is the caller's.
 */
export function removeSlicer(wb, sheet, name) {
  const all = readSlicers(wb);
  const s = all.find((x) => x.sheet === sheet && x.name === name);
  if (!s) throw new Error('no slicer "' + name + '" on ' + sheet);
  const sheetPart = wb.partNameFor(sheet);
  const xml = wb.pkg.text(s.part);
  const next = xml.replace(new RegExp('<slicer\\b[^>]*name="' + esc(name).replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '"[^>]*/>'), '');
  if (/<slicer\b/.test(next)) {
    wb.pkg.write_(s.part, next);
  } else {
    const rel = sheetSlicerParts(wb, sheetPart).find((p) => p.part === s.part);
    wb.pkg.removePart(s.part);
    if (rel) {
      removeRel(wb, sheetPart, rel.relId);
      const part = wb._sheetPart(sheet).part;
      part.setTailElement('extLst', removeFromExt(part.tailElement('extLst'), rel.relId));
    }
  }
  const stillUsed = all.some((x) => x !== s && x.cacheName === s.cacheName);
  if (!stillUsed && s.cache) {
    wb.pkg.removePart(s.cache.part);
    removeRel(wb, wb.mainPart, s.cache.relId);
    const ext = workbookExt(wb);
    ext.write(removeFromExt(ext.xml, s.cache.relId));
    wb.setDefinedName(s.cacheName, null);
  }
  return s;
}

/** Drop one relationship from a part's rels. */
function removeRel(wb, fromPart, relId) {
  const relsPath = OoxmlPackage.relsPathFor(fromPart);
  if (!wb.pkg.has(relsPath)) return;
  const xml = wb.pkg.text(relsPath);
  wb.pkg.write_(relsPath, xml.replace(new RegExp('<Relationship\\b[^>]*Id="' + relId + '"[^>]*/>'), ''));
}

/** Rename a slicer's caption, or its column count, in its slicers part. */
export function setSlicerProps(wb, sheet, name, { caption, columns } = {}) {
  const s = readSlicers(wb).find((x) => x.sheet === sheet && x.name === name);
  if (!s) throw new Error('no slicer "' + name + '" on ' + sheet);
  const xml = wb.pkg.text(s.part);
  const re = new RegExp('<slicer\\b[^>]*name="' + esc(name).replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '"[^>]*/>');
  const next = xml.replace(re, (tag) => {
    let t = tag;
    if (caption !== undefined) t = t.replace(/\s+caption="[^"]*"/, '').replace(/\s*\/>$/, ' caption="' + esc(caption) + '"/>');
    if (columns !== undefined) t = t.replace(/\s+columnCount="[^"]*"/, '').replace(/\s*\/>$/, (columns > 1 ? ' columnCount="' + columns + '"' : '') + '/>');
    return t;
  });
  if (next !== xml) wb.pkg.write_(s.part, next);
}

export { REL_SLICER, REL_SLICER_CACHE, CT_SLICER, CT_SLICER_CACHE };
