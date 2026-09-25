// Insert → Slicer: a floating panel with a button per item that filters a
// table or a pivot table.
//
// Written as Excel 2010 and later write slicers: `xl/slicers/slicerN.xml`
// (the panels), `xl/slicerCaches/slicerCacheN.xml` (one per field), the
// workbook's extension-list entries (`x14:slicerCaches` for a pivot's,
// `x15:slicerCaches` with `x15:tableSlicerCache` for a table's), the sheet's
// `x14:slicerList`, the drawing's `sle:slicer` graphic frame inside
// `mc:AlternateContent`, the defined name, content types and relationships.
// A table's filter is its autoFilter; a pivot's is hidden items in its field,
// with the slicer cache keeping each item's state. Excel's own slicers are
// read back.

import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { readSlicers, EXT } from '@rutba/ooxml/slicers';
import { SheetView } from '@rutba/sheet-view';

const ROWS = [
  ['Region', 'Rep', 'Units'],
  ['North', 'Amira', 10],
  ['South', 'Ben', 20],
  ['North', 'Chen', 5],
  ['East', 'Dana', 7],
  ['South', 'Amira', 3],
  ['East', 'Ben', 12],
];
const BOOK = () => buildXlsx({ sheets: [{ name: 'Sales', rows: ROWS }] });

/** A table over the data, and the view with the cursor in it. */
const tableView = () => {
  const view = new SheetView(BOOK());
  view.select(1, 0);
  view.formatAsTable({ style: 'TableStyleMedium2', stripes: true });
  view.select(2, 1);
  return view;
};
const slicer = (view, name) => view.render().drawings.find((d) => d.kind === 'slicer' && d.slicer?.name === name)?.slicer;

test('a table slicer is written as Excel writes one: slicers part, cache, workbook and sheet lists, drawing, name, types', () => {
  const view = tableView();
  assert.deepEqual(view.slicerSources().fields.map((f) => f.name), ['Region', 'Rep', 'Units']);
  const made = view.insertSlicers({ fields: ['Region'] });
  assert.deepEqual(made, ['Region']);
  const saved = SheetView.open(view.save());
  const pkg = saved.pkg;
  assert.ok(pkg.has('xl/slicers/slicer1.xml') && pkg.has('xl/slicerCaches/slicerCache1.xml'));
  assert.equal(pkg.contentTypeOf('/xl/slicers/slicer1.xml') ?? pkg.contentTypeOf('xl/slicers/slicer1.xml'), 'application/vnd.ms-excel.slicer+xml');
  assert.equal(pkg.contentTypeOf('/xl/slicerCaches/slicerCache1.xml') ?? pkg.contentTypeOf('xl/slicerCaches/slicerCache1.xml'), 'application/vnd.ms-excel.slicerCache+xml');
  const cache = pkg.text('xl/slicerCaches/slicerCache1.xml');
  assert.match(cache, /<slicerCacheDefinition xmlns="http:\/\/schemas.microsoft.com\/office\/spreadsheetml\/2009\/9\/main"[^>]*name="Slicer_Region" sourceName="Region">/);
  assert.match(cache, /<x15:tableSlicerCache tableId="1" column="1"\/>/);
  const panel = pkg.text('xl/slicers/slicer1.xml');
  assert.match(panel, /<slicer name="Region" cache="Slicer_Region" caption="Region" rowHeight="241300"\/>/);
  const wb = pkg.text('xl/workbook.xml');
  assert.ok(wb.includes('<ext uri="' + EXT.workbookTableCaches + '"'), 'the workbook lists the cache under x15:slicerCaches');
  assert.match(wb, /<x15:slicerCaches xmlns:x14="[^"]+"><x14:slicerCache r:id="rId\d+"\/><\/x15:slicerCaches>/);
  assert.match(wb, /<definedName name="Slicer_Region">#N\/A<\/definedName>/);
  assert.match(wb, /<extLst>[\s\S]*<\/extLst><\/workbook>\s*$/, 'the extension list stays last');
  const wbRel = pkg.rels('xl/workbook.xml').find((r) => r.Type.endsWith('/slicerCache'));
  assert.equal(wbRel.Target, 'slicerCaches/slicerCache1.xml');
  const sheetPart = saved.workbook.partNameFor('Sales');
  const sheetXml = pkg.text(sheetPart);
  assert.ok(sheetXml.includes('<ext uri="' + EXT.sheetTableSlicers + '"'), 'the sheet lists its table slicers');
  assert.match(sheetXml, /<x14:slicerList xmlns:x14="[^"]+"><x14:slicer r:id="rId\d+"\/><\/x14:slicerList>/);
  const drawingPart = pkg.rels(sheetPart).find((r) => r.Type.endsWith('/drawing'));
  const drawing = pkg.text('xl/' + drawingPart.Target.replace(/^\.\.\//, ''));
  assert.match(drawing, /<mc:AlternateContent xmlns:mc="[^"]+"><mc:Choice xmlns:sle15="http:\/\/schemas.microsoft.com\/office\/drawing\/2012\/slicer" Requires="sle15">/);
  assert.match(drawing, /<a:graphicData uri="http:\/\/schemas.microsoft.com\/office\/drawing\/2010\/slicer"><sle:slicer xmlns:sle="[^"]+" name="Region"\/>/);
  assert.match(drawing, /<mc:Fallback><xdr:sp\b/, 'a rectangle for readers without slicers');
  // Read back.
  const read = readSlicers(saved.workbook);
  assert.equal(read.length, 1);
  assert.deepEqual([read[0].name, read[0].cache.kind, read[0].cache.table.id], ['Region', 'table', 1]);
  const s = slicer(saved, 'Region');
  assert.deepEqual(s.items.map((i) => i.label), ['East', 'North', 'South']);
  assert.ok(s.items.every((i) => i.selected && i.hasData));
});

test('pressing a table slicer\'s buttons filters the table; another slicer greys items with no data; Clear Filter; undo', () => {
  const view = tableView();
  view.insertSlicers({ fields: ['Region', 'Rep'] });
  view.setSlicerSelection({ name: 'Region', values: ['North'] });
  const region = slicer(view, 'Region');
  assert.deepEqual(region.items.filter((i) => i.selected).map((i) => i.label), ['North']);
  assert.ok(region.filtered);
  // Rows 3, 5, 6, 7 (South, East, South, East) hide.
  const t = view.workbook.tables()[0];
  assert.deepEqual([...view.workbook.tableFilters(t.part).get(0)], ['North']);
  const shown = view.render().rows.map((r) => r.index).filter((r) => r <= 6);
  assert.deepEqual(shown, [0, 1, 3], 'only North rows show');
  const rep = slicer(view, 'Rep');
  assert.deepEqual(rep.items.map((i) => [i.label, i.hasData]), [['Amira', true], ['Chen', true], ['Ben', false], ['Dana', false]],
    'reps with no North rows sink to the end, greyed');
  view.setSlicerSelection({ name: 'Region', values: null });
  assert.ok(slicer(view, 'Region').items.every((i) => i.selected), 'Clear Filter selects every item');
  assert.equal(view.workbook.tableFilters(t.part).get(0), undefined);
  view.undo();
  assert.deepEqual(slicer(view, 'Region').items.filter((i) => i.selected).map((i) => i.label), ['North'], 'undo puts the filter back');
  view.redo();
  assert.ok(slicer(view, 'Region').items.every((i) => i.selected));
  assert.throws(() => view.setSlicerSelection({ name: 'Region', values: [] }), /at least one item/);
});

test('a pivot slicer: cache with the pivot and item states, x14 lists, hidden items; a field off the axes filters too', () => {
  const view = new SheetView(BOOK());
  view.createPivot({ source: 'Sales!A1:C7', rowFields: ['Region'], dataFields: [{ field: 'Units' }] });
  view.select(9, 0);
  assert.equal(view.slicerSources().kind, 'pivot');
  view.insertSlicers({ fields: ['Region', 'Rep'] });
  const saved = () => SheetView.open(view.save());
  let pkg = saved().pkg;
  const cache = pkg.text('xl/slicerCaches/slicerCache1.xml');
  assert.match(cache, /<pivotTables><pivotTable tabId="1" name="PivotTable1"\/><\/pivotTables>/);
  const id = /<tabular pivotCacheId="(\d+)">/.exec(cache)?.[1];
  assert.ok(id, 'a tabular cache');
  assert.match(pkg.text('xl/pivotCache/pivotCacheDefinition1.xml'), new RegExp('<x14:pivotCacheDefinition pivotCacheId="' + id + '"/>'), 'the pivot cache is known to slicers by that id');
  assert.match(cache, /<items count="3"><i x="0" s="1"\/><i x="1" s="1"\/><i x="2" s="1"\/><\/items>/);
  const wb = pkg.text('xl/workbook.xml');
  assert.ok(wb.includes('<ext uri="' + EXT.workbookPivotCaches + '"'));
  assert.match(wb, /<x14:slicerCaches><x14:slicerCache r:id="rId\d+"\/><x14:slicerCache r:id="rId\d+"\/><\/x14:slicerCaches>/);
  const sheetXml = pkg.text('xl/worksheets/sheet1.xml');
  assert.ok(sheetXml.includes('<ext uri="' + EXT.sheetPivotSlicers + '"'));
  assert.match(pkg.text('xl/slicers/slicer1.xml'), /<slicer name="Region"[^>]*\/><slicer name="Rep"[^>]*\/>/, 'one panel part for both');
  assert.match(pkg.text(pkg.partNames().find((p) => /drawings\/drawing\d+\.xml$/.test(p))), /Requires="a14"/);

  // Rep is on no axis: the grid keeps its rows, the totals drop Amira's units.
  view.setSlicerSelection({ name: 'Rep', values: ['Ben', 'Chen', 'Dana'] });
  const cell = (r, c) => view.calc.getValue('Sales', r, c);
  assert.deepEqual([cell(9, 0), cell(9, 1), cell(10, 0), cell(10, 1), cell(11, 0), cell(11, 1)], ['East', 19, 'North', 5, 'South', 20]);
  pkg = saved().pkg;
  assert.match(pkg.text('xl/pivotTables/pivotTable1.xml'), /<pivotField showAll="0"><items count="4"><item x="0" h="1"\/><item x="1"\/><item x="2"\/><item x="3"\/><\/items><\/pivotField>/,
    'Amira hidden in the Rep field');
  assert.match(pkg.text('xl/slicerCaches/slicerCache2.xml'), /<i x="0"\/><i x="1" s="1"\/>/, 'the cache keeps the states');
  // Region is the row field: North alone.
  view.setSlicerSelection({ name: 'Region', values: ['North'] });
  assert.deepEqual([cell(9, 0), cell(9, 1), cell(10, 0), cell(10, 1)], ['North', 5, 'Grand Total', 5]);
  const rep = slicer(view, 'Rep');
  assert.deepEqual(rep.items.map((i) => [i.label, i.selected, i.hasData]), [['Amira', false, true], ['Chen', true, true], ['Ben', true, false], ['Dana', true, false]]);
  view.undo();
  assert.equal(cell(11, 0), 'South', 'undo brings the other regions back');
  // Reopened, the states read back.
  const again = saved();
  assert.deepEqual(slicer(again, 'Rep').items.filter((i) => !i.selected).map((i) => i.label), ['Amira']);
});

test('Excel\'s own slicers read back: panel, caption, columns, items and states', () => {
  const view = new SheetView(BOOK());
  view.createPivot({ source: 'Sales!A1:C7', rowFields: ['Region'], dataFields: [{ field: 'Units' }] });
  const bytes = view.save();
  // The parts an Excel 365 file carries for one pivot slicer, as it writes them.
  const wb = SheetView.open(bytes);
  const pkg = wb.pkg;
  pkg.write_('xl/pivotTables/pivotTable1.xml', pkg.text('xl/pivotTables/pivotTable1.xml')
    .replace('<pivotField axis="axisRow" showAll="0"><items count="4"><item x="0"/><item x="1"/><item x="2"/>', '<pivotField axis="axisRow" showAll="0"><items count="4"><item x="0" h="1"/><item x="1"/><item x="2"/>'));
  pkg.addPart('xl/slicerCaches/slicerCache1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    + '<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x xr10" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" name="Slicer_Region" xr10:uid="{6F8C1A37-0D2B-4B55-9E6B-2C6C4B8D3A11}" sourceName="Region">'
    + '<pivotTables><pivotTable tabId="1" name="PivotTable1"/></pivotTables><data><tabular pivotCacheId="1813658795"><items count="3"><i x="0"/><i x="1" s="1"/><i x="2" s="1"/></items></tabular></data></slicerCacheDefinition>',
  'application/vnd.ms-excel.slicerCache+xml');
  pkg.addPart('xl/slicers/slicer1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    + '<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x xr10" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10">'
    + '<slicer name="Region 1" xr10:uid="{A1B2C3D4-0000-4000-8000-000000000001}" cache="Slicer_Region" caption="Pick a region" columnCount="2" rowHeight="241300"/></slicers>',
  'application/vnd.ms-excel.slicer+xml');
  const wbRel = pkg.addRelationshipTo('xl/workbook.xml', 'http://schemas.microsoft.com/office/2007/relationships/slicerCache', 'slicerCaches/slicerCache1.xml');
  pkg.write_('xl/workbook.xml', pkg.text('xl/workbook.xml').replace(/<\/workbook>\s*$/, '<extLst><ext uri="{BBE1A952-AA13-448e-AADC-164F8A28A991}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerCaches><x14:slicerCache r:id="' + wbRel + '"/></x14:slicerCaches></ext></extLst></workbook>'));
  const sheetRel = pkg.addRelationshipTo('xl/worksheets/sheet1.xml', 'http://schemas.microsoft.com/office/2007/relationships/slicer', '../slicers/slicer1.xml');
  const drawing = wb.workbook.ensureSheetDrawing('Sales');
  wb.workbook.appendDrawingAnchor(drawing, () => '<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>14</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
    + '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14"><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Region 1"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{00000000-0000-0000-0000-000000000000}"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2010/slicer"><sle:slicer xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer" name="Region 1"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice>'
    + '<mc:Fallback xmlns=""><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="0" name=""/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="3048000" y="190500"/><a:ext cx="1828800" cy="2476500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp></mc:Fallback></mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>');
  const part = wb.workbook._sheetPart('Sales').part;
  part.setTailElement('extLst', '<extLst><ext uri="{A8765BA9-456A-4dab-B4F3-ACF838C121DE}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerList><x14:slicer r:id="' + sheetRel + '"/></x14:slicerList></ext></extLst>');
  const excel = SheetView.open(wb.save());

  const read = readSlicers(excel.workbook);
  assert.equal(read.length, 1);
  assert.deepEqual([read[0].name, read[0].caption, read[0].columns, read[0].cache.pivotCacheId], ['Region 1', 'Pick a region', 2, 1813658795]);
  const d = excel.render().drawings.find((x) => x.kind === 'slicer');
  assert.ok(d, 'the panel is placed by its drawing');
  assert.equal(d.x, excel.geo.colOffset(5));
  const s = d.slicer;
  assert.deepEqual([s.caption, s.columns, s.kind, s.source], ['Pick a region', 2, 'pivot', 'PivotTable1']);
  assert.deepEqual(s.items.map((i) => [i.label, i.selected]), [['East', true], ['North', false], ['South', true]], 'North hidden in the pivot reads as not selected');
  // And it drives the pivot here too.
  excel.setSlicerSelection({ name: 'Region 1', values: ['North'] });
  assert.equal(excel.calc.getValue('Sales', 9, 0), 'North');
});

test('a slicer is moved and resized by its anchor, and deleted with its parts; undo brings it back', () => {
  const view = tableView();
  view.insertSlicers({ fields: ['Region'] });
  const d = view.render().drawings.find((x) => x.kind === 'slicer');
  view.setDrawingBox({ id: d.id, x: d.x + 30, y: d.y + 40, width: 220, height: 200 });
  const moved = view.render().drawings.find((x) => x.kind === 'slicer');
  assert.deepEqual([moved.x, moved.y, moved.width, moved.height], [d.x + 30, d.y + 40, 220, 200]);
  view.deleteDrawings({ ids: [moved.id] });
  assert.equal(view.render().drawings.filter((x) => x.kind === 'slicer').length, 0);
  const pkg = SheetView.open(view.save()).pkg;
  assert.ok(!pkg.has('xl/slicers/slicer1.xml') && !pkg.has('xl/slicerCaches/slicerCache1.xml'), 'its parts go');
  assert.doesNotMatch(pkg.text('xl/workbook.xml'), /Slicer_Region|slicerCache/, 'the name and the list entry go');
  view.undo();
  assert.equal(view.render().drawings.filter((x) => x.kind === 'slicer').length, 1, 'undo puts it back');
  assert.ok(view.pkg.has('xl/slicers/slicer1.xml'));
});
