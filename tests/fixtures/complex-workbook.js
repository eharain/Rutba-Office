/**
 * A workbook that carries the things we do NOT model.
 *
 * Charts, a pivot cache, a drawing, binary media, custom XML, styles, shared
 * strings, merged cells, conditional formatting, data validation and sheet
 * protection. Every one of these is something a naive rewriter drops, and
 * several of them — data validation and sheet protection especially — are the
 * kind of loss that matters when the file is going to a bank.
 *
 * Synthetic, and honest about it: this proves the MECHANISM preserves parts it
 * does not understand. Only real customer files prove the claim, which is what
 * `npm run fidelity -- <dir>` is for.
 */
import { writeZip, ZipEntry } from '@rutba/ooxml';

const entry = (name, data, method = 8) => {
  const e = new ZipEntry({
    name, method, crc: 0, compressedSize: 0, uncompressedSize: 0,
    compressed: Buffer.alloc(0), flags: 0, dosTime: 0, dosDate: 0x2821,
    externalAttrs: 0, comment: Buffer.alloc(0),
  });
  // Setting .data marks the entry modified, which is what makes writeZip
  // compress it. The fixture is "pristine" only once it has been written and
  // read back — which is exactly how a real file arrives.
  e.data = Buffer.isBuffer(data) ? data : Buffer.from(data, 'utf8');
  return e;
};

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** A tiny but genuinely binary part, so the test covers non-XML preservation. */
function pngBytes() {
  return Buffer.from(
    '89504e470d0a1a0a0000000d494844520000000100000001080600000' +
    '01f15c4890000000a49444154789c6360000002000100ffff03000006000557bfabd40000000049454e44ae426082',
    'hex',
  );
}

/**
 * `protect: false` builds the same workbook without its `<sheetProtection>`.
 * The default stays protected, because a customer file CAN arrive that way
 * and the fidelity tests should face it — but suites about editing semantics
 * (undo, history) are not about protection and should not have to thaw it.
 */
export function buildComplexWorkbook({ protect = true } = {}) {
  const contentTypes = XML +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Default Extension="png" ContentType="image/png"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
    '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
    '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' +
    '<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' +
    '<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>' +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
    '</Types>';

  const rootRels = XML +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
    '</Relationships>';

  const workbook = XML +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="26925"/>' +
    '<workbookPr defaultThemeVersion="166925"/>' +
    '<sheets>' +
    '<sheet name="Stock" sheetId="1" r:id="rId1"/>' +
    '<sheet name="Summary" sheetId="2" r:id="rId2"/>' +
    '</sheets>' +
    '<definedNames>' +
    '<definedName name="EXISTING_RANGE">Stock!$A$1:$C$3</definedName>' +
    '<definedName name="_xlnm.Print_Area" localSheetId="0">Stock!$A$1:$D$20</definedName>' +
    '</definedNames>' +
    '<calcPr calcId="191029"/>' +
    '</workbook>';

  const workbookRels = XML +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>' +
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' +
    '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '</Relationships>';

  // Everything outside <sheetData> here is what a careless rewriter loses.
  const sheet1 = XML +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<sheetPr filterMode="false"><pageSetUpPr fitToPage="true"/></sheetPr>' +
    '<dimension ref="A1:D20"/>' +
    '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
    '<sheetFormatPr defaultRowHeight="15"/>' +
    '<cols><col min="1" max="1" width="18.5" customWidth="1"/><col min="2" max="4" width="12.75" customWidth="1"/></cols>' +
    '<sheetData>' +
    '<row r="1" spans="1:4" s="2" customFormat="1" ht="22.5">' +
      '<c r="A1" s="2" t="s"><v>0</v></c>' +
      '<c r="B1" s="2" t="s"><v>1</v></c>' +
      '<c r="C1" s="2" t="s"><v>2</v></c>' +
    '</row>' +
    '<row r="2" spans="1:4">' +
      '<c r="A2" t="s"><v>3</v></c>' +
      '<c r="B2" s="4"><v>1420</v></c>' +
      '<c r="C2" s="5"><f>B2*1.17</f><v>1661.4</v></c>' +
    '</row>' +
    '<row r="3" spans="1:4">' +
      '<c r="A3" t="s"><v>4</v></c>' +
      '<c r="B3" s="4"><v>860</v></c>' +
      '<c r="C3" s="5"><f>B3*1.17</f><v>1006.2</v></c>' +
    '</row>' +
    // A merged banner with a themed fill, and a wrapped note beneath it —
    // between them they exercise merge spanning, theme colours, tint, wrapping,
    // vertical alignment and indent.
    '<row r="20" spans="1:4"><c r="A20" s="6" t="inlineStr"><is><t>Warehouse stock — WH1</t></is></c></row>' +
    '<row r="22" spans="1:4"><c r="A22" s="7" t="inlineStr"><is><t>Figures refresh from ERP Inventory on open.</t></is></c></row>' +
    '</sheetData>' +
    (protect ? '<sheetProtection sheet="1" objects="1" scenarios="1" selectLockedCells="1"/>' : '') +
    '<autoFilter ref="A1:D3"/>' +
    '<mergeCells count="2"><mergeCell ref="A20:D20"/><mergeCell ref="A22:D23"/></mergeCells>' +
    '<conditionalFormatting sqref="B2:B3"><cfRule type="cellIs" dxfId="0" priority="1" operator="lessThan"><formula>100</formula></cfRule></conditionalFormatting>' +
    '<dataValidations count="1"><dataValidation type="whole" operator="greaterThanOrEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="Invalid" error="Quantity cannot be negative" sqref="B2:B100"><formula1>0</formula1></dataValidation></dataValidations>' +
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
    '<pageSetup paperSize="9" orientation="landscape" r:id="rId9" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>' +
    '<drawing r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>' +
    // An extension Excel does NOT recognise, carrying our own namespace — the
    // preservation challenge is that it must ride through untouched. It was an
    // EMPTY x14:conditionalFormattings under the x14 URI until 2026-08-31,
    // and real Excel REFUSED the whole file over it: a consumer ignores an
    // unknown ext but validates one whose URI it owns, and the x14 schema
    // requires at least one child. The Office gate is what caught it.
    '<extLst><ext uri="{6E8A54F2-1D3B-4C0A-9E27-52AB1C4D9F70}"><rutba:sheetMeta xmlns:rutba="urn:rutba:erp" refreshedBy="erp"/></ext></extLst>' +
    '</worksheet>';

  const sheet2 = XML +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<dimension ref="A1"/><sheetData><row r="1"><c r="A1" t="s"><v>5</v></c></row></sheetData>' +
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
    '</worksheet>';

  const sheet1Rels = XML +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>' +
    '</Relationships>';

  const sharedStrings = XML +
    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="6" uniqueCount="6">' +
    '<si><t>Item</t></si><si><t>On hand</t></si><si><t>With tax</t></si>' +
    '<si><t>Steel bracket 40mm</t></si><si><t>Steel bracket 60mm</t></si><si><t>Summary</t></si>' +
    '</sst>';

  const styles = XML +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00&quot; PKR&quot;"/></numFmts>' +
    '<fonts count="5"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font><font><i/><sz val="11"/><name val="Calibri"/></font>' +
      '<font><b/><sz val="14"/><color theme="0"/><name val="Calibri"/></font>' +
      '<font><sz val="11"/><color rgb="FFC00000"/><u/><name val="Calibri"/></font></fonts>' +
    '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/></patternFill></fill>' +
      '<fill><patternFill patternType="solid"><fgColor theme="4"/><bgColor indexed="64"/></patternFill></fill>' +
      '<fill><patternFill patternType="solid"><fgColor theme="4" tint="0.8"/><bgColor indexed="64"/></patternFill></fill></fills>' +
    '<borders count="3"><border><left/><right/><top/><bottom/><diagonal/></border>' +
      '<border><left style="thin"><color rgb="FF7F7F7F"/></left><right style="thin"><color rgb="FF7F7F7F"/></right>' +
        '<top style="thin"><color rgb="FF7F7F7F"/></top><bottom style="thin"><color rgb="FF7F7F7F"/></bottom></border>' +
      '<border><top style="none"/><bottom style="double"><color theme="4"/></bottom></border></borders>' +
    '<cellXfs count="9">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="1" fillId="1" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>' +
    '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="3" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>' +
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>' +
    '<xf numFmtId="0" fontId="3" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1">' +
      '<alignment horizontal="center" vertical="center"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="3" borderId="2" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1">' +
      '<alignment horizontal="left" vertical="top" wrapText="1" indent="1"/></xf>' +
    '<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="0"/>' +
    '</cellXfs>' +
    '<dxfs count="1"><dxf><font><color rgb="FF9C0006"/></font></dxf></dxfs>' +
    '</styleSheet>';

  const theme = XML +
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">' +
    '<a:themeElements><a:clrScheme name="Office">' +
    '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>' +
    '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
    '<a:dk2><a:srgbClr val="44546A"/></a:dk2>' +
    '<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>' +
    '<a:accent1><a:srgbClr val="4472C4"/></a:accent1>' +
    '<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>' +
    '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>' +
    '<a:accent4><a:srgbClr val="FFC000"/></a:accent4>' +
    '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>' +
    '<a:accent6><a:srgbClr val="70AD47"/></a:accent6>' +
    '<a:hlink><a:srgbClr val="0563C1"/></a:hlink>' +
    '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>' +
    '</a:clrScheme></a:themeElements></a:theme>';

  // Schema-complete since 2026-08-31: the anchors carry their colOff/rowOff,
  // the graphicFrame its cNvGraphicFramePr and xfrm, the shape its cNvSpPr
  // and bodyPr, the picture its cNvPicPr and spPr. The fixture used to omit
  // all of these and every one of our tests still passed — real Excel is the
  // consumer that refused, which is exactly what the Office gate is for. A
  // customer file is written BY Excel, so customer-shaped means valid here.
  const drawing = XML +
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<xdr:twoCellAnchor>' +
    '<xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
    '<xdr:to><xdr:col>12</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>16</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
    '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>' +
    '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
    '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>' +
    '</a:graphicData></a:graphic></xdr:graphicFrame>' +
    '<xdr:clientData/></xdr:twoCellAnchor>' +
    // A callout someone drew round a total, and the company logo — both as
    // common in a real workbook as the chart, and both previously invisible.
    '<xdr:twoCellAnchor>' +
    '<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>18</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
    '<xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>21</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
    '<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id=\"3\" name=\"Callout\"/><xdr:cNvSpPr/></xdr:nvSpPr>' +
    '<xdr:spPr><a:prstGeom prst=\"roundRect\"><a:avLst/></a:prstGeom>' +
    '<a:solidFill><a:schemeClr val=\"accent2\"/></a:solidFill></xdr:spPr>' +
    '<xdr:txBody><a:bodyPr/><a:p><a:r><a:rPr lang=\"en-GB\" b=\"1\"/><a:t>Confirm with supplier</a:t></a:r></a:p></xdr:txBody>' +
    '</xdr:sp><xdr:clientData/></xdr:twoCellAnchor>' +
    '<xdr:oneCellAnchor>' +
    '<xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>23</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
    '<xdr:ext cx=\"914400\" cy=\"457200\"/>' +
    '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"4\" name=\"Logo\"/><xdr:cNvPicPr/></xdr:nvPicPr>' +
    '<xdr:blipFill><a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"rId2\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>' +
    '<xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>';

  // Also schema-complete since 2026-08-31: the chartSpace declares xmlns:a
  // (its title always used a:p WITHOUT declaring it — not even well-formed,
  // and no test noticed because our reader is namespace-blind), the series
  // carry their order element, and the barChart names its two axes, which
  // ECMA-376 requires and Excel enforces.
  const chart = XML +
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<c:chart>' +
    '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Stock on hand</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>' +
    '<c:autoTitleDeleted val="0"/>' +
    '<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>' +
    '<c:ser><c:idx val="0"/><c:order val="0"/>' +
      '<c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>On hand</c:v></c:pt></c:strCache></c:strRef></c:tx>' +
      '<c:cat><c:strRef><c:f>Stock!$A$2:$A$3</c:f><c:strCache>' +
        '<c:pt idx="0"><c:v>Steel bracket 40mm</c:v></c:pt>' +
        '<c:pt idx="1"><c:v>Steel bracket 60mm</c:v></c:pt>' +
      '</c:strCache></c:strRef></c:cat>' +
      '<c:val><c:numRef><c:f>Stock!$B$2:$B$3</c:f><c:numCache>' +
        '<c:pt idx="0"><c:v>1420</c:v></c:pt><c:pt idx="1"><c:v>860</c:v></c:pt>' +
      '</c:numCache></c:numRef></c:val>' +
    '</c:ser>' +
    '<c:ser><c:idx val="1"/><c:order val="1"/>' +
      '<c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Reserved</c:v></c:pt></c:strCache></c:strRef></c:tx>' +
      '<c:val><c:numRef><c:f>Stock!$C$2:$C$3</c:f><c:numCache>' +
        '<c:pt idx="0"><c:v>120</c:v></c:pt><c:pt idx="1"><c:v>40</c:v></c:pt>' +
      '</c:numCache></c:numRef></c:val>' +
    '</c:ser>' +
    '<c:axId val="111111111"/><c:axId val="222222222"/>' +
    '</c:barChart>' +
    '<c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
    '<c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222222222"/></c:catAx>' +
    '<c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
    '<c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111111111"/></c:valAx>' +
    '</c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>';

  const drawingRels = XML +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>' +
    '</Relationships>';

  const pivotCache = XML +
    '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="2">' +
    '<cacheSource type="worksheet"><worksheetSource ref="A1:C3" sheet="Stock"/></cacheSource>' +
    '<cacheFields count="3"><cacheField name="Item" numFmtId="0"/><cacheField name="On hand" numFmtId="0"/><cacheField name="With tax" numFmtId="0"/></cacheFields>' +
    '</pivotCacheDefinition>';

  const customXml = XML + '<rutba:meta xmlns:rutba="urn:rutba:erp"><binding query="stock.levels" range="EXISTING_RANGE"/></rutba:meta>';

  const core = XML +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
    'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/">' +
    '<dc:title>Stock report</dc:title><dc:creator>Rutba ERP</dc:creator>' +
    '<dcterms:created xsi:type="dcterms:W3CDTF" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">2026-08-01T09:00:00Z</dcterms:created>' +
    '</cp:coreProperties>';

  const app = XML +
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">' +
    '<Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><AppVersion>16.0300</AppVersion>' +
    '</Properties>';

  return writeZip([
    entry('[Content_Types].xml', contentTypes),
    entry('_rels/.rels', rootRels),
    entry('xl/workbook.xml', workbook),
    entry('xl/_rels/workbook.xml.rels', workbookRels),
    entry('xl/worksheets/sheet1.xml', sheet1),
    entry('xl/worksheets/_rels/sheet1.xml.rels', sheet1Rels),
    entry('xl/worksheets/sheet2.xml', sheet2),
    entry('xl/sharedStrings.xml', sharedStrings),
    entry('xl/styles.xml', styles),
    entry('xl/theme/theme1.xml', theme),
    entry('xl/drawings/drawing1.xml', drawing),
    entry('xl/drawings/_rels/drawing1.xml.rels', drawingRels),
    entry('xl/charts/chart1.xml', chart),
    entry('xl/pivotCache/pivotCacheDefinition1.xml', pivotCache),
    entry('xl/media/image1.png', pngBytes(), 0), // stored, not deflated
    entry('customXml/item1.xml', customXml),
    entry('docProps/core.xml', core),
    entry('docProps/app.xml', app),
  ]);
}

/** The parts a careless implementation loses, named so tests can assert on them. */
export const FRAGILE_PARTS = [
  'xl/styles.xml',
  'xl/sharedStrings.xml',
  'xl/theme/theme1.xml',
  'xl/drawings/drawing1.xml',
  'xl/drawings/_rels/drawing1.xml.rels',
  'xl/charts/chart1.xml',
  'xl/pivotCache/pivotCacheDefinition1.xml',
  'xl/media/image1.png',
  'customXml/item1.xml',
  'docProps/core.xml',
  'docProps/app.xml',
  'xl/worksheets/sheet2.xml',
];

/** Elements INSIDE the edited sheet that must survive the edit. */
export const FRAGILE_SHEET_ELEMENTS = [
  // sheetProtection GRADUATED out of this list on 2026-08-31: it is modelled
  // and enforced now, and editing the fixture requires deliberately removing
  // it first — its preservation-under-refusal is pinned by its own tests.
  'dataValidations',
  'conditionalFormatting',
  'mergeCells',
  'autoFilter',
  'cols',
  'pageSetup',
  'extLst',
  'pane',
];
