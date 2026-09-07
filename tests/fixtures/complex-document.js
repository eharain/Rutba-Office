/**
 * A .docx carrying the things we do NOT model.
 *
 * Styles, numbering, a header, a footer, footnotes, comments, a content control,
 * a bookmark, a field code, a hyperlink relationship, binary media, custom XML
 * and a section properties block. Every one of these is something a careless
 * rewriter drops, and several — the field code and the content control
 * especially — are how a real template carries its intelligence.
 *
 * Shaped like a document that would actually leave the building: a bank letter
 * with a reference number in a field, an amount in a content control, and a
 * numbered schedule.
 */
import { writeZip, ZipEntry } from '@rutba/ooxml';

const entry = (name, data, method = 8) => {
  const e = new ZipEntry({
    name, method, crc: 0, compressedSize: 0, uncompressedSize: 0,
    compressed: Buffer.alloc(0), flags: 0, dosTime: 0, dosDate: 0x2821,
    externalAttrs: 0, comment: Buffer.alloc(0),
  });
  e.data = Buffer.isBuffer(data) ? data : Buffer.from(data, 'utf8');
  return e;
};

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const pngBytes = () => Buffer.from(
  '89504e470d0a1a0a0000000d494844520000000100000001080600000' +
  '01f15c4890000000a49444154789c6360000002000100ffff03000006000557bfabd40000000049454e44ae426082',
  'hex',
);

export function buildComplexDocument() {
  const contentTypes = XML +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Default Extension="png" ContentType="image/png"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
    '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
    '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' +
    '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>' +
    '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>' +
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    '</Types>';

  const rootRels = XML +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
    '</Relationships>';

  const documentXml = XML +
    '<w:document ' + W + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<w:body>' +

    // heading, with a paragraph style we must not disturb
    '<w:p w:rsidR="00A1" w:rsidRDefault="00A1">' +
      '<w:pPr><w:pStyle w:val="Heading1"/><w:spacing w:before="240" w:after="120"/></w:pPr>' +
      '<w:r><w:rPr><w:b/><w:sz w:val="32"/></w:rPr><w:t>Confirmation of balance</w:t></w:r>' +
    '</w:p>' +

    // a field code — the reference number. Losing this breaks the template.
    '<w:p>' +
      '<w:pPr><w:jc w:val="right"/></w:pPr>' +
      '<w:r><w:t xml:space="preserve">Ref: </w:t></w:r>' +
      '<w:fldSimple w:instr=" DOCPROPERTY &quot;RefNo&quot; \\* MERGEFORMAT ">' +
        '<w:r><w:rPr><w:noProof/></w:rPr><w:t>RB-2026-0001</w:t></w:r>' +
      '</w:fldSimple>' +
    '</w:p>' +

    // a bookmark around body text
    '<w:p>' +
      '<w:bookmarkStart w:id="1" w:name="Salutation"/>' +
      '<w:r><w:rPr><w:i/></w:rPr><w:t>Dear Sir or Madam,</w:t></w:r>' +
      '<w:bookmarkEnd w:id="1"/>' +
    '</w:p>' +

    '<w:p>' +
      '<w:r><w:t xml:space="preserve">We confirm the balance outstanding as at the date below is </w:t></w:r>' +
      // a content control — the docx equivalent of a named range, and our binding target
      '<w:sdt>' +
        '<w:sdtPr>' +
          '<w:alias w:val="Outstanding balance"/>' +
          '<w:tag w:val="RUTBA_AR_TOTAL"/>' +
          '<w:id w:val="123456789"/>' +
          '<w:text/>' +
        '</w:sdtPr>' +
        '<w:sdtContent><w:r><w:rPr><w:b/></w:rPr><w:t>0.00</w:t></w:r></w:sdtContent>' +
      '</w:sdt>' +
      '<w:r><w:t xml:space="preserve"> PKR.</w:t></w:r>' +
    '</w:p>' +

    // a second content control, for the as-of date
    '<w:p>' +
      '<w:sdt>' +
        '<w:sdtPr><w:alias w:val="As of"/><w:tag w:val="RUTBA_AS_OF"/><w:id w:val="123456790"/><w:text/></w:sdtPr>' +
        '<w:sdtContent><w:r><w:t>1970-01-01</w:t></w:r></w:sdtContent>' +
      '</w:sdt>' +
    '</w:p>' +

    // a numbered list, driven by numbering.xml
    '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
      '<w:r><w:t>Invoices raised in the period</w:t></w:r></w:p>' +
    '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
      '<w:r><w:t>Payments received</w:t></w:r></w:p>' +

    // a table — a repeating header row, shading, a vertical merge, a spanning
    // total row and a nested table, because a real statement has all of them and
    // each one is a different way for a renderer to get it wrong
    '<w:tbl>' +
      '<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="5000" w:type="pct"/>' +
        '<w:tblBorders><w:top w:val="single" w:sz="4"/>' +
        '<w:insideH w:val="single" w:sz="4" w:color="D9D9D9"/></w:tblBorders></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="4508"/><w:gridCol w:w="2254"/><w:gridCol w:w="2254"/></w:tblGrid>' +
      '<w:tr><w:trPr><w:tblHeader/></w:trPr>' +
        '<w:tc><w:tcPr><w:tcW w:w="4508" w:type="dxa"/><w:shd w:fill="D9E2F3"/></w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Description</w:t></w:r></w:p></w:tc>' +
        '<w:tc><w:tcPr><w:tcW w:w="2254" w:type="dxa"/><w:shd w:fill="D9E2F3"/></w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Ref</w:t></w:r></w:p></w:tc>' +
        '<w:tc><w:tcPr><w:tcW w:w="2254" w:type="dxa"/><w:shd w:fill="D9E2F3"/></w:tcPr><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Amount</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr>' +
        '<w:tc><w:tcPr><w:tcW w:w="4508" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>Opening balance</w:t></w:r></w:p></w:tc>' +
        '<w:tc><w:tcPr><w:tcW w:w="2254" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>INV-1001</w:t></w:r></w:p></w:tc>' +
        '<w:tc><w:tcPr><w:tcW w:w="2254" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>1,842,600</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr>' +
        '<w:tc><w:tcPr><w:tcW w:w="4508" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc>' +
        '<w:tc><w:tcPr><w:tcW w:w="2254" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>INV-1002</w:t></w:r></w:p>' +
          // a nested table, inside a cell
          '<w:tbl><w:tblGrid><w:gridCol w:w="1127"/></w:tblGrid>' +
            '<w:tr><w:tc><w:p><w:r><w:t>part-paid</w:t></w:r></w:p></w:tc></w:tr>' +
          '</w:tbl>' +
          '<w:p><w:r><w:t>see note 4</w:t></w:r></w:p></w:tc>' +
        '<w:tc><w:tcPr><w:tcW w:w="2254" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>240,000</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr>' +
        '<w:tc><w:tcPr><w:tcW w:w="6762" w:type="dxa"/><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Total</w:t></w:r></w:p></w:tc>' +
        '<w:tc><w:tcPr><w:tcW w:w="2254" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>2,082,600</w:t></w:r></w:p></w:tc></w:tr>' +
    '</w:tbl>' +

    // the company logo, inline — an embedded picture with a real extent.
    // Schema-complete since 2026-08-31: pic:pic requires nvPicPr, blipFill
    // AND spPr, and the fixture shipped blipFill alone — every one of our
    // tests passed over it and real Word refused the whole file. The shape
    // below mirrors what insertImageParagraph writes, which was always
    // complete; only the fixture was truncated. The Office gate caught it.
    '<w:p><w:r><w:drawing>' +
      '<wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
      '<wp:extent cx="914400" cy="457200"/><wp:docPr id="7" name="Company logo"/>' +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:nvPicPr><pic:cNvPr id="7" name="Company logo"/><pic:cNvPicPr/></pic:nvPicPr>' +
      '<pic:blipFill>' +
      '<a:blip r:embed="rId8" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>' +
      '<a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
      '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
      '</pic:pic></a:graphicData></a:graphic></wp:inline>' +
    '</w:drawing></w:r></w:p>' +

    // a paragraph carrying a comment reference
    '<w:p>' +
      '<w:commentRangeStart w:id="0"/>' +
      '<w:r><w:t>Subject to audit.</w:t></w:r>' +
      '<w:commentRangeEnd w:id="0"/>' +
      '<w:r><w:commentReference w:id="0"/></w:r>' +
    '</w:p>' +

    // section properties: margins, header/footer wiring, page size
    '<w:sectPr w:rsidR="00A1">' +
      '<w:headerReference w:type="default" r:id="rIdHdr"/>' +
      '<w:footerReference w:type="default" r:id="rIdFtr"/>' +
      '<w:footnotePr><w:numFmt w:val="decimal"/></w:footnotePr>' +
      '<w:pgSz w:w="11906" w:h="16838"/>' +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>' +
      '<w:cols w:space="708"/>' +
      '<w:docGrid w:linePitch="360"/>' +
    '</w:sectPr>' +
    '</w:body></w:document>';

  const documentRels = XML +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' +
    '<Relationship Id="rIdHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' +
    '<Relationship Id="rIdFtr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>' +
    '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>' +
    '<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>' +
    '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>' +
    '<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo.png"/>' +
    '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://rutba.io" TargetMode="External"/>' +
    '</Relationships>';

  const styles = XML +
    '<w:styles ' + W + '>' +
    '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr>' +
      '<w:rPr><w:b/><w:sz w:val="32"/><w:color w:val="1F5F8B"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:pPr><w:ind w:left="720"/></w:pPr></w:style>' +
    '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style>' +
    '</w:styles>';

  const numbering = XML +
    '<w:numbering ' + W + '>' +
    '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
    '</w:numbering>';

  const header = XML + '<w:hdr ' + W + '><w:p><w:r><w:t>Rutba Trading Company</w:t></w:r></w:p></w:hdr>';
  const footer = XML + '<w:ftr ' + W + '><w:p><w:r><w:t>Page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>';
  const footnotes = XML + '<w:footnotes ' + W + '><w:footnote w:id="1"><w:p><w:r><w:t>Balances exclude disputed items.</w:t></w:r></w:p></w:footnote></w:footnotes>';
  const comments = XML + '<w:comments ' + W + '><w:comment w:id="0" w:author="Finance" w:date="2026-08-01T10:00:00Z"><w:p><w:r><w:t>Confirm with the ledger before sending.</w:t></w:r></w:p></w:comment></w:comments>';
  const settings = XML + '<w:settings ' + W + '><w:zoom w:percent="100"/><w:trackChanges/><w:defaultTabStop w:val="720"/></w:settings>';

  const core = XML +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
    'xmlns:dc="http://purl.org/dc/elements/1.1/">' +
    '<dc:title>Balance confirmation</dc:title><dc:creator>Rutba ERP</dc:creator></cp:coreProperties>';

  return writeZip([
    entry('[Content_Types].xml', contentTypes),
    entry('_rels/.rels', rootRels),
    entry('word/document.xml', documentXml),
    entry('word/_rels/document.xml.rels', documentRels),
    entry('word/styles.xml', styles),
    entry('word/numbering.xml', numbering),
    entry('word/header1.xml', header),
    entry('word/footer1.xml', footer),
    entry('word/footnotes.xml', footnotes),
    entry('word/comments.xml', comments),
    entry('word/settings.xml', settings),
    entry('word/media/logo.png', pngBytes(), 0),
    entry('docProps/core.xml', core),
  ]);
}

/** Parts a careless implementation loses. */
export const FRAGILE_DOC_PARTS = [
  'word/styles.xml',
  'word/numbering.xml',
  'word/header1.xml',
  'word/footer1.xml',
  'word/footnotes.xml',
  'word/comments.xml',
  'word/settings.xml',
  'word/media/logo.png',
  'word/_rels/document.xml.rels',
  'docProps/core.xml',
];

/** Elements inside document.xml that must survive an edit. */
export const FRAGILE_DOC_ELEMENTS = [
  'w:sectPr',
  'w:headerReference',
  'w:footerReference',
  'w:pgMar',
  'w:fldSimple',
  'w:bookmarkStart',
  'w:commentRangeStart',
  'w:numPr',
  'w:tbl',
  'w:tblBorders',
  'w:pStyle',
];
