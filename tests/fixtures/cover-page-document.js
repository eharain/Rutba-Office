/**
 * A .docx shaped like a tender's first pages — the constructs a page has to
 * DRAW rather than merely preserve.
 *
 * A cover page in a body-level content control, whose paragraph anchors a
 * text box filled from the theme; a table of contents in another control,
 * with the dotted right tab on its style; a paragraph with its own tab stops,
 * shading and a rule beneath; a footnote reference and the note it points at;
 * fonts named through the theme, so that without the theme part every word
 * would be Calibri; a first-page header switched on by titlePg.
 *
 * Every one of these was invisible on the page before it was read: the
 * Leicestershire LMS tender opened on its second page, the Selection
 * Questionnaire had no footnotes, and everything was Calibri.
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
const NS = W +
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
  ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
  ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
  ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"' +
  ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
  ' xmlns:v="urn:schemas-microsoft-com:vml"';

/** The text box: a shape filled with accent1 lightened, carrying two paragraphs. */
const textBox = () =>
  '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>' +
    '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
      '<wp:simplePos x="0" y="0"/>' +
      '<wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>' +
      '<wp:positionV relativeFrom="page"><wp:posOffset>1828800</wp:posOffset></wp:positionV>' +
      '<wp:extent cx="5486400" cy="1371600"/>' +
      '<wp:wrapTopAndBottom/>' +
      '<wp:docPr id="1" name="Title Box"/>' +
      '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
        '<wps:wsp><wps:cNvSpPr txBox="1"/>' +
          '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5486400" cy="1371600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
            '<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr></a:solidFill>' +
            '<a:ln w="6350"><a:solidFill><a:srgbClr val="1F2123"/></a:solidFill></a:ln>' +
          '</wps:spPr>' +
          '<wps:txbx><w:txbxContent>' +
            '<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:rPr><w:color w:val="FFFFFF"/><w:sz w:val="96"/></w:rPr><w:t>Learning Management System</w:t></w:r></w:p>' +
            '<w:p><w:pPr><w:pStyle w:val="Subtitle"/><w:jc w:val="right"/></w:pPr><w:r><w:t>Scope of requirements, 2017</w:t></w:r></w:p>' +
          '</w:txbxContent></wps:txbx>' +
        '</wps:wsp>' +
      '</a:graphicData></a:graphic>' +
    '</wp:anchor>' +
  '</w:drawing></mc:Choice>' +
  // The fallback repeats the words for older Words; a reader must take one.
  '<mc:Fallback><w:pict><v:rect style="width:432pt;height:108pt" fillcolor="#7E97AD"><v:textbox><w:txbxContent>' +
    '<w:p><w:r><w:t>Learning Management System</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t>Scope of requirements, 2017</w:t></w:r></w:p>' +
  '</w:txbxContent></v:textbox></v:rect></w:pict></mc:Fallback>' +
  '</mc:AlternateContent></w:r>';

export function buildCoverPageDocument() {
  const contentTypes = XML +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
    '<Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
    '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>' +
    '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
    '</Types>';

  const rootRels = XML +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>';

  const documentXml = XML +
    '<w:document ' + NS + '><w:body>' +

    // The cover page: a body-level content control whose one paragraph
    // anchors the title box. Word puts the whole cover in an sdt like this.
    '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Cover Pages"/><w:docPartUnique/></w:docPartObj><w:id w:val="1001"/></w:sdtPr><w:sdtContent>' +
      '<w:p><w:pPr><w:pStyle w:val="NoSpacing"/></w:pPr>' + textBox() + '<w:r><w:t>Cover</w:t></w:r></w:p>' +
      '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' +
    '</w:sdtContent></w:sdt>' +

    // The table of contents: another body-level control, entries on TOC1
    // whose style carries the dotted right tab.
    '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj><w:id w:val="1002"/></w:sdtPr><w:sdtContent>' +
      '<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr><w:r><w:t>Contents</w:t></w:r></w:p>' +
      // An entry as Word writes one: a hyperlink to the heading's bookmark,
      // every run wearing the Hyperlink style Word then does not draw.
      '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:hyperlink w:anchor="_Toc1" w:history="1">' +
        '<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Introduction</w:t></w:r>' +
        '<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:tab/><w:t>2</w:t></w:r></w:hyperlink></w:p>' +
    '</w:sdtContent></w:sdt>' +

    // A form line: its own tab stops, shading, a rule beneath, a hanging indent.
    '<w:p><w:pPr>' +
      '<w:tabs><w:tab w:val="left" w:pos="2160"/><w:tab w:val="right" w:leader="dot" w:pos="8640"/></w:tabs>' +
      '<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/>' +
      '<w:pBdr><w:bottom w:val="single" w:sz="8" w:space="1" w:color="1F5F8B"/></w:pBdr>' +
      '<w:ind w:left="720" w:hanging="360"/>' +
      '<w:rPr><w:shd w:val="clear" w:fill="FF0000"/></w:rPr>' +
    '</w:pPr>' +
      // The first tab in a run of its own, as Word writes it; the second
      // sharing a run with the words after it.
      '<w:r><w:t>Name:</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>Value</w:t></w:r><w:r><w:tab/><w:t>Page 3</w:t></w:r>' +
    '</w:p>' +

    // A heading on the theme's major face, body on the minor.
    '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Introduction</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t xml:space="preserve">The supplier must not meet any exclusion ground</w:t></w:r>' +
      '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r>' +
      '<w:r><w:t xml:space="preserve"> and must say so</w:t></w:r>' +
      '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>' +
      '<w:r><w:t>.</w:t></w:r>' +
    '</w:p>' +
    '<w:p><w:r><w:t>Plain, editable.</w:t></w:r></w:p>' +
    // Two links: one wearing the Hyperlink character style, one bare.
    '<w:p><w:hyperlink r:id="rIdLink"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>gov.uk guidance</w:t></w:r></w:hyperlink>' +
      '<w:r><w:t xml:space="preserve"> and </w:t></w:r>' +
      '<w:hyperlink w:anchor="_Toc1"><w:r><w:t>a plain TOC link</w:t></w:r></w:hyperlink></w:p>' +

    '<w:sectPr>' +
      '<w:headerReference w:type="default" r:id="rIdHdr"/>' +
      '<w:headerReference w:type="first" r:id="rIdHdrFirst"/>' +
      '<w:pgSz w:w="11906" w:h="16838"/>' +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="567" w:footer="567" w:gutter="0"/>' +
      '<w:titlePg/>' +
    '</w:sectPr>' +
    '</w:body></w:document>';

  const documentRels = XML +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '<Relationship Id="rIdHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' +
    '<Relationship Id="rIdHdrFirst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>' +
    '<Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>' +
    '<Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' +
    '<Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://www.gov.uk/guidance" TargetMode="External"/>' +
    '</Relationships>';

  const styles = XML + '<w:styles ' + W + '>' +
    '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:sz w:val="22"/><w:color w:val="595959"/></w:rPr></w:rPrDefault>' +
      '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="NoSpacing"><w:name w:val="No Spacing"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>' +
      '<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="0"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi"/><w:color w:val="2F5496"/><w:sz w:val="32"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/>' +
      '<w:pPr><w:pBdr><w:top w:val="single" w:sz="4" w:space="10" w:color="7E97AD"/></w:pBdr><w:shd w:val="clear" w:fill="7E97AD"/>' +
      '<w:spacing w:before="120" w:after="120" w:line="1200" w:lineRule="exact"/><w:ind w:left="115" w:right="115"/></w:pPr>' +
      '<w:rPr><w:caps/><w:color w:val="FFFFFF"/><w:sz w:val="136"/></w:rPr></w:style>' +
    '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>' +
    '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>' +
    '<w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:basedOn w:val="DefaultParagraphFont"/><w:rPr><w:b/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:rPr><w:i/><w:color w:val="7E97AD"/><w:sz w:val="32"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/><w:basedOn w:val="Heading1"/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/>' +
      '<w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9016"/></w:tabs><w:spacing w:after="100"/></w:pPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/></w:rPr></w:style>' +
    '</w:styles>';

  const theme = XML +
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Tender"><a:themeElements>' +
      '<a:clrScheme name="Tender"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
        '<a:dk2><a:srgbClr val="1F2123"/></a:dk2><a:lt2><a:srgbClr val="DC9E1F"/></a:lt2>' +
        '<a:accent1><a:srgbClr val="7E97AD"/></a:accent1><a:accent2><a:srgbClr val="CC8E60"/></a:accent2>' +
        '<a:accent3><a:srgbClr val="7A6A60"/></a:accent3><a:accent4><a:srgbClr val="B4936D"/></a:accent4>' +
        '<a:accent5><a:srgbClr val="67787B"/></a:accent5><a:accent6><a:srgbClr val="9D936F"/></a:accent6>' +
        '<a:hlink><a:srgbClr val="646464"/></a:hlink><a:folHlink><a:srgbClr val="969696"/></a:folHlink></a:clrScheme>' +
      '<a:fontScheme name="Tender"><a:majorFont><a:latin typeface="Georgia"/></a:majorFont><a:minorFont><a:latin typeface="Cambria"/></a:minorFont></a:fontScheme>' +
    '</a:themeElements></a:theme>';

  const header = XML + '<w:hdr ' + NS + '>' +
    // The watermark Word writes: a WordArt shape whose words are a text path.
    '<w:p><w:r><w:pict><v:shape id="PowerPlusWaterMarkObject1" style="position:absolute;width:527.85pt;height:131.95pt;rotation:315;z-index:-251657216" fillcolor="silver" stroked="f">' +
      '<v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" string="DRAFT"/></v:shape></w:pict></w:r></w:p>' +
    '<w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="9016"/></w:tabs></w:pPr>' +
    '<w:r><w:t>Selection Questionnaire</w:t></w:r><w:r><w:tab/><w:t xml:space="preserve">Page </w:t></w:r>' +
    '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:hdr>';
  const headerFirst = XML + '<w:hdr ' + W + '><w:p><w:r><w:t>Leicestershire Fire and Rescue Service</w:t></w:r></w:p></w:hdr>';

  const footnotes = XML + '<w:footnotes ' + W + '>' +
    '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
    '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
    '<w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> See PCR 2015 regulation 57.</w:t></w:r></w:p></w:footnote>' +
    '<w:footnote w:id="2"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> For the list of exclusions see the guidance.</w:t></w:r></w:p></w:footnote>' +
    '</w:footnotes>';

  return writeZip([
    entry('[Content_Types].xml', contentTypes),
    entry('_rels/.rels', rootRels),
    entry('word/document.xml', documentXml),
    entry('word/_rels/document.xml.rels', documentRels),
    entry('word/styles.xml', styles),
    entry('word/theme/theme1.xml', theme),
    entry('word/header1.xml', header),
    entry('word/header2.xml', headerFirst),
    entry('word/footnotes.xml', footnotes),
  ]);
}
