// Fixtures the review checks share (word-redact, inspect, deck-redact,
// sheet-inspect): a small landscape as a PNG, and a document that keeps a
// name, an address and a telephone number everywhere Word can keep words.

import zlib from 'node:zlib';
import { OoxmlPackage, crc32, buildDocx } from '@rutba/ooxml';

/** A small landscape as a PNG: sky, a sun, a green hill. */
export function landscapePng(w = 200, h = 140) {
  const rows = [];
  for (let y = 0; y < h; y++) {
    const row = Buffer.alloc(1 + w * 3);
    for (let x = 0; x < w; x++) {
      let c = [120 + Math.round((y / h) * 80), 170 + Math.round((y / h) * 50), 235];
      if ((x - w * 0.74) ** 2 + (y - h * 0.27) ** 2 < (h * 0.12) ** 2) c = [250, 210, 80];
      if (y > h * 0.62 + h * 0.13 * Math.sin(x / (w / 6))) c = [70, 140 - Math.round((y - h * 0.6) / 2), 70];
      row.set(c, 1 + x * 3);
    }
    rows.push(row);
  }
  const chunk = (type, data) => {
    const len = Buffer.alloc(4);
    len.writeUInt32BE(data.length);
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crc32(Buffer.concat([Buffer.from(type), data])) >>> 0);
    return Buffer.concat([len, Buffer.from(type), data, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0);
  ihdr.writeUInt32BE(h, 4);
  ihdr[8] = 8;
  ihdr[9] = 2;
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(Buffer.concat(rows))), chunk('IEND', Buffer.alloc(0))]);
}

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/**
 * A letter that names Ayesha Khan, her address and her telephone number in
 * the body (a run apart), a header, a footer, a footnote, a comment and its
 * author, a tracked insertion and deletion, a link's address, a picture's
 * alt text and the file's properties, with a picture of a landscape.
 */
export function secretLetter() {
  const pkg = OoxmlPackage.read(buildDocx({ styles: true, paragraphs: [{ text: 'Offer letter' }] }));
  const png = landscapePng();
  const body = '<w:p><w:r><w:t xml:space="preserve">Dear Ayesha </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>Khan</w:t></w:r><w:r><w:t xml:space="preserve">, we write to ayesha.khan@example.pk and +92 300 1234567 about INV-2026.</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>'
    + '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t xml:space="preserve">Terms agreed </w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>'
    + '<w:ins w:id="1" w:author="Ayesha Khan" w:date="2026-10-01T10:00:00Z"><w:r><w:t xml:space="preserve">by Ayesha Khan </w:t></w:r></w:ins><w:del w:id="2" w:author="Reviewer" w:date="2026-10-01T10:00:00Z"><w:r><w:delText>and ayesha.k@old.example.pk</w:delText></w:r></w:del></w:p>'
    + '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> HYPERLINK "mailto:ayesha.khan@example.pk" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Write to her</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    + '<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="1905000" cy="1333500"/><wp:docPr id="7" name="Picture 7" descr="Ayesha Khan at her desk"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="7" name="Picture 7" descr="Ayesha Khan at her desk"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdPic" xmlns:r="' + REL + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="1333500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
  pkg.write_('word/document.xml', pkg.text('word/document.xml').replace('<w:document ', '<w:document xmlns:r="' + REL + '" ').replace('<w:sectPr/>', body + '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/></w:sectPr>'));
  const part = (name, xml, type) => pkg.addPart(name, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + xml, type);
  part('word/header1.xml', `<w:hdr ${W}><w:p><w:r><w:t>Prepared for Ayesha Khan</w:t></w:r></w:p></w:hdr>`, 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml');
  part('word/footer1.xml', `<w:ftr ${W}><w:p><w:r><w:t xml:space="preserve">Questions: +92 300 1234567</w:t></w:r></w:p></w:ftr>`, 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml');
  part('word/comments.xml', `<w:comments ${W}><w:comment w:id="0" w:author="Ayesha Khan" w:initials="AK"><w:p><w:r><w:t>Agreed with Ayesha Khan on the phone</w:t></w:r></w:p></w:comment></w:comments>`, 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml');
  part('word/footnotes.xml', `<w:footnotes ${W}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Her own address is ayesha.khan@example.pk.</w:t></w:r></w:p></w:footnote></w:footnotes>`, 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml');
  pkg.ensureDefault('png', 'image/png');
  pkg.addPart('word/media/image1.png', png);
  pkg.write_('word/_rels/document.xml.rels', pkg.text('word/_rels/document.xml.rels').replace('</Relationships>', `<Relationship Id="rIdPic" Type="${REL}/image" Target="media/image1.png"/><Relationship Id="rIdH" Type="${REL}/header" Target="header1.xml"/><Relationship Id="rIdF" Type="${REL}/footer" Target="footer1.xml"/><Relationship Id="rIdC" Type="${REL}/comments" Target="comments.xml"/><Relationship Id="rIdN" Type="${REL}/footnotes" Target="footnotes.xml"/></Relationships>`));
  part('docProps/core.xml', '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Offer for Ayesha Khan</dc:title><dc:creator>Ayesha Khan</dc:creator><cp:lastModifiedBy>Bilal Ahmed</cp:lastModifiedBy></cp:coreProperties>', 'application/vnd.openxmlformats-package.core-properties+xml');
  pkg.write_('_rels/.rels', pkg.text('_rels/.rels').replace('</Relationships>', '<Relationship Id="rIdCore" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>'));
  return pkg.write();
}

/**
 * A deck that names Ayesha Khan on a slide, in its notes, a comment and a
 * picture's alt text, with her address and a picture of a landscape; a
 * hidden box and one parked off the slide.
 */
export async function secretDeck() {
  const { Deck, buildPptx } = await import('@rutba/presentation');
  const d = Deck.open(buildPptx({ title: 'Offer for Ayesha Khan', author: 'Bilal Ahmed', slides: [
    { layout: 'obj', title: 'Offer', body: ['For Ayesha Khan', 'Write to ayesha.khan@example.pk'] },
    { layout: 'obj', title: 'Terms', body: ['Start in March'] },
  ] }));
  d.addPicture(0, { data: landscapePng(), contentType: 'image/png', name: 'Photo', x: 620, y: 300, w: 240, h: 168, descr: 'Ayesha Khan at her desk' });
  d.setNotes(0, 'Remember: Ayesha Khan prefers mornings.');
  d.addComment(1, { text: 'Check with Ayesha Khan', author: 'Sana Iqbal' });
  const box = d.addTextBox(1, { x: 60, y: 60, w: 200, h: 40, paragraphs: [{ runs: [{ text: 'Draft only' }] }], name: 'Hidden note' });
  d.setShapeHidden(1, box, true);
  d.addTextBox(1, { x: 5000, y: 60, w: 200, h: 40, paragraphs: [{ runs: [{ text: 'Parked outside' }] }], name: 'Parked' });
  return Buffer.from(d.save());
}

/** A workbook with a name and telephone numbers in cells, a note, a header, a hidden sheet, a hidden row and column, a hidden name and an author. */
export async function secretWorkbook() {
  const { buildXlsx } = await import('@rutba/ooxml');
  const { SheetView } = await import('@rutba/sheet-view');
  const pkg = OoxmlPackage.read(buildXlsx({ sheets: [
    { name: 'Staff', rows: [['Name', 'Phone', 'Pay'], ['Ayesha Khan', '+92 300 1234567', 120000], ['Bilal Ahmed', '+92 321 7654321', 95000], ['Sana Iqbal', '+92 333 1112223', 99000]] },
    { name: 'Salaries', rows: [['Ayesha Khan', 120000]] },
  ] }));
  const s1 = 'xl/worksheets/sheet1.xml';
  let x = pkg.text(s1);
  x = x.replace(/<row r="4"/, '<row r="4" hidden="1"').replace('<sheetData>', '<cols><col min="3" max="3" width="10" hidden="1" customWidth="1"/></cols><sheetData>');
  x = x.replace('</worksheet>', '<headerFooter><oddHeader>&amp;LPrepared by Ayesha Khan</oddHeader></headerFooter></worksheet>');
  pkg.write_(s1, x);
  let wb = pkg.text('xl/workbook.xml');
  wb = wb.replace(/(<sheet\b[^>]*name="Salaries"[^>]*)\/>/, '$1 state="hidden"/>');
  wb = wb.includes('<definedNames>') ? wb.replace('<definedNames>', '<definedNames><definedName name="Secret" hidden="1">Staff!$C$2</definedName>') : wb.replace('</sheets>', '</sheets><definedNames><definedName name="Secret" hidden="1">Staff!$C$2</definedName></definedNames>');
  pkg.write_('xl/workbook.xml', wb);
  if (!pkg.has('docProps/core.xml')) {
    pkg.addPart('docProps/core.xml', '<?xml version="1.0"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>Ayesha Khan</dc:creator><dc:title>Payroll</dc:title></cp:coreProperties>', 'application/vnd.openxmlformats-package.core-properties+xml');
    pkg.write_('_rels/.rels', pkg.text('_rels/.rels').replace('</Relationships>', '<Relationship Id="rIdCore" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>'));
  }
  const view = SheetView.open(Buffer.from(pkg.write()));
  view.setNote({ row: 1, col: 0, text: 'Ayesha Khan joined in May', author: 'Sana Iqbal' });
  return Buffer.from(view.save());
}
