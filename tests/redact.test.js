// Redaction across a whole package: the marked words found wherever a
// document keeps them, the body (across runs), headers and footers, notes,
// comments and their authors, tracked insertions and deletions, a field's
// code and result, alt text, the document's properties, an embedded
// workbook, and replaced by black bars of the same length; a picture's
// bytes replaced; the thumbnail dropped; and the saved package read back
// with the words nowhere in it.
import test from 'node:test';
import assert from 'node:assert/strict';
import { buildDocx, buildXlsx, OoxmlPackage } from '@rutba/ooxml';
import { redactPackage, partsHolding, compileMarks, findSpans, barsOver, BAR, luhn } from '@rutba/ooxml/redact';
import { openDocx } from '@rutba/doc-view/backends/ooxml';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** A document that keeps a name, an address and a telephone number everywhere Word can keep words. */
export function secretDocx() {
  const pkg = OoxmlPackage.read(buildDocx({ styles: true, paragraphs: [{ text: 'Cover' }] }));
  const body = '<w:p><w:r><w:t xml:space="preserve">Client: Ayesha </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>Khan</w:t></w:r><w:r><w:t xml:space="preserve">, ayesha.khan@example.pk, +92 300 1234567.</w:t></w:r></w:p>'
    + '<w:p><w:ins w:id="1" w:author="Ayesha Khan" w:date="2026-10-01T10:00:00Z"><w:r><w:t>Added by Ayesha Khan.</w:t></w:r></w:ins><w:del w:id="2" w:author="Reviewer" w:date="2026-10-01T10:00:00Z"><w:r><w:delText>Old card 4111 1111 1111 1111</w:delText></w:r></w:del></w:p>'
    + '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> HYPERLINK "mailto:ayesha.khan@example.pk" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>write to Ayesha Khan</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    + '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Commented</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>'
    + '<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="95250" cy="95250"/><wp:docPr id="7" name="Photo of Ayesha Khan" descr="Ayesha Khan at her desk"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="7" name="Photo" descr="Ayesha Khan at her desk"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdPic" xmlns:r="' + REL + '"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
  pkg.write_('word/document.xml', pkg.text('word/document.xml').replace('<w:sectPr/>', body + '<w:sectPr><w:headerReference w:type="default" r:id="rIdH" xmlns:r="' + REL + '"/><w:footerReference w:type="default" r:id="rIdF" xmlns:r="' + REL + '"/></w:sectPr>'));
  const add = (name, xml, type, rel, target) => {
    pkg.addPart(name, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + xml, type);
    if (rel) pkg.addRelationshipTo('word/document.xml', REL + '/' + rel, target);
  };
  add('word/header1.xml', `<w:hdr ${W}><w:p><w:r><w:t>Prepared for Ayesha Khan</w:t></w:r></w:p></w:hdr>`, 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml', 'header', 'header1.xml');
  add('word/footer1.xml', `<w:ftr ${W}><w:p><w:r><w:t xml:space="preserve">Call 042-35761234 </w:t></w:r></w:p></w:ftr>`, 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml', 'footer', 'footer1.xml');
  add('word/comments.xml', `<w:comments ${W}><w:comment w:id="0" w:author="Ayesha Khan" w:initials="AK"><w:p><w:r><w:t>Ask Ayesha Khan first</w:t></w:r></w:p></w:comment></w:comments>`, 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml', 'comments', 'comments.xml');
  add('word/footnotes.xml', `<w:footnotes ${W}><w:footnote w:id="1"><w:p><w:r><w:t>Source: ayesha.khan@example.pk</w:t></w:r></w:p></w:footnote></w:footnotes>`, 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml', 'footnotes', 'footnotes.xml');
  pkg.ensureDefault('png', 'image/png');
  pkg.addPart('word/media/image1.png', Buffer.from([0x89, 0x50, 0x4e, 0x47, 1, 2, 3]));
  pkg.write_('word/_rels/document.xml.rels', pkg.text('word/_rels/document.xml.rels').replace('</Relationships>', '<Relationship Id="rIdPic" Type="' + REL + '/image" Target="media/image1.png"/><Relationship Id="rIdH" Type="' + REL + '/header" Target="header1.xml"/><Relationship Id="rIdF" Type="' + REL + '/footer" Target="footer1.xml"/></Relationships>').replace(/<Relationship Id="rId\d+" Type="[^"]*\/(?:header|footer)" Target="(?:header|footer)1\.xml"\/>/g, ''));
  add('docProps/core.xml', '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>Ayesha Khan</dc:creator><cp:lastModifiedBy>Ayesha Khan</cp:lastModifiedBy><dc:title>Offer for Ayesha Khan</dc:title></cp:coreProperties>', 'application/vnd.openxmlformats-package.core-properties+xml');
  pkg.addPart('docProps/thumbnail.jpeg', Buffer.from('a picture of Ayesha Khan'));
  pkg.write_('_rels/.rels', pkg.text('_rels/.rels').replace('</Relationships>', '<Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="docProps/thumbnail.jpeg"/></Relationships>'));
  // An embedded workbook with the name in a cell.
  pkg.addPart('word/embeddings/Microsoft_Excel_Worksheet.xlsx', Buffer.from(buildXlsx({ sheets: [{ name: 'Sheet1', rows: [['Ayesha Khan', 42]] }] })));
  return pkg.write();
}

const MARKS = [{ text: 'Ayesha Khan' }, { pattern: 'email' }, { pattern: 'phone' }, { pattern: 'card' }];

test('the patterns find what they say and no more', () => {
  const find = (text, mark) => findSpans(text, compileMarks([mark])).map(([a, b]) => text.slice(a, b));
  assert.deepEqual(find('Mail ayesha.khan@example.pk or b@x.co.uk.', { pattern: 'email' }), ['ayesha.khan@example.pk', 'b@x.co.uk']);
  assert.deepEqual(find('Call +92 300 1234567 or (042) 3576-1234, not 2026-10-14 or 1999.', { pattern: 'phone' }), ['+92 300 1234567', '(042) 3576-1234']);
  assert.deepEqual(find('Card 4111 1111 1111 1111; not 4111 1111 1111 1112.', { pattern: 'card' }), ['4111 1111 1111 1111']);
  assert.equal(luhn('4111-1111-1111-1111'), true);
  assert.deepEqual(find('INV-2026-0042 and inv-2026-9', { regex: 'INV-\\d{4}-\\d{4}' }), ['INV-2026-0042']);
  assert.deepEqual(find('Khan, KHAN and khanate', { text: 'khan', wholeWords: true }), ['Khan', 'KHAN']);
  assert.deepEqual(find('Khan and KHAN', { text: 'Khan', matchCase: true }), ['Khan']);
  assert.equal(barsOver('Hi Ayesha!', [[3, 9]]), 'Hi ' + BAR.repeat(6) + '!');
  assert.throws(() => compileMarks([{ regex: '(' }]), /not a pattern that can be searched for/);
});

test('a scan says what would go, part by part, and changes nothing', () => {
  const bytes = secretDocx();
  const pkg = OoxmlPackage.read(bytes);
  const scan = redactPackage(pkg, { marks: MARKS, apply: false });
  const kinds = new Set(scan.parts.map((p) => p.kind));
  for (const k of ['body', 'headers', 'footers', 'comments', 'notes', 'properties', 'embedded']) assert.ok(kinds.has(k), 'found in ' + k);
  assert.ok(scan.total >= 14, 'every place: ' + scan.total);
  assert.ok(scan.words.includes('Ayesha Khan'));
  assert.equal(pkg.modifiedParts().length, 0, 'a scan writes nothing');
});

test('Apply takes the words away everywhere, black bars of the same length in their place, and the package read back holds none of them', () => {
  const pkg = OoxmlPackage.read(secretDocx());
  const black = Buffer.from([0x89, 0x50, 0x4e, 0x47, 9, 9, 9]);
  const result = redactPackage(pkg, { marks: MARKS, pictures: { 'word/media/image1.png': black } });
  assert.ok(result.total >= 14);
  const out = Buffer.from(pkg.write());
  assert.deepEqual(partsHolding(out, ['Ayesha Khan', 'ayesha.khan@example.pk', '300 1234567', '4111 1111 1111 1111', '042-35761234', 'Ayesha', 'a picture of']), []);
  const saved = OoxmlPackage.read(out);
  const body = saved.text('word/document.xml');
  assert.match(body, /<w:t xml:space="preserve">Client: ███████<\/w:t><\/w:r><w:r><w:rPr><w:b\/><\/w:rPr><w:t>████<\/w:t>/, 'a name across two runs, each run keeping its look and its length');
  assert.match(body, /, ██████████████████████, ███████████████\./);
  assert.match(body, /w:author="███████████"/, 'a tracked change\'s author');
  assert.match(body, /<w:delText>Old card ███████████████████<\/w:delText>/, 'deleted words too');
  assert.match(body, /HYPERLINK "mailto:██████████████████████"/, 'a field\'s code');
  assert.match(body, /descr="███████████ at her desk"/, 'alt text');
  assert.match(saved.text('word/header1.xml'), /Prepared for ███████████/);
  assert.match(saved.text('word/comments.xml'), /w:author="███████████" w:initials="AK"/);
  assert.match(saved.text('docProps/core.xml'), /<dc:creator>███████████<\/dc:creator>/);
  assert.deepEqual([...saved.read('word/media/image1.png')], [...black], 'the picture is the blacked-out one');
  assert.equal(saved.has('docProps/thumbnail.jpeg'), false, 'the thumbnail is dropped');
  assert.doesNotMatch(saved.text('_rels/.rels'), /thumbnail/);
  // Still a document Word and this suite open.
  const view = openDocx(out);
  assert.match(view.blocks.map((b) => b.text).join('\n'), /Client: ███████████/);
});

test('a worksheet\'s number a pattern finds becomes a text cell of bars', () => {
  const pkg = OoxmlPackage.read(buildXlsx({ sheets: [{ name: 'S', rows: [['Card', 4111111111111111, 7]] }] }));
  redactPackage(pkg, { marks: [{ pattern: 'card' }] });
  const sheet = pkg.text('xl/worksheets/sheet1.xml');
  assert.match(sheet, /<c r="B1"[^>]* t="inlineStr"><is><t>████████████████<\/t><\/is><\/c>/);
  assert.match(sheet, /<c r="C1"[^>]*><v>7<\/v><\/c>/);
  assert.deepEqual(partsHolding(Buffer.from(pkg.write()), ['4111111111111111']), []);
});
