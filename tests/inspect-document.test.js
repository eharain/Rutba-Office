// Inspect Document: what a document carries that its author may not want to
// hand on, comments, tracked changes, personal details, properties, custom
// XML, headers and footers, hidden words, invisible objects, each found,
// and each removed on request as Word's Document Inspector removes it, the
// file still one Word and the suite open.
import test from 'node:test';
import assert from 'node:assert/strict';
import { buildDocx, OoxmlPackage } from '@rutba/ooxml';
import { inspectPackage, removeFromPackage, acceptAllXml } from '@rutba/ooxml/inspect';
import { openDocx } from '@rutba/doc-view/backends/ooxml';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function busyDocx() {
  const pkg = OoxmlPackage.read(buildDocx({ styles: true, paragraphs: [{ text: 'Plain' }] }));
  const body = '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t xml:space="preserve">Seen </w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>'
    + '<w:ins w:id="1" w:author="Bilal" w:date="2026-10-01T10:00:00Z"><w:r><w:t>added</w:t></w:r></w:ins><w:del w:id="2" w:author="Bilal" w:date="2026-10-01T10:00:00Z"><w:r><w:delText>gone</w:delText></w:r></w:del>'
    + '<w:r><w:rPr><w:vanish/></w:rPr><w:t>secret note</w:t></w:r><w:r><w:t xml:space="preserve"> end</w:t></w:r></w:p>'
    + '<w:p><w:r><w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1" relativeHeight="251658240" distT="0" distB="0" distL="0" distR="0" simplePos="0"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="9525" cy="9525"/><wp:wrapNone/><wp:docPr id="9" name="Hidden box" hidden="1"/></wp:anchor></w:drawing></w:r><w:r><w:t>visible</w:t></w:r></w:p>';
  pkg.write_('word/document.xml', pkg.text('word/document.xml').replace('<w:sectPr/>', body + '<w:sectPr><w:headerReference w:type="default" r:id="rIdH" xmlns:r="' + REL + '"/></w:sectPr>'));
  pkg.addPart('word/header1.xml', `<?xml version="1.0"?><w:hdr ${W}><w:p><w:r><w:t>Internal</w:t></w:r></w:p></w:hdr>`, 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml');
  pkg.addPart('word/comments.xml', `<?xml version="1.0"?><w:comments ${W}><w:comment w:id="0" w:author="Sana" w:initials="S"><w:p><w:r><w:t>Check this figure</w:t></w:r></w:p></w:comment></w:comments>`, 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml');
  pkg.addPart('word/people.xml', `<?xml version="1.0"?><w15:people xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:person w15:author="Sana"><w15:presenceInfo w15:providerId="None" w15:userId="Sana"/></w15:person></w15:people>`, 'application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml');
  pkg.addPart('word/settings.xml', `<?xml version="1.0"?><w:settings ${W}><w:zoom w:percent="100"/></w:settings>`, 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');
  let rels = pkg.text('word/_rels/document.xml.rels');
  rels = rels.replace('</Relationships>', `<Relationship Id="rIdH" Type="${REL}/header" Target="header1.xml"/><Relationship Id="rIdC" Type="${REL}/comments" Target="comments.xml"/><Relationship Id="rIdP" Type="http://schemas.microsoft.com/office/2011/relationships/people" Target="people.xml"/><Relationship Id="rIdS" Type="${REL}/settings" Target="settings.xml"/><Relationship Id="rIdX" Type="${REL}/customXml" Target="../customXml/item1.xml"/></Relationships>`);
  pkg.write_('word/_rels/document.xml.rels', rels);
  pkg.addPart('customXml/item1.xml', '<?xml version="1.0"?><client><name>Ayesha</name></client>', 'application/xml');
  pkg.addPart('docProps/core.xml', '<?xml version="1.0"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Quarterly plan</dc:title><dc:creator>Bilal Ahmed</dc:creator><cp:lastModifiedBy>Sana Iqbal</cp:lastModifiedBy><cp:keywords>budget</cp:keywords></cp:coreProperties>', 'application/vnd.openxmlformats-package.core-properties+xml');
  pkg.addPart('docProps/app.xml', '<?xml version="1.0"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Company>Tech Style</Company><Pages>1</Pages></Properties>', 'application/vnd.openxmlformats-officedocument.extended-properties+xml');
  return pkg.write();
}

const byId = (items) => Object.fromEntries(items.map((k) => [k.id, k]));

test('Inspect finds each kind of content Word\'s Document Inspector looks for', () => {
  const found = byId(inspectPackage(OoxmlPackage.read(busyDocx()), 'document'));
  assert.equal(found.comments.found, 1);
  assert.deepEqual(found.comments.items, ['Check this figure']);
  assert.equal(found.revisions.found, 2);
  assert.ok(found.personal.found >= 4, 'author, last saved by, company and reviewers: ' + JSON.stringify(found.personal.items));
  assert.ok(found.personal.items.includes('Bilal Ahmed') && found.personal.items.includes('Sana'));
  assert.deepEqual(found.properties.items.sort(), ['Quarterly plan', 'budget']);
  assert.equal(found.customXml.found, 1);
  assert.equal(found.headers.found, 1);
  assert.deepEqual(found.hidden.items, ['secret note']);
  assert.equal(found.invisible.found, 1);
});

test('Remove All takes each kind out, and the document still opens', () => {
  const pkg = OoxmlPackage.read(busyDocx());
  const removed = removeFromPackage(pkg, 'document', ['comments', 'revisions', 'personal', 'properties', 'customXml', 'headers', 'hidden', 'invisible']);
  assert.equal(removed.length, 8);
  const after = byId(inspectPackage(pkg, 'document'));
  for (const id of removed) assert.equal(after[id].found, 0, id + ' is gone');
  const body = pkg.text('word/document.xml');
  assert.doesNotMatch(body, /commentRangeStart|commentReference|<w:ins\b|<w:del\b|secret note|Hidden box|headerReference/);
  assert.match(body, /<w:t>added<\/w:t>/, 'an insertion is accepted');
  assert.doesNotMatch(body, /gone/, 'a deletion is accepted');
  assert.equal(pkg.has('word/comments.xml'), false);
  assert.equal(pkg.has('word/header1.xml'), false);
  assert.equal(pkg.has('word/people.xml'), false);
  assert.equal(pkg.has('customXml/item1.xml'), false);
  assert.doesNotMatch(pkg.text('word/_rels/document.xml.rels'), /comments\.xml|header1|people\.xml|customXml/);
  assert.match(pkg.text('word/settings.xml'), /<w:zoom w:percent="100"\/><w:removePersonalInformation\/><w:removeDateAndTime\/>/);
  assert.doesNotMatch(pkg.text('docProps/core.xml'), /Bilal|Sana|Quarterly|budget/);
  assert.doesNotMatch(pkg.text('docProps/app.xml'), /Tech Style/);
  const view = openDocx(pkg.write());
  assert.equal(view.blocks.map((b) => b.text).join('|'), 'Plain|Seen added end|visible');
});

test('accepting every change by hand leaves the inserted words and drops the deleted ones and the records of before', () => {
  const xml = '<w:p><w:pPr><w:pPrChange w:id="5" w:author="A"><w:pPr/></w:pPrChange></w:pPr><w:moveFrom w:id="6"><w:r><w:t>away</w:t></w:r></w:moveFrom><w:moveTo w:id="7"><w:r><w:rPr><w:b/><w:rPrChange w:id="8" w:author="A"><w:rPr/></w:rPrChange></w:rPr><w:t>here</w:t></w:r></w:moveTo></w:p>';
  assert.equal(acceptAllXml(xml), '<w:p><w:pPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>here</w:t></w:r></w:p>');
});
