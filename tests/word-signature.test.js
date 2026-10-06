// Insert → Signature Line in Documents: written as Word writes one — a VML
// picture shape carrying o:signatureline with the suggested signer's name,
// title and e-mail — read back, drawn, and undone; and a VML picture of an
// older Word's is drawn on the page now too.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { OoxmlPackage } from '@rutba/ooxml';
import { buildDocx } from '@rutba/ooxml/build';
import { createDocumentService } from '../apps/desktop/main/documents.js';

const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC', 'base64');
const images = (model) => model.blocks.flatMap((b) => b.images || []);

test('a signature line is written as Word writes one and read back with its signer', () => {
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const { id } = docs.new({ kind: 'word', template: 'doc' });
  docs.apply({ id, ops: [{ op: 'insertSignatureLine', png: PNG, widthPx: 256, heightPx: 128, signer: 'Jo Bloggs', title: 'Manager & Co', email: 'jo@example.com', instructions: 'Sign before Friday', showDate: true }] });
  const [sig] = images(docs.model({ id })).filter((im) => im.kind === 'signatureLine');
  assert.ok(sig, 'drawn on the page');
  assert.deepEqual(sig.signatureLine, { signer: 'Jo Bloggs', title: 'Manager & Co', email: 'jo@example.com' });
  assert.equal(sig.vml, true);
  assert.equal(Math.round(sig.widthPx), 256);
  assert.ok(sig.href, 'with its picture');

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-sig-'));
  const file = path.join(dir, 'signed.docx');
  docs.save({ id, path: file });
  const xml = OoxmlPackage.read(fs.readFileSync(file)).text('word/document.xml');
  assert.match(xml, /<w:pict xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"/);
  assert.match(xml, /<v:shapetype id="_x0000_t75"/);
  assert.match(xml, /<v:shape id="_x0000_i\d+" type="#_x0000_t75"[^>]*style="width:192pt;height:96pt"><v:imagedata r:id="rId\d+" o:title=""\/>/);
  assert.match(xml, /<o:signatureline v:ext="edit" id="\{[0-9A-F-]{36}\}" provid="\{00000000-0000-0000-0000-000000000000\}" o:suggestedsigner="Jo Bloggs" o:suggestedsigner2="Manager &amp; Co" o:suggestedsigneremail="jo@example.com" signinginstructionsset="t" o:signinginstructions="Sign before Friday" issignatureline="t" showsigndate="t"\/>/);
});

test('a second signature line does not declare the picture shape type twice, and one undo takes a line out', () => {
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const { id } = docs.new({ kind: 'word', template: 'doc' });
  docs.apply({ id, ops: [{ op: 'insertSignatureLine', png: PNG, signer: 'One' }] });
  docs.apply({ id, ops: [{ op: 'insertSignatureLine', png: PNG, signer: 'Two' }] });
  assert.equal(images(docs.model({ id })).filter((im) => im.kind === 'signatureLine').length, 2);
  docs.undo({ id });
  assert.deepEqual(images(docs.model({ id })).filter((im) => im.kind === 'signatureLine').map((im) => im.signatureLine.signer), ['One']);
});

test('an older Word\'s VML picture is drawn on the page, after the paragraph\'s DrawingML pictures', async () => {
  const pkg = OoxmlPackage.read(buildDocx({ paragraphs: [{ text: 'Logo:' }] }));
  pkg.addPart('word/media/old.png', PNG, 'image/png');
  const relsPart = 'word/_rels/document.xml.rels';
  const image = '<Relationship Id="rId90" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/old.png"/>';
  if (pkg.has(relsPart)) pkg.write_(relsPart, Buffer.from(pkg.text(relsPart).replace('</Relationships>', () => image + '</Relationships>'), 'utf8'));
  else pkg.addPart(relsPart, Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${image}</Relationships>`, 'utf8'), 'application/vnd.openxmlformats-package.relationships+xml');
  const pict = '<w:r><w:pict xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:72pt;height:36pt"><v:imagedata r:id="rId90" o:title=""/></v:shape></w:pict></w:r>';
  pkg.write_('word/document.xml', Buffer.from(pkg.text('word/document.xml').replace('</w:p>', () => pict + '</w:p>'), 'utf8'));
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-vml-'));
  const file = path.join(dir, 'old.docx');
  fs.writeFileSync(file, pkg.write());
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const { id } = await docs.open({ path: file });
  const [pic] = images(docs.model({ id }));
  assert.ok(pic?.href, 'the picture is drawn');
  assert.equal(pic.vml, true);
  assert.deepEqual([Math.round(pic.widthPx), Math.round(pic.heightPx)], [96, 48]);
});
