// Insert → Object in a document — the engine's half.
//
// An Office document is embedded as Word embeds one: the file whole in
// word/embeddings, related as a package, and in a paragraph of its own a
// w:object whose VML picture shows its icon and whose o:OLEObject names its
// program and the file. The page draws its icon and knows what it holds;
// the service writes it out for its app to open.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { createDocumentService } from '../apps/desktop/main/documents.js';

const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64');

test('a worksheet is embedded as Word embeds one, and the page draws its icon knowing what it holds', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'See the figures:' }, { text: 'After.' }] }));
  view.setSelection({ block: 0, offset: 5 });
  view.insertObject({ data: buildXlsx({ sheets: [{ name: 'Sheet1', rows: [['Total', 42]] }] }), ext: 'xlsx', name: 'Figures', icon: PNG });
  const doc = view.doc.doc;
  const xml = doc.xml;
  assert.match(xml, /<w:object xmlns:v="urn:schemas-microsoft-com:vml"[^>]*w:dxaOrig="1920" w:dyaOrig="1800">/);
  assert.match(xml, /<v:shape id="_x0000_i\d+" type="#_x0000_t75" alt="Figures" style="width:96pt;height:90pt" o:ole=""><v:imagedata r:id="rId\d+" o:title=""\/><\/v:shape>/);
  assert.match(xml, /<o:OLEObject Type="Embed" ProgID="Excel\.Sheet\.12" ShapeID="_x0000_i\d+" DrawAspect="Icon" ObjectID="_\d+" r:id="rId\d+"><o:FieldCodes>\\s<\/o:FieldCodes><\/o:OLEObject>/);
  assert.ok(doc.pkg.has('word/embeddings/Microsoft_Excel_Worksheet1.xlsx'));
  assert.equal(view.blocks[0].text, 'See the figures:', 'the words stay where they were');
  const images = view.render().blocks.flatMap((b) => b.images || []);
  const obj = images.find((im) => im.kind === 'object');
  assert.ok(obj, JSON.stringify(images.map((im) => im.kind)));
  assert.equal(obj.object.progId, 'Excel.Sheet.12');
  assert.equal(obj.object.part, 'word/embeddings/Microsoft_Excel_Worksheet1.xlsx');
  assert.match(obj.href, /^data:image\/png/);
  // Saved and opened again, still an object.
  const back = openDocx(view.save());
  assert.equal(back.render().blocks.flatMap((b) => b.images || []).find((im) => im.kind === 'object')?.object.progId, 'Excel.Sheet.12');
});

test('the embedded document comes out as a file its app opens, and a new blank one can be embedded', () => {
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const { id } = docs.new({ kind: 'doc' });
  const letter = buildDocx({ styles: true, paragraphs: [{ text: 'Inside.' }] });
  docs.apply({ id, ops: [{ op: 'insertObject', data: new Uint8Array(letter), ext: 'docx', name: 'Letter', icon: new Uint8Array(PNG) }] });
  docs.apply({ id, ops: [{ op: 'insertObject', ext: 'pptx', name: 'Deck', icon: new Uint8Array(PNG) }] });
  const objects = (docs.model({ id }).blocks || []).flatMap((b) => b.images || []).filter((im) => im.kind === 'object');
  assert.deepEqual(objects.map((o) => o.object.progId).sort(), ['PowerPoint.Show.12', 'Word.Document.12']);
  const out = docs.objectFile({ id, part: objects.find((o) => o.object.progId === 'Word.Document.12').object.part });
  assert.equal(openDocx(fs.readFileSync(out.path)).blocks[0].text, 'Inside.');
});
