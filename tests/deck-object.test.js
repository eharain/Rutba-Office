// Insert → Object on a slide — the engine's half.
//
// An Office document is embedded as PowerPoint embeds one: the file whole
// in ppt/embeddings, related as a package, and a graphic frame whose
// p:oleObj names its program and shows its icon. The reader gives it back
// as a picture of its icon that knows what it holds; the service writes the
// embedded document out to a file its app opens.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Deck, buildPptx, renderSlide } from '@rutba/presentation';
import { buildDocx } from '@rutba/ooxml';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { createDocumentService } from '../apps/desktop/main/documents.js';

const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64');
const letter = buildDocx({ styles: true, paragraphs: [{ text: 'Dear reader,' }, { text: 'This letter rides inside a slide.' }] });

test('a document is embedded as PowerPoint embeds one, and read back as its icon that knows what it holds', () => {
  const deck = Deck.open(buildPptx({ title: 'Objects', slides: [{ layout: 'blank' }] }));
  const id = deck.addObject(0, { data: letter, ext: 'docx', name: 'Letter', icon: PNG, x: 100, y: 100, w: 120, h: 110 });
  const xml = deck.pkg.text('ppt/slides/slide1.xml');
  const rel = [...deck.pkg.rels('ppt/slides/slide1.xml')].find((r) => /\/package$/.test(r.Type));
  assert.equal(rel.Target, '../embeddings/Microsoft_Word_Document1.docx');
  assert.match(xml, new RegExp(`<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Object \\d+"/>`));
  assert.match(xml, new RegExp(`<a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj name="Letter" r:id="${rel.Id}" imgW="\\d+" imgH="\\d+" progId="Word.Document.12" showAsIcon="1"><p:embed/><p:pic>`));
  assert.match(deck.pkg.text('[Content_Types].xml'), /<Default Extension="docx" ContentType="application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.document"\/>/);
  const shape = deck.slide(0).shapes.find((s) => String(s.id) === String(id));
  assert.equal(shape.kind, 'picture');
  assert.equal(shape.object.progId, 'Word.Document.12');
  assert.equal(shape.object.source.part, 'ppt/embeddings/Microsoft_Word_Document1.docx');
  assert.ok(shape.source?.part?.startsWith('ppt/media/'), 'drawn as its icon');
  assert.match(renderSlide(deck.slide(0), { width: 1280, resolveImage: () => 'blob:icon' }), /<image\b[^>]*href="blob:icon"/);
  assert.throws(() => deck.addObject(0, { data: letter, ext: 'pdf', icon: PNG, w: 10, h: 10 }), /not a document this embeds/);
});

test('the embedded document comes out as a file its app opens, and a new blank one can be embedded', () => {
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const { id } = docs.new({ kind: 'deck' });
  docs.apply({ id, ops: [{ op: 'addObject', slide: 0, data: new Uint8Array(letter), ext: 'docx', name: 'Letter', icon: new Uint8Array(PNG), x: 10, y: 10, w: 120, h: 110 }] });
  docs.apply({ id, ops: [{ op: 'addObject', slide: 0, ext: 'xlsx', name: 'Worksheet', icon: new Uint8Array(PNG), x: 200, y: 10, w: 120, h: 110 }] });
  const objects = docs.model({ id }).slide.shapes.filter((s) => s.object);
  assert.deepEqual(objects.map((s) => s.object.progId), ['Word.Document.12', 'Excel.Sheet.12']);
  const out = docs.objectFile({ id, part: objects[0].object.part });
  assert.match(out.path, /Microsoft_Word_Document\.docx$/);
  const view = openDocx(fs.readFileSync(out.path));
  assert.equal(view.blocks[1].text, 'This letter rides inside a slide.');
  assert.throws(() => docs.objectFile({ id, part: 'ppt/slides/slide1.xml' }), /not an embedded document/);
});
