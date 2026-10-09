// A template of one's own makes a new, untitled document, as Word, Excel and
// PowerPoint make one from it: the template's words and look, labelled as
// the document it makes, and the template itself left alone.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { buildDocx, buildXlsx, OoxmlPackage } from '@rutba/ooxml';
import { buildPptx } from '@rutba/presentation';
import { createDocumentService } from '../apps/desktop/main/documents.js';

const docs = () => createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
const labelOf = (bytes) => { const pkg = OoxmlPackage.read(bytes); return pkg.contentTypeOf(pkg.mainDocument()); };

test('a .dotx makes an untitled document with its words, saved as a document, the template untouched', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-tpl-'));
  const source = path.join(dir, 'letter.docx');
  fs.writeFileSync(source, buildDocx({ styles: true, paragraphs: [{ text: 'Letterhead', style: 'Heading1' }, { text: 'Kept for every letter.' }] }));
  const service = docs();
  const { id } = await service.open({ path: source });
  const template = path.join(dir, 'Letterhead.dotx');
  await service.save({ id, path: template });
  assert.match(labelOf(fs.readFileSync(template)), /wordprocessingml\.template\.main\+xml$/, 'saved as a template');
  const before = fs.readFileSync(template);

  const made = service.new({ kind: 'word', templateFile: template });
  assert.equal(made.kind, 'doc');
  assert.equal(made.path, null, 'untitled');
  assert.equal(made.name, 'Document1.docx');
  assert.match((made.model.blocks || []).map((b) => b.text).join(' '), /Letterhead Kept for every letter\./);
  const out = path.join(dir, 'first-letter.docx');
  await service.save({ id: made.id, path: out });
  assert.match(labelOf(fs.readFileSync(out)), /wordprocessingml\.document\.main\+xml$/, 'labelled as the document it makes');
  assert.ok(before.equals(fs.readFileSync(template)), 'the template left as it was');
});

test('a .xltx and a .potx make a workbook and a deck; anything else is not a template', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-tpl-'));
  const service = docs();
  const book = path.join(dir, 'b.xlsx');
  fs.writeFileSync(book, buildXlsx({ sheets: [{ name: 'Budget', rows: [['Item', 'Cost']] }] }));
  const b = await service.open({ path: book });
  await service.save({ id: b.id, path: path.join(dir, 'Budget.xltx') });
  const sheet = service.new({ kind: 'sheets', templateFile: path.join(dir, 'Budget.xltx') });
  assert.equal(sheet.kind, 'sheet');
  assert.equal(sheet.path, null);
  const deck = path.join(dir, 'd.pptx');
  fs.writeFileSync(deck, buildPptx({ title: 'Pitch', slides: [{ layout: 'title', title: 'Our idea' }] }));
  const d = await service.open({ path: deck });
  await service.save({ id: d.id, path: path.join(dir, 'Pitch.potx') });
  const show = service.new({ kind: 'slides', templateFile: path.join(dir, 'Pitch.potx') });
  assert.equal(show.kind, 'deck');
  assert.equal(show.path, null);
  assert.throws(() => service.new({ kind: 'word', templateFile: book }), /not a template/);
});
