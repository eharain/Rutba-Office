/**
 * PowerPoint 97-2003 presentations, opened in full.
 *
 * showcase.ppt is tests/fixtures/rich/showcase.pptx as PowerPoint itself
 * saved it in its 97-2003 format (tools/make-binary-fixtures.ps1): eight
 * slides of titles, bullets at three levels, sixteen preset shapes (some
 * PowerPoint can only save as freeforms), a table, two charts, a picture,
 * WordArt, notes. The reader is judged by whether the slides come back
 * with the words, shapes, pictures and notes the original had. tables.ppt
 * is a table with merged cells PowerPoint made and saved the same way.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { readPpt } from '@rutba/office-formats/msppt';
import { OoxmlPackage } from '@rutba/ooxml/package';
import { Deck } from '@rutba/presentation';
import { pptModelToDeck } from '../apps/desktop/main/legacy-deck.js';
import { createDocumentService } from '../apps/desktop/main/documents.js';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const fixture = (...p) => new Uint8Array(fs.readFileSync(path.join(HERE, 'fixtures', ...p)));
const textOf = (shape) => (shape.paragraphs || []).map((p) => p.runs.map((r) => r.text).join('')).join('\n');

test('a PowerPoint 97-2003 presentation reads as its slides: size, backgrounds, titles and notes', () => {
  const model = readPpt(fixture('binary', 'showcase.ppt'));
  assert.deepEqual(model.size, { width: 1280, height: 720 });
  assert.equal(model.slides.length, 8);
  const [first, second] = model.slides;
  assert.equal(first.background, '#1F4E79');
  const title = first.shapes.find((s) => textOf(s) === 'The Rutba Office showcase');
  assert.ok(title, 'the title slide\'s title');
  assert.equal(title.paragraphs[0].runs[0].size, 60);
  assert.equal(title.paragraphs[0].runs[0].color, '#FFFFFF');
  assert.match(first.notes, /^Speaker notes for the title slide/);
  assert.equal(second.notes, 'Notes on slide two: the picture is generated, not fetched.');
  // The slide number field shows the slide's number.
  assert.ok(second.shapes.some((s) => textOf(s) === '2'));
});

test('bullets keep their levels, and the master\'s sizes and fonts', () => {
  const model = readPpt(fixture('binary', 'showcase.ppt'));
  const body = model.slides[1].shapes.find((s) => textOf(s).startsWith('First point'));
  assert.deepEqual(body.paragraphs.map((p) => p.level), [0, 0, 1, 2, 0]);
  assert.ok(body.paragraphs[0].bullet && body.paragraphs[0].bullet.type === 'char', 'a bullet');
  assert.equal(body.paragraphs[0].runs[0].size, 28);
  assert.equal(body.paragraphs[0].runs[0].font, 'Aptos');
});

test('preset shapes, and the ones PowerPoint saved as freeforms, come back with their fills and words', () => {
  const model = readPpt(fixture('binary', 'showcase.ppt'));
  const shapes = model.slides[2].shapes.filter((s) => s.type === 'shape');
  const byText = Object.fromEntries(shapes.map((s) => [textOf(s), s]));
  for (const word of ['Rectangle', 'Rounded', 'Oval', 'Triangle', 'Star', 'Arrow', 'Chevron', 'Hexagon', 'Heart', 'Smile', 'Process', 'Callout', 'Cloud', 'Up arrow', 'Diamond', 'Parallelogram']) {
    assert.ok(byText[word], `the ${word} shape`);
  }
  assert.equal(byText.Rectangle.preset, 'rect');
  assert.equal(byText.Rectangle.fill, '#C00000');
  assert.equal(byText.Chevron.preset, 'chevron');
  assert.ok(byText.Star.path && byText.Heart.path && byText.Cloud.path, 'freeforms drawn from their own points');
  assert.ok(byText.Star.path.commands.length > 5);
  // The slide 6 text box: white, with a two-point outline, as in the original.
  const box = model.slides[5].shapes.find((s) => textOf(s).startsWith('A text box'));
  assert.equal(box.fill, '#FFFFFF');
  assert.deepEqual(box.line, { color: '#000000', width: 2 });
  assert.ok(box.paragraphs[1].runs[0].bold, 'the bold paragraph');
});

test('pictures and charts come with their bytes', () => {
  const model = readPpt(fixture('binary', 'showcase.ppt'));
  const pictures = model.slides.flatMap((s) => s.shapes.filter((x) => x.type === 'picture'));
  assert.ok(pictures.length >= 3, `${pictures.length} pictures`);
  for (const p of pictures) {
    const img = model.images[p.image];
    assert.ok(img && img.bytes.length > 100, 'bytes');
    assert.ok(['image/png', 'image/jpeg', 'image/x-emf', 'image/x-wmf', 'image/bmp'].includes(img.contentType), img.contentType);
  }
});

test('written as a .pptx: eight slides, each shape where it stood, the notes kept', () => {
  const model = readPpt(fixture('binary', 'showcase.ppt'));
  const bytes = pptModelToDeck(model);
  const pkg = OoxmlPackage.read(bytes);
  const slides = pkg.partNames().filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n));
  assert.equal(slides.length, 8);
  const slide3 = pkg.text('ppt/slides/slide3.xml');
  assert.match(slide3, /<a:prstGeom prst="chevron">/);
  assert.match(slide3, /<a:custGeom>/, 'a freeform\'s own outline');
  assert.match(pkg.text('ppt/slides/slide1.xml'), /<p:bg><p:bgPr><a:solidFill><a:srgbClr val="1F4E79"\/>/);
  const deck = Deck.open(Buffer.from(bytes));
  assert.ok(deck, 'the deck opens');
  assert.ok(pkg.partNames().some((n) => /^ppt\/notesSlides\//.test(n)), 'notes pages');
});

test('a table comes back a table: its columns and rows, each cell its words, fill and borders', () => {
  const model = readPpt(fixture('binary', 'showcase.ppt'));
  const table = model.slides[3].shapes.find((s) => s.type === 'table');
  assert.ok(table, 'not twenty boxes and eleven lines');
  assert.deepEqual([table.rows.length, table.columns.length], [5, 4]);
  assert.deepEqual(table.cells.map((row) => row.map((c) => textOf(c))), [
    ['Region', 'Q1', 'Q2', 'Total'], ['North', '1,200', '1,350', '2,550'], ['South', '980', '1,120', '2,100'],
    ['East', '1,430', '1,510', '2,940'], ['All', '3,610', '3,980', '7,590'],
  ]);
  const header = table.cells[0][0];
  assert.equal(header.fill, '#156082');
  assert.equal(header.paragraphs[0].runs[0].bold, true);
  assert.equal(header.borders.bottom.width, 3, 'the line under the header, thicker');
  assert.equal(table.cells[4][3].fill, '#FFE699', 'the one cell picked out');

  const xml = OoxmlPackage.read(pptModelToDeck(model)).text('ppt/slides/slide4.xml');
  const tbl = /<a:tbl>[\s\S]*<\/a:tbl>/.exec(xml)?.[0] ?? '';
  assert.equal((tbl.match(/<a:tr /g) || []).length, 5);
  assert.equal((tbl.match(/<a:gridCol /g) || []).length, 4);
  assert.match(tbl, /<a:tblPr\/>/, 'its own look, no table style over it');
  assert.match(tbl, /<a:lnB w="38100"><a:solidFill><a:srgbClr val="FFFFFF"\/><\/a:solidFill><\/a:lnB><a:solidFill><a:srgbClr val="156082"\/>/);
});

test('a table\'s merged cells, a cell set in the middle and one with no fill come too', () => {
  const model = readPpt(fixture('binary', 'tables.ppt'));
  const table = model.slides[0].shapes.find((s) => s.type === 'table');
  assert.deepEqual([table.rows.length, table.columns.length], [4, 3]);
  const [[across, coveredRight], [, , down], [south, , below]] = table.cells;
  assert.equal(across.colSpan, 2);
  assert.equal(coveredRight.hMerge, true);
  assert.equal(down.rowSpan, 3);
  assert.equal(down.anchor, 'middle');
  assert.equal(below.vMerge, true);
  assert.equal(south.fill, 'none');
  assert.equal(across.borders.right.color, '#FFFFFF', 'a merged cell\'s far edge from the last cell it covers');

  const tbl = /<a:tbl>[\s\S]*<\/a:tbl>/.exec(OoxmlPackage.read(pptModelToDeck(model)).text('ppt/slides/slide1.xml'))?.[0] ?? '';
  assert.match(tbl, /<a:tc gridSpan="2">[\s\S]*?<a:t[^>]*>Across two<\/a:t>/);
  assert.match(tbl, /<a:tc hMerge="1">/);
  assert.match(tbl, /<a:tc rowSpan="3">[\s\S]*?<a:tcPr anchor="ctr">/);
  assert.equal((tbl.match(/<a:tc vMerge="1">/g) || []).length, 2);
});

test('the suite opens a .ppt as a presentation, read in full, and says it saves a .pptx', () => {
  const doc = createDocumentService({ holdBlob: () => ({ url: 'blob://held' }) });
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-ppt-'));
  const p = path.join(dir, 'showcase.ppt');
  fs.writeFileSync(p, fixture('binary', 'showcase.ppt'));
  const s = doc.open({ path: p });
  doc.close({ id: s.id });
  assert.equal(s.kind, 'deck');
  assert.equal(s.converted.from, 'ppt');
  assert.ok(!s.converted.partial, 'not the old "could not be converted" slide');
});
