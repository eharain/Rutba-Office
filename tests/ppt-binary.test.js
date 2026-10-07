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
 *
 * PowerPoint 2007 and later keep the deck as they made it beside the older
 * records, and a .ppt they saved is rebuilt from that; the older records'
 * own reading, which a .ppt from PowerPoint 97 to 2003 gets, is tested on the
 * model with what they kept taken away.
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
/** The model as an older PowerPoint would have written it: the older records alone. */
const older = (model) => ({ ...model, kept: null });

test('a PowerPoint 97-2003 presentation reads as its slides: size, backgrounds, titles and notes', () => {
  const model = readPpt(fixture('binary', 'showcase.ppt'));
  assert.deepEqual(model.size, { width: 1280, height: 720 });
  assert.equal(model.slides.length, 8);
  const [first, second] = model.slides;
  assert.equal(first.background.gradient.stops[0].color, '#1F4E79', 'the title slide\'s gradient, from its first colour');
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

test('written as a .pptx from the older records: eight slides, each shape where it stood, the notes kept', () => {
  const model = readPpt(fixture('binary', 'showcase.ppt'));
  const bytes = pptModelToDeck(older(model));
  const pkg = OoxmlPackage.read(bytes);
  const slides = pkg.partNames().filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n));
  assert.equal(slides.length, 8);
  const slide3 = pkg.text('ppt/slides/slide3.xml');
  assert.match(slide3, /<a:prstGeom prst="chevron">/);
  assert.match(slide3, /<a:custGeom>/, 'a freeform\'s own outline');
  assert.match(pkg.text('ppt/slides/slide1.xml'), /<p:bg><p:bgPr><a:gradFill/);
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

  const xml = OoxmlPackage.read(pptModelToDeck(older(model))).text('ppt/slides/slide4.xml');
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

  const tbl = /<a:tbl>[\s\S]*<\/a:tbl>/.exec(OoxmlPackage.read(pptModelToDeck(older(model))).text('ppt/slides/slide1.xml'))?.[0] ?? '';
  assert.match(tbl, /<a:tc gridSpan="2">[\s\S]*?<a:t[^>]*>Across two<\/a:t>/);
  assert.match(tbl, /<a:tc hMerge="1">/);
  assert.match(tbl, /<a:tc rowSpan="3">[\s\S]*?<a:tcPr anchor="ctr">/);
  assert.equal((tbl.match(/<a:tc vMerge="1">/g) || []).length, 2);
});

test('gradients, transparency, shadows and where words sit come as the original has them', () => {
  const model = readPpt(fixture('binary', 'showcase.ppt'));
  // The title slide's background: Office Art's shaded fill read as the gradient the .pptx had (blue to red at 135°).
  assert.deepEqual(model.slides[0].background, { gradient: { stops: [{ pos: 0, color: '#1F4E79' }, { pos: 1, color: '#C00000' }], angle: 135 } });
  // A centred title takes the title style's font where its own style says none.
  assert.equal(model.slides[0].shapes[0].paragraphs[0].runs[0].font, 'Aptos Display');
  const shapes = Object.fromEntries(model.slides[2].shapes.map((s) => [textOf(s), s]));
  assert.deepEqual(shapes.Rounded.fill.gradient, { stops: [{ pos: 0, color: '#ED7D31' }, { pos: 1, color: '#FFFFFF' }], angle: 90 }, 'the second colour white where it does not say');
  assert.equal(Math.round(shapes.Smile.fill.alpha * 100), 60);
  assert.ok(Math.abs(shapes.Triangle.shadow.dist * 12700 - 37357) < 2, 'as far off as the original\'s shadow');
  assert.equal(shapes.Triangle.shadow.dir, 45);
  assert.equal(shapes.Rectangle.shadow, undefined);

  const pkg = OoxmlPackage.read(pptModelToDeck(older(model)));
  assert.match(pkg.text('ppt/slides/slide1.xml'), /<p:bg><p:bgPr><a:gradFill[^>]*><a:gsLst><a:gs pos="0"><a:srgbClr val="1F4E79"\/><\/a:gs><a:gs pos="100000"><a:srgbClr val="C00000"\/><\/a:gs><\/a:gsLst><a:lin ang="8100000"/);
  const slide3 = pkg.text('ppt/slides/slide3.xml');
  assert.equal((slide3.match(/<a:outerShdw /g) || []).length, 4);
  assert.match(slide3, /<a:srgbClr val="806000"><a:alpha val="60000"\/><\/a:srgbClr>/);
  assert.match(pkg.text('ppt/slides/slide4.xml'), /<a:bodyPr[^>]*anchor="ctr"/, 'the title in the middle of its box, as the placeholder says');
});

test('a slide hidden from the show stays hidden, and each comes on with its transition', () => {
  const model = readPpt(fixture('binary', 'showcase.ppt'));
  assert.deepEqual(model.slides.map((s) => s.hidden), [false, false, false, false, false, false, true, false]);
  assert.deepEqual(model.slides.map((s) => s.transition?.type ?? null), [null, null, 'cut', null, null, 'cut', null, null], 'the two cuts the original has, and no more');
  const deck = Deck.open(Buffer.from(pptModelToDeck(model)));
  assert.equal(deck.isSlideHidden(6), true);
  assert.equal(deck.transition(2)?.type, 'cut');
  assert.equal(deck.transition(0), null);
});

test('a slide\'s effects come as a later PowerPoint kept them, the PowerPoint 97 builds beside them read too', () => {
  const model = readPpt(fixture('binary', 'showcase.ppt'));
  const slide = model.slides[6];
  assert.deepEqual(slide.timing.map(({ kind, effect, direction, trigger }) => [kind, effect, direction, trigger]), [
    ['entr', 'fly', 'bottom', 'onClick'], ['entr', 'fade', null, 'onClick'], ['entr', 'zoom', null, 'withPrevious'],
  ]);
  // PowerPoint 97 had no fade and no "with previous": what it kept is the nearest.
  assert.deepEqual(slide.shapes.filter((s) => s.animation).map((s) => [s.animation.effect, s.animation.direction, s.animation.trigger]), [
    ['fly', 'bottom', 'onClick'], ['appear', null, 'onClick'], ['zoom', null, 'afterPrevious'],
  ]);
  const deck = Deck.open(Buffer.from(pptModelToDeck(model)));
  assert.deepEqual(deck.animations(6).map((a) => [a.effect, a.direction, a.trigger]), [['fly', 'bottom', 'onClick'], ['fade', null, 'onClick'], ['zoom', null, 'withPrevious']]);
});

test('a .ppt PowerPoint 2007 or later saved is rebuilt as they kept it: its theme, master and layouts, each slide on its layout', () => {
  const model = readPpt(fixture('binary', 'showcase.ppt'));
  assert.equal(model.kept.masters.filter((m) => m.layout).length, 11, 'eleven layouts, each kept whole');
  const pkg = OoxmlPackage.read(pptModelToDeck(model));
  const original = OoxmlPackage.read(Buffer.from(fixture('rich', 'showcase.pptx')));
  assert.equal(pkg.partNames().filter((n) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(n)).length, 11);
  const accents = (p) => /<a:accent1>[\s\S]*?<\/a:accent6>/.exec(p.text('ppt/theme/theme1.xml'))[0];
  assert.equal(accents(pkg), accents(original), 'the deck\'s own theme');
  const layoutOf = (p, n) => /<p:cSld name="([^"]*)"/.exec(p.text(p.partNames().find((x) => x === 'ppt/slideLayouts/' + /slideLayout\d+\.xml/.exec(p.text(`ppt/slides/_rels/slide${n}.xml.rels`))[0])))[1];
  for (let n = 1; n <= 8; n++) assert.equal(layoutOf(pkg, n), layoutOf(original, n), `slide ${n} on its layout`);
  // Placeholders of their layouts, in only their own looks; the master's styles as kept.
  const last = pkg.text('ppt/slides/slide8.xml');
  assert.match(last, /<p:ph type="title"\/>[\s\S]*<a:t>Thank you<\/a:t>/);
  assert.match(last, /<p:ph idx="1"\/>[\s\S]*?<a:p><a:r><a:rPr lang="en-GB"\/><a:t>Eight slides/);
  assert.match(last, /<a:rPr lang="ja-JP" altLang="en-US"\/><a:t>日本語 <\/a:t>[\s\S]*?<a:rPr lang="ar-SA"\/><a:t>العربية/, 'each run in its language, which orders the Arabic among the rest');
  assert.match(pkg.text('ppt/slideMasters/slideMaster1.xml'), /<p:txStyles><p:titleStyle>/);
  // Drawings as kept, their masked words put back: the heart its preset and link, the connector a connector, the slide number its own.
  const shapes = pkg.text('ppt/slides/slide3.xml');
  assert.match(shapes, /name="Shape Heart">[\s\S]*?<a:hlinkClick [^>]*r:id="(rId\d+)"[\s\S]*?prst="heart"[\s\S]*?<a:t>Heart<\/a:t>/);
  assert.match(pkg.text('ppt/slides/_rels/slide3.xml.rels'), /Target="https:\/\/office.rutba.io\/" TargetMode="External"/);
  assert.match(shapes, /<p:cxnSp\b[\s\S]*?prst="bentConnector3"/);
  assert.match(shapes, /type="slidenum">[\s\S]*?<a:t>3<\/a:t>/);
  assert.match(pkg.text('ppt/slides/slide4.xml'), /<p:graphicFrame\b[\s\S]*<a:tableStyleId>[\s\S]*<a:t>Region<\/a:t>/, 'the table as kept, in its style');
  assert.ok(pkg.partNames().includes('ppt/tableStyles.xml'));
});

test('a drawing an older PowerPoint changed after it was kept comes from the older records it changed', () => {
  const model = readPpt(fixture('binary', 'showcase.ppt'));
  const items = model.slides[2].kept.items;
  const rectangle = items.find((it) => textOf(it.shapes[0] || {}) === 'Rectangle').shapes[0];
  rectangle.words = 'Square';
  rectangle.paragraphs = [{ runs: [{ text: 'Square' }] }];
  const oval = items.find((it) => textOf(it.shapes[0] || {}) === 'Oval').shapes[0];
  oval.x += 40;
  const shapes = OoxmlPackage.read(pptModelToDeck(model)).text('ppt/slides/slide3.xml');
  const rect = /<p:sp\b[^>]*><p:nvSpPr><p:cNvPr id="\d+" name="Shape Rectangle"[\s\S]*?<\/p:sp>/.exec(shapes)[0];
  assert.match(rect, /<a:t[^>]*>Square<\/a:t>/, 'its words changed: drawn from the older records');
  assert.doesNotMatch(rect, /<p:style>/, 'not the kept rectangle');
  assert.doesNotMatch(shapes, /name="Shape Oval">[\s\S]*?<a:off x="3302000"/, 'moved: not where it was kept');
  assert.match(shapes, /name="Shape Star"/, 'the rest as kept');
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
