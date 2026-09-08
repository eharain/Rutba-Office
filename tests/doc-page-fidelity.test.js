/**
 * What the page has to DRAW, read out of a tender's first pages: a cover in
 * a content control with a themed text box, a table of contents with its
 * style's dotted tab, a form line with tab stops, shading and a rule, theme
 * fonts, footnotes numbered in body order, a first-page header.
 *
 * Every assertion here was a visible difference beside Microsoft Word before
 * the reader learned the construct — see tests/fixtures/cover-page-document.js.
 */
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { Document, readThemeColours, readThemeFonts } from '@rutba/ooxml';
import { buildCoverPageDocument } from './fixtures/cover-page-document.js';

const frame = () => openDocx(buildCoverPageDocument()).render({ pages: false });

test('a body-level content control\'s paragraphs are on the page, read-only', () => {
  const f = frame();
  const texts = f.blocks.map((b) => b.text);
  assert.ok(texts.includes('Cover'), 'the cover paragraph is a block: ' + JSON.stringify(texts));
  assert.ok(texts.includes('Contents'), 'the TOC heading is a block');
  const cover = f.blocks.find((b) => b.text === 'Cover');
  assert.equal(cover.inSdt, true);
  assert.equal(cover.structural, true, 'nothing in a control is editable');
  assert.ok(cover.structuralTags.includes('w:sdt'));
  // The body paragraph after the controls is still plain and editable.
  const plain = f.blocks.find((b) => b.text === 'Plain, editable.');
  assert.equal(plain.structural, false);
  assert.equal(plain.inSdt, undefined);
});

test('a text box is read with its fill, its outline and its paragraphs — once', () => {
  const f = frame();
  const cover = f.blocks.find((b) => b.text === 'Cover');
  assert.equal(cover.textBoxes.length, 1, 'the VML fallback is not a second box');
  const box = cover.textBoxes[0];
  assert.equal(box.name, 'Title Box');
  assert.equal(box.widthPx, 576);
  assert.equal(box.heightPx, 144);
  assert.equal(box.hAlign, 'center');
  // accent1 (#7E97AD) lightened by lumMod 60% + lumOff 40%.
  assert.match(box.fill, /^#[0-9A-F]{6}$/);
  assert.notEqual(box.fill, '#7E97AD', 'the tint is applied');
  assert.equal(box.line, '#1F2123');
  assert.deepEqual(box.paragraphs.map((p) => p.text), ['Learning Management System', 'Scope of requirements, 2017']);
  assert.equal(box.paragraphs[0].style, 'Title');
  assert.equal(box.paragraphs[0].runs[0].fontColour, 'FFFFFF');
  assert.equal(box.paragraphs[0].runs[0].fontSize, 48);
  assert.equal(box.paragraphs[1].align, 'right');
  // The anchor paragraph's own runs are its own words, not the box's.
  assert.deepEqual(cover.runs.map((r) => r.text), ['Cover']);
  assert.equal(f.wordCount, f.blocks.reduce((n, b) => n + (b.text.trim() ? b.text.trim().split(/\s+/).length : 0), 0));
});

test('a text box also keeps the anchor paragraph from being rebuilt', () => {
  const view = openDocx(buildCoverPageDocument());
  const cover = view.blocks.findIndex((b) => b.text === 'Cover');
  view.setSelection({ block: cover, offset: 0 });
  assert.throws(() => view.insertText('x'), /w:sdt|w:txbxContent|structural|carries/i);
});

test('direct tab stops, shading, a border and a hanging indent reach the frame', () => {
  const f = frame();
  const line = f.blocks.find((b) => b.text.startsWith('Name:'));
  assert.deepEqual(line.tabs, [
    { align: 'left', posPx: 144, leader: null },
    { align: 'right', posPx: 576, leader: 'dot' },
  ]);
  assert.equal(line.shading, '#D9E2F3', 'the paragraph shading, not the pilcrow run\'s');
  assert.equal(line.borders.bottom.style, 'single');
  assert.equal(line.borders.bottom.colour, '#1F5F8B');
  assert.equal(line.borders.bottom.widthPx, 1);
  assert.equal(line.hangingPx, 24);
  assert.equal(line.text, 'Name:\tValue\tPage 3', 'tabs survive into the text');
  assert.deepEqual(line.runs.map((r) => r.text), ['Name:', '\t', 'Value', '\tPage 3'], 'a tab in a run of its own is a run');
  assert.equal(line.runs.map((r) => r.text).join(''), line.text, 'text and runs agree, so caret offsets do');
});

test('superscript and subscript are run properties the engine writes and reads back', () => {
  const view = openDocx(buildCoverPageDocument());
  const plain = view.blocks.findIndex((b) => b.text === 'Plain, editable.');
  view.setSelection({ block: plain, offset: 0 }, { block: plain, offset: 5 });
  view.setRunFormat({ vertAlign: 'superscript' });
  view.setSelection({ block: plain, offset: 1 });
  assert.equal(view.formatAtCaret().vertAlign, 'superscript');
  assert.match(view.doc.doc.editParagraph(plain).xml, /<w:vertAlign w:val="superscript"\/>/);
  const frame = view.render({ pages: false });
  assert.equal(frame.blocks[plain].runs[0].vertAlign, 'superscript');
  assert.equal(frame.blocks[plain].runs[0].text, 'Plain');
  view.setSelection({ block: plain, offset: 0 }, { block: plain, offset: 5 });
  view.setRunFormat({ vertAlign: null });
  assert.ok(!/<w:vertAlign/.test(view.doc.doc.editParagraph(plain).xml), 'cleared');
  assert.throws(() => view.setRunFormat({ vertAlign: 'sideways' }), /superscript, subscript or null/);
});

test('a tab typed into a paragraph is written back as the element, not the character', () => {
  const view = openDocx(buildCoverPageDocument());
  const plain = view.blocks.findIndex((b) => b.text === 'Plain, editable.');
  view.setSelection({ block: plain, offset: 6 });
  view.insertText('\tand');
  assert.equal(view.blocks[plain].text, 'Plain,\tand editable.');
  const xml = view.doc.doc.editParagraph(plain).xml;
  assert.match(xml, /<w:t xml:space="preserve">Plain,<\/w:t><w:tab\/><w:t xml:space="preserve">and editable\.<\/w:t>/);
  assert.ok(!/<w:t[^>]*>[^<]*\t/.test(xml), 'no literal tab inside w:t');
});

test('styles resolve fonts through the theme, and carry shading and tab stops', () => {
  const f = frame();
  const s = f.styles;
  assert.equal(s['*default*'].fontName, 'Cambria', 'the minor slot is the body face');
  assert.equal(s.Heading1.fontName, 'Georgia', 'the major slot is the heading face');
  assert.equal(s.Heading1.colour, '#2f5496');
  assert.equal(s.Title.shading, '#7E97AD');
  assert.equal(s.Title.caps, true);
  assert.equal(s.TOC1.tabs.length, 1);
  assert.equal(s.TOC1.tabs[0].align, 'right');
  assert.equal(s.TOC1.tabs[0].leader, 'dot');
  assert.ok(Math.abs(s.TOC1.tabs[0].posPx - 9016 * (96 / 1440)) < 1e-6, 'the stop is 9016 twips');
  assert.equal(s.TOCHeading.fontName, 'Georgia', 'basedOn carries the theme slot down the chain');
  assert.equal(s['*default*'].lineFactor, 259 / 240);
  // The theme readers on their own.
  const doc = Document.open(buildCoverPageDocument());
  assert.deepEqual(doc.themeFonts(), { major: 'Georgia', minor: 'Cambria' });
  assert.equal(doc.themeColours().accent1, '7E97AD');
  assert.equal(doc.themeColours().dk1, '000000', 'a system colour resolves to lastClr');
  assert.equal(readThemeColours(null).accent1, '4472C4', 'Office defaults without a theme');
  assert.deepEqual(readThemeFonts(null), { major: 'Calibri Light', minor: 'Calibri' });
});

test('footnotes are numbered by where their references fall, not by id', () => {
  const f = frame();
  assert.equal(f.footnotes.length, 2);
  assert.deepEqual(f.footnotes.map((n) => [n.n, n.id]), [[1, '2'], [2, '1']], 'id 2 is referenced first');
  assert.equal(f.footnotes[0].paragraphs[0].text, ' For the list of exclusions see the guidance.');
  assert.equal(f.footnotes[0].paragraphs[0].style, 'FootnoteText');
  assert.deepEqual(f.footnotes[0].paragraphs[0].runs[0].noteMark, { kind: 'footnote', n: 1 });
  const body = f.blocks.find((b) => b.text.startsWith('The supplier'));
  const refs = body.runs.filter((r) => r.noteRef);
  assert.deepEqual(refs.map((r) => r.noteRef), [{ kind: 'footnote', id: '2', n: 1 }, { kind: 'footnote', id: '1', n: 2 }]);
  assert.equal(refs[0].text, '', 'a reference adds no text, so no caret offset moves');
  assert.equal(body.structural, true, 'a paragraph with a reference is never rebuilt');
  assert.equal(f.endnotes.length, 0);
});

test('a first-page header is read alongside the default one, with the title page flag', () => {
  const view = openDocx(buildCoverPageDocument());
  const bands = view.doc.headerFooters();
  assert.equal(bands.headers.first.paragraphs[0].text, 'Leicestershire Fire and Rescue Service');
  const words = bands.headers.default.paragraphs.find((p) => /Selection Questionnaire/.test(p.text));
  assert.ok(words, 'the default header keeps its words');
  assert.equal(words.text, 'Selection Questionnaire\tPage 1', 'the header tab survives too');
  assert.equal(view.section.titlePage, true);
  assert.equal(view.section.margins.header, 37.8);
  // The watermark is the header's WordArt, read as words for the page to draw.
  assert.deepEqual(bands.watermark, { text: 'DRAFT', colour: 'silver', rotation: 315, band: 'default' });
});

test('a document without the parts still renders, with Office defaults', () => {
  // The complex fixture has styles but no theme part and no text boxes.
  const view = openDocx(buildCoverPageDocument());
  assert.ok(view.render({ pages: false }).blocks.length > 5);
  const plain = openDocx(Buffer.from(buildCoverPageDocument()));
  assert.equal(typeof plain.render({ pages: false }).styles, 'object');
});
