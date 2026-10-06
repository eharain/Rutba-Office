// Design → Themes, Colours, Fonts and Effects in Documents: the theme part
// written as Word keeps it, and the page following it — a colour named by
// theme slot (Word's headings are accent 1, darkened) drawn from the theme
// rather than the hex beside it, the heading and body faces from its fonts.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { OoxmlPackage, themeColourHex, readThemeColours } from '@rutba/ooxml';
import { buildDocx } from '@rutba/ooxml/build';
import { PALETTES } from '@rutba/office-formats/themes';
import { createDocumentService } from '../apps/desktop/main/documents.js';

const near = (a, b, by = 4) => [0, 2, 4].every((i) => Math.abs(parseInt(a.slice(i, i + 2), 16) - parseInt(b.slice(i, i + 2), 16)) <= by);

test('a theme colour is drawn as Word draws it: the slot, darkened by its shade or lightened by its tint', () => {
  const office = readThemeColours(null);
  assert.equal(themeColourHex(office, 'accent1'), '4472C4');
  assert.ok(near(themeColourHex(office, 'accent1', null, 'BF'), '2F5496'), 'Word\'s Heading 1 blue is accent 1 shaded to three quarters');
  assert.ok(near(themeColourHex(office, 'accent1', '99', null), '8FAADC'), 'and its light accent is accent 1 tinted');
  assert.equal(themeColourHex(office, 'text1'), '000000', 'text 1 is dark 1');
  assert.equal(themeColourHex(office, 'none'), null);
});

/** A document as Word writes it: Heading 1 in accent 1 shaded, in the heading face; a run in accent 2. */
function wordLike() {
  const pkg = OoxmlPackage.read(buildDocx({ styles: true, paragraphs: [{ text: 'A heading', style: 'Heading1' }, { text: 'Body words' }] }));
  const styles = pkg.text('word/styles.xml').replace(/(<w:style w:type="paragraph" w:styleId="Heading1">[\s\S]*?<w:rPr>)([\s\S]*?)(<\/w:rPr>)/, (m, open, props, close) =>
    open + '<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>' + props.replace(/<w:color\b[^>]*\/>/, '') + '<w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>' + close);
  pkg.write_('word/styles.xml', Buffer.from(styles, 'utf8'));
  const doc = pkg.text('word/document.xml').replace('<w:r><w:t xml:space="preserve">Body words', () => '<w:r><w:rPr><w:color w:val="ED7D31" w:themeColor="accent2"/></w:rPr><w:t xml:space="preserve">Body words');
  pkg.write_('word/document.xml', Buffer.from(doc, 'utf8'));
  return pkg.write();
}

test('the document service: Colours, Fonts, Effects and a whole theme rewrite the theme part, and the page follows each; one undo each', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-doctheme-'));
  const file = path.join(dir, 'themed.docx');
  fs.writeFileSync(file, wordLike());
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const { id } = await docs.open({ path: file });
  const model = () => docs.model({ id });
  const body = () => model().blocks.find((b) => b.text === 'Body words').runs[0];
  assert.equal(model().design.exists, false, 'no theme part: Office\'s');
  assert.ok(near(model().resolvedStyles.Heading1.colour.slice(1).toUpperCase(), '2F5496'), 'Office accent 1, shaded');
  assert.equal(body().fontColour, 'ED7D31');

  // Colours: Berry.
  const berry = PALETTES.find((p) => p.id === 'berry').colors;
  docs.apply({ id, ops: [{ op: 'setDocTheme', spec: { colors: 'berry' } }] });
  assert.equal(model().design.colorName, 'Berry');
  assert.equal(body().fontColour, berry.accent2, 'the run follows accent 2');
  assert.ok(near(model().resolvedStyles.Heading1.colour.slice(1).toUpperCase(), themeColourHex(berry, 'accent1', null, 'BF')), 'the heading follows accent 1, still shaded');

  // Fonts: Classic.
  docs.apply({ id, ops: [{ op: 'setDocTheme', spec: { fonts: 'classic' } }] });
  assert.equal(model().resolvedStyles.Heading1.fontName, 'Georgia', 'the heading face is the theme\'s');
  assert.equal(model().design.colorName, 'Berry', 'and the colours stay');

  // Effects, then a whole theme.
  docs.apply({ id, ops: [{ op: 'setDocTheme', spec: { effects: 'lifted' } }] });
  assert.equal(model().design.effects, 'lifted');
  docs.apply({ id, ops: [{ op: 'setDocTheme', spec: { theme: 'harbour' } }] });
  assert.equal(model().design.builtIn, 'harbour');

  // One undo puts the effects back, another the fonts.
  docs.undo({ id });
  assert.equal(model().design.effects, 'lifted');
  docs.undo({ id });
  docs.undo({ id });
  assert.equal(model().design.fontName, 'Office', 'back to before Fonts: Office\'s faces');
  assert.equal(model().design.colorName, 'Berry', 'with Berry\'s colours');
  assert.notEqual(model().resolvedStyles.Heading1.fontName, 'Georgia');

  // Saved: the part, its relationship and its content type.
  const out = path.join(dir, 'saved.docx');
  docs.save({ id, path: out });
  const saved = OoxmlPackage.read(fs.readFileSync(out));
  assert.match(saved.text('word/theme/theme1.xml'), /<a:clrScheme name="Berry">/);
  assert.match(saved.text('word/_rels/document.xml.rels'), /relationships\/theme" Target="theme\/theme1\.xml"/);
  assert.equal(saved.contentTypeOf('word/theme/theme1.xml'), 'application/vnd.openxmlformats-officedocument.theme+xml');
  assert.throws(() => docs.apply({ id, ops: [{ op: 'setDocTheme', spec: { colors: 'no-such-palette' } }] }), /no palette/);
});

test('Set as Default: a document\'s theme and styles, and a new blank document made in them, unmodified', async () => {
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const source = docs.new({ kind: 'word', template: 'doc' });
  docs.apply({ id: source.id, ops: [{ op: 'setDocTheme', spec: { theme: 'meadow' } }] });
  const design = docs.design({ id: source.id });
  assert.equal(design.name, 'Meadow');
  assert.match(design.theme, /<a:theme\b[^>]*name="Meadow"/);
  assert.match(design.styles, /<w:styles\b/);

  const made = docs.new({ kind: 'word', template: 'doc', design: { theme: design.theme, styles: design.styles } });
  assert.equal(made.model.design.name, 'Meadow', 'the new document starts in the default theme');
  assert.equal(made.dirty, false, 'and nothing has been done to it yet');
  const plain = docs.new({ kind: 'word', template: 'doc' });
  assert.notEqual(plain.model.design.name, 'Meadow', 'without the default, the suite\'s own');
  const junk = docs.new({ kind: 'word', template: 'doc', design: { theme: '<html>not a theme</html>', styles: 42 } });
  assert.notEqual(junk.model.design.name, 'Meadow', 'what is not a theme part is not put on');
});
