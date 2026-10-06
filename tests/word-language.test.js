// Review → Language in Documents: words marked as a language, or as not to
// be checked, kept in the run as Word keeps them (w:lang, w:noProof), read
// past by the spelling pass — and the document's default language, kept in
// the stylesheet's defaults.

import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml/build';
import { OoxmlPackage, langElement } from '@rutba/ooxml';
import { wordSegments, wordLanguage, setWordDefaultLanguage, nextMisspelling } from '@rutba/proofing';
import { createDocumentService } from '../apps/desktop/main/documents.js';
import { createProofing } from '../apps/desktop/main/proofing.js';

const fakeDictionary = (wrong) => async (list) => new Set(list.filter((w) => wrong.includes(w)));

test('a language fills the slot for its script, keeps the others, and a tag that is not one is refused', () => {
  assert.equal(langElement(null, 'fr-FR'), '<w:lang w:val="fr-FR"/>');
  assert.equal(langElement(' w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"', 'de-DE'), '<w:lang w:val="de-DE" w:eastAsia="zh-CN" w:bidi="ar-SA"/>');
  assert.equal(langElement(' w:val="en-US" w:bidi="ar-SA"', 'ur-PK'), '<w:lang w:val="ur-PK" w:bidi="ur-PK"/>', 'Urdu proofs right-to-left text too');
  assert.equal(langElement(null, 'ja-JP'), '<w:lang w:val="ja-JP" w:eastAsia="ja-JP"/>');
  assert.equal(langElement(' w:val="fr-FR"', null), null, 'clearing the only slot leaves no element');
  assert.throws(() => langElement(null, 'fr"FR'), /not a language tag/);
});

test('marking words writes w:lang and w:noProof in their schema places, reads back at the caret and survives a save', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Hello bonjour world' }] }));
  view.setSelection({ block: 0, offset: 6 }, { block: 0, offset: 13 });
  view.toggleFormat('b');
  view.setRunFormat({ lang: 'fr-FR', fontColour: 'C00000' });
  view.setSelection({ block: 0, offset: 14 }, { block: 0, offset: 19 });
  view.setRunFormat({ noProof: true });
  view.setSelection({ block: 0, offset: 8 });
  assert.equal(view.formatAtCaret().lang, 'fr-FR');
  assert.equal(view.formatAtCaret().noProof, false);
  view.setSelection({ block: 0, offset: 16 });
  assert.equal(view.formatAtCaret().noProof, true);

  const xml = OoxmlPackage.read(view.save()).text('word/document.xml');
  assert.match(xml, /<w:rPr><w:b\/><w:color w:val="C00000"\/><w:lang w:val="fr-FR"\/><\/w:rPr><w:t[^>]*>bonjour<\/w:t>/, 'lang after colour, as the schema orders them');
  assert.match(xml, /<w:rPr><w:noProof\/><\/w:rPr><w:t[^>]*>world<\/w:t>/);
  const frame = view.render({ pages: false });
  const runs = frame.blocks[0].runs;
  assert.equal(runs.find((r) => r.text === 'bonjour').lang, 'fr-FR', 'the painter is told the language');
  assert.equal(runs.find((r) => r.text === 'world').noProof, true);
});

test('Spelling reads past French words, words marked not to be checked, and words in a script English cannot spell', async () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Teh chat noir est ici, wrold مرحبا dun.' }] }));
  view.setSelection({ block: 0, offset: 4 }, { block: 0, offset: 21 });
  view.setRunFormat({ lang: 'fr-FR' });
  view.setSelection({ block: 0, offset: 23 }, { block: 0, offset: 28 });
  view.setRunFormat({ noProof: true });
  const segments = wordSegments(view);
  assert.deepEqual(segments[0].skip, [[4, 21], [23, 28]]);

  const misspelt = fakeDictionary(['Teh', 'chat', 'noir', 'est', 'ici', 'wrold', 'مرحبا', 'dun']);
  const seen = [];
  const stop = { key: segments[0].key, offset: 0 };
  let step = await nextMisspelling({ segments, from: stop, stop, misspelt });
  while (step.found) {
    seen.push(step.found.word);
    step = await nextMisspelling({ segments, from: step.next, stop, wrapped: step.wrapped, misspelt });
  }
  assert.deepEqual(seen, ['Teh', 'dun']);
});

test('a text box, a footnote and a header carry their runs\' marks to the pass too', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Body' }] }));
  view.setBand('header', ['Le titre']);
  const doc = view.doc.doc;
  const band = Object.values(doc.headerFooters().headers)[0];
  doc.pkg.write_(band.part, Buffer.from(doc.pkg.text(band.part).replace('<w:r>', () => '<w:r><w:rPr><w:lang w:val="fr-FR"/></w:rPr>'), 'utf8'));
  const header = wordSegments(view).find((s) => s.where.story === 'header');
  assert.equal(header.text, 'Le titre');
  assert.deepEqual(header.skip, [[0, 8]]);
});

test('Set As Default writes the language into the stylesheet\'s defaults, keeps what was there, and undoes', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Colour' }] }));
  setWordDefaultLanguage(view, { lang: 'en-US' });
  const doc = view.doc.doc;
  assert.match(doc.pkg.text('word/styles.xml'), /<w:styles\b[^>]*><w:docDefaults><w:rPrDefault><w:rPr><w:lang w:val="en-US"\/><\/w:rPr><\/w:rPrDefault><\/w:docDefaults><w:style /);
  assert.equal(wordLanguage(view), 'en-US');

  // Word's own defaults: the Latin slot changes, the other two and the fonts stay.
  const styles = doc.pkg.text('word/styles.xml').replace(/<w:docDefaults>[\s\S]*?<\/w:docDefaults>/, () =>
    '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi"/><w:sz w:val="22"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160"/></w:pPr></w:pPrDefault></w:docDefaults>');
  doc.pkg.write_('word/styles.xml', Buffer.from(styles, 'utf8'));
  setWordDefaultLanguage(view, { lang: 'en-GB' });
  assert.match(doc.pkg.text('word/styles.xml'), /<w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi"\/><w:sz w:val="22"\/><w:lang w:val="en-GB" w:eastAsia="en-US" w:bidi="ar-SA"\/><\/w:rPr><\/w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160"\/><\/w:pPr><\/w:pPrDefault>/);
  assert.equal(wordLanguage(view), 'en-GB');
  view.undo();
  assert.equal(wordLanguage(view), 'en-US', 'one undo puts the stylesheet back');
  assert.equal(OoxmlPackage.read(view.save()).text('word/styles.xml').includes('w:val="en-US" w:eastAsia'), true, 'and the saved file agrees');
});

test('a document with no stylesheet is given one to keep its default in', () => {
  const view = openDocx(buildDocx({ paragraphs: [{ text: 'Plain' }] }));
  assert.equal(view.doc.doc.pkg.has('word/styles.xml'), false);
  setWordDefaultLanguage(view, { lang: 'cy-GB' });
  const saved = OoxmlPackage.read(view.save());
  assert.match(saved.text('word/styles.xml'), /<w:docDefaults><w:rPrDefault><w:rPr><w:lang w:val="cy-GB"\/>/);
  assert.match(saved.text('word/_rels/document.xml.rels'), /relationships\/styles" Target="styles\.xml"/);
});

test('the document service: the Language dialog\'s query, its marks as one op, and Set As Default as another', async () => {
  const proofing = createProofing({ worker: false, locale: () => 'en-GB' });
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }), proofing });
  const { id } = docs.new({ kind: 'doc' });
  docs.apply({ id, ops: [{ op: 'selectAll' }, { op: 'insertText', text: 'Guten Tag, Zorblat.' }] });
  docs.apply({ id, ops: [{ op: 'setSelection', anchor: { block: 0, offset: 0 }, focus: { block: 0, offset: 9 } }, { op: 'setRunFormat', delta: { lang: 'de-DE' } }] });
  docs.apply({ id, ops: [{ op: 'setDefaultLanguage', lang: 'en-US' }] });
  const info = await docs.proof({ id, action: 'language' });
  assert.deepEqual(info, { document: 'en-US', default: 'en-US', checking: 'en-US', checkingName: 'English (United States)' });
  const start = await docs.proof({ id, action: 'spellStart' });
  const step = await docs.proof({ id, action: 'spellNext', from: start.start, stop: start.start });
  assert.equal(step.found.word, 'Zorblat', 'the German words are read past');
  proofing.close();
});
