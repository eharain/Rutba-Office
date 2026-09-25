// Review → Spelling: the words a pass checks, the dictionaries it checks
// them against, the person's own dictionary, the order it reads a document
// in from the caret, and each app's text — a Word document's body, tables,
// text boxes, notes and headers; a sheet's cells; a deck's shapes, tables
// and notes — changed where each word is.

import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml';
import { SheetView } from '@rutba/sheet-view';
import { Deck, buildPptx } from '@rutba/presentation';
import {
  tokenize, matchCase, contextOf, acceptedBy, toDic, fromDic, chooseLanguage, dictionaryFor,
  loadDictionary, checkWords, suggestWords, nextMisspelling, changeAllEdits,
  wordSegments, wordStart, wordLanguage, replaceWordText,
  sheetSegments, replaceSheetText, deckSegments, deckLanguage, replaceDeckText,
} from '@rutba/proofing';
import { createDocumentService } from '../apps/desktop/main/documents.js';
import { createProofing } from '../apps/desktop/main/proofing.js';

const words = (text, o) => tokenize(text, o).map((t) => t.word);

test('the tokenizer skips UPPERCASE, words with numbers, web and e-mail addresses and file paths, as Office does by default', () => {
  const text = 'Send the NATO brief to jo@example.com, see https://rutba.io/office and C:\\Reports\\q3.docx; version2 is teh one’s.';
  assert.deepEqual(words(text), ['Send', 'the', 'brief', 'to', 'see', 'and', 'is', 'teh', "one's"]);
  const t = tokenize(text).find((w) => w.word === 'teh');
  assert.equal(text.slice(t.offset, t.offset + t.length), 'teh', 'offsets are the text\'s own');
  assert.ok(words(text, { ignoreUppercase: false }).includes('NATO'));
  assert.ok(words(text, { ignoreNumbers: false }).includes('version2'));
  assert.ok(words('see https://rutba.io', { ignoreAddresses: false }).includes('rutba'));
  assert.deepEqual(words('well-known state-of-the-art'), ['well', 'known', 'state', 'of', 'the', 'art'], 'a hyphenated word a part at a time');
  assert.equal(matchCase('Recieve', 'receive'), 'Receive');
  assert.equal(matchCase('recieve', 'receive'), 'receive');
  assert.deepEqual(contextOf('The quick brown fox', 10, 5), { before: 'The quick ', word: 'brown', after: ' fox' });
});

test('the dictionaries: English (UK) and English (US) each know their own spellings, with suggestions', async () => {
  const gb = await loadDictionary('en-GB');
  const us = await loadDictionary('en-US');
  assert.deepEqual(checkWords(gb.spell, ['colour', 'organisation', 'color']), ['color']);
  assert.deepEqual(checkWords(us.spell, ['color', 'colour']), ['colour']);
  assert.deepEqual(checkWords(gb.spell, ["Northwind's", 'The', "don't"]), ["Northwind's"], 'a sentence\'s first capital and a contraction are words');
  assert.ok(suggestWords(gb.spell, 'recieve').includes('receive'));
  assert.ok(suggestWords(gb.spell, 'accomodate').includes('accommodate'));
});

test('the language: the document\'s own where it has one we spell, else the system\'s English', () => {
  assert.deepEqual(chooseLanguage('en-US', 'en-GB').lang, 'en-US');
  assert.deepEqual(chooseLanguage('en-AU', 'en-US').lang, 'en-GB', 'Australian English is spelt the British way');
  assert.deepEqual(chooseLanguage(null, 'en-US'), { lang: 'en-US', from: 'system', documentLang: null });
  assert.deepEqual(chooseLanguage(null, 'de-DE').lang, 'en-GB');
  assert.deepEqual(chooseLanguage('fr-FR', 'en-US'), { lang: 'en-US', from: 'fallback', documentLang: 'fr-FR' });
  assert.equal(dictionaryFor('en_US'), 'en-US');
  assert.equal(dictionaryFor('cy-GB'), null);

  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Colour' }] }));
  const doc = view.doc.doc;
  doc.xml = doc.xml.replace('<w:r>', '<w:r><w:rPr><w:lang w:val="en-US"/></w:rPr>');
  view._invalidate();
  assert.equal(wordLanguage(view), 'en-US', 'a document\'s runs say which English');
  const deck = Deck.open(buildPptx({ slides: [{ layout: 'title', title: 'Hello', body: 'World' }] }));
  assert.match(deckLanguage(deck) || '', /^en-/, 'a deck\'s runs carry lang');
});

test('the person\'s own dictionary: accepted as written or in any capitals of a lower-case entry, and kept as a .dic file', () => {
  const personal = new Set(['Rutba', 'kerning']);
  assert.ok(acceptedBy(personal, 'Rutba'));
  assert.ok(!acceptedBy(personal, 'rutba'), 'a proper name stays capitalised');
  assert.ok(acceptedBy(personal, 'Kerning') && acceptedBy(personal, 'kerning'));
  const dic = toDic(['zeta', 'Rutba', 'alpha', 'alpha']);
  assert.equal(dic, '3\nalpha\nRutba\nzeta\n');
  assert.deepEqual(fromDic('\uFEFF3\r\nalpha/MS\r\n# a note\r\n\r\nRutba\r\nzeta\r\n'), ['alpha', 'Rutba', 'zeta']);
  assert.deepEqual(fromDic('one\ntwo words\nthree'), ['one', 'three'], 'a list without the count; a line with a space is not a word');
});

const SEGMENTS = [
  { key: 'a', text: 'Frist words here.' },
  { key: 'b', text: 'Then teh caret sits here, recieve.' },
  { key: 'c', text: 'After the caret, speling.' },
];
const fakeDictionary = (wrong) => async (list) => new Set(list.filter((w) => wrong.includes(w)));

test('the pass reads from the caret to the end, goes round to the start and stops where it began', async () => {
  const misspelt = fakeDictionary(['Frist', 'teh', 'recieve', 'speling']);
  const stop = { key: 'b', offset: 9 }; // the caret after "teh"
  const seen = [];
  let step = await nextMisspelling({ segments: SEGMENTS, from: stop, stop, misspelt });
  while (step.found) {
    seen.push(`${step.found.word}${step.wrapped ? ' (after going round)' : ''}`);
    step = await nextMisspelling({ segments: SEGMENTS, from: step.next, stop, wrapped: step.wrapped, misspelt });
  }
  assert.deepEqual(seen, ['recieve', 'speling', 'Frist (after going round)', 'teh (after going round)']);
  assert.equal(step.done, true);

  // Ignore All and the person's dictionary are both "accepted".
  const quiet = await nextMisspelling({ segments: SEGMENTS, from: { key: 'a', offset: 0 }, stop: { key: 'a', offset: 0 }, misspelt, accepted: (w) => ['Frist', 'teh'].includes(w) });
  assert.equal(quiet.found.word, 'recieve');
  assert.deepEqual(quiet.found.context, { before: 'Then teh caret sits here, ', word: 'recieve', after: '.' });
});

/** A document with a misspelling in every story Word has. */
function wordFixture() {
  const bytes = buildDocx({ styles: true, paragraphs: [
    { text: 'The frist paragraph.' },
    { table: { rows: [['Name', 'Adress'], ['Jo', 'York']] } },
    { text: 'A lone paragraph.' },
    { text: 'Anchor' },
  ] });
  // A text box, anchored in the last paragraph.
  const pkg = OoxmlPackage.read(bytes);
  const box = '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor><wp:extent cx="914400" cy="457200"/><wp:docPr id="9" name="Box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:t>Boxed wrods here</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:textbox><w:txbxContent><w:p><w:r><w:t>Boxed wrods here</w:t></w:r></w:p></w:txbxContent></v:textbox></w:pict></mc:Fallback></mc:AlternateContent></w:r>';
  const ns = ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml"';
  // The box rides the last paragraph, which cannot then be typed into.
  const xml = pkg.text('word/document.xml').replace('<w:document ', '<w:document' + ns + ' ');
  const last = xml.lastIndexOf('<w:p>');
  pkg.write_('word/document.xml', Buffer.from(xml.slice(0, last + 5) + box + xml.slice(last + 5), 'utf8'));
  const view = openDocx(pkg.write());
  view.setSelection({ block: 0, offset: 3 });
  view.insertNote('footnote', 'A footnoet here.');
  view.setBand('header', ['The headder line']);
  return view;
}

test('Word\'s walker reads the body, the table\'s cells, the text box, the footnote and the header — and changes each where it is', () => {
  const view = wordFixture();
  const segments = wordSegments(view);
  const stories = new Set(segments.map((s) => s.where.story));
  for (const story of ['body', 'textbox', 'footnote', 'header']) assert.ok(stories.has(story), story);
  const find = (w) => segments.find((s) => s.text.includes(w));
  for (const w of ['frist', 'Adress', 'wrods', 'footnoet', 'headder']) assert.ok(find(w), `${w} is read`);
  assert.equal(wordStart(view).key, 'b:0', 'the pass starts at the caret\'s paragraph');

  const edits = ['frist', 'Adress', 'wrods', 'footnoet', 'headder'].map((w) => {
    const s = find(w);
    const at = s.text.indexOf(w);
    return { key: s.key, from: at, to: at + w.length, text: { frist: 'first', Adress: 'Address', wrods: 'words', footnoet: 'footnote', headder: 'header' }[w] };
  });
  replaceWordText(view, edits);
  const after = wordSegments(view).map((s) => s.text).join(' | ');
  for (const w of ['first', 'Address', 'Boxed words here', 'footnote here', 'The header line']) assert.ok(after.includes(w), `${w}: ${after}`);
  const saved = OoxmlPackage.read(view.save());
  assert.match(saved.text('word/footnotes.xml'), /footnote here/);
  view.undo();
  assert.ok(wordSegments(view).some((s) => s.text.includes('headder')), 'one undo puts every story back');
});

test('Change All changes the word everywhere in a document, each in its own capitals', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Recieve it. We recieve daily.' }, { text: 'They will recieve.' }] }));
  view.setBand('footer', ['Please recieve']);
  const edits = changeAllEdits(wordSegments(view), 'recieve', 'receive');
  assert.equal(edits.length, 4);
  replaceWordText(view, edits);
  assert.deepEqual(wordSegments(view).map((s) => s.text), ['Receive it. We receive daily.', 'They will receive.', 'Please receive']);
});

test('the sheet walker reads text cells row by row — never formulas or numbers — and changes a cell\'s text', () => {
  const view = new SheetView(buildXlsx({ sheets: [
    { name: 'Notes', rows: [['Itme', 12, '=A1'], ['Blue pnes', 'NATO', 'ok']] },
    { name: 'Other', rows: [['Anothr']] },
  ] }));
  const segments = sheetSegments(view);
  assert.deepEqual(segments.map((s) => s.text), ['Itme', 'Blue pnes', 'NATO', 'ok']);
  assert.deepEqual(sheetSegments(view, { sheets: ['Other'] }).map((s) => s.where.ref), ['A1']);
  const pnes = segments[1];
  replaceSheetText(view, [{ key: pnes.key, from: 5, to: 9, text: 'pens' }]);
  assert.equal(view.editValue(1, 0), 'Blue pens');
  view.undo();
  assert.equal(view.editValue(1, 0), 'Blue pnes');
});

test('the deck walker reads every shape\'s runs, a table\'s cells and the notes, and Change All rewrites them in place', () => {
  const deck = Deck.open(buildPptx({ slides: [
    { layout: 'title', title: 'Quartely review', body: 'Spring' },
    { layout: 'obj', title: 'Results', body: ['Quartely orders up'] },
  ] }));
  const table = deck.addTable(1, { rows: 1, cols: 2, cells: [['Quartely', 'Total']] });
  deck.setNotes(1, 'Mention the quartely numbers.');
  const segments = deckSegments(deck);
  assert.ok(segments.some((s) => s.where.notes && /quartely/.test(s.text)));
  assert.ok(segments.some((s) => String(s.where.shape) === String(table) && s.where.row === 0));
  const edits = changeAllEdits(segments, 'quartely', 'quarterly');
  assert.equal(edits.length, 4);
  replaceDeckText(deck, edits);
  const text = deckSegments(deck).map((s) => s.text).join(' | ');
  assert.ok(!/quartely/i.test(text), text);
  assert.match(text, /Quarterly review/);
  assert.match(text, /Mention the quarterly numbers/);
});

test('the document service: a pass from the caret finds the planted word, Change and Add to Dictionary, then the end', async () => {
  const stores = (() => { const m = new Map(); return { settings: { get: (k, f) => (m.has(k) ? m.get(k) : f), set: (k, v) => m.set(k, v) } }; })();
  const proofing = createProofing({ stores, worker: false, locale: () => 'en-GB' });
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }), proofing });
  const { id } = docs.new({ kind: 'doc' });
  // A window that opens warms its dictionary without waiting on it.
  const warm = await docs.proof({ id, action: 'spellWarm' });
  assert.equal(warm.lang, 'en-GB', 'the warm-up answers at once with the language it reads');
  docs.apply({ id, ops: [{ op: 'selectAll' }, { op: 'insertText', text: 'Our Zorblat plan will recieve support.' }, { op: 'setSelection', anchor: { block: 0, offset: 0 } }] });
  const start = await docs.proof({ id, action: 'spellStart' });
  assert.equal(start.lang, 'en-GB');
  let step = await docs.proof({ id, action: 'spellNext', from: start.start, stop: start.start });
  assert.equal(step.found.word, 'Zorblat');
  await docs.proof({ id, action: 'addWord', word: 'Zorblat' });
  assert.deepEqual(stores.settings.get('proofing.dictionary'), ['Zorblat'], 'kept in the settings store');
  step = await docs.proof({ id, action: 'spellNext', from: step.next, stop: start.start });
  assert.equal(step.found.word, 'recieve');
  assert.equal(step.found.suggestions, undefined, 'the word comes back at once; its suggestions are asked for next');
  assert.ok((await docs.proof({ id, action: 'spellSuggest', word: step.found.word })).suggestions.includes('receive'));
  docs.apply({ id, ops: [{ op: 'spellReplace', edits: [{ key: step.found.key, from: step.found.offset, to: step.found.offset + step.found.length, text: 'receive' }] }] });
  step = await docs.proof({ id, action: 'spellNext', from: { key: step.found.key, offset: step.found.offset + 7 }, stop: start.start });
  assert.equal(step.done, true);
  assert.equal(docs.model({ id }).blocks[0].text, 'Our Zorblat plan will receive support.');
  const again = await docs.proof({ id, action: 'spellNext', from: start.start, stop: start.start });
  assert.equal(again.done, true, 'the added word stays accepted on the next pass');
  const word = await docs.proof({ id, action: 'spellCheckWord', word: 'plann' });
  assert.equal(word.misspelt, true);
  assert.ok(word.suggestions.includes('plan'));
  proofing.close();
});

test('a word split across runs is changed where it fell, each run keeping its look, and escaped text stays escaped', async () => {
  const { paragraphText, replaceInParagraph, createSpellerHost } = await import('@rutba/proofing');
  const p = '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Fish &amp; chi</w:t></w:r><w:r><w:t>sp shop</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t xml:space="preserve"> open</w:t></w:r></w:p>';
  assert.equal(paragraphText(p).text, 'Fish & chisp shop\t open');
  const fixed = replaceInParagraph(p, 7, 12, 'chips');
  assert.equal(paragraphText(fixed).text, 'Fish & chips shop\t open');
  assert.match(fixed, /<w:rPr><w:b\/><\/w:rPr><w:t>Fish &amp; chips<\/w:t>/, 'the bold run takes the word');
  assert.match(fixed, /<w:r><w:t xml:space="preserve"> shop<\/w:t><\/w:r>/, 'the next run keeps the rest, its space kept');

  // The dictionaries on a worker thread, as the application asks them.
  const host = createSpellerHost({ worker: true });
  const wrong = await host.misspelt('en-GB', ['colour', 'colur', 'recieve']);
  assert.deepEqual([...wrong], ['colur', 'recieve']);
  assert.ok((await host.suggest('en-GB', 'colur')).includes('colour'));
  assert.ok(Number.isFinite(host.timings['en-GB']), 'how long the dictionary took is kept');
  host.close();
});
