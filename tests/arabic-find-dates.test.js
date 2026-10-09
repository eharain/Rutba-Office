// Arabic in the engines: Find matches a word whatever marks it is written
// with, as Office in Arabic does by default (Match diacritics, kashida and
// alef hamza off), in a worksheet, a document and a deck alike; and a number
// format can name the Hijri calendar and a set of digits, as Excel's do.
import test from 'node:test';
import assert from 'node:assert/strict';
import { foldForSearch, findAll, replaceAllIn, includesFolded } from '@rutba/editing/find';
import { SheetView, formatValue } from '@rutba/sheet-view';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx, buildXlsx } from '@rutba/ooxml';
import { Deck, buildPptx } from '@rutba/presentation';
import { dateToSerial } from '@rutba/formula';

const WITH_MARKS = 'كَتَبَ';
const STRETCHED = 'كتـــب';
const BARE = 'كتب';

test('find passes over the vowel marks, the kashida and the forms of alef and yeh, unless asked to match them', () => {
  const text = `قال ${WITH_MARKS} ثم ${STRETCHED} أحمد إلى احمد، على علي. Hello WORLD`;
  assert.deepEqual(findAll(text, BARE).map((m) => text.slice(m.start, m.end)), [WITH_MARKS, STRETCHED]);
  assert.deepEqual(findAll(text, 'احمد').map((m) => text.slice(m.start, m.end)), ['أحمد', 'احمد']);
  assert.deepEqual(findAll(text, 'علي').map((m) => text.slice(m.start, m.end)), ['على', 'علي'], 'alef maqsura is yeh');
  assert.deepEqual(findAll(text, BARE, { matchDiacritics: true }), [], 'matched as written, the bare word is not there');
  assert.equal(findAll(text, WITH_MARKS, { matchDiacritics: true }).length, 1);
  assert.equal(findAll(text, 'world').length, 1);
  assert.equal(findAll(text, 'world', { matchCase: true }).length, 0);
  // Urdu's yeh and kaf are Arabic's, so an Urdu word is found either way.
  assert.ok(includesFolded('کتاب یہ ہے', 'كتاب'));
  assert.ok(includesFolded('کتاب یہ ہے', 'يہ'));
});

test('a match runs from its first letter to the end of the marks on its last, so replacing it takes them along', () => {
  const { text, count } = replaceAllIn(`أ ${WITH_MARKS} ب`, BARE, 'قرأ');
  assert.equal(count, 1);
  assert.equal(text, 'أ قرأ ب');
  const folded = foldForSearch(WITH_MARKS);
  assert.equal(folded.text, BARE);
  assert.equal(folded.map[folded.text.length], WITH_MARKS.length);
  // A needle that is only a mark looks for that mark.
  assert.equal(findAll('\u0627\u064E\u0628\u064E', '\u064E').length, 2);
});

test('a worksheet finds and replaces Arabic as Excel in Arabic does', () => {
  const view = SheetView.open(buildXlsx({ sheets: [{ name: 'Data', rows: [['English'], [`المبلغ ${WITH_MARKS}`], ['أحمد']] }] }));
  assert.equal(view.findNext(BARE), true);
  assert.equal(view.selection.active.row, 1);
  assert.equal(view.findNext('احمد'), true);
  assert.equal(view.selection.active.row, 2);
  assert.equal(view.replaceAll(BARE, 'قرأ'), 1);
  assert.equal(view.displayValue(1, 0).text, 'المبلغ قرأ');
});

test('a document replaces every Arabic match, marks and all, and only matches as written when asked', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: `قال ${WITH_MARKS} ثم ${STRETCHED}.` }] }));
  assert.equal(view.replaceAll(BARE, 'قرأ', { matchDiacritics: true }), 0);
  assert.equal(view.replaceAll(BARE, 'قرأ'), 2);
  assert.equal(view.blocks[0].text, 'قال قرأ ثم قرأ.');
});

test('a deck finds and replaces Arabic in its shapes the same way', () => {
  const deck = Deck.open(buildPptx({ title: 'بحث', slides: [{ layout: 'title', title: `عنوان ${WITH_MARKS}`, body: 'أحمد' }] }));
  assert.equal(deck.find(BARE).length, 1);
  assert.equal(deck.find('احمد').length, 1);
  assert.equal(deck.find(BARE, { matchDiacritics: true }).length, 0);
  assert.equal(deck.replaceAll(BARE, 'قرأ'), 1);
  assert.equal(deck.find('قرأ').length, 1);
});

test('a number format names the Hijri calendar (B2, or the calendar byte of [$-…]) and a set of digits', () => {
  const day = dateToSerial(new Date(Date.UTC(2026, 9, 10)));
  assert.equal(formatValue(day, 'dd/mm/yyyy').text, '10/10/2026');
  assert.equal(formatValue(day, 'B1dd/mm/yyyy').text, '10/10/2026', 'B1 is Gregorian');
  assert.equal(formatValue(day, 'B2dd/mm/yyyy').text, '29/04/1448', 'B2 is Hijri, Umm al-Qura');
  assert.equal(formatValue(day, '[$-1060401]dd/mm/yyyy').text, '29/04/1448', 'the calendar byte 06 is Hijri too');
  assert.match(formatValue(day, '[$-60401]d mmmm yyyy').text, /^29 ربيع الآخر 1448$/, 'Arabic month names for an Arabic locale');
  assert.match(formatValue(day, 'B2dddd d mmmm yyyy').text, /^Saturday 29 \S+ II 1448$/, 'the weekday is the same day\'s');
  assert.equal(formatValue(1234.5, '[$-2000000]#,##0.00').text, '١,٢٣٤.٥٠', 'Arabic-Indic digits');
  assert.equal(formatValue(25, '[$-3000000]0').text, '۲۵', 'Urdu\'s and Persian\'s digits');
  assert.equal(formatValue(day, '[$-2000000]dd/mm/yyyy').text, '١٠/١٠/٢٠٢٦');
  assert.equal(formatValue(12, '[$-409]0').text, '12', 'a locale alone changes nothing');
  assert.equal(formatValue(5, '[$€-407]0.00 €').text, '5.00 €', 'a currency with its locale still reads as before');
});
