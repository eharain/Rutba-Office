// Find in Sheets, as Excel's: Match case, Match entire cell contents and
// Within Sheet or Workbook; Find Next and Find Previous with "3 of 12";
// Find All's list; Replace and Replace All on one sheet or every sheet, as
// one undo step. Then what the window reads back, the keys that open Find
// and Replace, and the words, in every language the suite ships.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { SheetView } from '@rutba/sheet-view';
import { buildXlsx } from '@rutba/ooxml/build';
import { createDocumentService } from '../apps/desktop/main/documents.js';
import { readFindResult, readFindList, findCountText, findOptions, findKey, goToMatch } from '../apps/desktop/renderer/apps/sheets/find-model.js';
import { SHEETS_WORDS } from '../apps/desktop/renderer/apps/sheets-translations.js';
import ur from '../packages/office-ui/src/catalogues/ur.js';
import ar from '../packages/office-ui/src/catalogues/ar.js';
import hi from '../packages/office-ui/src/catalogues/hi.js';

const book = () => SheetView.open(buildXlsx({
  sheets: [
    { name: 'Stock', rows: [['Item', 'Note'], ['Tea', 'Green tea'], ['Apple', 'apple pie'], ['tea', 'Tea']] },
    { name: 'Orders', rows: [['Tea', 2], ['Milk', 'tea later']] },
    { name: 'Spare', rows: [['Tea']] },
  ],
}));
const at = (v) => `${v.activeSheet}!${v.selection.active.row}:${v.selection.active.col}`;

test('Find Next walks the matches by rows, wrapping, and says which of how many it is on', () => {
  const v = book();
  const r1 = v.findStep('tea');
  assert.deepEqual({ found: r1.found, index: r1.index, total: r1.total, ref: r1.ref }, { found: true, index: 1, total: 4, ref: 'A2' });
  assert.equal(v.findStep('tea').ref, 'B2');
  assert.equal(v.findStep('tea').ref, 'A4');
  assert.equal(v.findStep('tea').ref, 'B4');
  const back = v.findStep('tea', { back: true });
  assert.deepEqual([back.ref, back.index], ['A4', 3], 'Find Previous goes back one');
  v.findStep('tea'); // B4, the last
  const wrapped = v.findStep('tea');
  assert.deepEqual([wrapped.ref, wrapped.index], ['A2', 1], 'past the last it starts again at the first');
  assert.equal(v.findStep('tea', { back: true }).ref, 'B4', 'and before the first it goes to the last');
  assert.equal(v.findNext('Apple'), true, 'findNext still says whether anything matched');
  assert.equal(at(v), 'Stock!2:0');
});

test('nothing found says so and leaves the selection where it was', () => {
  const v = book();
  v.select(2, 1);
  const r = v.findStep('coffee');
  assert.deepEqual({ found: r.found, index: r.index, total: r.total }, { found: false, index: 0, total: 0 });
  assert.equal(at(v), 'Stock!2:1');
  assert.equal(v.findNext(''), false);
});

test('Match case and Match entire cell contents narrow the search as Excel\'s do', () => {
  const v = book();
  assert.deepEqual(v.findAll('Tea', { matchCase: true }).map((m) => m.ref), ['A2', 'B4']);
  assert.deepEqual(v.findAll('tea', { wholeCell: true }).map((m) => m.ref), ['A2', 'A4', 'B4'], '"Green tea" is not the whole cell');
  assert.deepEqual(v.findAll('tea', { wholeCell: true, matchCase: true }).map((m) => m.ref), ['A4']);
  assert.deepEqual(v.findAll('apple').map((m) => m.ref), ['A3', 'B3']);
});

test('Within Workbook looks in every sheet showing, in tab order, and goes to the sheet a match is on', () => {
  const v = book();
  v.hideSheet('Spare');
  const all = v.findAll('tea', { within: 'workbook' });
  assert.deepEqual(all.map((m) => `${m.sheet}!${m.ref}`), ['Stock!A2', 'Stock!B2', 'Stock!A4', 'Stock!B4', 'Orders!A1', 'Orders!B2'], 'a hidden sheet is not searched');
  assert.equal(v.findAll('tea').length, 4, 'Within Sheet stays on this sheet');
  v.select(3, 1); // Stock!B4, the last on this sheet
  const r = v.findStep('tea', { within: 'workbook' });
  assert.deepEqual([r.sheet, r.ref, r.index, r.total], ['Orders', 'A1', 5, 6]);
  assert.equal(v.activeSheet, 'Orders');
  assert.equal(at(v), 'Orders!0:0');
  const back = v.findStep('tea', { within: 'workbook', back: true });
  assert.deepEqual([back.sheet, back.ref], ['Stock', 'B4'], 'and back again to the sheet before');
  assert.equal(v.findStep('tea', { within: 'workbook' }).sheet, 'Orders');
  assert.equal(v.findStep('tea', { within: 'workbook' }).ref, 'B2');
  const round = v.findStep('tea', { within: 'workbook' });
  assert.deepEqual([round.sheet, round.ref], ['Stock', 'A2'], 'past the last sheet it starts at the first');
});

test('Find All lists each match with its sheet, cell, what it shows and its formula', () => {
  const v = SheetView.open(buildXlsx({ sheets: [{ name: 'S', rows: [['Tea', 3], ['=A1&" cup"', 4]] }] }));
  const all = v.findAll('tea');
  assert.deepEqual(all.map(({ sheet, row, col, ref, value, formula }) => ({ sheet, row, col, ref, value, formula })), [
    { sheet: 'S', row: 0, col: 0, ref: 'A1', value: 'Tea', formula: null },
    { sheet: 'S', row: 1, col: 0, ref: 'A2', value: 'Tea cup', formula: '=A1&" cup"' },
  ]);
});

test('Replace and Replace All keep to the choices: case, the whole cell, this sheet', () => {
  const v = book();
  assert.equal(v.replaceAll('Tea', 'Coffee', { matchCase: true }), 2);
  assert.deepEqual([v.editValue(1, 0), v.editValue(1, 1), v.editValue(3, 0), v.editValue(3, 1)], ['Coffee', 'Green tea', 'tea', 'Coffee']);
  const w = book();
  assert.equal(w.replaceAll('tea', 'Chai', { wholeCell: true }), 3);
  assert.deepEqual([w.editValue(1, 0), w.editValue(1, 1), w.editValue(3, 0)], ['Chai', 'Green tea', 'Chai']);
  w.selectSheet('Orders');
  assert.equal(w.editValue(0, 0), 'Tea', 'another sheet is left alone Within Sheet');
  const x = book();
  x.select(0, 0);
  assert.equal(x.replaceNext('apple', 'pear', { matchCase: true }), true);
  assert.equal(x.editValue(2, 1), 'pear pie', 'Replace with Match case skips Apple');
  assert.equal(x.editValue(2, 0), 'Apple');
});

test('Replace All Within Workbook changes every sheet, and one undo puts every sheet back', () => {
  const v = book();
  const n = v.replaceAll('tea', 'Chai', { within: 'workbook' });
  assert.equal(n, 7);
  assert.equal(v.activeSheet, 'Stock', 'the sheet in view stays');
  assert.equal(v.editValue(1, 1), 'Green Chai');
  const on = (s, r, c) => { const was = v.activeSheet; v.selectSheet(s); const out = v.editValue(r, c); v.selectSheet(was); return out; };
  assert.deepEqual([on('Orders', 0, 0), on('Orders', 1, 1), on('Spare', 0, 0)], ['Chai', 'Chai later', 'Chai']);
  v.undo();
  assert.deepEqual([v.editValue(1, 1), on('Orders', 0, 0), on('Orders', 1, 1), on('Spare', 0, 0)], ['Green tea', 'Tea', 'tea later', 'Tea']);
  v.redo();
  assert.deepEqual([on('Orders', 0, 0), on('Spare', 0, 0)], ['Chai', 'Chai']);
});

test('Replace All Within Workbook refuses a locked cell on a protected sheet, and changes nothing', () => {
  const v = book();
  v.selectSheet('Orders');
  v.protect();
  v.selectSheet('Stock');
  assert.throws(() => v.replaceAll('tea', 'Chai', { within: 'workbook' }), /protected/);
  assert.equal(v.editValue(1, 0), 'Tea');
  assert.equal(v.replaceAll('tea', 'Chai'), 4, 'Within Sheet, this unprotected sheet still changes');
});

test('the window\'s operations: a step and Find All come back as their answers, and neither makes the file dirty', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-sheet-find-'));
  const doc = createDocumentService({ holdBlob: () => ({ url: 'blob://held' }) });
  try {
    const file = path.join(dir, 'find.xlsx');
    fs.writeFileSync(file, buildXlsx({ sheets: [{ name: 'A', rows: [['Tea'], ['Milk'], ['tea']] }, { name: 'B', rows: [['Tea']] }] }));
    const s = doc.open({ path: file });
    const step = doc.apply({ id: s.id, ops: [{ op: 'findStep', text: 'tea', within: 'workbook' }] });
    assert.equal(step.dirty, false);
    assert.deepEqual(readFindResult(step.opResult), { found: true, index: 2, total: 3, sheet: 'A', ref: 'A3' }, 'the next match after A1');
    const on = doc.apply({ id: s.id, ops: [{ op: 'findStep', text: 'tea', within: 'workbook' }] });
    assert.deepEqual(readFindResult(on.opResult), { found: true, index: 3, total: 3, sheet: 'B', ref: 'A1' }, 'onto the next sheet');
    const back = doc.apply({ id: s.id, ops: [{ op: 'findStep', text: 'tea', within: 'workbook', back: true }] });
    assert.deepEqual(readFindResult(back.opResult), { found: true, index: 2, total: 3, sheet: 'A', ref: 'A3' }, 'and back');
    const list = doc.apply({ id: s.id, ops: [{ op: 'findList', text: 'TEA', matchCase: true, within: 'workbook' }] });
    assert.equal(list.dirty, false);
    assert.deepEqual(readFindList(list.opResult), []);
    const listed = doc.apply({ id: s.id, ops: [{ op: 'findList', text: 'tea', wholeCell: true }] });
    assert.deepEqual(readFindList(listed.opResult).map((m) => `${m.sheet}!${m.ref}`), ['A!A1', 'A!A3']);
    const replaced = doc.apply({ id: s.id, ops: [{ op: 'replaceAll', find: 'tea', replace: 'Chai', within: 'workbook' }] });
    assert.equal(replaced.opResult, 3);
    assert.equal(replaced.dirty, true);
    doc.close({ id: s.id });
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('what the box and the panel read back, and the keys that open them', () => {
  assert.deepEqual(readFindResult('{"found":true,"index":3,"total":12,"sheet":"S","ref":"B4"}'), { found: true, index: 3, total: 12, sheet: 'S', ref: 'B4' });
  assert.equal(readFindResult('nope').found, false);
  assert.equal(readFindResult(undefined).total, 0);
  const say = (w, v) => w.replace(/\{(\w+)\}/g, (_, k) => String(v?.[k]));
  assert.equal(findCountText({ found: true, index: 3, total: 12 }, say), '3 of 12');
  assert.equal(findCountText({ found: false, index: 0, total: 0 }, say), 'No matches');
  assert.deepEqual(readFindList('[{"sheet":"S","row":1,"col":0,"ref":"A2"},{"bad":1}]').map((m) => m.ref), ['A2']);
  assert.deepEqual(readFindList('{}'), []);
  assert.deepEqual(findOptions({ matchCase: 1, within: 'elsewhere' }), { matchCase: true, wholeCell: false, within: 'sheet' });
  assert.equal(findKey({ ctrlKey: true, key: 'f' }), 'find');
  assert.equal(findKey({ metaKey: true, key: 'H' }), 'replace');
  assert.equal(findKey({ ctrlKey: true, shiftKey: true, key: 'f' }), null);
  assert.equal(findKey({ ctrlKey: true, altKey: true, key: 'h' }), null);
  assert.equal(findKey({ key: 'f' }), null);
  assert.deepEqual(goToMatch({ sheet: 'B', row: 2, col: 1 }, 'A'), [{ op: 'sheet', name: 'B' }, { op: 'select', row: 2, col: 1 }]);
  assert.deepEqual(goToMatch({ sheet: 'A', row: 2, col: 1 }, 'A'), [{ op: 'select', row: 2, col: 1 }]);
});

test('Sheets opens Find with Ctrl+F and Replace with Ctrl+H, Find & Select\'s Replace opens Replace, and no dialog covers the grid', () => {
  const read = (f) => fs.readFileSync(new URL(`../apps/desktop/renderer/apps/${f}`, import.meta.url), 'utf8');
  const sheets = read('sheets.js');
  assert.match(sheets, /'edit\.find': \{[^}]*key: 'Mod\+F'/);
  assert.match(sheets, /'edit\.replace': \{[^}]*key: 'Mod\+H'/);
  assert.match(sheets, /key === 'f' && !e\.shiftKey\) \{ e\.preventDefault\(\); focusFind\(\)/, 'the grid hands Ctrl+F to the Find box');
  assert.match(sheets, /key === 'h' && !e\.shiftKey\) \{ e\.preventDefault\(\); openFind\('replace'\)/);
  assert.match(sheets, /<FindBox ref=\{findBoxRef\}/);
  assert.doesNotMatch(sheets, /FindDialog/, 'the modal dialog is gone');
  assert.match(read('sheets/ribbon.js'), /label: t\('Replace…'\), icon: 'find', run: \(\) => openDialog\('replace'\)/);
  assert.match(read('sheets/find.js'), /aria-modal="false"/);
});

const SHARED = { ur, ar, hi };
const SCRIPTS = { ur: /[؀-ۿ]/, ar: /[؀-ۿ]/, hi: /[ऀ-ॿ]/ };
const unescape = (s) => s.replace(/\\(n|.)/g, (m, c) => (c === 'n' ? '\n' : c));
const places = (s) => [...String(s).matchAll(/\{(\w+)\}/g)].map((m) => m[1]).sort();
const wordsIn = (f) => {
  const text = fs.readFileSync(new URL(`../apps/desktop/renderer/apps/${f}`, import.meta.url), 'utf8');
  return { text, words: [...text.matchAll(/(?<![\w.$])t\(\s*'((?:[^'\\\n]|\\.)*)'/g)].map((m) => unescape(m[1])) };
};

test('every word Find shows is translated in every language the suite ships, with its places, in its own script', () => {
  const { text, words: own } = wordsIn('sheets/find.js');
  // find-model.js says its words through the t() find.js hands it, as say().
  const model = fs.readFileSync(new URL('../apps/desktop/renderer/apps/sheets/find-model.js', import.meta.url), 'utf8');
  const words = [...own, ...[...model.matchAll(/(?<![\w.$])say\(\s*'((?:[^'\\\n]|\\.)*)'/g)].map((m) => unescape(m[1]))];
  assert.ok(words.includes('{index} of {total}'), 'the count is read from find-model.js');
  assert.match(text, /from '\.\.\/sheets-words\.js'/, 'find.js takes its words from sheets-words.js');
  assert.ok(words.length >= 8, 'the words were read');
  assert.deepEqual(Object.keys(SHEETS_WORDS).sort(), Object.keys(SHARED).sort());
  for (const [tag, list] of Object.entries(SHEETS_WORDS)) {
    assert.deepEqual(words.filter((w) => !list[w] && !SHARED[tag][w]), [], `${tag}: not translated`);
    assert.deepEqual(Object.keys(list).filter((w) => !words.includes(w)), [], `${tag}: no longer used`);
    assert.deepEqual(Object.keys(list).filter((w) => SHARED[tag][w] !== undefined), [], `${tag}: already in the shared catalogue`);
    for (const [english, v] of Object.entries(list)) {
      assert.deepEqual(places(v), places(english), `${tag}: "${english}" keeps its places`);
      assert.match(v, SCRIPTS[tag], `${tag}: "${english}" is written in ${tag}`);
    }
  }
  const shared = wordsIn('sheets/find-names.js');
  assert.match(shared.text, /import \{ t \} from '@rutba\/office-ui'/);
  for (const [tag, list] of Object.entries(SHARED)) assert.deepEqual(shared.words.filter((w) => !list[w]), [], `${tag}: a shared word is missing`);
});
