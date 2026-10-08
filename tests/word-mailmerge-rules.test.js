// Mail merge Rules → Ask, Fill-in and Set Bookmark: the fields Word writes
// (ASK, FILLIN, SET, and a REF to show a bookmark), the questions a merge
// asks before it runs, and each copy taking its record's answer — or the one
// answer given once, or the default.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml/build';
import { parseDelimited, sourceFromRows, fillInInstr, askInstr, setInstr, refInstr, bookmarkName, readPrompt, evaluateField } from '@rutba/ooxml/mailmerge';

const CSV = 'First Name,City\r\nJoshua,London\r\nCynthia,Leeds\r\nAmira,Paris\r\n';
const source = () => sourceFromRows(parseDelimited(CSV), { kind: 'csv', name: 'list.csv', path: 'list.csv' });

test('the fields are written as Word writes them, a bookmark\'s name made one Word allows', () => {
  assert.equal(fillInInstr({ prompt: 'Your table number?', def: '1', once: true }), ' FILLIN "Your table number?" \\d "1" \\o ');
  assert.equal(askInstr({ name: 'Event', prompt: 'Which event?', def: 'the dinner' }), ' ASK Event "Which event?" \\d "the dinner" ');
  assert.equal(setInstr({ name: 'Venue', value: 'The "Old" Hall' }), ' SET Venue "The \\"Old\\" Hall" ');
  assert.equal(refInstr('Venue'), ' REF Venue ');
  assert.equal(bookmarkName('2nd guest list!'), 'nd_guest_list_');
  assert.throws(() => bookmarkName('42'), /starts with a letter/);
  assert.deepEqual(readPrompt(askInstr({ name: 'Event', prompt: 'Which event?', def: 'the dinner', once: true })), { kind: 'ask', name: 'Event', prompt: 'Which event?', def: 'the dinner', once: true });
  assert.deepEqual(readPrompt(' FILLIN "Seat?" '), { kind: 'fillin', name: null, prompt: 'Seat?', def: '', once: false });
});

test('a Fill-in takes the record\'s own answer, else the one given for all, else its default', () => {
  const instr = fillInInstr({ prompt: 'Seat?', def: 'any' });
  assert.equal(evaluateField(instr, { recordNumber: 2, answers: { 'fillin:Seat?': { 2: 'B4' } } }).text, 'B4');
  assert.equal(evaluateField(instr, { recordNumber: 3, answers: { 'fillin:Seat?': { 2: 'B4', all: 'C1' } } }).text, 'C1');
  assert.equal(evaluateField(instr, { recordNumber: 3, answers: { 'fillin:Seat?': 'D9' } }).text, 'D9');
  assert.equal(evaluateField(instr, { recordNumber: 1 }).text, 'any');
});

test('a merge asks its questions first, and each copy shows its answers — an Ask\'s and a Set\'s through the bookmark', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Dear ' }, { text: 'Seat: ' }, { text: 'See you at ' }] }));
  view.startMailMerge('formLetters');
  view.attachMergeSource(source());
  view.setSelection({ block: 0, offset: 5 });
  view.insertMergeField('First Name');
  view.setSelection({ block: 1, offset: 6 });
  view.insertMergeRule('fillin', { prompt: 'Seat?', def: 'any' });
  view.setSelection({ block: 2, offset: 11 });
  view.insertMergeRule('ask', { name: 'Event', prompt: 'Which event?', def: 'the dinner', once: true, show: true });
  view.insertText(' in ');
  view.insertMergeRule('set', { name: 'Venue', value: 'the Old Hall', show: true });
  assert.match(view.block(2).text, /«Ask Event».*«Set Venue»/, 'the main document shows what they keep');

  const { prompts, records } = view.mergePrompts();
  assert.deepEqual(prompts.map((p) => [p.key, p.prompt, p.def, p.once]), [['fillin:Seat?', 'Seat?', 'any', false], ['ask:Event', 'Which event?', 'the dinner', true]]);
  assert.deepEqual(records.map((r) => [r.number, r.label]), [[1, 'Joshua London'], [2, 'Cynthia Leeds'], [3, 'Amira Paris']]);

  const merged = openDocx(view.mergeToDocument({ answers: { 'fillin:Seat?': { 1: 'A1', 3: 'C3' }, 'ask:Event': 'the summer party' } }).bytes);
  const texts = merged.blocks.map((b) => b.text).filter((t) => t);
  assert.deepEqual(texts, [
    'Dear Joshua', 'Seat: A1', 'See you at the summer party in the Old Hall',
    'Dear Cynthia', 'Seat: any', 'See you at the summer party in the Old Hall',
    'Dear Amira', 'Seat: C3', 'See you at the summer party in the Old Hall',
  ]);
  // Without answers: the defaults.
  const plain = openDocx(view.mergeToDocument({ range: 'current' }).bytes).blocks.map((b) => b.text).filter((t) => t);
  assert.deepEqual(plain, ['Dear Joshua', 'Seat: any', 'See you at the dinner in the Old Hall']);
});
