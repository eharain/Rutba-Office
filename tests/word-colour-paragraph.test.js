// Documents: a Font colour or a highlight chosen with nothing selected
// colours the paragraph the caret is in (the owner's rule, D-2026-11-11-11),
// as one undo step, and what is typed there next takes it. Before, the
// choice waited for typing alone, and a click anywhere dropped it, so a
// person who picked a colour and then clicked into the text saw nothing.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';

const doc = (paragraphs) => openDocx(buildDocx({ styles: true, paragraphs }));
const at = (block, offset) => ({ block, offset });
const props = (view, block) => view.blocks[block].runs.filter((r) => r.text).map((r) => [r.text, view.doc.readRunProps(r.rPr)]);
const colours = (view, block) => props(view, block).map(([text, p]) => [text, p.fontColour ?? null]);
const highlights = (view, block) => props(view, block).map(([text, p]) => [text, p.highlight ?? null]);

test('a Font colour with nothing selected colours the whole paragraph the caret is in, and no other', () => {
  const view = doc([{ text: 'Before' }, { text: 'The quick brown fox' }, { text: 'After' }]);
  view.setSelection(at(1, 6), at(1, 6));
  view.setRunFormat({ fontColour: 'FF0000' });
  assert.deepEqual(colours(view, 1), [['The quick brown fox', 'FF0000']]);
  assert.deepEqual(colours(view, 0), [['Before', null]], 'the paragraph above is left');
  assert.deepEqual(colours(view, 2), [['After', null]], 'the paragraph below is left');
  assert.ok(view.collapsed, 'nothing is selected after');
  assert.deepEqual({ ...view.focus }, at(1, 6), 'the caret stays where it was');
});

test('every run of the paragraph takes it, one already coloured too', () => {
  const view = doc([{ runs: [{ text: 'Mixed ' }, { text: 'blue part', colour: '2B5FD9' }, { text: ' tail', bold: true }] }]);
  view.setSelection(at(0, 2), at(0, 2));
  view.setRunFormat({ fontColour: '00B050' });
  assert.ok(colours(view, 0).every(([, c]) => c === '00B050'), JSON.stringify(colours(view, 0)));
  const tail = view.blocks[0].runs.find((r) => r.text.includes('tail'));
  assert.match(tail.rPr || '', /<w:b\/>/, 'the bold stays bold');
});

test('one Undo takes the paragraph\'s colour back', () => {
  const view = doc([{ text: 'One two three' }]);
  view.setSelection(at(0, 3), at(0, 3));
  view.setRunFormat({ fontColour: 'FF0000' });
  assert.deepEqual(colours(view, 0), [['One two three', 'FF0000']]);
  assert.equal(view.undo(), true);
  assert.deepEqual(colours(view, 0), [['One two three', null]]);
  assert.equal(view.canUndo, false, 'it was one step');
});

test('words typed next, at the caret, take the colour', () => {
  const view = doc([{ text: 'Second line here' }]);
  view.setSelection(at(0, 16), at(0, 16));
  view.setRunFormat({ fontColour: 'FF0000' });
  view.insertText(' typed');
  assert.deepEqual(colours(view, 0), [['Second line here typed', 'FF0000']]);
});

test('words typed after a click elsewhere in the coloured paragraph take it from the words beside them', () => {
  const view = doc([{ text: 'Second line here' }]);
  view.setSelection(at(0, 16), at(0, 16));
  view.setRunFormat({ fontColour: 'FF0000' });
  view.setSelection(at(0, 6), at(0, 6));
  view.insertText('!');
  assert.ok(colours(view, 0).every(([, c]) => c === 'FF0000'), JSON.stringify(colours(view, 0)));
});

test('in an empty paragraph nothing is coloured yet, and what is typed there takes it', () => {
  const view = doc([{ text: 'Above' }, { text: '' }]);
  view.setSelection(at(1, 0), at(1, 0));
  view.setRunFormat({ fontColour: 'C00000' });
  assert.deepEqual(colours(view, 0), [['Above', null]]);
  view.insertText('New words');
  assert.deepEqual(colours(view, 1), [['New words', 'C00000']]);
});

test('a highlight with nothing selected highlights the whole paragraph the same way, as one undo step', () => {
  const view = doc([{ text: 'First' }, { text: 'Highlight me whole' }]);
  view.setSelection(at(1, 4), at(1, 4));
  view.setRunFormat({ highlight: 'yellow' });
  assert.deepEqual(highlights(view, 1), [['Highlight me whole', 'yellow']]);
  assert.deepEqual(highlights(view, 0), [['First', null]]);
  view.undo();
  assert.deepEqual(highlights(view, 1), [['Highlight me whole', null]]);
});

test('Automatic and No highlight with nothing selected take the paragraph\'s colour and highlight off', () => {
  const view = doc([{ runs: [{ text: 'Red words', colour: 'FF0000' }] }]);
  view.setSelection(at(0, 2), at(0, 2));
  view.setRunFormat({ highlight: 'yellow' });
  view.setRunFormat({ fontColour: null });
  view.setRunFormat({ highlight: null });
  assert.deepEqual(colours(view, 0), [['Red words', null]]);
  assert.deepEqual(highlights(view, 0), [['Red words', null]]);
});

test('in a table cell only the cell\'s own paragraph is coloured', () => {
  const view = doc([{ text: 'Above' }, { table: { rows: [['cell words', 'other cell']] } }, { text: 'Below' }]);
  const cell = view.blocks.findIndex((b) => b.text === 'cell words');
  const other = view.blocks.findIndex((b) => b.text === 'other cell');
  assert.ok(cell > 0 && other > cell, 'the cells are paragraphs of their own');
  view.setSelection(at(cell, 2), at(cell, 2));
  view.setRunFormat({ fontColour: 'FF0000' });
  assert.deepEqual(colours(view, cell), [['cell words', 'FF0000']]);
  assert.deepEqual(colours(view, other), [['other cell', null]]);
});

test('a font or a size with nothing selected still waits for typing: the paragraph is left', () => {
  const view = doc([{ text: 'Plain words' }]);
  view.setSelection(at(0, 5), at(0, 5));
  view.setRunFormat({ fontSize: 20 });
  view.setRunFormat({ fontName: 'Arial' });
  assert.deepEqual(props(view, 0).map(([, p]) => [p.fontSize ?? null, p.fontName ?? null]), [[null, null]]);
  view.insertText('X');
  const typed = props(view, 0).find(([t]) => t === 'X');
  assert.equal(typed?.[1].fontSize, 20, JSON.stringify(props(view, 0)));
});

test('with words selected nothing changes: only the selection is coloured', () => {
  const view = doc([{ text: 'The quick brown fox' }]);
  view.setSelection(at(0, 4), at(0, 9));
  view.setRunFormat({ fontColour: 'FF0000' });
  assert.deepEqual(colours(view, 0), [['The ', null], ['quick', 'FF0000'], [' brown fox', null]]);
});
