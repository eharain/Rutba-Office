// The keyboard while Documents reads aloud (read-aloud/keys.js): the sentence
// being read is the selection, and a key must not put itself in its place.

import test from 'node:test';
import assert from 'node:assert/strict';
import { inputWhileReading, movesCaret } from '../apps/desktop/renderer/read-aloud/keys.js';

// Every way the keyboard, the mouse or a menu changes the words of a page.
const EDITS = [
  ['insertText', 'x'], ['insertCompositionText', 'ا'], ['insertParagraph', null], ['insertLineBreak', null],
  ['deleteContentBackward', null], ['deleteContentForward', null], ['deleteByCut', null],
  ['insertFromPaste', null], ['insertFromDrop', null], ['insertReplacementText', 'word'],
  ['formatBold', null], ['formatItalic', null], ['formatUnderline', null],
];

test('with Read Aloud off, every edit is made as ever', () => {
  for (const [inputType, data] of [...EDITS, ['insertText', ' ']]) {
    assert.equal(inputWhileReading({ reading: false, own: false, inputType, data }), 'edit', inputType);
    // A mark left behind by a reading that ended changes nothing.
    assert.equal(inputWhileReading({ reading: false, own: true, inputType, data }), 'edit', inputType);
  }
});

test('while the selection is the sentence being read, no key changes the words', () => {
  for (const [inputType, data] of EDITS) {
    assert.equal(inputWhileReading({ reading: true, own: true, inputType, data }), 'keep', inputType);
  }
});

test('Space pauses and reads on, and only a typed space does', () => {
  assert.equal(inputWhileReading({ reading: true, own: true, inputType: 'insertText', data: ' ' }), 'pause');
  assert.equal(inputWhileReading({ reading: true, own: true, inputType: 'insertText', data: '  ' }), 'keep');
  assert.equal(inputWhileReading({ reading: true, own: true, inputType: 'insertFromPaste', data: ' ' }), 'keep', 'a pasted space is a paste');
});

test('once the person has put the caret somewhere, an edit stops the reading and is made, a space among them', () => {
  for (const [inputType, data] of [...EDITS, ['insertText', ' ']]) {
    assert.equal(inputWhileReading({ reading: true, own: false, inputType, data }), 'stop', inputType);
  }
});

test('the keys that move the caret, and no key that types', () => {
  for (const key of ['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End', 'PageUp', 'PageDown']) assert.ok(movesCaret(key), key);
  for (const key of ['a', ' ', 'Enter', 'Backspace', 'Delete', 'Tab', 'Escape', 'Shift', 'Control']) assert.ok(!movesCaret(key), key);
});
