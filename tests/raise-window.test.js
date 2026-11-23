// Which window comes forward when Office is started again while it runs
// (office-shell raise-window.js): never one that is not shown to anyone.

import test from 'node:test';
import assert from 'node:assert/strict';
import { windowToRaise } from '../packages/office-shell/src/electron/raise-window.js';

const win = (name, { focusable = true, destroyed = false } = {}) => ({ name, isDestroyed: () => destroyed, isFocusable: () => focusable });

test('the first window a person can be shown comes forward', () => {
  const word = win('word');
  const talk = win('talk');
  assert.equal(windowToRaise([word, talk]), word);
  assert.equal(windowToRaise([win('gone', { destroyed: true }), talk]), talk);
});

test('a window that is never shown is not brought forward: with only the ring for a call alive, there is none, and Home opens', () => {
  const ring = win('ring', { focusable: false });
  assert.equal(windowToRaise([ring]), null);
  const word = win('word');
  assert.equal(windowToRaise([ring, word]), word, 'a real window behind it is the one');
  assert.equal(windowToRaise([]), null);
  assert.equal(windowToRaise(), null);
});

test('a window that fails as it is asked is passed over', () => {
  const broken = { isDestroyed: () => { throw new Error('Object has been destroyed'); }, isFocusable: () => true };
  const word = win('word');
  assert.equal(windowToRaise([broken, word]), word);
  assert.equal(windowToRaise([null, undefined, broken]), null);
});
