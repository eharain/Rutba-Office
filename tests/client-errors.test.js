// What reaches a person when a request from a window fails: the sentence, not the plumbing.
import test from 'node:test';
import assert from 'node:assert/strict';
import { CHANNEL_NOISE } from '../packages/office-shell/src/client.js';

const shown = (m) => m.replace(CHANNEL_NOISE, '');

test('the channel and the name of the error class do not reach the person', () => {
  assert.equal(shown("Error invoking remote method 'rutba-office/mail:importScan': NdbError: That is no good."), 'That is no good.');
  assert.equal(shown("Error invoking remote method 'x': Error: Plain."), 'Plain.');
  assert.equal(shown("Error invoking remote method 'x': ShellError: Named."), 'Named.');
});

test('a sentence of its own is left alone', () => {
  assert.equal(shown('Error: not from a channel'), 'Error: not from a channel');
  assert.equal(shown("Error invoking remote method 'x': The server said no."), 'The server said no.');
});
