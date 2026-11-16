// The default voice (2.26.1): which voice Read Aloud takes for a language,
// and the order the spoken words offer the added voices in.
import test from 'node:test';
import assert from 'node:assert/strict';
import { voiceFor, reads, defaultFirst, DEFAULT_VOICE_KEY } from '../apps/desktop/renderer/read-aloud/voice-choice.js';

const added = [
  { id: 'voice-us', language: 'en-US' },
  { id: 'voice-gb', language: 'en-GB', also: ['en-US'] },
  { id: 'voice-hi', language: 'hi-IN', also: ['ur'] },
];

test('the default voice is kept under one key, an added voice\'s id', () => {
  assert.equal(DEFAULT_VOICE_KEY, 'speech.defaultVoice');
});

test('with no voice chosen for a language, the default voice reads it when it is added and reads that language', () => {
  assert.deepEqual(voiceFor({ defaultId: 'voice-gb', added, language: 'en-US' }), { kind: 'added', id: 'voice-gb' });
  assert.deepEqual(voiceFor({ defaultId: 'voice-gb', added, language: 'en' }), { kind: 'added', id: 'voice-gb' });
  // Text with no language said: the default reads it.
  assert.deepEqual(voiceFor({ defaultId: 'voice-gb', added, language: null }), { kind: 'added', id: 'voice-gb' });
  // Another language: the system's own, never an English voice over Urdu.
  assert.equal(voiceFor({ defaultId: 'voice-gb', added, language: 'ur' }), null);
  // A voice that also reads a language is its default too.
  assert.deepEqual(voiceFor({ defaultId: 'voice-hi', added, language: 'ur-PK' }), { kind: 'added', id: 'voice-hi' });
});

test('no default, a default that is not added, or an engine that stopped: the system\'s own voice', () => {
  assert.equal(voiceFor({ defaultId: '', added, language: 'en-US' }), null);
  assert.equal(voiceFor({ defaultId: 'voice-gone', added, language: 'en-US' }), null);
  assert.equal(voiceFor({ defaultId: 'voice-gb', added: [], language: 'en-US' }), null);
  assert.equal(voiceFor({ defaultId: 'voice-gb', added, language: 'en-US', engineFailed: true }), null);
});

test('a voice chosen in the bar stands over the default, and so does the system\'s own when it was chosen', () => {
  const pick = { kind: 'added', id: 'voice-us' };
  assert.deepEqual(voiceFor({ chosen: { 'en-us': pick }, defaultId: 'voice-gb', added, language: 'en-US' }), pick);
  // Chosen for the language's base, read for a region of it.
  assert.deepEqual(voiceFor({ chosen: { en: pick }, defaultId: 'voice-gb', added, language: 'en-GB' }), pick);
  // System voice chosen in the bar: kept as a choice, not passed over for the default.
  assert.deepEqual(voiceFor({ chosen: { 'en-us': { kind: 'system' } }, defaultId: 'voice-gb', added, language: 'en-US' }), { kind: 'system' });
  assert.deepEqual(voiceFor({ chosen: { 'en-us': { kind: 'system', name: 'Microsoft Zira' } }, defaultId: 'voice-gb', added, language: 'en-US' }), { kind: 'system', name: 'Microsoft Zira' });
  // What 2.26.0 kept for the system's own (null) is no choice: the default reads.
  assert.deepEqual(voiceFor({ chosen: { 'en-us': null }, defaultId: 'voice-gb', added, language: 'en-US' }), { kind: 'added', id: 'voice-gb' });
  // An added voice chosen while the engine has stopped: the system's own.
  assert.equal(voiceFor({ chosen: { 'en-us': pick }, defaultId: 'voice-gb', added, language: 'en-US', engineFailed: true }), null);
});

test('a voice reads its own language, those it also reads, and anything when no language is said', () => {
  assert.equal(reads(added[1], 'en-US'), true);
  assert.equal(reads(added[1], 'fr'), false);
  assert.equal(reads(added[2], 'ur'), true);
  assert.equal(reads(added[0], null), true);
});

test('the spoken words offer the default voice first, and the rest as they came', () => {
  assert.deepEqual(defaultFirst(added, 'voice-hi').map((v) => v.id), ['voice-hi', 'voice-us', 'voice-gb']);
  assert.deepEqual(defaultFirst(added, '').map((v) => v.id), ['voice-us', 'voice-gb', 'voice-hi']);
  assert.deepEqual(defaultFirst(added, 'voice-gone').map((v) => v.id), ['voice-us', 'voice-gb', 'voice-hi']);
});
