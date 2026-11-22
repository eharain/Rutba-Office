// Read Aloud when nothing on the computer reads the language (reader.js):
// the system's voice still tries, and the reader's state names the language,
// so the bar can say why the words sound wrong or are not heard.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import esbuild from 'esbuild';

// The reader takes its words from the interface package, which is written for the bundler: bundled here as the windows' code is.
const here = path.dirname(fileURLToPath(import.meta.url));
const out = await esbuild.build({
  stdin: { contents: "export { createReader } from './apps/desktop/renderer/read-aloud/reader.js';", resolveDir: path.join(here, '..'), loader: 'js' },
  bundle: true,
  format: 'esm',
  platform: 'node',
  write: false,
  logLevel: 'silent',
  loader: { '.js': 'jsx', '.css': 'empty' },
  jsx: 'automatic',
  external: ['react', 'react-dom', 'react/jsx-runtime'],
});
const bundled = path.join(here, '..', 'node_modules', `.read-aloud-unread-${process.pid}.mjs`);
fs.writeFileSync(bundled, out.outputFiles[0].text);
let createReader;
try { ({ createReader } = await import(`file://${bundled.replace(/\\/g, '/')}`)); } finally { fs.rmSync(bundled, { force: true }); }

// The system's speech as the reader uses it: its voices, what it was asked to say, and its one event.
function fakeSynth(voices) {
  const heard = [];
  const listeners = [];
  return {
    heard,
    voices,
    getVoices: () => voices.slice(),
    speak: (u) => heard.push(u),
    cancel: () => {},
    pause: () => {},
    resume: () => {},
    addEventListener: (name, fn) => { if (name === 'voiceschanged') listeners.push(fn); },
    arrive(list) { voices.push(...list); for (const fn of listeners) fn(); },
  };
}

const shell = { store: { get: async ({ fallback }) => fallback, set: async () => {} } };

function setUp(voices) {
  globalThis.SpeechSynthesisUtterance = class { constructor(text) { this.text = text; } };
  const synth = fakeSynth(voices);
  const states = [];
  const reader = createReader({ shell, synth, onState: (s) => states.push(s) });
  return { synth, states, reader, last: () => states[states.length - 1] };
}

test('a language no voice reads is named in the state, and the words are still handed to the system', async () => {
  const { synth, reader, last } = setUp([{ name: 'David', lang: 'en-US' }]);
  await reader.read([{ text: 'یہ اردو ہے۔', language: 'ur-PK', source: {} }]);
  assert.equal(last().unread, 'ur-PK');
  assert.equal(last().playing, true);
  assert.equal(synth.heard.length, 1, 'the system is still asked to read it');
  assert.equal(synth.heard[0].lang, 'ur-PK');
  reader.stop();
});

test('a language the system reads, and words with no language, name nothing', async () => {
  const { reader, last } = setUp([{ name: 'David', lang: 'en-US' }]);
  await reader.read([{ text: 'This is English.', language: 'en-GB', source: {} }]);
  assert.equal(last().unread, null);
  await reader.read([{ text: 'No language is said.', language: null, source: {} }]);
  assert.equal(last().unread, null);
  reader.stop();
});

test('each sentence is judged by its own language', async () => {
  const { reader, last } = setUp([{ name: 'David', lang: 'en-US' }]);
  await reader.read([{ text: 'English first.', language: 'en-US', source: {} }, { text: 'پھر اردو۔', language: 'ur', source: {} }]);
  assert.equal(last().unread, null);
  reader.forward();
  assert.equal(last().unread, 'ur');
  reader.back();
  assert.equal(last().unread, null);
  reader.stop();
});

test('voices the system lists a moment late are waited for: nothing is said before, and the state is told when they come', async () => {
  const { synth, reader, states, last } = setUp([]);
  await reader.read([{ text: 'یہ اردو ہے۔', language: 'ur-PK', source: {} }]);
  assert.equal(last().unread, null, 'an empty list says nothing about what the computer reads');
  const told = states.length;
  synth.arrive([{ name: 'David', lang: 'en-US' }]);
  assert.equal(states.length, told + 1);
  assert.equal(last().unread, 'ur-PK');
  reader.stop();
});

test('the reader opens Voices and languages in a window of its own', () => {
  const made = [];
  globalThis.SpeechSynthesisUtterance = class {};
  const reader = createReader({ shell: { ...shell, win: { create: async (o) => { made.push(o); } } }, synth: fakeSynth([]) });
  reader.voicesWindow();
  assert.deepEqual(made, [{ app: 'home', query: { voices: 1 } }]);
  // A window with no way to open another does nothing, and does not throw.
  createReader({ shell, synth: fakeSynth([]) }).voicesWindow();
});
