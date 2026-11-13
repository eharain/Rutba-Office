// The English switch (L14, D-2026-11-13-30; renderer/english-switch-rule.js):
// shown while a language other than English is, one click to English and
// the same click back, its words in both languages, and in every window.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { switchState, bothLanguages, afterClick, BEFORE_ENGLISH_KEY } from '../apps/desktop/renderer/english-switch-rule.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

test('shown in any language but English, and in English only when switched from another', () => {
  assert.deepEqual(switchState({ showing: 'ur' }), { to: 'en', from: 'ur' });
  assert.deepEqual(switchState({ showing: 'fr_CA' }), { to: 'en', from: 'fr-ca' });
  assert.equal(switchState({ showing: 'en' }), null, 'English chosen: nothing to switch back to');
  assert.equal(switchState({ showing: 'en-GB', before: { setting: 'en-us', shown: 'en-us' } }), null);
  assert.deepEqual(switchState({ showing: 'en', before: { setting: 'ur', shown: 'ur' } }), { to: 'ur', setting: 'ur', from: 'en' });
  assert.deepEqual(switchState({ showing: 'en', before: { setting: null, shown: 'ar' } }), { to: 'ar', setting: null, from: 'en' });
  assert.equal(switchState({ showing: 'en', before: 'nonsense' }), null);
});

test('one click to English keeps where it came from; the same click goes back to it', () => {
  // The system's language (no setting) was Urdu: English now, the system's again after.
  const there = afterClick(switchState({ showing: 'ur' }), { setting: null });
  assert.deepEqual(there, { setting: 'en', before: { setting: null, shown: 'ur' }, lang: 'en' });
  const back = afterClick(switchState({ showing: 'en', before: there.before }));
  assert.deepEqual(back, { setting: null, before: null, lang: 'ur' });
  // Chosen in the menu: chosen again.
  const chosen = afterClick(switchState({ showing: 'hi' }), { setting: 'hi' });
  assert.deepEqual(afterClick(switchState({ showing: 'en', before: chosen.before })), { setting: 'hi', before: null, lang: 'hi' });
  assert.equal(afterClick(null), null);
  assert.equal(BEFORE_ENGLISH_KEY, 'language.beforeEnglish');
});

test('its words are in the window\'s language and in English, so either reader finds it', () => {
  assert.equal(bothLanguages('اس ونڈو کو انگریزی میں دکھائیں', 'Show this window in English'), 'اس ونڈو کو انگریزی میں دکھائیں · Show this window in English');
  assert.equal(bothLanguages('Show this window in English', 'Show this window in English'), 'Show this window in English');
});

test('every window has it, in the same place: the frame\'s title bar, a real button with its name', () => {
  const shell = fs.readFileSync(path.join(ROOT, 'apps', 'desktop', 'renderer', 'shell.js'), 'utf8');
  assert.match(shell, /right=\{<>\{right\}<EnglishSwitch shell=\{shell\} dirty=\{dirty\}/);
  const button = fs.readFileSync(path.join(ROOT, 'apps', 'desktop', 'renderer', 'english-switch.js'), 'utf8');
  assert.match(button, /<button type="button"[^>]*data-tip=\{words\} aria-label=\{words\}/);
  assert.match(button, /window\.location\.replace/, 'no restart: the window opens again in the language');
  assert.match(button, /if \(dirty\)/, 'never over changes not saved');
});
