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

test('a second click to English (after save first, or from a second window) keeps the way back', () => {
  // Urdu chosen in the menu; first click: English stored, the way back to Urdu kept.
  const first = afterClick(switchState({ showing: 'ur' }), { setting: 'ur' });
  assert.deepEqual(first.before, { setting: 'ur', shown: 'ur' });
  // The window had unsaved work, so it is still in Urdu; saved, the same click again.
  const second = afterClick(switchState({ showing: 'ur' }), { setting: first.setting, before: first.before });
  assert.deepEqual(second, { setting: 'en', before: { setting: 'ur', shown: 'ur' }, lang: 'en' });
  // And the way back still goes to Urdu.
  assert.deepEqual(afterClick(switchState({ showing: 'en', before: second.before })), { setting: 'ur', before: null, lang: 'ur' });
  // The system language kept as the way back is kept too.
  const system = afterClick(switchState({ showing: 'ar' }), { setting: 'en', before: { setting: null, shown: 'ar' } });
  assert.deepEqual(system.before, { setting: null, shown: 'ar' });
  // English stored with no way back left: back to the language this window shows.
  assert.deepEqual(afterClick(switchState({ showing: 'hi' }), { setting: 'en-GB', before: null }).before, { setting: 'hi', shown: 'hi' });
  assert.deepEqual(afterClick(switchState({ showing: 'hi' }), { setting: 'en', before: { setting: 'en', shown: 'en' } }).before, { setting: 'hi', shown: 'hi' });
});

test('the button reads the way back at the click, not only when the window opened', () => {
  const button = fs.readFileSync(path.join(ROOT, 'apps', 'desktop', 'renderer', 'english-switch.js'), 'utf8');
  assert.match(button, /const stored = await shell\.store\.get\(\{ key: BEFORE_ENGLISH_KEY[^]*afterClick\(state, \{ setting, before: stored \}\)/);
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
