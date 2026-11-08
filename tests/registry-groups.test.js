// The two sets, Office and Studio (D-2026-11-05-17): every app in exactly one,
// the names and order the same wherever the sets show, and which apps are on
// until chosen left as it was by the move.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { APPS, APP_GROUPS, APP_ORDER, OFF_UNTIL_CHOSEN, chosenApps } from '../packages/office-formats/src/registry.js';
import { menuAppSets } from '../packages/office-shell/src/menu-sets.js';
import { WORDS } from '../tools/make-installer-words.mjs';
import { installerScript } from '../apps/desktop/build/make-app-packs.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OFFICE = ['mail', 'calendar', 'contacts', 'talk', 'word', 'sheets', 'slides', 'pdf', 'base'];
const STUDIO = ['pictures', 'image', 'video', 'recorder', 'broadcast'];

test('the sets are Office and Studio, in that order, with Talk and Data Files in Office and Pictures in Studio', () => {
  assert.deepEqual(APP_GROUPS.map((g) => g.id), ['office', 'studio']);
  assert.deepEqual(APP_GROUPS[0].apps, OFFICE);
  assert.deepEqual(APP_GROUPS[1].apps, STUDIO);
});

test('every app is in exactly one set, and the launcher\'s order is the sets\' order', () => {
  assert.deepEqual([...APP_ORDER].sort(), Object.keys(APPS).sort());
  assert.equal(new Set(APP_ORDER).size, APP_ORDER.length);
  assert.deepEqual(APP_ORDER, [...OFFICE, ...STUDIO]);
});

test('which apps are on until chosen does not change with the sets', () => {
  assert.deepEqual([...OFF_UNTIL_CHOSEN].sort(), ['base', 'broadcast', 'image', 'recorder', 'talk', 'video']);
  assert.deepEqual(chosenApps(null), ['mail', 'calendar', 'contacts', 'word', 'sheets', 'slides', 'pdf', 'pictures']);
});

test('the macOS menu lists the two sets, only what is installed, and draws no line for a set with nothing installed', () => {
  assert.deepEqual(menuAppSets(APP_GROUPS), [OFFICE, STUDIO]);
  assert.deepEqual(menuAppSets(APP_GROUPS, ['video', 'mail', 'pictures']), [['mail'], ['pictures', 'video']]);
  assert.deepEqual(menuAppSets(APP_GROUPS, ['mail']), [['mail']]);
});

const lines = (text, re) => text.split(/\r?\n/).filter((l) => re.test(l));

test('the installer\'s Choose apps page has the two names, each set\'s boxes in order, inside the page\'s height', () => {
  const page = fs.readFileSync(path.join(ROOT, 'apps', 'desktop', 'resources', 'installer.nsh'), 'utf8');
  assert.ok(WORDS.some(([v, en]) => v === 'office' && en === 'Office'));
  assert.ok(WORDS.some(([v, en]) => v === 'studio' && en === 'Studio'));
  assert.ok(!WORDS.some(([v]) => v === 'more'));
  assert.ok(!page.includes('$rutbaW_more'));
  const boxes = lines(page, /!insertmacro rutbaAppBox \w+ /).map((l) => l.trim().split(/\s+/));
  assert.deepEqual(boxes.filter((b) => b[4] === '0').map((b) => b[2]), OFFICE);
  assert.deepEqual(boxes.filter((b) => b[4] === '52%').map((b) => b[2]), STUDIO);
  for (const b of boxes) assert.ok(parseInt(b[5], 10) + 12 <= 140, `${b[2]} is inside the page`);
  assert.ok(page.indexOf('rutbaW_office') < page.indexOf('rutbaW_studio'));
});

test('the installer\'s Apps to install page has the two names and the add-ons, each set\'s boxes in order, inside the page\'s height', () => {
  const packs = [...STUDIO, ...OFFICE, 'ocr', 'ai'].map((id) => ({ id, name: id, kinds: [], files: [], extensions: [] }));
  const text = installerScript({ dir: 'x', packs });
  const from = text.indexOf('!macro rutbaPackBoxes');
  const macro = text.slice(from, text.indexOf('!macroend', from));
  const boxes = lines(macro, /!insertmacro rutbaPackBox /).map((l) => l.trim().split(/\s+/));
  assert.deepEqual(boxes.filter((b) => b[4] === '0').map((b) => b[2]), OFFICE);
  assert.deepEqual(boxes.filter((b) => b[4] === '52%').map((b) => b[2]), [...STUDIO, 'ocr', 'ai']);
  for (const b of boxes) assert.ok(parseInt(b[5], 10) + 12 <= 140, `${b[2]} is inside the page`);
  for (const word of ['office', 'studio', 'addons']) assert.ok(macro.includes(`"$rutbaW_${word}"`), word);
});
