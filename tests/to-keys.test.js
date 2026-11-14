// Moving calls from English to meaning keys (tools/to-keys.mjs, D-2026-11-14-16):
// t, msg and tn with literal English become their keys, a counted call loses
// its English plural, the quotes and spacing around stay as they were, a key
// already there and another function called t are left alone, and English
// with no definition is reported, not guessed.
import test from 'node:test';
import assert from 'node:assert/strict';
import { toKeys, usesCatalogue } from '../tools/to-keys.mjs';

const LABELS = {
  'common.close': { english: 'Close' },
  'common.saved-name': { english: 'Saved {name}' },
  'common.count-file': { english: '{count} file', plural: '{count} files' },
  'common.it-s': { english: "It's done" },
  'common.line': { english: 'One\nTwo' },
};

test('literal English in t, msg and tn becomes its key; keys, other calls and unknown English are left', () => {
  const src = [
    "import { t, tn, msg } from '@rutba/office-ui';",
    "const a = t('Close');",
    'const b = t("Saved {name}", { name });',
    "const c = tn(files.length, '{count} file', '{count} files', { extra });",
    "const d = tn(n,\n  '{count} file', '{count} files');",
    "const e = msg( 'Close' );",
    "const f = t('It\\'s done');",
    "const g = t('One\\nTwo');",
    "const h = t('common.close');",
    "const i = obj.t('Close');",
    "const j = format('Close');",
    "const k = t('Nobody defined this');",
    'const l = t(variable);',
  ].join('\n');
  const r = toKeys(src, LABELS);
  const lines = r.text.split('\n');
  assert.equal(lines[1], "const a = t('common.close');");
  assert.equal(lines[2], "const b = t('common.saved-name', { name });", 'the values stay');
  assert.equal(lines[3], "const c = tn(files.length, 'common.count-file', { extra });", 'the plural goes, the values stay');
  assert.equal(r.text.includes("const d = tn(n, 'common.count-file');"), true, 'across a line break');
  assert.ok(r.text.includes("const e = msg( 'common.close' );"), 'its spacing kept');
  assert.ok(r.text.includes("const f = t('common.it-s');"), 'an escaped quote');
  assert.ok(r.text.includes("const g = t('common.line');"), 'an escaped line break');
  assert.ok(r.text.includes("const h = t('common.close');"));
  assert.ok(r.text.includes("const i = obj.t('Close');"), 'a method called t is not the catalogue');
  assert.ok(r.text.includes("const j = format('Close');"));
  assert.ok(r.text.includes("const k = t('Nobody defined this');"), 'no key, no change');
  assert.deepEqual(r.unknown, ['Nobody defined this']);
  assert.equal(r.moved, 7);
  assert.deepEqual(toKeys(r.text, LABELS).moved, 0, 'running it again changes nothing');
});

test('only a file that takes t from the catalogue or a part is moved', () => {
  assert.ok(usesCatalogue("import { t } from '@rutba/office-ui';"));
  assert.ok(usesCatalogue("import { Button, t, tn } from '@rutba/office-ui';"));
  assert.ok(usesCatalogue("import { t } from './words.js';"));
  assert.ok(usesCatalogue("import { t } from '../locker-words.js';"));
  assert.ok(usesCatalogue("import { t, tn } from '../talk-words.js';"));
  assert.ok(!usesCatalogue("import { t } from './i18n.js';"));
  assert.ok(!usesCatalogue("const t = (s) => s;"));
});
