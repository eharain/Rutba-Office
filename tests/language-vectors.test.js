// The language cases (L9): small packs as an add-ons folder holds them, and
// the words a window must get from each through language-packs.js. Made by
// tools/make-addon-vectors.mjs, from nothing outside this repository.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createLanguageWords } from '../packages/office-shell/src/electron/language-packs.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const DIR = path.join(ROOT, 'packages', 'office-shell', 'test-vectors', 'addons', 'languages');
const { office, keys, cases } = JSON.parse(fs.readFileSync(path.join(DIR, 'cases.json'), 'utf8'));
const contract = JSON.parse(fs.readFileSync(path.join(ROOT, 'packages', 'office-shell', 'src', 'addons-contract.json'), 'utf8'));
const manifestOf = (name) => JSON.parse(fs.readFileSync(path.join(DIR, name, 'manifest.json'), 'utf8'));

test('the language cases were made for this contract\'s Office', () => {
  assert.equal(office, contract.office, 'run node tools/make-addon-vectors.mjs');
  assert.ok(cases.length >= 13);
});

/** Office's English copy, laid out as resources/languages is: the envelope beside lang-en/. */
function copyFrom(name, root) {
  const dir = path.join(root, 'languages');
  fs.mkdirSync(path.join(dir, 'lang-en'), { recursive: true });
  fs.copyFileSync(path.join(DIR, name, 'messages.json'), path.join(dir, 'lang-en', 'messages.json'));
  fs.writeFileSync(path.join(dir, 'lang-en.pack.json'), JSON.stringify({ format: 'rutba-office-pack', keyId: 'x', manifest: manifestOf(name), signature: 'x' }));
  return dir;
}

for (const c of cases) {
  test(`language case: ${c.name}`, () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-langvec-'));
    try {
      const installed = c.installed.map((name) => ({ id: manifestOf(name).id, kind: 'language', dir: path.join(DIR, name), manifest: manifestOf(name) }));
      const words = createLanguageWords({ keys, copyDir: c.copy ? copyFrom(c.copy, root) : null, addons: () => installed });
      const got = words.wordsFor(c.ask);
      assert.equal(got.tag, c.expect.tag);
      assert.deepEqual(got.table, c.expect.table);
      if ('english' in c.expect) assert.deepEqual(got.english, c.expect.english);
      if ('from' in c.expect) assert.equal(got.pack?.from ?? null, c.expect.from);
      if ('direction' in c.expect) assert.equal(got.direction, c.expect.direction);
    } finally {
      fs.rmSync(root, { recursive: true, force: true });
    }
  });
}
