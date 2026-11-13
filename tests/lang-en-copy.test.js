// The English copy Office carries (C15, L9): checked as Office would use it,
// its files as lang-en.source.json says they were taken, and, when it is at
// this Office's version, every label the code uses in it. A copy for an
// earlier Office is allowed: the uncovered and stale counts are reported
// instead (RELEASING.md, the packs steps). With no copy taken yet, there is
// nothing to check and the test says so.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { usable, contractFor } from '../packages/office-shell/src/electron/addons-compat.js';
import { checkFolder, sha512Of } from '../packages/office-shell/src/electron/addons.js';
import { keyOf, hashOf } from '../packages/office-shell/src/electron/language-packs.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const DIR = path.join(ROOT, 'apps', 'desktop', 'resources', 'languages');
const ENVELOPE = path.join(DIR, 'lang-en.pack.json');
const none = fs.existsSync(ENVELOPE) ? false : 'no English copy taken yet (tools/take-lang-en.mjs)';
const read = (file) => JSON.parse(fs.readFileSync(file, 'utf8'));
const office = read(path.join(ROOT, 'apps', 'desktop', 'package.json')).version;
const contract = read(path.join(ROOT, 'packages', 'office-shell', 'src', 'addons-contract.json'));

test('the copy is a lang-en whose files are as signed and as taken', { skip: none }, () => {
  const envelope = read(ENVELOPE);
  const m = envelope.manifest;
  assert.equal(m.id, 'lang-en');
  assert.equal(m.kind, 'language');
  assert.equal(checkFolder(path.join(DIR, 'lang-en'), m, { hashAll: true }), null);
  const record = read(path.join(DIR, 'lang-en.source.json'));
  assert.equal(record.version, m.version);
  assert.equal(record.revision, m.revision || 0);
  assert.equal(record.envelope.sha512, sha512Of(ENVELOPE));
  for (const [p, f] of Object.entries(record.files)) assert.equal(sha512Of(path.join(DIR, 'lang-en', ...p.split('/'))), f.sha512, p);
});

test('the copy is signed by the packages key', { skip: none || (!contract.keys?.packages?.length && 'no packages keys yet: a release refuses an unchecked copy') }, () => {
  const envelope = read(ENVELOPE);
  const verdict = usable(envelope, { contract: contractFor(contract, { office }), office, keys: { packages: contract.keys.packages } });
  assert.equal(verdict.usable, true, verdict.reason);
});

test('at this Office\'s version the copy has every label the code uses; before it, the gaps are said', { skip: none }, (t) => {
  const envelope = read(ENVELOPE);
  const entries = read(path.join(DIR, 'lang-en', 'messages.json'));
  const keys = read(path.join(ROOT, 'packages', 'office-ui', 'src', 'catalogues', 'messages.json')).messages;
  const sources = new Set(Object.values(entries).map((e) => e?.source));
  const wanted = new Set(keys.map((k) => hashOf(keyOf(k))));
  const uncovered = keys.filter((k) => !sources.has(hashOf(keyOf(k)))).map((k) => k.message);
  const stale = Object.values(entries).filter((e) => !wanted.has(e?.source)).length;
  if (envelope.manifest.version === office) {
    assert.deepEqual(uncovered, [], `lang-en ${office} lacks labels the code uses: take a lang-en built from this Office`);
  } else {
    t.diagnostic(`the copy is lang-en ${envelope.manifest.version} r${envelope.manifest.revision}, for an earlier Office: ${uncovered.length} labels uncovered, ${stale} stale`);
  }
});
