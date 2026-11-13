// The English copy Office carries (C15, L9; private design 1.4), as the
// release build checks it (tools/check-english-copy.mjs, run by dist.js):
// no copy, a copy taken unchecked or signed with another key, a changed
// file, and a copy at this Office's version that lacks labels or has stale
// ones are all refused; a copy for an earlier Office passes with its gaps
// reported. Then the copy this branch carries, through the same check.
import test from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { keyOf, hashOf } from '../packages/office-shell/src/electron/language-packs.js';
import { envelopeOf, fileEntry } from '../tools/make-addon-vectors.mjs';
import { takeLangEn } from '../tools/take-lang-en.mjs';
import { englishCopyFaults, COPY_DIR } from '../tools/check-english-copy.mjs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const read = (file) => JSON.parse(fs.readFileSync(file, 'utf8'));
const OFFICE = read(path.join(ROOT, 'apps', 'desktop', 'package.json')).version;
const CONTRACT = read(path.join(ROOT, 'packages', 'office-shell', 'src', 'addons-contract.json'));

const LABELS = [{ message: 'Close' }, { message: 'Saved {name}' }];
const H = (m) => hashOf(keyOf(m));
const pair = crypto.generateKeyPairSync('ed25519');
const pem = pair.publicKey.export({ type: 'spki', format: 'pem' }).toString();
const contract = { ...CONTRACT, keys: { packages: [pem], release: [] } };

/** A released lang-en, as the packages repository leaves one, into a folder. */
function released(dir, { version = OFFICE, entries = { 'shell.close': { text: 'Close', source: H(LABELS[0]) }, 'shell.saved': { text: 'Saved {name}', source: H(LABELS[1]) } }, key = pair } = {}) {
  const body = `${JSON.stringify(entries, null, 2)}\n`;
  fs.mkdirSync(path.join(dir, 'lang-en'), { recursive: true });
  fs.writeFileSync(path.join(dir, 'lang-en', 'messages.json'), body);
  const manifest = { id: 'lang-en', kind: 'language', format: 2, version, revision: 1, tag: 'en', name: 'English', files: { 'messages.json': fileEntry(body) } };
  fs.writeFileSync(path.join(dir, 'lang-en.pack.json'), JSON.stringify(envelopeOf(manifest, key.privateKey)));
  return dir;
}

test('a release refuses no copy, an unchecked one, another key, a changed file, and gaps at this version', () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-en-copy-'));
  try {
    const check = (dir, more = {}) => englishCopyFaults({ dir, office: OFFICE, contract, labels: LABELS, ...more });
    assert.match(check(path.join(root, 'none')).faults.join(), /there is no English copy/);

    const good = path.join(root, 'good');
    takeLangEn({ from: released(path.join(root, 'r1')), to: good, keys: [pem], labels: LABELS, office: OFFICE });
    assert.deepEqual(check(good), { faults: [], report: null, version: OFFICE });

    const unchecked = path.join(root, 'unchecked');
    takeLangEn({ from: released(path.join(root, 'r2')), to: unchecked, keys: [], unchecked: true, labels: LABELS });
    assert.match(check(unchecked).faults.join(), /without its signature checked/);

    const other = path.join(root, 'other');
    takeLangEn({ from: released(path.join(root, 'r3'), { key: crypto.generateKeyPairSync('ed25519') }), to: other, keys: [], unchecked: true, labels: LABELS });
    assert.match(check(other).faults.join(), /not signed by Rutba/);

    fs.appendFileSync(path.join(good, 'lang-en', 'messages.json'), ' ');
    assert.match(check(good).faults.join(), /changed since it was signed/);

    const lacking = path.join(root, 'lacking');
    takeLangEn({ from: released(path.join(root, 'r4'), { entries: { 'shell.close': { text: 'Close', source: H(LABELS[0]) } } }), to: lacking, keys: [pem], labels: LABELS, office: OFFICE });
    assert.match(check(lacking).faults.join(), /lacks 1 labels the code uses \(first: Saved \{name\}\)/);

    const stale = path.join(root, 'stale');
    takeLangEn({ from: released(path.join(root, 'r5'), { entries: { 'shell.close': { text: 'Close', source: H(LABELS[0]) }, 'shell.saved': { text: 'Saved {name}', source: H(LABELS[1]) }, 'shell.gone': { text: 'Gone', source: H({ message: 'Gone' }) } } }), to: stale, keys: [pem], labels: LABELS, office: OFFICE });
    assert.match(check(stale).faults.join(), /1 entries the code no longer uses \(first: shell\.gone\)/);

    // The fallback: a copy for an earlier Office passes, its gaps said.
    const earlier = path.join(root, 'earlier');
    takeLangEn({ from: released(path.join(root, 'r6'), { version: '2.0.0', entries: { 'shell.close': { text: 'Close', source: H(LABELS[0]) } } }), to: earlier, keys: [pem], labels: LABELS, office: '2.0.0' });
    const r = check(earlier);
    assert.deepEqual(r.faults, []);
    assert.equal(r.report, 'the copy is lang-en 2.0.0 r1, for an earlier Office: 1 labels uncovered, 0 stale');
  } finally {
    fs.rmSync(root, { recursive: true, force: true });
  }
});

test('the copy this branch carries, as the release build checks it', (t) => {
  const labels = read(path.join(ROOT, 'packages', 'office-ui', 'src', 'catalogues', 'messages.json')).messages;
  const r = englishCopyFaults({ dir: COPY_DIR, office: OFFICE, contract: CONTRACT, labels });
  if (!fs.existsSync(path.join(COPY_DIR, 'lang-en.pack.json'))) {
    // Taken at the release step (RELEASING.md): until then the release build refuses, as it must.
    assert.match(r.faults.join(), /there is no English copy/);
    t.diagnostic('no English copy yet: the release build refuses until the released lang-en is taken');
    return;
  }
  assert.deepEqual(r.faults, []);
  if (r.report) t.diagnostic(r.report);
});
