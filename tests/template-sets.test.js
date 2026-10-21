// Newer template sets fetched at the update check (apps/desktop/main/
// template-sets.js): a set signed with the key is taken, its files written
// under the data folder and used from then on; one not newer is left alone;
// one needing a newer Office is reported and not taken; and a set unsigned,
// signed with another key, with its description altered, or whose bundle is
// not the one it signs, or names a file outside itself, is never used.
import test from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createTemplateSets, verifyDescription, writeBundle } from '../apps/desktop/main/template-sets.js';

const keys = crypto.generateKeyPairSync('ed25519');
const PUBLIC = keys.publicKey.export({ type: 'spki', format: 'pem' });
const BUNDLED = { set: '1.1.0', minOffice: '2.12.0' };

/** A release as GitHub answers it, with a set made and signed here. */
function release({ set = '1.2.0', minOffice = '2.13.0', files = { 'manifest.json': JSON.stringify({ set, minOffice }), 'data-files/x.js': 'export default 1;\n' }, key = keys.privateKey, tamper = null, badBundle = false } = {}) {
  const enc = Object.fromEntries(Object.entries(files).map(([k, v]) => [k, Buffer.from(v).toString('base64')]));
  const bundleText = JSON.stringify({ set, files: enc });
  let description = JSON.stringify({ set, minOffice, parts: {}, bundle: { name: `templates-${set}.bundle.json`, size: Buffer.byteLength(bundleText), sha256: crypto.createHash('sha256').update(bundleText).digest('hex') } });
  const signature = crypto.sign(null, Buffer.from(description), key).toString('base64');
  if (tamper) description = tamper(description);
  const assets = { 'templates-set.json': JSON.stringify({ description, signature }), [`templates-${set}.bundle.json`]: badBundle ? bundleText.replace('"set":', '"set" :') : bundleText };
  return {
    fetchJson: async () => ({ assets: Object.keys(assets).map((name) => ({ name, browser_download_url: `mem://${name}` })) }),
    fetchText: async (url) => assets[url.slice('mem://'.length)],
  };
}
const folder = () => fs.mkdtempSync(path.join(os.tmpdir(), 'sets-'));

test('a newer set signed with the key is taken, written under the data folder and used from then on', async () => {
  const root = folder();
  try {
    const sets = createTemplateSets({ root, bundled: BUNDLED, office: '2.13.0', keys: [PUBLIC], ...release() });
    assert.equal(sets.version(), '1.1.0');
    assert.deepEqual(await sets.check(), { phase: 'ready', set: '1.2.0', needs: null, error: null });
    assert.equal(fs.readFileSync(path.join(root, '1.2.0', 'data-files', 'x.js'), 'utf8'), 'export default 1;\n');
    assert.equal(sets.inUse().set, '1.2.0');
    assert.equal(sets.version(), '1.2.0');
    // Checked again: the same set is not newer.
    assert.equal((await sets.check()).phase, 'current');
    // An Office too old for the downloaded set goes back to the bundled one.
    assert.equal(createTemplateSets({ root, bundled: BUNDLED, office: '2.12.0', keys: [PUBLIC], ...release() }).inUse(), null);
  } finally {
    fs.rmSync(root, { recursive: true, force: true });
  }
});

test('a set that needs a newer Office is reported, not taken; one not newer is left alone', async () => {
  const root = folder();
  try {
    const old = createTemplateSets({ root, bundled: BUNDLED, office: '2.12.0', keys: [PUBLIC], ...release() });
    assert.deepEqual(await old.check(), { phase: 'needs-office', set: '1.2.0', needs: '2.13.0', error: null });
    assert.equal(fs.existsSync(path.join(root, '1.2.0')), false);
    const same = createTemplateSets({ root, bundled: BUNDLED, office: '2.13.0', keys: [PUBLIC], ...release({ set: '1.1.0' }) });
    assert.equal((await same.check()).phase, 'current');
  } finally {
    fs.rmSync(root, { recursive: true, force: true });
  }
});

test('a set unsigned, signed with another key, altered, or with a bundle not the one it signs is never used', async () => {
  const other = crypto.generateKeyPairSync('ed25519').privateKey;
  for (const [what, r, why] of [
    ['another key', release({ key: other }), /not signed by Rutba/],
    ['an altered description', release({ tamper: (d) => d.replace('"2.13.0"', '"2.0.0"') }), /not signed by Rutba/],
    ['another bundle', release({ badBundle: true }), /bundle is not the one its description signs/],
  ]) {
    const root = folder();
    try {
      const sets = createTemplateSets({ root, bundled: BUNDLED, office: '2.13.0', keys: [PUBLIC], ...r });
      const s = await sets.check();
      assert.equal(s.phase, 'error', what);
      assert.match(s.error, why, what);
      assert.equal(sets.inUse(), null, what);
      assert.equal(fs.existsSync(path.join(root, '1.2.0')), false, what);
    } finally {
      fs.rmSync(root, { recursive: true, force: true });
    }
  }
  assert.throws(() => verifyDescription({ description: '{}' }), /no signed description/);
});

test('a bundle naming a file outside its folder is refused', () => {
  const root = folder();
  try {
    assert.throws(() => writeBundle(path.join(root, 'set'), { files: { '../escape.js': 'eA==' } }), /outside itself/);
    assert.equal(fs.existsSync(path.join(root, 'escape.js')), false);
    assert.equal(writeBundle(path.join(root, 'set'), { files: { 'a/b.txt': 'eA==' } }), 1);
  } finally {
    fs.rmSync(root, { recursive: true, force: true });
  }
});
