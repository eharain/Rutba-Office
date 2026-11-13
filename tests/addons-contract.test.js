// The add-ons contract (packages/office-shell/src/addons-contract.json) and
// Office's one rule for it (addons-compat.js): the shared cases in
// test-vectors/addons/, which the packages repository runs too, and the
// feed checks. Private design: office-pack-compatibility-design.md 1.2, 3.1.
import test from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { usable, readiness, verifyFeed, acceptFeed, contractFor, compareVersions, FEED_WORDS } from '../packages/office-shell/src/electron/addons-compat.js';
import { canonical } from '../packages/office-shell/src/electron/packs.js';
import { VECTORS_DIR, CONTRACT_FILE, VECTORS, contractHash, cases } from '../tools/make-addon-vectors.mjs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const contract = JSON.parse(fs.readFileSync(CONTRACT_FILE, 'utf8'));
const vectors = JSON.parse(fs.readFileSync(path.join(VECTORS_DIR, 'vectors.json'), 'utf8'));
const testKey = crypto.createPublicKey(fs.readFileSync(path.join(VECTORS_DIR, 'test-key.pub')));
const inside = (rel) => JSON.parse(fs.readFileSync(path.join(VECTORS_DIR, rel), 'utf8'));
const named = (got, expect) => Object.fromEntries(Object.keys(expect).map((k) => [k, got?.[k]]));

test('the contract is for this Office version', () => {
  const pkg = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'));
  assert.equal(contract.contract, 1);
  assert.equal(contract.office, pkg.version);
});

test('the engine is trusted to the release key and is exactly its Office version; data kinds to the packages key', () => {
  assert.equal(contract.kinds.engine.trust, 'release');
  assert.equal(contract.kinds.engine.exactVersion, true);
  for (const [kind, k] of Object.entries(contract.kinds)) if (kind !== 'engine') assert.equal(k.trust, 'packages', kind);
});

test('the shared cases were made from this contract', () => {
  assert.equal(vectors.vectors, VECTORS);
  assert.equal(vectors.contract, contract.contract);
  assert.equal(vectors.from, contractHash(), 'the contract changed: make the cases again with node tools/make-addon-vectors.mjs');
  // The same cases the generator would make now, by name and answer.
  const fresh = cases(contract, crypto.generateKeyPairSync('ed25519').privateKey).cases;
  assert.deepEqual(vectors.cases.map((c) => [c.name, c.expect]), fresh.map((c) => [c.name, c.expect]));
});

test('the test key is trusted for nothing', () => {
  const pem = testKey.export({ type: 'spki', format: 'pem' }).toString();
  assert.ok(!contract.keys.packages.includes(pem));
  assert.ok(!contract.keys.release.includes(pem));
});

test('every shared case gives what it expects', () => {
  const wrong = [];
  for (const c of vectors.cases) {
    const caseContract = c.contract || contract;
    let got;
    if (c.type === 'load') {
      const keys = { packages: [], release: [], [c.signedAs]: [testKey] };
      got = usable(inside(c.pack), { contract: caseContract, office: c.office, keys, engine: c.engine ? inside(c.engine).manifest : null });
    } else if (c.type === 'ready') {
      got = readiness({ installed: c.installed, feed: c.feed ? inside(c.feed) : null, target: c.target, contract: caseContract });
    } else {
      wrong.push(`${c.name}: no case type ${c.type}`);
      continue;
    }
    if (JSON.stringify(named(got, c.expect)) !== JSON.stringify(c.expect)) wrong.push(`${c.name}: ${JSON.stringify(got)}`);
  }
  assert.deepEqual(wrong, []);
});

test('usable always answers an object, with a reason when it refuses', () => {
  for (const env of [null, undefined, {}, 'text', { format: 'rutba-office-pack', manifest: null }]) {
    const u = usable(env, { contract, office: contract.office, keys: {} });
    assert.equal(u.usable, false);
    assert.equal(u.code, 'envelope');
    assert.equal(typeof u.reason, 'string');
  }
  assert.equal(usable(inside('packs/lang-xx.pack.json'), { contract: { ...contract, contract: 2 }, office: contract.office, keys: { packages: [testKey] } }).code, 'envelope');
});

test('a newer version or revision orders higher', () => {
  assert.ok(compareVersions({ version: '2.25.1', revision: 1 }, { version: '2.25.0', revision: 9 }) > 0);
  assert.ok(compareVersions({ version: '2.25.1', revision: 2 }, { version: '2.25.1', revision: 1 }) > 0);
  assert.equal(compareVersions({ version: '2.25.1', revision: 1 }, { version: '2.25.1', revision: 1 }), 0);
  assert.ok(compareVersions({ version: '2.10.0' }, { version: '2.9.0' }) > 0);
});

test('a feed is verified with its words, and taken only for its own version and only when newer', () => {
  const { privateKey, publicKey } = crypto.generateKeyPairSync('ed25519');
  const feed = { format: 1, readiness: 1, office: '2.25.1', sequence: 4, state: 'partial', retired: [], packs: [] };
  const sig = { signature: crypto.sign(null, Buffer.from(FEED_WORDS + canonical(feed), 'utf8'), privateKey).toString('base64') };
  assert.ok(verifyFeed(feed, sig, [publicKey]));
  assert.ok(!verifyFeed({ ...feed, sequence: 5 }, sig, [publicKey]));
  assert.ok(!verifyFeed(feed, { signature: crypto.sign(null, Buffer.from(canonical(feed)), privateKey).toString('base64') }, [publicKey]), 'a manifest signature is not a feed signature');
  assert.ok(!verifyFeed(feed, sig, []));

  const first = acceptFeed(feed, { office: '2.25.1', seen: { '2.26.0': 9 } });
  assert.deepEqual(first, { ok: true, seen: { '2.26.0': 9, '2.25.1': 4 } });
  assert.equal(acceptFeed(feed, { office: '2.25.1', seen: first.seen }).ok, false, 'the same sequence again is old');
  assert.equal(acceptFeed({ ...feed, sequence: 3 }, { office: '2.25.1', seen: first.seen }).ok, false);
  assert.equal(acceptFeed(feed, { office: '2.26.0', seen: {} }).ok, false, 'a feed is for the version it names');
  assert.equal(acceptFeed({ ...feed, format: 2 }, { office: '2.25.1' }).ok, false);
  assert.equal(acceptFeed({ ...feed, state: 'done' }, { office: '2.25.1' }).ok, false);
  // The target's feed, read for the update gate, never makes the running version's look old.
  const target = acceptFeed({ ...feed, office: '2.26.0', sequence: 50 }, { office: '2.26.0', seen: first.seen });
  assert.equal(acceptFeed({ ...feed, sequence: 5 }, { office: '2.25.1', seen: target.seen }).ok, true);
});

test('the running contract takes the core\'s release key for the engine and keeps the packages keys', () => {
  const { publicKey } = crypto.generateKeyPairSync('ed25519');
  const pem = publicKey.export({ type: 'spki', format: 'pem' }).toString();
  const running = contractFor({ ...contract, keys: { packages: ['P'], release: [] } }, { office: '2.25.1', releaseKeys: [publicKey, pem] });
  assert.deepEqual(running.keys, { packages: ['P'], release: [pem] });
  assert.equal(running.office, '2.25.1');
});
