/**
 * An update is installed only if it was signed with a key that is not on
 * GitHub. A stolen token can change an installer and its latest.yml together;
 * it cannot make this check pass.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { UPDATE_KEYS, signedWords, sha512OfFile, signatureHolds, whyNotOurs, guardDownloads, installerIn, signatureAddress } from '../apps/desktop/main/update-signature.js';
import { sha512Of, signAndPublish, publishRelease, parseFeed } from '../tools/publish-release.js';

const pair = () => {
  const { publicKey, privateKey } = crypto.generateKeyPairSync('ed25519');
  return { pub: publicKey.export({ type: 'spki', format: 'pem' }), key: privateKey };
};
const main = pair();
const recovery = pair();
const stranger = pair();
const KEYS = [main.pub, recovery.pub];
const sign = (key, words) => crypto.sign(null, Buffer.from(words, 'utf8'), key).toString('base64');
const VERSION = '2.25.0';
const FILE = `Rutba-Office-${VERSION}-win-x64.exe`;

function download(bytes = 'the installer') {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-signature-'));
  const filePath = path.join(dir, FILE);
  fs.writeFileSync(filePath, bytes);
  return filePath;
}
const wordsFor = async (filePath, over = {}) => signedWords({ version: VERSION, file: FILE, sha512: await sha512OfFile(filePath), ...over });
const check = (filePath, text, keys = KEYS) => whyNotOurs({ filePath, version: VERSION, file: FILE, fetchSignature: async () => text, keys });

test('the real keys are two different Ed25519 public keys', () => {
  assert.equal(UPDATE_KEYS.length, 2);
  assert.notEqual(UPDATE_KEYS[0], UPDATE_KEYS[1]);
  for (const pem of UPDATE_KEYS) assert.equal(crypto.createPublicKey(pem).asymmetricKeyType, 'ed25519');
});

test('an update signed with either key is taken', async () => {
  const file = download();
  assert.equal(await check(file, sign(main.key, await wordsFor(file))), null);
  assert.equal(await check(file, sign(recovery.key, await wordsFor(file))), null);
});

test('an installer changed after signing is refused', async () => {
  const file = download();
  const text = sign(main.key, await wordsFor(file));
  fs.appendFileSync(file, 'x');
  assert.match(await check(file, text), /not signed by the makers/);
});

test('a signature made by anyone else is refused', async () => {
  const file = download();
  assert.match(await check(file, sign(stranger.key, await wordsFor(file))), /not signed by the makers/);
});

test('an old signed installer offered as a new version is refused', async () => {
  const file = download();
  assert.match(await check(file, sign(main.key, await wordsFor(file, { version: '2.20.0' }))), /not signed by the makers/);
});

test('a signature for another file is refused', async () => {
  const file = download();
  assert.match(await check(file, sign(main.key, await wordsFor(file, { file: 'Other.exe' }))), /not signed by the makers/);
});

test('a signature that is missing, empty, too long or not a signature is refused', async () => {
  const file = download();
  for (const text of ['', '   ', null, 'not base64 at all!', 'QUJD', 'A'.repeat(100000), Buffer.alloc(64).toString('base64')]) {
    assert.match(await check(file, text), /not signed by the makers/, String(text).slice(0, 20));
  }
});

test('a signature that cannot be fetched stops the update', async () => {
  const file = download();
  const why = await whyNotOurs({ filePath: file, version: VERSION, file: FILE, fetchSignature: async () => { throw new Error('offline'); }, keys: KEYS });
  assert.match(why, /could not be fetched/);
});

test('an update that does not name itself, or a download that is gone, stops the update', async () => {
  assert.match(await whyNotOurs({ filePath: download(), version: null, file: FILE, fetchSignature: async () => '', keys: KEYS }), /does not say what it is/);
  assert.match(await whyNotOurs({ filePath: path.join(os.tmpdir(), 'no-such-file.exe'), version: VERSION, file: FILE, fetchSignature: async () => '', keys: KEYS }), /could not be read/);
});

test('the installer is found in the feed, and the signature beside it', () => {
  assert.equal(installerIn({ files: [{ url: 'a.blockmap' }, { url: 'Rutba%20Office.exe' }] }), 'Rutba Office.exe');
  assert.equal(installerIn({ path: 'x.exe' }), 'x.exe');
  assert.equal(installerIn({}), null);
  assert.equal(signatureAddress({ owner: 'o', repo: 'r', version: '1.2.3', file: 'a b.exe' }), 'https://github.com/o/r/releases/download/v1.2.3/a%20b.exe.sig');
});

test('the updater is given the check, and answers from the version being offered', async () => {
  const file = download();
  const asked = [];
  const updater = {};
  const text = sign(main.key, await wordsFor(file));
  guardDownloads(updater, { current: () => ({ version: VERSION, files: [{ url: FILE }] }), fetchSignature: async (a) => { asked.push(a); return text; }, keys: KEYS });
  assert.equal(await updater.verifySignature(file), null);
  assert.deepEqual(asked, [{ version: VERSION, file: FILE }]);
  fs.appendFileSync(file, 'x');
  assert.match(await updater.verifySignature(file), /not signed/);
});

test('the update service installs the check on Windows, and asks the network through Electron', () => {
  const source = fs.readFileSync(path.join(import.meta.dirname, '..', 'apps', 'desktop', 'main', 'updates.js'), 'utf8');
  assert.match(source, /guardDownloads\(updater,/);
  assert.match(source, /process\.platform === 'win32'/);
  assert.match(source, /net\.fetch\(/);
});

test('electron-updater still calls the check and still refuses on an answer', () => {
  // The hook is its method by name: an upgrade that renames it would turn the check off without a sound.
  const source = fs.readFileSync(path.join(import.meta.dirname, '..', 'node_modules', 'electron-updater', 'out', 'NsisUpdater.js'), 'utf8');
  assert.match(source, /await this\.verifySignature\(destinationFile\)/);
  assert.match(source, /signatureVerificationStatus != null[\s\S]{0,200}throw/);
});

// ---- signing a release

const TAG = `v${VERSION}`;
const quiet = () => {};

function draftOnGitHub({ installer = Buffer.from('the installer'), feedSha = null, draft = true, feedVersion = VERSION } = {}) {
  const feed = `version: ${feedVersion}\nfiles:\n  - url: ${FILE}\n    sha512: ${feedSha ?? sha512Of(installer)}\n    size: ${installer.length}\npath: ${FILE}\nsha512: ${feedSha ?? sha512Of(installer)}\n`;
  let id = 1;
  const assets = new Map([
    [FILE, { id: id++, name: FILE, state: 'uploaded', size: installer.length, bytes: installer, uploader: { login: 'github-actions[bot]' } }],
    [`${FILE}.blockmap`, { id: id++, name: `${FILE}.blockmap`, state: 'uploaded', size: 3, bytes: Buffer.from('map') }],
    ['latest.yml', { id: id++, name: 'latest.yml', state: 'uploaded', size: feed.length, bytes: Buffer.from(feed), uploader: { login: 'github-actions[bot]' } }],
  ]);
  const release = { id: 9, tag_name: TAG, draft, html_url: 'x', body: 'N' };
  const calls = [];
  return {
    calls,
    assets,
    release,
    client: {
      listReleases: async () => [release],
      listAssets: async () => [...assets.values()],
      deleteAsset: async (_r, a) => { calls.push(`delete ${a.name}`); assets.delete(a.name); },
      downloadAsset: async (_r, a) => a.bytes,
      uploadAsset: async (_r, name, file, size) => {
        calls.push(`upload ${name}`);
        assets.set(name, { id: id++, name, state: 'uploaded', size, bytes: fs.readFileSync(file) });
      },
      updateRelease: async (r, body) => { calls.push(`update ${JSON.stringify(body)}`); Object.assign(r, body); return r; },
      createRelease: async () => { throw new Error('signing makes no release'); },
    },
  };
}
const signer = (key) => (words) => sign(key, words);

test('signing: the installer is read back from the draft, signed, and only then published', async () => {
  const gh = draftOnGitHub();
  await signAndPublish({ client: gh.client, tag: TAG, sign: signer(main.key), keys: KEYS, log: quiet });
  assert.deepEqual(gh.calls.map((c) => c.split(' ')[0]), ['upload', 'update']);
  assert.equal(gh.release.draft, false);
  const text = gh.assets.get(`${FILE}.sig`).bytes.toString();
  assert.equal(signatureHolds(signedWords({ version: VERSION, file: FILE, sha512: sha512Of(Buffer.from('the installer')) }), text, KEYS), true);
});

test('signing: an installer that is not what the feed says is not signed or published', async () => {
  const gh = draftOnGitHub({ feedSha: sha512Of(Buffer.from('something else')) });
  await assert.rejects(signAndPublish({ client: gh.client, tag: TAG, sign: signer(main.key), keys: KEYS, log: quiet }), /not what latest\.yml says/);
  assert.deepEqual(gh.calls, []);
  assert.equal(gh.release.draft, true);
});

test('signing: a key the program does not trust is found before anything is published', async () => {
  const gh = draftOnGitHub();
  await assert.rejects(signAndPublish({ client: gh.client, tag: TAG, sign: signer(stranger.key), keys: KEYS, log: quiet }), /would refuse/);
  assert.equal(gh.calls.some((c) => c.startsWith('update')), false);
  assert.equal(gh.release.draft, true);
});

test('signing: no draft, or a feed for another version, stops', async () => {
  const none = draftOnGitHub();
  none.client.listReleases = async () => [];
  await assert.rejects(signAndPublish({ client: none.client, tag: TAG, sign: signer(main.key), keys: KEYS, log: quiet }), /no release/);
  const other = draftOnGitHub({ feedVersion: '2.24.0' });
  await assert.rejects(signAndPublish({ client: other.client, tag: TAG, sign: signer(main.key), keys: KEYS, log: quiet }), /do not agree/);
  assert.equal(other.release.draft, true);
});

test('a lost signature can be put back on a published release, and a good one is never taken down', async () => {
  const gh = draftOnGitHub({ draft: false });
  await signAndPublish({ client: gh.client, tag: TAG, sign: signer(main.key), keys: KEYS, log: quiet });
  await signAndPublish({ client: gh.client, tag: TAG, sign: signer(recovery.key), keys: KEYS, log: quiet });
  assert.deepEqual(gh.calls.map((c) => c.split(' ')[0]), ['upload']);
  gh.assets.get(`${FILE}.sig`).bytes = Buffer.from('spoiled');
  await signAndPublish({ client: gh.client, tag: TAG, sign: signer(recovery.key), keys: KEYS, log: quiet });
  assert.deepEqual(gh.calls.map((c) => c.split(' ')[0]), ['upload', 'delete', 'upload']);
});

test('signing refuses a draft built from another commit, and says which', async () => {
  const gh = draftOnGitHub();
  gh.release.target_commitish = 'aaa111';
  await assert.rejects(signAndPublish({ client: gh.client, tag: TAG, commit: 'bbb222', sign: signer(main.key), keys: KEYS, log: quiet }), /built from aaa111.*at bbb222/);
  assert.deepEqual(gh.calls, []);
  await signAndPublish({ client: gh.client, tag: TAG, commit: 'aaa111', sign: signer(main.key), keys: KEYS, log: quiet });
  assert.equal(gh.release.draft, false);
});

test('signing refuses an installer the release workflow did not upload', async () => {
  const gh = draftOnGitHub();
  gh.assets.get(FILE).uploader = { login: 'someone' };
  await assert.rejects(signAndPublish({ client: gh.client, tag: TAG, sign: signer(main.key), keys: KEYS, log: quiet }), /did not come from the release workflow/);
  assert.deepEqual(gh.calls, []);
});

test('two releases of one tag are refused until one has gone', async () => {
  const gh = draftOnGitHub();
  gh.client.listReleases = async () => [gh.release, { ...gh.release, id: 10 }];
  await assert.rejects(signAndPublish({ client: gh.client, tag: TAG, sign: signer(main.key), keys: KEYS, log: quiet }), /2 releases/);
  assert.deepEqual(gh.calls, []);
});

test('each app pack is signed as the installer is', async () => {
  const gh = draftOnGitHub();
  const pack = 'Rutba-Office-2.25.0-sheets.pack';
  const bytes = Buffer.from('the pack');
  gh.assets.set(pack, { id: 90, name: pack, state: 'uploaded', size: bytes.length, bytes, uploader: { login: 'github-actions[bot]' } });
  await signAndPublish({ client: gh.client, tag: TAG, sign: signer(main.key), keys: KEYS, log: quiet });
  assert.equal(signatureHolds(signedWords({ version: VERSION, file: pack, sha512: sha512Of(bytes) }), gh.assets.get(`${pack}.sig`).bytes.toString(), KEYS), true);
});

test('publishing from a build leaves a draft when told to, so nothing goes out unsigned', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-publish-'));
  const installer = Buffer.from('the installer');
  fs.writeFileSync(path.join(dir, FILE), installer);
  fs.writeFileSync(path.join(dir, `${FILE}.blockmap`), 'map');
  fs.writeFileSync(path.join(dir, 'latest.yml'), `version: ${VERSION}\nfiles:\n  - url: ${FILE}\n    sha512: ${sha512Of(installer)}\n    size: ${installer.length}\npath: ${FILE}\nsha512: ${sha512Of(installer)}\n`);
  const gh = draftOnGitHub();
  gh.assets.clear();
  gh.client.listReleases = async () => [gh.release];
  const release = await publishRelease({ client: gh.client, dir, tag: TAG, notes: 'N', leaveDraft: true, log: quiet });
  assert.equal(release.draft, true);
  assert.equal(gh.calls.some((c) => c.startsWith('update')), false);
});

test('the release workflow holds no update key and the tool stops at a draft without --sign', () => {
  const root = path.join(import.meta.dirname, '..');
  assert.doesNotMatch(fs.readFileSync(path.join(root, '.github', 'workflows', 'release.yml'), 'utf8'), /RUTBA_UPDATE_KEY/);
  const tool = fs.readFileSync(path.join(root, 'tools', 'publish-release.js'), 'utf8');
  assert.match(tool, /leaveDraft: true/);
  assert.match(tool, /RUTBA_UPDATE_KEY_FILE/);
});

test('what the release PC signs is what an installed copy asks for and accepts', async () => {
  const gh = draftOnGitHub();
  await signAndPublish({ client: gh.client, tag: TAG, sign: signer(main.key), keys: KEYS, log: quiet });
  const info = parseFeed(gh.assets.get('latest.yml').bytes.toString());
  const asked = [];
  const updater = {};
  guardDownloads(updater, {
    current: () => ({ version: info.version, files: info.files, path: info.path }),
    fetchSignature: async ({ version, file }) => {
      asked.push(signatureAddress({ owner: 'o', repo: 'r', version, file }));
      return gh.assets.get(`${file}.sig`).bytes.toString();
    },
    keys: KEYS,
  });
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-real-')), FILE);
  fs.writeFileSync(file, 'the installer');
  assert.equal(await updater.verifySignature(file), null);
  assert.equal(asked[0], `https://github.com/o/r/releases/download/v${VERSION}/${encodeURIComponent(FILE)}.sig`);
  fs.writeFileSync(file, 'something else');
  assert.match(await updater.verifySignature(file), /not signed/);
});
