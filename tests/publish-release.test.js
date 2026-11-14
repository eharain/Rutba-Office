/**
 * Publishing a release (tools/publish-release.js): made as a draft, every
 * file uploaded, the update feed read back from the release and held to the
 * files, and only then published. A stand-in for GitHub here records what
 * was asked of it; nothing touches the network.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { releaseAssets, parseFeed, checkFeed, sha512Of, publishRelease, gateProblem } from '../tools/publish-release.js';

const TAG = 'v2.1.0';
const INSTALLER = 'Rutba-Office-2.1.0-win-x64.exe';
const PORTABLE = 'Rutba-Office-2.1.0-portable.exe';

/** A build directory as `npm run dist:win` leaves it, with this version's files and an older one's. */
function buildDir({ feedFor = null, feedVersion = '2.1.0', blockmap = true, feed = true } = {}) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-publish-'));
  const installer = Buffer.from('the installer of 2.1.0, a stand-in for a hundred megabytes');
  fs.writeFileSync(path.join(dir, INSTALLER), installer);
  fs.writeFileSync(path.join(dir, PORTABLE), Buffer.from('the portable executable'));
  if (blockmap) fs.writeFileSync(path.join(dir, `${INSTALLER}.blockmap`), Buffer.from('blockmap'));
  fs.writeFileSync(path.join(dir, 'Rutba-Office-2.0.0-win-x64.exe'), Buffer.from('last version'));
  fs.writeFileSync(path.join(dir, 'builder-debug.yml'), 'x');
  const described = feedFor ?? installer;
  const sha = sha512Of(described);
  if (feed) {
    fs.writeFileSync(
      path.join(dir, 'latest.yml'),
      `version: ${feedVersion}\nfiles:\n  - url: ${INSTALLER}\n    sha512: ${sha}\n    size: ${described.length}\npath: ${INSTALLER}\nsha512: ${sha}\nreleaseDate: '2026-10-15T11:10:26.493Z'\n`
    );
  }
  return dir;
}

/** GitHub, as far as publishing goes, in memory. `failUpload` names a file whose upload fails once. */
function fakeGitHub({ releases = [], failUpload = null } = {}) {
  const calls = [];
  let nextId = 100;
  const state = { releases: releases.map((r) => ({ assets: [], ...r })) };
  const find = (release) => state.releases.find((r) => r.id === release.id);
  const client = {
    listReleases: async () => {
      calls.push('list');
      return state.releases.map(({ assets, ...r }) => r);
    },
    createRelease: async (body) => {
      calls.push(`create draft=${body.draft} prerelease=${body.prerelease}`);
      const r = { id: nextId++, ...body, html_url: `https://github.com/r/${body.tag_name}`, upload_url: 'https://uploads/x{?name}', assets: [] };
      state.releases.push(r);
      return r;
    },
    listAssets: async (release) => find(release).assets.map(({ bytes, ...a }) => a),
    deleteAsset: async (release, asset) => {
      calls.push(`delete ${asset.name}`);
      find(release).assets = find(release).assets.filter((a) => a.id !== asset.id);
    },
    uploadAsset: async (release, name, file) => {
      if (failUpload === name) {
        failUpload = null;
        calls.push(`upload ${name} FAILED`);
        throw new Error('POST upload → 502 Bad Gateway');
      }
      calls.push(`upload ${name}`);
      const bytes = fs.readFileSync(file);
      find(release).assets.push({ id: nextId++, name, size: bytes.length, state: 'uploaded', bytes, digest: `sha256:${crypto.createHash('sha256').update(bytes).digest('hex')}` });
    },
    downloadAsset: async (release, asset) => {
      calls.push(`download ${asset.name}`);
      return find(release).assets.find((a) => a.id === asset.id).bytes;
    },
    updateRelease: async (release, body) => {
      calls.push(`update draft=${body.draft} tag=${body.tag_name}`);
      Object.assign(find(release), body);
      const { assets, ...r } = find(release);
      return r;
    },
  };
  return { client, calls, state };
}

const quiet = () => {};

test('the files of this version go up, the update feeds last, nothing of an older version and no build log', () => {
  const names = ['latest.yml', INSTALLER, `${INSTALLER}.blockmap`, PORTABLE, 'Rutba-Office-2.0.0-win-x64.exe', 'builder-debug.yml', 'builder-effective-config.yaml', 'win-unpacked', 'Rutba-Office-2.1.0-x86_64.AppImage', 'latest-linux.yml', 'Rutba-Office-2.1.0-sbom.cdx.json', 'Rutba-Office-2.0.0-sbom.cdx.json'];
  assert.deepEqual(releaseAssets(names, TAG), ['Rutba-Office-2.1.0-portable.exe', 'Rutba-Office-2.1.0-sbom.cdx.json','Rutba-Office-2.1.0-win-x64.exe', 'Rutba-Office-2.1.0-win-x64.exe.blockmap', 'Rutba-Office-2.1.0-x86_64.AppImage', 'latest-linux.yml', 'latest.yml']);
});

test("each app's pack of this version goes up, for Add or remove apps to fetch", () => {
  const names = ['latest.yml', INSTALLER, 'Rutba-Office-2.1.0-mail.pack', 'Rutba-Office-2.1.0-word.pack', 'Rutba-Office-2.0.0-mail.pack'];
  assert.deepEqual(releaseAssets(names, TAG), ['Rutba-Office-2.1.0-mail.pack', 'Rutba-Office-2.1.0-win-x64.exe', 'Rutba-Office-2.1.0-word.pack', 'latest.yml']);
});

test('the update feed is read as electron-builder writes it', () => {
  const feed = parseFeed("version: 2.0.1\r\nfiles:\r\n  - url: Rutba-Office-2.0.1-win-x64.exe\r\n    sha512: YCra/x==\r\n    size: 119234685\r\npath: Rutba-Office-2.0.1-win-x64.exe\r\nsha512: YCra/x==\r\nreleaseDate: '2026-10-15T11:10:26.493Z'\r\n");
  assert.equal(feed.version, '2.0.1');
  assert.deepEqual(feed.files, [{ url: 'Rutba-Office-2.0.1-win-x64.exe', sha512: 'YCra/x==', size: 119234685 }]);
  assert.equal(feed.path, 'Rutba-Office-2.0.1-win-x64.exe');
  assert.equal(feed.releaseDate, '2026-10-15T11:10:26.493Z');
});

test('the leftCore line of a feed is read, and held to the contract when the release names one', () => {
  assert.deepEqual(parseFeed("version: 2.1.0\nleftCore: ['ar', 'hi', 'ur']\n").leftCore, ['ar', 'hi', 'ur']);
  assert.deepEqual(parseFeed('version: 2.1.0\r\nleftCore: []\r\n').leftCore, []);
  const local = new Map([[INSTALLER, { size: 10, sha512: 'right' }]]);
  const uploaded = new Map([[INSTALLER, 10], [`${INSTALLER}.blockmap`, 3]]);
  const feed = (over = {}) => ({ version: '2.1.0', files: [{ url: INSTALLER, sha512: 'right', size: 10 }], path: INSTALLER, sha512: 'right', ...over });
  const check = (f, leftCore) => checkFeed({ name: 'latest.yml', feed: f, tag: TAG, uploaded, local, leftCore });
  assert.deepEqual(check(feed({ leftCore: ['ur', 'ar'] }), ['ar', 'ur']), []);
  assert.deepEqual(check(feed({ leftCore: [] }), []), []);
  assert.deepEqual(check(feed(), null), [], 'not asked: not checked');
  assert.match(check(feed(), [])[0], /does not say which languages left the core/);
  assert.match(check(feed({ leftCore: ['ur'] }), ['ar', 'ur'])[0], /\[ur\], and the contract says \[ar, ur\]/);
});

test('a feed that disagrees with the release says each way it does', () => {
  const local = new Map([[INSTALLER, { size: 10, sha512: 'right' }]]);
  const uploaded = new Map([[INSTALLER, 10], [`${INSTALLER}.blockmap`, 3]]);
  const feed = (over = {}) => ({ version: '2.1.0', files: [{ url: INSTALLER, sha512: 'right', size: 10 }], path: INSTALLER, sha512: 'right', ...over });
  assert.deepEqual(checkFeed({ name: 'latest.yml', feed: feed(), tag: TAG, uploaded, local }), []);
  assert.deepEqual(checkFeed({ name: 'latest.yml', feed: feed({ version: '2.0.0' }), tag: TAG, uploaded, local }), ['latest.yml is for version 2.0.0, not 2.1.0']);
  assert.match(checkFeed({ name: 'latest.yml', feed: feed({ files: [{ url: 'Other.exe', sha512: 'right', size: 10 }] }), tag: TAG, uploaded, local })[0], /names Other\.exe, which is not in the release/);
  assert.match(checkFeed({ name: 'latest.yml', feed: feed({ files: [{ url: INSTALLER, sha512: 'right', size: 11 }] }), tag: TAG, uploaded, local })[0], /says .* is 11 bytes; the release has 10/);
  assert.match(checkFeed({ name: 'latest.yml', feed: feed({ files: [{ url: INSTALLER, sha512: 'wrong', size: 10 }], sha512: 'wrong' }), tag: TAG, uploaded, local })[0], /SHA-512 .* is not the SHA-512 of the file that was built/);
  assert.match(checkFeed({ name: 'latest.yml', feed: feed(), tag: TAG, uploaded: new Map([[INSTALLER, 10]]), local })[0], /blockmap is not in the release/);
  assert.match(checkFeed({ name: 'latest.yml', feed: { version: '2.1.0', files: [] }, tag: TAG, uploaded, local })[0], /names no file/);
});

test('a release is made as a draft, its files uploaded, the feed read back from it and checked, and only then published', async () => {
  const dir = buildDir();
  try {
    const gh = fakeGitHub();
    const release = await publishRelease({ client: gh.client, dir, tag: TAG, notes: 'Notes.', commit: 'abc123', log: quiet });
    assert.deepEqual(gh.calls, [
      'list',
      'create draft=true prerelease=false',
      `upload ${PORTABLE}`,
      `upload ${INSTALLER}`,
      `upload ${INSTALLER}.blockmap`,
      'upload latest.yml',
      'download latest.yml',
      `update draft=false tag=${TAG}`,
    ]);
    assert.equal(release.draft, false);
    assert.equal(gh.state.releases[0].target_commitish, 'abc123', 'the tag goes on the commit released');
    assert.equal(gh.state.releases[0].assets.length, 4, 'the older version\'s installer stayed behind');
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a feed that does not describe the installer leaves the release a draft', async () => {
  const dir = buildDir({ feedFor: Buffer.from('an installer from another build') });
  try {
    const gh = fakeGitHub();
    const lines = [];
    await assert.rejects(publishRelease({ client: gh.client, dir, tag: TAG, notes: 'N', log: (s) => lines.push(s) }), /do not agree, so the draft was not published: 2 problems/);
    assert.ok(!gh.calls.some((c) => c.startsWith('update')), 'never published');
    assert.equal(gh.state.releases[0].draft, true);
    assert.ok(lines.some((l) => /size|bytes/.test(l)) && lines.some((l) => /SHA-512/.test(l)), lines.join('\n'));
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a feed for another version, or no blockmap, also stops it', async () => {
  for (const [opts, pattern] of [[{ feedVersion: '2.0.9' }, /1 problem$/], [{ blockmap: false }, /1 problem$/]]) {
    const dir = buildDir(opts);
    try {
      const gh = fakeGitHub();
      await assert.rejects(publishRelease({ client: gh.client, dir, tag: TAG, notes: 'N', log: quiet }), pattern);
      assert.equal(gh.state.releases[0].draft, true);
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  }
});

test('a build with no update feed is refused before anything is made', async () => {
  const dir = buildDir({ feed: false });
  try {
    const gh = fakeGitHub();
    await assert.rejects(publishRelease({ client: gh.client, dir, tag: TAG, notes: 'N', log: quiet }), /no update feed/);
    assert.deepEqual(gh.calls, []);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('an upload that fails leaves a draft; the next run finds it, uploads only what is missing, and publishes', async () => {
  const dir = buildDir();
  try {
    const gh = fakeGitHub({ failUpload: `${INSTALLER}.blockmap` });
    await assert.rejects(publishRelease({ client: gh.client, dir, tag: TAG, notes: 'N', log: quiet }), /502/);
    assert.equal(gh.state.releases[0].draft, true, 'nobody saw a release with no blockmap and no feed');
    gh.calls.length = 0;
    await publishRelease({ client: gh.client, dir, tag: TAG, notes: 'N', log: quiet });
    assert.deepEqual(gh.calls, ['list', `upload ${INSTALLER}.blockmap`, 'upload latest.yml', 'download latest.yml', `update draft=false tag=${TAG}`]);
    assert.equal(gh.state.releases.length, 1, 'the same release, not a second');
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a file left half-uploaded, or at another size, is removed and uploaded again', async () => {
  const dir = buildDir();
  try {
    const gh = fakeGitHub({
      releases: [{ id: 7, tag_name: TAG, draft: true, html_url: 'h', upload_url: 'u{?name}', assets: [
        { id: 1, name: INSTALLER, size: 3, state: 'uploaded', bytes: Buffer.from('abc') },
        { id: 2, name: PORTABLE, size: 0, state: 'starter', bytes: Buffer.alloc(0) },
      ] }],
    });
    await publishRelease({ client: gh.client, dir, tag: TAG, notes: 'N', log: quiet });
    assert.ok(gh.calls.includes(`delete ${INSTALLER}`) && gh.calls.includes(`delete ${PORTABLE}`));
    assert.ok(gh.calls.indexOf(`upload ${INSTALLER}`) > gh.calls.indexOf(`delete ${INSTALLER}`));
    assert.equal(gh.state.releases[0].draft, false);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a release already published is mended and checked, not taken back to a draft', async () => {
  const dir = buildDir();
  try {
    const gh = fakeGitHub({ releases: [{ id: 9, tag_name: TAG, draft: false, html_url: 'h', upload_url: 'u{?name}', assets: [] }] });
    await publishRelease({ client: gh.client, dir, tag: TAG, notes: 'N', log: quiet });
    assert.ok(!gh.calls.some((c) => c.startsWith('create') || c.startsWith('update')));
    assert.equal(gh.state.releases[0].assets.length, 4);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a release is a full release, and a pre-release only when a caller asks for one', async () => {
  const dir = buildDir();
  try {
    const gh = fakeGitHub();
    await publishRelease({ client: gh.client, dir, tag: TAG, notes: 'N', prerelease: true, log: quiet });
    assert.ok(gh.calls.includes('create draft=true prerelease=true'));
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

// A run against a release that already has files: what is there is held to what was built.
const inRelease = (id, dir, name, over = {}) => {
  const bytes = over.bytes ?? fs.readFileSync(path.join(dir, name));
  return { id, name, size: bytes.length, state: 'uploaded', bytes, ...over };
};

test('a file of another size in a published release stops the run, and nothing there is touched', async () => {
  const dir = buildDir();
  try {
    const gh = fakeGitHub({ releases: [{ id: 9, tag_name: TAG, draft: false, html_url: 'h', upload_url: 'u{?name}', assets: [inRelease(1, dir, INSTALLER, { bytes: Buffer.from('what people are downloading') })] }] });
    await assert.rejects(publishRelease({ client: gh.client, dir, tag: TAG, notes: 'N', log: quiet }), /never changed/);
    assert.ok(!gh.calls.some((c) => c.startsWith('delete') || c.startsWith('upload')), gh.calls.join());
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a feed of the same length with another hash in it stops a published release too', async () => {
  const dir = buildDir();
  try {
    const feed = fs.readFileSync(path.join(dir, 'latest.yml'), 'utf8');
    const other = Buffer.from(feed.replace(/sha512: (.)/, (_, c) => `sha512: ${c === 'A' ? 'B' : 'A'}`));
    const gh = fakeGitHub({ releases: [{ id: 9, tag_name: TAG, draft: false, html_url: 'h', upload_url: 'u{?name}', assets: [inRelease(1, dir, 'latest.yml', { bytes: other })] }] });
    await assert.rejects(publishRelease({ client: gh.client, dir, tag: TAG, notes: 'N', log: quiet }), /never changed/);
    assert.ok(!gh.calls.some((c) => c.startsWith('delete')));
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a rebuild over a draft replaces a same-length feed and an installer whose digest differs', async () => {
  const dir = buildDir();
  try {
    const feed = fs.readFileSync(path.join(dir, 'latest.yml'), 'utf8');
    const other = Buffer.from(feed.replace(/sha512: (.)/, (_, c) => `sha512: ${c === 'A' ? 'B' : 'A'}`));
    const gh = fakeGitHub({ releases: [{ id: 9, tag_name: TAG, draft: true, target_commitish: 'old', html_url: 'h', upload_url: 'u{?name}', assets: [
      inRelease(1, dir, 'latest.yml', { bytes: other }),
      inRelease(2, dir, INSTALLER, { digest: `sha256:${'0'.repeat(64)}` }),
    ] }] });
    await publishRelease({ client: gh.client, dir, tag: TAG, notes: 'N', commit: 'new', leaveDraft: true, log: quiet });
    assert.ok(gh.calls.includes('delete latest.yml'));
    assert.ok(gh.calls.includes(`delete ${INSTALLER}`));
    assert.equal(gh.state.releases[0].assets.find((a) => a.name === 'latest.yml').bytes.toString(), feed);
    assert.equal(gh.state.releases[0].target_commitish, 'new', 'and the draft tags this build');
    assert.equal(gh.state.releases[0].draft, true);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a file already there with the right digest is left alone', async () => {
  const dir = buildDir();
  try {
    const bytes = fs.readFileSync(path.join(dir, INSTALLER));
    const digest = `sha256:${crypto.createHash('sha256').update(bytes).digest('hex')}`;
    const gh = fakeGitHub({ releases: [{ id: 9, tag_name: TAG, draft: true, html_url: 'h', upload_url: 'u{?name}', assets: [inRelease(2, dir, INSTALLER, { digest })] }] });
    await publishRelease({ client: gh.client, dir, tag: TAG, notes: 'N', leaveDraft: true, log: quiet });
    assert.ok(!gh.calls.includes(`delete ${INSTALLER}`));
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a tag that is not the version in package.json is refused before anything else', () => {
  const r = spawnSync(process.execPath, ['tools/publish-release.js', 'v0.0.1'], { cwd: new URL('..', import.meta.url), encoding: 'utf8', env: { ...process.env, GH_TOKEN: '' } });
  assert.notEqual(r.status, 0);
  assert.match(r.stdout + r.stderr, /must be the same/);
});

test('a file of the right size with no digest to hold it to goes up again', async () => {
  const dir = buildDir();
  try {
    const gh = fakeGitHub({ releases: [{ id: 9, tag_name: TAG, draft: true, html_url: 'h', upload_url: 'u{?name}', body: 'N', assets: [inRelease(2, dir, INSTALLER)] }] });
    await publishRelease({ client: gh.client, dir, tag: TAG, notes: 'N', log: quiet });
    assert.ok(gh.calls.includes(`delete ${INSTALLER}`));
    assert.ok(gh.calls.includes(`upload ${INSTALLER}`));
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a draft left by an earlier build takes the notes of this build', async () => {
  const dir = buildDir();
  try {
    const gh = fakeGitHub({ releases: [{ id: 9, tag_name: TAG, draft: true, html_url: 'h', upload_url: 'u{?name}', body: 'old words', assets: [] }] });
    await publishRelease({ client: gh.client, dir, tag: TAG, notes: 'new words', log: quiet });
    assert.equal(gh.state.releases[0].body, 'new words');
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a draft moved to a new commit loses the old files before it says so', async () => {
  const order = [];
  const draft = { id: 1, tag_name: 'v1.2.3', draft: true, body: 'n', target_commitish: 'OLD' };
  const client = {
    listReleases: async () => [draft],
    listAssets: async () => [{ name: 'old.exe', state: 'uploaded', size: 1 }],
    deleteAsset: async (_r, a) => { order.push(`delete ${a.name}`); },
    updateRelease: async (_r, patch) => { order.push(`target ${patch.target_commitish}`); throw new Error('a run that dies here'); },
  };
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-move-'));
  fs.writeFileSync(path.join(dir, 'latest.yml'), 'version: 1.2.3\n');
  fs.writeFileSync(path.join(dir, 'Rutba-Office-1.2.3-win-x64.exe'), 'x');
  await assert.rejects(publishRelease({ client, dir, tag: 'v1.2.3', notes: 'n', commit: 'NEW', log: () => {} }), /a run that dies here/);
  assert.deepEqual(order, ['delete old.exe', 'target NEW']);
  fs.rmSync(dir, { recursive: true, force: true });
});

test('the record the gate writes does not stop the publish', () => {
  const note = 'x\n\n## Gate\n\nRun on Windows.\n\nCommit checked: ' + 'a'.repeat(40) + '\n';
  assert.equal(gateProblem({ note, tag: 'v1.2.3', changed: ['README.md', 'tools/gate-floors.json', 'docs/releases/v1.2.3.md'] }), null);
  assert.match(gateProblem({ note, tag: 'v1.2.3', changed: ['tools/gate.js'] }), /changed since/);
});
