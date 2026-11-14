/**
 * The languages a release leaves out of the core, written into its update
 * feed (apps/desktop/build/write-left-core.js), so that an installed Office
 * holds the update for a person in one of them even when the packages feed
 * cannot be had (D-2026-11-13-36).
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { contractLeftCore, withLeftCore, CONTRACT } from '../apps/desktop/build/write-left-core.js';
import { parseFeed } from '../tools/publish-release.js';

const FEED = "version: 2.26.0\r\nfiles:\r\n  - url: Rutba-Office-2.26.0-win-x64.exe\r\n    sha512: YCra/x==\r\n    size: 119234685\r\npath: Rutba-Office-2.26.0-win-x64.exe\r\nsha512: YCra/x==\r\nreleaseDate: '2026-11-14T11:10:26.493Z'\r\n";

test('the feed gets one leftCore line, the rest of it as it was, its line endings kept', () => {
  const once = withLeftCore(FEED, ['ar', 'hi', 'ur']);
  assert.equal(once, `${FEED}leftCore: ['ar', 'hi', 'ur']\r\n`);
  // Written again (a build run twice): still one line, the newer list.
  const twice = withLeftCore(once, ['ur']);
  assert.equal(twice, `${FEED}leftCore: ['ur']\r\n`);
  assert.equal(withLeftCore('version: 1.0.0\n\n', []), "version: 1.0.0\nleftCore: []\n");
  // Read back as publish-release.js reads it, and the files as before.
  const feed = parseFeed(once);
  assert.deepEqual(feed.leftCore, ['ar', 'hi', 'ur']);
  assert.deepEqual(feed.files, [{ url: 'Rutba-Office-2.26.0-win-x64.exe', sha512: 'YCra/x==', size: 119234685 }]);
  assert.equal(feed.version, '2.26.0');
});

test('the list is the contract\'s languages.leftCore, sorted, each a language tag', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-left-core-'));
  try {
    const file = (languages) => {
      const f = path.join(dir, `c${Math.random().toString(36).slice(2)}.json`);
      fs.writeFileSync(f, JSON.stringify({ languages }));
      return f;
    };
    assert.deepEqual(contractLeftCore(file({ builtIn: [], leftCore: ['ur', 'ar', 'hi', 'ur'] })), ['ar', 'hi', 'ur']);
    assert.deepEqual(contractLeftCore(file({ builtIn: ['ar'], leftCore: [] })), []);
    for (const bad of [undefined, 'ur', ['ur-PK'], ['UR'], [7]]) {
      assert.throws(() => contractLeftCore(file({ leftCore: bad })), /languages\.leftCore is not a list of language tags/, JSON.stringify(bad));
    }
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
  // The real contract reads.
  assert.ok(Array.isArray(contractLeftCore(CONTRACT)));
});
