// Removing a tree without ever going through a link (D-2026-11-13-37):
// packages/office-shell/src/remove-tree.js.
//
// The case that matters is Windows' junction, which Electron 44's Node 24
// follows in a recursive fs.rm. So this file runs on the Node the tests run
// on and on Office's own Electron (tools/test-removal-runtimes.mjs, with
// ELECTRON_RUN_AS_NODE), and it imports nothing but the helper and Node
// itself, so either can run it from a bare checkout. Elsewhere a junction is
// made as a folder symlink, which the walk treats the same way.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { removeTree, removeTreeSync } from '../packages/office-shell/src/remove-tree.js';

const WINDOWS = process.platform === 'win32';
const runtime = process.versions.electron ? `Electron ${process.versions.electron} (Node ${process.versions.node})` : `Node ${process.versions.node}`;

/** A scratch place, removed after the test. Every junction made here points inside it, so nothing outside can go. */
function scratch(t) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-remove-tree-'));
  t.after(() => fs.rmSync(dir, { recursive: true, force: true, maxRetries: 3 }));
  return dir;
}

/** What a person keeps somewhere else: a folder with a file and a folder in it. */
function treasure(base) {
  const keep = path.join(base, 'keep');
  fs.mkdirSync(path.join(keep, 'inner'), { recursive: true });
  fs.writeFileSync(path.join(keep, 'precious.txt'), 'mine');
  fs.writeFileSync(path.join(keep, 'inner', 'deep.txt'), 'mine too');
  return keep;
}
const intact = (keep) => {
  assert.equal(fs.readFileSync(path.join(keep, 'precious.txt'), 'utf8'), 'mine');
  assert.equal(fs.readFileSync(path.join(keep, 'inner', 'deep.txt'), 'utf8'), 'mine too');
};

const junction = (target, at) => fs.symlinkSync(target, at, 'junction');
/** A symbolic link, where this machine allows one: Windows asks for Developer Mode or the right to make them. */
function trySymlink(target, at, type) {
  try {
    fs.symlinkSync(target, at, type);
    return true;
  } catch (err) {
    if (WINDOWS && (err.code === 'EPERM' || err.code === 'EACCES')) return false;
    throw err;
  }
}

const exists = (p) => {
  try {
    fs.lstatSync(p);
    return true;
  } catch {
    return false;
  }
};

for (const [how, remove] of [['at once', (p, o) => removeTreeSync(p, o)], ['in the background', (p, o) => removeTree(p, o)]]) {
  test(`${how}: a folder and all in it goes, deep folders and files too (${runtime})`, async (t) => {
    const base = scratch(t);
    const tree = path.join(base, 'tree');
    fs.mkdirSync(path.join(tree, 'a', 'b', 'c'), { recursive: true });
    for (let i = 0; i < 30; i++) fs.writeFileSync(path.join(tree, 'a', `f${i}.txt`), String(i));
    fs.writeFileSync(path.join(tree, 'a', 'b', 'c', 'deep.bin'), Buffer.alloc(1024));
    await remove(tree);
    assert.equal(exists(tree), false);
  });

  test(`${how}: a junction inside is removed as itself, and what it points at is left whole (${runtime})`, async (t) => {
    const base = scratch(t);
    const keep = treasure(base);
    const tree = path.join(base, 'tree');
    fs.mkdirSync(path.join(tree, 'sub'), { recursive: true });
    fs.writeFileSync(path.join(tree, 'sub', 'own.txt'), 'x');
    junction(keep, path.join(tree, 'link-here'));
    junction(keep, path.join(tree, 'sub', 'link-deeper'));
    await remove(tree);
    assert.equal(exists(tree), false);
    intact(keep);
  });

  test(`${how}: a path that is itself a junction is unlinked, never followed (${runtime})`, async (t) => {
    const base = scratch(t);
    const keep = treasure(base);
    const link = path.join(base, 'link');
    junction(keep, link);
    await remove(link);
    assert.equal(exists(link), false);
    intact(keep);
  });

  test(`${how}: a junction to the tree itself, or to a folder above it, ends nothing but itself (${runtime})`, async (t) => {
    const base = scratch(t);
    const keep = treasure(base);
    const tree = path.join(keep, 'inner', 'tree');
    fs.mkdirSync(tree);
    fs.writeFileSync(path.join(tree, 'own.txt'), 'x');
    junction(tree, path.join(tree, 'loop'));
    junction(keep, path.join(tree, 'up'));
    await remove(tree);
    assert.equal(exists(tree), false);
    intact(keep);
  });

  test(`${how}: a junction whose folder is gone goes too (${runtime})`, async (t) => {
    const base = scratch(t);
    const gone = path.join(base, 'gone');
    fs.mkdirSync(gone);
    const tree = path.join(base, 'tree');
    fs.mkdirSync(tree);
    junction(gone, path.join(tree, 'dangling'));
    fs.rmdirSync(gone);
    await remove(tree);
    assert.equal(exists(tree), false);
  });

  test(`${how}: symbolic links to a folder and to a file are removed as themselves (${runtime})`, async (t) => {
    const base = scratch(t);
    const keep = treasure(base);
    const tree = path.join(base, 'tree');
    fs.mkdirSync(tree);
    const dirLink = trySymlink(keep, path.join(tree, 'dir-link'), 'dir');
    const fileLink = trySymlink(path.join(keep, 'precious.txt'), path.join(tree, 'file-link'), 'file');
    if (!dirLink && !fileLink) t.diagnostic('this machine makes no symbolic links (no Developer Mode): only junctions are tried');
    await remove(tree);
    assert.equal(exists(tree), false);
    intact(keep);
    if (dirLink) {
      const link = path.join(base, 'own-dir-link');
      trySymlink(keep, link, 'dir');
      await remove(link);
      assert.equal(exists(link), false);
      intact(keep);
    }
  });

  test(`${how}: a read-only file, a lone file, and a path that is not there (${runtime})`, async (t) => {
    const base = scratch(t);
    const tree = path.join(base, 'tree');
    fs.mkdirSync(tree);
    const ro = path.join(tree, 'read-only.txt');
    fs.writeFileSync(ro, 'x');
    fs.chmodSync(ro, 0o444);
    await remove(tree);
    assert.equal(exists(tree), false);
    const lone = path.join(base, 'lone.txt');
    fs.writeFileSync(lone, 'x');
    await remove(lone);
    assert.equal(exists(lone), false);
    await remove(path.join(base, 'never-there'));
    await remove(path.join(base, 'never-there', 'nor-this'));
    await remove(path.join(lone, 'under-a-file'), { maxRetries: 3, retryDelay: 1 });
  });

  test(`${how}: a root, an empty path or no path is refused before anything goes (${runtime})`, async () => {
    const root = path.parse(process.cwd()).root;
    await assert.rejects(async () => remove(root), /refusing to remove a root/);
    await assert.rejects(async () => remove(''), /a path is needed/);
    await assert.rejects(async () => remove(undefined), /a path is needed/);
  });
}

test(`what fs.rmSync itself does with a junction on this runtime, for the record (${runtime})`, (t) => {
  const base = scratch(t);
  const keep = treasure(base);
  const tree = path.join(base, 'tree');
  fs.mkdirSync(tree);
  junction(keep, path.join(tree, 'link'));
  fs.rmSync(tree, { recursive: true, force: true });
  const followed = !exists(path.join(keep, 'precious.txt'));
  // Not a pass or a fault of Office's: the reason removeTree exists, written down where it was seen.
  t.diagnostic(`${runtime} on ${process.platform}: fs.rmSync ${followed ? 'FOLLOWS' : 'does not follow'} a junction inside the folder it removes`);
});
