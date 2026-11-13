// No recursive removal in product code but removeTree (D-2026-11-13-37):
// tools/check-recursive-removals.mjs, on this checkout and on what it must
// and must not take for one.

import test from 'node:test';
import assert from 'node:assert/strict';
import { removalFaults, recursiveRemovals, isProduct, codeOnly, ALLOWED, HELPER } from '../tools/check-recursive-removals.mjs';

test('this checkout removes trees only through removeTree, the Locker sites still listed for their owner', () => {
  assert.deepEqual(removalFaults(), []);
  for (const [rel, entry] of Object.entries(ALLOWED)) {
    assert.match(rel, /^apps\/desktop\/main\/extras\/locker\//, 'only the Locker is still to move');
    assert.ok(entry.count > 0 && entry.owner);
  }
});

test('what is taken for a recursive removal: every way of writing one', () => {
  const lines = (src) => recursiveRemovals(src).map((f) => f.line);
  assert.deepEqual(lines("fs.rmSync(dir, { recursive: true, force: true });"), [1]);
  assert.deepEqual(lines("await fsp.rm(dir, { force: true, recursive: true });"), [1]);
  assert.deepEqual(lines("await fs.promises.rm(at, { recursive: true });"), [1]);
  assert.deepEqual(lines("ops.rm(dir, { recursive: true, force: true });"), [1], 'any object\'s rm');
  assert.deepEqual(lines("fs.rmdirSync(dir, { recursive: true });"), [1]);
  assert.deepEqual(lines("const recursive = true;\nfs.rmSync(dir, { recursive, force: true });"), [2], 'shorthand');
  assert.deepEqual(lines("x();\nfs.rmSync(\n  path.join(root, name(a, b)),\n  {\n    recursive: true,\n  },\n);"), [2], 'over several lines');
  assert.deepEqual(lines("fs.rmSync(dir, OPTIONS);"), [1], 'options that cannot be read where they are passed');
  assert.match(recursiveRemovals("fs.rmSync(dir, OPTIONS);")[0].what, /not written out/);
  // Options that cannot be read where they are written: a spread, a quoted or a computed key.
  for (const src of ["fs.rmSync(dir, { ...OPTS });", "fs.rmSync(dir, { force: true, ...more });", "fs.rmSync(dir, { 'recursive': true });", "fs.rmSync(dir, { \"recursive\": true });", "fs.rmSync(dir, { [KEY]: true });"]) {
    assert.match(recursiveRemovals(src)[0]?.what || '', /not written out/, src);
  }
  // Taken out of fs by name, so the calls go by another name: imported, destructured, or given as a value.
  assert.deepEqual(lines("import { rmSync } from 'node:fs';"), [1]);
  assert.deepEqual(lines("import { readFileSync,\n  rm as remove } from 'fs/promises';"), [1], 'renamed, over two lines');
  assert.deepEqual(lines("import { rmdirSync } from 'original-fs';"), [1]);
  assert.deepEqual(lines("const { rmSync } = fs;"), [1]);
  assert.deepEqual(lines("const { rm: remove } = require('node:fs/promises');"), [1], 'destructured and renamed');
  assert.deepEqual(lines("let { readFileSync, rmdir } = fsp;"), [1]);
  assert.deepEqual(lines("const remove = fs.rmSync;"), [1]);
  assert.deepEqual(lines("const ops = { rename: fs.renameSync, rm: fs.rmSync };"), [1], 'an injected default');
  assert.deepEqual(lines("later(fs.promises.rm, dir);"), [1], 'passed to another function');
  assert.deepEqual(lines("const remove = fs.rmSync.bind(fs);"), [1]);
  assert.match(recursiveRemovals("const { rmSync } = fs;")[0].what, /destructured/);
  assert.match(recursiveRemovals("const remove = fs.rmSync;")[0].what, /taken as a value/);
  assert.deepEqual(lines("import rimraf from 'rimraf';"), [1]);
  assert.deepEqual(lines("const { removeSync } = require('fs-extra');"), [1]);
  assert.deepEqual(lines("const r = await import('rimraf/sync');"), [1]);
});

test('what is not: one file, recursive false, a comment, a message, a regular expression, its own definition', () => {
  const none = (src) => assert.deepEqual(recursiveRemovals(src), [], src);
  none("fs.rmSync(file, { force: true });");
  none("fs.rmSync(file);");
  none("await fsp.rmdir(dir);");
  none("fs.rmSync(dir, { recursive: false, force: true });");
  none("// fs.rmSync(dir, { recursive: true });\nremoveTree(dir);");
  none("/* fsp.rm(x, { recursive: true }) */ removeTree(x);");
  none("throw new Error('fs.rmSync(dir, { recursive: true }) follows a junction');");
  none("const said = `rm(${dir}, { recursive: true })`;");
  none("const re = /rm\\(x, { recursive: true }\\)/;");
  none("function rm(p, o) { return removeTree(p); }");
  // A method's own definition, in a class or an object, is not a call.
  none("class Ops { rm(p, o) { return removeTree(p); } }");
  none("const ops = {\n  rmdir(p, o) {\n    return removeTree(p);\n  },\n};");
  none("class Ops { async rmSync(p, o) { return removeTree(p); } }");
  // fs taken by other names, a comparison, a name in an import that only reads.
  none("import { readFileSync, rmSync as notThis } from './elsewhere.js';");
  none("import { readFileSync } from 'node:fs';");
  none("const { readFileSync } = fs;");
  none("if (typeof fs.rmSync !== 'function') throw new Error('old Node');");
  none("if (remove === fs.rmSync) x();");
  // The limits named in the check's header: not seen, and said so there.
  none("fs['rmSync'](dir, { recursive: true });");
  none("fs.rmSync.call(fs, dir, { recursive: true });");
  none("execFileSync('cmd', ['/c', 'rd', '/s', '/q', dir]);");
  none("const note = 'rimraf is not used here';");
  none("fs.mkdirSync(dir, { recursive: true });");
  none("farm(dir, { recursive: true });");
  // Quotes inside a regular expression do not lose the place: the call after it is still seen.
  assert.deepEqual(recursiveRemovals("const q = /[\"']/;\nfs.rmSync(d, { recursive: true });").map((f) => f.line), [2]);
  assert.equal(codeOnly("a('x\\'y'); // c\nb").split('\n').length, 2, 'lines kept');
});

test('product code is what ships: main, renderer and every package\'s src, less verification runs, tests, tools and the helper', () => {
  for (const rel of ['apps/desktop/main/print.js', 'apps/desktop/renderer/apps/mail.js', 'packages/mailbox/src/store.js', 'packages/office-shell/src/electron/thumbs.js']) assert.equal(isProduct(rel), true, rel);
  for (const rel of ['apps/desktop/main/verify-apps.js', 'apps/desktop/build/bundle.js', 'tools/corpus.js', 'tests/print.test.js', 'packages/pdf/test/x.test.js', HELPER, 'packages/mailbox/src/notes.md']) assert.equal(isProduct(rel), false, rel);
});

test('the list of sites still to move only shrinks: a new one, one moved and its count left, a file gone, all refused', () => {
  const files = {
    'apps/desktop/main/a.js': "fs.rmSync(x, { recursive: true });",
    'apps/desktop/main/b.js': "fs.rmSync(x, { recursive: true });\nfs.rmSync(y, { recursive: true });",
    'apps/desktop/main/c.js': "removeTree(x);",
  };
  const root = '/nowhere';
  const faults = (allowed) => removalFaults({ root, allowed, read: (rel) => files[rel], files: Object.keys(files) });
  assert.deepEqual(faults({ 'apps/desktop/main/a.js': { count: 1, owner: 'A' }, 'apps/desktop/main/b.js': { count: 2, owner: 'B' } }), []);
  assert.match(faults({ 'apps/desktop/main/a.js': { count: 1, owner: 'A' } }).join('\n'), /b\.js:1: a recursive rmSync\(\)/);
  assert.match(faults({ 'apps/desktop/main/a.js': { count: 1, owner: 'A' }, 'apps/desktop/main/b.js': { count: 1, owner: 'B' } }).join('\n'), /b\.js: 2 recursive removals where 1 are listed for B: a new one/);
  assert.match(faults({ 'apps/desktop/main/a.js': { count: 1, owner: 'A' }, 'apps/desktop/main/b.js': { count: 3, owner: 'B' } }).join('\n'), /b\.js: 2 recursive removals left of the 3 listed for B: bring its count/);
  assert.match(faults({ 'apps/desktop/main/a.js': { count: 1, owner: 'A' }, 'apps/desktop/main/b.js': { count: 2, owner: 'B' }, 'apps/desktop/main/c.js': { count: 1, owner: 'C' } }).join('\n'), /c\.js: 0 recursive removals left of the 1 listed for C: bring its count in ALLOWED down, or drop the entry/);
  assert.match(faults({ 'apps/desktop/main/a.js': { count: 1, owner: 'A' }, 'apps/desktop/main/b.js': { count: 2, owner: 'B' }, 'apps/desktop/main/gone.js': { count: 1, owner: 'D' } }).join('\n'), /gone\.js: listed in ALLOWED but not found/);
});
