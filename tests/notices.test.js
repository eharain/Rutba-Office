// The licences of the packages built into the window code travel with an
// installed copy: one text file beside the program, written by the build from
// what it read (apps/desktop/build/notices.js). Checked on made-up packages,
// on the packages this tree really has, and on what the installers and a
// removal are told about the file.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { tempDir } from './temp-dirs.js';
import { packageDirOf, packagesFrom, licenceTextOf, noticesText, noticesFor, KEPT_TEXTS } from '../apps/desktop/build/notices.js';
import { buildSbom } from '../tools/sbom.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const desktop = path.join(root, 'apps', 'desktop');
const posix = (p) => p.split(path.sep).join('/');

/** A made-up installed package: its package.json and the files given. */
function pkg(dir, json, files = {}) {
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(path.join(dir, 'package.json'), JSON.stringify(json));
  for (const [name, text] of Object.entries(files)) fs.writeFileSync(path.join(dir, name), text);
  return dir;
}

/** A project with: one (LICENSE) → two (LICENCE.md, nested under one) and, optionally, three; @s/four with no licence file. */
function madeUp() {
  const base = tempDir('rutba-notices-');
  const nm = path.join(base, 'node_modules');
  pkg(path.join(nm, 'one'), { name: 'one', version: '1.2.3', license: 'MIT', dependencies: { two: '^2.0.0' }, optionalDependencies: { three: '*' }, peerDependencies: { five: '*' } }, { LICENSE: '\uFEFFThe licence of one.\r\nSecond line.   \r\n\r\n', 'index.js': 'export default 1;' });
  pkg(path.join(nm, 'one', 'node_modules', 'two'), { name: 'two', version: '2.0.0', license: 'BSD-3-Clause' }, { 'LICENCE.md': 'The licence of two.', NOTICE: 'The notice of two.', 'licensed-things.js': 'not a licence' });
  pkg(path.join(nm, 'two'), { name: 'two', version: '9.9.9', license: 'ISC' }, { LICENSE: 'Another two, which one does not use.' });
  pkg(path.join(nm, 'three'), { name: 'three', version: '3.0.0', license: 'MIT' }, { LICENSE: 'The licence of three.' });
  pkg(path.join(nm, 'five'), { name: 'five', version: '5.0.0', license: 'MIT' }, { LICENSE: 'The licence of five.' });
  pkg(path.join(nm, '@s', 'four'), { name: '@s/four', version: '4.0.0', license: '(Apache-2.0 AND MIT)' }, { 'README.md': 'No licence file here.' });
  return { base, nm, kept: path.join(base, 'kept') };
}

test('a file is of the package whose folder holds it, the innermost one, and a file of our own is of none', () => {
  const at = (p) => packageDirOf(path.join(root, ...p.split('/')));
  const dir = (p) => posix(path.join(root, ...p.split('/')));
  assert.equal(at('node_modules/react/index.js'), dir('node_modules/react'));
  assert.equal(at('node_modules/@livekit/protocol/dist/index.mjs'), dir('node_modules/@livekit/protocol'));
  assert.equal(at('node_modules/one/node_modules/two/lib/a.js'), dir('node_modules/one/node_modules/two'));
  assert.equal(at('node_modules/one/node_modules/@s/four/a.js'), dir('node_modules/one/node_modules/@s/four'));
  assert.equal(at('apps/desktop/renderer/index.js'), null);
  assert.equal(at('packages/ocr/vendor/engine.js'), null);
  assert.equal(at('node_modules/stray.js'), null);
  assert.equal(at('node_modules/@s/stray.js'), null);
});

test('the packages are those read and all they depend on, as installed for them, without optional ones or peers', () => {
  const { nm } = madeUp();
  const got = packagesFrom([path.join(nm, 'one')]);
  assert.deepEqual(got.map((p) => `${p.name}@${p.version} ${p.license}`), ['one@1.2.3 MIT', 'two@2.0.0 BSD-3-Clause']);
  // Read twice, or by two of its files, a package is still one.
  assert.equal(packagesFrom([path.join(nm, 'one'), path.join(nm, 'one'), path.join(nm, 'one', 'node_modules', 'two')]).length, 2);
  // An optional one or a peer that was built in was read, and so is an input of its own.
  assert.deepEqual(packagesFrom([path.join(nm, 'three'), path.join(nm, 'one'), path.join(nm, 'five')]).map((p) => p.name), ['five', 'one', 'three', 'two']);
});

test('a dependency that is not installed stops the build: its licence cannot be read', () => {
  const { nm } = madeUp();
  fs.rmSync(path.join(nm, 'one', 'node_modules'), { recursive: true });
  fs.rmSync(path.join(nm, 'two'), { recursive: true });
  assert.throws(() => packagesFrom([path.join(nm, 'one')]), /one@1\.2\.3 depends on two, which is not installed/);
});

test('the text names each package with its version and licence and carries its licence files whole, in one order', () => {
  const { nm } = madeUp();
  const text = noticesText(packagesFrom([path.join(nm, 'three'), path.join(nm, 'one')]));
  assert.ok(text.startsWith('Rutba Office: third-party notices\n'));
  const where = ['one 1.2.3\nLicence: MIT\n', 'The licence of one.\nSecond line.\n', 'three 3.0.0\nLicence: MIT\n', 'The licence of three.', 'two 2.0.0\nLicence: BSD-3-Clause\n', 'The licence of two.\n\nThe notice of two.'].map((s) => text.indexOf(s));
  assert.ok(where.every((n) => n > 0), `each part is there: ${where}`);
  assert.deepEqual(where, [...where].sort((a, b) => a - b), 'by name, each heading before its text');
  assert.ok(!text.includes('Another two'), 'the two that one does not use is not named');
  assert.ok(!text.includes('not a licence'), 'a file that only starts like a licence is not one');
  assert.ok(!/[\r\uFEFF]/.test(text) && !/[ \t]\n/.test(text), 'plain lines, whatever the package wrote');
  assert.ok(text.endsWith('The licence of two.\n\nThe notice of two.\n'));
  assert.equal(noticesText(packagesFrom([path.join(nm, 'one'), path.join(nm, 'three')])), text, 'the same packages give the same text, whatever order they were read in');
});

test('a package with no licence file stops the build, unless its text is kept for that very version', () => {
  const { nm, kept } = madeUp();
  const four = packagesFrom([path.join(nm, '@s', 'four')]);
  assert.throws(() => noticesText(four, kept), /@s\/four 4\.0\.0 is built into the window code and carries no licence file/);
  fs.mkdirSync(path.join(kept, '@s'), { recursive: true });
  fs.writeFileSync(path.join(kept, '@s', 'four@3.9.0.txt'), 'The licence of an older four.');
  assert.throws(() => noticesText(four, kept), /carries no licence file/, 'the text kept for another version does not count');
  fs.writeFileSync(path.join(kept, '@s', 'four@4.0.0.txt'), '   \n');
  assert.throws(() => noticesText(four, kept), /carries no licence file/, 'nor does an empty one');
  fs.writeFileSync(path.join(kept, '@s', 'four@4.0.0.txt'), 'The licence of four, kept.\n');
  assert.ok(noticesText(four, kept).endsWith('@s/four 4.0.0\nLicence: (Apache-2.0 AND MIT)\n' + '='.repeat(78) + '\n\nThe licence of four, kept.\n'));
  // A licence file the package gains is used before the kept text.
  fs.writeFileSync(path.join(nm, '@s', 'four', 'COPYING'), 'The licence of four, its own.');
  assert.equal(licenceTextOf(four[0], kept), 'The licence of four, its own.');
});

test('the build\'s record of what it read decides the packages: our own files and made-up inputs are nobody\'s', () => {
  const { base, kept } = madeUp();
  const meta = [{ inputs: { 'src/index.js': {}, 'node_modules/one/index.js': {}, '<define:process.env>': {}, 'dataurl:x': {} } }, { inputs: { [path.join(base, 'node_modules', 'three', 'index.js')]: {} } }, {}];
  const text = noticesFor(meta, { cwd: base, kept });
  assert.deepEqual([...text.matchAll(/^={78}\n(\S+) /gm)].map((m) => m[1]), ['one', 'three', 'two']);
  assert.ok(noticesFor([{ inputs: { 'src/index.js': {} } }], { cwd: base, kept }).startsWith('Rutba Office: third-party notices\n'), 'a build of our own files alone still says so');
});

test('every package this tree builds into the window code has its licence text, and the bill of materials lists each', () => {
  const dev = Object.keys(JSON.parse(fs.readFileSync(path.join(desktop, 'package.json'), 'utf8')).devDependencies);
  assert.ok(dev.includes('react') && dev.includes('livekit-client'), 'the packages built in are the desktop app\'s devDependencies');
  const packages = packagesFrom(dev.map((name) => path.join(root, 'node_modules', ...name.split('/'))));
  const text = noticesText(packages);
  const listed = new Set(buildSbom().components.map((c) => `${c.name}@${c.version}`));
  for (const p of packages) {
    assert.ok(listed.has(`${p.name}@${p.version}`), `${p.name} ${p.version} is in the notices and not in the bill of materials`);
    assert.ok(licenceTextOf(p).length > 100, `${p.name} has a licence text and not a line`);
  }
  for (const name of ['react', 'react-dom', 'scheduler', 'livekit-client', '@livekit/protocol', '@bufbuild/protobuf', 'webrtc-adapter', 'tslib']) assert.ok(packages.some((p) => p.name === name), `${name} is among them`);
  assert.ok(text.includes('Permission is hereby granted, free of charge'), 'the MIT text');
  assert.ok(text.includes('Apache License') && text.includes('Version 2.0, January 2004'), 'the Apache text');
  assert.ok(text.includes('Redistributions in binary form must reproduce the above'), 'the BSD text');
});

test('a text kept for a package is for a version this tree has, and that package still carries no licence file', () => {
  const kept = [];
  const walk = (dir) => { for (const e of fs.readdirSync(dir, { withFileTypes: true })) (e.isDirectory() ? walk(path.join(dir, e.name)) : kept.push(posix(path.relative(KEPT_TEXTS, path.join(dir, e.name))))); };
  walk(KEPT_TEXTS);
  assert.ok(kept.length >= 1);
  for (const file of kept) {
    const m = /^((?:@[^/]+\/)?[^/@]+)@(\d[^/]*)\.txt$/.exec(file);
    assert.ok(m, `${file} is named <package>@<version>.txt`);
    const dir = path.join(root, 'node_modules', ...m[1].split('/'));
    assert.equal(JSON.parse(fs.readFileSync(path.join(dir, 'package.json'), 'utf8')).version, m[2], `${file} is kept for a version that is not the one installed: remove it`);
    assert.deepEqual(fs.readdirSync(dir).filter((n) => /^(licen[cs]e|copying|notice)/i.test(n)), [], `${m[1]} has a licence file of its own now: remove ${file}`);
  }
});

test('the installers put the file beside the program from where the build writes it, and a removal knows it', async () => {
  const require = createRequire(import.meta.url);
  const config = await require('../apps/desktop/electron-builder.config.cjs')();
  const entry = config.extraFiles.find((e) => e.to === 'THIRD-PARTY-NOTICES.txt');
  assert.ok(entry, 'electron-builder.config.cjs extraFiles');
  const bundle = fs.readFileSync(path.join(desktop, 'build', 'bundle.js'), 'utf8');
  assert.match(bundle, /export const NOTICES = [^\n]*path\.join\(app, 'build', 'meta', 'THIRD-PARTY-NOTICES\.txt'\)/);
  assert.equal(posix(entry.from), posix(path.join(desktop, 'build', 'meta', 'THIRD-PARTY-NOTICES.txt')));
  assert.match(fs.readFileSync(path.join(desktop, 'build', 'dist.js'), 'utf8'), /existsSync\(path\.join\(APP, 'build', 'meta', 'THIRD-PARTY-NOTICES\.txt'\)\)/, 'no installer is made without it');
  const nsh = fs.readFileSync(path.join(desktop, 'resources', 'installer.nsh'), 'utf8');
  assert.ok((nsh.match(/!macro rutbaOurFiles\b[\s\S]*?!macroend/)?.[0] || '').includes('"THIRD-PARTY-NOTICES.txt"'), 'rutbaOurFiles names it');
});

test('a real build writes the file beside its output, never in it, naming what its bundles read', () => {
  const base = tempDir('rutba-notices-build-');
  const out = path.join(base, 'out');
  execFileSync(process.execPath, ['build/bundle.js'], { cwd: desktop, stdio: 'pipe', env: { ...process.env, RUTBA_BUNDLE_OUT: out } });
  const text = fs.readFileSync(path.join(base, 'THIRD-PARTY-NOTICES.txt'), 'utf8');
  const named = [...text.matchAll(/^={78}\n(\S+) \d/gm)].map((m) => m[1]);
  for (const name of ['react', 'react-dom', 'scheduler', 'livekit-client']) assert.ok(named.includes(name), `${name}, which a bundle read, is named`);
  assert.ok(!fs.existsSync(path.join(out, 'THIRD-PARTY-NOTICES.txt')), 'the archive does not carry a second copy');
  // The renderer's record for the packs is still written, by the same build.
  assert.ok(Object.keys(JSON.parse(fs.readFileSync(path.join(base, 'renderer-meta.json'), 'utf8')).outputs).includes('app.js'));
});
