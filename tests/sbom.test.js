/**
 * The software bill of materials (tools/sbom.mjs): every vendored file has
 * a licence, packages are resolved as the lock resolves them, and the
 * document is CycloneDX, complete, and the same for the same commit.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { VENDORED, vendoredFiles, spdxOf, resolveInLock, packageClosure, buildSbom, compareSboms } from '../tools/sbom.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

test('every vendored file has a licence, and each licence file named is there', () => {
  const files = vendoredFiles();
  assert.ok(files.some((f) => f.file.endsWith('pdf.worker.min.mjs')));
  assert.ok(files.some((f) => f.file.endsWith('urd.traineddata')));
  assert.deepEqual(files.filter((f) => !f.row).map((f) => f.file), [], 'a vendored file with no row in VENDORED');
  for (const row of VENDORED) assert.ok(fs.existsSync(path.join(root, row.dir, row.licenseFile)), `${row.dir}/${row.licenseFile}`);
});

test('a licence reads as an SPDX expression however package.json wrote it', () => {
  assert.equal(spdxOf('MIT'), 'MIT');
  assert.equal(spdxOf({ type: 'ISC' }), 'ISC');
  assert.equal(spdxOf([{ type: 'MIT' }, { type: 'Apache-2.0' }]), 'MIT OR Apache-2.0');
  assert.equal(spdxOf(undefined), null);
});

test('a package is found as Node finds it: nested first, then each node_modules above, links followed', () => {
  const lock = {
    packages: {
      'apps/desktop': { dependencies: { a: '*', '@rutba/x': '*' } },
      'node_modules/a': { version: '1.0.0', license: 'MIT', dependencies: { b: '*' } },
      'node_modules/a/node_modules/b': { version: '2.0.0', license: 'ISC' },
      'node_modules/b': { version: '1.0.0', license: 'MIT' },
      'node_modules/@rutba/x': { resolved: 'packages/x', link: true },
      'packages/x': { name: '@rutba/x', version: '1.0.0', license: 'AGPL-3.0-or-later', dependencies: { b: '*', c: '*' } },
      'node_modules/c': { version: '3.0.0', license: 'MIT', optionalDependencies: { d: '*' } },
      'node_modules/dev-only': { version: '9.9.9', dev: true },
    },
  };
  assert.equal(resolveInLock(lock.packages, 'node_modules/a', 'b'), 'node_modules/a/node_modules/b');
  assert.equal(resolveInLock(lock.packages, 'packages/x', 'b'), 'node_modules/b');
  assert.equal(resolveInLock(lock.packages, 'packages/x', 'nothing'), null);
  const closure = packageClosure(lock, 'apps/desktop');
  assert.deepEqual(closure.map((p) => `${p.name}@${p.version}`), ['@rutba/x@1.0.0', 'a@1.0.0', 'b@1.0.0', 'b@2.0.0', 'c@3.0.0'], 'an optional package not installed is left out, as is what nothing ships');
  assert.equal(closure.find((p) => p.name === '@rutba/x').workspace, true);
});

test('the bill is CycloneDX 1.5 and names Electron, the shipped packages, React in the bundle, the vendored files and the app\'s own files', () => {
  const bom = buildSbom();
  assert.equal(bom.bomFormat, 'CycloneDX');
  assert.equal(bom.specVersion, '1.5');
  assert.match(bom.serialNumber, /^urn:uuid:[0-9a-f]{8}-[0-9a-f]{4}-5[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/);
  const app = JSON.parse(fs.readFileSync(path.join(root, 'apps', 'desktop', 'package.json'), 'utf8'));
  assert.equal(bom.metadata.component.version, app.version);
  const names = bom.components.map((c) => c.name);
  for (const want of ['electron', 'electron-updater', 'imapflow', 'nodemailer', '@rutba/ocr', '@rutba/pdf', 'react', 'react-dom', 'pdf.js', 'Liberation fonts', 'Tesseract language models (tessdata_fast)']) assert.ok(names.includes(want), want);
  assert.ok(!names.includes('electron-builder') && !names.includes('esbuild'), 'tools that build the installer are not in it');
  const pdfjs = bom.components.find((c) => c.name === 'pdf.js');
  assert.equal(pdfjs.version, fs.readFileSync(path.join(root, 'apps/desktop/vendor/pdfjs/VERSION'), 'utf8').match(/\d+\.\d+\.\d+/)[0]);
  assert.ok(pdfjs.components.every((f) => /^[0-9a-f]{64}$/.test(f.hashes[0].content)));
  for (const c of bom.components) assert.ok(c.licenses.length, `${c.name} has a licence`);
  assert.ok(bom.metadata.component.components.some((f) => f.name === 'apps/desktop/main/main.js'));
  assert.ok(!bom.metadata.component.components.some((f) => /\.map$/.test(f.name)));
  const imap = bom.components.find((c) => c.name === 'imapflow');
  assert.equal(imap.hashes[0].alg, 'SHA-512');
  assert.match(imap.hashes[0].content, /^[0-9a-f]{128}$/);
});

test('the same commit gives the same bill, and a changed file is the one difference', () => {
  const a = buildSbom();
  const b = buildSbom();
  assert.deepEqual(a, b);
  assert.deepEqual(compareSboms(a, b), []);
  const changed = structuredClone(b);
  changed.metadata.component.components[0].hashes[0].content = '0'.repeat(64);
  changed.components = changed.components.filter((c) => c.name !== 'imapflow');
  const diff = compareSboms(a, changed);
  assert.ok(diff.includes(`differs: ${a.metadata.component.components[0]['bom-ref']}`), diff.join(' | '));
  assert.ok(diff.some((d) => /^only in the first: pkg:npm\/imapflow@/.test(d)));
});
