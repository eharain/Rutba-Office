/**
 * Every workspace package the desktop application imports is declared as its
 * dependency.
 *
 * The installer packs the desktop package and its declared dependency tree
 * into an archive, and nothing else. A main-process file that imports a
 * workspace package nobody declared still runs from the repository — Node
 * climbs from the unpacked build to the repository's own node_modules and
 * finds it there — and fails only once installed under Program Files, where
 * there is nothing above it to climb to. 1.8.0 shipped its installer that way:
 * every check passed, the packaged smoke passed from inside the repository,
 * and the installed copy opened with "Cannot find package '@rutba/contacts'".
 * This reads the imports and checks the declaration, because no run inside
 * the repository can.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const DESKTOP = join(ROOT, 'apps', 'desktop');

function files(dir, out = []) {
  for (const name of readdirSync(dir)) {
    const p = join(dir, name);
    if (statSync(p).isDirectory()) {
      if (name !== 'node_modules' && name !== 'build' && name !== 'release') files(p, out);
    } else if (p.endsWith('.js')) out.push(p);
  }
  return out;
}

/** The `@rutba/<name>` packages a file imports, by package name. */
function imported(file) {
  const text = readFileSync(file, 'utf8');
  const names = new Set();
  for (const m of text.matchAll(/(?:from|import)\s*\(?\s*['"]@rutba\/([a-z0-9-]+)(?:\/[^'"]*)?['"]/g)) names.add(`@rutba/${m[1]}`);
  return names;
}

test('every @rutba package the desktop application imports is a declared dependency', () => {
  const pkg = JSON.parse(readFileSync(join(DESKTOP, 'package.json'), 'utf8'));
  const declared = new Set(Object.keys(pkg.dependencies || {}));
  const missing = new Map();
  for (const dir of ['main', 'renderer']) {
    for (const file of files(join(DESKTOP, dir))) {
      for (const name of imported(file)) {
        if (!declared.has(name)) missing.set(name, [...(missing.get(name) || []), file.slice(DESKTOP.length + 1)]);
      }
    }
  }
  assert.deepEqual(
    [...missing.entries()].map(([name, from]) => `${name} (from ${from.join(', ')})`),
    [],
    'undeclared workspace packages would be left out of the installer'
  );
});

test('every declared @rutba dependency is a workspace package that exists', () => {
  const pkg = JSON.parse(readFileSync(join(DESKTOP, 'package.json'), 'utf8'));
  for (const name of Object.keys(pkg.dependencies || {}).filter((n) => n.startsWith('@rutba/'))) {
    const dir = join(ROOT, 'packages', name.slice('@rutba/'.length));
    assert.ok(statSync(dir).isDirectory(), `${name} is declared but packages/${name.slice('@rutba/'.length)} is not there`);
  }
});
