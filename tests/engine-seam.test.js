/**
 * The seam between the shared editor packages and the format layer.
 *
 * `@rutba/drawing`, `@rutba/doc-view` and `@rutba/editing` are consumed beyond
 * this suite — the Rutba consumer line takes them by file: link for Workspace,
 * Sign, Studio and comms, and Mail is the built-for HTML-backend consumer — and
 * their documentation promises they carry no format dependency outside a
 * backend. That promise is easy to break by accident and impossible to notice
 * by reading: an import added to a core file still passes every test, because
 * this repo has the dependency available.
 *
 * It broke exactly that way once, in Workspace: pagination needed two helpers
 * that happened to live in `@rutba/ooxml`, so the format-neutral editor core
 * imported the format layer — while a comment three files away still said only
 * the backend did. These tests read the source and check, because nothing else
 * will. Workspace keeps the same guard on its side (tests/shared-seam.test.js),
 * reading the packages through its links; this one sits where the code is.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join, dirname, relative } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');

/** Every .js file under a package's src, recursively. */
function sources(pkg) {
  const out = [];
  const walk = (dir) => {
    for (const entry of readdirSync(dir)) {
      const path = join(dir, entry);
      if (statSync(path).isDirectory()) walk(path);
      else if (entry.endsWith('.js')) out.push(path);
    }
  };
  walk(join(ROOT, 'packages', pkg, 'src'));
  return out;
}

/** The packages a file actually imports — not the ones its prose mentions. */
function imports(path) {
  const source = readFileSync(path, 'utf8');
  const found = new Set();
  const re = /(?:^|\n)\s*(?:import|export)[^'"\n]*from\s+['"]([^'"]+)['"]/g;
  let m;
  while ((m = re.exec(source))) found.add(m[1]);
  for (const dynamic of source.matchAll(/\bimport\(\s*['"]([^'"]+)['"]/g)) found.add(dynamic[1]);
  // A subpath import (`@rutba/drawing/glyphs`) is a dependency on the package,
  // so it is judged as the package.
  return [...new Set([...found]
    .filter((s) => s.startsWith('@rutba/'))
    .map((s) => s.split('/').slice(0, 2).join('/')))];
}

/** What each shared package is allowed to depend on, and nothing more. */
const ALLOWED = {
  drawing: [],
  editing: [],
  // The editor may use the drawing engine's text metrics and the shared undo
  // stack; both are shared packages themselves, so they travel with it. The
  // PDF writer is allowed for the one file that prints a paginated frame
  // (src/export/pdf.js) — a shared package with no format dependency, so the
  // editor still travels whole.
  'doc-view': ['@rutba/drawing', '@rutba/editing', '@rutba/pdf'],
};

for (const [pkg, allowed] of Object.entries(ALLOWED)) {
  test('@rutba/' + pkg + ' core carries no format-layer dependency', () => {
    const offenders = [];
    for (const path of sources(pkg)) {
      // A backend is the format edge — that is the whole point of the port, and
      // the one place a format import belongs.
      if (path.includes(join('src', 'backends'))) continue;
      for (const dependency of imports(path)) {
        if (allowed.includes(dependency)) continue;
        offenders.push(relative(ROOT, path) + ' imports ' + dependency);
      }
    }
    assert.deepEqual(offenders, [],
      'a shared package may not import the format layer outside a backend');
  });
}

test('only a backend may import the format layer', () => {
  const backends = sources('doc-view').filter((p) => p.includes(join('src', 'backends')));
  assert.ok(backends.length >= 2, 'there are at least the ooxml and html backends');

  const html = backends.find((p) => p.endsWith('html.js'));
  assert.deepEqual(imports(html), [],
    'the HTML backend is what Mail uses — it must not pull the OOXML layer in');

  const ooxml = backends.find((p) => p.endsWith('ooxml.js'));
  assert.ok(imports(ooxml).includes('@rutba/ooxml'), 'the OOXML backend is where that import belongs');
});

test('a shared package declares the dependencies it actually has', () => {
  // A manifest that over-declares is as misleading as one that under-declares:
  // it is what made `@rutba/doc-view` look unmovable when its core was clean.
  for (const pkg of Object.keys(ALLOWED)) {
    const manifest = JSON.parse(readFileSync(join(ROOT, 'packages', pkg, 'package.json'), 'utf8'));
    const declared = Object.keys(manifest.dependencies ?? {});
    const used = new Set(sources(pkg).flatMap(imports));
    for (const dependency of declared) {
      assert.ok(used.has(dependency),
        pkg + ' declares ' + dependency + ' but no source imports it');
    }
  }
});

test('every subpath the consumer line imports is still exported', () => {
  // The export maps are the contract the consumer line compiles against. A
  // subpath dropped here is a build that breaks in another repo, so the list
  // of what is actually imported over there is pinned here, where the change
  // would be made.
  const wanted = {
    ooxml: ['.', './build', './pivot', './recalc', './zip', './document', './package', './workbook', './fidelity', './capabilities'],
    formula: ['.', './catalog'],
    'doc-view': ['.', './backends/ooxml', './backends/html', './export/pdf', './bands', './lists', './paginate', './positions'],
    drawing: ['.', './palette', './svg', './scene', './chart', './ooxml', './shapes', './measure'],
    editing: ['.', './history'],
    pdf: ['.', './metrics', './encoding', './document', './sheet'],
    'sheet-view': ['.', './numfmt', './geometry', './selection', './styles-write'],
  };
  for (const [pkg, subpaths] of Object.entries(wanted)) {
    const manifest = JSON.parse(readFileSync(join(ROOT, 'packages', pkg, 'package.json'), 'utf8'));
    for (const subpath of subpaths) {
      assert.ok(manifest.exports && manifest.exports[subpath],
        '@rutba/' + pkg + ' no longer exports ' + subpath + ' — the consumer line imports it');
    }
  }
});
