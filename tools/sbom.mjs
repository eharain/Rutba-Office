// The software bill of materials: everything the installers carry, in
// CycloneDX 1.5 JSON, with each part's version, licence and hash.
//
//   node tools/sbom.mjs                  apps/desktop/release/Rutba-Office-<version>-sbom.cdx.json
//   node tools/sbom.mjs --out <file>     somewhere else
//   node tools/sbom.mjs --compare <a> <b>  what differs between two (docs/REPRODUCIBLE.md)
//
// What it lists:
//   - the application, and its own files as the installer packs them
//     (main/, the renderer bundle in build/out/, the icons), each with its
//     SHA-256, so two builds of one commit can be compared file by file;
//   - Electron, the runtime every copy runs on;
//   - every npm package the installer ships: the desktop app's
//     dependencies and theirs, the suite's own @rutba packages among them,
//     resolved as package-lock.json resolves them (what `npm ci` installs is
//     what electron-builder packs), with the lock's licence and SHA-512;
//   - the packages the renderer bundle is built from (React), which ship
//     inside build/out/app.js rather than as packages;
//   - the files vendored rather than installed: pdf.js and its fonts, wasm
//     decoders and colour profile (apps/desktop/vendor/pdfjs), and the text
//     recognition engine and its six language models (packages/ocr/vendor),
//     each file with its licence and SHA-256.
//
// The same commit gives the same document: the time is the commit's
// (SOURCE_DATE_EPOCH when it is set), the serial number is made from the
// content, and every list is sorted.

import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const desktop = path.join(root, 'apps', 'desktop');

const sha256 = (file) => crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');
const rel = (file) => path.relative(root, file).split(path.sep).join('/');

/**
 * The licence of every vendored file, by where it sits: the file and the
 * licence file beside it that says so. tests/sbom.test.js holds every
 * vendored file to having a row here.
 */
export const VENDORED = [
  { group: 'pdf.js', dir: 'apps/desktop/vendor/pdfjs', match: /^(pdf\.min\.mjs|pdf\.worker\.min\.mjs|LICENSE|VERSION)$/, license: 'Apache-2.0', licenseFile: 'LICENSE', supplier: 'Mozilla Foundation' },
  { group: 'pdf.js colour profile', dir: 'apps/desktop/vendor/pdfjs/iccs', match: /./, license: 'CC0-1.0', licenseFile: 'LICENSE', supplier: 'Mozilla Foundation' },
  { group: 'Foxit fonts (PDFium)', dir: 'apps/desktop/vendor/pdfjs/standard_fonts', match: /^(Foxit.*\.pfb|LICENSE_FOXIT)$/, license: 'BSD-3-Clause', licenseFile: 'LICENSE_FOXIT', supplier: 'PDFium Authors' },
  { group: 'Liberation fonts', dir: 'apps/desktop/vendor/pdfjs/standard_fonts', match: /^(Liberation.*\.ttf|LICENSE_LIBERATION)$/, license: 'OFL-1.1', licenseFile: 'LICENSE_LIBERATION', supplier: 'Red Hat, Inc.' },
  { group: 'JBIG2 decoder (PDFium, built for pdf.js)', dir: 'apps/desktop/vendor/pdfjs/wasm', match: /^(jbig2\.wasm|LICENSE_JBIG2|LICENSE_PDFJS_JBIG2)$/, license: 'BSD-3-Clause AND Apache-2.0', licenseFile: 'LICENSE_JBIG2', supplier: 'PDFium Authors; Mozilla Foundation' },
  { group: 'OpenJPEG decoder (built for pdf.js)', dir: 'apps/desktop/vendor/pdfjs/wasm', match: /^(openjpeg\.wasm|LICENSE_OPENJPEG|LICENSE_PDFJS_OPENJPEG)$/, license: 'BSD-2-Clause', licenseFile: 'LICENSE_OPENJPEG', supplier: 'OpenJPEG authors; Mozilla Foundation' },
  { group: 'qcms colour management (built for pdf.js)', dir: 'apps/desktop/vendor/pdfjs/wasm', match: /^(qcms_bg\.wasm|LICENSE_QCMS|LICENSE_PDFJS_QCMS)$/, license: 'MIT AND BSD-2-Clause', licenseFile: 'LICENSE_QCMS', supplier: 'Mozilla Corporation' },
  { group: 'Tesseract engine (tesseract.js-core)', dir: 'packages/ocr/vendor/core', match: /./, license: 'Apache-2.0', licenseFile: 'LICENSE', supplier: 'tesseract.js contributors' },
  { group: 'Tesseract language models (tessdata_fast)', dir: 'packages/ocr/vendor/tessdata', match: /./, license: 'Apache-2.0', licenseFile: 'LICENSE', supplier: 'Tesseract OCR contributors' },
];

/** Every vendored file, each with the row that licenses it (null when none does). */
export function vendoredFiles(base = root) {
  const out = [];
  for (const dir of ['apps/desktop/vendor', 'packages/ocr/vendor']) {
    const walk = (at) => {
      for (const e of fs.readdirSync(path.join(base, at), { withFileTypes: true })) {
        const p = `${at}/${e.name}`;
        if (e.isDirectory()) walk(p);
        else out.push({ file: p, row: VENDORED.find((v) => v.dir === at && v.match.test(e.name)) || null });
      }
    };
    if (fs.existsSync(path.join(base, dir))) walk(dir);
  }
  return out.sort((a, b) => a.file.localeCompare(b.file));
}

/** The version a vendored group says it is, from its VERSION file. */
function vendoredVersion(dir) {
  const file = path.join(root, dir, 'VERSION');
  if (!fs.existsSync(file)) return null;
  const first = fs.readFileSync(file, 'utf8').split(/\r?\n/)[0];
  return /(\d+\.\d+\.\d+)/.exec(first)?.[1] || null;
}

/** A licence as package.json or the lock writes it, as an SPDX expression. */
export function spdxOf(license) {
  if (!license) return null;
  if (typeof license === 'string') return license;
  if (Array.isArray(license)) return license.map((l) => spdxOf(l)).filter(Boolean).join(' OR ') || null;
  return license.type || null;
}

/** Where a package named `name` is found from the lock entry at `from`, as Node would find it. */
export function resolveInLock(packages, from, name) {
  let base = from;
  for (;;) {
    const candidate = `${base ? `${base}/` : ''}node_modules/${name}`;
    if (packages[candidate]) return candidate;
    if (!base) return null;
    const at = base.lastIndexOf('/node_modules/');
    base = at >= 0 ? base.slice(0, at) : '';
  }
}

/**
 * The packages shipped: the closure of `dependencies` (and installed
 * optional ones) from the lock entry `start`, links followed to the
 * workspace package they name. Answers [{ key, name, version, license,
 * integrity, workspace }], one per name and version, sorted.
 */
export function packageClosure(lock, start) {
  const packages = lock.packages;
  const seen = new Set();
  const found = new Map();
  const visit = (key) => {
    let entry = packages[key];
    let at = key;
    if (entry?.link) {
      at = entry.resolved;
      entry = packages[at];
    }
    if (!entry || seen.has(at)) return;
    seen.add(at);
    const name = entry.name || at.slice(at.lastIndexOf('node_modules/') + 'node_modules/'.length);
    if (at !== start) {
      const id = `${name}@${entry.version}`;
      if (!found.has(id)) found.set(id, { key: at, name, version: entry.version, license: spdxOf(entry.license), integrity: entry.integrity || null, workspace: !at.includes('node_modules/') });
    }
    for (const dep of Object.keys({ ...(entry.dependencies || {}), ...(entry.optionalDependencies || {}) })) {
      const where = resolveInLock(packages, at, dep);
      if (where) visit(where);
    }
  };
  visit(start);
  return [...found.values()].sort((a, b) => a.name.localeCompare(b.name) || String(a.version).localeCompare(String(b.version)));
}

/** A lock integrity (sha512-<base64>) as a CycloneDX hash. */
const hashOfIntegrity = (integrity) => {
  const m = /^sha512-(.+)$/.exec(integrity || '');
  return m ? [{ alg: 'SHA-512', content: Buffer.from(m[1], 'base64').toString('hex') }] : [];
};

const licenses = (expr) => {
  if (!expr) return [];
  if (/\s(AND|OR|WITH)\s/.test(expr)) return [{ expression: expr }];
  // A single identifier is an id; anything else a lock carries ("SEE LICENSE IN ...") is a name.
  return /^[A-Za-z0-9.+-]+$/.test(expr) ? [{ license: { id: expr } }] : [{ license: { name: expr } }];
};

/** The application's own files the installer packs (electron-builder.config.cjs `files`). */
function ownFiles() {
  const out = [];
  const walk = (dir) => {
    if (!fs.existsSync(dir)) return;
    for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
      const p = path.join(dir, e.name);
      if (e.isDirectory()) {
        if (!['test', 'tests'].includes(e.name)) walk(p);
      } else if (!/\.(map|md)$/.test(e.name)) out.push(p);
    }
  };
  walk(path.join(desktop, 'main'));
  walk(path.join(desktop, 'build', 'out'));
  for (const f of ['resources/icon.png', 'resources/icon.ico', 'package.json']) if (fs.existsSync(path.join(desktop, f))) out.push(path.join(desktop, f));
  const tiles = path.join(desktop, 'resources', 'apps');
  if (fs.existsSync(tiles)) for (const f of fs.readdirSync(tiles)) if (f.endsWith('.ico')) out.push(path.join(tiles, f));
  return out.sort((a, b) => rel(a).localeCompare(rel(b)));
}

/** When the commit was made, or SOURCE_DATE_EPOCH. */
function buildTime() {
  if (process.env.SOURCE_DATE_EPOCH) return new Date(Number(process.env.SOURCE_DATE_EPOCH) * 1000).toISOString();
  const r = spawnSync('git', ['log', '-1', '--format=%cI'], { cwd: root, encoding: 'utf8' });
  return r.status === 0 && r.stdout.trim() ? new Date(r.stdout.trim()).toISOString() : new Date(0).toISOString();
}

function commitOf() {
  const r = spawnSync('git', ['rev-parse', 'HEAD'], { cwd: root, encoding: 'utf8' });
  return r.status === 0 ? r.stdout.trim() : null;
}

/** The bill of materials, as an object. */
export function buildSbom({ lock = JSON.parse(fs.readFileSync(path.join(root, 'package-lock.json'), 'utf8')) } = {}) {
  const app = JSON.parse(fs.readFileSync(path.join(desktop, 'package.json'), 'utf8'));
  const top = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
  const electron = lock.packages['node_modules/electron'];
  const components = [];

  components.push({
    type: 'framework',
    'bom-ref': `pkg:npm/electron@${electron.version}`,
    name: 'electron',
    version: electron.version,
    purl: `pkg:npm/electron@${electron.version}`,
    licenses: licenses(spdxOf(electron.license) || 'MIT'),
    description: 'The runtime (Chromium and Node.js); Chromium\'s own notices ship beside the program as LICENSES.chromium.html.',
    hashes: hashOfIntegrity(electron.integrity),
  });

  for (const p of packageClosure(lock, 'apps/desktop')) {
    const ref = p.workspace ? `workspace:${p.name}@${p.version}` : `pkg:npm/${p.name}@${p.version}`;
    components.push({
      type: 'library',
      'bom-ref': ref,
      name: p.name,
      version: p.version,
      ...(p.workspace ? {} : { purl: ref }),
      licenses: licenses(p.license),
      ...(p.workspace ? { description: `Part of Rutba Office (${p.key})` } : {}),
      hashes: hashOfIntegrity(p.integrity),
      properties: [{ name: 'rutba:ships', value: p.workspace ? 'workspace package, packed into app.asar' : 'npm package, packed into app.asar' }],
    });
  }

  // Bundled into the renderer by esbuild: devDependencies only in name.
  const bundled = { packages: { ...lock.packages, 'apps/desktop': { dependencies: lock.packages['apps/desktop'].devDependencies } } };
  for (const p of packageClosure(bundled, 'apps/desktop')) {
    const ref = `pkg:npm/${p.name}@${p.version}`;
    if (components.some((c) => c['bom-ref'] === ref)) continue;
    components.push({
      type: 'library', 'bom-ref': ref, name: p.name, version: p.version, purl: ref, licenses: licenses(p.license), hashes: hashOfIntegrity(p.integrity),
      properties: [{ name: 'rutba:ships', value: 'bundled into build/out/app.js' }],
    });
  }

  const groups = new Map();
  for (const { file, row } of vendoredFiles()) {
    if (!row) throw new Error(`${file} is vendored with no licence in tools/sbom.mjs VENDORED`);
    if (!groups.has(row.group)) groups.set(row.group, { row, files: [] });
    groups.get(row.group).files.push(file);
  }
  for (const { row, files } of [...groups.values()].sort((a, b) => a.row.group.localeCompare(b.row.group))) {
    const version = vendoredVersion(row.dir);
    components.push({
      type: 'library',
      'bom-ref': `vendored:${row.group}`,
      name: row.group,
      ...(version ? { version } : {}),
      supplier: { name: row.supplier },
      licenses: licenses(row.license),
      description: `Vendored in ${row.dir}; licence text in ${row.dir}/${row.licenseFile}`,
      components: files.map((f) => ({ type: 'file', 'bom-ref': `file:${f}`, name: f, hashes: [{ alg: 'SHA-256', content: sha256(path.join(root, f)) }] })),
    });
  }

  const own = ownFiles().map((f) => ({ type: 'file', 'bom-ref': `file:${rel(f)}`, name: rel(f), hashes: [{ alg: 'SHA-256', content: sha256(f) }] }));

  const bom = {
    bomFormat: 'CycloneDX',
    specVersion: '1.5',
    serialNumber: null,
    version: 1,
    metadata: {
      timestamp: buildTime(),
      tools: { components: [{ type: 'application', name: 'tools/sbom.mjs', version: top.version }] },
      component: {
        type: 'application',
        'bom-ref': 'rutba-office',
        name: 'Rutba Office',
        version: app.version,
        supplier: { name: 'Tech Style Ltd', url: ['https://tech-style.co'] },
        licenses: licenses(app.license),
        externalReferences: [{ type: 'vcs', url: 'https://github.com/eharain/Rutba-Office' }, { type: 'website', url: 'https://office.rutba.io' }],
        properties: [{ name: 'rutba:commit', value: commitOf() || 'unknown' }],
        components: own,
      },
    },
    components,
    dependencies: [{ ref: 'rutba-office', dependsOn: components.map((c) => c['bom-ref']) }],
  };
  // The serial number from the content: the same build, the same number.
  const digest = crypto.createHash('sha256').update(JSON.stringify({ ...bom, metadata: { ...bom.metadata, timestamp: null } })).digest('hex');
  bom.serialNumber = `urn:uuid:${digest.slice(0, 8)}-${digest.slice(8, 12)}-5${digest.slice(13, 16)}-${((parseInt(digest[16], 16) & 3) | 8).toString(16)}${digest.slice(17, 20)}-${digest.slice(20, 32)}`;
  return bom;
}

/** Every file and package with its hash or version, for comparing two bills. */
export function inventory(bom) {
  const out = new Map();
  const add = (c) => {
    const hash = (c.hashes || []).map((h) => `${h.alg}:${h.content}`).join(' ');
    out.set(c['bom-ref'], `${c.version || ''} ${hash}`.trim());
    for (const child of c.components || []) add(child);
  };
  for (const c of bom.metadata?.component?.components || []) add(c);
  for (const c of bom.components || []) add(c);
  return out;
}

/** What differs between two bills, as sentences. */
export function compareSboms(a, b) {
  const x = inventory(a);
  const y = inventory(b);
  const out = [];
  for (const [ref, v] of x) {
    if (!y.has(ref)) out.push(`only in the first: ${ref}`);
    else if (y.get(ref) !== v) out.push(`differs: ${ref}`);
  }
  for (const ref of y.keys()) if (!x.has(ref)) out.push(`only in the second: ${ref}`);
  return out;
}

function main() {
  const args = process.argv.slice(2);
  if (args[0] === '--compare') {
    const [a, b] = args.slice(1).map((f) => JSON.parse(fs.readFileSync(f, 'utf8')));
    const diff = compareSboms(a, b);
    for (const line of diff) console.log(line);
    console.log(diff.length ? `${diff.length} differences` : 'the same files, packages and versions');
    process.exit(diff.length ? 1 : 0);
  }
  const bom = buildSbom();
  const at = args.indexOf('--out');
  const out = at >= 0 ? path.resolve(args[at + 1]) : path.join(desktop, 'release', `Rutba-Office-${bom.metadata.component.version}-sbom.cdx.json`);
  fs.mkdirSync(path.dirname(out), { recursive: true });
  fs.writeFileSync(out, JSON.stringify(bom, null, 2) + '\n');
  const files = bom.components.reduce((n, c) => n + (c.components?.length || 0), 0) + bom.metadata.component.components.length;
  console.log(`${bom.components.length} components, ${files} files with their hashes: ${out}`);
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) main();
