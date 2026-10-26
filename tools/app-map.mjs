// The map of what each app takes from another and from the shared engines
// (docs/APP-PACKS.md, the plan's Part F item 7): where the core's line falls,
// drawn from the code as it is rather than from what it was meant to be.
//
// The import graph is esbuild's own (the renderer's build resolves the same
// way), walked from each app's entry points: its window module, the lazy
// import renderer/index.js makes, and the main-process service main.js
// builds for its IPC namespace. Every file is given a home: an app's folder
// in the renderer, a main-process file by its name (MAIN_HOMES), or the
// core. A walk does not enter another app's files; the import that would is
// a cross-app import, one of the edges to cut before the code moves.
//
//   node tools/app-map.mjs                the headline, on screen
//   node tools/app-map.mjs --json <file>  the whole map, as JSON
//   node tools/app-map.mjs --doc          the map's tables, into docs/APP-PACKS.md
//   node tools/app-map.mjs --write-edges  the cross-app imports allowed today,
//                                         into tests/fixtures/app-map-edges.json
//
// tests/app-map.test.js reads that fixture: a new cross-app import, or an
// engine one app owned reached by another, fails the test until it is cut or
// written into the fixture on purpose.

import esbuild from 'esbuild';
import fs from 'node:fs';
import path from 'node:path';
import { builtinModules } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';
// Who owns each namespace, event and AI tool: the tables the core itself runs on.
import { PACKS, NAMESPACE_PACK, EVENT_PACK } from '../packages/office-shell/src/contract.js';
import { TOOL_PACK } from '../apps/desktop/main/pack-tools.js';

export const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const DESKTOP = 'apps/desktop';
const MAIN = `${DESKTOP}/main`;
const RENDERER = `${DESKTOP}/renderer`;
export const EDGES_FIXTURE = 'tests/fixtures/app-map-edges.json';
const DOC = 'docs/APP-PACKS.md';

/** The packs, by the keys the registry and the windows already use: the contract's list. */
export { PACKS };
const PACK_IDS = new Set(PACKS.map((p) => p.id));
const isPack = (home) => PACK_IDS.has(home);

/** A main-process file's home, by its name (relative to apps/desktop/main). The rest are the core's. */
const MAIN_HOMES = [
  [/^(verify-|smoke\.js$|screen-|seed-mail\.js$|sample-picture\.js$|upgrade-report\.js$|console-message\.js$)/, 'checks'],
  [/^mail(-[\w-]+)?\.js$|^discover\.js$/, 'mail'],
  [/^calendar\.js$/, 'calendar'],
  [/^contacts\.js$/, 'contacts'],
  [/^base(-[\w-]+)?\.js$|^engines\//, 'base'],
  [/^pdf\.js$/, 'pdf'],
  [/^ocr\.js$/, 'ocr'],
  // The program AI apps run (mcp/node.js) is the core's: it loads AI Tools' pack.
  [/^ai-tools\.js$|^mcp\/(?!node\.js$)/, 'ai'],
  [/^live\.js$/, 'recorder'],
  [/^talk(-[\w-]+)?\.js$/, 'talk'],
  // The old .ppt reader (legacy-deck*.js) and mail merge's sources
  // (mailmerge-source.js) are the document service's, so the core's.
  [/^present\.js$/, 'slides'],
];
/**
 * Engines whose files go each by who reaches them rather than as one: text
 * recognition's engine and data are one pack's, the result helpers the
 * windows read are shared; Data Files' databases are its own, the gallery
 * the core's; Data Files' engine is its own but for the file reader
 * (`@rutba/database/base`) Documents' mail merge reads a database with. The
 * others move whole until the move itself splits them.
 */
const SPLIT_PACKAGES = new Set(['ocr', 'templates', 'database']);

/**
 * Files in the renderer's shared folder that are one app's own. Text
 * recognition's window part (ocr-ui.js) is the core's: the five apps that
 * read text show it, and its calls go to the ocr namespace, which says Text
 * recognition is not installed when it is not.
 */
const RENDERER_HOMES = { 'ai-tools-dialog.js': 'ai' };

/**
 * Windows an app's pack carries that the launcher does not list: the AI Tools
 * window, which the shell opens (its badge, the confirmations and the step
 * pickers in other apps are the core's, in ai-tools.js).
 */
export const PACK_WINDOWS = { ai: `${RENDERER}/ai-tools-dialog.js` };

/** Entry points no import reaches: workers started by name, a program run on its own. */
const EXTRA_ENTRIES = {
  renderer: {
    // build/bundle.js builds it as a bundle of its own (scripts-worker.js).
    sheets: [`${RENDERER}/apps/sheets/scripts-worker.js`],
    ai: [PACK_WINDOWS.ai],
  },
  main: {},
};
const CORE_ENTRIES = {
  // The image worker (image-worker.js, a bundle of its own) is the core's: Image and Pictures' batch conversion both start it.
  renderer: [`${RENDERER}/index.js`, `${RENDERER}/imaging/worker.js`],
  // The AI tools for an AI app on this computer: Rutba Office's program run as Node, which loads AI Tools' pack.
  main: [`${MAIN}/main.js`, 'packages/office-shell/src/electron/preload.js', `${MAIN}/mcp/node.js`],
};
/** The core's own wiring: what it imports from an app here is the registration the loader takes over. */
const WIRING = new Set([`${RENDERER}/index.js`, `${RENDERER}/shell.js`, `${MAIN}/main.js`, `${MAIN}/app-packs.js`]);

/** The window checks, by their names: one pack's when the name says so, the core's otherwise. */
const CHECK_PACK = [
  [/^verify-(word|mailmerge)/, 'word'],
  [/^verify-(sheet|old-sheets|slicers|receipt-row|datatools|analysis|evaluate)/, 'sheets'],
  [/^verify-deck/, 'slides'],
  [/^verify-pdf/, 'pdf'],
  [/^verify-(picture|gallery)/, 'pictures'],
  [/^verify-image/, 'image'],
  [/^verify-(video|intro-video)/, 'video'],
  [/^verify-recorder/, 'recorder'],
  [/^verify-(mail|smime)|^seed-mail/, 'mail'],
  [/^verify-calendar/, 'calendar'],
  [/^verify-base/, 'base'],
  [/^verify-talk/, 'talk'],
  [/^verify-ai/, 'ai'],
  [/^verify-ocr/, 'ocr'],
];

/** The heavy parts, named so their size and owner are read off the map. */
const HEAVY = [
  { what: "Text recognition's engine and language data", dirs: ['packages/ocr/vendor'] },
  { what: 'pdf.js, which draws PDF pages (copied beside the bundle)', dirs: [`${DESKTOP}/vendor/pdfjs`], through: `${RENDERER}/pdfjs.js` },
  { what: "Data Files' ready-made and sample databases", dirs: ['packages/templates/data-files'] },
  { what: 'The template gallery: files and thumbnails', dirs: ['packages/templates/files', 'packages/templates/thumbs'], through: 'packages/templates/src/catalogue.js' },
  { what: 'The words in Urdu, Arabic and Hindi', dirs: ['packages/office-ui/src/catalogues'] },
  { what: 'The studio renderer (layers for Video, Image and Recorder)', dirs: ['packages/studio/src'] },
  { what: "Video's containers and export plan (the codecs are Chromium's own WebCodecs; none ship)", dirs: ['packages/media/src'] },
  { what: "AI Tools' server (the Model Context Protocol, local and over the network)", dirs: [`${MAIN}/mcp`] },
  { what: "Talk's calling and messages", dirs: [`${RENDERER}/apps/talk`, `${RENDERER}/apps/talk.js`, `${MAIN}/talk.js`, `${MAIN}/talk-signin.js`] },
  { what: 'Spelling dictionaries', npm: ['dictionary-en', 'dictionary-en-gb'] },
  { what: 'Database engine clients', npm: ['pg', 'mysql2'] },
  { what: 'Mail protocols', npm: ['imapflow', 'nodemailer'] },
];

const LOADERS = {
  '.js': 'jsx', '.png': 'file', '.jpg': 'file', '.gif': 'file', '.svg': 'file', '.woff': 'file', '.woff2': 'file',
  '.ttf': 'file', '.otf': 'file', '.wasm': 'file', '.glb': 'file', '.mp3': 'file', '.wav': 'file', '.txt': 'text',
};
const BUILTIN = new Set([...builtinModules, 'electron', 'original-fs']);

const posix = (p) => p.split(path.sep).join('/');
const rel = (abs) => posix(path.relative(ROOT, abs));
const exists = (p) => fs.existsSync(path.join(ROOT, p));

/** An npm package's name from a specifier or a node_modules path. */
function npmName(spec) {
  const parts = spec.replace(/^.*node_modules\//, '').split('/');
  return parts[0].startsWith('@') ? `${parts[0]}/${parts[1]}` : parts[0];
}

/** A file's home: a pack's id, 'core', 'checks', 'pkg:<name>' or 'npm:<name>'. */
export function homeOf(file) {
  if (file.startsWith('npm:')) return file;
  if (file.startsWith(`${MAIN}/`)) {
    const r = file.slice(MAIN.length + 1);
    const own = /^([a-z]+)-pack\.js$/.exec(r);
    if (own && isPack(own[1])) return own[1];
    for (const [re, home] of MAIN_HOMES) if (re.test(r)) return home;
    return 'core';
  }
  if (file.startsWith(`${RENDERER}/`)) {
    const r = file.slice(RENDERER.length + 1);
    if (RENDERER_HOMES[r]) return RENDERER_HOMES[r];
    const m = /^apps\/([\w-]+)(?:\.js$|\/)/.exec(r);
    return m && isPack(m[1]) ? m[1] : 'core';
  }
  if (file.startsWith('packages/')) return `pkg:${file.split('/')[1]}`;
  if (file.startsWith('node_modules/')) return `npm:${npmName(file)}`;
  // The tests' fixtures, reached only from the checks.
  return 'checks';
}

/* ── reading the wiring ─────────────────────────────────────────────── */

/** renderer/index.js's APP_LOADERS: each app's window module. */
export function rendererEntries() {
  const text = fs.readFileSync(path.join(ROOT, RENDERER, 'index.js'), 'utf8');
  const out = {};
  for (const m of text.matchAll(/^\s*(\w+):\s*\(\)\s*=>\s*import\(\s*'([^']+)'\s*\)/gm)) {
    out[m[1]] = posix(path.join(RENDERER, m[2]));
  }
  return out;
}

/** A main-process file's factories, by the names it imports them as, and the namespaces it builds with them. */
function factoriesIn(file) {
  const text = fs.readFileSync(path.join(ROOT, MAIN, file), 'utf8');
  const factories = {};
  // `import { createX } from './x.js'`, or `const { createX } = await import('./x.js')` where the module is loaded when it registers.
  for (const m of text.matchAll(/(?:import\s*\{([^}]+)\}\s*from\s*|const\s*\{([^}]+)\}\s*=\s*await\s+import\(\s*)'(\.\/[^']+)'/g)) {
    for (const name of (m[1] || m[2]).split(',').map((s) => s.trim().split(/\s+as\s+|\s*:\s*/).pop())) if (name) factories[name] = posix(path.join(MAIN, m[3]));
  }
  const byName = {};
  // `const contacts = createContactsService(`, `mail: createMailService(`, `update: updates = createUpdateService(` and `services.dav = createDavService(`.
  for (const m of text.matchAll(/(?:const\s+(\w+)\s*=|services\.(\w+)\s*=|(\w+):\s*(?:\w+\s*=\s*)?)\s*(create\w+)\s*\(/g)) {
    const ns = m[1] || m[2] || m[3];
    if (factories[m[4]]) byName[ns] = factories[m[4]];
  }
  return { text, byName };
}

/** Each app's register module (app-packs.js BUILT_IN), by the pack's id. */
export function registerModules() {
  const text = fs.readFileSync(path.join(ROOT, MAIN, 'app-packs.js'), 'utf8');
  const out = {};
  for (const m of text.matchAll(/^\s*(\w+):\s*\(\)\s*=>\s*import\(\s*'\.\/([\w-]+\.js)'\s*\)/gm)) out[m[1]] = posix(path.join(MAIN, m[2]));
  return out;
}

/**
 * The main process's namespaces: each IPC namespace and the module whose
 * factory builds it. The core's are built in main.js; each app's in its
 * register module, which the core calls through the loader.
 */
export function mainWiring() {
  const { text, byName } = factoriesIn('main.js');
  // The document service is built by the document host, and is the doc namespace.
  if (byName.docHost) byName.doc = byName.docHost;
  const wired = {};
  const services = text.slice(text.indexOf('services = {'));
  for (const m of services.matchAll(/^\s{6}(\w+)(?=[:,])/gm)) if (byName[m[1]]) wired[m[1]] = byName[m[1]];
  for (const m of text.matchAll(/services\.(\w+)\s*=\s*create\w+\s*\(/g)) if (byName[m[1]]) wired[m[1]] = byName[m[1]];
  if (byName.doc) wired.doc = byName.doc;
  for (const file of Object.values(registerModules())) {
    const pack = factoriesIn(path.basename(file)).byName;
    for (const [ns, module] of Object.entries(pack)) wired[ns] = module;
  }
  return wired;
}

/** The IPC contract: namespace → methods, and the events. */
async function contract() {
  const mod = await import(pathToFileURL(path.join(ROOT, 'packages/office-shell/src/contract.js')).href);
  return { methods: mod.METHODS, events: mod.EVENTS, prefix: mod.CHANNEL_PREFIX };
}

/* ── the graph ──────────────────────────────────────────────────────── */

/** Everything outside the suite's own code is left out of the walk and named. */
const outside = {
  name: 'outside',
  setup(build) {
    build.onResolve({ filter: /^[^./]/ }, (a) => {
      if (path.isAbsolute(a.path) || a.path.startsWith('@rutba/')) return undefined;
      const name = a.path.startsWith('node:') ? a.path : npmName(a.path);
      return { path: BUILTIN.has(name.replace(/^node:/, '')) || name.startsWith('node:') ? `node:${name.replace(/^node:/, '')}` : name, external: true };
    });
  },
};

const lineCache = new Map();
/** The line an import of `spec` is written on (1-based; 0 when not found). */
function lineOf(file, spec) {
  if (!lineCache.has(file)) lineCache.set(file, fs.readFileSync(path.join(ROOT, file), 'utf8').split(/\r?\n/));
  const lines = lineCache.get(file);
  const quoted = [`'${spec}'`, `"${spec}"`, `\`${spec}\``];
  const has = (l) => quoted.some((q) => l.includes(q));
  let i = lines.findIndex((l) => has(l) && /\b(import|from|require|export|URL)\b/.test(l));
  if (i < 0) i = lines.findIndex(has);
  return i + 1;
}

/** One side's graph: file → [{ to, spec, line, kind }], from esbuild's metafile and the workers it does not follow. */
async function sideGraph(entries, platform) {
  const graph = new Map();
  let wanted = [...new Set(entries)].filter(exists);
  for (let round = 0; round < 6 && wanted.length; round++) {
    const result = await esbuild.build({
      absWorkingDir: ROOT,
      entryPoints: wanted,
      bundle: true,
      write: false,
      metafile: true,
      outdir: path.join(ROOT, '.app-map-out'),
      format: 'esm',
      splitting: true,
      platform,
      target: platform === 'node' ? ['node20'] : ['chrome120'],
      loader: LOADERS,
      jsx: 'automatic',
      logLevel: 'silent',
      plugins: [outside],
    });
    const added = [];
    for (const [file, input] of Object.entries(result.metafile.inputs)) {
      if (file.includes(':') || file.startsWith('(')) continue;
      const edges = [];
      for (const imp of input.imports) {
        if (imp.path.startsWith('node:') || BUILTIN.has(imp.path)) continue;
        const to = imp.external ? `npm:${imp.path}` : imp.path;
        if (to.includes(':') && !to.startsWith('npm:')) continue;
        edges.push({ to, spec: imp.original || imp.path, line: imp.original ? lineOf(file, imp.original) : 0, kind: imp.kind });
      }
      // A worker or a file read by its own address: new URL('./x.js', import.meta.url).
      if (/\.(m?js|cjs)$/.test(file)) {
        const text = fs.readFileSync(path.join(ROOT, file), 'utf8');
        for (const m of text.matchAll(/new URL\(\s*(['"])(\.{1,2}\/[^'"]+)\1\s*,\s*import\.meta\.url/g)) {
          const to = posix(path.join(path.dirname(file), m[2])).replace(/\/$/, '');
          if (!exists(to)) continue;
          // A folder or a data file read at run time: whoever reaches this file reaches it.
          if (!/\.(m?js|cjs)$/.test(to)) {
            edges.push({ to, spec: m[2], line: lineOf(file, m[2]), kind: 'asset' });
            continue;
          }
          edges.push({ to, spec: m[2], line: lineOf(file, m[2]), kind: 'worker' });
          if (!result.metafile.inputs[to]) added.push(to);
        }
      }
      graph.set(file, edges);
    }
    // Only what the last round did not reach: the files it did are in the graph already.
    wanted = [...new Set(added)].filter((f) => !graph.has(f));
  }
  return graph;
}

/**
 * Walks from one owner's entries. Another app's file is not entered; the
 * import that reaches it is kept as an edge. The checks are never entered.
 */
function walk(graph, starts, owner, reach, edges, side) {
  const seen = new Set();
  const queue = starts.filter((s) => graph.has(s));
  while (queue.length) {
    const file = queue.pop();
    if (seen.has(file)) continue;
    seen.add(file);
    if (!reach.has(file)) reach.set(file, new Map());
    const by = reach.get(file);
    by.set(owner, new Set([...(by.get(owner) || []), side]));
    if (file.startsWith('npm:')) continue;
    for (const e of graph.get(file) || []) {
      const to = homeOf(e.to);
      if (to === 'checks') continue;
      edges.set(`${file} -> ${e.to}`, { from: file, ...e });
      if (isPack(to) && to !== owner) continue;
      if (e.to.startsWith('npm:') || e.kind === 'asset') {
        if (!reach.has(e.to)) reach.set(e.to, new Map());
        const r = reach.get(e.to);
        r.set(owner, new Set([...(r.get(owner) || []), side]));
        continue;
      }
      queue.push(e.to);
    }
  }
}

/** Who a set of reachers makes the owner: the core when it or two apps reach it, else the one app. */
function ownerFrom(reachers) {
  if (!reachers || !reachers.size) return 'unreached';
  if (reachers.has('core')) return 'core';
  const packs = [...reachers.keys()].filter(isPack);
  return packs.length === 1 ? packs[0] : 'core';
}

/**
 * The map's import graph alone: homes, owners and the cross-app imports. What
 * the test needs; the sizes are in buildMap.
 */
export async function mapImports() {
  const loaders = rendererEntries();
  const wired = mainWiring();
  const starts = { renderer: { core: [...CORE_ENTRIES.renderer] }, main: { core: [...CORE_ENTRIES.main] } };
  const unknownApps = Object.keys(loaders).filter((k) => k !== 'home' && !isPack(k));
  for (const [key, file] of Object.entries(loaders)) {
    if (key === 'home') continue;
    (starts.renderer[key] ||= []).push(file);
  }
  const namespaces = {};
  for (const [ns, file] of Object.entries(wired)) {
    const owner = NAMESPACE_PACK[ns] || 'core';
    namespaces[ns] = { owner, module: file };
    if (owner !== 'core') (starts.main[owner] ||= []).push(file);
  }
  // Each app's register module is where its main-process side starts.
  const registers = registerModules();
  for (const [id, file] of Object.entries(registers)) if (isPack(id)) (starts.main[id] ||= []).push(file);
  for (const side of ['renderer', 'main']) {
    for (const [id, files] of Object.entries(EXTRA_ENTRIES[side])) (starts[side][id] ||= []).push(...files);
  }

  const [renderer, main] = await Promise.all([
    sideGraph(Object.values(starts.renderer).flat(), 'browser'),
    sideGraph(Object.values(starts.main).flat(), 'node'),
  ]);
  const graphs = { renderer, main };
  const reach = new Map();
  const edges = new Map();
  for (const side of ['renderer', 'main']) {
    for (const [owner, files] of Object.entries(starts[side])) walk(graphs[side], files, owner, reach, edges, side);
  }

  // A package moves as one: its owner is the owner of everything that reaches any file of it.
  const packageReach = new Map();
  for (const [file, by] of reach) {
    const h = homeOf(file);
    if (!h.startsWith('pkg:') && !h.startsWith('npm:')) continue;
    if (!packageReach.has(h)) packageReach.set(h, new Map());
    const pr = packageReach.get(h);
    for (const [owner, sides] of by) pr.set(owner, new Set([...(pr.get(owner) || []), ...sides]));
  }
  const packageOwner = new Map([...packageReach].map(([h, by]) => [h, ownerFrom(by)]));
  // The folders and data files read at run time, longest first, for the files inside them.
  const assets = [...reach.keys()].filter((f) => !f.startsWith('npm:') && !graphs.renderer.has(f) && !graphs.main.has(f)).sort((a, b) => b.length - a.length);

  /**
   * The owner a file has in the proposal: its home; for a core file one app
   * alone uses, that app; for an engine, the engine's owner, except in an
   * engine split by file (SPLIT_PACKAGES), where each file goes by who reaches
   * it and a data file by who reads its folder.
   */
  const ownerOf = (file) => {
    const h = homeOf(file);
    if (isPack(h) || h === 'checks') return h;
    if (h.startsWith('npm:')) return packageOwner.get(h) || 'unreached';
    if (h.startsWith('pkg:')) {
      if (!SPLIT_PACKAGES.has(h.slice(4))) return packageOwner.get(h) || 'unreached';
      if (reach.has(file)) return ownerFrom(reach.get(file));
      const folder = assets.find((a) => file === a || file.startsWith(`${a}/`));
      return folder ? ownerFrom(reach.get(folder)) : 'core';
    }
    return ownerFrom(reach.get(file));
  };
  const entryModules = new Set([...Object.values(loaders), ...Object.values(PACK_WINDOWS), ...Object.values(wired), ...Object.values(registers)]);

  const cross = [];
  const registration = [];
  for (const e of edges.values()) {
    const fromOwner = ownerOf(e.from);
    const toOwner = ownerOf(e.to.startsWith('npm:') ? e.to : e.to);
    if (fromOwner === toOwner || toOwner === 'checks' || fromOwner === 'checks') continue;
    // An edge into a core file or a core engine is what the core is for.
    if (!isPack(toOwner)) continue;
    const row = { from: e.from, line: e.line, spec: e.spec, to: e.to, fromOwner, toOwner, kind: e.kind };
    // The window module and the service module of each app are registration; anything else the wiring imports is a cut.
    if (WIRING.has(e.from) && entryModules.has(e.to)) registration.push(row);
    else cross.push(row);
  }
  const order = (a, b) => (a.fromOwner + a.from + a.line).localeCompare(b.fromOwner + b.from + b.line);
  cross.sort(order);
  registration.sort(order);

  // Who reaches each package, and through which import first.
  const packages = {};
  for (const [h, by] of [...packageReach].sort()) {
    const importers = {};
    for (const e of edges.values()) {
      if (homeOf(e.to) !== h || homeOf(e.from) === h) continue;
      const o = ownerOf(e.from);
      (importers[o] ||= []).push(`${e.from}:${e.line}`);
    }
    for (const o of Object.keys(importers)) importers[o] = [...new Set(importers[o])].sort().slice(0, 1);
    const sides = new Set([...by.values()].flatMap((s) => [...s]));
    // An engine split by file: the owners of its parts.
    const parts = SPLIT_PACKAGES.has(h.slice(4)) ? [...new Set([...reach.keys()].filter((f) => homeOf(f) === h).map(ownerOf))].sort() : null;
    packages[h] = { owner: packageOwner.get(h), parts, reachedBy: [...by.keys()].sort(), sides: [...sides].sort(), importers };
  }

  // Core files one app alone uses: they move with it.
  const moves = [];
  for (const [file, by] of reach) {
    const h = homeOf(file);
    if (h !== 'core' || by.has('core')) continue;
    const o = ownerFrom(by);
    if (isPack(o)) moves.push({ file, to: o });
  }
  moves.sort((a, b) => (a.to + a.file).localeCompare(b.to + b.file));

  return { loaders, namespaces, starts, graphs, reach, edges, cross, registration, packages, packageOwner, ownerOf, moves, unknownApps };
}

/* ── sizes ──────────────────────────────────────────────────────────── */

const sizeOf = (p) => {
  try {
    return fs.statSync(path.join(ROOT, p)).size;
  } catch {
    return 0;
  }
};

/** A folder's files as electron-builder ships them, with their sizes: no tests, no notes, no source maps. */
function shippedFiles(dir, out = []) {
  const abs = path.join(ROOT, dir);
  if (!fs.existsSync(abs)) return out;
  const st = fs.statSync(abs);
  if (st.isFile()) {
    out.push([dir, st.size]);
    return out;
  }
  for (const ent of fs.readdirSync(abs, { withFileTypes: true })) {
    const child = posix(path.join(dir, ent.name));
    if (ent.isDirectory()) {
      if (!['test', 'tests', '.git', 'node_modules'].includes(ent.name)) shippedFiles(child, out);
    } else if (!/\.(md|map)$/i.test(ent.name)) {
      out.push([child, fs.statSync(path.join(abs, ent.name)).size]);
    }
  }
  return out;
}
const shippedBytes = (dir) => shippedFiles(dir).reduce((n, [, s]) => n + s, 0);

/** An npm package and everything it needs, as folders under node_modules. */
function npmClosure(name, from = 'node_modules', seen = new Set()) {
  const nested = posix(path.join(from, name));
  const dir = from !== 'node_modules' && exists(nested) ? nested : `node_modules/${name}`;
  if (seen.has(dir) || !exists(`${dir}/package.json`)) return seen;
  seen.add(dir);
  const pkg = JSON.parse(fs.readFileSync(path.join(ROOT, dir, 'package.json'), 'utf8'));
  for (const dep of Object.keys(pkg.dependencies || {})) npmClosure(dep, `${dir}/node_modules`, seen);
  return seen;
}

/** The renderer as it ships: minified, split by app, bytes per source file. */
async function rendererOutput() {
  const result = await esbuild.build({
    absWorkingDir: ROOT,
    entryPoints: [`${RENDERER}/index.js`, `${RENDERER}/apps/sheets/scripts-worker.js`, `${RENDERER}/imaging/worker.js`],
    bundle: true,
    write: false,
    metafile: true,
    outdir: path.join(ROOT, '.app-map-out'),
    format: 'esm',
    splitting: true,
    platform: 'browser',
    target: ['chrome120'],
    loader: { '.js': 'jsx', '.png': 'dataurl', '.svg': 'dataurl', '.woff2': 'dataurl' },
    jsx: 'automatic',
    minify: true,
    legalComments: 'none',
    define: { 'process.env.NODE_ENV': '"production"' },
    logLevel: 'silent',
  });
  const bytes = new Map();
  for (const out of Object.values(result.metafile.outputs)) {
    for (const [file, { bytesInOutput }] of Object.entries(out.inputs)) bytes.set(file, (bytes.get(file) || 0) + bytesInOutput);
  }
  return bytes;
}

/** The whole map: the graph, and per pack its sizes, IPC, file types, tiles, AI tools and checks. */
export async function buildMap() {
  const g = await mapImports();
  const { methods, events, prefix } = await contract();
  const { fileAssociations, APP_EXTENSIONS, OLD_EXTENSIONS } = await import(pathToFileURL(path.join(ROOT, 'packages/office-formats/src/registry.js')).href);
  const owners = ['core', ...PACKS.map((p) => p.id)];
  const blank = () => ({ renderer: 0, main: 0, packages: 0, npm: 0, assets: 0 });
  const size = Object.fromEntries(owners.map((o) => [o, blank()]));
  const bucket = (o) => size[isPack(o) ? o : 'core'];

  // The renderer, minified and split as the build makes it.
  for (const [file, n] of await rendererOutput()) {
    const unit = file.startsWith('node_modules/') ? `npm:${npmName(file)}` : file;
    bucket(g.ownerOf(unit)).renderer += n;
  }
  // The main process's own files, as they ship.
  for (const [file, by] of g.reach) {
    if (!file.startsWith(`${DESKTOP}/`)) continue;
    const sides = new Set([...by.values()].flatMap((s) => [...s]));
    if (sides.has('main')) bucket(g.ownerOf(file)).main += sizeOf(file);
  }
  // A package the main process reaches ships as its folder; one only a window reaches is in the bundle already.
  const packageRows = [];
  for (const [h, p] of Object.entries(g.packages)) {
    if (!h.startsWith('pkg:')) continue;
    const name = h.slice(4);
    const files = shippedFiles(`packages/${name}`);
    const bytes = files.reduce((n, [, s]) => n + s, 0);
    // An engine split by file: each part to its owner.
    const split = {};
    for (const [f, s] of files) {
      const o = SPLIT_PACKAGES.has(name) ? g.ownerOf(f) : p.owner;
      split[isPack(o) ? o : 'core'] = (split[isPack(o) ? o : 'core'] || 0) + s;
    }
    if (p.sides.includes('main')) for (const [o, s] of Object.entries(split)) bucket(o).packages += s;
    packageRows.push({ name: `@rutba/${name}`, owner: p.owner, sides: p.sides, bytes, split: SPLIT_PACKAGES.has(name) ? split : null, reachedBy: p.reachedBy, importers: p.importers });
  }
  const npmRows = [];
  const npmFolders = new Map();
  for (const [h, p] of Object.entries(g.packages)) {
    if (!h.startsWith('npm:') || !p.sides.includes('main')) continue;
    for (const dir of npmClosure(h.slice(4))) {
      if (!npmFolders.has(dir)) npmFolders.set(dir, new Set());
      npmFolders.get(dir).add(p.owner);
    }
    npmRows.push({ name: h.slice(4), owner: p.owner, reachedBy: p.reachedBy });
  }
  for (const [dir, by] of npmFolders) {
    const o = by.size === 1 ? [...by][0] : 'core';
    bucket(o).npm += shippedBytes(dir);
  }
  for (const row of npmRows) row.bytes = [...npmClosure(row.name)].reduce((n, d) => n + shippedBytes(d), 0);

  // What the archive carries that the main process never reaches: electron-builder copies every
  // production dependency's folder, and an engine the windows alone use is in their bundle already.
  const desktop = JSON.parse(fs.readFileSync(path.join(ROOT, DESKTOP, 'package.json'), 'utf8'));
  const unneeded = [];
  for (const dep of Object.keys(desktop.dependencies || {})) {
    const ours = dep.startsWith('@rutba/');
    const p = g.packages[ours ? `pkg:${dep.slice(7)}` : `npm:${dep}`];
    if (p?.sides.includes('main')) continue;
    const bytes = ours ? shippedBytes(`packages/${dep.slice(7)}`) : [...npmClosure(dep)].reduce((n, d) => n + shippedBytes(d), 0);
    unneeded.push({ name: dep, bytes, why: p ? 'the windows only: bundled' : 'reached by nothing' });
  }

  // The heavy parts: their size, and who owns them.
  const heavy = HEAVY.map((h) => {
    if (h.npm) {
      const rows = npmRows.filter((r) => h.npm.includes(r.name));
      const o = [...new Set(rows.map((r) => r.owner))];
      return { what: h.what, where: h.npm.join(', '), bytes: rows.reduce((n, r) => n + r.bytes, 0), owner: o.length === 1 ? o[0] : o.length ? 'core' : 'unreached', counted: 'npm' };
    }
    const files = h.dirs.flatMap((d) => shippedFiles(d));
    const bytes = files.reduce((n, [, s]) => n + s, 0);
    // Its owner: who loads it, or who owns most of its bytes.
    let owner = h.through ? g.ownerOf(h.through) : null;
    if (!owner) {
      const by = {};
      for (const [f, s] of files) by[g.ownerOf(f)] = (by[g.ownerOf(f)] || 0) + s;
      owner = Object.entries(by).sort((a, b) => b[1] - a[1])[0]?.[0] || 'unreached';
      if (!isPack(owner)) owner = 'core';
    }
    // pdf.js is copied beside the bundle, not imported: counted here, once.
    const inPackage = h.dirs.every((d) => d.startsWith('packages/'));
    const inMain = h.dirs.every((d) => d.startsWith(`${MAIN}`));
    const inRenderer = h.dirs.every((d) => d.startsWith(`${RENDERER}`) || d.startsWith(`${MAIN}`));
    if (!inPackage && !inMain && !inRenderer) bucket(owner).assets += bytes;
    return { what: h.what, where: h.dirs.join(', '), bytes, owner, counted: inPackage ? 'package' : inRenderer ? 'code' : 'asset' };
  });

  // IPC: each pack's namespaces and events, and the channels they make.
  const ipc = Object.fromEntries(owners.map((o) => [o, { namespaces: [], channels: 0, events: [] }]));
  for (const [ns, names] of Object.entries(methods)) {
    const o = NAMESPACE_PACK[ns] || 'core';
    ipc[o].namespaces.push(ns);
    ipc[o].channels += names.length;
  }
  for (const ev of events) ipc[EVENT_PACK[ev] || 'core'].events.push(ev);

  // What each app's windows call of another app's namespace, and which other apps' windows they open.
  const calls = Object.fromEntries(PACKS.map((p) => [p.id, {}]));
  const opens = Object.fromEntries(PACKS.map((p) => [p.id, {}]));
  for (const [file, by] of g.reach) {
    if (!file.startsWith(`${RENDERER}/`)) continue;
    const o = g.ownerOf(file);
    if (!isPack(o)) continue;
    const lines = fs.readFileSync(path.join(ROOT, file), 'utf8').split(/\r?\n/);
    lines.forEach((l, i) => {
      for (const m of l.matchAll(/\bshell\??\.(\w+)\??\.(\w+)\s*\(/g)) {
        const owner = NAMESPACE_PACK[m[1]];
        if (owner && owner !== o) (calls[o][owner] ||= []).push(`${file}:${i + 1} ${m[1]}.${m[2]}`);
      }
      for (const m of l.matchAll(/win\??\.create\(\{\s*app:\s*'(\w+)'/g)) {
        if (m[1] !== o && isPack(m[1])) (opens[o][m[1]] ||= []).push(`${file}:${i + 1}`);
      }
    });
    void by;
  }

  // File types: what the installer registers for each app, and what else it opens.
  const claimed = {};
  for (const a of fileAssociations()) (claimed[a.app] ||= []).push(a.ext);
  const fileTypes = Object.fromEntries(PACKS.map((p) => {
    const own = claimed[p.id] || [];
    const also = [...(APP_EXTENSIONS[p.id] || []), ...(OLD_EXTENSIONS[p.id] || [])].filter((e) => !own.includes(e));
    return [p.id, { claims: own, alsoOpens: also }];
  }));

  // Tiles: the taskbar's and the file types' icons.
  const tiles = Object.fromEntries(PACKS.map((p) => [p.id, ['apps', 'filetypes'].flatMap((d) => ['ico', 'png'].map((x) => `${DESKTOP}/resources/${d}/${p.id}.${x}`)).filter(exists)]));

  // The AI tools, by what each works with.
  const toolsText = fs.readFileSync(path.join(ROOT, MAIN, 'mcp', 'tools.js'), 'utf8');
  const list = toolsText.slice(toolsText.indexOf('export const TOOL_LIST'));
  const tools = Object.fromEntries(owners.map((o) => [o, []]));
  for (const m of list.matchAll(/^ {4}name: '(\w+)'/gm)) tools[TOOL_PACK[m[1]] || 'ai'].push(m[1]);

  // The window checks, by name.
  const checks = Object.fromEntries(owners.map((o) => [o, { count: 0, bytes: 0 }]));
  for (const f of fs.readdirSync(path.join(ROOT, MAIN))) {
    if (homeOf(`${MAIN}/${f}`) !== 'checks') continue;
    const o = CHECK_PACK.find(([re]) => re.test(f))?.[1] || 'core';
    checks[o].count += 1;
    checks[o].bytes += sizeOf(`${MAIN}/${f}`);
  }

  // What each pack needs: the imports not yet cut, and (softer) what it calls and opens.
  const needs = Object.fromEntries(PACKS.map((p) => [p.id, { imports: {}, calls: Object.keys(calls[p.id]).sort(), opens: Object.keys(opens[p.id]).sort() }]));
  for (const e of g.cross) if (isPack(e.fromOwner)) needs[e.fromOwner].imports[e.toOwner] = (needs[e.fromOwner].imports[e.toOwner] || 0) + 1;

  const packs = PACKS.map((p) => {
    const s = size[p.id];
    return {
      ...p,
      bytes: s.renderer + s.main + s.packages + s.npm + s.assets,
      size: s,
      needs: needs[p.id],
      ipc: ipc[p.id],
      fileTypes: fileTypes[p.id],
      tiles: tiles[p.id],
      tools: tools[p.id],
      checks: checks[p.id],
      calls: calls[p.id],
      opens: opens[p.id],
      entries: { renderer: g.starts.renderer[p.id] || [], main: g.starts.main[p.id] || [] },
    };
  });
  const c = size.core;
  const core = { bytes: c.renderer + c.main + c.packages + c.npm + c.assets, size: c, ipc: ipc.core, tools: tools.core, checks: checks.core };

  return {
    prefix,
    packs,
    core,
    packages: packageRows.sort((a, b) => a.name.localeCompare(b.name)),
    npm: npmRows.sort((a, b) => a.name.localeCompare(b.name)),
    heavy,
    unneeded,
    cross: g.cross,
    registration: g.registration,
    moves: g.moves,
    namespaces: g.namespaces,
    unknownApps: g.unknownApps,
  };
}

/* ── the allowed edges, for the test ────────────────────────────────── */

/** What the fixture keeps: the cross-app imports by file, and each engine's owner. */
export function edgesOf(g) {
  return {
    cross: [...new Set(g.cross.map((e) => `${e.from} -> ${e.to}`))].sort(),
    owners: Object.fromEntries(Object.entries(g.packages).map(([h, p]) => [h, p.parts ? p.parts.join('+') : p.owner]).sort()),
  };
}

/* ── the document's tables ──────────────────────────────────────────── */

const kb = (n) => (n >= 1024 * 1024 ? `${(n / 1024 / 1024).toFixed(1)} MB` : `${Math.round(n / 1024)} KB`);
const ownerName = (o) => PACKS.find((p) => p.id === o)?.name || (o === 'core' ? 'core' : o);
const at = (file, line) => `\`${file}:${line}\``;

export function tables(map) {
  const out = [];
  out.push('### The packs\n');
  out.push('Sizes as shipped before compression: the window code minified as the build makes it, the main-process files, the engines and npm packages only that pack reaches (as folders, without tests and notes) and its other files. "Imports" are the imports to cut (below); "calls" and "opens" are the other packs an app reaches at run time, through IPC or by opening a window.\n');
  out.push('| Pack | Size | Windows | Main | Engines | npm | Other | Imports to cut, to | Calls | Opens | File types it claims | IPC namespaces (channels) | AI tools | Checks |');
  out.push('|---|---:|---:|---:|---:|---:|---:|---|---|---|---|---|---|---:|');
  for (const p of map.packs) {
    const imports = Object.entries(p.needs.imports).map(([o, n]) => `${ownerName(o)} (${n})`).join(', ') || 'none';
    const s = p.size;
    out.push(`| ${p.name} (\`${p.id}\`) | ${kb(p.bytes)} | ${kb(s.renderer)} | ${kb(s.main)} | ${kb(s.packages)} | ${kb(s.npm)} | ${kb(s.assets)} | ${imports} | ${p.needs.calls.map(ownerName).join(', ') || 'none'} | ${p.needs.opens.map(ownerName).join(', ') || 'none'} | ${p.fileTypes.claims.join(' ') || 'none'} | ${p.ipc.namespaces.join(', ') || 'none'} (${p.ipc.channels}) | ${p.tools.join(', ') || 'none'} | ${p.checks.count} |`);
  }
  const c = map.core;
  const total = map.packs.reduce((n, p) => n + p.bytes, c.bytes);
  out.push(`| **Core** | ${kb(c.bytes)} | ${kb(c.size.renderer)} | ${kb(c.size.main)} | ${kb(c.size.packages)} | ${kb(c.size.npm)} | ${kb(c.size.assets)} | | | | | ${c.ipc.namespaces.join(', ')} (${c.ipc.channels}) | ${c.tools.join(', ')} | ${c.checks.count} |`);
  out.push(`\nAll together: ${kb(total)}. The window checks (\`main/verify-*.js\` and the smoke and screens runs, ${kb(map.packs.reduce((n, p) => n + p.checks.bytes, c.checks.bytes))}) are not in these sizes.\n`);

  out.push('### The engines and who reaches them\n');
  out.push('An engine one pack alone reaches moves into that pack; one the core or two packs reach stays in the core. "Windows" means it reaches the windows only, where it is bundled; its folder in the archive is then not needed at run time.\n');
  out.push('| Engine | Owner | Size | Reached from | Reached by | First import, by owner |');
  out.push('|---|---|---:|---|---|---|');
  for (const r of map.packages) {
    const first = Object.entries(r.importers).map(([o, l]) => `${ownerName(o)}: ${l.map((x) => `\`${x}\``).join(', ')}`).join('; ');
    const owner = r.split ? Object.entries(r.split).sort((a, b) => b[1] - a[1]).map(([o, n]) => `${ownerName(o)} ${kb(n)}`).join(', ') : ownerName(r.owner);
    out.push(`| ${r.name} | ${r.split ? `split: ${owner}` : owner} | ${kb(r.bytes)} | ${r.sides.map((s) => (s === 'main' ? 'main process' : 'windows')).join(', ')} | ${r.reachedBy.map(ownerName).join(', ')} | ${first} |`);
  }
  out.push('\n| npm package (main process) | Owner | Size with what it needs | Reached by |');
  out.push('|---|---|---:|---|');
  for (const r of map.npm) out.push(`| ${r.name} | ${ownerName(r.owner)} | ${kb(r.bytes)} | ${r.reachedBy.map(ownerName).join(', ')} |`);

  out.push('\n### The heavy parts\n');
  out.push('| What | Where | Size | Owner |');
  out.push('|---|---|---:|---|');
  for (const h of map.heavy) out.push(`| ${h.what} | \`${h.where}\` | ${kb(h.bytes)} | ${ownerName(h.owner)} |`);

  out.push('\n### In the archive, never reached by the main process\n');
  out.push("electron-builder copies the folder of every dependency `apps/desktop/package.json` names. These the main process never imports: either the windows' bundle already holds what they use, or nothing uses them. Neither the core nor a pack needs their folders.\n");
  out.push('| Dependency | Size | Why |');
  out.push('|---|---:|---|');
  for (const u of map.unneeded) out.push(`| ${u.name} | ${kb(u.bytes)} | ${u.why} |`);

  const groups = {};
  for (const e of map.cross) (groups[`${e.fromOwner} -> ${e.toOwner}`] ||= []).push(e);
  out.push(`\n### The imports to cut (${map.cross.length})\n`);
  out.push('Each is a file of one owner importing a file of another app. Before the move each is cut (the app asks through IPC or an event, or the shared part moves to the core). Grouped by who imports whom.\n');
  out.push('| From | To | Where | Imports | Kind |');
  out.push('|---|---|---|---|---|');
  for (const [k, rows] of Object.entries(groups).sort()) {
    const [from, to] = k.split(' -> ');
    for (const e of rows) out.push(`| ${ownerName(from)} | ${ownerName(to)} | ${at(e.from, e.line)} | \`${e.to}\` | ${e.kind === 'dynamic-import' ? 'lazy' : e.kind === 'worker' ? 'worker' : 'static'} |`);
  }

  out.push(`\n### The core's wiring (${map.registration.length})\n`);
  out.push('What the core imports from each app today: `renderer/index.js` its window module, and `main/app-packs.js` its register module (`<id>-pack.js`), which builds its services. These are not cut: they are the registration, which a pack carrying its own code takes over from the core.\n');
  out.push('| Pack | Where | Imports |');
  out.push('|---|---|---|');
  for (const e of map.registration) out.push(`| ${ownerName(e.toOwner)} | ${at(e.from, e.line)} | \`${e.to}\` |`);

  out.push(`\n### Shared files one app alone uses (${map.moves.length})\n`);
  out.push('Files in the shared folders that only one app reaches: they move into its pack.\n');
  out.push('| File | Moves into |');
  out.push('|---|---|');
  for (const m of map.moves) out.push(`| \`${m.file}\` | ${ownerName(m.to)} |`);

  out.push('\n### Calls and windows across packs\n');
  out.push('Not imports, so nothing to cut, but a pack that is absent must not be offered: the button that would call it or open its window is shown only when it is installed.\n');
  out.push('| Pack | Calls (namespace: first place) | Opens a window of |');
  out.push('|---|---|---|');
  for (const p of map.packs) {
    const calls = Object.entries(p.calls).map(([o, l]) => `${ownerName(o)} (${l.length}): \`${l[0]}\``).join('; ');
    const opens = Object.entries(p.opens).map(([o, l]) => `${ownerName(o)}: \`${l[0]}\``).join('; ');
    if (calls || opens) out.push(`| ${p.name} | ${calls || 'none'} | ${opens || 'none'} |`);
  }
  return out.join('\n');
}

/* ── the command line ───────────────────────────────────────────────── */

async function cli() {
  const args = process.argv.slice(2);
  if (args.includes('--write-edges')) {
    const g = await mapImports();
    fs.writeFileSync(path.join(ROOT, EDGES_FIXTURE), `${JSON.stringify(edgesOf(g), null, 2)}\n`);
    console.log(`${EDGES_FIXTURE}: ${g.cross.length} cross-app imports, ${Object.keys(g.packages).length} engines.`);
    return;
  }
  const map = await buildMap();
  const json = args.indexOf('--json');
  if (json >= 0) fs.writeFileSync(path.resolve(args[json + 1]), `${JSON.stringify(map, null, 2)}\n`);
  if (args.includes('--doc')) {
    const file = path.join(ROOT, DOC);
    const text = fs.readFileSync(file, 'utf8');
    const begin = '<!-- app-map:begin -->';
    const end = '<!-- app-map:end -->';
    const a = text.indexOf(begin);
    const b = text.indexOf(end);
    if (a < 0 || b < a) throw new Error(`${DOC} has no ${begin} … ${end} block to write the map into.`);
    fs.writeFileSync(file, `${text.slice(0, a + begin.length)}\n${tables(map)}\n${text.slice(b)}`);
    console.log(`${DOC}: the map's tables written.`);
  }
  for (const p of map.packs) {
    const imports = Object.entries(p.needs.imports).map(([o, n]) => `${o}×${n}`).join(' ');
    console.log(`${p.name.padEnd(17)} ${kb(p.bytes).padStart(8)}  imports to cut: ${imports || 'none'}`);
  }
  console.log(`${'Core'.padEnd(17)} ${kb(map.core.bytes).padStart(8)}`);
  console.log(`${map.cross.length} cross-app imports to cut, ${map.registration.length} wiring imports, ${map.moves.length} shared files that move.`);
  if (map.unknownApps.length) console.log(`Apps with no pack: ${map.unknownApps.join(', ')}`);
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  cli().catch((err) => {
    console.error(err?.errors?.length ? err.errors.map((e) => `${e.location?.file}:${e.location?.line} ${e.text}`).join('\n') : err);
    process.exit(1);
  });
}
