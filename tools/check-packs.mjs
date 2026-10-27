// The release check for app packs (docs/APP-PACKS.md 5.6): a build's files
// sorted by owner, and each pack checked alone against the core it ships
// with.
//
//   node tools/check-packs.mjs [--core <app.asar or folder>] [--packs <folder>]
//                              [--keys <pack-keys.json>] [--installer <file>]
//                              [--version <v>] [--meta <renderer.json>]
//                              [--json <file>]
//
// The core is the packaged archive (release/win-unpacked/resources/app.asar
// after `npm run dist`), or a folder laid out the same way; the packs are the
// ones the installer carries (build/packs/out), with the include it writes
// them by (build/packs/installer-packs.nsh). Every file in the core is given
// an owner: the app whose code it is, by the same map the plan is drawn from
// (tools/app-map.mjs), or the core. A renderer chunk's owner is read off what
// went into it (build/bundle.js writes that, --meta): an app's when every
// input is that app's; a chunk that mixes an app's code with another's is the
// core's, and fails the check for an app that carries its own code.
//
// Then, for each pack, a copy with that pack alone is started the way a start
// plans it (main/app-packs.js), with the core's keys and version, and must
// hold nothing of another app: no IPC channel or event but the core's and its
// own, no tile, no kind of file in the Open dialog or the system's list, and
// no AI tool. The installer's part for that pack must write its files only
// and register its kinds of file only; where it carries them as a 7-Zip file
// of its own (rutba-pack-<id>.7z, unpacked as it installs), that file must be
// the pack's own archive and hold the pack's files only.
//
// A pack whose main module is the transitional one (it registers the core's
// own code, `core.builtIn`) still has its code in the core, and the check
// says how much. A pack that carries its code must have none left in the
// core: a file of it found there fails the check, as does anything above.
// Exits 1 on a failure.

import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import { createRequire } from 'node:module';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { readArchive } from '../packages/office-shell/src/electron/packs.js';
import { packFiles } from '../packages/office-shell/src/electron/pack-install.js';
import { PACKS, NAMESPACE_PACK, EVENT_PACK, coreMethods } from '../packages/office-shell/src/contract.js';
import { fileAssociations, openFilters, APP_EXTENSIONS, OLD_EXTENSIONS } from '../packages/office-formats/src/registry.js';
import { planPacks, readPackKeys } from '../apps/desktop/main/app-packs.js';
import { TOOL_PACK, TOOL_NEEDS, LOADED_KEY, toolOffered } from '../apps/desktop/main/pack-tools.js';
import { TOOL_LIST } from '../apps/desktop/main/mcp/tools.js';
import { REACH_TOOL_LIST } from '../apps/desktop/main/mcp/reach-tools.js';
import { ROOT, homeOf, mapImports } from './app-map.mjs';
import { rendererOwners, windowReach, asBuild, appWindows } from '../apps/desktop/build/pack-contents.js';

const DESKTOP = path.join(ROOT, 'apps', 'desktop');
const IDS = PACKS.map((p) => p.id);
const isPack = (o) => IDS.includes(o);
const nameOf = (id) => PACKS.find((p) => p.id === id)?.name || id;

/* ── the core's files ───────────────────────────────────────────────── */

/** The core's files and their sizes, from its archive or a folder laid out like one. */
export function coreFiles(core) {
  if (fs.statSync(core).isDirectory()) {
    const out = new Map();
    const walk = (dir, prefix) => {
      for (const ent of fs.readdirSync(dir, { withFileTypes: true })) {
        const p = prefix ? `${prefix}/${ent.name}` : ent.name;
        if (ent.isDirectory()) walk(path.join(dir, ent.name), p);
        else out.set(p, fs.statSync(path.join(dir, ent.name)).size);
      }
    };
    walk(core, '');
    return out;
  }
  return new Map([...readArchive(core)].map(([p, e]) => [p, e.size]));
}

/** One file of the core's, as text: from its archive (or the folder beside it) or a folder. */
function readCoreFile(core, name) {
  if (fs.statSync(core).isDirectory()) return fs.readFileSync(path.join(core, ...name.split('/')), 'utf8');
  const entry = readArchive(core).get(name);
  if (!entry) throw new Error(`${name} is not in ${core}.`);
  if (entry.unpacked) return fs.readFileSync(path.join(`${core}.unpacked`, ...name.split('/')), 'utf8');
  const fd = fs.openSync(core, 'r');
  try {
    const buf = Buffer.alloc(entry.size);
    fs.readSync(fd, buf, 0, entry.size, entry.offset);
    return buf.toString('utf8');
  } finally {
    fs.closeSync(fd);
  }
}

/** The workspace packages by their npm names, to their folders. */
function workspacePackages() {
  const out = {};
  for (const dir of fs.readdirSync(path.join(ROOT, 'packages'))) {
    try {
      out[JSON.parse(fs.readFileSync(path.join(ROOT, 'packages', dir, 'package.json'), 'utf8')).name] = dir;
    } catch {
      // not a package
    }
  }
  return out;
}

/** An npm package's own dependencies, from the repository's node_modules. */
function dependenciesOf(name) {
  try {
    return Object.keys(JSON.parse(fs.readFileSync(path.join(ROOT, 'node_modules', name, 'package.json'), 'utf8')).dependencies || {});
  } catch {
    return [];
  }
}

/**
 * Who owns a file the core ships, by its path in the archive: the map's
 * owner for the sources and the engines, an npm package's owner for what that
 * package brings with it, a renderer chunk by the windows that reach it (the
 * renderer's build, when it says what imports what), by what went into it
 * (its outputs alone), or else by the app it is named after. Anything no
 * single app owns is the core's.
 */
export function ownerResolver(g, outputs = null) {
  const build = outputs ? asBuild(outputs) : null;
  const reach = build ? windowReach(build, appWindows()) : null;
  const chunkOwners = reach || (build ? rendererOwners(g, build.outputs) : new Map());
  const workspace = workspacePackages();
  // An npm package reached by one owner brings its own dependencies with it.
  const npmOwner = new Map();
  for (const [h, owner] of g.packageOwner) if (h.startsWith('npm:')) npmOwner.set(h.slice(4), owner);
  const pending = [...npmOwner.keys()];
  while (pending.length) {
    const name = pending.pop();
    for (const dep of dependenciesOf(name)) {
      const was = npmOwner.get(dep);
      const now = was === undefined || was === npmOwner.get(name) ? npmOwner.get(name) : 'core';
      if (was !== now) {
        npmOwner.set(dep, now);
        pending.push(dep);
      }
    }
  }
  const sourceOwner = (file) => {
    const o = g.ownerOf(file);
    return isPack(o) ? o : 'core';
  };
  return (p) => {
    let m = /^main\/(.+)$/.exec(p);
    if (m) return homeOf(`apps/desktop/main/${m[1]}`) === 'checks' ? 'checks' : sourceOwner(`apps/desktop/main/${m[1]}`);
    const made = p.startsWith('build/out/') && chunkOwners.get(p.slice('build/out/'.length));
    if (made) return made.size === 1 && isPack([...made][0]) ? [...made][0] : 'core';
    m = /^build\/out\/chunks\/([a-z]+)-[A-Z0-9]+\.js$/.exec(p);
    if (m) return isPack(m[1]) ? m[1] : 'core';
    if (p === 'build/out/scripts-worker.js') return 'sheets';
    m = /^resources\/apps\/([a-z]+)\.ico$/.exec(p);
    if (m) return isPack(m[1]) ? m[1] : 'core';
    m = /^node_modules\/((?:@[^/]+\/)?[^/]+)\/(.*)$/.exec(p);
    if (m) {
      // A nested copy belongs with the package it sits in.
      const top = m[1];
      if (workspace[top]) return sourceOwner(`packages/${workspace[top]}/${m[2]}`);
      const owner = npmOwner.get(top);
      return isPack(owner) ? owner : 'core';
    }
    return 'core';
  };
}

/** The core's files, counted and measured by owner. */
export function filesByOwner(files, ownerOf) {
  const out = {};
  for (const [p, size] of files) {
    const o = ownerOf(p);
    const row = (out[o] ||= { files: 0, bytes: 0, sample: [] });
    row.files += 1;
    row.bytes += size;
    if (row.sample.length < 5) row.sample.push(p);
  }
  return out;
}

/* ── each pack alone ────────────────────────────────────────────────── */

/** Whether a pack's main module is the transitional one, which registers the core's own code. */
export function registersCoreCode(pack) {
  if (!pack.main || !pack.archive) return true;
  const inside = path.relative(pack.archive, pack.main).split(path.sep).join('/');
  const unpacked = path.join(`${pack.archive}.unpacked`, ...inside.split('/'));
  let text;
  if (fs.existsSync(unpacked)) text = fs.readFileSync(unpacked, 'utf8');
  else {
    const entry = readArchive(pack.archive).get(inside);
    if (!entry) return false;
    const fd = fs.openSync(pack.archive, 'r');
    try {
      const buf = Buffer.alloc(entry.size);
      fs.readSync(fd, buf, 0, entry.size, entry.offset);
      text = buf.toString('utf8');
    } finally {
      fs.closeSync(fd);
    }
  }
  return /\bcore\.builtIn\s*\(/.test(text);
}

/** The pack and the packs it needs, copied into a folder of their own. */
function aloneIn(packsDir, ids, into) {
  fs.rmSync(into, { recursive: true, force: true });
  fs.mkdirSync(into, { recursive: true });
  for (const id of ids) {
    const f = packFiles(id);
    for (const name of [f.archive, f.unpacked, f.manifest]) {
      const from = path.join(packsDir, name);
      if (fs.existsSync(from)) fs.cpSync(from, path.join(into, name), { recursive: true });
    }
  }
  return into;
}

/** What one pack started alone offers, and what of it is another app's. */
export function checkAlone({ id, packsDir, keys, version, scratch, toolNames = [...TOOL_LIST, ...REACH_TOOL_LIST].map((t) => t.name) }) {
  const failures = [];
  const fail = (what) => failures.push(`${nameOf(id)} alone: ${what}`);
  const manifest = JSON.parse(fs.readFileSync(path.join(packsDir, packFiles(id).manifest), 'utf8')).manifest;
  const needs = manifest.needs || [];
  const plan = planPacks({ dir: aloneIn(packsDir, [id, ...needs], path.join(scratch, id)), keys, version });
  const expected = [id, ...needs].sort();
  if (plan.refused.length) fail(`refused: ${plan.refused.map((r) => r.reason).join(' ')}`);
  if (plan.ids().sort().join(',') !== expected.join(',')) fail(`loaded ${plan.ids().join(', ') || 'nothing'}, not ${expected.join(', ')}`);
  const mine = new Set(expected);
  const offer = plan.offer();

  // IPC: the core's namespaces as the core has them, and this pack's own.
  const core = coreMethods();
  for (const [ns, names] of Object.entries(offer.methods)) {
    const owner = NAMESPACE_PACK[ns] || 'core';
    if (owner === 'core') {
      if (!core[ns]) fail(`offers the ${ns} namespace, which is no one's`);
      else if ([...names].sort().join() !== [...core[ns]].sort().join()) fail(`the core's ${ns} namespace is not as the core has it`);
    } else if (!mine.has(owner)) fail(`offers the ${ns} namespace, which is ${nameOf(owner)}'s`);
  }
  for (const e of offer.events) if (EVENT_PACK[e] && !mine.has(EVENT_PACK[e])) fail(`offers the ${e} event, which is ${nameOf(EVENT_PACK[e])}'s`);
  const channels = Object.values(offer.methods).reduce((n, names) => n + names.length, 0);

  // The registry the windows read: the packs loaded, their tiles and kinds of file.
  for (const p of offer.packs) if (!mine.has(p.id)) fail(`the windows are told ${nameOf(p.id)} is installed`);
  const loaded = plan.packs;
  for (const p of loaded) {
    for (const t of p.tiles || []) if (!mine.has(t.app)) fail(`has a tile for ${nameOf(t.app)}`);
    const allowed = new Set(fileAssociations({ apps: [p.id] }).map((a) => a.ext));
    for (const t of p.fileTypes || []) if (!allowed.has(t.ext)) fail(`registers .${t.ext}, which is not ${nameOf(p.id)}'s kind of file`);
    for (const n of p.tools || []) if ((TOOL_PACK[n] || 'ai') !== p.id) fail(`names the AI tool ${n}, which is ${nameOf(TOOL_PACK[n] || 'ai')}'s`);
  }
  const own = new Set([...mine].flatMap((x) => [...(APP_EXTENSIONS[x] || []), ...(OLD_EXTENSIONS[x] || [])]));
  const dialog = openFilters(null, { apps: plan.ids() }).flatMap((f) => f.extensions).filter((e) => e !== '*');
  for (const ext of new Set(dialog)) if (!own.has(ext)) fail(`the Open dialog offers .${ext}`);

  // AI tools: none without AI Tools; with it, its own and this pack's.
  const settings = { [LOADED_KEY]: plan.ids() };
  const offered = toolNames.filter((n) => toolOffered(n, settings));
  for (const n of offered) if (!mine.has(TOOL_PACK[n] || 'ai')) fail(`offers the AI tool ${n}, which is ${nameOf(TOOL_PACK[n] || 'ai')}'s`);
  if (!mine.has('ai') && offered.length) fail(`offers AI tools without AI Tools: ${offered.join(', ')}`);
  // With AI Tools: no tool that reaches into an app that is not there.
  const withAi = { [LOADED_KEY]: [...plan.ids(), 'ai'] };
  for (const n of toolNames.filter((x) => toolOffered(x, withAi))) {
    const missing = (TOOL_NEEDS[n] || []).filter((x) => !withAi[LOADED_KEY].includes(x));
    if (missing.length) fail(`offers the AI tool ${n} with ${missing.map(nameOf).join(' and ')} not installed`);
  }
  const tools = toolNames.filter((n) => (TOOL_PACK[n] || 'ai') === id);
  const reachable = (n) => (TOOL_NEEDS[n] || []).every((x) => withAi[LOADED_KEY].includes(x));
  for (const n of tools.filter(reachable)) if (!toolOffered(n, withAi)) fail(`its own AI tool ${n} is not offered with AI Tools there`);

  return { id, loaded: plan.ids(), channels, tiles: loaded.flatMap((p) => p.tiles || []).length, fileTypes: loaded.flatMap((p) => p.fileTypes || []).length, tools, failures };
}

/* ── the installer's part ───────────────────────────────────────────── */

const baseName = (p) => p.split(/[\\/]/).pop();

/**
 * Each pack's part of the installer include: the files it writes, the kinds
 * of file it registers, the 7-Zip files it carries (`File "/oname=<name>"
 * "<source>"`) and those it unpacks.
 */
export function installerParts(text) {
  const parts = {};
  for (const m of text.matchAll(/!macro rutbaWritePack_([a-z0-9-]+)\r?\n([\s\S]*?)!macroend/g)) {
    parts[m[1]] = {
      files: [...m[2].matchAll(/^\s*File (?:\/r )?"([^"]+)"/gm)].map((f) => baseName(f[1])),
      types: [...m[2].matchAll(/rutbaAssociate "([^"]+)"/g)].map((t) => t[1]),
      carried: [...m[2].matchAll(/^\s*File "\/oname=([^"]+)" "([^"]+)"/gm)].map((f) => ({ name: baseName(f[1]), source: f[2] })),
      extracted: [...m[2].matchAll(/Nsis7z::Extract "([^"]+)"/g)].map((f) => baseName(f[1])),
    };
  }
  return parts;
}

/** The names at the top of a 7-Zip file, with the 7-Zip electron-builder uses; null where it cannot be had. */
export async function sevenZipNames(file) {
  try {
    const require = createRequire(import.meta.url);
    const exe = await require('app-builder-lib/out/toolsets/7zip.js').getPath7za();
    const out = execFileSync(exe, ['l', '-ba', '-slt', file], { encoding: 'utf8', windowsHide: true, maxBuffer: 64 * 1024 * 1024 });
    return [...new Set([...out.matchAll(/^Path = (.+)$/gm)].map((m) => m[1].trim().split(/[\\/]/)[0]))];
  } catch {
    return null;
  }
}

/**
 * What the installer writes for each pack is that pack's alone. `contents`
 * answers the names in a 7-Zip file a part carries ({ source: names }); a
 * carried file not in it is checked by its name only.
 */
export function checkInstaller(text, ids, contents = {}) {
  const failures = [];
  const parts = installerParts(text);
  for (const id of ids) {
    const part = parts[id];
    if (!part) {
      failures.push(`The installer has no part for ${nameOf(id)}.`);
      continue;
    }
    const f = packFiles(id);
    const own = `rutba-pack-${id}.7z`;
    for (const c of part.carried) {
      if (c.name !== own || baseName(c.source) !== `${id}.7z`) failures.push(`The installer's part for ${nameOf(id)} carries ${baseName(c.source)} as ${c.name}.`);
      else if (!part.extracted.includes(own)) failures.push(`The installer's part for ${nameOf(id)} carries ${own} and never unpacks it.`);
      const names = contents[c.source];
      if (names) {
        for (const n of names) if (![f.archive, f.unpacked].includes(n)) failures.push(`The installer's part for ${nameOf(id)} unpacks ${n}.`);
        if (!names.includes(f.archive)) failures.push(`The installer's part for ${nameOf(id)} carries no ${f.archive}.`);
      }
    }
    const carriesOwn = part.carried.some((c) => c.name === own);
    for (const file of part.files) if (![f.archive, f.unpacked, f.manifest].includes(file) && !(file === own && carriesOwn)) failures.push(`The installer's part for ${nameOf(id)} writes ${file}.`);
    if (!part.files.includes(f.manifest)) failures.push(`The installer's part for ${nameOf(id)} does not write its manifest.`);
    const allowed = new Set(fileAssociations({ apps: [id] }).map((a) => a.ext));
    for (const ext of part.types) if (!allowed.has(ext)) failures.push(`The installer's part for ${nameOf(id)} registers .${ext}.`);
  }
  for (const id of Object.keys(parts)) if (!ids.includes(id)) failures.push(`The installer writes ${id}, which is not among the packs.`);
  return { parts, failures };
}

/* ── the whole check ────────────────────────────────────────────────── */

/**
 * The check, from a core, its packs and its installer include.
 * @returns {Promise<{ owners, packs, installer, failures: string[] }>}
 */
export async function checkPacks({ core, packsDir, keysFile = null, installer = null, version = null, scratch = null, outputs = null }) {
  const failures = [];
  const keys = keysFile ? readPackKeys(keysFile) : readPackKeysFromText(readCoreFile(core, 'build/out/pack-keys.json'));
  if (!keys.length) failures.push('The core holds no key to check its packs with.');
  const v = version || JSON.parse(readCoreFile(core, 'package.json')).version;
  const files = coreFiles(core);
  const g = await mapImports();
  const ownerOf = ownerResolver(g, outputs);
  const owners = filesByOwner(files, ownerOf);
  // The renderer files that mix an app's code with another's: the core's, so an app that carries its own code must have none in them.
  const mixed = outputs ? [...rendererOwners(g, asBuild(outputs).outputs)].filter(([out, o]) => files.has(`build/out/${out}`) && ownerOf(`build/out/${out}`) === 'core' && [...o].some(isPack)) : [];
  // The packaged checks (smoke.js and what it seeds) stay; the window checks and the screens build run from the source.
  const unwanted = [...files.keys()].filter((p) => /^main\/(verify-|screen-)/.test(p));
  if (unwanted.length) failures.push(`The core carries the window checks: ${unwanted.slice(0, 5).join(', ')}.`);

  const ids = IDS.filter((id) => fs.existsSync(path.join(packsDir, packFiles(id).manifest)));
  if (!ids.length) failures.push(`There are no packs in ${packsDir}.`);
  const tmp = scratch || fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-check-packs-'));
  const packs = [];
  try {
    for (const id of ids) {
      const alone = checkAlone({ id, packsDir, keys, version: v, scratch: tmp });
      const found = planPacks({ dir: aloneIn(packsDir, [id], path.join(tmp, `${id}-main`)), keys, version: v }).packs.find((p) => p.id === id);
      const builtIn = found ? registersCoreCode(found) : true;
      const inCore = owners[id] || { files: 0, bytes: 0, sample: [] };
      for (const [out, o] of builtIn ? [] : mixed) if (o.has(id)) alone.failures.push(`${nameOf(id)} carries its own code, yet the core's ${out} mixes some of it with ${[...o].filter((x) => x !== id).join(', ')}'s.`);
      if (!builtIn && inCore.files) alone.failures.push(`${nameOf(id)} carries its own code, yet the core still holds ${inCore.files} of its files: ${inCore.sample.join(', ')}.`);
      packs.push({ ...alone, builtIn, inCore: { files: inCore.files, bytes: inCore.bytes } });
      failures.push(...alone.failures);
    }
  } finally {
    if (!scratch) fs.rmSync(tmp, { recursive: true, force: true });
  }
  let installerResult = null;
  if (installer) {
    const text = fs.readFileSync(installer, 'utf8');
    const contents = {};
    for (const part of Object.values(installerParts(text))) {
      for (const c of part.carried) {
        const names = await sevenZipNames(c.source);
        if (names) contents[c.source] = names;
        else failures.push(`The installer carries ${baseName(c.source)}, which could not be read.`);
      }
    }
    installerResult = checkInstaller(text, ids, contents);
    failures.push(...installerResult.failures);
  }
  return { version: v, owners, packs, installer: installerResult, failures };
}

function readPackKeysFromText(text) {
  const file = path.join(os.tmpdir(), `rutba-pack-keys-${process.pid}-${Date.now()}.json`);
  fs.writeFileSync(file, text);
  try {
    return readPackKeys(file);
  } finally {
    fs.rmSync(file, { force: true });
  }
}

const kb = (n) => `${(n / 1024).toFixed(0)} KB`;
const mb = (n) => `${(n / 1024 / 1024).toFixed(1)} MB`;

/** The check's result, on screen. */
export function report(r) {
  const lines = [`Rutba Office ${r.version}: the core's files by owner`];
  const rows = Object.entries(r.owners).sort((a, b) => b[1].bytes - a[1].bytes);
  for (const [o, row] of rows) lines.push(`  ${(isPack(o) ? nameOf(o) : o).padEnd(18)} ${String(row.files).padStart(6)} files  ${mb(row.bytes).padStart(9)}`);
  lines.push('', 'Each pack alone');
  for (const p of r.packs) {
    const code = p.builtIn ? `code still in the core (${p.inCore.files} files, ${kb(p.inCore.bytes)})` : 'carries its own code';
    lines.push(`  ${nameOf(p.id).padEnd(18)} ${String(p.channels).padStart(4)} channels, ${p.tiles} tile, ${p.fileTypes} kinds of file, ${p.tools.length} AI tools; ${code}${p.failures.length ? '  FAILS' : ''}`);
  }
  if (r.installer) lines.push('', `The installer: a part for each of ${Object.keys(r.installer.parts).length} packs`);
  lines.push('', r.failures.length ? `${r.failures.length} failures:` : 'No failures.');
  for (const f of r.failures) lines.push(`  ${f}`);
  return lines.join('\n');
}

function option(args, name) {
  const i = args.indexOf(name);
  return i >= 0 ? args[i + 1] : null;
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const args = process.argv.slice(2);
  const core = option(args, '--core') || path.join(DESKTOP, 'release', 'win-unpacked', 'resources', 'app.asar');
  const packsDir = option(args, '--packs') || path.join(DESKTOP, 'build', 'packs', 'out');
  const installer = option(args, '--installer') || path.join(DESKTOP, 'build', 'packs', 'installer-packs.nsh');
  if (!fs.existsSync(core)) {
    console.error(`No core at ${core}. Build one with npm run dist, or name it with --core.`);
    process.exit(2);
  }
  try {
    const meta = option(args, '--meta') || path.join(DESKTOP, 'build', 'meta', 'renderer.json');
    const outputs = fs.existsSync(meta) ? JSON.parse(fs.readFileSync(meta, 'utf8')) : null;
    const r = await checkPacks({ core, packsDir, keysFile: option(args, '--keys'), installer: fs.existsSync(installer) ? installer : null, version: option(args, '--version'), outputs });
    console.log(report(r));
    const json = option(args, '--json');
    if (json) fs.writeFileSync(json, `${JSON.stringify(r, null, 2)}\n`);
    process.exit(r.failures.length ? 1 : 0);
  } catch (err) {
    console.error(err.message);
    process.exit(1);
  }
}
