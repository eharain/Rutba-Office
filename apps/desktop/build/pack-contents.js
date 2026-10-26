// What an app's pack carries once its code moves out of the core
// (docs/APP-PACKS.md 5.2), read off the map (tools/app-map.mjs) and the
// renderer's build (build/bundle.js writes what went into each file):
//
//   - its main-process files: those the map gives the app, its register
//     module among them, kept as they lie under main/;
//   - its window chunks: each chunk of the renderer whose every input is
//     the app's (the chunk its window module starts, and any only it uses);
//   - its npm packages and engines: those only the app reaches.
//
// And what the core leaves out for it: the same files, by their paths in
// the core's archive, and the app's tile, which the pack has.
//
// A move that cannot be made is refused with the reason: a chunk that mixes
// the app's code with another's (the import to cut is named), or a file of
// the app that imports one of the core's by a relative path (which would not
// be there beside it in the pack; the core's own is reached by its package
// name, or handed over in the context the core registers it with).

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { PACKS } from '@rutba/office-shell/contract';
import { mapImports, homeOf, registerModules, rendererEntries, ROOT } from '../../../tools/app-map.mjs';
import { RENDERER_META } from './bundle.js';

const PACK_IDS = new Set(PACKS.map((p) => p.id));

/**
 * The bundles of their own a window starts by name (build/bundle.js), by the
 * source each starts from: the pack whose code it is carries it beside its
 * chunks, and the app's protocol finds it there when the core has none.
 */
const WORKER_BUNDLES = {
  'scripts-worker.js': 'apps/desktop/renderer/apps/sheets/scripts-worker.js',
  'image-worker.js': 'apps/desktop/renderer/imaging/worker.js',
};

const APP = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const MAIN = 'apps/desktop/main/';
const OUT = 'apps/desktop/build/out/';

/** The renderer's outputs and their inputs, as the last build wrote them. */
export function readRendererMeta(file = RENDERER_META) {
  return readRendererBuild(file).outputs;
}

/** All the last build said: each output's inputs, the outputs it imports, and the entry each starts from. */
export function readRendererBuild(file = RENDERER_META) {
  try {
    const m = JSON.parse(fs.readFileSync(file, 'utf8'));
    return { outputs: m.outputs || {}, imports: m.imports || null, entries: m.entries || null };
  } catch {
    throw new Error(`The renderer's build has not said what went into its files (${path.relative(APP, file)}). Run node build/bundle.js first.`);
  }
}

/** The renderer's build in its full form, from what a caller hands over: the build, or its outputs alone. */
export const asBuild = (meta) => (meta && meta.outputs ? meta : { outputs: meta || {}, imports: null, entries: null });

/**
 * Which windows reach each output of the renderer: the app whose window
 * module starts it, or 'core' for the launcher's bundle and anything else no
 * import reaches. A window's walk stops at another app's window module. An
 * output only one app's window reaches holds code nothing else runs, whatever
 * the map says of its files' owners (a file the main process shares, say):
 * it goes with that app. Null when the build did not say what imports what.
 * @param {{ outputs, imports, entries }} build
 * @param {Record<string, string>} windows each app's window module, by the app's id
 */
export function windowReach(build, windows) {
  const { outputs, imports, entries } = build;
  if (!imports || !entries) return null;
  const byFile = Object.fromEntries(Object.entries(windows).map(([id, file]) => [file, id]));
  const rootOf = {};
  for (const [out, entry] of Object.entries(entries)) if (byFile[entry]) rootOf[out] = byFile[entry];
  const imported = new Set(Object.values(imports).flatMap((list) => list.map((i) => i.path)));
  const roots = Object.keys(entries).map((out) => [out, rootOf[out] || (imported.has(out) ? null : 'core')]).filter(([, o]) => o);
  const reach = new Map(Object.keys(outputs).map((o) => [o, new Set()]));
  for (const [root, owner] of roots) {
    const seen = new Set();
    const queue = [root];
    while (queue.length) {
      const o = queue.pop();
      if (seen.has(o)) continue;
      seen.add(o);
      reach.get(o)?.add(owner);
      for (const i of imports[o] || []) if (!rootOf[i.path] || rootOf[i.path] === owner) queue.push(i.path);
    }
  }
  for (const set of reach.values()) if (!set.size) set.add('core');
  return reach;
}

/** Each app's window module, from the launcher's list (the launcher's own is the core's). */
export function appWindows() {
  return Object.fromEntries(Object.entries(rendererEntries()).filter(([id]) => PACK_IDS.has(id)));
}

/**
 * A renderer input's owner, by the map: an npm package's by the package, a
 * file the map never reached by its home.
 */
export function inputOwnerFor(map) {
  const ownerOf = (file) => (file.startsWith('npm:') ? map.packageOwner.get(file) : map.ownerOf(file));
  return (i) => {
    const npm = /^(?:.*\/)?node_modules\/((?:@[^/]+\/)?[^/]+)/.exec(i);
    const o = npm ? map.packageOwner.get(`npm:${npm[1]}`) : ownerOf(i);
    return o && o !== 'unreached' ? o : npm ? 'core' : homeOf(i);
  };
}

/** Each renderer output's owners: every owner of the inputs that went into it. */
export function rendererOwners(map, outputs) {
  const inputOwner = inputOwnerFor(map);
  return new Map(Object.entries(outputs).map(([out, inputs]) => [out, new Set(inputs.map(inputOwner))]));
}

/** A workspace package's files, by their paths in it: what a copy of it ships (no tests, notes or maps). */
function packageFiles(dir) {
  const out = [];
  const walk = (abs, prefix) => {
    for (const ent of fs.readdirSync(abs, { withFileTypes: true })) {
      if (ent.isDirectory() && ['test', 'tests', '.git', 'node_modules'].includes(ent.name)) continue;
      if (!ent.isDirectory() && /\.(md|map)$/i.test(ent.name)) continue;
      const p = prefix ? `${prefix}/${ent.name}` : ent.name;
      if (ent.isDirectory()) walk(path.join(abs, ent.name), p);
      else out.push(p);
    }
  };
  walk(path.join(ROOT, 'packages', dir), '');
  return out.sort();
}

/** The relative imports a file makes. */
function relativeImports(text) {
  const out = [];
  for (const m of text.matchAll(/(?:\bfrom\s*|\bimport\s*\(\s*|\bimport\s+)['"](\.{1,2}\/[^'"]+)['"]/g)) out.push(m[1]);
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
 * Each npm package's owner, with what it needs: a package one owner's
 * packages alone need is that owner's; one two owners need (iconv-lite, which
 * Mail's parsers and the MySQL client share) is the core's, and stays in it.
 * The same reckoning the release check makes (tools/check-packs.mjs).
 */
export function npmOwners(map) {
  const owner = new Map();
  for (const [h, o] of map.packageOwner) if (h.startsWith('npm:')) owner.set(h.slice(4), o);
  const pending = [...owner.keys()];
  while (pending.length) {
    const name = pending.pop();
    for (const dep of dependenciesOf(name)) {
      const was = owner.get(dep);
      const now = was === undefined || was === owner.get(name) ? owner.get(name) : 'core';
      if (was !== now) {
        owner.set(dep, now);
        pending.push(dep);
      }
    }
  }
  return owner;
}

/**
 * What the packs named carry, and what the core leaves out for them.
 * @param {string[]} ids
 * @param {{ g?: object, meta?: object }} [o] the map and the renderer's outputs, when already read
 * @returns {Promise<Record<string, { main: string[], register: string, chunks: string[], npm: string[], packages: { name: string, dir: string }[], core: string[] }>>}
 */
export async function packContents(ids, { g = null, meta = null } = {}) {
  if (!ids.length) return {};
  const map = g || (await mapImports());
  const build = meta ? asBuild(meta) : readRendererBuild();
  const { outputs } = build;
  const reach = windowReach(build, appWindows());
  const npmByOwner = npmOwners(map);
  const registers = registerModules();
  const ownerOf = (file) => (file.startsWith('npm:') ? map.packageOwner.get(file) : map.ownerOf(file));
  const workspace = {};
  for (const dir of fs.readdirSync(path.join(ROOT, 'packages'))) {
    try {
      workspace[dir] = JSON.parse(fs.readFileSync(path.join(ROOT, 'packages', dir, 'package.json'), 'utf8')).name;
    } catch {
      // not a package
    }
  }
  const inputOwner = inputOwnerFor(map);
  const result = {};
  for (const id of ids) {
    const problems = [];
    const register = registers[id];
    if (!register) problems.push(`${id} has no register module in main/app-packs.js.`);
    const main = [...new Set([...map.graphs.main.keys(), register].filter(Boolean))]
      .filter((f) => f.startsWith(MAIN) && homeOf(f) !== 'checks' && ownerOf(f) === id)
      .sort();
    for (const file of main) {
      const text = fs.readFileSync(path.join(ROOT, file), 'utf8');
      for (const spec of relativeImports(text)) {
        const to = path.posix.normalize(path.posix.join(path.posix.dirname(file), spec));
        if (!main.includes(to)) problems.push(`${file} imports ${spec}, which stays in the core: reach it by its package name or through the context the core hands the pack.`);
      }
    }
    // Nothing left in the core may import a file the pack takes, but for the register module the core loads it by.
    for (const [file, edges] of map.graphs.main) {
      if (main.includes(file) || file.startsWith('npm:') || homeOf(file) === 'checks') continue;
      for (const e of edges) {
        if (main.includes(e.to) && e.to !== register) problems.push(`${file} imports ${e.to}, which would go with ${id}'s pack: move it to the core, or reach it through ${id}'s service.`);
      }
    }
    const chunks = [];
    for (const [out, inputs] of Object.entries(outputs)) {
      const owners = new Set(inputs.map(inputOwner));
      if (reach) {
        // By the windows that reach it: one only this app's window reaches goes with it, unless it holds another app's code.
        const by = reach.get(out) || new Set(['core']);
        const others = [...owners].filter((o) => PACK_IDS.has(o) && o !== id);
        if (by.size === 1 && by.has(id)) {
          if (!out.startsWith('chunks/')) problems.push(`${out} is reached only by ${id}'s window, but it is not a chunk the core imports when a window asks for it.`);
          else if (others.length) problems.push(`${out}, which only ${id}'s window reaches, holds ${others.join(', ')}'s code (${inputs.filter((i) => others.includes(inputOwner(i))).slice(0, 3).join(', ')}); cut the import that brings it in.`);
          else chunks.push(out);
        } else if (owners.has(id)) {
          const mine = inputs.filter((i) => inputOwner(i) === id);
          problems.push(`${out} holds ${id}'s code (${mine.slice(0, 3).join(', ')}) and is reached by ${[...by].filter((o) => o !== id).join(', ')} too; cut the import that joins them.`);
        }
        continue;
      }
      if (owners.size === 1 && owners.has(id)) {
        if (!out.startsWith('chunks/')) problems.push(`${out} is all ${id}'s code, but it is not a chunk the core imports when a window asks for it.`);
        else chunks.push(out);
      } else if (owners.has(id)) {
        const mine = inputs.filter((i) => inputOwner(i) === id);
        problems.push(`${out} mixes ${id}'s code (${mine.slice(0, 3).join(', ')}) with ${[...owners].filter((o) => o !== id).join(', ')}'s; cut the import that joins them.`);
      }
    }
    if (!chunks.length && appWindows()[id]) problems.push(`No chunk of the renderer is ${id}'s alone.`);
    const workers = Object.entries(WORKER_BUNDLES).filter(([, src]) => ownerOf(src) === id).map(([out]) => out);
    const npm = [...npmByOwner].filter(([n, o]) => o === id && fs.existsSync(path.join(ROOT, 'node_modules', n, 'package.json'))).map(([n]) => n).sort();
    const packages = [...map.packageOwner].filter(([h, o]) => h.startsWith('pkg:') && o === id).map(([h]) => ({ dir: h.slice(4), name: workspace[h.slice(4)] })).filter((p) => p.name);
    // An engine split by file (the map's SPLIT_PACKAGES: text recognition's,
    // the templates): the pack carries the package's code whole and the data
    // that is this app's; the core leaves out the files only this app reaches.
    const split = [];
    for (const [dir, name] of Object.entries(workspace)) {
      if (map.packageOwner.get(`pkg:${dir}`) === id) continue;
      const files = packageFiles(dir);
      const mine = files.filter((f) => ownerOf(`packages/${dir}/${f}`) === id);
      if (!mine.length || !files.some((f) => ownerOf(`packages/${dir}/${f}`) !== id)) continue;
      const code = (f) => /\.(m?js|cjs|json)$/i.test(f) && !/^(vendor|data-files|files|thumbs)\//.test(f);
      split.push({ name, dir, files: files.filter((f) => mine.includes(f) || code(f)), mine });
    }
    if (problems.length) throw new Error(`${id} cannot move into its pack yet:\n  ${problems.join('\n  ')}`);
    const core = [
      ...main.map((f) => f.slice('apps/desktop/'.length)),
      ...chunks.map((c) => `build/out/${c}`),
      ...workers.map((w) => `build/out/${w}`),
      `resources/apps/${id}.ico`,
      ...npm.map((n) => `node_modules/${n}/**`),
      ...packages.map((p) => `node_modules/${p.name}/**`),
      ...split.flatMap((s) => s.mine.map((f) => `node_modules/${s.name}/${f}`)),
    ];
    result[id] = { main, register, chunks, workers, npm, packages, split: split.map(({ name, dir, files }) => ({ name, dir, files })), core };
  }
  return result;
}

export { OUT };
