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
import { mapImports, homeOf, registerModules, ROOT } from '../../../tools/app-map.mjs';
import { RENDERER_META } from './bundle.js';

const APP = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const MAIN = 'apps/desktop/main/';
const OUT = 'apps/desktop/build/out/';

/** The renderer's outputs and their inputs, as the last build wrote them. */
export function readRendererMeta(file = RENDERER_META) {
  try {
    return JSON.parse(fs.readFileSync(file, 'utf8')).outputs || {};
  } catch {
    throw new Error(`The renderer's build has not said what went into its files (${path.relative(APP, file)}). Run node build/bundle.js first.`);
  }
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

/** The relative imports a file makes. */
function relativeImports(text) {
  const out = [];
  for (const m of text.matchAll(/(?:\bfrom\s*|\bimport\s*\(\s*|\bimport\s+)['"](\.{1,2}\/[^'"]+)['"]/g)) out.push(m[1]);
  return out;
}

/** An npm package and what it needs, by name, from the repository's node_modules. */
function npmClosure(name, seen = new Set()) {
  if (seen.has(name)) return seen;
  const file = path.join(ROOT, 'node_modules', name, 'package.json');
  if (!fs.existsSync(file)) return seen;
  seen.add(name);
  for (const dep of Object.keys(JSON.parse(fs.readFileSync(file, 'utf8')).dependencies || {})) npmClosure(dep, seen);
  return seen;
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
  const outputs = meta || readRendererMeta();
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
    const chunks = [];
    for (const [out, inputs] of Object.entries(outputs)) {
      const owners = new Set(inputs.map(inputOwner));
      if (owners.size === 1 && owners.has(id)) {
        if (!out.startsWith('chunks/')) problems.push(`${out} is all ${id}'s code, but it is not a chunk the core imports when a window asks for it.`);
        else chunks.push(out);
      } else if (owners.has(id)) {
        const mine = inputs.filter((i) => inputOwner(i) === id);
        problems.push(`${out} mixes ${id}'s code (${mine.slice(0, 3).join(', ')}) with ${[...owners].filter((o) => o !== id).join(', ')}'s; cut the import that joins them.`);
      }
    }
    if (!chunks.length && fs.existsSync(path.join(ROOT, 'apps/desktop/renderer/apps', `${id}.js`))) problems.push(`No chunk of the renderer is ${id}'s alone.`);
    const npm = [...new Set([...map.packageOwner].filter(([h, o]) => h.startsWith('npm:') && o === id).flatMap(([h]) => [...npmClosure(h.slice(4))]))].sort();
    const packages = [...map.packageOwner].filter(([h, o]) => h.startsWith('pkg:') && o === id).map(([h]) => ({ dir: h.slice(4), name: workspace[h.slice(4)] })).filter((p) => p.name);
    if (problems.length) throw new Error(`${id} cannot move into its pack yet:\n  ${problems.join('\n  ')}`);
    const core = [
      ...main.map((f) => f.slice('apps/desktop/'.length)),
      ...chunks.map((c) => `build/out/${c}`),
      `resources/apps/${id}.ico`,
      ...npm.map((n) => `node_modules/${n}/**`),
      ...packages.map((p) => `node_modules/${p.name}/**`),
    ];
    result[id] = { main, register, chunks, npm, packages, core };
  }
  return result;
}

export { OUT };
