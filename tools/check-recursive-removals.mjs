// No recursive removal in Office's own code but removeTree (D-2026-11-13-37).
//
// Electron 44's Node 24 follows a Windows junction in a recursive fs.rm,
// rmSync or rmdir: what a junction inside the folder points at, or what a
// junction given as the path points at, is deleted with it. removeTree
// (@rutba/office-shell/electron/remove-tree) walks the tree and unlinks links
// instead. This refuses, in product code, any other way to remove a tree:
//   - fs.rm, fs.rmSync, fs.promises.rm, fsp.rm, rmdir and rmdirSync (any
//     object's) with `recursive` that is not false;
//   - those calls with options not written out where they are called, which
//     cannot be told apart;
//   - rimraf and fs-extra, which do the same thing a package away.
// Product code is what ships: apps/desktop/main and renderer, and every
// package's src, less the in-app verification runs (verify-*.js). Build
// scripts, tools and tests run on the Node they are given, on a developer's
// or the lab's own folders, and are not checked.
//
// Sites another session is moving are listed by file in ALLOWED with how many
// each holds. An entry with more is a new site; with fewer, its count must
// come down (or the entry go) in the same change that moved them, so the
// list only shrinks.
//
//   node tools/check-recursive-removals.mjs

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

/** The one place a tree is removed: the helper itself. */
export const HELPER = 'packages/office-shell/src/electron/remove-tree.js';

/** Sites their owners are moving to removeTree: file, how many it holds, whose they are. */
export const ALLOWED = Object.freeze({
  'apps/desktop/main/extras/locker/folder.js': { count: 7, owner: 'the Locker session' },
  'apps/desktop/main/extras/locker/index.js': { count: 1, owner: 'the Locker session' },
  'apps/desktop/main/extras/locker/locker.js': { count: 2, owner: 'the Locker session' },
  'apps/desktop/main/extras/locker/working.js': { count: 1, owner: 'the Locker session' },
});

const CODE = /\.(?:js|mjs|cjs|jsx)$/;
const SKIP_DIRS = new Set(['node_modules', 'dist', 'out', 'build', 'test', 'tests', 'vendor', '.gate']);

/** Is this file (a path from the repository's root, with "/") product code? */
export function isProduct(rel) {
  if (!CODE.test(rel) || rel === HELPER) return false;
  if (/(^|\/)verify-[^/]*$/.test(rel)) return false;
  return /^apps\/desktop\/(?:main|renderer)\//.test(rel) || /^packages\/[^/]+\/src\//.test(rel);
}

/**
 * The source with its comments and the words of its strings blanked (the
 * quotes kept, every newline kept), so a call written in a comment or named
 * in a message is not taken for one, and lines still count true.
 */
export function codeOnly(src) {
  const s = String(src);
  let out = '';
  let i = 0;
  let last = ''; // the last character of code that was not a space, to tell a regex from a division
  const blank = (text) => text.replace(/[^\n]/g, ' ');
  while (i < s.length) {
    const c = s[i];
    const next = s[i + 1];
    if (c === '/' && next === '/') {
      const end = s.indexOf('\n', i);
      const stop = end < 0 ? s.length : end;
      out += blank(s.slice(i, stop));
      i = stop;
      continue;
    }
    if (c === '/' && next === '*') {
      const end = s.indexOf('*/', i + 2);
      const stop = end < 0 ? s.length : end + 2;
      out += blank(s.slice(i, stop));
      i = stop;
      continue;
    }
    if (c === '"' || c === "'" || c === '`') {
      let j = i + 1;
      while (j < s.length && s[j] !== c) {
        if (s[j] === '\\') j += 1;
        else if (c !== '`' && s[j] === '\n') break; // an unclosed quote ends with its line
        j += 1;
      }
      out += c + blank(s.slice(i + 1, j)) + (j < s.length && s[j] === c ? c : '');
      i = j < s.length && s[j] === c ? j + 1 : j;
      last = c;
      continue;
    }
    if (c === '/' && (last === '' || '(,=:[!&|?{};+-*%<>~^'.includes(last) || /\b(?:return|typeof|case|in|of)$/.test(out.trimEnd()))) {
      // A regular expression: to its closing "/", a "/" in a [class] not counting.
      let j = i + 1;
      let inClass = false;
      while (j < s.length && s[j] !== '\n') {
        if (s[j] === '\\') j += 1;
        else if (s[j] === '[') inClass = true;
        else if (s[j] === ']') inClass = false;
        else if (s[j] === '/' && !inClass) break;
        j += 1;
      }
      out += '/' + blank(s.slice(i + 1, j)) + (s[j] === '/' ? '/' : '');
      i = s[j] === '/' ? j + 1 : j;
      last = '/';
      continue;
    }
    out += c;
    if (!/\s/.test(c)) last = c;
    i += 1;
  }
  return out;
}

/** The text between a call's "(" at `open` and its matching ")", in code already blanked of strings. */
function argsAt(code, open) {
  let depth = 0;
  for (let j = open; j < code.length && j < open + 8000; j++) {
    if (code[j] === '(' || code[j] === '[' || code[j] === '{') depth += 1;
    else if (code[j] === ')' || code[j] === ']' || code[j] === '}') {
      depth -= 1;
      if (depth === 0) return code.slice(open + 1, j);
    }
  }
  return code.slice(open + 1, open + 8000);
}

/** The arguments split at the commas of their own level. */
function splitArgs(args) {
  const parts = [];
  let depth = 0;
  let from = 0;
  for (let j = 0; j < args.length; j++) {
    const c = args[j];
    if (c === '(' || c === '[' || c === '{') depth += 1;
    else if (c === ')' || c === ']' || c === '}') depth -= 1;
    else if (c === ',' && depth === 0) {
      parts.push(args.slice(from, j));
      from = j + 1;
    }
  }
  parts.push(args.slice(from));
  return parts.map((p) => p.trim()).filter((p, k, all) => p || k < all.length - 1);
}

const lineOf = (code, at) => code.slice(0, at).split('\n').length;

/**
 * Every way this source removes a tree other than removeTree, with its line.
 * @returns {Array<{ line: number, what: string }>}
 */
export function recursiveRemovals(src) {
  const code = codeOnly(src);
  const found = [];
  const call = /(?<![\w$])(rmSync|rmdirSync|rm|rmdir)\s*\(/g;
  for (let m = call.exec(code); m; m = call.exec(code)) {
    const before = code.slice(Math.max(0, m.index - 12), m.index);
    if (/\bfunction\s*$/.test(before) || /\basync\s*$/.test(before) && !/\.\s*$/.test(before)) continue; // its own definition
    const open = m.index + m[0].length - 1;
    const parts = splitArgs(argsAt(code, open));
    const options = parts[1];
    if (options === undefined) continue;
    if (!options.startsWith('{')) {
      found.push({ line: lineOf(code, m.index), what: `${m[1]}() with options not written out: write them out, or use removeTree` });
      continue;
    }
    if (/(?<![\w$])recursive\s*(?::|,|\})/.test(options) && !/(?<![\w$])recursive\s*:\s*false\b/.test(options)) {
      found.push({ line: lineOf(code, m.index), what: `a recursive ${m[1]}(): use removeTree from @rutba/office-shell/electron/remove-tree` });
    }
  }
  const pkg = /(?:\brequire\s*\(\s*|\bimport\s*\(\s*|\bfrom\s+|\bimport\s+)(['"])(rimraf|fs-extra)(?:\/[^'"]*)?\1/g;
  // Package names are in strings, which codeOnly blanks: looked for in the source, on lines that are code.
  for (let m = pkg.exec(src); m; m = pkg.exec(src)) {
    const line = lineOf(src, m.index);
    const codeLine = code.split('\n')[line - 1] || '';
    if (/\b(?:require|import|from)\b/.test(codeLine)) found.push({ line, what: `${m[2]}: use removeTree from @rutba/office-shell/electron/remove-tree` });
  }
  return found.sort((a, b) => a.line - b.line);
}

/** The product files under the repository's root, as paths from it with "/". */
export function productFiles(root = ROOT) {
  const out = [];
  const walk = (dir) => {
    let entries;
    try {
      entries = fs.readdirSync(dir, { withFileTypes: true });
    } catch {
      return;
    }
    for (const e of entries) {
      if (e.isSymbolicLink()) continue;
      const full = path.join(dir, e.name);
      const rel = path.relative(root, full).split(path.sep).join('/');
      if (e.isDirectory()) {
        if (!SKIP_DIRS.has(e.name) && !e.name.startsWith('.')) walk(full);
      } else if (isProduct(rel)) out.push(rel);
    }
  };
  for (const start of ['apps/desktop/main', 'apps/desktop/renderer', 'packages']) walk(path.join(root, start));
  return out.sort();
}

/**
 * The faults: none when product code removes trees only through removeTree,
 * and the allow-list says exactly what is left to move.
 * @param {{ root?: string, allowed?: Record<string, { count: number, owner: string }>, read?: (rel: string) => string, files?: string[] }} [o]
 *   `files`: the product files to look at, found under `root` when not given
 * @returns {string[]}
 */
export function removalFaults({ root = ROOT, allowed = ALLOWED, read = (rel) => fs.readFileSync(path.join(root, rel), 'utf8'), files = productFiles(root) } = {}) {
  const faults = [];
  for (const rel of files) {
    const found = recursiveRemovals(read(rel));
    const allow = allowed[rel];
    if (allow) {
      if (found.length > allow.count) faults.push(`${rel}: ${found.length} recursive removals where ${allow.count} are listed for ${allow.owner}: a new one (lines ${found.map((f) => f.line).join(', ')}): use removeTree`);
      else if (found.length < allow.count) faults.push(`${rel}: ${found.length} recursive removals left of the ${allow.count} listed for ${allow.owner}: bring its count in ALLOWED down${found.length ? '' : ', or drop the entry'} (tools/check-recursive-removals.mjs)`);
      continue;
    }
    for (const f of found) faults.push(`${rel}:${f.line}: ${f.what}`);
  }
  for (const rel of Object.keys(allowed)) if (!files.includes(rel)) faults.push(`${rel}: listed in ALLOWED but not found: drop the entry`);
  return faults;
}

const isMain = process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (isMain) {
  const faults = removalFaults();
  if (faults.length) {
    process.stderr.write(`Recursive removals outside removeTree (D-2026-11-13-37):\n${faults.map((f) => `  ${f}`).join('\n')}\n`);
    process.exitCode = 1;
  } else {
    const left = Object.entries(ALLOWED).map(([rel, a]) => `${rel} (${a.count}, ${a.owner})`);
    process.stdout.write(`No recursive removal outside removeTree${left.length ? `; still to move: ${left.join('; ')}` : ''}\n`);
  }
}
