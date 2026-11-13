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
//     cannot be told apart: a variable, a spread ({ ...OPTS }), or a quoted
//     or computed key ({ 'recursive': true });
//   - rm, rmSync, rmdir or rmdirSync taken out of fs by name: imported from
//     fs, fs/promises or original-fs, destructured from any object, or taken
//     as a value (`const del = fs.rmSync`, `{ rm: fs.rmSync }`), since the
//     calls under another name are not seen;
//   - rimraf and fs-extra, which do the same thing a package away.
// A method's own definition (`rm(p, o) {` in a class or an object) is not a
// call. Known limits, not looked for: bracket access (fs['rmSync']), a
// statement that starts with .call or .apply (fs.rmSync.call(fs, p, o)), and
// shell commands (rd /s, rmdir /s, rm -rf, Remove-Item -Recurse) run through
// child_process.
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
export const ALLOWED = Object.freeze({});

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

/** The text between a call's "(" at `open` and its matching ")" at `end`, in code already blanked of strings. */
function argsAt(code, open) {
  let depth = 0;
  for (let j = open; j < code.length && j < open + 8000; j++) {
    if (code[j] === '(' || code[j] === '[' || code[j] === '{') depth += 1;
    else if (code[j] === ')' || code[j] === ']' || code[j] === '}') {
      depth -= 1;
      if (depth === 0) return { args: code.slice(open + 1, j), end: j };
    }
  }
  return { args: code.slice(open + 1, open + 8000), end: -1 };
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

/** The names that remove: as fs, fs/promises and original-fs have them. */
const REMOVERS = new Set(['rm', 'rmSync', 'rmdir', 'rmdirSync']);

/**
 * Options that cannot be read where they are written: not an object literal,
 * or one with a spread, or a quoted or computed key at its own level.
 */
function unreadable(options) {
  if (!options.startsWith('{')) return true;
  const close = options.lastIndexOf('}');
  if (close < 0 || options.slice(close + 1).trim()) return true;
  return splitArgs(options.slice(1, close)).some((p) => p.startsWith('...') || /^['"`[]/.test(p));
}

/** The names taken out of a brace list: `a, b as c` (an import) or `a, b: c, d = 1` (destructuring). */
function namesIn(list, sep) {
  return list.split(',').map((p) => p.trim().split(sep)[0].trim()).filter(Boolean);
}

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
    const { args, end } = argsAt(code, open);
    // A method's own definition, in a class or an object: rm(p, o) { ... }
    if (end >= 0 && /^\s*\{/.test(code.slice(end + 1, end + 40)) && !/\.\s*$/.test(before)) continue;
    const parts = splitArgs(args);
    const options = parts[1];
    if (options === undefined) continue;
    if (unreadable(options)) {
      found.push({ line: lineOf(code, m.index), what: `${m[1]}() with options not written out: write them out, or use removeTree` });
      continue;
    }
    if (/(?<![\w$])recursive\s*(?::|,|\})/.test(options) && !/(?<![\w$])recursive\s*:\s*false\b/.test(options)) {
      found.push({ line: lineOf(code, m.index), what: `a recursive ${m[1]}(): use removeTree from @rutba/office-shell/electron/remove-tree` });
    }
  }
  // Taken out by name, so that later calls go by another name or none.
  const isCode = (i) => code[i] === src[i];
  const imports = /\bimport\s*\{([^}]*)\}\s*from\s*(['"])(?:node:)?(fs|fs\/promises|original-fs)\2/g;
  for (let m = imports.exec(src); m; m = imports.exec(src)) {
    if (!isCode(m.index)) continue;
    for (const name of namesIn(m[1], /\s+as\s+/)) {
      if (REMOVERS.has(name)) found.push({ line: lineOf(src, m.index), what: `${name} imported from ${m[3]} by name: use removeTree, or call fs.${name} so its options are seen` });
    }
  }
  const destructured = /\b(?:const|let|var)\s*\{([^}]*)\}\s*=(?!=)/g;
  for (let m = destructured.exec(code); m; m = destructured.exec(code)) {
    for (const name of namesIn(m[1], /[:=]/)) {
      if (REMOVERS.has(name)) found.push({ line: lineOf(code, m.index), what: `${name} destructured: use removeTree, or call fs.${name} so its options are seen` });
    }
  }
  const asValue = /(?:(?<![=!<>])=(?![=>])|[:,([?]|&&|\|\||\breturn)\s*[\w$]+(?:\s*\.\s*[\w$]+)*\s*\.\s*(rmSync|rmdirSync|rm|rmdir)\b(?!\s*[(\w$])/g;
  for (let m = asValue.exec(code); m; m = asValue.exec(code)) {
    found.push({ line: lineOf(code, m.index), what: `${m[1]} taken as a value: use removeTree, or call it where it is so its options are seen` });
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
