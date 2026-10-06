// No engine splices data into XML through a replacement string.
//
// `xml.replace(old, newXml)` reads `$&`, `$'`, `` $` `` and `$$` in newXml as
// instructions. Whenever newXml carries anything a person typed or a file
// held — a cell, a comment, a sheet or section name, a link, a subject — those
// four corrupt the part: a workbook lost its sheets to "Sales $'000", a comment
// took the rest of the comments part into itself. Every such splice now hands
// `replace` a function, which inserts its result as it is.
//
// This reads the sources and fails on any `.replace(pattern, replacement)`
// whose replacement is a computed string rather than a function or a fixed
// literal, so the splice cannot come back unnoticed. A fixed literal may use
// `$1` on purpose: it carries no data. The few calls below are not
// String.prototype.replace at all, or insert a constant.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const BS = String.fromCharCode(92);

/** Allowed, by file and the replacement as written: engine methods named replace, and a constant. */
const ALLOWED = new Set([
  'apps/desktop/main/documents.js|String.fromCharCode(10)',
  'apps/desktop/main/documents.js|a.replace',
  'apps/desktop/main/documents.js|a.replacement',
  'packages/presentation/src/deck.js|replacement',
  'packages/proofing/src/deck.js|e.text',
]);

function sources() {
  const out = [];
  const walk = (dir) => {
    for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
      const p = path.join(dir, e.name);
      if (e.isDirectory()) { if (e.name !== 'node_modules') walk(p); }
      else if (e.name.endsWith('.js')) out.push(p);
    }
  };
  for (const pkg of fs.readdirSync(path.join(root, 'packages'))) {
    const src = path.join(root, 'packages', pkg, 'src');
    if (fs.existsSync(src)) walk(src);
  }
  for (const f of fs.readdirSync(path.join(root, 'apps/desktop/main'))) {
    if (f.endsWith('.js') && !f.startsWith('verify')) out.push(path.join(root, 'apps/desktop/main', f));
  }
  return out;
}

/** The top-level arguments of a call whose '(' ends just before `i`. */
function argsAt(src, i) {
  const out = [];
  let depth = 0, start = i, str = null;
  const tpl = [];
  for (let j = i; j < src.length; j++) {
    const c = src[j];
    if (str) {
      if (c === BS) { j++; continue; }
      if (str === '`' && c === '$' && src[j + 1] === '{') { tpl.push(depth); depth++; str = null; j++; continue; }
      if (c === str) str = null;
      continue;
    }
    if (c === "'" || c === '"' || c === '`') { str = c; continue; }
    if (c === '/' && /[(,=:[!&|?{};]\s*$/.test(src.slice(Math.max(0, j - 20), j))) {
      let cls = false;
      for (j++; j < src.length; j++) {
        const d = src[j];
        if (d === BS) { j++; continue; }
        if (d === '[') cls = true; else if (d === ']') cls = false;
        else if (d === '/' && !cls) break;
      }
      continue;
    }
    if (c === '(' || c === '[' || c === '{') depth++;
    else if (c === ')' || c === ']' || c === '}') {
      if (c === '}' && tpl.length && tpl[tpl.length - 1] === depth - 1) { tpl.pop(); depth--; str = '`'; continue; }
      if (depth === 0) { out.push(src.slice(start, j).trim()); return out; }
      depth--;
    } else if (c === ',' && depth === 0) { out.push(src.slice(start, j).trim()); start = j + 1; }
  }
  return out;
}

const isFunction = (e) => /^(\([^)]*\)|[A-Za-z_$][\w$]*)\s*=>/.test(e) || /^(async\s+)?function\b/.test(e);
const isFixedLiteral = (e) => /^'(?:[^'\\]|\\.)*'$/.test(e) || /^"(?:[^"\\]|\\.)*"$/.test(e) || (/^`[^`]*`$/.test(e) && !e.includes('${'));

test('every splice into a part inserts its text as written', () => {
  const offending = [];
  for (const file of sources()) {
    const rel = path.relative(root, file).split(path.sep).join('/');
    const src = fs.readFileSync(file, 'utf8');
    for (const m of src.matchAll(/\.replace(?:All)?\(/g)) {
      const args = argsAt(src, m.index + m[0].length);
      if (args.length < 2) continue;
      const rep = args[1];
      if (isFunction(rep) || isFixedLiteral(rep) || ALLOWED.has(`${rel}|${rep}`)) continue;
      const line = src.slice(0, m.index).split('\n').length;
      offending.push(`${rel}:${line}: ${rep.replace(/\s+/g, ' ').slice(0, 100)}`);
    }
  }
  assert.deepEqual(offending, [], 'a replacement built from data must be a function — () => text — or $-patterns in it act on the data:\n' + offending.join('\n'));
});
