// The words a window shows that do not go through the message catalogue:
// English left where a translation can never reach it.
//
//   node tools/find-english.mjs                 every window file, a count each
//   node tools/find-english.mjs <file> [...]    those files, each finding by line
//
// It reads the source, not a parse of it, so it looks for the shapes words
// take in this code — an attribute a person reads (label="…", title="…",
// placeholder="…"), a property that is shown (label: '…', message: '…'), a
// toast, and words standing between two tags — and it says where. A finding
// that is not words (a font's name, a format code) is marked where it is
// written, with a comment on the same line or the line before:
//
//   // words-ok: a font's name
//
// tests/english-left.test.js holds the window files to none.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const SOURCES = ['packages/office-ui/src', 'apps/desktop/renderer'];

/** Files whose strings are not shown words: a worker's code, glyph tables, the catalogues themselves. */
export const NOT_WORDS = new Set([
  'packages/office-ui/src/catalogues/index.js',
  'packages/office-ui/src/messages.js',
  'packages/office-ui/src/icons.js',
  'apps/desktop/renderer/apps/sheets/scripts-worker.js',
]);

const ATTRS = 'label|title|placeholder|aria-label|aria-description|hint|alt|data-tip|heading|caption';
const PROPS = 'label|title|heading|hint|message|detail|description|tip|placeholder|caption|subtitle|confirm|empty';
const hasWords = (s) => /[A-Za-z]{2,}/.test(s) && !/^[a-z][\w.-]*$/.test(s) && !/^(https?:|mailto:|data:|#[0-9a-f]{3,8}$)/i.test(s);

/** The findings in one file's text: { line, text, kind }. */
export function findEnglish(text) {
  const lines = text.split(/\r?\n/);
  const out = [];
  const okAt = (i) => /words-ok/.test(lines[i] || '') || /words-ok/.test(lines[i - 1] || '');
  lines.forEach((line, i) => {
    if (okAt(i)) return;
    const code = line.replace(/^\s*(\/\/|\*|\/\*).*$/, '');
    if (!code.trim()) return;
    for (const m of code.matchAll(new RegExp(String.raw`(?<![\w-])(?:${ATTRS})="([^"]*)"`, 'g'))) if (hasWords(m[1])) out.push({ line: i + 1, text: m[1], kind: 'attribute' });
    for (const m of code.matchAll(new RegExp(String.raw`(?<![\w$.])(?:${PROPS})\s*:\s*(?:'((?:[^'\\]|\\.)*)'|"((?:[^"\\]|\\.)*)"|\x60([^\x60$]*)\x60)`, 'g'))) {
      const s = m[1] ?? m[2] ?? m[3];
      if (hasWords(s)) out.push({ line: i + 1, text: s, kind: 'property' });
    }
    for (const m of code.matchAll(/(?<![\w$.])toast\(\s*(?:'((?:[^'\\]|\\.)*)'|"((?:[^"\\]|\\.)*)"|\x60((?:[^\x60\\]|\\.)*)\x60)/g)) {
      const s = m[1] ?? m[2] ?? m[3];
      if (hasWords(s)) out.push({ line: i + 1, text: s, kind: 'toast' });
    }
    // Words between two tags on one line: >Save<, >Save as {name}<.
    for (const m of code.matchAll(/>([^<>{}`'"=;]*[A-Za-z]{2,}[^<>{}`'"=;]*)</g)) {
      const s = m[1].trim();
      // Code between a > and a < is not words: a comparison, a ternary, an index, an optional chain.
      const looksLikeCode = /&&|\|\||=>|\?\.|\?\?|[[\]]|\s\?\s|^:\s|\s:\s|^[\w.]+\s*[<>]=?/.test(s);
      if (s && hasWords(s) && !looksLikeCode) out.push({ line: i + 1, text: s, kind: 'text' });
    }
    // Words on a line of their own inside an element: the line before ends a tag.
    const bare = code.trim();
    if (/^[A-Z][A-Za-z’'.,:;!?() -]*[A-Za-z.!?:)]$/.test(bare) && bare.includes(' ') && !/[{}=;<>`"]/.test(bare)) {
      const before = lines.slice(0, i).reverse().find((l) => l.trim());
      if (before && />\s*$/.test(before) && !/=>\s*$/.test(before)) out.push({ line: i + 1, text: bare, kind: 'text' });
    }
  });
  return out;
}

/** Every window file: { file, findings }. */
export function scan(root = ROOT) {
  const files = [];
  const walk = (dir) => {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      const p = path.join(dir, entry.name);
      if (entry.isDirectory()) walk(p);
      else if (entry.name.endsWith('.js')) files.push(p);
    }
  };
  for (const dir of SOURCES) walk(path.join(root, dir));
  return files
    .map((p) => ({ file: path.relative(root, p).split(path.sep).join('/'), path: p }))
    .filter((f) => !NOT_WORDS.has(f.file))
    .map((f) => ({ file: f.file, findings: findEnglish(fs.readFileSync(f.path, 'utf8')) }));
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const asked = process.argv.slice(2);
  if (asked.length) {
    let total = 0;
    for (const file of asked) {
      const findings = findEnglish(fs.readFileSync(path.resolve(file), 'utf8'));
      total += findings.length;
      for (const f of findings) console.log(`${file}:${f.line}  ${f.kind}  ${f.text}`);
    }
    console.log(`${total} left in English`);
    process.exit(total ? 1 : 0);
  }
  const all = scan().filter((f) => f.findings.length).sort((a, b) => b.findings.length - a.findings.length);
  for (const f of all) console.log(`${String(f.findings.length).padStart(5)}  ${f.file}`);
  console.log(`${all.reduce((n, f) => n + f.findings.length, 0)} left in English, in ${all.length} files`);
}
