// The window checks in groups small enough to run alone (owner, 11 Nov
// 2026: "re group tests to smaller groups to be independently runnable ones
// and resumable"). Every check is in a part or a named section; no part sits
// inside a section, so asking for a part runs that part and nothing else; and
// every section and part has an id of its own.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { sectionsIn, matchBrace } from '../tools/verify-sections.js';
import { partSlug } from '../apps/desktop/main/verify-select.js';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const src = fs.readFileSync(path.join(ROOT, 'apps', 'desktop', 'main', 'verify-apps.js'), 'utf8').replace(/\r\n/g, '\n');
const sections = sectionsIn(src);

test('the scanner steps over strings, templates, comments and regular expressions', () => {
  const code = "{ const a = '}'; const b = `}${ { x: 1 }.x }`; // }\n /* } */ const r = /\\}/; if (a) { b; } }";
  assert.equal(matchBrace(code, 0), code.length - 1);
  const s = sectionsIn("x;\n  await section('mail', 'reading', async () => {\n    check('a }', true);\n  });\n");
  assert.deepEqual(s.map((x) => [x.app, x.name]), [['mail', 'reading']]);
  assert.match(s[0].body, /check\('a \}'/);
});

test('every inline section is named and found whole, and the old unnamed form is gone', () => {
  assert.ok(sections.length >= 30, `${sections.length} sections`);
  for (const s of sections) assert.ok(s.close > s.open, `${s.app}#${s.name} has no end`);
  assert.equal(/\bif \(runs\('/.test(src), false, 'a section written as if (runs(...)) has no name and no end');
});

test('no part sits inside a section: a part asked for runs alone', () => {
  const inside = sections.filter((s) => /\bpart\(\s*'/.test(s.body)).map((s) => `${s.app}#${s.name}`);
  assert.deepEqual(inside, []);
  const nested = sections.filter((s) => /\bawait section\(/.test(s.body)).map((s) => `${s.app}#${s.name}`);
  assert.deepEqual(nested, []);
});

test('each section and part has an id of its own', () => {
  const ids = new Map();
  const add = (id, where) => ids.set(id, [...(ids.get(id) || []), where]);
  for (const s of sections) add(`${s.app}#${s.name}`, 'section');
  for (const m of src.matchAll(/\bpart\('(\w+)', (\(\) => [^\n]+)/g)) {
    const slug = partSlug(m[2]);
    if (slug) add(`${m[1]}#${slug}`, 'part');
  }
  // A part written twice on purpose, under different conditions, is one id: only one of them runs in a run.
  const twice = [...ids].filter(([, w]) => w.length > 1 && w.includes('section'));
  assert.deepEqual(twice, []);
});
