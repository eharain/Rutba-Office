import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { unwrapNotes } from '../tools/release-notes.js';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

test('a wrapped paragraph and a wrapped list item each become one line', () => {
  const md = [
    'The first line of a paragraph',
    'and the rest of it.',
    '',
    '## Documents',
    '',
    '- **A Table Design tab**: Table Styles puts a table in Table Grid,',
    '  Plain Table 1 or Grid Table 4.',
    '- **Draw Table**: the pen draws a new table.',
  ].join('\n');
  assert.equal(unwrapNotes(md), [
    'The first line of a paragraph and the rest of it.',
    '',
    '## Documents',
    '',
    '- **A Table Design tab**: Table Styles puts a table in Table Grid, Plain Table 1 or Grid Table 4.',
    '- **Draw Table**: the pen draws a new table.',
  ].join('\n'));
});

test('headings, numbered items, nested items, quotes, tables, rules and code keep their lines', () => {
  const md = [
    '## Heading',
    'A line under a heading starts its own paragraph.',
    '1. first',
    '2. second',
    '- outer',
    '  - inner',
    '> a quote',
    '| a | b |',
    '| - | - |',
    '---',
    '```',
    'code that',
    '  keeps its lines',
    '```',
  ].join('\n');
  assert.equal(unwrapNotes(md), md);
});

test('line endings from Windows are read the same way', () => {
  assert.equal(unwrapNotes('one\r\ntwo\r\n\r\n- a\r\n  b'), 'one two\n\n- a b');
});

test('every release note unwraps with no wrapped continuation left over', () => {
  const dir = path.join(root, 'docs', 'releases');
  for (const name of fs.readdirSync(dir).filter((n) => /^v\d+\.\d+\.\d+\.md$/.test(n))) {
    const body = unwrapNotes(fs.readFileSync(path.join(dir, name), 'utf8'));
    let inFence = false;
    for (const line of body.split('\n')) {
      if (/^\s*```/.test(line)) inFence = !inFence;
      if (inFence) continue;
      // A continuation is an indented line of plain words, not a nested item.
      assert.ok(!/^\s{2,}[^\s\-*+\d>|]/.test(line), `${name}: a wrapped line was left: "${line.slice(0, 60)}"`);
    }
  }
});
