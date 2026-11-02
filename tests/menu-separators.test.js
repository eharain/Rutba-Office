// A menu's line between groups is the bare string '-'. An item whose label is
// '-' is drawn as a row reading "-" (Worksheets' Format menu did, in 2.21).
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';

const roots = ['apps/desktop/renderer', 'packages/office-ui/src', 'packages/studio'];

function sources(dir, out = []) {
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    if (e.name === 'node_modules' || e.name === 'out') continue;
    const p = path.join(dir, e.name);
    if (e.isDirectory()) sources(p, out);
    else if (/\.(js|jsx|mjs)$/.test(e.name)) out.push(p);
  }
  return out;
}

test('no menu item has "-" for its label', () => {
  const found = [];
  for (const root of roots) {
    for (const file of sources(root)) {
      fs.readFileSync(file, 'utf8').split('\n').forEach((line, i) => {
        if (/\{\s*label:\s*['"]-['"]\s*\}/.test(line)) found.push(`${file}:${i + 1}`);
      });
    }
  }
  assert.deepEqual(found, [], `use the bare string '-' for a menu's line: ${found.join(', ')}`);
});
