// Each app's commands go through one switch on the command's name, and a
// name written twice there is not an error: the first case takes every call
// and the second is never reached. Worksheets' View → Arrange All once took
// the name the shape Arrange commands used, and Align, Bring to Front, Rotate
// and Group all arranged windows instead. No name appears twice in an app.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const apps = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'apps', 'desktop', 'renderer', 'apps');

test('no app names a command case twice', () => {
  const twice = [];
  for (const name of fs.readdirSync(apps).filter((f) => f.endsWith('.js'))) {
    const seen = new Map();
    for (const m of fs.readFileSync(path.join(apps, name), 'utf8').matchAll(/^\s*case '([A-Za-z]+)':/gm)) {
      seen.set(m[1], (seen.get(m[1]) || 0) + 1);
    }
    for (const [label, n] of seen) if (n > 1) twice.push(`${name}: '${label}' ×${n}`);
  }
  assert.deepEqual(twice, []);
});
