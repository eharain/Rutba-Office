// An add-on whose code runs (the speech engine) imports only what its own
// folder carries (office-shell's pack-resolve.js, closeAddonsFolder): Node's
// built-ins and its own files and packages, never a node_modules above the
// add-ons folder, the core's packages or another add-on's files.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-addon-resolve-'));
test.after(() => fs.rmSync(tmp, { recursive: true, force: true }));

test('an add-on engine imports only what its own folder carries: no node_modules above, no core, no other add-on', async () => {
  const { closeAddonsFolder } = await import('../packages/office-shell/src/electron/pack-resolve.js');
  const above = path.join(tmp, 'addons-above', 'node_modules', 'only-above');
  fs.mkdirSync(above, { recursive: true });
  fs.writeFileSync(path.join(above, 'package.json'), JSON.stringify({ name: 'only-above', type: 'module', exports: './c.js' }));
  fs.writeFileSync(path.join(above, 'c.js'), "export const where = 'above';\n");
  const dir = path.join(tmp, 'addons-above', 'addons');
  const engine = path.join(dir, 'engine-xx');
  const own = path.join(engine, 'node_modules', 'its-own');
  fs.mkdirSync(own, { recursive: true });
  fs.writeFileSync(path.join(own, 'package.json'), JSON.stringify({ name: 'its-own', type: 'module', exports: './c.js' }));
  fs.writeFileSync(path.join(own, 'c.js'), "export const where = 'engine';\n");
  fs.mkdirSync(path.join(dir, 'other-xx'), { recursive: true });
  fs.writeFileSync(path.join(dir, 'other-xx', 'x.mjs'), "export const where = 'other';\n");
  const main = path.join(engine, 'main');
  fs.mkdirSync(main, { recursive: true });
  fs.writeFileSync(path.join(main, 'own.mjs'), "export { where } from 'its-own';\nexport { readFileSync } from 'fs';\nexport { x } from './beside.mjs';\n");
  fs.writeFileSync(path.join(main, 'beside.mjs'), 'export const x = 1;\n');
  fs.writeFileSync(path.join(main, 'above.mjs'), "export { where } from 'only-above';\n");
  fs.writeFileSync(path.join(main, 'core.mjs'), "export { PACKS } from '@rutba/office-shell/contract';\n");
  fs.writeFileSync(path.join(main, 'other.mjs'), "export { where } from '../../other-xx/x.mjs';\n");
  fs.writeFileSync(path.join(main, 'requires.mjs'), "import { createRequire } from 'node:module';\nexport const r = createRequire(import.meta.url)('only-above');\n");
  closeAddonsFolder(dir);
  const ok = await import(pathToFileURL(path.join(main, 'own.mjs')).href);
  assert.equal(ok.where, 'engine');
  assert.equal(typeof ok.readFileSync, 'function', 'the runtime\'s own modules are untouched');
  for (const f of ['above.mjs', 'other.mjs', 'requires.mjs']) {
    await assert.rejects(import(pathToFileURL(path.join(main, f)).href), /only what it carries/, f);
  }
  // Nor is the core asked, as it is for an app pack.
  await assert.rejects(import(pathToFileURL(path.join(main, 'core.mjs')).href), /@rutba\/office-shell/);
  assert.match(process.env.RUTBA_PACK_RESOLUTION, /addons-above/, 'its worker threads are told too');
});
