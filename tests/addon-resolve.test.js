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

test('a process an engine forks is held to its folder from its first import: it starts as the core bootstrap, with no execArgv', async () => {
  const { closeAddonsFolder } = await import('../packages/office-shell/src/electron/pack-resolve.js');
  const { heldUtilityProcess } = await import('../apps/desktop/main/engine-host.js');
  const { fork } = await import('node:child_process');
  const top = path.join(tmp, 'child-above');
  const above = path.join(top, 'node_modules', 'only-above');
  fs.mkdirSync(above, { recursive: true });
  fs.writeFileSync(path.join(above, 'package.json'), JSON.stringify({ name: 'only-above', type: 'module', exports: './c.js' }));
  fs.writeFileSync(path.join(above, 'c.js'), "export const where = 'above';\n");
  const dir = path.join(top, 'addons');
  const own = path.join(dir, 'engine-yy', 'node_modules', 'its-own');
  fs.mkdirSync(own, { recursive: true });
  fs.writeFileSync(path.join(own, 'package.json'), JSON.stringify({ name: 'its-own', type: 'module', exports: './c.js' }));
  fs.writeFileSync(path.join(own, 'c.js'), "export const where = 'engine';\n");
  // Its first line imports its own package, as a module is read before any of its code runs.
  const worker = path.join(dir, 'engine-yy', 'worker.mjs');
  fs.writeFileSync(worker, [
    "import { where as first } from 'its-own';",
    "const tried = async (name) => { try { return (await import(name)).where; } catch (err) { return String(err.message); } };",
    "process.send({ first, above: await tried('only-above'), told: process.env.RUTBA_ADDON_CHILD ?? null }, () => process.exit(0));",
    '',
  ].join('\n'));
  // And one whose first import climbs out: with the hooks in only after it, it would run.
  const climbs = path.join(dir, 'engine-yy', 'climbs.mjs');
  fs.writeFileSync(climbs, "import { where } from 'only-above';\nprocess.send({ where }, () => process.exit(0));\n");
  closeAddonsFolder(dir);
  // Node's fork takes the same arguments as Electron's utilityProcess.fork; neither is told to import anything first.
  const start = (file) => {
    const forked = [];
    const child = heldUtilityProcess({ fork: (...a) => { forked.push(a); return fork(...a); } }).fork(file, [], { stdio: ['ignore', 'ignore', 'pipe', 'ipc'], execArgv: ['--import', 'data:text/javascript,'] });
    let errors = '';
    child.stderr.on('data', (d) => { errors += d; });
    const ended = new Promise((resolve) => {
      let said = null;
      child.on('message', (m) => { said = m; });
      child.on('exit', (code) => resolve({ code, said, errors }));
    });
    return { forked, ended };
  };
  const ran = start(worker);
  assert.equal(path.basename(ran.forked[0][0]), 'pack-resolve-child.js', 'the core bootstrap is started, not the engine module');
  assert.deepEqual(ran.forked[0][2].execArgv, []);
  const r = await ran.ended;
  assert.equal(r.code, 0, r.errors);
  assert.equal(r.said.first, 'engine');
  assert.match(r.said.above, /only what it carries/);
  assert.equal(r.said.told, null, 'the module it was started for is not left in its environment');
  const out = await start(climbs).ended;
  assert.notEqual(out.code, 0);
  assert.equal(out.said, null, 'the module whose first import climbs out never ran');
  assert.match(out.errors, /only what it carries/);
  // Not in a held add-on: not started at all.
  const loose = path.join(tmp, 'loose-child.mjs');
  fs.writeFileSync(loose, 'process.exit(0);\n');
  assert.throws(() => heldUtilityProcess({ fork }).fork(loose, []), /not inside an add-on whose imports are held/);
});

test('the upgrade report probe starts a held child and finds its own package and not the one above', async () => {
  const { probeEngineChild, probeHeld } = await import('../apps/desktop/main/engine-guard-probe.js');
  const { fork } = await import('node:child_process');
  const root = fs.mkdtempSync(path.join(tmp, 'probe-'));
  const p = await probeEngineChild({ utilityProcess: { fork }, root, options: { stdio: ['ignore', 'ignore', 'pipe', 'ipc'] } });
  assert.equal(p.error, null);
  assert.equal(path.basename(p.bootstrap), 'pack-resolve-child.js');
  assert.equal(p.fromArchive, false, 'run from the source, not an archive');
  assert.equal(p.first, 'engine');
  assert.match(p.above, /only what it carries/);
  assert.equal(probeHeld(p), true);
  assert.deepEqual(fs.readdirSync(root), [], 'its temporary folder is cleared');
  assert.equal(probeHeld({ ...p, above: 'above' }), false);
  // A fork that cannot start is said, not thrown.
  const broken = await probeEngineChild({ utilityProcess: { fork: () => { throw new Error('no utility process'); } }, root });
  assert.match(broken.error, /no utility process/);
});
