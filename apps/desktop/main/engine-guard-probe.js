// Whether a process the speech engine starts is held to its folder, in this
// copy as it was built: for the upgrade test's report (upgrade-report.js),
// which the newer copy writes from its installer or its portable executable,
// so the answer comes from inside the archive.
//
// A stand-in add-on is made in a temporary add-ons folder, held as the real
// one is, and its module started through the utilityProcess an engine is
// handed (engine-host.js, heldUtilityProcess). Its first line imports its own
// package; then it tries one from a node_modules above the add-ons folder,
// which must be refused. No engine or model is needed, so it proves the start
// with the guard on whatever add-ons the profile has.

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { closeAddonsFolder } from '@rutba/office-shell/electron/pack-resolve';
import { heldUtilityProcess } from './engine-host.js';

const CHILD = [
  "import { where as first } from 'its-own';",
  "const tried = async (name) => { try { return (await import(name)).where; } catch (err) { return String(err?.message || err); } };",
  'const said = { first, above: await tried(\'only-above\') };',
  'if (process.parentPort) process.parentPort.postMessage(said);',
  'else process.send(said, () => process.exit(0));',
  '',
].join('\n');

const pkg = (dir, name, where) => {
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(path.join(dir, 'package.json'), JSON.stringify({ name, type: 'module', exports: './c.js' }));
  fs.writeFileSync(path.join(dir, 'c.js'), `export const where = '${where}';\n`);
};

/**
 * @param {object} o
 * @param {{ fork: Function }} o.utilityProcess  Electron's (Node's child_process in a test)
 * @param {string} [o.root]       where the temporary folder is made
 * @param {number} [o.timeoutMs]
 * @param {object} [o.options]    fork options (a test's stdio)
 * @returns {Promise<{ bootstrap: string|null, fromArchive: boolean, first: string|null, above: string|null, error: string|null }>}
 */
export async function probeEngineChild({ utilityProcess, root = os.tmpdir(), timeoutMs = 20000, options = {} }) {
  const tmp = fs.mkdtempSync(path.join(root, 'rutba-engine-child-'));
  const out = { bootstrap: null, fromArchive: false, first: null, above: null, error: null };
  let child = null;
  try {
    pkg(path.join(tmp, 'node_modules', 'only-above'), 'only-above', 'above');
    const dir = path.join(tmp, 'addons');
    pkg(path.join(dir, 'engine-check', 'node_modules', 'its-own'), 'its-own', 'engine');
    const file = path.join(dir, 'engine-check', 'child.mjs');
    fs.writeFileSync(file, CHILD);
    closeAddonsFolder(dir);
    const handed = heldUtilityProcess({ fork: (...a) => { out.bootstrap = a[0]; return utilityProcess.fork(...a); } });
    child = handed.fork(file, [], { serviceName: 'Rutba Office engine check', ...options });
    out.fromArchive = /[\\/]app\.asar[\\/]/i.test(out.bootstrap || '');
    const said = await new Promise((resolve) => {
      const timer = setTimeout(() => resolve({ error: `no answer in ${timeoutMs / 1000} seconds` }), timeoutMs);
      child.on('message', (m) => { clearTimeout(timer); resolve(m || {}); });
      child.on('exit', (code) => { clearTimeout(timer); resolve({ error: `the check process ended (${code}) without answering` }); });
    });
    out.first = said.first ?? null;
    out.above = said.above ?? null;
    out.error = said.error ?? null;
  } catch (err) {
    out.error = String(err?.message || err).slice(0, 400);
  } finally {
    try { child?.kill(); } catch { /* already ended */ }
    try { fs.rmSync(tmp, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 }); } catch { /* the temporary folder's to clear */ }
  }
  return out;
}

/** Whether a probe's answer shows the guard on: its own package found, the one above refused. */
export const probeHeld = (p) => Boolean(p && !p.error && p.first === 'engine' && /only what it carries/.test(String(p.above || '')));
