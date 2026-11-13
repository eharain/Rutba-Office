// The removal tests on both runtimes that matter (D-2026-11-13-37): the Node
// the unit tests run on, and Office's own Electron as Node
// (ELECTRON_RUN_AS_NODE), which is the Node 24 Office ships. The two differ on
// the one thing these tests are about: Electron 44's Node follows a Windows
// junction in a recursive fs.rm, the Node the tests run on does not. A suite
// that passes only on the first proves nothing about the second.
//
//   node tools/run-temp.js node tools/test-removal-runtimes.mjs [test files...]
//   npm run test:removal
//
// With no files, SUITES run: the helper's, the check's, and those of every
// place that removes a tree with it. A suite listed but not in the checkout
// is a fault, not a silent drop (tests/remove-tree.test.js comes with
// packs-core, and this fails until it is in). Electron is the one this
// checkout installed (node_modules/electron), or the binary named by
// RUTBA_ELECTRON. A checkout without Electron is a fault here, not a skip: a
// pass on one runtime is half the answer.

import { spawnSync } from 'node:child_process';
import { createRequire } from 'node:module';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

/** The suites run on both: the helper's own, and the suites of the places that remove trees with it. */
export const SUITES = [
  // The helper, and the check that nothing else removes a tree.
  'tests/remove-tree.test.js', // packs-core's, with the add-ons
  'tests/remove-tree-runtimes.test.js',
  'tests/recursive-removals.test.js',
  // office-shell: profiles, temp folders, thumbnails, the packs.
  'tests/profiles.test.js',
  'tests/profile-folder.test.js',
  'tests/profile-keystore.test.js',
  'tests/temp-sweep.test.js',
  'tests/thumbs-queue.test.js',
  'tests/locker-thumbnails-cleared.test.js',
  'tests/app-packs.test.js',
  'tests/packs-service.test.js',
  // The desktop's own sites.
  'tests/template-sets.test.js',
  'tests/print.test.js',
  'tests/mail-eml.test.js',
  'tests/locker-memory.test.js', // memory-files' sweep
  // The mailbox store: messages removed, folders moved and removed.
  'tests/mailbox.test.js',
  'tests/mail-never-lose.test.js',
  'tests/mail-many.test.js',
  'tests/mail-server-sync.test.js',
  'tests/mail-third-review.test.js',
  'tests/mail-crash-at-each-write.test.js',
];

/** The listed suites this checkout does not have. */
export function missingSuites(root = ROOT, suites = SUITES) {
  return suites.filter((f) => !fs.existsSync(path.join(root, f)));
}

/** Office's Electron: RUTBA_ELECTRON, else the electron package's own binary path. */
export function electronBinary(env = process.env) {
  if (env.RUTBA_ELECTRON) return env.RUTBA_ELECTRON;
  try {
    const found = createRequire(path.join(ROOT, 'package.json'))('electron');
    return typeof found === 'string' ? found : null;
  } catch {
    return null;
  }
}

function run(label, binary, files, extraEnv) {
  process.stdout.write(`\n== ${label}: ${binary}\n`);
  const r = spawnSync(binary, ['--test', '--test-reporter=tap', ...files], {
    cwd: ROOT,
    stdio: 'inherit',
    env: { ...process.env, ...extraEnv },
  });
  if (r.error) {
    process.stdout.write(`${label}: could not start (${r.error.message})\n`);
    return false;
  }
  return r.status === 0;
}

const isMain = process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (isMain) {
  const named = process.argv.slice(2);
  const missing = named.length ? missingSuites(ROOT, named) : missingSuites();
  const files = (named.length ? named : SUITES).filter((f) => !missing.includes(f));
  const results = [];
  if (missing.length) {
    process.stdout.write(`\nListed but not in this checkout: ${missing.join(', ')}\n`);
    results.push(['Suites', false]);
  }
  const electron = electronBinary();
  results.push(['Node', run(`Node ${process.versions.node}`, process.execPath, files, {})]);
  if (!electron || !fs.existsSync(electron)) {
    process.stdout.write('\nElectron: not found (install the checkout, or set RUTBA_ELECTRON to its binary)\n');
    results.push(['Electron', false]);
  } else {
    results.push(['Electron', run('Electron as Node', electron, files, { ELECTRON_RUN_AS_NODE: '1' })]);
  }
  process.stdout.write(`\n${results.map(([name, ok]) => `${name}: ${ok ? 'pass' : 'FAIL'}`).join(', ')}\n`);
  process.exitCode = results.every(([, ok]) => ok) ? 0 : 1;
}
