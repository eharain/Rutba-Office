// Everything, in the order that finds problems soonest.
//
//   npm run gate
//
// Four passes, each answering a question the one before it cannot:
//
//   test          do the engines do what they say? 619 checks, no windows.
//   verify:edit   do keystrokes reach the document? two real windows.
//   verify:apps   does each app open, change and save a real file? every window.
//   smoke         does every window paint without logging anything?
//
// Unit tests first because they take seconds and catch most of it. The window
// runs are slower, need a display, and are the only ones that can prove the
// thing a person actually came for.
//
// Nothing here is optional before a release. A suite that cannot save is not a
// suite, and the only way to know it still saves is to have watched it.

import { spawn } from 'node:child_process';

const PASSES = [
  ['test', 'the engines', 'npm', ['test']],
  ['verify:edit', 'typing reaches the document', 'npm', ['run', 'verify:edit']],
  ['verify:apps', 'each app opens, changes and saves', 'npm', ['run', 'verify:apps']],
  ['smoke', 'every window paints', 'npm', ['run', 'smoke']],
];

const only = process.argv.slice(2).filter((a) => !a.startsWith('-'));
const quiet = process.argv.includes('--quiet');

function run(command, args) {
  return new Promise((resolve) => {
    const child = spawn(command, args, {
      stdio: quiet ? ['ignore', 'pipe', 'pipe'] : 'inherit',
      shell: process.platform === 'win32',
    });
    let tail = '';
    if (quiet) {
      const keep = (chunk) => {
        tail = (tail + chunk).slice(-4000);
      };
      child.stdout.on('data', keep);
      child.stderr.on('data', keep);
    }
    child.on('exit', (code) => resolve({ code: code ?? 1, tail }));
  });
}

const results = [];
let failed = false;

for (const [name, what, command, args] of PASSES) {
  if (only.length && !only.includes(name)) continue;

  const started = Date.now();
  if (!quiet) console.log(`\n[1m── ${name} — ${what} ─────────────────────────────[0m\n`);
  const { code, tail } = await run(command, args);
  const seconds = ((Date.now() - started) / 1000).toFixed(1);
  results.push({ name, what, ok: code === 0, seconds });

  if (code !== 0) {
    failed = true;
    if (quiet && tail) console.log(tail);
    // The later passes are slower and mostly re-report the same fault, so a
    // failure stops here rather than spending four minutes confirming it.
    break;
  }
}

console.log('\n────────────────────────────────────────────────');
for (const r of results) {
  console.log(`${r.ok ? '[32mok  [0m' : '[31mFAIL[0m'} ${r.name.padEnd(13)} ${r.what.padEnd(38)} ${r.seconds}s`);
}
const total = results.reduce((n, r) => n + Number(r.seconds), 0).toFixed(1);
console.log(
  failed
    ? `\n[31mThe gate is closed.[0m Fix the pass above and run it again.`
    : `\n[32mThe gate is open.[0m ${results.length} passes in ${total}s.`
);

process.exit(failed ? 1 : 0);
