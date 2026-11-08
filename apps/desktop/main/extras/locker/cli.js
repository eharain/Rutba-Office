#!/usr/bin/env node
// The Locker's command line. Its name (rutba-locker) is a working name.
//
//   rutba-locker add <folder> [--keep <dir> | --delete] [--visible]
//   rutba-locker status [<folder>]        what is locked, waiting or excluded
//   rutba-locker why <folder> <file>      the rule that decided
//   rutba-locker lock <folder>            lock everything waiting
//   rutba-locker unlock <folder> <file>   back to a plain file, for good
//   rutba-locker export <folder> <file> <dest>
//   rutba-locker import <folder> <source> [<as>]
//   rutba-locker log <folder> <file>      the versions kept
//   rutba-locker restore <folder> <file> <version>
//   rutba-locker change-password
//   rutba-locker recovery                 replace the recovery key
//
// The password is read at the prompt, or from standard input when piped;
// never from an argument, so it never appears in a process list or a shell's
// history. It opens nothing that the app has not; it is run by Office's own
// program started as Node (ELECTRON_RUN_AS_NODE) so nothing more is installed.

import os from 'node:os';
import path from 'node:path';
import readline from 'node:readline';
import { createLocker } from './locker.js';
import { LockerError } from './format.js';

function parseArgs(argv) {
  const args = [];
  const flags = {};
  for (let i = 0; i < argv.length; i += 1) {
    const a = argv[i];
    if (a === '--visible' || a === '--delete' || a === '--json') flags[a.slice(2)] = true;
    else if (a === '--keep' || a === '--dir' || a === '--versions') flags[a.slice(2)] = argv[++i];
    else args.push(a);
  }
  return { args, flags };
}

/** The locker's folder: named, or the default profile's. */
function lockerDir(flags) {
  if (flags.dir) return path.resolve(flags.dir);
  const base = process.env.APPDATA || path.join(os.homedir(), '.config');
  return path.join(base, 'Rutba Office', 'locker');
}

/** A line of a password, from standard input or the prompt, never shown. */
let piped = null;

function readSecret(prompt) {
  if (!process.stdin.isTTY) {
    // Standard input is read once; each secret asked for takes the next line.
    piped ||= new Promise((resolve) => {
      let data = '';
      process.stdin.setEncoding('utf8');
      process.stdin.on('data', (c) => { data += c; });
      process.stdin.on('end', () => resolve(data.split(/\r?\n/)));
    });
    return piped.then((lines) => lines.shift() ?? '');
  }
  return new Promise((resolve) => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout, terminal: true });
    rl.question(prompt, (answer) => {
      rl.close();
      process.stdout.write('\n');
      resolve(answer);
    });
    rl._writeToOutput = (s) => { if (s.includes(prompt)) rl.output.write(s); };
  });
}

function say(flags, value) {
  if (flags.json) process.stdout.write(`${JSON.stringify(value)}\n`);
  else process.stdout.write(`${typeof value === 'string' ? value : JSON.stringify(value, null, 2)}\n`);
}

/** The folder a person named: its folder record. */
async function need(locker, folderPath) {
  const list = await locker.folders();
  const target = path.resolve(folderPath);
  const f = list.find((x) => path.resolve(x.root).toLowerCase() === target.toLowerCase());
  if (!f) throw new LockerError('state', 'That folder is not in the locker. Add it with "add".');
  return f;
}

async function main() {
  const { args, flags } = parseArgs(process.argv.slice(2));
  const [verb, ...rest] = args;
  if (!verb || verb === 'help') {
    say(flags, 'rutba-locker add|status|why|lock|unlock|export|import|log|restore|change-password|recovery');
    return 0;
  }
  const dir = lockerDir(flags);
  const locker = createLocker({ dir });

  if (verb === 'setup') {
    if (await locker.exists()) throw new LockerError('exists', 'This profile already has a locker.');
    const password = await readSecret('New locker password: ');
    const { recovery } = await locker.setup(password);
    say(flags, flags.json ? { recovery } : `Locker made. Keep this recovery key: ${recovery}`);
    return 0;
  }

  const password = await readSecret('Locker password: ');
  await locker.unlock({ password });
  try {
    switch (verb) {
      case 'add': {
        const choices = {};
        if (flags.visible) choices.hidden = false;
        if (flags.delete) choices.afterLock = 'delete';
        if (flags.keep) {
          choices.afterLock = 'keep';
          choices.keepDir = path.resolve(flags.keep);
        }
        if (flags.versions) choices.keepVersions = Number(flags.versions);
        const out = await locker.addFolder(rest[0], choices, { password });
        say(flags, out);
        break;
      }
      case 'status': {
        const folders = rest[0] ? [await need(locker, rest[0])] : await locker.folders();
        const rows = [];
        for (const f of folders) rows.push({ folder: f.root, ...(await locker.summary(f.id)), entries: await locker.scan(f.id) });
        say(flags, flags.json ? rows : rows.map((r) => `${r.folder}: ${r.locked} locked, ${r.waiting} waiting, ${r.excluded} not locked${r.conflict ? `, ${r.conflict} with two versions` : ''}`).join('\n'));
        break;
      }
      case 'why': {
        const f = await need(locker, rest[0]);
        const out = await locker.why(f.id, rest[1]);
        say(flags, flags.json ? out : `${out.state}: ${out.why ?? 'locked'}`);
        break;
      }
      case 'lock': {
        const f = await need(locker, rest[0]);
        const out = await locker.lockWaiting(f.id);
        say(flags, flags.json ? out : `${out.locked.length} locked${out.failed.length ? `, ${out.failed.length} could not be locked` : ''}`);
        if (out.failed.length) {
          for (const x of out.failed) say(flags, `  ${x.rel}: ${x.message}`);
          return 2;
        }
        break;
      }
      case 'unlock': {
        const f = await need(locker, rest[0]);
        say(flags, await locker.unlockForGood(f.id, rest[1]));
        break;
      }
      case 'export': {
        const f = await need(locker, rest[0]);
        say(flags, await locker.exportCopy(f.id, rest[1], path.resolve(rest[2])));
        break;
      }
      case 'import': {
        const f = await need(locker, rest[0]);
        say(flags, await locker.importFile(f.id, path.resolve(rest[1]), rest[2]));
        break;
      }
      case 'log': {
        const f = await need(locker, rest[0]);
        say(flags, await locker.versions(f.id, rest[1]));
        break;
      }
      case 'restore': {
        const f = await need(locker, rest[0]);
        say(flags, await locker.restore(f.id, rest[1], Number(rest[2])));
        break;
      }
      case 'change-password': {
        const next = await readSecret('New password: ');
        say(flags, await locker.changePassword(password, next));
        break;
      }
      case 'recovery': {
        const out = await locker.replaceRecovery(password);
        say(flags, flags.json ? out : `New recovery key: ${out.recovery}`);
        break;
      }
      default:
        say(flags, `Unknown command: ${verb}`);
        return 1;
    }
  } finally {
    await locker.lock();
  }
  return 0;
}

main().then((code) => { process.exitCode = code; }, (err) => {
  process.stderr.write(`${err?.message || err}\n`);
  process.exitCode = err?.code === 'bad-password' ? 3 : 1;
});
