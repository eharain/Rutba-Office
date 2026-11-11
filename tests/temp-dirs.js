// A test's own folder in the system's temp folder, removed when the test file's process ends.
// `npm test` gives every run a temp folder of its own and removes it (tools/run-temp.js), but a
// test file run by hand used to leave its folders behind: the Locker tests made one a case, and
// eleven thousand of them filled the disk on 11 Nov 2026.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

const made = [];
let hooked = false;

/** A new empty folder named `prefix` and six random letters, gone when this process exits. */
export function tempDir(prefix = 'rutba-test-') {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), prefix));
  made.push(dir);
  if (!hooked) {
    hooked = true;
    process.on('exit', removeTempDirs);
  }
  return dir;
}

/**
 * Remove every folder tempDir made, read-only files included, and any a test moved it to beside
 * it under a longer name (a locked folder moved to `<folder>-moved`); one still in use is left.
 */
export function removeTempDirs() {
  for (const dir of made.splice(0)) {
    const parent = path.dirname(dir);
    const name = path.basename(dir);
    let beside = [];
    try { beside = fs.readdirSync(parent).filter((n) => n.startsWith(`${name}-`)); } catch {}
    for (const target of [dir, ...beside.map((n) => path.join(parent, n))]) {
      try {
        fs.rmSync(target, { recursive: true, force: true, maxRetries: 2 });
      } catch {
        // A file another process holds open stays, and run-temp clears it with the run.
      }
    }
  }
}
