// The Locker's write path in a process that stops dead (exit 9) at a named
// step, so the test beside it can look at what a real crash leaves behind.

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';

const { sealFile } = await import(process.env.LOCKER_PLACE);
const { newHeader, openLocked } = await import(process.env.LOCKER_FORMAT);

const [src, edited, target, stageDir, keyHex, step, mode] = process.argv.slice(2);
const rootKey = Buffer.from(keyHex, 'hex');

if (mode === 'read') {
  const o = await openLocked(target, rootKey);
  const sha = crypto.createHash('sha256');
  for await (const c of o.frames()) sha.update(c);
  await o.close();
  process.stdout.write(`${sha.digest('hex')}\n`);
  process.exit(0);
}

const hooks = { at: (s) => { if (s === step) process.exit(9); } };
const headerBuf = mode === 'first' ? newHeader() : fs.readFileSync(target).subarray(0, 128);
await sealFile({ source: mode === 'first' ? src : edited, target, rootKey, headerBuf, meta: { name: path.basename(src) }, stageDir, previous: mode === 'first' ? null : target, hooks });
process.exit(0);
