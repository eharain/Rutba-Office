// Writing a file so that it is either the old one or the new one, never half.

import fsp from 'node:fs/promises';
import path from 'node:path';

/**
 * A file written whole or not at all: the bytes go to a file beside it first
 * and are renamed over it, so a crash, a full disk or a lost drive part way
 * through leaves the file as it was rather than half of the new one — as
 * the document service saves (documents.js, `writeWhole`). Where Windows will
 * not let the rename replace a file another program holds open, the bytes
 * already safe on disk are copied over it instead.
 */
export async function writeWhole(target, data, encoding) {
  let real = target;
  try { real = await fsp.realpath(target); } catch { /* a new file: there is no link to follow */ }
  const old = await fsp.stat(real).catch(() => null);
  const temp = path.join(path.dirname(real), `.${path.basename(real)}.${process.pid}.saving`);
  try {
    await fsp.writeFile(temp, data, encoding);
    if (old) await fsp.chmod(temp, old.mode).catch(() => {});
    try {
      await fsp.rename(temp, real);
    } catch (err) {
      if (!old || !['EPERM', 'EACCES', 'EBUSY'].includes(err.code)) throw err;
      await fsp.copyFile(temp, real);
    }
  } finally {
    await fsp.rm(temp, { force: true }).catch(() => {});
  }
}
