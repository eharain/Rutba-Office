// The taskbar's tiles, kept apart from windows.js so a test can reach them without Electron.

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';

/**
 * The taskbar's copy of an app's tile, named by what is in it. Windows keeps
 * an icon by its path: after an update replaced Mail's tile while Mail was
 * open, Mail showed a blank page on the taskbar from then on, though the
 * file was whole. A new tile is a new name, so Windows reads it afresh. The
 * copies are kept (a pinned app keeps the path it was pinned with); they are
 * small and change only with the icon. The tile itself is the answer when a
 * copy cannot be made.
 */
export function taskbarTile(appKey, tile, dir) {
  try {
    const bytes = fs.readFileSync(tile);
    const name = `${appKey}-${crypto.createHash('sha256').update(bytes).digest('hex').slice(0, 12)}.ico`;
    const copy = path.join(dir, name);
    if (!fs.existsSync(copy)) {
      fs.mkdirSync(dir, { recursive: true });
      fs.writeFileSync(`${copy}.part`, bytes);
      fs.renameSync(`${copy}.part`, copy);
    }
    return copy;
  } catch {
    return tile;
  }
}
