// Every file association the installer registers wears its owning app's
// file-type icon, and that icon must exist: 1.44.0's installers never built
// because .rimg, Image's own project file, named an image.ico nobody drew.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { fileAssociations } from '@rutba/office-formats/registry';

const dir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', 'apps', 'desktop', 'resources', 'filetypes');

test('every associated file type has its icon on disk, as electron-builder.config.cjs names it', () => {
  const missing = [];
  for (const a of fileAssociations()) {
    const icon = `${a.ext === 'pdf' ? 'pdf' : a.app}.ico`;
    if (!fs.existsSync(path.join(dir, icon))) missing.push(`.${a.ext} (${a.app}) -> ${icon}`);
  }
  assert.deepEqual(missing, []);
});
