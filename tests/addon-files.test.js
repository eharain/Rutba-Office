// A window loads from an add-on only a font pack's fonts, at
// rutba://app/addons/<id>/<file> (office-shell's addon-files.js), never
// through the app packs' roots, whose chunks and files are window code.

import test from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import { addonFile, addonFileUrl, isAddonFileUrl } from '../packages/office-shell/src/electron/addon-files.js';

const roots = { 'font-aran': path.resolve('C:/addons/font-aran') };

test('a font add-on serves its fonts at its own address, and nothing else it holds', () => {
  assert.equal(addonFileUrl('font-aran', 'fonts/Noto Nastaliq.ttf'), 'rutba://app/addons/font-aran/fonts/Noto%20Nastaliq.ttf');
  assert.ok(isAddonFileUrl(addonFileUrl('font-aran', 'a.woff2')));
  assert.ok(!isAddonFileUrl('rutba://app/packs/font-aran/a.ttf'));
  assert.equal(addonFile('addons/font-aran/fonts/Nastaliq.ttf', roots), path.join(roots['font-aran'], 'fonts', 'Nastaliq.ttf'));
  assert.equal(addonFile('addons/font-aran/X.WOFF2', roots), path.join(roots['font-aran'], 'X.WOFF2'));
  for (const rel of [
    'addons/font-aran/renderer/chunks/talk.js', // window code
    'addons/font-aran/index.html',
    'addons/font-aran/font-aran.pack.json',
    'addons/lang-ur/a.ttf', // not a font add-on served
    'addons/font-aran/../lang-ur/a.ttf',
    'addons/font-aran/fonts//a.ttf',
    'addons/font-aran/C:a.ttf',
    'addons/font-aran',
    'packs/font-aran/a.ttf',
    'addons/__proto__/a.ttf',
  ]) assert.equal(addonFile(rel, roots), null, rel);
});
