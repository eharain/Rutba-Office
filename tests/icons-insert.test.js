// Insert → Icons draws the suite's own icons to pictures: each as a
// standalone SVG in the colour asked, and every icon the gallery offers is
// one the set has.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { iconSvg, iconNames } from '@rutba/office-ui/icons';

test('an icon is a standalone SVG on the 24 grid, stroked in the colour asked, with its own extras kept', () => {
  const svg = iconSvg('pictures', { colour: '#C62828', size: 512 });
  assert.match(svg, /^<svg xmlns="http:\/\/www\.w3\.org\/2000\/svg" viewBox="0 0 24 24" width="512" height="512" fill="none" stroke="#C62828"/);
  assert.match(svg, /<path d="M15 8\.5h\.01" stroke-width="2\.4"\/>/, 'a path\'s own stroke width, written as SVG writes it');
  assert.equal(iconSvg('nothing-of-the-kind'), null);
  assert.match(iconSvg('mail', { colour: 'red" onload="x' }), /stroke="#000000"/, 'a colour that is not one is not written');
});

test('every icon the gallery offers is in the set', () => {
  const source = fs.readFileSync(path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'apps', 'desktop', 'renderer', 'icons-insert.js'), 'utf8');
  const groups = /export const ICON_GROUPS = \[([\s\S]*?)\n\];/.exec(source)[1];
  const offered = [...groups.matchAll(/\[\s*'[^']+',\s*\[([^\]]*)\]\s*\]/g)].flatMap((m) => [...m[1].matchAll(/'([^']+)'/g)].map((n) => n[1]));
  assert.ok(offered.length > 40, `${offered.length} icons offered`);
  assert.deepEqual(offered.filter((n) => !iconNames.includes(n)), []);
});
