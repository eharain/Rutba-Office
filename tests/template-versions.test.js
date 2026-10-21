// Which Office a template needs (Rutba-Office-Templates, src/versions.js and
// manifest.json): versions compared as numbers, an entry with no minOffice
// usable by any Office, a development build counted as the release it leads
// to, and the set's manifest saying what each part needs.
import fs from 'node:fs';
import test from 'node:test';
import assert from 'node:assert/strict';
import { compareVersions, usableBy, needsFor } from '@rutba/templates/versions';
import { readManifest } from '@rutba/templates/manifest';
import { SAMPLE_IDS, TEMPLATE_IDS, sampleNeeds, templateNeeds } from '@rutba/templates/data-files';

test('versions compare as numbers, a development build as the release it leads to', () => {
  assert.equal(compareVersions('2.13.0', '2.9.4'), 1);
  assert.equal(compareVersions('2.12.0', '2.13.0'), -1);
  assert.equal(compareVersions('2.13.0-dev', '2.13.0'), 0);
  assert.equal(compareVersions('3', '2.99.99'), 1);
  assert.equal(usableBy('2.12.0', null), true, 'no minOffice: any Office');
  assert.equal(usableBy('2.12.0', '2.13.0'), false);
  assert.equal(needsFor('2.12.0', { minOffice: '2.13.0' }), '2.13.0');
  assert.equal(needsFor('2.13.0', { minOffice: '2.13.0' }), null);
});

test("the set's manifest says its version and what each part needs, and this Office can use all of it", () => {
  const m = readManifest();
  assert.match(m.set, /^\d+\.\d+\.\d+$/);
  for (const part of Object.values(m.parts)) assert.match(part.minOffice, /^\d+\.\d+\.\d+$/);
  const office = JSON.parse(fs.readFileSync(new URL('../package.json', import.meta.url), 'utf8')).version;
  assert.equal(usableBy(office, m.minOffice), true);
  for (const part of Object.values(m.parts)) assert.equal(usableBy(office, part.minOffice), true, `this Office (${office}) can use the part needing ${part.minOffice}`);
  for (const id of SAMPLE_IDS) assert.equal(usableBy(office, sampleNeeds(id)), true, id);
  for (const id of TEMPLATE_IDS) assert.equal(usableBy(office, templateNeeds(id)), true, id);
});
