// The map of what each app takes from another (tools/app-map.mjs,
// docs/APP-PACKS.md). Before the apps become packs, every import of one
// app's file by another is cut; until then the ones there are today are
// written down (tests/fixtures/app-map-edges.json), and a new one, or an
// engine one app owned now reached by another, is noticed here rather than
// found when the move meets it.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { mapImports, edgesOf, EDGES_FIXTURE, ROOT, PACKS } from '../tools/app-map.mjs';

const allowed = JSON.parse(fs.readFileSync(path.join(ROOT, EDGES_FIXTURE), 'utf8'));
const map = await mapImports();
const now = edgesOf(map);
const HOW = 'Cut it (ask through IPC or an event, or move the shared part into the core), or, when it is meant, run `node tools/app-map.mjs --write-edges` and commit the fixture with it.';

test('every app the windows load is a pack, and every pack is found', () => {
  assert.deepEqual(map.unknownApps, [], 'renderer/index.js loads an app tools/app-map.mjs has no pack for: add it to PACKS');
  for (const p of PACKS) {
    const found = (map.starts.renderer[p.id] || []).length + (map.starts.main[p.id] || []).length;
    assert.ok(found > 0, `${p.name}: no window module or service found for it`);
  }
  // The main process's services the map reads from main.js: the apps' own among them.
  for (const ns of ['mail', 'calendar', 'contacts', 'pdf', 'ocr', 'base', 'talk', 'live', 'ai', 'present']) {
    assert.ok(map.namespaces[ns], `main.js no longer builds the ${ns} namespace where the map looks for it`);
  }
});

test('no new import of one app by another', () => {
  const extra = now.cross.filter((e) => !allowed.cross.includes(e)).map((key) => {
    const e = map.cross.find((c) => `${c.from} -> ${c.to}` === key);
    return `${e.from}:${e.line} imports ${e.to} (${e.fromOwner} to ${e.toOwner})`;
  });
  assert.deepEqual(extra, [], `A new cross-app import. ${HOW}`);
});

test('every engine keeps its owner', () => {
  const moved = Object.entries(now.owners)
    .filter(([h, o]) => allowed.owners[h] && allowed.owners[h] !== o)
    .map(([h, o]) => `${h.replace(/^pkg:/, '@rutba/').replace(/^npm:/, '')} was ${allowed.owners[h]}'s and is now ${o}'s (reached by ${map.packages[h].reachedBy.join(', ')})`);
  assert.deepEqual(moved, [], `An engine changed hands. ${HOW}`);
});

test('the cuts and the wiring the map lists are real imports with their lines', () => {
  assert.ok(map.registration.length > 0);
  for (const e of [...map.cross, ...map.registration]) {
    const text = fs.readFileSync(path.join(ROOT, e.from), 'utf8').split(/\r?\n/);
    assert.ok(e.line > 0 && text[e.line - 1].includes(e.spec), `${e.from}:${e.line} does not name ${e.spec}`);
  }
});
