// The window checks cut into chunks by the seconds measured before, never by a
// guess (PM, 11 Nov 2026), a long app split into pieces of its parts, and the
// pieces read back as the app's one record.
import test from 'node:test';
import assert from 'node:assert/strict';
import { planChunks, noteSeconds, mergePieces, sourceIds, MAX_PIECES } from '../tools/gate-plan.js';
import { partsToRerun } from '../tools/gate.js';
import { createSelector } from '../apps/desktop/main/verify-select.js';
import { leftToMachine } from '../tools/gate.js';

const seconds = {
  parts: { 'video#a': 200, 'video#b': 150, 'video#c': 100, 'video#d': 40, 'mail#x': 30, 'mail#y': 20 },
  apps: { video: { of: 90, inParts: 90, seconds: 500, ids: ['video#a', 'video#b', 'video#c', 'video#d'] }, mail: { of: 40, inParts: 40, seconds: 60, ids: ['mail#x', 'mail#y'] } },
};

test('a chunk never measured runs first; the rest longest first', () => {
  const plan = planChunks(['mail', 'talk', 'video'], { timing: { mail: 60, video: 120 }, seconds: { parts: {}, apps: {} }, jobs: 3 });
  assert.deepEqual(plan.map((c) => [c.label, c.seconds]), [['talk', null], ['video', 120], ['mail', 60]]);
});

test('a long app with every check in a part is cut into pieces of about a fair share each, one running the rest', () => {
  const plan = planChunks(['video', 'mail', 'talk'], { timing: { video: 500, mail: 60, talk: 100 }, seconds, jobs: 3 });
  const pieces = plan.filter((c) => c.app === 'video');
  assert.ok(pieces.length >= 2 && pieces.length <= MAX_PIECES, `${pieces.length} pieces`);
  // Each named part in one piece only; the rest piece skips exactly what the others name.
  const named = pieces.filter((p) => p.ids).flatMap((p) => p.ids);
  assert.equal(new Set(named).size, named.length);
  const rest = pieces.filter((p) => p.skip);
  assert.equal(rest.length, 1);
  assert.deepEqual([...rest[0].skip].sort(), [...named].sort());
  assert.equal(rest[0].ids, null);
  // Every measured part runs somewhere: named in a piece, or not skipped by the rest.
  for (const id of Object.keys(seconds.parts).filter((x) => x.startsWith('video#'))) assert.ok(named.includes(id) || !rest[0].skip.includes(id), id);
  assert.deepEqual(plan.filter((c) => c.app !== 'video').map((c) => c.label).sort(), ['mail', 'talk']);
  // Longest first, the pieces among them.
  for (let i = 1; i < plan.length; i++) assert.ok(plan[i - 1].seconds >= plan[i].seconds);
});

test('an app is left whole when a check of it sat outside its parts, or it is short, or its parts were never measured', () => {
  const outside = { ...seconds, apps: { ...seconds.apps, video: { of: 95, inParts: 90, seconds: 500, ids: seconds.apps.video.ids } } };
  assert.deepEqual(planChunks(['video', 'mail'], { timing: { video: 500, mail: 60 }, seconds: outside, jobs: 3 }).map((c) => c.label), ['video', 'mail']);
  assert.deepEqual(planChunks(['video', 'mail'], { timing: { video: 500, mail: 60 }, seconds: { parts: {}, apps: {} }, jobs: 3 }).map((c) => c.label), ['video', 'mail']);
  // With one lane there is no other lane to give a piece to: its share is the whole.
  assert.deepEqual(planChunks(['video', 'mail'], { timing: { video: 500, mail: 60 }, seconds, jobs: 1 }).map((c) => c.label), ['video', 'mail']);
});

test('the seconds noted are of clean parallel parts only, and an app counts as all in parts only when it ran whole', () => {
  const s = { parts: {}, apps: {} };
  noteSeconds(s, {
    word: { of: 12, seconds: 80, parts: { 'word#a': { of: 5, seconds: 30.04 }, 'word#b': { of: 4, seconds: 20, threw: true }, 'word#c': { of: 3, seconds: 9, failed: ['x'] }, 'word#m': { of: 2, seconds: 5, machine: true } } },
  }, { whole: ['word'] });
  assert.deepEqual(s.parts, { 'word#a': 30 });
  assert.deepEqual(s.apps.word, { of: 12, inParts: 12, seconds: 80, ids: ['word#a', 'word#b', 'word#c'] });
  noteSeconds(s, { sheets: { of: 3, parts: { 'sheets#a': { of: 3, seconds: 4 } } } });
  assert.equal(s.apps.sheets, undefined);
  assert.equal(s.parts['sheets#a'], 4);
});

test("the pieces read back as one app's record; a piece that crashed makes it the app whole", () => {
  const seen = (ids) => ids.map((id) => ({ id, machine: true, because: 'lane' }));
  const a = { ok: true, seconds: 100, seen: seen(['video#m']), parts: { video: { passed: 10, of: 10, failed: [], passedNames: ['n1'], parts: { 'video#a': { of: 10 } } } }, sum: { passed: 10, of: 10, failed: [] } };
  const b = { ok: false, seconds: 80, seen: seen(['video#m']), parts: { video: { passed: 4, of: 5, failed: ['bad'], realFailed: ['bad'], failedParts: ['video#b'], passedNames: ['n2'], parts: { 'video#b': { of: 5 } } } }, sum: { passed: 4, of: 5, failed: ['bad'] } };
  const m = mergePieces([a, b], 'video', leftToMachine);
  assert.deepEqual([m.ok, m.passed, m.of, m.failed, m.failedParts, m.passedNames, Object.keys(m.parts), m.toMachine, m.seconds], [false, 14, 15, ['bad'], ['video#b'], ['n1', 'n2'], ['video#a', 'video#b'], ['video#m'], 180]);
  const crashed = mergePieces([a, { ok: false, seconds: 5, seen: [], parts: null, sum: { passed: 0, of: 0, failed: ['the run ended before it counted'] } }], 'video', leftToMachine);
  assert.deepEqual(crashed.parts, {});
  assert.ok(crashed.failed.includes('the run ended before it counted'));
});

test('a piece skips what the others run, sections and parts alike, and nothing else', () => {
  const s = createSelector({ lane: 'parallel', apps: ['video'], skip: ['video#verifyA', 'video#intro'] });
  assert.equal(s.part('video', () => verifyA({})).runs, false);
  assert.equal(s.part('video', () => verifyB({})).runs, true);
  assert.equal(s.section('video', 'intro').runs, false);
  assert.equal(s.section('video', 'outro').runs, true);
  // A machine part is the machine lane's whatever a piece skips.
  assert.equal(s.part('video', () => verifyM({}), { machine: true }).because, 'lane');
  assert.equal(s.wants('video'), true);
});

test("the checks outside any part are not counted as in one, so an app with them is never split (review of 11 Nov)", () => {
  const s = { parts: {}, apps: {} };
  noteSeconds(s, { mail: { of: 10, seconds: 600, parts: { '*': { of: 2, seconds: 3 }, 'mail#a': { of: 4, seconds: 300 }, 'mail#b': { of: 4, seconds: 290 } } } }, { whole: ['mail'] });
  assert.deepEqual(s.apps.mail, { of: 10, inParts: 8, seconds: 600, ids: ['mail#a', 'mail#b'] });
  assert.equal(s.parts['*'], undefined);
  assert.deepEqual(planChunks(['mail', 'word'], { timing: { mail: 600, word: 100 }, seconds: s, jobs: 3 }).map((c) => c.label), ['mail', 'word']);
});

test('a part measured and since renamed or removed is never given a piece: forgotten by the next whole run, and not named while the checks do not name it', () => {
  const s = JSON.parse(JSON.stringify(seconds));
  // A whole run that ran a, b, c and the new e: d is gone.
  noteSeconds(s, { video: { of: 90, seconds: 500, parts: { 'video#a': { of: 30, seconds: 200 }, 'video#b': { of: 30, seconds: 150 }, 'video#c': { of: 20, seconds: 100 }, 'video#e': { of: 10, seconds: 40 } } } }, { whole: ['video'] });
  assert.equal(s.parts['video#d'], undefined);
  assert.deepEqual(s.apps.video.ids, ['video#a', 'video#b', 'video#c', 'video#e']);
  // The checks no longer name a (renamed since): no piece names it, and it runs in the piece that runs the rest.
  const current = new Set(['video#b', 'video#c', 'video#e', 'video#a2']);
  const plan = planChunks(['video', 'mail'], { timing: { video: 500, mail: 60 }, seconds: s, jobs: 3, current }).filter((c) => c.app === 'video');
  assert.ok(!plan.some((c) => (c.ids || []).includes('video#a') || (c.skip || []).includes('video#a')));
  // A file from before ids were kept: no app is split.
  const old = { parts: seconds.parts, apps: { video: { of: 90, inParts: 90, seconds: 500 } } };
  assert.deepEqual(planChunks(['video'], { timing: { video: 500 }, seconds: old, jobs: 3 }).map((c) => c.label), ['video']);
});

test('parts of no seconds never leave a piece with no part, which would run the app whole beside the others', () => {
  const s = { parts: { 'video#a': 600, 'video#b': 0, 'video#c': 0 }, apps: { video: { of: 30, inParts: 30, seconds: 700, ids: ['video#a', 'video#b', 'video#c'] } } };
  const plan = planChunks(['video', 'mail'], { timing: { video: 700, mail: 60 }, seconds: s, jobs: 4 });
  for (const c of plan.filter((x) => x.app === 'video' && x.ids)) assert.ok(c.ids.length > 0, c.label);
});

test('a crashed piece asks for the app whole, and a crashed app clears only on a run again of the app itself (review of 11 Nov)', () => {
  const ok = { ok: true, seconds: 10, seen: [], parts: { word: { passed: 5, of: 5, failed: [], parts: { 'word#a': { of: 5 } } } }, sum: { passed: 5, of: 5, failed: [] } };
  const crash = { ok: false, seconds: 5, seen: [], parts: null, sum: { passed: 0, of: 0, failed: ['the run ended before it counted'] } };
  const word = mergePieces([ok, crash], 'word', leftToMachine);
  assert.deepEqual(word.failedParts, ['*']);
  const mail = { ok: false, failed: ['x'], failedParts: ['mail#rules'] };
  const again = partsToRerun({ word, mail }, null);
  assert.deepEqual(again.parallel.ids.sort(), ['mail#rules', 'word#*']);
  // Failed with no part named: the app whole too.
  assert.deepEqual(partsToRerun({ sheets: { ok: false, failed: ['y'], failedParts: [] } }, null).parallel.ids, ['sheets#*']);
});

test("the part and section ids are read from the checks' own text", () => {
  const ids = sourceIds("await part('word', () => verifyWordColour({ open }, { dir }));\n  await part('video', async () => await verifyIntroScenes(x), { machine: true });\n  await section('mail', 'rules', async () => {\n  await part(app, () => verifyDyn());");
  assert.deepEqual([...ids].sort(), ['mail#rules', 'word#verifyWordColour'], 'a machine part is no piece\'s');
});

test('every part and section the checks name now is read by sourceIds', async () => {
  const fs = await import('node:fs');
  const text = fs.readFileSync(new URL('../apps/desktop/main/verify-apps.js', import.meta.url), 'utf8');
  const ids = sourceIds(text);
  const calls = text.split('\n').filter((l) => /\bawait (?:part|section)\(/.test(l) && !/\bmachine:\s*true\b/.test(l)).length;
  // A call named another way (its app in a variable) is not split out, only run in the rest: allowed, but few.
  assert.ok(ids.size >= calls - 5, `${ids.size} ids read of ${calls} calls`);
});

function verifyA() {}
function verifyB() {}
function verifyM() {}
