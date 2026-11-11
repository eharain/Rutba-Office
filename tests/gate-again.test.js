// The gate's lanes and its run again at the end (owner, 2026-10-21: "Queue
// the failed ones to run at the end"): lanes started only while there is
// memory for them and added as it comes free; the parts to run again picked
// from what failed, a failure outside any part running its app whole; a run
// again that clears a failed part only when that part runs again clean, and
// says when it found none of them; and the record naming each check that ran
// again (review of 11 Nov, after gates 7 to 9).
import test from 'node:test';
import assert from 'node:assert/strict';
import { laneCap, poolByMemory, partsToRerun, rerunCleared, againSaid, partsById, machinePartsUnrun, leftToMachine, settle, onBattery, gateSection } from '../tools/gate.js';
import { createSelector, partSlug } from '../apps/desktop/main/verify-select.js';

test('the lanes the processors allow: one to each two and a half, one to four', () => {
  assert.equal(laneCap(1), 1);
  assert.equal(laneCap(8), 3);
  assert.equal(laneCap(32), 4);
});

test('a lane starts beside others only while memory is free for it, and more are added as it comes free', async () => {
  let free = 2.7;
  let clock = 0;
  let running = 0;
  let most = 0;
  const shorts = [];
  const gates = [];
  const tasks = Array.from({ length: 4 }, () => () => new Promise((resolve) => {
    running++;
    most = Math.max(most, running);
    gates.push(() => { running--; resolve(); });
  }));
  const done = poolByMemory(tasks, { cap: 3, freeGb: () => free, now: () => clock, tickMs: 5, onShort: (f, r) => shorts.push([f, r]) });
  await new Promise((r) => setTimeout(r, 20));
  assert.equal(running, 1, 'with 2.7 GB free, one lane: the second would leave less than 1.5 GB');
  assert.deepEqual(shorts[0], [2.7, 1]);
  free = 6; clock = 60_000; // memory came free, and the first lane has settled
  await new Promise((r) => setTimeout(r, 30));
  assert.equal(running, 3, 'the cap reached once there is room');
  while (gates.length) { gates.shift()(); await new Promise((r) => setTimeout(r, 10)); }
  await done;
  assert.equal(most, 3);
});

test('a lane just started counts as taking its share, so one reading of free memory does not start them all', async () => {
  let running = 0;
  const release = [];
  const tasks = Array.from({ length: 4 }, () => () => new Promise((resolve) => { running++; release.push(resolve); }));
  const done = poolByMemory(tasks, { cap: 4, freeGb: () => 3.5, now: () => 0, tickMs: 5 });
  await new Promise((r) => setTimeout(r, 20));
  assert.equal(running, 2, '3.5 GB: the first lane, and one more with 1.5 GB left for it after the first');
  for (;;) { const r = release.shift(); if (!r) { await new Promise((x) => setTimeout(x, 10)); if (!release.length) break; continue; } r(); await new Promise((x) => setTimeout(x, 10)); }
  await done;
});

test('the parts to run again: those that failed, by lane; a failure outside the parts asks for its app whole, and only that app', () => {
  const apps = {
    sheets: { ok: false, failed: ['sheets: a picture'], failedParts: ['sheets#verifySheetPictures'] },
    word: { ok: true, failed: [] },
    access: { ok: false, failed: ['a11y-audit: mail'], failedParts: ['access#verifyA11yAudit'] },
    look: { ok: false, failed: ['look: set-up'], failedParts: ['*'] },
  };
  const machine = { ok: false, failed: ['machine: probe'], byApp: { shell: { failed: ['machine: probe'], failedParts: ['shell#verifyMachine'] }, mail: { failed: [], failedParts: [] } } };
  const r = partsToRerun(apps, machine);
  assert.deepEqual(r.parallel.apps, ['sheets', 'access', 'look']);
  assert.deepEqual(r.parallel.ids, ['sheets#verifySheetPictures', 'access#verifyA11yAudit', 'look#*'], 'the whole app is asked for by its own id, so the others still run only what failed');
  assert.deepEqual(r.parallel.whole, ['look']);
  assert.deepEqual(r.parallel.names, ['sheets: a picture', 'a11y-audit: mail', 'look: set-up']);
  assert.deepEqual(r.machine.ids, ['shell#verifyMachine']);
  assert.deepEqual(partsToRerun({ word: { ok: true, failed: [] } }, { ok: true }).parallel.names, []);
});

// Parts files as verify-apps.js writes them (`apps`), cut down to what the gate reads.
const partsFile = (parts) => {
  const apps = {};
  for (const [id, p] of Object.entries(parts)) {
    const app = id.split('#')[0];
    const a = (apps[app] ||= { passed: 0, of: 0, failed: [], parts: {} });
    // `real`: the failed checks that are the product's own (not a catch-all); `names`: the checks that passed.
    a.parts[id] = { of: p.of, passed: p.passed ?? p.of - (p.failed || []).length, failed: p.failed || [], realFailed: p.real || [], passedNames: p.names || [], threw: Boolean(p.threw) };
    a.of += p.of;
    a.passed += a.parts[id].passed;
    a.failed.push(...a.parts[id].failed);
  }
  return apps;
};

test('a failed part is cleared by the same part running again clean, whatever its failed checks were called', () => {
  // Gate 8 and 9: a catch-all called "the fixture was made" or a console check could never be seen passing by name.
  const first = partsFile({ 'word#verifyWordRedact': { of: 3, passed: 2, failed: ['word-redact: the fixture was made'] }, 'word#verifyWordInk': { of: 9 } });
  const again = partsFile({ 'word#verifyWordRedact': { of: 11 } });
  const r = rerunCleared(['word#verifyWordRedact'], first, again);
  assert.equal(r.ok, true, againSaid(r));
  assert.equal(againSaid(r), 'each passed alone');
});

test('a part that threw again, failed again, or ran fewer checks than passed before is not cleared, and says which', () => {
  const first = partsFile({ 'video#verifyVideoEditor': { of: 30, passed: 26, failed: ['video editor: Export writes', 'video editor: the exported file runs', 'video editor: the sound', 'video editor: a frame'] } });
  const threw = rerunCleared(['video#verifyVideoEditor'], first, partsFile({ 'video#verifyVideoEditor': { of: 30, threw: true } }));
  assert.equal(threw.ok, false);
  assert.deepEqual(threw.threw, ['video#verifyVideoEditor']);
  const failing = rerunCleared(['video#verifyVideoEditor'], first, partsFile({ 'video#verifyVideoEditor': { of: 30, failed: ['video editor: a frame'] } }));
  assert.equal(failing.ok, false);
  assert.match(againSaid(failing), /failed again: video editor: a frame/);
  const shrunk = rerunCleared(['video#verifyVideoEditor'], first, partsFile({ 'video#verifyVideoEditor': { of: 12 } }));
  assert.equal(shrunk.ok, false, 'twelve checks cannot stand for the twenty-six that passed before');
  assert.match(againSaid(shrunk), /fewer checks than passed the first time: video#verifyVideoEditor \(12 checks, 26 passed before\)/);
  const elsewhere = rerunCleared(['video#verifyVideoEditor'], first, partsFile({ 'video#verifyVideoEditor': { of: 30 }, 'video#verifyVideoMain': { of: 4, failed: ['video: other'] } }));
  assert.equal(elsewhere.ok, false, 'nothing else in the run again may fail');
});

test('a run again that found none of the failed parts is the gate\'s own fault, said as one, not "still failing" (gate 7 ran 0/0)', () => {
  const first = partsFile({ 'mail#verifyMailIndexWorker': { of: 8, passed: 7, failed: ['mail index worker: keys'] } });
  const none = rerunCleared(['mail#verifyMailIndexWorker'], first, partsFile({ 'mail#main': { of: 5 } }));
  assert.equal(none.ok, false);
  assert.deepEqual(none.missing, ['mail#verifyMailIndexWorker']);
  assert.match(againSaid(none), /^the gate's run again found none of these parts: mail#verifyMailIndexWorker \(a fault of the gate, not of Office\)/);
  const crashed = rerunCleared(['mail#verifyMailIndexWorker'], first, null);
  assert.equal(crashed.ok, false);
  assert.deepEqual(crashed.missing, [], 'a run again that wrote no parts file crashed: not the gate missing a part');
  assert.match(againSaid(crashed), /ended without writing its parts file/);
  const empty = rerunCleared(['mail#verifyMailIndexWorker'], first, { mail: { of: 0, passed: 0, failed: [], parts: { 'mail#verifyMailIndexWorker': { of: 0, passed: 0, failed: [] } } } });
  assert.deepEqual(empty.missing, ['mail#verifyMailIndexWorker'], 'a part asked for that counted nothing is missing too');
});

test('a failed check of the product\'s own must pass again by name: a run again down another path is not a pass (review of 11 Nov)', () => {
  const first = partsFile({ 'video#verifyVideoEditor': { of: 20, passed: 19, failed: ['video editor: a frame'], real: ['video editor: a frame'] } });
  const other = rerunCleared(['video#verifyVideoEditor'], first, partsFile({ 'video#verifyVideoEditor': { of: 19, names: Array.from({ length: 19 }, (_, i) => `video editor: ${i}`) } }));
  assert.equal(other.ok, false, 'nineteen checks that passed, none of them the one that failed');
  assert.match(againSaid(other), /did not pass again by name: video#verifyVideoEditor: video editor: a frame/);
  const same = rerunCleared(['video#verifyVideoEditor'], first, partsFile({ 'video#verifyVideoEditor': { of: 20, names: ['video editor: a frame', ...Array.from({ length: 19 }, (_, i) => `video editor: ${i}`)] } }));
  assert.equal(same.ok, true, againSaid(same));
});

test('a part\'s own catch-all, written only when it threw, is cleared by a clean run again that ran at least as much (second review of 11 Nov)', () => {
  // As done() writes it: the catch in wordPages said "the pages check ran" (not synthetic), after four checks of the nine it has.
  const first = partsFile({ 'word#wordPages': { of: 5, passed: 4, failed: ['word: the pages check ran'], real: ['word: the pages check ran'], names: ['word: a', 'word: b', 'word: c', 'word: d'] } });
  const clean = rerunCleared(['word#wordPages'], first, partsFile({ 'word#wordPages': { of: 9, names: ['word: a', 'word: b', 'word: c', 'word: d', 'word: e', 'word: f', 'word: g', 'word: h', 'word: i'] } }));
  assert.equal(clean.ok, true, againSaid(clean));
  // Fewer checks than the first run had in all: then each failed one must pass by name.
  const less = rerunCleared(['word#wordPages'], first, partsFile({ 'word#wordPages': { of: 4, names: ['word: a', 'word: b', 'word: c', 'word: d'] } }));
  assert.equal(less.ok, false);
  assert.match(againSaid(less), /did not pass again by name: word#wordPages: word: the pages check ran/);
});

test('a machine lane that wrote no parts file is not run again under a made-up name', () => {
  const r = partsToRerun({}, { ok: false, failed: [], byApp: null });
  assert.deepEqual(r.machine.ids, []);
  assert.equal(r.machine.unnamed, 'ended without its parts file');
  assert.equal(partsToRerun({}, { ok: false, failed: [], notMeasured: true }).machine.unnamed, 'not measured');
});

test('an app asked for whole is cleared by the app running clean with as many checks as passed before', () => {
  const first = { look: { of: 22, passed: 21, failed: ['look: set-up'], parts: { '*': { of: 1, passed: 0, failed: ['look: set-up'] }, 'look#verifyPanels': { of: 21, passed: 21, failed: [] } } } };
  assert.equal(rerunCleared(['look#*'], first, { look: { of: 22, passed: 22, failed: [], parts: {} } }).ok, true);
  assert.equal(rerunCleared(['look#*'], first, { look: { of: 5, passed: 5, failed: [], parts: {} } }).ok, false);
});

test('partsById reads each part by its id, and every app whole', () => {
  const by = partsById(partsFile({ 'mail#verifyMailLinks': { of: 3 }, 'mail#main': { of: 5, failed: ['mail: x'] } }));
  assert.equal(by['mail#verifyMailLinks'].of, 3);
  assert.equal(by['mail#*'].of, 8);
  assert.deepEqual(by['mail#*'].failed, ['mail: x']);
});

test('a part marked for the machine lane that the chunks passed by and the machine lane never reached is found', () => {
  const chunk = [{ id: 'mail#verifyMailSafety', machine: true, because: 'lane' }, { id: 'mail#verifyMailLinks', machine: false, because: null }, { id: 'mail#verifyNested', machine: true, because: 'lane' }];
  assert.deepEqual(leftToMachine(chunk), ['mail#verifyMailSafety', 'mail#verifyNested']);
  const machineSeen = [{ id: 'mail#verifyMailSafety', machine: true, because: null }, { id: 'mail#verifyMailLinks', machine: false, because: 'lane' }];
  assert.deepEqual(machinePartsUnrun(leftToMachine(chunk), machineSeen), ['mail#verifyNested']);
});

test('settle waits for a quiet machine, and says when it never was one rather than going on', async () => {
  const lines = [];
  let n = 0;
  const calm = await settle(60_000, { sample: async () => (n++ < 2 ? 0.9 : 0.1), log: (l) => lines.push(l) });
  assert.equal(calm.calm, true);
  assert.match(lines.at(-1), /settled for the budgets/);
  const busy = await settle(1500, { sample: async () => { await new Promise((r) => setTimeout(r, 300)); return 0.6; }, log: (l) => lines.push(l) });
  assert.equal(busy.calm, false);
  assert.equal(busy.busy, 60);
  assert.match(lines.at(-1), /did not settle: 60% busy/);
});

test('the battery is read from Win32_Battery: discharging is 1, 4 or 5, no battery is the mains', () => {
  assert.equal(onBattery({ platform: 'win32', query: () => '1\r\n' }), true);
  assert.equal(onBattery({ platform: 'win32', query: () => '2\r\n' }), false);
  assert.equal(onBattery({ platform: 'win32', query: () => '4' }), true, 'low is still on the battery');
  assert.equal(onBattery({ platform: 'win32', query: () => '5' }), true, 'and critical');
  assert.equal(onBattery({ platform: 'win32', query: () => '' }), false);
  assert.equal(onBattery({ platform: 'linux', query: () => '1' }), false);
});

test('a part is named by the function it runs, the same in every lane whatever was skipped before it', () => {
  assert.equal(partSlug(() => verifyMailLinks({ a: 1 })), 'verifyMailLinks');
  assert.equal(partSlug(async () => await verifyX()), 'verifyX');
  assert.equal(partSlug(() => { const a = 1; return a; }), null);
  // The same run in each lane: the machine lane enters no section, and the ids still agree.
  const walk = (sel) => {
    const out = [];
    if (sel.section('mail', 'reading').runs) { out.push(sel.part('mail', () => verifyMailLinks())); }
    out.push(sel.part('mail', () => verifyMailSafety(), { machine: true }));
    out.push(sel.part('mail', () => verifyMailIndexWorker(), { machine: true }));
    out.push(sel.part('talk', () => verifyTalk()));
    return out;
  };
  const parallel = walk(createSelector({ lane: 'parallel' }));
  const machine = walk(createSelector({ lane: 'machine' }));
  assert.deepEqual(parallel.map((p) => p.id), ['mail#verifyMailLinks', 'mail#verifyMailSafety', 'mail#verifyMailIndexWorker', 'talk#verifyTalk']);
  assert.deepEqual(machine.map((p) => p.id), ['mail#verifyMailSafety', 'mail#verifyMailIndexWorker', 'talk#verifyTalk']);
  assert.deepEqual(parallel.filter((p) => p.runs).map((p) => p.id), ['mail#verifyMailLinks', 'talk#verifyTalk']);
  assert.deepEqual(machine.filter((p) => p.runs).map((p) => p.id), ['mail#verifyMailSafety', 'mail#verifyMailIndexWorker']);
});

test('asked for one part, a run runs that part alone: no section of its app beside it (owner, 11 Nov: smaller groups)', () => {
  const walk = (sel) => {
    const ran = [];
    for (const [app, name] of [['mail', 'reading'], ['mail', 'accountDialog']]) if (sel.section(app, name).runs) ran.push(`${app}#${name}`);
    for (const run of [() => verifyMailLinks(), () => verifyMailDrafts(), () => verifyTalk()]) { const app = String(run).includes('Talk') ? 'talk' : 'mail'; const p = sel.part(app, run); if (p.runs) ran.push(p.id); }
    if (sel.section('talk', 'connect').runs) ran.push('talk#connect');
    return ran;
  };
  assert.deepEqual(walk(createSelector({ lane: 'parallel', apps: ['mail'], ids: ['mail#verifyMailLinks'] })), ['mail#verifyMailLinks']);
  assert.deepEqual(walk(createSelector({ lane: 'parallel', apps: ['mail'], ids: ['mail#reading'] })), ['mail#reading'], 'a section alone');
  assert.deepEqual(walk(createSelector({ lane: 'parallel', apps: ['talk'], ids: ['talk#verifyTalk'] })), ['talk#verifyTalk']);
  assert.deepEqual(walk(createSelector({ lane: 'parallel', apps: ['mail'], ids: ['mail#*'] })), ['mail#reading', 'mail#accountDialog', 'mail#verifyMailLinks', 'mail#verifyMailDrafts'], 'the app whole');
  assert.deepEqual(walk(createSelector({ lane: 'machine' })), [], "sections are the parallel lane's");
  const sel = createSelector({ lane: 'parallel', ids: ['talk#verifyTalk'] });
  assert.equal(sel.wants('talk'), true);
  assert.equal(sel.wants('mail'), false);
});

test('a section named twice, or as a part is, is refused rather than run under a shared id', () => {
  const sel = createSelector({});
  assert.equal(sel.section('mail', 'reading').runs, true);
  assert.deepEqual(sel.section('mail', 'reading'), { id: 'mail#reading', runs: false, because: 'twice' });
  assert.equal(sel.part('mail', () => verifyMailLinks()).runs, true);
  assert.equal(sel.section('mail', 'verifyMailLinks').because, 'twice');
  assert.equal(sel.section('mail', 'a b').because, 'unnamed');
});

test('two parts that run the same function, or one not written as () => verifySomething(), are refused rather than run under a shared id', () => {
  const sel = createSelector({});
  assert.equal(sel.part('mail', () => verifyMailLinks()).runs, true);
  assert.deepEqual(sel.part('mail', () => verifyMailLinks({ other: true })), { id: 'mail#verifyMailLinks', runs: false, because: 'twice' });
  assert.equal(sel.part('mail', () => { verifyMailLinks(); }).because, 'unnamed');
});

test('every part in verify-apps.js runs one function of its own, so each has an id the gate can ask for', async () => {
  const fs = await import('node:fs');
  const src = fs.readFileSync(new URL('../apps/desktop/main/verify-apps.js', import.meta.url), 'utf8');
  const calls = [...src.matchAll(/await part\('([a-z]+)', ((?:async )?\(\) => [^\n]*)/g)];
  assert.ok(calls.length > 300, `${calls.length} parts found`);
  const ids = calls.map(([, app, run]) => `${app}#${/=>\s*(?:await\s+)?([A-Za-z_$][\w$]*)\s*\(/.exec(run)?.[1]}`);
  assert.deepEqual(ids.filter((id) => id.endsWith('#undefined')), []);
  assert.deepEqual(ids.filter((id, i) => ids.indexOf(id) !== i), [], 'no two parts share an id');
  assert.equal((src.match(/await part\(/g) || []).length, calls.length, 'every part is in the one shape');
});

test('the record names each check that ran again', () => {
  const counts = { 'verify:apps': { passed: 10, of: 10 } };
  const plain = gateSection(counts, { version: '9.9.9', date: '2026-10-21', platform: 'win32', arch: 'x64' });
  assert.doesNotMatch(plain, /run again/);
  const again = gateSection(counts, { version: '9.9.9', date: '2026-10-21', platform: 'win32', arch: 'x64', reran: ['sheets: a picture'] });
  assert.match(again, /One check failed beside the others and passed when run again alone/);
  assert.match(again, /\n- sheets: a picture\n$/);
});
