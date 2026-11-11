// The gate's failure list and the check numbers (owner, 11 Nov 2026: "find
// all the failures and then fix the failure before re running", and a number
// beside each check's words that says exactly where it failed), with what the
// review of the same day found wrong with the first version.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { emptyLedger, readLedger, writeLedger, runFindings, recordRun, openEntries, findEntry, closeEntry, moveEntry, openParts, ledgerLines, ledgerMarkdown, reasonClosures, markRecorded, OWNER_QUOTED, CRASH_PLACEHOLDER } from '../tools/gate-ledger.js';
import { readRegistry, writeRegistry, idOf, assign, callSite, REGISTRY } from '../apps/desktop/main/verify-ids.js';
import { dirtyForRecord, gateSection, treeKey, UNKEYED } from '../tools/gate.js';

// Commits A, then B, then C, each holding the one before.
const A = 'a'.repeat(40);
const B = 'b'.repeat(40);
const C = 'c'.repeat(40);
const D = 'd'.repeat(40);
const order = [A, B, C, D];
const descends = (head, commit) => order.indexOf(commit) >= 0 && order.indexOf(commit) <= order.indexOf(head);

const SIGNIN = 'mail: sign-in shows the browsers';
const failing = (name, extra = {}) => ({ id: '18.0007', name, detail: 'no select', at: 'verify-mail-signin.js:54', cause: { what: 'the browser list', ms: 15000, at: 'verify-mail-signin.js:31' }, ...extra });
/** One app's counts as the gate keeps them, its parts given as { id: { failures, passedNames, threw } }. */
const app = (parts) => ({ parts: Object.fromEntries(Object.entries(parts).map(([id, p]) => [id, { of: (p.passedNames || []).length + (p.failures || []).length, threw: false, passedNames: [], failures: [], ...p }])) });
const mailFails = () => ({ mail: app({ 'mail#verifyMailSignIn': { failures: [failing(SIGNIN)] } }) });
const mailPasses = () => ({ mail: app({ 'mail#verifyMailSignIn': { passedNames: [SIGNIN] } }) });

test("a run's failures, with their numbers, lines, parts and lanes, open entries in the list", () => {
  const findings = runFindings({
    apps: mailFails(),
    machine: { video: app({ 'video#verifyVideoBudget': { failures: [{ id: '12.0003', name: 'video: export within 60 s' }] } }) },
  });
  assert.deepEqual(findings.failures.map((f) => [f.app, f.part, f.lane]), [['mail', 'mail#verifyMailSignIn', 'parallel'], ['video', 'video#verifyVideoBudget', 'machine']]);
  const ledger = recordRun(emptyLedger(), { head: A }, findings);
  assert.equal(openEntries(ledger).length, 2);
  const e = findEntry(ledger, '18.0007').entry;
  assert.equal(e.name, SIGNIN);
  assert.equal(e.at, 'verify-mail-signin.js:54');
  assert.equal(e.last.cause.what, 'the browser list');
  assert.deepEqual(openParts(ledger).sort(), ['mail#verifyMailSignIn', 'video#verifyVideoBudget']);
});

test('the same words failing in two apps are two entries, each run again in its own app', () => {
  const catchAll = "the checks outside a part logged no error in a window's console";
  const findings = runFindings({ apps: { mail: app({ '*': { failures: [{ name: catchAll, synthetic: true }] } }), word: app({ '*': { failures: [{ name: catchAll, synthetic: true }] } }) } });
  const ledger = recordRun(emptyLedger(), { head: A }, findings);
  assert.equal(openEntries(ledger).length, 2);
  assert.deepEqual(openParts(ledger).sort(), ['mail#*', 'word#*']);
  // Each has a number of its own from the list, as a check the run said of itself.
  assert.deepEqual(openEntries(ledger).map((e) => e.id).sort(), ['L.0001', 'L.0002']);
  // Its words name both: closing by them is refused, by number it is not.
  assert.match(closeEntry(ledger, catchAll, { reason: 'the console noise was a test window' }).said, /names 2 entries/);
});

test('a check the run says of itself is proved by its part running clean', () => {
  const threw = 'mail: part mail#verifyMailLinks ran to its end';
  const ledger = recordRun(emptyLedger(), { head: A }, runFindings({ apps: { mail: app({ 'mail#verifyMailLinks': { threw: true, passedNames: ['mail: a link opens'], failures: [{ name: threw, synthetic: true }] } }) } }));
  recordRun(ledger, { head: B }, runFindings({ apps: { mail: app({ 'mail#verifyMailLinks': { passedNames: ['mail: a link opens', 'mail: a second link'] } }) } }));
  const r = closeEntry(ledger, 'L.0001', { by: B, descends });
  assert.equal(r.ok, true, r.said);
});

test('a window process that ended early is an entry, and its app whole runs again for it', () => {
  const findings = runFindings({ apps: { pdf: { parts: {}, failed: [CRASH_PLACEHOLDER] } }, machineCrashed: true, machineScope: ['video', 'recorder'] });
  const ledger = recordRun(emptyLedger(), { head: A }, findings);
  const names = openEntries(ledger).map((e) => e.name).sort();
  assert.deepEqual(names, ['machine lane: the window process ran to its end and wrote its counts', 'pdf: the window process ran to its end and wrote its counts']);
  assert.deepEqual(openParts(ledger).sort(), ['pdf#*', 'recorder#*', 'video#*']);
  // Not a check: given no number from the registry.
  assert.ok(openEntries(ledger).every((e) => e.id.startsWith('L.') && e.synthetic));
});

test('a failure that passed only when run again alone is flaky, and stays open', () => {
  const findings = runFindings({ apps: mailFails(), again: [mailPasses()] });
  const ledger = recordRun(emptyLedger(), { head: A }, findings);
  const e = findEntry(ledger, SIGNIN).entry;
  assert.equal(e.flaky, true);
  assert.equal(e.status, 'open');
  assert.equal(e.passedAt.length, 0, 'the pass in the run again is no proof');
  assert.match(ledgerLines(openEntries(ledger)).join('\n'), /flaky: passed alone/);
});

test('what failed only in the run again gets an entry of its own', () => {
  const findings = runFindings({ apps: mailFails(), again: [{ mail: app({ 'mail#verifyMailSignIn': { failures: [failing(SIGNIN), { id: '18.0008', name: 'mail: the browser chosen is remembered' }] } }) }] });
  const ledger = recordRun(emptyLedger(), { head: A }, findings);
  assert.equal(openEntries(ledger).length, 2);
  assert.equal(findEntry(ledger, '18.0008').entry.last.onlyAgain, true);
  assert.equal(findEntry(ledger, '18.0007').entry.failures, 1, 'failing in both runs is one failure');
});

test('a crashed chunk the run again cleared is still on the list, as flaky', () => {
  const findings = runFindings({ apps: { pdf: { parts: {}, failed: [CRASH_PLACEHOLDER] } }, again: [{ pdf: app({ 'pdf#verifyPdfApp': { passedNames: ['pdf: opens'] } }) }], cleared: new Set(['pdf']) });
  const ledger = recordRun(emptyLedger(), { head: A }, findings);
  assert.equal(openEntries(ledger)[0].flaky, true);
});

test('closing by a commit: made after the failure, and proved by a clean run at or after it', () => {
  const ledger = recordRun(emptyLedger(), { head: B }, runFindings({ apps: mailFails() }));
  // A commit the failing run already held is no fix.
  recordRun(ledger, { head: B }, runFindings({ apps: mailPasses() }));
  assert.match(closeEntry(ledger, '18.0007', { by: B, descends }).said, /a commit made after the failure/);
  assert.match(closeEntry(ledger, '18.0007', { by: A, descends }).said, /a commit made after the failure/);
  // The fix is C; a pass on a tree with changes beyond C is no proof.
  recordRun(ledger, { head: C, clean: false }, runFindings({ apps: mailPasses() }));
  assert.equal(closeEntry(ledger, '18.0007', { by: C, descends }).ok, false);
  recordRun(ledger, { head: D }, runFindings({ apps: mailPasses() }));
  const r = closeEntry(ledger, '18.0007', { by: C, descends });
  assert.equal(r.ok, true, r.said);
  assert.equal(openEntries(ledger).length, 0);
  assert.equal(closeEntry(ledger, '18.0007', { by: C, descends }).ok, false);
});

test('a flaky entry needs two clean passes after its fix', () => {
  const ledger = recordRun(emptyLedger(), { head: A }, runFindings({ apps: mailFails(), again: [mailPasses()] }));
  recordRun(ledger, { head: B }, runFindings({ apps: mailPasses() }));
  assert.match(closeEntry(ledger, '18.0007', { by: B, descends }).said, /1 of the 2 runs/);
  recordRun(ledger, { head: B }, runFindings({ apps: mailPasses() }));
  assert.equal(closeEntry(ledger, '18.0007', { by: B, descends }).ok, true);
});

test('a failure after a pass takes the pass away, and a closed entry that fails again is opened again', () => {
  const ledger = recordRun(emptyLedger(), { head: A }, runFindings({ apps: mailFails() }));
  recordRun(ledger, { head: B }, runFindings({ apps: mailPasses() }));
  recordRun(ledger, { head: B }, runFindings({ apps: mailFails() }));
  assert.equal(findEntry(ledger, SIGNIN).entry.passedAt.length, 0);
  assert.equal(findEntry(ledger, SIGNIN).entry.failures, 2);
  recordRun(ledger, { head: C }, runFindings({ apps: mailPasses() }));
  assert.equal(closeEntry(ledger, SIGNIN, { by: C, descends }).ok, true);
  recordRun(ledger, { head: C }, runFindings({ apps: mailFails() }));
  const e = findEntry(ledger, SIGNIN).entry;
  assert.equal(e.status, 'open');
  assert.equal(e.reopened, 1);
  assert.match(ledgerLines([e]).join('\n'), /failed again after a close/);
});

test('a reason closes an entry only when it says why; a flaky one only with the owner quoted; each said once in the record', () => {
  const ledger = recordRun(emptyLedger(), { head: A }, runFindings({ apps: { pdf: { failed: ['pdf: prints'], parts: {} } } }));
  assert.equal(closeEntry(ledger, 'pdf: prints', { reason: 'flaky' }).ok, false);
  assert.equal(closeEntry(ledger, 'nothing here', { reason: 'the printer was unplugged on purpose' }).ok, false);
  assert.equal(closeEntry(ledger, 'pdf: prints', { reason: 'the printer was unplugged on purpose' }).ok, true);
  recordRun(ledger, { head: A }, runFindings({ apps: mailFails(), again: [mailPasses()] }));
  assert.equal(closeEntry(ledger, '18.0007', { reason: 'the machine was busy that hour' }).ok, false);
  assert.equal(OWNER_QUOTED.test('owner: "leave it for 2.26, the browsers list is mail\'s"'), true);
  assert.equal(OWNER_QUOTED.test('the ownership was "unclear to me here"'), false);
  assert.equal(closeEntry(ledger, '18.0007', { reason: 'owner, 11 Nov: "the browser list goes to the mail session"' }).ok, true);
  const said = reasonClosures(ledger);
  assert.deepEqual(said.map((c) => c.name).sort(), [SIGNIN, 'pdf: prints']);
  const section = gateSection({}, { version: '2.25.0', date: '12 November 2026', closedByReason: said });
  assert.match(section, /2 failed checks were closed with a reason rather than a fixing commit/);
  assert.match(section, /- 18\.0007 mail: sign-in shows the browsers: owner, 11 Nov/);
  markRecorded(ledger, '2.25.0');
  assert.deepEqual(reasonClosures(ledger), []);
  assert.doesNotMatch(gateSection({}, { version: '2.25.0', date: 'x' }), /closed with a reason/);
});

test('the list goes to disk whole and back; one that does not read stops the gate rather than read as empty', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-ledger-'));
  try {
    const file = path.join(dir, 'ledger.json');
    assert.deepEqual(readLedger(file), emptyLedger());
    const ledger = recordRun(emptyLedger(), { head: A }, runFindings({ apps: { mail: app({ 'mail#verifyMailSignIn': { failures: [failing('mail: a | b')] } }) } }));
    writeLedger(ledger, file);
    assert.deepEqual(fs.readdirSync(dir), ['ledger.json'], 'no file left beside it');
    const back = readLedger(file);
    assert.equal(openEntries(back).length, 1);
    const page = ledgerMarkdown(back, { head: A });
    assert.match(page, /\| 18\.0007 \| mail \| mail#verifyMailSignIn \| verify-mail-signin\.js:54 \| 1 \|/);
    assert.match(page, /mail: a \/ b/);
    fs.writeFileSync(file, '{"entries": {');
    assert.throws(() => readLedger(file), /does not read/);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('check numbers: given once per name, by app, never reused, and steady on disk', () => {
  const reg = { apps: {}, next: {}, ids: {} };
  assert.deepEqual(assign(reg, [{ name: 'mail: one', app: 'mail' }, { name: 'word: one', app: 'word' }, { name: 'mail: two', app: 'mail' }], ['shell', 'word', 'mail']), ['mail: one', 'word: one', 'mail: two']);
  assert.equal(idOf(reg, 'mail: one'), '03.0001');
  assert.equal(idOf(reg, 'mail: two'), '03.0002');
  assert.equal(idOf(reg, 'word: one'), '02.0001');
  assert.equal(idOf(reg, 'never seen'), 'new');
  assert.deepEqual(assign(reg, [{ name: 'mail: one', app: 'mail' }]), []);
  assign(reg, [{ name: 'locker: one', app: 'locker' }], ['locker']);
  assert.equal(idOf(reg, 'locker: one'), '04.0001');
  assign(reg, [{ name: 'odd one', app: null }]);
  assert.equal(idOf(reg, 'odd one'), '01.0001');
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-ids-'));
  try {
    const file = path.join(dir, 'ids.json');
    writeRegistry(reg, file);
    const text = fs.readFileSync(file, 'utf8');
    writeRegistry(readRegistry(file), file);
    assert.equal(fs.readFileSync(file, 'utf8'), text);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('the committed registry numbers every app, holds no number twice, and no check the run says of itself', () => {
  const reg = readRegistry(REGISTRY);
  assert.ok(Object.keys(reg.ids).length > 1000);
  const seen = new Set();
  for (const id of Object.values(reg.ids)) {
    assert.match(id, /^\d{2}\.\d{4}$/);
    assert.ok(!seen.has(id), `${id} twice`);
    seen.add(id);
  }
  for (const [a, next] of Object.entries(reg.next)) {
    const no = String(reg.apps[a]).padStart(2, '0');
    assert.ok(Object.values(reg.ids).every((id) => !id.startsWith(`${no}.`) || Number(id.slice(3)) < next), `${a} next ${next}`);
  }
  assert.equal(reg.ids[CRASH_PLACEHOLDER], undefined);
  // The commit hook refuses these characters in an added line: written as escapes, the names unchanged.
  assert.doesNotMatch(fs.readFileSync(REGISTRY, 'utf8'), /[\u2014\u200b\u200c\u200e\u200f\u2060\u00a0\ufeff]/);
  assert.ok(Object.keys(reg.ids).some((n) => n.includes('\u2014')), 'a name with the long dash reads back with it');
});

test('the place of a check is the line that called it, past the waits', () => {
  const stack = [
    'Error',
    '    at check (D:\\o\\apps\\desktop\\main\\verify-apps.js:120:22)',
    '    at until (D:\\o\\apps\\desktop\\main\\verify-waits.js:40:9)',
    '    at verifyMailSignIn (file:///D:/o/apps/desktop/main/verify-mail-signin.js:54:5)',
    '    at helper (file:///D:/o/apps/desktop/main/verify-mail-signin.js:31:3)',
    '    at run (D:\\o\\apps\\desktop\\main\\verify-apps.js:300:1)',
  ].join('\n');
  assert.equal(callSite(stack, 2), 'verify-mail-signin.js:54');
  assert.equal(callSite(stack, 2, 2), 'verify-mail-signin.js:54 < :31');
  assert.equal(callSite('Error', 2), null);
});

test("the numbers a run gives new checks are the gate's own change: not a dirty tree, and not a new tree to resume", () => {
  assert.deepEqual(dirtyForRecord(' M tools/check-ids.json\n M tools/gate-floors.json\n', '2.25.0'), []);
  assert.equal(dirtyForRecord(' M tools/gate.js\n', '2.25.0').length, 1);
  assert.equal(UNKEYED, 'tools/check-ids.json');
  const file = path.join(path.dirname(REGISTRY), 'check-ids.json');
  const before = fs.readFileSync(file);
  const key = treeKey();
  try {
    fs.appendFileSync(file, ' ');
    assert.equal(treeKey(), key);
  } finally {
    fs.writeFileSync(file, before);
  }
});

test('a part that threw and then ran clean alone is flaky (review of 11 Nov, second pass)', () => {
  const threw = 'mail: part mail#verifyMailLinks ran to its end';
  const first = { mail: app({ 'mail#verifyMailLinks': { threw: true, failures: [{ name: threw, synthetic: true }] } }) };
  const again = { mail: app({ 'mail#verifyMailLinks': { passedNames: ['mail: a link opens'] } }) };
  const ledger = recordRun(emptyLedger(), { head: A }, runFindings({ apps: first, again: [again] }));
  const e = openEntries(ledger)[0];
  assert.equal(e.flaky, true);
  assert.equal(closeEntry(ledger, e.id, { reason: 'a busy hour on the machine' }).ok, false, 'a reason alone needs the owner quoted');
});

test('an app whole is proved only by a run of the app whole, and the machine lane only by a run with checks in it', () => {
  const catchAll = "the checks outside a part logged no error in a window's console";
  const ledger = recordRun(emptyLedger(), { head: A }, runFindings({ apps: { mail: app({ '*': { failures: [{ name: catchAll, synthetic: true }] } }) }, machineCrashed: true, machineScope: ['mail'] }));
  // Some of mail's parts, asked for by id, and a machine lane that ran nothing.
  recordRun(ledger, { head: B }, runFindings({ apps: { mail: app({ 'mail#reading': { passedNames: ['mail: rows'] } }) }, machine: { mail: { of: 0, parts: {} } }, narrowed: new Set(['mail']) }));
  assert.ok(openEntries(ledger).every((e) => e.passedAt.length === 0));
  // The app whole, and a machine lane with checks.
  recordRun(ledger, { head: C }, runFindings({ apps: { mail: app({ 'mail#reading': { passedNames: ['mail: rows'] } }) }, machine: { mail: app({ 'mail#verifyMailSafety': { passedNames: ['mail: safety'] } }) } }));
  assert.ok(openEntries(ledger).every((e) => e.passedAt.length === 1));
});

test("an engine test moved to another file to fix it takes its entry with it, and only passes there prove it", () => {
  const ledger = emptyLedger();
  const old = 'tests/locker-edges.test.js > a large file comes back';
  recordRun(ledger, { head: 'a1' }, { failures: [{ name: old, id: '00.0001', key: `engine|test#tests/locker-edges.test.js|${old}`, app: 'engine', part: 'test#tests/locker-edges.test.js', lane: 'local', flaky: true }] });
  const files = { 'tests/locker-edges.test.js': ['a large file comes back'], 'tests/locker-large.test.js': [] };
  const has = (file, title) => (files[file] || []).includes(title);
  // Not moved yet: refused, and the entry is as it was.
  assert.match(moveEntry(ledger, '00.0001', 'tests/locker-large.test.js', { has }).said, /still has/);
  files['tests/locker-edges.test.js'] = [];
  assert.match(moveEntry(ledger, '00.0001', 'tests/locker-large.test.js', { has }).said, /has no test/);
  files['tests/locker-large.test.js'] = ['a large file comes back'];
  assert.match(moveEntry(ledger, '00.0001', 'tests/locker-edges.test.js', { has }).said, /name the file/);
  const r = moveEntry(ledger, '00.0001', 'tests/locker-large.test.js', { has });
  assert.ok(r.ok, r.said);
  const e = findEntry(ledger, '00.0001').entry;
  assert.deepEqual([e.name, e.part, e.flaky, e.failures, e.passedAt.length], ['tests/locker-large.test.js > a large file comes back', 'test#tests/locker-large.test.js', true, 1, 0]);
  assert.equal(e.moved.length, 1);
  // A pass of the old name proves nothing now; the new name's, on the lab machine it belongs to, does.
  recordRun(ledger, { head: 'b2' }, { passed: new Set([`engine@home|${old}`]) });
  assert.equal(findEntry(ledger, '00.0001').entry.passedAt.length, 0);
  recordRun(ledger, { head: 'b2' }, { passed: new Set(['engine@home|tests/locker-large.test.js > a large file comes back']) });
  assert.equal(findEntry(ledger, '00.0001').entry.passedAt.length, 1);
  // A window check's entry has no file to move with.
  recordRun(ledger, { head: 'a1' }, { failures: [{ name: 'word: x', id: '07.0001', key: 'word|word#a|word: x', app: 'word', part: 'word#a', lane: 'parallel' }] });
  assert.match(moveEntry(ledger, '07.0001', 'tests/locker-large.test.js', { has }).said, /not an engine test/);
});
