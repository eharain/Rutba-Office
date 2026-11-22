// A part of the gate keeps its pass from an earlier commit only while no
// file it runs has changed since (tools/gate-carry.js, D-2026-11-16-02).
// What is held here: which changes void which parts, that nothing is carried
// on a doubt, that the record says what was carried, and that the passes
// taken to open no window really import nothing of the windows' own.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import esbuild from 'esbuild';
import { CLEAN_TREE, ENGINE_ONLY, kindOf, passKept, carry, carryOver, carriedNote, packsKey, isWindowsOwn } from '../tools/gate-carry.js';
import { VERIFY_APPS, appsForPath, treeKey, gateSection, FUZZERS } from '../tools/gate.js';

const ROOT = path.join(import.meta.dirname, '..');
const A = 'a'.repeat(40);
const B = 'b'.repeat(40);
const C = 'c'.repeat(40);
const PASSES = ['test', 'test:perf', 'test:removal', 'verify:edit', 'smoke', 'verify:packs', 'office:reopen', ...FUZZERS.map(([name]) => name), 'corpus'];

/** A state in which everything passed on commit A. */
const full = () => ({
  head: A,
  tree: CLEAN_TREE,
  passes: Object.fromEntries(PASSES.map((name) => [name, { ok: true, counts: { passed: 3, of: 3 } }])),
  apps: Object.fromEntries(VERIFY_APPS.map((a) => [a, { ok: true, passed: 2, of: 2, failed: [] }])),
  machine: { ok: true, passed: 5, of: 5, failed: [], scope: VERIFY_APPS.slice() },
  priority: ['raised'],
});
const over = (files, old = full(), head = B) => carry(old, { head, files, appsForPath, apps: VERIFY_APPS });
const sorted = (list) => list.slice().sort();

test('a clean checkout has the tree key that carrying asks for', (t) => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-gate-carry-'));
  t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
  const git = (...args) => execFileSync('git', ['-c', 'user.name=Gate test', '-c', 'user.email=gate@test.invalid', ...args], { cwd: dir, encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] });
  git('init', '-q');
  fs.writeFileSync(path.join(dir, 'a.txt'), 'one\n');
  git('add', 'a.txt');
  git('commit', '-q', '-m', 'one');
  assert.equal(treeKey(dir), CLEAN_TREE);
  // A file edited, and a file git does not know: neither is a clean tree.
  fs.writeFileSync(path.join(dir, 'a.txt'), 'two\n');
  assert.notEqual(treeKey(dir), CLEAN_TREE);
  git('checkout', '-q', '--', 'a.txt');
  assert.equal(treeKey(dir), CLEAN_TREE);
  fs.writeFileSync(path.join(dir, 'b.txt'), 'new\n');
  assert.notEqual(treeKey(dir), CLEAN_TREE);
});

test('what a changed file is to the passes', () => {
  for (const f of ['tools/gate.js', 'tools/gate-plan.js', 'tools/gate-carry.js', 'tools/gate-floors.json', 'tools/run-temp.js', 'docs/releases/v2.27.0.md']) assert.equal(kindOf(f), 'gate', f);
  for (const f of ['docs/APP-PACKS.md', 'README.md', 'CHANGELOG.md', 'tools/check-ids.json', '.github/workflows/upgrade.yml', '.github/workflows/candidate.yml']) assert.equal(kindOf(f), 'words', f);
  for (const f of ['apps/desktop/renderer/apps/slides/ink.js', 'apps/desktop/renderer/index.js', 'apps/desktop/main/verify-deck-ink.js', 'apps\\desktop\\main\\verify-apps.js']) assert.equal(kindOf(f), 'windows', f);
  // Everything else is code some pass runs: the engines, the main process, the tests, the tools, what is installed, the build.
  for (const f of ['packages/ooxml/src/index.js', 'apps/desktop/main/documents.js', 'apps/desktop/main/index.js', 'tests/ooxml.test.js', 'tests/perf-budget.perf.js', 'tools/fuzz-pdf.js', 'tools/corpus.js', 'package.json', 'package-lock.json', 'packages/templates', 'apps/desktop/build/bundle.js', 'apps/desktop/resources/installer.nsh', '.github/workflows/release.yml', 'docs.js', 'notes/README.md']) assert.equal(kindOf(f), 'code', f);
});

test('the engine tests never keep their pass over a change, and no pass keeps it over code', () => {
  for (const name of PASSES) assert.equal(passKept(name, []), true, name);
  assert.equal(passKept('test', ['README.md']), false);
  for (const name of PASSES.filter((n) => n !== 'test')) {
    assert.equal(passKept(name, ['README.md', 'docs/APP-PACKS.md', 'tools/check-ids.json']), true, name);
    assert.equal(passKept(name, ['README.md', 'packages/pdf/src/index.js']), false, name);
    assert.equal(passKept(name, ['tests/fixtures/pst-writer.js']), false, name);
    assert.equal(passKept(name, ['apps/desktop/renderer/apps/word/window.js']), ENGINE_ONLY.includes(name), name);
  }
  assert.deepEqual(sorted(ENGINE_ONLY), sorted(['test:perf', ...FUZZERS.map(([name]) => name)]));
});

test('a fix to one app: the others keep their passes, and what the fix could touch runs again', () => {
  const r = over(['apps/desktop/renderer/apps/slides/ink.js', 'apps/desktop/main/verify-deck-ink.js', 'tools/check-ids.json']);
  assert.deepEqual(sorted(Object.keys(r.state.apps)), sorted(VERIFY_APPS.filter((a) => a !== 'slides')));
  assert.equal(r.state.machine, null, 'the machine lane ran slides too');
  assert.deepEqual(sorted(Object.keys(r.state.passes)), sorted(ENGINE_ONLY));
  assert.deepEqual(sorted(r.again), sorted(['slides', 'machine', ...PASSES.filter((n) => !ENGINE_ONLY.includes(n))]));
  assert.equal(r.state.head, B);
  assert.equal(r.state.tree, CLEAN_TREE);
  assert.deepEqual(r.state.priority, ['raised'], 'how the kept runs were made stays with them');
  for (const e of [...Object.values(r.state.apps), ...Object.values(r.state.passes)]) assert.equal(e.at, A);
});

test('an engine several apps share: those apps and every pass run again', () => {
  const r = over(['packages/ooxml/src/workbook.js']);
  assert.deepEqual(sorted(Object.keys(r.state.apps)), sorted(VERIFY_APPS.filter((a) => !['word', 'sheets', 'slides', 'base', 'mail'].includes(a))));
  assert.deepEqual(r.state.passes, {});
  assert.equal(r.state.machine, null);
});

test('words alone: only the engine tests run again', () => {
  const r = over(['README.md', 'docs/APP-PACKS.md', '.github/workflows/upgrade.yml', 'tools/check-ids.json']);
  assert.deepEqual(sorted(Object.keys(r.state.apps)), sorted(VERIFY_APPS));
  assert.equal(r.state.machine.ok, true);
  assert.equal(r.state.machine.at, A);
  assert.deepEqual(sorted(Object.keys(r.state.passes)), sorted(PASSES.filter((n) => n !== 'test')));
  assert.deepEqual(r.again, ['test']);
});

test('a test or a tool changed: the windows keep their passes, the passes do not', () => {
  const r = over(['tests/ooxml.test.js', 'tools/fuzz-open.js']);
  assert.deepEqual(sorted(Object.keys(r.state.apps)), sorted(VERIFY_APPS));
  assert.equal(r.state.machine.ok, true);
  assert.deepEqual(r.state.passes, {});
});

test('what every part runs voids every pass', () => {
  for (const f of ['packages/office-ui/src/index.js', 'apps/desktop/main/index.js', 'apps/desktop/renderer/index.js', 'package-lock.json', 'packages/templates', 'apps/desktop/build/bundle.js', 'apps/desktop/package.json']) {
    const r = over([f, 'README.md']);
    assert.deepEqual(Object.keys(r.state?.apps || {}), [], f);
    assert.equal(r.state?.machine ?? null, null, f);
    assert.deepEqual(Object.keys(r.state?.passes || {}).filter((n) => !ENGINE_ONLY.includes(n)), [], f);
  }
  // The gate itself, and the note the build takes into the app: nothing at all is kept.
  for (const f of ['tools/gate.js', 'tools/gate-carry.js', 'tools/run-temp.js', 'docs/releases/v2.27.0.md']) {
    const r = over([f]);
    assert.equal(r.state, null, f);
    assert.deepEqual(r.kept, [], f);
    assert.equal(r.again.length, PASSES.length + VERIFY_APPS.length + 1, f);
  }
});

test('nothing changed between the two commits: every pass stands, the engine tests too', () => {
  const r = over([]);
  assert.deepEqual(sorted(Object.keys(r.state.passes)), sorted(PASSES));
  assert.deepEqual(sorted(Object.keys(r.state.apps)), sorted(VERIFY_APPS));
  assert.equal(r.state.machine.at, A);
  assert.deepEqual(r.again, []);
});

test('what had not passed is never carried, and a pass keeps the commit it really ran on', () => {
  const old = full();
  old.apps.word = { ok: false, passed: 1, of: 2, failed: ['word: x'] };
  old.passes.smoke = { ok: false, counts: null };
  old.machine = { ok: true, passed: 0, of: 0, failed: [], serial: true, scope: [] };
  const first = over(['README.md'], old);
  assert.equal(first.state.apps.word, undefined);
  assert.equal(first.state.passes.smoke, undefined);
  assert.equal(first.state.machine, null, 'a --serial run is no machine lane');
  assert.ok(!first.kept.includes('word') && !first.again.includes('word'));
  // On B the mail checks ran again; then to C over a change to slides alone.
  first.state.apps.mail = { ok: true, passed: 2, of: 2, failed: [] };
  const second = over(['apps/desktop/renderer/apps/slides/ink.js'], first.state, C);
  assert.equal(second.state.head, C);
  assert.equal(second.state.apps.sheets.at, A);
  assert.equal(second.state.apps.mail.at, B);
  assert.equal(second.state.apps.slides, undefined);
  assert.equal(second.state.passes['fuzz:pdf'].at, A);
});

test('nothing is carried on a doubt, and each reason is said', () => {
  const git = { isAncestor: () => true, changed: () => ['README.md'] };
  const ask = (old, o = {}) => carryOver(old, { head: B, tree: CLEAN_TREE, git, appsForPath, apps: VERIFY_APPS, ...o });
  assert.deepEqual(ask(null), { state: null, why: 'no run before', lines: [] });
  assert.equal(ask({}).state, null);
  const refused = [
    [ask({ ...full(), head: B, tree: 'edited' }), /on this commit with other changes/],
    [ask({ ...full(), tree: 'edited' }), /ran on a tree with changes beyond its commit/],
    [ask(full(), { tree: 'edited' }), /this tree has changes beyond its commit/],
    [ask(full(), { git: { ...git, isAncestor: () => false } }), /does not come after it/],
    [ask(full(), { git: { ...git, changed: () => null } }), /git could not list what changed/],
    [ask(full(), { git: { ...git, changed: () => ['tools/gate.js'] } }), /run by every part that had passed/],
  ];
  for (const [r, why] of refused) {
    assert.equal(r.state, null);
    assert.match(r.lines.join('\n'), why);
    assert.match(r.lines[0], /^Nothing is carried from the run on (aaaaaaaa|bbbbbbbb): /);
  }
  // The ancestor is asked of the two commits in order, and the files between them.
  const asked = [];
  const ok = ask(full(), { git: { isAncestor: (a, b) => { asked.push(['ancestor', a, b]); return true; }, changed: (a, b) => { asked.push(['changed', a, b]); return ['apps/desktop/renderer/apps/mail/list.js']; } } });
  assert.deepEqual(asked, [['ancestor', A, B], ['changed', A, B]]);
  assert.equal(ok.state.head, B);
  assert.equal(ok.state.apps.mail, undefined);
  assert.match(ok.lines[0], /^Carried from the run on aaaaaaaa \(1 file changed since, none of them run by these\): /);
  assert.match(ok.lines[1], /^Passed there and run again here, since a file they run changed: .*\bmail\b/);
});

test('the record says which parts were not run on the commit checked', () => {
  const state = over(['apps/desktop/renderer/apps/slides/ink.js']).state;
  // Run again on B: slides, the machine lane and the passes the change could touch.
  state.apps.slides = { ok: true, passed: 2, of: 2, failed: [] };
  state.machine = { ok: true, passed: 5, of: 5, failed: [], scope: VERIFY_APPS.slice() };
  for (const name of PASSES.filter((n) => !ENGINE_ONLY.includes(n))) state.passes[name] = { ok: true, counts: { passed: 3, of: 3 } };
  const note = carriedNote(state);
  assert.match(note, /^Not every part was run on the commit checked\./);
  assert.match(note, new RegExp(`on ${A}, the window checks of shell, .* and locker, the performance budgets, .* and the files made to do harm\\.`));
  assert.doesNotMatch(note, /slides|the engine tests|clipboard/);
  for (const dash of [0x2014, 0x2013]) assert.ok(!note.includes(String.fromCharCode(dash)), 'the note is written with no dashes');
  // Everything run on the commit: nothing to say.
  assert.equal(carriedNote({ ...full(), head: A }), '');
  assert.equal(carriedNote({ head: A, passes: {}, apps: {}, machine: null }), '');
  // Two earlier commits, each with its own.
  const two = over(['apps/desktop/renderer/apps/mail/list.js'], { ...state, tree: CLEAN_TREE }, C).state;
  assert.match(carriedNote(two), new RegExp(`on ${A}, .*; on ${B}, the window checks of slides\\.`));
  // In the note's Gate section, under the commit.
  const section = gateSection({ smoke: { passed: 21, of: 21 } }, { version: '9.9.9', date: '1 January 2030', platform: 'win32', arch: 'x64', commit: B, carried: note });
  assert.match(section, new RegExp(`Commit checked: ${B}\\n\\nNot every part was run on the commit checked\\.`));
  assert.doesNotMatch(gateSection({ smoke: { passed: 21, of: 21 } }, { version: '9.9.9', date: '1 January 2030', commit: B }), /Not every part/);
});

test('the packs the window checks add are named by their signed envelopes', (t) => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-gate-packs-key-'));
  t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
  assert.equal(packsKey(null, ['lang-ur']), null);
  assert.equal(packsKey('', ['lang-ur']), null);
  fs.writeFileSync(path.join(dir, 'lang-ur.pack.json'), '{"manifest":{"files":{"a":{"sha512":"one"}}}}');
  fs.writeFileSync(path.join(dir, 'lang-ar.pack.json'), '{"manifest":{"files":{}}}');
  const first = packsKey(dir, ['lang-ur', 'lang-ar']);
  assert.match(first, /^[0-9a-f]{64}$/);
  assert.equal(packsKey(dir, ['lang-ar', 'lang-ur']), first, 'the order they are named in is nothing');
  assert.notEqual(packsKey(dir, ['lang-ur', 'lang-ar', 'lang-hi']), first, 'a pack that is not there is not the same as one not asked for');
  fs.writeFileSync(path.join(dir, 'lang-ur.pack.json'), '{"manifest":{"files":{"a":{"sha512":"two"}}}}');
  assert.notEqual(packsKey(dir, ['lang-ur', 'lang-ar']), first, 'a pack built again is another pack');
});

test('the passes taken to open no window import nothing of the windows\' own', async () => {
  const perf = fs.readdirSync(path.join(ROOT, 'tests')).filter((f) => f.endsWith('.perf.js')).map((f) => `tests/${f}`);
  assert.ok(perf.length > 0);
  const entries = { 'test:perf': perf, ...Object.fromEntries(FUZZERS.map(([name, tool]) => [name, [tool]])) };
  assert.deepEqual(sorted(Object.keys(entries)), sorted(ENGINE_ONLY));
  // What is installed from outside is left out; the project's own packages are walked.
  const outside = {
    name: 'outside',
    setup(build) {
      build.onResolve({ filter: /^[^./]/ }, (a) => {
        if (path.isAbsolute(a.path) || a.path.startsWith('@rutba/')) return undefined;
        return { path: a.path, external: true };
      });
    },
  };
  for (const [name, files] of Object.entries(entries)) {
    const result = await esbuild.build({
      absWorkingDir: ROOT,
      entryPoints: files,
      bundle: true,
      write: false,
      metafile: true,
      outdir: path.join(ROOT, '.gate-carry-out'),
      format: 'esm',
      splitting: true,
      platform: 'node',
      target: ['node20'],
      loader: { '.js': 'jsx', '.png': 'file', '.jpg': 'file', '.svg': 'file', '.woff2': 'file', '.ttf': 'file', '.wasm': 'file', '.txt': 'text', '.node': 'file' },
      jsx: 'automatic',
      logLevel: 'silent',
      plugins: [outside],
    });
    const inputs = Object.keys(result.metafile.inputs).filter((f) => !f.includes(':') && !f.startsWith('('));
    assert.ok(inputs.length >= files.length, `${name}: its files were walked`);
    assert.ok(inputs.some((f) => f.startsWith('packages/')), `${name}: the walk went into the project's own packages (${inputs.length} files)`);
    // An import made while running (await import of a path written out) is walked too: the fuzzers reach the document service so.
    if (name === 'fuzz:open' || name === 'fuzz:grammar') assert.ok(inputs.includes('apps/desktop/main/documents.js'), `${name}: the walk follows an import made while running`);
    assert.deepEqual(inputs.filter(isWindowsOwn), [], `${name} has come to import a window's own file: take it out of ENGINE_ONLY (tools/gate-carry.js), or cut the import`);
    for (const f of files) assert.doesNotMatch(fs.readFileSync(path.join(ROOT, f), 'utf8'), /apps[\\/'", ]+desktop[\\/'", ]+renderer|main[\\/'", ]+verify-/, `${f} names a window's own file`);
  }
});
