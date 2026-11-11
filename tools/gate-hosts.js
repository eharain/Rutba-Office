// Where the gate's passes run: this computer, or a Linux machine of the
// project's own, reached by a runner script given in the environment (PM,
// 11 Nov 2026: the engine tests and the fuzz passes go to the lab machine,
// the windows stay on Windows).
//
//   RUTBA_GATE_VM       the runner: bash <runner> <clone> <label> <command...>
//                       snapshots the tree as it stands, runs the command
//                       there, and brings back the run's .gate/logs to
//                       .gate/vm/logs. No path or address of it is kept here.
//   RUTBA_GATE_VM_HOST  the machine's name, said in this computer's own output and
//                       failure list (.gate/, never committed). The record, which
//                       is public, says only what the machine is (labLabel).
//   --local             runs them on this computer instead, for this run.
//
// With the runner set, a machine that cannot be reached stops the gate and
// says so: nothing falls back to this computer unless --local asks for it.
// The engine tests that touch Windows itself (its paths, clipboard, fonts,
// registry, shells) are listed in tools/engine-hosts.json and stay here, and
// so do those that import, however far down, a module that touches it.

import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

/** The fuzz passes: name, tool, what it goes through, rounds. */
export const FUZZERS = [['fuzz:open', 'tools/fuzz-open.js', 'damaged Office files', 300], ['fuzz:pdf', 'tools/fuzz-pdf.js', 'damaged PDFs', 300], ['fuzz:smime', 'tools/fuzz-smime.js', 'damaged signed mail', 2000], ['fuzz:grammar', 'tools/fuzz-grammar.js', 'files made to do harm', 600]];

/** The passes that may run on the lab machine: nothing in them opens a window. */
export const VM_PASSES = ['test', ...FUZZERS.map(([name]) => name)];

/** The runner's own exit codes: the machine's fault, not a test's. */
/** The bash that runs the runner: RUTBA_GATE_VM_BASH, or the first on the path. */
export const vmBash = (env = process.env) => String(env.RUTBA_GATE_VM_BASH || '').trim() || 'bash';

export const VM_FAULTS = {
  3: 'the runner could not snapshot the tree',
  4: 'the lab machine could not be reached (ssh or the push failed)',
  5: 'the lab machine could not check out the tree',
  6: 'npm ci failed on the lab machine',
};

/** The file of engine tests that stay on Windows, and those checked and found portable. */
export const HOSTS_FILE = path.join('tools', 'engine-hosts.json');

/**
 * What in a test file touches the operating system itself: its platform,
 * Windows paths, the registry, the clipboard, its fonts folder, its shells.
 * A file with one of these is listed in engine-hosts.json, one way or the other.
 */
export const OS_MARKERS = /process\.platform|\bwin32\b|reg(?:\.exe)?['"`\s,]+\[?['"`]?(?:query|add|delete)|\bHK(?:CU|LM|EY_)|clipboard|Windows[\\/]+Fonts|fc-list|systemFonts|powershell|cmd\.exe|%?LOCALAPPDATA|\bAPPDATA\b|\b[A-Z]:\\\\/;

/**
 * The same for a module a test imports, read as code: its comments left out,
 * and the clipboard only as a call on it. The word in a comment or a message
 * ("There is nothing on the clipboard to paste") is no use of it.
 */
export const MODULE_MARKERS = /process\.platform|\bwin32\b|reg(?:\.exe)?['"`\s,]+\[?['"`]?(?:query|add|delete)|\bHK(?:CU|LM|EY_)|\bclipboard\s*\.\s*(?:read|write|clear|has|available)\w*|Windows[\\/]+Fonts|fc-list|systemFonts|powershell|cmd\.exe|%?LOCALAPPDATA|\bAPPDATA\b|\b[A-Z]:\\\\/;
export function moduleTouches(text) {
  const code = String(text).replace(/\/\*[\s\S]*?\*\//g, '').split('\n').filter((l) => !/^\s*\/\//.test(l)).join('\n');
  return MODULE_MARKERS.test(code);
}

/** Where this run sends the lab machine's passes: { on, runner, host, said }. */
export function vmConfig(env = process.env, argv = []) {
  const runner = String(env.RUTBA_GATE_VM || '').trim();
  const host = String(env.RUTBA_GATE_VM_HOST || '').trim() || 'the lab machine';
  if (!runner) return { on: false, runner: null, host: null, said: null };
  if (argv.includes('--local')) return { on: false, runner, host, said: `--local: the engine tests and fuzz passes run on this computer, not on ${host}` };
  return { on: true, runner, host, said: `the engine tests and fuzz passes run on ${host}, beside the rest` };
}

/** The command that runs `command` on the lab machine, as [program, args]. */
export function vmCommand(cfg, root, label, command, env = process.env) {
  return [vmBash(env), [cfg.runner, String(root).split(path.sep).join('/'), label, ...command]];
}

/** The test files `npm test` names, from package.json's own globs, in a steady order. */
export function engineGlobs(root) {
  const script = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8')).scripts.test;
  return [...String(script).matchAll(/"([^"]+)"/g)].map((m) => m[1]);
}

export function engineFiles(root) {
  const files = [];
  for (const glob of engineGlobs(root)) {
    const dir = path.posix.dirname(glob);
    const base = path.posix.basename(glob);
    // Only the shape npm test uses: a folder, then *<ending>. Another shape would be read wrongly, so it is refused.
    if (!/^\*[^*?[\]{}]+$/.test(base) || /[*?[\]{}]/.test(dir)) throw new Error(`npm test names ${glob}, which the gate cannot list: it reads <folder>/*<ending> only`);
    let names = [];
    try { names = fs.readdirSync(path.join(root, dir)); } catch { /* none */ }
    for (const f of names.sort()) if (f.endsWith(base.slice(1))) files.push(`${dir}/${f}`);
  }
  return files;
}

/** tools/engine-hosts.json: { windowsOnly: { file: why }, portable: { file: why } }. */
export function readHosts(root) {
  const file = path.join(root, HOSTS_FILE);
  const j = JSON.parse(fs.readFileSync(file, 'utf8'));
  return { windowsOnly: j.windowsOnly || {}, portable: j.portable || {} };
}

/** A relative import, re-export, dynamic import or require of a module of the project's own. */
const IMPORT = /(?:^|[^\w$.])(?:import\s*(?:[\w$*{}\s,]+from\s*)?|import\s*\(\s*|require\s*\(\s*|export\s+[\w$*{}\s,]+from\s*)['"`]((?:\.{1,2}\/|@rutba\/)[^'"`]+)['"`]/g;

/** The workspace's own packages by name (`@rutba/office-shell`): their folder and exports, from packages/<x>/package.json. */
function workspace(root, cache) {
  if (cache.has('\0packages')) return cache.get('\0packages');
  const out = new Map();
  let dirs = [];
  try { dirs = fs.readdirSync(path.join(root, 'packages')); } catch { /* none */ }
  for (const d of dirs) {
    try {
      const pkg = JSON.parse(fs.readFileSync(path.join(root, 'packages', d, 'package.json'), 'utf8'));
      if (pkg.name) out.set(pkg.name, { dir: path.join(root, 'packages', d), exports: pkg.exports, main: pkg.main });
    } catch { /* not a package */ }
  }
  cache.set('\0packages', out);
  return out;
}

/** `@rutba/<name>` or `@rutba/<name>/<sub>` as a file: by the package's exports, else its folder. */
function resolvePackage(root, spec, cache) {
  const m = /^(@rutba\/[^/]+)(?:\/(.+))?$/.exec(spec);
  const pkg = m && workspace(root, cache).get(m[1]);
  if (!pkg) return null;
  const key = m[2] ? `./${m[2]}` : '.';
  let target = typeof pkg.exports === 'string' ? (key === '.' ? pkg.exports : null) : pkg.exports?.[key];
  // A pattern, as Node reads one: `./lib/*` to `./src/lib/*.js`, the star standing for the same text in both.
  if (!target && pkg.exports && typeof pkg.exports === 'object') {
    for (const [k, v] of Object.entries(pkg.exports)) {
      const star = k.indexOf('*');
      if (star < 0 || typeof v !== 'string') continue;
      const [pre, post] = [k.slice(0, star), k.slice(star + 1)];
      if (key.startsWith(pre) && key.endsWith(post) && key.length >= pre.length + post.length) { target = v.split('*').join(key.slice(pre.length, key.length - post.length)); break; }
    }
  }
  if (target && typeof target === 'object') target = target.import || target.default || target.node || null;
  if (!target && key === '.') target = pkg.main || './index.js';
  return path.resolve(pkg.dir, target || `./${m[2]}`);
}

/**
 * The first module a test file reaches through its imports that touches the
 * operating system (OS_MARKERS), or null. A test of a module that branches on
 * Windows must run on Windows, though its own text says nothing of it (review
 * of 11 Nov: 31 files went to the lab machine through mail.js, documents.js,
 * profiles.js and others). The test file's own text is not read here: that is
 * engine-hosts.json's to say. `cache` is shared across calls.
 */
export function osReach(root, file, cache = new Map()) {
  // Per module, once: what it imports, and whether it touches the system itself.
  const node = (f) => {
    if (cache.has(f)) return cache.get(f);
    let text = '';
    try { text = fs.readFileSync(f, 'utf8'); } catch { /* none */ }
    const imports = [];
    for (const m of text.matchAll(IMPORT)) {
      const base = m[1].startsWith('@rutba/') ? resolvePackage(root, m[1], cache) : path.resolve(path.dirname(f), m[1]);
      if (!base) continue;
      for (const c of [base, `${base}.js`, `${base}.mjs`, `${base}.cjs`, path.join(base, 'index.js')]) {
        let isFile = false;
        try { isFile = fs.statSync(c).isFile(); } catch { /* not this one */ }
        if (isFile) { if (/\.[mc]?js$/.test(c)) imports.push(c); break; }
      }
    }
    const n = { imports, touches: moduleTouches(text) };
    cache.set(f, n);
    return n;
  };
  // Breadth first from the test file's imports, so the module named is the nearest.
  const start = path.join(root, file);
  const seen = new Set([start]);
  const queue = [...node(start).imports];
  while (queue.length) {
    const f = queue.shift();
    if (seen.has(f)) continue;
    seen.add(f);
    const n = node(f);
    if (n.touches) return path.relative(root, f).split(path.sep).join('/');
    queue.push(...n.imports);
  }
  return null;
}

/**
 * The test files for each host: those that touch Windows stay here, the rest
 * go to the lab machine. `reach(file)` names a module the file reaches that
 * touches Windows (osReach), which keeps it here too.
 */
export function splitEngine(files, hosts, reach = () => null) {
  const listed = new Set(Object.keys(hosts.windowsOnly || {}));
  const here = new Set(files.filter((f) => listed.has(f) || reach(f)));
  return { vm: files.filter((f) => !here.has(f)), local: files.filter((f) => here.has(f)) };
}

/** Test files that touch the operating system and are on neither list: each must be put on one. */
export function unlisted(root, files, hosts) {
  const known = new Set([...Object.keys(hosts.windowsOnly || {}), ...Object.keys(hosts.portable || {})]);
  return files.filter((f) => !known.has(f) && OS_MARKERS.test(fs.readFileSync(path.join(root, f), 'utf8')));
}

/** Whether a pass's output carries its counts: TAP totals, or a fuzzer's last line. */
const counted = (output) => /^# tests \d+$/m.test(output) || /^\d+ (?:damaged files|damaged PDFs|messages): /m.test(output);

/**
 * The lab machine's fault in a run, or null: one of the runner's own exit
 * codes with nothing counted, or the runner saying it could not go on.
 */
export function vmFault(code, output = '') {
  const text = String(output).replace(/\r/g, '');
  if (VM_FAULTS[code] && !counted(text)) return VM_FAULTS[code];
  if (code !== 0 && /^\[vm\] (?:could not snapshot|push to the VM failed|npm ci failed)/m.test(text)) return VM_FAULTS[/npm ci/.test(text) ? 6 : /snapshot/.test(text) ? 3 : 4];
  return null;
}

/** The snapshot the runner made of the tree, from its first line. */
export function snapshotOf(output) {
  // `[vm] office at <head>, snapshot <s>`, or with the route it took: `[vm] office via <address> at <head>, snapshot <s>`.
  const m = /^\[vm\] \S+(?: via \S+)? at \w+, snapshot ([0-9a-f]{7,40})\s*$/m.exec(String(output).replace(/\r/g, ''));
  return m ? m[1] : null;
}

/**
 * The tree as it stands, committed or not, as git names it: the runner
 * snapshots the same way, so the two can be compared.
 */
export function worktreeHash(root) {
  const tmp = path.join(os.tmpdir(), `rutba-gate-index-${process.pid}-${Date.now()}`);
  const gitDir = execFileSync('git', ['rev-parse', '--absolute-git-dir'], { cwd: root, encoding: 'utf8' }).trim();
  try {
    fs.copyFileSync(path.join(gitDir, 'index'), tmp);
    const env = { ...process.env, GIT_INDEX_FILE: tmp };
    execFileSync('git', ['add', '-A', '.'], { cwd: root, env, stdio: 'ignore' });
    return execFileSync('git', ['write-tree'], { cwd: root, env, encoding: 'utf8' }).trim();
  } finally {
    fs.rmSync(tmp, { force: true });
  }
}

const unquote = (v) => {
  const s = String(v || '').trim();
  const q = /^'(.*)'$/.exec(s) || /^"(.*)"$/.exec(s);
  return (q ? q[1].replace(/''/g, "'") : s).replace(/\\\\/g, '\\');
};

/** A test's location as a file the suite names, and its line. */
function placeOf(location, files) {
  const loc = String(location || '').replace(/\\/g, '/');
  const m = /^(.*?):(\d+):(\d+)$/.exec(loc);
  const at = m ? m[1] : loc;
  const file = files.find((f) => at === f || at.endsWith(`/${f}`)) || null;
  return { file, line: m ? Number(m[2]) : null };
}

/**
 * A node --test TAP stream read into what it found: each test that failed
 * itself (not a parent failing for its children), with its file and line,
 * each that passed, and the totals. `files` are the suite's test files,
 * relative to the clone, so a failure is placed whichever machine ran it.
 *   failures  [{ name, file, at, detail }]  name is `<file> > <test> > <subtest>`
 *   passed    [name]
 *   ran       the files that reported a test
 *   totals    { tests, passed, failed, skipped } or null when the run never got to them
 */
export function parseTap(text, files = []) {
  const lines = String(text || '').replace(/\u001b\[[0-9;]*m/g, '').replace(/\r/g, '').split('\n');
  const stack = [];
  const failures = [];
  const passed = [];
  const ran = new Set();
  const sum = (key) => lines.reduce((n, l) => { const m = new RegExp(`^# ${key} (\\d+)$`).exec(l); return m ? n + Number(m[1]) : n; }, 0);
  const unescape = (s) => s.replace(/\\#/g, '#').replace(/\\\\/g, '\\');
  for (let i = 0; i < lines.length; i++) {
    const sub = /^( *)# Subtest: (.*)$/.exec(lines[i]);
    if (sub) { const d = Math.floor(sub[1].length / 4); stack.length = d; stack[d] = unescape(sub[2]); continue; }
    const r = /^( *)(not ok|ok) \d+ - (.*?)(?: # (SKIP|TODO)\b.*)?$/.exec(lines[i]);
    if (!r) continue;
    const depth = Math.floor(r[1].length / 4);
    const title = unescape(r[3]);
    const titles = [...stack.slice(0, depth), title];
    stack.length = depth;
    // Its YAML block: the keys at its own depth, and the first line of a block value.
    const yaml = {};
    if (lines[i + 1]?.trim() === '---') {
      const own = ' '.repeat(depth * 4 + 2);
      let j = i + 2;
      for (; j < lines.length && lines[j].trim() !== '...'; j++) {
        const kv = new RegExp(`^${own}([A-Za-z_]+): ?(.*)$`).exec(lines[j]);
        if (!kv) continue;
        yaml[kv[1]] = /^[|>]-?$/.test(kv[2]) ? (lines[j + 1] || '').trim() : unquote(kv[2]);
      }
      i = j;
    }
    if (r[4]) continue;
    let { file, line } = placeOf(yaml.location, files);
    // A file that did not run to its end is reported as a test named by the file.
    // Compared with forward slashes: on Windows a crashed file is titled tests\x.test.js.
    const plainTitle = title.replace(/\\/g, '/');
    const fileLevel = depth === 0 && files.find((f) => plainTitle === f || plainTitle.endsWith(`/${f}`) || f.endsWith(`/${plainTitle}`));
    if (fileLevel) file = fileLevel;
    const name = fileLevel ? `${file}: the file ran to its end` : `${file || 'a test file'} > ${titles.join(' > ')}`;
    if (file) ran.add(file);
    if (r[2] === 'ok') { passed.push(name); continue; }
    // A parent that failed only because a child did: the child is the failure.
    if (yaml.failureType === 'subtestsFailed') continue;
    failures.push({ name, file, at: file && line ? `${file}:${line}` : file, detail: yaml.error || yaml.failureType || null });
  }
  // A file that ran to its end, by the same name its failure would have: so a pass later closes that entry.
  const fileFailed = new Set(failures.filter((f) => f.name.endsWith(': the file ran to its end')).map((f) => f.file));
  for (const f of ran) if (!fileFailed.has(f)) passed.push(`${f}: the file ran to its end`);
  const tests = sum('tests');
  return { failures, passed, ran: [...ran], totals: lines.some((l) => /^# tests \d+$/.test(l)) ? { tests, passed: sum('pass'), failed: sum('fail'), skipped: sum('skipped') } : null };
}

/** The reporter the gate runs beside TAP: a line a test, with its file (tools/gate-reporter.js). */
export const REPORTER = './tools/gate-reporter.js';

/** node --test's arguments for the gate: TAP to the log, the gate's own lines to `events`, a relative path. */
export function testArgs(files, { events, alone = false }) {
  return ['--test', '--test-reporter=tap', '--test-reporter-destination=stdout', `--test-reporter=${REPORTER}`, `--test-reporter-destination=${events}`, ...(alone ? ['--test-concurrency=1'] : []), ...files];
}

/**
 * The gate reporter's lines read into what TAP cannot say: the file of each
 * test that passed as well as of each that failed. The same shape as
 * parseTap's, without the totals.
 */
export function readEvents(text, files = []) {
  const stacks = new Map();
  const failures = [];
  const passed = [];
  const ran = new Set();
  const fileFailed = new Set();
  for (const raw of String(text || '').split(/\r?\n/)) {
    if (!raw.trim()) continue;
    let e;
    try { e = JSON.parse(raw); } catch { continue; }
    const file = e.file ? placeOf(e.file, files).file : null;
    const key = file || e.file || '?';
    const stack = stacks.get(key) || [];
    stacks.set(key, stack);
    const depth = Number(e.nesting) || 0;
    if (e.event === 'start') { stack.length = depth; stack[depth] = e.name; continue; }
    if (e.event !== 'pass' && e.event !== 'fail') continue;
    const titles = [...stack.slice(0, depth), e.name];
    stack.length = depth;
    if (file) ran.add(file);
    if (e.skip || e.todo) continue;
    const plain = String(e.name || '').replace(/\\/g, '/');
    const fileLevel = depth === 0 && file && (plain === file || plain.endsWith(`/${file}`) || file.endsWith(`/${plain}`));
    const name = fileLevel ? `${file}: the file ran to its end` : `${file || 'a test file'} > ${titles.join(' > ')}`;
    if (e.event === 'pass') { passed.push(name); continue; }
    if (e.failureType === 'subtestsFailed') continue;
    if (fileLevel) fileFailed.add(file);
    failures.push({ name, file, at: file && e.line ? `${file}:${e.line}` : file, detail: e.error || e.failureType || null });
  }
  for (const f of ran) if (!fileFailed.has(f)) passed.push(`${f}: the file ran to its end`);
  return { failures, passed, ran: [...ran] };
}

/**
 * A test run as the gate reads it: the totals from its TAP, the tests and
 * their files from the reporter's lines. Without those lines (an older run,
 * or a reporter that did not start) it is TAP alone, whose passes name no file.
 */
export function readRun(tapText, eventsText, files = []) {
  const tap = parseTap(tapText, files);
  if (!String(eventsText || '').trim()) return tap;
  const ev = readEvents(eventsText, files);
  // TAP saw failures the lines did not: the lines are short, so TAP's are the failures.
  const failures = ev.failures.length || !tap.failures.length ? ev.failures : tap.failures;
  return { failures, passed: ev.passed, ran: ev.ran, totals: tap.totals };
}

/** The part a test failure is rerun and closed by: its file, or the suite whole. */
export const testPart = (file) => `test#${file || '*'}`;

/**
 * What the engine passes found, in the failure list's own shape (as
 * gate-ledger.js's runFindings answers): app `engine`, a test's part its
 * file, a fuzz pass's its name. `tests` is [{ host, tap, again, whole }]
 * (`code` its exit, `again` the TAP of its run again, `againCode` that one's exit);
 * `fuzz` is [{ name, host, ok, counted, again }]. `wholeSuite` when every
 * test file ran: only then is the suite whole (`test#*`) clean.
 */
export function engineFindings({ tests = [], fuzz = [], wholeSuite = false, home = () => null, fuzzHome = null } = {}) {
  // A pass on the host a file belongs to now is said as `engine@home|...` too: an entry whose file moved hosts since it failed (listed as
  // Windows-only to fix a Linux failure, say) is proved by a pass where it runs now (gate-ledger.js recordRun; review of 11 Nov).
  const fileOf = (name) => String(name).split(' > ')[0].replace(/: the file ran to its end$/, '');
  const failures = [];
  const passed = new Set();
  const clean = new Set();
  const keyOf = (part, name) => `engine|${part}|${name}`;
  for (const t of tests) {
    const againPassed = new Set(t.again?.passed || []);
    const failedFiles = new Set();
    if (!t.tap?.totals) {
      const name = `test: the engine tests on ${t.host} ran to their end and counted`;
      failures.push({ key: keyOf('test#*', name), name, id: 'new', app: 'engine', part: 'test#*', lane: t.host, synthetic: true, flaky: false, crash: true });
      continue;
    }
    // An exit that was not a pass, with no test named for it: the run's own failure.
    if (t.code && !t.tap.failures.length) {
      const name = `test: the engine tests on ${t.host} ended with exit ${t.code}, no test named`;
      failures.push({ key: keyOf('test#*', name), name, id: 'new', app: 'engine', part: 'test#*', lane: t.host, synthetic: true, flaky: Boolean(t.again && !t.again.failures.length && t.againCode === 0), crash: true });
    }
    for (const f of t.tap.failures) {
      const part = testPart(f.file);
      failedFiles.add(part);
      failures.push({ key: keyOf(part, f.name), name: f.name, id: 'new', app: 'engine', part, lane: t.host, at: f.at, detail: f.detail, synthetic: false, flaky: againPassed.has(f.name) });
    }
    // Each pass said twice: as itself, and as on its host, for an entry whose failure was on one host (gate-ledger.js recordRun).
    for (const n of t.tap.passed) {
      passed.add(`engine|${n}`);
      passed.add(`engine@${t.host}|${n}`);
      if (home(fileOf(n)) === t.host) passed.add(`engine@home|${n}`);
    }
    for (const file of t.tap.ran) {
      if (failedFiles.has(testPart(file))) continue;
      clean.add(`engine|${testPart(file)}`);
      clean.add(`engine@${t.host}|${testPart(file)}`);
      if (home(file) === t.host) clean.add(`engine@home|${testPart(file)}`);
    }
    // What failed only when run again: an entry of its own.
    const first = new Set(t.tap.failures.map((f) => f.name));
    for (const f of t.again?.failures || []) if (!first.has(f.name)) failures.push({ key: keyOf(testPart(f.file), f.name), name: f.name, id: 'new', app: 'engine', part: testPart(f.file), lane: t.host, at: f.at, detail: f.detail, synthetic: false, flaky: false, again: true });
  }
  if (wholeSuite && tests.length && tests.every((t) => t.tap?.totals && !t.tap.failures.length && !t.code)) {
    clean.add('engine|test#*');
    for (const t of tests) clean.add(`engine@${t.host}|test#*`);
    if (tests.every((t) => t.tap.ran.every((f) => home(f) === t.host))) clean.add('engine@home|test#*');
  }
  for (const z of fuzz) {
    const part = `${z.name}#*`;
    if (z.ok) {
      clean.add(`engine|${part}`);
      clean.add(`engine@${z.host}|${part}`);
      if (fuzzHome && z.host === fuzzHome) clean.add(`engine@home|${part}`);
      continue;
    }
    const name = `${z.name}: every file read or refused in a sentence`;
    failures.push({ key: keyOf(part, name), name, id: 'new', app: 'engine', part, lane: z.host, synthetic: true, flaky: Boolean(z.again?.ok), detail: z.counted ? `${z.counted.of - z.counted.passed} of ${z.counted.of} faulted` : 'it ended without its count' });
  }
  return { failures, passed, clean };
}

/** Two sets of findings as one, for the failure list. */
export function mergeFindings(a, b) {
  return { failures: [...a.failures, ...b.failures], passed: new Set([...a.passed, ...b.passed]), clean: new Set([...a.clean, ...b.clean]) };
}

/**
 * The parts asked for by id that are the engine's: `test#<file>`, `test#*`,
 * `fuzz:<name>#*`. Answers { tests, fuzz, strange }: the files asked for
 * (or `*` for the suite), the fuzz passes, and the ids that name nothing.
 */
export function engineParts(ids, files) {
  const tests = new Set();
  const fuzz = new Set();
  const strange = [];
  for (const id of ids) {
    const [pass, rest] = [id.slice(0, id.indexOf('#')), id.slice(id.indexOf('#') + 1)];
    if (pass === 'test' && (rest === '*' || files.includes(rest))) tests.add(rest);
    else if (FUZZERS.some(([n]) => n === pass) && rest === '*') fuzz.add(pass);
    else strange.push(id);
  }
  return { tests: tests.has('*') ? ['*'] : [...tests], fuzz: [...fuzz], strange };
}

/** Whether an id is the engine's rather than a window part's. */
export const isEngineId = (id) => /^(?:test|fuzz:[\w-]+)#/.test(id);

/**
 * What the lab machine ran, as each pass's state: the host and its operating
 * system as the machine itself said, and the tree it ran.
 */
export function hostRecord(host, result, extra = {}) {
  return { host, os: result?.os || null, platform: result?.platform || null, ...extra };
}

/**
 * The machine as the public record names it: what it is, never its name, user
 * or address (PM, 11 Nov 2026: the record goes into a public note and README,
 * and the machine is a private project's). `Linux lab machine (Rocky Linux 9.5, 12 cores)`.
 */
export function labLabel(osName, cpus) {
  const system = String(osName || '').replace(/\s*\([^)]*\)\s*$/, '').replace(/[^\w .+-]/g, '').trim();
  const cores = Number(cpus) > 0 ? `${Number(cpus)} cores` : null;
  const inside = [system || null, cores].filter(Boolean).join(', ');
  return `the Linux lab machine${inside ? ` (${inside})` : ''}`;
}

/** The record's sentence on where the passes ran, or null when all ran here. `host` says only whether it was here. */
export function hostsSentence(groups, { n = (v) => Number(v).toLocaleString('en-GB') } = {}) {
  const away = groups.filter((g) => g.host !== 'local');
  if (!away.length) return null;
  const where = (g) => (g.host === 'local' ? 'on this computer, Windows' : `on ${labLabel(g.osName, g.cpus)}`);
  const tests = groups.filter((g) => g.pass === 'test' && g.tests).map((g) => `${n(g.tests)} engine tests ${where(g)}${g.host === 'local' && away.some((a) => a.pass === 'test') ? ', those that touch Windows itself' : ''}`);
  const fuzzBy = new Map();
  for (const g of groups.filter((x) => x.pass !== 'test')) fuzzBy.set(where(g), [...(fuzzBy.get(where(g)) || []), g.pass]);
  const fuzz = [...fuzzBy].map(([w, names]) => `${names.join(', ')} ${w}`);
  const parts = [...tests, ...fuzz];
  return `Where they ran: ${parts.length > 1 ? `${parts.slice(0, -1).join('; ')}; and ${parts.at(-1)}` : parts[0]}. The windows ran on this computer.`;
}
