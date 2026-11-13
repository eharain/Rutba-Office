// The speech engine add-on, loaded (apps/desktop/main/engine-host.js): only
// the contract's context handed to it, only the contract's speech namespace
// taken from it, and a start that keeps the Windows voice when it is slow,
// throws or answers anything else. The engines here are stand-ins written
// for the test; no model is run.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { loadEngine, engineContext, heldUtilityProcess } from '../apps/desktop/main/engine-host.js';
import { closeAddonsFolder, isHeld } from '../packages/office-shell/src/electron/pack-resolve.js';
import { CONTRACT_FILE } from '../tools/make-addon-vectors.mjs';

const contract = JSON.parse(fs.readFileSync(CONTRACT_FILE, 'utf8'));
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-engine-host-'));
// The engines here sit in a folder held as an add-ons folder, as the run's is (pack-resolve.js).
const addonsDir = path.join(tmp, 'addons');
fs.mkdirSync(addonsDir, { recursive: true });
closeAddonsFolder(addonsDir);
test.after(() => fs.rmSync(tmp, { recursive: true, force: true }));
let n = 0;

/** A registry holding one engine whose main is the source given. */
function withEngine(source, { main = 'main/index.mjs', listed = true } = {}) {
  const dir = path.join(addonsDir, `e${(n += 1)}`);
  fs.mkdirSync(path.join(dir, 'main'), { recursive: true });
  fs.writeFileSync(path.join(dir, ...main.split('/')), source);
  const manifest = { id: 'speech-engine', name: 'Test engine', kind: 'engine', main, files: listed ? { [main]: { size: 1, sha512: 'x' } } : {} };
  const voice = { id: 'voice-a', kind: 'voice', dir: path.join(tmp, 'voice'), manifest: { id: 'voice-a', kind: 'voice' } };
  const engine = { id: 'speech-engine', kind: 'engine', dir, manifest };
  const all = [engine, voice];
  return { list: () => all, find: (name) => all.find((a) => a.id === name) || null, engine: () => engine };
}

const GOOD = `
export function register(context) {
  globalThis.__seen = { keys: Object.keys(context).sort(), addons: Object.keys(context.addons).sort(), electron: Object.keys(context.electron || {}), voices: context.addons.list().map((a) => a.id) };
  return { speech: {
    voices: async () => context.addons.list().filter((a) => a.kind === 'voice').map((a) => ({ id: a.id })),
    say: async ({ text }) => ({ samples: new Float32Array(text.length), rate: 24000 }),
    cancel: async () => ({}),
    status: async () => ({ running: false }),
    secret: () => 'not offered',
    stop: () => { globalThis.__stopped = true; },
  } };
}`;

test('no engine added: nothing loaded, nothing to say', async () => {
  const r = await loadEngine({ registry: { engine: () => null }, contract });
  assert.deepEqual(r.namespaces, {});
  assert.equal(r.problem, null);
});

test('an engine is handed only the contract\'s context and serves exactly the contract\'s calls', async () => {
  const fakeElectron = { utilityProcess: { fork: () => null }, BrowserWindow: {}, shell: {} };
  const r = await loadEngine({ registry: withEngine(GOOD), contract, electron: fakeElectron });
  assert.equal(r.problem, null);
  assert.deepEqual(Object.keys(r.namespaces), ['speech']);
  assert.deepEqual(Object.keys(r.namespaces.speech).sort(), [...contract.speech.calls].sort());
  assert.deepEqual(r.methods, { speech: contract.speech.calls });
  assert.deepEqual(globalThis.__seen, { keys: ['addons', 'electron'], addons: ['find', 'list'], electron: ['utilityProcess'], voices: ['speech-engine', 'voice-a'] });
  const said = await r.namespaces.speech.say({ text: 'Hello' });
  assert.equal(said.samples.length, 5);
  assert.deepEqual(await r.namespaces.speech.voices(), [{ id: 'voice-a' }]);
  r.stop();
  assert.equal(globalThis.__stopped, true);
});

test('the context is frozen and holds nothing the contract does not list', () => {
  const ctx = engineContext({ registry: { list: () => [], find: () => null }, electron: { utilityProcess: 1, app: 2 }, names: ['addons.list', 'electron.utilityProcess', 'electron.app2', 'fs.readFile'] });
  assert.deepEqual(Object.keys(ctx).sort(), ['addons', 'electron']);
  assert.deepEqual(Object.keys(ctx.addons), ['list']);
  assert.ok(Object.isFrozen(ctx) && Object.isFrozen(ctx.addons));
});

test('an engine answering another namespace, or missing a call, is not used', async () => {
  const other = await loadEngine({ registry: withEngine('export function register() { return { speech: { voices() {}, say() {}, cancel() {}, status() {} }, mail: {} }; }'), contract });
  assert.deepEqual(other.namespaces, {});
  assert.match(other.problem, /not speech alone/);
  const missing = await loadEngine({ registry: withEngine('export function register() { return { speech: { voices() {}, say() {} } }; }'), contract });
  assert.match(missing.problem, /does not answer cancel, status/);
});

test('an engine that throws, never answers, or names no listed main is not used, and the start goes on', async () => {
  const throws = await loadEngine({ registry: withEngine('export function register() { throw new Error("no model"); }'), contract });
  assert.match(throws.problem, /could not start: no model/);
  const slow = await loadEngine({ registry: withEngine('export function register() { return new Promise(() => {}); }'), contract, budgetMs: 50 });
  assert.match(slow.problem, /longer than/);
  const unlisted = await loadEngine({ registry: withEngine(GOOD, { listed: false }), contract });
  assert.match(unlisted.problem, /names no main module/);
  const outside = await loadEngine({ registry: withEngine(GOOD, { main: '../x.mjs' }), contract });
  assert.match(outside.problem, /names no main module/);
});

test('a process the engine forks imports the resolution hooks first, with the resolution of this run, whatever it asks', () => {
  const calls = [];
  const raw = { fork: (...a) => { calls.push(a); return { pid: 1 }; }, other: () => 'not offered' };
  const env = { PATH: 'p', RUTBA_PACK_RESOLUTION: '{"closed":["D:\\addons"]}' };
  const held = heldUtilityProcess(raw, { hooks: 'file:///hooks.js', env });
  assert.deepEqual(Object.keys(held), ['fork']);
  assert.ok(Object.isFrozen(held));
  held.fork('w.mjs', [], { serviceName: 'speech', stdio: 'ignore' });
  assert.deepEqual(calls[0], ['w.mjs', [], { serviceName: 'speech', stdio: 'ignore', env, execArgv: ['--import', 'file:///hooks.js'] }]);
  // Its own environment and arguments are kept, the resolution put back and the hooks still first.
  held.fork('w.mjs', ['a'], { env: { ONLY: '1', RUTBA_PACK_RESOLUTION: '{}' }, execArgv: ['--import', 'file:///mine.js'] });
  assert.deepEqual(calls[1][2], { env: { ONLY: '1', RUTBA_PACK_RESOLUTION: env.RUTBA_PACK_RESOLUTION }, execArgv: ['--import', 'file:///hooks.js', '--import', 'file:///mine.js'] });
  // Through the context, as the engine is handed it.
  const ctx = engineContext({ registry: { list: () => [], find: () => null }, electron: { utilityProcess: raw }, names: ['electron.utilityProcess'] });
  assert.deepEqual(Object.keys(ctx.electron.utilityProcess), ['fork']);
  assert.notEqual(ctx.electron.utilityProcess, raw);
});

test('an engine is started only from a held add-ons folder, which a junction does not get round', { skip: process.platform !== 'win32' && 'junctions are for Windows' }, async () => {
  // Outside any held folder: not imported at all.
  const loose = path.join(tmp, 'loose', 'engine');
  fs.mkdirSync(path.join(loose, 'main'), { recursive: true });
  fs.writeFileSync(path.join(loose, 'main', 'index.mjs'), GOOD);
  const engine = { id: 'speech-engine', kind: 'engine', dir: loose, manifest: { id: 'speech-engine', name: 'Test engine', kind: 'engine', main: 'main/index.mjs', files: { 'main/index.mjs': { size: 1, sha512: 'x' } } } };
  let imported = false;
  const r = await loadEngine({ registry: { list: () => [engine], find: () => engine, engine: () => engine }, contract, importer: async () => { imported = true; return {}; } });
  assert.match(r.problem, /not in an add-ons folder whose imports are held/);
  assert.equal(imported, false);
  // A folder held by the path it was named by, which is a junction: Node names its modules by the real path, and those are held too.
  const real = path.join(tmp, 'real-addons');
  fs.mkdirSync(path.join(real, 'e-x'), { recursive: true });
  fs.writeFileSync(path.join(real, 'e-x', 'm.mjs'), '');
  const via = path.join(tmp, 'via-junction');
  fs.symlinkSync(real, via, 'junction');
  assert.equal(isHeld(path.join(real, 'e-x', 'm.mjs')), false);
  closeAddonsFolder(via);
  assert.equal(isHeld(path.join(via, 'e-x', 'm.mjs')), true);
  assert.equal(isHeld(path.join(real, 'e-x', 'm.mjs')), true);
});
