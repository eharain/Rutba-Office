// The IPC contract per pack (docs/APP-PACKS.md 5.1.1): the contract's
// namespaces divided by owner, the preload's bridge built from the list one
// run offers (asked for once, synchronously, as a window starts), and the
// window's client answering a call into a pack that is not installed in
// plain words, without reaching the main process.
import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import esbuild from 'esbuild';
import {
  METHODS, EVENTS, PACKS, NAMESPACE_PACK, EVENT_PACK, CONTRACT_CHANNEL,
  coreMethods, coreEvents, methodsFor, eventsFor, packIpc, packEvents, channels,
} from '../packages/office-shell/src/contract.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const ids = PACKS.map((p) => p.id);

test('every namespace is the core\'s or one pack\'s, and every owner is a pack', () => {
  for (const owner of Object.values(NAMESPACE_PACK)) assert.ok(ids.includes(owner), `${owner} is a pack`);
  for (const owner of Object.values(EVENT_PACK)) assert.ok(ids.includes(owner), `${owner} is a pack`);
  for (const ns of Object.keys(NAMESPACE_PACK)) assert.ok(METHODS[ns], `${ns} is in the contract`);
  for (const e of Object.keys(EVENT_PACK)) assert.ok(EVENTS.includes(e), `${e} is in the contract`);
  const core = coreMethods();
  for (const ns of Object.keys(METHODS)) assert.equal(Boolean(core[ns]), !NAMESPACE_PACK[ns], ns);
  // What the core keeps whatever is installed: the shared document service, printing, the DAV accounts, recording, Add or remove apps.
  for (const ns of ['app', 'win', 'fs', 'doc', 'print', 'dav', 'capture', 'oauth', 'packs', 'store']) assert.ok(core[ns], ns);
  assert.ok(coreEvents().includes('mail:oauth'), 'signing in is the core\'s');
  assert.ok(!coreEvents().includes('mail:new'));
});

test('a run offers the core\'s methods and those of the packs it loaded, no others', () => {
  const none = methodsFor([]);
  assert.deepEqual(Object.keys(none).sort(), Object.keys(coreMethods()).sort());
  assert.ok(!none.mail && !none.calendar && !none.base);
  const some = methodsFor(['mail', 'contacts']);
  assert.deepEqual(some.mail, METHODS.mail);
  assert.deepEqual(some.discover, METHODS.discover, 'Mail\'s other namespace comes with it');
  assert.deepEqual(some.contacts, METHODS.contacts);
  assert.ok(!some.calendar && !some.talk && !some.pdf);
  const every = methodsFor(ids);
  assert.deepEqual(Object.keys(every).sort(), Object.keys(METHODS).sort(), 'every pack loaded is the whole contract');
  assert.equal(channels(every).length, channels().length);
  assert.ok(!channels(none).some((c) => c.startsWith('mail:')));

  assert.deepEqual(packIpc('mail'), { mail: METHODS.mail, discover: METHODS.discover });
  assert.deepEqual(packIpc('word'), {}, 'Documents has no namespace of its own: the document service is the core\'s');
  assert.deepEqual(packIpc('slides'), { present: METHODS.present });
  assert.ok(packEvents('mail').includes('mail:new') && !packEvents('mail').includes('mail:oauth'));
  assert.ok(!eventsFor([]).includes('talk:state') && eventsFor(['talk']).includes('talk:state'));
  assert.deepEqual(eventsFor(ids), EVENTS);
});

/** The preload, bundled as the build bundles it, with a stand-in for Electron that answers its one synchronous request. */
async function preloadWith(answer) {
  const built = await esbuild.build({
    entryPoints: [path.join(ROOT, 'packages/office-shell/src/electron/preload.js')],
    bundle: true,
    write: false,
    format: 'cjs',
    platform: 'node',
    external: ['electron'],
    logLevel: 'silent',
  });
  const exposed = {};
  const asked = [];
  const invoked = [];
  const electron = {
    contextBridge: { exposeInMainWorld: (name, api) => { exposed[name] = api; } },
    ipcRenderer: {
      sendSync: (channel) => {
        asked.push(channel);
        if (answer instanceof Error) throw answer;
        return answer;
      },
      invoke: (channel, payload) => { invoked.push([channel, payload]); return Promise.resolve('answered'); },
      on: () => {},
      removeListener: () => {},
    },
    webUtils: { getPathForFile: () => null },
  };
  const sandbox = { require: (m) => (m === 'electron' ? electron : null), module: { exports: {} }, process: { platform: 'win32', arch: 'x64', versions: {} }, URLSearchParams, globalThis: {} };
  sandbox.exports = sandbox.module.exports;
  vm.runInNewContext(built.outputFiles[0].text, sandbox);
  return { bridge: exposed.rutbaOffice, asked, invoked };
}

test('the preload asks once for this run\'s list and offers exactly that', async () => {
  const methods = methodsFor(['mail']);
  const { bridge, asked, invoked } = await preloadWith({ methods, events: eventsFor(['mail']), packs: [{ id: 'mail', name: 'Mail', windows: null }], refused: [] });
  assert.deepEqual(asked, [CONTRACT_CHANNEL], 'one synchronous request');
  assert.ok(bridge.mail && bridge.discover && bridge.app && bridge.doc);
  assert.equal(bridge.calendar, undefined, 'a method of an absent pack does not exist in the window');
  assert.equal(bridge.talk, undefined);
  assert.deepEqual(Object.keys(bridge.__contract.methods).sort(), Object.keys(methods).sort());
  assert.deepEqual(bridge.__contract.packs.map((p) => p.id), ['mail']);
  await bridge.mail.accounts();
  assert.equal(invoked[0][0], 'rutba-office/mail:accounts');
  assert.throws(() => bridge.on('calendar:changed', () => {}), /unknown event/);
  assert.equal(typeof bridge.on('mail:new', () => {}), 'function');
});

test('a preload whose request goes unanswered offers the core alone, never every pack', async () => {
  const { bridge } = await preloadWith(new Error('no handler'));
  assert.ok(bridge.app && bridge.win && bridge.fs && bridge.packs);
  for (const ns of Object.keys(NAMESPACE_PACK)) assert.equal(bridge[ns], undefined, ns);
});

test('the window\'s client answers a call into an absent pack in plain words, without reaching the main process', async () => {
  const calls = [];
  const fake = { app: {}, mail: {}, discover: {}, on: (event) => { calls.push(['on', event]); return () => {}; } };
  const offered = methodsFor(['mail']);
  for (const [ns, names] of Object.entries(offered)) {
    fake[ns] = {};
    for (const n of names) fake[ns][n] = (payload) => { calls.push([`${ns}.${n}`, payload]); return Promise.resolve(`${ns}.${n}`); };
  }
  fake.__contract = { methods: offered, events: eventsFor(['mail']), packs: [{ id: 'mail', name: 'Mail', windows: null }], refused: [{ id: 'calendar', name: 'Calendar', reason: 'The Calendar pack is version 1 and this Rutba Office is 2.' }] };
  globalThis.rutbaOffice = fake;
  try {
    const client = await import(`${pathToFileURL(path.join(ROOT, 'packages/office-shell/src/client.js')).href}?with-mail`);
    assert.equal(await client.mail.accounts(), 'mail.accounts');
    const err = await client.calendar.calendars().then(() => null, (e) => e);
    assert.ok(err, 'the call is refused');
    assert.equal(err.message, 'Calendar is not installed on this computer. Add it with Add or remove apps.');
    assert.equal(err.code, 'absent');
    assert.equal(err.pack, 'calendar');
    assert.equal(err.operation, 'calendar.calendars');
    const talk = await client.talk.state().then(() => null, (e) => e);
    assert.equal(talk.message, 'Talk is not installed on this computer. Add it with Add or remove apps.');
    assert.ok(!calls.some(([c]) => c.startsWith('calendar') || c.startsWith('talk')), 'nothing crossed for an absent pack');
    // Listening to an absent pack's event is listening to nothing; an unknown event is still an error.
    assert.equal(typeof client.on('calendar:changed', () => {}), 'function');
    assert.ok(!calls.some(([c, e]) => c === 'on' && e === 'calendar:changed'));
    client.on('mail:new', () => {});
    assert.ok(calls.some(([c, e]) => c === 'on' && e === 'mail:new'));
    assert.throws(() => client.on('nothing:here', () => {}), /unknown event/);
    assert.deepEqual(client.installedPacks(), ['mail']);
    assert.equal(client.isInstalled('mail'), true);
    assert.equal(client.isInstalled('calendar'), false);
    assert.equal(client.refusedPacks()[0].id, 'calendar');
  } finally {
    delete globalThis.rutbaOffice;
  }
});

test('outside the desktop shell the client stands in the whole contract', async () => {
  const client = await import(`${pathToFileURL(path.join(ROOT, 'packages/office-shell/src/client.js')).href}?bare`);
  assert.deepEqual(client.installedPacks(), ids);
  assert.equal(typeof client.calendar.calendars, 'function');
  await assert.rejects(client.calendar.calendars(), /not attached/);
});
