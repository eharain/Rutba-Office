// The apps registered through the loader (apps/desktop/main/app-packs.js,
// docs/APP-PACKS.md 5.1): what each pack is, which packs a run has (every
// app built in, or the signed packs of a folder), how each registers, and
// the services that reach another app coping with it being absent. The
// packs here are made and signed for the test with a key thrown away with
// its folder.
import test from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { METHODS, EVENTS, PACKS, NAMESPACE_PACK, packIpc, coreMethods } from '../packages/office-shell/src/contract.js';
import { markRemoval, pendingRemovals } from '../packages/office-shell/src/electron/pack-install.js';
import { fileAssociations } from '../packages/office-formats/src/registry.js';
import { makePacks } from '../tools/make-packs.mjs';
import { makeKeyPair } from '../tools/make-pack-key.mjs';
import { describePack, planPacks, registerPacks, packView, readPackKeys, BUILT_IN } from '../apps/desktop/main/app-packs.js';
import { TOOL_PACK, toolOffered, toolsOf, LOADED_KEY } from '../apps/desktop/main/pack-tools.js';

const VERSION = '9.9.9';
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-app-packs-'));
test.after(() => fs.rmSync(tmp, { recursive: true, force: true }));
const keys = makeKeyPair();
const privateKey = crypto.createPrivateKey(keys.privatePem);

/** A built tree of packs as the build describes them, each with a main module of its own. */
async function built(ids, { name = 'built', patch = {} } = {}) {
  const tree = path.join(tmp, `${name}-tree`);
  for (const id of ids) {
    const d = describePack(id);
    const folder = path.join(tree, id);
    fs.mkdirSync(path.join(folder, 'main'), { recursive: true });
    fs.writeFileSync(path.join(folder, 'main', 'index.js'), 'export function register(core) { return core.builtIn ? core.builtIn(core) : {}; }\n');
    const { id: _id, tiles, fileTypes, ...rest } = d;
    fs.writeFileSync(path.join(folder, 'pack.json'), JSON.stringify({ ...rest, ...(patch[id] || {}), main: 'main/index.js', tiles: [], fileTypes }));
  }
  const out = path.join(tmp, name);
  await makePacks({ tree, out, version: VERSION, privateKey });
  return out;
}

test('each pack says what it is: its channels are the contract\'s for it, its file types and AI tools its own', () => {
  for (const { id, name } of PACKS) {
    const d = describePack(id);
    assert.equal(d.name, name);
    assert.deepEqual(d.ipc, packIpc(id), id);
    for (const ns of Object.keys(d.ipc)) assert.equal(NAMESPACE_PACK[ns], id, `${id} declares ${ns}`);
    assert.deepEqual(d.fileTypes.map((t) => t.ext), fileAssociations({ apps: [id] }).map((t) => t.ext), id);
    for (const tool of d.tools) assert.equal(TOOL_PACK[tool], id, `${id} offers ${tool}`);
    assert.ok(BUILT_IN[id], `${id} has a register module in the core`);
  }
  // PDF's own icon wherever .pdf is claimed.
  assert.equal(describePack('pdf').fileTypes.find((t) => t.ext === 'pdf').icon, 'filetypes/pdf.ico');
  assert.deepEqual(describePack('word').tiles, [{ app: 'word', icon: 'tiles/word.ico' }]);
  assert.deepEqual(describePack('ocr').tiles, [], 'Text recognition has no tile of its own');
  assert.throws(() => describePack('nonsense'), /no app pack called nonsense/);
  // Every namespace of the contract is the core's or one pack's.
  const owned = new Set(Object.keys(NAMESPACE_PACK));
  for (const ns of Object.keys(METHODS)) assert.ok(owned.has(ns) !== Boolean(coreMethods()[ns]), ns);
});

test('with no packs folder every app is built in, and the run offers the whole contract', () => {
  const plan = planPacks({ dir: null, version: VERSION, tilesDir: path.join(tmp, 'apps') });
  assert.equal(plan.source, 'built-in');
  assert.deepEqual(plan.ids(), PACKS.map((p) => p.id));
  const offer = plan.offer();
  assert.deepEqual(Object.keys(offer.methods).sort(), Object.keys(METHODS).sort());
  assert.deepEqual([...offer.events].sort(), [...EVENTS].sort());
  assert.deepEqual(offer.refused, []);
  assert.equal(plan.packs.find((p) => p.id === 'mail').tiles[0].icon, path.join(tmp, 'apps', 'mail.ico'));
});

test('with a packs folder only the signed packs found are loaded, and an absent app offers nothing', async () => {
  const dir = await built(['contacts', 'mail', 'word'], { name: 'three' });
  const plan = planPacks({ dir, keys: readPackKeysFrom([keys.publicPem]), version: VERSION });
  assert.equal(plan.source, 'files');
  assert.deepEqual(plan.ids().sort(), ['contacts', 'mail', 'word']);
  assert.equal(plan.has('calendar'), false);
  const { methods, events, packs, refused } = plan.offer();
  assert.ok(methods.mail && methods.discover && methods.contacts, 'the loaded packs\' namespaces');
  assert.ok(methods.doc && methods.win && methods.packs, 'the core\'s');
  for (const ns of ['calendar', 'talk', 'base', 'pdf', 'ocr', 'ai', 'live', 'present']) assert.equal(methods[ns], undefined, ns);
  assert.ok(!events.includes('talk:state'));
  assert.deepEqual(packs.map((p) => p.id).sort(), ['contacts', 'mail', 'word']);
  assert.deepEqual(refused, []);
  // The manifest carries the app's rows of the registry.
  assert.deepEqual(plan.packs.find((p) => p.id === 'word').registry.extensions, describePack('word').registry.extensions);
});

/** The keys as main.js reads them from the core's build file. */
function readPackKeysFrom(pems) {
  const file = path.join(tmp, `keys-${crypto.randomUUID()}.json`);
  fs.writeFileSync(file, JSON.stringify({ keys: pems }));
  return readPackKeys(file);
}

test('a pack is refused for a namespace the core or another app has, a wrong key, or no key at all', async () => {
  // Talk's pack claiming Mail's namespace, and Calendar's claiming the core's documents.
  const dir = await built(['mail', 'talk', 'calendar'], { name: 'clash', patch: { talk: { ipc: { ...packIpc('talk'), mail: ['send'] } }, calendar: { ipc: { ...packIpc('calendar'), doc: ['open'] } } } });
  const plan = planPacks({ dir, keys: readPackKeysFrom([keys.publicPem]), version: VERSION });
  assert.deepEqual(plan.ids(), ['mail']);
  assert.match(plan.refused.find((r) => r.id === 'talk').reason, /declares the mail namespace, which is the Mail pack's/);
  assert.match(plan.refused.find((r) => r.id === 'calendar').reason, /declares the doc namespace, which is the core's/);

  const other = makeKeyPair();
  const wrong = planPacks({ dir, keys: readPackKeysFrom([other.publicPem]), version: VERSION });
  assert.deepEqual(wrong.ids(), []);
  assert.equal(wrong.refused.length, 3);

  const none = planPacks({ dir, keys: [], version: VERSION });
  assert.deepEqual(none.ids(), []);
  assert.match(none.refused[0].reason, /holds no key to check its app packs with/);
  assert.deepEqual(readPackKeys(path.join(tmp, 'no-such-file.json')), []);
});

test('a pack written down to go is gone before the loader runs', async () => {
  const dir = await built(['contacts', 'word'], { name: 'going' });
  markRemoval(dir, 'word');
  const plan = planPacks({ dir, keys: readPackKeysFrom([keys.publicPem]), version: VERSION });
  assert.deepEqual(plan.ids(), ['contacts']);
  assert.equal(fs.existsSync(path.join(dir, 'word.pack.json')), false);
  assert.equal(fs.existsSync(path.join(dir, 'word.asar')), false);
  assert.deepEqual(pendingRemovals(dir), []);
});

/** A plan of packs whose main modules are written here, for registerPacks. */
function planOf(mains) {
  const dir = path.join(tmp, `mains-${crypto.randomUUID()}`);
  fs.mkdirSync(dir);
  const packs = Object.entries(mains).map(([id, body]) => {
    const main = path.join(dir, `${id}.mjs`);
    fs.writeFileSync(main, body);
    return { ...describePack(id), main, source: 'files' };
  });
  const plan = planPacks({ dir: null, version: VERSION });
  plan.packs.splice(0, plan.packs.length, ...packs);
  return plan;
}

test('each pack registers through its main module, and serves only what its manifest declares', async () => {
  const plan = planOf({
    contacts: 'export function register({ pack, stores }) { return { contacts: { who: () => pack.id + stores.tag } }; }\n',
    mail: 'export function register() { return { mail: { accounts: () => [] }, contacts: { list: () => [] } }; }\n',
    calendar: 'export function register() { throw new Error("its store would not open"); }\n',
    word: 'export const nothing = 1;\n',
  });
  const services = { doc: { core: true } };
  const recorded = [];
  await registerPacks(plan, { stores: { tag: '!' } }, { services, record: (kind) => recorded.push(kind) });
  assert.equal(services.contacts.who(), 'contacts!');
  assert.equal(services.mail, undefined, 'a pack that serves more than it declared is not served at all');
  assert.equal(services.calendar, undefined);
  assert.ok(services.doc.core, 'the core\'s services stay');
  assert.deepEqual(plan.ids(), ['contacts']);
  assert.match(plan.refused.find((r) => r.id === 'mail').reason, /^The Mail pack could not be started: it serves contacts, which its manifest does not declare\. Install it again with Add or remove apps\.$/);
  assert.match(plan.refused.find((r) => r.id === 'calendar').reason, /its store would not open/);
  assert.match(plan.refused.find((r) => r.id === 'word').reason, /no register function/);
  assert.deepEqual(recorded.sort(), ['pack calendar', 'pack mail', 'pack word']);
  // What the windows are told follows: a pack that failed is refused there too.
  const offer = plan.offer();
  assert.equal(offer.methods.mail, undefined);
  assert.deepEqual(offer.refused.map((r) => r.id).sort(), ['calendar', 'mail', 'word']);
});

test('one pack reaches another only through the view, which answers nothing while it is absent', () => {
  const plan = planOf({ contacts: '' });
  const services = { contacts: { list: () => ['card'] }, mail: { accounts: () => ['left over'] } };
  const view = packView({ plan, services });
  assert.equal(view.has('contacts'), true);
  assert.equal(view.has('mail'), false);
  assert.deepEqual(view.service('contacts').list(), ['card']);
  assert.equal(view.service('mail'), undefined, 'Mail is absent, whatever the services hold');
  assert.equal(view.service('discover'), undefined, 'and so is its namespace for other mail clients');
  assert.equal(view.service('doc'), undefined, 'the core\'s services are not reached through the view');
});

/** The context the core hands a pack, with stores in a folder of their own. */
function context(view) {
  const map = new Map();
  const dir = fs.mkdtempSync(path.join(tmp, 'stores-'));
  return {
    stores: {
      dir,
      settings: { get: (k, d) => (map.has(k) ? map.get(k) : d), set: (k, v) => map.set(k, v), delete: (k) => map.delete(k), all: () => Object.fromEntries(map) },
      secrets: { available: () => true, get: () => null, set: () => {}, delete: () => {}, keys: () => [] },
    },
    broadcast: () => {},
    holdBlob: () => ({ url: '' }),
    oauth: null,
    packs: view,
  };
}

/** A view whose packs are switched on and off by the test, over services it holds. */
function switchable() {
  const on = new Set();
  const services = {};
  return { on, services, view: { has: (id) => on.has(id), service: (ns) => (on.has(NAMESPACE_PACK[ns]) ? services[ns] : undefined) } };
}

test('Contacts with Mail absent suggests its cards alone, and Mail\'s people once Mail is there', async () => {
  const { on, services, view } = switchable();
  const { contacts } = (await BUILT_IN.contacts()).register(context(view));
  contacts.save({ contact: { names: { given: 'Sam', family: 'Kaur' }, emails: [{ value: 'sam@example.com' }] } });
  assert.deepEqual(contacts.suggest({ query: 'sa' }).map((s) => s.email), ['sam@example.com']);
  services.mail = { accounts: () => [{ id: 'a' }], people: () => ({ rows: [{ name: 'Sara Ali', address: 'sara@example.com' }] }) };
  assert.deepEqual(contacts.suggest({ query: 'sa' }).map((s) => s.email), ['sam@example.com'], 'Mail\'s service left over, but Mail not installed');
  on.add('mail');
  assert.deepEqual(contacts.suggest({ query: 'sa' }).map((s) => s.email), ['sam@example.com', 'sara@example.com']);
  // A Mail that throws is no people, never a broken address line.
  services.mail = { accounts: () => { throw new Error('closed'); } };
  assert.deepEqual(contacts.suggest({ query: 'sa' }).map((s) => s.email), ['sam@example.com']);
});

test('Calendar with Mail absent registers, with no accounts to answer from', async () => {
  const { on, services, view } = switchable();
  const { mailFor, register } = await BUILT_IN.calendar();
  assert.ok(register(context(view)).calendar);
  const mail = mailFor(view);
  assert.deepEqual(mail.accounts(), []);
  services.mail = { accounts: () => [{ id: 'a', email: 'me@example.com' }] };
  on.add('mail');
  assert.deepEqual(mail.accounts().map((a) => a.id), ['a']);
});

test('Talk with Mail absent registers, and refuses in plain words to send an invitation from Mail', async () => {
  const { on, services, view } = switchable();
  const { mailFor, register } = await BUILT_IN.talk();
  assert.ok(register(context(view)).talk);
  const mail = mailFor(view);
  assert.deepEqual(mail.accounts(), []);
  await assert.rejects(mail.send({ accountId: 'a' }), /^Error: Mail is not installed on this computer, so the invitation cannot be sent from your own address\. Add Mail with Add or remove apps, or let Rutba send it\.$/);
  const sent = [];
  services.mail = { accounts: () => [{ id: 'a' }], send: async (args) => { sent.push(args); return { messageId: 'x' }; } };
  on.add('mail');
  assert.deepEqual(await mail.send({ accountId: 'a' }), { messageId: 'x' });
  assert.equal(sent.length, 1);
});

test('Mail with Contacts absent registers, and its check for a known contact matches no one', async () => {
  process.env.RUTBA_OFFICE_VERIFY_APPS ||= '1';
  const { on, services, view } = switchable();
  const { contactsFor, register } = await BUILT_IN.mail();
  const out = register(context(view));
  assert.deepEqual(Object.keys(out).sort(), Object.keys(describePack('mail').ipc).sort());
  const contacts = contactsFor(view);
  assert.deepEqual(contacts.list({ query: 'sam@example.com' }), []);
  services.contacts = { list: () => [{ id: 1, emails: [{ value: 'sam@example.com' }] }] };
  on.add('contacts');
  assert.equal(contacts.list({ query: 'sam@example.com' }).length, 1);
});

test('the AI tools offered are those of the packs loaded, and none without AI Tools', () => {
  const names = Object.keys(TOOL_PACK);
  const settings = (loaded) => (loaded ? { [LOADED_KEY]: loaded } : {});
  assert.ok(names.every((n) => toolOffered(n, settings(undefined))), 'a run that has not said: every tool, as before');
  assert.ok(!names.some((n) => toolOffered(n, settings(['word', 'sheets']))), 'without AI Tools, none');
  const some = names.filter((n) => toolOffered(n, settings(['ai', 'word'])));
  assert.ok(some.length > 0);
  for (const n of some) assert.ok(['ai', 'word'].includes(TOOL_PACK[n]), n);
  assert.ok(toolsOf('word', names).includes('create_document'));
  assert.ok(!toolsOf('word', names).includes('create_workbook'));
});
