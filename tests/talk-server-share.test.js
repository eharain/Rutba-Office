// Who else may hold meetings on a Talk Server of one's own
// (apps/desktop/renderer/apps/talk/server-share.js, in Talk's settings),
// drawn without a window: the four ways, the people chosen ticked among the
// owner's contacts, somebody chosen who is no longer a contact, and that
// somebody with no server is shown nothing of it.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import esbuild from 'esbuild';
import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';

const here = path.dirname(fileURLToPath(import.meta.url));

/** The screens bundled for node, React shared with this test. */
let loaded = null;
function screens() {
  loaded ||= (async () => {
    const out = await esbuild.build({
      stdin: {
        contents: "export * from './apps/desktop/renderer/apps/talk/server-share.js'; export { TalkSettings } from './apps/desktop/renderer/apps/talk/settings.js';",
        resolveDir: path.join(here, '..'),
        loader: 'js',
      },
      bundle: true,
      format: 'esm',
      platform: 'node',
      write: false,
      logLevel: 'silent',
      loader: { '.js': 'jsx', '.css': 'empty' },
      jsx: 'automatic',
      external: ['react', 'react-dom', 'react/jsx-runtime'],
    });
    const file = path.join(here, '..', 'node_modules', `.talk-server-share-${process.pid}.mjs`);
    fs.writeFileSync(file, out.outputFiles[0].text);
    try { return await import(`file://${file.replace(/\\/g, '/')}`); } finally { fs.rmSync(file, { force: true }); }
  })();
  return loaded;
}

const textOf = (html) => html.replace(/<[^>]*>/g, ' ').replace(/&#x27;/g, "'").replace(/\s+/g, ' ');
const count = (html, what) => html.split(what).length - 1;
const draw = (Component, props) => renderToStaticMarkup(React.createElement(Component, props));
const noop = () => {};

const BEN = { id: 'u-ben', name: 'Ben Ortiz', email: 'ben@example.com', presence: 'online' };
const CARA = { id: 'u-cara', name: 'Cara Lund', email: 'cara@example.com', presence: 'offline' };
const DEV = { id: 'u-dev', name: '', email: 'dev@example.com', presence: 'offline' };
const STATE = { phase: 'online', me: { id: 'u-a', name: 'Amira Khan' }, contacts: [BEN, CARA, DEV], blocked: [] };
const OWN = { id: 'srv_1', name: 'Study PC', state: 'active', answering: true, canCarry: true, why: null, who: 'me', named: [] };
const shell = { talk: { serverMine: async () => ({ error: 'offline' }), serverShare: async () => ({ error: 'offline' }) } };
const props = (servers, more = {}) => ({ shell, state: STATE, toast: noop, mine: servers ? { plan: {}, servers, knows: true } : null, ...more });

test('the four ways are named, each with what it means', async () => {
  const { WHO, whoLabel, whoNote } = await screens();
  assert.deepEqual(WHO, ['me', 'contacts', 'named', 'org'], 'the four the Talk service takes, and no other');
  assert.deepEqual(WHO.map(whoLabel), ['Only me', 'My contacts', 'The people I choose', 'Everybody in my organisation']);
  assert.equal(whoLabel('something-new'), 'Only me');
  assert.equal(whoNote('me'), 'Nobody else can hold a meeting on it.');
  assert.equal(whoNote('contacts'), 'Whoever is a contact of yours at the moment they start a meeting.');
  assert.equal(whoNote('org'), 'Everybody in the organisation your Rutba account belongs to. If it belongs to none, nobody else.');
  assert.equal(whoNote('named', 0), 'Nobody is chosen yet. Tick the contacts who may hold meetings on it.');
  assert.ok(whoNote('named', 1).startsWith('1 person chosen. They are not told'));
  assert.ok(whoNote('named', 3).startsWith('3 people chosen. They are not told'));
  assert.ok(whoNote('named', 3).endsWith('does not say whose.'));
});

test('the people chosen are the owner\'s contacts ticked, and whoever is chosen and no longer a contact', async () => {
  const { namedList } = await screens();
  const { people, gone } = namedList(['u-cara', 'u-zed', 'u-dev'], [BEN, CARA, DEV]);
  assert.deepEqual(people, [
    { id: 'u-ben', name: 'Ben Ortiz', email: 'ben@example.com', on: false },
    { id: 'u-cara', name: 'Cara Lund', email: 'cara@example.com', on: true },
    { id: 'u-dev', name: 'dev@example.com', email: 'dev@example.com', on: true },
  ]);
  assert.deepEqual(gone, ['u-zed'], 'an id with no name to show');
  assert.deepEqual(namedList(null, null), { people: [], gone: [] });
  assert.deepEqual(namedList(undefined, [BEN]).people.map((p) => p.on), [false]);
});

test('the list is sent whole: as it should stand with one person put in or taken out', async () => {
  const { namedWith, NAMED_MOST } = await screens();
  assert.deepEqual(namedWith([], 'u-ben', true), ['u-ben']);
  assert.deepEqual(namedWith(['u-ben', 'u-cara'], 'u-ben', false), ['u-cara']);
  assert.deepEqual(namedWith(['u-ben'], 'u-ben', true), ['u-ben'], 'nobody twice');
  assert.deepEqual(namedWith(['u-ben', 'u-ben', 'u-cara'], 'u-dev', true), ['u-ben', 'u-cara', 'u-dev']);
  assert.deepEqual(namedWith(['u-ben'], 'u-zed', false), ['u-ben'], 'taking out somebody who is not in changes nothing');
  assert.deepEqual(namedWith(null, 'u-ben', false), []);
  // The Talk service takes no more than five hundred: one more is not put in.
  const full = Array.from({ length: NAMED_MOST }, (_, i) => `u-${i}`);
  assert.equal(NAMED_MOST, 500);
  assert.equal(namedWith(full, 'u-new', true).length, 500);
  assert.equal(namedWith(full, 'u-new', true).includes('u-new'), false);
  assert.equal(namedWith(full, 'u-7', false).length, 499);
});

test('somebody with no Talk Server of their own is shown nothing of this', async () => {
  const { ServerShare, TalkSettings } = await screens();
  // Nor is anybody where the Talk service does not say it knows a meeting held on a server, whatever servers they have.
  for (const mine of [null, { error: 'no-service' }, { plan: {}, servers: [], knows: true }, { plan: {}, servers: null, knows: true }, { plan: {}, servers: [OWN] }, { plan: {}, servers: [OWN], knows: false }, { plan: {}, servers: [OWN], knows: 'yes' }]) {
    assert.equal(draw(ServerShare, { shell, state: STATE, toast: noop, mine }), '', JSON.stringify(mine));
  }
  // Talk's settings are then what they always were.
  const plain = draw(TalkSettings, { shell, state: STATE, toast: noop, onClose: noop });
  assert.doesNotMatch(plain, /tk-server/);
  assert.ok(!textOf(plain).includes('Talk Server'));
});

test('an owner is shown each server with who may hold meetings on it, and no list of people until they choose that way', async () => {
  const { ServerShare } = await screens();
  const html = draw(ServerShare, props([OWN, { ...OWN, id: 'srv_2', name: 'Office box', who: 'contacts' }]));
  const text = textOf(html);
  assert.ok(text.includes('Your Talk Servers'));
  assert.ok(text.includes('Who may hold meetings on Study PC'));
  assert.ok(text.includes('Who may hold meetings on Office box'));
  assert.equal(count(html, 'tk-server-who'), 2);
  assert.equal(count(html, '<option'), 8, 'the four ways, for each');
  assert.match(html, /data-server="srv_1" data-who="me"/);
  assert.match(html, /data-server="srv_2" data-who="contacts"/);
  assert.match(html, /<option value="me" selected="">Only me<\/option>/);
  assert.match(html, /<option value="contacts" selected="">My contacts<\/option>/);
  assert.ok(text.includes('Nobody else can hold a meeting on it.'));
  assert.ok(text.includes('Whoever is a contact of yours at the moment they start a meeting.'));
  assert.doesNotMatch(html, /type="checkbox"/, 'nobody to tick');
  assert.doesNotMatch(html, /type="text"|<input(?![^>]*type="checkbox")/, 'and no place to type anything: no address, no id');
  // A way the app does not know is shown as the safest one.
  assert.match(draw(ServerShare, props([{ ...OWN, who: 'everyone' }])), /data-who="me"/);
  // A server with no name is still named.
  assert.ok(textOf(draw(ServerShare, props([{ ...OWN, name: '' }]))).includes('Who may hold meetings on your Talk Server'));
});

test('choosing people: the owner\'s contacts, the chosen ones ticked, and somebody no longer a contact with the way to take them out', async () => {
  const { ServerShare } = await screens();
  const html = draw(ServerShare, props([{ ...OWN, who: 'named', named: ['u-cara', 'u-zed'] }]));
  const text = textOf(html);
  assert.equal(count(html, 'type="checkbox"'), 3, 'one for each contact');
  assert.match(html, /data-pick="u-cara"><input type="checkbox"[^>]*checked=""/);
  assert.doesNotMatch(html, /data-pick="u-ben"><input type="checkbox"[^>]*checked=""/);
  assert.ok(text.includes('Ben Ortiz') && text.includes('Cara Lund') && text.includes('dev@example.com'));
  assert.ok(text.includes('2 people chosen. They are not told: where they set a meeting up, Talk offers to hold it on a Talk Server and does not say whose.'));
  assert.match(html, /data-gone="u-zed"/);
  assert.ok(text.includes('Somebody no longer in your contacts'));
  assert.ok(text.includes('Take out'));
  assert.ok(!text.includes('u-zed'), 'an id is never shown as if it were a name');
  // Nobody chosen yet, and nobody to choose from.
  const empty = textOf(draw(ServerShare, props([{ ...OWN, who: 'named' }])));
  assert.ok(empty.includes('Nobody is chosen yet. Tick the contacts who may hold meetings on it.'));
  const alone = draw(ServerShare, props([{ ...OWN, who: 'named' }], { state: { ...STATE, contacts: [] } }));
  assert.doesNotMatch(alone, /type="checkbox"/);
  assert.ok(textOf(alone).includes('No contacts yet') || /tk-none/.test(alone));
});

test('a server that cannot hold a meeting now says why beside it, and with Talk not connected nothing can be changed', async () => {
  const { ServerShare } = await screens();
  const text = textOf(draw(ServerShare, props([{ ...OWN, canCarry: false, answering: false, why: 'not-answering' }])));
  assert.ok(text.includes('Study PC is not answering. Check that it is running and connected.'));
  assert.doesNotMatch(draw(ServerShare, props([OWN])), /tk-server-why/);

  const off = draw(ServerShare, props([{ ...OWN, who: 'named', named: ['u-ben', 'u-zed'] }], { state: { ...STATE, phase: 'offline' } }));
  assert.match(off, /<select[^>]*disabled=""/);
  assert.equal(count(off, 'disabled=""'), 5, 'the choice, the three contacts and the button');
});

test('Talk\'s settings show it after the people blocked, for an owner', async () => {
  const { TalkSettings } = await screens();
  const html = draw(TalkSettings, { shell, state: STATE, toast: noop, onClose: noop, servers: { plan: {}, servers: [OWN], knows: true } });
  assert.ok(html.indexOf('Blocked people') < html.indexOf('Your Talk Servers'));
  assert.match(html, /tk-server-who/);
});
