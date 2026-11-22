// The Guests dialog of Live Broadcast and the Recorder
// (apps/desktop/renderer/studio/guests.js, Office 2.27.0) drawn without a
// window. What it holds to: the notice that everyone asked is told, before
// anything else; who is here and whether each is on air; no more than three
// asked in; and nothing to ask with while Talk is not connected.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import esbuild from 'esbuild';
import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';

const here = path.dirname(fileURLToPath(import.meta.url));

/** The dialog bundled for node, React shared with this test. */
let loaded = null;
function screens() {
  loaded ||= (async () => {
    const out = await esbuild.build({
      stdin: {
        contents: "export { GuestsDialog, whereText, guestRefusal, sessionName } from './apps/desktop/renderer/studio/guests.js'; export { liveNotice, recordedNotice } from './apps/desktop/renderer/apps/talk/meet-lines.js';",
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
    const file = path.join(here, '..', 'node_modules', `.studio-guests-screen-${process.pid}.mjs`);
    fs.writeFileSync(file, out.outputFiles[0].text);
    try { return await import(`file://${file.replace(/\\/g, '/')}`); } finally { fs.rmSync(file, { force: true }); }
  })();
  return loaded;
}

const textOf = (html) => html.replace(/<[^>]*>/g, ' ').replace(/&#x27;/g, "'").replace(/\s+/g, ' ');
const count = (html, what) => html.split(what).length - 1;
const noop = () => {};
const BEN = { id: 'u-ben', name: 'Ben Ortiz', email: 'ben@example.com', presence: 'online' };
const CARA = { id: 'u-cara', name: 'Cara Lund', email: 'cara@example.com', presence: 'offline' };
const DAN = { id: 'u-dan', name: 'Dan Moyo', email: 'dan@example.com', presence: 'online' };
const EVE = { id: 'u-eve', name: 'Eve Park', email: 'eve@example.com', presence: 'online' };
const TALK = { phase: 'online', me: { id: 'u-a', name: 'Amira Khan' }, contacts: [BEN, CARA, DAN, EVE] };
const here1 = (p, where, more = {}) => ({ id: p.id, name: p.name, muted: false, connected: where !== 'connecting' && where !== 'failed', where, ...more });

async function draw(props) {
  const { GuestsDialog } = await screens();
  return renderToStaticMarkup(React.createElement(GuestsDialog, {
    how: 'live', talk: TALK, holding: false, people: [], onAir: false, busy: false,
    onInvite: noop, onCopy: noop, onMute: noop, onTakeOut: noop, onToTalk: noop, onEnd: noop, onOpenTalk: noop, onClose: noop, ...props,
  }));
}

test('while Talk is not connected there is nothing to ask with, only the way to Talk', async () => {
  const waiting = await draw({ talk: null });
  assert.equal(count(waiting, 'st-guests-invite'), 0);
  assert.equal(count(waiting, 'type="checkbox"'), 0);

  const off = await draw({ talk: { phase: 'signed-out', contacts: [BEN] } });
  assert.match(textOf(off), /Guests come in through Talk\. Open Talk and sign in, then come back here\./);
  assert.equal(count(off, 'st-guests-open-talk'), 1);
  assert.equal(count(off, 'st-guests-invite'), 0, 'no Invite');
  assert.equal(count(off, 'type="checkbox"'), 0, 'and nobody to tick');
});

test('before anyone is asked: what everyone will be told, and the contacts to choose from', async () => {
  const { liveNotice, recordedNotice } = await screens();
  const live = await draw({});
  assert.match(live, /role="note" data-how="live"/);
  assert.ok(textOf(live).includes(`Everyone asked is told this before they join: ${liveNotice()}`), textOf(live));
  assert.match(textOf(live), /Each waits off air until you put them in a scene, and sees and hears your programme\./);
  assert.equal(count(live, 'data-contact='), 4);
  assert.equal(count(live, 'Not in Talk just now'), 1, 'Cara is not in Talk just now, and can still be asked');
  assert.match(live, /st-guests-invite[^>]*disabled/, 'nothing to send until someone is chosen');
  assert.equal(count(live, 'st-guests-foot'), 0, 'no meeting yet: no link, nothing to end');
  assert.equal(count(live, 'st-guests-here'), 0);

  const rec = await draw({ how: 'record' });
  assert.match(rec, /role="note" data-how="record"/);
  assert.ok(textOf(rec).includes(`Everyone asked is told this before they join: ${recordedNotice()}`));
  assert.match(textOf(rec), /Each stays out of the recording until you put them in a scene, and sees and hears what you record\./);
  assert.notEqual(liveNotice(), recordedNotice());

  const alone = await draw({ talk: { ...TALK, contacts: [] } });
  assert.match(textOf(alone), /You have no other Talk contacts to ask\. Add people in Talk first\./);
});

test('who is here, each on air or off it, and only the others left to ask', async () => {
  const html = await draw({ holding: true, onAir: true, people: [here1(BEN, 'on-air'), here1(CARA, 'off-air', { muted: true }), here1(DAN, 'connecting')] });
  assert.match(html, /data-person="u-ben" data-where="on-air"/);
  assert.match(html, /data-person="u-cara" data-where="off-air"/);
  assert.equal(count(html, '<small class="air">'), 1, 'one person is marked as on air');
  const text = textOf(html);
  assert.match(text, /Ben Ortiz On air/);
  assert.match(text, /Cara Lund Off air · Muted/);
  assert.match(text, /Dan Moyo Connecting/);
  // Three are here: nobody more can be asked.
  assert.match(text, /All three seats are taken\./);
  assert.equal(count(html, 'data-contact='), 0);
  assert.equal(count(html, 'st-guests-mute'), 3);
  assert.equal(count(html, 'st-guests-out'), 3);

  const two = await draw({ holding: true, people: [here1(BEN, 'off-air'), here1(CARA, 'off-air')] });
  assert.deepEqual([...two.matchAll(/data-contact="([^"]+)"/g)].map((m) => m[1]), ['u-dan', 'u-eve'], 'those here are not offered again');
});

test('back to Talk as a meeting only while nothing is on air; the link and ending it are there once a meeting is', async () => {
  const quiet = await draw({ holding: true, people: [here1(BEN, 'off-air')] });
  assert.equal(count(quiet, 'st-guests-copy'), 1);
  assert.equal(count(quiet, 'st-guests-end'), 1);
  assert.doesNotMatch(quiet, /st-guests-talk[^>]*disabled/);
  assert.match(quiet, /This window leaves and Talk comes in: the same people, as a meeting\./);

  const live = await draw({ holding: true, onAir: true, people: [here1(BEN, 'on-air')] });
  assert.match(live, /st-guests-talk[^>]*disabled/);
  assert.match(live, /End the broadcast first\./);
  const rec = await draw({ how: 'record', holding: true, onAir: true, people: [here1(BEN, 'on-air')] });
  assert.match(rec, /Stop the recording first\./);
});

test('the words for where a guest is, for a meeting that cannot be held here, and for the meeting\'s name', async () => {
  const { whereText, guestRefusal, sessionName } = await screens();
  const words = ['on-air', 'off-air', 'unseated', 'failed', 'connecting'].map(whereText);
  assert.equal(new Set(words).size, 5, 'each place has its own words');
  assert.ok(words.every((w) => w && !/^[a-z-]+\.[a-z.-]+$/.test(w)), words.join(' | '));
  assert.match(guestRefusal({ error: 'not-live' }), /not set up to be streamed live/);
  assert.match(guestRefusal({ error: 'not-recorded' }), /not set up to be recorded/);
  assert.ok(guestRefusal({ error: 'offline' }));
  assert.notEqual(guestRefusal({ error: 'not-yours' }), guestRefusal({ error: 'offline' }));
  assert.equal(sessionName('live', { name: 'Amira Khan' }), 'Live broadcast with Amira');
  assert.equal(sessionName('record', { name: 'Amira Khan' }), 'Recording with Amira');
  assert.ok(sessionName('live', null) && sessionName('record', {}));
});
