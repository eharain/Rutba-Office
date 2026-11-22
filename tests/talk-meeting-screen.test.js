// The screens of a small meeting (Office 2.28, renderer/apps/talk/meeting.js)
// drawn without a window: setting one up, looking at it before joining, the
// card that asks someone in, and the meeting itself. What they hold to: the
// one switch at set-up, the notice wherever a meeting will be streamed live,
// and what only the person who set the meeting up is offered.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import esbuild from 'esbuild';
import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';

const here = path.dirname(fileURLToPath(import.meta.url));

// The meeting's own picture is a MediaStream made in the window.
class FakeStream {
  constructor() { this.tracks = []; }
  getTracks() { return this.tracks; }
  addTrack(t) { this.tracks.push(t); }
  removeTrack(t) { this.tracks = this.tracks.filter((x) => x !== t); }
}
globalThis.MediaStream ||= FakeStream;

/** The screens bundled for node, React shared with this test. */
let loaded = null;
function screens() {
  loaded ||= (async () => {
    const out = await esbuild.build({
      stdin: {
        contents: "export * from './apps/desktop/renderer/apps/talk/meeting.js'; export { liveNotice } from './apps/desktop/renderer/apps/talk/meet-lines.js';",
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
    const file = path.join(here, '..', 'node_modules', `.talk-meeting-screen-${process.pid}.mjs`);
    fs.writeFileSync(file, out.outputFiles[0].text);
    try { return await import(`file://${file.replace(/\\/g, '/')}`); } finally { fs.rmSync(file, { force: true }); }
  })();
  return loaded;
}

const textOf = (html) => html.replace(/<[^>]*>/g, ' ').replace(/&#x27;/g, "'").replace(/\s+/g, ' ');
const count = (html, what) => html.split(what).length - 1;
const draw = (Component, props) => renderToStaticMarkup(React.createElement(Component, props));

const ME = { id: 'u-a', name: 'Amira Khan' };
const STATE = { me: ME, contacts: [{ id: 'u-ben', name: 'Ben Ortiz', email: 'ben@example.com', presence: 'online' }, { id: 'u-cara', name: 'Cara Lund', email: 'cara@example.com', presence: 'offline' }] };
const ID = 'meet-AAAAAAAAAAAAAAAAAAAA';
const shell = { talk: { meetPeek: async () => ({ error: 'offline' }) }, clipboard: { writeText: async () => {} } };
const noop = () => {};

/** A room as the window holds it, standing still. */
function room({ mine = false, live = false, locked = false, people = null, says = [], unseen = 0, mic = true } = {}) {
  const meeting = { id: live ? 'live-BBBBBBBBBBBBBBBBBBBB' : ID, name: 'Monday plan', starter: { id: mine ? 'u-a' : 'u-ben', name: mine ? 'Amira Khan' : 'Ben Ortiz' }, live, locked, link: `https://introducer.example/m/${ID}`, most: 4 };
  const list = people || [
    { id: 'u-a', name: 'Amira Khan', me: true, starter: mine, connected: true, muted: false, camera: false, sharing: false, hand: false },
    { id: 'u-ben', name: 'Ben Ortiz', me: false, starter: !mine, connected: true, muted: true, camera: false, sharing: false, hand: true, stream: null },
    { id: 'u-cara', name: 'Cara Lund', me: false, starter: false, connected: false, failed: false, muted: false, camera: false, sharing: false, hand: false, stream: null },
  ];
  return {
    state: 'in', meeting, mine, says, unseen,
    mesh: { mic: mic ? {} : null, camera: null, screen: null, muted: !mic, hand: false, setMuted: noop, setCamera: noop, setHand: noop },
    people: () => list, seen: noop, say: noop, leave: noop, lock: noop, remove: noop, mute: noop, end: noop,
  };
}

test('Meet now: a name, up to three contacts, and the one switch, off until chosen', async () => {
  const { MeetSetup } = await screens();
  const html = draw(MeetSetup, { shell, state: STATE, toast: noop, onClose: noop, onStarted: noop });
  const text = textOf(html);
  assert.match(html, /class="[^"]*tk-meet-name[^"]*"[^>]*value="Meeting with Amira"/);
  assert.ok(text.includes('Ben Ortiz') && text.includes('Cara Lund'));
  assert.equal(count(html, 'type="checkbox"'), 3, 'two contacts and the switch');
  assert.ok(text.includes('This meeting will be streamed live'));
  assert.ok(text.includes('cannot be changed once the meeting is set up'));
  assert.doesNotMatch(html, /tk-meet-live"[^>]*checked/, 'not live unless chosen');
  // Nobody chosen yet: nothing to start.
  assert.match(html, /<button[^>]*tk-meet-start[^>]*disabled/);
});

test('Meet now in a group asks nobody by name: everyone in it is told', async () => {
  const { MeetSetup } = await screens();
  const html = draw(MeetSetup, { shell, state: STATE, group: { id: 'g_1', title: 'Launch' }, toast: noop, onClose: noop, onStarted: noop });
  assert.match(html, /tk-meet-name[^>]*value="Launch"/);
  assert.equal(count(html, 'type="checkbox"'), 1, 'only the switch');
  assert.ok(textOf(html).includes('Everyone in this group is told'));
  assert.doesNotMatch(html, /<button[^>]*tk-meet-start"[^>]*disabled/, 'ready at once');
});

test('the card that asks someone in says who, and that the meeting will be streamed live', async () => {
  const { MeetCard } = await screens();
  const by = { id: 'u-ben', name: 'Ben Ortiz' };
  const plain = textOf(draw(MeetCard, { asked: { meeting: { id: ID, name: 'Monday plan', live: false }, by }, onJoin: noop, onDismiss: noop }));
  assert.ok(plain.includes('Monday plan') && plain.includes('Ben Ortiz started a meeting') && plain.includes('Join') && plain.includes('Not now'));
  assert.ok(!plain.includes('streamed live'));
  const live = textOf(draw(MeetCard, { asked: { meeting: { id: ID, name: 'Town hall', live: true }, by }, onJoin: noop, onDismiss: noop }));
  assert.ok(live.includes('Will be streamed live'));
});

test('before joining, a meeting waits to be looked at, and says so when Talk is not connected', async () => {
  const { MeetJoin } = await screens();
  const waiting = draw(MeetJoin, { shell, link: ID, online: true, onClose: noop, onJoin: noop });
  assert.doesNotMatch(waiting, /tk-meet-join-voice/, 'no Join until the meeting has been seen');
  const off = textOf(draw(MeetJoin, { shell, link: ID, online: false, onClose: noop, onJoin: noop }));
  assert.ok(off.includes('not connected'));
});

test('the meeting: everyone on a tile, with who is muted, whose hand is up and who is still connecting', async () => {
  const { MeetingView } = await screens();
  const html = draw(MeetingView, { shell, room: room(), toast: noop, asked: async () => false, onShare: noop });
  assert.equal((html.match(/class="tk-tile( hand)?"/g) || []).length, 3, 'three tiles');
  assert.match(html, /data-person="u-ben"[^>]*data-connected="yes"/);
  assert.match(html, /class="tk-tile hand" data-person="u-ben"/);
  assert.match(html, /data-person="u-cara" data-connected="no"/);
  const text = textOf(html);
  assert.ok(text.includes('Monday plan') && text.includes('3 people') && text.includes('Connecting'));
  assert.match(html, /grid-template-columns:repeat\(2,/);
  // One row of buttons: mute, camera, share, hand, messages, people, leave.
  assert.equal(count(html, 'class="tk-round'), 7);
  assert.match(html, /tk-meet-leave/);
  assert.ok(!text.includes('streamed live'));
  assert.match(html, /data-live="no"/);
});

test('a meeting set up to be streamed live says so all the while, and has no switch for it', async () => {
  const { MeetingView, liveNotice } = await screens();
  const html = draw(MeetingView, { shell, room: room({ live: true, mine: true }), toast: noop, asked: async () => false, onShare: noop });
  assert.match(html, /data-live="yes"/);
  assert.ok(textOf(html).includes(liveNotice()));
  assert.match(html, /class="tk-live in" role="note"/);
  assert.equal(count(html, 'type="checkbox"'), 0, 'nothing in the meeting turns it on or off');
});

test('a locked meeting shows it, and someone with no microphone cannot unmute', async () => {
  const { MeetingView } = await screens();
  const html = draw(MeetingView, { shell, room: room({ locked: true, mic: false }), toast: noop, asked: async () => false, onShare: noop });
  assert.match(html, /tk-meet-locked/);
  assert.match(html, /<button[^>]*disabled[^>]*title="No microphone was found"|<button[^>]*title="No microphone was found"[^>]*disabled/);
});

test('messages not yet seen are counted on their button', async () => {
  const { MeetingView } = await screens();
  const some = draw(MeetingView, { shell, room: room({ unseen: 3 }), toast: noop, asked: async () => false, onShare: noop });
  assert.match(some, /class="tk-round-count">3</);
  const many = draw(MeetingView, { shell, room: room({ unseen: 12 }), toast: noop, asked: async () => false, onShare: noop });
  assert.match(many, /class="tk-round-count">9\+</);
  const none = draw(MeetingView, { shell, room: room(), toast: noop, asked: async () => false, onShare: noop });
  assert.doesNotMatch(none, /tk-round-count/);
});

test('a room still coming in draws a wait, never a broken screen', async () => {
  const { MeetingView } = await screens();
  const html = draw(MeetingView, { shell, room: { state: 'joining', mesh: null, meeting: null, says: [], unseen: 0, people: () => [], seen: noop }, toast: noop, asked: async () => false, onShare: noop });
  assert.match(html, /data-state="joining"/);
  assert.doesNotMatch(html, /tk-round/);
});
