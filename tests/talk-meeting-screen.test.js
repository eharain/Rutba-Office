// The screens of a small meeting (Office 2.27, renderer/apps/talk/meeting.js)
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
  constructor(tracks = []) { this.tracks = [...tracks]; }
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
        contents: "export * from './apps/desktop/renderer/apps/talk/meeting.js'; export { liveNotice, recordedNotice, airText } from './apps/desktop/renderer/apps/talk/meet-lines.js';",
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
function room({ mine = false, live = false, recorded = false, air = null, locked = false, people = null, says = [], unseen = 0, mic = true, micFault = null, cameraFault = null } = {}) {
  const meeting = { id: live ? 'live-BBBBBBBBBBBBBBBBBBBB' : ID, name: 'Monday plan', starter: { id: mine ? 'u-a' : 'u-ben', name: mine ? 'Amira Khan' : 'Ben Ortiz' }, live, recorded, locked, link: `https://introducer.example/m/${ID}`, most: 4 };
  const list = people || [
    { id: 'u-a', name: 'Amira Khan', me: true, starter: mine, connected: true, muted: false, camera: false, sharing: false, hand: false },
    { id: 'u-ben', name: 'Ben Ortiz', me: false, starter: !mine, connected: true, muted: true, camera: false, sharing: false, hand: true, stream: null },
    { id: 'u-cara', name: 'Cara Lund', me: false, starter: false, connected: false, failed: false, muted: false, camera: false, sharing: false, hand: false, stream: null },
  ];
  return {
    state: 'in', me: ME, meeting, mine, says, unseen, air: air || { on: false, you: false, how: null },
    mesh: { mic: mic ? {} : null, camera: null, screen: null, muted: !mic, hand: false, micFault: mic ? null : micFault || 'none', cameraFault, setMuted: noop, setCamera: noop, setHand: noop, retryMic: noop },
    people: () => list, seen: noop, say: noop, leave: noop, lock: noop, remove: noop, mute: noop, end: noop,
  };
}

test('Meet now: a name, up to three contacts, and the one switch, off until chosen', async () => {
  const { MeetSetup } = await screens();
  const html = draw(MeetSetup, { shell, state: STATE, toast: noop, onClose: noop, onStarted: noop });
  const text = textOf(html);
  assert.match(html, /class="[^"]*tk-meet-name[^"]*"[^>]*value="Meeting with Amira"/);
  assert.ok(text.includes('Ben Ortiz') && text.includes('Cara Lund'));
  assert.equal(count(html, 'type="checkbox"'), 5, 'two contacts, the switch, and the microphone and camera');
  assert.equal(count(html, 'tk-meet-start'), 1, 'one button starts it');
  assert.ok(!text.includes('Start with video'), 'the camera is chosen with its switch');
  assert.ok(text.includes('This meeting will be streamed live'));
  assert.ok(text.includes('a meeting cannot be made live once it is set up'), 'with the switch off, what is said is about leaving it off');
  assert.ok(!text.includes('Everyone asked is told'), 'and not what happens when it is on');
  assert.match(html, /data-says="off"/);
  assert.doesNotMatch(html, /tk-meet-live"[^>]*checked/, 'not live unless chosen');
  // Nobody chosen yet: nothing to start.
  assert.match(html, /<button[^>]*tk-meet-start[^>]*disabled/);
});

test('Meet now in a group asks nobody by name: everyone in it is told', async () => {
  const { MeetSetup } = await screens();
  const html = draw(MeetSetup, { shell, state: STATE, group: { id: 'g_1', title: 'Launch' }, toast: noop, onClose: noop, onStarted: noop });
  assert.match(html, /tk-meet-name[^>]*value="Launch"/);
  assert.equal(count(html, 'type="checkbox"'), 3, 'only the switch, and the microphone and camera');
  assert.ok(textOf(html).includes('Everyone in this group is told'));
  assert.doesNotMatch(html, /<button[^>]*tk-meet-start"[^>]*disabled/, 'ready at once');
});

test('the card that asks someone in says who, and that the meeting will be streamed live', async () => {
  const { MeetCard } = await screens();
  const by = { id: 'u-ben', name: 'Ben Ortiz' };
  const plain = textOf(draw(MeetCard, { asked: { meeting: { id: ID, name: 'Monday plan', live: false }, by }, onJoin: noop, onDismiss: noop }));
  assert.ok(plain.includes('Monday plan') && plain.includes('Ben Ortiz started a meeting') && plain.includes('Join') && plain.includes('Not now'));
  assert.ok(!plain.includes('streamed live'));
  assert.doesNotMatch(draw(MeetCard, { asked: { meeting: { id: ID, name: 'Monday plan', live: false }, by }, onJoin: noop, onDismiss: noop }), /tk-air-chip/);
  const liveHtml = draw(MeetCard, { asked: { meeting: { id: ID, name: 'Town hall', live: true }, by }, onJoin: noop, onDismiss: noop });
  assert.ok(textOf(liveHtml).includes('Will be streamed live'));
  // A mark of its own with its sign, not a few grey words after the name.
  assert.match(liveHtml, /class="tk-air-chip" data-mark="live"><svg/);
  assert.doesNotMatch(liveHtml, /<small>[^<]*streamed live/);
});

test('before joining, a meeting waits to be looked at, and says so when Talk is not connected', async () => {
  const { MeetJoin } = await screens();
  const waiting = draw(MeetJoin, { shell, link: ID, online: true, onClose: noop, onJoin: noop });
  assert.doesNotMatch(waiting, /tk-meet-join-go/, 'no Join until the meeting has been seen');
  assert.doesNotMatch(waiting, /tk-devices/, 'and no microphone or camera to choose');
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
  assert.doesNotMatch(html, /tk-tile-fault/, 'still connecting is not a failure');
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
  assert.ok(textOf(html).includes(liveNotice({ mine: true })), 'said to the person who set it up as their own doing');
  assert.ok(textOf(html).includes('You will stream this meeting live on the internet.'));
  assert.match(html, /class="tk-live in" role="note"/);
  assert.equal(count(html, 'type="checkbox"'), 0, 'nothing in the meeting turns it on or off');
});

test('a locked meeting shows it, and someone with no microphone is told so in words, with a way to try again', async () => {
  const { MeetingView } = await screens();
  const html = draw(MeetingView, { shell, room: room({ locked: true, mic: false }), toast: noop, asked: async () => false, onShare: noop });
  assert.match(html, /tk-meet-locked/);
  assert.match(html, /class="tk-device-fault in" role="alert" data-fault="microphone"/);
  const text = textOf(html);
  assert.ok(text.includes('No microphone was found on this computer. Others cannot hear you.'));
  assert.match(html, /tk-device-again/);
  // The button is never a dead one: it tries the microphone again.
  assert.match(html, /<button[^>]*tk-meet-mic[^>]*data-microphone="no"[^>]*title="Try the microphone again"/);
  assert.doesNotMatch(html, /<button[^>]*tk-meet-mic[^>]*disabled/);
});

test('a microphone or a camera that could not be used says what to check, and names no one system', async () => {
  const { MeetingView } = await screens();
  const html = draw(MeetingView, { shell, room: room({ mic: false, micFault: 'blocked', cameraFault: 'blocked' }), toast: noop, asked: async () => false, onShare: noop });
  const text = textOf(html);
  assert.ok(text.includes('Talk could not use your microphone. Others cannot hear you.'));
  assert.ok(text.includes('Talk could not use your camera.'));
  assert.match(html, /data-fault="camera"/);
  assert.ok(text.includes('privacy settings of this computer'));
  assert.ok(!/Windows/.test(text), 'Office also runs on a Mac and on Linux');
  const fine = draw(MeetingView, { shell, room: room(), toast: noop, asked: async () => false, onShare: noop });
  assert.doesNotMatch(fine, /tk-device-fault/);
});

/* ── before coming in: the microphone, the camera, and who cannot come in (the screen review of 2026-11-22) ── */

const pickOf = (over = {}) => ({ micOn: true, camera: null, cameraFault: null, opening: false, setMic: noop, setCamera: noop, take: () => ({ video: false, muted: false, camera: null }), ...over });
const seenOf = (over = {}) => ({ meeting: { id: ID, name: 'Monday plan', starter: { id: 'u-ben', name: 'Ben Ortiz' }, at: null, live: false, recorded: false, locked: false, most: 4, people: [{ id: 'u-ben', name: 'Ben Ortiz' }], ...over } });

test('before coming in: two switches, the microphone on and said so, and the camera off until chosen', async () => {
  const { DevicePick } = await screens();
  const html = draw(DevicePick, { pick: pickOf() });
  assert.match(html, /class="tk-devices" data-mic="on" data-camera="off"/);
  assert.match(html, /tk-device-mic"[^>]*checked/);
  assert.doesNotMatch(html, /tk-device-camera"[^>]*checked/);
  assert.doesNotMatch(html, /<video/, 'no picture until the camera is on');
  assert.ok(textOf(html).includes('Your microphone will be on. You can turn it off here.'));

  const off = draw(DevicePick, { pick: pickOf({ micOn: false }) });
  assert.match(off, /data-mic="off"/);
  assert.ok(textOf(off).includes('Your microphone will be off. You can turn it on in the meeting.'));

  const withCamera = draw(DevicePick, { pick: pickOf({ camera: { kind: 'video' } }) });
  assert.match(withCamera, /data-camera="on"/);
  assert.match(withCamera, /<video[^>]*class="tk-devices-picture"/, 'their own picture');
  assert.match(withCamera, /tk-device-camera"[^>]*checked/);

  const opening = draw(DevicePick, { pick: pickOf({ opening: true }) });
  assert.match(opening, /tk-device-camera"[^>]*disabled/);

  const none = draw(DevicePick, { pick: pickOf({ cameraFault: 'none' }) });
  assert.match(none, /class="tk-device-fault" role="alert" data-fault="camera"/);
  assert.ok(textOf(none).includes('No camera was found on this computer.'));
});

test('a meeting that can be joined shows who is in it and the microphone and camera to come in with', async () => {
  const { MeetSeen } = await screens();
  const html = draw(MeetSeen, { online: true, seen: seenOf(), me: ME, pick: pickOf() });
  assert.match(html, /data-can="yes"/);
  assert.match(html, /tk-devices/);
  assert.doesNotMatch(html, /tk-meet-refused/);
  assert.ok(textOf(html).includes('Ben Ortiz'));
});

test('a full meeting says so before any microphone or camera is offered', async () => {
  const { MeetSeen } = await screens();
  const four = [{ id: 'u-ben', name: 'Ben Ortiz' }, { id: 'u-c', name: 'Cara Lund' }, { id: 'u-d', name: 'Dev Rao' }, { id: 'u-e', name: 'Eli Stone' }];
  const html = draw(MeetSeen, { online: true, seen: seenOf({ people: four }), me: ME, pick: pickOf() });
  assert.match(html, /data-can="no"/);
  assert.match(html, /class="tk-lead tk-meet-refused" role="alert" data-refused="full"/);
  assert.doesNotMatch(html, /tk-devices/);
  assert.ok(/full/i.test(textOf(html)) && textOf(html).includes('4'));
  // Someone already in it on another computer brings it here.
  const mine = draw(MeetSeen, { online: true, seen: seenOf({ people: [...four.slice(0, 3), { id: 'u-a', name: 'Amira Khan' }] }), me: ME, pick: pickOf() });
  assert.match(mine, /data-can="yes"/);
});

test('a locked meeting says so before joining, except to the person who set it up', async () => {
  const { MeetSeen } = await screens();
  const html = draw(MeetSeen, { online: true, seen: seenOf({ locked: true }), me: ME, pick: pickOf() });
  assert.match(html, /data-refused="locked"/);
  assert.doesNotMatch(html, /tk-devices/);
  const own = draw(MeetSeen, { online: true, seen: seenOf({ locked: true, starter: { id: 'u-a', name: 'Amira Khan' } }), me: ME, pick: pickOf() });
  assert.match(own, /data-can="yes"/);
});

test('a meeting that is over, or that one may not join, says why, with nothing to come in with', async () => {
  const { MeetSeen } = await screens();
  const over = draw(MeetSeen, { online: true, seen: { error: 'no-meeting' }, me: ME, pick: pickOf() });
  assert.match(over, /data-refused="no-meeting"/);
  assert.doesNotMatch(over, /tk-devices/);
  const notAllowed = draw(MeetSeen, { online: true, seen: { error: 'not-allowed', starter: 'Ben Ortiz' }, me: ME, pick: pickOf() });
  assert.match(notAllowed, /data-refused="not-allowed"/);
});

test('a meeting to be streamed live says so first, above who set it up', async () => {
  const { MeetSeen, liveNotice } = await screens();
  const html = draw(MeetSeen, { online: true, seen: seenOf({ live: true }), me: ME, pick: pickOf() });
  const text = textOf(html);
  const notice = liveNotice({ name: 'Ben Ortiz' });
  assert.ok(text.includes(notice), 'it names who streams it');
  assert.ok(text.indexOf(notice) < text.indexOf('Set up by Ben Ortiz'));
});

/* ── leaving, or ending for everyone ── */

test('the person who set a meeting up is asked whether they leave or end it for everyone', async () => {
  const { MeetingView, LeaveAsk } = await screens();
  const mine = draw(MeetingView, { shell, room: room({ mine: true }), toast: noop, asked: async () => false, onShare: noop });
  assert.match(mine, /tk-meet-leave" data-asks="yes"/);
  const guest = draw(MeetingView, { shell, room: room(), toast: noop, asked: async () => false, onShare: noop });
  assert.match(guest, /tk-meet-leave" data-asks="no"/, 'anyone else just leaves');
  const ask = draw(LeaveAsk, { onLeave: noop, onEnd: noop, onCancel: noop });
  const text = textOf(ask);
  assert.ok(text.includes('Leave, or end the meeting for everyone?'));
  assert.match(ask, /tk-meet-leave-only/);
  assert.match(ask, /tk-meet-end-all/);
  assert.ok(text.includes('End for everyone') && text.includes('Cancel'));
  assert.ok(text.includes('its link still works') && text.includes('the link stops working'));
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
  // In words, with a way out.
  assert.ok(textOf(html).includes('Joining…'));
  assert.match(html, /<button[^>]*tk-join-cancel/);
  const named = draw(MeetingView, { shell, room: { state: 'joining', wanted: 'Monday plan', mesh: null, meeting: null, says: [], unseen: 0, people: () => [], seen: noop, leave: noop }, toast: noop, asked: async () => false, onShare: noop });
  assert.ok(textOf(named).includes('Joining Monday plan…'));
});

/* ── on air, recorded, and taken to Live Broadcast or the Recorder (2.27.0, the plan's phase 3) ── */

test('a meeting set up to be recorded says so on its card, before joining and all the while', async () => {
  const { MeetCard, MeetingView, RecordedNotice, recordedNotice } = await screens();
  const card = draw(MeetCard, { asked: { meeting: { id: ID, name: 'Interview', recorded: true }, by: { id: 'u-ben', name: 'Ben Ortiz' } }, onJoin: noop, onDismiss: noop });
  assert.ok(textOf(card).includes('Will be recorded'));
  const plain = draw(MeetCard, { asked: { meeting: { id: ID, name: 'Interview' }, by: { id: 'u-ben', name: 'Ben Ortiz' } }, onJoin: noop, onDismiss: noop });
  assert.ok(!textOf(plain).includes('Will be recorded'));
  assert.match(draw(RecordedNotice, {}), /role="note" data-recorded="yes"/);
  const html = draw(MeetingView, { shell, room: room({ recorded: true }), toast: noop, asked: async () => false, onShare: noop });
  assert.match(html, /data-recorded="yes"/);
  assert.ok(textOf(html).includes(recordedNotice({ name: 'Ben Ortiz' })), 'it names who records it');
  assert.ok(textOf(html).includes('Ben Ortiz will record this meeting on their computer and may share the recording.'));
  assert.equal(count(html, 'type="checkbox"'), 0, 'nothing in the meeting turns it on or off');
});

test('on air: everyone in the meeting is told, and each whether they are in the picture', async () => {
  const { MeetingView, AirMark, airText, liveNotice } = await screens();
  const props = { shell, toast: noop, asked: async () => false, onShare: noop };
  const waiting = draw(MeetingView, { ...props, room: room({ live: true }) });
  assert.doesNotMatch(waiting, /data-air=/, 'off air: the notice only');
  assert.ok(textOf(waiting).includes(liveNotice({ name: 'Ben Ortiz' })));

  const on = { on: true, you: true, how: 'live' };
  const inPicture = draw(MeetingView, { ...props, room: room({ live: true, air: on }) });
  assert.match(inPicture, /class="tk-air you" role="status" aria-live="assertive" data-air="live" data-you="yes"/);
  assert.ok(textOf(inPicture).includes(airText(on)));
  const off = { on: true, you: false, how: 'live' };
  const beside = draw(MeetingView, { ...props, room: room({ live: true, air: off }) });
  assert.match(beside, /data-air="live" data-you="no"/);
  assert.ok(textOf(beside).includes(airText(off)));
  assert.notEqual(airText(on), airText(off));

  const taped = { on: true, you: true, how: 'record' };
  assert.match(draw(MeetingView, { ...props, room: room({ recorded: true, air: taped }) }), /data-air="record" data-you="yes"/);
  assert.equal(new Set([airText(on), airText(off), airText(taped), airText({ on: true, you: false, how: 'record' })]).size, 4, 'four things to say, each its own words');
  assert.equal(draw(AirMark, { air: { on: false, you: false, how: null } }), '');
  assert.equal(airText(null), '');
});

test('only the person who set up a meeting to be streamed or recorded is offered Live Broadcast or the Recorder', async () => {
  const { MeetingView } = await screens();
  const props = { shell, toast: noop, asked: async () => false, onShare: noop, onStudio: noop };
  const live = draw(MeetingView, { ...props, room: room({ live: true, mine: true }) });
  assert.equal(count(live, 'tk-meet-studio'), 1);
  assert.ok(textOf(live).includes('Open in Live Broadcast'));
  const taped = draw(MeetingView, { ...props, room: room({ recorded: true, mine: true }) });
  assert.ok(textOf(taped).includes('Open in the Recorder'));
  assert.doesNotMatch(draw(MeetingView, { ...props, room: room({ live: true }) }), /tk-meet-studio/, 'not for a guest');
  assert.doesNotMatch(draw(MeetingView, { ...props, room: room({ mine: true }) }), /tk-meet-studio/, 'not in a meeting set up for neither');
  assert.doesNotMatch(draw(MeetingView, { ...props, onStudio: null, room: room({ live: true, mine: true }) }), /tk-meet-studio/);
});

/* ── the screen review's "should fix" (2.27.0) ── */

test('a meeting set up neither to be streamed nor recorded says so before joining, and one that will be says that instead', async () => {
  const { MeetSeen } = await screens();
  const plain = draw(MeetSeen, { online: true, seen: seenOf(), me: ME, pick: pickOf() });
  assert.match(plain, /class="tk-lead tk-meet-plain" data-plain="yes"/);
  assert.ok(textOf(plain).includes('This meeting was not set up to be streamed live or recorded. That cannot be changed later.'));
  for (const over of [{ live: true }, { recorded: true }]) {
    assert.doesNotMatch(draw(MeetSeen, { online: true, seen: seenOf(over), me: ME, pick: pickOf() }), /tk-meet-plain/, JSON.stringify(over));
  }
  // Nothing is promised about a meeting one cannot come into.
  assert.doesNotMatch(draw(MeetSeen, { online: true, seen: { error: 'no-meeting' }, me: ME, pick: pickOf() }), /tk-meet-plain/);
});

test('Meet now with the switch off says what leaving it off means', async () => {
  const { MeetSetup } = await screens();
  const html = draw(MeetSetup, { shell, state: STATE, toast: noop, onClose: noop, onStarted: noop });
  assert.equal(count(html, 'data-says='), 1, 'one line for the switch, said for the way it stands');
});

test('someone who could not be connected to is named on their tile, with the likely reason and a way to try again', async () => {
  const { MeetingView } = await screens();
  const people = [
    { id: 'u-a', name: 'Amira Khan', me: true, starter: false, connected: true, muted: false, camera: false, sharing: false, hand: false },
    { id: 'u-cara', name: 'Cara Lund', me: false, starter: false, connected: false, failed: true, muted: false, camera: false, sharing: false, hand: false, stream: null },
  ];
  const html = draw(MeetingView, { shell, room: { ...room({ people }), redial: noop }, toast: noop, asked: async () => false, onShare: noop });
  assert.match(html, /class="tk-tile-fault" role="status" data-failed="yes"/);
  const text = textOf(html);
  assert.ok(text.includes('Could not connect to Cara Lund. The network between you may be blocking it.'));
  assert.match(html, /<button[^>]*tk-tile-again/);
  assert.ok(text.includes('Try again'));
  assert.ok(!text.includes('Connecting'), 'failed is not said as still connecting');
  assert.equal(count(html, 'tk-tile-fault'), 1, 'only on the tile of the person it failed with');
});

test('each round button of a meeting has its word under it, said for the way it stands', async () => {
  const { MeetingView } = await screens();
  const html = draw(MeetingView, { shell, room: room(), toast: noop, asked: async () => false, onShare: noop });
  assert.equal(count(html, 'class="tk-round'), 7);
  assert.equal(count(html, 'class="tk-ctl"'), 7, 'a word for each');
  const words = [...html.matchAll(/class="tk-ctl">.*?<small aria-hidden="true">([^<]*)<\/small>/g)].map((m) => m[1]);
  assert.deepEqual(words, ['Microphone', 'Camera off', 'Screen', 'Hand', 'Messages', 'People', 'Leave']);
  // The camera turned off is struck through, as the microphone is, and each button says what it does to a screen reader.
  assert.match(html, /<button[^>]*tk-meet-camera off[^>]*aria-label="[^"]+"/);
  assert.match(html, /<button[^>]*tk-meet-screen[^>]*aria-label="[^"]+"/);

  const alone = [{ id: 'u-a', name: 'Amira Khan', me: true, starter: false, connected: true, muted: true, camera: false, sharing: false, hand: false }];
  const muted = draw(MeetingView, { shell, room: room({ people: alone }), toast: noop, asked: async () => false, onShare: noop });
  assert.equal([...muted.matchAll(/class="tk-ctl">.*?<small aria-hidden="true">([^<]*)<\/small>/g)][0][1], 'Muted');
});

test('a locked meeting says "Meeting locked", not one bare word', async () => {
  const { MeetingView } = await screens();
  const html = draw(MeetingView, { shell, room: room({ locked: true }), toast: noop, asked: async () => false, onShare: noop });
  assert.match(html, /class="tk-meet-locked">.*?Meeting locked<\/span>/);
});

/* ── a meeting that is streamed and recorded says both (2.27.0, the screen review) ── */

test('a meeting set up to be streamed and recorded says both before joining and in it, and keeps the other in view once one is on', async () => {
  const { MeetSeen, MeetingView, liveNotice, recordedNotice } = await screens();
  const by = { name: 'Ben Ortiz' };
  const before = textOf(draw(MeetSeen, { online: true, seen: seenOf({ live: true, recorded: true }), me: ME, pick: pickOf() }));
  assert.ok(before.includes(liveNotice(by)) && before.includes(recordedNotice(by)));

  const props = { shell, toast: noop, asked: async () => false, onShare: noop };
  const waiting = textOf(draw(MeetingView, { ...props, room: room({ live: true, recorded: true }) }));
  assert.ok(waiting.includes(liveNotice(by)) && waiting.includes(recordedNotice(by)), 'both while it waits, not the first of the two only');

  // On air: the mark says it is on air, and that it is recorded stays said.
  const onAir = textOf(draw(MeetingView, { ...props, room: room({ live: true, recorded: true, air: { on: true, you: true, how: 'live' } }) }));
  assert.ok(onAir.includes('You are on air.'));
  assert.ok(onAir.includes(recordedNotice(by)), 'a copy may be kept of it: still said');
  assert.ok(!onAir.includes(liveNotice(by)), 'the mark has taken the place of the notice it makes true');
  const taped = textOf(draw(MeetingView, { ...props, room: room({ live: true, recorded: true, air: { on: true, you: true, how: 'record' } }) }));
  assert.ok(taped.includes('You are being recorded.') && taped.includes(liveNotice(by)) && !taped.includes(recordedNotice(by)));

  // One set up for one of them only says nothing of the other, on air or off.
  const liveOnly = textOf(draw(MeetingView, { ...props, room: room({ live: true, air: { on: true, you: false, how: 'live' } }) }));
  assert.ok(!liveOnly.includes('will record') && !liveOnly.includes('will stream'));
});

/* ── a meeting held on a Talk Server ──────────────────────────────── */

const PLAN = { id: 'team', maxPeople: 25, maxServers: 2, rutbaServers: false, recording: false };
const OWN = { id: 'srv-1', name: 'Office box', state: 'active', answering: true, canCarry: true, why: null, who: 'me', named: [] };
const setup = (more = {}) => ({ shell, state: STATE, toast: noop, onClose: noop, onStarted: noop, ...more });

test('Meet now offers a Talk Server only to someone who has one that can hold a meeting now', async () => {
  const { MeetSetup } = await screens();
  // Nothing known yet, a service with no servers, a service that did not answer: the dialog as it always was.
  const plain = draw(MeetSetup, setup());
  for (const servers of [null, { error: 'no-service' }, { error: 'offline' }, { plan: PLAN, servers: [] }]) {
    const html = draw(MeetSetup, setup({ servers }));
    assert.equal(html, plain, JSON.stringify(servers));
    assert.doesNotMatch(html, /tk-meet-server/);
  }
  // A server of their own that can: one switch more, off until chosen, and the rest as before.
  const html = draw(MeetSetup, setup({ servers: { plan: PLAN, servers: [OWN] } }));
  const text = textOf(html);
  assert.equal(count(html, 'type="checkbox"'), 6, 'two contacts, the two switches, and the microphone and camera');
  assert.ok(text.includes('Hold it on a Talk Server'));
  assert.ok(text.includes('For more than four people. Choose now: it cannot be changed later.'));
  assert.doesNotMatch(html, /tk-meet-server"[^>]*checked/);
  assert.ok(text.includes('A meeting holds four people'));
  assert.ok(text.includes('This meeting will be streamed live'), 'the other switch is still there');
  // A plan with the service's own servers offers it with no server of one's own.
  assert.match(draw(MeetSetup, setup({ servers: { plan: { ...PLAN, rutbaServers: true }, servers: [] } })), /tk-meet-server"/);
});

test('with the Talk Server switch on, the dialog says how many it holds and that it is not streamed or recorded', async () => {
  const { MeetSetup } = await screens();
  const html = draw(MeetSetup, setup({ servers: { plan: PLAN, servers: [OWN] }, startHeld: true }));
  const text = textOf(html);
  assert.match(html, /tk-meet-server"[^>]*checked/);
  assert.ok(text.includes('This meeting holds up to 25 people: you and the contacts you ask.'));
  assert.ok(!text.includes('A meeting holds four people'));
  assert.ok(text.includes('Up to 25 people. Sound, picture and messages are encrypted by the apps of the people in it, and the server cannot open them.'));
  assert.ok(text.includes('A meeting held on a Talk Server cannot be streamed live or recorded yet.'));
  assert.doesNotMatch(html, /tk-meet-live"/, 'no switch for what it cannot be');
  assert.equal(count(html, 'type="checkbox"'), 5, 'two contacts, the one switch, and the microphone and camera');
  // The switch on with nothing behind it (the servers went away) is a small meeting.
  const gone = textOf(draw(MeetSetup, setup({ servers: { plan: PLAN, servers: [] }, startHeld: true })));
  assert.ok(gone.includes('A meeting holds four people'));
});

test('someone whose own server cannot hold a meeting now is told why, and offered no switch for it', async () => {
  const { MeetSetup } = await screens();
  const html = draw(MeetSetup, setup({ servers: { plan: PLAN, servers: [{ ...OWN, canCarry: false, answering: false, why: 'not-answering' }] } }));
  const text = textOf(html);
  assert.doesNotMatch(html, /tk-meet-server"/);
  assert.match(html, /tk-meet-server-why/);
  assert.ok(text.includes('A meeting of more than four people is held on a Talk Server. Office box is not answering. Check that it is running and connected.'));
  assert.equal(count(html, 'type="checkbox"'), 5);
});

test('before joining, a meeting held on a Talk Server says so, and that the person who set it up lets people in', async () => {
  const { MeetSeen } = await screens();
  const there = textOf(draw(MeetSeen, { online: true, seen: seenOf({ carried: 'server', most: 25 }), me: ME, pick: pickOf() }));
  assert.ok(there.includes('Held on a Talk Server, and encrypted by the apps of the people in it.'));
  assert.ok(!there.includes('lets people in'), 'they are in it: nothing to wait for');
  const away = draw(MeetSeen, { online: true, seen: seenOf({ carried: 'server', most: 25, people: [{ id: 'u-cara', name: 'Cara Lund' }] }), me: ME, pick: pickOf() });
  assert.ok(textOf(away).includes('Ben Ortiz lets people in: you can come in once they are in the meeting.'));
  assert.match(away, /data-can="yes"/, 'they can still come in, and wait there');
  // More than four are in it, and it is not full.
  const many = draw(MeetSeen, { online: true, seen: seenOf({ carried: 'server', most: 25, people: ['u-ben', 'u-1', 'u-2', 'u-3', 'u-4', 'u-5'].map((id) => ({ id, name: id })) }), me: ME, pick: pickOf() });
  assert.match(many, /data-can="yes"/);
  // A small meeting says nothing of servers.
  assert.ok(!textOf(draw(MeetSeen, { online: true, seen: seenOf(), me: ME, pick: pickOf() })).includes('Talk Server'));
});

test('in a meeting held on a Talk Server: where the person is on the way in, then that it is encrypted by the apps', async () => {
  const { MeetingView } = await screens();
  const props = { shell, toast: noop, asked: async () => false, onShare: noop };
  const held = (server, mesh = {}) => { const r = room(); r.meeting = { ...r.meeting, carried: 'server', most: 25 }; r.server = server; Object.assign(r.mesh, { encrypted: true, unreadable: false, unreadableAll: false }, mesh); return r; };
  for (const [step, says] of [['waiting', 'Waiting for Ben Ortiz, who lets people into this meeting.'], ['asking', 'Asking Ben Ortiz to let you in…'], ['entering', "Connecting to the meeting's Talk Server…"], ['starting', 'Finding a Talk Server for the meeting…']]) {
    const html = draw(MeetingView, { ...props, room: held(step) });
    assert.match(html, new RegExp(`class="tk-meet-step" role="status" data-step="${step}"`));
    assert.ok(textOf(html).includes(says), step);
    assert.doesNotMatch(html, /tk-meet-held/, 'not said to be encrypted before they are in');
    assert.match(html, /tk-meet-leave/, 'with the way out');
  }
  const inIt = draw(MeetingView, { ...props, room: held('in') });
  assert.doesNotMatch(inIt, /tk-meet-step/);
  assert.match(inIt, /class="tk-meet-held in" role="note"/);
  assert.ok(textOf(inIt).includes('Held on a Talk Server. Sound, picture and messages are encrypted by the apps of the people in it.'));
  // Something arrives that was not encrypted by the others' apps: the line is not said.
  assert.doesNotMatch(draw(MeetingView, { ...props, room: held('in', { encrypted: false }) }), /tk-meet-held/);
  // Sound or picture that does not open here is said, in place of that line.
  const some = draw(MeetingView, { ...props, room: held('in', { unreadable: true }) });
  assert.match(some, /data-fault="unreadable"/);
  assert.doesNotMatch(some, /tk-meet-held/);
  assert.ok(textOf(some).includes('Somebody in the meeting cannot be heard or seen here'));
  assert.ok(textOf(draw(MeetingView, { ...props, room: held('in', { unreadable: true, unreadableAll: true }) })).includes('You cannot hear or see the others'));
  // A small meeting has none of these lines.
  const small = draw(MeetingView, { ...props, room: room() });
  assert.doesNotMatch(small, /tk-meet-step|tk-meet-held|data-fault="unreadable"/);
});

test('more than four people: three tiles across, then four, and a long list scrolls', async () => {
  const { MeetingView } = await screens();
  const props = { shell, toast: noop, asked: async () => false, onShare: noop };
  const crowd = (n) => { const people = Array.from({ length: n }, (_, i) => ({ id: `u-${i}`, name: `Person ${i}`, me: i === 0, starter: i === 1, connected: true, muted: false, camera: false, sharing: false, hand: false, stream: null })); const r = room({ people }); r.server = 'in'; Object.assign(r.mesh, { encrypted: true }); return r; };
  const across = (html) => Number(/grid-template-columns:repeat\((\d+)/.exec(html)[1]);
  assert.deepEqual([3, 4, 5, 9, 10, 16, 17, 40].map((n) => across(draw(MeetingView, { ...props, room: crowd(n) }))), [2, 2, 3, 3, 4, 4, 4, 4]);
  assert.doesNotMatch(draw(MeetingView, { ...props, room: crowd(16) }), /class="tk-tiles many"/);
  const many = draw(MeetingView, { ...props, room: crowd(17) });
  assert.match(many, /class="tk-tiles many"/);
  assert.equal(count(many, 'data-person="u-') >= 17, true);
});
