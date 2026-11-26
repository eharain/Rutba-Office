// Talk rings (Office 2.27.0), the window's side: the sounds a call makes,
// when each is made, and the settings dialog with the people blocked, drawn
// without a window. Also what Talk must not do itself: it shows no Windows
// notification, flashes nothing and puts no icon by the clock; the notifier
// does, from what main/talk-notify.js says.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import esbuild from 'esbuild';
import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { SOUNDS, soundFor, createRinger, incomingCall } from '../apps/desktop/renderer/apps/talk/ring.js';

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.join(here, '..');

/** A sound device that keeps what it was asked to play. */
function device() {
  const notes = [];
  const made = { contexts: 0, closed: 0, resumed: 0 };
  const makeContext = () => {
    made.contexts += 1;
    return {
      currentTime: 5,
      destination: {},
      resume: () => { made.resumed += 1; return Promise.resolve(); },
      close: () => { made.closed += 1; return Promise.resolve(); },
      createOscillator: () => {
        const note = { hz: 0, from: null, to: null, level: 0, stopped: false };
        notes.push(note);
        const osc = { type: '', frequency: {}, connect: () => {}, start: (at) => { note.from = at; }, stop: (at) => { if (at == null) note.stopped = true; else note.to = at; }, onended: null };
        Object.defineProperty(osc.frequency, 'value', { set: (v) => { note.hz = v; } });
        note.osc = osc;
        return osc;
      },
      createGain: () => ({ gain: { setValueAtTime: (v) => { const n = notes.at(-1); n.level = Math.max(n.level, v); }, linearRampToValueAtTime: (v) => { const n = notes.at(-1); n.level = Math.max(n.level, v); } }, connect: () => {} }),
    };
  };
  const timers = new Map();
  let ids = 0;
  return {
    notes,
    made,
    timers,
    ringer: createRinger({ makeContext, setTimer: (fn, ms) => { timers.set(++ids, { fn, ms }); return ids; }, clearTimer: (id) => timers.delete(id) }),
    /** The clock reaches every waiting timer once. */
    pass() { for (const t of [...timers.values()]) t.fn(); },
  };
}

test('which sound: a ring for a call coming in, a tone while mine rings there, none when turned off or in a build for the checks', () => {
  assert.equal(soundFor({ on: true, incoming: true }), 'ring');
  assert.equal(soundFor({ on: true, incoming: false, call: 'ringing' }), 'calling');
  // Not before it rings at the other side, and not once connected.
  for (const call of ['calling', 'connecting', 'connected', 'ended', null]) assert.equal(soundFor({ on: true, incoming: false, call }), null, String(call));
  assert.equal(soundFor({ on: false, incoming: true, call: 'ringing' }), null);
  assert.equal(soundFor({ on: true, silent: true, incoming: true }), null);
});

test('a ring plays its notes, comes again until stopped, and is not started twice', () => {
  const d = device();
  // Nothing is asked of the sound device until there is something to play.
  assert.equal(d.made.contexts, 0);
  assert.equal(d.ringer.play('ring'), true);
  assert.equal(d.ringer.playing(), 'ring');
  assert.deepEqual(d.notes.map((n) => n.hz), SOUNDS.ring.notes.map((n) => n[2]));
  assert.deepEqual(d.notes.map((n) => n.from), SOUNDS.ring.notes.map((n) => 5 + n[0]));
  assert.ok(d.notes.every((n) => n.to > n.from && n.level === SOUNDS.ring.level));
  assert.deepEqual([...d.timers.values()].map((t) => t.ms), [SOUNDS.ring.every]);
  // Asked for again while it plays: it goes on as it was.
  d.ringer.play('ring');
  assert.equal(d.notes.length, 4);
  assert.equal(d.timers.size, 1);
  d.pass();
  assert.equal(d.notes.length, 8);
  d.ringer.stop();
  assert.equal(d.ringer.playing(), null);
  assert.equal(d.timers.size, 0);
  assert.ok(d.notes.every((n) => n.stopped), 'what was sounding is cut off');
  d.pass();
  assert.equal(d.notes.length, 8);
  assert.equal(d.made.contexts, 1);
});

test('one sound gives way to another, a tone for a meeting plays once, and closing lets the sound device go', () => {
  const d = device();
  d.ringer.play('ring');
  d.ringer.play('calling');
  assert.equal(d.ringer.playing(), 'calling');
  assert.deepEqual([...d.timers.values()].map((t) => t.ms), [SOUNDS.calling.every]);
  assert.ok(SOUNDS.calling.level < SOUNDS.ring.level, 'the tone for my own call is the softer one');
  const before = d.notes.length;
  assert.equal(d.ringer.once('chime'), true);
  assert.equal(d.notes.length, before + SOUNDS.chime.notes.length);
  // Once: it sets no timer and does not stop what plays.
  assert.equal(d.timers.size, 1);
  assert.equal(d.ringer.playing(), 'calling');
  // A sound made once is not one to keep playing, and an unknown one is nothing.
  assert.equal(d.ringer.play('chime'), false);
  assert.equal(d.ringer.play('siren'), false);
  assert.equal(d.ringer.once('siren'), false);
  d.ringer.close();
  assert.equal(d.ringer.playing(), null);
  assert.equal(d.timers.size, 0);
  assert.equal(d.made.closed, 1);
  // Used again after closing, it starts afresh.
  d.ringer.play('ring');
  assert.equal(d.made.contexts, 2);
});

test('a computer with no sound says no, quietly: nothing is thrown at the window', () => {
  const ringer = createRinger({ makeContext: () => { throw new Error('no sound device'); }, setTimer: () => 1, clearTimer: () => {} });
  assert.equal(ringer.play('ring'), false);
  assert.equal(ringer.once('chime'), false);
  assert.equal(ringer.playing(), null);
  ringer.stop();
  ringer.close();
});

/* ── the settings dialog, drawn without a window ── */

let loaded = null;
function screens() {
  loaded ||= (async () => {
    const out = await esbuild.build({
      stdin: { contents: "export * from './apps/desktop/renderer/apps/talk/settings.js';", resolveDir: root, loader: 'js' },
      bundle: true,
      format: 'esm',
      platform: 'node',
      write: false,
      logLevel: 'silent',
      loader: { '.js': 'jsx', '.css': 'empty' },
      jsx: 'automatic',
      external: ['react', 'react-dom', 'react/jsx-runtime'],
    });
    const file = path.join(root, 'node_modules', `.talk-settings-screen-${process.pid}.mjs`);
    fs.writeFileSync(file, out.outputFiles[0].text);
    try { return await import(`file://${file.replace(/\\/g, '/')}`); } finally { fs.rmSync(file, { force: true }); }
  })();
  return loaded;
}
const textOf = (html) => html.replace(/<[^>]*>/g, ' ').replace(/&#x27;/g, "'").replace(/\s+/g, ' ');
const count = (html, what) => html.split(what).length - 1;
const noop = () => {};
const shell = { talk: { setOptions: async () => ({}), unblock: async () => ({ unblocked: true }) } };
const STATE = { phase: 'online', startWithOffice: true, ring: true, blocked: [], blockedPeople: [] };

test('settings: two choices, both on until turned off, and nowhere to type a server', async () => {
  const { TalkSettings } = await screens();
  const html = renderToStaticMarkup(React.createElement(TalkSettings, { shell, state: STATE, toast: noop, onClose: noop }));
  const text = textOf(html);
  assert.match(text, /Talk settings/);
  assert.match(text, /Connect when Rutba Office starts/);
  assert.match(text, /whenever Rutba Office is open, without Talk being opened first/);
  assert.match(text, /Play a sound for calls/);
  assert.equal(count(html, 'type="checkbox"'), 2);
  assert.equal(count(html, 'checked=""'), 2);
  assert.match(text, /Nobody is blocked\./);
  assert.equal(count(html, 'tk-unblock'), 0);
  // No box to type in at all: a server comes only from the introducer.
  assert.equal(count(html, '<input type="text"') + count(html, '<textarea') + count(html, 'type="url"'), 0);

  const off = renderToStaticMarkup(React.createElement(TalkSettings, { shell, state: { ...STATE, startWithOffice: false, ring: false }, toast: noop, onClose: noop }));
  assert.equal(count(off, 'checked=""'), 0);
});

test('settings: each person blocked is listed by name with Unblock, and what unblocking does is said', async () => {
  const { TalkSettings, blockedName } = await screens();
  const blockedPeople = [{ id: 'u-ben', name: 'Ben Ortiz', email: 'ben@example.com' }, { id: 'u-eve', name: 'Eve Marsh', email: '' }, { id: 'u-old', name: '', email: '' }];
  const html = renderToStaticMarkup(React.createElement(TalkSettings, { shell, state: { ...STATE, blockedPeople }, toast: noop, onClose: noop }));
  const text = textOf(html);
  assert.equal(count(html, 'data-blocked='), 3);
  assert.equal(count(html, 'tk-unblock'), 3);
  assert.match(text, /Ben Ortiz ben@example\.com/);
  assert.match(text, /Eve Marsh/);
  // Blocked from another computer before names were kept: still listed, still unblockable.
  assert.match(text, /Someone you blocked/);
  assert.equal(blockedName({ id: 'u-x', name: '', email: 'x@example.com' }), 'x@example.com');
  assert.match(text, /can ask to be your contact again\. They are not put back in your contacts\./);
  assert.doesNotMatch(text, /Nobody is blocked/);
  assert.equal(count(html, 'disabled=""'), 0);

  // Not connected: the list still shows, and Unblock waits.
  const offline = renderToStaticMarkup(React.createElement(TalkSettings, { shell, state: { ...STATE, phase: 'offline', blockedPeople }, toast: noop, onClose: noop }));
  assert.equal(count(offline, 'disabled=""'), 3);
  assert.match(offline, /Talk is not connected\./);

  // A state from a service that names nobody (ids only) is still a list.
  const ids = renderToStaticMarkup(React.createElement(TalkSettings, { shell, state: { phase: 'online', blocked: ['u-1', 'u-2'] }, toast: noop, onClose: noop }));
  assert.equal(count(ids, 'data-blocked='), 2);
});

/* ── what Talk leaves to the notifier ── */

test('Talk shows no Windows notification, flashes nothing and puts no icon by the clock: that is the notifier\'s', () => {
  const files = [
    ...fs.readdirSync(path.join(root, 'apps/desktop/main')).filter((f) => /^talk.*\.js$/.test(f)).map((f) => path.join(root, 'apps/desktop/main', f)),
    path.join(root, 'apps/desktop/renderer/apps/talk.js'),
    ...fs.readdirSync(path.join(root, 'apps/desktop/renderer/apps/talk')).filter((f) => f.endsWith('.js')).map((f) => path.join(root, 'apps/desktop/renderer/apps/talk', f)),
  ];
  assert.ok(files.length > 15);
  for (const file of files) {
    const code = fs.readFileSync(file, 'utf8');
    for (const not of [/new Notification\b/, /\bNotification\.requestPermission\b/, /\.flashFrame\b/, /new Tray\b/, /setOverlayIcon\b/]) assert.doesNotMatch(code, not, `${path.basename(file)} ${not}`);
  }
});

test('the window makes its sounds from Talk\'s own state, asks for a call already ringing, and opens where a notification points', () => {
  const code = fs.readFileSync(path.join(root, 'apps/desktop/renderer/apps/talk.js'), 'utf8');
  assert.match(code, /soundFor\(\{ on: sounds, incoming: Boolean\(ringing\), call: call\?\.state \|\| null \}\)/);
  assert.match(code, /state\.ring !== false && !state\.testServers/);
  assert.match(code, /shell\.talk\.ringing\?\.\(\)/);
  assert.match(code, /command === 'notify\.open'/);
  assert.match(code, /q\?\.show === 'calls'/);
  assert.match(code, /className="tk-settings-btn"/);
});

/* ── one call at a time (2.27.0, the screen review) ── */

test('a second call while the first is still ringing is turned away as busy, and the first keeps ringing', () => {
  assert.equal(incomingCall({ callId: 'c-1' }), 'ring');
  assert.equal(incomingCall({ callId: 'c-2', ringing: 'c-1' }), 'busy');
  assert.equal(incomingCall({ callId: 'c-1', ringing: 'c-1' }), 'same', 'the same call told twice is not turned away');
  for (const call of ['calling', 'ringing', 'connecting', 'connected']) assert.equal(incomingCall({ callId: 'c-2', call }), 'busy', call);
  assert.equal(incomingCall({ callId: 'c-2', call: 'ended' }), 'ring', 'a call that is over holds nothing');
  assert.equal(incomingCall({ callId: 'c-2', room: true }), 'busy');
  assert.equal(incomingCall({ callId: 'c-2', call: 'ended', ringing: 'c-1' }), 'busy');
});

test('the window asks what a call coming in does, turns the busy one away by its own id, and never puts a second card over the first', () => {
  const code = fs.readFileSync(path.join(root, 'apps/desktop/renderer/apps/talk.js'), 'utf8');
  for (const has of [
    'incomingCall({ callId: msg.callId, call: c?.state || null, room: Boolean(roomRef.current), ringing: ringRef.current?.callId || null, onAir })',
    "if (what === 'busy') shell.talk.send({ type: 'call.decline', callId: msg.callId, reason: 'busy' })",
    "if (what === 'ring') { ringRef.current = msg; setRinging(msg); }",
  ]) assert.ok(code.includes(has), has);
  assert.equal(code.split('setRinging(msg)').length, 2, 'one place takes a call coming in');
});
