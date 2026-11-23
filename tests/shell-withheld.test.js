// Windows kept out of capture (Office 2.27.0, the screen review of 22 Nov
// 2026): a Talk window that holds a call or a meeting is no source for a
// screenshot, a recording or a broadcast, and a call or a meeting not set up
// to be streamed or recorded is kept out of a capture of the screen as well.
//
// The rules (withheld.js) are run here with stand-in windows; the shell's
// use of them (ipc.js needs Electron) and the two screens are read as code.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { createWithheld } from '../packages/office-shell/src/electron/withheld.js';
import { METHODS } from '../packages/office-shell/src/contract.js';
import { withholdOf } from '../apps/desktop/renderer/apps/talk/meet-lines.js';

const read = (f) => fs.readFileSync(new URL(`../${f}`, import.meta.url), 'utf8').replace(/\r\n/g, '\n');

/** A window as the shell hands one in: an id, a media source id, and what was asked of it. */
function fakeWin(id, { protects = true } = {}) {
  const win = {
    id,
    destroyed: false,
    asked: [],
    listeners: [],
    isDestroyed() { return this.destroyed; },
    getMediaSourceId() { return `window:${id * 100}:0`; },
    on(name, fn) { this.listeners.push([name, fn]); },
    close() { this.destroyed = true; for (const [name, fn] of this.listeners) if (name === 'closed') fn(); },
  };
  if (protects) win.setContentProtection = (on) => win.asked.push(on);
  return win;
}

test('a window that says it holds a call or a meeting is no source until it says it no longer does', () => {
  const w = createWithheld();
  const talk = fakeWin(3);
  assert.equal(w.count(), 0);
  assert.equal(w.has('window:300:0'), false);
  assert.deepEqual(w.set(talk, { on: true, protect: true }), { on: true, protected: true });
  assert.equal(w.has('window:300:0'), true);
  assert.equal(w.has('window:400:0'), false, 'another window is not');
  assert.equal(w.has('screen:0:0'), false, 'a screen is never one');
  assert.equal(w.has(''), false);
  assert.equal(w.has(null), false);
  assert.equal(w.count(), 1);
  assert.deepEqual(w.set(talk, { on: false }), { on: false, protected: false });
  assert.equal(w.has('window:300:0'), false);
  assert.equal(w.count(), 0);
});

test('its contents are kept out of a capture of the screen only when asked, and given back when it lets go', () => {
  const w = createWithheld();
  const talk = fakeWin(3);
  // A meeting set up to be streamed or recorded: no source, and not protected.
  assert.deepEqual(w.set(talk, { on: true, protect: false }), { on: true, protected: false });
  assert.deepEqual(talk.asked, [], 'nothing asked of the system');
  // A call follows it in the same window.
  w.set(talk, { on: true, protect: true });
  assert.deepEqual(talk.asked, [true]);
  // Said again as it was: asked once.
  w.set(talk, { on: true, protect: true });
  assert.deepEqual(talk.asked, [true]);
  // Back to a meeting everyone was told about.
  w.set(talk, { on: true, protect: false });
  assert.deepEqual(talk.asked, [true, false]);
  w.set(talk, { on: true, protect: true });
  w.set(talk, { on: false });
  assert.deepEqual(talk.asked, [true, false, true, false], 'let go: the window can be captured again');
  // Letting go of what was never held asks nothing.
  w.set(talk, { on: false });
  assert.deepEqual(talk.asked, [true, false, true, false]);
});

test('a window that was never protected is not unprotected: another reason for it may stand', () => {
  const w = createWithheld();
  const talk = fakeWin(3);
  w.set(talk, { on: true, protect: false });
  w.set(talk, { on: false });
  assert.deepEqual(talk.asked, []);
});

test('a closed window is forgotten, and its closing is watched once however many calls it held', () => {
  const w = createWithheld();
  const talk = fakeWin(3);
  const other = fakeWin(4);
  for (let i = 0; i < 3; i += 1) {
    w.set(talk, { on: true, protect: true });
    w.set(talk, { on: false });
  }
  assert.equal(talk.listeners.length, 1);
  w.set(talk, { on: true, protect: true });
  w.set(other, { on: true, protect: false });
  assert.equal(w.count(), 2);
  talk.close();
  assert.equal(w.count(), 1);
  assert.equal(w.has('window:300:0'), false);
  assert.equal(w.has('window:400:0'), true);
  // Destroyed without the event reaching here: still not counted, and nothing is asked of it.
  other.destroyed = true;
  assert.equal(w.count(), 0);
  assert.equal(w.has('window:400:0'), false);
  assert.deepEqual(w.set(other, { on: true, protect: true }), { on: false, protected: false });
  assert.deepEqual(other.asked, []);
  assert.deepEqual(w.set(null, { on: true, protect: true }), { on: false, protected: false });
});

test('a system that cannot protect a window, or refuses, still leaves it out of the sources', () => {
  const w = createWithheld();
  const plain = fakeWin(5, { protects: false });
  assert.deepEqual(w.set(plain, { on: true, protect: true }), { on: true, protected: true });
  assert.equal(w.has('window:500:0'), true);
  const cross = fakeWin(6);
  cross.setContentProtection = () => { throw new Error('not here'); };
  assert.doesNotThrow(() => w.set(cross, { on: true, protect: true }));
  assert.equal(w.has('window:600:0'), true);
  assert.doesNotThrow(() => w.set(cross, { on: false }));
  assert.equal(w.count(), 1);
});

test('the Talk window asks for it while it holds a call or a meeting, and for its contents to be kept out unless everyone was told', () => {
  assert.deepEqual(withholdOf(), { on: false, protect: false });
  assert.deepEqual(withholdOf({ call: null, room: false }), { on: false, protect: false });
  assert.deepEqual(withholdOf({ call: 'ended', room: false }), { on: false, protect: false }, 'a call that is over');
  for (const state of ['calling', 'ringing', 'connecting', 'connected']) {
    assert.deepEqual(withholdOf({ call: state }), { on: true, protect: true }, state);
  }
  // A meeting like any other: kept out of a capture of the screen.
  assert.deepEqual(withholdOf({ room: true, meeting: { id: 'm1' } }), { on: true, protect: true });
  // Still coming in, the meeting not yet known: taken as one nobody was told anything about.
  assert.deepEqual(withholdOf({ room: true, meeting: null }), { on: true, protect: true });
  // Set up to be streamed or recorded: everyone was told before joining. Still no source: it goes on air through Guests.
  assert.deepEqual(withholdOf({ room: true, meeting: { live: true } }), { on: true, protect: false });
  assert.deepEqual(withholdOf({ room: true, meeting: { recorded: true } }), { on: true, protect: false });
  assert.deepEqual(withholdOf({ room: true, meeting: { live: true, recorded: true } }), { on: true, protect: false });
});

test('the shell answers from it: such a window is in no list, is never chosen or given to a recording, and no picture is taken of it', () => {
  assert.ok(METHODS.capture.includes('withhold'));
  assert.ok(METHODS.capture.includes('withheld'));
  const ipc = read('packages/office-shell/src/electron/ipc.js');
  for (const has of [
    "import { withheld } from './withheld.js';",
    // The list every chooser reads, and a check run's stand-in for it.
    '.filter((s) => s.id !== own && !withheld.has(s.id) && !s.thumbnail.isEmpty() && (!CHECK || ours.has(s.id)))',
    'if (id === own || withheld.has(id)) continue;',
    // Chosen by its id all the same, or chosen before the call began.
    "if (withheld.has(id)) throw new Error(t('shell.ipc.window-call-meeting'));\n      installDisplayHandler();",
    'if (withheld.has(choice.id)) return callback({});',
    // A window speaks for itself only: the asking window is the one kept out.
    'withhold: ({ on = false, protect = false } = {}, win) => withheld.set(win, { on: on === true, protect: protect === true }),',
    'withheld: () => ({ count: withheld.count() }),',
    "grab: async ({ id }) => {\n      if (withheld.has(id)) throw new Error(t('shell.ipc.window-call-meeting'));",
  ]) assert.ok(ipc.includes(has), has);
  // The rules themselves need no Electron, so they are what was run above.
  assert.ok(!/from 'electron'|require\(/.test(read('packages/office-shell/src/electron/withheld.js')));
});

test('the Talk window says so as its call or meeting comes and goes, and the studio says why a Talk window is not in its list', () => {
  const talk = read('apps/desktop/renderer/apps/talk.js');
  for (const has of [
    'const hold = withholdOf({ call: call?.state || null, room: Boolean(room), meeting: room?.meeting || null });',
    'useEffect(() => { Promise.resolve(shell.capture?.withhold?.({ on: hold.on, protect: hold.protect })).catch(() => {}); }, [shell, hold.on, hold.protect]);',
    'useEffect(() => () => { Promise.resolve(shell.capture?.withhold?.({ on: false })).catch(() => {}); }, [shell]);',
  ]) assert.ok(talk.includes(has), has);
  const studio = read('apps/desktop/renderer/studio/window.js');
  for (const has of [
    'Promise.resolve().then(() => shell.capture.withheld?.()).then((r) => r?.count || 0, () => 0),',
    'setKeptOut(out);',
    "...(k === 'window' && keptOut ? ['-', { label: t('studio.window.talk-windows-cannot'), disabled: true }] : []),",
  ]) assert.ok(studio.includes(has), has);
});

test('the two new lines read as written, and neither promises what no program can', () => {
  const labels = read('packages/office-ui/src/catalogues/labels.json');
  for (const line of [
    '"studio.window.talk-windows-cannot": {"english":"Talk windows cannot be shown here. Put a meeting on air with Guests."',
    '"shell.ipc.window-call-meeting": {"english":"That window is in a call or a meeting, so it cannot be shown or recorded."',
  ]) assert.ok(labels.includes(line), line);
});
