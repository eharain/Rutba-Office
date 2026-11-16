// Spoken words made simpler (D-2026-11-16-01): one choice of three, the same
// control in Video, the Recorder and Live (shown and read aloud, read aloud
// only, shown only), and words typed and said in each.
//
// In the Recorder and in Live, Text to speech under the frame takes two lines;
// Shown only (which needs no voice) shows them one after the other across the
// foot of the frame, the band drawn over every source, and Stop takes it
// away. In Video, Insert has Text to speech: two lines added at the playhead
// land as one spoken layer, a text clip each, one after the other, and the
// clip picked after says Shown only in its Spoken lines. With no voice added,
// the two choices that read aloud cannot be chosen, and the box offers a
// voice; with one added they can.
// Run alone with RUTBA_VERIFY_PART_IDS=recorder#verifySpeakRecorder (or
// broadcast#verifySpeakLive, video#verifySpeakVideo).

import { emptyStack } from '@rutba/studio/stack';
import { seedSetup, restoreSetup, CHECK_SOUNDS, PANEL, IN_PANEL } from './verify-studio-seed.js';
import { labelKeyed } from './verify-keys.js';
import { letGoOfStudioWindows } from './verify-let-go.js';
import { BrowserWindow } from 'electron';

const TWO_LINES = 'Welcome to the stream\nThe next part starts now';

/** Words put in a box as typed, so React hears them. */
const typeInto = (sel, text) => `(() => {
  const n = document.querySelector(${JSON.stringify(sel)});
  if (!n) return 'no ' + ${JSON.stringify(sel)};
  Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(n, ${JSON.stringify(text)});
  n.dispatchEvent(new Event('input', { bubbles: true }));
  return 1;
})()`;

const click = (sel) => `(() => { const n = document.querySelector(${JSON.stringify(sel)}); if (!n) return 'no ' + ${JSON.stringify(sel)}; if (n.disabled) return 'disabled ' + ${JSON.stringify(sel)}; n.click(); return 1; })()`;

/** The choice of three inside `root`: each one's id, whether it is chosen, and whether it can be. */
const choicesIn = (root) => `[...document.querySelectorAll(${JSON.stringify(`${root} .spoken-choice input[type=radio]`)})].map((n) => ({ id: n.value, on: n.checked, can: !n.disabled }))`;

/** What the three choices say with or without a voice: all three there, and those that read aloud only with one. */
function choicesRight(list, voiced) {
  if (list.map((c) => c.id).join(' ') !== 'both read show') return false;
  const can = Object.fromEntries(list.map((c) => [c.id, c.can]));
  return can.show && (voiced ? can.both && can.read : !can.both && !can.read);
}

/** The studio's Text to speech, in the window `js` drives, on the panel now showing. */
async function studioTyped(h, js, name) {
  const { check, until, wait } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  const opened = await js(click('.st-typed-open'));
  const shown = await ok(until(() => js(`Boolean(document.querySelector('.st-typed .spoken-box-text'))`), 'the Text to speech box', 4000));
  const voiced = await js(`Boolean(document.querySelector('.st-typed .spoken-voice'))`);
  const offered = await js(`Boolean(document.querySelector('.st-typed .spoken-box-novoice .spoken-add-voice'))`);
  const choices = await js(choicesIn('.st-typed'));
  check(`${name}: Text to speech under the frame opens a box with the choice of three, shown and read aloud, read aloud only and shown only, those that read aloud only with a voice (and a voice offered without one)`,
    opened === 1 && shown && choicesRight(choices, voiced) && (voiced || offered), JSON.stringify({ opened, shown, voiced, offered, choices }));

  await js(typeInto('.st-typed .spoken-box-text', TWO_LINES));
  await js(click('.st-typed input.spoken-choice-show'));
  await wait(150);
  const said = await js(click('.st-typed .spoken-box-say'));
  const band = await ok(until(() => js(`document.querySelector('.st-canvas')?.dataset.typed === '1'`), 'the typed band', 6000));
  const first = await js(`document.querySelector('.st-typed-now')?.textContent || ''`);
  const second = await ok(until(() => js(`/^2\\/2 The next part starts now/.test(document.querySelector('.st-typed-now')?.textContent || '')`), 'the second line', 8000));
  const gone = await ok(until(() => js(`document.querySelector('.st-canvas')?.dataset.typed === '0' && !document.querySelector('.st-typed-now')`), 'the band to go', 8000));
  check(`${name}: two lines typed and said Shown only are shown one after the other across the foot of the frame, with no voice needed, and the band goes when they are done`,
    said === 1 && band && /^1\/2 Welcome to the stream/.test(first) && second && gone, JSON.stringify({ said, band, first, second, gone }));

  await js(click('.st-typed .spoken-box-say'));
  const again = await ok(until(() => js(`document.querySelector('.st-canvas')?.dataset.typed === '1'`), 'the band again', 6000));
  const stopped = await js(click('.st-typed .spoken-box-stop'));
  const off = await ok(until(() => js(`document.querySelector('.st-canvas')?.dataset.typed === '0' && Boolean(document.querySelector('.st-typed .spoken-box-say'))`), 'stopped', 4000));
  check(`${name}: while words are said its button is Stop, which takes the band away at once`, again && stopped === 1 && off, JSON.stringify({ again, stopped, off }));
  return { voiced };
}

/** The Recorder: Text to speech on its Record panel. */
export async function verifySpeakRecorder(h, { dir }) {
  const { open, check, until, errorsIn } = h;
  const win = await open('recorder');
  const js = (code) => win.webContents.executeJavaScript(code);
  let kept = null;
  try {
    kept = await seedSetup(js, { app: 'recorder', stack: { ...emptyStack(), sounds: CHECK_SOUNDS(false) }, studio: { size: 720, fps: 30, countdown: 0, folder: dir, minimise: false } });
    await win.webContents.reload();
    await until(() => js(`Boolean(document.querySelector('.st-canvas'))`), 'the Recorder window', 15000);
    await js(PANEL('main'));
    await until(() => js(IN_PANEL('record')), 'the Record panel', 6000).catch(() => {});
    await studioTyped(h, js, 'speak recorder');
    const complaints = await errorsIn(win);
    check('speak recorder: nothing went wrong on the way', !complaints.length, complaints.join(' | '));
  } finally {
    if (kept) await restoreSetup(js, kept).catch(() => {});
    letGoOfStudioWindows();
  }
}

/** Live Broadcast: Text to speech on its Live panel, before going live (the band is the program's, live or not). */
export async function verifySpeakLive(h, { dir }) {
  const { open, check, until, errorsIn } = h;
  const win = await open('broadcast');
  const js = (code) => win.webContents.executeJavaScript(code);
  let kept = null;
  try {
    kept = await seedSetup(js, { app: 'broadcast', stack: { ...emptyStack(), sounds: CHECK_SOUNDS(false) }, studio: { size: 720, fps: 30, countdown: 0, folder: dir, liveCopy: false, minimise: false } });
    await win.webContents.reload();
    await until(() => js(`Boolean(document.querySelector('.st-canvas')) && Boolean(document.querySelector('.st-livebar'))`), 'the Live Broadcast window', 15000);
    await js(PANEL('main'));
    await until(() => js(IN_PANEL('live')), 'the Live panel', 6000).catch(() => {});
    await studioTyped(h, js, 'speak live');
    const complaints = await errorsIn(win);
    check('speak live: nothing went wrong on the way', !complaints.length, complaints.join(' | '));
  } finally {
    if (kept) await restoreSetup(js, kept).catch(() => {});
    letGoOfStudioWindows();
  }
}

/** Video: Insert, Text to speech, two lines at the playhead. */
export async function verifySpeakVideo(h) {
  const { open, check, until, wait, errorsIn } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  const win = await open('video');
  const js = (code) => win.webContents.executeJavaScript(code);
  try {
    await until(() => js(`document.querySelectorAll('.rw-tab').length > 0`), 'the ribbon', 10000);
    await js(`((list) => list.find((n) => ${labelKeyed('Insert')}(n)) || list.find((n) => n.textContent.trim() === 'Insert'))([...document.querySelectorAll('.rw-tab')])?.click(), 1`);
    const pressed = await ok(until(() => js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => ${labelKeyed('Text to speech', { tip: true })}(n)); if (!b || b.disabled) return false; b.click(); return true; })()`), 'Text to speech on Insert', 6000));
    const dialog = await ok(until(() => js(`Boolean(document.querySelector('.vd-say .spoken-box-text'))`), 'the Text to speech dialog', 4000));
    // The voices are looked for as the dialog opens.
    await wait(600);
    const voiced = await js(`Boolean(document.querySelector('.vd-say .spoken-voice'))`);
    const choices = await js(choicesIn('.vd-say'));
    check('speak video: Insert has Text to speech, which opens a box with the same choice of three as the Recorder and Live',
      pressed && dialog && choicesRight(choices, voiced), JSON.stringify({ pressed, dialog, voiced, choices }));

    const before = await js(`document.querySelectorAll('.vd-clip[data-kind="text"]').length`);
    await js(typeInto('.vd-say .spoken-box-text', TWO_LINES));
    await js(click('.vd-say input.spoken-choice-show'));
    await wait(150);
    const added = await js(click('.vd-say .spoken-box-say'));
    const closed = await ok(until(() => js(`!document.querySelector('.vd-say')`), 'the dialog to close', voiced ? 30000 : 6000));
    const clips = await js(`[...document.querySelectorAll('.vd-clip[data-kind="text"]')].map((n) => ({ id: n.dataset.id, start: Number(n.dataset.start) })).sort((a, b) => a.start - b.start)`);
    const made = clips.length - before;
    check('speak video: two lines added at the playhead land as two text clips, one after the other from the playhead, in one step',
      added === 1 && closed && made === 2 && clips[0]?.start === 0 && clips[1]?.start > 0, JSON.stringify({ added, closed, before, clips }));

    const section = await ok(until(() => js(`Boolean(document.querySelector('.vd-spoken .spoken-choice'))`), 'its Spoken lines', 5000));
    const picked = await js(choicesIn('.vd-spoken'));
    check('speak video: the clip picked after says Shown only in its Spoken lines, the same control',
      section && picked.find((c) => c.on)?.id === 'show', JSON.stringify({ section, picked }));
    const complaints = await errorsIn(win);
    check('speak video: nothing went wrong on the way', !complaints.length, complaints.join(' | '));
  } finally {
    // Its project is left unsaved: closed without asking, unless it is the last window (which would end the run).
    if (!win.isDestroyed() && BrowserWindow.getAllWindows().filter((w) => !w.isDestroyed()).length > 1) win.destroy();
  }
}
