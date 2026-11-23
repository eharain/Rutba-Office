// Guests from Talk in the Recorder, seen in its window (Office 2.27.0,
// D-2026-11-22-19): what the people asked in are told before they join is
// what the window then does.
//
// The Recorder's own meeting is set up to be recorded and not streamed, and
// the people asked in are told so at the door. A guest who comes in is
// connected to this window, waits, and sees the programme. Heard here, their
// voice is in everything this computer plays, so Computer sound is kept out
// of the mix for as long. The seat they sit in is a source like any other:
// in the frame of a recording that is running they are in the recording, and
// not before. While it records the meeting cannot be moved to Talk. Stop
// keeps one file with picture and sound and leaves the guest here.
//
// Talk is the stand-in introducer of verify-talk.js with the meetings of
// verify-talk-meet.js; the guest is Ben's hidden page from there, with a
// stand-in camera and microphone. The recording goes into the run's folder.
// Run alone with RUTBA_VERIFY_ONLY=recorder-guests.

import fs from 'node:fs';
import path from 'node:path';
import { BrowserWindow } from 'electron';
import { probeMedia } from '@rutba/media/probe';
import { emptyStack } from '@rutba/studio/stack';
import { seedSetup, restoreSetup, CHECK_SOUNDS, PANEL, IN_PANEL } from './verify-studio-seed.js';
import { letGoOfStudioWindows } from './verify-let-go.js';
import { standInIntroducer } from './verify-talk.js';
import { standInMeetings, BEN } from './verify-talk-meet.js';

export async function verifyRecorderGuests(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  const meetings = standInMeetings();
  const intro = standInIntroducer({ page: BEN, more: meetings.more, contacts: [['u-amira', 'u-ben']] });
  const folder = path.join(dir, 'recorder-guests');
  let ben = null;
  let talk = null;
  let win = null;
  let js = null;
  let seeded = null;
  try {
    fs.mkdirSync(folder, { recursive: true });
    const port = await intro.listen();
    const url = `ws://127.0.0.1:${port}/ws`;

    // Amira signs in with a test server in Talk's window; Ben is her contact already.
    talk = await open('talk');
    const tjs = (code) => talk.webContents.executeJavaScript(code);
    const set = (cls, value) => tjs(`(() => { const n = document.querySelector('.${cls}'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, ${JSON.stringify(value)}); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await tjs(`window.rutbaOffice.talk.signOut().then(() => 1, () => 1)`);
    await until(() => tjs(`Boolean(document.querySelector('.tk-signin'))`), 'the sign-in', 8000).catch(() => {});
    await tjs(`document.querySelector('.tk-test')?.setAttribute('open', ''), 1`);
    await set('tk-server', url);
    await set('tk-token', 'amira');
    await tjs(`(document.querySelector('.tk-connect')?.click(), 1)`);
    await until(() => tjs(`Boolean(document.querySelector('.tk-contact[data-contact="u-ben"]'))`), 'signed in, Ben among the contacts', 10000).catch(() => {});
    ben = new BrowserWindow({ show: false, webPreferences: { contextIsolation: true, nodeIntegration: false, backgroundThrottling: false, partition: 'talk-check-ben-recorder' } });
    await ben.loadURL(`http://127.0.0.1:${port}/ben`);
    const bjs = (code) => ben.webContents.executeJavaScript(code);
    const benHeard = (test, what, ms = 8000) => ok(until(() => bjs(`window.ben.log.some((m) => ${test})`), what, ms));
    await until(() => bjs(`window.ben.log.some((m) => m.type === 'welcome')`), 'Ben signed in', 6000).catch(() => {});

    win = await open('recorder');
    js = (code) => win.webContents.executeJavaScript(code);
    // A setup of the check's own: nothing in the frame, a microphone and Computer sound open.
    const sounds = CHECK_SOUNDS(true).map((x) => (x.kind === 'system' ? { ...x, live: true } : x));
    seeded = await seedSetup(js, { app: 'recorder', stack: { ...emptyStack(), sounds }, studio: { size: 720, fps: 30, countdown: 0, folder, minimise: false, shot: 'program', shotTimer: 0 } });
    await win.webContents.reload();
    await until(() => js(`Boolean(document.querySelector('.st-canvas')) && Boolean(document.querySelector('.st-record'))`), 'the Recorder window', 15000).catch(() => {});
    const q = (sel) => `document.querySelector(${JSON.stringify(sel)})`;
    const click = (sel) => js(`(${q(sel)}?.click(), 1)`);
    const seen = (sel, what, ms = 8000) => ok(until(() => js(`Boolean(${q(sel)})`), what, ms));
    const gone = (sel, what, ms = 8000) => ok(until(() => js(`!${q(sel)}`), what, ms));
    // The last button of the dialog on top: the one that says yes to its question.
    const yes = () => js(`([...[...document.querySelectorAll('.rw-dialog')].at(-1).querySelectorAll('button')].at(-1)?.click(), 1)`);
    const toSetup = async () => { await js(PANEL('setup')); return ok(until(() => js(IN_PANEL('setup')), 'Set up', 6000)); };
    const toRecord = async () => { await js(PANEL('main')); return ok(until(() => js(IN_PANEL('main')), 'back on Record', 6000)); };
    // Computer sound's row in Set up: how it stands, and whether it says it is kept out.
    const computerSound = () => js(`(() => { const n = ${q('.st-snd[data-kind="system"]')}; return n ? { state: n.dataset.state, says: Boolean(n.querySelector('.st-snd-held')?.textContent.trim()) } : null; })()`);
    const soundIs = (state, says, what) => ok(until(() => computerSound().then((c) => Boolean(c) && c.state === state && c.says === says), what, 12000));
    const openGuests = async () => { await click('.st-guests-btn'); return seen('.st-guests', 'Guests'); };
    const closeGuests = async () => { await js(`(${q('.st-guests-invite')}?.previousElementSibling?.click(), 1)`); return gone('.st-guests', 'Guests closed'); };
    const BEN_ROW = '.st-guests-here .st-guests-row[data-person="u-ben"]';
    const benIs = (where) => ok(until(() => js(`${q(BEN_ROW)}?.dataset.where === ${JSON.stringify(where)}`), `Ben ${where}`, 12000));
    // The eye of the seat the guest sits in, among the sources in Set up.
    const SEAT = '.st-card[data-kind="guest"][data-state="live"]';
    const files = () => fs.readdirSync(folder).filter((f) => /\.webm$/i.test(f)).map((name) => ({ name, size: fs.statSync(path.join(folder, name)).size }));

    const there = await seen('.st-guests-btn', 'Guests in the Recorder', 15000);
    // Computer sound is open and in the mix before anyone is here.
    await toSetup();
    const soundBefore = await soundIs('live', false, 'Computer sound in the mix');
    await toRecord();

    // Asked in: the meeting is one that is recorded and not streamed, the guest is told so before joining, and
    // Guests says what everyone asked is told.
    await click('.st-guests-btn');
    await seen('.st-guests-row[data-contact="u-ben"] input', 'Ben to ask in');
    const saysBefore = await js(`(() => { const n = ${q('.st-guests .st-guests-note')}; return n ? { how: n.dataset.how, words: n.textContent.trim().length } : null; })()`);
    await click('.st-guests-row[data-contact="u-ben"] input');
    await ok(until(() => js(`${q('.st-guests-invite')}?.disabled === false`), 'Invite ready', 6000));
    const had = meetings.asked.length;
    await click('.st-guests-invite');
    await until(() => meetings.asked.length > had || null, 'the meeting set up', 8000).catch(() => {});
    const made = meetings.asked.length > had ? meetings.asked.at(-1) : {};
    const told = await benHeard(`m.type === 'meet.invited' && m.meeting.id === ${JSON.stringify(made.id || '')} && m.meeting.live === false && m.meeting.recorded === true && m.by.name === 'Amira Khan'`, 'Ben asked in, told it is recorded');
    const held = await seen('.st-guests-btn.on', 'the meeting held here');
    check('recorder guests: asking someone in sets up a meeting that is recorded and not streamed live, they are told so before they join, and Guests says what everyone asked is told',
      there && made.by === 'u-amira' && made.live === false && made.recorded === true && made.tell === true && made.invite?.join() === 'u-ben' && told && held && saysBefore?.how === 'record' && saysBefore.words > 30,
      JSON.stringify({ there, made, told, held, saysBefore }));

    // The guest comes in: connected to this window, waiting, counted on the button, and seeing the programme.
    await bjs(`window.ben.send({ type: 'meet.join', id: ${JSON.stringify(made.id || '')}, ref: 'g-join' }), 1`);
    const here = await seen(BEN_ROW, 'Ben among those here', 20000);
    const waits = await ok(until(() => js(`${q(BEN_ROW)}?.dataset.where === 'off-air'`), 'Ben waiting', 20000));
    const where = await js(`${q(BEN_ROW)}?.dataset.where || ''`);
    const counted = await ok(until(() => js(`${q('.st-guests-btn')}?.dataset.guests === '1'`), 'one guest on the button', 8000));
    const programme = await ok(until(() => bjs(`window.ben.remoteWidth > 0`), 'the programme at Ben', 15000));
    const canMove = await js(`${q('.st-guests-talk')}?.disabled === false`);
    check('recorder guests: a guest who comes in is connected to this window, waits out of the picture, is counted on the Guests button and sees the programme; the meeting can still be moved to Talk',
      here && waits && counted && programme && canMove, JSON.stringify({ here, where, counted, programme, canMove }));

    // Heard here, the guest's voice is in everything this computer plays: Computer sound is kept out of the mix
    // and its row says so. The seat is a source like any other; in the frame they are in a recording only once
    // one is running.
    await closeGuests();
    await toSetup();
    const soundHeld = await soundIs('held', true, 'Computer sound kept out, saying why');
    const seated = await seen(SEAT, 'the seat Ben sits in');
    await click(`${SEAT} .st-eye`);
    const shown = await seen(`${SEAT} .st-eye.on`, 'the seat in the frame');
    await toRecord();
    await openGuests();
    const notYet = await benIs('off-air');
    await closeGuests();
    await click('.st-record');
    const recording = await ok(until(() => js(`(() => { const b = ${q('.st-record')}; return Boolean(b) && !b.disabled && b.classList.contains('on') && b.getAttribute('aria-pressed') === 'true'; })()`), 'the recording to start', 15000));
    const grew = await ok(until(() => files().some((f) => f.size > 5000) || null, 'the recording on disk', 10000));
    await openGuests();
    const inIt = await benIs('on-air');
    const cannotMove = await js(`${q('.st-guests-talk')}?.disabled === true`);
    await closeGuests();
    await toSetup();
    const heldWhile = await soundIs('held', true, 'Computer sound still kept out');
    await toRecord();
    await wait(2500);
    await click('.st-record');
    const saved = await ok(until(() => js(`Boolean(${q('.st-saved')})`), 'the saved bar', 30000));
    const made1 = files();
    let probe = null;
    try { probe = made1.length === 1 ? probeMedia(new Uint8Array(fs.readFileSync(path.join(folder, made1[0].name))), { complete: true }) : null; } catch { probe = null; }
    await openGuests();
    const after = await benIs('off-air');
    const stays = await js(`Boolean(${q('.st-guests-btn.on')}) && ${q('.st-guests-btn')}?.dataset.guests === '1'`);
    const canMoveAgain = await js(`${q('.st-guests-talk')}?.disabled === false`);
    check('recorder guests: while a guest is heard here Computer sound is kept out of the mix and its row says why; in the frame they are in the recording once one runs and not before; the meeting cannot be moved to Talk while it records; Stop keeps one file with picture and sound and leaves the guest here',
      soundBefore && soundHeld && seated && shown && notYet && recording && grew && inIt && cannotMove && heldWhile && saved && made1.length === 1 && Boolean(probe?.hasVideo) && Boolean(probe?.hasAudio) && probe.duration > 2 && after && stays && canMoveAgain,
      JSON.stringify({ soundBefore, soundHeld, seated, shown, notYet, recording, grew, inIt, cannotMove, heldWhile, saved, files: made1, probe: probe ? { video: probe.hasVideo, audio: probe.hasAudio, w: probe.width, h: probe.height, duration: probe.duration } : null, after, stays, canMoveAgain }));

    // Ended for everyone: the guest is told, and Computer sound is back in the mix.
    await click('.st-guests-end');
    await ok(until(() => js(`document.querySelectorAll('.rw-dialog').length >= 2`), 'the question', 6000));
    await yes();
    const ended = await gone('.st-guests-btn.on', 'no meeting held', 8000);
    const benOut = await ok(until(() => bjs(`window.ben.over === 'ended'`), 'ended at Ben', 8000));
    await closeGuests();
    await toSetup();
    const soundBack = await soundIs('live', false, 'Computer sound in the mix again');
    await toRecord();
    check('recorder guests: End for everyone ends the meeting for the guest too, and Computer sound is back in the mix',
      ended && benOut && soundBack && !meetings.get(made.id), JSON.stringify({ ended, benOut, soundBack, kept: Boolean(meetings.get(made.id)) }));

    await wait(300);
    const complaints = await errorsIn(win);
    const benFaults = await bjs(`window.ben.log.filter((m) => m.type === 'page-error').map((m) => m.message)`);
    check('recorder guests: none of that reported an error, in Office or on the guest\'s side', complaints.length === 0 && benFaults.length === 0, [...complaints, ...benFaults].join(' | ') || 'nothing reported');
    await tjs(`window.rutbaOffice.talk.signOut().then(() => 1, () => 1)`);
  } catch (err) {
    check('recorder guests: the check ran to the end', false, err?.message || String(err));
  } finally {
    if (js && seeded) await restoreSetup(js, seeded).catch(() => {});
    letGoOfStudioWindows();
    try { ben?.destroy(); } catch { /* gone */ }
    await intro.close().catch(() => {});
  }
}
