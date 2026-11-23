// Guests from Talk in Live Broadcast, seen in its window (Office 2.27.0,
// D-2026-11-22-19, and the screen review of 22 November 2026): what the
// people asked in are told before they join is what the window then does.
//
// Live Broadcast's own meeting is set up to be streamed live. While it is
// held no copy of the broadcast can be kept, since its people were not told
// of a recording; with Keep a copy on before anyone is asked, the meeting is
// set up as recorded too and both are said at the door. A guest who comes in
// is connected to this window, waits off air, and sees the programme.
//
// Talk is the stand-in introducer of verify-talk.js with the meetings of
// verify-talk-meet.js; the guest is Ben's hidden page from there, with a
// stand-in camera and microphone. The one broadcast goes to a local ingest
// (verify-broadcast-live.js), nowhere else.

import path from 'node:path';
import { BrowserWindow } from 'electron';
import { emptyStack } from '@rutba/studio/stack';
import { seedSetup, restoreSetup, CHECK_SOUNDS, PANEL, IN_PANEL } from './verify-studio-seed.js';
import { startIngest } from './verify-broadcast-live.js';
import { letGoOfStudioWindows } from './verify-let-go.js';
import { standInIntroducer } from './verify-talk.js';
import { standInMeetings, BEN } from './verify-talk-meet.js';

export async function verifyBroadcastGuests(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  const meetings = standInMeetings();
  const intro = standInIntroducer({ page: BEN, more: meetings.more, contacts: [['u-amira', 'u-ben']] });
  let ben = null;
  let talk = null;
  let win = null;
  let js = null;
  let seeded = null;
  let ingest = null;
  try {
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
    ben = new BrowserWindow({ show: false, webPreferences: { contextIsolation: true, nodeIntegration: false, backgroundThrottling: false, partition: 'talk-check-ben-guest' } });
    await ben.loadURL(`http://127.0.0.1:${port}/ben`);
    const bjs = (code) => ben.webContents.executeJavaScript(code);
    const benHeard = (test, what, ms = 8000) => ok(until(() => bjs(`window.ben.log.some((m) => ${test})`), what, ms));
    await until(() => bjs(`window.ben.log.some((m) => m.type === 'welcome')`), 'Ben signed in', 6000).catch(() => {});

    win = await open('broadcast');
    js = (code) => win.webContents.executeJavaScript(code);
    // Somewhere to broadcast to: a local ingest, the only place saved.
    ingest = await startIngest();
    for (const d of await js(`window.rutbaOffice.live.destinations()`)) await js(`window.rutbaOffice.live.removeDestination({ id: ${JSON.stringify(d.id)} })`);
    await js(`window.rutbaOffice.live.saveDestination(${JSON.stringify({ platform: 'custom', name: 'Guest stream', url: `rtmp://127.0.0.1:${ingest.port}/live2`, key: 'guest-key-1G', shape: 'wide' })})`);
    // A setup of the check's own: nothing in the frame, Keep a copy off, a microphone and Computer sound open.
    const sounds = CHECK_SOUNDS(true).map((x) => (x.kind === 'system' ? { ...x, live: true } : x));
    seeded = await seedSetup(js, { app: 'broadcast', stack: { ...emptyStack(), sounds }, studio: { size: 720, fps: 30, countdown: 0, folder: path.join(dir, 'guest-recordings'), liveCopy: false, minimise: false, shot: 'program', shotTimer: 0 } });
    await win.webContents.reload();
    await until(() => js(`Boolean(document.querySelector('.st-canvas')) && Boolean(document.querySelector('.st-livebar'))`), 'the Live Broadcast window', 15000).catch(() => {});
    const q = (sel) => `document.querySelector(${JSON.stringify(sel)})`;
    const click = (sel) => js(`(${q(sel)}?.click(), 1)`);
    const seen = (sel, what, ms = 8000) => ok(until(() => js(`Boolean(${q(sel)})`), what, ms));
    const gone = (sel, what, ms = 8000) => ok(until(() => js(`!${q(sel)}`), what, ms));
    // Keep a copy as the window has it: on, greyed, and whether it says why.
    const copy = () => js(`(() => { const n = ${q('.st-keep-copy input')}; return n ? { on: n.checked, greyed: n.disabled, says: Boolean(${q('.st-no-copy')}) } : null; })()`);
    const copyIs = (want, what) => ok(until(() => copy().then((c) => Boolean(c) && c.on === want.on && c.greyed === want.greyed && c.says === want.says), what, 8000));
    // The last button of the dialog on top: the one that says yes to its question.
    const yes = () => js(`([...[...document.querySelectorAll('.rw-dialog')].at(-1).querySelectorAll('button')].at(-1)?.click(), 1)`);
    const ask = async () => {
      await click('.st-guests-btn');
      await seen('.st-guests-row[data-contact="u-ben"] input', 'Ben to ask in');
      await click('.st-guests-row[data-contact="u-ben"] input');
      await ok(until(() => js(`${q('.st-guests-invite')}?.disabled === false`), 'Invite ready', 6000));
      const had = meetings.asked.length;
      await click('.st-guests-invite');
      await until(() => meetings.asked.length > had || null, 'the meeting set up', 8000).catch(() => {});
      return meetings.asked.length > had ? meetings.asked.at(-1) : {};
    };
    const endAll = async () => {
      await click('.st-guests-end');
      await ok(until(() => js(`document.querySelectorAll('.rw-dialog').length >= 2`), 'the question', 6000));
      await yes();
      return gone('.st-guests-btn.on', 'no meeting held', 8000);
    };

    const toSetup = async () => { await js(PANEL('setup')); return ok(until(() => js(IN_PANEL('setup')), 'Set up', 6000)); };
    const toLive = async () => { await js(PANEL('main')); return ok(until(() => js(IN_PANEL('live')), 'back on Live', 6000)); };
    // Computer sound's row in Set up: how it stands, and whether it says it is kept out.
    const computerSound = () => js(`(() => { const n = ${q('.st-snd[data-kind="system"]')}; return n ? { state: n.dataset.state, says: Boolean(n.querySelector('.st-snd-held')?.textContent.trim()) } : null; })()`);
    const soundIs = (state, says, what) => ok(until(() => computerSound().then((c) => Boolean(c) && c.state === state && c.says === says), what, 12000));
    const openGuests = async () => { await click('.st-guests-btn'); return seen('.st-guests', 'Guests'); };
    const closeGuests = async () => { await js(`(${q('.st-guests-invite')}?.previousElementSibling?.click(), 1)`); return gone('.st-guests', 'Guests closed'); };
    const benIs = (where) => ok(until(() => js(`${q('.st-guests-here .st-guests-row[data-person="u-ben"]')}?.dataset.where === ${JSON.stringify(where)}`), `Ben ${where}`, 12000));
    // The eye of the seat the guest sits in, among the sources in Set up.
    const SEAT = '.st-card[data-kind="guest"][data-state="live"]';

    const there = await seen('.st-guests-btn', 'Guests in Live Broadcast', 15000);
    await until(() => js(`Boolean(${q('.st-keep-copy input')})`), 'Keep a copy', 8000).catch(() => {});
    const asFound = await copy();
    if (asFound?.on) await click('.st-keep-copy input');
    const free = await copyIs({ on: false, greyed: false, says: false }, 'Keep a copy off and free');
    // Computer sound is open and in the mix before anyone is here.
    await toSetup();
    const soundBefore = await soundIs('live', false, 'Computer sound in the mix');
    await toLive();

    // Asked in with no copy kept: the meeting is one streamed live and not recorded, the guest is told so, and from
    // then on no copy can be kept, with the reason beside the switch and in Guests.
    const made = await ask();
    const told = await benHeard(`m.type === 'meet.invited' && m.meeting.id === ${JSON.stringify(made.id || '')} && m.meeting.live === true && m.meeting.recorded === false && m.by.name === 'Amira Khan'`, 'Ben asked in, told it is streamed');
    const held = await seen('.st-guests-btn.on', 'the meeting held here');
    const barred = await copyIs({ on: false, greyed: true, says: true }, 'Keep a copy greyed with its reason');
    const saidInGuests = await js(`Boolean(${q('.st-guests .st-guests-nocopy')}) && ${q('.st-guests .st-guests-note')}?.dataset.how === 'live'`);
    check('live broadcast guests: asking someone in with no copy kept sets up a meeting streamed live and not recorded, they are told so, and while it is here Keep a copy is off and greyed with the reason beside it and in Guests',
      there && free && made.by === 'u-amira' && made.live === true && made.recorded === false && made.tell === true && made.invite?.join() === 'u-ben' && told && held && barred && saidInGuests,
      JSON.stringify({ there, asFound, free, made, told, held, barred, saidInGuests }));

    // The guest comes in: connected to this window, waiting off air, counted on the button, and seeing the programme.
    await bjs(`window.ben.send({ type: 'meet.join', id: ${JSON.stringify(made.id || '')}, ref: 'g-join' }), 1`);
    const here = await seen('.st-guests-here .st-guests-row[data-person="u-ben"]', 'Ben among those here', 20000);
    const waits = await ok(until(() => js(`${q('.st-guests-here .st-guests-row[data-person="u-ben"]')}?.dataset.where === 'off-air'`), 'Ben waiting off air', 20000));
    const where = await js(`${q('.st-guests-here .st-guests-row[data-person="u-ben"]')}?.dataset.where || ''`);
    const counted = await ok(until(() => js(`${q('.st-guests-btn')}?.dataset.guests === '1'`), 'one guest on the button', 8000));
    const programme = await ok(until(() => bjs(`window.ben.remoteWidth > 0`), 'the programme at Ben', 15000));
    const stillBarred = await copyIs({ on: false, greyed: true, says: true }, 'Keep a copy still greyed');
    check('live broadcast guests: a guest who comes in is connected to this window, waits off air, is counted on the Guests button and sees the programme; no copy can be kept while they are here',
      here && waits && counted && programme && stillBarred, JSON.stringify({ here, where, counted, programme, stillBarred }));

    // Heard here, the guest's voice is in everything this computer plays: Computer sound is kept out of the mix for as
    // long, and its row says so. The seat they sit in is a source like any other: in the frame of a broadcast that is
    // going out they are on air. A change made in Set up during a broadcast waits in Next: taken off the frame there
    // the guest stays on air until Take is pressed, and then waits off air again, the broadcast going on.
    await closeGuests();
    await toSetup();
    const soundHeld = await soundIs('held', true, 'Computer sound kept out, saying why');
    const seated = await seen(SEAT, 'the seat Ben sits in');
    await click(`${SEAT} .st-eye`);
    const shown = await seen(`${SEAT} .st-eye.on`, 'the seat in the frame');
    await toLive();
    await openGuests();
    const notYet = await benIs('off-air');
    await closeGuests();
    const tile = `.lv-tile[data-id="${(await js(`window.rutbaOffice.live.destinations()`))[0]?.id}"]`;
    await seen(tile, 'the place to broadcast to');
    await click(`${tile} .lv-tile-go`);
    await ok(until(() => js(`${q('.lv-go')}?.disabled === false`), 'Go live ready', 8000));
    await click('.lv-go');
    const live = await ok(until(() => js(`${q(tile)}?.dataset.state === 'live'`), 'the broadcast going out', 30000));
    const out = () => ingest.seen.connections.find((c) => c.key === 'guest-key-1G');
    const arrives = await ok(until(() => (out()?.video.length || 0) > 30 && (out()?.audio.length || 0) > 5 || null, 'picture and sound at the ingest', 30000));
    await openGuests();
    const onAir = await benIs('on-air');
    await closeGuests();
    await toSetup();
    const heldOnAir = await soundIs('held', true, 'Computer sound still kept out');
    await click(`${SEAT} .st-eye`);
    const takenOff = await gone(`${SEAT} .st-eye.on`, 'the seat out of the frame');
    await toLive();
    await openGuests();
    const untilTake = await js(`${q('.st-guests-here .st-guests-row[data-person="u-ben"]')}?.dataset.where === 'on-air'`);
    await closeGuests();
    const waiting = await seen('.st-take-btn', 'Take, with the change waiting');
    await click('.st-take-btn');
    await openGuests();
    const offAgain = await benIs('off-air');
    const stillLive = await js(`${q(tile)}?.dataset.state === 'live'`);
    await closeGuests();
    await click('.lv-end');
    await seen('.lv-ask', 'the question');
    await click('.lv-ask .lv-end-now');
    const over = await ok(until(() => js(`Boolean(${q('.lv-go')}) && !${q('.lv-end')}`), 'the broadcast ended', 15000));
    const guestStays = await js(`Boolean(${q('.st-guests-btn.on')}) && ${q('.st-guests-btn')}?.dataset.guests === '1'`);
    await openGuests();
    check('live broadcast guests: while a guest is heard here Computer sound is kept out of the mix and its row says why; the seat the guest sits in is among the sources, and in the frame of a broadcast that is going out they are on air, not before; taken off the frame in Set up they stay on air until Take, then wait off air again with the broadcast going on, and ending the broadcast leaves the guest here',
      soundBefore && soundHeld && seated && shown && notYet && live && arrives && onAir && heldOnAir && takenOff && untilTake && waiting && offAgain && stillLive && over && guestStays,
      JSON.stringify({ soundBefore, soundHeld, seated, shown, notYet, live, arrives, onAir, heldOnAir, takenOff, untilTake, waiting, offAgain, stillLive, over, guestStays }));

    // Ended for everyone: the guest is told, Keep a copy is the person's again, and Computer sound is back in the mix.
    const ended = await endAll();
    const benOut = await ok(until(() => bjs(`window.ben.over === 'ended'`), 'ended at Ben', 8000));
    const freed = await copyIs({ on: false, greyed: false, says: false }, 'Keep a copy free again');
    await toSetup();
    const soundBack = await soundIs('live', false, 'Computer sound in the mix again');
    await toLive();
    check('live broadcast guests: End for everyone ends the meeting for the guest too, Keep a copy can be chosen again and Computer sound is back in the mix',
      ended && benOut && freed && soundBack && !meetings.get(made.id), JSON.stringify({ ended, benOut, freed, soundBack, kept: Boolean(meetings.get(made.id)) }));

    // With Keep a copy on first, the meeting is one that is recorded as well: the guest is told both before joining,
    // and the switch stays the person's.
    await click('.st-keep-copy input');
    const on = await copyIs({ on: true, greyed: false, says: false }, 'Keep a copy on');
    const made2 = await ask();
    const told2 = await benHeard(`m.type === 'meet.invited' && m.meeting.id === ${JSON.stringify(made2.id || '')} && m.meeting.live === true && m.meeting.recorded === true`, 'Ben asked in, told it is streamed and recorded');
    await seen('.st-guests-btn.on', 'the second meeting held here');
    const kept = await copyIs({ on: true, greyed: false, says: false }, 'Keep a copy still on');
    const noReason = await js(`!${q('.st-guests .st-guests-nocopy')}`);
    check('live broadcast guests: with Keep a copy on before anyone is asked, the meeting is set up as streamed live and recorded, the guest is told both, and the copy stays on',
      on && Boolean(made2.id) && made2.id !== made.id && made2.live === true && made2.recorded === true && told2 && kept && noReason, JSON.stringify({ on, made2, told2, kept, noReason }));
    await endAll();

    await wait(300);
    const complaints = await errorsIn(win);
    const benFaults = await bjs(`window.ben.log.filter((m) => m.type === 'page-error').map((m) => m.message)`);
    check('live broadcast guests: none of that reported an error, in Office or on the guest\'s side', complaints.length === 0 && benFaults.length === 0, [...complaints, ...benFaults].join(' | ') || 'nothing reported');
    await tjs(`window.rutbaOffice.talk.signOut().then(() => 1, () => 1)`);
  } catch (err) {
    check('live broadcast guests: the check ran to the end', false, err?.message || String(err));
  } finally {
    if (js) {
      await js(`window.rutbaOffice.live.stop()`).catch(() => {});
      for (const d of (await js(`window.rutbaOffice.live.destinations()`).catch(() => [])) || []) await js(`window.rutbaOffice.live.removeDestination({ id: ${JSON.stringify(d.id)} })`).catch(() => {});
      if (seeded) await restoreSetup(js, seeded).catch(() => {});
    }
    ingest?.cut();
    ingest?.server.close();
    letGoOfStudioWindows();
    try { ben?.destroy(); } catch { /* gone */ }
    await intro.close().catch(() => {});
  }
}
