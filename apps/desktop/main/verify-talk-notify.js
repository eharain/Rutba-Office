// Talk's notifications, seen with no Talk window open (Office 2.27.0,
// D-2026-11-22-23): a message, a call and a meeting someone asks the person
// into each make a notification while Talk has no window, and pressing one
// opens Talk at the thing.
//
// A check run shows nothing by the clock and no Windows notification: what
// would be shown is recorded instead (office-shell's notify-electron.js,
// globalThis.__rutbaNotifications), each with the two things a person can do
// to it, press it and press one of its buttons. So this sees the whole way
// from the introducer to the notification and back to a window; the
// notification as Windows draws it, and the icon, are not seen here.
//
// The other side is Ben, a plain connection to the stand-in introducer
// (verify-talk.js, with the meetings of verify-talk-meet.js): no picture or
// sound is needed to ring.

import { BrowserWindow } from 'electron';
import { standInIntroducer } from './verify-talk.js';
import { standInMeetings } from './verify-talk-meet.js';

/** The windows of an app that are open now. */
const windowsOf = (app) => BrowserWindow.getAllWindows().filter((w) => {
  if (w.isDestroyed()) return false;
  try { return new URL(w.webContents.getURL()).searchParams.get('app') === app; } catch { return false; }
});

export async function verifyTalkNotify(h) {
  const { open, check, until, wait, errorsIn } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  const notes = () => globalThis.__rutbaNotifications || [];
  const noteOf = (test) => notes().filter(test).at(-1) || null;
  // Each ring the notifier would have sounded for a call no window rings for: { at, stopped }. A check run is silent.
  const rings = () => globalThis.__rutbaRings || [];
  const ringsAtStart = rings().length;
  const meetings = standInMeetings();
  const intro = standInIntroducer({ more: meetings.more, contacts: [['u-amira', 'u-ben']] });
  let ben = null;
  let other = null;
  const opened = [];
  try {
    const port = await intro.listen();
    const url = `ws://127.0.0.1:${port}/ws`;

    // Another window stays open throughout, so Talk's can close with Office still there; it also asks for the settings.
    // Talk windows left by the checks before this one are done with: this one is about there being none.
    const left = windowsOf('talk');
    other = await open('word');
    for (const w of left) w.destroy();
    const ojs = (code) => other.webContents.executeJavaScript(code);
    await ojs(`window.rutbaOffice.notify.resume().then(() => 1, () => 1)`);

    // Amira signs in with a test server in Talk's window; Ben is her contact already.
    const first = await open('talk');
    const fjs = (code) => first.webContents.executeJavaScript(code);
    const set = (cls, value) => fjs(`(() => { const n = document.querySelector('.${cls}'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, ${JSON.stringify(value)}); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await fjs(`window.rutbaOffice.talk.signOut().then(() => 1, () => 1)`);
    await until(() => fjs(`Boolean(document.querySelector('.tk-signin'))`), 'the sign-in', 8000).catch(() => {});
    await fjs(`document.querySelector('.tk-test')?.setAttribute('open', ''), 1`);
    await set('tk-server', url);
    await set('tk-token', 'amira');
    await fjs(`(document.querySelector('.tk-connect')?.click(), 1)`);
    const signedIn = await ok(until(() => fjs(`Boolean(document.querySelector('.tk-contact[data-contact="u-ben"]'))`), 'signed in, Ben among the contacts', 10000));

    // Ben, on the other side.
    const heard = [];
    ben = new WebSocket(url);
    await new Promise((resolve, reject) => { ben.onopen = resolve; ben.onerror = () => reject(new Error('Ben could not connect')); });
    ben.onmessage = (e) => { try { heard.push(JSON.parse(e.data)); } catch { /* not for this check */ } };
    const say = (m) => ben.send(JSON.stringify(m));
    const benHeard = (test, what, ms = 8000) => ok(until(() => heard.some(test) || null, what, ms));
    say({ type: 'hello', token: 'ben' });
    await benHeard((m) => m.type === 'welcome', 'Ben signed in');

    // Talk's window closes. Office goes on with Talk connected: Ben still sees Amira there to be called.
    first.close();
    const closed = await ok(until(() => windowsOf('talk').length === 0 || null, 'no Talk window', 8000));

    // The test notification of the settings, and what the settings say.
    const before = notes().length;
    const tested = await ojs(`window.rutbaOffice.notify.test()`);
    const state = await ojs(`window.rutbaOffice.notify.state()`);
    const testNote = notes().slice(before).at(-1) || null;
    check('notifications: with every Talk window closed the settings still answer, and their test shows one notification',
      signedIn && closed && tested?.shown === true && Boolean(testNote) && Boolean(testNote.title) && Boolean(testNote.body) && testNote.actions.length === 0 && state.supported === true && state.paused === false,
      JSON.stringify({ signedIn, closed, tested, test: testNote && { kind: testNote.kind, title: testNote.title }, supported: state.supported, paused: state.paused }));

    // A message: one notification with who wrote and the words; a second from the same person replaces it and counts.
    say({ type: 'msg.send', to: 'u-amira', id: 'n-msg-1', text: 'Are you there?' });
    const told1 = await ok(until(() => noteOf((n) => n.id === 'chat:u-ben' && n.body === 'Are you there?'), 'the message notified', 8000));
    say({ type: 'msg.send', to: 'u-amira', id: 'n-msg-2', text: 'The agenda is ready.' });
    const told2 = await ok(until(() => noteOf((n) => n.id === 'chat:u-ben' && n.body.startsWith('The agenda is ready.') && n.body !== 'The agenda is ready.'), 'the second, counted', 8000));
    const chat = noteOf((n) => n.id === 'chat:u-ben');
    const waiting = (await ojs(`window.rutbaOffice.notify.state()`)).waiting;
    // Pressed: Talk opens at Ben's conversation, with what he wrote.
    chat?.click();
    const atBen = await ok(until(() => windowsOf('talk')[0] || null, 'a Talk window', 10000));
    const talk1 = windowsOf('talk')[0] || null;
    if (talk1) opened.push(talk1);
    // The last two he wrote: in a run of all of Talk's checks the conversation also holds what he wrote in the checks before this one.
    const read = Boolean(talk1) && await ok(until(() => talk1.webContents.executeJavaScript(`[...document.querySelectorAll('.tk-bubble:not(.mine) .tk-text')].map((n) => n.textContent).slice(-2).join('|') === 'Are you there?|The agenda is ready.'`), 'Ben\'s messages in the window', 10000));
    check('notifications: a message that comes with no Talk window open is notified with who wrote it and its words, a second from them takes its place and is counted, and pressing it opens Talk at their conversation',
      told1 && told2 && chat?.title === 'Ben Ortiz' && chat?.actions.length === 0 && chat?.sticky === false && waiting?.chat === 2 && atBen && read,
      JSON.stringify({ told1, told2, title: chat?.title, body: chat?.body, waiting, atBen, read }));

    // A call: a notification that stays, with Answer and Decline. Decline there turns the call away and takes it down.
    talk1?.close();
    await until(() => windowsOf('talk').length === 0 || null, 'no Talk window again', 8000).catch(() => {});
    say({ type: 'call.offer', to: 'u-amira', callId: 'n-call-1', sdp: 'v=0', video: false });
    const rang = await ok(until(() => noteOf((n) => n.id === 'call:n-call-1'), 'the call notified', 8000));
    const call1 = noteOf((n) => n.id === 'call:n-call-1');
    // No window is there to ring, so the notifier rings: one ring, going on, and the notification silent under it.
    const ring1 = rings().at(-1) || null;
    const ringing1 = rings().length === ringsAtStart + 1 && ring1?.stopped === false && call1?.silent === true;
    call1?.act('decline');
    const declined = await benHeard((m) => m.type === 'call.declined' && m.callId === 'n-call-1', 'the call turned away at Ben');
    const down1 = await ok(until(() => call1?.closed === true || null, 'its notification taken down', 6000));
    const noWindow = windowsOf('talk').length === 0;
    const hushed1 = ring1?.stopped === true;
    check('notifications: a call that comes with no Talk window open is notified until it is dealt with, with Answer and Decline; Decline there turns the call away, takes the notification down and opens nothing',
      rang && Boolean(call1?.title) && call1?.body === 'Ben Ortiz' && call1?.sticky === true && call1?.actions.join() === 'answer,decline' && declined && down1 && noWindow,
      JSON.stringify({ rang, title: call1?.title, body: call1?.body, sticky: call1?.sticky, actions: call1?.actions, declined, down1, noWindow }));
    check('notifications: with no Talk window open the call rings by the clock, its notification silent under the ring, and Decline stops the ring',
      ringing1 && hushed1, JSON.stringify({ rings: rings().length - ringsAtStart, ringing1, silent: call1?.silent, hushed1 }));

    // Answer there opens Talk at the caller with the call still ringing in its window; the caller giving up leaves a missed call.
    say({ type: 'call.offer', to: 'u-amira', callId: 'n-call-2', sdp: 'v=0', video: true });
    await until(() => noteOf((n) => n.id === 'call:n-call-2'), 'the second call notified', 8000).catch(() => {});
    const call2 = noteOf((n) => n.id === 'call:n-call-2');
    const ring2 = rings().at(-1) || null;
    const ringing2 = rings().length === ringsAtStart + 2 && ring2 !== ring1 && ring2?.stopped === false;
    call2?.act('answer');
    await until(() => windowsOf('talk')[0] || null, 'a Talk window for the call', 10000).catch(() => {});
    const talk2 = windowsOf('talk')[0] || null;
    if (talk2) opened.push(talk2);
    const tjs = (code) => talk2.webContents.executeJavaScript(code);
    const ringing = Boolean(talk2) && await ok(until(() => tjs(`Boolean(document.querySelector('.tk-ring .tk-answer-video'))`), 'the call ringing in the window', 10000));
    const notAnswered = !heard.some((m) => m.type === 'call.answered' && m.callId === 'n-call-2');
    const mark = notes().length;
    say({ type: 'call.end', callId: 'n-call-2' });
    const stopped = Boolean(talk2) && await ok(until(() => tjs(`!document.querySelector('.tk-ring')`), 'the ringing stopped', 8000));
    const down2 = await ok(until(() => call2?.closed === true || null, 'the call\'s notification taken down', 6000));
    // Missed calls are counted by who called, so the count by the clock need not move: the notification is what is new.
    const missed = await ok(until(() => notes().slice(mark).find((n) => n.id === 'missed:u-ben' && n.body === 'Ben Ortiz') || null, 'the missed call notified', 6000));
    const waits = (await ojs(`window.rutbaOffice.talk.state()`)).missed;
    check('notifications: Answer in a call\'s notification opens Talk at the caller with the call still ringing there and answers nothing by itself; when the caller gives up the ringing stops, the notification comes down and a missed call is notified in its place',
      Boolean(call2) && ringing && notAnswered && stopped && down2 && missed, JSON.stringify({ notified: Boolean(call2), ringing, notAnswered, stopped, down2, missed, waits }));

    // The ring by the clock stopped as Answer was pressed, and the window rang from there. A call that comes
    // with the Talk window open rings in the window and makes no ring by the clock: never both.
    const hushed2 = ring2?.stopped === true;
    say({ type: 'call.offer', to: 'u-amira', callId: 'n-call-3', sdp: 'v=0', video: false });
    const ringing3 = Boolean(talk2) && await ok(until(() => tjs(`Boolean(document.querySelector('.tk-ring'))`), 'the third call ringing in the window', 10000));
    const noRing3 = rings().length === ringsAtStart + 2;
    const call3 = noteOf((n) => n.id === 'call:n-call-3');
    say({ type: 'call.end', callId: 'n-call-3' });
    if (talk2) await until(() => tjs(`!document.querySelector('.tk-ring')`), 'the third call over', 8000).catch(() => {});
    check('notifications: the ring by the clock stops when Answer opens Talk, and a call that comes with a Talk window open rings in the window and never by the clock as well',
      ringing2 && hushed2 && ringing3 && noRing3 && call3?.silent !== true, JSON.stringify({ ringing2, hushed2, ringing3, noRing3, rings: rings().length - ringsAtStart, notified3: Boolean(call3), silent3: call3?.silent ?? null }));

    // A meeting someone asks her into: a notification with Join, which opens Talk at the meeting before joining.
    talk2?.close();
    await until(() => windowsOf('talk').length === 0 || null, 'no Talk window once more', 8000).catch(() => {});
    say({ type: 'meet.create', name: 'Planning', live: false, invite: ['u-amira'], tell: true, ref: 'n-meet' });
    await benHeard((m) => m.type === 'meet.created' && m.ref === 'n-meet', 'the meeting set up');
    const meetId = heard.find((m) => m.type === 'meet.created' && m.ref === 'n-meet')?.meeting?.id || '';
    const asked = await ok(until(() => noteOf((n) => n.id === `meet:${meetId}`), 'the meeting notified', 8000));
    const meet = noteOf((n) => n.id === `meet:${meetId}`);
    meet?.act('join');
    await until(() => windowsOf('talk')[0] || null, 'a Talk window for the meeting', 10000).catch(() => {});
    const talk3 = windowsOf('talk')[0] || null;
    if (talk3) opened.push(talk3);
    const atDoor = Boolean(talk3) && await ok(until(() => talk3.webContents.executeJavaScript(`Boolean(document.querySelector('.tk-meet-seen[data-meeting=${JSON.stringify(meetId)}][data-can="yes"]')) && !document.querySelector('.tk-meeting')`), 'the meeting before joining', 10000));
    const notIn = !meetings.inside(meetId).includes('u-amira');
    say({ type: 'meet.end', id: meetId, ref: 'n-end' });
    await benHeard((m) => m.type === 'meet.ended' && m.ref === 'n-end', 'the meeting ended');
    check('notifications: a meeting she is asked into with no Talk window open is notified with Join, which opens Talk at the meeting before joining and joins nothing by itself',
      Boolean(meetId) && asked && meet?.actions.join() === 'join' && atDoor && notIn, JSON.stringify({ meetId, asked, actions: meet?.actions, body: meet?.body, atDoor, notIn, takenDown: meet?.closed }));

    // Paused: what comes in is counted and not shown; resumed, the next is shown again.
    talk3?.close();
    await until(() => windowsOf('talk').length === 0 || null, 'no Talk window at the end', 8000).catch(() => {});
    const paused = await ojs(`window.rutbaOffice.notify.pause({ minutes: 60 })`);
    say({ type: 'msg.send', to: 'u-amira', id: 'n-msg-3', text: 'Sent while paused.' });
    await until(() => ojs(`window.rutbaOffice.talk.state().then((s) => (s.unread?.['u-ben'] || 0) >= 1)`), 'the paused message kept', 8000).catch(() => {});
    const resumed = await ojs(`window.rutbaOffice.notify.resume()`);
    say({ type: 'msg.send', to: 'u-amira', id: 'n-msg-4', text: 'Sent after.' });
    const again = await ok(until(() => noteOf((n) => n.id === 'chat:u-ben' && n.body.startsWith('Sent after.')), 'notified again', 8000));
    const quiet = !notes().some((n) => n.body.startsWith('Sent while paused.'));
    check('notifications: paused, a message is kept and not notified; resumed, the next one is notified',
      paused?.paused === true && resumed?.paused === false && again && quiet, JSON.stringify({ paused: paused?.paused, resumed: resumed?.paused, again, quiet }));

    await wait(300);
    const complaints = [...await errorsIn(other)];
    check('notifications: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    await ojs(`window.rutbaOffice.talk.signOut().then(() => 1, () => 1)`);
  } catch (err) {
    check('notifications: the check ran to the end', false, err?.message || String(err));
  } finally {
    try { ben?.close(); } catch { /* gone */ }
    for (const w of opened) { try { if (!w.isDestroyed()) w.destroy(); } catch { /* gone */ } }
    try { other?.destroy(); } catch { /* gone */ }
    await intro.close().catch(() => {});
  }
}
