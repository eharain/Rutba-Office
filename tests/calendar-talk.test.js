// Talk behind the calendar's own adapter (renderer/apps/calendar-talk.js):
// the people Talk is connected to join the Invite line's suggestions after
// Contacts', each address once, and a contact is asked into Talk through a
// Talk window opened with { inviteEmail, inviteName }. A Talk not installed,
// or one without `connections`, gives nobody and offers no button.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { talkPeople, mergePeople, invitePeople, inviteToTalk, canInviteToTalk, talkableCard, inviteCardToTalk, canMeet, meetingOf, addMeeting, joinMeeting, canJoinMeeting, talkReady, openTalk, meetingGuests } from '../apps/desktop/renderer/apps/calendar-talk.js';
import { liveNotice } from '../apps/desktop/renderer/apps/talk/meet-lines.js';
import { asEnglish } from './label-english.js';

const CONNECTIONS = [
  { id: 't1', name: 'Kim Lee', email: 'kim@example.org', presence: 'online' },
  { id: 't2', name: 'Sam Roy', email: 'SAM@example.org', presence: 'away' },
  { id: 't3', name: 'No Address', email: '', presence: 'online' },
  { id: 't4', name: '', email: 'ola@example.net', presence: null },
];

function fakeShell({ talk = true, connections = true, cards = [] } = {}) {
  const opened = [];
  const shell = {
    win: { create: (o) => { opened.push(o); return Promise.resolve({ id: 1 }); } },
    contacts: { suggest: async ({ query }) => cards.filter((c) => c.email.toLowerCase().includes(query) || (c.name || '').toLowerCase().includes(query)) },
  };
  if (talk) shell.talk = connections ? { connections: async () => CONNECTIONS } : { state: async () => ({}) };
  return { shell, opened };
}

test("Talk's people matching what is typed, by name or address, as Contacts' suggestions are", async () => {
  const { shell } = fakeShell();
  assert.deepEqual(await talkPeople(shell, 'kim'), [{ name: 'Kim Lee', email: 'kim@example.org', source: 'talk', presence: 'online' }]);
  assert.deepEqual((await talkPeople(shell, 'sam@')).map((p) => p.email), ['SAM@example.org']);
  assert.deepEqual((await talkPeople(shell, 'example')).map((p) => p.email), ['kim@example.org', 'SAM@example.org', 'ola@example.net'], 'nobody without an address');
  assert.equal((await talkPeople(shell, 'ola'))[0].name, null);
  assert.equal((await talkPeople(shell, 'example', 2)).length, 2);
});

test('no Talk, a Talk without connections, or one that fails, gives nobody', async () => {
  assert.deepEqual(await talkPeople(fakeShell({ talk: false }).shell, 'kim'), []);
  assert.deepEqual(await talkPeople(fakeShell({ connections: false }).shell, 'kim'), []);
  assert.deepEqual(await talkPeople({ talk: { connections: async () => { throw new Error('signed out'); } } }, 'kim'), []);
  assert.deepEqual(await talkPeople({ talk: { connections: async () => null } }, 'kim'), []);
  assert.deepEqual(await talkPeople(undefined, 'kim'), []);
});

test("the Invite line's suggestions: Contacts' first, then Talk's, each address once", async () => {
  const cards = [{ name: 'Kim Lee (card)', email: 'Kim@Example.org', source: 'card' }, { name: 'Kit', email: 'kit@example.com', source: 'card' }];
  const { shell } = fakeShell({ cards });
  const people = await invitePeople(shell, 'ki');
  assert.deepEqual(people.map((p) => [p.email, p.source]), [['Kim@Example.org', 'card'], ['kit@example.com', 'card']], "Kim from Talk is the card's Kim");
  const both = await invitePeople(shell, 'example');
  assert.deepEqual(both.map((p) => p.source), ['card', 'card', 'talk', 'talk']);
  assert.deepEqual(both.slice(2).map((p) => p.email), ['SAM@example.org', 'ola@example.net']);
  // Without Talk the suggestions are Contacts' alone; without Contacts, Talk's.
  assert.equal((await invitePeople(fakeShell({ talk: false, cards }).shell, 'example')).length, 2);
  assert.equal((await invitePeople({ talk: { connections: async () => CONNECTIONS } }, 'example')).length, 3);
  assert.equal((await invitePeople({ contacts: { suggest: async () => { throw new Error('x'); } }, talk: { connections: async () => CONNECTIONS } }, 'kim')).length, 1);
  assert.deepEqual(mergePeople([{ email: 'a@x' }], [{ email: 'A@X' }, { email: 'b@x' }], 6).map((p) => p.email), ['a@x', 'b@x']);
});

test('a contact asked into Talk opens a Talk window on its invitation, address and name filled', () => {
  const { shell, opened } = fakeShell();
  assert.equal(canInviteToTalk(shell), true);
  assert.equal(inviteToTalk(shell, { email: ' kim@example.org ', name: 'Kim Lee' }), true);
  assert.deepEqual(opened, [{ app: 'talk', query: { inviteEmail: 'kim@example.org', inviteName: 'Kim Lee' } }]);
  inviteToTalk(shell, { email: 'ola@example.net' });
  assert.deepEqual(opened[1], { app: 'talk', query: { inviteEmail: 'ola@example.net' } });
  assert.equal(inviteToTalk(shell, { email: '' }), false);
  assert.equal(opened.length, 2);
});

test('with Talk not installed there is nothing to ask into, and Contacts shows no Talk button', () => {
  const { shell, opened } = fakeShell({ talk: false });
  assert.equal(canInviteToTalk(shell), false);
  assert.equal(inviteToTalk(shell, { email: 'kim@example.org' }), false);
  assert.equal(opened.length, 0);
  const contacts = asEnglish(fs.readFileSync(new URL('../apps/desktop/renderer/apps/contacts.js', import.meta.url), 'utf8'));
  assert.match(contacts, /\{talkHere \? <Button[^\n]*ct-talk-invite/, 'the button is drawn only when Talk is there');
  assert.match(contacts, /const talkable = talkableCard\(selected\)/, 'a room or a resource is not asked into Talk');
});

test('a card is asked into Talk when it is a person with an address, by its first address and the name it shows', () => {
  const kim = { display: 'Kim Lee', kind: 'individual', emails: [{ value: '' }, { value: 'kim@example.org' }, { value: 'kim@home.example' }] };
  assert.equal(talkableCard(kim), true);
  assert.equal(talkableCard({ display: 'No address', emails: [{ value: '' }] }), false);
  assert.equal(talkableCard({ display: 'Board room', kind: 'room', emails: [{ value: 'board@rooms.example' }] }), false);
  assert.equal(talkableCard({ display: 'Projector', kind: 'resource', emails: [{ value: 'beam@rooms.example' }] }), false);
  assert.equal(talkableCard(null), false);
  const { shell, opened } = fakeShell();
  assert.equal(inviteCardToTalk(shell, kim), true);
  assert.deepEqual(opened, [{ app: 'talk', query: { inviteEmail: 'kim@example.org', inviteName: 'Kim Lee' } }]);
  assert.equal(inviteCardToTalk(shell, { display: 'Board room', kind: 'room', emails: [{ value: 'board@rooms.example' }] }), false);
  assert.equal(opened.length, 1);
});

test("the contact list's right-click menu offers Invite to Talk for the card right-clicked, only with Talk installed", () => {
  // Read as LF: a Windows checkout has the file with CRLF, and the patterns below end lines with \n.
  const contacts = asEnglish(fs.readFileSync(new URL('../apps/desktop/renderer/apps/contacts.js', import.meta.url), 'utf8')).replace(/\r\n/g, '\n');
  const menu = /onContextMenu=\{\(e\) => \{([\s\S]*?)\n\s*\}\}/.exec(contacts)?.[1] || '';
  assert.match(menu, /talkHere \? \[\{ label: t\('Invite to Talk'\)/, 'only with Talk installed');
  assert.match(menu, /disabled: !talkableCard\(c\), run: \(\) => talkTo\(c\)/, 'the card right-clicked, not the one chosen before');
  assert.match(menu, /'contact\.mail'\]\), \.\.\.talk, '-'/, 'after Send mail, before Delete');
  // The menu is drawn (it was opened but never mounted), and Edit, Send mail and Delete act on the card right-clicked.
  assert.match(contacts, /\n\s*\{menu\.node\}\n/);
  assert.match(contacts, /run: \(\) => commandsNow\.current\[ids\[i\]\]\?\.run\(\)/);
  assert.doesNotMatch(menu, /menuItems\(commands/);
});

test("only the adapter calls Talk: Calendar and Contacts go through it", () => {
  for (const f of ['calendar.js', 'contacts.js']) {
    const text = asEnglish(fs.readFileSync(new URL(`../apps/desktop/renderer/apps/${f}`, import.meta.url), 'utf8'));
    assert.doesNotMatch(text, /shell\.talk\b|app: 'talk'/, `${f} reaches Talk only through calendar-talk.js`);
    assert.match(text, /from '\.\/calendar-talk\.js'/);
  }
});

/* ── a meeting on an event (Office 2.27) ──────────────────────────── */

const MEET = 'meet-AAAAAAAAAAAAAAAAAAAA';
const LIVE = 'live-BBBBBBBBBBBBBBBBBBBB';
const pageOf = (code) => `https://introducer.rutba.io/m/${code}`;

/** A shell whose Talk sets meetings up as `create` says, noting what it was asked. */
function meetShell(create = ({ live }) => ({ meeting: { id: live ? LIVE : MEET, link: pageOf(live ? LIVE : MEET), live: Boolean(live) } })) {
  const asked = [];
  const opened = [];
  return {
    asked,
    opened,
    shell: {
      win: { create: (o) => { opened.push(o); return Promise.resolve({ id: 1 }); } },
      talk: { meetCreate: async (args) => { asked.push(args); return create(args); } },
    },
  };
}

test('an event can carry a meeting only where Talk holds meetings', () => {
  assert.equal(canMeet(meetShell().shell), true);
  assert.equal(canMeet(fakeShell().shell), false, 'a Talk from before meetings');
  assert.equal(canMeet(fakeShell({ talk: false }).shell), false);
  assert.equal(canMeet(null), false);
  assert.equal(canJoinMeeting(fakeShell().shell), true, 'joining only needs Talk');
  assert.equal(canJoinMeeting(fakeShell({ talk: false }).shell), false);
});

test("the meeting is set up with the event's name and time, and the event takes its link and how to join", async () => {
  const { shell, asked } = meetShell();
  const made = await addMeeting(shell, { name: ' Monday plan ', at: 1_800_000_000_000, location: '', description: 'Bring the figures.' });
  assert.deepEqual(asked, [{ name: 'Monday plan', at: 1_800_000_000_000, live: false }]);
  assert.equal(made.location, pageOf(MEET));
  assert.ok(made.description.startsWith('Bring the figures.\n\n'));
  assert.ok(made.description.includes(pageOf(MEET)));
  assert.ok(!made.description.includes(liveNotice()));
  assert.deepEqual(meetingOf(made), { code: MEET, link: pageOf(MEET), live: false });
});

test('"This meeting will be streamed live" is chosen at set-up and written into the invitation', async () => {
  const { shell, asked } = meetShell();
  const made = await addMeeting(shell, { name: 'Town hall', at: null, live: true, location: 'Main hall', description: '' });
  assert.deepEqual(asked, [{ name: 'Town hall', at: null, live: true }]);
  assert.equal(made.location, 'Main hall', 'a place the person wrote stays');
  assert.ok(made.description.includes(pageOf(LIVE)));
  assert.ok(made.description.includes(liveNotice()), 'the notice goes out with the invitation');
  assert.equal(meetingOf(made).live, true);
  // Anything but a plain yes is no.
  await addMeeting(shell, { name: 'x', live: 'yes' });
  assert.equal(asked[1].live, false);
});

test('a refusal, a Talk that is not there, and a Talk that fails all come back as a reason, never as a half-made event', async () => {
  assert.deepEqual(await addMeeting(meetShell(() => ({ error: 'time' })).shell, { name: 'x' }), { error: 'time' });
  assert.deepEqual(await addMeeting(meetShell(() => ({ meeting: { id: MEET } })).shell, { name: 'x' }), { error: 'offline' }, 'a meeting without a link is no meeting');
  assert.deepEqual(await addMeeting(meetShell(() => { throw new Error('gone'); }).shell, { name: 'x' }), { error: 'offline' });
  assert.deepEqual(await addMeeting(fakeShell().shell, { name: 'x' }), { error: 'no-talk' });
  assert.deepEqual(await addMeeting(fakeShell({ talk: false }).shell, { name: 'x' }), { error: 'no-talk' });
});

test("the meeting an event or an invitation carries is read from its place and notes", () => {
  assert.deepEqual(meetingOf({ location: pageOf(LIVE), description: '' }), { code: LIVE, link: pageOf(LIVE), live: true });
  assert.equal(meetingOf({ location: 'Room 4', description: `Join: ${pageOf(MEET)}` }).code, MEET);
  assert.equal(meetingOf({ location: 'Room 4', description: 'https://example.com/m/not-a-meeting' }), null);
  assert.equal(meetingOf(null), null);
});

test('Join brings Talk forward on the meeting, by its code alone, in the window already open', () => {
  const { shell, opened } = meetShell();
  assert.equal(joinMeeting(shell, 'https://somewhere-else.example/m/' + MEET), true);
  assert.deepEqual(opened, [{ app: 'talk', query: { meet: MEET }, reuse: true }]);
  assert.equal(joinMeeting(shell, 'https://example.com/whatever'), false);
  assert.equal(joinMeeting(fakeShell({ talk: false }).shell, pageOf(MEET)), false, 'no Talk, no window');
  assert.equal(opened.length, 1);
});

test("Calendar's form offers the meeting and the one switch, and the event's card and Mail's invitation card offer Join with the notice", () => {
  const calendar = asEnglish(fs.readFileSync(new URL('../apps/desktop/renderer/apps/calendar.js', import.meta.url), 'utf8')).replace(/\r\n/g, '\n');
  assert.match(calendar, /canMeet\(shell\) \? \(/, 'only where Talk holds meetings');
  assert.match(calendar, /t\('Add a Talk meeting'\)/);
  assert.match(calendar, /t\('This meeting will be streamed live'\)/);
  // The switch is only there while the meeting is being asked for, and is cleared with it.
  assert.match(calendar, /set\(\{ meet: e\.target\.checked, meetLive: e\.target\.checked \? form\.meetLive : false, meetProblem: null \}\)/);
  // The meeting is set up before the event is kept or sent, and a refusal stops both: the form says why.
  const made = /const eventWithMeeting = useCallback\(async \(how\) => \{([\s\S]*?)\n {2}\}, \[form, shell, toast\]\);/.exec(calendar)?.[1] || '';
  assert.match(made, /if \(!form\.meet \|\| how\?\.meeting === false \|\| meetingOf\(event\)\) return event;/);
  assert.match(made, /addMeeting\(shell, \{ name: event\.summary/);
  assert.match(made, /if \(made\.error\) \{\n\s+setForm\(\(f\) => \(f \? \{ \.\.\.f, meetProblem: made\.error \} : f\)\);[\s\S]*?return null;/);
  for (const name of ['saveForm', 'sendForm']) {
    const body = new RegExp(`const ${name} = useCallback\\(async \\((?:how)?\\) => \\{([\\s\\S]*?)\\n {2}\\}, \\[`).exec(calendar)?.[1] || '';
    assert.match(body, /const event = await eventWithMeeting\((?:how)?\);\n\s+if \(!event\) return;[\s\S]*shell\.calendar\.save\(/, `${name} keeps nothing without the meeting asked for`);
  }
  // What went wrong is said beside the meeting, with a way to Talk and a way to save without it.
  assert.match(calendar, /className="cal-meet-problem" role="alert" data-problem=\{form\.meetProblem\}/);
  assert.match(calendar, /className="cal-meet-open-talk" onClick=\{\(\) => openTalk\(shell\)\}/);
  assert.match(calendar, /className="cal-meet-save-without" onClick=\{\(\) => saveForm\(\{ meeting: false \}\)\}/);
  // Who of the people invited the meeting cannot take is said before it is saved.
  assert.match(calendar, /data-warn="too-many"/);
  assert.match(calendar, /data-warn="not-contacts"/);
  // An event that has its meeting says so in place of the tick: the choice is never offered again.
  assert.match(calendar, /meetingOf\(form\) \? \(/);
  assert.match(calendar, /canJoinMeeting\(shell\) \? <Button primary icon="camera" className="cal-join"/);
  assert.ok(calendar.includes(`openedMeeting.live ? <span className="cal-meet-live">{liveNotice({ name: opened.event?.organizer?.name || '' })}</span>`), 'the notice names the organiser where the event does');
  assert.ok(calendar.includes(`meetingOf(form).live ? <span className="cal-meet-live">{liveNotice({ name: form.organizer?.name || '' })}</span>`));

  const card = asEnglish(fs.readFileSync(new URL('../apps/desktop/renderer/apps/calendar-invite-card.js', import.meta.url), 'utf8')).replace(/\r\n/g, '\n');
  assert.match(card, /meeting: meetingOf\(first\)/);
  assert.ok(card.includes(`card.meeting && card.meeting.live && card.method !== 'CANCEL' ? <div className="cal-card-live" data-live="yes">{liveNotice({ name: card.organizer })}</div>`));
  assert.match(card, /card\.method !== 'CANCEL' && card\.method !== 'REPLY' && canJoinMeeting\(shell\) \? <Button icon="camera" className="cal-card-join"/);
  assert.doesNotMatch(card, /shell\.talk\b|app: 'talk'/, 'Mail reaches Talk only through calendar-talk.js');
});

/* ── Talk made ready before a meeting is set up, and who the meeting cannot take (the screen review of 2026-11-22) ── */

/** Talk as Calendar sees it, standing at each of `phases` in turn as it is asked, the last for good. */
function readyShell(phases, { contacts = [], me = null } = {}) {
  let i = 0;
  const calls = [];
  const opened = [];
  const next = () => ({ phase: phases[Math.min(i++, phases.length - 1)], contacts, me });
  const talk = {
    resume: async () => { calls.push('resume'); return next(); },
    state: async () => { calls.push('state'); return next(); },
    meetCreate: async () => { calls.push('meetCreate'); return { meeting: { id: MEET, link: pageOf(MEET), live: false } }; },
  };
  return { shell: { talk, win: { create: (o) => { opened.push(o); return Promise.resolve({ id: 1 }); } } }, calls, opened };
}
/** A wait that takes no time: a second in all, asked every quarter. */
const quick = { wait: 1000, step: 250, sleep: async () => {} };

test('Talk is connected first when it is signed in but not connected, with no Talk window open', async () => {
  const { shell, calls } = readyShell(['connecting', 'connecting', 'online']);
  assert.equal(await talkReady(shell, quick), 'online');
  assert.deepEqual(calls, ['resume', 'state', 'state']);
  const already = readyShell(['online']);
  assert.equal(await talkReady(already.shell, quick), 'online');
  assert.deepEqual(already.calls, ['resume']);
  assert.equal(await talkReady(readyShell(['offline', 'connecting', 'online']).shell, quick), 'online', 'a dropped connection that comes back in time');
});

test('nobody signed in to Talk, and a Talk that cannot connect, are told apart', async () => {
  assert.equal(await talkReady(readyShell(['signed-out']).shell, quick), 'signed-out');
  assert.equal(await talkReady(readyShell(['signing-in']).shell, quick), 'signed-out');
  const stuck = readyShell(['connecting']);
  assert.equal(await talkReady(stuck.shell, quick), 'offline');
  assert.deepEqual(stuck.calls, ['resume', 'state', 'state', 'state', 'state'], 'asked for as long as the wait, then given up');
  assert.equal(await talkReady({ talk: { resume: async () => { throw new Error('gone'); }, state: async () => ({ phase: 'online' }) } }, quick), 'offline');
  assert.equal(await talkReady({ talk: { resume: async () => null, state: async () => null } }, quick), 'offline');
  // An earlier Talk that does not say how it stands is asked as before.
  assert.equal(await talkReady(meetShell().shell, quick), 'online');
});

test('a meeting is not asked for until Talk is connected, and the reason comes back in place of a half-made event', async () => {
  const signedOut = readyShell(['signed-out']);
  assert.deepEqual(await addMeeting(signedOut.shell, { name: 'x' }, quick), { error: 'signed-out' });
  const off = readyShell(['connecting']);
  assert.deepEqual(await addMeeting(off.shell, { name: 'x' }, quick), { error: 'offline' });
  assert.ok(!signedOut.calls.includes('meetCreate') && !off.calls.includes('meetCreate'));
  const late = readyShell(['connecting', 'online']);
  const made = await addMeeting(late.shell, { name: 'Monday plan' }, quick);
  assert.equal(made.location, pageOf(MEET));
  assert.deepEqual(late.calls, ['resume', 'state', 'meetCreate']);
});

test('Talk is brought forward from Calendar, in the window already open, to sign in or to see why it is not connected', () => {
  const { shell, opened } = readyShell(['signed-out']);
  assert.equal(openTalk(shell), true);
  assert.deepEqual(opened, [{ app: 'talk', reuse: true }]);
  assert.equal(openTalk(fakeShell({ talk: false }).shell), false);
});

test('who of the people invited a Talk meeting cannot take: more than three, and those who are not Talk contacts', async () => {
  const contacts = [{ id: 'u1', email: 'Kim@Example.org' }, { id: 'u2', email: 'sam@example.org' }, { id: 'u3', email: '' }];
  const { shell } = readyShell(['online'], { contacts, me: { id: 'me', email: 'me@example.org' } });
  assert.deepEqual(await meetingGuests(shell, ['kim@example.org', 'SAM@example.org']), { count: 2, tooMany: false, strangers: [] });
  assert.deepEqual(await meetingGuests(shell, ['kim@example.org', 'ola@example.net', 'me@example.org', 'ola@example.net', '', 'not an address']), { count: 3, tooMany: false, strangers: ['ola@example.net'] }, 'each address once, and the person themself is no stranger');
  const four = await meetingGuests(shell, ['kim@example.org', 'sam@example.org', 'a@example.org', 'b@example.org']);
  assert.deepEqual(four, { count: 4, tooMany: true, strangers: ['a@example.org', 'b@example.org'] });
  // Talk not connected cannot say who is a contact: the count is still said, and nobody is called a stranger.
  assert.deepEqual(await meetingGuests(readyShell(['offline'], { contacts }).shell, ['kim@example.org', 'x@example.org']), { count: 2, tooMany: false, strangers: null });
  assert.deepEqual(await meetingGuests({ talk: { state: async () => { throw new Error('gone'); } } }, ['x@example.org']), { count: 1, tooMany: false, strangers: null });
  assert.deepEqual(await meetingGuests({}, []), { count: 0, tooMany: false, strangers: null });
});
