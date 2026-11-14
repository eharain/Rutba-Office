// Talk behind the calendar's own adapter (renderer/apps/calendar-talk.js):
// the people Talk is connected to join the Invite line's suggestions after
// Contacts', each address once, and a contact is asked into Talk through a
// Talk window opened with { inviteEmail, inviteName }. A Talk not installed,
// or one without `connections`, gives nobody and offers no button.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { talkPeople, mergePeople, invitePeople, inviteToTalk, canInviteToTalk, talkableCard, inviteCardToTalk } from '../apps/desktop/renderer/apps/calendar-talk.js';
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
