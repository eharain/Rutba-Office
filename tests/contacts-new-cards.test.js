/**
 * Where new cards go, across the update to 2.19. A contacts store kept from
 * before 2.19 keeps new cards on this computer (or in the book chosen then),
 * recorded once as 'local', sealed or not, so an account added later never
 * starts taking them to its server unasked. A store 2.19 makes sends new
 * cards to the first address book of the account added, until the person
 * chooses; and a book that goes never hands them to another account's.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createCalendarService } from '../apps/desktop/main/calendar.js';
import { createContactsService } from '../apps/desktop/main/contacts.js';
import { createDavService } from '../apps/desktop/main/dav-sync.js';
import { createVault, makeLock } from '../packages/office-shell/src/profiles.js';
import { startDavServer } from './fixtures/dav-server.js';

function memory() {
  const map = new Map();
  return { get: (k) => map.get(k), set: (k, v) => { map.set(k, v); }, delete: (k) => { map.delete(k); } };
}

const tmp = () => fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-newcards-'));
// The store is written a moment after the step: wait for it there, not for a set time (a busy machine takes longer).
async function written(read, ms = 10000) {
  const end = Date.now() + ms;
  for (;;) {
    let got = null;
    try { got = read(); } catch { /* not there yet, or half written */ }
    if (got || Date.now() > end) return got;
    await new Promise((r) => setTimeout(r, 100));
  }
}
const card = (name, email) => ({ name: { full: name, given: name.split(' ')[0], family: name.split(' ')[1] || '' }, emails: [{ value: email }], phones: [], addresses: [] });

/** A contacts store as 2.18 left it: cards, and no choice of where new ones go. */
const older = () => ({ contacts: [{ id: 'c1', name: { full: 'Kept Here', given: 'Kept', family: 'Here' }, emails: [{ value: 'kept@example.com' }], createdAt: 1, updatedAt: 1 }] });

function services(dir, vault = null) {
  const stores = { dir, settings: memory(), secrets: memory(), ...(vault ? { vault } : {}) };
  const calendar = createCalendarService({ stores, broadcast: null });
  const contacts = createContactsService({ stores, broadcast: null });
  const dav = createDavService({ stores, calendar, contacts, timers: false });
  return { contacts, dav };
}

test('A store from before 2.19 with no choice keeps new cards on this computer, recorded once, even after an account is added', async () => {
  const server = await startDavServer();
  const dir = tmp();
  fs.writeFileSync(path.join(dir, 'contacts.json'), JSON.stringify(older()));
  try {
    const { contacts, dav } = services(dir);
    assert.equal(contacts._dav.state().defaultBook, 'local', 'recorded as this computer at the first start');
    const onDisk = await written(() => { const j = JSON.parse(fs.readFileSync(path.join(dir, 'contacts.json'), 'utf8')); return j.newCardsSettled ? j : null; });
    assert.equal(onDisk.defaultBook, 'local');
    assert.equal(onDisk.newCardsSettled, 1);
    assert.equal(onDisk.contacts.length, 1, 'the cards are untouched');

    await dav.add({ url: server.url, user: 'ann', password: 'secret' });
    assert.ok(dav.books().books.length, 'the account brought its address book');
    assert.equal(dav.books().defaultBook, null, 'new cards still stay here');
    const made = contacts.save({ contact: card('Ann Novak', 'ann@example.com') });
    await dav.sync();
    const sent = [...server.collections.get(server.bookPath('friends')).items.values()].map((i) => i.data).join('\n');
    assert.doesNotMatch(sent, /Ann Novak/, 'the new card never went to the server');
    assert.equal(contacts.get({ id: made.id }).book, undefined);

    // The person chooses the book: from then on, new cards go there.
    const [{ id }] = dav.books().books;
    dav.setDefaultBook({ id });
    assert.equal(dav.books().defaultBook, id);
    dav.stop();
  } finally {
    await server.close();
  }
});

test('The step runs once: a store already settled is left as the person set it', async () => {
  const dir = tmp();
  fs.writeFileSync(path.join(dir, 'contacts.json'), JSON.stringify({ ...older(), newCardsSettled: 1, defaultBook: null, books: [] }));
  const { contacts } = services(dir);
  assert.equal(contacts._dav.state().defaultBook, null);
});

test('A store from before 2.19 with a book chosen keeps that book', () => {
  const dir = tmp();
  const books = [{ id: 'b1', account: 'a1', href: '/dav/ann/friends/', name: 'Friends', ctag: null }];
  fs.writeFileSync(path.join(dir, 'contacts.json'), JSON.stringify({ ...older(), books, defaultBook: 'b1' }));
  const { contacts } = services(dir);
  assert.equal(contacts._dav.state().defaultBook, 'b1');
  assert.equal(contacts._dav.state().newCardsSettled, 1);
});

test('A sealed store from before 2.19 is recorded through the vault and stays sealed', async () => {
  const dir = tmp();
  const file = path.join(dir, 'contacts.json');
  const { key } = await makeLock('harbour-7');
  const sealer = createVault({ key });
  sealer.cover(file);
  sealer.writeJson(file, older());
  assert.doesNotMatch(fs.readFileSync(file, 'utf8'), /Kept Here/, 'sealed to begin with');

  const { contacts } = services(dir, createVault({ key }));
  assert.equal(contacts._dav.state().defaultBook, 'local');
  assert.equal(contacts.list().length, 1);
  const reopened = createVault({ key });
  reopened.cover(file);
  const again = await written(() => { const j = reopened.readJson(file, null); return j?.newCardsSettled ? j : null; });
  assert.doesNotMatch(fs.readFileSync(file, 'utf8'), /Kept Here|newCardsSettled/, 'still sealed after the step');
  assert.equal(again.defaultBook, 'local');
  assert.equal(again.newCardsSettled, 1);
});

test('A store 2.19 makes sends new cards to the first book of an account added, and never to another account\'s when that one goes', async () => {
  const server = await startDavServer();
  try {
    const { contacts, dav } = services(tmp());
    assert.equal(dav.books().defaultBook, null);
    const first = await dav.add({ url: server.url, user: 'ann', password: 'secret', name: 'Family' });
    const ofFirst = dav.books().books.find((b) => b.account === first.id);
    assert.equal(dav.books().defaultBook, ofFirst.id, 'an account added now takes new cards into its book');
    const made = contacts.save({ contact: card('Bo Lund', 'bo@example.com') });
    assert.equal(contacts.get({ id: made.id }).book, ofFirst.id);

    // A second account brings its own book; the first stays where new cards go.
    const second = await dav.add({ url: server.url, user: 'ann', password: 'secret', name: 'Work' });
    assert.equal(dav.books().defaultBook, ofFirst.id, 'a second account does not take the choice over');
    dav.remove({ id: first.id });
    assert.equal(dav.books().defaultBook, null, 'its book gone, new cards stay here, not in the other account\'s book');
    assert.ok(dav.books().books.some((b) => b.account === second.id));
    await dav.sync();
    assert.equal(dav.books().defaultBook, null, 'and a sync does not pick the other book either');
    dav.stop();
  } finally {
    await server.close();
  }
});
