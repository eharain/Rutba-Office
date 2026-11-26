// A saved draft reaches the account's server, and Archive on Gmail goes to
// All Mail (2.28, B19 and B20).
//
// Against the test IMAP server (packages/mailbox/test/imap-server.js), run
// in this process. Before 2.28 Save draft kept a draft on this computer only,
// in a folder of Mail's own called Drafts, so a phone or webmail never had
// it; and Archive on a Gmail account, which has no \Archive folder, made a
// label called "Archive" there. Here: a draft is APPENDed to the folder the
// server calls Drafts, marked \Draft; saved again, one copy stays; sent or
// deleted, the server's copy goes; with the server away it waits and goes
// when the server is back; a server that says no loses nothing typed; a draft
// written in another mail program opens with its Bcc, is saved over and sent
// without leaving a copy behind; a server with no Drafts folder is given one;
// and Archive finds the account's own folder (\Archive, or Gmail's \All)
// without ever making a label on Gmail.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createMailService } from '../apps/desktop/main/mail.js';
import { imapServer } from '../packages/mailbox/test/imap-server.js';

function fakeStores(settings = new Map(), secrets = new Map()) {
  return {
    settings: {
      get: (k, fallback) => (settings.has(k) ? structuredClone(settings.get(k)) : fallback),
      set: (k, v) => settings.set(k, structuredClone(v)),
      delete: (k) => settings.delete(k),
      all: () => Object.fromEntries(settings),
    },
    secrets: {
      available: () => true,
      get: (k) => secrets.get(k) ?? null,
      set: (k, v) => secrets.set(k, v),
      delete: (k) => secrets.delete(k),
      keys: () => [...secrets.keys()],
    },
  };
}

const inboxMessage = (n) => [
  'From: Amina Yusuf <amina@work.example>', 'To: Sam <sam@checks.example>', `Subject: Message ${n}`,
  `Date: ${new Date(Date.UTC(2026, 9, 1, 9, n)).toUTCString()}`, `Message-ID: <m${n}@work.example>`, '', `Body of message ${n}.`, '',
].join('\r\n');

/** Gmail's folders as its IMAP lists them, All Mail among them. */
const GMAIL = [['INBOX', null], ['[Gmail]/All Mail', '\\All'], ['[Gmail]/Sent Mail', '\\Sent'], ['[Gmail]/Drafts', '\\Drafts'], ['[Gmail]/Bin', '\\Trash'], ['[Gmail]/Spam', '\\Junk']];

/** A server, a service with its account fetched once, and a way to make another service on the same folder (a restart, or a second look). */
async function setUp({ capabilities = {}, folders = null, fetch = true } = {}) {
  const server = await imapServer({ capabilities, folders });
  for (let n = 1; n <= 2; n++) server.deliver('INBOX', inboxMessage(n));
  const userData = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-maildrafts-'));
  const stores = fakeStores();
  const heard = [];
  const make = (dir = userData, st = stores) => createMailService({
    stores: st,
    holdBlob: () => ({ url: '' }),
    broadcast: (event, payload) => heard.push({ event, payload }),
    userData: dir,
    smtp: { async send() { return { messageId: '<ignored@smtp>' }; } },
  });
  const service = make();
  const { id: accountId } = service.addAccount({ account: server.account(), password: 'check' });
  if (fetch) await service.sync({ accountId });
  const rows = (folder, svc = service) => svc.messages({ accountId, folder, limit: 100 }).rows;
  const at = (folder) => server.messages(folder);
  const draft = (words, extra = {}) => ({ to: 'Amina Yusuf <amina@work.example>', subject: 'Plan for Friday', text: words, ...extra });
  /** Save draft, then the line sent now (the window's save waits a moment and sends it itself). */
  const save = async (d, replaces = null, svc = service) => {
    const { id } = await svc.saveDraft({ accountId, draft: d, replaces });
    await svc.flushChanges({ accountId });
    return id;
  };
  /** The id a draft is known by, as the window reads it off the message it opened. */
  const idOf = (folder, id, svc = service) => String(svc.message({ accountId, folder, id }).messageId).slice(6);
  return { server, service, make, stores, userData, accountId, heard, rows, at, draft, save, idOf, close: () => server.close() };
}

test('a saved draft is kept in the Drafts folder of the server, marked as a draft, with its Bcc', async () => {
  const s = await setUp();
  try {
    const id = await s.save(s.draft('First words.', { cc: 'kofi@work.example', bcc: 'hidden@work.example' }));
    const there = s.at('Drafts');
    assert.equal(there.length, 1, 'one copy at the server');
    assert.ok(there[0].flags.includes('\\Draft'), `marked as a draft: ${there[0].flags.join(' ')}`);
    assert.ok(there[0].flags.includes('\\Seen'));
    assert.match(there[0].raw, /First words\./);
    assert.match(there[0].raw, /^Bcc: hidden@work\.example/im, 'the server copy keeps who was on Bcc');
    assert.match(there[0].messageId, /^<draft-[0-9a-f-]+@checks\.example>$/, there[0].messageId);
    // Here it sits in the same folder, with the number the server gave it, so the next fetch does not double it.
    const here = s.rows('Drafts');
    assert.equal(here.length, 1);
    assert.equal(here[0].id, id);
    assert.equal(here[0].uid, there[0].uid);
    await s.service.sync({ accountId: s.accountId, folder: 'Drafts' });
    assert.equal(s.rows('Drafts').length, 1, 'fetched again, still one');
    assert.equal(s.service.message({ accountId: s.accountId, folder: 'Drafts', id }).draftBcc, 'hidden@work.example');
    assert.deepEqual(s.heard.filter((h) => h.event === 'mail:syncProblem'), []);
  } finally {
    await s.close();
  }
});

test('a draft saved again leaves one copy at the server, the newer one', async () => {
  const s = await setUp();
  try {
    const id = await s.save(s.draft('First words.'));
    const first = s.at('Drafts')[0].uid;
    // Opened from Drafts and saved: the window names the draft it came from.
    const again = await s.save(s.draft('Second words.', { id: s.idOf('Drafts', id) }), { accountId: s.accountId, folder: 'Drafts', id });
    assert.equal(again, id, 'the same draft');
    const there = s.at('Drafts');
    assert.equal(there.length, 1, there.map((m) => m.uid).join(','));
    assert.notEqual(there[0].uid, first);
    assert.match(there[0].raw, /Second words\./);
    assert.equal(s.rows('Drafts').length, 1);
    assert.equal(s.rows('Drafts')[0].uid, there[0].uid);
    // The new copy went up before the old one came down.
    const said = s.server.heard.filter((l) => /APPEND|EXPUNGE/.test(l));
    const lastAppend = said.map((l) => / APPEND /.test(l)).lastIndexOf(true);
    const lastExpunge = said.map((l) => /EXPUNGE/.test(l)).lastIndexOf(true);
    assert.ok(lastAppend >= 0 && lastExpunge > lastAppend, said.join(' | '));
    await s.service.sync({ accountId: s.accountId, folder: 'Drafts' });
    assert.equal(s.rows('Drafts').length, 1);
    assert.match(s.service.message({ accountId: s.accountId, folder: 'Drafts', id }).text, /Second words\./);
  } finally {
    await s.close();
  }
});

test('a draft sent or deleted goes from the server too', async () => {
  const s = await setUp();
  try {
    const id = await s.save(s.draft('To be sent.'));
    assert.equal(s.at('Drafts').length, 1);
    // What the window does once a draft has been sent: the draft it was opened from is deleted for good.
    s.service.delete({ accountId: s.accountId, folder: 'Drafts', ids: [id] });
    await s.service.flushChanges({ accountId: s.accountId });
    assert.equal(s.at('Drafts').length, 0, 'gone at the server');
    assert.equal(s.rows('Drafts').length, 0);

    // One sent before the server ever heard of it is never put there.
    const { id: quick } = await s.service.saveDraft({ accountId: s.accountId, draft: s.draft('Sent at once.') });
    s.service.delete({ accountId: s.accountId, folder: 'Drafts', ids: [quick] });
    await s.service.flushChanges({ accountId: s.accountId });
    assert.equal(s.at('Drafts').length, 0);
    assert.equal(s.rows('Drafts').length, 0);

    // One the person deletes goes to the Trash, at the server as here.
    const binned = await s.save(s.draft('Thrown away.'));
    s.service.delete({ accountId: s.accountId, folder: 'Drafts', ids: [binned], toTrash: true });
    await s.service.flushChanges({ accountId: s.accountId });
    assert.equal(s.at('Drafts').length, 0);
    assert.deepEqual(s.at('Trash').map((m) => m.subject), ['Plan for Friday']);
    assert.deepEqual(s.heard.filter((h) => h.event === 'mail:syncProblem'), []);
  } finally {
    await s.close();
  }
});

test('with the server away a draft is kept here, survives a restart, and reaches the server when it is back', async () => {
  const s = await setUp();
  try {
    await s.server.stop();
    const { id } = await s.service.saveDraft({ accountId: s.accountId, draft: s.draft('Written on the train.') });
    const left = await s.service.flushChanges({ accountId: s.accountId });
    assert.ok(left.left >= 1, 'the change waits');
    assert.equal(s.rows('Drafts').length, 1, 'kept on the computer');
    assert.match(s.service.message({ accountId: s.accountId, folder: 'Drafts', id }).text, /Written on the train\./);
    // Saved twice more with no connection: only the last words need go.
    await s.service.saveDraft({ accountId: s.accountId, draft: s.draft('Written on the train, and changed.', { id: s.idOf('Drafts', id) }), replaces: { accountId: s.accountId, folder: 'Drafts', id } });
    await s.service.saveDraft({ accountId: s.accountId, draft: s.draft('Written on the train, and finished.', { id: s.idOf('Drafts', id) }), replaces: { accountId: s.accountId, folder: 'Drafts', id } });
    // Office closed and opened again before the server came back. Closing
    // ends the old copy's wait to send its line (250 ms after each save), as
    // the end of the process does; left waiting, it can send the draft as well
    // as the copy opened again, once the server is back.
    await s.service.flushChanges({ accountId: s.accountId });
    const later = s.make();
    assert.equal(s.rows('Drafts', later).length, 1);
    await s.server.start();
    const before = s.server.heard.length;
    await later.sync({ accountId: s.accountId });
    const there = s.at('Drafts');
    assert.equal(there.length, 1, 'one copy at the server');
    assert.match(there[0].raw, /and finished\./);
    assert.equal(s.server.heard.slice(before).filter((l) => / APPEND /.test(l)).length, 1, 'the versions in between were not sent');
    assert.equal(s.rows('Drafts', later).length, 1);
    assert.equal(s.rows('Drafts', later)[0].uid, there[0].uid);
    assert.equal(later.syncState({ accountId: s.accountId }).waiting, 0);
  } finally {
    await s.close();
  }
});

test('a server that will not take the draft loses nothing: the words stay here and its own copy stays there', async () => {
  const s = await setUp();
  try {
    const id = await s.save(s.draft('First words.'));
    s.server.refuse(/ APPEND /);
    await s.save(s.draft('Second words, which the server refuses.', { id: s.idOf('Drafts', id) }), { accountId: s.accountId, folder: 'Drafts', id });
    // What was typed is here.
    assert.equal(s.rows('Drafts').length, 1);
    assert.match(s.service.message({ accountId: s.accountId, folder: 'Drafts', id }).text, /which the server refuses\./);
    // What the server had, it still has.
    const there = s.at('Drafts');
    assert.equal(there.length, 1);
    assert.match(there[0].raw, /First words\./);
    // And the person is told, once.
    const told = s.heard.filter((h) => h.event === 'mail:syncProblem');
    assert.equal(told.length, 1);
    assert.equal(told[0].payload.kind, 'draft');
    assert.equal(told[0].payload.folder, 'Drafts');
    // A fetch does not take the words away either.
    await s.service.sync({ accountId: s.accountId, folder: 'Drafts' });
    assert.match(s.service.message({ accountId: s.accountId, folder: 'Drafts', id }).text, /which the server refuses\./);
    // Saved once the server takes drafts again, it goes and the old copy with it.
    s.server.allow();
    await s.save(s.draft('Third words.', { id: s.idOf('Drafts', id) }), { accountId: s.accountId, folder: 'Drafts', id });
    assert.deepEqual(s.at('Drafts').map((m) => /Third words\./.test(m.raw)), [true]);
  } finally {
    await s.close();
  }
});

test('a draft written in another mail program opens with its Bcc, is saved over and sent, and leaves no copy behind', async () => {
  const s = await setUp();
  try {
    s.server.deliver('Drafts', [
      'From: Sam <sam@checks.example>', 'To: Amina Yusuf <amina@work.example>', 'Bcc: Hidden Reader <hidden@work.example>', 'Subject: Started on the phone',
      `Date: ${new Date(Date.UTC(2026, 9, 2, 8, 0)).toUTCString()}`, 'Message-ID: <phone-1@other.example>', '', 'Typed on the phone.', '',
    ].join('\r\n'), { flags: ['\\Draft', '\\Seen'] });
    await s.service.sync({ accountId: s.accountId, folder: 'Drafts' });
    const [found] = s.rows('Drafts');
    assert.equal(found.subject, 'Started on the phone');
    const opened = s.service.message({ accountId: s.accountId, folder: 'Drafts', id: found.id });
    // What the form is filled from: the words, and the Bcc line the other program wrote.
    assert.match(opened.text, /Typed on the phone\./);
    assert.deepEqual(opened.bcc.map((a) => a.address), ['hidden@work.example']);
    assert.equal(s.service.folders({ accountId: s.accountId }).find((f) => f.path === 'Drafts').role, 'drafts', 'so the window opens it in the form');

    // Carried on here and saved: the phone's copy goes, ours is there.
    const id = await s.save(s.draft('Typed on the phone, finished at the desk.', { subject: 'Started on the phone', bcc: 'Hidden Reader <hidden@work.example>' }), { accountId: s.accountId, folder: 'Drafts', id: found.id });
    assert.notEqual(id, found.id);
    let there = s.at('Drafts');
    assert.equal(there.length, 1, there.map((m) => m.messageId).join(','));
    assert.match(there[0].raw, /finished at the desk\./);
    assert.match(there[0].messageId, /^<draft-/);
    assert.deepEqual(s.rows('Drafts').map((r) => r.id), [id]);
    await s.service.sync({ accountId: s.accountId, folder: 'Drafts' });
    assert.deepEqual(s.rows('Drafts').map((r) => r.id), [id], 'the phone copy does not come back');

    // Sent: nothing of it is left in Drafts anywhere.
    s.service.delete({ accountId: s.accountId, folder: 'Drafts', ids: [id] });
    await s.service.flushChanges({ accountId: s.accountId });
    there = s.at('Drafts');
    assert.equal(there.length, 0);
    assert.equal(s.rows('Drafts').length, 0);
  } finally {
    await s.close();
  }
});

test('a draft from another program that the server will not replace is shown again as the server has it', async () => {
  const s = await setUp();
  try {
    s.server.deliver('Drafts', ['From: Sam <sam@checks.example>', 'To: amina@work.example', 'Subject: Started on the phone', 'Message-ID: <phone-2@other.example>', '', 'Typed on the phone.', ''].join('\r\n'), { flags: ['\\Draft'] });
    await s.service.sync({ accountId: s.accountId, folder: 'Drafts' });
    const [found] = s.rows('Drafts');
    s.server.refuse(/ APPEND /);
    const id = await s.save(s.draft('Finished at the desk.', { subject: 'Started on the phone' }), { accountId: s.accountId, folder: 'Drafts', id: found.id });
    // The server still has the phone's, and so the list shows it beside what was typed here.
    assert.deepEqual(s.at('Drafts').map((m) => m.messageId), ['<phone-2@other.example>']);
    assert.deepEqual(s.rows('Drafts').map((r) => r.id).sort(), [found.id, id].sort());
    assert.match(s.service.message({ accountId: s.accountId, folder: 'Drafts', id }).text, /Finished at the desk\./);
  } finally {
    await s.close();
  }
});

test('a draft saved again on another computer shows its newer words here', async () => {
  const s = await setUp();
  try {
    const id = await s.save(s.draft('Words from this computer.'));
    await s.service.sync({ accountId: s.accountId, folder: 'Drafts' });
    // A second computer with the same account: its own store, the same server.
    const elsewhere = s.make(fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-maildrafts-b-')), fakeStores());
    const { id: other } = elsewhere.addAccount({ account: s.server.account(), password: 'check' });
    await elsewhere.sync({ accountId: other, folder: 'Drafts' });
    const [theirs] = elsewhere.messages({ accountId: other, folder: 'Drafts', limit: 10 }).rows;
    assert.equal(theirs.id, id, 'the same draft there');
    await elsewhere.saveDraft({ accountId: other, draft: s.draft('Words from the other computer.', { id: String(elsewhere.message({ accountId: other, folder: 'Drafts', id }).messageId).slice(6) }), replaces: { accountId: other, folder: 'Drafts', id } });
    await elsewhere.flushChanges({ accountId: other });
    assert.equal(s.at('Drafts').length, 1);
    // Fetched here: one draft, with the words the other computer saved.
    await s.service.sync({ accountId: s.accountId, folder: 'Drafts' });
    assert.equal(s.rows('Drafts').length, 1);
    assert.match(s.service.message({ accountId: s.accountId, folder: 'Drafts', id }).text, /Words from the other computer\./);
    assert.equal(s.rows('Drafts')[0].uid, s.at('Drafts')[0].uid);
  } finally {
    await s.close();
  }
});

test('a server with no Drafts folder is given one, and the draft kept in it', async () => {
  const s = await setUp({ folders: [['INBOX', null], ['Sent', '\\Sent'], ['Trash', '\\Trash']] });
  try {
    assert.ok(!s.server.folders().includes('Drafts'));
    const id = await s.save(s.draft('Nowhere to go yet.'));
    assert.ok(s.server.folders().includes('Drafts'), s.server.folders().join(','));
    const there = s.at('Drafts');
    assert.equal(there.length, 1);
    assert.ok(there[0].flags.includes('\\Draft'));
    await s.service.sync({ accountId: s.accountId, folder: 'Drafts' });
    assert.deepEqual(s.rows('Drafts').map((r) => r.id), [id]);
    const listed = s.service.folders({ accountId: s.accountId }).filter((f) => f.role === 'drafts');
    assert.deepEqual(listed.map((f) => f.path), ['Drafts'], 'one Drafts folder in the list');
    assert.ok(!listed[0].local, 'and it is the server\'s');
  } finally {
    await s.close();
  }
});

test('a server that will not make a Drafts folder leaves the draft here, and says so', async () => {
  const s = await setUp({ folders: [['INBOX', null], ['Sent', '\\Sent'], ['Trash', '\\Trash']] });
  try {
    s.server.refuse(/ CREATE /);
    const id = await s.save(s.draft('Kept on this computer.'));
    assert.ok(!s.server.folders().includes('Drafts'));
    assert.deepEqual(s.rows('Drafts').map((r) => r.id), [id]);
    assert.equal(s.heard.filter((h) => h.event === 'mail:syncProblem' && h.payload.kind === 'draft').length, 1);
    await s.service.sync({ accountId: s.accountId });
    assert.deepEqual(s.rows('Drafts').map((r) => r.id), [id], 'a fetch does not take it away');
  } finally {
    await s.close();
  }
});

test('on Gmail a draft goes to [Gmail]/Drafts, before the first fetch and after it, and no Drafts label is made', async () => {
  const s = await setUp({ capabilities: { gmail: true }, folders: GMAIL, fetch: false });
  try {
    // Saved before the account was ever fetched: Mail does not know the server's folders yet.
    const early = await s.save(s.draft('Before the first fetch.', { subject: 'Early' }));
    assert.deepEqual(s.at('[Gmail]/Drafts').map((m) => m.subject), ['Early'], 'found the server\'s own Drafts');
    assert.ok(!s.server.folders().includes('Drafts'), 'and made no label called Drafts');
    assert.deepEqual(s.rows('[Gmail]/Drafts').map((r) => r.id), [early]);
    assert.equal(s.rows('Drafts').length, 0);
    await s.service.sync({ accountId: s.accountId });

    const id = await s.save(s.draft('After it.', { subject: 'Later' }));
    assert.deepEqual(s.at('[Gmail]/Drafts').map((m) => m.subject).sort(), ['Early', 'Later']);
    assert.deepEqual(s.rows('[Gmail]/Drafts').map((r) => r.id).sort(), [early, id].sort());
    assert.ok(!s.server.folders().includes('Drafts'));
    assert.ok(!s.service.folders({ accountId: s.accountId }).some((f) => f.path === 'Drafts'), 'no second Drafts in the list');
  } finally {
    await s.close();
  }
});

test('a draft kept only on this computer by an earlier version reaches the server when it is saved again', async () => {
  const s = await setUp({ capabilities: { gmail: true }, folders: GMAIL });
  try {
    // As 2.27 left it: a message in a folder of Mail's own called Drafts, under an id with no domain.
    const { MailStore } = await import('../packages/mailbox/src/store.js');
    const { parseMessage } = await import('../packages/mailbox/src/mime.js');
    const store = new MailStore(path.join(s.userData, 'mail'));
    const old = parseMessage(['From: sam@checks.example', 'To: amina@work.example', 'Subject: From last month', 'Message-ID: <draft-11111111-2222-4333-8444-555555555555>', '', 'Old words.', ''].join('\r\n'));
    old.messageId = 'draft-11111111-2222-4333-8444-555555555555';
    const oldId = store.put(s.accountId, 'Drafts', old, { force: true });
    store.upsertFolder(s.accountId, { path: 'Drafts', name: 'Drafts', role: 'drafts' });
    assert.deepEqual(s.rows('Drafts').map((r) => r.id), [oldId]);

    const id = await s.save(s.draft('Old words, carried on.', { subject: 'From last month', id: '11111111-2222-4333-8444-555555555555' }), { accountId: s.accountId, folder: 'Drafts', id: oldId });
    assert.equal(id, oldId, 'the same draft, under the id it had');
    assert.deepEqual(s.at('[Gmail]/Drafts').map((m) => m.messageId), ['<draft-11111111-2222-4333-8444-555555555555>']);
    assert.deepEqual(s.rows('[Gmail]/Drafts').map((r) => r.id), [id]);
    assert.equal(s.rows('Drafts').length, 0, 'it left Mail\'s own folder');
    assert.ok(!s.server.folders().includes('Drafts'));
    assert.deepEqual(s.heard.filter((h) => h.event === 'mail:syncProblem'), []);
  } finally {
    await s.close();
  }
});

test('an imported mailbox, which has no server, keeps its drafts as before', async () => {
  const s = await setUp();
  try {
    const { id: local } = s.service.addAccount({ account: { email: 'old@archive.example', name: 'Old mail', local: true }, password: '' });
    const before = s.server.heard.length;
    const { id } = await s.service.saveDraft({ accountId: local, draft: s.draft('Local only.') });
    await s.service.flushChanges({ accountId: local });
    assert.deepEqual(s.service.messages({ accountId: local, folder: 'Drafts', limit: 10 }).rows.map((r) => r.id), [id]);
    assert.equal(s.server.heard.length, before, 'nothing said to any server');
    const again = await s.service.saveDraft({ accountId: local, draft: s.draft('Local only, again.', { id: String(s.service.message({ accountId: local, folder: 'Drafts', id }).messageId).slice(6) }), replaces: { accountId: local, folder: 'Drafts', id } });
    assert.equal(again.id, id);
    assert.equal(s.service.messages({ accountId: local, folder: 'Drafts', limit: 10 }).rows.length, 1);
  } finally {
    await s.close();
  }
});

/* ── Archive (B20) ─────────────────────────────────────────────────── */

test('Archive on Gmail moves the message to All Mail and makes no Archive label', async () => {
  const s = await setUp({ capabilities: { gmail: true }, folders: GMAIL });
  try {
    const one = s.rows('INBOX').find((r) => r.subject === 'Message 1');
    // What the Archive button asks for: no folder named, the account's own found by Mail.
    const r = s.service.move({ accountId: s.accountId, folder: 'INBOX', ids: [one.id], archive: true });
    assert.equal(r.moved, 1);
    await s.service.flushChanges({ accountId: s.accountId });
    assert.deepEqual(s.at('INBOX').map((m) => m.subject), ['Message 2'], 'out of the Inbox at the server');
    assert.deepEqual(s.at('[Gmail]/All Mail').map((m) => m.subject), ['Message 1']);
    assert.ok(!s.server.folders().some((f) => /archive/i.test(f)), s.server.folders().join(','));
    assert.ok(!s.server.heard.some((l) => / CREATE /.test(l)), 'no folder was made');
    // Here too: out of the Inbox, in All Mail, and no Archive folder in the list.
    assert.deepEqual(s.rows('INBOX').map((r) => r.subject), ['Message 2']);
    assert.deepEqual(s.rows('[Gmail]/All Mail').map((r) => r.subject), ['Message 1']);
    const folders = s.service.folders({ accountId: s.accountId });
    assert.ok(!folders.some((f) => f.path === 'Archive'));
    // All Mail is listed under the server's name for it, with the archive icon, after the Trash and before ordinary folders.
    const all = folders.find((f) => f.path === '[Gmail]/All Mail');
    assert.deepEqual({ name: all.name, role: all.role, icon: all.icon }, { name: 'All Mail', role: 'allmail', icon: 'archive' });
    assert.deepEqual(folders.map((f) => f.role), ['inbox', 'drafts', 'sent', 'junk', 'trash', 'allmail']);
    await s.service.sync({ accountId: s.accountId });
    assert.deepEqual(s.rows('INBOX').map((r) => r.subject), ['Message 2'], 'and it does not come back');
    assert.deepEqual(s.heard.filter((h) => h.event === 'mail:syncProblem'), []);
  } finally {
    await s.close();
  }
});

test('Archive on Gmail goes to All Mail even where an earlier version left an Archive label', async () => {
  const s = await setUp({ capabilities: { gmail: true }, folders: [...GMAIL, ['Archive', null]] });
  try {
    const one = s.rows('INBOX').find((r) => r.subject === 'Message 1');
    s.service.move({ accountId: s.accountId, folder: 'INBOX', ids: [one.id], archive: true });
    await s.service.flushChanges({ accountId: s.accountId });
    assert.deepEqual(s.at('[Gmail]/All Mail').map((m) => m.subject), ['Message 1']);
    assert.equal(s.at('Archive').length, 0, 'the label is left as it was');
  } finally {
    await s.close();
  }
});

test('Archive on a server with its own Archive folder goes there, as before', async () => {
  const s = await setUp();
  try {
    const one = s.rows('INBOX').find((r) => r.subject === 'Message 1');
    s.service.move({ accountId: s.accountId, folder: 'INBOX', ids: [one.id], archive: true });
    await s.service.flushChanges({ accountId: s.accountId });
    assert.deepEqual(s.at('Archive').map((m) => m.subject), ['Message 1']);
    assert.deepEqual(s.at('INBOX').map((m) => m.subject), ['Message 2']);
  } finally {
    await s.close();
  }
});

test('Archive prefers the folder the server calls its archive, then one named Archive, over All Mail', async () => {
  // \Archive under another name, beside \All.
  const flagged = await setUp({ folders: [['INBOX', null], ['Everything', '\\All'], ['Kept', '\\Archive'], ['Trash', '\\Trash']] });
  try {
    const one = flagged.rows('INBOX')[0];
    flagged.service.move({ accountId: flagged.accountId, folder: 'INBOX', ids: [one.id], archive: true });
    await flagged.service.flushChanges({ accountId: flagged.accountId });
    assert.equal(flagged.at('Kept').length, 1);
    assert.equal(flagged.at('Everything').length, 0);
  } finally {
    await flagged.close();
  }
  // Not Gmail, an Archive folder by name only, beside \All: the Archive folder, as before.
  const named = await setUp({ folders: [['INBOX', null], ['Everything', '\\All'], ['Archive', null], ['Trash', '\\Trash']] });
  try {
    const one = named.rows('INBOX')[0];
    named.service.move({ accountId: named.accountId, folder: 'INBOX', ids: [one.id], archive: true });
    await named.service.flushChanges({ accountId: named.accountId });
    assert.equal(named.at('Archive').length, 1);
    assert.equal(named.at('Everything').length, 0);
  } finally {
    await named.close();
  }
});

test('Archive on a server with \\All and no Archive folder goes to \\All and makes nothing', async () => {
  const s = await setUp({ folders: [['INBOX', null], ['Everything', '\\All'], ['Trash', '\\Trash']] });
  try {
    const one = s.rows('INBOX')[0];
    s.service.move({ accountId: s.accountId, folder: 'INBOX', ids: [one.id], archive: true });
    await s.service.flushChanges({ accountId: s.accountId });
    assert.equal(s.at('Everything').length, 1);
    assert.ok(!s.server.folders().includes('Archive'));
  } finally {
    await s.close();
  }
});

test('Archive on a server with neither makes an Archive folder, as before', async () => {
  const s = await setUp({ folders: [['INBOX', null], ['Sent', '\\Sent'], ['Trash', '\\Trash']] });
  try {
    const one = s.rows('INBOX')[0];
    s.service.move({ accountId: s.accountId, folder: 'INBOX', ids: [one.id], archive: true });
    await s.service.flushChanges({ accountId: s.accountId });
    assert.ok(s.server.folders().includes('Archive'));
    assert.equal(s.at('Archive').length, 1);
  } finally {
    await s.close();
  }
});

test('a move that names its folder still goes where it is told', async () => {
  const s = await setUp({ capabilities: { gmail: true }, folders: [...GMAIL, ['Receipts', null]] });
  try {
    const one = s.rows('INBOX')[0];
    s.service.move({ accountId: s.accountId, folder: 'INBOX', ids: [one.id], to: 'Receipts' });
    await s.service.flushChanges({ accountId: s.accountId });
    assert.equal(s.at('Receipts').length, 1);
    assert.equal(s.at('[Gmail]/All Mail').length, 0);
  } finally {
    await s.close();
  }
});
