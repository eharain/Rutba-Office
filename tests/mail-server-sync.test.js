// Mail's changes reach the server, and the server's reach Mail (A6.1).
//
// Against the test IMAP server (packages/mailbox/test/imap-server.js), run
// in this process: a message deleted here is gone at the server and does not
// come back; one filed here is filed there; read and flagged marks go both
// ways; changes made with the server away wait on disk, survive a restart,
// and are made once, in order, when it is back; a folder renumbered (a new
// UIDVALIDITY) is matched again without doubling a message; the Sent copy is
// kept at the server unless the server keeps its own; a change the server
// refuses is said and put back; and a store from before all this is matched
// with its server without bringing back what it had deleted.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createMailService } from '../apps/desktop/main/mail.js';
import { MailStore } from '../packages/mailbox/src/store.js';
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
    held: { settings, secrets },
  };
}

const message = (n, extra = '') => [
  'From: Amina Yusuf <amina@work.example>', 'To: Sam <sam@checks.example>', `Subject: Message ${n}`,
  `Date: ${new Date(Date.UTC(2026, 9, 1, 9, n)).toUTCString()}`, `Message-ID: <m${n}@work.example>`, ...(extra ? [extra] : []), '', `Body of message ${n}.`, '',
].join('\r\n');

/** A server with three messages in its Inbox, a service with its account, and a way to make another service on the same folder (a restart). */
async function setUp({ capabilities = {}, count = 3 } = {}) {
  const server = await imapServer({ capabilities });
  for (let n = 1; n <= count; n++) server.deliver('INBOX', message(n));
  const userData = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-mailsync-'));
  const stores = fakeStores();
  const heard = [];
  const sent = [];
  const make = () => createMailService({
    stores,
    holdBlob: () => ({ url: '' }),
    broadcast: (event, payload) => heard.push({ event, payload }),
    userData,
    smtp: { async send({ draft }) { sent.push(draft); return { messageId: '<ignored@smtp>' }; } },
  });
  const service = make();
  const { id: accountId } = service.addAccount({ account: server.account(), password: 'check' });
  const subjects = (folder, svc = service) => svc.messages({ accountId, folder, limit: 100 }).rows.map((r) => r.subject).sort();
  const row = (folder, subject, svc = service) => svc.messages({ accountId, folder, limit: 100 }).rows.find((r) => r.subject === subject);
  const at = (folder) => server.messages(folder).map((m) => m.subject).sort();
  return { server, service, make, stores, userData, accountId, heard, sent, subjects, row, at, close: () => server.close() };
}

test('a fetch goes by UID, and fetching again brings nothing twice', async () => {
  const s = await setUp();
  try {
    const first = await s.service.sync({ accountId: s.accountId });
    assert.equal(first.added, 3);
    assert.deepEqual(s.subjects('INBOX'), ['Message 1', 'Message 2', 'Message 3']);
    assert.ok(s.service.messages({ accountId: s.accountId, folder: 'INBOX' }).rows.every((r) => Number.isInteger(r.uid)), 'each message knows its UID');
    s.server.deliver('INBOX', message(4));
    const before = s.server.heard.length;
    const second = await s.service.sync({ accountId: s.accountId });
    assert.equal(second.added, 1);
    assert.deepEqual(s.subjects('INBOX'), ['Message 1', 'Message 2', 'Message 3', 'Message 4']);
    // Only the new one was fetched whole.
    const wholes = s.server.heard.slice(before).filter((l) => /BODY\.PEEK\[\]/.test(l));
    assert.equal(wholes.length, 1, wholes.join(' | '));
    assert.match(wholes[0], /UID FETCH 4 /);
  } finally {
    await s.close();
  }
});

test('a message deleted here goes to the Trash at the server, deleted for good there too, and never comes back', async () => {
  const s = await setUp();
  try {
    await s.service.sync({ accountId: s.accountId });
    const two = s.row('INBOX', 'Message 2');
    // The window sends it to the account's Trash.
    s.service.move({ accountId: s.accountId, folder: 'INBOX', ids: [two.id], to: 'Trash' });
    await s.service.flushChanges({ accountId: s.accountId });
    assert.deepEqual(s.at('INBOX'), ['Message 1', 'Message 3'], 'gone from the Inbox at the server');
    assert.deepEqual(s.at('Trash'), ['Message 2'], 'and in its Trash');
    await s.service.sync({ accountId: s.accountId, folder: 'Trash' });
    assert.deepEqual(s.subjects('INBOX'), ['Message 1', 'Message 3'], 'it does not come back');
    assert.deepEqual(s.subjects('Trash'), ['Message 2'], 'one copy in the Trash here');
    const inTrash = s.row('Trash', 'Message 2');
    assert.equal(inTrash.uid, s.server.messages('Trash')[0].uid, 'with the UID the server gave it there');
    // Deleted from the Trash: for good, at the server too.
    s.service.delete({ accountId: s.accountId, folder: 'Trash', ids: [inTrash.id] });
    await s.service.flushChanges({ accountId: s.accountId });
    assert.deepEqual(s.at('Trash'), []);
    assert.ok(s.server.heard.some((l) => /UID EXPUNGE/.test(l)), 'by EXPUNGE');
    await s.service.sync({ accountId: s.accountId, folder: 'Trash' });
    assert.deepEqual(s.subjects('Trash'), []);
    assert.deepEqual(s.subjects('INBOX'), ['Message 1', 'Message 3']);
    // Empty folder empties it at the server.
    s.service.move({ accountId: s.accountId, folder: 'INBOX', ids: [s.row('INBOX', 'Message 1').id], to: 'Junk' });
    s.service.emptyFolder({ accountId: s.accountId, folder: 'Junk' });
    await s.service.flushChanges({ accountId: s.accountId });
    assert.deepEqual(s.at('Junk'), []);
    assert.deepEqual(s.at('INBOX'), ['Message 3']);
  } finally {
    await s.close();
  }
});

test('read, unread and flagged go to the server, and marks made at the server elsewhere show here', async () => {
  for (const condstore of [true, false]) {
    const s = await setUp({ capabilities: { condstore } });
    try {
      await s.service.sync({ accountId: s.accountId });
      const one = s.row('INBOX', 'Message 1');
      s.service.flag({ accountId: s.accountId, folder: 'INBOX', ids: [one.id], patch: { unread: false, flagged: true } });
      await s.service.flushChanges({ accountId: s.accountId });
      assert.deepEqual(s.server.messages('INBOX')[0].flags.sort(), ['\\Flagged', '\\Seen'], `condstore ${condstore}: marked at the server`);
      // A pin is Rutba Office's own: nothing goes to the server.
      const before = s.server.heard.length;
      s.service.flag({ accountId: s.accountId, folder: 'INBOX', ids: [one.id], patch: { pinned: true } });
      await s.service.flushChanges({ accountId: s.accountId });
      assert.equal(s.server.heard.length, before, 'a pin sends nothing');
      // A phone reads message 3 and flags message 2, and unflags message 1.
      const [m1, m2, m3] = s.server.messages('INBOX');
      s.server.setFlags('INBOX', m3.uid, ['\\Seen']);
      s.server.setFlags('INBOX', m2.uid, ['\\Flagged']);
      s.server.setFlags('INBOX', m1.uid, ['\\Seen']);
      await s.service.sync({ accountId: s.accountId });
      assert.equal(s.row('INBOX', 'Message 3').unread, false, `condstore ${condstore}: read elsewhere`);
      assert.equal(s.row('INBOX', 'Message 2').flagged, true, 'flagged elsewhere');
      assert.equal(s.row('INBOX', 'Message 2').unread, true);
      assert.equal(s.row('INBOX', 'Message 1').flagged, false, 'unflagged elsewhere');
      assert.equal(s.row('INBOX', 'Message 1').pinned, true, 'its pin kept');
      if (condstore) assert.ok(s.server.heard.some((l) => /CHANGEDSINCE/.test(l)), 'asked with CHANGEDSINCE');
      else assert.ok(!s.server.heard.some((l) => /CHANGEDSINCE/.test(l)));
    } finally {
      await s.close();
    }
  }
});

test('deleted or filed at the server elsewhere, a message goes here too', async () => {
  const s = await setUp();
  try {
    await s.service.sync({ accountId: s.accountId });
    await s.service.sync({ accountId: s.accountId, folder: 'Archive' });
    const [m1, m2] = s.server.messages('INBOX');
    s.server.expunge('INBOX', m1.uid);
    s.server.move('INBOX', m2.uid, 'Archive');
    await s.service.sync({ accountId: s.accountId });
    assert.deepEqual(s.subjects('INBOX'), ['Message 3']);
    // Archive was fetched before, so it is fetched again with the Inbox.
    assert.deepEqual(s.subjects('Archive'), ['Message 2']);
    assert.ok(s.heard.some((h) => h.event === 'mail:new'), 'the window is told');
  } finally {
    await s.close();
  }
});

test('changes made with the server away wait on disk, outlive a restart, and are made once, in order, when it is back', async () => {
  const s = await setUp();
  try {
    await s.service.sync({ accountId: s.accountId });
    await s.server.stop();
    const one = s.row('INBOX', 'Message 1');
    const two = s.row('INBOX', 'Message 2');
    s.service.flag({ accountId: s.accountId, folder: 'INBOX', ids: [one.id], patch: { unread: false } });
    s.service.move({ accountId: s.accountId, folder: 'INBOX', ids: [one.id], to: 'Archive' });
    s.service.move({ accountId: s.accountId, folder: 'INBOX', ids: [two.id], to: 'Trash' });
    const tried = await s.service.flushChanges({ accountId: s.accountId });
    assert.equal(tried.left, 3, 'nothing lost while the server is away');
    await assert.rejects(s.service.sync({ accountId: s.accountId }), 'Get mail says the server could not be reached');
    // Rutba Office closed and opened again: a new service on the same store.
    const again = s.make();
    assert.equal(again.syncState({ accountId: s.accountId }).waiting, 3, 'still waiting after a restart');
    await s.server.start();
    const before = s.server.heard.length;
    await again.flushChanges({ accountId: s.accountId });
    const made = s.server.heard.slice(before).filter((l) => /UID (STORE|MOVE)/.test(l));
    assert.equal(made.length, 3, made.join(' | '));
    assert.match(made[0], /UID STORE \d+ \+FLAGS \(\\Seen\)/);
    assert.match(made[1], /UID MOVE \d+ Archive/);
    assert.match(made[2], /UID MOVE \d+ Trash/);
    assert.deepEqual(s.at('INBOX'), ['Message 3']);
    assert.deepEqual(s.at('Archive'), ['Message 1']);
    assert.ok(s.server.messages('Archive')[0].flags.includes('\\Seen'), 'read, then filed');
    assert.deepEqual(s.at('Trash'), ['Message 2']);
    assert.equal(again.syncState({ accountId: s.accountId }).waiting, 0);
    // Made again (a crash before the line was written): nothing happens twice.
    const store = new MailStore(path.join(s.userData, 'mail'));
    store.addChange(s.accountId, { kind: 'move', folder: 'INBOX', to: 'Archive', items: [{ id: one.id, uid: one.uid, messageId: 'm1@work.example' }] });
    await again.flushChanges({ accountId: s.accountId });
    assert.deepEqual(s.at('Archive'), ['Message 1'], 'filed once');
    await again.sync({ accountId: s.accountId, folder: 'Archive' });
    assert.deepEqual(s.subjects('INBOX', again), ['Message 3']);
    assert.deepEqual(s.subjects('Archive', again), ['Message 1']);
  } finally {
    await s.close();
  }
});

test('a folder renumbered at the server (a new UIDVALIDITY) is matched again without doubling a message', async () => {
  const s = await setUp();
  try {
    await s.service.sync({ accountId: s.accountId });
    const was = s.row('INBOX', 'Message 1').uid;
    s.server.renumber('INBOX');
    s.server.deliver('INBOX', message(4));
    await s.service.sync({ accountId: s.accountId });
    assert.deepEqual(s.subjects('INBOX'), ['Message 1', 'Message 2', 'Message 3', 'Message 4']);
    const now = s.row('INBOX', 'Message 1').uid;
    assert.notEqual(now, was);
    assert.equal(now, s.server.messages('INBOX')[0].uid, 'each message has its new UID');
    // And changes go by the new UIDs.
    s.service.flag({ accountId: s.accountId, folder: 'INBOX', ids: [s.row('INBOX', 'Message 1').id], patch: { flagged: true } });
    await s.service.flushChanges({ accountId: s.accountId });
    assert.ok(s.server.messages('INBOX')[0].flags.includes('\\Flagged'));
  } finally {
    await s.close();
  }
});

test('the Sent copy is kept at the server once, unless the server keeps its own', async () => {
  const s = await setUp();
  try {
    await s.service.sync({ accountId: s.accountId });
    await s.service.send({ accountId: s.accountId, draft: { to: 'amina@work.example', subject: 'The minutes', text: 'Attached are the minutes.' } });
    await s.service.flushChanges({ accountId: s.accountId });
    assert.deepEqual(s.at('Sent'), ['The minutes']);
    assert.equal(s.sent[0].raw && s.server.messages('Sent')[0].raw.includes('Attached are the minutes.'), true, 'the copy is the message as it went');
    assert.deepEqual(s.subjects('Sent'), ['The minutes']);
    assert.equal(s.row('Sent', 'The minutes').uid, s.server.messages('Sent')[0].uid, 'with its UID there');
    // The same change again (a crash after the copy was made): not twice.
    const store = new MailStore(path.join(s.userData, 'mail'));
    const file = store.outgoingPath(s.accountId, 'again.eml');
    fs.writeFileSync(file, s.server.messages('Sent')[0].raw);
    store.addChange(s.accountId, { kind: 'append', folder: 'Sent', file, messageId: s.server.messages('Sent')[0].messageId.replace(/^<|>$/g, '') });
    await s.service.flushChanges({ accountId: s.accountId });
    assert.equal(s.server.messages('Sent').length, 1);
    await s.service.sync({ accountId: s.accountId, folder: 'Sent' });
    assert.deepEqual(s.subjects('Sent'), ['The minutes'], 'and once here');
  } finally {
    await s.close();
  }
  const g = await setUp({ capabilities: { gmail: true } });
  try {
    await g.service.sync({ accountId: g.accountId });
    await g.service.send({ accountId: g.accountId, draft: { to: 'amina@work.example', subject: 'From Gmail', text: 'Hello.' } });
    await g.service.flushChanges({ accountId: g.accountId });
    assert.ok(!g.server.heard.some((l) => /APPEND/.test(l)), 'a Gmail server keeps its own: nothing is appended');
    assert.deepEqual(g.subjects('[Gmail]/Sent Mail'), ['From Gmail'], 'and the copy here is in its Sent');
  } finally {
    await g.close();
  }
});

test('a change the server refuses is said in plain words and the message put back as the server has it', async () => {
  const s = await setUp();
  try {
    await s.service.sync({ accountId: s.accountId });
    s.server.refuse(/UID MOVE/);
    const two = s.row('INBOX', 'Message 2');
    s.service.move({ accountId: s.accountId, folder: 'INBOX', ids: [two.id], to: 'Archive' });
    assert.deepEqual(s.subjects('Archive'), ['Message 2'], 'filed here at once');
    await s.service.flushChanges({ accountId: s.accountId });
    const problem = s.heard.find((h) => h.event === 'mail:syncProblem');
    assert.ok(problem, 'the window is told');
    assert.equal(problem.payload.kind, 'move');
    assert.equal(problem.payload.to, 'Archive');
    assert.deepEqual(s.subjects('Archive'), [], 'not left where the server would not put it');
    assert.deepEqual(s.subjects('INBOX'), ['Message 1', 'Message 2', 'Message 3'], 'back in the Inbox, where the server keeps it');
    assert.equal(s.service.syncState({ accountId: s.accountId }).waiting, 0, 'and not tried for ever');
    s.server.allow();
    s.server.refuse(/UID STORE/);
    s.service.flag({ accountId: s.accountId, folder: 'INBOX', ids: [two.id], patch: { flagged: true } });
    await s.service.flushChanges({ accountId: s.accountId });
    assert.equal(s.row('INBOX', 'Message 2').flagged, false, 'a mark refused is taken off again');
  } finally {
    await s.close();
  }
});

test('a server without MOVE or UIDPLUS: filed by copy and delete, and the new UID found by Message-ID', async () => {
  const s = await setUp({ capabilities: { move: false, uidplus: false } });
  try {
    await s.service.sync({ accountId: s.accountId });
    await s.service.sync({ accountId: s.accountId, folder: 'Archive' });
    s.service.move({ accountId: s.accountId, folder: 'INBOX', ids: [s.row('INBOX', 'Message 3').id], to: 'Archive' });
    await s.service.flushChanges({ accountId: s.accountId });
    assert.deepEqual(s.at('INBOX'), ['Message 1', 'Message 2']);
    assert.deepEqual(s.at('Archive'), ['Message 3']);
    assert.ok(s.server.heard.some((l) => /UID COPY/.test(l)));
    await s.service.sync({ accountId: s.accountId });
    assert.deepEqual(s.subjects('Archive'), ['Message 3'], 'one copy');
    assert.equal(s.row('Archive', 'Message 3').uid, s.server.messages('Archive')[0].uid);
  } finally {
    await s.close();
  }
});

test('a store from before 2.19 is matched with its server: what it deleted stays deleted, what it filed is filed there', async () => {
  const s = await setUp({ count: 5 });
  try {
    // The store as 2.18 left it: the Inbox fetched by position with UIDs, then
    // message 2 deleted and message 3 filed in Archive here only, no sync.json.
    const store = new MailStore(path.join(s.userData, 'mail'));
    const { parseMessage } = await import('../packages/mailbox/src/mime.js');
    const old = s.server.messages('INBOX').map((m) => Object.assign(parseMessage(Buffer.from(m.raw)), { uid: m.uid, unread: true }));
    store.setFolders(s.accountId, [{ path: 'INBOX', name: 'INBOX' }, { path: 'Archive', name: 'Archive' }, { path: 'Sent', name: 'Sent' }]);
    store.putMany(s.accountId, 'INBOX', old.filter((m) => !/Message [23]/.test(m.subject)));
    store.putMany(s.accountId, 'Archive', old.filter((m) => m.subject === 'Message 3'));
    // A Sent copy 2.18 kept here only.
    store.putMany(s.accountId, 'Sent', [Object.assign(parseMessage(Buffer.from(message(9))), { unread: false })]);
    s.server.deliver('INBOX', message(6));
    await s.service.sync({ accountId: s.accountId });
    assert.deepEqual(s.subjects('INBOX'), ['Message 1', 'Message 4', 'Message 5', 'Message 6'], 'message 2 does not come back; message 6 arrives');
    assert.deepEqual(s.at('Archive'), ['Message 3'], 'what was filed here is filed at the server');
    assert.deepEqual(s.at('INBOX').sort(), ['Message 1', 'Message 2', 'Message 4', 'Message 5', 'Message 6'], 'what was deleted here is left at the server, not deleted by a guess');
    assert.deepEqual(s.subjects('Sent'), ['Message 9'], 'nothing kept here is lost');
    assert.ok(store.syncState(s.accountId, 'INBOX'), 'the Inbox now has its place at the server');
  } finally {
    await s.close();
  }
});

test('folders made, renamed and deleted here are so at the server, and a message filed in one goes there (A6.4)', async () => {
  const s = await setUp();
  try {
    await s.service.sync({ accountId: s.accountId });
    const folders = () => s.service.folders({ accountId: s.accountId }).map((f) => f.path);
    assert.deepEqual(s.service.createFolder({ accountId: s.accountId, name: 'Projects' }), { path: 'Projects' });
    assert.deepEqual(s.service.createFolder({ accountId: s.accountId, parent: 'Projects', name: 'Old' }), { path: 'Projects/Old' });
    assert.throws(() => s.service.createFolder({ accountId: s.accountId, name: 'Projects' }), /already/);
    assert.throws(() => s.service.createFolder({ accountId: s.accountId, name: 'a/b' }), /character/);
    assert.ok(folders().includes('Projects/Old'), 'in the list at once');
    // Dragged onto it: a move, as Move to is.
    s.service.move({ accountId: s.accountId, folder: 'INBOX', ids: [s.row('INBOX', 'Message 1').id], to: 'Projects/Old' });
    await s.service.flushChanges({ accountId: s.accountId });
    assert.ok(s.server.folders().includes('Projects') && s.server.folders().includes('Projects/Old'));
    assert.deepEqual(s.at('Projects/Old'), ['Message 1']);

    // Renamed, with the folder inside it and its mail.
    assert.deepEqual(s.service.renameFolder({ accountId: s.accountId, folder: 'Projects', name: 'Work' }), { path: 'Work' });
    assert.deepEqual(s.subjects('Work/Old'), ['Message 1']);
    assert.ok(!folders().includes('Projects') && folders().includes('Work/Old'));
    await s.service.flushChanges({ accountId: s.accountId });
    assert.ok(!s.server.folders().includes('Projects') && s.server.folders().includes('Work/Old'));
    assert.deepEqual(s.at('Work/Old'), ['Message 1']);
    await s.service.sync({ accountId: s.accountId, folder: 'Work/Old' });
    assert.deepEqual(s.subjects('Work/Old'), ['Message 1'], 'fetched again, not doubled');

    // The account's own folders keep their names.
    assert.throws(() => s.service.renameFolder({ accountId: s.accountId, folder: 'INBOX', name: 'Post' }), /keeps its name/);
    assert.throws(() => s.service.deleteFolder({ accountId: s.accountId, folder: 'Trash' }), /cannot be deleted/);

    // Deleted, with what is inside it.
    assert.deepEqual(s.service.deleteFolder({ accountId: s.accountId, folder: 'Work' }), { folders: 2, removed: 1 });
    assert.ok(!folders().some((p) => p.startsWith('Work')));
    await s.service.flushChanges({ accountId: s.accountId });
    assert.ok(!s.server.folders().some((p) => p.startsWith('Work')));
    await s.service.sync({ accountId: s.accountId });
    assert.ok(!folders().some((p) => p.startsWith('Work')), 'and it does not come back');

    // Made with the server away: made once it is back.
    await s.server.stop();
    s.service.createFolder({ accountId: s.accountId, name: 'Later' });
    assert.equal((await s.service.flushChanges({ accountId: s.accountId })).left, 1);
    await s.server.start();
    assert.equal((await s.service.flushChanges({ accountId: s.accountId })).left, 0);
    assert.ok(s.server.folders().includes('Later'));
  } finally {
    await s.close();
  }
});

test('folders made, renamed and deleted at the server elsewhere show here (A6.4)', async () => {
  const s = await setUp();
  try {
    await s.service.sync({ accountId: s.accountId });
    const folders = () => s.service.folders({ accountId: s.accountId }).map((f) => f.path);
    s.server.makeFolder('Receipts');
    s.server.deliver('Receipts', message(7));
    s.heard.length = 0;
    await s.service.sync({ accountId: s.accountId });
    assert.ok(folders().includes('Receipts'));
    assert.ok(s.heard.some((h) => h.event === 'mail:new'), 'the window is told to show the list again');
    await s.service.sync({ accountId: s.accountId, folder: 'Receipts' });
    assert.deepEqual(s.subjects('Receipts'), ['Message 7']);

    s.server.renameFolder('Receipts', 'Bills');
    await s.service.sync({ accountId: s.accountId });
    assert.ok(!folders().includes('Receipts') && folders().includes('Bills'));
    await s.service.sync({ accountId: s.accountId, folder: 'Bills' });
    assert.deepEqual(s.subjects('Bills'), ['Message 7']);

    s.server.deleteFolder('Bills');
    await s.service.sync({ accountId: s.accountId });
    assert.ok(!folders().includes('Bills'));
  } finally {
    await s.close();
  }
});
