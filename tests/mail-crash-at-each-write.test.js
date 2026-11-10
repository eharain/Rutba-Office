// The mail store stopped at every write it makes, one at a time, and at each
// of them half way through: no message is lost, and doing the work again
// finishes it.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { MailStore } from '../packages/mailbox/src/store.js';
import { createMailService } from '../apps/desktop/main/mail.js';

class Stopped extends Error {}
const msg = (n, extra = {}) => ({ id: `m${n}`, messageId: `m${n}@x`, subject: `Message ${n}`, from: { address: 'a@b.c' }, to: [], date: '2026-01-01T00:00:00Z', text: `body ${n}`, attachments: [], ...extra });

/** Runs `work`, stopping at write number `at`; `torn` leaves that write half done. Says whether it stopped. */
function stopAt(t, at, torn, work) {
  let count = 0;
  // A stop is the end of the process: nothing after it reaches the disk, also when the code goes on past a write it caught.
  let tripped = false;
  const real = { writeFileSync: fs.writeFileSync, appendFileSync: fs.appendFileSync, renameSync: fs.renameSync, rmSync: fs.rmSync };
  const guard = (name, half) => t.mock.method(fs, name, (...args) => {
    // appendFileSync is made of this one: the one write, counted once.
    if (name === 'writeFileSync' && args[2]?.flag === 'a') return real[name](...args);
    if (tripped) throw new Stopped(name);
    if (++count === at) {
      tripped = true;
      if (torn && half) half(args);
      throw new Stopped(name);
    }
    return real[name](...args);
  });
  const halfOf = (fn) => ([file, data]) => {
    if (typeof file === 'number') return;
    real[fn](file, typeof data === 'string' ? data.slice(0, Math.ceil(data.length / 2)) : data.subarray(0, Math.ceil(data.length / 2)));
  };
  guard('writeFileSync', halfOf('writeFileSync'));
  guard('appendFileSync', halfOf('appendFileSync'));
  guard('renameSync');
  guard('rmSync');
  try {
    work();
    return tripped;
  } catch (err) {
    if (!(err instanceof Stopped)) throw err;
    return true;
  } finally {
    t.mock.restoreAll();
  }
}

function seeded() {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-crash-'));
  const s = new MailStore(root);
  for (let n = 1; n <= 4; n++) s.put('acc', 'INBOX', msg(n));
  return root;
}
const ids = (s, folder) => s.list('acc', folder, { limit: 100 }).rows.map((r) => r.subject).sort();
const eachCase = (fn) => async (t) => {
  for (const torn of [false, true]) {
    let stops = 0;
    for (let at = 1; at < 60; at++) {
      const root = seeded();
      const stopped = stopAt(t, at, torn, () => fn.work(new MailStore(root)));
      fn.after(new MailStore(root), `stopped at write ${at}${torn ? ', half done' : ''}`);
      fs.rmSync(root, { recursive: true, force: true });
      if (!stopped) break;
      stops += 1;
      assert.ok(at < 59, 'the work ends');
    }
    assert.ok(stops > 0, 'the work was stopped at least once');
  }
};
const names = (...ns) => ns.map((n) => `Message ${n}`).sort();
/** Each of the four is in one of the folders, listed once in it, and opens with its own words. */
const allThere = (s, folders, why) => {
  const seen = new Set();
  for (const f of folders) {
    const rows = s.list('acc', f, { limit: 100 }).rows;
    assert.equal(new Set(rows.map((r) => r.subject)).size, rows.length, `nothing listed twice in ${f}, ${why}`);
    for (const r of rows) {
      seen.add(r.subject);
      assert.equal(s.get('acc', f, r.id)?.text, `body ${r.subject.slice('Message '.length)}`, `${r.subject} opens with its words, ${why}`);
    }
  }
  for (let n = 1; n <= 4; n++) assert.ok(seen.has(`Message ${n}`), `Message ${n}, ${why}`);
};

test('a new message: the others stay, and fetching it again brings it in', eachCase({
  work: (s) => s.put('acc', 'INBOX', msg(5)),
  after: (s, why) => {
    allThere(s, ['INBOX'], why);
    s.put('acc', 'INBOX', msg(5));
    assert.ok(ids(s, 'INBOX').includes('Message 5'), `fetched again, ${why}`);
    s.put('acc', 'INBOX', msg(6));
    assert.deepEqual(ids(s, 'INBOX'), names(1, 2, 3, 4, 5, 6), `and the next one is not lost, ${why}`);
  },
}));

test('many new messages at once', eachCase({
  work: (s) => s.putMany('acc', 'INBOX', [msg(5), msg(6), msg(7)]),
  after: (s, why) => {
    allThere(s, ['INBOX'], why);
    s.putMany('acc', 'INBOX', [msg(5), msg(6), msg(7)]);
    assert.deepEqual(ids(s, 'INBOX'), names(1, 2, 3, 4, 5, 6, 7), `fetched again, ${why}`);
  },
}));

const markOneAndTwo = (s) => s.setFlagsMany('acc', 'INBOX', ['m1', 'm2'].map((id) => s.list('acc', 'INBOX', { limit: 100 }).rows.find((r) => r.subject === `Message ${id.slice(1)}`).id), { flagged: true });
test('marks on a message', eachCase({
  work: markOneAndTwo,
  after: (s, why) => {
    allThere(s, ['INBOX'], why);
    markOneAndTwo(s);
    for (const r of s.list('acc', 'INBOX', { limit: 100 }).rows) {
      const marked = ['Message 1', 'Message 2'].includes(r.subject);
      assert.equal(r.flagged, marked, `${r.subject} in the list, marked again, ${why}`);
      assert.equal(Boolean(s.get('acc', 'INBOX', r.id).flagged), marked, `${r.subject} opened, marked again, ${why}`);
    }
  },
}));

test('a message filed elsewhere is in one folder or the other, or both, and never neither', eachCase({
  work: (s) => {
    const rows = s.list('acc', 'INBOX', { limit: 100 }).rows.slice(0, 2);
    for (const r of rows) s.put('acc', 'Archive', { ...s.get('acc', 'INBOX', r.id), folder: 'INBOX' }, { force: true });
    s.removeMany('acc', 'INBOX', rows.map((r) => r.id));
  },
  after: (s, why) => allThere(s, ['INBOX', 'Archive'], why),
}));

const deleteTwo = (s) => s.removeMany('acc', 'INBOX', s.list('acc', 'INBOX', { limit: 100 }).rows.filter((r) => r.subject === 'Message 2').map((r) => r.id));
test('deleting some keeps the rest', eachCase({
  work: deleteTwo,
  after: (s, why) => {
    const left = ids(s, 'INBOX');
    for (const n of [1, 3, 4]) assert.ok(left.includes(`Message ${n}`), `Message ${n}, ${why}`);
    deleteTwo(s);
    assert.deepEqual(ids(s, 'INBOX'), names(1, 3, 4), `deleted again, ${why}`);
  },
}));

test('a folder renamed is under one name or the other', eachCase({
  work: (s) => {
    s.setFolders('acc', [{ path: 'INBOX', name: 'INBOX' }]);
    s.renameFolder('acc', 'INBOX', 'Kept');
  },
  after: (s, why) => allThere(s, ['INBOX', 'Kept'], why),
}));

test('filing through the mail service itself: each message is in one folder or the other', async (t) => {
  for (const torn of [false, true]) {
    let stops = 0;
    for (let at = 1; at < 60; at++) {
      const userData = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-crash-service-'));
      const settings = new Map();
      const stores = {
        settings: { get: (k, f) => (settings.has(k) ? settings.get(k) : f), set: (k, v) => settings.set(k, v), delete: (k) => settings.delete(k), all: () => Object.fromEntries(settings) },
        secrets: { available: () => true, get: () => null, set: () => {}, delete: () => {}, keys: () => [] },
      };
      const service = createMailService({ stores, holdBlob: () => ({ url: '' }), broadcast: () => {}, userData });
      const account = service.addAccount({ account: { email: 'me@example.com', name: 'Me', local: true }, password: 'x' });
      const root = path.join(userData, 'mail');
      const seed = new MailStore(root);
      for (let n = 1; n <= 4; n++) seed.put(account.id, 'INBOX', msg(n));
      const ids2 = seed.list(account.id, 'INBOX', { limit: 100 }).rows.slice(0, 2).map((r) => r.id);
      const stopped = stopAt(t, at, torn, () => service.move({ accountId: account.id, folder: 'INBOX', ids: ids2, to: 'Archive' }));
      const after = new MailStore(root);
      const seen = new Set(['INBOX', 'Archive'].flatMap((f) => after.list(account.id, f, { limit: 100 }).rows.map((r) => r.subject)));
      for (let n = 1; n <= 4; n++) assert.ok(seen.has(`Message ${n}`), `Message ${n}, stopped at write ${at}${torn ? ', half done' : ''}`);
      fs.rmSync(userData, { recursive: true, force: true });
      if (!stopped) break;
      stops += 1;
    }
    assert.ok(stops > 0, 'the filing was stopped at least once');
  }
});
