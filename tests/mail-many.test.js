// Messages leaving a folder take it out in one write of its index: rules
// that move or delete many, and Move to, call removeMany once rather than
// remove once a message; and the store's parsed folders are held to a limit
// without a folder read past it coming back wrong.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { MailStore } from '@rutba/mailbox/store';
import { applyPlan } from '../apps/desktop/main/mail-rules.js';

const message = (n) => ({ messageId: `m${n}@x`, subject: `Message ${n}`, from: [{ address: 'ana@work.example' }], text: 'Hello.' });

/** A store whose removals are counted. */
function counted(root) {
  const store = new MailStore(root);
  const calls = { remove: 0, removeMany: 0 };
  const remove = store.remove.bind(store);
  const removeMany = store.removeMany.bind(store);
  store.remove = (...a) => { calls.remove++; return remove(...a); };
  store.removeMany = (...a) => { calls.removeMany++; return removeMany(...a); };
  return { store, calls };
}

test('rules that move and delete many take them out of their folder in one write', () => {
  const { store, calls } = counted(fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-many-')));
  store.putMany('acc', 'INBOX', Array.from({ length: 30 }, (_, i) => message(i)));
  const rows = store.list('acc', 'INBOX', { limit: 100 }).rows;
  applyPlan(rows.map((row, i) => ({ row: { ...row, folder: 'INBOX' }, actions: [i % 3 ? { type: 'move', value: 'Filed' } : { type: 'delete' }] })), { store, accountId: 'acc', folderFor: () => null });
  assert.equal(calls.remove, 0);
  assert.equal(calls.removeMany, 1);
  assert.equal(store.list('acc', 'INBOX', { limit: 100 }).rows.length, 0);
  assert.equal(store.list('acc', 'Filed', { limit: 100 }).rows.length, 20);
});

test('a store reading more folders than it keeps parsed still answers each one rightly', () => {
  const store = new MailStore(fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-many-')));
  for (let f = 0; f < 80; f++) store.putMany('acc', `F${f}`, Array.from({ length: (f % 5) + 1 }, (_, i) => message(`${f}-${i}`)));
  for (let pass = 0; pass < 2; pass++) {
    for (let f = 0; f < 80; f++) assert.equal(store.list('acc', `F${f}`, { limit: 10 }).rows.length, (f % 5) + 1, `F${f}`);
  }
});
