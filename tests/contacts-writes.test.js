// Contacts written once for a burst of changes: fifty cards saved in a row
// are one write a moment after the last, not fifty rewrites of the file, and
// the file then holds them all for the next run to read.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createContactsService } from '../apps/desktop/main/contacts.js';

test('a burst of saves is one write, and the file holds every card', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-contacts-'));
  const stores = { dir };
  const file = path.join(dir, 'contacts.json');
  const writes = [];
  const realRename = fs.promises.rename;
  fs.promises.rename = async (from, to) => { if (to === file) writes.push(to); return realRename(from, to); };
  try {
    const contacts = createContactsService({ stores, broadcast: null });
    for (let i = 0; i < 50; i++) contacts.save({ contact: { names: { given: `Person ${i}` }, emails: [{ value: `p${i}@example.com` }] } });
    assert.equal(contacts.count(), 50, 'all there at once');
    await new Promise((r) => setTimeout(r, 600));
    assert.equal(writes.length, 1, 'one write for the burst');
    assert.equal(JSON.parse(fs.readFileSync(file, 'utf8')).contacts.length, 50);
    assert.equal(createContactsService({ stores, broadcast: null }).count(), 50, 'read back by the next run');
  } finally {
    fs.promises.rename = realRename;
  }
});
