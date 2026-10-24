// Import mail's look for other clients' mail (A6.12): it can be pointed at a
// folder of its own, so a check or a film never reads the mail of the person
// whose computer it runs on, and an empty folder finds nothing; and the
// Importing box names a file, not the folders it is in.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createDiscoveryService, placesOf } from '../apps/desktop/main/discover.js';
import { fileNameOf } from '../apps/desktop/renderer/apps/mail/parts.js';

const scanAt = (dir) => createDiscoveryService({ places: () => placesOf({ RUTBA_MAIL_IMPORT_HOME: dir }) }).scan();

test('pointed at an empty folder, the scan finds nothing and looks nowhere else', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-import-empty-'));
  try {
    const found = await scanAt(dir);
    assert.deepEqual(found.accounts, []);
    assert.deepEqual(found.files, []);
    assert.deepEqual(found.scanned, []);
    assert.equal(placesOf({ RUTBA_MAIL_IMPORT_HOME: dir }).registry, false, 'the Outlook registry is not read');
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('pointed at a folder made for it, the scan finds what is there and only that', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-import-made-'));
  try {
    const outlook = process.platform === 'win32' ? path.join(dir, 'AppData', 'Local', 'Microsoft', 'Outlook') : path.join(dir, 'Documents', 'Outlook Files');
    fs.mkdirSync(outlook, { recursive: true });
    fs.writeFileSync(path.join(outlook, 'sam@example.com.pst'), Buffer.alloc(64, 1));
    const found = await scanAt(dir);
    assert.equal(found.files.length, 1);
    assert.equal(fileNameOf(found.files[0].path), 'sam@example.com.pst');
    assert.ok(found.scanned.every((p) => p.startsWith(dir)), JSON.stringify(found.scanned));
    const account = found.accounts.find((a) => a.source === 'Outlook');
    assert.equal(account?.storeCount, 1);
    assert.equal(account.incoming, null, 'no host read from the registry');
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('with no folder named, the scan looks in the person\'s own folders', () => {
  const at = placesOf({ APPDATA: 'C:/Users/sam/AppData/Roaming', LOCALAPPDATA: 'C:/Users/sam/AppData/Local' });
  assert.equal(at.registry, true);
  assert.equal(at.appData, 'C:/Users/sam/AppData/Roaming');
  assert.equal(at.home, os.homedir());
});

test('a file found or chosen is shown by its name alone', () => {
  assert.equal(fileNameOf('C:\\Users\\sam\\AppData\\Local\\Microsoft\\Outlook\\sam@example.com.ost'), 'sam@example.com.ost');
  assert.equal(fileNameOf('/home/sam/.thunderbird/abc.default/Mail/Local Folders/Inbox'), 'Inbox');
  assert.equal(fileNameOf('archive.mbox'), 'archive.mbox');
  assert.equal(fileNameOf(''), '');
});
