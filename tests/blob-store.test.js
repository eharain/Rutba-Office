// What the main process holds for the windows, and when it lets go.
//
// Nothing ever did: every message read held its pictures again, every
// attachment saved held the file again, and the main process grew for as
// long as the suite ran. A held blob now belongs to the window whose request
// made it and goes when that window closes; a blob held for the message
// being read gives way to the next message's.

import test from 'node:test';
import assert from 'node:assert/strict';
import { blobOwner, hold, heldBlob, releaseOwner, heldTotals } from '../packages/office-shell/src/electron/blobs.js';

const asWindow = (id, fn) => blobOwner.run(id, fn);

test('a blob belongs to the window that asked for it, across awaits, and goes when that window closes', async () => {
  const mine = await asWindow(101, async () => {
    await new Promise((resolve) => setTimeout(resolve, 5));
    return hold(Buffer.from('a picture'), 'image/png', 'a.png');
  });
  const theirs = asWindow(102, () => hold(Buffer.from('another'), 'image/png', 'b.png'));
  assert.equal(heldBlob(mine.id).owner, 101, 'the owner survives an await inside the request');
  releaseOwner(101);
  assert.equal(heldBlob(mine.id), undefined, 'the closed window\'s blob is gone');
  assert.ok(heldBlob(theirs.id), 'another window\'s is not touched');
  releaseOwner(102);
});

test('the next message read replaces the last one\'s pictures, and a message\'s own pictures keep each other', () => {
  const first = asWindow(201, () => [
    hold(Buffer.from('logo'), 'image/png', 'logo.png', { group: 'mail-message', generation: 'm1' }),
    hold(Buffer.from('chart'), 'image/png', 'chart.png', { group: 'mail-message', generation: 'm1' }),
  ]);
  assert.ok(first.every((b) => heldBlob(b.id)), 'both pictures of one message are held');
  const second = asWindow(201, () => hold(Buffer.from('photo'), 'image/jpeg', 'photo.jpg', { group: 'mail-message', generation: 'm2' }));
  assert.ok(first.every((b) => heldBlob(b.id) === undefined), 'the first message\'s pictures went');
  assert.ok(heldBlob(second.id), 'the second message\'s is held');
  // The same message in another window keeps its own.
  const elsewhere = asWindow(202, () => hold(Buffer.from('logo'), 'image/png', 'logo.png', { group: 'mail-message', generation: 'm1' }));
  asWindow(201, () => hold(Buffer.from('x'), 'image/png', 'x.png', { group: 'mail-message', generation: 'm3' }));
  assert.ok(heldBlob(elsewhere.id), 'a group is per window');
  releaseOwner(201);
  releaseOwner(202);
});

test('reading a hundred messages in one window holds one message\'s pictures, not a hundred', () => {
  const before = heldTotals();
  for (let i = 0; i < 100; i++) {
    asWindow(301, () => hold(Buffer.alloc(10_000), 'image/png', 'p.png', { group: 'mail-message', generation: `m${i}` }));
  }
  const after = heldTotals();
  assert.equal(after.count - before.count, 1);
  assert.equal(after.bytes - before.bytes, 10_000);
  releaseOwner(301);
  assert.deepEqual(heldTotals(), before);
});
