// The checks' mailbox (apps/desktop/main/seed-mail.js) is written by the core
// itself: @rutba/mailbox is Mail's pack's, which an installed core cannot
// reach, so the packaged smoke once seeded nothing and photographed an empty
// Mail. The file must be exactly what Mail's own writer makes, and read back
// by Mail's own reader as every message.
import test from 'node:test';
import assert from 'node:assert/strict';
import { seedMessages, toMbox } from '../apps/desktop/main/seed-mail.js';
import { writeMbox, scanMbox } from '@rutba/mailbox/mbox';

test('the checks\' mailbox is the same bytes Mail writes, and reads back whole', () => {
  const messages = seedMessages(Date.UTC(2026, 9, 27, 12));
  const ours = toMbox(messages);
  assert.ok(Buffer.from(writeMbox(messages)).equals(ours));
  assert.match(ours.toString('utf8'), /^From amina@northwind\.example /);
  assert.equal(scanMbox(ours).ranges.length, messages.length);
});

test('a From at the start of a body line is quoted, as an mbox reader expects', () => {
  const m = { from: { address: 'a@b.example' }, date: '2026-10-27T00:00:00Z', raw: 'Subject: x\r\n\r\nFrom here on, yes.\r\n>From before.\r\n' };
  assert.ok(Buffer.from(writeMbox([m])).equals(toMbox([m])));
  assert.match(toMbox([m]).toString('utf8'), /\n>From here on/);
  assert.match(toMbox([m]).toString('utf8'), /\n>>From before/);
});
