// Mail's View source, for a message with no server to ask: written out again
// from what Mail kept, it reads back as the same message, and says it was
// rebuilt.
import test from 'node:test';
import assert from 'node:assert/strict';
import { parseMessage } from '@rutba/mailbox/mime';
import { rebuildSource } from '../apps/desktop/main/mail.js';

const RAW = [
  'From: Amira Khan <amira@example.com>',
  'To: Ben Ortiz <ben@example.com>',
  'Subject: Board pack, =?utf-8?B?ZmluYWw=?=',
  'Date: Tue, 20 Oct 2026 09:00:00 +0000',
  'Message-ID: <source-1@example.com>',
  'X-Mailer: Something',
  'MIME-Version: 1.0',
  'Content-Type: multipart/mixed; boundary="b1"',
  '',
  '--b1',
  'Content-Type: multipart/alternative; boundary="b2"',
  '',
  '--b2',
  'Content-Type: text/plain; charset=utf-8',
  '',
  'The board pack is attached. Café at nine.',
  '--b2',
  'Content-Type: text/html; charset=utf-8',
  '',
  '<p>The board pack is <b>attached</b>. Café at nine.</p>',
  '--b2--',
  '--b1',
  'Content-Type: application/pdf; name="pack.pdf"',
  'Content-Disposition: attachment; filename="pack.pdf"',
  'Content-Transfer-Encoding: base64',
  '',
  Buffer.from('%PDF-1.4 a small pack').toString('base64'),
  '--b1--',
  '',
].join('\r\n');

test('View source, rebuilt: the same headers, text, page and attachment, and a header saying it was rebuilt', () => {
  const kept = parseMessage(RAW);
  const att = kept.attachments.find((a) => a.filename === 'pack.pdf');
  assert.ok(att, 'the parser keeps the attachment');
  const bytes = Buffer.from('%PDF-1.4 a small pack');
  const text = rebuildSource({ ...kept, attachments: kept.attachments.map((a) => ({ ...a, stored: 'x' })) }, () => bytes);
  assert.match(text, /^X-Rutba-Rebuilt: /m);
  assert.match(text, /^Message-ID: <source-1@example\.com>/m);
  assert.match(text, /^X-Mailer: Something/m);
  assert.equal((text.match(/^MIME-Version:/gm) || []).length, 1);
  const again = parseMessage(text);
  assert.equal(again.subject, kept.subject);
  assert.equal(again.from?.[0]?.address, 'amira@example.com');
  assert.match(again.text, /Café at nine/);
  assert.match(again.html, /<b>attached<\/b>/);
  const back = again.attachments.find((a) => a.filename === 'pack.pdf');
  assert.ok(back, 'the attachment comes back');
  assert.equal(back.type, 'application/pdf');
});
