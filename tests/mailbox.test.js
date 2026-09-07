// The mail readers, on messages built to contain what real mail contains:
// encoded words, folded headers, multipart bodies, an attachment, and a
// character set that is not UTF-8.

import test from 'node:test';
import assert from 'node:assert/strict';
import { parseMessage, parseAddresses, decodeWords, parseParameters, stripHtml } from '@rutba/mailbox/mime';
import { Mbox, writeMbox } from '@rutba/mailbox/mbox';
import { MailStore, messageKey } from '@rutba/mailbox/store';
import { verifyTables } from '../packages/mailbox/src/pst/tables.js';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

const CRLF = '\r\n';
const message = [
  'Return-Path: <sender@example.com>',
  'Received: from mail.example.com (mail.example.com [203.0.113.4])',
  '\tby mx.example.net with ESMTPS id abc123',
  '\tfor <someone@example.net>; Tue, 11 Mar 2026 09:14:02 +0000',
  'From: =?utf-8?B?w4RubmUgU2NobWlkdA==?= <anne@example.com>',
  'To: "Person, A" <a@example.net>, b@example.net',
  'Cc: =?iso-8859-1?Q?Jos=E9?= <jose@example.com>',
  'Subject: =?utf-8?Q?Quarterly_review_=E2=80=94_Q3?=',
  'Date: Tue, 11 Mar 2026 09:14:02 +0000',
  'Message-ID: <abc-123@example.com>',
  'In-Reply-To: <prev-1@example.com>',
  'MIME-Version: 1.0',
  'Content-Type: multipart/mixed; boundary="OUTER"',
  '',
  '--OUTER',
  'Content-Type: multipart/alternative; boundary="INNER"',
  '',
  '--INNER',
  'Content-Type: text/plain; charset=utf-8',
  'Content-Transfer-Encoding: quoted-printable',
  '',
  'Revenue rose 18=25 this quarter =E2=80=94 details attached.',
  '',
  '--INNER',
  'Content-Type: text/html; charset=utf-8',
  'Content-Transfer-Encoding: base64',
  '',
  Buffer.from('<p>Revenue rose <b>18%</b> this quarter.</p>', 'utf8').toString('base64'),
  '',
  '--INNER--',
  '',
  '--OUTER',
  'Content-Type: text/csv; charset=utf-8',
  'Content-Disposition: attachment; filename*0*=utf-8\'\'%51%33%20; filename*1*=figures.csv',
  'Content-Transfer-Encoding: base64',
  '',
  Buffer.from('region,revenue\nNorth,1420\n', 'utf8').toString('base64'),
  '',
  '--OUTER--',
  '',
].join(CRLF);

test('a multipart message parses into text, html and an attachment', () => {
  const m = parseMessage(message);
  assert.equal(m.subject, 'Quarterly review — Q3');
  assert.equal(m.from[0].name, 'Änne Schmidt');
  assert.equal(m.from[0].address, 'anne@example.com');
  assert.equal(m.to.length, 2);
  assert.equal(m.to[0].name, 'Person, A', 'a comma inside quotes is not an address separator');
  assert.equal(m.cc[0].name, 'José', 'the header was latin-1');
  assert.equal(m.messageId, 'abc-123@example.com');
  assert.equal(m.inReplyTo, 'prev-1@example.com');
  assert.match(m.text, /Revenue rose 18% this quarter — details attached\./);
  assert.match(m.html, /<b>18%<\/b>/);
  assert.equal(m.attachments.length, 1);
  assert.equal(m.attachments[0].filename, 'Q3 figures.csv', 'RFC 2231 continuation and percent-encoding');
  assert.equal(m.attachments[0].type, 'text/csv');
  assert.match(new TextDecoder().decode(m.attachments[0].bytes), /North,1420/);
  assert.equal(m.hasAttachments, true);
});

test('a folded header is unfolded before it is read', () => {
  const m = parseMessage(message);
  const received = m.headers.find((h) => h.key === 'received');
  assert.match(received.value, /mx\.example\.net with ESMTPS/);
  assert.doesNotMatch(received.value, /\n/);
});

test('the date becomes a real timestamp', () => {
  const m = parseMessage(message);
  assert.equal(m.date, '2026-03-11T09:14:02.000Z');
});

test('addresses parse in every legal shape', () => {
  assert.deepEqual(
    parseAddresses('a@b.c, "Last, First" <d@e.f>, g@h.i (A Person)').map((x) => x.address),
    ['a@b.c', 'd@e.f', 'g@h.i']
  );
  assert.equal(parseAddresses('g@h.i (A Person)')[0].name, 'A Person');
  assert.deepEqual(parseAddresses(''), []);
});

test('encoded words that sit next to each other join without a space', () => {
  assert.equal(decodeWords('=?utf-8?Q?One?= =?utf-8?Q?Two?='), 'OneTwo');
  assert.equal(decodeWords('plain text'), 'plain text');
});

test('content-type parameters survive quoting and continuation', () => {
  const { value, params } = parseParameters('text/plain; charset="utf-8"; name="a b.txt"');
  assert.equal(value, 'text/plain');
  assert.equal(params.charset, 'utf-8');
  assert.equal(params.name, 'a b.txt');
});

test('html is reduced to readable text for previews', () => {
  assert.equal(stripHtml('<p>Hello<br>world</p><script>bad()</script>'), 'Hello\nworld');
});

test('an mbox splits, unescapes and reads back', () => {
  const archive = [
    'From sender@example.com Tue Mar 11 09:14:02 2026',
    'From: a@b.c',
    'Subject: First',
    '',
    'Body one.',
    '>From a line that needed escaping',
    '',
    'From other@example.com Tue Mar 11 10:00:00 2026',
    'From: d@e.f',
    'Subject: Second',
    '',
    'Body two.',
    '',
  ].join('\n');

  const mbox = new Mbox(archive);
  assert.equal(mbox.length, 2);
  assert.equal(mbox.message(0).subject, 'First');
  assert.match(mbox.message(0).text, /^From a line that needed escaping$/m, 'the > was removed');
  assert.equal(mbox.message(1).subject, 'Second');

  const summaries = mbox.summaries();
  assert.deepEqual(summaries.map((s) => s.subject), ['First', 'Second']);
});

test('messages written back as mbox can be read again', () => {
  const out = writeMbox([{ raw: message, from: { address: 'anne@example.com' }, date: '2026-03-11T09:14:02Z' }]);
  const back = new Mbox(out);
  assert.equal(back.length, 1);
  assert.equal(back.message(0).subject, 'Quarterly review — Q3');
});

test('the PST cipher tables are the ones the specification publishes', () => {
  const result = verifyTables();
  assert.deepEqual(result.problems, []);
  assert.equal(result.ok, true);
});

test('the store keeps messages, indexes them, and does not double them', () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-mail-'));
  try {
    const store = new MailStore(root);
    const m = parseMessage(message);
    const id = store.put('acct', 'Inbox', m);
    assert.ok(id);
    assert.equal(store.put('acct', 'Inbox', m), null, 'the same message is not stored twice');

    const listed = store.list('acct', 'Inbox');
    assert.equal(listed.total, 1);
    assert.equal(listed.rows[0].subject, 'Quarterly review — Q3');
    assert.equal(listed.rows[0].hasAttachments, true);

    const full = store.get('acct', 'Inbox', id);
    assert.match(full.text, /Revenue rose 18%/);
    assert.equal(full.attachments[0].stored, '00-Q3_figures.csv');
    assert.ok(fs.existsSync(store.attachmentPath('acct', 'Inbox', id, full.attachments[0].stored)));

    store.setFlags('acct', 'Inbox', id, { unread: false, flagged: true });
    assert.equal(store.list('acct', 'Inbox').rows[0].unread, false);
    assert.equal(store.list('acct', 'Inbox').rows[0].flagged, true);

    assert.equal(store.list('acct', 'Inbox', { query: 'quarterly' }).total, 1);
    assert.equal(store.list('acct', 'Inbox', { query: 'nothing here' }).total, 0);

    store.remove('acct', 'Inbox', id);
    assert.equal(store.list('acct', 'Inbox').total, 0);
  } finally {
    fs.rmSync(root, { recursive: true, force: true });
  }
});

test('the same message from two sources gets the same id', () => {
  const a = parseMessage(message);
  const b = parseMessage(message);
  assert.equal(messageKey(a), messageKey(b));
});
