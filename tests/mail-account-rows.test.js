// Many accounts at once, and each set up by trying hard: rows of addresses
// and passwords read however they were written, and the servers discovery
// found tried in turn, with both user names, until one signs in - with no
// network, the tries handed in.

import test from 'node:test';
import assert from 'node:assert/strict';
import { parseAccountRows, isAddress, splitRow } from '@rutba/mailbox/account-rows';
import { setupAccount, failureOf, serverSoftware } from '../apps/desktop/main/mail-setup.js';

test('rows of addresses and passwords read whatever separates them, the password whole', () => {
  const text = [
    '# the office, set up on Monday',
    'ali@northwind.example secret one',
    'bo@northwind.example\tTab$eparated',
    'cy@northwind.example,comma,in,it',
    'di@northwind.example;semi',
    'ed@northwind.example:colon:too',
    'fa@northwind.example | barred',
    'gu@northwind.example, "quoted, with a comma"',
    '',
    'not an address at all',
    'ali@northwind.example again',
    'hu@northwind.example',
  ].join('\r\n');
  const { rows, header } = parseAccountRows(`\uFEFF${text}`);
  assert.equal(header, null);
  const by = Object.fromEntries(rows.map((r) => [`${r.line}`, r]));
  assert.deepEqual(rows.filter((r) => !r.problem).map((r) => [r.email, r.password]), [
    ['ali@northwind.example', 'secret one'],
    ['bo@northwind.example', 'Tab$eparated'],
    ['cy@northwind.example', 'comma,in,it'],
    ['di@northwind.example', 'semi'],
    ['ed@northwind.example', 'colon:too'],
    ['fa@northwind.example', 'barred'],
    ['gu@northwind.example', 'quoted, with a comma'],
  ]);
  assert.equal(by['10'].problem, 'unreadable', 'a line that does not start with an address');
  assert.equal(by['10'].email, '', 'and its text is not kept to be shown back');
  assert.equal(by['11'].problem, 'repeated');
  assert.equal(by['12'].problem, 'no-password');
  assert.equal(rows.length, 10, 'the comment and the blank line are not rows');
});

test('a table with a header naming its columns reads as a spreadsheet writes it', () => {
  const csv = 'Email,Password,Name,IMAP server,IMAP port,SMTP server\r\nali@northwind.example,"p,w ""1""",Ali Khan,mail.northwind.example,993,smtp.northwind.example\r\nbo@contoso.example,second,,,,\r\n';
  const { rows, header, separator } = parseAccountRows(csv);
  assert.equal(separator, ',');
  assert.deepEqual(header, ['email', 'password', 'name', 'imapHost', 'imapPort', 'smtpHost']);
  assert.equal(rows[0].password, 'p,w "1"');
  assert.equal(rows[0].name, 'Ali Khan');
  assert.deepEqual(rows[0].imap, { host: 'mail.northwind.example', port: 993 });
  assert.deepEqual(rows[0].smtp, { host: 'smtp.northwind.example', port: 465 });
  assert.equal(rows[1].imap, undefined);
  const tsv = parseAccountRows('e-mail\tpassword\nali@northwind.example\tone two\n');
  assert.equal(tsv.separator, '\t');
  assert.equal(tsv.rows[0].password, 'one two');
  assert.ok(isAddress('a.b+c@sub.example.co.uk') && !isAddress('a@b') && !isAddress('ali northwind'));
  assert.deepEqual(splitRow('a  b   c', ' '), ['a', 'b', 'c']);
});

/** A discovery and servers of the test's own: who answers, and to which name and password. */
function servers({ found, accept = {}, unreachable = [] }) {
  const tried = [];
  const tryAt = (kind) => async (server, user, password) => {
    tried.push(`${kind} ${server.host}:${server.port} ${user}`);
    if (unreachable.includes(server.host)) return { ok: false, message: 'connect ECONNREFUSED' };
    const want = accept[`${kind} ${server.host}`];
    if (want && want.user === user && want.password === password) return { ok: true, folders: 7, banner: want.banner };
    return { ok: false, message: 'AUTHENTICATIONFAILED Invalid credentials (Failure)' };
  };
  return { discover: async () => found, tryImap: tryAt('imap'), trySmtp: tryAt('smtp'), tried };
}

test('setting up tries every server found, best first, with both user names, until one signs in', async () => {
  const found = {
    provider: null,
    imap: { host: 'imap.northwind.example', port: 993, secure: true },
    smtp: { host: 'smtp.northwind.example', port: 465, secure: true },
    candidates: [
      { kind: 'imap', host: 'imap.northwind.example', port: 993, secure: true, confidence: 2 },
      { kind: 'imap', host: 'mail.northwind.example', port: 993, secure: true, confidence: 1, verified: true, banner: '* OK Dovecot ready.' },
      { kind: 'smtp', host: 'smtp.northwind.example', port: 465, secure: true, confidence: 2 },
      { kind: 'smtp', host: 'mail.northwind.example', port: 587, secure: false, confidence: 1 },
    ],
  };
  const s = servers({
    found,
    unreachable: ['imap.northwind.example'],
    accept: { 'imap mail.northwind.example': { user: 'ali', password: 'pw', banner: 'Dovecot IMAP4rev1' }, 'smtp mail.northwind.example': { user: 'ali', password: 'pw' } },
  });
  const r = await setupAccount({ email: 'ali@northwind.example', password: 'pw', ...s });
  assert.equal(r.outcome, 'ready');
  assert.deepEqual(r.account.imap, { host: 'mail.northwind.example', port: 993, secure: true, starttls: false, user: 'ali' });
  assert.equal(r.account.smtp.host, 'mail.northwind.example');
  assert.equal(r.server, 'Dovecot');
  assert.equal(r.folders, 7);
  // The unreachable server once, not once a name; the name before the @ after the whole address.
  assert.deepEqual(s.tried.slice(0, 3), ['imap imap.northwind.example:993 ali@northwind.example', 'imap mail.northwind.example:993 ali@northwind.example', 'imap mail.northwind.example:993 ali']);
  assert.ok(r.tried.every((t) => !('password' in t)), 'no password in what was tried');
});

test('setting up says which failure it was: a refused password, an app password wanted, a browser sign-in, nothing found, nothing answering', async () => {
  const google = { provider: { id: 'google', label: 'Google' }, oauth: 'google', appPassword: { url: 'x' }, imap: { host: 'imap.gmail.com', port: 993, secure: true }, smtp: { host: 'smtp.gmail.com', port: 465, secure: true }, candidates: [] };
  assert.equal((await setupAccount({ email: 'a@gmail.com', password: '', ...servers({ found: google }) })).outcome, 'sign-in');
  assert.equal((await setupAccount({ email: 'a@gmail.com', password: 'wrong', ...servers({ found: google }) })).outcome, 'app-password');
  const plain = { provider: null, imap: { host: 'mail.x.example', port: 993, secure: true }, smtp: null, candidates: [] };
  assert.equal((await setupAccount({ email: 'a@x.example', password: 'wrong', ...servers({ found: plain }) })).outcome, 'password');
  assert.equal((await setupAccount({ email: 'a@x.example', password: 'pw', ...servers({ found: plain, unreachable: ['mail.x.example'] }) })).outcome, 'unreachable');
  assert.equal((await setupAccount({ email: 'a@y.example', password: 'pw', ...servers({ found: { imap: null, smtp: null, candidates: [] } }) })).outcome, 'not-found');
  // A row's own servers are tried before anything found.
  const s = servers({ found: plain, accept: { 'imap own.y.example': { user: 'a@y.example', password: 'pw' }, 'smtp own.y.example': { user: 'a@y.example', password: 'pw' } } });
  const own = await setupAccount({ email: 'a@y.example', password: 'pw', imap: { host: 'own.y.example', port: 993 }, smtp: { host: 'own.y.example', port: 465 }, ...s });
  assert.equal(own.outcome, 'ready');
  assert.equal(own.account.imap.host, 'own.y.example');
});

test('a failure is read from what the server said, and the server named from its greeting', () => {
  assert.equal(failureOf('[ALERT] Application-specific password required: https://support.google.com/accounts/answer/185833 (Failure)'), 'app-password');
  assert.equal(failureOf('535 5.7.8 Username and Password not accepted'), 'password');
  assert.equal(failureOf('getaddrinfo ENOTFOUND imap.nowhere.example'), 'no-server');
  assert.equal(failureOf('connect ETIMEDOUT 10.0.0.1:993'), 'unreachable');
  assert.equal(failureOf('self signed certificate in certificate chain'), 'certificate');
  assert.equal(serverSoftware('* OK Gimap ready for requests from 1.2.3.4'), 'Gmail');
  assert.equal(serverSoftware('* OK The Microsoft Exchange IMAP4 service is ready.'), 'Microsoft Exchange');
  assert.equal(serverSoftware('* OK [CAPABILITY IMAP4rev1] Dovecot ready.'), 'Dovecot');
  assert.equal(serverSoftware('220 mx.example ESMTP Postfix'), 'Postfix');
  assert.equal(serverSoftware('* OK hello'), null);
});
