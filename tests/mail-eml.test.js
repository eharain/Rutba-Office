// Messages as .eml files (Mail 2.25): saved one to a file or several into a
// folder; brought in with their date and read state, APPENDed to a live
// account's server; a file opened from Explorer kept apart, read only.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { MailStore } from '@rutba/mailbox/store';
import { parseMessage } from '@rutba/mailbox/mime';
import { createMailService, readableBody } from '../apps/desktop/main/mail.js';
import { emlMarks, appendFlags, emlFileName, freeName, looksLikeMessage, isOpened } from '../apps/desktop/main/mail-eml.js';
import { imapServer } from '../packages/mailbox/test/imap-server.js';

const dirs = [];
test.after(() => {
  for (const d of dirs) fs.rmSync(d, { recursive: true, force: true });
});
const tempDir = (name) => {
  const d = fs.mkdtempSync(path.join(os.tmpdir(), `rutba-mail-eml-${name}-`));
  dirs.push(d);
  return d;
};

const eml = (n, extra = []) => [
  'From: Amina Yusuf <amina@work.example>', 'To: Sam <sam@checks.example>', `Subject: File message ${n}`,
  `Date: ${new Date(Date.UTC(2025, 2, 4, 8, n)).toUTCString()}`, `Message-ID: <file${n}@work.example>`, ...extra,
  'MIME-Version: 1.0', 'Content-Type: multipart/mixed; boundary="b"', '',
  '--b', 'Content-Type: text/plain; charset=utf-8', '', `Body of file ${n}.`, '',
  '--b', 'Content-Type: application/pdf; name="note.pdf"', 'Content-Disposition: attachment; filename="note.pdf"', 'Content-Transfer-Encoding: base64', '',
  Buffer.from('%PDF-1.4 a note').toString('base64'), '--b--', '',
].join('\r\n');

function fixture() {
  const settings = new Map();
  const stores = {
    settings: { get: (k, f) => (settings.has(k) ? structuredClone(settings.get(k)) : f), set: (k, v) => settings.set(k, structuredClone(v)), delete: (k) => settings.delete(k), all: () => Object.fromEntries(settings) },
    secrets: { available: () => true, get: () => 'check', set: () => {}, delete: () => {}, keys: () => [] },
  };
  const userData = tempDir('data');
  const service = createMailService({ stores, holdBlob: (bytes) => ({ url: 'blob:x', bytes }), broadcast: () => {}, userData, smtp: { send: async () => ({ messageId: 'x' }) }, autoFetch: false });
  const add = (email, { live = false } = {}) => service.addAccount({
    account: { email, name: email, local: !live, imap: { host: live ? 'imap.invalid' : '', port: 993, secure: true }, smtp: { host: live ? 'smtp.invalid' : '', port: 465 } },
    password: 'x',
  });
  const store = new MailStore(path.join(userData, 'mail'));
  const rows = (accountId, folder) => service.messages({ accountId, folder }).rows;
  return { service, store, add, rows, userData, stores };
}

test('the marks a file carries: Status, X-Status and X-Mozilla-Status, and read where none is written', () => {
  const b = (headers) => Buffer.from([...headers, 'Subject: x', '', 'body'].join('\r\n'));
  assert.deepEqual(emlMarks(b([])), { unread: false, flagged: false, answered: false });
  assert.deepEqual(emlMarks(b(['Status: O'])), { unread: true, flagged: false, answered: false });
  assert.deepEqual(emlMarks(b(['Status: RO', 'X-Status: AF'])), { unread: false, flagged: true, answered: true });
  assert.deepEqual(emlMarks(b(['X-Mozilla-Status: 0004'])), { unread: true, flagged: true, answered: false });
  assert.deepEqual(emlMarks(b(['X-Mozilla-Status: 0003'])), { unread: false, flagged: false, answered: true });
  assert.deepEqual(appendFlags({ unread: false, flagged: true, answered: true }), ['\\Seen', '\\Flagged', '\\Answered']);
  assert.deepEqual(appendFlags({ unread: true }), []);
  assert.equal(looksLikeMessage(Buffer.from(eml(1))), true);
  assert.equal(looksLikeMessage(Buffer.from('%PDF-1.4 not mail at all')), false);
  assert.equal(looksLikeMessage(Buffer.from('')), false);
});

test('file names: the subject made safe, never empty, and never one already taken', () => {
  assert.equal(emlFileName({ subject: 'Re: Q3 / numbers? <final>' }), 'Re Q3 numbers final.eml');
  assert.equal(emlFileName({ subject: '' }), 'message.eml');
  assert.equal(emlFileName({ subject: 'CON' }), 'CON_.eml');
  assert.equal(emlFileName({ subject: 'Trailing dots...' }), 'Trailing dots.eml');
  const dir = tempDir('names');
  fs.writeFileSync(path.join(dir, 'Same.eml'), 'x');
  const taken = new Set();
  assert.equal(freeName(dir, 'Same.eml', taken), 'Same (2).eml');
  assert.equal(freeName(dir, 'Same.eml', taken), 'Same (3).eml');
});

test('import: files on disk and files handed over as bytes come in with their date, read state and attachment; what is not mail is passed over', async () => {
  const f = fixture();
  const a = f.add('old@example.com');
  const dir = tempDir('in');
  const one = path.join(dir, 'one.eml');
  const two = path.join(dir, 'two.eml');
  const notMail = path.join(dir, 'report.eml');
  fs.writeFileSync(one, eml(1));
  fs.writeFileSync(two, eml(2, ['Status: O', 'X-Status: F']));
  fs.writeFileSync(notMail, '%PDF-1.4 not a message');
  const r = await f.service.importEml({ accountId: a.id, folder: 'Kept', paths: [one, two, notMail, path.join(dir, 'gone.eml')], files: [{ name: 'three.eml', bytes: new Uint8Array(Buffer.from(eml(3))) }] });
  assert.deepEqual(r, { added: 3, skipped: 2, folder: 'Kept' });
  const rows = f.rows(a.id, 'Kept');
  const by = Object.fromEntries(rows.map((x) => [x.subject, x]));
  assert.deepEqual(Object.keys(by).sort(), ['File message 1', 'File message 2', 'File message 3']);
  assert.equal(by['File message 1'].unread, false, 'a file with no marks was read');
  assert.equal(by['File message 2'].unread, true, 'Status: O is unread');
  assert.equal(by['File message 2'].flagged, true);
  assert.equal(new Date(by['File message 1'].date).toISOString(), '2025-03-04T08:01:00.000Z', 'the date it was written, not today');
  const kept = f.store.get(a.id, 'Kept', by['File message 1'].id);
  assert.equal(Buffer.from(f.store.attachmentBytes(a.id, 'Kept', kept.id, kept.attachments[0].stored)).toString(), '%PDF-1.4 a note');
  assert.ok(f.service.folders({ accountId: a.id }).some((x) => x.path === 'Kept'), 'the folder is listed');
  await assert.rejects(() => f.service.importEml({ accountId: a.id, folder: '', paths: [one] }), /folder to go to/);
  await assert.rejects(() => f.service.importEml({ accountId: 'archive:x', folder: 'Inbox', paths: [one] }), /read only/);
  await assert.rejects(() => f.service.importEml({ accountId: 'nobody', folder: 'Inbox', paths: [one] }), /not set up/);
});

test('import into a live account: an APPEND with the file as it is, its marks and its date', async () => {
  const f = fixture();
  const a = f.add('me@example.com', { live: true });
  const dir = tempDir('live');
  const p = path.join(dir, 'x.eml');
  fs.writeFileSync(p, eml(4, ['X-Mozilla-Status: 0005']));
  await f.service.importEml({ accountId: a.id, folder: 'INBOX', paths: [p] });
  const appends = f.store.changes(a.id).items.filter((c) => c.kind === 'append');
  assert.equal(appends.length, 1);
  assert.deepEqual(appends[0].flags.sort(), ['\\Flagged', '\\Seen']);
  assert.equal(new Date(appends[0].date).toISOString(), '2025-03-04T08:04:00.000Z');
  assert.equal(fs.readFileSync(appends[0].file, 'latin1'), eml(4, ['X-Mozilla-Status: 0005']), 'the bytes of the file, unchanged');
});

test('save as .eml: one message to a file, several into a folder, each named by its subject', async () => {
  const f = fixture();
  const a = f.add('old@example.com');
  const dir = tempDir('src');
  const files = [1, 2, 3].map((n) => {
    const p = path.join(dir, `${n}.eml`);
    fs.writeFileSync(p, eml(n));
    return p;
  });
  // Two with the same subject: the second file takes a name of its own.
  fs.writeFileSync(files[2], eml(3).replace('Subject: File message 3', 'Subject: File message 2'));
  await f.service.importEml({ accountId: a.id, folder: 'Inbox', paths: files });
  const ids = f.rows(a.id, 'Inbox').map((r) => r.id);
  const out = tempDir('out');
  const single = path.join(out, 'chosen.eml');
  const r1 = await f.service.saveEml({ accountId: a.id, folder: 'Inbox', ids: [ids[0]], path: single });
  assert.deepEqual(r1.written, ['chosen.eml']);
  const again = await f.service.importEml({ accountId: a.id, folder: 'Again', paths: [single] });
  assert.equal(again.added, 1, 'a saved file is one Mail reads back');
  const folder = path.join(out, 'several');
  const r2 = await f.service.saveEml({ accountId: a.id, folder: 'Inbox', ids, path: folder });
  assert.deepEqual(r2.written.slice().sort(), ['File message 1.eml', 'File message 2 (2).eml', 'File message 2.eml']);
  for (const name of r2.written) assert.match(fs.readFileSync(path.join(folder, name), 'latin1'), /^Subject: File message/m);
  await assert.rejects(() => f.service.saveEml({ accountId: a.id, folder: 'Inbox', ids: ['gone'], path: single }), /not there/);
});

test('save as .eml: the file saved is the message, its words readable as they are, whatever their script', async () => {
  const f = fixture();
  const a = f.add('words@example.com');
  const urdu = 'سلام، کل دس بجے ملتے ہیں۔ Café at ten.';
  const long = 'A long line. '.repeat(100).trim();
  const raws = [
    eml(7),
    ['From: Amina Yusuf <amina@work.example>', 'To: Sam <sam@checks.example>', 'Subject: =?UTF-8?B?' + Buffer.from('ملاقات').toString('base64') + '?=',
      `Date: ${new Date(Date.UTC(2025, 2, 4, 9, 0)).toUTCString()}`, 'Message-ID: <urdu@work.example>', 'MIME-Version: 1.0',
      'Content-Type: text/plain; charset=utf-8', 'Content-Transfer-Encoding: base64', '', Buffer.from(`${urdu}\r\n${long}\r\n`).toString('base64'), ''].join('\r\n'),
  ];
  const dir = tempDir('words');
  const paths = raws.map((raw, n) => {
    const p = path.join(dir, `${n}.eml`);
    fs.writeFileSync(p, raw);
    return p;
  });
  await f.service.importEml({ accountId: a.id, folder: 'Inbox', paths });
  const out = tempDir('words-out');
  for (const row of f.rows(a.id, 'Inbox')) {
    const kept = f.service.message({ accountId: a.id, folder: 'Inbox', id: row.id });
    const file = path.join(out, `${row.id}.eml`);
    await f.service.saveEml({ accountId: a.id, folder: 'Inbox', ids: [row.id], path: file });
    const bytes = fs.readFileSync(file);
    const back = parseMessage(bytes);
    for (const k of ['subject', 'messageId', 'text']) assert.equal(String(back[k]).trim(), String(kept[k]).trim(), `${k} of ${kept.subject}`);
    assert.equal(new Date(back.date).toISOString(), new Date(kept.date).toISOString());
    assert.deepEqual(back.from.map((x) => x.address), kept.from.map((x) => x.address));
    assert.deepEqual(back.attachments.map((x) => [x.filename, x.size]), kept.attachments.map((x) => [x.filename, x.size]));
    // Readable as it is: ASCII words as they were, others as quoted-printable, no line over 998.
    const text = bytes.toString('latin1');
    if (/File message 7/.test(kept.subject)) assert.match(text, /^Body of file 7\.$/m, 'the words in the file as they were written');
    else {
      assert.match(text, /^Content-Transfer-Encoding: quoted-printable$/m);
      assert.match(text, /Caf=C3=A9 at ten\./);
      assert.match(text, /A long line\. A long line\./);
    }
    assert.ok(text.split('\r\n').every((l) => l.length <= 998), 'no line too long for mail');
  }
});

test('the words of a rebuilt message: 7-bit when they can be, else quoted-printable, lines kept to 76', () => {
  assert.deepEqual(readableBody('Hello\nthere'), ['Content-Transfer-Encoding: 7bit', '', 'Hello\r\nthere']);
  const [cte, blank, body] = readableBody(`naïve = ${'x'.repeat(200)} \ttab`);
  assert.equal(cte, 'Content-Transfer-Encoding: quoted-printable');
  assert.equal(blank, '');
  assert.ok(body.split('\r\n').every((l) => l.length <= 76));
  assert.match(body, /^na=C3=AFve =3D x/);
  assert.equal(parseMessage(Buffer.from(['Content-Type: text/plain; charset=utf-8', cte, '', body, ''].join('\r\n'))).text.trim(), `naïve = ${'x'.repeat(200)} \ttab`);
  // A space or tab that ends a line is written as =20 or =09, as a server may drop it otherwise.
  assert.match(readableBody('é \nnext')[2], /^=C3=A9=20\r\nnext$/);
});

test('a file opened from Explorer is kept apart, read only, with no account of its own, and gone at the next start', async () => {
  const f = fixture();
  const a = f.add('old@example.com');
  const dir = tempDir('open');
  const p = path.join(dir, 'invoice.eml');
  fs.writeFileSync(p, eml(5, ['Status: O']));
  const opened = f.service.openEml({ path: p });
  assert.ok(isOpened(opened.accountId));
  assert.equal(opened.name, 'invoice.eml');
  const m = f.service.message(opened);
  assert.equal(m.subject, 'File message 5');
  assert.equal(m.readOnly, true, 'read only');
  assert.equal(m.opened?.path, p, 'and says which file');
  assert.ok(!f.service.accounts().some((x) => x.id === opened.accountId), 'no account appears for it');
  await assert.rejects(async () => f.service.importEml({ accountId: opened.accountId, folder: 'Inbox', paths: [p] }), /read only/);
  assert.throws(() => f.service.delete({ accountId: opened.accountId, folder: opened.folder, ids: [opened.id] }), /read only/);
  // Add to a folder: the file itself brought in, unread as it was.
  const added = await f.service.importEml({ accountId: a.id, folder: 'Inbox', paths: [opened.path] });
  assert.equal(added.added, 1);
  assert.equal(f.rows(a.id, 'Inbox')[0].unread, true);
  assert.throws(() => f.service.openEml({ path: path.join(dir, 'gone.eml') }), /not there/);
  fs.writeFileSync(path.join(dir, 'not.eml'), 'just words');
  assert.throws(() => f.service.openEml({ path: path.join(dir, 'not.eml') }), /not a mail message/);
  // The next start finds it there and takes it away.
  assert.ok(fs.existsSync(path.join(f.userData, 'mail', opened.accountId)));
  createMailService({ stores: f.stores, holdBlob: () => ({ url: '' }), broadcast: () => {}, userData: f.userData, smtp: { send: async () => ({}) }, autoFetch: false });
  assert.ok(!fs.existsSync(path.join(f.userData, 'mail', opened.accountId)));
});

test('at the server: an imported file is APPENDed with its marks', async () => {
  const server = await imapServer();
  const f = fixture();
  try {
    const a = f.service.addAccount({ account: server.account(), password: 'check' });
    await f.service.sync({ accountId: a.id });
    const dir = tempDir('srv');
    const p = path.join(dir, 'x.eml');
    fs.writeFileSync(p, eml(6, ['X-Status: F']));
    const r = await f.service.importEml({ accountId: a.id, folder: 'INBOX', paths: [p] });
    assert.equal(r.added, 1);
    await f.service.flushChanges({ accountId: a.id });
    const there = server.messages('INBOX');
    assert.deepEqual(there.map((m) => m.subject), ['File message 6']);
    assert.deepEqual(there[0].flags.sort(), ['\\Flagged', '\\Seen']);
    assert.match(there[0].raw, /Body of file 6\./);
  } finally {
    await server.close();
  }
});
