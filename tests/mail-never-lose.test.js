// What the review found could lose mail, each as the failure it would be.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createMailService } from '../apps/desktop/main/mail.js';
import { MailStore } from '../packages/mailbox/src/store.js';
import { parseMessage } from '../packages/mailbox/src/mime.js';
import { scanMbox, isSeparator, openMbox as openMboxFile } from '../packages/mailbox/src/mbox.js';
import { freeNameIn } from '../apps/desktop/renderer/apps/mail/parts.js';

function fakeStores() {
  const settings = new Map();
  const secrets = new Map();
  return {
    settings: {
      get: (k, fallback) => (settings.has(k) ? structuredClone(settings.get(k)) : fallback),
      set: (k, v) => settings.set(k, structuredClone(v)),
      delete: (k) => settings.delete(k),
      all: () => Object.fromEntries(settings),
    },
    secrets: { available: () => true, get: (k) => secrets.get(k) ?? null, set: (k, v) => secrets.set(k, v), delete: (k) => secrets.delete(k), keys: () => [...secrets.keys()] },
  };
}

function fixture({ smtp } = {}) {
  const userData = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-mail-lose-'));
  const stores = fakeStores();
  const clock = { at: Date.parse('2026-03-02T09:00:00Z') };
  const sent = [];
  const ship = smtp || { async send({ draft }) { sent.push(draft); return { messageId: `m-${sent.length}` }; } };
  const told = [];
  const make = (over = ship) => createMailService({ stores, holdBlob: () => ({ url: '' }), broadcast: (name, data) => told.push({ name, data }), userData, now: () => clock.at, smtp: over });
  const service = make();
  const account = service.addAccount({ account: { email: 'me@example.com', name: 'Me', imap: { host: 'h', port: 993 }, smtp: { host: 'h', port: 465 } }, password: 'x' });
  return { service, make, account, stores, sent, clock, userData, told };
}
const dueNow = (f) => new Date(f.clock.at).toISOString();
const later = (f) => new Date(f.clock.at + 3600_000).toISOString();

test('Send now that fails leaves the message in the Outbox', async () => {
  const f = fixture({ smtp: { async send() { throw new Error('the server said no'); } } });
  const item = f.service.queue({ accountId: f.account.id, draft: { to: 'a@b.com', subject: 'Keep me' }, at: later(f) });
  const done = await f.service.sendNow({ id: item.id });
  assert.equal(done.sent, false);
  assert.match(done.error, /said no/);
  assert.equal(f.service.outbox().length, 1);
  assert.equal(f.service.outbox()[0].draft.subject, 'Keep me');
});

test('Send now that works takes it off, once', async () => {
  const f = fixture();
  const item = f.service.queue({ accountId: f.account.id, draft: { to: 'a@b.com', subject: 'Go' }, at: later(f) });
  assert.deepEqual(await f.service.sendNow({ id: item.id }), { sent: true });
  assert.equal(f.sent.length, 1);
  assert.equal(f.service.outbox().length, 0);
});

test('Send now of something no longer waiting says so', async () => {
  const f = fixture();
  assert.match((await f.service.sendNow({ id: 'gone' })).error, /no longer waiting/);
});

test('a stop in the middle of sending loses nothing: the rest are still due and the one going is held', async () => {
  let release;
  const stuck = new Promise((r) => { release = r; });
  const f = fixture({ smtp: { async send() { await stuck; return { messageId: 'x' }; } } });
  for (const subject of ['One', 'Two', 'Three']) f.service.queue({ accountId: f.account.id, draft: { to: 'a@b.com', subject }, at: dueNow(f) });
  // Queue pumps by itself; give it the moment to reach the first send.
  for (let i = 0; i < 50 && !f.service.outbox().some((o) => o.sending); i++) await new Promise((r) => setTimeout(r, 5));
  const during = f.service.outbox();
  assert.equal(during.length, 3, 'all three still written down while the first goes');
  assert.deepEqual(during.filter((o) => o.sending).map((o) => o.draft.subject), ['One']);

  // The program stops here. Another run starts from what was written.
  const sent = [];
  const after = f.make({ async send({ draft }) { sent.push(draft.subject); return { messageId: 'y' }; } });
  await after.pumpOutbox();
  assert.deepEqual(sent, ['Two', 'Three'], 'those not reached go; the one in doubt does not go twice');
  const left = after.outbox();
  assert.equal(left.length, 1);
  assert.equal(left[0].draft.subject, 'One');
  assert.equal(left[0].held, true, 'held for the person to decide');
  await after.pumpOutbox();
  assert.deepEqual(sent, ['Two', 'Three'], 'and not sent by the next pass');

  // The person looked, and sends it.
  assert.deepEqual(await after.sendNow({ id: left[0].id }), { sent: true });
  assert.deepEqual(sent, ['Two', 'Three', 'One']);
  release();
});

test('a message taken back while the queue is working is not sent', async () => {
  let release;
  const gate = new Promise((r) => { release = r; });
  const going = [];
  const f = fixture({ smtp: { async send({ draft }) { going.push(draft.subject); if (draft.subject === 'A') await gate; return { messageId: `m-${draft.subject}` }; } } });
  f.service.queue({ accountId: f.account.id, draft: { to: 'a@b.com', subject: 'A' }, at: later(f) });
  const b = f.service.queue({ accountId: f.account.id, draft: { to: 'a@b.com', subject: 'B' }, at: new Date(f.clock.at + 3660_000).toISOString() });
  f.clock.at += 7200_000;
  const pass = f.service.pumpOutbox();
  for (let i = 0; i < 100 && !going.includes('A'); i++) await new Promise((r) => setTimeout(r, 5));
  assert.deepEqual(going, ['A'], 'the pass is sending A, with B still to come');
  // Taken back while A is on its way, after the pass has already planned B.
  assert.equal(f.service.unsend({ id: b.id }).draft.subject, 'B');
  release();
  await pass;
  assert.deepEqual(going, ['A']);
  assert.equal(f.service.outbox().length, 0);
});

test('a scheduled message keeps its attachments when the temporary folder is cleared', async () => {
  const f = fixture();
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-att-'));
  const file = path.join(dir, 'scratch.pdf');
  fs.writeFileSync(file, 'the attached bytes');
  f.service.queue({ accountId: f.account.id, draft: { to: 'a@b.com', subject: 'With file', attachments: [{ filename: 'report.pdf', path: file }] }, at: later(f) });
  fs.rmSync(dir, { recursive: true });
  assert.equal(f.service.outbox()[0].draft.attachments[0].path, undefined, 'no temporary path is relied on');

  await f.service.pumpOutbox();
  const restarted = f.make();
  f.clock.at += 7200_000;
  await restarted.pumpOutbox();
  assert.equal(f.sent.length, 1);
  assert.equal(f.sent[0].attachments[0].filename, 'report.pdf');
  assert.equal(Buffer.from(f.sent[0].attachments[0].content).toString(), 'the attached bytes');
});

test('a scheduled attachment that cannot be read stops the scheduling, out loud', () => {
  const f = fixture();
  assert.throws(() => f.service.queue({ accountId: f.account.id, draft: { to: 'a@b.com', attachments: [{ filename: 'x', path: path.join(os.tmpdir(), 'not-there-at-all.pdf') }] }, at: later(f) }), /ENOENT/);
  assert.equal(f.service.outbox().length, 0);
});

test('taken back, a scheduled message has its attachments as files again', () => {
  const f = fixture();
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-att-')), 'a.txt');
  fs.writeFileSync(file, 'words');
  const item = f.service.queue({ accountId: f.account.id, draft: { to: 'a@b.com', attachments: [{ filename: 'a.txt', path: file }] }, at: later(f) });
  const kept = f.service.outbox()[0].draft.attachments[0].kept;
  assert.ok(fs.existsSync(kept));
  const back = f.service.unsend({ id: item.id });
  assert.equal(fs.readFileSync(back.draft.attachments[0].path, 'utf8'), 'words');
  assert.equal(fs.existsSync(kept), false, 'and the copy is let go');
});

// ---- the store

function store() {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-store-lose-'));
  return { root, s: new MailStore(root) };
}
const msg = (n, extra = {}) => ({ id: `m${n}`, subject: `Message ${n}`, from: { address: 'a@b.c' }, to: [], date: '2026-01-01T00:00:00Z', text: `body ${n}`, attachments: [], ...extra });
const subjects = (s, folder) => s.list('acc', folder, { limit: 100 }).rows.map((r) => r.subject).sort();

test('folders whose names come to one directory keep their own mail', () => {
  const { s } = store();
  ['Old Mail', 'Old-Mail', 'Old/Mail', 'INBOX', 'Inbox'].forEach((folder, n) => s.put('acc', folder, msg(n)));
  assert.deepEqual(subjects(s, 'Old Mail'), ['Message 0']);
  assert.deepEqual(subjects(s, 'Old-Mail'), ['Message 1']);
  assert.deepEqual(subjects(s, 'Old/Mail'), ['Message 2']);
  assert.deepEqual(subjects(s, 'INBOX'), ['Message 3']);
  assert.deepEqual(subjects(s, 'Inbox'), ['Message 4']);
});

test('removing one of those folders leaves the others', () => {
  const { s } = store();
  s.put('acc', 'Old Mail', msg(1));
  s.put('acc', 'Old/Mail', msg(2));
  s.setFolders('acc', [{ path: 'Old Mail', name: 'Old Mail' }, { path: 'Old/Mail', name: 'Mail' }]);
  s.removeFolder('acc', 'Old Mail');
  assert.deepEqual(subjects(s, 'Old/Mail'), ['Message 2']);
});

test('a folder called outgoing is not the queue of sent mail, and a folder keeps its directory across a restart', () => {
  const { s, root } = store();
  s.put('acc', 'outgoing', msg(1));
  s.put('acc', 'Old Mail', msg(2));
  s.put('acc', 'Old/Mail', msg(3));
  const again = new MailStore(root);
  assert.deepEqual(subjects(again, 'outgoing'), ['Message 1']);
  assert.deepEqual(subjects(again, 'Old/Mail'), ['Message 3']);
  assert.equal(fs.existsSync(path.join(root, 'acc', 'outgoing', 'index.jsonl')), false);
});

test('a directory from before is still found by its folder, and a folder named like it later gets its own', () => {
  const { s, root } = store();
  // Made before the list of directories was kept: the folder's directory has a changed name and nothing says whose it is.
  s.put('acc', 'Old/Mail', msg(1));
  const [dir] = fs.readdirSync(path.join(root, 'acc'), { withFileTypes: true }).filter((e) => e.isDirectory()).map((e) => e.name);
  assert.notEqual(dir, 'Old/Mail');
  fs.rmSync(path.join(root, 'acc', 'folder-dirs.json'), { force: true });
  const again = new MailStore(root);
  assert.deepEqual(subjects(again, 'Old/Mail'), ['Message 1']);
  again.put('acc', dir, msg(2));
  assert.deepEqual(subjects(again, 'Old/Mail'), ['Message 1']);
  assert.deepEqual(subjects(new MailStore(root), dir), ['Message 2']);
});

test('a store that holds an older list reads the one written since, and writes nothing over it', () => {
  const { s, root } = store();
  s.put('acc', 'Saved', msg(1));
  const searching = new MailStore(root);
  assert.deepEqual(subjects(searching, 'Saved'), ['Message 1']);
  // Since the search store read the list: a folder takes the plain directory name, so the next folder that would come to it is given another.
  const plain = 'Old_Mail';
  s.put('acc', plain, msg(2));
  assert.ok(fs.existsSync(path.join(root, 'acc', plain)), 'the plain name is the one Old/Mail would have had');
  s.put('acc', 'Old/Mail', msg(3));
  assert.deepEqual(subjects(searching, 'Old/Mail'), ['Message 3'], 'not the mail of the folder on the plain name');
  assert.deepEqual(subjects(searching, plain), ['Message 2']);
  const fresh = new MailStore(root);
  assert.deepEqual(subjects(fresh, 'Saved'), ['Message 1']);
  assert.deepEqual(subjects(fresh, plain), ['Message 2']);
  assert.deepEqual(subjects(fresh, 'Old/Mail'), ['Message 3']);
});

test('renaming a folder onto a name that would collide keeps both', () => {
  const { s } = store();
  s.put('acc', 'A B', msg(1));
  s.put('acc', 'C', msg(2));
  s.setFolders('acc', [{ path: 'A B', name: 'A B' }, { path: 'C', name: 'C' }]);
  s.renameFolder('acc', 'C', 'A/B');
  assert.deepEqual(subjects(s, 'A B'), ['Message 1']);
  assert.deepEqual(subjects(s, 'A/B'), ['Message 2']);
});

test('a crash while a folder index is being rewritten leaves the old one whole', (t) => {
  const { s, root } = store();
  s.put('acc', 'INBOX', msg(1));
  s.put('acc', 'INBOX', msg(2));
  const index = path.join(root, 'acc', 'INBOX', 'index.jsonl');
  const before = fs.readFileSync(index, 'utf8');
  t.mock.method(fs, 'renameSync', () => { const e = new Error('the machine stopped'); e.code = 'EIO'; throw e; });
  assert.throws(() => s.removeMany('acc', 'INBOX', ['m1']));
  t.mock.restoreAll();
  assert.equal(fs.readFileSync(index, 'utf8'), before);
  assert.deepEqual(fs.readdirSync(path.dirname(index)).filter((n) => n.endsWith('.tmp')), [], 'and no half file left about');
});

test('a message file is written whole or not at all', (t) => {
  const { s, root } = store();
  s.put('acc', 'INBOX', msg(1));
  const dir = path.join(root, 'acc', 'INBOX', 'm');
  const file = path.join(dir, fs.readdirSync(dir)[0]);
  const before = fs.readFileSync(file, 'utf8');
  t.mock.method(fs, 'renameSync', () => { const e = new Error('stopped'); e.code = 'EIO'; throw e; });
  assert.throws(() => s.put('acc', 'INBOX', msg(1, { subject: 'Changed' }), { force: true }));
  t.mock.restoreAll();
  assert.equal(fs.readFileSync(file, 'utf8'), before);
});

test('a row added after a torn last line is not lost with it', () => {
  const { s, root } = store();
  s.put('acc', 'INBOX', msg(1));
  fs.appendFileSync(path.join(root, 'acc', 'INBOX', 'index.jsonl'), '{"id":"half');
  const again = new MailStore(root);
  again.put('acc', 'INBOX', msg(2));
  assert.deepEqual(subjects(new MailStore(root), 'INBOX'), ['Message 1', 'Message 2']);
});

test('Save all: two attachments of one name, or a file already there, are both kept', async () => {
  const there = new Set(['C:/out/report.pdf', 'C:/out/report (2).pdf']);
  const exists = async (p) => there.has(p);
  assert.equal(await freeNameIn('C:/out', 'report.pdf', exists), 'C:/out/report (3).pdf');
  assert.equal(await freeNameIn('C:/out', 'new.pdf', exists), 'C:/out/new.pdf');
  assert.equal(await freeNameIn('C:/out', '..\\..\\evil.exe', exists), 'C:/out/.._.._evil.exe', 'a name cannot climb out of the folder');
  assert.equal(await freeNameIn('C:/out', 'no extension', exists), 'C:/out/no extension');
});

const WORDS = 'the words of a private letter';

test('a queued message keeps its words out of the settings, and sends them', async () => {
  const f = fixture();
  const item = f.service.queue({ accountId: f.account.id, draft: { to: 'a@b.com', subject: 'Later', text: WORDS, html: `<p>${WORDS}</p>` }, at: later(f) });
  assert.ok(!JSON.stringify(f.stores.settings.all()).includes(WORDS));
  assert.equal(f.service.outbox()[0].draft.subject, 'Later');
  const file = item.draft.keptWords;
  assert.ok(fs.existsSync(file));
  await f.service.sendNow({ id: item.id });
  assert.equal(f.sent[0].text, WORDS);
  assert.equal(f.sent[0].html, `<p>${WORDS}</p>`);
  assert.equal(fs.existsSync(file), false);
});

test('taken back, a queued message has its words again', () => {
  const f = fixture();
  const item = f.service.queue({ accountId: f.account.id, draft: { to: 'a@b.com', text: WORDS }, at: later(f) });
  const back = f.service.unsend({ id: item.id });
  assert.equal(back.draft.text, WORDS);
  assert.equal(back.draft.keptWords, undefined);
  assert.equal(fs.existsSync(item.draft.keptWords), false);
});

test('a message queued before the words were kept apart still sends them', async () => {
  const f = fixture();
  f.stores.settings.set('mail.outbox', [{ id: 'old', accountId: f.account.id, draft: { to: 'a@b.com', text: WORDS }, at: dueNow(f) }]);
  await f.service.sendNow({ id: 'old' });
  assert.equal(f.sent[0].text, WORDS);
});

test('a recipient the server refused is said, though the message went to the others', async () => {
  const f = fixture({ smtp: { async send() { return { messageId: 'm', rejected: ['gone@b.com', { address: 'no@b.com' }] }; } } });
  const item = f.service.queue({ accountId: f.account.id, draft: { to: 'a@b.com, gone@b.com, no@b.com', text: 'x' }, at: later(f) });
  await f.service.sendNow({ id: item.id });
  assert.deepEqual(f.told.find((e) => e.name === 'mail:sent').data.refused, ['gone@b.com', 'no@b.com']);
});

test('a folder exported keeps its attachments, its copies and its page', async () => {
  const f = fixture();
  new MailStore(path.join(f.userData, 'mail')).putMany(f.account.id, 'INBOX', [{
    messageId: 'kept@x', subject: 'The report', date: '2026-03-01T09:00:00Z',
    from: [{ address: 'ana@work.example', name: 'Ana' }], to: [{ address: 'me@example.com' }], cc: [{ address: 'bo@work.example' }],
    text: 'See attached.', html: '<p>See <b>attached</b>.</p>',
    attachments: [{ filename: 'report.pdf', type: 'application/pdf', bytes: Buffer.from('%PDF-1.4 the report') }],
  }]);
  const out = path.join(f.userData, 'out');
  assert.equal((await f.service.export({ accountId: f.account.id, folder: 'INBOX', path: out, format: 'eml' })).messages, 1);
  const eml = fs.readFileSync(path.join(out, fs.readdirSync(out)[0]));
  const back = parseMessage(eml);
  assert.equal(back.subject, 'The report');
  assert.equal(back.messageId, 'kept@x');
  assert.equal(back.cc[0].address, 'bo@work.example');
  assert.match(back.html, /<b>attached<\/b>/);
  assert.equal(back.attachments.length, 1);
  assert.equal(back.attachments[0].filename, 'report.pdf');
  assert.equal(Buffer.from(back.attachments[0].bytes).toString(), '%PDF-1.4 the report');

  // The same folder as one archive, the form a backup takes.
  const archive = path.join(f.userData, 'INBOX.mbox');
  assert.equal((await f.service.export({ accountId: f.account.id, folder: 'INBOX', path: archive, format: 'mbox' })).messages, 1);
  const open = openMboxFile(archive);
  try {
    assert.equal(open.length, 1);
    const inMbox = parseMessage(Buffer.from(open.raw(0)));
    assert.equal(inMbox.subject, 'The report');
    assert.equal(inMbox.messageId, 'kept@x');
    assert.equal(inMbox.cc[0].address, 'bo@work.example');
    assert.match(inMbox.html, /<b>attached<\/b>/);
    assert.equal(inMbox.attachments.length, 1);
    assert.equal(inMbox.attachments[0].filename, 'report.pdf');
    assert.equal(Buffer.from(inMbox.attachments[0].bytes).toString(), '%PDF-1.4 the report');
  } finally {
    open.close?.();
  }
  assert.equal(fs.existsSync(`${archive}.part`), false, 'nothing half written left beside it');
});

test('a line of a letter that starts with From does not start a message', () => {
  const text = [
    'From ana@work.example Mon Mar  2 09:00:00 2026',
    'Subject: One',
    '',
    'From here on we use the new office.',
    'From Monday 9:00 to 17:00 in 2026 we are open.',
    '',
    'From - Tue, 03 Mar 2026 10:00:00 GMT',
    'Subject: Two',
    '',
    'Second.',
    '',
  ].join('\n');
  assert.equal(scanMbox(Buffer.from(text)).ranges.length, 2);
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-mbox-from-'));
  const file = path.join(dir, 'a.mbox');
  fs.writeFileSync(file, text);
  const open = openMboxFile(file);
  try {
    assert.equal(open.length, 2);
  } finally {
    open.close?.();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('separators that are not dated are still separators when a header follows', () => {
  const sep = (s) => isSeparator(Buffer.from(s));
  assert.equal(sep('From a@b.example\nSubject: hello\n'), true);
  assert.equal(sep('From - \nReceived: by x\n'), true);
  assert.equal(sep('From the desk of\nNote: read this\n'), false);
  assert.equal(sep('From a@b.example\nand then some words\n'), false);
  assert.equal(sep('From root Sun Jan  5 11:02:03 +0000 2025 remote from host\n'), true);
});
