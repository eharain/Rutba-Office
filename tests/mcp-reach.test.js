// The AI tools beyond files (owner, 2026-10-26, D-2026-10-26-04, -05, -06):
// each mail account, archive, calendar, address book, Data Files database
// and folder carries its own setting, Off until chosen; each step allows
// what it says and no more; a refusal names the thing and where its setting
// is; what leaves the computer waits for the person's press, and Keep as
// draft and Cancel do what they say; a protected profile is reached only
// while it is open; the log names the tool, the thing and the step, never
// what was read.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createReachTools } from '../apps/desktop/main/mcp/reach-tools.js';
import { createConfirmations } from '../apps/desktop/main/mcp/confirm.js';
import { createChannelServer, createChannelClient, CHANNEL_FILE } from '../apps/desktop/main/mcp/channel.js';
import { createTools, READERS } from '../apps/desktop/main/mcp/tools.js';
import { MODE_KEY, FOLDERS_KEY, resolveIn, foldersOf, readLog } from '../apps/desktop/main/mcp/access.js';
import { stepOf, stepsFor, folderEntries, withFolderStep, DATABASES_KEY, refusalFor } from '../apps/desktop/main/mcp/reach.js';
import { createCalendarService } from '../apps/desktop/main/calendar.js';
import { createContactsService } from '../apps/desktop/main/contacts.js';
import { createBaseService } from '../apps/desktop/main/base.js';
import { Base, buildFromSpec } from '../packages/database/src/index.js';

const scratch = () => fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-reach-'));
const textOf = (r) => (r?.content || []).map((c) => c.text).join('\n');

/** Settings as the store keeps them, in memory. */
function memorySettings(initial = {}) {
  const s = { ...initial };
  return { get: (k, f) => (k in s ? s[k] : f), set: (k, v) => { s[k] = v; }, all: () => ({ ...s }) };
}

/** Mail as the tools see it: two accounts, one brought in from a file, and an archive; what it sends and keeps is noted. */
function fakeMail(settings) {
  const sent = [];
  const drafts = [];
  const message = { id: 'm1', subject: 'Invoice 42', from: [{ name: 'Supplier', address: 'bills@supplier.example' }], to: [{ address: 'me@home.example' }], date: '2026-10-20T09:00:00.000Z', text: 'Please find the invoice attached.', attachments: [{ filename: 'invoice.txt', type: 'text/plain', size: 12 }] };
  const mail = {
    accounts: () => settings.get('mail.accounts', []),
    archives: () => settings.get('mail.archives', []).map((a) => ({ ...a, archive: true })),
    updateAccount: ({ id, patch }) => settings.set('mail.accounts', mail.accounts().map((a) => (a.id === id ? { ...a, ...patch } : a))),
    searchAll: async ({ sources }) => ({ rows: sources.map((s) => ({ accountId: s, folder: 'Inbox', id: 'm1', subject: message.subject, from: message.from, date: message.date })) }),
    folders: ({ accountId }) => [{ path: 'Inbox', total: 1, unread: 0, accountId }],
    messages: () => ({ total: 1, rows: [{ ...message, unread: false }] }),
    message: () => message,
    attachment: () => ({ id: 'b1', name: 'invoice.txt', type: 'text/plain' }),
    attachmentText: () => ({ name: 'invoice.txt', text: 'Total: 42.00' }),
    saveDraft: async (args) => { drafts.push(args); return { id: `d${drafts.length}` }; },
    send: async (args) => { sent.push(args); return { messageId: `s${sent.length}` }; },
  };
  return { mail, sent, drafts };
}

function world({ steps = {}, waitMs = 2000 } = {}) {
  const dir = scratch();
  const work = path.join(dir, 'work');
  const readOnly = path.join(dir, 'read-only');
  fs.mkdirSync(work);
  fs.mkdirSync(readOnly);
  fs.writeFileSync(path.join(work, 'notes.txt'), 'the notes');
  const settings = memorySettings({
    [MODE_KEY]: 'always',
    [FOLDERS_KEY]: [work, { path: readOnly, step: 'read' }],
    'mail.accounts': [
      { id: 'a1', email: 'me@home.example', name: 'Me', ai: steps.mail },
      { id: 'a2', email: 'old@file.example', name: 'Old mailbox', local: true, ai: steps.local },
    ],
    'mail.archives': [{ id: 'archive:x1', name: 'Old archive', path: path.join(dir, 'old.pst'), ai: steps.archive }],
  });
  const stores = { dir, settings, vault: null };
  const calendar = createCalendarService({ stores, mail: { accounts: () => settings.get('mail.accounts', []) } });
  const contacts = createContactsService({ stores });
  // A calendar and an address book a server keeps in step, as dav-sync.js makes them.
  calendar._dav.state().calendars.push({ id: 'server-cal', name: 'Work', colour: '#2b5fd9', visible: true, events: [], remote: { account: 'dav1', href: '/cal/work/' }, dirty: [], deleted: [] });
  contacts._dav.state().books = [{ id: 'server-book', account: 'dav1', href: '/book/', name: 'Office cards' }];
  const { mail, sent, drafts } = fakeMail(settings);
  const base = createBaseService();
  const confirm = createConfirmations({ waitMs });
  const tools = createReachTools({
    settings, mail, calendar, contacts, base, confirm,
    files: { readable: (given) => resolveIn(foldersOf(settings.all()), given) },
    readers: READERS,
    attachmentBytes: async () => ({ bytes: Buffer.from('Total: 42.00'), name: 'invoice.txt', type: 'text/plain' }),
  });
  const call = async (name, args = {}) => (await tools.call(name, args)).result;
  return { dir, work, readOnly, settings, calendar, contacts, mail, sent, drafts, base, confirm, tools, call };
}

test('the steps: four everywhere, an archive read only, a folder and a mailbox from a file never past prepare, anything else Off', () => {
  assert.deepEqual(stepsFor('mail'), ['off', 'read', 'prepare', 'act']);
  assert.deepEqual(stepsFor('archive'), ['off', 'read']);
  assert.deepEqual(stepsFor('folder'), ['off', 'read', 'prepare']);
  assert.deepEqual(stepsFor('mail', { local: true }), ['off', 'read', 'prepare']);
  assert.equal(stepOf('mail', undefined), 'off', 'Off until chosen');
  assert.equal(stepOf('archive', 'act'), 'off');
  assert.equal(stepOf('calendar', 'nonsense'), 'off');
  // A folder given before 2.19 keeps what it allowed: read and new files.
  assert.deepEqual(folderEntries(['C:\\Given', { path: 'C:\\Read', step: 'read' }]), [{ path: 'C:\\Given', step: 'prepare' }, { path: 'C:\\Read', step: 'read' }]);
  assert.deepEqual(withFolderStep(['C:\\Given'], 'C:\\given', 'read'), [{ path: 'C:\\given', step: 'read' }]);
  assert.deepEqual(withFolderStep([{ path: 'C:\\Given', step: 'read' }], 'C:\\Given', 'prepare'), ['C:\\Given'], 'kept as before 2.19 when it means the same');
  assert.match(refusalFor('mail', 'me@home.example', 'read', 'act'), /me@home\.example.*Read.*Read, prepare and act.*in Mail under Accounts/);
});

test('Off until chosen: nothing of mail, calendars, contacts or databases is reached, and each refusal names the thing and where its setting is', async () => {
  const w = world();
  assert.match(textOf(await w.call('list_allowed')), /^kind\tname\tallows\nfolder\t[^\n]*\nfolder\t[^\n]*\n\n7 more are Off/);
  assert.match(textOf(await w.call('search_mail', { query: 'invoice' })), /No mail account or archive is set to Read yet.*Mail under Accounts/);
  const read = await w.call('read_message', { account: 'me@home.example', folder: 'Inbox', id: 'm1' });
  assert.equal(read.isError, true);
  assert.match(textOf(read), /the mail account me@home\.example: it is set to Off.*in Mail under Accounts, or in AI Tools/);
  assert.doesNotMatch(textOf(read), /invoice attached/, 'nothing of the message');
  assert.match(textOf(await w.call('list_events')), /No calendar is set to Read yet.*Calendar under Accounts/);
  assert.match(textOf(await w.call('search_contacts', { query: '' })), /No address book is set to Read yet.*Contacts under Accounts/);
  assert.match(textOf(await w.call('create_event', { calendar: 'My calendar', summary: 'x', start: '2026-11-03T10:00:00Z' })), /the calendar My calendar: it is set to Off/);
  assert.match(textOf(await w.call('read_records', { database: 'Shop', table: 'Products' })), /no database called Shop/);
  assert.equal(w.sent.length + w.drafts.length, 0);
});

test('mail, step by step: Read reads and searches, Read and prepare makes drafts, only Read, prepare and act may send; an archive and a mailbox from a file stay below', async () => {
  const w = world({ steps: { mail: 'read', archive: 'read', local: 'act' } });
  const found = textOf(await w.call('search_mail', { query: 'invoice' }));
  assert.match(found, /me@home\.example\tInbox\tm1/);
  assert.match(found, /Old archive\tInbox\tm1/);
  assert.doesNotMatch(found, /old@file\.example/, 'a mailbox from a file set to act is Off: act is not one of its steps');
  assert.match(textOf(await w.call('read_message', { account: 'me@home.example', folder: 'Inbox', id: 'm1' })), /Subject: Invoice 42[\s\S]*Please find the invoice attached[\s\S]*0\tinvoice\.txt/);
  assert.match(textOf(await w.call('read_attachment', { account: 'me@home.example', folder: 'Inbox', id: 'm1', index: 0 })), /Total: 42\.00/);
  assert.match(textOf(await w.call('create_draft', { account: 'me@home.example', to: 'a@b.example', text: 'x' })), /set to Read, and this needs Read and prepare/);
  assert.match(textOf(await w.call('create_draft', { account: 'Old archive', to: 'a@b.example' })), /only read the archive Old archive/);
  // A refusal is logged as a call that worked is: the thing and the step it needed.
  assert.deepEqual((await w.tools.call('create_draft', { account: 'me@home.example', to: 'a@b.example' })).log, { kind: 'mail', name: 'me@home.example', step: 'prepare' });
  assert.deepEqual((await w.tools.call('read_message', { account: 'me@home.example', folder: 'Inbox', id: 'm1' })).log, { kind: 'mail', name: 'me@home.example', step: 'read' });

  w.mail.updateAccount({ id: 'a1', patch: { ai: 'prepare' } });
  assert.match(textOf(await w.call('create_draft', { account: 'me@home.example', to: 'a@b.example', subject: 'Hello', text: 'Hi', attachments: ['notes.txt'] })), /Kept a draft.*Nothing was sent/);
  assert.equal(w.drafts.length, 1);
  assert.equal(w.drafts[0].draft.attachments[0].path, path.join(w.work, 'notes.txt'));
  const send = await w.call('send_message', { account: 'me@home.example', to: 'a@b.example', subject: 'Hello', text: 'Hi' });
  assert.match(textOf(send), /set to Read and prepare, and this needs Read, prepare and act/);
  assert.equal(w.confirm.pending().length, 0, 'nothing waits for a call refused');
  assert.equal(w.sent.length, 0);
});

test('an attachment goes only from a folder the tools may read', async () => {
  const w = world({ steps: { mail: 'prepare' } });
  const outside = path.join(os.tmpdir(), 'elsewhere.txt');
  const r = await w.call('create_draft', { account: 'me@home.example', to: 'a@b.example', attachments: [outside] });
  assert.equal(r.isError, true);
  assert.match(textOf(r), /may only use these folders/);
  assert.equal(w.drafts.length, 0);
});

/** A send through the tools, answered by the person once it waits. */
async function sendAndAnswer(w, choice) {
  const before = w.sent.length;
  const asked = w.call('send_message', { account: 'me@home.example', to: 'friend@there.example', subject: 'Plans', text: 'See you at nine.' });
  for (let i = 0; i < 50 && !w.confirm.pending().length; i++) await new Promise((r) => setTimeout(r, 10));
  const [waiting] = w.confirm.pending();
  assert.ok(waiting, 'shown in Rutba Office');
  assert.equal(waiting.kind, 'mail');
  assert.equal(waiting.fields.to, 'friend@there.example');
  assert.equal(waiting.fields.subject, 'Plans');
  assert.equal(waiting.draft, true, 'with Keep as draft');
  assert.equal(w.sent.length, before, 'nothing sent before the press');
  await w.confirm.answer(waiting.id, choice);
  return { result: await asked, id: waiting.id };
}

test('send with confirmation: it goes only on the press of Send; Keep as draft makes a draft; Cancel does nothing', async () => {
  const w = world({ steps: { mail: 'act' } });
  const sent = await sendAndAnswer(w, 'act');
  assert.match(textOf(sent.result), /Sent from me@home\.example to friend@there\.example/);
  assert.equal(w.sent.length, 1);
  assert.equal(w.sent[0].draft.to, 'friend@there.example');
  assert.equal(w.drafts.length, 0);

  const kept = await sendAndAnswer(w, 'draft');
  assert.match(textOf(kept.result), /Kept as a draft.*Nothing was sent/);
  assert.equal(w.sent.length, 1, 'not sent');
  assert.equal(w.drafts.length, 1);

  const cancelled = await sendAndAnswer(w, 'cancel');
  assert.match(textOf(cancelled.result), /Cancelled.*nothing was done/);
  assert.equal(w.sent.length, 1);
  assert.equal(w.drafts.length, 1);
  await assert.rejects(w.confirm.answer(cancelled.id, 'act'), /no longer waiting/, 'a request answered cannot be answered again');
});

test('a request the person has not answered yet: the call says it waits, with an id, and the answer can be asked for later', async () => {
  const w = world({ steps: { mail: 'act' }, waitMs: 30 });
  const r = await w.call('send_message', { account: 'me@home.example', to: 'friend@there.example', subject: 'Later', text: 'x' });
  const id = /id ([0-9a-f-]{36})/.exec(textOf(r))?.[1];
  assert.ok(id, textOf(r));
  assert.match(textOf(await w.call('confirmation_answer', { id })), /Still waiting/);
  assert.equal(w.sent.length, 0);
  await w.confirm.answer(id, 'act');
  assert.match(textOf(await w.call('confirmation_answer', { id })), /Sent from me@home\.example/);
  assert.equal(w.sent.length, 1);
});

test('calendars: events read, new ones made on this computer at prepare, one for a server calendar confirmed, and invitations sent only on the press', async () => {
  const w = world({ steps: { mail: 'act' } });
  const [mine] = w.calendar.calendars();
  w.calendar.saveCalendar({ calendar: { id: mine.id, ai: 'prepare' } });
  const made = textOf(await w.call('create_event', { calendar: mine.name, summary: 'Dentist', start: '2026-11-03T10:00:00Z', end: '2026-11-03T10:30:00Z', attendees: ['dr@clinic.example'] }));
  assert.match(made, /Made the event Dentist in My calendar, on this computer.*Nobody was told/);
  const listed = textOf(await w.call('list_events', { from: '2026-11-01', to: '2026-11-30' }));
  assert.match(listed, /Dentist/);
  const id = /\(id ([^)]+)\)/.exec(made)[1];
  // Sending its invitation tells another person: Read, prepare and act on the calendar.
  assert.match(textOf(await w.call('send_invitation', { id })), /the calendar My calendar: it is set to Read and prepare, and this needs Read, prepare and act/);
  w.calendar.saveCalendar({ calendar: { id: mine.id, ai: 'act' } });
  const asked = w.call('send_invitation', { id });
  for (let i = 0; i < 50 && !w.confirm.pending().length; i++) await new Promise((r) => setTimeout(r, 10));
  const [inv] = w.confirm.pending();
  assert.equal(inv.kind, 'invitation');
  assert.equal(inv.fields.to, 'dr@clinic.example');
  assert.equal(w.sent.length, 0);
  await w.confirm.answer(inv.id, 'act');
  assert.match(textOf(await asked), /Sent the invitation to dr@clinic\.example/);
  assert.equal(w.sent.length, 1);
  assert.equal(path.extname(w.sent[0].draft.attachments[0].path), '.ics');

  // A server's calendar: a new event goes to the server, so it is confirmed; Keep as draft keeps it here.
  w.calendar.saveCalendar({ calendar: { id: 'server-cal', ai: 'prepare' } });
  assert.match(textOf(await w.call('create_event', { calendar: 'Work', summary: 'Review', start: '2026-11-04T10:00:00Z' })), /the calendar Work: it is set to Read and prepare, and this needs Read, prepare and act/);
  w.calendar.saveCalendar({ calendar: { id: 'server-cal', ai: 'act' } });
  const ev = w.call('create_event', { calendar: 'Work', summary: 'Review', start: '2026-11-04T10:00:00Z' });
  for (let i = 0; i < 50 && !w.confirm.pending().length; i++) await new Promise((r) => setTimeout(r, 10));
  await w.confirm.answer(w.confirm.pending()[0].id, 'draft');
  assert.match(textOf(await ev), /Kept the event Review in My calendar, on this computer only/);
  assert.equal(w.calendar._dav.state().calendars.find((c) => c.id === 'server-cal').events.length, 0, 'nothing for the server');
});

test('contacts: cards read, a new one on this computer at prepare, one for a server book confirmed', async () => {
  const w = world();
  w.contacts.setBookAi({ id: 'local', ai: 'prepare' });
  assert.match(textOf(await w.call('create_contact', { name: 'Amira Khan', emails: ['amira@example.com'] })), /Added Amira Khan to the address book on this computer/);
  assert.match(textOf(await w.call('search_contacts', { query: 'amira' })), /Amira Khan\tamira@example\.com/);
  assert.match(textOf(await w.call('create_contact', { book: 'Office cards', name: 'Ben' })), /the address book Office cards: it is set to Off/);
  w.contacts.setBookAi({ id: 'server-book', ai: 'act' });
  const asked = w.call('create_contact', { book: 'Office cards', name: 'Ben Ali', emails: ['ben@example.com'] });
  for (let i = 0; i < 50 && !w.confirm.pending().length; i++) await new Promise((r) => setTimeout(r, 10));
  const [req] = w.confirm.pending();
  assert.equal(req.kind, 'contact');
  assert.equal(w.contacts.list({ query: 'ben' }).length, 0, 'not added before the press');
  await w.confirm.answer(req.id, 'act');
  assert.match(textOf(await asked), /Added Ben Ali to Office cards/);
  const ben = w.contacts.list({ query: 'ben' })[0];
  assert.equal(ben.book, 'server-book');
  assert.ok(w.contacts._dav.state().dirty.includes(ben.id), 'queued for the server');
});

test('Data Files: records read and written within the step, never the design; a protected database only as the role chosen for the tools', async () => {
  const w = world();
  const file = path.join(w.dir, 'Shop.rbase');
  const b = Base.open(file);
  buildFromSpec(b, { tables: [{ name: 'Products', fields: [{ name: 'Name', type: 'text' }, { name: 'Price', type: 'currency' }] }, { name: 'Wages', fields: [{ name: 'Who', type: 'text' }, { name: 'Amount', type: 'currency' }] }] });
  b.insert('Products', { Name: 'Bread', Price: 2.5 });
  b.insert('Wages', { Who: 'Amira', Amount: 900 });
  b.close();
  w.settings.set(DATABASES_KEY, { [file]: { step: 'read' } });
  assert.match(textOf(await w.call('read_records', { database: 'Shop', table: 'Products' })), /Bread/);
  assert.match(textOf(await w.call('list_tables', { database: file })), /Products\t1 records[\s\S]*Wages/);
  assert.match(textOf(await w.call('add_record', { database: 'Shop', table: 'Products', values: { Name: 'Milk' } })), /the database Shop: it is set to Read, and this needs Read and prepare.*Data Files/);
  w.settings.set(DATABASES_KEY, { [file]: { step: 'prepare' } });
  assert.match(textOf(await w.call('add_record', { database: 'Shop', table: 'Products', values: { Name: 'Milk', Price: 1.2 } })), /Added a record to Products/);
  assert.match(textOf(await w.call('change_record', { database: 'Shop', table: 'Products', rowid: 1, values: { Price: 2.75 } })), /Changed record 1/);
  assert.ok(!w.tools.list().some((t) => /field|design|table_create|drop|delete|remove|move/i.test(t.name)), 'no tool touches the design, deletes or moves');

  // Protected: the tools work as the role an Admin chose for them, and never as one that changes the design.
  const p = Base.open(file);
  p.protect({ name: 'Amira', password: 'owner-pass' });
  p.saveRole('Till', { tables: { default: 'read', each: { Wages: 'hidden' } } });
  p.close();
  assert.match(textOf(await w.call('read_records', { database: 'Shop', table: 'Products' })), /is protected, and no role has been chosen for the AI tools/);
  w.settings.set(DATABASES_KEY, { [file]: { step: 'prepare', role: 'Admin' } });
  assert.match(textOf(await w.call('read_records', { database: 'Shop', table: 'Products' })), /never work as a role that may change the design/);
  w.settings.set(DATABASES_KEY, { [file]: { step: 'prepare', role: 'Till' } });
  assert.match(textOf(await w.call('read_records', { database: 'Shop', table: 'Products' })), /Milk/);
  assert.match(textOf(await w.call('read_records', { database: 'Shop', table: 'Wages' })), /no table called Wages that the AI tools may see/);
  assert.match(textOf(await w.call('add_record', { database: 'Shop', table: 'Products', values: { Name: 'Eggs' } })), /may not add/);
});

test('folders, each with its own step: a folder set to Read is read and never written in, and the refusal names it', async () => {
  const dir = scratch();
  const readOnly = path.join(dir, 'read-only');
  fs.mkdirSync(readOnly);
  fs.writeFileSync(path.join(readOnly, 'kept.txt'), 'kept words');
  const tools = createTools({ settings: () => ({ [MODE_KEY]: 'always', [FOLDERS_KEY]: [{ path: readOnly, step: 'read' }] }), dataDir: dir });
  assert.match(textOf(await tools.call('read_file', { path: path.join(readOnly, 'kept.txt') })), /kept words/);
  const made = await tools.call('create_document', { path: path.join(readOnly, 'new.docx'), blocks: [{ type: 'paragraph', text: 'x' }] });
  assert.equal(made.isError, true);
  assert.match(textOf(made), /the folder .*read-only: it is set to Read, and this needs Read and prepare\. Its setting is in AI Tools, under Folders/);
  assert.ok(!fs.existsSync(path.join(readOnly, 'new.docx')));
  assert.match(textOf(await tools.call('list_folders')), /read-only\tread only/);
});

test('a protected profile: while it is not open and unlocked, nothing is reached, its folders included', async () => {
  const dir = scratch();
  const work = path.join(dir, 'work');
  fs.mkdirSync(work);
  fs.writeFileSync(path.join(work, 'a.txt'), 'private words');
  let open = false;
  const tools = createTools({
    settings: () => ({ [MODE_KEY]: 'always', [FOLDERS_KEY]: [work] }),
    dataDir: dir,
    locked: async () => (open ? null : 'This profile has a password, so the AI tools reach nothing of it.'),
    reach: { call: async () => ({ result: { content: [{ type: 'text', text: 'mail read' }] }, log: { kind: 'mail', name: 'me@home.example', step: 'read' } }) },
  });
  for (const [name, args] of [['read_file', { path: 'a.txt' }], ['list_files', {}], ['search_mail', { query: 'x' }]]) {
    const r = await tools.call(name, args);
    assert.equal(r.isError, true, name);
    assert.match(textOf(r), /has a password/);
    assert.doesNotMatch(textOf(r), /private words|mail read/);
  }
  open = true;
  assert.match(textOf(await tools.call('read_file', { path: 'a.txt' })), /private words/);
  assert.match(textOf(await tools.call('search_mail', { query: 'x' })), /mail read/);
  // The log names the tool, the thing and the step: never what was read.
  const log = readLog(dir, 10);
  assert.deepEqual({ tool: log[0].tool, kind: log[0].kind, thing: log[0].thing, step: log[0].step, ok: log[0].ok }, { tool: 'search_mail', kind: 'mail', thing: 'me@home.example', step: 'read', ok: true });
  assert.ok(!JSON.stringify(log).includes('private words') && !JSON.stringify(log).includes('mail read'));
});

test('with Rutba Office closed, the mail tools say so, and one that needs a press says nothing was done and nothing waits', async () => {
  const dir = scratch();
  const client = createChannelClient({ folder: dir, timeoutMs: 2000 });
  const tools = createTools({ settings: () => ({ [MODE_KEY]: 'always', [FOLDERS_KEY]: [dir] }), dataDir: dir, reach: { call: (name, args) => client.request('call', { name, args }) } });
  const read = await tools.call('search_mail', { query: 'x' });
  assert.match(textOf(read), /Rutba Office is not open in this profile/);
  const send = await tools.call('send_message', { account: 'me@home.example', to: 'a@b', subject: 's', text: 't' });
  assert.match(textOf(send), /needs the person to see it.*nothing was done and nothing waits/);
});

test('the channel: the running copy answers its own profile\'s tools with the key it wrote, and nobody without it', async () => {
  const dir = scratch();
  const server = createChannelServer({ folder: dir, handle: async (method, params) => ({ method, echo: params.n }) });
  await server.start();
  try {
    const client = createChannelClient({ folder: dir, timeoutMs: 5000 });
    assert.deepEqual(await client.request('call', { n: 7 }), { method: 'call', echo: 7 });
    client.close();
    const kept = JSON.parse(fs.readFileSync(path.join(dir, CHANNEL_FILE), 'utf8'));
    assert.ok(kept.key && kept.where);
    // A key that is not the one written is refused.
    fs.writeFileSync(path.join(dir, CHANNEL_FILE), JSON.stringify({ ...kept, key: 'wrong' }));
    const stranger = createChannelClient({ folder: dir, timeoutMs: 5000 });
    await assert.rejects(stranger.request('call', { n: 1 }), /key is missing or wrong/);
    stranger.close();
    fs.writeFileSync(path.join(dir, CHANNEL_FILE), JSON.stringify(kept));
  } finally {
    await server.stop();
  }
  assert.ok(!fs.existsSync(path.join(dir, CHANNEL_FILE)), 'the file goes when the channel closes');
  await assert.rejects(createChannelClient({ folder: dir }).request('hello'), (err) => err.notRunning === true);
});
