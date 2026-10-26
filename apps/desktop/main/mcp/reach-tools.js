// The AI tools for mail, calendars, address books and Data Files databases
// (owner, 2026-10-26, D-2026-10-26-04, -05, -06; plan Part E items 1 to 4).
// They run in the Rutba Office that holds those stores (channel.js): an AI
// app on this computer reaches them through the tools' own process, one on
// another machine through this copy's network side (remote.js).
//
// Each call names the thing it uses and is asked that thing's own setting
// (reach.js) before anything is read: Off refuses, naming the thing and
// where its setting is. What each step allows, per kind:
//   mail       read and search; drafts; send, each message confirmed (an
//              archive opened in place is read only)
//   calendar   events; new events (in a calendar a server keeps, confirmed,
//              as they go to the server); invitations sent and answered,
//              each confirmed
//   contacts   cards; new cards (in a book a server keeps, confirmed)
//   databases  records; new and changed records, never the design, within
//              the role chosen for the tools when the database is
//              protected; a change written to a linked engine confirmed
// Attachments go only from folders the tools may read, invitations only from
// calendars they may read, and an attachment is read only from mail they
// may read. Nothing is ever deleted or moved: there is no tool for it.
//
// Nothing here needs Electron: the services are handed in, and `confirm`
// (confirm.js) is how the person is asked.

import path from 'node:path';
import { readCalendar, dateTimeAt } from '@rutba/calendar';
import { STEPS, STEP_WORDS, LOCAL_BOOK, DATABASES_KEY, FOLDERS_KEY, stepOf, allows, refuse, folderEntries, databaseSetting } from './reach.js';

const MAX_CHARS = 60000;
const text = (t) => ({ content: [{ type: 'text', text: String(t) }] });
const trimTo = (s, max) => (s.length > max ? `${s.slice(0, max)}\n\n[… cut at ${max} characters of ${s.length}]` : s);
const addressList = (list) => (Array.isArray(list) ? list : list ? [list] : []).map((a) => (a?.name ? `${a.name} <${a.address}>` : a?.address || String(a || ''))).filter(Boolean).join(', ');
const words = (v) => String(v ?? '').trim();
const same = (a, b) => words(a).toLowerCase() === words(b).toLowerCase();
const stripHtml = (html) => String(html || '').replace(/<style[\s\S]*?<\/style>|<script[\s\S]*?<\/script>/gi, '').replace(/<br\s*\/?>|<\/p>|<\/div>|<\/li>|<\/tr>/gi, '\n').replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, '&').replace(/\n{3,}/g, '\n\n').trim();
const notThere = (what, name) => Object.assign(new Error(`Rutba Office has no ${what} called ${name}. list_allowed names those the AI tools may use.`), { refusal: true });

const STR = { type: 'string' };
const ACCOUNT = { type: 'string', description: 'The mail account\'s address, or an archive\'s name, as list_allowed gives it.' };
const MESSAGE = {
  account: ACCOUNT,
  folder: { type: 'string', description: 'The folder, as search_mail or list_messages gives it.' },
  id: { type: 'string', description: 'The message\'s id, as search_mail or list_messages gives it.' },
};
const OUTGOING = {
  account: { type: 'string', description: 'The mail account it is from (its address).' },
  to: { type: 'string', description: 'Addresses, separated by commas.' },
  cc: STR,
  bcc: STR,
  subject: STR,
  text: { type: 'string', description: 'The message, as plain text.' },
  attachments: { type: 'array', items: { type: 'string' }, description: 'Files to attach, each in a folder the AI tools may read.' },
};
const WHEN = { type: 'string', description: 'A date and time, as ISO 8601 (2026-11-03T14:00:00Z, or 2026-11-03 for a day).' };
const DATABASE = { type: 'string', description: 'The database\'s path, or its name as list_allowed gives it.' };

export const REACH_TOOL_LIST = [
  {
    name: 'list_allowed',
    description: 'Everything the person has let the Rutba Office AI tools use, beyond files: each mail account, archive, calendar, address book, Data Files database and folder that is not Off, with what each allows (Read; Read and prepare; Read, prepare and act). Anything that leaves this computer or tells another person is shown to the person in Rutba Office first and happens only when they press it.',
    inputSchema: { type: 'object', properties: {}, additionalProperties: false },
  },
  {
    name: 'search_mail',
    description: 'Search the mail of the accounts and archives the tools may read: words in the subject, sender, recipients and text. Gives each message\'s account, folder and id for read_message.',
    inputSchema: { type: 'object', properties: { query: STR, account: { ...ACCOUNT, description: `${ACCOUNT.description} Every one the tools may read when left out.` }, limit: { type: 'integer', minimum: 1, maximum: 200 } }, required: ['query'], additionalProperties: false },
  },
  {
    name: 'list_mail_folders',
    description: 'The folders of a mail account or archive the tools may read, with how many messages each holds.',
    inputSchema: { type: 'object', properties: { account: ACCOUNT }, required: ['account'], additionalProperties: false },
  },
  {
    name: 'list_messages',
    description: 'The newest messages in a folder of a mail account or archive the tools may read (the Inbox when no folder is named).',
    inputSchema: { type: 'object', properties: { account: ACCOUNT, folder: MESSAGE.folder, offset: { type: 'integer', minimum: 0 }, limit: { type: 'integer', minimum: 1, maximum: 200 } }, required: ['account'], additionalProperties: false },
  },
  {
    name: 'read_message',
    description: 'One message: who it is from and to, when, its subject, its text, and its attachments by number for read_attachment.',
    inputSchema: { type: 'object', properties: MESSAGE, required: ['account', 'folder', 'id'], additionalProperties: false },
  },
  {
    name: 'read_attachment',
    description: 'The words of a message\'s attachment: a document, workbook, presentation, PDF, or text (an invitation\'s .ics, a .vcf card, a .csv).',
    inputSchema: { type: 'object', properties: { ...MESSAGE, index: { type: 'integer', minimum: 0, description: 'The attachment\'s number, as read_message gives it.' } }, required: ['account', 'folder', 'id', 'index'], additionalProperties: false },
  },
  {
    name: 'create_draft',
    description: 'Make a draft message in a mail account\'s Drafts, on this computer, for the person to look at and send themselves. Needs the account set to Read and prepare or more.',
    inputSchema: { type: 'object', properties: OUTGOING, required: ['account', 'to'], additionalProperties: false },
  },
  {
    name: 'send_message',
    description: 'Send a message from a mail account set to Read, prepare and act. It is shown to the person in Rutba Office first, with Send, Keep as draft and Cancel, and goes only if they press Send. The answer says what they chose; if they have not chosen yet it says so, with an id for confirmation_answer.',
    inputSchema: { type: 'object', properties: OUTGOING, required: ['account', 'to', 'subject', 'text'], additionalProperties: false },
  },
  {
    name: 'confirmation_answer',
    description: 'What the person chose for a request shown in Rutba Office (a message to send, an invitation, a change for a server): done, kept as a draft, cancelled, or still waiting.',
    inputSchema: { type: 'object', properties: { id: { type: 'string', description: 'The id the request was given.' } }, required: ['id'], additionalProperties: false },
  },
  {
    name: 'list_events',
    description: 'The events of the calendars the tools may read between two dates (the next 30 days when left out), each with its id for read_event.',
    inputSchema: { type: 'object', properties: { calendar: { type: 'string', description: 'A calendar\'s name; every one the tools may read when left out.' }, from: WHEN, to: WHEN }, additionalProperties: false },
  },
  {
    name: 'read_event',
    description: 'One event in full: when, where, its description, organizer and who is invited, with their answers.',
    inputSchema: { type: 'object', properties: { id: STR }, required: ['id'], additionalProperties: false },
  },
  {
    name: 'create_event',
    description: 'Make a new event in a calendar set to Read and prepare or more. People named as attendees are not told: send_invitation does that. In a calendar a server keeps in step, the event goes to the server, so it needs Read, prepare and act and is shown to the person first.',
    inputSchema: {
      type: 'object',
      properties: {
        calendar: { type: 'string', description: 'The calendar\'s name.' },
        summary: STR,
        start: WHEN,
        end: WHEN,
        allDay: { type: 'boolean' },
        location: STR,
        description: STR,
        attendees: { type: 'array', items: { type: 'string' }, description: 'Addresses of people to invite later.' },
      },
      required: ['calendar', 'summary', 'start'],
      additionalProperties: false,
    },
  },
  {
    name: 'send_invitation',
    description: 'Send an event\'s invitation to its attendees, from a mail account set to Read, prepare and act, for an event in a calendar set to Read, prepare and act. Shown to the person first, with Send, Keep as draft and Cancel.',
    inputSchema: { type: 'object', properties: { id: { type: 'string', description: 'The event\'s id.' }, account: { type: 'string', description: 'The mail account it goes from; the first one set to act when left out.' } }, required: ['id'], additionalProperties: false },
  },
  {
    name: 'answer_invitation',
    description: 'Answer an invitation: accept, tentatively accept or decline. The invitation is an event in a calendar the tools may use, or an .ics attachment of a message they may read. The answer is mailed to the organizer from the person\'s own account, so both need Read, prepare and act, and it is shown to the person first.',
    inputSchema: {
      type: 'object',
      properties: {
        answer: { type: 'string', enum: ['accept', 'tentative', 'decline'] },
        id: { type: 'string', description: 'The event\'s id, for an invitation in a calendar.' },
        message: { type: 'object', properties: { ...MESSAGE, index: { type: 'integer', minimum: 0 } }, description: 'The message and attachment number of an invitation that came by mail.' },
        calendar: { type: 'string', description: 'The calendar to keep it in; the first one the tools may use when left out.' },
      },
      required: ['answer'],
      additionalProperties: false,
    },
  },
  {
    name: 'search_contacts',
    description: 'The people in the address books the tools may read, matched on name, address, phone, company or group; every card when the words are left out.',
    inputSchema: { type: 'object', properties: { query: STR, book: { type: 'string', description: 'An address book\'s name.' } }, additionalProperties: false },
  },
  {
    name: 'read_contact',
    description: 'One person\'s card in full.',
    inputSchema: { type: 'object', properties: { id: STR }, required: ['id'], additionalProperties: false },
  },
  {
    name: 'create_contact',
    description: 'Add a person to an address book set to Read and prepare or more: on this computer when no book is named. In a book a server keeps in step the card goes to the server, so it needs Read, prepare and act and is shown to the person first.',
    inputSchema: {
      type: 'object',
      properties: {
        book: { type: 'string', description: 'An address book\'s name.' },
        name: STR,
        emails: { type: 'array', items: STR },
        phones: { type: 'array', items: STR },
        organisation: STR,
        title: STR,
        note: STR,
      },
      required: ['name'],
      additionalProperties: false,
    },
  },
  {
    name: 'list_tables',
    description: 'The tables of a Data Files database the tools may read, with their fields and how many records each has (only those the role chosen for the tools sees, in a protected database).',
    inputSchema: { type: 'object', properties: { database: DATABASE }, required: ['database'], additionalProperties: false },
  },
  {
    name: 'read_records',
    description: 'Records of a table, a page at a time, searched for words when given.',
    inputSchema: { type: 'object', properties: { database: DATABASE, table: STR, search: STR, offset: { type: 'integer', minimum: 0 }, limit: { type: 'integer', minimum: 1, maximum: 500 } }, required: ['database', 'table'], additionalProperties: false },
  },
  {
    name: 'add_record',
    description: 'Add a record to a table of a database set to Read and prepare or more, never its design. A table linked to a database engine writes to the engine, so it needs Read, prepare and act and is shown to the person first.',
    inputSchema: { type: 'object', properties: { database: DATABASE, table: STR, values: { type: 'object', description: 'Field names and their values.' } }, required: ['database', 'table', 'values'], additionalProperties: false },
  },
  {
    name: 'change_record',
    description: 'Change fields of a record (by its _rowid, as read_records gives it) in a database set to Read and prepare or more. A table linked to a database engine writes to the engine, so it needs Read, prepare and act and is shown to the person first.',
    inputSchema: { type: 'object', properties: { database: DATABASE, table: STR, rowid: { type: 'integer' }, values: { type: 'object' } }, required: ['database', 'table', 'rowid', 'values'], additionalProperties: false },
  },
];

export const REACH_TOOLS = new Set(REACH_TOOL_LIST.map((t) => t.name));

/** A date given by an AI app, as milliseconds: an ISO date or date and time; a day alone is that day. */
function whenOf(given, name) {
  const s = words(given);
  if (!s) return null;
  const at = /^\d{4}-\d{2}-\d{2}$/.test(s) ? Date.parse(`${s}T00:00:00Z`) : Date.parse(s);
  if (Number.isNaN(at)) throw Object.assign(new Error(`${name} is not a date Rutba Office reads: give it as ISO 8601, like 2026-11-03T14:00:00Z.`), { refusal: true });
  return at;
}

/**
 * The tools, answering for real. `settings` reads the profile's settings;
 * `mail`, `calendar`, `contacts`, `base` are the running copy's services;
 * `confirm` asks the person (confirm.js); `files.readable(path)` is a file
 * in a folder the tools may read, or throws why not; `readers` turn a
 * document's bytes into its words; `attachmentBytes(args)` an attachment's
 * bytes.
 */
export function createReachTools({ settings, mail = null, calendar = null, contacts = null, base = null, confirm, files, readers = {}, attachmentBytes = null }) {
  /* ── what there is, each with its step ─────────────────────────────── */

  const mailThings = () => [
    ...(mail?.accounts?.() || []).map((a) => ({ kind: 'mail', id: a.id, name: a.email, label: a.name || a.email, local: Boolean(a.local), step: stepOf('mail', a.ai, a) })),
    ...(mail?.archives?.() || []).map((a) => ({ kind: 'archive', id: a.id, name: a.name || path.basename(a.path || ''), label: a.name, step: stepOf('archive', a.ai) })),
  ];
  const calendarThings = () => (calendar?.calendars?.() || []).map((c) => ({ kind: 'calendar', id: c.id, name: c.name, server: Boolean(c.account), step: stepOf('calendar', c.ai) }));
  const bookThings = () => (contacts?.books?.() || []).map((b) => ({ kind: 'book', id: b.id, name: b.id === LOCAL_BOOK ? 'On this computer' : b.name, server: Boolean(b.server), step: stepOf('book', b.ai) }));
  const databaseThings = () => Object.entries(settings.get(DATABASES_KEY, {}) || {}).map(([file]) => {
    const s = databaseSetting(settings.get(DATABASES_KEY, {}), file);
    return { kind: 'database', id: file, name: path.basename(file).replace(/\.[^.]+$/, ''), file, role: s.role, step: s.step };
  });
  const folderThings = () => folderEntries(settings.get(FOLDERS_KEY, [])).map((f) => ({ kind: 'folder', id: f.path, name: f.path, step: f.step }));

  const need = (thing, want) => {
    if (!allows(thing.step, want)) throw refuse(thing.kind, thing.name, thing.step, want);
    return thing;
  };
  const pick = (list, given, what) => {
    const g = words(given);
    const hit = list.find((x) => x.id === g) || list.find((x) => same(x.name, g)) || list.find((x) => same(x.label, g));
    if (!hit) throw notThere(what, g || '(none named)');
    return hit;
  };
  const mailOf = (given, want) => need(pick(mailThings(), given, 'mail account or archive'), want);
  const calendarOf = (id) => calendarThings().find((c) => c.id === id) || null;
  const bookOf = (id) => bookThings().find((b) => b.id === (id || LOCAL_BOOK)) || null;

  /** A file to attach, from a folder the tools may read (item 4). */
  const attachment = (given) => {
    const full = files.readable(given);
    return { filename: path.basename(full), path: full };
  };

  /** The outgoing message as Mail keeps a draft. */
  const draftOf = (args) => ({
    to: words(args.to),
    cc: words(args.cc) || undefined,
    bcc: words(args.bcc) || undefined,
    subject: words(args.subject),
    text: String(args.text ?? ''),
    attachments: (Array.isArray(args.attachments) ? args.attachments : []).map(attachment),
  });

  /** The answer for a request shown to the person: what they chose, or that it waits. */
  const answered = (r, waitingWords) => {
    if (r.waiting) return text(`${waitingWords} It is waiting in Rutba Office for the person to choose. Ask confirmation_answer with id ${r.id} later.`);
    if (r.error) return { content: [{ type: 'text', text: r.error }], isError: true };
    return text(r.outcome);
  };

  /** One database opened for a call, as the role chosen for the tools, and let go after. */
  async function withDatabase(given, want, work) {
    const list = databaseThings();
    const thing = need(pick(list, given, 'database'), want);
    const summary = await base.openForTools({ file: thing.file, role: thing.role });
    try {
      return await work(summary, thing);
    } finally {
      base.close({ id: summary.id });
    }
  }
  const tableOf = (summary, table) => {
    const t = (summary.tables || []).find((x) => x.name === table) || (summary.tables || []).find((x) => same(x.name, table));
    if (!t) throw Object.assign(new Error(`${summary.name} has no table called ${table} that the AI tools may see.`), { refusal: true });
    return t;
  };
  /** A record change: in the file at prepare; to a linked engine (live, or a copy caught up later) at act, confirmed. */
  async function writeRecord(args, kind) {
    const want = 'prepare';
    return withDatabase(args.database, want, async (summary, thing) => {
      const t = tableOf(summary, args.table);
      const values = args.values && typeof args.values === 'object' ? args.values : {};
      const run = async () => {
        // Opened again for the press: the person may answer after this call has let its copy go.
        const again = await base.openForTools({ file: thing.file, role: thing.role });
        try {
          if (kind === 'add') {
            const r = await base.insert({ id: again.id, table: t.name, values });
            return `Added a record to ${t.name} in ${thing.name}${r?.rowid != null ? ` (_rowid ${r.rowid})` : ''}.`;
          }
          const r = await base.update({ id: again.id, table: t.name, rowid: Number(args.rowid), values });
          return r?.changed ? `Changed record ${args.rowid} of ${t.name} in ${thing.name}.` : `Record ${args.rowid} of ${t.name} in ${thing.name} was not changed.`;
        } finally {
          base.close({ id: again.id });
        }
      };
      if (!t.linked) return { result: text(await run()), log: thing };
      need(thing, 'act');
      const r = await confirm.ask(
        { kind: 'database', thing: thing.name, fields: { database: thing.name, table: t.name, record: kind === 'add' ? null : String(args.rowid), values: JSON.stringify(values, null, 1).slice(0, 2000) }, change: kind },
        { act: run }
      );
      return { result: answered(r, `The change to ${t.name} goes to its database engine.`), log: thing, step: 'act' };
    });
  }

  /** An invitation's event as the calendar reads it, from a calendar or a message, with the address it was sent to. */
  async function invitationOf(args) {
    if (args.id) {
      const event = calendar.get({ id: args.id });
      if (!event) throw notThere('event', args.id);
      const cal = calendarOf(event.calendarId);
      return { event, cal, from: { id: args.id } };
    }
    const m = args.message || {};
    const source = mailOf(m.account, 'read');
    const got = await mail.attachmentText({ accountId: source.id, folder: m.folder, id: m.id, index: Number(m.index) });
    if (!got?.text || !/BEGIN:VCALENDAR/i.test(got.text)) throw Object.assign(new Error('That attachment is not an invitation.'), { refusal: true });
    const read = readCalendar(got.text);
    return { event: read.events[0] || null, cal: null, from: { text: got.text }, source };
  }

  const handlers = {
    list_allowed: async () => {
      const all = [...mailThings(), ...calendarThings(), ...bookThings(), ...databaseThings(), ...folderThings()];
      const on = all.filter((x) => x.step !== 'off');
      const kindWords = { mail: 'mail account', archive: 'mail archive', calendar: 'calendar', book: 'address book', database: 'database', folder: 'folder' };
      const lines = on.map((x) => `${kindWords[x.kind]}\t${x.name}\t${STEP_WORDS[x.step]}${x.kind === 'calendar' || x.kind === 'book' ? (x.server ? '\tkept in step with a server' : '\ton this computer') : ''}${x.kind === 'database' && x.role ? `\tas the role ${x.role}` : ''}`);
      const off = all.length - on.length;
      return text([
        on.length ? `kind\tname\tallows\n${lines.join('\n')}` : 'Nothing beyond files is set for the AI tools yet.',
        off ? `${off} more ${off === 1 ? 'is' : 'are'} Off. The person sets each in its own app (Mail, Calendar and Contacts under Accounts, a database in Data Files) or in Rutba Office's AI Tools window.` : '',
      ].filter(Boolean).join('\n\n'));
    },

    /* ── mail ─────────────────────────────────────────────────────────── */

    search_mail: async ({ query, account, limit = 50 }) => {
      const sources = account ? [mailOf(account, 'read')] : mailThings().filter((x) => allows(x.step, 'read'));
      if (!sources.length) throw Object.assign(new Error('No mail account or archive is set to Read yet. The person sets one in Mail under Accounts, or in AI Tools.'), { refusal: true });
      const found = await mail.searchAll({ query: words(query), limit: Math.min(200, Number(limit) || 50), sources: sources.map((s) => s.id) });
      const names = new Map(sources.map((s) => [s.id, s.name]));
      const rows = (found?.rows || []).filter((r) => names.has(r.accountId));
      const out = rows.map((r) => `${names.get(r.accountId)}\t${r.folder}\t${r.id}\t${r.date || ''}\t${addressList(r.from)}\t${r.subject || ''}`);
      return { result: text(out.length ? `account\tfolder\tid\tdate\tfrom\tsubject\n${out.join('\n')}` : `No messages match "${words(query)}".`), log: sources.length === 1 ? sources[0] : { kind: 'mail', name: sources.map((s) => s.name).join(', '), step: 'read' } };
    },

    list_mail_folders: async ({ account }) => {
      const thing = mailOf(account, 'read');
      const folders = (await mail.folders({ accountId: thing.id })) || [];
      return { result: text(folders.length ? `folder\tmessages\tunread\n${folders.map((f) => `${f.path}\t${f.total ?? ''}\t${f.unread ?? ''}`).join('\n')}` : `${thing.name} has no folders yet.`), log: thing };
    },

    list_messages: async ({ account, folder = 'Inbox', offset = 0, limit = 50 }) => {
      const thing = mailOf(account, 'read');
      const page = (await mail.messages({ accountId: thing.id, folder, offset: Number(offset) || 0, limit: Math.min(200, Number(limit) || 50) })) || { rows: [] };
      const out = (page.rows || []).map((r) => `${r.id}\t${r.date || ''}\t${addressList(r.from)}\t${r.subject || ''}${r.unread ? '\tunread' : ''}`);
      return { result: text(out.length ? `${page.total ?? out.length} in ${folder}\nid\tdate\tfrom\tsubject\n${out.join('\n')}` : `No messages in ${folder} of ${thing.name}.`), log: thing };
    },

    read_message: async ({ account, folder, id }) => {
      const thing = mailOf(account, 'read');
      const m = await mail.message({ accountId: thing.id, folder, id });
      if (!m) throw notThere('message', id);
      const body = m.text || stripHtml(m.html) || '';
      const atts = (m.attachments || []).map((a, i) => `${i}\t${a.filename || 'attachment'}\t${a.type || ''}\t${a.size ?? ''}`);
      return {
        result: text(trimTo([
          `From: ${addressList(m.from)}`, `To: ${addressList(m.to)}`, m.cc?.length ? `Cc: ${addressList(m.cc)}` : null, `Date: ${m.date || ''}`, `Subject: ${m.subject || ''}`, '',
          body,
          atts.length ? `\nAttachments (read_attachment):\nindex\tname\ttype\tbytes\n${atts.join('\n')}` : null,
        ].filter((x) => x !== null).join('\n'), MAX_CHARS)),
        log: thing,
      };
    },

    read_attachment: async ({ account, folder, id, index }) => {
      const thing = mailOf(account, 'read');
      const got = attachmentBytes ? await attachmentBytes({ accountId: thing.id, folder, id, index: Number(index) }) : null;
      if (!got?.bytes) throw notThere('attachment', `${index} of that message`);
      const name = String(got.name || 'attachment');
      const ext = path.extname(name).toLowerCase();
      let out;
      if (ext === '.docx' && readers.docx) out = readers.docx(got.bytes);
      else if (ext === '.xlsx' && readers.xlsx) out = readers.xlsx(got.bytes);
      else if (ext === '.pptx' && readers.pptx) out = readers.pptx(got.bytes);
      else if (ext === '.pdf' && readers.pdf) out = readers.pdf(got.bytes);
      else if (/^text\//.test(got.type || '') || ['.txt', '.md', '.csv', '.ics', '.vcf', '.json', '.xml', '.html', '.htm'].includes(ext)) out = Buffer.from(got.bytes).toString('utf8');
      else return { result: { content: [{ type: 'text', text: `${name} is not a kind of file the AI tools read the words of.` }], isError: true }, log: thing };
      return { result: text(trimTo(String(out), MAX_CHARS)), log: thing };
    },

    create_draft: async (args) => {
      const thing = mailOf(args.account, 'prepare');
      if (thing.kind !== 'mail') throw refuse(thing.kind, thing.name, thing.step, 'prepare');
      const draft = draftOf(args);
      await mail.saveDraft({ accountId: thing.id, draft });
      return { result: text(`Kept a draft in ${thing.name}'s Drafts, on this computer, for the person to look at. Nothing was sent.`), log: thing };
    },

    send_message: async (args) => {
      const thing = mailOf(args.account, 'act');
      const draft = draftOf(args);
      if (!draft.to) throw Object.assign(new Error('A message needs someone to go to.'), { refusal: true });
      const keep = async () => {
        await mail.saveDraft({ accountId: thing.id, draft });
        return `Kept as a draft in ${thing.name}'s Drafts, as the person chose. Nothing was sent.`;
      };
      const r = await confirm.ask(
        { kind: 'mail', thing: thing.name, fields: { from: thing.name, to: draft.to, cc: draft.cc || null, bcc: draft.bcc || null, subject: draft.subject, attachments: draft.attachments.map((a) => a.filename).join(', ') || null }, body: draft.text.slice(0, 4000) },
        {
          act: async () => {
            try {
              await mail.send({ accountId: thing.id, draft });
            } catch (err) {
              // Never lost: what the person agreed to send is kept to send again.
              await mail.saveDraft({ accountId: thing.id, draft }).catch(() => {});
              throw new Error(`Sending failed: ${err?.message || err}. The message was kept as a draft in ${thing.name}'s Drafts.`);
            }
            return `Sent from ${thing.name} to ${draft.to}, as the person chose.`;
          },
          draft: keep,
        }
      );
      return { result: answered(r, 'The message is not sent yet.'), log: thing };
    },

    confirmation_answer: async ({ id }) => {
      const s = confirm.status(String(id || ''));
      if (!s) return { result: { content: [{ type: 'text', text: 'Rutba Office knows of no request with that id. It may have been asked before Rutba Office last started.' }], isError: true } };
      if (s.waiting) return { result: text('Still waiting in Rutba Office for the person to choose.') };
      return { result: s.error ? { content: [{ type: 'text', text: s.error }], isError: true } : text(s.outcome) };
    },

    /* ── calendars ────────────────────────────────────────────────────── */

    list_events: async ({ calendar: named, from, to }) => {
      const cals = named ? [need(pick(calendarThings(), named, 'calendar'), 'read')] : calendarThings().filter((c) => allows(c.step, 'read'));
      if (!cals.length) throw Object.assign(new Error('No calendar is set to Read yet. The person sets one in Calendar under Accounts, or in AI Tools.'), { refusal: true });
      const start = whenOf(from, 'from') ?? Date.now();
      const end = whenOf(to, 'to') ?? start + 30 * 86400000;
      const ids = new Set(cals.map((c) => c.id));
      const events = (calendar.events({ from: start, to: end, all: true }) || []).filter((e) => ids.has(e.calendarId));
      const iso = (ms) => (ms == null ? '' : new Date(ms).toISOString());
      const out = events.slice(0, 500).map((e) => `${e.id || ''}\t${e.calendar}\t${iso(e.start)}\t${iso(e.end)}\t${e.allDay ? 'all day' : ''}\t${e.summary}\t${e.location || ''}${e.attendees ? `\t${e.attendees} invited` : ''}`);
      return { result: text(out.length ? `id\tcalendar\tstart\tend\t\tsummary\tlocation\n${out.join('\n')}` : 'No events in those days.'), log: cals.length === 1 ? cals[0] : { kind: 'calendar', name: cals.map((c) => c.name).join(', '), step: 'read' } };
    },

    read_event: async ({ id }) => {
      const event = calendar.get({ id: String(id || '') });
      if (!event) throw notThere('event', id);
      const cal = need(calendarOf(event.calendarId) || { kind: 'calendar', name: event.calendar, step: 'off' }, 'read');
      const at = (d) => (d?.at != null ? new Date(d.at).toISOString() : '');
      return {
        result: text([
          `Summary: ${event.summary || ''}`, `Calendar: ${event.calendar}`, `Start: ${at(event.start)}${event.allDay ? ' (all day)' : ''}`, `End: ${at(event.end)}`,
          event.location ? `Location: ${event.location}` : null, `Repeats: ${event.rule || 'Does not repeat'}`,
          event.organizer?.email ? `Organizer: ${event.organizer.name ? `${event.organizer.name} <${event.organizer.email}>` : event.organizer.email}` : null,
          event.attendees?.length ? `Attendees:\n${event.attendees.map((a) => `${a.name ? `${a.name} <${a.email}>` : a.email}\t${a.partstat || ''}`).join('\n')}` : null,
          event.description ? `\n${event.description}` : null,
        ].filter((x) => x !== null).join('\n')),
        log: cal,
      };
    },

    create_event: async (args) => {
      const cal = need(pick(calendarThings(), args.calendar, 'calendar'), 'prepare');
      const allDay = Boolean(args.allDay);
      const start = whenOf(args.start, 'start');
      const end = whenOf(args.end, 'end') ?? start + (allDay ? 86400000 : 3600000);
      if (end < start) throw Object.assign(new Error('The event ends before it starts.'), { refusal: true });
      const event = {
        summary: words(args.summary),
        location: words(args.location),
        description: String(args.description ?? ''),
        start: dateTimeAt(start, null, allDay),
        end: dateTimeAt(end, null, allDay),
        allDay,
        attendees: (Array.isArray(args.attendees) ? args.attendees : []).map((e) => ({ email: words(e), name: null, partstat: 'NEEDS-ACTION', rsvp: true, role: 'REQ-PARTICIPANT' })).filter((a) => a.email),
      };
      const make = (calendarId) => calendar.save({ calendarId, event });
      if (!cal.server) {
        const kept = make(cal.id);
        return { result: text(`Made the event ${event.summary} in ${cal.name}, on this computer (id ${kept.id}). Nobody was told.`), log: cal };
      }
      need(cal, 'act');
      const local = calendarThings().find((c) => !c.server && allows(c.step, 'prepare'));
      const r = await confirm.ask(
        { kind: 'event', thing: cal.name, fields: { calendar: cal.name, summary: event.summary, start: new Date(start).toISOString(), end: new Date(end).toISOString(), location: event.location || null } },
        {
          act: async () => `Made the event ${event.summary} in ${cal.name} (id ${make(cal.id).id}); it goes to the calendar's server at the next sync.`,
          ...(local ? { draft: async () => `Kept the event ${event.summary} in ${local.name}, on this computer only (id ${make(local.id).id}), as the person chose.` } : {}),
        }
      );
      return { result: answered(r, 'The event is not made yet.'), log: cal };
    },

    send_invitation: async ({ id, account }) => {
      const event = calendar.get({ id: String(id || '') });
      if (!event) throw notThere('event', id);
      const cal = need(calendarOf(event.calendarId) || { kind: 'calendar', name: event.calendar, step: 'off' }, 'act');
      if (!event.attendees?.length) throw Object.assign(new Error(`${event.summary || 'That event'} has nobody to invite.`), { refusal: true });
      const from = account ? mailOf(account, 'act') : mailThings().find((m) => m.kind === 'mail' && allows(m.step, 'act'));
      if (!from) throw Object.assign(new Error('An invitation goes by mail, and no mail account is set to Read, prepare and act. The person sets one in Mail under Accounts, or in AI Tools.'), { refusal: true });
      if (from.kind !== 'mail') throw refuse(from.kind, from.name, from.step, 'act');
      const made = () => {
        const inv = calendar.invitationFile({ id: event.id });
        return { to: inv.to.join(', '), subject: inv.subject, text: '', attachments: [{ filename: 'invite.ics', path: inv.path }] };
      };
      const to = event.attendees.map((a) => a.email).filter(Boolean).join(', ');
      const r = await confirm.ask(
        { kind: 'invitation', thing: cal.name, fields: { from: from.name, to, subject: `Invitation: ${event.summary || 'Event'}`, start: event.start?.at != null ? new Date(event.start.at).toISOString() : null } },
        {
          act: async () => { await mail.send({ accountId: from.id, draft: made() }); return `Sent the invitation to ${to} from ${from.name}, as the person chose.`; },
          draft: async () => { await mail.saveDraft({ accountId: from.id, draft: made() }); return `Kept the invitation as a draft in ${from.name}'s Drafts, as the person chose. Nobody was told.`; },
        }
      );
      return { result: answered(r, 'The invitation is not sent yet.'), log: { ...cal, name: `${cal.name}, from ${from.name}` } };
    },

    answer_invitation: async (args) => {
      const partstat = { accept: 'ACCEPTED', tentative: 'TENTATIVE', decline: 'DECLINED' }[args.answer];
      if (!partstat) throw Object.assign(new Error('The answer is accept, tentative or decline.'), { refusal: true });
      const inv = await invitationOf(args);
      if (!inv.event) throw Object.assign(new Error('There is no invitation to answer there.'), { refusal: true });
      // Kept in the calendar it is in, the one named, or the first the tools may prepare in.
      const cal = need(inv.cal || (args.calendar ? pick(calendarThings(), args.calendar, 'calendar') : calendarThings().find((c) => allows(c.step, 'act')) || calendarThings()[0] || { kind: 'calendar', name: 'a calendar', step: 'off' }), 'act');
      const mine = mailThings().filter((m) => m.kind === 'mail');
      const me = (inv.event.attendees || []).map((a) => mine.find((m) => same(m.name, a.email))).find(Boolean);
      if (!me) throw Object.assign(new Error('None of the person\'s mail accounts is among those invited.'), { refusal: true });
      need(me, 'act');
      const respond = () => {
        const r = calendar.respond({ ...inv.from, partstat, calendarId: cal.id });
        return { to: r.organizer?.email || '', subject: r.subject, text: '', attachments: r.replyPath ? [{ filename: 'reply.ics', path: r.replyPath }] : [] };
      };
      const r = await confirm.ask(
        { kind: 'answer', thing: cal.name, fields: { from: me.name, to: inv.event.organizer?.email || '', subject: inv.event.summary || '', answer: args.answer } },
        {
          act: async () => { const d = respond(); if (d.to) await mail.send({ accountId: me.id, draft: d }); return `Answered ${args.answer} to ${inv.event.summary || 'the invitation'}${d.to ? `, mailed to ${d.to}` : ''}, as the person chose.`; },
          draft: async () => { const d = respond(); await mail.saveDraft({ accountId: me.id, draft: d }); return `Kept the answer in ${cal.name} and its message as a draft in ${me.name}'s Drafts, as the person chose. The organizer was not told.`; },
        }
      );
      return { result: answered(r, 'The answer is not sent yet.'), log: { ...cal, name: `${cal.name}, from ${me.name}` } };
    },

    /* ── contacts ─────────────────────────────────────────────────────── */

    search_contacts: async ({ query = '', book }) => {
      const books = book ? [need(pick(bookThings(), book, 'address book'), 'read')] : bookThings().filter((b) => allows(b.step, 'read'));
      if (!books.length) throw Object.assign(new Error('No address book is set to Read yet. The person sets one in Contacts under Accounts, or in AI Tools.'), { refusal: true });
      const ids = new Set(books.map((b) => b.id));
      const cards = (contacts.list({ query: words(query) }) || []).filter((c) => ids.has(c.book || LOCAL_BOOK));
      const out = cards.slice(0, 300).map((c) => `${c.id}\t${c.display || ''}\t${(c.emails || []).map((e) => e.value).join(', ')}\t${(c.phones || []).map((p) => p.value).join(', ')}\t${c.org || ''}`);
      return { result: text(out.length ? `id\tname\temails\tphones\torganisation\n${out.join('\n')}` : 'Nobody matches.'), log: books.length === 1 ? books[0] : { kind: 'book', name: books.map((b) => b.name).join(', '), step: 'read' } };
    },

    read_contact: async ({ id }) => {
      const c = contacts.get({ id: String(id || '') });
      if (!c) throw notThere('card', id);
      const book = need(bookOf(c.book) || { kind: 'book', name: 'its address book', step: 'off' }, 'read');
      const list = (xs) => (xs || []).map((x) => `${x.value}${x.type ? ` (${x.type})` : ''}`).join(', ');
      return {
        result: text([
          `Name: ${c.display || ''}`, c.org ? `Organisation: ${c.org}${c.department ? `, ${c.department}` : ''}` : null, c.title ? `Title: ${c.title}` : null,
          c.emails?.length ? `Email: ${list(c.emails)}` : null, c.phones?.length ? `Phone: ${list(c.phones)}` : null,
          (c.addresses || []).length ? `Address: ${c.addresses.map((a) => [a.street, a.locality, a.region, a.postcode, a.country].filter(Boolean).join(', ')).join('; ')}` : null,
          c.birthday ? `Birthday: ${c.birthday}` : null, (c.categories || []).length ? `Groups: ${c.categories.join(', ')}` : null, c.note ? `Note: ${c.note}` : null,
          `Address book: ${book.name}`,
        ].filter((x) => x !== null).join('\n')),
        log: book,
      };
    },

    create_contact: async (args) => {
      const book = need(args.book ? pick(bookThings(), args.book, 'address book') : bookOf(LOCAL_BOOK) || { kind: 'book', name: 'On this computer', step: 'off' }, 'prepare');
      const name = words(args.name);
      const [given = '', ...rest] = name.split(/\s+/);
      const contact = {
        uid: null, version: '3.0',
        name: { full: name, given, family: rest.join(' '), middle: '', prefix: '', suffix: '' },
        nickname: null, org: words(args.organisation) || null, department: null, title: words(args.title) || null, role: null,
        emails: (args.emails || []).map((v, i) => ({ value: words(v), type: 'other', pref: i === 0, label: null })).filter((e) => e.value),
        phones: (args.phones || []).map((v, i) => ({ value: words(v), type: 'cell', pref: i === 0, label: null })).filter((p) => p.value),
        addresses: [], birthday: null, anniversary: null, note: words(args.note) || null, urls: [], photo: null, categories: [], rev: null, kind: 'individual', extra: [],
      };
      if (!book.server) {
        const kept = contacts.save({ contact, book: LOCAL_BOOK });
        return { result: text(`Added ${name} to the address book on this computer (id ${kept.id}).`), log: book };
      }
      need(book, 'act');
      const local = bookOf(LOCAL_BOOK);
      const r = await confirm.ask(
        { kind: 'contact', thing: book.name, fields: { book: book.name, name, emails: contact.emails.map((e) => e.value).join(', ') || null, phones: contact.phones.map((p) => p.value).join(', ') || null } },
        {
          act: async () => `Added ${name} to ${book.name} (id ${contacts.save({ contact, book: book.id }).id}); the card goes to the server at the next sync.`,
          ...(local && allows(local.step, 'prepare') ? { draft: async () => `Kept ${name} on this computer only (id ${contacts.save({ contact, book: LOCAL_BOOK }).id}), as the person chose.` } : {}),
        }
      );
      return { result: answered(r, 'The card is not added yet.'), log: book };
    },

    /* ── Data Files ───────────────────────────────────────────────────── */

    list_tables: async ({ database }) => withDatabase(database, 'read', async (summary, thing) => {
      const out = [];
      for (const t of summary.tables || []) {
        const fields = (base.fields({ id: summary.id, table: t.name }) || []).map((f) => `${f.name} (${f.type}${f.readOnly ? ', read only' : ''})`).join(', ');
        out.push(`${t.name}\t${t.count ?? ''} records${t.linked ? `\tlinked to ${t.mode === 'copy' ? 'an engine, kept as a copy' : 'an engine'}` : ''}\n  ${fields}`);
      }
      return { result: text(out.length ? out.join('\n') : `${thing.name} has no tables the AI tools may see.`), log: thing };
    }),

    read_records: async ({ database, table, search = '', offset = 0, limit = 100 }) => withDatabase(database, 'read', async (summary, thing) => {
      const t = tableOf(summary, table);
      // A live table's records are read from its engine as the tools ask, as a window's are.
      if (t.linked && t.mode !== 'copy') await base.refreshLinked({ id: summary.id, table: t.name, now: true });
      const page = base.rows({ id: summary.id, table: t.name, offset: Number(offset) || 0, limit: Math.min(500, Number(limit) || 100), search: words(search) });
      const rows = page.rows || [];
      const head = rows.length ? Object.keys(rows[0]) : [];
      const cell = (v) => (v == null ? '' : typeof v === 'object' ? JSON.stringify(v) : String(v)).replace(/[\t\n]/g, ' ');
      return { result: text(trimTo(rows.length ? `${page.count} records; ${rows.length} from ${Number(offset) || 0}\n${head.join('\t')}\n${rows.map((r) => head.map((h) => cell(r[h])).join('\t')).join('\n')}` : `No records in ${t.name}${search ? ` match "${search}"` : ''}.`, MAX_CHARS)), log: thing };
    }),

    add_record: async (args) => writeRecord(args, 'add'),
    change_record: async (args) => writeRecord(args, 'change'),
  };

  return {
    list: () => REACH_TOOL_LIST.map((t) => ({ ...t })),
    /** Everything there is, each with its step: the AI Tools window lists them. */
    things: () => [...mailThings(), ...calendarThings(), ...bookThings(), ...databaseThings(), ...folderThings()],
    /**
     * One call: { result, log }, `log` the thing it used and the step that
     * call needed ({ kind, name, step }) for the log, never what was read.
     */
    async call(name, args = {}) {
      const handler = handlers[name];
      if (!handler) return { result: { content: [{ type: 'text', text: `Rutba Office has no tool called ${name}.` }], isError: true } };
      try {
        const out = await handler(args || {});
        const thing = out?.log || null;
        return { result: out?.result || out, log: thing ? { kind: thing.kind, name: thing.name, step: out.step || stepNeeded(name, thing) } : null };
      } catch (err) {
        return { result: { content: [{ type: 'text', text: err?.message || String(err) }], isError: true }, log: err?.refusal ? err.log : null };
      }
    },
  };
}

/** The step a call used, for the log: what it needed of the thing. */
function stepNeeded(name, thing) {
  if (['send_message', 'send_invitation', 'answer_invitation'].includes(name)) return 'act';
  if (['create_draft', 'add_record', 'change_record'].includes(name)) return 'prepare';
  if (['create_event', 'create_contact'].includes(name)) return thing.server ? 'act' : 'prepare';
  return 'read';
}

export { STEPS };
