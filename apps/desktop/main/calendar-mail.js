// Mail as the calendar reaches it: the accounts, the one new mail goes from,
// a way to send, and the Inbox to read invitations and answers out of.
//
// This is the only file that knows Mail's calls. Mail is reached through
// `packs` when it is asked, so it may register before or after Calendar, or
// be absent: then there are no accounts, nothing is sent, and nothing is
// read. Sending goes through Mail's `sendFrom` (the 2.25 contract: from an
// account, by id or address, or the one chosen for new mail; filed in Sent;
// attachments as content only, the first typed text/calendar with a method
// sent as the RFC 6047 invitation part by Mail itself); a Mail without it is
// sent through with its `send` and the account's id, the calendar file
// handed over as a file on disk. A failure keeps Mail's `code`
// ('no-account', 'cannot-send') for the calendar to say which it was.
// What arrives is heard from Mail's `onArrived` as it comes; a Mail without
// it has its Inboxes read on a timer (calendar-pack.js).

import fs from 'node:fs';

const norm = (s) => String(s || '').trim().toLowerCase();

/** A Mail account as the calendar speaks of it. */
const asAccount = (a) => (a ? { id: a.id, address: a.address || a.email, name: a.name || null, room: a.room || null } : null);

/** Mail's arrival as the calendar reads it: what arrived where, and each calendar part's text and sender. */
function asArrival(e) {
  const invitations = (Array.isArray(e?.invitations) ? e.invitations : []).filter((x) => x && x.id != null).map((x) => ({
    id: x.id,
    subject: String(x.subject || ''),
    from: x.from?.address ? { name: x.from.name || '', address: String(x.from.address) } : null,
    method: x.method ? String(x.method).toUpperCase() : null,
    text: typeof x.text === 'string' && x.text ? x.text : null,
  }));
  return {
    accountId: e?.accountId ?? null,
    folder: e?.folder ?? null,
    ids: Array.isArray(e?.ids) ? e.ids : invitations.map((x) => x.id),
    at: e?.at || new Date().toISOString(),
    invitations,
  };
}

/**
 * @param {() => object | null} get Mail's service, or null while it is absent
 */
export function createMailLink(get) {
  const mail = () => {
    try {
      return get() || null;
    } catch {
      return null;
    }
  };
  const accounts = () => {
    try {
      return mail()?.accounts?.() || [];
    } catch {
      return [];
    }
  };
  const link = {
    /** Mail's own accounts, as it lists them. */
    accounts,

    /** Whether there is a Mail to send through. */
    present: () => Boolean(mail()),

    /** The account new mail goes from: the one chosen in Mail, else the first. */
    defaultAccount: () => {
      const m = mail();
      try {
        if (typeof m?.defaultAccount === 'function') return asAccount(m.defaultAccount());
      } catch {
        /* the first account, below */
      }
      // An imported mailbox is named after its file: it has no address to send from.
      return asAccount(accounts().find((a) => /@/.test(a.email || '')));
    },

    /** The account of one of this person's addresses, or null. */
    accountFor: (address) => asAccount(accounts().find((a) => norm(a.email) === norm(address))),

    /**
     * Send a message: `{ account, to, cc, subject, text, html, attachments, headers }`,
     * `account` an id or an address (none: the one new mail goes from), `to`
     * and `cc` lists of addresses, each attachment `{ filename, content,
     * contentType, path }` (the calendar's carry both). Answers
     * `{ messageId, refused }` as Mail does; throws with Mail's `code`.
     */
    sendFrom: async ({ account = null, to = [], cc = [], subject = '', text = '', html = undefined, attachments = [], headers = undefined }) => {
      const m = mail();
      if (!m) throw Object.assign(new Error('Mail is not installed, so there is no account to send from.'), { code: 'no-account' });
      if (typeof m.sendFrom === 'function') {
        // Content only: Mail refuses a path. Mail builds the invitation part from the first calendar attachment itself.
        const files = attachments.map((a) => ({ filename: a.filename, content: a.content ?? fs.readFileSync(a.path, 'utf8'), ...(a.contentType ? { contentType: a.contentType } : {}) }));
        return m.sendFrom({ ...(account ? { account } : {}), to, ...(cc.length ? { cc } : {}), subject, text, ...(html ? { html } : {}), attachments: files, ...(headers ? { headers } : {}) });
      }
      // Mail's older send takes an account's id: an address is looked up.
      const accountId = /@/.test(String(account || '')) ? accounts().find((a) => norm(a.email) === norm(account))?.id : account || link.defaultAccount()?.id;
      if (!accountId) throw Object.assign(new Error('There is no Mail account to send from.'), { code: 'no-account' });
      const draft = {
        to: to.join(', '),
        cc: cc.length ? cc.join(', ') : undefined,
        subject,
        text,
        html,
        headers,
        attachments: attachments.map((a) => ({ filename: a.filename, path: a.path, ...(a.path ? {} : { content: a.content }) })),
      };
      return m.send({ accountId, draft });
    },

    /** An account's Inbox, as Mail names it: the folder whose role is inbox, else INBOX. */
    inboxOf: (accountId) => {
      try {
        const folders = mail()?.folders?.({ accountId }) || [];
        return (folders.find((f) => f.role === 'inbox') || folders.find((f) => /^inbox$/i.test(f.path)))?.path || null;
      } catch {
        return null;
      }
    },

    /** The newest messages of a folder, as Mail's list gives them. */
    latest: ({ accountId, folder, limit = 50 }) => {
      try {
        return mail()?.messages?.({ accountId, folder, limit })?.rows || [];
      } catch {
        return [];
      }
    },

    /** A message's calendar part, as text with who sent it, or null when it has none. */
    calendarPart: ({ accountId, folder, id }) => {
      const m = mail();
      try {
        const message = m?.message?.({ accountId, folder, id });
        const index = (message?.attachments || []).findIndex((a) => /text\/calendar|application\/ics/i.test(a.type || '') || /\.ics$/i.test(a.filename || ''));
        if (index < 0) return { message, part: null };
        return { message, part: m.attachmentText({ accountId, folder, id, index }) };
      } catch {
        return { message: null, part: null };
      }
    },

    /**
     * Mail arrived, as it comes (Mail's `onArrived`, 2.25): `listener({
     * accountId, folder, ids, at, invitations: [{ id, subject, from: { name,
     * address }, method, text }] })`, junk left out, once each arrival has
     * been through Mail's rules. Answers the function that stops it, or null
     * when this Mail cannot tell (absent, or older than 2.25): then the
     * Inboxes are read on a timer instead.
     */
    onArrived: (listener) => {
      const m = mail();
      if (typeof m?.onArrived !== 'function') return null;
      try {
        const stop = m.onArrived((e) => listener(asArrival(e)));
        return typeof stop === 'function' ? stop : () => {};
      } catch {
        return null;
      }
    },

    /** A message read whole (the words of an enquiry). */
    message: ({ accountId, folder, id }) => {
      try {
        return mail()?.message?.({ accountId, folder, id }) || null;
      } catch {
        return null;
      }
    },
  };
  return link;
}

/** Mail through the packs, as the calendar's pack hands it over. */
export function mailFor(packs) {
  return createMailLink(() => packs.service('mail'));
}
