// Mail as the calendar reaches it: the accounts, the one new mail goes from,
// a way to send, and the Inbox to read invitations and answers out of.
//
// This is the only file that knows Mail's calls. Mail is reached through
// `packs` when it is asked, so it may register before or after Calendar, or
// be absent: then there are no accounts, nothing is sent, and nothing is
// read. Sending goes through Mail's `sendFrom` (the 2.25 contract: from an
// account, or the one chosen for new mail, filed in Sent as compose files
// it); a Mail without it is sent through with its `send` and the account
// named, the calendar file handed over as a file on disk.

const norm = (s) => String(s || '').trim().toLowerCase();

/** A Mail account as the calendar speaks of it. */
const asAccount = (a) => (a ? { id: a.id, address: a.address || a.email, name: a.name || null, room: a.room || null } : null);

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
     * `to` and `cc` lists of addresses, each attachment `{ filename, content,
     * contentType, path }` (the calendar's always carry a path). Answers
     * `{ messageId, refused }` as Mail does.
     */
    sendFrom: async ({ account = null, to = [], cc = [], subject = '', text = '', html = undefined, attachments = [], headers = undefined }) => {
      const m = mail();
      if (!m) throw new Error('Mail is not installed, so there is no account to send from.');
      if (typeof m.sendFrom === 'function') return m.sendFrom({ account, to, cc, subject, text, html, attachments, headers });
      const accountId = account || link.defaultAccount()?.id;
      if (!accountId) throw new Error('There is no Mail account to send from.');
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
