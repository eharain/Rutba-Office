// The mail service.
//
// Two halves that meet in the local store: an IMAP/SMTP client for live
// accounts, and importers for the archives other clients leave behind. Both
// write the same records, so an imported .pst folder and a synchronised IMAP
// folder are the same thing to the window drawing them.
//
// Passwords never live here. They go to the operating system's keystore through
// the shell's `secrets` namespace and come back only for the moment a
// connection is being made.

import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { MailStore } from '@rutba/mailbox/store';
import { parseMessage } from '@rutba/mailbox/mime';
import { scan, read as readArchive, identify } from '@rutba/mailbox/import';
import { writeMbox } from '@rutba/mailbox/mbox';
import { insightFor } from './mail-insight.js';
import { planRules, applyPlan } from './mail-rules.js';

/** Well-known providers, so most people never type a server name. */
const PROVIDERS = [
  { match: /@(gmail|googlemail)\.com$/i, imap: { host: 'imap.gmail.com', port: 993, secure: true }, smtp: { host: 'smtp.gmail.com', port: 465, secure: true }, note: 'Gmail needs an app password when two-step verification is on.' },
  { match: /@(outlook|hotmail|live|msn)\.[a-z.]+$/i, imap: { host: 'outlook.office365.com', port: 993, secure: true }, smtp: { host: 'smtp.office365.com', port: 587, secure: false } },
  { match: /@yahoo\.[a-z.]+$/i, imap: { host: 'imap.mail.yahoo.com', port: 993, secure: true }, smtp: { host: 'smtp.mail.yahoo.com', port: 465, secure: true }, note: 'Yahoo requires an app password.' },
  { match: /@(icloud|me|mac)\.com$/i, imap: { host: 'imap.mail.me.com', port: 993, secure: true }, smtp: { host: 'smtp.mail.me.com', port: 587, secure: false }, note: 'iCloud requires an app-specific password.' },
  { match: /@(proton|protonmail)\.(com|me)$/i, imap: { host: '127.0.0.1', port: 1143, secure: false }, smtp: { host: '127.0.0.1', port: 1025, secure: false }, note: 'Proton Mail needs its Bridge running locally.' },
  { match: /@fastmail\.[a-z.]+$/i, imap: { host: 'imap.fastmail.com', port: 993, secure: true }, smtp: { host: 'smtp.fastmail.com', port: 465, secure: true } },
  { match: /@zoho\.[a-z.]+$/i, imap: { host: 'imap.zoho.com', port: 993, secure: true }, smtp: { host: 'smtp.zoho.com', port: 465, secure: true } },
];

/** The conventional names, tried when a domain is not one we know. */
function guessFromDomain(domain) {
  return {
    imap: { host: `imap.${domain}`, port: 993, secure: true },
    smtp: { host: `smtp.${domain}`, port: 465, secure: true },
    guessed: true,
  };
}

const SPECIAL = [
  { test: /^inbox$/i, role: 'inbox', icon: 'inbox', order: 0 },
  { test: /(drafts?)$/i, role: 'drafts', icon: 'file', order: 1 },
  { test: /(sent)/i, role: 'sent', icon: 'send', order: 2 },
  { test: /(junk|spam|bulk)/i, role: 'junk', icon: 'spam', order: 3 },
  { test: /(trash|deleted|bin)/i, role: 'trash', icon: 'trash', order: 4 },
  { test: /(archive)/i, role: 'archive', icon: 'archive', order: 5 },
];

function classify(name) {
  const leaf = String(name).split(/[/.]/).pop();
  for (const s of SPECIAL) if (s.test.test(leaf)) return s;
  return { role: 'folder', icon: 'folder', order: 9 };
}

export function createMailService({ stores, holdBlob, broadcast, userData, oauth = null }) {
  const store = new MailStore(path.join(userData, 'mail'));

  const accounts = () => stores.settings.get('mail.accounts', []);
  const setAccounts = (list) => {
    stores.settings.set('mail.accounts', list);
    return list;
  };

  const passwordKey = (id) => `mail:${id}`;

  /**
   * How to prove who we are to this account's servers.
   *
   * An account added by signing in to Google or Microsoft has no password to
   * store — it has a refresh token in the keystore, and a fresh access token is
   * fetched for each connection. Both IMAP and SMTP take the same shape, so the
   * decision is made once here rather than in two places that could disagree.
   */
  async function credentialsFor(account) {
    const user = account.imap?.user || account.email;
    if (account.auth === 'oauth' || account.imap?.auth === 'oauth') {
      if (!oauth) throw new Error('This build cannot sign in to that provider.');
      const { accessToken } = await oauth.accessToken({ email: account.email, provider: account.provider });
      return { user, accessToken };
    }
    const password = stores.secrets.get(passwordKey(account.id));
    if (!password) throw new Error(`No password is stored for ${account.email}.`);
    return { user, pass: password };
  }

  /** ImapFlow is loaded when an account is actually used, not at start-up. */
  async function imapFor(account) {
    const { ImapFlow } = await import('imapflow');
    const client = new ImapFlow({
      host: account.imap.host,
      port: account.imap.port,
      secure: account.imap.secure,
      auth: await credentialsFor(account),
      logger: false,
      tls: { rejectUnauthorized: account.imap.rejectUnauthorized !== false },
    });
    await client.connect();
    return client;
  }

  /**
   * The folder this account calls its Trash, its Junk, its Archive.
   * A rule says "junk"; the account may call it Spam, Bulk Mail or
   * [Gmail]/Spam, and putting a message in a folder that does not exist is
   * how mail gets lost.
   */
  const folderFor = (accountId, role) =>
    (store.folders(accountId) || []).find((f) => classify(f.name || f.path).role === role)?.path || null;

  const find = (id) => {
    const account = accounts().find((a) => a.id === id);
    if (!account) throw new Error('That account is not set up.');
    return account;
  };

  /**
   * Send whatever is due, then sleep until the next one.
   *
   * One timer for the whole queue rather than one per message: a queue read
   * from settings after a restart has no timers at all, and rebuilding N of
   * them would be a way to get N sends out of one message. The queue is the
   * truth; the timer is only how we wake up.
   */
  let outboxTimer = null;
  function pumpOutbox() {
    if (outboxTimer) clearTimeout(outboxTimer);
    outboxTimer = null;

    const outbox = stores.settings.get('mail.outbox', []);
    if (!outbox.length) return;

    const now = Date.now();
    const due = outbox.filter((o) => new Date(o.at).getTime() <= now);
    const waiting = outbox.filter((o) => new Date(o.at).getTime() > now);

    if (due.length) {
      // Taken off the queue *before* sending, so a crash mid-send cannot send
      // the same message twice on the next start. A failure puts it back.
      stores.settings.set('mail.outbox', waiting);
      for (const item of due) {
        service
          .send({ accountId: item.accountId, draft: item.draft })
          .then(() => broadcast?.('mail:sent', { id: item.id, to: item.draft?.to || '' }))
          .catch((error) => {
            const message = String(error?.message || error);
            const attempts = (item.attempts || 0) + 1;
            if (attempts < 3) {
              const held = stores.settings.get('mail.outbox', []);
              stores.settings.set('mail.outbox', [
                ...held,
                { ...item, attempts, error: message, at: new Date(Date.now() + 5 * 60_000).toISOString() },
              ]);
              pumpOutbox();
            } else {
              // Three failures is not a network blip. Keep the words, put them
              // back in Drafts, and say so — losing it silently is the one
              // outcome that is never acceptable.
              try {
                service.saveDraft({ accountId: item.accountId, draft: item.draft });
              } catch {
                /* the account itself is gone */
              }
            }
            broadcast?.('mail:sendFailed', { id: item.id, message, attempts, gaveUp: attempts >= 3 });
          });
      }
    }

    if (waiting.length) {
      const soonest = Math.min(...waiting.map((o) => new Date(o.at).getTime()));
      outboxTimer = setTimeout(pumpOutbox, Math.min(Math.max(soonest - Date.now(), 500), 60_000));
      outboxTimer.unref?.();
    }
  }

  const service = {
    accounts: () =>
      accounts().map((a) => ({
        ...a,
        hasPassword: Boolean(stores.secrets.get(passwordKey(a.id))),
        counts: (store.folders(a.id) || []).reduce(
          (acc, f) => {
            const c = store.counts(a.id, f.path);
            return { total: acc.total + c.total, unread: acc.unread + c.unread };
          },
          { total: 0, unread: 0 }
        ),
      })),

    autodiscover: ({ email }) => {
      const domain = String(email).split('@')[1] || '';
      const known = PROVIDERS.find((p) => p.match.test(email));
      if (known) return { imap: known.imap, smtp: known.smtp, note: known.note || null, source: 'known' };
      if (!domain) return { imap: null, smtp: null, source: 'none' };
      return { ...guessFromDomain(domain), note: 'These server names are a guess from the domain.', source: 'guess' };
    },

    addAccount: ({ account, password }) => {
      const id = account.id || crypto.randomUUID().slice(0, 8);
      const record = {
        id,
        email: account.email,
        name: account.name || account.email,
        imap: account.imap,
        smtp: account.smtp,
        colour: account.colour || null,
        signature: account.signature || '',
        addedAt: new Date().toISOString(),
        local: Boolean(account.local),
        // A signed-in account has no password to keep. What it has is a refresh
        // token, already in the keystore under the address, and these two say
        // which door to knock on for a live one.
        auth: account.auth === 'oauth' ? 'oauth' : 'password',
        provider: account.provider || null,
      };
      if (password) stores.secrets.set(passwordKey(id), password);
      setAccounts([...accounts().filter((a) => a.id !== id), record]);
      return record;
    },

    updateAccount: ({ id, patch, password }) => {
      const list = accounts();
      const at = list.findIndex((a) => a.id === id);
      if (at < 0) throw new Error('That account is not set up.');
      list[at] = { ...list[at], ...patch };
      if (password) stores.secrets.set(passwordKey(id), password);
      setAccounts(list);
      return list[at];
    },

    removeAccount: ({ id, keepMail }) => {
      setAccounts(accounts().filter((a) => a.id !== id));
      stores.secrets.delete(passwordKey(id));
      if (!keepMail) {
        try {
          fs.rmSync(path.join(userData, 'mail', id), { recursive: true, force: true });
        } catch {
          /* the store may already be gone */
        }
      }
      return true;
    },

    testAccount: async ({ account, password }) => {
      const result = { imap: null, smtp: null, error: null };
      try {
        const { ImapFlow } = await import('imapflow');
        const client = new ImapFlow({
          host: account.imap.host,
          port: account.imap.port,
          secure: account.imap.secure,
          auth: { user: account.imap.user || account.email, pass: password },
          logger: false,
        });
        await client.connect();
        const list = await client.list();
        result.imap = { ok: true, folders: list.length };
        await client.logout();
      } catch (err) {
        result.imap = { ok: false, message: err.message };
        result.error = err.message;
      }
      try {
        const nodemailer = (await import('nodemailer')).default;
        const transport = nodemailer.createTransport({
          host: account.smtp.host,
          port: account.smtp.port,
          secure: account.smtp.secure,
          auth: { user: account.smtp.user || account.email, pass: password },
        });
        await transport.verify();
        result.smtp = { ok: true };
      } catch (err) {
        result.smtp = { ok: false, message: err.message };
        result.error = result.error || err.message;
      }
      return result;
    },

    folders: ({ accountId }) => {
      const stored = store.folders(accountId);
      return stored
        .map((f) => ({ ...f, ...classify(f.name || f.path), ...store.counts(accountId, f.path) }))
        .sort((a, b) => a.order - b.order || a.path.localeCompare(b.path));
    },

    /** Pull new mail. Live accounts only; an imported archive has nothing to pull. */
    sync: async ({ accountId, folder, limit = 300 }) => {
      const account = find(accountId);
      if (account.local) return { added: 0, total: store.counts(accountId, folder || 'Inbox').total, local: true };

      const client = await imapFor(account);
      let added = 0;
      try {
        const list = await client.list();
        const folders = list
          .filter((f) => !f.flags?.has?.('\\Noselect'))
          .map((f) => ({ path: f.path, name: f.name, ...classify(f.name) }));
        store.setFolders(accountId, folders);

        const targets = folder ? folders.filter((f) => f.path === folder) : folders.filter((f) => f.role === 'inbox');
        for (const target of targets) {
          const lock = await client.getMailboxLock(target.path);
          try {
            const total = client.mailbox.exists;
            if (!total) continue;
            const from = Math.max(1, total - limit + 1);
            const batch = [];
            for await (const msg of client.fetch(`${from}:*`, { uid: true, flags: true, source: true })) {
              const parsed = parseMessage(msg.source);
              parsed.uid = msg.uid;
              parsed.unread = !msg.flags?.has?.('\\Seen');
              parsed.flagged = Boolean(msg.flags?.has?.('\\Flagged'));
              batch.push(parsed);
              if (batch.length >= 50) {
                added += store.putMany(accountId, target.path, batch.splice(0));
                broadcast?.('mail:progress', { accountId, folder: target.path, done: added, total, phase: 'fetch' });
              }
            }
            if (batch.length) added += store.putMany(accountId, target.path, batch);
          } finally {
            lock.release();
          }
        }
      } finally {
        await client.logout().catch(() => {});
      }
      // Rules run on arrival, over what has just arrived rather than over the
      // whole folder — a rule is about the message coming in, and re-running
      // one across 50,000 old messages every fetch would be a different and
      // much slower feature.
      let filed = null;
      if (added) {
        const rules = stores.settings.get('mail.rules', []);
        if (rules.some((r) => r.enabled !== false)) {
          const { rows } = store.list(accountId, folder || 'Inbox', { limit: added });
          const plan = planRules(rows.map((r) => ({ ...r, folder: folder || 'Inbox' })), rules);
          if (plan.length) filed = applyPlan(plan, { store, accountId, folderFor });
        }
      }

      if (added) broadcast?.('mail:new', { accountId, folder, count: added, filed });
      return { added, filed, total: store.counts(accountId, folder || 'Inbox').total };
    },

    messages: ({ accountId, folder, offset = 0, limit = 100, query = '', unreadOnly = false }) =>
      store.list(accountId, folder, { offset, limit, query, unreadOnly }),

    /**
     * One list across every account.
     *
     * A unified inbox is not a view of one folder; it is the same folder *role*
     * in each account, merged and re-sorted. Each row keeps the account and
     * folder it came from, because every action on it — flag, move, delete —
     * has to go back to the right store.
     */
    unified: ({ role = 'inbox', offset = 0, limit = 200, query = '', unreadOnly = false }) => {
      const rows = [];
      for (const account of accounts()) {
        const folders = store.folders(account.id) || [];
        const wanted = folders.filter((f) => (role === 'all' ? true : classify(f.name || f.path).role === role));
        for (const folder of wanted) {
          const page = store.list(account.id, folder.path, { limit: 1000, query, unreadOnly });
          for (const row of page.rows) {
            rows.push({ ...row, accountId: account.id, accountName: account.name || account.email, folder: folder.path });
          }
        }
      }
      rows.sort((a, b) => (b.date || '').localeCompare(a.date || ''));
      return { total: rows.length, rows: rows.slice(offset, offset + limit) };
    },

    /* ── rules ────────────────────────────────────────────────────────── */

    rules: () => stores.settings.get('mail.rules', []),

    saveRule: ({ rule }) => {
      const rules = stores.settings.get('mail.rules', []);
      const id = rule.id || crypto.randomUUID().slice(0, 8);
      const record = { ...rule, id };
      const at = rules.findIndex((r) => r.id === id);
      if (at >= 0) rules[at] = record;
      else rules.push(record);
      stores.settings.set('mail.rules', rules);
      return record;
    },

    deleteRule: ({ id }) => {
      stores.settings.set('mail.rules', stores.settings.get('mail.rules', []).filter((r) => r.id !== id));
      return { removed: true };
    },

    /**
     * What the rules would do, without doing any of it.
     *
     * The reason anybody writes a rule is the mail already sitting in the
     * folder, and the reason people distrust rules is that a wrong one moves a
     * thousand messages before they can look. So this answers first.
     */
    testRules: ({ accountId, folder, rules: only = null, limit = 40 }) => {
      const rules = only || stores.settings.get('mail.rules', []);
      const { rows } = store.list(accountId, folder, { limit: 100000 });
      const plan = planRules(rows.map((r) => ({ ...r, folder })), rules);
      return {
        matched: plan.length,
        of: rows.length,
        sample: plan.slice(0, limit).map(({ row, rule, actions }) => ({
          id: row.id,
          subject: row.subject,
          from: row.from,
          date: row.date,
          rule: rule?.name || rule?.id || null,
          actions: actions.map((a) => (a.value ? `${a.type} → ${a.value}` : a.type)),
        })),
      };
    },

    /** Run them for real, over one folder. */
    runRules: ({ accountId, folder, rules: only = null }) => {
      const rules = only || stores.settings.get('mail.rules', []);
      const { rows } = store.list(accountId, folder, { limit: 100000 });
      const plan = planRules(rows.map((r) => ({ ...r, folder })), rules);
      const tally = applyPlan(plan, { store, accountId, folderFor });
      if (tally.matched) broadcast?.('mail:new', { accountId, folder, count: 0 });
      return tally;
    },

    /** Mark a whole folder read — the button every list needs and few have. */
    markAllRead: ({ accountId, folder }) => {
      const { rows } = store.list(accountId, folder, { limit: 100000, unreadOnly: true });
      for (const row of rows) store.setFlags(accountId, folder, row.id, { unread: false });
      return { changed: rows.length };
    },

    /** Empty a folder. Only ever offered for Trash and Junk. */
    emptyFolder: ({ accountId, folder }) => {
      const { rows } = store.list(accountId, folder, { limit: 100000 });
      for (const row of rows) store.remove(accountId, folder, row.id);
      return { removed: rows.length };
    },

    message: ({ accountId, folder, id }) => {
      const message = store.get(accountId, folder, id);
      if (!message) return null;
      // Inline images are handed to the window as URLs so the message body can
      // reference them without the bytes crossing as JSON.
      const inline = {};
      for (const a of message.attachments || []) {
        if (!a.inline || !a.stored || !a.contentId) continue;
        try {
          const bytes = fs.readFileSync(store.attachmentPath(accountId, folder, id, a.stored));
          inline[a.contentId] = holdBlob(bytes, a.type, a.filename).url;
        } catch {
          /* an attachment that did not survive a copy */
        }
      }
      return { ...message, inlineUrls: inline };
    },

    /**
     * Who is watching, can I leave, and is it really them.
     *
     * Computed from the stored message alone — no network call — so it is as
     * true for a 2009 message in an imported .pst as for one that arrived a
     * minute ago. See mail-insight.js for what each part means.
     */
    insight: ({ accountId, folder, id }) => {
      const message = store.get(accountId, folder, id);
      return message ? insightFor(message) : null;
    },

    /**
     * Every attachment in the mailbox, newest first.
     *
     * A person looking for "that spreadsheet Ali sent" is not looking for a
     * message, and making them find the message first is a design that only
     * exists because listing attachments is harder. Only messages the index
     * already marks as carrying one are opened, so the cost is proportional to
     * the attachments, not to the mailbox.
     */
    files: ({ accountId, query = '', limit = 400 }) => {
      const q = query.trim().toLowerCase();
      const out = [];
      for (const folder of store.folders(accountId)) {
        const { rows } = store.list(accountId, folder.path, { limit: 100000 });
        for (const row of rows) {
          if (!row.hasAttachments) continue;
          const message = store.get(accountId, folder.path, row.id);
          for (const [index, a] of (message?.attachments || []).entries()) {
            if (a.inline || !a.filename) continue;
            if (q && !a.filename.toLowerCase().includes(q) && !(row.subject || '').toLowerCase().includes(q)) continue;
            out.push({
              filename: a.filename,
              type: a.type,
              size: a.size || 0,
              stored: Boolean(a.stored),
              index,
              id: row.id,
              folder: folder.path,
              subject: row.subject,
              from: row.from,
              date: row.date,
            });
          }
          if (out.length >= limit * 4) break;
        }
      }
      out.sort((a, b) => (b.date || '').localeCompare(a.date || ''));
      return { total: out.length, rows: out.slice(0, limit) };
    },

    /**
     * The people in this mailbox, ranked by how much you actually deal with
     * them. Derived from the folder indexes, so it needs no address book and
     * is correct the moment an archive finishes importing.
     */
    people: ({ accountId, limit = 500 }) => {
      const seen = new Map();
      const note = (person, direction) => {
        const address = String(person?.address || '').toLowerCase();
        if (!address || !address.includes('@')) return;
        const entry = seen.get(address) || { address, name: person.name || null, received: 0, sent: 0, last: null };
        if (!entry.name && person.name) entry.name = person.name;
        entry[direction]++;
        seen.set(address, entry);
      };
      for (const folder of store.folders(accountId)) {
        const outbound = folder.role === 'sent' || folder.role === 'drafts';
        const { rows } = store.list(accountId, folder.path, { limit: 100000 });
        for (const row of rows) {
          if (outbound) for (const to of row.to || []) note(to, 'sent');
          else note(row.from, 'received');
          const entry = seen.get(String((outbound ? row.to?.[0] : row.from)?.address || '').toLowerCase());
          if (entry && (!entry.last || (row.date || '') > entry.last)) entry.last = row.date;
        }
      }
      const rows = [...seen.values()].sort((a, b) => b.received + b.sent - (a.received + a.sent));
      return { total: rows.length, rows: rows.slice(0, limit) };
    },

    flag: ({ accountId, folder, ids, patch }) => {
      for (const id of ids || []) store.setFlags(accountId, folder, id, patch);
      return store.counts(accountId, folder);
    },

    move: ({ accountId, folder, ids, to }) => {
      for (const id of ids || []) {
        const message = store.get(accountId, folder, id);
        if (!message) continue;
        store.put(accountId, to, message, { force: true });
        store.remove(accountId, folder, id);
      }
      store.upsertFolder(accountId, { path: to, name: to.split('/').pop() });
      return { moved: (ids || []).length };
    },

    delete: ({ accountId, folder, ids }) => {
      for (const id of ids || []) store.remove(accountId, folder, id);
      return { deleted: (ids || []).length };
    },

    /**
     * Hold a message before it goes, or until a chosen time.
     *
     * Everything pressed Send goes here first. With a delay of a few seconds
     * that is the undo window every web client has and almost no desktop client
     * does; with a date it is a scheduled send. It is the same queue either way,
     * so a scheduled message can still be recalled, and the queue survives a
     * restart because it is written to settings rather than held in a timer.
     */
    queue: ({ accountId, draft, at = null, holdSeconds = 0 }) => {
      const when = at ? new Date(at).toISOString() : new Date(Date.now() + holdSeconds * 1000).toISOString();
      const item = { id: crypto.randomUUID(), accountId, draft, at: when, queuedAt: new Date().toISOString() };
      const outbox = stores.settings.get('mail.outbox', []);
      stores.settings.set('mail.outbox', [...outbox, item]);
      pumpOutbox();
      return item;
    },

    /** Everything still waiting, so the window can show an Outbox that is real. */
    outbox: () => stores.settings.get('mail.outbox', []),

    /** Take one back. Nothing has left, so this always works. */
    unsend: ({ id }) => {
      const outbox = stores.settings.get('mail.outbox', []);
      const item = outbox.find((o) => o.id === id) || null;
      stores.settings.set('mail.outbox', outbox.filter((o) => o.id !== id));
      return item;
    },

    send: async ({ accountId, draft }) => {
      const account = find(accountId);
      const nodemailer = (await import('nodemailer')).default;
      const credentials = await credentialsFor(account);
      const transport = nodemailer.createTransport({
        host: account.smtp.host,
        port: account.smtp.port,
        secure: account.smtp.secure,
        auth: credentials.accessToken
          ? { type: 'OAuth2', user: account.smtp.user || account.email, accessToken: credentials.accessToken }
          : { user: account.smtp.user || account.email, pass: credentials.pass },
      });
      const info = await transport.sendMail({
        from: { name: account.name, address: account.email },
        to: draft.to,
        cc: draft.cc || undefined,
        bcc: draft.bcc || undefined,
        subject: draft.subject || '',
        text: draft.text || '',
        html: draft.html || undefined,
        attachments: (draft.attachments || []).map((a) => ({ filename: a.filename, path: a.path })),
      });
      // A sent message belongs in Sent whether or not the server puts it there.
      const stored = parseMessage(
        `From: ${account.email}\r\nTo: ${draft.to}\r\nSubject: ${draft.subject || ''}\r\n` +
          `Date: ${new Date().toUTCString()}\r\nMessage-ID: <${info.messageId || crypto.randomUUID()}>\r\n\r\n${draft.text || ''}`
      );
      stored.unread = false;
      store.put(accountId, 'Sent', stored);
      store.upsertFolder(accountId, { path: 'Sent', name: 'Sent', role: 'sent' });
      return { messageId: info.messageId };
    },

    saveDraft: ({ accountId, draft }) => {
      const message = parseMessage(
        `From: ${draft.from || ''}\r\nTo: ${draft.to || ''}\r\nSubject: ${draft.subject || ''}\r\n` +
          `Date: ${new Date().toUTCString()}\r\nMessage-ID: <draft-${draft.id || crypto.randomUUID()}>\r\n\r\n${draft.text || ''}`
      );
      const id = store.put(accountId, 'Drafts', message, { force: true });
      store.upsertFolder(accountId, { path: 'Drafts', name: 'Drafts', role: 'drafts' });
      return { id };
    },

    /** Look inside an archive without importing any of it. */
    importScan: ({ path: target }) => {
      const found = scan(target);
      return { ...found, kindLabel: found.label, identified: identify(target).kind };
    },

    /**
     * Bring an archive in. Progress is broadcast so the window can show it, and
     * the import is resumable in the only sense that matters: running it twice
     * adds nothing, because a message's id is derived from the message.
     */
    import: ({ path: target, accountId, folders: only }) => {
      const account = accountId ? find(accountId) : null;
      const id = account
        ? account.id
        : (() => {
            const made = {
              id: crypto.randomUUID().slice(0, 8),
              email: path.basename(target),
              name: path.basename(target, path.extname(target)),
              local: true,
              imported: target,
              addedAt: new Date().toISOString(),
            };
            setAccounts([...accounts(), made]);
            return made.id;
          })();

      const seenFolders = new Map();
      let count = 0;
      let batchFolder = null;
      let batch = [];

      const flush = () => {
        if (!batch.length || !batchFolder) return;
        count += store.putMany(id, batchFolder, batch);
        batch = [];
        broadcast?.('mail:progress', { accountId: id, folder: batchFolder, done: count, total: 0, phase: 'import' });
      };

      for (const { folder, message } of readArchive(target, { folders: only })) {
        if (folder !== batchFolder) {
          flush();
          batchFolder = folder;
        }
        batch.push(message);
        if (!seenFolders.has(folder)) seenFolders.set(folder, classify(folder));
        if (batch.length >= 100) flush();
      }
      flush();

      store.setFolders(
        id,
        [...seenFolders.entries()].map(([folderPath, meta]) => ({
          path: folderPath,
          name: folderPath.split('/').pop() || folderPath,
          ...meta,
        }))
      );

      broadcast?.('mail:new', { accountId: id, count });
      return { accountId: id, messages: count, folders: seenFolders.size };
    },

    export: ({ accountId, folder, path: target, format = 'mbox' }) => {
      const { rows } = store.list(accountId, folder, { limit: 1e6 });
      const messages = rows.map((r) => store.get(accountId, folder, r.id)).filter(Boolean);
      if (format === 'mbox') {
        const raw = messages.map((m) => ({
          from: m.from?.[0],
          date: m.date,
          raw:
            `From: ${m.from?.[0]?.display || ''}\r\nTo: ${(m.to || []).map((t) => t.display).join(', ')}\r\n` +
            `Subject: ${m.subject}\r\nDate: ${m.date ? new Date(m.date).toUTCString() : ''}\r\n\r\n${m.text || m.html || ''}`,
        }));
        fs.writeFileSync(target, Buffer.from(writeMbox(raw)));
      } else {
        fs.mkdirSync(target, { recursive: true });
        messages.forEach((m, i) => {
          const name = `${String(i + 1).padStart(5, '0')}-${(m.subject || 'message').replace(/[^\w. -]/g, '_').slice(0, 60)}.eml`;
          fs.writeFileSync(
            path.join(target, name),
            `From: ${m.from?.[0]?.display || ''}\r\nTo: ${(m.to || []).map((t) => t.display).join(', ')}\r\n` +
              `Subject: ${m.subject}\r\nDate: ${m.date ? new Date(m.date).toUTCString() : ''}\r\n\r\n${m.text || ''}`
          );
        });
      }
      return { messages: messages.length, path: target };
    },

    attachment: ({ accountId, folder, id, index }) => {
      const message = store.get(accountId, folder, id);
      const a = message?.attachments?.[index];
      if (!a?.stored) return null;
      const bytes = fs.readFileSync(store.attachmentPath(accountId, folder, id, a.stored));
      return { ...holdBlob(bytes, a.type, a.filename), name: a.filename, type: a.type, size: bytes.length };
    },

    search: ({ accountId, query, limit }) => store.search(accountId, query, { limit }),
  };

  // Anything left over from the last run goes out now.
  pumpOutbox();

  return service;
}
