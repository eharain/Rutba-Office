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

export function createMailService({ stores, holdBlob, broadcast, userData }) {
  const store = new MailStore(path.join(userData, 'mail'));

  const accounts = () => stores.settings.get('mail.accounts', []);
  const setAccounts = (list) => {
    stores.settings.set('mail.accounts', list);
    return list;
  };

  const passwordKey = (id) => `mail:${id}`;

  /** ImapFlow is loaded when an account is actually used, not at start-up. */
  async function imapFor(account) {
    const { ImapFlow } = await import('imapflow');
    const password = stores.secrets.get(passwordKey(account.id));
    if (!password) throw new Error(`No password is stored for ${account.email}.`);
    const client = new ImapFlow({
      host: account.imap.host,
      port: account.imap.port,
      secure: account.imap.secure,
      auth: { user: account.imap.user || account.email, pass: password },
      logger: false,
      tls: { rejectUnauthorized: account.imap.rejectUnauthorized !== false },
    });
    await client.connect();
    return client;
  }

  const find = (id) => {
    const account = accounts().find((a) => a.id === id);
    if (!account) throw new Error('That account is not set up.');
    return account;
  };

  return {
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
      if (added) broadcast?.('mail:new', { accountId, folder, count: added });
      return { added, total: store.counts(accountId, folder || 'Inbox').total };
    },

    messages: ({ accountId, folder, offset = 0, limit = 100, query = '', unreadOnly = false }) =>
      store.list(accountId, folder, { offset, limit, query, unreadOnly }),

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

    send: async ({ accountId, draft }) => {
      const account = find(accountId);
      const nodemailer = (await import('nodemailer')).default;
      const password = stores.secrets.get(passwordKey(account.id));
      const transport = nodemailer.createTransport({
        host: account.smtp.host,
        port: account.smtp.port,
        secure: account.smtp.secure,
        auth: password ? { user: account.smtp.user || account.email, pass: password } : undefined,
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
}
