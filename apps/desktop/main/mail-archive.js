// Mail as the archive tool: archives opened in place, one search over every
// account and archive, and the export for legal hold.
//
// An archive opened here is not imported. Its file is opened for reading and
// stays where it is. Everything about it - opening it, listing its folders
// and messages, reading one, indexing it for search and searching it - runs
// on a thread of its own (@rutba/mailbox/archive-host), never on this
// process, which answers every window: a 5 GB .pst indexing for half a minute
// holds nobody's typing. A search while it indexes answers from what has
// been indexed so far and says how far that is; indexing can be stopped and
// carried on. The list of opened archives is a setting; nothing else about
// them is kept.
//
// The accounts' mail is searched the same way: their part of the index is
// kept and searched on a thread of its own too, which reads the mail store's
// folders as this process writes them, so a search over all mail never
// indexes anything here.
//
// To the window an archive looks like an account whose id begins
// `archive:` - mail.js sends that id's folders, lists, messages and
// attachments here - except that nothing can change it: no flags, no moves,
// no deletes.

import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { createArchiveHost } from '@rutba/mailbox/archive-host';
import { exportMessages, storeLoader, freshFolder } from '@rutba/mailbox/export';
import { attachmentFileName } from '@rutba/mailbox/mime';

export const ARCHIVE = 'archive:';
export const isArchive = (id) => typeof id === 'string' && id.startsWith(ARCHIVE);

const READ_ONLY = 'An archive opened in place is read only. Import it to change its mail.';

/** Rows of several searches as one, newest first, as the index orders its own. */
const newestFirst = (a, b) => (b.t || 0) - (a.t || 0) || (a.id < b.id ? -1 : a.id > b.id ? 1 : 0);

/**
 * @param {{ stores, store, broadcast, holdBlob, insightFor, classify, accounts: () => object[], unicodeFont?: () => any, host?: object }} deps
 */
export function createArchiveDesk({ stores, store, broadcast, holdBlob, insightFor, classify, accounts, unicodeFont = () => null, host = null }) {
  const listed = () => stores.settings.get('mail.archives', []);
  const setListed = (list) => stores.settings.set('mail.archives', list);
  const status = new Map(); // id -> { state, done, total, error }
  const jobs = new Map(); // export id -> { cancelled }

  const threads =
    host ||
    createArchiveHost({
      onProgress: (id, p) => {
        if (!listed().some((a) => a.id === id)) return;
        const held = status.get(id) || {};
        // A thread lost is said once; the next question starts it again.
        status.set(id, { ...held, ...p, done: p.state === 'failed' && !p.done ? held.done || 0 : p.done, total: p.total || held.total || 0 });
        broadcast?.('mail:indexing', { id, ...status.get(id) });
      },
    });

  const entryOf = (id) => {
    const entry = listed().find((a) => a.id === id);
    if (!entry) throw new Error('That archive is not open.');
    return entry;
  };
  /** Ask an archive's thread, starting it (the file opened again) when it is not running. */
  const ask = (id, op, args) => {
    const entry = entryOf(id);
    return threads.call(id, op, args, { file: entry.path, name: entry.name });
  };

  const describe = (entry) => {
    const s = status.get(entry.id) || { state: 'waiting', done: 0, total: entry.messages || 0 };
    let missing = false;
    try {
      fs.accessSync(entry.path, fs.constants.R_OK);
    } catch {
      missing = true;
    }
    return { ...entry, archive: true, missing, index: { state: s.state, done: s.done, total: s.total, error: s.error || null } };
  };

  /** Start indexing one archive on its thread (or carry on where it was stopped). */
  async function index(id, { resume = false } = {}) {
    const entry = listed().find((a) => a.id === id);
    if (!entry) return null;
    const held = status.get(id);
    if (held?.state === 'ready' || held?.state === 'indexing') return held;
    if (held?.state === 'paused' && !resume) return held;
    status.set(id, { state: 'indexing', done: held?.done || 0, total: held?.total || entry.messages || 0 });
    broadcast?.('mail:indexing', { id, ...status.get(id) });
    try {
      const s = await ask(id, 'index');
      status.set(id, { ...status.get(id), ...s });
    } catch (err) {
      status.set(id, { state: 'failed', done: 0, total: 0, error: err.message });
      broadcast?.('mail:indexing', { id, ...status.get(id) });
    }
    return status.get(id);
  }

  /** The accounts' thread: their part of the index, read from the store's folders. */
  const ACCOUNTS = 'accounts';
  const askAccounts = (op, args) => threads.call(ACCOUNTS, op, args, { file: store.root, name: ACCOUNTS, kind: 'accounts' });

  /** Every account's part of the index brought up to date, on its thread. */
  async function syncAccounts() {
    await askAccounts('sync', { accounts: accounts().map((a) => ({ id: a.id, name: a.name || a.email })) });
  }

  /** The archives still being indexed, for a search to say its answer is not yet whole. */
  const indexing = () =>
    listed()
      .map((a) => ({ id: a.id, name: a.name, ...(status.get(a.id) || { state: 'waiting', done: 0, total: a.messages || 0 }) }))
      .filter((a) => a.state !== 'ready' && a.state !== 'failed')
      .map(({ id, name, state, done, total }) => ({ id, name, state, done, total }));

  /**
   * One search over the accounts named (here, from the accounts' index) and
   * the archives named (each on its thread), as one answer. Each is asked
   * for enough rows to fill the page after `offset`, and the page is cut from
   * them all, newest first.
   */
  async function searchSources(query, { limit, offset, sources, slim = false }) {
    const archiveIds = sources.filter(isArchive);
    const accountIds = sources.filter((id) => !isArchive(id));
    const want = offset + limit;
    const parts = [];
    if (accountIds.length) parts.push(await askAccounts('search', { query, limit: want, offset: 0, sources: accountIds, slim }));
    const found = await Promise.all(
      archiveIds.map((id) =>
        ask(id, 'search', { query, limit: want, offset: 0, slim }).catch(() => ({ rows: [], total: 0, exact: true, searched: 0 }))
      )
    );
    parts.push(...found);
    const rows = parts.flatMap((p) => p.rows);
    rows.sort(newestFirst);
    return {
      rows: rows.slice(offset, offset + limit),
      total: parts.reduce((n, p) => n + p.total, 0),
      exact: parts.every((p) => p.exact),
      searched: parts.reduce((n, p) => n + (p.searched || 0), 0),
    };
  }

  const desk = {
    /** Opened archives, with how far their indexing has got. */
    archives: () => listed().map(describe),

    /** Open an archive where it lies. The same file twice is the one archive. */
    openArchive: async ({ path: file }) => {
      const full = path.resolve(file);
      const already = listed().find((a) => path.resolve(a.path).toLowerCase() === full.toLowerCase());
      if (already) {
        index(already.id).catch(() => {});
        return describe(already);
      }
      const id = `${ARCHIVE}${crypto.createHash('sha1').update(full.toLowerCase()).digest('hex').slice(0, 10)}`;
      const name = path.basename(full, path.extname(full));
      const opened = await threads.open(id, full, name);
      const entry = {
        id,
        path: full,
        name,
        kind: opened.kind,
        label: opened.label,
        bytes: opened.bytes,
        folders: opened.folders.length,
        messages: opened.folders.reduce((n, f) => n + (f.total || 0), 0),
        openedAt: new Date().toISOString(),
      };
      setListed([...listed(), entry]);
      broadcast?.('mail:archives', { id });
      index(id).catch(() => {});
      return describe(entry);
    },

    /** Close it: forgotten, its thread and its part of the index gone. The file is untouched. */
    closeArchive: async ({ id }) => {
      setListed(listed().filter((a) => a.id !== id));
      status.delete(id);
      await threads.close(id);
      broadcast?.('mail:archives', { id, closed: true });
      return { closed: true };
    },

    /** Pause an archive's indexing; what has been indexed stays searchable. */
    stopIndexing: async ({ id }) => {
      if (!threads.has(id)) return status.get(id) || null;
      const s = await ask(id, 'pause');
      status.set(id, { ...status.get(id), ...s });
      broadcast?.('mail:indexing', { id, ...status.get(id) });
      return status.get(id);
    },

    /** Carry on indexing where it was paused. */
    resumeIndexing: ({ id }) => index(id, { resume: true }),

    /** Start indexing every archive listed (at start-up, and when the window first asks). */
    indexAll: () => {
      for (const a of listed()) index(a.id).catch(() => {});
      return { archives: listed().length };
    },

    folders: async ({ accountId }) =>
      (await ask(accountId, 'folders')).map((f) => {
        const kind = classify(f.name || f.path);
        return { path: f.path, name: f.name, depth: f.depth, total: f.total, unread: f.unread || 0, ...kind, order: kind.order, readOnly: true };
      }),

    messages: async ({ accountId, folder, offset = 0, limit = 100, query = '', unreadOnly = false }) => {
      const page = await ask(accountId, 'list', { folder, offset, limit, query, unreadOnly });
      return { total: page.total, rows: page.rows.map((r) => ({ ...r, accountId, folder })) };
    },

    message: async ({ accountId, folder, id }) => {
      const message = await ask(accountId, 'message', { folder, id });
      if (!message) return null;
      const inline = {};
      for (const a of message.attachments || []) {
        if (!a.inline || !a.bytes || !a.contentId) continue;
        inline[String(a.contentId).replace(/^<|>$/g, '')] = holdBlob(Buffer.from(a.bytes), a.type, a.filename, { group: 'mail-message', generation: `${accountId}|${folder}|${id}` }).url;
      }
      const entry = listed().find((a) => a.id === accountId);
      return {
        ...message,
        id,
        folder,
        accountId,
        readOnly: true,
        archiveName: entry?.name || null,
        // The bytes stay here; the window fetches one when it is opened or saved.
        attachments: (message.attachments || []).map((a) => ({ ...a, bytes: undefined, stored: Boolean(a.bytes), size: a.size ?? a.bytes?.length ?? 0 })),
        inlineUrls: inline,
      };
    },

    insight: async ({ accountId, folder, id }) => {
      const message = await ask(accountId, 'message', { folder, id });
      return message ? insightFor(message) : null;
    },

    attachment: async ({ accountId, folder, id, index: at }) => {
      const message = await ask(accountId, 'message', { folder, id });
      const a = message?.attachments?.[at];
      if (!a?.bytes) return null;
      const bytes = Buffer.from(a.bytes);
      const name = attachmentFileName(a.filename);
      return { ...holdBlob(bytes, a.type, name, { group: 'mail-attachment', generation: `${accountId}|${folder}|${id}|${at}` }), name, type: a.type, size: bytes.length };
    },

    attachmentText: async ({ accountId, folder, id, index: at }) => {
      const message = await ask(accountId, 'message', { folder, id });
      const a = message?.attachments?.[at];
      if (!a?.bytes) return null;
      return { name: a.filename, type: a.type, text: Buffer.from(a.bytes).toString('utf8') };
    },

    /** One account's search, on the accounts' thread, its index brought up to date first. */
    searchAccount: async ({ accountId, query, limit = 200 }) => {
      await syncAccounts();
      return (await askAccounts('search', { query, limit, sources: [accountId] })).rows;
    },

    /** An account removed: its part of the index goes. */
    dropAccount: async ({ id }) => {
      if (threads.has(ACCOUNTS)) await askAccounts('drop', { accountId: id });
    },

    /** One archive's search (its own window's search box), from what is indexed so far. */
    search: async ({ accountId, query, limit = 200 }) => (await ask(accountId, 'search', { query, limit })).rows,

    readOnly: () => {
      throw new Error(READ_ONLY);
    },

    /**
     * One search over every account and every opened archive (or the
     * `sources` named), newest first, each row saying where it lives, and
     * the archives still being indexed with how far each has got.
     */
    searchAll: async ({ query, limit = 500, offset = 0, sources = null }) => {
      await syncAccounts();
      const all = [...accounts().map((a) => a.id), ...listed().map((a) => a.id)];
      const wanted = sources?.length ? all.filter((id) => sources.includes(id)) : all;
      const found = await searchSources(query, { limit, offset, sources: wanted });
      const files = new Map(listed().map((a) => [a.id, a]));
      return {
        ...found,
        rows: found.rows.map((r) => ({ ...r, sourceFile: files.get(r.accountId)?.path || null, readOnly: isArchive(r.accountId) })),
        indexing: indexing().filter((a) => wanted.includes(a.id)),
      };
    },

    /**
     * Export for legal hold: a search's results (`query`, across `sources`),
     * folders (`folders: [{ accountId, folder }]`) or messages named (`items`),
     * into a new folder under `dir`. Progress goes out as mail:exportProgress;
     * `cancelExport` stops it.
     */
    exportHold: async ({ job = crypto.randomUUID(), query = null, sources = null, folders = null, items = null, dir, mbox = true, pdf = 'combined', attachments = false, labels = {}, selection = '', folderName = null }) => {
      if (!dir) throw new Error('Choose a folder to export to.');
      const control = { cancelled: false };
      jobs.set(job, control);
      try {
        await syncAccounts();
        const names = new Map([...accounts().map((a) => [a.id, a.name || a.email]), ...listed().map((a) => [a.id, a.name])]);
        const files = new Map(listed().map((a) => [a.id, a.path]));
        const kinds = new Map(listed().map((a) => [a.id, a.kind]));
        const itemOf = (accountId, folder, row) => ({
          sourceId: accountId,
          sourceName: names.get(accountId) || accountId,
          sourceKind: kinds.get(accountId) || 'account',
          sourceFile: files.get(accountId) || null,
          folder,
          id: row.id,
          date: row.date || (row.t ? new Date(row.t).toISOString() : null),
        });
        let list = [];
        if (items?.length) list = items.map((i) => itemOf(i.accountId, i.folder, i));
        else if (query) {
          const all = [...accounts().map((a) => a.id), ...listed().map((a) => a.id)];
          const wanted = sources?.length ? all.filter((id) => sources.includes(id)) : all;
          list = (await searchSources(query, { limit: 1e9, offset: 0, sources: wanted, slim: true })).rows.map((r) => itemOf(r.accountId, r.folder, r));
        } else if (folders?.length) {
          for (const { accountId, folder } of folders) {
            const rows = isArchive(accountId) ? (await ask(accountId, 'list', { folder, limit: 1e9, slim: true })).rows : store.list(accountId, folder, { limit: 1e9 }).rows;
            for (const r of rows) list.push(itemOf(accountId, folder, r));
          }
        }
        const loadStore = storeLoader(store);
        // An archive's messages are read on its thread, the bytes as it holds them.
        const load = (item) => (isArchive(item.sourceId) ? ask(item.sourceId, 'load', { folder: item.folder, id: item.id }) : loadStore(item));
        const stamp = new Date();
        const p = (n) => String(n).padStart(2, '0');
        const out = freshFolder(dir, folderName || `Mail export ${stamp.getFullYear()}-${p(stamp.getMonth() + 1)}-${p(stamp.getDate())} ${p(stamp.getHours())}${p(stamp.getMinutes())}`);
        broadcast?.('mail:exportProgress', { job, done: 0, total: list.length, dir: out });
        let told = 0;
        const result = await exportMessages({
          items: list,
          load,
          dir: out,
          mbox,
          pdf,
          attachments,
          selection,
          labels,
          unicodeFont: unicodeFont(),
          cancelled: () => control.cancelled,
          onProgress: ({ done, total }) => {
            if (done === total || Date.now() - told > 150) {
              told = Date.now();
              broadcast?.('mail:exportProgress', { job, done, total, dir: out });
            }
          },
        });
        return { job, ...result };
      } finally {
        jobs.delete(job);
      }
    },

    cancelExport: ({ job }) => {
      const control = jobs.get(job);
      if (control) control.cancelled = true;
      return { cancelled: Boolean(control) };
    },

    /** Stop every archive's thread (the service closing). */
    stop: () => threads.stop(),

    /** Threads running, for the checks. */
    threads: () => threads.stats,
  };
  return desk;
}
