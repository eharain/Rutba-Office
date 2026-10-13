// Mail as the archive tool: archives opened in place, one search over every
// account and archive, and the export for legal hold.
//
// An archive opened here is not imported. Its file is opened for reading and
// stays where it is; its folders and messages are read from it when asked
// for, and it is indexed into the same in-memory index as the accounts
// (MailStore's `index`), a slice at a time so the windows keep answering. The
// list of opened archives is a setting; nothing else about them is kept.
//
// To the window an archive looks like an account whose id begins
// `archive:` - mail.js sends that id's folders, lists, messages and
// attachments here - except that nothing can change it: no flags, no moves,
// no deletes.

import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { openArchive, indexSteps, runSliced } from '@rutba/mailbox/archive';
import { segmentKey } from '@rutba/mailbox/store';
import { exportMessages, storeLoader, archiveLoader, freshFolder } from '@rutba/mailbox/export';
import { attachmentFileName } from '@rutba/mailbox/mime';

export const ARCHIVE = 'archive:';
export const isArchive = (id) => typeof id === 'string' && id.startsWith(ARCHIVE);

const READ_ONLY = 'An archive opened in place is read only. Import it to change its mail.';

/**
 * @param {{ stores, store, broadcast, holdBlob, insightFor, classify, accounts: () => object[], unicodeFont?: () => any }} deps
 */
export function createArchiveDesk({ stores, store, broadcast, holdBlob, insightFor, classify, accounts, unicodeFont = () => null }) {
  const listed = () => stores.settings.get('mail.archives', []);
  const setListed = (list) => stores.settings.set('mail.archives', list);
  const opened = new Map(); // id -> source
  const status = new Map(); // id -> { state, done, total, error }
  const jobs = new Map(); // export id -> { cancelled }

  const sourceOf = (id) => {
    if (opened.has(id)) return opened.get(id);
    const entry = listed().find((a) => a.id === id);
    if (!entry) throw new Error('That archive is not open.');
    const source = openArchive(entry.path);
    opened.set(id, source);
    return source;
  };

  const describe = (entry) => {
    const s = status.get(entry.id) || { state: 'waiting', done: 0, total: 0 };
    let missing = false;
    try {
      fs.accessSync(entry.path, fs.constants.R_OK);
    } catch {
      missing = true;
    }
    return { ...entry, archive: true, missing, index: { state: s.state, done: s.done, total: s.total, error: s.error || null } };
  };

  /** Index one archive, a slice at a time, saying how far it has got. */
  async function index(id) {
    const entry = listed().find((a) => a.id === id);
    if (!entry) return;
    const held = status.get(id);
    if (held?.state === 'indexing' || held?.state === 'ready') return;
    let source;
    try {
      source = sourceOf(id);
    } catch (err) {
      status.set(id, { state: 'failed', done: 0, total: 0, error: err.message });
      broadcast?.('mail:indexing', { id, ...status.get(id) });
      return;
    }
    const total = source.folders().reduce((n, f) => n + (f.total || 0), 0);
    const s = { state: 'indexing', done: 0, total };
    status.set(id, s);
    let told = 0;
    try {
      await runSliced(indexSteps(source, store.index, { id, name: entry.name }), {
        slice: 24,
        onStep: (step) => {
          s.done = step.done;
          if (step.done - told >= 500 || Date.now() - told > 400) {
            told = step.done;
            broadcast?.('mail:indexing', { id, state: s.state, done: s.done, total: s.total });
          }
        },
        cancelled: () => !listed().some((a) => a.id === id),
      });
      s.state = 'ready';
      s.done = total;
    } catch (err) {
      s.state = 'failed';
      s.error = err.message;
    }
    broadcast?.('mail:indexing', { id, state: s.state, done: s.done, total: s.total });
  }

  /** Every account's part of the index brought up to date, in slices. */
  async function syncAccounts() {
    for (const a of accounts()) {
      await runSliced(store.indexSteps(a.id, { name: a.name || a.email }), { slice: 24 });
    }
  }

  /** The archives still being indexed, for a search to say its answer is not yet whole. */
  const indexing = () =>
    listed()
      .map((a) => ({ id: a.id, name: a.name, ...(status.get(a.id) || { state: 'waiting', done: 0, total: 0 }) }))
      .filter((a) => a.state !== 'ready' && a.state !== 'failed');

  const folderRows = (id, folder) => {
    const segment = store.index.segments.get(segmentKey(id, folder));
    if (!segment) return null;
    return segment.docs;
  };

  const rowOf = (id, folder, doc) => ({
    ...doc,
    date: doc.t ? new Date(doc.t).toISOString() : doc.date || null,
    accountId: id,
    folder,
  });

  const desk = {
    /** Opened archives, with how far their indexing has got. */
    archives: () => listed().map(describe),

    /** Open an archive where it lies. The same file twice is the one archive. */
    openArchive: ({ path: file }) => {
      const full = path.resolve(file);
      const already = listed().find((a) => path.resolve(a.path).toLowerCase() === full.toLowerCase());
      if (already) {
        index(already.id).catch(() => {});
        return describe(already);
      }
      const source = openArchive(full);
      const folders = source.folders();
      const id = `${ARCHIVE}${crypto.createHash('sha1').update(full.toLowerCase()).digest('hex').slice(0, 10)}`;
      const entry = {
        id,
        path: full,
        name: path.basename(full, path.extname(full)),
        kind: source.kind,
        label: source.label,
        bytes: source.bytes,
        folders: folders.length,
        messages: folders.reduce((n, f) => n + (f.total || 0), 0),
        openedAt: new Date().toISOString(),
      };
      opened.set(id, source);
      setListed([...listed(), entry]);
      broadcast?.('mail:archives', { id });
      index(id).catch(() => {});
      return describe(entry);
    },

    /** Close it: forgotten, its part of the index dropped. The file is untouched. */
    closeArchive: ({ id }) => {
      setListed(listed().filter((a) => a.id !== id));
      try {
        opened.get(id)?.close();
      } catch {
        /* already closed */
      }
      opened.delete(id);
      status.delete(id);
      store.dropIndex(id);
      broadcast?.('mail:archives', { id, closed: true });
      return { closed: true };
    },

    /** Start indexing every archive listed (at start-up, and when the window first asks). */
    indexAll: () => {
      for (const a of listed()) index(a.id).catch(() => {});
      return { archives: listed().length };
    },

    folders: ({ accountId }) =>
      sourceOf(accountId)
        .folders()
        .map((f) => {
          const kind = classify(f.name || f.path);
          return { path: f.path, name: f.name, depth: f.depth, total: f.total, unread: f.unread || 0, ...kind, order: kind.order, readOnly: true };
        }),

    messages: ({ accountId, folder, offset = 0, limit = 100, query = '', unreadOnly = false }) => {
      if (query) {
        const found = store.index.search(query, { keys: [segmentKey(accountId, folder)], limit, offset });
        if (store.index.segments.has(segmentKey(accountId, folder))) return { total: found.total, rows: found.rows };
      }
      const docs = folderRows(accountId, folder);
      if (docs) {
        let rows = docs;
        if (unreadOnly) rows = rows.filter((d) => d.unread);
        rows = rows.slice().sort((a, b) => b.t - a.t);
        return { total: rows.length, rows: rows.slice(offset, offset + limit).map((d) => rowOf(accountId, folder, d)) };
      }
      // A folder the archive does not have (the window still on the folder of
      // the mailbox it was showing a moment ago) is an empty list, not an error.
      const source = sourceOf(accountId);
      if (!source.folders().some((f) => f.path === folder)) return { total: 0, rows: [] };
      const page = source.list(folder, { offset, limit });
      return { total: page.total, rows: page.rows.map((r) => ({ ...r, accountId, folder })) };
    },

    message: ({ accountId, folder, id }) => {
      const message = sourceOf(accountId).message(folder, id);
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

    insight: ({ accountId, folder, id }) => {
      const message = sourceOf(accountId).message(folder, id);
      return message ? insightFor(message) : null;
    },

    attachment: ({ accountId, folder, id, index: at }) => {
      const message = sourceOf(accountId).message(folder, id);
      const a = message?.attachments?.[at];
      if (!a?.bytes) return null;
      const bytes = Buffer.from(a.bytes);
      const name = attachmentFileName(a.filename);
      return { ...holdBlob(bytes, a.type, name, { group: 'mail-attachment', generation: `${accountId}|${folder}|${id}|${at}` }), name, type: a.type, size: bytes.length };
    },

    attachmentText: ({ accountId, folder, id, index: at }) => {
      const message = sourceOf(accountId).message(folder, id);
      const a = message?.attachments?.[at];
      if (!a?.bytes) return null;
      return { name: a.filename, type: a.type, text: Buffer.from(a.bytes).toString('utf8') };
    },

    readOnly: () => {
      throw new Error(READ_ONLY);
    },

    /**
     * One search over every account and every opened archive (or the
     * `sources` named), newest first, each row saying where it lives.
     */
    searchAll: async ({ query, limit = 500, offset = 0, sources = null }) => {
      await syncAccounts();
      const all = [...accounts().map((a) => a.id), ...listed().map((a) => a.id)];
      const wanted = sources?.length ? all.filter((id) => sources.includes(id)) : all;
      const found = store.index.search(query, { limit, offset, sources: wanted });
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
          list = store.index.search(query, { limit: 1e9, sources: wanted }).rows.map((r) => itemOf(r.accountId, r.folder, r));
        } else if (folders?.length) {
          for (const { accountId, folder } of folders) {
            const rows = isArchive(accountId) ? desk.messages({ accountId, folder, limit: 1e9 }).rows : store.list(accountId, folder, { limit: 1e9 }).rows;
            for (const r of rows) list.push(itemOf(accountId, folder, r));
          }
        }
        const loadStore = storeLoader(store);
        const loaders = new Map();
        const load = (item) => {
          if (!isArchive(item.sourceId)) return loadStore(item);
          if (!loaders.has(item.sourceId)) loaders.set(item.sourceId, archiveLoader(sourceOf(item.sourceId)));
          return loaders.get(item.sourceId)(item);
        };
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
  };
  return desk;
}
