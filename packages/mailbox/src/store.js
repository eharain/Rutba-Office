// The local message store.
//
// Files, not a database. A mail client's store has to survive being copied to
// another machine, being backed up by something that knows nothing about it,
// and being opened next year by a version of the software that has moved on —
// and it must never lose a message because an index got out of step with a
// table. Plain files with an append-only index give all three.
//
//   <root>/<account>/folders.json          the folder list
//   <root>/<account>/<folder>/index.jsonl  one header per line, append-only
//   <root>/<account>/<folder>/m/<id>.json  the message
//   <root>/<account>/<folder>/a/<id>/…     its attachments, written on demand
//   <root>/<account>/<folder>/sync.json    where the folder stands with its server
//   <root>/<account>/changes.json          changes made here, waiting for the server
//   <root>/<account>/outgoing/<n>.eml      a sent message waiting to be kept in Sent there
//
// The index is the only thing read to draw a list, so a 50,000-message folder
// costs one file read. The message bodies are read when a message is opened.
//
// The last three came with Mail's changes reaching the server (2.19). A store
// from before has none of them and reads as it always did: a folder with no
// sync.json is one the server has not been matched against yet, and its first
// fetch does that (main/mail-sync.js), keeping every message it holds.
//
// Sealed (2.19, D-2026-10-26-02): in a profile with a password, every file
// here is sealed with the profile's key (sealed.js): the messages, their
// attachments, the folder list, each index line, sync.json, changes.json
// and outgoing/. Nothing is read from here but through this class, which
// opens them; an attachment leaves it only as bytes (attachmentBytes), never
// as a path to the sealed file. Without a password nothing is sealed and the
// files are as they always were. When the key changes every file is sealed
// again (rekey), a folder at a time, with
//
//   <root>/.reseal.json   the run under way: the folders done, and the key
//                         it replaces sealed with the new one
//
// so a run stopped by a crash carries on at the next open (resumeReseal),
// and no file is left sealed with a key nobody can make again.

import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { SearchIndex, docOf, headlineText } from './search-index.js';
import { stripHtml } from './mime.js';
import { sealBytes, openBytes, isSealed, sealLine, openLine, isSealedLine, sealedWith } from './sealed.js';

/** The run of sealing again under way, at the store's root. */
const RESEAL = '.reseal.json';
/** How many files are sealed again between looks at everything else waiting. */
const RESEAL_SLICE = 40;

/** A file written whole or not at all: a crash mid-write leaves the one before. */
function writeAtomic(file, data) {
  const temp = `${file}.${process.pid}.tmp`;
  fs.writeFileSync(temp, data);
  fs.renameSync(temp, file);
}

// Anything a file system anywhere objects to, plus whitespace, so a stored
// name is safe to type, safe to script over, and identical on every platform.
const UNSAFE = new RegExp('[<>:\\/\\\\|?*\\u0000-\\u001f\\s-]', 'g');

const safe = (name) =>
  String(name || 'Folder')
    .replace(UNSAFE, '_')
    .replace(/\.+$/, '')
    .slice(0, 120) || 'Folder';

/** A stable id for a message, so importing the same archive twice does not double it. */
export function messageKey(message) {
  if (message.messageId) return crypto.createHash('sha1').update(message.messageId).digest('hex').slice(0, 24);
  const basis = [
    message.subject || '',
    message.date || '',
    message.from?.[0]?.address || '',
    String(message.text || message.html || '').slice(0, 512),
  ].join('\u0000');
  return crypto.createHash('sha1').update(basis).digest('hex').slice(0, 24);
}

function headerOf(message, id, folder) {
  return {
    id,
    folder,
    uid: message.uid ?? null,
    subject: message.subject || '',
    from: message.from?.[0] || null,
    to: (message.to || []).slice(0, 8),
    date: message.date || null,
    preview: message.preview || (message.text || '').replace(/\s+/g, ' ').slice(0, 220),
    unread: Boolean(message.unread),
    flagged: Boolean(message.flagged),
    hasAttachments: Boolean(message.hasAttachments || message.attachments?.length),
    attachmentCount: message.attachments?.length || 0,
    // Kept in the index because conversation grouping happens over the list,
    // and reading 10,000 message files to find out what answers what is not a
    // thing a list can afford to do. Truncated: a reference chain fifty deep
    // adds nothing once the first few have matched.
    messageId: message.messageId || null,
    inReplyTo: message.inReplyTo || null,
    references: (message.references || []).slice(-6),
    listId: message.listId || null,
    size: message.size || 0,
    source: message.source || 'imap',
  };
}

/* ── search ──────────────────────────────────────────────────────────────── */

/** A message's body as plain text, for its words. */
const bodyOf = (m) => m?.text || (m?.html ? stripHtml(m.html) : '') || '';

/** The search index's key for one folder of one account. */
export const segmentKey = (accountId, folder) => `${accountId}\u0000${folder}`;

/** How many folders' parsed indexes are kept at once. */
const CACHED_FOLDERS = 64;

export class MailStore {
  /**
   * Parsed folder indexes, keyed by account and folder, each stamped with the
   * file's size and modification time. See `#index`. At most `CACHED_FOLDERS`
   * are kept, the least lately read let go first: an imported .pst of
   * hundreds of folders held every one of them parsed for the life of the app.
   */
  #cache = new Map();

  /**
   * The search index: a segment per folder of every account, and of every
   * archive opened in place, which puts its segments here too. See search().
   */
  index = new SearchIndex();

  /** Each message's words by folder and id, so a changed folder is indexed again without reading every message. */
  #words = new Map();

  /** The profile's key files are sealed with now, or null for none; and the one it replaced while a rekey runs. */
  #key = null;
  #previous = [];
  /** Runs of sealing again, one at a time: Mail's start carrying one on and the vault finishing a change must not undo each other's keys. */
  #runs = Promise.resolve();

  /**
   * @param {string} root directory the store owns
   * @param {{ key?: Buffer|null }} [options] the profile's key, for a profile with a password
   */
  constructor(root, { key = null } = {}) {
    this.root = root;
    fs.mkdirSync(root, { recursive: true });
    this.useKey(key);
  }

  /* ── sealed with the profile's key ───────────────────────────────────── */

  /**
   * The key files are sealed with from now on (null: plain). `also`: keys
   * files may still be sealed with, while a password change a crash stopped
   * is finished (the profile's vault holds them; a password taken away
   * leaves none in the marker).
   */
  useKey(key, also = []) {
    this.#key = key ? Buffer.from(key) : null;
    this.#previous = (also || []).filter(Boolean).map((k) => Buffer.from(k));
    // A run a crash stopped: the files it had not reached open with the key it was replacing, until resumeReseal ends it.
    const marker = this.#key ? this.#readMarker() : null;
    if (marker?.from) {
      try {
        this.#previous.push(openBytes([this.#key], Buffer.from(marker.from, 'base64')));
      } catch {
        /* not this key's run */
      }
    }
    this.#cache.clear();
  }

  /** Every key that opens a file here now: the current one, and while a rekey runs the one it replaces. */
  keys() {
    return [this.#key, ...this.#previous].filter(Boolean);
  }

  /** The keys as `keys()` gave them, in another store on the same folder (the search thread's). */
  useKeys(keys = []) {
    const same = keys.length === this.keys().length && keys.every((k, i) => Buffer.from(k).equals(this.keys()[i]));
    if (same) return;
    this.#key = keys[0] ? Buffer.from(keys[0]) : null;
    this.#previous = keys.slice(1).map((k) => Buffer.from(k));
    this.#cache.clear();
  }

  /** Whether files written now are sealed. */
  get sealed() {
    return Boolean(this.#key);
  }

  #seal(bytes) {
    const b = Buffer.isBuffer(bytes) ? bytes : Buffer.from(String(bytes), 'utf8');
    return this.#key ? sealBytes(this.#key, b) : b;
  }

  #open(raw) {
    return isSealed(raw) ? openBytes(this.keys(), raw) : raw;
  }

  /** A file's bytes as they were before they were sealed. Throws as fs does for a missing file, LockedError for one no key opens. */
  #readBytes(file) {
    return this.#open(fs.readFileSync(file));
  }

  #readText(file) {
    return this.#readBytes(file).toString('utf8');
  }

  #write(file, data) {
    fs.writeFileSync(file, this.#seal(data));
  }

  /** One index line as written: sealed alone, so the next is an append. */
  #line(text) {
    return this.#key ? sealLine(this.#key, text) : text;
  }

  #parseLine(line) {
    return JSON.parse(isSealedLine(line) ? openLine(this.keys(), line) : line);
  }

  /** A folder's index written whole. */
  #writeRows(file, rows) {
    fs.writeFileSync(file, rows.map((r) => this.#line(JSON.stringify(r))).join('\n') + (rows.length ? '\n' : ''));
  }

  #appendRows(file, rows) {
    if (rows.length) fs.appendFileSync(file, `${rows.map((r) => this.#line(JSON.stringify(r))).join('\n')}\n`);
  }

  #accountDir(accountId) {
    const dir = path.join(this.root, safe(accountId));
    fs.mkdirSync(dir, { recursive: true });
    return dir;
  }

  #folderDir(accountId, folder) {
    const dir = path.join(this.#accountDir(accountId), safe(folder));
    fs.mkdirSync(path.join(dir, 'm'), { recursive: true });
    return dir;
  }

  /** The folder list, with counts kept alongside so a list draws instantly. */
  folders(accountId) {
    const file = path.join(this.#accountDir(accountId), 'folders.json');
    try {
      return JSON.parse(this.#readText(file));
    } catch {
      return [];
    }
  }

  setFolders(accountId, folders) {
    const file = path.join(this.#accountDir(accountId), 'folders.json');
    this.#writeWhole(file, JSON.stringify(folders, null, 2));
    return folders;
  }

  upsertFolder(accountId, folder) {
    const folders = this.folders(accountId);
    const at = folders.findIndex((f) => f.path === folder.path);
    if (at >= 0) folders[at] = { ...folders[at], ...folder };
    else folders.push(folder);
    return this.setFolders(accountId, folders);
  }

  /** Append one message. Returns its id, or null if it was already here. */
  put(accountId, folder, message, { force = false } = {}) {
    const dir = this.#folderDir(accountId, folder);
    const id = messageKey(message);
    const file = path.join(dir, 'm', `${id}.json`);
    const existed = fs.existsSync(file);
    if (!force && existed) return null;

    const attachments = (message.attachments || []).map((a, i) => {
      if (!a.bytes) {
        // A message put here from another folder — moved to Archive or
        // Junk, filed by a rule — brings its stored files with it: the
        // record alone would keep their names and lose the bytes once the
        // folder it came from lets go of them.
        const from = typeof a.stored === 'string' && message.folder != null && message.id
          ? path.join(this.#folderDir(accountId, message.folder), 'a', message.id, a.stored)
          : null;
        if (from && fs.existsSync(from)) {
          const adir = path.join(dir, 'a', id);
          fs.mkdirSync(adir, { recursive: true });
          const target = path.join(adir, a.stored);
          // Read and written again, not copied: sealed with the key held now, whatever the one it came with.
          if (path.resolve(from) !== path.resolve(target)) this.#write(target, this.#readBytes(from));
          return { ...a, bytes: undefined };
        }
        return { ...a, bytes: undefined, stored: false };
      }
      const adir = path.join(dir, 'a', id);
      fs.mkdirSync(adir, { recursive: true });
      const name = `${String(i).padStart(2, '0')}-${safe(a.filename)}`;
      this.#write(path.join(adir, name), Buffer.from(a.bytes));
      return { filename: a.filename, type: a.type, size: a.size ?? a.bytes.length, contentId: a.contentId, inline: a.inline, stored: name };
    });

    const record = { ...message, id, folder, attachments };
    delete record.structure; // the parse tree is not worth storing
    this.#keepOriginal(record, dir, id, message, accountId);
    this.#write(file, JSON.stringify(record));
    if (existed) {
      // Put again over itself (a draft saved again): its one index row is replaced, not joined by a second.
      const rows = this.#index(accountId, folder).filter((row) => row.id !== id);
      rows.push(headerOf(record, id, folder));
      this.#writeRows(path.join(dir, 'index.jsonl'), rows);
    } else {
      this.#appendRows(path.join(dir, 'index.jsonl'), [headerOf(record, id, folder)]);
    }
    return id;
  }

  /**
   * A signed or encrypted message's bytes as it arrived, kept beside its
   * attachments: its signature is over them and the record cannot give them
   * back. A message moved here from another folder brings its own with it.
   */
  #keepOriginal(record, dir, id, message, accountId) {
    const name = 'smime-original.eml';
    let bytes = record.smimeRaw || null;
    delete record.smimeRaw;
    if (!bytes && message.smime?.original && message.folder != null && message.id) {
      try {
        bytes = this.#readBytes(path.join(this.#folderDir(accountId, message.folder), 'a', message.id, name));
      } catch {
        bytes = null;
      }
    }
    if (!bytes) return;
    const adir = path.join(dir, 'a', id);
    fs.mkdirSync(adir, { recursive: true });
    this.#write(path.join(adir, name), Buffer.from(bytes));
    record.smime = { original: name };
  }

  /** A signed or encrypted message's bytes as it arrived, or null. */
  original(accountId, folder, id) {
    const record = this.get(accountId, folder, id);
    if (!record?.smime?.original) return null;
    try {
      return this.#readBytes(path.join(this.#folderDir(accountId, folder), 'a', id, record.smime.original));
    } catch {
      return null;
    }
  }

  /**
   * Many messages at once — one index append instead of N. `ids`, when
   * given, collects the id of each one added, so what arrived is known by
   * name rather than guessed from the newest dates.
   */
  putMany(accountId, folder, messages, { ids = null } = {}) {
    const dir = this.#folderDir(accountId, folder);
    const lines = [];
    let added = 0;
    for (const message of messages) {
      const id = messageKey(message);
      const file = path.join(dir, 'm', `${id}.json`);
      if (fs.existsSync(file)) continue;
      const attachments = (message.attachments || []).map((a, i) => {
        if (!a.bytes) return { filename: a.filename, type: a.type, size: a.size || 0, contentId: a.contentId, inline: a.inline, stored: false };
        const adir = path.join(dir, 'a', id);
        fs.mkdirSync(adir, { recursive: true });
        const name = `${String(i).padStart(2, '0')}-${safe(a.filename)}`;
        this.#write(path.join(adir, name), Buffer.from(a.bytes));
        return { filename: a.filename, type: a.type, size: a.size ?? a.bytes.length, contentId: a.contentId, inline: a.inline, stored: name };
      });
      const record = { ...message, id, folder, attachments };
      delete record.structure;
      this.#keepOriginal(record, dir, id, message, accountId);
      this.#write(file, JSON.stringify(record));
      lines.push(headerOf(record, id, folder));
      ids?.push(id);
      added++;
    }
    this.#appendRows(path.join(dir, 'index.jsonl'), lines);
    return added;
  }

  /**
   * A folder's index, parsed once.
   *
   * Every list, every search, every count read the file and parsed it again.
   * That is fine for one folder and quadratic for a mailbox: searching an
   * imported .pst with 467 folders re-parsed 12,000 lines of JSON per folder,
   * per keystroke. The cache is keyed on the file's size and modification time,
   * so a folder written by anything — this process, a sync, a restore from
   * backup — invalidates itself without being told.
   */
  #index(accountId, folder) {
    const file = path.join(this.#folderDir(accountId, folder), 'index.jsonl');
    let stat = null;
    try {
      stat = fs.statSync(file);
    } catch {
      return [];
    }

    const key = `${accountId}\u0000${folder}`;
    const stamp = `${stat.size}:${stat.mtimeMs}`;
    const cached = this.#cache.get(key);
    if (cached && cached.stamp === stamp) {
      // Read again: the most lately read, so the last to be let go.
      this.#cache.delete(key);
      this.#cache.set(key, cached);
      return cached.rows;
    }

    const rows = [];
    for (const line of fs.readFileSync(file, 'utf8').split('\n')) {
      if (!line.trim()) continue;
      try {
        rows.push(this.#parseLine(line));
      } catch {
        // A torn last line after a crash: skip it, keep the rest.
      }
    }

    this.#cache.delete(key);
    this.#cache.set(key, { stamp, rows });
    while (this.#cache.size > CACHED_FOLDERS) this.#cache.delete(this.#cache.keys().next().value);
    return rows;
  }

  /**
   * A page of a folder's message list, newest first.
   * @param {{ offset?: number, limit?: number, query?: string, unreadOnly?: boolean }} [opts]
   */
  list(accountId, folder, opts = {}) {
    const { offset = 0, limit = 100, query = '', unreadOnly = false } = opts;
    let rows = this.#index(accountId, folder);
    if (unreadOnly) rows = rows.filter((r) => r.unread);
    if (query) {
      const q = query.toLowerCase();
      rows = rows.filter(
        (r) =>
          r.subject.toLowerCase().includes(q) ||
          (r.preview || '').toLowerCase().includes(q) ||
          (r.from?.address || '').toLowerCase().includes(q) ||
          (r.from?.name || '').toLowerCase().includes(q)
      );
    }
    rows.sort((a, b) => (b.date || '').localeCompare(a.date || ''));
    return { total: rows.length, rows: rows.slice(offset, offset + limit) };
  }

  /** One message in full. */
  /** The index rows of the messages named, in the folder's order. */
  headers(accountId, folder, ids) {
    const wanted = new Set(ids || []);
    return this.#index(accountId, folder).filter((r) => wanted.has(r.id));
  }

  get(accountId, folder, id) {
    const file = path.join(this.#folderDir(accountId, folder), 'm', `${id}.json`);
    try {
      return JSON.parse(this.#readText(file));
    } catch {
      return null;
    }
  }

  /** Where an attachment's file is: sealed in a profile with a password, so read it with attachmentBytes. */
  attachmentPath(accountId, folder, id, stored) {
    return path.join(this.#folderDir(accountId, folder), 'a', id, stored);
  }

  /** An attachment's bytes, opened. Throws when it is not stored. */
  attachmentBytes(accountId, folder, id, stored) {
    return this.#readBytes(this.attachmentPath(accountId, folder, id, stored));
  }

  /** Change flags on a message; the index line is rewritten in place. */
  setFlags(accountId, folder, id, patch) {
    return this.setFlagsMany(accountId, folder, [id], patch)[0] ?? null;
  }

  /**
   * Change flags on several messages with ONE rewrite of the folder's index.
   * A loop of `setFlags` rewrote and re-parsed the whole index per message —
   * 20 s to mark 2,000 messages read, and quadratic from there — because each
   * write invalidates the cache the next call reads through.
   * @returns {object[]} the messages that were found and changed
   */
  setFlagsMany(accountId, folder, ids, patch) {
    const dir = this.#folderDir(accountId, folder);
    const changed = new Set();
    const messages = [];
    for (const id of ids) {
      const message = this.get(accountId, folder, id);
      if (!message) continue;
      Object.assign(message, patch);
      this.#write(path.join(dir, 'm', `${id}.json`), JSON.stringify(message));
      changed.add(id);
      messages.push(message);
    }
    if (changed.size) {
      const rows = this.#index(accountId, folder).map((r) => (changed.has(r.id) ? { ...r, ...patch } : r));
      this.#writeRows(path.join(dir, 'index.jsonl'), rows);
    }
    return messages;
  }

  /**
   * Keep what was read from one of a message's attachments (an invoice's
   * details) with the message, so it is not read again: in the message's own
   * record, which a move to another folder carries with it. The folder's
   * index is not touched; nothing in a list shows it.
   * @returns {boolean} whether the message and the attachment were there
   */
  setAttachmentReading(accountId, folder, id, index, reading) {
    const message = this.get(accountId, folder, id);
    const a = message?.attachments?.[index];
    if (!a) return false;
    if (reading === null || reading === undefined) delete a.reading;
    else a.reading = reading;
    this.#write(path.join(this.#folderDir(accountId, folder), 'm', `${id}.json`), JSON.stringify(message));
    return true;
  }

  remove(accountId, folder, id) {
    return this.removeMany(accountId, folder, [id]);
  }

  /** Remove several messages with one rewrite of the index; see `setFlagsMany`. */
  removeMany(accountId, folder, ids) {
    const dir = this.#folderDir(accountId, folder);
    const gone = new Set(ids);
    for (const id of gone) {
      try {
        fs.rmSync(path.join(dir, 'm', `${id}.json`), { force: true });
        fs.rmSync(path.join(dir, 'a', id), { recursive: true, force: true });
      } catch {
        /* already gone */
      }
    }
    const rows = this.#index(accountId, folder).filter((r) => !gone.has(r.id));
    this.#writeRows(path.join(dir, 'index.jsonl'), rows);
    return true;
  }

  /**
   * Search every folder of an account.
   *
   * Backed by an inverted index rather than a scan. The scan version compared
   * every query against every field of every message in every folder, which is
   * fine at ten thousand messages and unusable at the two hundred thousand a
   * decade of imported .pst actually contains.
   *
   * The index (search-index.js) is built on first use and kept in memory, a
   * segment per folder, each stamped with its folder's index file: a folder
   * that changed is indexed again, and only that folder. It is not written to
   * disk on purpose: an index file that can disagree with the messages is a
   * class of bug this store was designed to avoid - the folder indexes are
   * the truth, and anything derived from them has to be cheap enough to throw
   * away. Message bodies are indexed too; each message's words are kept by
   * id, so indexing a folder again reads only the messages new to it.
   *
   * The syntax is query.js's: from:, to:, cc:, subject:, filename:,
   * has:attachment, is:unread, is:starred, after:, before:, size>, quoted
   * phrases, a minus to leave out and OR.
   */
  search(accountId, query, { limit = 200 } = {}) {
    this.syncIndex(accountId);
    return this.index.search(query, { limit, sources: [accountId] }).rows;
  }

  /** Bring an account's segments of the index up to date, at once. */
  syncIndex(accountId, info = {}) {
    const steps = this.indexSteps(accountId, info);
    let r = steps.next();
    while (!r.done) r = steps.next();
    return r.value;
  }

  /**
   * The same, a message at a time: each `yield` is one message read from
   * disk, so a caller can run it in slices (archive.js `runSliced`).
   */
  *indexSteps(accountId, { name = null } = {}) {
    const source = { id: accountId, name: name || accountId, kind: 'account' };
    const folders = this.folders(accountId);
    const live = new Set();
    let read = 0;
    for (const folder of folders) {
      const key = segmentKey(accountId, folder.path);
      live.add(key);
      let stamp;
      try {
        const st = fs.statSync(path.join(this.#folderDir(accountId, folder.path), 'index.jsonl'));
        stamp = `${st.size}:${st.mtimeMs}`;
      } catch {
        stamp = '0';
      }
      const held = this.index.segments.get(key);
      if (held && held.stamp === stamp) {
        held.source = source;
        continue;
      }
      const known = this.#words.get(key) || new Map();
      const kept = new Map();
      const builder = this.index.builder({
        key,
        stamp,
        source,
        folder: { path: folder.path, name: folder.name || folder.path },
        textOf: (doc) => bodyOf(this.get(accountId, folder.path, doc.id)),
      });
      for (const row of this.#index(accountId, folder.path)) {
        let entry = known.get(row.id);
        if (!entry) {
          const m = this.get(accountId, folder.path, row.id);
          const extra = m ? { to: m.to, cc: m.cc, bcc: m.bcc, attachments: (m.attachments || []).map((a) => a.filename).filter(Boolean) } : {};
          const doc = docOf({ ...row, ...extra }, row.id);
          const { words, sig } = this.index.analyse(headlineText(doc), m ? bodyOf(m) : '');
          entry = { extra, ids: this.index.ids(words), sig };
          read++;
          builder.add(doc, entry.ids, sig);
          kept.set(row.id, entry);
          yield { read };
          continue;
        }
        builder.add(docOf({ ...row, ...entry.extra }, row.id), entry.ids, entry.sig);
        kept.set(row.id, entry);
      }
      builder.finish();
      this.#words.set(key, kept);
    }
    // Folders that are gone take their segments with them.
    const prefix = segmentKey(accountId, '');
    for (const key of [...this.index.segments.keys()]) {
      if (key.startsWith(prefix) && !live.has(key)) {
        this.index.drop(key);
        this.#words.delete(key);
      }
    }
    return { read };
  }

  /** Forget an account's part of the index - the account was removed. */
  dropIndex(accountId) {
    const prefix = segmentKey(accountId, '');
    for (const key of [...this.index.segments.keys()]) {
      if (key.startsWith(prefix)) {
        this.index.drop(key);
        this.#words.delete(key);
      }
    }
  }

  counts(accountId, folder) {
    const rows = this.#index(accountId, folder);
    return { total: rows.length, unread: rows.filter((r) => r.unread).length };
  }

  /* ── where each folder stands with its server ─────────────────────────── */

  /** A file written whole or not at all: a crash mid-write leaves the one before. */
  #writeWhole(file, text) {
    writeAtomic(file, this.#seal(text));
  }

  /** Does the store hold this folder (its directory), without making it? */
  hasFolder(accountId, folder) {
    return fs.existsSync(path.join(this.root, safe(accountId), safe(folder)));
  }

  /**
   * A folder's place at its server: { uidValidity, uidNext, highestModseq,
   * lastSeen, messages }, the first three as the server last said them and
   * `lastSeen` the highest UID fetched. Null for a folder never matched
   * against its server: a store from before 2.19, or a folder never fetched.
   */
  syncState(accountId, folder) {
    try {
      return JSON.parse(this.#readText(path.join(this.#folderDir(accountId, folder), 'sync.json')));
    } catch {
      return null;
    }
  }

  setSyncState(accountId, folder, state) {
    const file = path.join(this.#folderDir(accountId, folder), 'sync.json');
    if (state === null) fs.rmSync(file, { force: true });
    else this.#writeWhole(file, JSON.stringify(state));
    return state;
  }

  /** The index row of the message with this UID, or null. */
  byUid(accountId, folder, uid) {
    return this.#index(accountId, folder).find((r) => r.uid === uid) || null;
  }

  /**
   * Each message's UID at its server, by id ({ id: uid }, null for none), in
   * one write of the index. A message moved here keeps the UID its old folder
   * gave it until the server says the new one; this is how it is told.
   */
  setUids(accountId, folder, uids) {
    const dir = this.#folderDir(accountId, folder);
    const ids = Object.keys(uids);
    let changed = 0;
    for (const id of ids) {
      const message = this.get(accountId, folder, id);
      if (!message || message.uid === uids[id]) continue;
      message.uid = uids[id];
      this.#write(path.join(dir, 'm', `${id}.json`), JSON.stringify(message));
      changed++;
    }
    if (changed) {
      const rows = this.#index(accountId, folder).map((r) => (r.id in uids ? { ...r, uid: uids[r.id] } : r));
      this.#writeRows(path.join(dir, 'index.jsonl'), rows);
    }
    return changed;
  }

  /**
   * Changes made here that the server has not heard yet, oldest first: what
   * main/mail-sync.js replays, in order, each taken off once the server has
   * it. Kept beside the account's folders so it travels with them.
   */
  changes(accountId) {
    try {
      const kept = JSON.parse(this.#readText(path.join(this.#accountDir(accountId), 'changes.json')));
      return Array.isArray(kept.items) ? kept : { next: 1, items: [] };
    } catch {
      return { next: 1, items: [] };
    }
  }

  /** One more change at the end of the line: it is on disk before this returns, with its number. */
  addChange(accountId, change) {
    const kept = this.changes(accountId);
    const item = { ...change, n: kept.next, at: change.at || new Date().toISOString() };
    this.#writeWhole(path.join(this.#accountDir(accountId), 'changes.json'), JSON.stringify({ next: kept.next + 1, items: [...kept.items, item] }));
    return item;
  }

  /** A change the server now has (or has refused), taken off the line; `patch` keeps it there, changed. */
  settleChange(accountId, n, patch = null) {
    const kept = this.changes(accountId);
    const items = patch ? kept.items.map((c) => (c.n === n ? { ...c, ...patch } : c)) : kept.items.filter((c) => c.n !== n);
    this.#writeWhole(path.join(this.#accountDir(accountId), 'changes.json'), JSON.stringify({ next: kept.next, items }));
    return items.length;
  }

  /** Where a sent message's bytes wait to be kept in Sent at the server. */
  outgoingPath(accountId, name) {
    const dir = path.join(this.#accountDir(accountId), 'outgoing');
    fs.mkdirSync(dir, { recursive: true });
    return path.join(dir, safe(name));
  }

  /** A sent message's bytes kept to wait for the server, sealed like the rest: the file they are in. */
  writeOutgoing(accountId, name, bytes) {
    const file = this.outgoingPath(accountId, name);
    writeAtomic(file, this.#seal(Buffer.from(bytes)));
    return file;
  }

  /** What writeOutgoing kept, opened. */
  readOutgoing(file) {
    return this.#readBytes(file);
  }

  /**
   * A folder renamed here, with the folders inside it (`delimiter` between
   * the names): their files moved, each message told its new folder, the
   * folder list changed. The server is told separately.
   */
  renameFolder(accountId, from, to, delimiter = '/') {
    const folders = this.folders(accountId);
    const inside = (p) => p === from || p.startsWith(`${from}${delimiter}`);
    for (const f of folders.filter((x) => inside(x.path)).sort((a, b) => b.path.length - a.path.length)) {
      const next = to + f.path.slice(from.length);
      const oldDir = path.join(this.#accountDir(accountId), safe(f.path));
      const newDir = path.join(this.#accountDir(accountId), safe(next));
      if (fs.existsSync(oldDir) && !fs.existsSync(newDir)) fs.renameSync(oldDir, newDir);
      this.#cache.delete(`${accountId}\u0000${f.path}`);
      // Each message and index row says which folder it is in.
      const rows = this.#index(accountId, next).map((r) => ({ ...r, folder: next }));
      for (const r of rows) {
        const m = this.get(accountId, next, r.id);
        if (m) this.#write(path.join(newDir, 'm', `${r.id}.json`), JSON.stringify({ ...m, folder: next }));
      }
      if (rows.length) this.#writeRows(path.join(this.#folderDir(accountId, next), 'index.jsonl'), rows);
    }
    this.setFolders(accountId, folders.map((f) => (inside(f.path) ? { ...f, path: to + f.path.slice(from.length), name: (to + f.path.slice(from.length)).split(delimiter).pop() } : f)));
    return true;
  }

  /** A folder gone, with everything in it, and out of the folder list. */
  removeFolder(accountId, folder) {
    fs.rmSync(path.join(this.#accountDir(accountId), safe(folder)), { recursive: true, force: true });
    this.#cache.delete(`${accountId}\u0000${folder}`);
    this.setFolders(accountId, this.folders(accountId).filter((f) => f.path !== folder));
    return true;
  }

  /* ── sealing every file again when the key changes ───────────────────── */

  #marker() {
    return path.join(this.root, RESEAL);
  }

  #readMarker() {
    try {
      const m = JSON.parse(fs.readFileSync(this.#marker(), 'utf8'));
      return { from: typeof m.from === 'string' ? m.from : null, done: Array.isArray(m.done) ? m.done : [] };
    } catch {
      return null;
    }
  }

  /**
   * Before the profile's lock changes: the run about to start written down,
   * with the key being replaced sealed with the new one (never on its own,
   * and not at all when a password is set or taken away: then one of the two
   * is no key). Should anything stop between the lock changing and the last
   * file, the next open has what it needs to carry on.
   */
  prepareReseal(old, next) {
    return this.#oneAtATime(() => this.#prepare(old, next));
  }

  #oneAtATime(step) {
    const run = this.#runs.catch(() => {}).then(step);
    this.#runs = run;
    return run;
  }

  async #prepare(old, next) {
    // A run a crash stopped is finished first: its marker holds the only copy of the key before it.
    if (this.#readMarker()?.done.length) await this.#resume();
    const from = old && next ? sealBytes(Buffer.from(next), Buffer.from(old)).toString('base64') : null;
    writeAtomic(this.#marker(), JSON.stringify({ v: 1, from, done: [] }));
  }

  /** A run of sealing again begun and not finished. */
  resealPending() {
    return fs.existsSync(this.#marker());
  }

  /**
   * Every file sealed again with `next` (or made plain, with none), the
   * files it replaces opened with `old`. Reads and writes go on meanwhile:
   * a file is read with either key, a file written is sealed with the new.
   * @param {{ progress?: (p: { done: number, total: number }) => void }} [options]
   * @returns {Promise<{ files: number, sealed: number, unopened: number, ms: number }>}
   */
  rekey(old, next, { progress } = {}) {
    return this.#oneAtATime(() => this.#rekey(old, next, progress));
  }

  async #rekey(old, next, progress) {
    // Not prepared for (a vault used without prepare), or prepared for another run: written now.
    const marker = this.#readMarker();
    if (!marker || marker.done.length || (old && next && !marker.from)) await this.#prepare(old, next);
    this.#key = next ? Buffer.from(next) : null;
    this.#previous = old ? [Buffer.from(old)] : [];
    this.#cache.clear();
    return this.#resealAll(progress);
  }

  /**
   * A run a crash stopped, carried on: the key it replaced read from the
   * marker (sealed with the key held now), and every folder not yet done
   * sealed with the key held now. Null when there is none to carry on.
   */
  resumeReseal({ progress } = {}) {
    return this.#oneAtATime(() => this.#resume(progress));
  }

  async #resume(progress) {
    const marker = this.#readMarker();
    if (!marker) return null;
    let previous = null;
    if (marker.from && this.#key) {
      try {
        previous = openBytes([this.#key], Buffer.from(marker.from, 'base64'));
      } catch {
        // Sealed with a key this lock does not make: the lock never changed, and the files are still its.
        previous = null;
      }
    }
    // Added to the keys held, not in place of them: the vault may hold the one that opens the rest (a password taken away leaves none in the marker).
    if (previous && !this.#previous.some((k) => k.equals(previous))) this.#previous.push(previous);
    this.#cache.clear();
    return this.#resealAll(progress);
  }

  /** The units a run goes through: each account's own files, and each of its folders. */
  #units() {
    const units = [];
    let accounts = [];
    try {
      accounts = fs.readdirSync(this.root, { withFileTypes: true }).filter((e) => e.isDirectory());
    } catch {
      return units;
    }
    for (const a of accounts) {
      const dir = path.join(this.root, a.name);
      units.push({ name: `${a.name}/`, dir, own: true });
      for (const f of fs.readdirSync(dir, { withFileTypes: true })) {
        if (f.isDirectory() && f.name !== 'outgoing') units.push({ name: `${a.name}/${f.name}`, dir: path.join(dir, f.name), own: false });
      }
    }
    return units;
  }

  /** The files of one unit, each with how it is sealed. */
  #filesOf(unit) {
    const out = [];
    const add = (file, kind) => out.push({ file, kind });
    const list = (dir) => {
      try {
        return fs.readdirSync(dir, { withFileTypes: true });
      } catch {
        return [];
      }
    };
    // Not a write's own temporary file (writeAtomic names them <file>.<pid>.tmp).
    const plainFile = (n) => !/\.\d+\.tmp$/.test(n);
    if (unit.own) {
      for (const e of list(unit.dir)) if (e.isFile() && plainFile(e.name)) add(path.join(unit.dir, e.name), 'whole');
      for (const e of list(path.join(unit.dir, 'outgoing'))) if (e.isFile() && plainFile(e.name)) add(path.join(unit.dir, 'outgoing', e.name), 'whole');
      return out;
    }
    for (const e of list(unit.dir)) {
      if (!e.isFile() || !plainFile(e.name)) continue;
      add(path.join(unit.dir, e.name), e.name === 'index.jsonl' ? 'lines' : 'whole');
    }
    for (const e of list(path.join(unit.dir, 'm'))) if (e.isFile() && plainFile(e.name)) add(path.join(unit.dir, 'm', e.name), 'whole');
    for (const d of list(path.join(unit.dir, 'a'))) {
      if (!d.isDirectory()) continue;
      for (const e of list(path.join(unit.dir, 'a', d.name))) if (e.isFile() && plainFile(e.name)) add(path.join(unit.dir, 'a', d.name, e.name), 'whole');
    }
    return out;
  }

  /**
   * One file sealed with the key held now, written whole or not at all; a
   * file sealed with it already is left. Synchronous, so nothing written to
   * the same file meanwhile can be lost between its read and its write.
   * @returns {'sealed'|'kept'|'unopened'|'gone'}
   */
  #resealFile({ file, kind }) {
    let raw;
    try {
      raw = fs.readFileSync(file);
    } catch {
      return 'gone';
    }
    const keys = [this.#key, ...this.#previous];
    if (kind === 'lines') {
      const lines = raw.toString('utf8').split('\n');
      let changed = false;
      const out = [];
      for (const line of lines) {
        if (!line.trim()) {
          out.push(line);
          continue;
        }
        const by = sealedWith(keys, line);
        if (by === -1) return 'unopened';
        if ((by === 0 && this.#key) || (by === null && !this.#key)) {
          out.push(line);
          continue;
        }
        const text = by === null ? line : openLine(keys, line);
        out.push(this.#line(text));
        changed = true;
      }
      if (!changed) return 'kept';
      writeAtomic(file, out.join('\n'));
      return 'sealed';
    }
    const by = sealedWith(keys, raw);
    if (by === -1) return 'unopened';
    if ((by === 0 && this.#key) || (by === null && !this.#key)) return 'kept';
    const plain = by === null ? raw : openBytes(keys, raw);
    writeAtomic(file, this.#seal(plain));
    return 'sealed';
  }

  async #resealAll(progress) {
    const began = Date.now();
    const marker = this.#readMarker() || { from: null, done: [] };
    const done = new Set(marker.done);
    let files = 0;
    let sealed = 0;
    let unopened = 0;
    let total = 0;
    // Gone round again until a pass finds nothing new: a folder renamed or made while the run goes on is done too.
    for (;;) {
      const units = this.#units().filter((u) => !done.has(u.name));
      if (!units.length) break;
      const work = units.map((u) => ({ unit: u, files: this.#filesOf(u) }));
      total += work.reduce((n, w) => n + w.files.length, 0);
      for (const { unit, files: list } of work) {
        for (let i = 0; i < list.length; i++) {
          const r = this.#resealFile(list[i]);
          files++;
          if (r === 'sealed') sealed++;
          else if (r === 'unopened') unopened++;
          if (files % RESEAL_SLICE === 0) {
            progress?.({ done: files, total });
            // Everything else waiting (a window's question, a fetch) is answered between slices.
            await new Promise((resolve) => setImmediate(resolve));
          }
        }
        done.add(unit.name);
        marker.done = [...done];
        writeAtomic(this.#marker(), JSON.stringify({ v: 1, from: marker.from, done: marker.done }));
      }
    }
    progress?.({ done: files, total });
    fs.rmSync(this.#marker(), { force: true });
    this.#previous = [];
    this.#cache.clear();
    return { files, sealed, unopened, ms: Date.now() - began };
  }
}
