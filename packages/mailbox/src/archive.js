// An archive opened where it lies.
//
// Importing copies an archive into the local store, which is right for mail a
// person will go on working with and wrong for the ten-year .pst a lawyer has
// asked about: copying it doubles the disk it takes, takes as long as reading
// it, and leaves two things that can be told apart. So an archive can also
// be opened in place - its folders listed, its messages read and searched,
// exported - from the file itself, opened for reading only. Nothing here
// writes to it, or beside it.
//
// Nothing is read whole either. A .pst is paged from disk; an mbox is read
// once in blocks to find where its messages begin, and an .olm's zip
// directory is read from the end of the file; after that a message is read
// from the file when it is asked for. So an archive of any size opens - an
// mbox or .olm past 2 GB, more than a Node buffer holds, as well as a small
// one - and what it costs in memory is its index, not its size.
//
// One shape for every kind: an Outlook data file (.pst or .ost), an Outlook
// for Mac archive (.olm) or an mbox. Folders are named by their path as a
// person sees it ("Inbox/Projects"), and messages by an id that is stable for
// the file: the node id in a .pst, the place in an .olm or an mbox.

import fs from 'node:fs';
import path from 'node:path';
import { PstFile, HIDDEN_FOLDERS } from './pst/index.js';
import { OlmArchive, messageFrom, folderOf } from './olm.js';
import { openMbox } from './mbox.js';
import { parse, all } from '@rutba/office-formats/xml';
import { fileReader, identify } from './import.js';
import { stripHtml } from './mime.js';
import { docOf, headlineText } from './search-index.js';

/** A message's body as plain text, for the index and for phrases. */
export function bodyText(m) {
  return m?.text || (m?.html ? stripHtml(m.html) : '') || '';
}

/**
 * Past this size an mbox or .olm is listed without reading every message
 * whole first: an mbox from its messages' headers alone, an .olm from what
 * has been indexed of it so far (`listsCheaply` false). Under it, as before.
 */
export const LARGE_ARCHIVE = 256 * 1048576;

const LABELS = {
  pst: 'Outlook data file (.pst)',
  ost: 'Outlook offline file (.ost)',
  olm: 'Outlook for Mac archive (.olm)',
  mbox: 'Mbox archive',
};

/** Rows of a list, newest first, a page of them. */
function page(rows, { offset = 0, limit = 200 } = {}) {
  const sorted = rows.slice().sort((a, b) => String(b.date || '').localeCompare(String(a.date || '')));
  return { total: sorted.length, rows: sorted.slice(offset, offset + limit) };
}

class PstSource {
  constructor(filePath) {
    this.path = filePath;
    this.pst = new PstFile(fileReader(filePath));
    this.kind = this.pst.info.kind;
    this.label = LABELS[this.kind];
    this.bytes = this.pst.info.bytes;
    this.listsCheaply = true;
    this.storeName = null;
    this.#lists = new Map();
  }

  #folders = null;
  #lists;

  folders() {
    if (this.#folders) return this.#folders;
    const out = [];
    const seen = new Map();
    for (const f of this.pst.folders().list) {
      if (f.hidden && !f.messageCount) continue;
      const visible = String(f.path || f.name)
        .split('/')
        .filter((part) => !HIDDEN_FOLDERS.has(part));
      let key = visible.join('/') || f.name;
      // Two folders of one name side by side are both somebody's mail.
      const n = (seen.get(key) || 0) + 1;
      seen.set(key, n);
      if (n > 1) key = `${key} (${n})`;
      out.push({ path: key, name: visible[visible.length - 1] || f.name, depth: Math.max(0, visible.length - 1), total: f.messageCount, unread: f.unreadCount, nid: f.nid });
    }
    this.#folders = out;
    return out;
  }

  #nid(folder) {
    const f = this.folders().find((x) => x.path === folder);
    if (!f) throw new Error(`There is no folder "${folder}" in this archive.`);
    return f.nid;
  }

  list(folder, opts) {
    let rows = this.#lists.get(folder);
    if (!rows) {
      const nid = this.#nid(folder);
      rows = this.pst.messages(nid, { offset: 0, limit: 1e9 }).map((r) => ({
        id: String(r.nid),
        subject: r.subject,
        from: r.from,
        to: r.to ? [{ name: r.to, address: '' }] : [],
        date: r.date,
        size: r.size,
        unread: r.unread,
        hasAttachments: r.hasAttachments,
        preview: r.preview,
      }));
      this.#lists.set(folder, rows);
      while (this.#lists.size > 8) this.#lists.delete(this.#lists.keys().next().value);
    }
    return page(rows, opts);
  }

  count(folder) {
    return this.pst.messageNids(this.#nid(folder)).length;
  }

  *messages(folder) {
    for (const nid of this.pst.messageNids(this.#nid(folder))) {
      const m = this.pst.message(nid, { withBytes: false });
      if (m) yield { id: String(nid), message: m };
    }
  }

  message(folder, id) {
    const m = this.pst.message(Number(id));
    return m ? { ...m, id: String(id), folder } : null;
  }

  text(folder, id) {
    return bodyText(this.pst.body(Number(id)));
  }

  raw() {
    return null;
  }

  close() {
    this.pst.close();
  }
}

class OlmSource {
  constructor(filePath) {
    this.path = filePath;
    this.kind = 'olm';
    this.label = LABELS.olm;
    this.bytes = fs.statSync(filePath).size;
    this.olm = OlmArchive.open(filePath);
    this.listsCheaply = this.bytes <= LARGE_ARCHIVE;
    this.#decoder = new TextDecoder('utf-8', { fatal: false });
  }

  #decoder;
  #folders = null;
  #lists = new Map();

  /** The messages of one entry, parsed. */
  /**
   * The messages of one entry, parsed. `light` leaves the attachments' base64
   * out before the XML is parsed - their names and sizes are attributes - so
   * a list or the index never holds a megabyte of text for a picture.
   */
  #entry(index, light = false) {
    const entry = this.olm.entries[index];
    // Read into the archive's reused buffer when it is decoded at once and let go.
    let xml = this.#decoder.decode(light && this.olm.zip ? this.olm.zip.data(entry.entry, { reuse: true }) : entry.data);
    if (light) xml = xml.replace(/(<messageAttachment\b[^>]*>)[^<]*(<\/messageAttachment>)/g, '$1$2');
    return all(parse(xml), 'email');
  }

  folders() {
    if (this.#folders) return this.#folders;
    const out = new Map();
    // A large archive is counted from its directory, an entry a message, as
    // Outlook for Mac writes them; reading every entry to count them would
    // read the whole file before its folders could be shown. Its unread
    // count waits for the index.
    const counted = !this.listsCheaply;
    this.olm.entries.forEach((entry) => {
      const folder = folderOf(entry.name);
      let n = 1;
      let unread = 0;
      if (!counted) {
        const xml = this.#decoder.decode(entry.data);
        n = (xml.match(/<email\b/g) || []).length;
        unread = (xml.match(/OPFMessageIsRead="0"/g) || []).length;
      }
      if (!n) return;
      const f = out.get(folder) || { path: folder, name: folder.split('/').pop(), depth: folder.split('/').length - 1, total: 0, unread: 0 };
      f.total += n;
      f.unread += unread;
      out.set(folder, f);
    });
    this.#folders = [...out.values()];
    return this.#folders;
  }

  /** A folder's message ids from the directory alone, an entry a message: for an export of a large archive before it is indexed. */
  ids(folder) {
    const out = [];
    this.olm.entries.forEach((entry, e) => {
      if (folderOf(entry.name) === folder) out.push(`${e}.0`);
    });
    return out;
  }

  *messages(folder) {
    for (let e = 0; e < this.olm.entries.length; e++) {
      if (folderOf(this.olm.entries[e].name) !== folder) continue;
      const emails = this.#entry(e, true);
      // The attachments' names and sizes, not their bytes: a list and the index want no more.
      for (let k = 0; k < emails.length; k++) yield { id: `${e}.${k}`, message: messageFrom(emails[k], { withBytes: false }) };
    }
  }

  list(folder, opts) {
    let rows = this.#lists.get(folder);
    if (!rows) {
      rows = [];
      for (const { id, message: m } of this.messages(folder)) {
        rows.push({ id, subject: m.subject, from: m.from[0] || null, to: m.to, cc: m.cc, date: m.date, size: 0, unread: m.unread, hasAttachments: m.hasAttachments, attachmentCount: m.attachments.length, preview: m.preview, messageId: m.messageId });
      }
      this.#lists.set(folder, rows);
    }
    return page(rows, opts);
  }

  count(folder) {
    return this.folders().find((f) => f.path === folder)?.total || 0;
  }

  message(folder, id) {
    const [e, k] = String(id).split('.').map(Number);
    const el = this.olm.entries[e] ? this.#entry(e)[k] : null;
    return el ? { ...messageFrom(el), id: String(id), folder } : null;
  }

  text(folder, id) {
    return bodyText(this.message(folder, id));
  }

  raw() {
    return null;
  }

  close() {
    this.olm.close();
  }
}

class MboxSource {
  constructor(filePath) {
    this.path = filePath;
    this.kind = 'mbox';
    this.label = LABELS.mbox;
    this.mbox = openMbox(filePath);
    this.bytes = this.mbox.bytes;
    this.listsCheaply = true;
    this.folder = path.basename(filePath, path.extname(filePath)) || 'Mbox';
    this.#rows = null;
  }

  #rows;

  folders() {
    return [{ path: this.folder, name: this.folder, depth: 0, total: this.mbox.length, unread: 0 }];
  }

  *messages() {
    for (let i = 0; i < this.mbox.length; i++) {
      // Attachments listed, not decoded: what the index and a list want.
      const m = this.mbox.message(i, { light: true });
      if (m) yield { id: String(i), message: { ...m, size: this.mbox.size(i) } };
    }
  }

  list(folder, opts) {
    if (!this.#rows && this.bytes > LARGE_ARCHIVE) {
      // A large mbox is listed from its messages' headers, read a few
      // kilobytes each, not from the messages whole; the index gives the rest.
      this.#rows = [];
      for (let i = 0; i < this.mbox.length; i++) {
        const m = this.mbox.headers(i, { most: 16384 });
        if (m) this.#rows.push({ id: String(i), subject: m.subject, from: m.from[0] || null, to: m.to, cc: m.cc, date: m.date, size: this.mbox.size(i), unread: false, hasAttachments: m.hasAttachments, attachmentCount: m.hasAttachments ? 1 : 0, preview: '', messageId: m.messageId, inReplyTo: m.inReplyTo, references: m.references });
      }
    }
    if (!this.#rows) {
      this.#rows = [];
      for (const { id, message: m } of this.messages()) {
        this.#rows.push({ id, subject: m.subject, from: m.from[0] || null, to: m.to, cc: m.cc, date: m.date, size: m.size, unread: false, hasAttachments: m.hasAttachments, attachmentCount: m.attachments.length, preview: (m.text || stripHtml(m.html || '')).replace(/\s+/g, ' ').slice(0, 200), messageId: m.messageId, inReplyTo: m.inReplyTo, references: m.references });
      }
    }
    return page(this.#rows, opts);
  }

  count() {
    return this.mbox.length;
  }

  message(folder, id) {
    const m = this.mbox.message(Number(id));
    if (!m) return null;
    delete m.structure;
    return { ...m, id: String(id), folder, source: 'mbox' };
  }

  text(folder, id) {
    return bodyText(this.mbox.message(Number(id)));
  }

  /** The message exactly as the archive holds it. */
  raw(folder, id) {
    return this.mbox.raw(Number(id));
  }

  close() {
    this.mbox.close();
  }
}

/**
 * Open an archive where it lies, for reading only.
 * @param {string} filePath
 */
export function openArchive(filePath) {
  const found = identify(filePath);
  if (found.kind === 'pst') return new PstSource(filePath);
  if (found.kind === 'olm') return new OlmSource(filePath);
  if (found.kind === 'mbox') return new MboxSource(filePath);
  const err = new Error('That file is not a mail archive this can open in place (.pst, .ost, .olm or mbox).');
  err.code = 'not-archive';
  throw err;
}

/**
 * Index an opened archive into `index`, a segment per folder, a step at a
 * time: each `yield` is one message read, so a caller can run it in slices
 * and keep a window answering. The segments carry `textOf`, which reads a
 * message's body again for a phrase.
 *
 * With `part`, a folder is made searchable in parts as it goes rather than
 * when it is finished: every `part` messages, and whenever the caller calls
 * `hooks.flush()` between steps (before a search, or when indexing is
 * paused), what has been read of the folder so far becomes a segment of its
 * own, keyed by the folder's key and the part's number. A search then
 * answers from everything indexed so far.
 * @param {object} source from openArchive
 * @param {import('./search-index.js').SearchIndex} index
 * @param {{ id: string, name: string }} info the source as results name it
 * @param {{ part?: number, hooks?: { flush?: () => void } }} [opts]
 */
export function* indexSteps(source, index, info, { part = 0, hooks = null } = {}) {
  const folders = source.folders();
  const total = folders.reduce((n, f) => n + (f.total || 0), 0);
  let done = 0;
  for (const folder of folders) {
    const key = `${info.id}\u0000${folder.path}`;
    if (index.segments.get(key)?.stamp === 'archive') {
      done += folder.total || 0;
      continue;
    }
    let parts = 0;
    let held = 0;
    const builderOf = () =>
      index.builder({
        key: parts ? `${key}\u0000${parts}` : key,
        stamp: 'archive',
        source: { id: info.id, name: info.name, kind: source.kind },
        folder: { path: folder.path, name: folder.name },
        textOf: (doc) => source.text(folder.path, doc.id),
      });
    let builder = builderOf();
    const flush = () => {
      if (!held) return;
      builder.finish();
      parts++;
      held = 0;
      builder = builderOf();
    };
    if (hooks) hooks.flush = flush;
    for (const { id, message } of source.messages(folder.path)) {
      const doc = docOf({ ...message, preview: message.preview || bodyText(message).slice(0, 200) }, id);
      const { words, sig } = index.analyse(headlineText(doc), bodyText(message));
      builder.add(doc, words, sig);
      done++;
      held++;
      if (part && held >= part) flush();
      yield { done, total, folder: folder.path };
    }
    if (held || !parts) builder.finish();
    if (hooks) hooks.flush = () => {};
  }
  return { done, total };
}

/** Run `indexSteps` to the end at once. */
export function indexArchive(source, index, info) {
  const steps = indexSteps(source, index, info);
  let r = steps.next();
  while (!r.done) r = steps.next();
  return r.value;
}

/**
 * Run a generator of steps in slices of about `slice` ms, giving the event
 * loop back between them. Resolves with the generator's return value.
 */
export async function runSliced(steps, { slice = 16, onStep = null, cancelled = () => false } = {}) {
  let r = steps.next();
  while (!r.done) {
    const until = Date.now() + slice;
    while (!r.done && Date.now() < until) {
      onStep?.(r.value);
      r = steps.next();
    }
    if (cancelled()) {
      steps.return?.();
      return null;
    }
    if (!r.done) await new Promise((resolve) => setImmediate(resolve));
  }
  return r.value;
}
