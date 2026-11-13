// Messages as .eml files (Mail 2.25): saved out, one to a file or several
// into a folder; brought in, from File, Open, from files dropped on a folder
// or the list, or from files copied in Explorer and pasted; and a file opened
// from Explorer shown on its own, read only, until it is added to a folder.
//
// A message brought in keeps the date it was written and whether it was read:
// the file's own Status, X-Status or X-Mozilla-Status headers say so where a
// mail program wrote them, and a file without them is mail already read. In
// a live account it is APPENDed to the server's folder with those marks and
// that date, as Paste does between accounts (mail-paste.js).

import fs from 'node:fs';
import { removeTreeSync } from '@rutba/office-shell/remove-tree';
import path from 'node:path';
import crypto from 'node:crypto';

/** The most of one file read as a message: anything larger is not one. */
export const EML_LIMIT = 100 * 1024 * 1024;

/** What a file opened from Explorer is kept under while its window is open: never an account's id. */
export const OPENED = 'opened_';
export const isOpened = (id) => typeof id === 'string' && id.startsWith(OPENED);

/** The header block of a message's bytes, unfolded, as name/value pairs (lower-case names). */
function headerPairs(bytes) {
  const text = Buffer.from(bytes).subarray(0, 64 * 1024).toString('latin1');
  const end = text.search(/\r?\n\r?\n/);
  const block = (end < 0 ? text : text.slice(0, end)).replace(/\r?\n[ \t]+/g, ' ');
  return block.split(/\r?\n/).map((line) => {
    const at = line.indexOf(':');
    return at > 0 ? [line.slice(0, at).trim().toLowerCase(), line.slice(at + 1).trim()] : null;
  }).filter(Boolean);
}

/** Whether the bytes begin as a message does: header lines, and one of the headers every message has. */
export function looksLikeMessage(bytes) {
  const pairs = headerPairs(bytes);
  if (!pairs.length) return false;
  const names = new Set(pairs.map(([n]) => n));
  return ['from', 'subject', 'date', 'message-id', 'to', 'received', 'mime-version'].some((n) => names.has(n));
}

/**
 * Whether the file's message was read, flagged and answered. A file says so
 * only where a mail program wrote it in: mbox's Status and X-Status, and
 * Thunderbird's X-Mozilla-Status bits (1 read, 2 answered, 4 flagged). With
 * none of them it was read: a message saved as a file is one already seen.
 */
export function emlMarks(bytes) {
  const pairs = headerPairs(bytes);
  const get = (name) => pairs.find(([n]) => n === name)?.[1] ?? null;
  const marks = { unread: false, flagged: false, answered: false };
  const mozilla = get('x-mozilla-status');
  if (mozilla && /^[0-9a-f]{1,4}$/i.test(mozilla)) {
    const bits = parseInt(mozilla, 16);
    marks.unread = !(bits & 0x1);
    marks.answered = Boolean(bits & 0x2);
    marks.flagged = Boolean(bits & 0x4);
  }
  const status = get('status');
  if (status != null) marks.unread = !/R/.test(status);
  const x = get('x-status');
  if (x != null) {
    if (/F/.test(x)) marks.flagged = true;
    if (/A/.test(x)) marks.answered = true;
  }
  return marks;
}

/** The IMAP marks for those, as the APPEND carries them. */
export function appendFlags(marks) {
  const flags = [];
  if (!marks.unread) flags.push('\\Seen');
  if (marks.flagged) flags.push('\\Flagged');
  if (marks.answered) flags.push('\\Answered');
  return flags;
}

/** A file name for a message saved as .eml: its subject, made safe on every system, and never empty. */
export function emlFileName(message, fallback = 'message') {
  const base = String(message?.subject || '')
    .replace(/[\u0000-\u001f<>:"/\\|?*]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[. ]+$/, '')
    .slice(0, 80);
  const safe = /^(con|prn|aux|nul|com[0-9]|lpt[0-9])$/i.test(base) ? `${base}_` : base;
  return `${safe || fallback}.eml`;
}

/** A name not yet taken in `dir` (or by `taken`): "Name.eml", then "Name (2).eml" and on. */
export function freeName(dir, name, taken = new Set()) {
  const ext = path.extname(name);
  const stem = name.slice(0, name.length - ext.length);
  for (let n = 1; ; n++) {
    const candidate = n === 1 ? name : `${stem} (${n})${ext}`;
    const key = candidate.toLowerCase();
    if (!taken.has(key) && !fs.existsSync(path.join(dir, candidate))) {
      taken.add(key);
      return candidate;
    }
  }
}

/** One file's bytes, refused when it is not a file Mail can read as a message. */
function readFileBytes(p) {
  let size;
  try {
    size = fs.statSync(p).size;
  } catch {
    throw new Error('That file is not there any more.');
  }
  if (size > EML_LIMIT) throw new Error('That file is too large to be one message.');
  return fs.readFileSync(p);
}

/**
 * The mail service's part (mail.js): `ctx` is its store, its accounts and the
 * changes it sends its servers.
 * @param {{ store, find, isArchive, live, record, kick, listFolder, sourceBytes, parseMessage, broadcast?, marked?, userData }} ctx
 */
export function createEmlDesk(ctx) {
  const { store, find, isArchive, live, record, kick, listFolder, sourceBytes, parseMessage, broadcast = null, marked = async () => {}, userData } = ctx;

  /** A message's bytes, read as one, with its marks: or why not. */
  const asMessage = (bytes) => {
    if (!looksLikeMessage(bytes)) return null;
    const message = parseMessage(bytes);
    return { message, marks: emlMarks(bytes) };
  };

  /**
   * .eml files brought into a folder, from `paths` on disk or `files` handed
   * over as bytes (a drop or a paste whose files have no path). Answers
   * { added, skipped, folder }: a file that is not a message is passed over and counted.
   */
  async function importEml({ accountId, folder, paths = [], files = [] } = {}) {
    if (isArchive(accountId) || isOpened(accountId)) throw new Error('An archive opened in place is read only: messages are not copied out of it or into it.');
    find(accountId);
    const to = String(folder || '').trim();
    if (!to) throw new Error('Paste needs a folder to go to.');
    const sources = [
      ...(paths || []).map((p) => ({ name: path.basename(String(p)), read: () => readFileBytes(String(p)) })),
      ...(files || []).map((f) => ({ name: String(f?.name || 'message.eml'), read: () => Buffer.from(f?.bytes || []) })),
    ];
    const serverOf = live(accountId);
    let added = 0;
    let skipped = 0;
    for (const source of sources) {
      let bytes;
      try {
        bytes = source.read();
      } catch {
        skipped++;
        continue;
      }
      if (!bytes.length || bytes.length > EML_LIMIT) {
        skipped++;
        continue;
      }
      const read = asMessage(bytes);
      if (!read) {
        skipped++;
        continue;
      }
      const { message, marks } = read;
      const id = store.put(accountId, to, { ...message, ...marks }, { force: true, flush: true });
      if (!id) {
        skipped++;
        continue;
      }
      if (serverOf) {
        const file = store.writeOutgoing(accountId, `import-${id}-${Date.now()}.eml`, bytes);
        record(accountId, { kind: 'append', folder: to, file, id, messageId: message.messageId || null, flags: appendFlags(marks), date: message.date || null });
      }
      added++;
    }
    if (added) {
      listFolder(accountId, to);
      if (serverOf) kick(accountId);
      broadcast?.('mail:new', { accountId, folder: to, count: added });
    }
    return { added, skipped, folder: to };
  }

  /**
   * Messages saved as .eml: one message to the file `path`, or several (or
   * one, when `path` is a folder) each to a file of its own in the folder
   * `path`, named by its subject. Answers { written: [file names], path }.
   */
  async function saveEml({ accountId, folder, ids = [], path: target } = {}) {
    if (!target) throw new Error('Saving needs a place to save to.');
    const here = [...new Set(ids || [])].map((id) => ({ id, message: store.get(accountId, folder, id) })).filter((m) => m.message);
    if (!here.length) throw new Error('That message is not there any more.');
    const toFile = here.length === 1 && /\.eml$/i.test(String(target)) && !(fs.existsSync(target) && fs.statSync(target).isDirectory());
    if (toFile) {
      const bytes = await sourceBytes(accountId, folder, here[0].id, here[0].message);
      fs.mkdirSync(path.dirname(target), { recursive: true });
      fs.writeFileSync(target, bytes);
      await marked(target);
      return { written: [path.basename(target)], path: target };
    }
    fs.mkdirSync(target, { recursive: true });
    const taken = new Set();
    const written = [];
    for (const { id, message } of here) {
      const name = freeName(target, emlFileName(message), taken);
      fs.writeFileSync(path.join(target, name), await sourceBytes(accountId, folder, id, message));
      written.push(name);
    }
    for (const name of written) await marked(path.join(target, name));
    return { written, path: target };
  }

  /** Where the files opened from Explorer are kept while their windows are open. */
  const openedRoot = () => path.join(userData, 'mail');

  /**
   * A .eml opened from Explorer: kept apart from every account, read only,
   * for its window to show. Answers { accountId, folder, id, name, path }:
   * the reader reads it as it reads any message, and Add to a folder brings
   * the file in with importEml.
   */
  function openEml({ path: target } = {}) {
    const bytes = readFileBytes(String(target || ''));
    const read = asMessage(bytes);
    if (!read) throw new Error('That file is not a mail message.');
    const key = crypto.createHash('sha256').update(path.resolve(String(target))).update(String(bytes.length)).digest('hex').slice(0, 16);
    const accountId = `${OPENED}${key}`;
    const folder = 'Opened';
    const id = store.put(accountId, folder, { ...read.message, unread: false, openedFrom: String(target) }, { force: true, flush: true });
    return { accountId, folder, id, name: path.basename(String(target)), path: String(target) };
  }

  /** The files opened before, gone: they were only ever kept for their windows. */
  function sweepOpened() {
    let gone = 0;
    try {
      for (const entry of fs.readdirSync(openedRoot(), { withFileTypes: true })) {
        if (!entry.isDirectory() || !isOpened(entry.name)) continue;
        removeTreeSync(path.join(openedRoot(), entry.name));
        gone++;
      }
    } catch {
      /* no mail kept yet */
    }
    return gone;
  }

  return { importEml, saveEml, openEml, sweepOpened };
}
