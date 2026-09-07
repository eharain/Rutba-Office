// One door for every mail archive.
//
// A person with mail to bring across does not know or care whether their
// history is a .pst, an .olm, an mbox or a directory of .eml files. They point
// at what they have. This decides what it is, reports what is inside it before
// anything is imported, and then streams the messages out folder by folder.
//
// Scanning is deliberately separate from importing. Telling someone "this file
// holds 12,172 messages in 467 folders — import all of it?" before spending ten
// minutes on it is the difference between a tool and a trap.

import fs from 'node:fs';
import path from 'node:path';
import { CompoundFile } from '@rutba/office-formats/cfb';
import { sniff } from '@rutba/office-formats/sniff';
import { PstFile } from './pst/index.js';
import { OlmArchive } from './olm.js';
import { OutlookMessage } from './msg.js';
import { Mbox } from './mbox.js';
import { parseMessage, summarize } from './mime.js';

/** A reader that pages a file from disk instead of loading it. */
export function fileReader(filePath) {
  const fd = fs.openSync(filePath, 'r');
  const size = fs.fstatSync(fd).size;
  return {
    size,
    read(offset, length) {
      const n = Math.max(0, Math.min(length, size - offset));
      const buf = Buffer.allocUnsafe(n);
      if (n) fs.readSync(fd, buf, 0, n, offset);
      return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
    },
    close: () => fs.closeSync(fd),
  };
}

const MAIL_EXTENSIONS = new Set(['.eml', '.emlx', '.msg', '.mbox', '.mbx', '.pst', '.ost', '.olm']);

function head(filePath, n = 4096) {
  const fd = fs.openSync(filePath, 'r');
  try {
    const buf = Buffer.allocUnsafe(Math.min(n, fs.fstatSync(fd).size));
    fs.readSync(fd, buf, 0, buf.length, 0);
    return new Uint8Array(buf);
  } finally {
    fs.closeSync(fd);
  }
}

/** What kind of mail archive is this? */
export function identify(target) {
  const stat = fs.statSync(target);
  if (stat.isDirectory()) {
    const names = fs.readdirSync(target);
    const maildir = ['cur', 'new', 'tmp'].every((d) => names.includes(d));
    return { kind: maildir ? 'maildir' : 'directory', path: target };
  }
  const bytes = head(target);
  const sniffed = sniff(bytes, path.basename(target));
  const ext = path.extname(target).toLowerCase();
  if (sniffed.kind === 'pst' || sniffed.kind === 'ost') return { kind: 'pst', path: target };
  if (ext === '.olm' || OlmArchive.is(fs.readFileSync(target))) return { kind: 'olm', path: target };
  if (sniffed.kind === 'msg' || (CompoundFile.is(bytes) && ext === '.msg')) return { kind: 'msg', path: target };
  if (sniffed.kind === 'mbox' || ext === '.mbox' || ext === '.mbx') return { kind: 'mbox', path: target };
  if (sniffed.kind === 'eml' || ext === '.eml' || ext === '.emlx') return { kind: 'eml', path: target };
  return { kind: 'unknown', path: target, detected: sniffed.kind };
}

/**
 * Look inside without importing: what folders, how many messages, and anything
 * that will not work — reported before the user commits to it.
 */
export function scan(target) {
  const found = identify(target);
  switch (found.kind) {
    case 'pst': {
      const pst = new PstFile(fileReader(target));
      try {
        const summary = pst.summary();
        return {
          kind: summary.kind,
          path: target,
          label: `Outlook data file (${summary.kind.toUpperCase()})`,
          readable: summary.readable,
          encoding: summary.encoding,
          folders: summary.folders,
          messages: summary.messages,
          bytes: summary.bytes,
        };
      } finally {
        pst.close();
      }
    }
    case 'olm': {
      const olm = new OlmArchive(fs.readFileSync(target));
      const summary = olm.summary();
      return { kind: 'olm', path: target, label: 'Outlook for Mac archive', readable: true, ...summary, bytes: fs.statSync(target).size };
    }
    case 'mbox': {
      const mbox = new Mbox(fs.readFileSync(target));
      return {
        kind: 'mbox',
        path: target,
        label: 'Mbox archive',
        readable: true,
        folders: [{ name: path.basename(target, path.extname(target)), path: '', depth: 0, messages: mbox.length }],
        messages: mbox.length,
        bytes: fs.statSync(target).size,
      };
    }
    case 'msg':
    case 'eml': {
      return {
        kind: found.kind,
        path: target,
        label: found.kind === 'msg' ? 'Outlook message' : 'Email message',
        readable: true,
        folders: [{ name: 'Imported', path: '', depth: 0, messages: 1 }],
        messages: 1,
        bytes: fs.statSync(target).size,
      };
    }
    case 'maildir':
    case 'directory': {
      const files = walkFiles(target).filter((f) => MAIL_EXTENSIONS.has(path.extname(f).toLowerCase()));
      const byFolder = new Map();
      for (const f of files) {
        const folder = path.relative(target, path.dirname(f)) || path.basename(target);
        byFolder.set(folder, (byFolder.get(folder) || 0) + 1);
      }
      return {
        kind: 'directory',
        path: target,
        label: found.kind === 'maildir' ? 'Maildir' : 'Folder of mail files',
        readable: true,
        folders: [...byFolder.entries()].map(([name, messages]) => ({ name, path: name, depth: name.split(/[\\/]/).length - 1, messages })),
        messages: files.length,
        bytes: files.reduce((n, f) => n + fs.statSync(f).size, 0),
      };
    }
    default:
      return { kind: 'unknown', path: target, label: 'Not a mail archive', readable: false, folders: [], messages: 0 };
  }
}

function walkFiles(dir, out = [], depth = 0) {
  if (depth > 12) return out;
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) walkFiles(full, out, depth + 1);
    else out.push(full);
  }
  return out;
}

/**
 * Stream messages out of an archive.
 * @param {string} target
 * @param {{ folders?: string[], onProgress?: (done, total, folder) => void, withBytes?: boolean }} [opts]
 * @yields {{ folder: string, message: object }}
 */
export function* read(target, opts = {}) {
  const found = identify(target);
  const wanted = opts.folders?.length ? new Set(opts.folders) : null;
  let done = 0;

  const emit = (folder, message) => {
    done++;
    opts.onProgress?.(done, opts.total ?? 0, folder);
    return { folder, message };
  };

  if (found.kind === 'pst') {
    const pst = new PstFile(fileReader(target));
    try {
      for (const folder of pst.folders().list) {
        if (folder.hidden || !folder.messageCount) continue;
        if (wanted && !wanted.has(folder.path) && !wanted.has(String(folder.nid))) continue;
        const size = 200;
        for (let offset = 0; offset < folder.messageCount; offset += size) {
          const page = pst.messages(folder.nid, { offset, limit: size });
          if (!page.length) break;
          for (const head of page) {
            const full = pst.message(head.nid);
            if (full) yield emit(folder.path || folder.name, full);
          }
        }
      }
    } finally {
      pst.close();
    }
    return;
  }

  if (found.kind === 'olm') {
    const olm = new OlmArchive(fs.readFileSync(target));
    for (const message of olm.messages()) {
      if (wanted && !wanted.has(message.folder)) continue;
      yield emit(message.folder, message);
    }
    return;
  }

  if (found.kind === 'mbox') {
    const mbox = new Mbox(fs.readFileSync(target));
    const folder = path.basename(target, path.extname(target));
    for (let i = 0; i < mbox.length; i++) {
      const message = mbox.message(i);
      if (message) yield emit(folder, message);
    }
    return;
  }

  if (found.kind === 'msg') {
    yield emit('Imported', OutlookMessage.open(fs.readFileSync(target)).toMessage());
    return;
  }

  if (found.kind === 'eml') {
    yield emit('Imported', parseMessage(fs.readFileSync(target)));
    return;
  }

  if (found.kind === 'maildir' || found.kind === 'directory') {
    for (const file of walkFiles(target)) {
      const ext = path.extname(file).toLowerCase();
      if (!MAIL_EXTENSIONS.has(ext) && ext !== '') continue;
      const folder = path.relative(target, path.dirname(file)) || path.basename(target);
      if (wanted && !wanted.has(folder)) continue;
      try {
        if (ext === '.msg') yield emit(folder, OutlookMessage.open(fs.readFileSync(file)).toMessage());
        else if (ext === '.mbox' || ext === '.mbx') {
          const mbox = new Mbox(fs.readFileSync(file));
          for (let i = 0; i < mbox.length; i++) {
            const m = mbox.message(i);
            if (m) yield emit(folder, m);
          }
        } else yield emit(folder, parseMessage(fs.readFileSync(file)));
      } catch {
        // One unreadable file must not end the import.
      }
    }
  }
}

/** Headers only, for a preview list. */
export function preview(target, { limit = 20 } = {}) {
  const out = [];
  for (const { folder, message } of read(target)) {
    out.push({ folder, ...summarize(message) });
    if (out.length >= limit) break;
  }
  return out;
}
