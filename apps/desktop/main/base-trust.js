// Whether a database's sources and engine links may be read (A2.18). A
// database someone else made and sent, with a table linked to a web
// address, a file path or an engine, would otherwise be read as soon as it
// opened: the address's owner learns it was opened and from where, a
// \\server\share path makes Windows offer the person's sign-in to that
// server, and a path on this computer reads the person's own file into the
// database, which goes back with it if it is sent on. So nothing is read
// or connected to until the person on this computer has allowed it.
//
// What is allowed is kept in this computer's settings, never in the file,
// which could claim anything about itself: for each database, by where it
// is on this computer, the addresses of the sources and engines linked
// there by the person or allowed by them. A link made here is allowed as it
// is made, and a database made here has nothing else in it, so neither
// asks. A link added to the file elsewhere, or one whose address was
// changed elsewhere, is an address not on the list, and is asked about. A
// database from before this rule asks once: nothing in it says reliably
// where its links were made.

import os from 'node:os';
import path from 'node:path';
import { SOURCE_KINDS } from '@rutba/database';

const WINDOWS = process.platform === 'win32';
/** Most databases remembered; past it, the one allowed longest ago is let go (and asks again). */
const MOST = 1000;

const isWeb = (where) => /^https?:\/\//i.test(String(where || '').trim());
const hostOf = (url) => { try { return new URL(url).host; } catch { return String(url); } };

/** A file as this computer names it, the same file spelt alike: resolved, and on Windows in one case. */
export const placeKey = (file) => {
  const p = path.resolve(String(file || ''));
  return WINDOWS ? p.toLowerCase() : p;
};

/** What a link reads or connects to, the same address spelt alike: its kind and where, or its engine's host, port, database and account. */
export function addressOf(link) {
  const s = link?.settings || {};
  if (SOURCE_KINDS.has(link?.engine)) {
    const where = String(s.where || '').trim();
    return `${link.engine} ${!where || isWeb(where) ? where : placeKey(where)}`;
  }
  return `${link?.engine} ${String(s.host || '').trim().toLowerCase()}:${Number(s.port) || ''}/${s.database || ''}/${s.user || ''}`;
}

/**
 * Where a path is, as the question says it: on a network share (and which
 * server), in the person's own folders, or outside them.
 */
export function pathPlace(where, { home = os.homedir() } = {}) {
  let p = String(where || '').trim();
  // \\?\ and \\.\ name a path as the system does: \\?\UNC\server\share is a share, \\?\C:\ a drive.
  const device = /^[\\/]{2}[?.][\\/]/.test(p);
  if (device) {
    p = p.slice(4);
    if (/^UNC[\\/]/i.test(p)) return { place: 'share', server: p.slice(4).split(/[\\/]/)[0], path: String(where).trim() };
    if (!/^[a-z]:/i.test(p)) return { place: 'outside', path: String(where).trim() };
  }
  if (/^[\\/]{2}[^\\/]/.test(p)) return { place: 'share', server: p.slice(2).split(/[\\/]/)[0], path: String(where).trim() };
  const full = path.resolve(p);
  const fold = (x) => (WINDOWS ? x.toLowerCase() : x);
  const rel = path.relative(fold(path.resolve(home)), fold(full));
  const mine = Boolean(home) && rel !== '' && !rel.startsWith('..') && !path.isAbsolute(rel);
  return { place: mine ? 'mine' : 'outside', path: full };
}

/** A link as the question names it: { place, host?, url?, server?, path?, port?, database?, user? }. */
export function placeOf(link, opts = {}) {
  const s = link?.settings || {};
  if (SOURCE_KINDS.has(link?.engine)) {
    const where = String(s.where || '').trim();
    if (isWeb(where)) return { place: 'web', host: hostOf(where), url: where };
    return pathPlace(where, opts);
  }
  return { place: 'engine', host: String(s.host || ''), port: Number(s.port) || null, database: String(s.database || ''), user: String(s.user || '') };
}

/** A store kept in memory: a service made without this computer's settings (the engine tests) remembers for its own life. */
export function memoryStore() {
  let kept = {};
  return { get: () => kept, set: (v) => { kept = v; } };
}

/**
 * This computer's list of what each database may read, over `store` ({
 * get, set } of one value in this computer's settings).
 */
export function createTrust(store = memoryStore()) {
  const read = () => { const v = store.get(); return v && typeof v === 'object' && !Array.isArray(v) ? v : {}; };
  const listOf = (all, file) => (Array.isArray(all[placeKey(file)]?.allowed) ? all[placeKey(file)].allowed : []);
  const write = (all) => {
    const keys = Object.keys(all);
    if (keys.length > MOST) {
      for (const k of keys.sort((a, b) => (all[a].at || 0) - (all[b].at || 0)).slice(0, keys.length - MOST)) delete all[k];
    }
    store.set(all);
  };
  return {
    /** Whether the database at `file` may read or connect to `link` on this computer. */
    allowed: (file, link) => Boolean(file && link) && listOf(read(), file).includes(addressOf(link)),
    /** `links` allowed for the database at `file`: linked here by the person, or allowed when asked. */
    allow(file, links) {
      const all = { ...read() };
      const list = new Set(listOf(all, file));
      for (const l of links || []) if (l) list.add(addressOf(l));
      all[placeKey(file)] = { at: Date.now(), allowed: [...list] };
      write(all);
    },
    /** A new database at `file`: what it may read is only what it was made with, nothing a file there before was allowed. */
    reset(file, links = []) {
      const all = { ...read() };
      delete all[placeKey(file)];
      if (links.length) all[placeKey(file)] = { at: Date.now(), allowed: [...new Set(links.map(addressOf))] };
      write(all);
    },
    /** The database at `from` now at `to` too (Save As) or only (a rename): what it may read goes with it. */
    carry(from, to, { move = false } = {}) {
      const all = { ...read() };
      const was = all[placeKey(from)];
      if (move) delete all[placeKey(from)];
      if (was) all[placeKey(to)] = { ...was, at: Date.now() };
      else delete all[placeKey(to)];
      write(all);
    },
  };
}
