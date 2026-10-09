// Calendar and contacts accounts: CalDAV and CardDAV servers, kept in step.
//
// An account is a server address, a user name and a password (an app
// password, for iCloud and the like), kept as Mail keeps its own: the
// password in the encrypted secrets store, the rest in settings. Adding one
// finds the person's calendars and address books on it; each becomes a
// calendar here, or an address book whose cards join the rest. A sync then
// does, per collection: send what changed here (dav.js writes with If-Match,
// so an event changed on the server meanwhile is never overwritten — the
// server's copy wins and comes back), ask whether anything changed there
// (the collection's change tag), and if so list every item's ETag and fetch
// the ones that differ, dropping the ones that went.
//
// It runs when an account is added, when Sync is pressed, a little after a
// change made here, and every quarter of an hour while the suite is open.

import crypto from 'node:crypto';
import { readCalendar, writeCalendar } from '@rutba/calendar';
import { readVCards, writeVCard } from '@rutba/contacts';
import { createDavClient } from './dav.js';

const SETTING = 'dav.accounts';
const secretKey = (id) => `dav:${id}`;
const EVERY = 15 * 60 * 1000;
const AFTER_CHANGE = 8000;

/** A name for an item's file on the server, from its UID. */
const fileName = (uid, ext) => `${String(uid).replace(/[^A-Za-z0-9._@-]/g, '_').slice(0, 120) || crypto.randomUUID()}.${ext}`;
const join = (collection, name) => (collection.endsWith('/') ? collection : `${collection}/`) + name;

/** An event as the server keeps it: without what only this computer needs. */
const forServer = (e) => {
  const { id, createdAt, updatedAt, remote, ...rest } = e;
  return rest;
};

export function createDavService({ stores, calendar, contacts, broadcast = null, fetch: fetchImpl = globalThis.fetch, timers = true }) {
  const read = () => stores.settings.get(SETTING) || [];
  const write = (list) => stores.settings.set(SETTING, list);
  const passwordOf = (id) => stores.secrets.get(secretKey(id)) || '';
  const clientFor = (acc) => createDavClient({ url: acc.url, user: acc.user, password: passwordOf(acc.id), fetch: fetchImpl });
  const announce = () => broadcast?.('dav:changed', {});
  const listed = () => read().map(({ id, url, user, name, addedAt, lastSync, lastError, calendars, books }) => ({ id, url, user, name, addedAt, lastSync: lastSync || null, lastError: lastError || null, calendars: calendars || 0, books: books || 0 }));
  const cal = calendar?._dav || null;
  const con = contacts?._dav || null;
  const running = new Map();

  /** Bring the account's collections into step with the server's list: new ones made, gone ones dropped. */
  function reconcileCalendars(acc, remote) {
    const state = cal.state();
    for (const r of remote) {
      const kept = state.calendars.find((c) => c.remote?.account === acc.id && c.remote.href === r.href);
      if (kept) { kept.name = r.name || kept.name; if (r.colour) kept.colour = r.colour; continue; }
      state.calendars.push({
        id: crypto.randomUUID(), name: r.name, colour: r.colour || cal.colours[state.calendars.length % cal.colours.length], visible: true, events: [],
        remote: { account: acc.id, href: r.href, ctag: null }, dirty: [], deleted: [],
      });
    }
    const hrefs = new Set(remote.map((r) => r.href));
    state.calendars = state.calendars.filter((c) => c.remote?.account !== acc.id || hrefs.has(c.remote.href));
  }

  function reconcileBooks(acc, remote) {
    const state = con.state();
    state.books = state.books || [];
    for (const r of remote) {
      const kept = state.books.find((b) => b.account === acc.id && b.href === r.href);
      if (kept) { kept.name = r.name || kept.name; continue; }
      state.books.push({ id: crypto.randomUUID(), account: acc.id, href: r.href, name: r.name, ctag: null });
    }
    const hrefs = new Set(remote.map((r) => r.href));
    const gone = new Set(state.books.filter((b) => b.account === acc.id && !hrefs.has(b.href)).map((b) => b.id));
    state.books = state.books.filter((b) => !gone.has(b.id));
    state.contacts = state.contacts.filter((c) => !gone.has(c.book));
  }

  /** One calendar: what changed here sent, then what changed there fetched. Answers counts. */
  async function syncCalendar(client, c) {
    const done = { sent: 0, received: 0, removed: 0, conflicts: 0 };
    const href = c.remote.href;
    const refetch = new Set();
    // What is queued now is what this sync sends; anything queued while it
    // runs (an event made or removed meanwhile) is left for the next.
    const removing = [...(c.deleted || [])];
    for (const d of removing) {
      try { await client.remove(d.href, d.etag); done.sent += 1; } catch (err) { if (!err.conflict) throw err; done.conflicts += 1; refetch.add(d.href); }
    }
    c.deleted = (c.deleted || []).filter((d) => !removing.includes(d));
    const failed = [];
    const sending = [...(c.dirty || [])];
    for (const uid of sending) {
      const events = c.events.filter((e) => e.uid === uid);
      if (!events.length) continue;
      const at = events.find((e) => e.remote?.href)?.remote || null;
      const target = at?.href || join(new URL(href, client.origin()).pathname, fileName(uid, 'ics'));
      try {
        const etag = await client.put(target, writeCalendar({ events: events.map(forServer) }), { etag: at?.etag || null, kind: 'caldav' });
        for (const e of events) e.remote = { href: target, etag: etag || null };
        if (!etag) refetch.add(target);
        done.sent += 1;
      } catch (err) {
        if (err.conflict) { done.conflicts += 1; refetch.add(target); } else failed.push(uid);
      }
    }
    c.dirty = [...new Set([...failed, ...(c.dirty || []).filter((u) => !sending.includes(u))])];
    const tag = await client.ctag(href);
    if (tag && tag === c.remote.ctag && !refetch.size && !done.sent) return done;
    const server = await client.etags(href, 'caldav');
    const local = new Map();
    for (const e of c.events) if (e.remote?.href) local.set(e.remote.href, e.remote.etag);
    const wanted = [...server].filter(([h, etag]) => refetch.has(h) || local.get(h) !== etag || !etag).map(([h]) => h);
    for (let i = 0; i < wanted.length; i += 50) {
      for (const item of await client.multiget(href, wanted.slice(i, i + 50), 'caldav')) {
        let events;
        try { events = readCalendar(item.data).events; } catch { continue; }
        const before = c.events.filter((e) => e.remote?.href === item.href || (events.some((x) => x.uid && x.uid === e.uid) && !e.remote?.href));
        const now = Date.now();
        const idFor = (e) => before.find((b) => b.uid === e.uid && Boolean(b.recurrenceId) === Boolean(e.recurrenceId) && (!e.recurrenceId || Math.abs((b.recurrenceId?.at ?? 0) - e.recurrenceId.at) < 1000))?.id;
        c.events = c.events.filter((e) => !before.includes(e));
        for (const e of events) c.events.push({ ...e, id: idFor(e) || crypto.randomUUID(), createdAt: now, updatedAt: now, remote: { href: item.href, etag: item.etag } });
        done.received += 1;
      }
    }
    const kept = c.events.length;
    c.events = c.events.filter((e) => !e.remote?.href || server.has(e.remote.href));
    done.removed += kept - c.events.length;
    c.remote.ctag = (await client.ctag(href)) || tag;
    return done;
  }

  /** One address book, the same way. */
  async function syncBook(client, book) {
    const state = con.state();
    const done = { sent: 0, received: 0, removed: 0, conflicts: 0 };
    const refetch = new Set();
    const mine = (d) => d.book === book.id;
    const removing = (state.deleted || []).filter(mine);
    for (const d of removing) {
      try { await client.remove(d.href, d.etag); done.sent += 1; } catch (err) { if (!err.conflict) throw err; done.conflicts += 1; refetch.add(d.href); }
    }
    state.deleted = (state.deleted || []).filter((d) => !removing.includes(d));
    const failed = [];
    const sending = [...(state.dirty || [])];
    for (const id of sending) {
      const c = state.contacts.find((x) => x.id === id);
      if (!c) continue;
      if (c.book !== book.id) { if (c.book) failed.push(id); continue; }
      if (!c.uid) c.uid = crypto.randomUUID();
      const target = c.remote?.href || join(new URL(book.href, client.origin()).pathname, fileName(c.uid, 'vcf'));
      try {
        const etag = await client.put(target, writeVCard(con.forWriting(c), { version: '3.0' }), { etag: c.remote?.etag || null, kind: 'carddav' });
        c.remote = { href: target, etag: etag || null };
        if (!etag) refetch.add(target);
        done.sent += 1;
      } catch (err) {
        if (err.conflict) { done.conflicts += 1; refetch.add(target); } else failed.push(id);
      }
    }
    state.dirty = [...new Set([...failed, ...(state.dirty || []).filter((id) => !sending.includes(id))])];
    const tag = await client.ctag(book.href);
    if (tag && tag === book.ctag && !refetch.size && !done.sent) return done;
    const server = await client.etags(book.href, 'carddav');
    const local = new Map();
    for (const c of state.contacts) if (c.book === book.id && c.remote?.href) local.set(c.remote.href, c.remote.etag);
    const wanted = [...server].filter(([h, etag]) => refetch.has(h) || local.get(h) !== etag || !etag).map(([h]) => h);
    for (let i = 0; i < wanted.length; i += 50) {
      for (const item of await client.multiget(book.href, wanted.slice(i, i + 50), 'carddav')) {
        let card;
        try { [card] = readVCards(item.data); } catch { continue; }
        if (!card) continue;
        const existing = state.contacts.find((c) => c.book === book.id && (c.remote?.href === item.href || (card.uid && c.uid === card.uid)));
        const now = Date.now();
        const fresh = { ...con.storable(card), id: existing?.id || crypto.randomUUID(), createdAt: existing?.createdAt || now, updatedAt: now, book: book.id, remote: { href: item.href, etag: item.etag } };
        if (existing) state.contacts[state.contacts.indexOf(existing)] = fresh;
        else state.contacts.push(fresh);
        done.received += 1;
      }
    }
    const kept = state.contacts.length;
    state.contacts = state.contacts.filter((c) => c.book !== book.id || !c.remote?.href || server.has(c.remote.href));
    done.removed += kept - state.contacts.length;
    book.ctag = (await client.ctag(book.href)) || tag;
    return done;
  }

  async function syncAccount(acc) {
    const client = clientFor(acc);
    const totals = { sent: 0, received: 0, removed: 0, conflicts: 0 };
    const add = (d) => { for (const k of Object.keys(totals)) totals[k] += d[k]; };
    if (cal) {
      const remote = await client.discover('caldav');
      reconcileCalendars(acc, remote);
      for (const c of cal.state().calendars.filter((x) => x.remote?.account === acc.id)) add(await syncCalendar(client, c));
      cal.save();
      acc.calendars = remote.length;
    }
    if (con) {
      const remote = await client.discover('carddav');
      reconcileBooks(acc, remote);
      for (const b of (con.state().books || []).filter((x) => x.account === acc.id)) add(await syncBook(client, b));
      con.save();
      acc.books = remote.length;
    }
    return totals;
  }

  /** Sync one account, or every one. Answers what moved, per account. */
  async function sync({ id = null } = {}) {
    const out = [];
    for (const acc of read().filter((a) => !id || a.id === id)) {
      if (running.has(acc.id)) {
        // A sync asked for while one runs can be for a change that one began
        // too soon to see (an event made a moment ago): it waits for it, then
        // runs once more, a run that later asks share.
        await running.get(acc.id);
        if (running.has(acc.id)) { out.push({ id: acc.id, ...(await running.get(acc.id)) }); continue; }
      }
      const job = (async () => {
        try {
          const totals = await syncAccount(acc);
          const list = read();
          const kept = list.find((a) => a.id === acc.id);
          if (kept) Object.assign(kept, { lastSync: Date.now(), lastError: null, calendars: acc.calendars, books: acc.books });
          write(list);
          return totals;
        } catch (err) {
          const list = read();
          const kept = list.find((a) => a.id === acc.id);
          if (kept) Object.assign(kept, { lastError: String(err.message || err) });
          write(list);
          return { error: String(err.message || err) };
        } finally {
          running.delete(acc.id);
          announce();
        }
      })();
      running.set(acc.id, job);
      out.push({ id: acc.id, ...(await job) });
    }
    return out;
  }

  let soon = null;
  const later = () => {
    if (!timers) return;
    clearTimeout(soon);
    soon = setTimeout(() => { soon = null; sync().catch(() => {}); }, AFTER_CHANGE);
  };
  if (cal) cal.hooks.local = later;
  if (con) con.hooks.local = later;
  const every = timers ? setInterval(() => { if (read().length) sync().catch(() => {}); }, EVERY) : null;
  every?.unref?.();

  return {
    accounts: () => listed(),

    /**
     * Add an account: the server reached and asked for its calendars and
     * address books, which are brought in at once. Refused, with why, when
     * the server cannot be reached, will not take the password, or keeps
     * neither.
     */
    add: async ({ url, user = '', password = '', name = '' }) => {
      const client = createDavClient({ url, user, password, fetch: fetchImpl });
      const [calendars, books] = await Promise.all([
        cal ? client.discover('caldav') : [],
        con ? client.discover('carddav') : [],
      ]);
      if (!calendars.length && !books.length) throw new Error('No calendars or address books were found at that address.');
      const id = crypto.randomUUID();
      const acc = { id, url: String(url).trim(), user, name: name || new URL(client.origin()).hostname, addedAt: Date.now(), calendars: calendars.length, books: books.length };
      stores.secrets.set(secretKey(id), password);
      write([...read(), acc]);
      const [result] = await sync({ id });
      return { ...listed().find((a) => a.id === id), result };
    },

    /** Take an account away, and the calendars and cards it brought. */
    remove: ({ id }) => {
      const list = read();
      if (!list.some((a) => a.id === id)) return { removed: 0 };
      write(list.filter((a) => a.id !== id));
      stores.secrets.delete?.(secretKey(id));
      if (cal) {
        const state = cal.state();
        state.calendars = state.calendars.filter((c) => c.remote?.account !== id);
        if (!state.calendars.length) state.calendars.push({ id: crypto.randomUUID(), name: 'My calendar', colour: cal.colours[0], visible: true, events: [] });
        cal.save();
      }
      if (con) {
        const state = con.state();
        const books = new Set((state.books || []).filter((b) => b.account === id).map((b) => b.id));
        state.books = (state.books || []).filter((b) => !books.has(b.id));
        state.contacts = state.contacts.filter((c) => !books.has(c.book));
        if (books.has(state.defaultBook)) state.defaultBook = null;
        con.save();
      }
      announce();
      return { removed: 1 };
    },

    sync,

    /** The address books, and which one a new card goes to (null: this computer only). */
    books: () => ({ books: (con?.state().books || []).map(({ id, name, account }) => ({ id, name, account })), defaultBook: con?.state().defaultBook || null }),

    setDefaultBook: ({ id = null }) => {
      if (!con) return null;
      con.state().defaultBook = id && (con.state().books || []).some((b) => b.id === id) ? id : null;
      con.save();
      return con.state().defaultBook;
    },

    stop: () => { clearInterval(every); clearTimeout(soon); },
  };
}
