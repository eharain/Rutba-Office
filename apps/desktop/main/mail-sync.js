// Mail and its server, kept in step.
//
// Before 2.19 every change stayed on this computer: a message marked read,
// filed or deleted here was so only here, and Get mail fetched the newest few
// hundred again by position, so what was deleted came back and a phone saw
// none of it. Now:
//
//   - A change made here (read and unread, flagged, moved, junk and not junk,
//     deleted to the Trash or for good, a sent message's copy in Sent, a
//     folder made, renamed or deleted) is first written to the account's
//     line of changes (store.addChange), then made to the local store, then
//     replayed to the server in order. A change on disk before the store is
//     touched is never lost; one the server has is taken off the line, and
//     each is one the server can be told twice without harm (STORE and MOVE
//     and EXPUNGE by UID; an APPEND only after a search by Message-ID finds
//     the message is not there yet). Offline, or with the server unreachable,
//     the line waits and is replayed on the next connection.
//   - A change the server refuses is taken off the line, said in plain words
//     (mail:syncProblem), and the message put back as the server has it.
//   - A fetch goes by UID from the last one seen in each folder. A folder
//     whose UIDVALIDITY changed is matched again by Message-ID, so nothing is
//     doubled. Marks made at the server elsewhere come with CONDSTORE's
//     CHANGEDSINCE where the server has it, otherwise from a flags fetch over
//     the newest FETCH_WINDOW messages; messages gone at the server (deleted
//     or filed elsewhere) go here too, from a UID SEARCH over the same window.
//   - One account is fetched or replayed by one job at a time.
//
// Local only, and said so where it happens: imported mailboxes (account.local)
// and archives opened in place (archive:… ids) have no server, so their
// changes go to the store alone, as before. A message's pin is Rutba Office's
// own mark, with no IMAP flag to carry it. A draft saved here stays here
// (one opened from the server's Drafts and sent is deleted there, as any
// message is). Mail has no POP3 accounts.

import fs from 'node:fs';
import { messageKey } from '@rutba/mailbox/store';
import { parseMessage } from '@rutba/mailbox/mime';

/** How many of a folder's newest messages have their marks and presence compared with the server's on each fetch. */
export const FETCH_WINDOW = 2000;
/** The most messages read by Message-ID when a folder is first matched with its server, or matched again. */
const ADOPT_LIMIT = 20000;
/** Messages fetched whole at once. */
const BATCH = 50;
/** How long a change waits before it is sent, so a run of them goes on one connection. */
const KICK_MS = 250;
/** How long the line waits to try again after the server could not be reached. */
const RETRY_MS = 60_000;

/** The Message-ID a header block carries, as the store keys it (one pair of angle brackets off). */
export function messageIdOf(headers) {
  const m = /^message-id:[ \t]*((?:.*)(?:\r?\n[ \t].*)*)/im.exec(Buffer.isBuffer(headers) ? headers.toString('utf8') : String(headers || ''));
  const value = m ? m[1].replace(/\s+/g, '') : '';
  return value.replace(/^<|>$/g, '') || null;
}

/** The store's id for a message known only by its Message-ID. */
const keyOf = (messageId) => (messageId ? messageKey({ messageId }) : null);

/** A message's marks, as the store keeps them, from its IMAP flags. */
const marksOf = (flags) => ({ unread: !flags?.has?.('\\Seen'), flagged: Boolean(flags?.has?.('\\Flagged')) });

/** IMAP flags to add and to take off for a change of marks; a pin is local only. */
export function flagsFor(patch) {
  const add = [];
  const remove = [];
  if ('unread' in patch) (patch.unread ? remove : add).push('\\Seen');
  if ('flagged' in patch) (patch.flagged ? add : remove).push('\\Flagged');
  return { add, remove };
}

/** A connection that is gone, as against a server that said no. */
const lost = (client, err) => !client?.usable || /ECONN|ETIMEDOUT|EPIPE|ENOTFOUND|EAI_AGAIN|NoConnection|closed|socket/i.test(`${err?.code || ''} ${err?.message || ''}`);

const uidSet = (uids) => [...new Set(uids)].sort((a, b) => a - b).join(',');

/**
 * @param {object} o
 * @param {import('@rutba/mailbox/store').MailStore} o.store
 * @param {(accountId: string) => object|null} o.account the live account, or null for one with no server
 * @param {(account: object) => Promise<object>} o.connect an ImapFlow client, signed in
 * @param {(entry: object) => object} o.describe what the store keeps of a folder the server listed
 * @param {(accountId: string, arrived: {folder: string, ids: string[]}[]) => Promise<object>} [o.arrived] run over what a fetch brought, before the line is replayed again
 * @param {(accountId: string, capabilities: Map) => void} [o.seen] what the server said it offers, each fetch
 * @param {(event: string, payload: object) => void} [o.broadcast]
 */
export function createMailSync({ store, account: liveAccount, connect, describe, arrived: onArrived = null, seen = null, broadcast = null }) {
  const chains = new Map();
  const waiting = new Map();
  const kicks = new Map();
  const retries = new Map();
  const watchers = new Map();
  const states = new Map();

  /** One job at a time per account: each waits for the one before. */
  function serial(accountId, job) {
    const run = (chains.get(accountId) || Promise.resolve()).catch(() => {}).then(job);
    chains.set(accountId, run.catch(() => {}));
    return run;
  }

  const note = (accountId, patch) => {
    const next = { ...(states.get(accountId) || {}), ...patch };
    states.set(accountId, next);
    broadcast?.('mail:syncState', { accountId, ...next });
  };

  /* ── the line of changes ─────────────────────────────────────────── */

  /**
   * A change made here, on the line before the store is touched. Answers
   * null, and keeps nothing, for an account with no server: its changes are
   * the store's alone.
   */
  function record(accountId, change) {
    if (!liveAccount(accountId)) return null;
    // The UIDs kept are good only under the UIDVALIDITY they were given in.
    const uidValidity = change.items ? store.syncState(accountId, change.folder)?.uidValidity || null : null;
    return store.addChange(accountId, { ...change, uidValidity });
  }

  /** The messages a change is about, as the line keeps them: by id, with the UID and Message-ID they have now. */
  function itemsOf(accountId, folder, ids) {
    const state = store.syncState(accountId, folder);
    return store.headers(accountId, folder, ids).map((r) => ({ id: r.id, uid: state ? r.uid ?? null : null, messageId: r.messageId || null }));
  }

  /** Send the line soon, on one connection for a run of changes. */
  function kick(accountId) {
    if (!liveAccount(accountId)) return;
    clearTimeout(kicks.get(accountId));
    const timer = setTimeout(() => {
      kicks.delete(accountId);
      flush(accountId).catch(() => {});
    }, KICK_MS);
    timer.unref?.();
    kicks.set(accountId, timer);
  }

  /** The line replayed now: resolves once it is empty, or the server could not be reached (then it waits). */
  function flush(accountId) {
    clearTimeout(kicks.get(accountId));
    kicks.delete(accountId);
    const account = liveAccount(accountId);
    if (!account || !store.changes(accountId).items.length) return Promise.resolve({ left: 0 });
    return serial(accountId, async () => {
      if (!store.changes(accountId).items.length) return { left: 0 };
      let client = null;
      try {
        client = await connect(account);
        await replay(client, accountId);
        note(accountId, { offline: false, error: null });
        return { left: store.changes(accountId).items.length };
      } catch (err) {
        later(accountId);
        note(accountId, { offline: true, error: err.message });
        return { left: store.changes(accountId).items.length, error: err.message };
      } finally {
        await client?.logout().catch(() => client?.close?.());
      }
    });
  }

  /** Try again in a while: the server was not there. */
  function later(accountId) {
    if (retries.has(accountId)) return;
    const timer = setTimeout(() => {
      retries.delete(accountId);
      flush(accountId).catch(() => {});
    }, RETRY_MS);
    timer.unref?.();
    retries.set(accountId, timer);
  }

  /** The folders the server has, by path, read once a connection. */
  async function listed(client, ctx) {
    if (!ctx.listed) ctx.listed = new Map((await client.list()).map((f) => [f.path, f]));
    return ctx.listed;
  }

  /** Make sure a folder is there before something is put in it. */
  async function ensure(client, ctx, path) {
    const have = await listed(client, ctx);
    if (have.has(path)) return true;
    await client.mailboxCreate(path);
    have.set(path, { path });
    return true;
  }

  /** The UIDs a change's messages have in `folder` at the server now: the one kept, or found by Message-ID. */
  async function uidsOf(client, accountId, change, folder = change.folder) {
    const sameValidity = !change.uidValidity || String(client.mailbox?.uidValidity) === change.uidValidity;
    const out = [];
    for (const item of change.items || []) {
      let uid = sameValidity ? item.uid ?? null : null;
      if (uid == null && item.messageId) {
        const found = await client.search({ header: { 'message-id': `<${item.messageId}>` } }, { uid: true });
        uid = Array.isArray(found) && found.length ? found[found.length - 1] : null;
      }
      if (uid != null) out.push({ item, uid });
    }
    return out;
  }

  /** Replay the line in order. Throws when the connection is lost; the rest wait. */
  async function replay(client, accountId) {
    const ctx = {};
    for (;;) {
      const [change] = store.changes(accountId).items;
      if (!change) return;
      let refused = null;
      try {
        refused = await apply(client, accountId, change, ctx);
      } catch (err) {
        if (lost(client, err)) throw err;
        refused = err.message || 'refused';
      }
      if (refused) {
        await putBack(client, accountId, change, ctx).catch((err) => { if (lost(client, err)) throw err; });
        broadcast?.('mail:syncProblem', { accountId, kind: change.kind, folder: change.folder, to: change.to || null, count: (change.items || []).length, reason: typeof refused === 'string' ? refused : null });
      }
      store.settleChange(accountId, change.n);
    }
  }

  /** One change made at the server: null when done, or why it was refused. */
  async function apply(client, accountId, change, ctx) {
    const no = (what) => (client.usable ? what : Promise.reject(new Error('The connection was lost.')));
    switch (change.kind) {
      case 'flags': {
        const lock = await client.getMailboxLock(change.folder).catch(() => null);
        if (!lock) return no('The folder is not at the server.');
        try {
          const found = await uidsOf(client, accountId, change);
          if (!found.length) return null;
          const set = uidSet(found.map((f) => f.uid));
          if (change.add?.length && (await client.messageFlagsAdd(set, change.add, { uid: true })) === false) return no('The server did not take the change.');
          if (change.remove?.length && (await client.messageFlagsRemove(set, change.remove, { uid: true })) === false) return no('The server did not take the change.');
          return null;
        } finally {
          lock.release();
        }
      }
      case 'move': {
        if (change.folder === change.to) return null;
        await ensure(client, ctx, change.to);
        const lock = await client.getMailboxLock(change.folder).catch(() => null);
        if (!lock) return no('The folder is not at the server.');
        let moved = null;
        let found = [];
        try {
          found = await uidsOf(client, accountId, change);
          if (!found.length) return null;
          moved = await client.messageMove(uidSet(found.map((f) => f.uid)), change.to, { uid: true });
          if (moved === false) return no('The server did not move them.');
        } finally {
          lock.release();
        }
        // The server's new UIDs, where it says them (UIDPLUS); otherwise the next fetch of that folder finds them by Message-ID.
        const map = moved?.uidMap instanceof Map ? moved.uidMap : null;
        const uids = {};
        for (const { item, uid } of found) uids[item.id] = map?.get(uid) ?? null;
        store.setUids(accountId, change.to, uids);
        // Junk and not junk say so to the server's filter too, where it takes keywords.
        if (change.junk != null && map?.size) {
          const lockTo = await client.getMailboxLock(change.to).catch(() => null);
          if (lockTo) {
            try {
              const set = uidSet([...map.values()]);
              await client.messageFlagsAdd(set, [change.junk ? '$Junk' : '$NotJunk'], { uid: true });
              await client.messageFlagsRemove(set, [change.junk ? '$NotJunk' : '$Junk'], { uid: true });
            } finally {
              lockTo.release();
            }
          }
        }
        return null;
      }
      case 'copy': {
        await ensure(client, ctx, change.to);
        // A copy already there (this change replayed after a crash) is not made twice.
        const there = new Set();
        const lockTo = await client.getMailboxLock(change.to);
        try {
          for (const item of change.items || []) {
            if (!item.messageId) continue;
            const hit = await client.search({ header: { 'message-id': `<${item.messageId}>` } }, { uid: true });
            if (Array.isArray(hit) && hit.length) there.add(item.id);
          }
        } finally {
          lockTo.release();
        }
        const lock = await client.getMailboxLock(change.folder).catch(() => null);
        if (!lock) return no('The folder is not at the server.');
        try {
          const found = (await uidsOf(client, accountId, change)).filter(({ item }) => !there.has(item.id));
          if (!found.length) return null;
          const copied = await client.messageCopy(uidSet(found.map((f) => f.uid)), change.to, { uid: true });
          if (copied === false) return no('The server did not copy them.');
          const map = copied?.uidMap instanceof Map ? copied.uidMap : null;
          store.setUids(accountId, change.to, Object.fromEntries(found.map(({ item, uid }) => [item.id, map?.get(uid) ?? null])));
          return null;
        } finally {
          lock.release();
        }
      }
      case 'expunge': {
        const lock = await client.getMailboxLock(change.folder).catch(() => null);
        if (!lock) return null; // the folder is gone at the server, and its messages with it
        try {
          const found = await uidsOf(client, accountId, change);
          if (!found.length) return null;
          if ((await client.messageDelete(uidSet(found.map((f) => f.uid)), { uid: true })) === false) return no('The server did not delete them.');
          return null;
        } finally {
          lock.release();
        }
      }
      case 'append': {
        let raw = null;
        try {
          // Sealed like the rest of the store in a profile with a password.
          raw = store.readOutgoing(change.file);
        } catch {
          return null; // nothing left to keep
        }
        await ensure(client, ctx, change.folder);
        const lock = await client.getMailboxLock(change.folder);
        let uid = null;
        try {
          // Already there (a server that keeps its own copy, or this change replayed after a crash): not twice.
          const found = change.messageId ? await client.search({ header: { 'message-id': `<${change.messageId}>` } }, { uid: true }) : null;
          if (Array.isArray(found) && found.length) uid = found[found.length - 1];
        } finally {
          lock.release();
        }
        if (uid == null) {
          const r = await client.append(change.folder, raw, ['\\Seen']);
          if (r === false) return no('The server did not keep the copy.');
          uid = r?.uid ?? null;
        }
        if (change.id) store.setUids(accountId, change.folder, { [change.id]: uid });
        fs.rmSync(change.file, { force: true });
        return null;
      }
      case 'create': {
        const have = await listed(client, ctx);
        if (have.has(change.folder)) return null;
        await client.mailboxCreate(change.folder);
        have.set(change.folder, { path: change.folder });
        return null;
      }
      case 'rename': {
        const have = await listed(client, ctx);
        if (!have.has(change.folder) && have.has(change.to)) return null;
        if (!have.has(change.folder)) return no('The folder is not at the server.');
        await client.mailboxRename(change.folder, change.to);
        ctx.listed = null;
        return null;
      }
      case 'deleteFolder': {
        const have = await listed(client, ctx);
        if (!have.has(change.folder)) return null;
        await client.mailboxDelete(change.folder);
        ctx.listed = null;
        return null;
      }
      default:
        return null;
    }
  }

  /** A change refused: what it touched made again as the server has it. */
  async function putBack(client, accountId, change, ctx) {
    const ids = (change.items || []).map((i) => i.id);
    switch (change.kind) {
      case 'flags': {
        const lock = await client.getMailboxLock(change.folder).catch(() => null);
        if (!lock) return;
        try {
          const found = await uidsOf(client, accountId, change);
          if (!found.length) return;
          const marks = {};
          for await (const m of client.fetch(uidSet(found.map((f) => f.uid)), { uid: true, flags: true }, { uid: true })) marks[m.uid] = marksOf(m.flags);
          for (const { item, uid } of found) if (marks[uid]) store.setFlags(accountId, change.folder, item.id, marks[uid]);
        } finally {
          lock.release();
        }
        return;
      }
      case 'move':
      case 'expunge': {
        // Gone from where it was put here; back where the server still keeps it.
        if (change.kind === 'move') store.removeMany(accountId, change.to, ids);
        const lock = await client.getMailboxLock(change.folder).catch(() => null);
        if (!lock) return;
        try {
          const found = await uidsOf(client, accountId, change);
          await download(client, accountId, change.folder, found.map((f) => f.uid));
        } finally {
          lock.release();
        }
        return;
      }
      case 'copy':
        store.removeMany(accountId, change.to, ids);
        return;
      case 'rename':
        store.renameFolder(accountId, change.to, change.folder, change.delimiter || '/');
        return;
      case 'create':
        if (!store.counts(accountId, change.folder).total) store.removeFolder(accountId, change.folder);
        return;
      default:
        // A Sent copy the server would not keep stays here; a folder it would not delete comes back on the next fetch.
        void ctx;
    }
  }

  /* ── fetching ─────────────────────────────────────────────────────── */

  /** Messages' UID, marks and Message-ID, without their bodies. */
  async function heads(client, range, byUid = true) {
    const out = [];
    for await (const m of client.fetch(range, { uid: true, flags: true, headers: ['message-id'] }, { uid: byUid })) {
      out.push({ uid: m.uid, flags: m.flags, messageId: messageIdOf(m.headers) });
    }
    return out;
  }

  /** Messages fetched whole and kept: { ids } of those new here, the rest given the UID the server has. */
  async function download(client, accountId, folder, uids) {
    const ids = [];
    let added = 0;
    for (let i = 0; i < uids.length; i += BATCH) {
      const batch = [];
      const known = {};
      for await (const msg of client.fetch(uidSet(uids.slice(i, i + BATCH)), { uid: true, flags: true, source: true }, { uid: true })) {
        const parsed = parseMessage(msg.source);
        parsed.uid = msg.uid;
        Object.assign(parsed, marksOf(msg.flags));
        const id = messageKey(parsed);
        if (store.get(accountId, folder, id)) known[id] = msg.uid;
        else batch.push(parsed);
      }
      if (Object.keys(known).length) store.setUids(accountId, folder, known);
      added += store.putMany(accountId, folder, batch, { ids });
    }
    return { ids, added };
  }

  /** Marks the server has that differ from the store's, applied here. */
  function applyMarks(accountId, folder, rowsByUid, changes) {
    const groups = new Map();
    for (const { uid, flags } of changes) {
      const row = rowsByUid.get(uid);
      if (!row) continue;
      const marks = marksOf(flags);
      if (row.unread === marks.unread && Boolean(row.flagged) === marks.flagged) continue;
      const key = JSON.stringify(marks);
      if (!groups.has(key)) groups.set(key, { marks, ids: [] });
      groups.get(key).ids.push(row.id);
    }
    let changed = 0;
    for (const { marks, ids } of groups.values()) changed += store.setFlagsMany(accountId, folder, ids, marks).length;
    return changed;
  }

  /**
   * One folder brought into step: what is new fetched by UID, marks and
   * presence compared over the newest messages. `quick` leaves a folder whose
   * STATUS says nothing changed (all a server without CONDSTORE can say).
   */
  async function syncFolder(client, accountId, folder, { limit, quick, condstore }) {
    const prev = store.syncState(accountId, folder);
    if (prev) {
      const st = await client.status(folder, { messages: true, uidNext: true, uidValidity: true, ...(condstore ? { highestModseq: true } : {}) });
      const same = st && String(st.uidValidity) === prev.uidValidity && st.uidNext === prev.uidNext && st.messages === prev.messages;
      if (same && (condstore ? String(st.highestModseq) === String(prev.highestModseq) : quick)) return { folder, ids: [], added: 0, skipped: true };
    }
    const lock = await client.getMailboxLock(folder);
    try {
      const mb = client.mailbox;
      const snapshot = { uidValidity: String(mb.uidValidity), uidNext: mb.uidNext, highestModseq: condstore && mb.highestModseq ? String(mb.highestModseq) : null, messages: mb.exists };
      const result = !prev || prev.uidValidity !== snapshot.uidValidity
        ? await adopt(client, accountId, folder, { prev, limit, exists: mb.exists })
        : await incremental(client, accountId, folder, { prev, limit, condstore, snapshot });
      store.setSyncState(accountId, folder, { version: 1, ...snapshot, lastSeen: result.lastSeen });
      return { folder, ids: result.ids, added: result.added, changed: result.changed || 0 };
    } finally {
      lock.release();
    }
  }

  /** A folder already matched with its server: new by UID, then marks and presence. */
  async function incremental(client, accountId, folder, { prev, limit, condstore, snapshot }) {
    const lastSeen = prev.lastSeen || 0;
    let top = lastSeen;
    let ids = [];
    let added = 0;
    let changed = 0;
    if (snapshot.uidNext > lastSeen + 1 && snapshot.messages) {
      const fresh = (await heads(client, `${lastSeen + 1}:*`)).filter((h) => h.uid > lastSeen);
      const local = new Set(store.list(accountId, folder, { limit: 1e7 }).rows.map((r) => r.id));
      const known = {};
      const need = [];
      for (const h of fresh) {
        top = Math.max(top, h.uid);
        const key = keyOf(h.messageId);
        // Moved here from this computer: it is here already, and now has its UID.
        if (key && local.has(key)) known[key] = h.uid;
        else need.push(h.uid);
      }
      if (Object.keys(known).length) store.setUids(accountId, folder, known);
      ({ ids, added } = await download(client, accountId, folder, need.slice(-limit)));
    }
    // Marks and presence over the newest FETCH_WINDOW already here.
    const rows = store.list(accountId, folder, { limit: 1e7 }).rows.filter((r) => r.uid != null && r.uid <= lastSeen).sort((a, b) => a.uid - b.uid).slice(-FETCH_WINDOW);
    if (rows.length) {
      const range = `${rows[0].uid}:${lastSeen}`;
      const byUid = new Map(rows.map((r) => [r.uid, r]));
      let changes = [];
      if (condstore && prev.highestModseq && snapshot.highestModseq) {
        if (snapshot.highestModseq !== String(prev.highestModseq)) {
          for await (const m of client.fetch(range, { uid: true, flags: true }, { uid: true, changedSince: BigInt(prev.highestModseq) })) changes.push({ uid: m.uid, flags: m.flags });
        }
      } else {
        for await (const m of client.fetch(range, { uid: true, flags: true }, { uid: true })) changes.push({ uid: m.uid, flags: m.flags });
      }
      changed += applyMarks(accountId, folder, byUid, changes);
      const present = await client.search({ uid: range }, { uid: true });
      if (Array.isArray(present)) {
        const have = new Set(present);
        const gone = rows.filter((r) => !have.has(r.uid)).map((r) => r.id);
        if (gone.length) { store.removeMany(accountId, folder, gone); changed += gone.length; }
      }
    }
    return { ids, added, changed, lastSeen: top };
  }

  /**
   * A folder matched with its server by Message-ID: the first time (a store
   * from before 2.19, or a folder never fetched), or after its UIDVALIDITY
   * changed. Nothing here is doubled and nothing kept here is lost: a message
   * the server does not have stays, with no UID, unless the server listed
   * the whole folder under a new UIDVALIDITY and it is not there.
   */
  async function adopt(client, accountId, folder, { prev, limit, exists }) {
    const rows = store.list(accountId, folder, { limit: 1e7 }).rows;
    const complete = exists <= ADOPT_LIMIT;
    const listedHeads = exists ? await heads(client, complete ? '1:*' : `${exists - ADOPT_LIMIT + 1}:*`, !complete ? false : true) : [];
    const local = new Map(rows.map((r) => [r.id, r]));
    const uids = {};
    const matched = new Set();
    const unknown = [];
    let top = 0;
    const marks = [];
    for (const h of listedHeads) {
      top = Math.max(top, h.uid);
      const key = keyOf(h.messageId);
      if (key && local.has(key) && !matched.has(key)) {
        uids[key] = h.uid;
        matched.add(key);
        marks.push({ uid: h.uid, flags: h.flags });
      } else unknown.push(h);
    }
    const gone = [];
    for (const r of rows) {
      if (matched.has(r.id)) continue;
      if (prev && complete && r.uid != null) gone.push(r.id);
      else uids[r.id] = null;
    }
    store.setUids(accountId, folder, uids);
    if (gone.length) store.removeMany(accountId, folder, gone);
    const byUid = new Map(store.list(accountId, folder, { limit: 1e7 }).rows.filter((r) => r.uid != null).map((r) => [r.uid, r]));
    const changed = gone.length + applyMarks(accountId, folder, byUid, marks);

    let want = unknown;
    const legacy = !prev ? rows.filter((r) => r.uid != null).map((r) => r.uid) : [];
    if (legacy.length) {
      // A store from before 2.19 fetched these by position and kept their UIDs.
      // One it fetched that is no longer here was deleted or filed here, which
      // that version never told the server: it is not fetched again, and one
      // filed in another folder here is filed there at the server now.
      const lo = Math.min(...legacy);
      const hi = Math.max(...legacy);
      const elsewhere = new Map();
      for (const f of store.folders(accountId)) {
        if (f.path === folder) continue;
        for (const r of store.list(accountId, f.path, { limit: 1e7 }).rows) if (!elsewhere.has(r.id)) elsewhere.set(r.id, f.path);
      }
      const moves = new Map();
      for (const h of unknown) {
        if (h.uid < lo || h.uid > hi) continue;
        const to = elsewhere.get(keyOf(h.messageId));
        if (!to) continue;
        if (!moves.has(to)) moves.set(to, []);
        moves.get(to).push({ id: keyOf(h.messageId), uid: h.uid, messageId: h.messageId });
      }
      for (const [to, items] of moves) store.addChange(accountId, { kind: 'move', folder, to, items, uidValidity: String(client.mailbox.uidValidity) });
      want = unknown.filter((h) => h.uid > hi);
    }
    const fetched = await download(client, accountId, folder, want.slice(-limit).map((h) => h.uid));
    // Matched again after a new UIDVALIDITY, nothing has arrived: no rule or reply runs over it.
    return { ids: prev ? [] : fetched.ids, added: fetched.added, changed, lastSeen: top };
  }

  /** What the store keeps of the server's folder list: theirs, and ours with no server copy that still hold mail. */
  async function folders(client, accountId, ctx) {
    const pathsOf = (rows) => rows.map((f) => f.path).sort().join('\u0001');
    const had = pathsOf(store.folders(accountId));
    const list = await client.list();
    ctx.listed = new Map(list.map((f) => [f.path, f]));
    const server = list.filter((f) => !f.flags?.has?.('\\Noselect') && !f.flags?.has?.('\\NonExistent')).map((f) => describe(f));
    const paths = new Set(server.map((f) => f.path));
    const kept = [];
    for (const f of store.folders(accountId)) {
      if (paths.has(f.path)) continue;
      // Matched with the server before and gone from it: deleted or renamed elsewhere.
      if (store.syncState(accountId, f.path)) store.removeFolder(accountId, f.path);
      else if (store.counts(accountId, f.path).total) kept.push({ ...f, local: true });
    }
    const all = [...server, ...kept];
    // Made, renamed or deleted elsewhere: the window shows the list again.
    if (had !== pathsOf(all)) ctx.foldersChanged = true;
    store.setFolders(accountId, all);
    return all;
  }

  /**
   * Fetch an account: the line replayed, the folder list read, then the
   * Inbox, `folder` if one is named, and every folder fetched before (those
   * whose STATUS shows no change are left). Two asks while one is waiting
   * share it; never two at once.
   */
  function fetch(accountId, { folder = null, limit = 300 } = {}) {
    const queued = waiting.get(accountId);
    if (queued) {
      if (folder) queued.folders.add(folder);
      return queued.promise;
    }
    const entry = { folders: new Set(folder ? [folder] : []), limit };
    waiting.set(accountId, entry);
    entry.promise = serial(accountId, () => {
      waiting.delete(accountId);
      return runFetch(accountId, entry);
    });
    return entry.promise;
  }

  async function runFetch(accountId, { folders: asked, limit }) {
    const account = liveAccount(accountId);
    if (!account) return { added: 0, arrived: [], local: true };
    let client = null;
    try {
      client = await connect(account);
    } catch (err) {
      note(accountId, { offline: true, error: err.message });
      throw err;
    }
    try {
      await replay(client, accountId);
      const ctx = {};
      seen?.(accountId, client.capabilities);
      const all = await folders(client, accountId, ctx);
      const condstore = Boolean(client.enabled?.has?.('CONDSTORE') || client.capabilities?.has?.('CONDSTORE'));
      const inbox = all.find((f) => f.role === 'inbox')?.path || 'INBOX';
      const named = [...asked].filter((p) => ctx.listed.has(p));
      const before = all.filter((f) => !f.local && f.path !== inbox && !named.includes(f.path) && store.syncState(accountId, f.path)).map((f) => f.path);
      const targets = [{ path: inbox, quick: false }, ...named.filter((p) => p !== inbox).map((p) => ({ path: p, quick: false })), ...before.map((p) => ({ path: p, quick: true }))];
      const arrived = [];
      let added = 0;
      let changed = 0;
      for (const { path, quick } of targets) {
        if (!ctx.listed.has(path)) continue;
        const r = await syncFolder(client, accountId, path, { limit, quick, condstore });
        added += r.added;
        changed += r.changed || 0;
        if (r.ids.length) arrived.push({ folder: path, ids: r.ids });
        if (r.added) broadcast?.('mail:progress', { accountId, folder: path, done: r.added, total: r.added, phase: 'fetch' });
      }
      if (ctx.foldersChanged) changed += 1;
      const after = onArrived ? await onArrived(accountId, arrived) : {};
      // What arriving did (junk filed, rules run) goes to the server on this connection.
      await replay(client, accountId);
      note(accountId, { offline: false, error: null, at: new Date().toISOString() });
      return { added, changed, arrived, ...after };
    } catch (err) {
      if (lost(client, err)) {
        later(accountId);
        note(accountId, { offline: true, error: err.message });
      }
      throw err;
    } finally {
      await client.logout().catch(() => client.close?.());
    }
  }

  /* ── listening at the server ──────────────────────────────────────── */

  /**
   * Hold a connection open on the Inbox and fetch when the server says
   * something changed there (IDLE, RFC 2177; ImapFlow idles on its own once
   * a folder is open and nothing else is asked). A server without IDLE is
   * not watched: the timer's fetches cover it. A dropped connection is
   * forgotten, and the next timed fetch watches again.
   */
  async function watch(accountId, onChange) {
    const account = liveAccount(accountId);
    if (!account || watchers.has(accountId)) return Boolean(watchers.get(accountId));
    watchers.set(accountId, null);
    let client = null;
    try {
      // Listening starts at once, and is begun again every few minutes so nothing between drops it.
      client = await connect(account, { autoIdleDelay: 500, maxIdleTime: 4 * 60_000 });
      if (!client.capabilities?.has?.('IDLE')) {
        await client.logout().catch(() => {});
        watchers.delete(accountId);
        return false;
      }
      await client.mailboxOpen('INBOX');
      let timer = null;
      const changed = () => {
        clearTimeout(timer);
        timer = setTimeout(() => onChange(accountId), 400);
        timer.unref?.();
      };
      client.on('exists', changed);
      client.on('expunge', changed);
      client.on('flags', changed);
      client.on('close', () => { if (watchers.get(accountId) === client) watchers.delete(accountId); });
      client.on('error', () => {});
      watchers.set(accountId, client);
      return true;
    } catch {
      watchers.delete(accountId);
      try { client?.close(); } catch { /* never opened */ }
      return false;
    }
  }

  async function unwatch(accountId) {
    const client = watchers.get(accountId);
    watchers.delete(accountId);
    if (client) await client.logout().catch(() => client.close?.());
  }

  return {
    record,
    itemsOf,
    kick,
    flush,
    fetch,
    watch,
    unwatch,
    watching: (accountId) => Boolean(watchers.get(accountId)),
    async stop() {
      for (const t of [...kicks.values(), ...retries.values()]) clearTimeout(t);
      kicks.clear();
      retries.clear();
      await Promise.all([...watchers.keys()].map((id) => unwatch(id)));
    },
    state: (accountId) => ({ ...(states.get(accountId) || {}), waiting: store.changes(accountId).items.length }),
  };
}
