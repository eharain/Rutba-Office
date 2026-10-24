// Synced copies and other sources, the service's side (D-2026-10-22-13,
// steps 4 and 5). A table kept as a copy catches up with its engine as the
// database opens, every 15 seconds while it is on screen and on Sync now:
// the changes made here go to the engine by key first, then the engine's
// records are read and compared by key and hash (packages/database
// sync.js). An engine out of reach refuses nothing: the changes wait, and
// the table's bar says how many and since when. A source (a file, another
// database, a web address answering JSON) is read again on open, on its
// schedule and on Refresh, its records replaced; a changed set of columns
// changes the table's design, and the bar says so.

import {
  SOURCE_KINDS, SOURCE_EVERY, DEFAULT_EVERY, prepareSync, pushed, refused, reconcile, replaceRows,
  pending, clashList, clashCount, dropClash, markSynced, stampMax, copyReadOnly,
} from '@rutba/database';
import { readFileSource, readRbaseSource, readJsonSource, rbaseItems } from './base-sources.js';

/** Most records a copy reads at once; past it, an updated-at field reads only what changed. */
const MOST = 5_000_000;

/** A failure that says the engine is out of reach, not that it refused a record. */
export function unreachable(err) {
  const m = `${err?.code || ''} ${err?.message || err || ''}`;
  return Boolean(err?.unreachable) || /ECONNREFUSED|ECONNRESET|ETIMEDOUT|EPIPE|ENOTFOUND|EHOSTUNREACH|ENETUNREACH|could not be reached|Connection terminated|not queryable|connection is in closed state|PROTOCOL_CONNECTION_LOST|Client was closed|timeout/i.test(m);
}

/** A source's secret's place in the protected store: its password, or one of its headers. */
export const sourceSecret = (linkId, header = null) => (header ? `base-source:${linkId}:${header}` : `base-source:${linkId}`);

/** A source's settings as the file keeps them: nothing secret in them. */
export function cleanSourceSettings(kind, s = {}) {
  const every = SOURCE_EVERY.includes(Number(s.every)) ? Number(s.every) : DEFAULT_EVERY;
  const where = String(s.where || '').trim();
  if (kind === 'file') return { where, sheet: s.sheet || null, header: Math.max(1, Number(s.header) || 1), every };
  if (kind === 'rbase') return { where, person: s.person ? String(s.person) : null, every };
  if (kind === 'json') {
    const headers = (s.headers || []).filter((h) => String(h?.name || '').trim()).map((h) => (h.secret ? { name: String(h.name).trim(), secret: true } : { name: String(h.name).trim(), value: String(h.value ?? '') }));
    const p = s.paging || {};
    const paging = p.kind === 'next' ? { kind: 'next', next: String(p.next || '').trim() || null } : p.kind === 'page' ? { kind: 'page', param: String(p.param || 'page').trim() || 'page', start: Number.isFinite(Number(p.start)) ? Number(p.start) : 1 } : { kind: 'none' };
    return { where, headers, records: String(s.records || '').trim(), paging, every };
  }
  throw new Error(`${kind} is not a source Data Files reads.`);
}

/**
 * The sync side of the base service, given its windows (`get`), how it
 * reaches an engine (`engineFor`, `dropEngine`) and the protected store.
 */
export function createSync({ engineFor, dropEngine, secrets }) {
  const linkOf = (o, l) => o.base.links().find((x) => x.id === l.link) || null;
  const entryOf = (o, table) => o.base.linked().find((x) => x.table === table) || null;
  const kindOf = (o, l) => linkOf(o, l)?.engine || null;
  const isSourceTable = (o, l) => SOURCE_KINDS.has(kindOf(o, l));
  const stateOf = (o, table) => o.state.get(table) || {};
  const setState = (o, table, patch) => o.state.set(table, { ...stateOf(o, table), ...patch });
  // One catch-up at a time for each table: a second asked for meanwhile waits for the first.
  const once = (o, table, run) => {
    o.busy ||= new Map();
    if (o.busy.has(table)) return o.busy.get(table);
    const p = run().finally(() => o.busy.delete(table));
    o.busy.set(table, p);
    return p;
  };

  /** A copy caught up with its engine: what was changed here written by key, then the engine's records compared. */
  const syncCopy = (o, table) => once(o, table, async () => {
    const base = o.base;
    const l = entryOf(o, table);
    if (!l || l.mode !== 'copy') return;
    const link = linkOf(o, l);
    try {
      if (!link) throw new Error('Its link is gone.');
      const engine = await engineFor(o, link);
      const design = await engine.describe(l.remote);
      const diff = base.relinkDesign(table, design);
      const key = design.key;
      const fields = base.fields(table).map(({ readOnly, ...f }) => f);
      const keyType = fields.find((f) => f.name === key)?.type || 'text';
      const opts = { fields, key };
      const read = async () => {
        if (l.stamp && key) {
          const last = stampMax(base, table, l.stamp);
          if (last != null) {
            const rows = await engine.read(l.remote, { fields: design.fields, key, since: { field: l.stamp, value: last }, limit: MOST });
            const keys = await engine.keys(l.remote, { key, type: keyType });
            return { rows, keys, complete: true };
          }
        }
        const rows = await engine.read(l.remote, { fields: design.fields, key, limit: MOST });
        return { rows, complete: rows.length < MOST };
      };
      let snap = await read();
      let changed = false;
      let clashes = 0;
      if (!key) {
        changed = replaceRows(base, table, snap.rows, fields);
      } else {
        const prep = prepareSync(base, table, snap.rows, opts, snap.keys || null);
        clashes = prep.clashes;
        let stopped = null;
        for (const p of prep.pushes) {
          try {
            if (p.action === 'added') {
              const k = await engine.insert(l.remote, { key }, p.values);
              if (pushed(base, table, p, opts, k) === 'gone' && k != null) await engine.remove(l.remote, { key }, [k]);
            } else if (p.action === 'changed') {
              await engine.update(l.remote, { key }, p.key, p.values);
              pushed(base, table, p, opts);
            } else {
              await engine.remove(l.remote, { key }, [p.key]);
              pushed(base, table, p, opts);
            }
          } catch (err) {
            if (unreachable(err)) { stopped = err; break; }
            // The engine would not take it: the engine wins, and the version from here is kept to see.
            refused(base, table, p, opts, String(err.message || err));
            clashes++;
          }
        }
        if (stopped) throw stopped;
        if (prep.pushes.length || clashes) snap = await read();
        const r = reconcile(base, table, snap, opts);
        changed = Boolean(r.added || r.changed || r.removed || prep.pushes.length || clashes);
      }
      markSynced(base, table);
      const at = new Date().toISOString();
      setState(o, table, { error: null, at: changed || diff || !stateOf(o, table).at ? at : stateOf(o, table).at, note: diff || stateOf(o, table).note || null, out: null });
    } catch (err) {
      if (link && unreachable(err)) dropEngine(o, link);
      setState(o, table, { error: err.message || String(err), out: unreachable(err) });
    }
  });

  /** A source's records read again, its table's design following any change of its columns. */
  const readSource = async (link, l) => {
    const s = link.settings || {};
    if (link.engine === 'file') return readFileSource({ where: s.where, sheet: l?.remote?.sheet ?? s.sheet, header: s.header });
    if (link.engine === 'rbase') return readRbaseSource({ where: s.where, item: l?.remote?.item }, { person: s.person, password: secrets?.get?.(sourceSecret(link.id)) || null });
    if (link.engine === 'json') return readJsonSource(s, { secretOf: (name) => secrets?.get?.(sourceSecret(link.id, name)) || null });
    throw new Error(`${link.engine} is not a source Data Files reads.`);
  };

  const refreshSource = (o, table) => once(o, table, async () => {
    const base = o.base;
    const l = entryOf(o, table);
    const link = l ? linkOf(o, l) : null;
    if (!link) return;
    try {
      const got = await readSource(link, l);
      const diff = base.relinkDesign(table, { fields: got.fields, key: null });
      const changed = replaceRows(base, table, got.rows, base.fields(table).map(({ readOnly, ...f }) => f));
      markSynced(base, table);
      const at = new Date().toISOString();
      setState(o, table, { error: null, at: changed || diff || !stateOf(o, table).at ? at : stateOf(o, table).at, note: diff || null, noteAt: diff ? at : stateOf(o, table).noteAt || null });
    } catch (err) {
      setState(o, table, { error: err.message || String(err), who: Boolean(err.who) });
    }
  });

  /** Whether a source's table is due to be read again on its schedule. */
  const due = (o, l, at = Date.now()) => {
    const every = Number(linkOf(o, l)?.settings?.every ?? DEFAULT_EVERY);
    if (!every) return false;
    const last = Date.parse(l.synced || '') || 0;
    return at - last >= every * 60_000 - 1000;
  };

  /** A linked table's part of the summary: how it is kept, when it last caught up, what waits and what clashed. */
  const describe = (o, t) => {
    const base = o.base;
    const l = entryOf(o, t.name);
    const st = stateOf(o, t.name);
    if (!l || l.mode !== 'copy') return null;
    const link = linkOf(o, l);
    const source = SOURCE_KINDS.has(link?.engine) ? link.engine : null;
    return {
      count: base.count(t.name),
      linkError: st.error || null,
      readAt: st.at || l.synced || null,
      key: l.key || null,
      mode: 'copy',
      synced: l.synced || null,
      waiting: source ? { count: 0, since: null } : pending(base, t.name),
      out: Boolean(st.out),
      clashes: source ? 0 : clashCount(base, t.name),
      stamp: l.stamp || null,
      source,
      every: source ? Number(link.settings?.every ?? DEFAULT_EVERY) : null,
      note: st.note || null,
      who: Boolean(st.who),
      readOnly: Boolean(copyReadOnly(base, l)),
    };
  };

  return {
    syncCopy,
    refreshSource,
    isSourceTable,
    describe,
    due,
    entryOf,
    linkOf,
    /** Every source due on its schedule read again; answers whether any was. */
    async refreshDue(o) {
      if (o.base.access?.locked || o.base.rights()) return false;
      const list = o.base.linked().filter((l) => l.mode === 'copy' && isSourceTable(o, l) && due(o, l));
      await Promise.all(list.map((l) => refreshSource(o, l.table)));
      return list.length > 0;
    },
    /** What another Data Files database offers, for the Link box: its tables and queries, as the person named may see them. */
    rbaseItems: (where, who) => rbaseItems(String(where || '').trim(), who),
    /** A source tried before it is linked: its fields and first records, nothing kept. */
    async trySource(kind, settings, { password = null, headers = {}, item = null } = {}) {
      const s = cleanSourceSettings(kind, settings);
      let got;
      let extra = {};
      if (kind === 'file') {
        got = await readFileSource({ where: s.where, sheet: s.sheet, header: s.header });
        extra = { sheets: got.sheets, sheet: got.sheet, kind: got.kind };
      } else if (kind === 'rbase') {
        const listed = rbaseItems(s.where, { person: s.person, password });
        extra = { items: listed.items, protected: listed.protected };
        const pick = item || listed.items[0] || null;
        if (!pick) return { fields: [], rows: [], count: 0, ...extra };
        got = readRbaseSource({ where: s.where, item: pick }, { person: s.person, password });
        extra.item = pick;
      } else {
        got = await readJsonSource(s, { secretOf: (name) => headers[name] ?? null });
        extra = { pages: got.pages };
      }
      return { fields: got.fields, rows: got.rows.slice(0, 5), count: got.rows.length, ...extra };
    },
    readSource,
    clashes: (o, table) => clashList(o.base, table),
    /**
     * A clash settled: the engine's version kept as it is ('engine'), or the
     * version from here put back ('mine'), to go to the engine at the next
     * catch-up as any change made here does.
     */
    settleClash(o, table, id, keep) {
      const c = clashList(o.base, table).find((x) => x.id === id);
      if (!c) return false;
      if (keep === 'mine' && c.mine) {
        const l = entryOf(o, table);
        const fields = o.base.fields(table);
        const values = Object.fromEntries(fields.filter((f) => Object.hasOwn(c.mine, f.name)).map((f) => [f.name, c.mine[f.name]]));
        const row = l.key && c.key != null ? o.base.rows(table, { where: { [l.key]: c.key }, limit: 1 })[0] : null;
        if (row) {
          delete values[l.key];
          o.base.update(table, row._rowid, values);
        } else {
          if (fields.find((f) => f.name === l.key)?.type === 'autonumber') delete values[l.key];
          o.base.insert(table, values);
        }
      } else if (keep === 'mine' && c.action === 'deleted') {
        // Deleted here, changed there: deleted again.
        const l = entryOf(o, table);
        const row = l.key && c.key != null ? o.base.rows(table, { where: { [l.key]: c.key }, limit: 1 })[0] : null;
        if (row) o.base.remove(table, [row._rowid]);
      }
      dropClash(o.base, table, id);
      return true;
    },
  };
}
