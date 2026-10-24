// Synced copies and other sources (D-2026-10-22-13, steps 4 and 5). A table
// linked to a database engine can be kept as a copy in the file instead of
// read live: its records are a real table here, so the database opens and
// works with no engine at all, and each record's engine key and a hash of
// the engine's version last seen are kept in rutba_sync with whether it was
// changed, added or deleted here since. Catching up writes the changes made
// here by key first, then reads the engine's records and compares them by
// key and hash. When both changed the same record the engine wins, and the
// version from here is kept in rutba_clashes, shown, never lost without a
// word. Other sources (a CSV or spreadsheet file, another Data Files
// database, a web address answering JSON) are kept as copies the same way,
// read only: each read replaces their records.
//
// Nothing here is made in a file until it has a copy: a database with no
// links has no rutba_sync, and opens and works exactly as before.

import crypto from 'node:crypto';
import { q } from './types.js';

/** The kinds of source a table can be read from besides an engine, kept in rutba_links.engine. */
export const SOURCE_KINDS = new Set(['file', 'rbase', 'json']);

/** How often a source is read again, in minutes; 0 by hand. Every 15 minutes unless the person chose. */
export const SOURCE_EVERY = [5, 15, 60, 0];
export const DEFAULT_EVERY = 15;

const SYNC_DESIGN = `
CREATE TABLE IF NOT EXISTS rutba_sync (tbl TEXT NOT NULL COLLATE NOCASE, rid INTEGER NOT NULL, kj TEXT, hash TEXT, state TEXT, n INTEGER NOT NULL DEFAULT 0, since TEXT, PRIMARY KEY (tbl, rid));
CREATE INDEX IF NOT EXISTS rutba_sync_key ON rutba_sync (tbl, kj);
CREATE TABLE IF NOT EXISTS rutba_clashes (id INTEGER PRIMARY KEY, tbl TEXT NOT NULL COLLATE NOCASE, kj TEXT, at TEXT NOT NULL, action TEXT NOT NULL, mine TEXT, theirs TEXT, why TEXT);
`;

/** The design a copy needs, given to a file the first time it keeps one; an older rutba_linked gains its columns. */
export function ensureSync(base) {
  base.db.exec(SYNC_DESIGN);
  const cols = new Set(base.db.prepare('PRAGMA table_info(rutba_linked)').all().map((c) => c.name));
  for (const [c, def] of [['mode', "TEXT NOT NULL DEFAULT 'live'"], ['stamp', 'TEXT'], ['synced', 'TEXT']]) {
    if (!cols.has(c)) base.db.exec(`ALTER TABLE rutba_linked ADD COLUMN ${c} ${def}`);
  }
}

const hasSync = (base) => base.hasDesignTable('rutba_sync');
const now = () => new Date().toISOString();

/** A key as the file keeps it to compare: its JSON. */
export const kjOf = (v) => JSON.stringify(v === undefined ? null : v);

/** A value as a hash sees it: the same record from the engine and from the file hash alike. */
const norm = (v) => {
  if (v === null || v === undefined) return null;
  if (typeof v === 'number') return String(Number.isInteger(v) ? v : Number(v.toPrecision(15)));
  if (typeof v === 'boolean') return v ? '1' : '0';
  return String(v);
};

/** A hash of one record's values, in its fields' order. */
export function hashRecord(fields, row) {
  return crypto.createHash('sha1').update(JSON.stringify(fields.map((f) => norm(row?.[f.name])))).digest('hex');
}

/** A hash of a whole set of records, for a source read again with nothing changed. */
export function hashRows(fields, rows) {
  const h = crypto.createHash('sha1');
  h.update(JSON.stringify(fields.map((f) => [f.name, f.type])));
  for (const r of rows) h.update(JSON.stringify(fields.map((f) => norm(r?.[f.name]))));
  return h.digest('hex');
}

const AFFINITY = { autonumber: 'INTEGER', integer: 'INTEGER', number: 'REAL', currency: 'REAL', yesno: 'INTEGER', text: 'TEXT', longtext: 'TEXT', date: 'TEXT', datetime: 'TEXT' };

/**
 * A copy's table: its fields as columns with their affinity and no checks
 * (the records are the engine's or the source's), a whole-number key its
 * rowid, so a record keeps its place, and any other key unique.
 */
export function copyTableSql(table, fields, key) {
  const keyField = fields.find((f) => f.name === key) || null;
  const intKey = Boolean(keyField && ['autonumber', 'integer'].includes(keyField.type));
  const cols = fields.map((f) => {
    if (f === keyField) return intKey ? `${q(f.name)} INTEGER PRIMARY KEY` : `${q(f.name)} ${AFFINITY[f.type] || 'TEXT'} UNIQUE`;
    return `${q(f.name)} ${AFFINITY[f.type] || 'TEXT'}`;
  });
  return `CREATE TABLE ${q(table)} (${cols.join(', ')})`;
}

const intKeyOf = (fields, key) => {
  const f = fields.find((x) => x.name === key);
  return Boolean(f && ['autonumber', 'integer'].includes(f.type));
};

/** Whether a linked table is a source's (read only), by its link's kind. */
export function isSource(base, l) {
  const link = base.db.prepare('SELECT engine FROM rutba_links WHERE id = ?').get(l.link);
  return Boolean(link && SOURCE_KINDS.has(link.engine));
}

/** Why a copy cannot be changed here, or null when it can. */
export function copyReadOnly(base, l) {
  if (isSource(base, l)) return `${l.table} is read from its source: it can be read here, not changed.`;
  if (!l.key) return `${l.table} has no primary key in the engine: it can be read, not changed.`;
  return null;
}

/** A change to a copy's record may not touch its key: that is the engine's. */
export function checkCopyChange(base, l, rowid, values) {
  if (!l.key || !Object.hasOwn(values, l.key)) return;
  const row = base.db.prepare(`SELECT ${q(l.key)} AS k FROM main.${q(l.table)} WHERE rowid = ?`).get(rowid);
  const v = values[l.key];
  if (row && v !== undefined && String(v ?? '') !== String(row.k ?? '')) throw new Error(`The ${l.key} of a record kept as a copy is the engine's: it cannot be changed here.`);
}

/** A record of a copy added, changed or deleted here: kept to be written to the engine at the next catch-up. */
export function trackWrite(base, table, action, rids) {
  if (!hasSync(base) || !rids.length) return;
  const l = base.linked().find((x) => x.table === table);
  if (!l || l.mode !== 'copy' || !l.key) return;
  const at = now();
  const get = base.db.prepare('SELECT state, kj FROM rutba_sync WHERE tbl = ? AND rid = ?');
  for (const rid of rids) {
    const was = get.get(table, rid);
    if (action === 'add') {
      const row = base.db.prepare(`SELECT ${q(l.key)} AS k FROM main.${q(table)} WHERE rowid = ?`).get(rid);
      base.db.prepare(`INSERT INTO rutba_sync (tbl, rid, kj, hash, state, n, since) VALUES (?, ?, ?, NULL, 'added', 1, ?)
        ON CONFLICT(tbl, rid) DO UPDATE SET kj = excluded.kj, hash = NULL, state = 'added', n = n + 1, since = COALESCE(since, excluded.since)`).run(table, rid, kjOf(row?.k ?? null), at);
    } else if (action === 'change') {
      if (!was) {
        const row = base.db.prepare(`SELECT ${q(l.key)} AS k FROM main.${q(table)} WHERE rowid = ?`).get(rid);
        base.db.prepare("INSERT INTO rutba_sync (tbl, rid, kj, hash, state, n, since) VALUES (?, ?, ?, NULL, 'changed', 1, ?)").run(table, rid, kjOf(row?.k ?? null), at);
      } else {
        base.db.prepare("UPDATE rutba_sync SET state = CASE WHEN state IN ('added', 'deleted') THEN state ELSE 'changed' END, n = n + 1, since = COALESCE(since, ?) WHERE tbl = ? AND rid = ?").run(at, table, rid);
      }
    } else if (action === 'delete') {
      // One that never reached the engine is simply forgotten.
      if (was?.state === 'added') base.db.prepare('DELETE FROM rutba_sync WHERE tbl = ? AND rid = ?').run(table, rid);
      // Kept under a rowid below any record's, so a record added next in its place has its own entry.
      else if (was) {
        const low = base.db.prepare('SELECT MIN(MIN(rid), 0) - 1 AS r FROM rutba_sync WHERE tbl = ?').get(table).r;
        base.db.prepare("UPDATE rutba_sync SET state = 'deleted', rid = ?, n = n + 1, since = COALESCE(since, ?) WHERE tbl = ? AND rid = ?").run(low, at, table, rid);
      }
    }
  }
}

/** How many changes made here wait to go to the engine, and since when: { count, since }. */
export function pending(base, table) {
  if (!hasSync(base)) return { count: 0, since: null };
  const r = base.db.prepare('SELECT count(*) AS n, min(since) AS since FROM rutba_sync WHERE tbl = ? AND state IS NOT NULL').get(table);
  return { count: r.n, since: r.n ? r.since : null };
}

/** The versions from here the engine's replaced: [{ id, key, at, action, mine, theirs, why }], the latest first. */
export function clashList(base, table) {
  if (!hasSync(base)) return [];
  return base.db.prepare('SELECT id, kj, at, action, mine, theirs, why FROM rutba_clashes WHERE tbl = ? ORDER BY id DESC').all(table)
    .map((r) => ({ id: r.id, key: r.kj == null ? null : JSON.parse(r.kj), at: r.at, action: r.action, mine: r.mine ? JSON.parse(r.mine) : null, theirs: r.theirs ? JSON.parse(r.theirs) : null, why: r.why || null }));
}

export function clashCount(base, table) {
  if (!hasSync(base)) return 0;
  return base.db.prepare('SELECT count(*) AS n FROM rutba_clashes WHERE tbl = ?').get(table).n;
}

export function dropClash(base, table, id) {
  if (!hasSync(base)) return 0;
  return Number(base.db.prepare('DELETE FROM rutba_clashes WHERE tbl = ? AND id = ?').run(table, id).changes);
}

const keepClash = (base, table, kj, action, mine, theirs, why = null) => {
  base.db.prepare('INSERT INTO rutba_clashes (tbl, kj, at, action, mine, theirs, why) VALUES (?, ?, ?, ?, ?, ?, ?)')
    .run(table, kj, now(), action, mine ? JSON.stringify(mine) : null, theirs ? JSON.stringify(theirs) : null, why);
};

/** The fields of a record as plain values: { field: value }. */
const valuesOf = (fields, row) => Object.fromEntries(fields.map((f) => [f.name, row?.[f.name] ?? null]));

const inTransaction = (base, run) => {
  base.db.exec('BEGIN IMMEDIATE');
  try {
    const out = run();
    base.db.exec('COMMIT');
    return out;
  } catch (err) {
    base.db.exec('ROLLBACK');
    throw err;
  }
};

/** A record put into a copy from the engine: at its key's rowid for a whole-number key. Answers its rowid. */
function putRow(base, table, fields, key, row) {
  const cols = fields.map((f) => q(f.name)).join(', ');
  const marks = fields.map(() => '?').join(', ');
  const out = base.db.prepare(`INSERT INTO main.${q(table)} (${cols}) VALUES (${marks})`).run(...fields.map((f) => (row[f.name] === undefined ? null : row[f.name])));
  return intKeyOf(fields, key) ? Number(row[key]) : Number(out.lastInsertRowid);
}

/** A record of the copy that is only here moved to another rowid, its entry, and the records that point at it, with it. */
function moveRow(base, table, key, from, to) {
  base.db.prepare(`UPDATE main.${q(table)} SET ${q(key)} = ? WHERE rowid = ?`).run(to, from);
  base.db.prepare('UPDATE rutba_sync SET rid = ?, kj = CASE WHEN state = ? THEN ? ELSE kj END WHERE tbl = ? AND rid = ?').run(to, 'added', kjOf(to), table, from);
  for (const r of base.relationships().filter((x) => x.to.table === table && x.to.field === key && !base.isLinked(x.from.table))) {
    base.db.prepare(`UPDATE ${q(r.from.table)} SET ${q(r.from.field)} = ? WHERE ${q(r.from.field)} = ?`).run(to, from);
  }
}

/** Make room at `rid` in a copy with a whole-number key: a record added here and not yet sent moves along; any other goes. */
function freeRid(base, table, key, rid) {
  const there = base.db.prepare(`SELECT rowid AS r FROM main.${q(table)} WHERE rowid = ?`).get(rid);
  if (!there) return;
  const ent = base.db.prepare('SELECT state FROM rutba_sync WHERE tbl = ? AND rid = ?').get(table, rid);
  if (ent?.state === 'added') {
    const top = base.db.prepare(`SELECT COALESCE(MAX(rowid), 0) AS m FROM main.${q(table)}`).get().m;
    moveRow(base, table, key, rid, Math.max(top, rid) + 1);
  } else {
    base.db.prepare(`DELETE FROM main.${q(table)} WHERE rowid = ?`).run(rid);
    base.db.prepare('DELETE FROM rutba_sync WHERE tbl = ? AND rid = ?').run(table, rid);
  }
}

/**
 * Before the changes made here go to the engine: those the engine changed
 * too since the last catch-up are clashes, and the engine wins (the owner's
 * choice): its version is put here and the one from here kept in
 * rutba_clashes. Answers the rest, to write by key: { pushes: [{ rid, n,
 * action, key, values }], clashes }.
 */
export function prepareSync(base, table, engineRows, { fields, key }, keys = null) {
  ensureSync(base);
  const byKey = new Map(engineRows.map((r) => [kjOf(r[key]), r]));
  // Read by an updated-at field, `engineRows` are those changed there; `keys` are all there are.
  const known = keys ? new Set(keys.map(kjOf)) : null;
  const entries = base.db.prepare('SELECT rid, kj, hash, state, n FROM rutba_sync WHERE tbl = ? AND state IS NOT NULL ORDER BY rid').all(table);
  const local = base.db.prepare(`SELECT rowid AS _rid, * FROM main.${q(table)} WHERE rowid = ?`);
  const pushes = [];
  let clashes = 0;
  inTransaction(base, () => {
    for (const e of entries) {
      const row = local.get(e.rid);
      if (e.state === 'added') {
        if (!row) { base.db.prepare('DELETE FROM rutba_sync WHERE tbl = ? AND rid = ?').run(table, e.rid); continue; }
        const values = valuesOf(fields, row);
        // An AutoNumber is the engine's to give.
        if (fields.find((f) => f.name === key)?.type === 'autonumber') delete values[key];
        pushes.push({ rid: e.rid, n: e.n, action: 'added', key: null, values });
        continue;
      }
      const there = byKey.get(e.kj);
      const changedThere = there && e.hash && hashRecord(fields, there) !== e.hash;
      const gone = !there && !(known && known.has(e.kj));
      if (e.state === 'changed') {
        if (gone) {
          // Deleted in the engine: it goes here too, and what was changed here is kept to see.
          clashes++;
          keepClash(base, table, e.kj, 'gone', row ? valuesOf(fields, row) : null, null);
          base.db.prepare(`DELETE FROM main.${q(table)} WHERE rowid = ?`).run(e.rid);
          base.db.prepare('DELETE FROM rutba_sync WHERE tbl = ? AND rid = ?').run(table, e.rid);
        } else if (changedThere) {
          clashes++;
          keepClash(base, table, e.kj, 'changed', row ? valuesOf(fields, row) : null, valuesOf(fields, there));
          base.db.prepare(`UPDATE main.${q(table)} SET ${fields.map((f) => `${q(f.name)} = ?`).join(', ')} WHERE rowid = ?`).run(...fields.map((f) => there[f.name] ?? null), e.rid);
          base.db.prepare('UPDATE rutba_sync SET hash = ?, state = NULL, since = NULL WHERE tbl = ? AND rid = ?').run(hashRecord(fields, there), table, e.rid);
        } else if (row) {
          const values = valuesOf(fields, row);
          delete values[key];
          pushes.push({ rid: e.rid, n: e.n, action: 'changed', key: JSON.parse(e.kj), values });
        }
      } else if (e.state === 'deleted') {
        if (gone) {
          base.db.prepare('DELETE FROM rutba_sync WHERE tbl = ? AND rid = ?').run(table, e.rid);
        } else if (changedThere) {
          // Deleted here, changed in the engine: the engine's version comes back.
          clashes++;
          keepClash(base, table, e.kj, 'deleted', null, valuesOf(fields, there));
          base.db.prepare('DELETE FROM rutba_sync WHERE tbl = ? AND rid = ?').run(table, e.rid);
          if (intKeyOf(fields, key)) freeRid(base, table, key, Number(there[key]));
          const rid = putRow(base, table, fields, key, there);
          base.db.prepare('INSERT INTO rutba_sync (tbl, rid, kj, hash, state, n) VALUES (?, ?, ?, ?, NULL, 0) ON CONFLICT(tbl, rid) DO UPDATE SET kj = excluded.kj, hash = excluded.hash, state = NULL, since = NULL')
            .run(table, rid, e.kj, hashRecord(fields, there));
        } else {
          pushes.push({ rid: e.rid, n: e.n, action: 'deleted', key: JSON.parse(e.kj), values: null });
        }
      }
    }
  });
  return { pushes, clashes };
}

/**
 * A change from here the engine took. An added record takes the key the
 * engine gave it. Answers 'gone' when the record was deleted here while it
 * was on its way, for the caller to delete it in the engine too.
 */
export function pushed(base, table, push, { fields, key }, newKey = null) {
  return inTransaction(base, () => {
    const e = base.db.prepare('SELECT rid, state, n FROM rutba_sync WHERE tbl = ? AND rid = ?').get(table, push.rid);
    if (push.action === 'deleted') {
      if (e && e.state === 'deleted') base.db.prepare('DELETE FROM rutba_sync WHERE tbl = ? AND rid = ?').run(table, push.rid);
      return 'done';
    }
    if (push.action === 'added') {
      let k = newKey ?? push.values[key] ?? null;
      if (k != null && intKeyOf(fields, key)) k = Number(k);
      if (!e) return k == null ? 'done' : 'gone';
      let rid = push.rid;
      if (k != null && intKeyOf(fields, key) && k !== push.rid) {
        freeRid(base, table, key, k);
        moveRow(base, table, key, push.rid, k);
        rid = k;
      }
      const row = base.db.prepare(`SELECT * FROM main.${q(table)} WHERE rowid = ?`).get(rid);
      // Changed again while it was on its way: the change follows by key.
      const state = e.n !== push.n ? 'changed' : null;
      base.db.prepare('UPDATE rutba_sync SET kj = ?, hash = ?, state = ?, since = CASE WHEN ? IS NULL THEN NULL ELSE since END WHERE tbl = ? AND rid = ?')
        .run(kjOf(k), hashRecord(fields, row), state, state, table, rid);
      return 'done';
    }
    if (e && e.n === push.n) {
      const row = base.db.prepare(`SELECT * FROM main.${q(table)} WHERE rowid = ?`).get(push.rid);
      base.db.prepare('UPDATE rutba_sync SET hash = ?, state = NULL, since = NULL WHERE tbl = ? AND rid = ?').run(hashRecord(fields, row), table, push.rid);
    }
    return 'done';
  });
}

/** A change from here the engine refused (a value it would not take): the engine wins, and the version from here is kept to see. */
export function refused(base, table, push, { fields }, why) {
  inTransaction(base, () => {
    const row = base.db.prepare(`SELECT * FROM main.${q(table)} WHERE rowid = ?`).get(push.rid);
    const e = base.db.prepare('SELECT kj FROM rutba_sync WHERE tbl = ? AND rid = ?').get(table, push.rid);
    keepClash(base, table, e?.kj ?? null, 'refused', row ? valuesOf(fields, row) : push.values, null, why);
    if (push.action === 'added') {
      base.db.prepare(`DELETE FROM main.${q(table)} WHERE rowid = ?`).run(push.rid);
      base.db.prepare('DELETE FROM rutba_sync WHERE tbl = ? AND rid = ?').run(table, push.rid);
    } else {
      // Read again from the engine at the reconcile that follows.
      base.db.prepare('UPDATE rutba_sync SET hash = NULL, state = NULL, since = NULL WHERE tbl = ? AND rid = ?').run(table, push.rid);
    }
  });
}

/**
 * The engine's records brought into a copy, compared by key and hash: those
 * new there added, those changed there changed, and those gone from there
 * removed, leaving alone any record changed here and not yet sent. `rows`
 * are the engine's records (all of them, or those changed since the
 * updated-at field last read when `keys` lists every key there). Answers
 * { added, changed, removed }.
 */
export function reconcile(base, table, { rows, keys = null, complete = true }, { fields, key }) {
  ensureSync(base);
  const out = { added: 0, changed: 0, removed: 0 };
  if (!key) return replaceRows(base, table, rows, fields) ? { ...out, changed: rows.length } : out;
  const intKey = intKeyOf(fields, key);
  return inTransaction(base, () => {
    const entries = base.db.prepare("SELECT rid, kj, hash, state FROM rutba_sync WHERE tbl = ? AND (state IS NULL OR state <> 'added')").all(table);
    const byKj = new Map(entries.map((e) => [e.kj, e]));
    // Records here with no entry yet (a copy made from what was on screen): matched by their key.
    const loose = base.db.prepare(`SELECT rowid AS rid, ${q(key)} AS k FROM main.${q(table)} WHERE rowid NOT IN (SELECT rid FROM rutba_sync WHERE tbl = ?)`).all(table);
    for (const r of loose) {
      const kj = kjOf(r.k);
      if (!byKj.has(kj)) {
        base.db.prepare('INSERT INTO rutba_sync (tbl, rid, kj, hash, state, n) VALUES (?, ?, ?, NULL, NULL, 0)').run(table, r.rid, kj);
        byKj.set(kj, { rid: r.rid, kj, hash: null, state: null });
      }
    }
    const setRow = base.db.prepare(`UPDATE main.${q(table)} SET ${fields.map((f) => `${q(f.name)} = ?`).join(', ')} WHERE rowid = ?`);
    const there = new Set(keys ? keys.map(kjOf) : rows.map((r) => kjOf(r[key])));
    for (const row of rows) {
      const kj = kjOf(row[key]);
      const h = hashRecord(fields, row);
      const e = byKj.get(kj);
      if (!e) {
        if (intKey) freeRid(base, table, key, Number(row[key]));
        const rid = putRow(base, table, fields, key, row);
        base.db.prepare('INSERT INTO rutba_sync (tbl, rid, kj, hash, state, n) VALUES (?, ?, ?, ?, NULL, 0) ON CONFLICT(tbl, rid) DO UPDATE SET kj = excluded.kj, hash = excluded.hash, state = NULL, since = NULL')
          .run(table, rid, kj, h);
        out.added++;
      } else if (e.state == null && e.hash !== h) {
        setRow.run(...fields.map((f) => row[f.name] ?? null), e.rid);
        base.db.prepare('UPDATE rutba_sync SET hash = ? WHERE tbl = ? AND rid = ?').run(h, table, e.rid);
        if (e.hash != null) out.changed++;
      }
    }
    // Gone from the engine: gone here, unless what was read was cut short.
    if (complete) {
      for (const e of entries.concat(loose.map((r) => ({ rid: r.rid, kj: kjOf(r.k), state: null })))) {
        if (e.state != null || there.has(e.kj)) continue;
        base.db.prepare(`DELETE FROM main.${q(table)} WHERE rowid = ?`).run(e.rid);
        base.db.prepare('DELETE FROM rutba_sync WHERE tbl = ? AND rid = ?').run(table, e.rid);
        out.removed++;
      }
    }
    return out;
  });
}

/** A read-only copy's records replaced by those read (a source's, or a table with no key). Answers whether anything changed. */
export function replaceRows(base, table, rows, fields) {
  const before = hashRows(fields, base.db.prepare(`SELECT * FROM main.${q(table)} ORDER BY rowid`).all());
  if (before === hashRows(fields, rows)) return false;
  inTransaction(base, () => {
    base.db.exec(`DELETE FROM main.${q(table)}`);
    const put = base.db.prepare(`INSERT INTO main.${q(table)} (${fields.map((f) => q(f.name)).join(', ')}) VALUES (${fields.map(() => '?').join(', ')})`);
    for (const r of rows) put.run(...fields.map((f) => (r[f.name] === undefined ? null : r[f.name])));
  });
  return true;
}

/** What changed between two designs of a linked table: { added, removed, changed } field names. */
export function designDiff(before, after) {
  const was = new Map(before.map((f) => [f.name, f]));
  const now = new Map(after.map((f) => [f.name, f]));
  return {
    added: after.filter((f) => !was.has(f.name)).map((f) => f.name),
    removed: before.filter((f) => !now.has(f.name)).map((f) => f.name),
    changed: after.filter((f) => was.has(f.name) && was.get(f.name).type !== f.type).map((f) => f.name),
  };
}

/** A copy's table made again with a new design, the records of the fields it keeps kept with it. */
export function reshapeCopy(base, table, before, after, key) {
  const tmp = `rutba_reshape_${Date.now()}`;
  base.db.exec(copyTableSql(tmp, after, key));
  const keep = after.filter((f) => before.some((b) => b.name === f.name));
  if (keep.length) base.db.exec(`INSERT INTO ${q(tmp)} (rowid, ${keep.map((f) => q(f.name)).join(', ')}) SELECT rowid, ${keep.map((f) => q(f.name)).join(', ')} FROM main.${q(table)}`);
  base.db.exec(`DROP TABLE main.${q(table)}`);
  base.db.exec(`ALTER TABLE ${q(tmp)} RENAME TO ${q(table)}`);
  // The engine's versions are hashed again with the new design at the next read.
  if (hasSync(base)) base.db.prepare('UPDATE rutba_sync SET hash = NULL WHERE tbl = ? AND state IS NULL').run(table);
}

/** When a copy last caught up with its engine or source. */
export function markSynced(base, table, at = now()) {
  ensureSync(base);
  base.db.prepare('UPDATE rutba_linked SET synced = ? WHERE tbl = ?').run(at, table);
}

/** The largest value of a copy's updated-at field, to read only what changed after it. */
export function stampMax(base, table, field) {
  try { return base.db.prepare(`SELECT MAX(${q(field)}) AS m FROM main.${q(table)}`).get().m ?? null; } catch { return null; }
}
