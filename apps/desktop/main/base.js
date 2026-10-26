// Databases: the database app's service. A database is one SQLite file
// (@rutba/database), opened here and held while a window has it open; every
// design change goes through the file's numbered steps, and the window asks
// for its rows a page at a time, so a table of a million records opens as
// quickly as one of ten.
//
// Imports make a new table each: delimited text (CSV, TSV), workbooks (.xlsx
// and .ods, a table a sheet, read as Worksheets reads them) and the old files
// Worksheets reads (dBASE and FoxPro tables, Access databases, Lotus 1-2-3,
// SYLK, DIF), through office-formats' readers.

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { Base, importRows, importSheets, buildTemplate, buildSample, originOf, planUpgrade, applyUpgrade, toStored } from '@rutba/database';
import { ENGINES, connectEngine as defaultConnect } from './engines/index.js';
import { SOURCE_KINDS } from '@rutba/database';
import { createSync, unreachable, sourceSecret, cleanSourceSettings } from './base-sync.js';
import { createTrust, addressOf, placeOf, keptSecrets } from './base-trust.js';
import { TEMPLATE_IDS, SAMPLE_IDS, templateNeeds, sampleNeeds, pageOf } from '@rutba/templates/data-files/active';
import { needsFor } from '@rutba/templates/versions';
import { readManifest } from '@rutba/templates/manifest';
import { parseDelimited } from '@rutba/ooxml/mailmerge';
import { legacySheetKind, readLegacySheets } from '@rutba/office-formats/legacy-sheets';
import MDBReader from 'mdb-reader';
import { readWorkbook, WORKBOOK_EXTENSIONS } from './base-workbook.js';

/** The SQLite file's own first sixteen bytes. */
const SQLITE_MAGIC = 'SQLite format 3\u0000';

/** Whether bytes are a SQLite database. */
export const isSqlite = (bytes) => Buffer.from(bytes.subarray(0, 16)).toString('latin1') === SQLITE_MAGIC;

/**
 * `secrets` keeps a linked engine's password (D-2026-10-22-13), never the
 * file, each kept with the host or address it is for (base-trust.js,
 * `keptSecrets`); `connect` opens an engine (engines/index.js, or a stand-in in tests);
 * `trust` is this computer's list of what each database may read
 * (base-trust.js, A2.18), kept in memory for the service's life when it is
 * not given.
 */
export function createBaseService({ secrets = null, connect = defaultConnect, trust = createTrust() } = {}) {
  const open = new Map(); // id -> { base, file, engines: Map(link -> engine), state: Map(table -> { error, at, count }) }
  const kept = keptSecrets(secrets);
  let next = 1;

  const get = (id) => {
    const o = open.get(id);
    if (!o) throw new Error('That database is not open.');
    return o;
  };
  const summary = (id) => {
    const { base, file } = get(id);
    return {
      id,
      file,
      name: path.basename(file).replace(/\.[^.]+$/, ''),
      tables: base.tables().map((t) => {
        if (!t.linked) return { ...t, count: base.count(t.name) };
        // A copy kept in the file (D-2026-10-22-13, steps 4 and 5): when it caught up, what waits, what clashed.
        const copy = sync.describe(open.get(id), t);
        if (copy) return { ...t, ...copy };
        // A linked table: its records as last read from the engine, or why they could not be.
        const st = open.get(id).state.get(t.name) || {};
        const l = base.linked().find((x) => x.table === t.name);
        const link = l && base.links().find((x) => x.id === l.link);
        return { ...t, mode: 'live', count: base.linkedLoaded(t.name) ? base.count(t.name) : 0, linkError: st.error || null, readAt: st.at || null, key: l?.key || null, blocked: Boolean(link) && !allowed(open.get(id), link) };
      }),
      links: base.access?.locked || base.rights() ? [] : base.links().filter((l) => !SOURCE_KINDS.has(l.engine)).map((l) => ({ id: l.id, name: l.name, engine: l.engine, settings: l.settings, password: kept.has(secretKey(l.id), l) })),
      // The other sources tables are read from: a file, another database, a web address answering JSON.
      sources: base.access?.locked || base.rights() ? [] : base.links().filter((l) => SOURCE_KINDS.has(l.engine)).map((l) => ({ id: l.id, name: l.name, kind: l.engine, settings: l.settings })),
      // The sources and engines this computer has not allowed it to read (A2.18), for the question that names each.
      ask: base.access?.locked || base.rights() ? [] : askOf(open.get(id)),
      // Only those between tables the person's role shows.
      relationships: (() => { const seen = new Set(base.tables().map((t) => t.name)); return base.relationships().filter((r) => seen.has(r.from.table) && seen.has(r.to.table)); })(),
      queries: base.queries().map((x) => x.name),
      forms: base.forms().map((x) => x.name),
      reports: base.reports().map((x) => x.name),
      documents: base.documents().map((x) => ({ name: x.name, table: x.layout.table })),
      steps: base.steps().length,
      // Pinned items, first in the side panel, in the order pinned; those since renamed or removed left out.
      pinned: pinnedOf(base),
      folded: listMeta(base, 'folded'),
      // The template or sample it was made from, and its page online (D-2026-10-22-03).
      origin: (() => { const o = originOf(base); return o ? { kind: o.kind, id: o.id, set: o.set, page: pageOf(o.id) } : null; })(),
      // Who is using it (D-2026-10-22-11): protected or not, who is signed in, and what their role narrows.
      access: accessOf(base),
    };
  };
  /** Whether this computer allows the database to read or connect to `link` (A2.18). */
  const allowed = (o, link) => trust.allowed(o.file, link);
  /** Refused before anything is read or connected to. */
  const notAllowed = () => Object.assign(new Error('This database\'s sources are not read on this computer until you allow them.'), { blocked: true });
  /** The links with tables that this computer has not allowed: each by its kind and where it is, and the tables read from it. */
  const askOf = (o) => {
    const linked = o.base.linked();
    return o.base.links().filter((l) => linked.some((x) => x.link === l.id) && !allowed(o, l)).map((l) => ({
      id: l.id,
      name: l.name,
      kind: l.engine,
      engine: SOURCE_KINDS.has(l.engine) ? null : ENGINES.find((e) => e.id === l.engine)?.name || l.engine,
      address: addressOf(l),
      tables: linked.filter((x) => x.link === l.id).map((x) => x.table),
      ...placeOf(l),
    }));
  };
  const accessOf = (base) => {
    if (!base.isProtected()) return { protected: false };
    const person = base.person();
    return { protected: true, locked: !person, names: person ? [] : base.signInNames(), person, rights: base.rights() };
  };
  /** Only for someone who sees everything: a copy of the file, or a newer design, would reach past a role. */
  const whole = (base) => {
    if (base.access?.locked) throw new Error('Sign in to open this database.');
    if (base.rights()) throw new Error('Only someone whose role may change the design can do that.');
  };
  const listMeta = (base, key) => { try { const v = JSON.parse(base.meta(key) || '[]'); return Array.isArray(v) ? v : []; } catch { return []; } };
  const pinnedOf = (base) => {
    const have = {
      table: new Set(base.tables().map((x) => x.name)),
      query: new Set(base.queries().map((x) => x.name)),
      form: new Set(base.forms().map((x) => x.name)),
      report: new Set(base.reports().map((x) => x.name)),
      document: new Set(base.documents().map((x) => x.name)),
    };
    return listMeta(base, 'pinned').filter((p) => have[p?.kind]?.has(p.name));
  };
  const letGo = (id) => {
    const o = open.get(id);
    if (!o) return false;
    for (const e of o.engines.values()) e.then((x) => x.close()).catch(() => {});
    try { o.base.close(); } catch { /* already closed */ }
    open.delete(id);
    return true;
  };

  /* ── tables linked to a database engine (D-2026-10-22-13) ───────── */

  const secretKey = (linkId) => `base-link:${linkId}`;
  /** The engine behind `link`, connected once per window and kept while it lives. */
  const engineFor = (o, link) => {
    if (!allowed(o, link)) return Promise.reject(notAllowed());
    if (!o.engines.has(link.id)) {
      // A password kept for another host, port or database (the file's link changed elsewhere) is not sent: it is asked for again.
      const { value, elsewhere } = kept.read(secretKey(link.id), link);
      if (elsewhere) return Promise.reject(new Error(`The password this computer keeps for ${link.name} was saved for another engine. Type it again for ${link.settings?.host || ''} if you mean to send it there.`));
      const p = connect(link.engine, link.settings, value || '');
      o.engines.set(link.id, p);
      // A connection that failed is tried again next time, not kept.
      p.catch(() => { if (o.engines.get(link.id) === p) o.engines.delete(link.id); });
    }
    return o.engines.get(link.id);
  };
  /** A connection that broke let go, so the next read connects again (the engine stopped and started). */
  const dropEngine = (o, link) => {
    const p = o.engines.get(link.id);
    o.engines.delete(link.id);
    p?.then((x) => x.close()).catch(() => {});
  };
  const sync = createSync({ engineFor, dropEngine, secrets: kept, allowed });
  /** A linked table's records read from the engine into this window's copy, its design kept up to date. */
  const readLinked = async (o, table) => {
    const l = o.base.linked().find((x) => x.table === table);
    if (!l) return;
    const link = o.base.links().find((x) => x.id === l.link);
    try {
      if (!link) throw new Error('Its link is gone.');
      const engine = await engineFor(o, link);
      const design = await engine.describe(l.remote);
      o.base.relinkDesign(table, design);
      const rows = await engine.read(l.remote, { fields: design.fields, key: design.key });
      // Read again with nothing changed in the engine: the copy here, and when it was read, stay as they are.
      const hash = crypto.createHash('sha256').update(JSON.stringify([design, rows])).digest('hex');
      const was = o.state.get(table);
      if (was?.hash === hash && !was.error && o.base.linkedLoaded(table)) return;
      o.base.loadLinked(table, rows);
      o.state.set(table, { error: null, at: new Date().toISOString(), count: rows.length, hash });
    } catch (err) {
      if (link && unreachable(err)) dropEngine(o, link);
      o.state.set(table, { error: err.message || String(err), at: null });
    }
  };
  /**
   * A linked table caught up, as it is kept: read live, synced as a copy, or
   * read again from its source. One this computer has not allowed is left
   * as it is, nothing read or connected to.
   */
  const catchUp = (o, l) => {
    const link = o.base.links().find((x) => x.id === l.link);
    if (link && !allowed(o, link)) return Promise.resolve();
    return l.mode !== 'copy' ? readLinked(o, l.table) : sync.isSourceTable(o, l) ? sync.refreshSource(o, l.table) : sync.syncCopy(o, l.table);
  };
  /**
   * Every linked table read, as a database opens or is refreshed; one that
   * cannot be read says why. A copy is in the file already: `copiesFor`
   * waits that long for copies to catch up, the rest going on behind.
   */
  const readAllLinked = async (o, { copiesFor = Infinity } = {}) => {
    if (o.base.access?.locked || o.base.rights()) return;
    const all = o.base.linked();
    const copies = Promise.all(all.filter((l) => l.mode === 'copy').map((l) => catchUp(o, l)));
    const live = Promise.all(all.filter((l) => l.mode !== 'copy').map((l) => catchUp(o, l)));
    await Promise.all([live, Number.isFinite(copiesFor) ? Promise.race([copies, new Promise((r) => setTimeout(r, copiesFor))]) : copies]);
  };
  /** A change to a linked table: made in the engine, then the table read again. */
  const writeLinked = async (o, table, run) => {
    const l = o.base.linked().find((x) => x.table === table);
    const link = o.base.links().find((x) => x.id === l.link);
    if (o.base.rights()) throw new Error('Only someone whose role may change the design can change a table linked to an engine, for now.');
    const engine = await engineFor(o, link);
    const fields = o.base.fields(table);
    const stored = (values) => Object.fromEntries(Object.entries(values || {}).flatMap(([k, v]) => {
      const f = fields.find((x) => x.name === k);
      return f ? [[k, toStored(f, v)]] : [];
    }));
    const out = await run({ engine, remote: l.remote, design: { key: l.key }, stored });
    await readLinked(o, table);
    const st = o.state.get(table);
    if (st?.error) throw new Error(st.error);
    return out;
  };
  /** Whether a table is read live from its engine, its changes written there at once. */
  const isLive = (base, table) => base.isLinked(table) && base.linked().find((x) => x.table === table)?.mode !== 'copy';
  // Each window holds its own connection to the file, let go when the window
  // closes: two windows on one database are two readers and writers of it,
  // as SQLite allows, and one closing leaves the other's open.
  const hold = (file, win) => {
    const base = Base.open(file);
    const id = String(next++);
    open.set(id, { base, file, engines: new Map(), state: new Map() });
    win?.once?.('closed', () => letGo(id));
    return id;
  };

  /** A new database at `file`, a file there replaced, filled by `make(base)`; nothing is left behind if that fails. */
  const made = (file, win, make) => {
    if (!file || !path.isAbsolute(String(file))) throw new Error('A new database needs a place to be saved.');
    if (fs.existsSync(file)) {
      fs.rmSync(file, { force: true });
      for (const s of ['-wal', '-shm']) fs.rmSync(file + s, { force: true });
    }
    fs.mkdirSync(path.dirname(file), { recursive: true });
    const id = hold(file, win);
    try {
      make(get(id).base);
    } catch (err) {
      letGo(id);
      for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
      throw err;
    }
    // Made here: what it was made with is allowed, and nothing a file at that place was allowed before.
    trust.reset(file, get(id).base.links());
    return summary(id);
  };

  return {
    /** A new database at `file` (a .rbase, made empty). */
    create: ({ file }, win) => {
      if (!file || !path.isAbsolute(String(file))) throw new Error('A new database needs a place to be saved.');
      if (fs.existsSync(file)) {
        // Saving over a file that is there replaces it, as Save As does, once it has been asked.
        fs.rmSync(file, { force: true });
        for (const s of ['-wal', '-shm']) fs.rmSync(file + s, { force: true });
      }
      fs.mkdirSync(path.dirname(file), { recursive: true });
      const id = hold(file, win);
      trust.reset(file);
      return summary(id);
    },

    /**
     * The ready-made databases this Office has, and which need a newer one:
     * { set, office, templates: [{ id, needs }], samples: [{ id, needs }] },
     * `needs` the version wanted, or null when this Office can make it.
     */
    readyMade: () => {
      const office = process.env.RUTBA_APP_VERSION;
      const manifest = (() => { try { return readManifest(); } catch { return {}; } })();
      const part = manifest.parts?.dataFiles || {};
      const whole = needsFor(office, part) || needsFor(office, manifest);
      return {
        set: manifest.set || null,
        office,
        templates: TEMPLATE_IDS.map((id) => ({ id, needs: whole || needsFor(office, { minOffice: templateNeeds(id) }), page: pageOf(id) })),
        samples: SAMPLE_IDS.map((id) => ({ id, needs: whole || needsFor(office, { minOffice: sampleNeeds(id) }), page: pageOf(id) })),
      };
    },

    /** A new database at `file` made from a template: its tables, relationships, queries, forms and reports. */
    createFromTemplate: ({ file, template }, win) => made(file, win, (base) => buildTemplate(base, template)),

    /** A new database at `file` made from a sample: its design, and its invented records unless `withData` is false. */
    createFromSample: ({ file, sample, withData = true }, win) => made(file, win, (base) => buildSample(base, sample, { withData: withData !== false })),

    /** Open a database file: ours, or any SQLite database. */
    open: async ({ file }, win) => {
      const head = Buffer.alloc(16);
      const fd = fs.openSync(file, 'r');
      try { fs.readSync(fd, head, 0, 16, 0); } finally { fs.closeSync(fd); }
      if (!isSqlite(head)) throw new Error(`${path.basename(file)} is not a database this app opens. Import it into a database instead.`);
      const id = hold(file, win);
      await readAllLinked(get(id), { copiesFor: 2500 });
      return summary(id);
    },

    /**
     * A database opened for the AI tools (mcp/reach-tools.js), held apart
     * from any window's and let go with close: a protected one as the role
     * chosen for the tools on this computer (Base.asTools), never signed in
     * as anyone. Its linked tables are read when the tools ask for one.
     */
    openForTools: ({ file, role = null }) => {
      const head = Buffer.alloc(16);
      const fd = fs.openSync(file, 'r');
      try { fs.readSync(fd, head, 0, 16, 0); } finally { fs.closeSync(fd); }
      if (!isSqlite(head)) throw new Error(`${path.basename(file)} is not a database.`);
      const id = hold(file, null);
      const { base } = get(id);
      if (base.isProtected()) {
        try {
          if (!role) throw new Error(`${path.basename(file)} is protected, and no role has been chosen for the AI tools. An Admin chooses one in Data Files, in the database's AI tools.`);
          base.asTools(role);
        } catch (err) {
          letGo(id);
          throw err;
        }
      }
      return summary(id);
    },

    close: ({ id }) => ({ closed: letGo(id) }),

    summary: ({ id }) => summary(id),
    fields: ({ id, table }) => {
      const { base } = get(id);
      const fields = base.fields(table);
      // A table read from a source, or one with no key in its engine, is read here, not changed.
      const l = base.isLinked(table) ? base.linked().find((x) => x.table === table) : null;
      const fixed = l && (!l.key || SOURCE_KINDS.has(base.links().find((x) => x.id === l.link)?.engine));
      return fixed ? fields.map((f) => ({ ...f, readOnly: true })) : fields;
    },
    rows: ({ id, table, offset = 0, limit = 200, order = null, where = {}, search = '', filters = [] }) => {
      const { base } = get(id);
      return { rows: base.rows(table, { offset, limit, order, where, search, filters }), count: base.count(table, where, { search, filters }) };
    },
    // A query's results as a table's records are read: a page, sorted, searched and filtered (D-2026-10-22-07).
    queryRows: ({ id, name = null, design = null, offset = 0, limit = 200, order = null, search = '', filters = [] }) => get(id).base.queryRows({ name, design, offset, limit, order, search, filters }),
    // A linked table's changes go to its engine (D-2026-10-22-13); the rest to the file, a copy's
    // marked to go to its engine at the next catch-up.
    insert: async ({ id, table, values }) => {
      const o = get(id);
      if (!isLive(o.base, table)) return { rowid: o.base.insert(table, values || {}) };
      const key = await writeLinked(o, table, ({ engine, remote, design, stored }) => engine.insert(remote, design, stored(values)));
      const l = o.base.linked().find((x) => x.table === table);
      const row = key != null && l.key ? o.base.rows(table, { where: { [l.key]: key }, limit: 1 })[0] : null;
      return { rowid: row?._rowid ?? null };
    },
    update: async ({ id, table, rowid, values }) => {
      const o = get(id);
      if (!isLive(o.base, table)) return { changed: o.base.update(table, rowid, values || {}) };
      const key = o.base.linkedKey(table, rowid);
      return { changed: await writeLinked(o, table, ({ engine, remote, design, stored }) => engine.update(remote, design, key, stored(values))) };
    },
    remove: async ({ id, table, rowids }) => {
      const o = get(id);
      if (!isLive(o.base, table)) return { removed: o.base.remove(table, rowids || []) };
      const keys = (rowids || []).map((r) => o.base.linkedKey(table, r));
      return { removed: await writeLinked(o, table, ({ engine, remote, design }) => engine.remove(remote, design, keys)) };
    },

    /** The engines a table can be linked to: [{ id, name, port, ready }]. */
    engines: () => ENGINES,
    /** Try an engine's address and account: its tables and views, or why not. Nothing is kept. */
    engineTables: async ({ id = null, link = null, engine, settings, password }) => {
      // An existing link's tables, through the password this computer keeps for it.
      if (id && link) {
        const o = get(id);
        whole(o.base);
        const l = o.base.links().find((x) => x.id === link);
        if (!l) throw new Error('There is no such link.');
        return { tables: await (await engineFor(o, l)).tables() };
      }
      const e = await connect(engine, settings, password || '');
      try { return { tables: await e.tables() }; } finally { await e.close(); }
    },
    /** A link to an engine kept in the database, its password in this computer's protected store. Answers { link, summary }. */
    saveLink: async ({ id, link = null, name, engine, settings, password = null }) => {
      const o = get(id);
      whole(o.base);
      const linkId = o.base.saveLink({ id: link, name, engine, settings });
      // Linked here, by the person: allowed on this computer as it is.
      trust.allow(o.file, o.base.links().filter((l) => l.id === linkId));
      if (password != null) {
        if (!secrets?.available?.()) throw new Error('This computer offers no protected store for the password, so the link cannot keep it.');
        kept.write(secretKey(linkId), o.base.links().find((l) => l.id === linkId), String(password));
      }
      const was = o.engines.get(linkId);
      o.engines.delete(linkId);
      was?.then((x) => x.close()).catch(() => {});
      await readAllLinked(o);
      return { link: linkId, summary: summary(id) };
    },
    /** A link taken away, its password with it. */
    removeLink: async ({ id, link }) => {
      const o = get(id);
      whole(o.base);
      o.base.removeLink(link);
      secrets?.delete?.(secretKey(link));
      return summary(id);
    },
    /** Tables of the engine linked: remotes [{ schema, name, as? }]. Each is read at once. */
    linkTables: async ({ id, link, remotes }) => {
      const o = get(id);
      whole(o.base);
      const l = o.base.links().find((x) => x.id === link);
      if (!l) throw new Error('There is no such link.');
      const engine = await engineFor(o, l);
      const made = [];
      for (const r of remotes || []) {
        const design = await engine.describe({ schema: r.schema, name: r.name });
        // Each read live (as 2.17) or kept as a copy in this file (D-2026-10-22-13, step 4).
        const table = o.base.linkTable({ table: r.as || r.name, link, remote: { schema: r.schema, name: r.name }, key: design.key, fields: design.fields, mode: r.mode === 'copy' ? 'copy' : 'live' });
        made.push(table);
        await catchUp(o, o.base.linked().find((x) => x.table === table));
      }
      return { tables: made, summary: summary(id) };
    },
    /** A table unlinked: the engine keeps its records. A source's link goes with its table, its secrets with it. */
    unlinkTable: ({ id, table }) => {
      const o = get(id);
      whole(o.base);
      const l = o.base.linked().find((x) => x.table === table);
      const link = l ? o.base.links().find((x) => x.id === l.link) : null;
      o.base.unlinkTable(table);
      o.state.delete(table);
      if (link && SOURCE_KINDS.has(link.engine) && !o.base.linked().some((x) => x.link === link.id)) {
        o.base.removeLink(link.id);
        secrets?.delete?.(sourceSecret(link.id));
        for (const h of link.settings?.headers || []) if (h.secret) secrets?.delete?.(sourceSecret(link.id, h.name));
      }
      return summary(id);
    },
    /**
     * The linked tables caught up (one, or all): read again live, synced as
     * copies, and sources read again; a source on screen only when it is due
     * on its schedule, unless `now` (Refresh, Sync now).
     */
    refreshLinked: async ({ id, table = null, now = false }) => {
      const o = get(id);
      if (!table) { await readAllLinked(o); return summary(id); }
      const l = o.base.linked().find((x) => x.table === table);
      if (l && (now || !(l.mode === 'copy' && sync.isSourceTable(o, l)) || sync.due(o, l))) await catchUp(o, l);
      return summary(id);
    },
    /** Sources due on their schedules read again: the summary when any was, else null. */
    refreshDue: async ({ id }) => ((await sync.refreshDue(get(id))) ? summary(id) : null),
    /**
     * The sources and engines the person was asked about allowed for this
     * database on this computer (A2.18), by the addresses they were shown,
     * then read. Answers the summary.
     */
    allowSources: async ({ id, addresses = [] }) => {
      const o = get(id);
      whole(o.base);
      const shown = new Set(addresses || []);
      trust.allow(o.file, o.base.links().filter((l) => shown.has(addressOf(l))));
      await readAllLinked(o);
      return summary(id);
    },
    /** A linked table read live or kept as a copy in this file, its records kept, then caught up. */
    setLinkMode: async ({ id, table, mode }) => {
      const o = get(id);
      whole(o.base);
      o.base.setLinkMode(table, mode);
      o.state.delete(table);
      await catchUp(o, o.base.linked().find((x) => x.table === table));
      return summary(id);
    },
    /** A copy's updated-at field, to read only what changed in a large table; null to read all of it. */
    setLinkStamp: ({ id, table, field = null }) => { const o = get(id); whole(o.base); o.base.setLinkStamp(table, field); return summary(id); },
    /** The versions from here a copy's engine replaced: [{ id, key, at, action, mine, theirs, why }]. */
    clashes: ({ id, table }) => sync.clashes(get(id), table),
    /** A clash settled: the engine's version kept, or the one from here put back to go to the engine. */
    settleClash: ({ id, table, clash, keep = 'engine' }) => { sync.settleClash(get(id), table, clash, keep); return summary(id); },
    /** What another Data Files database offers to link: its tables and queries, as the person named may see them. */
    sourceItems: ({ where, person = null, password = null }) => sync.rbaseItems(where, { person, password }),
    /**
     * A source tried before it is linked (a file, another database, a web
     * address answering JSON): its fields, first records and count, or why
     * it cannot be read. Nothing is kept.
     */
    trySource: ({ kind, settings, password = null, headers = {}, item = null }) => sync.trySource(kind, settings, { password, headers, item }),
    /**
     * A source linked: kept in the database as a link of its own with one
     * table, its records a copy in the file, read at once. Its password, and
     * the headers marked secret, go in this computer's protected store,
     * never in the file. Answers { table, summary }.
     */
    linkSource: async ({ id, kind, name, settings, password = null, headers = {}, item = null }) => {
      const o = get(id);
      whole(o.base);
      if (!SOURCE_KINDS.has(kind)) throw new Error(`${kind} is not a source Data Files reads.`);
      const clean = cleanSourceSettings(kind, settings);
      const secret = (kind === 'rbase' && password) || clean.headers?.some((h) => h.secret);
      if (secret && !secrets?.available?.()) throw new Error('This computer offers no protected store for the key or password, so the source cannot keep it.');
      const got = await sync.trySource(kind, clean, { password, headers, item });
      if (!got.fields.length) throw new Error('There is nothing there to read.');
      const want = String(name || '').trim();
      if (o.base.tables().some((t) => t.name.toLowerCase() === want.toLowerCase())) throw new Error(`There is a table called ${want} already.`);
      // The source's link is named as its table, made its own beside an engine's of that name.
      const taken = new Set(o.base.links().map((l) => l.name.toLowerCase()));
      let linkName = want;
      for (let n = 2; taken.has(linkName.toLowerCase()); n++) linkName = `${want} ${n}`;
      const linkId = o.base.saveLink({ name: linkName, engine: kind, settings: clean });
      trust.allow(o.file, o.base.links().filter((l) => l.id === linkId));
      try {
        const made = o.base.links().find((l) => l.id === linkId);
        if (kind === 'rbase' && password) kept.write(sourceSecret(linkId), made, String(password));
        for (const h of clean.headers || []) if (h.secret) kept.write(sourceSecret(linkId, h.name), made, String(headers[h.name] ?? ''));
        const remote = kind === 'file' ? { sheet: got.sheet || null } : kind === 'rbase' ? { item: got.item } : {};
        const table = o.base.linkTable({ table: want, link: linkId, remote, key: null, fields: got.fields, mode: 'copy' });
        await sync.refreshSource(o, table);
        return { table, summary: summary(id) };
      } catch (err) {
        try {
          o.base.removeLink(linkId);
          secrets?.delete?.(sourceSecret(linkId));
          for (const h of clean.headers || []) if (h.secret) secrets?.delete?.(sourceSecret(linkId, h.name));
        } catch { /* a table was made: it stays, with its link */ }
        throw err;
      }
    },
    /** How often a source is read again: every 5, 15 or 60 minutes, or 0, by hand. */
    setSourceEvery: ({ id, table, every }) => {
      const o = get(id);
      whole(o.base);
      const l = o.base.linked().find((x) => x.table === table);
      const link = l && o.base.links().find((x) => x.id === l.link);
      if (!link || !SOURCE_KINDS.has(link.engine)) throw new Error(`${table} is not read from a source.`);
      o.base.saveLink({ id: link.id, name: link.name, engine: link.engine, settings: cleanSourceSettings(link.engine, { ...link.settings, every }) });
      return summary(id);
    },
    /** Who you are in a protected database a source reads: kept for the next reads, the password in the protected store. */
    sourceSignIn: async ({ id, table, person, password }) => {
      const o = get(id);
      whole(o.base);
      const l = o.base.linked().find((x) => x.table === table);
      const link = l && o.base.links().find((x) => x.id === l.link);
      if (!link || link.engine !== 'rbase') throw new Error(`${table} is not read from another database.`);
      if (!secrets?.available?.()) throw new Error('This computer offers no protected store for the password, so the source cannot keep it.');
      o.base.saveLink({ id: link.id, name: link.name, engine: link.engine, settings: cleanSourceSettings('rbase', { ...link.settings, person }) });
      kept.write(sourceSecret(link.id), link, String(password ?? ''));
      await sync.refreshSource(o, table);
      return summary(id);
    },
    /**
     * A web address's keys or tokens typed again (those its headers mark
     * secret), kept for that address as it is now, then read. Answers the summary.
     */
    sourceKeys: async ({ id, table, headers = {} }) => {
      const o = get(id);
      whole(o.base);
      const l = o.base.linked().find((x) => x.table === table);
      const link = l && o.base.links().find((x) => x.id === l.link);
      if (!link || link.engine !== 'json') throw new Error(`${table} is not read from a web address.`);
      if (!secrets?.available?.()) throw new Error('This computer offers no protected store for the key or password, so the source cannot keep it.');
      for (const h of link.settings?.headers || []) if (h.secret && headers[h.name] != null) kept.write(sourceSecret(link.id, h.name), link, String(headers[h.name]));
      await sync.refreshSource(o, table);
      return summary(id);
    },

    createTable: ({ id, name, fields = [] }) => { get(id).base.createTable({ name, fields }); return summary(id); },
    addField: ({ id, table, field }) => { get(id).base.addField(table, field); return summary(id); },
    // `clear`: the records whose value does not fit the new type are let go empty, as the person was asked.
    changeField: ({ id, table, name, change, clear = false }) => { get(id).base.changeField(table, name, change, { clear: Boolean(clear) }); return summary(id); },
    removeField: ({ id, table, name }) => { get(id).base.removeField(table, name); return summary(id); },
    renameTable: ({ id, table, to }) => { get(id).base.renameTable(table, to); return summary(id); },
    dropTable: ({ id, table }) => { get(id).base.dropTable(table); return summary(id); },
    addRelationship: ({ id, from, to, cascadeDelete = false }) => { get(id).base.addRelationship({ from, to, cascadeDelete }); return summary(id); },
    removeRelationship: ({ id, relation }) => { get(id).base.removeRelationship(relation); return summary(id); },
    steps: ({ id }) => get(id).base.steps(),
    /**
     * A newer design for a database made from a template or sample, when the
     * set in use has one (D-2026-10-22-04): what it would add, replace, or put
     * beside the person's own version. Records are never touched.
     */
    designUpdate: ({ id }) => {
      const { base } = get(id);
      // Offered to those who may change the design; to anyone else it is not there.
      if (base.access?.locked || base.rights()) return { available: false };
      const { available, tables, fields, added, replaced, beside } = planUpgrade(base);
      return { available, tables, fields, added, replaced, beside };
    },
    applyDesignUpdate: ({ id }) => {
      whole(get(id).base);
      const { done, tables, fields, added, replaced, beside } = applyUpgrade(get(id).base);
      return { done, tables, fields, added, replaced, beside, summary: summary(id) };
    },
    /** Pin an item to the top of the side panel, or unpin it (D-2026-10-22-05). */
    pin: ({ id, kind, name, pinned = true }) => {
      const { base } = get(id);
      if (!['table', 'query', 'form', 'report', 'document'].includes(kind)) throw new Error(`${kind} cannot be pinned.`);
      const list = listMeta(base, 'pinned').filter((p) => !(p.kind === kind && p.name === name));
      if (pinned) list.push({ kind, name: String(name) });
      base.setMeta('pinned', JSON.stringify(list));
      return summary(id);
    },
    /** A side panel section folded away, or open again. */
    fold: ({ id, section, folded = true }) => {
      const { base } = get(id);
      const list = listMeta(base, 'folded').filter((s) => s !== section);
      if (folded) list.push(String(section));
      base.setMeta('folded', JSON.stringify(list));
      return summary(id);
    },

    /* Queries built in the designer, and forms (2.10's Data Files). */
    queries: ({ id }) => get(id).base.queries(),
    saveQuery: ({ id, name, design }) => { get(id).base.saveQuery(name, design); return summary(id); },
    removeQuery: ({ id, name }) => { get(id).base.removeQuery(name); return summary(id); },
    runQuery: ({ id, name = null, design = null, limit = 5000 }) => get(id).base.runQuery({ name, design, limit }),
    sqlOf: ({ id, design }) => get(id).base.sqlOf(design),
    forms: ({ id }) => get(id).base.forms(),
    saveForm: ({ id, layout }) => { get(id).base.saveForm(layout); return summary(id); },
    removeForm: ({ id, name }) => { get(id).base.removeForm(name); return summary(id); },
    makeForm: ({ id, table, name = null }) => ({ layout: get(id).base.makeForm(table, { name }), summary: summary(id) }),
    reports: ({ id }) => get(id).base.reports(),
    saveReport: ({ id, layout }) => { get(id).base.saveReport(layout); return summary(id); },
    removeReport: ({ id, name }) => { get(id).base.removeReport(name); return summary(id); },
    makeReport: ({ id, table = null, query = null, name = null }) => ({ layout: get(id).base.makeReport({ table, query, name }), summary: summary(id) }),
    reportData: ({ id, name = null, layout = null }) => get(id).base.reportData({ name, layout }),
    documents: ({ id }) => get(id).base.documents(),
    saveDocument: ({ id, layout }) => { get(id).base.saveDocument(layout); return summary(id); },
    removeDocument: ({ id, name }) => { get(id).base.removeDocument(name); return summary(id); },
    makeDocument: ({ id, table, name = null }) => ({ layout: get(id).base.makeDocument({ table, name }), summary: summary(id) }),
    documentData: ({ id, name = null, layout = null, keys = null, limit = 1000 }) => get(id).base.documentData({ name, layout, keys, limit }),
    lookup: ({ id, table, field, show }) => get(id).base.lookup({ table, field, show }),
    // A page of a long list's choices, searched, or the words of given records (a form's list of 20,000 sales).
    lookupPage: ({ id, table, field, search = '', values = null, limit = 500 }) => get(id).base.lookupPage({ table, field, search, values, limit }),

    /** The person's own SELECT, read-only: { columns, rows }, at most `limit` rows. */
    query: ({ id, sql, limit = 5000 }) => {
      const out = get(id).base.query(sql);
      return { columns: out.columns, rows: out.rows.slice(0, limit), count: out.rows.length };
    },

    /**
     * A file imported as new tables: CSV or TSV, a workbook (each sheet, or
     * the `sheet` named), or an old file Worksheets reads. Answers { tables:
     * [{ table, count, skipped, key, dates }], notes, summary }. With `peek`, a
     * workbook's sheets are only named: { sheets }, nothing imported.
     */
    importFile: ({ id, file, sheet = null, peek = false }) => {
      const { base } = get(id);
      const bytes = new Uint8Array(fs.readFileSync(file));
      const name = path.basename(file).replace(/\.[^.]+$/, '');
      const ext = path.extname(file).toLowerCase();
      let made;
      let notes = [];
      if (WORKBOOK_EXTENSIONS.includes(ext.slice(1))) {
        let sheets;
        try {
          sheets = readWorkbook(bytes, ext, { only: peek ? null : sheet });
        } catch (err) {
          throw new Error(`${path.basename(file)} could not be read as a workbook: ${err.message}`);
        }
        if (peek) return { sheets: sheets.map((s) => s.name) };
        if (!sheets.length) throw new Error(`${path.basename(file)} has nothing in it to import.`);
        // One sheet is named after the file, as a CSV is; several after their sheets.
        made = importSheets(base, { sheets: sheets.length === 1 && !sheet ? [{ ...sheets[0], name }] : sheets });
      } else if (ext === '.csv' || ext === '.tsv' || ext === '.txt') {
        const text = new TextDecoder('utf-8').decode(bytes);
        made = [importRows(base, name, parseDelimited(text, ext === '.tsv' ? '\t' : null))];
      } else {
        const kind = legacySheetKind(bytes, ext);
        if (!kind) throw new Error(`${path.basename(file)} is not a file this app imports: CSV, TSV, a workbook (.xlsx or .ods), dBASE, Access, Lotus 1-2-3, SYLK or DIF.`);
        let memo = null;
        let memoKind = null;
        if (kind === 'dbf') {
          const stem = path.basename(file).replace(/\.dbf$/i, '').toLowerCase();
          for (const entry of fs.readdirSync(path.dirname(file))) {
            const m = /^(.*)\.(dbt|fpt)$/i.exec(entry);
            if (m && m[1].toLowerCase() === stem) { memo = new Uint8Array(fs.readFileSync(path.join(path.dirname(file), entry))); memoKind = m[2].toLowerCase(); break; }
          }
        }
        const model = readLegacySheets(bytes, kind, { name, memo, memoKind, Reader: MDBReader });
        notes = model.notes || [];
        made = importSheets(base, model);
      }
      return { tables: made.map((m) => ({ table: m.table, count: m.count, skipped: m.skipped, key: m.key || null, dates: m.dates || [] })), notes, summary: summary(id) };
    },

    /** A copy of the database at `file` (Save As), the copy then the one open. */
    saveAs: ({ id, file }, win) => {
      const o = get(id);
      whole(o.base);
      if (path.resolve(file).toLowerCase() === path.resolve(o.file).toLowerCase()) return summary(id);
      for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
      o.base.db.exec(`VACUUM INTO '${String(file).replace(/'/g, "''")}'`);
      // The person's own copy reads what the database read.
      trust.carry(o.file, file);
      const access = o.base.access;
      letGo(id);
      const copy = hold(file, win);
      // The copy opens as the person who made it, not locked in front of them.
      if (access && !access.locked) { get(copy).base.access = access; get(copy).base._applyViews(); }
      return summary(copy);
    },

    /* People, roles and protection (D-2026-10-22-11). Each answers the summary, its access in it. */
    signIn: async ({ id, name, password }) => { get(id).base.signIn(name, password); await readAllLinked(get(id)); return summary(id); },
    signOut: ({ id }) => { get(id).base.signOut(); return summary(id); },
    protect: ({ id, name, password }) => { get(id).base.protect({ name, password }); return summary(id); },
    unprotect: ({ id, password }) => { get(id).base.unprotect({ password }); return summary(id); },
    seeAs: ({ id, role = null }) => { get(id).base.seeAs(role); return summary(id); },
    roles: ({ id }) => get(id).base.roles(),
    saveRole: ({ id, name, rights }) => ({ name: get(id).base.saveRole(name, rights) }),
    removeRole: ({ id, name }) => { get(id).base.removeRole(name); return { removed: true }; },
    people: ({ id }) => get(id).base.people(),
    addPerson: ({ id, name, role, password }) => ({ id: get(id).base.addPerson({ name, role, password }) }),
    changePerson: ({ id, person, name = null, role = null }) => { get(id).base.changePerson(person, { name, role }); return { changed: true }; },
    removePerson: ({ id, person }) => { get(id).base.removePerson(person); return { removed: true }; },
    setPassword: ({ id, person, password, current = null }) => { get(id).base.setPassword(person, password, { current }); return { changed: true }; },
    trailOf: ({ id, table, rowid }) => get(id).base.trailOf(table, rowid),
    /** Every table's fields, all of them, for the Admin's People and roles grid. */
    allFields: ({ id }) => {
      const { base } = get(id);
      whole(base);
      return Object.fromEntries(base.tables().map((t) => [t.name, base.fields(t.name).map((f) => f.name)]));
    },

    /** Every database let go, as the app closes. */
    closeAll: () => {
      for (const id of [...open.keys()]) letGo(id);
    },
  };
}
