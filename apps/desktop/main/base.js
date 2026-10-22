// Databases: the database app's service. A database is one SQLite file
// (@rutba/database), opened here and held while a window has it open; every
// design change goes through the file's numbered steps, and the window asks
// for its rows a page at a time, so a table of a million records opens as
// quickly as one of ten.
//
// Imports make a new table each: delimited text (CSV, TSV) and the old files
// Worksheets reads (dBASE and FoxPro tables, Access databases, Lotus 1-2-3,
// SYLK, DIF), through office-formats' readers.

import fs from 'node:fs';
import path from 'node:path';
import { Base, importRows, importSheets, buildTemplate, buildSample, originOf, planUpgrade, applyUpgrade, toStored } from '@rutba/database';
import { ENGINES, connectEngine as defaultConnect } from './engines/index.js';
import { TEMPLATE_IDS, SAMPLE_IDS, templateNeeds, sampleNeeds, pageOf } from '@rutba/templates/data-files/active';
import { needsFor } from '@rutba/templates/versions';
import { readManifest } from '@rutba/templates/manifest';
import { parseDelimited } from '@rutba/ooxml/mailmerge';
import { legacySheetKind, readLegacySheets } from '@rutba/office-formats/legacy-sheets';
import MDBReader from 'mdb-reader';

/** The SQLite file's own first sixteen bytes. */
const SQLITE_MAGIC = 'SQLite format 3\u0000';

/** Whether bytes are a SQLite database. */
export const isSqlite = (bytes) => Buffer.from(bytes.subarray(0, 16)).toString('latin1') === SQLITE_MAGIC;

/**
 * `secrets` keeps a linked engine's password (D-2026-10-22-13), never the
 * file; `connect` opens an engine (engines/index.js, or a stand-in in tests).
 */
export function createBaseService({ secrets = null, connect = defaultConnect } = {}) {
  const open = new Map(); // id -> { base, file, engines: Map(link -> engine), state: Map(table -> { error, at, count }) }
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
        // A linked table: its records as last read from the engine, or why they could not be.
        const st = open.get(id).state.get(t.name) || {};
        return { ...t, count: base.linkedLoaded(t.name) ? base.count(t.name) : 0, linkError: st.error || null, readAt: st.at || null, key: base.linked().find((l) => l.table === t.name)?.key || null };
      }),
      links: base.access?.locked || base.rights() ? [] : base.links().map((l) => ({ id: l.id, name: l.name, engine: l.engine, settings: l.settings, password: Boolean(secrets?.get?.(secretKey(l.id))) })),
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
    if (!o.engines.has(link.id)) {
      const password = secrets?.get?.(secretKey(link.id)) || '';
      const p = connect(link.engine, link.settings, password);
      o.engines.set(link.id, p);
      // A connection that failed is tried again next time, not kept.
      p.catch(() => { if (o.engines.get(link.id) === p) o.engines.delete(link.id); });
    }
    return o.engines.get(link.id);
  };
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
      o.base.loadLinked(table, rows);
      o.state.set(table, { error: null, at: new Date().toISOString(), count: rows.length });
    } catch (err) {
      o.state.set(table, { error: err.message || String(err), at: null });
    }
  };
  /** Every linked table read, as a database opens or is refreshed; one that cannot be read says why. */
  const readAllLinked = async (o) => {
    if (o.base.access?.locked || o.base.rights()) return;
    await Promise.all(o.base.linked().map((l) => readLinked(o, l.table)));
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
      return summary(hold(file, win));
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
      await readAllLinked(get(id));
      return summary(id);
    },

    close: ({ id }) => ({ closed: letGo(id) }),

    summary: ({ id }) => summary(id),
    fields: ({ id, table }) => get(id).base.fields(table),
    rows: ({ id, table, offset = 0, limit = 200, order = null, where = {}, search = '', filters = [] }) => {
      const { base } = get(id);
      return { rows: base.rows(table, { offset, limit, order, where, search, filters }), count: base.count(table, where, { search, filters }) };
    },
    // A query's results as a table's records are read: a page, sorted, searched and filtered (D-2026-10-22-07).
    queryRows: ({ id, name = null, design = null, offset = 0, limit = 200, order = null, search = '', filters = [] }) => get(id).base.queryRows({ name, design, offset, limit, order, search, filters }),
    // A linked table's changes go to its engine (D-2026-10-22-13); the rest to the file.
    insert: async ({ id, table, values }) => {
      const o = get(id);
      if (!o.base.isLinked(table)) return { rowid: o.base.insert(table, values || {}) };
      const key = await writeLinked(o, table, ({ engine, remote, design, stored }) => engine.insert(remote, design, stored(values)));
      const l = o.base.linked().find((x) => x.table === table);
      const row = key != null && l.key ? o.base.rows(table, { where: { [l.key]: key }, limit: 1 })[0] : null;
      return { rowid: row?._rowid ?? null };
    },
    update: async ({ id, table, rowid, values }) => {
      const o = get(id);
      if (!o.base.isLinked(table)) return { changed: o.base.update(table, rowid, values || {}) };
      const key = o.base.linkedKey(table, rowid);
      return { changed: await writeLinked(o, table, ({ engine, remote, design, stored }) => engine.update(remote, design, key, stored(values))) };
    },
    remove: async ({ id, table, rowids }) => {
      const o = get(id);
      if (!o.base.isLinked(table)) return { removed: o.base.remove(table, rowids || []) };
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
      if (password != null) {
        if (!secrets?.available?.()) throw new Error('This computer offers no protected store for the password, so the link cannot keep it.');
        secrets.set(secretKey(linkId), String(password));
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
        const table = o.base.linkTable({ table: r.as || r.name, link, remote: { schema: r.schema, name: r.name }, key: design.key, fields: design.fields });
        made.push(table);
        await readLinked(o, table);
      }
      return { tables: made, summary: summary(id) };
    },
    /** A table unlinked: the engine keeps its records. */
    unlinkTable: ({ id, table }) => { const o = get(id); whole(o.base); o.base.unlinkTable(table); o.state.delete(table); return summary(id); },
    /** The linked tables read again from their engines (one, or all). */
    refreshLinked: async ({ id, table = null }) => {
      const o = get(id);
      if (table) await readLinked(o, table); else await readAllLinked(o);
      return summary(id);
    },

    createTable: ({ id, name, fields = [] }) => { get(id).base.createTable({ name, fields }); return summary(id); },
    addField: ({ id, table, field }) => { get(id).base.addField(table, field); return summary(id); },
    changeField: ({ id, table, name, change }) => { get(id).base.changeField(table, name, change); return summary(id); },
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

    /** The person's own SELECT, read-only: { columns, rows }, at most `limit` rows. */
    query: ({ id, sql, limit = 5000 }) => {
      const out = get(id).base.query(sql);
      return { columns: out.columns, rows: out.rows.slice(0, limit), count: out.rows.length };
    },

    /**
     * A file imported as new tables: CSV or TSV, or an old file Worksheets
     * reads. Answers { tables: [{ table, count, skipped }], notes, summary }.
     */
    importFile: ({ id, file }) => {
      const { base } = get(id);
      const bytes = new Uint8Array(fs.readFileSync(file));
      const name = path.basename(file).replace(/\.[^.]+$/, '');
      const ext = path.extname(file).toLowerCase();
      let made;
      let notes = [];
      if (ext === '.csv' || ext === '.tsv' || ext === '.txt') {
        const text = new TextDecoder('utf-8').decode(bytes);
        made = [importRows(base, name, parseDelimited(text, ext === '.tsv' ? '\t' : null))];
      } else {
        const kind = legacySheetKind(bytes, ext);
        if (!kind) throw new Error(`${path.basename(file)} is not a file this app imports: CSV, TSV, dBASE, Access, Lotus 1-2-3, SYLK or DIF.`);
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
      return { tables: made.map((m) => ({ table: m.table, count: m.count, skipped: m.skipped })), notes, summary: summary(id) };
    },

    /** A copy of the database at `file` (Save As), the copy then the one open. */
    saveAs: ({ id, file }, win) => {
      const o = get(id);
      whole(o.base);
      if (path.resolve(file).toLowerCase() === path.resolve(o.file).toLowerCase()) return summary(id);
      for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
      o.base.db.exec(`VACUUM INTO '${String(file).replace(/'/g, "''")}'`);
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
