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
import { Base, importRows, importSheets, buildTemplate, buildSample } from '@rutba/database';
import { TEMPLATE_IDS, SAMPLE_IDS, templateNeeds, sampleNeeds } from '@rutba/templates/data-files/active';
import { needsFor } from '@rutba/templates/versions';
import { readManifest } from '@rutba/templates/manifest';
import { parseDelimited } from '@rutba/ooxml/mailmerge';
import { legacySheetKind, readLegacySheets } from '@rutba/office-formats/legacy-sheets';
import MDBReader from 'mdb-reader';

/** The SQLite file's own first sixteen bytes. */
const SQLITE_MAGIC = 'SQLite format 3\u0000';

/** Whether bytes are a SQLite database. */
export const isSqlite = (bytes) => Buffer.from(bytes.subarray(0, 16)).toString('latin1') === SQLITE_MAGIC;

export function createBaseService() {
  const open = new Map(); // id -> { base, file }
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
      tables: base.tables().map((t) => ({ ...t, count: base.count(t.name) })),
      relationships: base.relationships(),
      queries: base.queries().map((x) => x.name),
      forms: base.forms().map((x) => x.name),
      reports: base.reports().map((x) => x.name),
      documents: base.documents().map((x) => ({ name: x.name, table: x.layout.table })),
      steps: base.steps().length,
    };
  };
  const letGo = (id) => {
    const o = open.get(id);
    if (!o) return false;
    try { o.base.close(); } catch { /* already closed */ }
    open.delete(id);
    return true;
  };
  // Each window holds its own connection to the file, let go when the window
  // closes: two windows on one database are two readers and writers of it,
  // as SQLite allows, and one closing leaves the other's open.
  const hold = (file, win) => {
    const base = Base.open(file);
    const id = String(next++);
    open.set(id, { base, file });
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
        templates: TEMPLATE_IDS.map((id) => ({ id, needs: whole || needsFor(office, { minOffice: templateNeeds(id) }) })),
        samples: SAMPLE_IDS.map((id) => ({ id, needs: whole || needsFor(office, { minOffice: sampleNeeds(id) }) })),
      };
    },

    /** A new database at `file` made from a template: its tables, relationships, queries, forms and reports. */
    createFromTemplate: ({ file, template }, win) => made(file, win, (base) => buildTemplate(base, template)),

    /** A new database at `file` made from a sample: its design, and its invented records unless `withData` is false. */
    createFromSample: ({ file, sample, withData = true }, win) => made(file, win, (base) => buildSample(base, sample, { withData: withData !== false })),

    /** Open a database file: ours, or any SQLite database. */
    open: ({ file }, win) => {
      const head = Buffer.alloc(16);
      const fd = fs.openSync(file, 'r');
      try { fs.readSync(fd, head, 0, 16, 0); } finally { fs.closeSync(fd); }
      if (!isSqlite(head)) throw new Error(`${path.basename(file)} is not a database this app opens. Import it into a database instead.`);
      return summary(hold(file, win));
    },

    close: ({ id }) => ({ closed: letGo(id) }),

    summary: ({ id }) => summary(id),
    fields: ({ id, table }) => get(id).base.fields(table),
    rows: ({ id, table, offset = 0, limit = 200, order = null, where = {} }) => {
      const { base } = get(id);
      return { rows: base.rows(table, { offset, limit, order, where }), count: base.count(table, where) };
    },
    insert: ({ id, table, values }) => ({ rowid: get(id).base.insert(table, values || {}) }),
    update: ({ id, table, rowid, values }) => ({ changed: get(id).base.update(table, rowid, values || {}) }),
    remove: ({ id, table, rowids }) => ({ removed: get(id).base.remove(table, rowids || []) }),

    createTable: ({ id, name, fields = [] }) => { get(id).base.createTable({ name, fields }); return summary(id); },
    addField: ({ id, table, field }) => { get(id).base.addField(table, field); return summary(id); },
    changeField: ({ id, table, name, change }) => { get(id).base.changeField(table, name, change); return summary(id); },
    removeField: ({ id, table, name }) => { get(id).base.removeField(table, name); return summary(id); },
    renameTable: ({ id, table, to }) => { get(id).base.renameTable(table, to); return summary(id); },
    dropTable: ({ id, table }) => { get(id).base.dropTable(table); return summary(id); },
    addRelationship: ({ id, from, to, cascadeDelete = false }) => { get(id).base.addRelationship({ from, to, cascadeDelete }); return summary(id); },
    removeRelationship: ({ id, relation }) => { get(id).base.removeRelationship(relation); return summary(id); },
    steps: ({ id }) => get(id).base.steps(),

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
      if (path.resolve(file).toLowerCase() === path.resolve(o.file).toLowerCase()) return summary(id);
      for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
      o.base.db.exec(`VACUUM INTO '${String(file).replace(/'/g, "''")}'`);
      letGo(id);
      return summary(hold(file, win));
    },

    /** Every database let go, as the app closes. */
    closeAll: () => {
      for (const o of open.values()) { try { o.base.close(); } catch { /* already closed */ } }
      open.clear();
    },
  };
}
