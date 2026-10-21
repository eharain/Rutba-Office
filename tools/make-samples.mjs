// Data Files' sample databases as files, made fresh so their dates count back
// from the day they are made, and their documentation from the same source.
//
//   node tools/make-samples.mjs [dir]   writes <dir>/<sample>.rbase for each sample and
//                                       <dir>/samples.json, the catalogue a website lists
//                                       them from (default: dist/samples)
//   node tools/make-samples.mjs --docs  writes docs/SAMPLES.md
//   node tools/make-samples.mjs --check builds each one and says what checkSample finds
//
// Each sample says what it is in its own `about` (packages/database/src/samples):
// what the business is, each table, field, query and report in plain words.
// The catalogue and docs/SAMPLES.md are made from that and from the design
// itself, so they say what the file holds. A file is written only when the
// sample passes checkSample: none is offered with an empty query or report.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { Base, SAMPLE_IDS, buildSample, sampleSpec, sampleAbout, checkSample } from '@rutba/database';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const TYPES = { text: 'Short text', longtext: 'Long text', integer: 'Whole number', number: 'Number', currency: 'Currency', date: 'Date', datetime: 'Date and time', yesno: 'Yes/No' };

/** A name as a file name: Northbank Trading is northbank-trading. */
export const slug = (s) => s.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');

/** A criterion as a person reads it: Status <> "Cancelled". */
const criterion = (c, text) => `${c.expr ? (c.as || 'the expression') : c.field} ${/^(<|>|=|Is |Not |In |Like |Between )/i.test(text) ? text : `= ${text}`}`;

/**
 * One sample described for a page: what `about` says, the design it is made
 * of, and how many records each table has when made with its records.
 */
export function describe(id, { base, counts = {} }) {
  const spec = sampleSpec(id);
  const about = sampleAbout(id);
  const fieldsOf = Object.fromEntries(spec.tables.map((t) => [t.name, t.fields]));
  return {
    id,
    slug: slug(about.name),
    business: about.name,
    kind: about.kind,
    line: about.line,
    summary: about.summary,
    suits: about.suits,
    empty: about.empty,
    adapt: about.adapt,
    tables: spec.tables.map((t) => ({
      name: t.name,
      about: about.tables[t.name].about,
      records: counts[t.name] || 0,
      fields: [{ name: 'ID', type: 'AutoNumber', about: 'The number Data Files gives each record. Other tables point at a record by it.' },
        ...fieldsOf[t.name].map((f) => ({
          name: f.name,
          type: f.points ? `Points at ${f.points}` : TYPES[f.type] || f.type,
          ...(f.required ? { required: true } : {}),
          ...(f.unique ? { unique: true } : {}),
          ...(f.default !== undefined ? { default: f.default } : {}),
          about: about.tables[t.name].fields[f.name],
        }))],
    })),
    relationships: spec.tables.flatMap((t) => t.fields.filter((f) => f.points).map((f) => ({ from: `${t.name}.${f.name}`, to: `${f.points}.ID` }))),
    queries: Object.entries(spec.queries).map(([name, design]) => {
      const columns = base.runQuery({ name, limit: 1 }).columns;
      return {
        name,
        about: about.queries[name],
        tables: design.tables,
        columns,
        calculated: design.columns.filter((c) => c.expr && c.show !== false && c.total !== 'where').map((c) => ({ as: c.as, expr: c.expr })),
        criteria: design.columns.flatMap((c) => (c.criteria || []).filter(Boolean).map((x) => criterion(c, x))),
        totals: design.columns.some((c) => c.total),
      };
    }),
    reports: spec.reports.map((r) => ({
      name: r.name,
      title: r.title || r.name,
      about: about.reports[r.name],
      source: r.query ? { query: r.query } : { table: r.table },
      group: r.group || null,
      columns: r.columns.map((c) => (typeof c === 'string' ? c : c.field)),
      totals: (r.totals || []).map((t) => `${t.fn === 'sum' ? 'Total' : t.fn === 'avg' ? 'Average' : t.fn === 'count' ? 'Count' : t.fn === 'min' ? 'Lowest' : 'Highest'} of ${t.field}`),
    })),
  };
}

/** Make sample `id` at `file` (with its records unless `withData` is false); answers its counts. */
export function makeFile(id, file, { withData = true, today = new Date() } = {}) {
  for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
  const base = Base.open(file);
  try {
    const out = buildSample(base, id, { withData, today });
    if (withData) {
      const problems = checkSample(base, sampleSpec(id));
      if (problems.length) throw new Error(`${id} is not ready:\n  ${problems.join('\n  ')}`);
    }
    // One file with nothing beside it: the write-ahead log folded in, and the space compacted.
    base.db.exec('PRAGMA journal_mode = DELETE');
    base.db.exec('VACUUM');
    return { counts: out.records, description: describe(id, { base, counts: out.records }) };
  } finally {
    base.close();
  }
}

/** docs/SAMPLES.md, from the catalogue. */
export function markdown(samples, made) {
  const out = [];
  const w = (s = '') => out.push(s);
  w('# Sample databases');
  w();
  w('Rutba Data Files comes with a working database for each kind of small business we make Office for.');
  w('Each one has related tables, the questions that business asks as saved queries, a form for every table and');
  w('reports to print, and it comes either with a year or more of invented records to learn from, or empty, ready');
  w('for your own. On the Data Files start screen, choose a business and then **With sample records** or **Empty**.');
  w('The same files are on [office.rutba.io/samples](https://office.rutba.io/samples) to download.');
  w();
  w('Everything in the records is made up: the businesses, the people, their addresses, e-mail addresses on');
  w('example domains and phone numbers from the ranges kept for drama and fiction. The dates count back from the');
  w(`day the file is made, so "due this week" always has something in it. This page was made on ${made}.`);
  w();
  w('The samples are made by `packages/database/src/samples` (one file each), checked by `check.js` and');
  w('`tests/database-samples.test.js`, and written as files by `node tools/make-samples.mjs`. This page is made by');
  w('`node tools/make-samples.mjs --docs` from each sample\'s own description, so it says what the files hold.');
  w();
  w('With each release, the downloads on office.rutba.io are made again so their dates are current: in the');
  w('management repository, `npm run import:samples` in `portal/apps/office` builds every sample from this');
  w('repository, zips it with a read-me into `public/samples/` and writes the pages\' catalogue,');
  w('`src/site/samples.json`. A sample that fails `checkSample` stops the import.');
  w();
  w('| Sample | For | Tables | Queries | Reports |');
  w('|---|---|---|---|---|');
  for (const s of samples) w(`| [${s.business}](#${slug(s.business)}) | ${s.kind} | ${s.tables.length} | ${s.queries.length} | ${s.reports.length} |`);
  for (const s of samples) {
    w();
    w(`## ${s.business}`);
    w();
    w(`**${s.kind}.** ${s.line}`);
    w();
    for (const p of s.summary) { w(p); w(); }
    w(`It suits, with small changes: ${s.suits.join('; ')}.`);
    w();
    w('### Tables');
    for (const t of s.tables) {
      w();
      w(`#### ${t.name}`);
      w();
      w(`${t.about}${t.records ? ` The sample has ${t.records.toLocaleString('en-GB')} records.` : ''}`);
      w();
      w('| Field | Type | What it holds |');
      w('|---|---|---|');
      for (const f of t.fields) w(`| ${f.name} | ${f.type}${f.required ? ', required' : ''}${f.unique ? ', unique' : ''} | ${f.about} |`);
    }
    w();
    w('### Relationships');
    w();
    for (const r of s.relationships) w(`- ${r.from} points at ${r.to}`);
    w();
    w('### Queries');
    for (const q of s.queries) {
      w();
      w(`#### ${q.name}`);
      w();
      w(q.about);
      w();
      w(`From ${q.tables.join(', ')}. Shows ${q.columns.join(', ')}.`);
      if (q.calculated.length) w(`Calculated: ${q.calculated.map((c) => `${c.as} is \`${c.expr}\``).join('; ')}.`);
      if (q.criteria.length) w(`Only where ${q.criteria.map((c) => `\`${c}\``).join(' and ')}.`);
    }
    w();
    w('### Reports');
    for (const r of s.reports) {
      w();
      w(`#### ${r.title}`);
      w();
      w(r.about);
      w();
      w(`From the ${r.source.query ? `query ${r.source.query}` : `table ${r.source.table}`}${r.group ? `, grouped by ${r.group}` : ''}: ${r.columns.join(', ')}${r.totals.length ? `, with ${r.totals.join(', ').toLowerCase()}` : ''}.`);
    }
    w();
    w('### Starting empty');
    w();
    w(s.empty);
    w();
    w('### Making it yours');
    w();
    for (const a of s.adapt) w(`- ${a}`);
  }
  w();
  return out.join('\n');
}

const isMain = process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (isMain) {
  const args = process.argv.slice(2);
  const today = new Date();
  const made = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
  if (args.includes('--check')) {
    let bad = 0;
    for (const id of SAMPLE_IDS) {
      const base = Base.open(':memory:');
      const t0 = Date.now();
      const out = buildSample(base, id);
      const ms = Date.now() - t0;
      const problems = checkSample(base, sampleSpec(id));
      base.close();
      bad += problems.length;
      console.log(`${id}: ${Object.values(out.records).reduce((a, n) => a + n, 0)} records in ${ms} ms, ${problems.length ? problems.join('; ') : 'ready'}`);
    }
    process.exit(bad ? 1 : 0);
  }
  const docs = args.includes('--docs');
  const dir = path.resolve(args.find((a) => !a.startsWith('--')) || path.join(ROOT, 'dist', 'samples'));
  fs.mkdirSync(dir, { recursive: true });
  const samples = [];
  for (const id of SAMPLE_IDS) {
    const about = sampleAbout(id);
    const file = path.join(dir, `${slug(about.name)}.rbase`);
    const { description } = makeFile(id, file, { today });
    samples.push({ ...description, file: path.basename(file), bytes: fs.statSync(file).size });
    console.log(`${path.basename(file)}  ${fs.statSync(file).size} bytes`);
  }
  if (docs) {
    for (const s of samples) fs.rmSync(path.join(dir, s.file), { force: true });
    fs.writeFileSync(path.join(ROOT, 'docs', 'SAMPLES.md'), markdown(samples, made));
    console.log('docs/SAMPLES.md written');
  } else {
    fs.writeFileSync(path.join(dir, 'samples.json'), `${JSON.stringify({ made, samples }, null, 2)}\n`);
    console.log(`${samples.length} samples and samples.json in ${dir}`);
  }
}
