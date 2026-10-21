// Data Files' sample databases as files, made fresh so their dates count back
// from the day they are made, and their documentation from the same source.
//
//   node tools/make-samples.mjs [dir]   writes <dir>/<sample>.rbase for each sample and
//                                       for each simple template, and <dir>/samples.json,
//                                       the catalogue a website lists them from
//                                       (default: dist/samples)
//   node tools/make-samples.mjs --docs  writes docs/SAMPLES.md
//   node tools/make-samples.mjs --check builds each one and says what checkSample finds
//
// Each database says what it is in its own `about` (packages/database/src/samples,
// and templates-about.js for the simple templates): what the business is, each
// table, field, query, report and print template in plain words, a tutorial,
// things to know and questions people ask. The catalogue and docs/SAMPLES.md
// are made from that and from the design read back out of the file itself,
// so they say what the file holds. A sample file is written only when it
// passes checkSample: none is offered with an empty query or report.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { Base, SAMPLE_IDS, TEMPLATE_IDS, buildSample, buildTemplate, sampleSpec, sampleAbout, checkSample } from '@rutba/database';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const TYPES = { autonumber: 'AutoNumber', text: 'Short text', longtext: 'Long text', integer: 'Whole number', number: 'Number', currency: 'Currency', date: 'Date', datetime: 'Date and time', yesno: 'Yes/No' };
const FN = { sum: 'total', avg: 'average', count: 'count', min: 'lowest', max: 'highest' };

/** The simple templates' documentation, when it is there. */
const TEMPLATE_ABOUT = await import('../packages/database/src/templates-about.js').then((m) => m.TEMPLATE_ABOUT).catch(() => ({}));

/** What every database's documentation says before its own notes. */
export const KNOW = [
  'Everything in a sample\'s records is invented: the business, the people, their addresses, e-mail addresses on example domains that reach nobody, and phone numbers from the ranges kept for drama and fiction. Tax, company and bank numbers are placeholders of zeros.',
  'The dates count back from the day the file was made, and some run on into the weeks after it, as real bookings and due dates do. Made from the Data Files start screen, a sample is dated from that day. Queries that compare with today (overdue, due this week) show more as a downloaded file gets older.',
  'To start empty, choose the same business on the start screen and then Empty: the same tables, queries, forms, reports and print templates, with nothing in them but the Start Here guide.',
  'Your records live in one file on your computer. Keep copies: File, Save As makes one, and a copy on another drive or in your own cloud storage keeps it safe if this computer is lost. Close the database before copying the file by hand.',
  'Every change to the design (a new field, a renamed table) is a numbered step that Data Files keeps, so you can see what was changed and when.',
];

/** A name as a file name: Northbank Trading is northbank-trading. */
export const slug = (s) => s.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');

/** A criterion as a person reads it: Status <> "Cancelled". */
const criterion = (c, text) => `${c.expr ? (c.as || 'the expression') : c.field} ${/^(<|>|=|Is |Not |In |Like |Between )/i.test(text) ? text : `= ${text}`}`;

/** A print template's blocks as a reader sees them. */
function blocksOf(layout) {
  return (layout.blocks || []).map((b) => {
    switch (b.type) {
      case 'heading': return { type: 'heading', text: b.text };
      case 'fields': return { type: 'fields', items: (b.items || []).map((i) => i.label || i.field.split('.').pop()) };
      case 'lines': return { type: 'lines', table: b.table, columns: (b.columns || []).map((c) => c.label || c.as || c.field.split('.').pop()) };
      case 'totals': return { type: 'totals', rows: (b.rows || []).map((r) => r.label) };
      case 'text': return { type: 'text', text: b.text };
      case 'signatures': return { type: 'signatures', labels: b.labels || ['Signed'] };
      default: return { type: b.type };
    }
  });
}

/**
 * One database described for a page: what `about` says, and the design read
 * back out of `base` (its tables, relationships, queries, forms, reports and
 * print templates), with how many records each table has.
 */
export function describe(base, about, { id, counts = {}, kind = 'sample' } = {}) {
  const rels = base.relationships();
  const pointing = new Map(rels.map((r) => [`${r.from.table}.${r.from.field}`, r.to.table]));
  const list = (v) => (Array.isArray(v) ? v : v ? [v] : []);
  return {
    id,
    kind,
    slug: slug(about.name),
    name: about.name,
    business: about.name,
    sector: about.kind,
    line: about.line,
    summary: list(about.summary),
    suits: list(about.suits),
    know: [...KNOW, ...list(about.notes)],
    tutorial: list(about.tutorial),
    faq: list(about.faq),
    empty: about.empty || '',
    adapt: list(about.adapt),
    tables: base.tables().map(({ name }) => ({
      name,
      about: about.tables?.[name]?.about || '',
      records: counts[name] || 0,
      fields: base.fields(name).map((f) => ({
        name: f.name,
        type: pointing.has(`${name}.${f.name}`) ? `Points at ${pointing.get(`${name}.${f.name}`)}` : TYPES[f.type] || f.type,
        ...(f.required ? { required: true } : {}),
        ...(f.unique ? { unique: true } : {}),
        about: f.name === 'ID' && f.type === 'autonumber' && !about.tables?.[name]?.fields?.ID
          ? 'The number Data Files gives each record. Other tables point at a record by it.'
          : about.tables?.[name]?.fields?.[f.name] || '',
      })),
    })),
    relationships: rels.map((r) => ({ from: r.from.table, field: r.from.field, to: r.to.table })),
    queries: base.queries().map(({ name, design }) => ({
      name,
      about: about.queries?.[name] || '',
      tables: design.tables,
      columns: base.runQuery({ name, limit: 1 }).columns,
      calculated: design.columns.filter((c) => c.expr && c.show !== false && c.total !== 'where').map((c) => ({ as: c.as, expr: c.expr })),
      criteria: design.columns.flatMap((c) => (c.criteria || []).filter(Boolean).map((x) => criterion(c, x))),
      totals: design.columns.some((c) => c.total),
    })),
    forms: base.forms().map(({ name, layout }) => ({ name, table: layout.table })),
    reports: base.reports().map(({ name, layout }) => ({
      name,
      title: layout.title || name,
      about: about.reports?.[name] || '',
      source: layout.source || {},
      group: layout.group || null,
      columns: (layout.columns || []).map((c) => c.label || c.field),
      totals: (layout.totals || []).map((t) => `${FN[t.fn] || t.fn} of ${t.field}`),
    })),
    documents: base.documents().map(({ name, layout }) => ({
      name,
      about: about.documents?.[name] || '',
      table: layout.table,
      paper: layout.paper || 'a4',
      blocks: blocksOf(layout),
    })),
  };
}

/** One file with nothing beside it: the write-ahead log folded in, and the space compacted. */
function seal(base) {
  base.db.exec('PRAGMA journal_mode = DELETE');
  base.db.exec('VACUUM');
}

/** Make sample `id` at `file` (with its records unless `withData` is false); answers its counts and description. */
export function makeFile(id, file, { withData = true, today = new Date() } = {}) {
  for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
  const base = Base.open(file);
  try {
    const out = buildSample(base, id, { withData, today });
    if (withData) {
      const problems = checkSample(base, sampleSpec(id));
      if (problems.length) throw new Error(`${id} is not ready:\n  ${problems.join('\n  ')}`);
    }
    seal(base);
    return { counts: out.records, description: describe(base, sampleAbout(id), { id, counts: out.records }) };
  } finally {
    base.close();
  }
}

/** Make the simple template `id` at `file`, empty as the start screen makes it; answers its description. */
export function makeTemplateFile(id, file) {
  for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
  const base = Base.open(file);
  try {
    buildTemplate(base, id);
    seal(base);
    const about = TEMPLATE_ABOUT[id] || { name: id, tables: {} };
    return { description: describe(base, about, { id, kind: 'template' }) };
  } finally {
    base.close();
  }
}

/** docs/SAMPLES.md, from the catalogue. */
export function markdown({ samples, templates }, made) {
  const out = [];
  const w = (s = '') => out.push(s);
  w('# Sample databases and templates');
  w();
  w('Rutba Data Files comes with a working database for each kind of small business we make Office for, and');
  w('five simple templates. Each sample has related tables, the questions that business asks as saved queries,');
  w('a form for every table, reports, and print templates for what it sends out (invoices, receipts,');
  w('agreements, payslips). It comes either with a year or more of invented records to learn from, or empty,');
  w('ready for your own: on the Data Files start screen, choose a business and then **With sample records** or');
  w('**Empty**. The same files, and this documentation a page each, are on');
  w('[office.rutba.io/samples](https://office.rutba.io/samples).');
  w();
  w('The samples are made by `packages/database/src/samples` (one file each), checked by `check.js` and');
  w('`tests/database-samples.test.js`, and written as files by `node tools/make-samples.mjs`. This page is made by');
  w('`node tools/make-samples.mjs --docs` from each database\'s own description and its design, so it says what the');
  w(`files hold. It was made on ${made}.`);
  w();
  w('With each release, the downloads on office.rutba.io are made again so their dates are current: in the');
  w('management repository, `npm run import:samples` in `portal/apps/office` builds every sample from this');
  w('repository, zips it with a read-me into `public/samples/` and writes the pages\' catalogue,');
  w('`src/site/samples.json`. A sample that fails `checkSample` stops the import.');
  w();
  w('| Database | For | Tables | Queries | Reports | Print templates |');
  w('|---|---|---|---|---|---|');
  for (const s of [...samples, ...templates]) w(`| [${s.name}](#${slug(s.name)}) | ${s.sector || 'Template'} | ${s.tables.length} | ${s.queries.length} | ${s.reports.length} | ${s.documents.length} |`);
  for (const s of [...samples, ...templates]) {
    w();
    w(`## ${s.name}`);
    w();
    if (s.sector) w(`**${s.sector}.** ${s.line}`);
    else w(s.line);
    w();
    for (const p of s.summary) { w(p); w(); }
    if (s.suits.length) { w(`It suits, with small changes: ${s.suits.join('; ')}.`); w(); }
    w('### Things you should know');
    w();
    for (const k of s.kind === 'template' ? [...KNOW.slice(3), ...s.know.slice(KNOW.length)] : s.know) w(`- ${k}`);
    if (s.tutorial.length) {
      w();
      w('### Tutorial');
      for (const t of s.tutorial) {
        w();
        w(`#### ${t.title}`);
        w();
        t.steps.forEach((x, i) => w(`${i + 1}. ${x}`));
      }
    }
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
    for (const r of s.relationships) w(`- ${r.from}.${r.field} points at a record in ${r.to}`);
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
    w('### Forms');
    w();
    w(`Every table has a form of the same name, made by the form wizard: ${s.forms.map((f) => f.name).join(', ')}. A form shows one record at a time; a field that points at another table is a list of that table's records, and the records that point at this one are listed below it.`);
    w();
    w('### Reports');
    for (const r of s.reports) {
      w();
      w(`#### ${r.title}`);
      w();
      w(r.about);
      w();
      w(`From the ${r.source.query ? `query ${r.source.query}` : `table ${r.source.table}`}${r.group ? `, a group for each ${r.group}` : ''}: ${r.columns.join(', ')}${r.totals.length ? `, with the ${r.totals.join(', ')}` : ''}.`);
    }
    if (s.documents.length) {
      w();
      w('### Print templates');
      for (const d of s.documents) {
        w();
        w(`#### ${d.name}`);
        w();
        w(`${d.about} One for each record of ${d.table}, on ${d.paper} paper, with the Business record at the head.`);
        w();
        for (const b of d.blocks) {
          if (b.type === 'heading') w(`- Heading: ${b.text}`);
          if (b.type === 'fields') w(`- Fields: ${b.items.join(', ')}`);
          if (b.type === 'lines') w(`- Lines from ${b.table}: ${b.columns.join(', ')}`);
          if (b.type === 'totals') w(`- Totals: ${b.rows.join(', ')}`);
          if (b.type === 'text') w(`- Words: ${b.text}`);
          if (b.type === 'signatures') w(`- Signed by: ${b.labels.join(', ')}`);
        }
      }
    }
    if (s.empty) {
      w();
      w('### Starting empty');
      w();
      w(s.empty);
    }
    if (s.adapt.length) {
      w();
      w('### Making it yours');
      w();
      for (const a of s.adapt) w(`- ${a}`);
    }
    if (s.faq.length) {
      w();
      w('### Questions people ask');
      for (const f of s.faq) {
        w();
        w(`**${f.q}**`);
        w();
        w(f.a);
      }
    }
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
  const templates = [];
  for (const id of SAMPLE_IDS) {
    const file = path.join(dir, `${slug(sampleAbout(id).name)}.rbase`);
    samples.push({ ...makeFile(id, file, { today }).description, file: path.basename(file), bytes: fs.statSync(file).size });
    console.log(`${path.basename(file)}  ${fs.statSync(file).size} bytes`);
  }
  for (const id of TEMPLATE_IDS) {
    const file = path.join(dir, `template-${id}.rbase`);
    templates.push({ ...makeTemplateFile(id, file).description, file: path.basename(file), bytes: fs.statSync(file).size });
  }
  if (docs) {
    for (const s of [...samples, ...templates]) fs.rmSync(path.join(dir, s.file), { force: true });
    fs.writeFileSync(path.join(ROOT, 'docs', 'SAMPLES.md'), markdown({ samples, templates }, made));
    console.log('docs/SAMPLES.md written');
  } else {
    fs.writeFileSync(path.join(dir, 'samples.json'), `${JSON.stringify({ made, samples, templates }, null, 2)}\n`);
    console.log(`${samples.length} samples, ${templates.length} templates and samples.json in ${dir}`);
  }
}
