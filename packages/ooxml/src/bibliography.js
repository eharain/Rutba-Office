/**
 * Citations and a bibliography, as Word keeps them.
 *
 * A document's sources live in a custom XML part, `customXml/itemN.xml`,
 * whose root is `<b:Sources>` in the bibliography namespace — the same
 * shape as the master list Word keeps in the profile (Sources.xml). Its
 * attributes name the style the citations are drawn in (`SelectedStyle`,
 * `StyleName`, `Version`); each `<b:Source>` carries a Tag the CITATION
 * fields name, a SourceType, a Guid, the people (`b:Author` holding
 * `b:Author`, `b:Editor`, `b:Translator`… each a `b:NameList` of
 * `b:Person`s, or a `b:Corporate` name) and plain fields (`b:Title`,
 * `b:Year`, `b:Publisher`…).
 *
 * Nothing here touches a package: this module reads and writes that XML and
 * formats a citation or a bibliography entry in a style, so the window can
 * preview exactly what the engine will write. Formatting comes back as
 * segments — `{ text, italic }` — because every style italicises something.
 *
 * Elements this does not model (a source type or a field it has no editor
 * for) ride through verbatim: a source is read into `fields`, `people` and
 * `extra`, the XML of everything else, written back after the rest.
 */

export const BIB_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/bibliography';

const escXml = (s) => String(s)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const unescXml = (s) => String(s)
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
  .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n)))
  .replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16)))
  .replace(/&amp;/g, '&');

/* ── styles ──────────────────────────────────────────────────────────────── */

/**
 * The styles offered in References → Style, as Word lists them, with the
 * attributes Word writes on `<b:Sources>` for each. `id` is ours.
 */
export const BIBLIOGRAPHY_STYLES = [
  { id: 'apa7', label: 'APA', edition: 'Seventh Edition', selected: '\\APASeventhEdition.xsl', name: 'APA', version: '7' },
  { id: 'apa6', label: 'APA', edition: 'Sixth Edition', selected: '\\APASixthEditionOfficeOnline.xsl', name: 'APA', version: '6' },
  { id: 'chicago16', label: 'Chicago', edition: 'Sixteenth Edition', selected: '\\ChicagoSixteenthEdition.xsl', name: 'Chicago', version: '16' },
  { id: 'harvard', label: 'Harvard - Anglia', edition: '2008', selected: '\\HarvardAnglia2008OfficeOnline.xsl', name: 'Harvard - Anglia', version: '2008' },
  { id: 'ieee', label: 'IEEE', edition: '2006', selected: '\\IEEE2006OfficeOnline.xsl', name: 'IEEE', version: '2006' },
  { id: 'mla7', label: 'MLA', edition: 'Seventh Edition', selected: '\\MLASeventhEditionOfficeOnline.xsl', name: 'MLA', version: '7' },
];
export const DEFAULT_STYLE = 'apa7';

/** Which of our styles a `<b:Sources>` element's attributes name — by the file name, then by name and version. */
export function styleFromAttributes({ selected, name, version } = {}) {
  const file = String(selected || '').toLowerCase().replace(/^.*[\\/]/, '');
  if (file) {
    const hit = BIBLIOGRAPHY_STYLES.find((s) => s.selected.toLowerCase().replace(/^.*[\\/]/, '') === file);
    if (hit) return hit.id;
    if (/^apa/.test(file)) return /six/.test(file) ? 'apa6' : 'apa7';
    if (/^mla/.test(file)) return 'mla7';
    if (/^chicago/.test(file)) return 'chicago16';
    if (/^harvard/.test(file)) return 'harvard';
    if (/^ieee/.test(file)) return 'ieee';
  }
  const n = String(name || '').toLowerCase();
  if (n === 'apa') return String(version) === '6' ? 'apa6' : 'apa7';
  if (n === 'mla') return 'mla7';
  if (n.startsWith('chicago')) return 'chicago16';
  if (n.startsWith('harvard')) return 'harvard';
  if (n === 'ieee') return 'ieee';
  return DEFAULT_STYLE;
}

export const styleById = (id) => BIBLIOGRAPHY_STYLES.find((s) => s.id === id) || BIBLIOGRAPHY_STYLES[0];

/* ── source types and their fields ───────────────────────────────────────── */

/** The people a source can name, as `b:Author`'s children. */
export const PERSON_ROLES = ['Author', 'Editor', 'Translator', 'BookAuthor', 'Compiler'];

/** Every plain field we edit, in the order Word writes them, with its label in Create Source. */
export const FIELD_LABELS = {
  Title: 'Title',
  ShortTitle: 'Short Title',
  BookTitle: 'Book Title',
  JournalName: 'Journal Name',
  ConferenceName: 'Conference Name',
  InternetSiteTitle: 'Name of Web Site',
  Year: 'Year',
  Month: 'Month',
  Day: 'Day',
  City: 'City',
  StateProvince: 'State/Province',
  Publisher: 'Publisher',
  Institution: 'Institution',
  Department: 'Department',
  ThesisType: 'Report Type',
  Edition: 'Edition',
  Volume: 'Volume',
  NumberVolumes: 'Number of Volumes',
  Issue: 'Issue',
  Pages: 'Pages',
  StandardNumber: 'Standard Number',
  ProductionCompany: 'Production Company',
  YearAccessed: 'Year Accessed',
  MonthAccessed: 'Month Accessed',
  DayAccessed: 'Day Accessed',
  URL: 'URL',
  DOI: 'DOI',
  Medium: 'Medium',
  Comments: 'Comments',
};
const FIELD_ORDER = Object.keys(FIELD_LABELS);

/**
 * The source types Create Source offers, each with the fields Word shows
 * for it (`fields`) and the further ones Show All Bibliography Fields adds
 * (`more`). A person field is named by its role ('Author', 'Editor'…).
 */
export const SOURCE_TYPES = [
  {
    id: 'Book', label: 'Book',
    fields: ['Author', 'Title', 'Year', 'City', 'Publisher'],
    more: ['Editor', 'Translator', 'Edition', 'Volume', 'NumberVolumes', 'Pages', 'ShortTitle', 'StandardNumber', 'DOI', 'URL', 'Medium', 'Comments'],
  },
  {
    id: 'BookSection', label: 'Book Section',
    fields: ['Author', 'Title', 'BookTitle', 'Year', 'Pages', 'City', 'Publisher'],
    more: ['BookAuthor', 'Editor', 'Translator', 'Edition', 'Volume', 'ShortTitle', 'StandardNumber', 'DOI', 'URL', 'Medium', 'Comments'],
  },
  {
    id: 'JournalArticle', label: 'Journal Article',
    fields: ['Author', 'Title', 'JournalName', 'Year', 'Pages', 'Volume', 'Issue'],
    more: ['Editor', 'Month', 'Day', 'ShortTitle', 'StandardNumber', 'DOI', 'URL', 'Medium', 'Comments'],
  },
  {
    id: 'ConferenceProceedings', label: 'Conference Proceedings',
    fields: ['Author', 'Title', 'Pages', 'Year', 'ConferenceName', 'City', 'Publisher'],
    more: ['Editor', 'Volume', 'ShortTitle', 'StandardNumber', 'DOI', 'URL', 'Medium', 'Comments'],
  },
  {
    id: 'Report', label: 'Report',
    fields: ['Author', 'Title', 'Year', 'Publisher', 'City'],
    more: ['Institution', 'Department', 'ThesisType', 'Pages', 'ShortTitle', 'StandardNumber', 'DOI', 'URL', 'Medium', 'Comments'],
  },
  {
    id: 'InternetSite', label: 'Web site',
    fields: ['Author', 'Title', 'InternetSiteTitle', 'Year', 'Month', 'Day', 'YearAccessed', 'MonthAccessed', 'DayAccessed', 'URL'],
    more: ['Editor', 'ProductionCompany', 'ShortTitle', 'Medium', 'Comments'],
  },
];
export const sourceType = (id) => SOURCE_TYPES.find((t) => t.id === id) || null;

/** A field's label in Create Source — the Web site names its title for what it is. */
export function fieldLabel(type, key) {
  if (type === 'InternetSite' && key === 'Title') return 'Name of Web Page';
  if (type === 'BookSection' && key === 'BookAuthor') return 'Book Author';
  if (PERSON_ROLES.includes(key)) return key === 'BookAuthor' ? 'Book Author' : key;
  return FIELD_LABELS[key] || key;
}

/* ── reading and writing the XML ─────────────────────────────────────────── */

const childText = (xml, tag) => {
  const m = new RegExp('<b:' + tag + '\\b[^>]*>([\\s\\S]*?)</b:' + tag + '>').exec(xml);
  return m ? unescXml(m[1]) : null;
};

/** The `b:` elements at the top of a fragment, each closed where its own nesting closes. */
function topElements(xml) {
  const out = [];
  const s = String(xml);
  const open = /<b:([A-Za-z]+)\b[^>]*?(\/?)>/g;
  let m;
  while ((m = open.exec(s))) {
    const tag = m[1];
    if (m[2] === '/') { out.push({ whole: m[0], tag, content: '' }); continue; }
    const re = new RegExp('<b:' + tag + '\\b[^>]*?(/?)>|</b:' + tag + '>', 'g');
    re.lastIndex = m.index + m[0].length;
    let depth = 1;
    let r;
    while (depth > 0 && (r = re.exec(s))) {
      if (r[0].startsWith('</')) depth -= 1;
      else if (r[1] !== '/') depth += 1;
    }
    const end = r ? r.index + r[0].length : s.length;
    out.push({ whole: s.slice(m.index, end), tag, content: s.slice(m.index + m[0].length, r ? r.index : s.length) });
    open.lastIndex = end;
  }
  return out;
}

/** One `b:NameList`'s people: `{ last, first, middle }` each. */
function readPeople(xml) {
  const out = [];
  for (const m of String(xml).matchAll(/<b:Person\b[^>]*>([\s\S]*?)<\/b:Person>/g)) {
    out.push({ last: childText(m[1], 'Last') || '', first: childText(m[1], 'First') || '', middle: childText(m[1], 'Middle') || '' });
  }
  return out;
}

/**
 * A `<b:Source>` read: `tag`, `type`, `guid`, `fields` (plain values),
 * `people` (role → `{ corporate }` or `{ list: [person] }`), `refOrder`,
 * and `extra` — the XML of every child this module does not edit, kept.
 */
export function parseSource(xml) {
  const inner = /<b:Source\b[^>]*>([\s\S]*)<\/b:Source>/.exec(String(xml));
  const body = inner ? inner[1] : String(xml);
  const source = { tag: '', type: 'Book', guid: '', fields: {}, people: {}, refOrder: null, extra: '' };
  const extra = [];
  // Top-level children only, element by element — `b:Author` holds a
  // `b:Author` of its own, so each is closed by a balanced scan.
  for (const { whole, tag, content } of topElements(body)) {
    if (tag === 'Tag') source.tag = unescXml(content);
    else if (tag === 'SourceType') source.type = unescXml(content);
    else if (tag === 'Guid') source.guid = unescXml(content);
    else if (tag === 'RefOrder') source.refOrder = Number(unescXml(content)) || null;
    else if (tag === 'Author') {
      // The outer `b:Author` holds one element per role.
      const unknown = [];
      for (const { whole: roleWhole, tag: role, content: roleBody } of topElements(content)) {
        if (!PERSON_ROLES.includes(role)) { unknown.push(roleWhole); continue; }
        const corporate = childText(roleBody, 'Corporate');
        if (corporate !== null) source.people[role] = { corporate };
        else source.people[role] = { list: readPeople(roleBody) };
      }
      if (unknown.length) source.peopleExtra = unknown.join('');
    } else if (FIELD_LABELS[tag]) source.fields[tag] = unescXml(content);
    else extra.push(whole);
  }
  source.extra = extra.join('');
  return source;
}

function peopleXml(source) {
  const parts = [];
  for (const role of PERSON_ROLES) {
    const p = source.people?.[role];
    if (!p) continue;
    if (p.corporate) {
      parts.push('<b:' + role + '><b:Corporate>' + escXml(p.corporate) + '</b:Corporate></b:' + role + '>');
      continue;
    }
    const list = (p.list || []).filter((x) => x && (x.last || x.first || x.middle));
    if (!list.length) continue;
    const persons = list.map((x) => '<b:Person>'
      + (x.last ? '<b:Last>' + escXml(x.last) + '</b:Last>' : '')
      + (x.first ? '<b:First>' + escXml(x.first) + '</b:First>' : '')
      + (x.middle ? '<b:Middle>' + escXml(x.middle) + '</b:Middle>' : '')
      + '</b:Person>').join('');
    parts.push('<b:' + role + '><b:NameList>' + persons + '</b:NameList></b:' + role + '>');
  }
  if (source.peopleExtra) parts.push(source.peopleExtra);
  return parts.length ? '<b:Author>' + parts.join('') + '</b:Author>' : '';
}

/** One source back to `<b:Source>`, in the order Word writes its children. */
export function sourceXml(source) {
  let xml = '<b:Source>'
    + '<b:Tag>' + escXml(source.tag) + '</b:Tag>'
    + '<b:SourceType>' + escXml(source.type || 'Book') + '</b:SourceType>'
    + '<b:Guid>' + escXml(source.guid || newGuid()) + '</b:Guid>'
    + peopleXml(source);
  for (const key of FIELD_ORDER) {
    const v = source.fields?.[key];
    if (v !== undefined && v !== null && String(v) !== '') xml += '<b:' + key + '>' + escXml(v) + '</b:' + key + '>';
  }
  xml += source.extra || '';
  if (source.refOrder) xml += '<b:RefOrder>' + source.refOrder + '</b:RefOrder>';
  return xml + '</b:Source>';
}

/** `<b:Sources>` read: the style it names and every source in it. */
export function parseSources(xml) {
  const text = String(xml || '');
  const open = /<b:Sources\b([^>]*)>/.exec(text);
  const attr = (name) => {
    const m = open ? new RegExp('\\b' + name + '="([^"]*)"').exec(open[1]) : null;
    return m ? unescXml(m[1]) : null;
  };
  const style = styleFromAttributes({ selected: attr('SelectedStyle'), name: attr('StyleName'), version: attr('Version') });
  const sources = [...text.matchAll(/<b:Source\b[^>]*>[\s\S]*?<\/b:Source>/g)].map((m) => parseSource(m[0]));
  return { style, sources };
}

/** The whole part, as Word writes it: the style in the root's attributes, the default namespace bound as well. */
export function sourcesXml({ style = DEFAULT_STYLE, sources = [] } = {}) {
  const s = styleById(style);
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<b:Sources SelectedStyle="' + escXml(s.selected) + '" StyleName="' + escXml(s.name) + '" Version="' + escXml(s.version) + '"'
    + ' xmlns:b="' + BIB_NS + '" xmlns="' + BIB_NS + '">'
    + sources.map(sourceXml).join('')
    + '</b:Sources>';
}

/** A Word-shaped GUID, braces and all. */
export function newGuid() {
  const uuid = globalThis.crypto?.randomUUID?.()
    || 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
      const r = (Math.random() * 16) | 0;
      return (c === 'x' ? r : (r & 0x3) | 0x8).toString(16);
    });
  return '{' + uuid.toUpperCase() + '}';
}

/**
 * Word's tag for a new source: the first three letters of the first
 * author's surname and the year's last two digits — `Smi20` — made unique
 * among `taken` by a number after it, as Word does (`Smi201`).
 */
export function makeTag(source, taken = []) {
  const who = firstName(source);
  const stem = (who || source.fields?.Title || 'Src').replace(/[^\p{L}\p{N}]/gu, '').slice(0, 3) || 'Src';
  const year = String(source.fields?.Year || '').replace(/\D/g, '').slice(-2);
  const base = stem.charAt(0).toUpperCase() + stem.slice(1) + year;
  const used = new Set((taken || []).map((t) => String(t).toLowerCase()));
  if (!used.has(base.toLowerCase())) return base;
  for (let n = 1; ; n++) if (!used.has((base + n).toLowerCase())) return base + n;
}

const firstName = (source) => {
  const a = source.people?.Author;
  if (a?.corporate) return a.corporate;
  return a?.list?.[0]?.last || '';
};

/**
 * Names typed into the Author box — "Smith, John A.; Jones, Bob" or
 * "John A. Smith; Bob Jones" — as people, the way Edit Name splits them.
 */
export function parseNames(text) {
  return String(text || '').split(/;/).map((one) => one.trim()).filter(Boolean).map((one) => {
    if (one.includes(',')) {
      const [last, rest = ''] = one.split(/,(.*)/s).map((x) => x.trim());
      const [first = '', ...middle] = rest.split(/\s+/).filter(Boolean);
      return { last, first, middle: middle.join(' ') };
    }
    const words = one.split(/\s+/).filter(Boolean);
    if (words.length === 1) return { last: words[0], first: '', middle: '' };
    return { last: words[words.length - 1], first: words[0], middle: words.slice(1, -1).join(' ') };
  });
}

/** People back to the Author box's words: "Smith, John A.; Jones, Bob". */
export function namesText(people) {
  if (!people) return '';
  if (people.corporate) return people.corporate;
  return (people.list || []).map((p) => [p.last, [p.first, p.middle].filter(Boolean).join(' ')].filter(Boolean).join(', ')).join('; ');
}

/* ── the CITATION field ──────────────────────────────────────────────────── */

/**
 * A CITATION field's code: ` CITATION Smi20 \l 2057 \p 23 \m Jon19 `.
 * `\l` the language, `\p` pages, `\n` \y` `\t` suppress author, year and
 * title, `\m` each further source cited in the same parentheses.
 */
export function citationInstr({ tags = [], lcid = 1033, pages = '', suppressAuthor = false, suppressYear = false, suppressTitle = false } = {}) {
  const [first, ...rest] = tags;
  let instr = ' CITATION ' + first + ' \\l ' + lcid + ' ';
  if (pages) instr += '\\p ' + (/\s/.test(pages) ? '"' + pages + '"' : pages) + ' ';
  if (suppressAuthor) instr += '\\n ';
  if (suppressYear) instr += '\\y ';
  if (suppressTitle) instr += '\\t ';
  for (const t of rest) instr += '\\m ' + t + ' ';
  return instr;
}

/** A CITATION field's code, read back. */
export function parseCitationInstr(instr) {
  const words = [];
  const re = /"([^"]*)"|(\S+)/g;
  let m;
  while ((m = re.exec(String(instr)))) words.push(m[1] !== undefined ? m[1] : m[2]);
  if (!words.length || words[0].toUpperCase() !== 'CITATION') return null;
  const out = { tags: [], lcid: 1033, pages: '', suppressAuthor: false, suppressYear: false, suppressTitle: false };
  if (words[1] && !words[1].startsWith('\\')) out.tags.push(words[1]);
  for (let i = 2; i < words.length; i++) {
    const w = words[i].toLowerCase();
    if (w === '\\l') out.lcid = Number(words[++i]) || 1033;
    else if (w === '\\p') out.pages = words[++i] || '';
    else if (w === '\\n') out.suppressAuthor = true;
    else if (w === '\\y') out.suppressYear = true;
    else if (w === '\\t') out.suppressTitle = true;
    else if (w === '\\m') { if (words[i + 1]) out.tags.push(words[++i]); }
  }
  return out;
}

/* ── formatting ──────────────────────────────────────────────────────────── */

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const MLA_MONTHS = ['Jan.', 'Feb.', 'Mar.', 'Apr.', 'May', 'June', 'July', 'Aug.', 'Sept.', 'Oct.', 'Nov.', 'Dec.'];
const monthIndex = (m) => {
  if (!m) return -1;
  const n = Number(m);
  if (Number.isFinite(n) && n >= 1 && n <= 12) return n - 1;
  const s = String(m).trim().toLowerCase().slice(0, 3);
  return MONTHS.findIndex((x) => x.toLowerCase().startsWith(s));
};
const monthName = (m) => { const i = monthIndex(m); return i >= 0 ? MONTHS[i] : String(m || ''); };
const mlaMonth = (m) => { const i = monthIndex(m); return i >= 0 ? MLA_MONTHS[i] : String(m || ''); };

/** Words end with one full stop, never two, and not after ? or !. */
const stop = (s) => { const t = String(s || '').trim(); return !t ? '' : /[.?!]$/.test(t) ? t : t + '.'; };
const initialsOf = (names, { spaced = true } = {}) => String(names || '').split(/\s+/).filter(Boolean)
  .map((w) => w.split('-').map((p) => (p ? p.charAt(0).toUpperCase() + '.' : '')).join('-'))
  .join(spaced ? ' ' : '');
const given = (p) => [p.first, p.middle].filter(Boolean).join(' ');
const range = (pages) => String(pages || '').replace(/\s*-+\s*/g, '–');
const isRange = (pages) => /[–,-]/.test(String(pages || ''));

/** Join names as prose: "A and B", "A, B, and C" (serial comma when `serial`), with `word` between the last two. */
function joinNames(names, word, serial) {
  if (names.length <= 1) return names.join('');
  if (names.length === 2) return names[0] + (serial === 'always' ? ', ' : ' ') + word + ' ' + names[1];
  return names.slice(0, -1).join(', ') + (serial ? ', ' : ' ') + word + ' ' + names[names.length - 1];
}

/** The people who stand in the author's place: authors, else editors, else a corporate name. */
function leadPeople(source) {
  const a = source.people?.Author;
  if (a?.corporate) return { corporate: a.corporate, list: [] };
  if (a?.list?.length) return { list: a.list };
  const e = source.people?.Editor;
  if (e?.corporate) return { corporate: e.corporate, list: [], editors: true };
  if (e?.list?.length) return { list: e.list, editors: true };
  return null;
}

const f = (source, key) => String(source.fields?.[key] ?? '').trim();
const seg = (text, italic = false) => ({ text, ...(italic ? { italic: true } : {}) });

/** Segments cleaned up: empty ones dropped, neighbours of the same look joined. */
function tidy(segs) {
  const out = [];
  for (const s of segs.flat().filter((x) => x && x.text)) {
    const last = out[out.length - 1];
    if (last && Boolean(last.italic) === Boolean(s.italic)) last.text += s.text;
    else out.push({ ...s });
  }
  if (out.length) out[out.length - 1].text = out[out.length - 1].text.replace(/\s+$/, '');
  return out;
}

export const segmentsText = (segs) => (segs || []).map((s) => s.text).join('');

/* In-text surnames, per style. */
function citeNames(source, styleId) {
  const lead = leadPeople(source);
  if (!lead) return null;
  if (lead.corporate) return lead.corporate;
  const last = lead.list.map((p) => p.last || p.first);
  const n = last.length;
  switch (styleId) {
    case 'apa7':
      return n === 1 ? last[0] : n === 2 ? last[0] + ' & ' + last[1] : last[0] + ' et al.';
    case 'apa6':
      return n === 1 ? last[0] : n === 2 ? last[0] + ' & ' + last[1] : n <= 5 ? last.slice(0, -1).join(', ') + ', & ' + last[n - 1] : last[0] + ' et al.';
    case 'mla7':
      return n === 1 ? last[0] : n === 2 ? last[0] + ' and ' + last[1] : n === 3 ? last[0] + ', ' + last[1] + ', and ' + last[2] : last[0] + ' et al.';
    case 'chicago16':
      return n === 1 ? last[0] : n === 2 ? last[0] + ' and ' + last[1] : n === 3 ? last[0] + ', ' + last[1] + ', and ' + last[2] : last[0] + ' et al.';
    case 'harvard':
      return n === 1 ? last[0] : n === 2 ? last[0] + ' and ' + last[1] : n === 3 ? last[0] + ', ' + last[1] + ' and ' + last[2] : last[0] + ' et al.';
    default:
      return last[0];
  }
}

/** A title as a citation shows it when there is no author: italic for a whole work, quoted for a part of one. */
function citeTitle(source, styleId) {
  const title = f(source, 'ShortTitle') || f(source, 'Title');
  if (!title) return [];
  const part = ['JournalArticle', 'BookSection', 'InternetSite', 'ConferenceProceedings'].includes(source.type);
  if (styleId === 'mla7' || styleId === 'chicago16') return part ? [seg('“' + title + '”')] : [seg(title, true)];
  return part ? [seg('“' + title + '”')] : [seg(title, true)];
}

const yearOf = (source) => f(source, 'Year') || 'n.d.';

/**
 * One source's part of an in-text citation — without the parentheses, which
 * hold every source the citation names — in `styleId`. `number` is the
 * source's number for a numeric style.
 */
function citePart(source, styleId, { pages = '', suppressAuthor = false, suppressYear = false, suppressTitle = false, number = null } = {}) {
  if (styleId === 'ieee') {
    const n = number ?? '?';
    return [seg(pages ? n + ', p. ' + range(pages) : String(n))];
  }
  const who = suppressAuthor ? null : citeNames(source, styleId);
  const title = !who && !suppressTitle && !suppressAuthor ? citeTitle(source, styleId) : [];
  const head = who ? [seg(who)] : title;
  const year = suppressYear ? '' : yearOf(source);
  const p = pages ? range(pages) : '';
  switch (styleId) {
    case 'mla7':
      // MLA cites no year: the name and the page, with no comma between.
      return tidy([head, p ? seg((head.length ? ' ' : '') + p) : null]);
    case 'chicago16':
      return tidy([head, year ? seg((head.length ? ' ' : '') + year) : null, p ? seg((head.length || year ? ', ' : '') + p) : null]);
    case 'harvard':
      return tidy([head, year ? seg((head.length ? ', ' : '') + year) : null, p ? seg((head.length || year ? ', ' : '') + (isRange(p) ? 'pp.' : 'p.') + p) : null]);
    default: // APA
      return tidy([head, year ? seg((head.length ? ', ' : '') + year) : null, p ? seg((head.length || year ? ', ' : '') + (isRange(p) ? 'pp. ' : 'p. ') + p) : null]);
  }
}

/**
 * An in-text citation, in `styleId`: the sources it names (a placeholder
 * — a tag with no source yet — is shown by its tag, as Word shows one), the
 * pages and the suppressions; `numbers` maps a tag to its number for a
 * numeric style. Segments, parentheses (or brackets) included.
 */
export function formatCitation(entries, styleId = DEFAULT_STYLE, opts = {}) {
  const list = (entries || []).filter(Boolean);
  if (!list.length) return [];
  const parts = list.map((src, i) => {
    if (src.placeholder) return [seg(src.tag)];
    return citePart(src, styleId, { ...(i === 0 ? opts : {}), number: opts.numbers?.get?.(src.tag) ?? opts.numbers?.[src.tag] ?? null });
  });
  if (styleId === 'ieee') return tidy([seg('['), ...parts.flatMap((p, i) => (i ? [seg(', '), ...p] : p)), seg(']')]);
  return tidy([seg('('), ...parts.flatMap((p, i) => (i ? [seg('; '), ...p] : p)), seg(')')]);
}

/* Reference-list names, per style. */
function listNames(source, styleId) {
  const lead = leadPeople(source);
  if (!lead) return null;
  const eds = lead.editors;
  if (lead.corporate) return lead.corporate + (eds ? (styleId.startsWith('apa') ? ' (Ed.)' : ', ed.') : '');
  const ps = lead.list;
  const n = ps.length;
  let text;
  switch (styleId) {
    case 'apa7':
    case 'apa6': {
      const one = (p) => (p.last ? p.last + (given(p) ? ', ' + initialsOf(given(p)) : '') : initialsOf(given(p)));
      const cap = styleId === 'apa7' ? 20 : 7;
      let names = ps.map(one);
      if (n > cap) names = [...names.slice(0, cap - 1), '. . . ' + names[n - 1]];
      text = n === 1 ? names[0] : n > cap ? names.slice(0, -1).join(', ') + ', ' + names[names.length - 1] : names.slice(0, -1).join(', ') + ', & ' + names[names.length - 1];
      if (eds) text += n > 1 ? ' (Eds.)' : ' (Ed.)';
      return text;
    }
    case 'mla7':
    case 'chicago16': {
      const first = ps[0].last + (given(ps[0]) ? ', ' + given(ps[0]) : '');
      const others = ps.slice(1).map((p) => [given(p), p.last].filter(Boolean).join(' '));
      if (styleId === 'mla7' && n >= 4) text = first + ', et al';
      else if (styleId === 'chicago16' && n > 10) text = [first, ...others.slice(0, 6)].join(', ') + ', et al';
      else text = n === 1 ? first : joinNames([first, ...others], 'and', 'always');
      if (eds) text += n > 1 ? ', eds' : ', ed';
      return text;
    }
    case 'harvard': {
      const one = (p) => String(p.last || '').toUpperCase() + (given(p) ? ', ' + initialsOf(given(p), { spaced: false }) : '');
      text = n === 1 ? one(ps[0]) : n <= 3 ? joinNames(ps.map(one), 'and', false) : one(ps[0]) + ' et al.';
      if (eds) text += ' ed.';
      return text;
    }
    case 'ieee': {
      const one = (p) => [initialsOf(given(p)), p.last].filter(Boolean).join(' ');
      text = n <= 2 ? ps.map(one).join(' and ') : n <= 6 ? ps.slice(0, -1).map(one).join(', ') + ' and ' + one(ps[n - 1]) : one(ps[0]) + ' et al.';
      if (eds) text += n > 1 ? ', Eds.' : ', Ed.';
      return text;
    }
    default:
      return ps.map((p) => p.last).join(', ');
  }
}

/** Editors as "In J. Smith (Ed.)," reads them — given names first. */
function editorsInline(source, styleId) {
  const e = source.people?.Editor;
  if (!e) return '';
  if (e.corporate) return e.corporate;
  const ps = e.list || [];
  if (styleId.startsWith('apa') || styleId === 'harvard') {
    const names = ps.map((p) => [initialsOf(given(p)), styleId === 'harvard' ? String(p.last || '').toUpperCase() : p.last].filter(Boolean).join(' '));
    return styleId === 'harvard' ? joinNames(names, 'and', false) : names.length === 2 ? names.join(' & ') : names.length > 2 ? names.slice(0, -1).join(', ') + ', & ' + names[names.length - 1] : names[0] || '';
  }
  return joinNames(ps.map((p) => [given(p), p.last].filter(Boolean).join(' ')), 'and', true);
}

const place = (source, sep = ': ') => [f(source, 'City'), f(source, 'Publisher')].filter(Boolean).join(sep);
const doiOrUrl = (source) => {
  const doi = f(source, 'DOI');
  if (doi) return /^https?:/i.test(doi) ? doi : 'https://doi.org/' + doi.replace(/^doi:\s*/i, '');
  return f(source, 'URL');
};
const dateLong = (y, m, d) => [monthName(m) && m ? monthName(m) : '', d, ].filter(Boolean).join(' ') + (y ? (m || d ? ', ' : '') + y : '');
const dayMonthYear = (y, m, d, months = MONTHS) => [d, m ? months[monthIndex(m)] || m : '', y].filter(Boolean).join(' ');

function apaEntry(source, styleId) {
  const t = source.type;
  const who = listNames(source, styleId);
  const title = f(source, 'Title');
  const y = f(source, 'Year');
  const m = f(source, 'Month');
  const d = f(source, 'Day');
  const date = '(' + (y ? [y, m ? monthName(m) + (d ? ' ' + d : '') : ''].filter(Boolean).join(', ') : 'n.d.') + ').';
  const whole = ['Book', 'Report'].includes(t);
  const ed = f(source, 'Edition');
  const edition = ed ? ' (' + (/^\d+$/.test(ed) ? ed + ordinal(ed) : ed.replace(/\s*ed(ition)?\.?$/i, '')) + ' ed.)' : '';
  // A whole work's title is italic (a web page's too, in APA); a part of
  // one — an article, a chapter, a paper — is plain.
  const titleSegs = title ? (whole || t === 'InternetSite' ? [seg(title, true), seg(edition + '. ')] : [seg(stop(title) + ' ')]) : [];
  const head = who
    ? [seg(stop(who) + ' '), seg(date + ' '), ...titleSegs]
    : [...titleSegs, seg(date + ' ')];
  const pub = styleId === 'apa6' ? place(source) : f(source, 'Publisher');
  const tail = [];
  switch (t) {
    case 'JournalArticle': {
      const vol = f(source, 'Volume');
      const iss = f(source, 'Issue');
      const pages = range(f(source, 'Pages'));
      tail.push(seg(f(source, 'JournalName'), true));
      if (vol) tail.push(seg(', '), seg(vol, true));
      if (iss) tail.push(seg('(' + iss + ')'));
      if (pages) tail.push(seg(', ' + pages));
      tail.push(seg('. '));
      break;
    }
    case 'BookSection': {
      const eds = editorsInline(source, styleId);
      const ne = (source.people?.Editor?.list || []).length;
      tail.push(seg('In ' + (eds ? eds + (ne > 1 ? ' (Eds.), ' : ' (Ed.), ') : '')), seg(f(source, 'BookTitle'), true));
      const pages = range(f(source, 'Pages'));
      if (pages) tail.push(seg(' (' + (isRange(pages) ? 'pp. ' : 'p. ') + pages + ')'));
      tail.push(seg('. '));
      if (pub) tail.push(seg(stop(pub) + ' '));
      break;
    }
    case 'ConferenceProceedings': {
      const pages = range(f(source, 'Pages'));
      if (f(source, 'ConferenceName')) tail.push(seg('In '), seg(f(source, 'ConferenceName'), true), seg(pages ? ' (' + (isRange(pages) ? 'pp. ' : 'p. ') + pages + '). ' : '. '));
      if (pub) tail.push(seg(stop(pub) + ' '));
      break;
    }
    case 'InternetSite': {
      if (f(source, 'InternetSiteTitle')) tail.push(seg(stop(f(source, 'InternetSiteTitle')) + ' '));
      const accessed = f(source, 'YearAccessed') ? dateLong(f(source, 'YearAccessed'), f(source, 'MonthAccessed'), f(source, 'DayAccessed')) : '';
      const url = f(source, 'URL');
      if (accessed && url) tail.push(seg('Retrieved ' + accessed + ', from ' + url));
      else if (url) tail.push(seg((styleId === 'apa6' ? 'Retrieved from ' : '') + url));
      return tidy([head, tail]);
    }
    default:
      if (pub) tail.push(seg(stop(pub) + ' '));
  }
  const link = t === 'InternetSite' ? '' : doiOrUrl(source);
  if (link) tail.push(seg(link));
  return tidy([head, tail]);
}

const ordinal = (n) => { const v = Number(n) % 100; return v >= 11 && v <= 13 ? 'th' : ['th', 'st', 'nd', 'rd'][Number(n) % 10] || 'th'; };

function mlaEntry(source) {
  const t = source.type;
  const who = listNames(source, 'mla7');
  const title = f(source, 'Title');
  const y = f(source, 'Year');
  const out = [];
  if (who) out.push(seg(stop(who) + ' '));
  const whole = ['Book', 'Report'].includes(t);
  if (title) out.push(whole ? seg(title, true) : seg('“' + stop(title) + '” '), whole ? seg('. ') : null);
  const pubLine = () => [place(source), y].filter(Boolean).join(', ');
  switch (t) {
    case 'JournalArticle': {
      const vi = [f(source, 'Volume'), f(source, 'Issue')].filter(Boolean).join('.');
      out.push(seg(f(source, 'JournalName'), true), seg((vi ? ' ' + vi : '') + ' (' + (y || 'n.d.') + ')' + (f(source, 'Pages') ? ': ' + range(f(source, 'Pages')) : '') + '. '));
      break;
    }
    case 'BookSection':
      out.push(seg(f(source, 'BookTitle'), true), seg('. '));
      if (source.people?.Editor) out.push(seg('Ed. ' + editorsInline(source, 'mla7') + '. '));
      if (pubLine()) out.push(seg(stop(pubLine()) + ' '));
      if (f(source, 'Pages')) out.push(seg(stop(range(f(source, 'Pages'))) + ' '));
      break;
    case 'ConferenceProceedings':
      if (f(source, 'ConferenceName')) out.push(seg(f(source, 'ConferenceName'), true), seg('. '));
      if (pubLine()) out.push(seg(stop(pubLine()) + ' '));
      if (f(source, 'Pages')) out.push(seg(stop(range(f(source, 'Pages'))) + ' '));
      break;
    case 'InternetSite': {
      if (f(source, 'InternetSiteTitle')) out.push(seg(f(source, 'InternetSiteTitle'), true), seg('. '));
      if (f(source, 'ProductionCompany')) out.push(seg(f(source, 'ProductionCompany') + ', '));
      out.push(seg(stop(dayMonthYear(y, f(source, 'Month'), f(source, 'Day'), MLA_MONTHS) || 'n.d.') + ' '), seg('Web. '));
      const acc = dayMonthYear(f(source, 'YearAccessed'), f(source, 'MonthAccessed'), f(source, 'DayAccessed'), MLA_MONTHS);
      if (acc) out.push(seg(stop(acc) + ' '));
      return tidy(out);
    }
    default:
      if (pubLine()) out.push(seg(stop(pubLine()) + ' '));
  }
  out.push(seg(stop(f(source, 'Medium') || 'Print')));
  return tidy(out);
}

function chicagoEntry(source) {
  const t = source.type;
  const who = listNames(source, 'chicago16');
  const title = f(source, 'Title');
  const y = f(source, 'Year');
  const out = [];
  if (who) out.push(seg(stop(who) + ' '));
  const whole = ['Book', 'Report'].includes(t);
  if (title) out.push(whole ? seg(title, true) : seg('“' + stop(title) + '” '), whole ? seg('. ') : null);
  const pubLine = [place(source), y].filter(Boolean).join(', ');
  switch (t) {
    case 'JournalArticle': {
      const iss = f(source, 'Issue');
      out.push(seg(f(source, 'JournalName'), true), seg((f(source, 'Volume') ? ' ' + f(source, 'Volume') : '') + (iss ? ', no. ' + iss : '') + ' (' + (y || 'n.d.') + ')' + (f(source, 'Pages') ? ': ' + range(f(source, 'Pages')) : '') + '. '));
      break;
    }
    case 'BookSection':
      out.push(seg('In '), seg(f(source, 'BookTitle'), true));
      if (source.people?.Editor) out.push(seg(', edited by ' + editorsInline(source, 'chicago16')));
      if (f(source, 'Pages')) out.push(seg(', ' + range(f(source, 'Pages'))));
      out.push(seg('. '));
      if (pubLine) out.push(seg(stop(pubLine) + ' '));
      break;
    case 'ConferenceProceedings':
      if (f(source, 'ConferenceName')) out.push(seg(f(source, 'ConferenceName'), true), seg('. '));
      if (pubLine) out.push(seg(stop(pubLine) + ' '));
      if (f(source, 'Pages')) out.push(seg(stop(range(f(source, 'Pages'))) + ' '));
      break;
    case 'InternetSite': {
      if (f(source, 'InternetSiteTitle')) out.push(seg(f(source, 'InternetSiteTitle'), true), seg('. '));
      const date = dateLong(y, f(source, 'Month'), f(source, 'Day'));
      if (date) out.push(seg(stop(date) + ' '));
      const acc = f(source, 'YearAccessed') ? dateLong(f(source, 'YearAccessed'), f(source, 'MonthAccessed'), f(source, 'DayAccessed')) : '';
      if (acc) out.push(seg('Accessed ' + stop(acc) + ' '));
      if (f(source, 'URL')) out.push(seg(stop(f(source, 'URL'))));
      return tidy(out);
    }
    default:
      if (pubLine) out.push(seg(stop(pubLine) + ' '));
  }
  const link = doiOrUrl(source);
  if (link) out.push(seg(stop(link)));
  return tidy(out);
}

function harvardEntry(source) {
  const t = source.type;
  const who = listNames(source, 'harvard');
  const title = f(source, 'Title');
  const y = f(source, 'Year') || 'n.d.';
  const out = [];
  const whole = ['Book', 'Report', 'InternetSite'].includes(t);
  if (who) out.push(seg(who + ', ' + y + '. '));
  if (title) out.push(whole ? seg(title, true) : seg(stop(title) + ' '), whole ? seg('. ') : null);
  if (!who) out.push(seg(y + '. '));
  const pages = range(f(source, 'Pages'));
  switch (t) {
    case 'JournalArticle':
      out.push(seg(f(source, 'JournalName'), true), seg((f(source, 'Volume') ? ', ' + f(source, 'Volume') : '') + (f(source, 'Issue') ? '(' + f(source, 'Issue') + ')' : '') + (pages ? ', ' + (isRange(pages) ? 'pp.' : 'p.') + pages : '') + '. '));
      break;
    case 'BookSection': {
      const eds = editorsInline(source, 'harvard');
      out.push(seg('In: ' + (eds ? eds + ', ed. ' : '')), seg(f(source, 'BookTitle'), true), seg('. '));
      if (place(source)) out.push(seg(stop(place(source)) + ' '));
      if (pages) out.push(seg((isRange(pages) ? 'pp.' : 'p.') + pages + '.'));
      break;
    }
    case 'ConferenceProceedings':
      out.push(seg('In: '), seg(f(source, 'ConferenceName'), true), seg('. '));
      if (place(source)) out.push(seg(stop(place(source)) + ' '));
      if (pages) out.push(seg((isRange(pages) ? 'pp.' : 'p.') + pages + '.'));
      break;
    case 'InternetSite': {
      out.push(seg('[Online] '));
      if (f(source, 'URL')) out.push(seg('Available at: ' + f(source, 'URL') + ' '));
      const acc = dayMonthYear(f(source, 'YearAccessed'), f(source, 'MonthAccessed'), f(source, 'DayAccessed'));
      if (acc) out.push(seg('[Accessed ' + acc + '].'));
      return tidy(out);
    }
    default:
      if (place(source)) out.push(seg(stop(place(source))));
  }
  return tidy(out);
}

function ieeeEntry(source, number) {
  const t = source.type;
  const who = listNames(source, 'ieee');
  const title = f(source, 'Title');
  const y = f(source, 'Year');
  const out = [seg('[' + (number ?? '?') + ']\t')];
  if (who) out.push(seg(who + ', '));
  const whole = ['Book', 'Report'].includes(t);
  if (title) out.push(whole ? seg(title, true) : seg('“' + title + ',” '), whole ? seg(', ') : null);
  const pages = range(f(source, 'Pages'));
  switch (t) {
    case 'JournalArticle':
      out.push(seg(f(source, 'JournalName'), true), seg((f(source, 'Volume') ? ', vol. ' + f(source, 'Volume') : '') + (f(source, 'Issue') ? ', no. ' + f(source, 'Issue') : '') + (pages ? ', pp. ' + pages : '') + (y ? ', ' + y : '') + '.'));
      break;
    case 'BookSection':
      out.push(seg('in '), seg(f(source, 'BookTitle'), true), seg(', ' + [place(source), y].filter(Boolean).join(', ') + (pages ? ', pp. ' + pages : '') + '.'));
      break;
    case 'ConferenceProceedings':
      out.push(seg('in '), seg(f(source, 'ConferenceName'), true), seg(', ' + [f(source, 'City'), y].filter(Boolean).join(', ') + (pages ? ', pp. ' + pages : '') + '.'));
      break;
    case 'InternetSite': {
      if (f(source, 'InternetSiteTitle')) out.push(seg(f(source, 'InternetSiteTitle') + ', '));
      out.push(seg((y || 'n.d.') + '. [Online]. '));
      if (f(source, 'URL')) out.push(seg('Available: ' + f(source, 'URL') + '. '));
      const acc = dayMonthYear(f(source, 'YearAccessed'), f(source, 'MonthAccessed'), f(source, 'DayAccessed'));
      if (acc) out.push(seg('[Accessed ' + acc + '].'));
      return tidy(out);
    }
    default:
      out.push(seg([place(source), y].filter(Boolean).join(', ') + '.'));
  }
  return tidy(out);
}

/** One bibliography entry, in `styleId`, as segments. `number` for a numeric style. */
export function formatBibliographyEntry(source, styleId = DEFAULT_STYLE, { number = null } = {}) {
  switch (styleId) {
    case 'apa7':
    case 'apa6': return apaEntry(source, styleId);
    case 'mla7': return mlaEntry(source);
    case 'chicago16': return chicagoEntry(source);
    case 'harvard': return harvardEntry(source);
    case 'ieee': return ieeeEntry(source, number);
    default: return apaEntry(source, 'apa7');
  }
}

/** What a source sorts by in an alphabetical list: its lead name, else its title without A/An/The; then the year. */
function sortKey(source) {
  const lead = leadPeople(source);
  const name = lead ? (lead.corporate || lead.list.map((p) => [p.last, p.first, p.middle].join(' ')).join(' ')) : '';
  const title = f(source, 'Title').replace(/^(the|a|an)\s+/i, '');
  return [(name || title).toLowerCase(), f(source, 'Year'), title.toLowerCase()];
}

/**
 * The numbers a numeric style gives: each source numbered by the order of
 * its first citation (`citedTags`, in document order), then the sources
 * never cited in the order of the list.
 */
export function citationNumbers(sources, citedTags = []) {
  const numbers = new Map();
  let n = 1;
  for (const tag of citedTags) if (!numbers.has(tag) && sources.some((s) => s.tag === tag)) numbers.set(tag, n++);
  for (const s of sources) if (!numbers.has(s.tag)) numbers.set(s.tag, n++);
  return numbers;
}

/**
 * The bibliography: every source in the document's list — Word lists them
 * all, cited or not — in the style's order, each as segments.
 */
export function formatBibliography(sources, styleId = DEFAULT_STYLE, { citedTags = [] } = {}) {
  const list = [...(sources || [])];
  if (styleId === 'ieee') {
    const numbers = citationNumbers(list, citedTags);
    return list
      .sort((a, b) => numbers.get(a.tag) - numbers.get(b.tag))
      .map((s) => ({ tag: s.tag, segments: formatBibliographyEntry(s, styleId, { number: numbers.get(s.tag) }) }));
  }
  return list
    .sort((a, b) => {
      const ka = sortKey(a);
      const kb = sortKey(b);
      for (let i = 0; i < ka.length; i++) {
        const c = ka[i].localeCompare(kb[i], undefined, { sensitivity: 'base' });
        if (c) return c;
      }
      return 0;
    })
    .map((s) => ({ tag: s.tag, segments: formatBibliographyEntry(s, styleId) }));
}

/**
 * How Manage Sources and the Insert Citation menu describe a source: the
 * people on one line, the title and year on the next.
 */
export function describeSource(source) {
  const lead = leadPeople(source);
  const who = lead ? (lead.corporate || lead.list.map((p) => [p.last, given(p)].filter(Boolean).join(', ')).join('; ')) : '';
  const title = f(source, 'Title');
  const y = f(source, 'Year');
  return { who, title, year: y, line: [who, title, y ? '(' + y + ')' : ''].filter(Boolean).join('; ') };
}
