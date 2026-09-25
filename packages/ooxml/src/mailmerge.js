/**
 * Mail merge — the rules, with no file in sight.
 *
 * Everything here is plain data in and plain data out: the field codes Word
 * writes (MERGEFIELD, ADDRESSBLOCK, GREETINGLINE, IF, NEXT, SKIPIF…), how
 * each one reads for a record, the recipient list a data source becomes, and
 * the order a merge walks it in. No Node module is imported, so the window
 * can draw the Address Block dialog's preview from the same rules the merge
 * itself runs on, and the two can never disagree about what "Mr. Joshua
 * Randall Jr." means.
 *
 * The XML half — putting these fields into a paragraph, and a merged copy of
 * the document per record — is `mailmerge-run.js`, which does know files.
 */

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** Start Mail Merge's six choices, in Word's own order, with the value `w:mainDocumentType` takes. */
export const MAIN_DOCUMENT_TYPES = [
  { id: 'formLetters', label: 'Letters' },
  { id: 'email', label: 'E-mail Messages' },
  { id: 'envelopes', label: 'Envelopes' },
  { id: 'mailingLabels', label: 'Labels' },
  { id: 'catalog', label: 'Directory' },
];

/** The field kinds a merge evaluates; every other field (REF, PAGE, TOC) is left as it is. */
export const MERGE_KINDS = new Set(['mergefield', 'addressblock', 'greetingline', 'if', 'next', 'nextif', 'skipif', 'mergerec', 'mergeseq']);

/** What each kind shows in the main document, between Word's chevrons, until a record is previewed. */
export const PLACEHOLDERS = {
  addressblock: '«AddressBlock»',
  greetingline: '«GreetingLine»',
  next: '«Next Record»',
  nextif: '«Next Record If...»',
  skipif: '«Skip Record If...»',
  mergerec: '«Merge Record #»',
  mergeseq: '«Merge Sequence #»',
};

/**
 * Type a New List's columns — Word's own, in its own order, so a list typed
 * here and one typed in Word carry the same names.
 */
export const NEW_LIST_FIELDS = [
  'Title', 'First Name', 'Last Name', 'Company Name', 'Address Line 1', 'Address Line 2',
  'City', 'State', 'ZIP Code', 'Country or Region', 'Home Phone', 'Work Phone', 'E-mail Address',
];

/* ── field names ───────────────────────────────────────────────────────── */

/**
 * A column as a MERGEFIELD names it. Word will not put a space in a field
 * name — "First Name" is written `MERGEFIELD First_Name` and shows as
 * «First_Name» — and matches the two back by the same substitution.
 */
export function mergeFieldName(header) {
  return String(header ?? '').trim().replace(/\s+/g, '_');
}

const fold = (s) => String(s ?? '').trim().toLowerCase().replace(/[\s_]+/g, ' ');

/** The value of a named field in a record: a column matched by name, spaces and underscores alike, case aside. */
export function fieldValue(source, record, name) {
  if (!source || !record) return '';
  const want = fold(name);
  const i = source.fields.findIndex((f) => fold(f) === want);
  return i < 0 ? '' : String(record[i] ?? '');
}

/** Does the source have a column by this name? */
export function hasField(source, name) {
  const want = fold(name);
  return Boolean(source?.fields?.some((f) => fold(f) === want));
}

/* ── data sources ──────────────────────────────────────────────────────── */

/**
 * Delimited text as rows. The byte-order mark Excel writes at the head of a
 * UTF-8 CSV goes (or the first column would be named "\uFEFFTitle" and never
 * match), a quoted field keeps its commas, doubled quotes and line breaks,
 * and a CRLF file reads the same as an LF one. `delimiter` defaults to the
 * one the first line uses most: a comma, a semicolon (Excel in much of
 * Europe) or a tab.
 */
export function parseDelimited(text, delimiter = null) {
  const s = String(text ?? '').replace(/^\uFEFF/, '');
  const d = delimiter || guessDelimiter(s);
  const rows = [];
  let row = [];
  let field = '';
  let quoted = false;
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (quoted) {
      if (c === '"') {
        if (s[i + 1] === '"') { field += '"'; i++; } else quoted = false;
      } else field += c;
      continue;
    }
    if (c === '"' && field === '') { quoted = true; continue; }
    if (c === d) { row.push(field); field = ''; continue; }
    if (c === '\r') continue;
    if (c === '\n') { row.push(field); rows.push(row); row = []; field = ''; continue; }
    field += c;
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  return rows.filter((r) => r.some((v) => String(v).trim() !== ''));
}

function guessDelimiter(text) {
  const first = text.split(/\r?\n/, 1)[0] || '';
  let best = ',';
  let count = -1;
  for (const d of [',', ';', '\t', '|']) {
    let n = 0;
    let q = false;
    for (const c of first) {
      if (c === '"') q = !q;
      else if (c === d && !q) n++;
    }
    if (n > count) { count = n; best = d; }
  }
  return best;
}

/**
 * Rows, the first of them the field names, as a recipient list: every name
 * made unique and non-empty (Word calls a nameless column by its position),
 * every record as long as the header, blank rows dropped.
 */
export function sourceFromRows(rows, meta = {}) {
  const list = (rows || []).filter((r) => Array.isArray(r) && r.some((v) => String(v ?? '').trim() !== ''));
  const head = (list[0] || []).map((h, i) => String(h ?? '').trim() || `Column ${i + 1}`);
  const seen = new Map();
  const fields = head.map((h) => {
    const k = h.toLowerCase();
    const n = (seen.get(k) || 0) + 1;
    seen.set(k, n);
    return n > 1 ? `${h} ${n}` : h;
  });
  const records = list.slice(1).map((r) => fields.map((_, i) => (r[i] == null ? '' : String(r[i]))));
  return { ...meta, fields, records };
}

/**
 * The suite's own address book as a recipient list, in Type a New List's
 * columns — so an Address Block matches it with no Match Fields at all.
 * One record per card; a card's first address is the one used.
 */
export function contactsToSource(contacts, meta = {}) {
  const rows = [NEW_LIST_FIELDS];
  for (const c of contacts || []) {
    const n = c.name || {};
    const a = (c.addresses || []).find((x) => x.pref) || (c.addresses || [])[0] || {};
    const street = String(a.street || '').split(/\r?\n/);
    const phone = (kind) => ((c.phones || []).find((p) => p.type === kind) || {}).value || '';
    const email = ((c.emails || []).find((e) => e.pref) || (c.emails || [])[0] || {}).value || '';
    const given = n.given || (!n.family && n.full ? String(n.full).split(/\s+/)[0] : '');
    const family = n.family || (!n.given && n.full ? String(n.full).split(/\s+/).slice(1).join(' ') : '');
    rows.push([
      n.prefix || '', given, family, c.org || '', street[0] || '', street.slice(1).join(', '),
      a.city || '', a.region || '', a.postcode || '', a.country || '',
      phone('home'), phone('work') || phone('cell') || ((c.phones || [])[0] || {}).value || '', email,
    ]);
  }
  return sourceFromRows(rows, { kind: 'contacts', name: 'Contacts', ...meta });
}

/**
 * The records a merge walks, in the order it walks them: the ones left
 * ticked in Edit Recipient List, sorted by the column the list was sorted
 * on (numbers as numbers), each with its index in the source.
 */
export function mergeOrder(source, { excluded = [], sort = null } = {}) {
  if (!source) return [];
  const off = new Set(excluded);
  const out = [];
  source.records.forEach((r, i) => { if (!off.has(i)) out.push(i); });
  const col = sort ? source.fields.findIndex((f) => fold(f) === fold(sort.field)) : -1;
  if (col >= 0) {
    const dir = sort.descending ? -1 : 1;
    out.sort((a, b) => dir * compareCells(source.records[a][col], source.records[b][col]) || a - b);
  }
  return out;
}

function compareCells(a, b) {
  const x = String(a ?? '');
  const y = String(b ?? '');
  const nx = Number(x);
  const ny = Number(y);
  if (x.trim() !== '' && y.trim() !== '' && Number.isFinite(nx) && Number.isFinite(ny)) return nx - ny;
  return x.localeCompare(y, undefined, { sensitivity: 'base', numeric: true });
}

/**
 * Records that say the same thing — every field alike, case and spacing
 * aside — grouped, for Edit Recipient List's duplicates hint. Blank records
 * are not duplicates of each other.
 */
export function findDuplicates(source) {
  const groups = new Map();
  (source?.records || []).forEach((r, i) => {
    const key = r.map((v) => fold(v)).join('\u0001');
    if (!key.replace(/\u0001/g, '')) return;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(i);
  });
  return [...groups.values()].filter((g) => g.length > 1);
}

/**
 * The records a Finish & Merge covers: every one in merge order, the one
 * being previewed (`current`, 1-based within the order), or From…To.
 */
export function rangeOf(order, range = 'all', current = 1) {
  if (range === 'current') {
    const k = Math.max(1, Math.min(order.length, Number(current) || 1));
    return order.slice(k - 1, k);
  }
  if (range && typeof range === 'object') {
    const from = Math.max(1, Number(range.from) || 1);
    const to = Math.min(order.length, Number(range.to) || order.length);
    return order.slice(from - 1, Math.max(from - 1, to));
  }
  return order.slice();
}

/* ── address fields: Match Fields ──────────────────────────────────────── */

/**
 * Word's address fields — what Address Block and Greeting Line are written
 * in — each with the internal token its format string uses and the column
 * names that match it without asking (Match Fields' own guesses).
 */
export const ADDRESS_FIELDS = [
  { token: '_TITLE0_', label: 'Courtesy Title', match: ['title', 'courtesy title', 'mr mrs', 'salutation', 'prefix', 'honorific'] },
  { token: '_FIRST0_', label: 'First Name', match: ['first name', 'firstname', 'first', 'given name', 'forename', 'fname'] },
  { token: '_MIDDLE0_', label: 'Middle Name', match: ['middle name', 'middle', 'middle initial'] },
  { token: '_LAST0_', label: 'Last Name', match: ['last name', 'lastname', 'last', 'surname', 'family name', 'lname'] },
  { token: '_SUFFIX0_', label: 'Suffix', match: ['suffix', 'name suffix'] },
  { token: '_NICK0_', label: 'Nickname', match: ['nickname', 'nick name', 'known as'] },
  { token: '_COMPANY_', label: 'Company', match: ['company', 'company name', 'organisation', 'organization', 'business', 'employer'] },
  { token: '_STREET1_', label: 'Address 1', match: ['address 1', 'address line 1', 'address1', 'street', 'street address', 'address', 'addr1'] },
  { token: '_STREET2_', label: 'Address 2', match: ['address 2', 'address line 2', 'address2', 'addr2'] },
  { token: '_CITY_', label: 'City', match: ['city', 'town', 'town city', 'locality'] },
  { token: '_STATE_', label: 'State', match: ['state', 'county', 'province', 'region', 'state province'] },
  { token: '_POSTAL_', label: 'Postal Code', match: ['postal code', 'postcode', 'post code', 'zip', 'zip code', 'postal'] },
  { token: '_COUNTRY_', label: 'Country or Region', match: ['country or region', 'country', 'country region', 'nation'] },
  { token: '_EMAIL_', label: 'E-mail Address', match: ['e mail address', 'email address', 'email', 'e mail', 'mail'] },
];

const loose = (s) => fold(s).replace(/[^a-z0-9 ]+/g, ' ').replace(/\s+/g, ' ').trim();

/**
 * Which column stands for each address field: an explicit choice from Match
 * Fields first, then the first column whose name Word would have guessed.
 * A field nothing matches is left out — it prints as nothing.
 */
export function matchFields(fields, overrides = {}) {
  const out = {};
  const list = fields || [];
  for (const f of ADDRESS_FIELDS) {
    if (Object.prototype.hasOwnProperty.call(overrides, f.token)) {
      if (overrides[f.token]) out[f.token] = overrides[f.token];
      continue;
    }
    const hit = list.find((c) => f.match.includes(loose(c)));
    if (hit) out[f.token] = hit;
  }
  return out;
}

/* ── Address Block and Greeting Line ───────────────────────────────────── */

/** Insert Address Block's name formats (one person), with Word's own sample as the label. */
export const ADDRESS_NAME_FORMATS = [
  { id: 'first', label: 'Joshua', format: '<<_FIRST0_>>' },
  { id: 'firstLastSuffix', label: 'Joshua Randall Jr.', format: '<<_FIRST0_>><< _LAST0_>><< _SUFFIX0_>>' },
  { id: 'firstMiddleLastSuffix', label: 'Joshua Q. Randall Jr.', format: '<<_FIRST0_>><< _MIDDLE0_>><< _LAST0_>><< _SUFFIX0_>>' },
  { id: 'titleNickLastSuffix', label: 'Mr. Josh Randall Jr.', format: '<<_TITLE0_ >><<_NICK0_>><< _LAST0_>><< _SUFFIX0_>>' },
  { id: 'titleNickMiddleLastSuffix', label: 'Mr. Josh Q. Randall Jr.', format: '<<_TITLE0_ >><<_NICK0_>><< _MIDDLE0_>><< _LAST0_>><< _SUFFIX0_>>' },
  { id: 'titleFirstLastSuffix', label: 'Mr. Joshua Randall Jr.', format: '<<_TITLE0_ >><<_FIRST0_>><< _LAST0_>><< _SUFFIX0_>>' },
  { id: 'titleFirstMiddleLastSuffix', label: 'Mr. Joshua Q. Randall Jr.', format: '<<_TITLE0_ >><<_FIRST0_>><< _MIDDLE0_>><< _LAST0_>><< _SUFFIX0_>>' },
];

/** Greeting Line's name formats (one person). */
export const GREETING_NAME_FORMATS = [
  { id: 'nick', label: 'Josh', format: '<<_NICK0_>>' },
  { id: 'first', label: 'Joshua', format: '<<_FIRST0_>>' },
  { id: 'titleLast', label: 'Mr. Randall', format: '<<_TITLE0_ >><<_LAST0_>>' },
  { id: 'titleNickLast', label: 'Mr. Josh Randall', format: '<<_TITLE0_ >><<_NICK0_>><< _LAST0_>>' },
  { id: 'titleFirstLast', label: 'Mr. Joshua Randall', format: '<<_TITLE0_ >><<_FIRST0_>><< _LAST0_>>' },
  { id: 'firstLast', label: 'Joshua Randall', format: '<<_FIRST0_>><< _LAST0_>>' },
  { id: 'nickLast', label: 'Josh Randall', format: '<<_NICK0_>><< _LAST0_>>' },
];

export const GREETING_SALUTATIONS = ['Dear', 'To', ''];
export const GREETING_PUNCTUATION = [',', ':', ''];
export const GREETING_FALLBACKS = ['Dear Sir or Madam,', 'To Whom It May Concern:', ''];

/** The address format Word writes for the choices Insert Address Block offers. */
export function addressBlockFormat({ name = 'titleFirstLastSuffix', company = true, postal = true, country = 'unlessEqual' } = {}) {
  // `name: null` is the dialog's "Insert recipient's name" left unticked,
  // `postal: false` its "Insert postal address".
  const person = name === null ? '' : (ADDRESS_NAME_FORMATS.find((f) => f.id === name) || ADDRESS_NAME_FORMATS[5]).format + '\r';
  return person + (company ? '<<_COMPANY_\r>>' : '') +
    (postal ? '<<_STREET1_\r>><<_STREET2_\r>><<_CITY_>><<, _STATE_>><< _POSTAL_>>' + (country === 'never' ? '' : '<<\r_COUNTRY_>>') : '');
}

/**
 * ` ADDRESSBLOCK \f "…" \l 1033 \c 2 \e "United States" \d ` — the field
 * code Word's dialog writes. `country`: 'never' (the format has no country),
 * 'always', or 'unlessEqual' with `except` the country left off.
 */
export function addressBlockInstr({ name = 'titleFirstLastSuffix', company = true, postal = true, country = 'unlessEqual', except = 'United Kingdom', language = 1033 } = {}) {
  const f = addressBlockFormat({ name, company, postal, country });
  return ' ADDRESSBLOCK \\f "' + f + '" \\l ' + language + ' \\c ' + (company ? 2 : 0) +
    (country === 'unlessEqual' && except ? ' \\e "' + except + '"' : '') + ' \\d ';
}

/** ` GREETINGLINE \f "<<_BEFORE_ Dear >><<_TITLE0_ >><<_LAST0_>><<_AFTER_ ,>>" \l 1033 \e "Dear Sir or Madam," `. */
export function greetingLineInstr({ salutation = 'Dear', name = 'titleLast', punctuation = ',', fallback = 'Dear Sir or Madam,', language = 1033 } = {}) {
  const person = (GREETING_NAME_FORMATS.find((f) => f.id === name) || GREETING_NAME_FORMATS[2]).format;
  return ' GREETINGLINE \\f "' + (salutation ? '<<_BEFORE_ ' + salutation + ' >>' : '') + person +
    (punctuation ? '<<_AFTER_ ' + punctuation + '>>' : '') + '" \\l ' + language + (fallback ? ' \\e "' + fallback + '"' : '') + ' ';
}

/**
 * A format string's `<<…>>` groups filled in: a group whose token has no
 * value is left out whole — its spaces, its comma and its line break with
 * it — which is how "Mr. Joshua Randall Jr." loses the "Jr." and an address
 * with no second line has no blank line in it. Text outside a group stays.
 * `_BEFORE_` and `_AFTER_` are the greeting's own words either side of the
 * name, the one space after the token only a separator.
 */
export function fillFormat(format, valueOf) {
  let out = '';
  let named = false;
  let person = false;
  const re = /<<([\s\S]*?)>>/g;
  let last = 0;
  let m;
  while ((m = re.exec(String(format)))) {
    out += String(format).slice(last, m.index);
    last = m.index + m[0].length;
    const inner = m[1];
    const t = /_([A-Z]+\d?)_/.exec(inner);
    if (!t) { out += inner; continue; }
    const token = t[0];
    if (token === '_BEFORE_' || token === '_AFTER_') {
      out += inner.slice(t.index + token.length).replace(/^ /, '');
      continue;
    }
    const value = String(valueOf(token) ?? '').trim();
    if (!value) continue;
    named = true;
    // A courtesy title alone is not a name: "Dear Dr.," is not a greeting.
    if (!/^_TITLE\d_$/.test(token)) person = true;
    out += inner.slice(0, t.index) + value + inner.slice(t.index + token.length);
  }
  out += String(format).slice(last);
  return { text: out, named, person };
}

/** Lines, each trimmed of trailing spaces, the empty ones gone. */
const cleanLines = (text) => String(text).split(/\r\n|\r|\n|\v/).map((l) => l.replace(/\s+$/, '')).filter((l) => l.trim() !== '');

/** The value an address token takes for a record, through the matched columns. */
function tokenValue(source, record, mapping, token) {
  const col = mapping?.[token];
  if (!col) {
    // A nickname nobody gave falls back to the first name, as Word's does.
    if (token === '_NICK0_' && mapping?._FIRST0_) return fieldValue(source, record, mapping._FIRST0_);
    return '';
  }
  return fieldValue(source, record, col);
}

/** An ADDRESSBLOCK's switches read back: `{ format, company, except, language }`. */
export function readAddressBlock(instr) {
  const sw = switches(instr);
  return {
    format: sw.f ?? addressBlockFormat(),
    company: sw.c !== '0',
    except: sw.e ?? null,
    language: sw.l ?? '1033',
  };
}

/**
 * An address block for one record, as lines. A company, a second address
 * line or a region the record leaves blank takes no line; the country goes
 * when it is the one the field says to leave off.
 */
export function formatAddressBlock(instr, source, record, mapping) {
  const spec = readAddressBlock(instr);
  let format = spec.format;
  if (!spec.company) format = format.replace(/<<_COMPANY_[^>]*>>/g, '');
  const country = tokenValue(source, record, mapping, '_COUNTRY_');
  const drop = spec.except && fold(country) === fold(spec.except);
  const { text } = fillFormat(format, (token) => (token === '_COUNTRY_' && drop ? '' : tokenValue(source, record, mapping, token)));
  return cleanLines(text);
}

/**
 * A greeting line for one record: "Dear Mr. Randall," — or, when the record
 * has none of the name the format asks for, the field's fallback ("Dear Sir
 * or Madam,"), never a bare "Dear ,".
 */
export function formatGreeting(instr, source, record, mapping) {
  const sw = switches(instr);
  const format = sw.f ?? '<<_BEFORE_ Dear >><<_TITLE0_ >><<_LAST0_>><<_AFTER_ ,>>';
  const { text, person } = fillFormat(format, (token) => tokenValue(source, record, mapping, token));
  if (!person) return sw.e ?? '';
  return text.replace(/\s+([,:;.!?])$/, '$1').replace(/\s{2,}/g, ' ').trim();
}

/* ── field codes ───────────────────────────────────────────────────────── */

/**
 * A field code's words: the first says the kind; quoted text is one word
 * with its quotes gone; `{…}` is a nested field, kept as it stands for the
 * caller to evaluate; a switch (`\f`, `\*`) is a word of its own.
 */
export function splitInstr(instr) {
  const s = String(instr ?? '');
  const out = [];
  let i = 0;
  while (i < s.length) {
    const c = s[i];
    if (/\s/.test(c)) { i++; continue; }
    if (c === '"') {
      let j = i + 1;
      let word = '';
      while (j < s.length && s[j] !== '"') {
        if (s[j] === '\\' && s[j + 1] === '"') { word += '"'; j += 2; continue; }
        word += s[j++];
      }
      out.push({ text: word, quoted: true });
      i = j + 1;
      continue;
    }
    if (c === '{') {
      let depth = 0;
      let j = i;
      for (; j < s.length; j++) {
        if (s[j] === '{') depth++;
        else if (s[j] === '}' && --depth === 0) break;
      }
      out.push({ text: s.slice(i + 1, j), nested: true });
      i = j + 1;
      continue;
    }
    let j = i;
    while (j < s.length && !/\s/.test(s[j]) && s[j] !== '"' && s[j] !== '{') j++;
    out.push({ text: s.slice(i, j) });
    i = j;
  }
  return out;
}

/** A field code's switches: `\f "…"` → `{ f: '…' }`; `\* Upper` → `{ '*': ['Upper'] }`. */
export function switches(instr) {
  const words = splitInstr(instr);
  const out = {};
  for (let i = 0; i < words.length; i++) {
    const w = words[i];
    if (w.quoted || w.nested || !/^\\[^\s]/.test(w.text)) continue;
    const key = w.text.slice(1);
    const next = words[i + 1];
    const arg = next && !(!next.quoted && !next.nested && /^\\/.test(next.text)) ? next.text : null;
    if (key === '*') (out['*'] ||= []).push(arg);
    else out[key] = arg;
    if (arg !== null) i++;
  }
  return out;
}

/** What kind of field an instruction is, and the name a MERGEFIELD reads. */
export function readInstr(instr) {
  const words = splitInstr(instr);
  const kind = (words[0]?.text || '').toLowerCase();
  const name = kind === 'mergefield' && words[1] ? words[1].text : null;
  return { kind, name, words };
}

/** A MERGEFIELD's `\*` formatting — Upper, Lower, Caps, FirstCap — applied as Word does. */
function applyCase(text, formats = []) {
  let t = text;
  for (const f of formats || []) {
    const k = String(f || '').toLowerCase();
    if (k === 'upper') t = t.toUpperCase();
    else if (k === 'lower') t = t.toLowerCase();
    else if (k === 'caps') t = t.replace(/\b(\p{L})(\p{L}*)/gu, (_, a, b) => a.toUpperCase() + b.toLowerCase());
    else if (k === 'firstcap') t = t.replace(/^\s*\p{L}/u, (m) => m.toUpperCase());
  }
  return t;
}

/** `=`/`<>` with `*` and `?` wildcards, as IF compares text. */
function wild(pattern, text) {
  const re = new RegExp('^' + String(pattern).replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '[\\s\\S]*').replace(/\?/g, '.') + '$');
  return re.test(String(text));
}

/** An IF's comparison, the way Word makes it: numbers as numbers, text as text, case counting. */
export function compare(left, op, right) {
  const a = String(left ?? '');
  const b = String(right ?? '');
  const na = Number(a);
  const nb = Number(b);
  const numeric = a.trim() !== '' && b.trim() !== '' && Number.isFinite(na) && Number.isFinite(nb);
  const c = numeric ? na - nb : a < b ? -1 : a > b ? 1 : 0;
  switch (op) {
    case '=': return numeric ? c === 0 : /[*?]/.test(b) ? wild(b, a) : a === b;
    case '<>': return numeric ? c !== 0 : /[*?]/.test(b) ? !wild(b, a) : a !== b;
    case '<': return c < 0;
    case '>': return c > 0;
    case '<=': return c <= 0;
    case '>=': return c >= 0;
    default: return false;
  }
}

const OPS = ['<>', '<=', '>=', '=', '<', '>'];

/**
 * A condition — `{ MERGEFIELD City } = "London"` — split into its three
 * parts, each nested field evaluated. An operator glued to a word (`X="a"`)
 * is found too.
 */
function condition(words, ctx) {
  const flat = [];
  for (const w of words) {
    if (w.nested || w.quoted) { flat.push(w); continue; }
    let s = w.text;
    while (s) {
      const op = OPS.find((o) => s.startsWith(o));
      if (op) { flat.push({ text: op, op: true }); s = s.slice(op.length); continue; }
      const at = Math.min(...OPS.map((o) => { const k = s.indexOf(o); return k < 0 ? Infinity : k; }));
      if (at === Infinity) { flat.push({ text: s }); break; }
      flat.push({ text: s.slice(0, at) });
      s = s.slice(at);
    }
  }
  const value = (w) => (w?.nested ? evaluateField(w.text, ctx).text : w?.text ?? '');
  const opAt = flat.findIndex((w) => w.op);
  if (opAt < 0) return { ok: false, rest: flat };
  return { ok: true, left: value(flat[opAt - 1]), op: flat[opAt].text, right: value(flat[opAt + 1]), rest: flat.slice(opAt + 2) };
}

/**
 * One field's result for the record a merge is at.
 *
 * `ctx`: `{ source, record, recordNumber, sequence, mapping }` — the record
 * being merged (an array in the source's field order), its 1-based number in
 * the source and in this merge, and the Match Fields mapping. Answers
 * `{ text, next, skip }`: the words to put in the field's place, whether it
 * moves the merge to the next record (NEXT, a true NEXTIF), and whether it
 * throws this record's copy away (a true SKIPIF). An ADDRESSBLOCK's lines
 * come back joined by "\n", which becomes a line break in the document.
 */
export function evaluateField(instr, ctx = {}) {
  const { kind, words } = readInstr(instr);
  const sw = switches(instr);
  const { source, record, mapping } = ctx;
  switch (kind) {
    case 'mergefield': {
      const name = words[1]?.text || '';
      let text = applyCase(fieldValue(source, record, name), sw['*']);
      if (text) text = (sw.b ?? '') + text + (sw.f ?? '');
      return { text };
    }
    case 'addressblock':
      return { text: formatAddressBlock(instr, source, record, mapping).join('\n') };
    case 'greetingline':
      return { text: formatGreeting(instr, source, record, mapping) };
    case 'if': {
      const cond = condition(words.slice(1), ctx);
      if (!cond.ok) return { text: '' };
      const [yes, no] = cond.rest.filter((w) => !(w.text === '' && !w.quoted));
      const pick = compare(cond.left, cond.op, cond.right) ? yes : no;
      return { text: pick ? (pick.nested ? evaluateField(pick.text, ctx).text : pick.text) : '' };
    }
    case 'next':
      return { text: '', next: true };
    case 'nextif': {
      const cond = condition(words.slice(1), ctx);
      return { text: '', next: cond.ok && compare(cond.left, cond.op, cond.right) };
    }
    case 'skipif': {
      const cond = condition(words.slice(1), ctx);
      return { text: '', skip: cond.ok && compare(cond.left, cond.op, cond.right) };
    }
    case 'mergerec':
      return { text: String(ctx.recordNumber ?? '') };
    case 'mergeseq':
      return { text: String(ctx.sequence ?? '') };
    default:
      return { text: '' };
  }
}

/* ── the fields as Word writes them ────────────────────────────────────── */

/** `MERGEFIELD FirstName \* MERGEFORMAT` for a column. */
export function mergeFieldInstr(header) {
  const name = mergeFieldName(header);
  return ' MERGEFIELD ' + (/[\s"]/.test(name) ? '"' + name.replace(/"/g, '\\"') + '"' : name) + ' \\* MERGEFORMAT ';
}

const q = (s) => '"' + String(s ?? '').replace(/"/g, '\\"') + '"';

/**
 * Rules → If…Then…Else: `IF { MERGEFIELD City } = "London" "…" "…"`.
 * `comparison`: '=', '<>', '<', '>', '<=', '>=', or 'blank'/'notBlank',
 * which Word writes as a comparison with "".
 */
export function ifInstr({ field, comparison = '=', value = '', then = '', otherwise = '' }) {
  const op = comparison === 'blank' ? '=' : comparison === 'notBlank' ? '<>' : comparison;
  const right = comparison === 'blank' || comparison === 'notBlank' ? '' : value;
  return ' IF {' + mergeFieldInstr(field).replace(/ \\\* MERGEFORMAT $/, ' ') + '} ' + op + ' ' + q(right) + ' ' + q(then) + ' ' + q(otherwise) + ' ';
}

/** Rules → Skip Record If / Next Record If: the same comparison, no texts. */
export function conditionInstr(kind, { field, comparison = '=', value = '' }) {
  const op = comparison === 'blank' ? '=' : comparison === 'notBlank' ? '<>' : comparison;
  const right = comparison === 'blank' || comparison === 'notBlank' ? '' : value;
  return ' ' + kind.toUpperCase() + ' {' + mergeFieldInstr(field).replace(/ \\\* MERGEFORMAT $/, ' ') + '} ' + op + ' ' + q(right) + ' ';
}

/** The words a field shows in the main document: «FirstName», «AddressBlock», an IF's result. */
export function placeholderFor(instr, ctx = null) {
  const { kind, words } = readInstr(instr);
  if (kind === 'mergefield') return '«' + (words[1]?.text || '') + '»';
  if (kind === 'if') {
    // An IF shows its result for the record in hand; with none, the words it
    // would print when the comparison fails — and never nothing, or the field
    // would be an invisible character nobody could find to delete.
    const got = ctx?.record ? evaluateField(instr, ctx).text : '';
    if (got) return got;
    const texts = splitInstr(instr).filter((w) => w.quoted);
    return texts[texts.length - 1]?.text || texts[texts.length - 2]?.text || '«IF»';
  }
  return PLACEHOLDERS[kind] || '«' + kind.toUpperCase() + '»';
}

/** Word's schema order inside `w:rPr`, for putting `w:noProof` where Word does. */
const AFTER_NO_PROOF = /<w:(?:snapToGrid|vanish|webHidden|color|spacing|w|kern|position|sz|szCs|highlight|u|effect|bdr|shd|fitText|vertAlign|rtl|cs|em|lang|eastAsianLayout|specVanish|oMath)\b/;

/** A run's properties with `<w:noProof/>` — what Word puts on a merge field's result. */
export function withNoProof(rPr) {
  if (!rPr) return '<w:rPr><w:noProof/></w:rPr>';
  if (/<w:noProof\b/.test(rPr)) return rPr;
  if (/<w:rPr\b[^>]*\/>$/.test(rPr)) return '<w:rPr><w:noProof/></w:rPr>';
  const m = AFTER_NO_PROOF.exec(rPr);
  return m ? rPr.slice(0, m.index) + '<w:noProof/>' + rPr.slice(m.index) : rPr.replace(/<\/w:rPr>$/, '<w:noProof/></w:rPr>');
}

/** One `w:r` of a field: its properties, then the one element. */
const fieldRun = (rPr, inner) => '<w:r>' + (rPr || '') + inner + '</w:r>';

/**
 * A complex field in Word's own five runs — begin, the instruction, separate,
 * the result, end — with a nested field (`{…}` in the instruction, the way
 * an IF names a MERGEFIELD) written as a field of its own inside it.
 */
export function complexFieldXml(instr, result, rPr = '') {
  const parts = [];
  let i = 0;
  const s = String(instr);
  while (i < s.length) {
    const open = s.indexOf('{', i);
    if (open < 0) { parts.push(s.slice(i)); break; }
    if (open > i) parts.push(s.slice(i, open));
    let depth = 0;
    let j = open;
    for (; j < s.length; j++) {
      if (s[j] === '{') depth++;
      else if (s[j] === '}' && --depth === 0) break;
    }
    parts.push({ nested: s.slice(open + 1, j) });
    i = j + 1;
  }
  let xml = fieldRun(rPr, '<w:fldChar w:fldCharType="begin"/>');
  for (const p of parts) {
    if (typeof p === 'string') {
      if (p) xml += fieldRun(rPr, '<w:instrText xml:space="preserve">' + esc(p).replace(/\r/g, '&#13;') + '</w:instrText>');
    } else {
      xml += complexFieldXml(p.nested, placeholderFor(p.nested), rPr);
    }
  }
  xml += fieldRun(rPr, '<w:fldChar w:fldCharType="separate"/>');
  const lines = String(result ?? '').split('\n');
  xml += '<w:r>' + withNoProof(rPr) + lines.map((l, k) => (k ? '<w:br/>' : '') + (l ? '<w:t xml:space="preserve">' + esc(l) + '</w:t>' : '')).join('') + '</w:r>';
  xml += fieldRun(rPr, '<w:fldChar w:fldCharType="end"/>');
  return xml;
}

/* ── sending ───────────────────────────────────────────────────────────── */

/**
 * Finish & Merge → Send E-mail Messages: one message at a time through
 * `send(message)` (the Mail service, or under the checks its fake
 * transport), `onProgress(done, total, message)` after each. A message that
 * fails is counted and named, and the rest still go — one bad address must
 * not stop a hundred good ones.
 */
export async function sendMergedMessages(messages, send, onProgress = null) {
  const out = { sent: 0, failed: [] };
  const list = messages || [];
  for (let i = 0; i < list.length; i++) {
    const m = list[i];
    try {
      if (!m.to) throw new Error('no address');
      await send(m);
      out.sent++;
    } catch (err) {
      out.failed.push({ record: m.record, to: m.to || '', error: String(err?.message || err) });
    }
    onProgress?.(i + 1, list.length, m);
  }
  return out;
}
