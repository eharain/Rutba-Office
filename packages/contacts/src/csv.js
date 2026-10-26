// The address books people export as CSV — Google Contacts and Outlook —
// read into the same contact shape a vCard gives. The columns are matched by
// name, so an export from either, in either of their layouts, comes in whole,
// and a spreadsheet somebody typed with "Name" and "Email" columns does too.
// What each column was matched to is there to be shown before anything is
// kept, and changed (`columns`); what was left out is said, not lost quietly.

import { splitName } from './vcard.js';

/**
 * A file's bytes as text: UTF-8 when they are UTF-8 (with or without its
 * mark), UTF-16 when a mark says so, and otherwise the Windows ANSI code page
 * Outlook and Excel write in (windows-1252), so "Zoë" is not "Zo�".
 */
export function decodeText(bytes) {
  if (typeof bytes === 'string') return bytes;
  const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (b[0] === 0xff && b[1] === 0xfe) return new TextDecoder('utf-16le').decode(b.subarray(2));
  if (b[0] === 0xfe && b[1] === 0xff) return new TextDecoder('utf-16be').decode(b.subarray(2));
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(b);
  } catch {
    return new TextDecoder('windows-1252').decode(b);
  }
}

/** The separator a file uses: a comma, or the semicolon or tab Excel writes in some languages, by its first line. */
function separatorOf(s) {
  let quoted = false;
  const count = { ',': 0, ';': 0, '\t': 0 };
  for (const ch of s) {
    if (ch === '"') quoted = !quoted;
    else if (!quoted && (ch === '\n' || ch === '\r')) break;
    else if (!quoted && ch in count) count[ch] += 1;
  }
  return count[';'] > count[','] && count[';'] >= count['\t'] ? ';' : count['\t'] > count[','] ? '\t' : ',';
}

/** A CSV row splitter that honours quotes, doubled quotes and embedded newlines. */
export function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = '';
  let quoted = false;
  const s = String(text || '').replace(/^﻿/, '');
  const sep = separatorOf(s);
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (quoted) {
      if (ch === '"') {
        if (s[i + 1] === '"') {
          field += '"';
          i += 1;
        } else quoted = false;
      } else field += ch;
    } else if (ch === '"') quoted = true;
    else if (ch === sep) {
      row.push(field);
      field = '';
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && s[i + 1] === '\n') i += 1;
      row.push(field);
      field = '';
      if (row.some((f) => f !== '')) rows.push(row);
      row = [];
    } else field += ch;
  }
  row.push(field);
  if (row.some((f) => f !== '')) rows.push(row);
  return rows;
}

const norm = (h) => String(h || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();

/** Which column holds what, from the header row. Each entry is a list of header spellings. */
const COLUMNS = {
  first: ['first name', 'given name', 'firstname'],
  last: ['last name', 'family name', 'surname', 'lastname'],
  middle: ['middle name', 'additional name'],
  prefix: ['title', 'name prefix', 'honorific prefix'],
  suffix: ['suffix', 'name suffix', 'honorific suffix'],
  full: ['name', 'full name', 'display name'],
  nickname: ['nickname'],
  // Google's own headings have changed: "Organization 1 - Name" until 2023, "Organization Name" now.
  org: ['company', 'organization name', 'organization 1 name', 'organisation', 'organization'],
  department: ['department', 'organization department', 'organization 1 department'],
  jobTitle: ['job title', 'organization title', 'organization 1 title'],
  birthday: ['birthday'],
  note: ['notes', 'note'],
  categories: ['categories', 'group membership', 'labels'],
  website: ['web page', 'website 1 value', 'website'],
};

/** What a column can be read as, in the order a preview offers them; '' leaves it out. */
export const CSV_FIELDS = ['full', 'prefix', 'first', 'middle', 'last', 'suffix', 'nickname', 'org', 'department', 'jobTitle', 'email', 'phone', 'street', 'city', 'region', 'postcode', 'country', 'birthday', 'website', 'note', 'categories'];

const EMAIL = [/^e mail( address)?$/, /^e mail (\d) value$/, /^e mail (\d) address$/, /^email( address)?$/, /^e mail (\d)$/, /^(home|work|other) email$/];
const PHONE = [/^(mobile|home|business|work|other|primary|home 2|business 2|car|pager|company main) phone$/, /^phone (\d) value$/, /^phone$/, /^telephone$/];
// Each address part, and Google's "Formatted" (the whole address on one line), "PO Box" and "Extended Address".
const ADDRESS_PARTS = { street: ['street', 'address'], city: ['city'], region: ['state', 'region', 'province', 'county'], postcode: ['postal code', 'post code', 'zip', 'postcode'], country: ['country', 'country region'], pobox: ['po box', 'p o box'], extended: ['extended address'], formatted: ['formatted'] };
const ADDRESS_RES = Object.entries(ADDRESS_PARTS).map(([part, names]) => [part, names.map((n) => new RegExp(`^(home|work|business|other|address \\d)( address)? (${n})$|^(${n})$`))]);
// Google pairs "Phone 1 - Value" with "Phone 1 - Label" (once "- Type"), and an address set with its own.
const PARTNER = /^(e mail|phone|website|address) (\d) (type|label)$/;

/** The kinds a card keeps for each, and the other words Google and Outlook use for them. */
const KINDS = {
  email: { home: 'home', work: 'work', business: 'work', other: 'other' },
  phone: { mobile: 'mobile', cell: 'mobile', 'mobile phone': 'mobile', home: 'home', work: 'work', business: 'work', fax: 'fax', 'work fax': 'fax', 'home fax': 'fax', 'business fax': 'fax', other: 'other' },
  address: { home: 'home', work: 'work', business: 'work', other: 'other' },
  website: { home: 'home', work: 'work', other: 'other' },
};

/** A label as a file wrote it ("* Work", "Mobile", "Google Voice") as a kind, and the words themselves when they are none of the kinds. */
function kindFrom(raw, list) {
  const words = String(raw || '').replace(/^\*\s*/, '').trim();
  if (!words) return null;
  const known = KINDS[list][words.toLowerCase()];
  return known ? { type: known, label: null } : { type: 'other', label: words };
}

function typeFromHeader(h) {
  const m = /^(home|work|business|mobile|other|main|cell|company main)/.exec(h);
  const t = m ? m[1] : 'other';
  return t === 'business' ? 'work' : t === 'cell' ? 'mobile' : t === 'company main' ? 'work' : t;
}

/** What one heading is read as, with the set it belongs to (an address's) or the column it labels. */
function guess(h, headers) {
  for (const [field, names] of Object.entries(COLUMNS)) if (names.includes(h)) return { field };
  if (EMAIL.some((re) => re.test(h))) return { field: 'email' };
  if (PHONE.some((re) => re.test(h))) return { field: 'phone' };
  const partner = PARTNER.exec(h);
  if (partner) {
    if (partner[1] === 'address') return { field: 'type', key: `address ${partner[2]}` };
    const of = headers.indexOf(`${partner[1]} ${partner[2]} value`);
    return of >= 0 ? { field: 'type', of } : { field: '' };
  }
  for (const [part, res] of ADDRESS_RES) {
    if (res.some((re) => re.test(h))) return { field: part, key: (/^(home|work|business|other|address \d)/.exec(h) || [, 'other'])[1] };
  }
  if (/^(home|work|business|other) address type$/.test(h)) return { field: 'type', key: h.replace(/ address type$/, '') };
  return { field: '' };
}

/**
 * Every column with what it is read as: `field` one of CSV_FIELDS, 'type'
 * (the kind of the column `of`, or of the address set `key`), or '' (left
 * out). `columns` ({ [index]: field }) changes what the person changed.
 */
function plan(headers, columns = {}) {
  return headers.map((header, index) => {
    const h = norm(header);
    const auto = guess(h, headers.map(norm));
    const asked = Object.prototype.hasOwnProperty.call(columns, index) ? columns[index] : undefined;
    if (asked === undefined || asked === auto.field) return { index, header, ...auto, auto: auto.field };
    if (asked && !CSV_FIELDS.includes(asked)) return { index, header, ...auto, auto: auto.field };
    // Moved to an address part: into its own set when the heading names one, else the "other" one.
    const key = ADDRESS_PARTS[asked] ? auto.key || 'other' : undefined;
    return { index, header, field: asked || '', key, auto: auto.field };
  });
}

/**
 * Contacts from a CSV export, with what the reading did: each column and what
 * it was read as (with the first value in it, for a preview), the rows left
 * out (nothing a card could keep: no name, address, company or phone) by
 * their line numbers, and the headings of the columns with values that were
 * not read. `columns` is a person's change to what a column is read as.
 */
export function readContactsCsvReport(text, { columns = {} } = {}) {
  const rows = parseCsv(decodeText(text));
  if (!rows.length) return { contacts: [], columns: [], skipped: [], unread: [] };
  const headers = rows[0];
  const cols = plan(headers, columns);
  const first = (field) => cols.find((c) => c.field === field)?.index ?? -1;
  const col = Object.fromEntries(Object.keys(COLUMNS).map((k) => [k, first(k)]));
  const emailCols = cols.filter((c) => c.field === 'email').map((c) => c.index);
  const phoneCols = cols.filter((c) => c.field === 'phone').map((c) => c.index);
  const websiteCols = cols.filter((c) => c.field === 'website').map((c) => c.index);
  // Address columns come in sets: "Home Street", "Home City"… or "Address 1 - Street".
  const addressSets = new Map();
  for (const c of cols) {
    if (!ADDRESS_PARTS[c.field]) continue;
    if (!addressSets.has(c.key)) addressSets.set(c.key, {});
    addressSets.get(c.key)[c.field] ??= c.index;
  }
  const typeOf = (i) => cols.find((c) => c.field === 'type' && c.of === i)?.index ?? -1;
  const addressTypeOf = (key) => cols.find((c) => c.field === 'type' && c.key === key)?.index ?? -1;

  const contacts = [];
  const skipped = [];
  const used = new Set();
  rows.slice(1).forEach((row, r) => {
    const get = (i) => (i >= 0 && i < row.length ? String(row[i] || '').trim() : '');
    const given = get(col.first);
    const family = get(col.last);
    const full = get(col.full) || [get(col.prefix), given, get(col.middle), family, get(col.suffix)].filter(Boolean).join(' ');
    const name = given || family ? { full, given, family, middle: get(col.middle), prefix: get(col.prefix), suffix: get(col.suffix) } : { ...splitName(full), ...(get(col.prefix) ? { prefix: get(col.prefix) } : {}), ...(get(col.suffix) ? { suffix: get(col.suffix) } : {}) };
    /** A list's entries from its columns, each with the kind its label or heading gives. */
    const entries = (list, indexes, keep) => {
      const out = [];
      for (const i of indexes) {
        const labelled = typeOf(i);
        const kind = (labelled >= 0 ? kindFrom(get(labelled), list) : null) || { type: (KINDS[list][typeFromHeader(norm(headers[i]))] || 'other'), label: null };
        for (const v of get(i).split(/\s*:::\s*|\s*;\s*/)) {
          if (!v || !keep(v)) continue;
          out.push({ value: v, ...kind, pref: out.length === 0 });
        }
      }
      return out.map(({ value, type, pref, label }) => ({ value, type, pref, label }));
    };
    const emails = entries('email', emailCols, (v) => v.includes('@'));
    const phones = entries('phone', phoneCols, () => true);
    const urls = entries('website', websiteCols, () => true).map(({ value, type, label }) => ({ value, type, label }));
    const addresses = [];
    for (const [key, parts] of addressSets) {
      const kind = kindFrom(get(addressTypeOf(key)), 'address') || { type: typeFromHeader(key), label: null };
      const street = [get(parts.pobox ?? -1), get(parts.extended ?? -1), get(parts.street ?? -1)].filter(Boolean).join('\n');
      const a = { street, city: get(parts.city ?? -1), region: get(parts.region ?? -1), postcode: get(parts.postcode ?? -1), country: get(parts.country ?? -1), type: kind.type, pref: addresses.length === 0, label: kind.label };
      // Only the one-line form: it is the street.
      if (!a.street && !a.city && !a.postcode && !a.country && get(parts.formatted ?? -1)) a.street = get(parts.formatted);
      if (a.street || a.city || a.postcode || a.country) addresses.push(a);
    }
    const org = get(col.org) || null;
    if (!name.full && !emails.length && !org && !phones.length) {
      skipped.push(r + 2);
      return;
    }
    row.forEach((v, i) => { if (String(v || '').trim()) used.add(i); });
    contacts.push({
      uid: null,
      version: '3.0',
      name,
      nickname: get(col.nickname) || null,
      org,
      department: get(col.department) || null,
      title: get(col.jobTitle) || null,
      role: null,
      emails,
      phones,
      addresses,
      birthday: get(col.birthday) ? get(col.birthday).replace(/\//g, '-') : null,
      anniversary: null,
      note: get(col.note) || null,
      urls,
      photo: null,
      // Google puts every contact in "* myContacts": its own group, not a person's tag.
      categories: get(col.categories) ? get(col.categories).split(/\s*:::\s*|\s*;\s*|\s*,\s*/).map((s) => s.replace(/^\*\s*/, '')).filter((s) => s && s.toLowerCase() !== 'mycontacts') : [],
      rev: null,
      kind: 'individual',
      extra: [],
    });
  });
  const sample = (i) => rows.slice(1).map((row) => String(row[i] || '').trim()).find(Boolean) || '';
  return {
    contacts,
    columns: cols.map(({ index, header, field, auto }) => ({ index, header, field, auto, sample: sample(index) })),
    skipped,
    unread: cols.filter((c) => !c.field && used.has(c.index)).map((c) => c.header),
  };
}

/** Contacts from a CSV export. Rows with nothing a card could keep are skipped. */
export function readContactsCsv(text, options = {}) {
  return readContactsCsvReport(text, options).contacts;
}

export default { readContactsCsv, readContactsCsvReport, parseCsv, decodeText, CSV_FIELDS };
