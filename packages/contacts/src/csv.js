// The address books people export as CSV — Google Contacts and Outlook —
// read into the same contact shape a vCard gives. The columns are matched by
// name, so an export from either, in either of their layouts, comes in whole,
// and a spreadsheet somebody typed with "Name" and "Email" columns does too.

import { splitName } from './vcard.js';

/** A CSV row splitter that honours quotes, doubled quotes and embedded newlines. */
export function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = '';
  let quoted = false;
  const s = String(text || '').replace(/^﻿/, '');
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
    else if (ch === ',') {
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
  full: ['name', 'full name', 'display name'],
  nickname: ['nickname'],
  org: ['company', 'organization 1 name', 'organisation', 'organization'],
  department: ['department', 'organization 1 department'],
  jobTitle: ['job title', 'organization 1 title'],
  birthday: ['birthday'],
  note: ['notes', 'note'],
  categories: ['categories', 'group membership', 'labels'],
  website: ['web page', 'website 1 value', 'website'],
};

const EMAIL = [/^e mail( address)?$/, /^e mail (\d) value$/, /^e mail (\d) address$/, /^email( address)?$/, /^e mail (\d)$/, /^(home|work|other) email$/];
const PHONE = [/^(mobile|home|business|work|other|primary|home 2|business 2|car|pager|company main) phone$/, /^phone (\d) value$/, /^phone$/, /^telephone$/];
const ADDRESS_PARTS = { street: ['street', 'address'], city: ['city'], region: ['state', 'region', 'province', 'county'], postcode: ['postal code', 'post code', 'zip', 'postcode'], country: ['country', 'country region'] };

function typeFromHeader(h) {
  const m = /^(home|work|business|mobile|other|main|cell|company main)/.exec(h);
  const t = m ? m[1] : 'other';
  return t === 'business' ? 'work' : t === 'cell' ? 'mobile' : t === 'company main' ? 'work' : t;
}

/** Contacts from a CSV export. Rows with neither a name nor an address are skipped. */
export function readContactsCsv(text) {
  const rows = parseCsv(text);
  if (rows.length < 2) return [];
  const headers = rows[0].map(norm);
  const index = (names) => headers.findIndex((h) => names.includes(h));
  const col = Object.fromEntries(Object.entries(COLUMNS).map(([k, names]) => [k, index(names)]));
  const emailCols = headers.map((h, i) => (EMAIL.some((re) => re.test(h)) ? i : -1)).filter((i) => i >= 0);
  const phoneCols = headers.map((h, i) => (PHONE.some((re) => re.test(h)) ? i : -1)).filter((i) => i >= 0);
  // Address columns come in sets: "Home Street", "Home City"… or "Address 1 - Street".
  const addressSets = new Map();
  headers.forEach((h, i) => {
    for (const [part, names] of Object.entries(ADDRESS_PARTS)) {
      const m = names.map((n) => new RegExp(`^(home|work|business|other|address \\d)( address)? (${n})$|^(${n})$`)).find((re) => re.test(h));
      if (m) {
        const key = (/^(home|work|business|other|address \d)/.exec(h) || [, 'other'])[1];
        if (!addressSets.has(key)) addressSets.set(key, {});
        addressSets.get(key)[part] = i;
      }
    }
  });
  // Google pairs "Phone 1 - Value" with "Phone 1 - Type"; a plain "Mobile
  // Phone" carries its type in its name and has no partner column.
  const typeCol = (i) => {
    const h = headers[i];
    if (!/ value$/.test(h)) return -1;
    return headers.findIndex((x) => x === h.replace(/ value$/, ' type'));
  };
  const addressTypeCol = (key) => headers.findIndex((x) => x === `${key} type`);

  const contacts = [];
  for (const row of rows.slice(1)) {
    const get = (i) => (i >= 0 && i < row.length ? String(row[i] || '').trim() : '');
    const given = get(col.first);
    const family = get(col.last);
    const full = get(col.full) || [get(col.prefix), given, get(col.middle), family].filter(Boolean).join(' ');
    const name = given || family ? { full, given, family, middle: get(col.middle), prefix: get(col.prefix), suffix: '' } : splitName(full);
    const emails = [];
    for (const i of emailCols) {
      for (const v of get(i).split(/\s*:::\s*|\s*;\s*/)) {
        if (!v || !v.includes('@')) continue;
        const t = typeCol(i);
        emails.push({ value: v, type: (t >= 0 ? get(t).replace(/^\*\s*/, '') : typeFromHeader(headers[i])).toLowerCase() || 'other', pref: emails.length === 0, label: null });
      }
    }
    const phones = [];
    for (const i of phoneCols) {
      for (const v of get(i).split(/\s*:::\s*|\s*;\s*/)) {
        if (!v) continue;
        const t = typeCol(i);
        phones.push({ value: v, type: (t >= 0 ? get(t).replace(/^\*\s*/, '') : typeFromHeader(headers[i])).toLowerCase() || 'other', pref: phones.length === 0, label: null });
      }
    }
    const addresses = [];
    for (const [key, parts] of addressSets) {
      const typed = addressTypeCol(key);
      const type = (typed >= 0 ? get(typed).replace(/^\*\s*/, '').toLowerCase() : '') || typeFromHeader(key);
      const a = { street: get(parts.street), city: get(parts.city), region: get(parts.region), postcode: get(parts.postcode), country: get(parts.country), type, pref: addresses.length === 0, label: null };
      if (a.street || a.city || a.postcode || a.country) addresses.push(a);
    }
    if (!name.full && !emails.length) continue;
    contacts.push({
      uid: null,
      version: '3.0',
      name,
      nickname: get(col.nickname) || null,
      org: get(col.org) || null,
      department: get(col.department) || null,
      title: get(col.jobTitle) || null,
      role: null,
      emails,
      phones,
      addresses,
      birthday: get(col.birthday) ? get(col.birthday).replace(/\//g, '-') : null,
      anniversary: null,
      note: get(col.note) || null,
      urls: get(col.website) ? [{ value: get(col.website), type: 'other', label: null }] : [],
      photo: null,
      categories: get(col.categories) ? get(col.categories).split(/\s*:::\s*|\s*;\s*|\s*,\s*/).filter(Boolean).map((s) => s.replace(/^\*\s*/, '')) : [],
      rev: null,
      kind: 'individual',
      extra: [],
    });
  }
  return contacts;
}

export default { readContactsCsv, parseCsv };
