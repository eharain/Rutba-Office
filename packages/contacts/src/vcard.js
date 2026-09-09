// vCard (RFC 6350, and the 3.0 and 2.1 that came before it), read and written.
//
// The line grammar is iCalendar's — NAME;PARAM=VALUE:VALUE, folded at 75
// octets — and is borrowed from @rutba/calendar. What is vCard's own: N and
// ADR are lists of components split on semicolons; 2.1 writes bare
// parameters (TEL;HOME;VOICE) and QUOTED-PRINTABLE text with soft line
// breaks; Apple groups properties under item1., item2. to hang a label on
// them; a photo arrives as base64 inline (3.0), as a data: URI (4.0), or as
// a link; a birthday is 1985-03-15 in 3.0 and 19850315 or --0315 in 4.0.
//
// What is read becomes a plain contact; what is written comes from one.
// Properties this model does not name ride along as `extra`.

import { parseLine, escapeText, unescapeText, fold } from '@rutba/calendar/ical';

/* ── values ────────────────────────────────────────────────────────────── */

function decodeQuotedPrintable(value, charset = 'utf-8') {
  const bytes = [];
  const s = String(value).replace(/=\r?\n/g, '');
  for (let i = 0; i < s.length; i++) {
    if (s[i] === '=' && /^[0-9A-Fa-f]{2}$/.test(s.slice(i + 1, i + 3))) {
      bytes.push(parseInt(s.slice(i + 1, i + 3), 16));
      i += 2;
    } else bytes.push(s.charCodeAt(i) & 0xff);
  }
  try {
    return new TextDecoder(charset).decode(Uint8Array.from(bytes));
  } catch {
    return new TextDecoder('utf-8').decode(Uint8Array.from(bytes));
  }
}

/** A property's text value, whatever encoding the file used. */
function valueOf(prop) {
  const enc = (prop.params.ENCODING || [])[0]?.toUpperCase();
  if (enc === 'QUOTED-PRINTABLE') return decodeQuotedPrintable(prop.value, (prop.params.CHARSET || [])[0]);
  return prop.value;
}

/** Split on unescaped semicolons — the components of N and ADR. */
function components(value) {
  const out = [];
  let current = '';
  for (let i = 0; i < value.length; i++) {
    const ch = value[i];
    if (ch === '\\' && i + 1 < value.length) {
      current += ch + value[i + 1];
      i += 1;
    } else if (ch === ';') {
      out.push(current);
      current = '';
    } else current += ch;
  }
  out.push(current);
  return out.map(unescapeText);
}

const types = (prop) => (prop.params.TYPE || []).flatMap((t) => t.split(',')).map((t) => t.trim().toUpperCase()).filter(Boolean);
const isPref = (prop) => types(prop).includes('PREF') || (prop.params.PREF || []).some((v) => Number(v) === 1);
const kindOf = (prop, fallback) => {
  const t = types(prop).filter((x) => !['PREF', 'VOICE', 'INTERNET', 'X-400', 'PLAIN'].includes(x));
  return (t[0] || fallback).toLowerCase();
};

/**
 * The logical lines of a file. Standard folding (a leading space or tab)
 * joins a line to the one before; on a quoted-printable line, vCard 2.1's
 * soft break — a "=" at the end — continues the value on the next raw line,
 * and the "=" itself goes. Only on such a line: a base64 photo ends in "="
 * too, and joining after it swallowed whatever property came next.
 */
function logicalLines(text) {
  const raw = String(text || '').replace(/^﻿/, '').split(/\r\n|\n|\r/);
  const lines = [];
  for (let i = 0; i < raw.length; i++) {
    let line = raw[i];
    if (!line.length) continue;
    if ((line.startsWith(' ') || line.startsWith('\t')) && lines.length) {
      lines[lines.length - 1] += line.slice(1);
      continue;
    }
    while (/ENCODING=QUOTED-PRINTABLE/i.test(line) && line.endsWith('=') && i + 1 < raw.length) line = line.slice(0, -1) + raw[++i];
    lines.push(line);
  }
  return lines;
}

/* ── reading ───────────────────────────────────────────────────────────── */

const KNOWN = new Set(['VERSION', 'UID', 'FN', 'N', 'NICKNAME', 'ORG', 'TITLE', 'ROLE', 'EMAIL', 'TEL', 'ADR', 'LABEL', 'BDAY', 'ANNIVERSARY', 'NOTE', 'URL', 'PHOTO', 'CATEGORIES', 'REV', 'PRODID', 'X-ABLABEL', 'GENDER', 'KIND', 'IMPP']);

function readCard(props) {
  // Apple hangs labels on grouped properties: item1.URL and item1.X-ABLabel.
  const labels = new Map();
  for (const p of props) if (p.group && p.name === 'X-ABLABEL') labels.set(p.group, unescapeText(valueOf(p)).replace(/^_\$!<|>!\$_$/g, ''));
  const label = (p) => (p.group && labels.get(p.group)) || null;
  const first = (name) => props.find((p) => p.name === name) || null;
  const all = (name) => props.filter((p) => p.name === name);
  const textOf = (name) => {
    const p = first(name);
    return p ? unescapeText(valueOf(p)) : null;
  };

  const n = first('N') ? components(valueOf(first('N'))) : [];
  const [family = '', given = '', middle = '', prefix = '', suffix = ''] = n;
  const full = textOf('FN') || [prefix, given, middle, family, suffix].filter(Boolean).join(' ');
  const org = first('ORG') ? components(valueOf(first('ORG'))) : [];

  const photoProp = first('PHOTO');
  let photo = null;
  if (photoProp) {
    const v = valueOf(photoProp).trim();
    const enc = (photoProp.params.ENCODING || [])[0]?.toUpperCase();
    const dataUri = /^data:([^;,]+);base64,(.*)$/is.exec(v);
    if (dataUri) photo = { data: Buffer.from(dataUri[2].replace(/\s+/g, ''), 'base64'), mediaType: dataUri[1], uri: null };
    else if (enc === 'B' || enc === 'BASE64') photo = { data: Buffer.from(v.replace(/\s+/g, ''), 'base64'), mediaType: photoProp.params.TYPE ? `image/${photoProp.params.TYPE[0].toLowerCase()}` : (photoProp.params.MEDIATYPE || [])[0] || 'image/jpeg', uri: null };
    else photo = { data: null, mediaType: (photoProp.params.MEDIATYPE || [])[0] || null, uri: v };
  }

  const bday = textOf('BDAY');
  return {
    uid: textOf('UID'),
    version: textOf('VERSION') || '3.0',
    name: { full, given, family, middle, prefix, suffix },
    nickname: textOf('NICKNAME'),
    org: org[0] || null,
    department: org[1] || null,
    title: textOf('TITLE'),
    role: textOf('ROLE'),
    emails: all('EMAIL').map((p) => ({ value: unescapeText(valueOf(p)).trim(), type: kindOf(p, 'other'), pref: isPref(p), label: label(p) })),
    phones: all('TEL').map((p) => ({ value: unescapeText(valueOf(p)).replace(/^tel:/i, '').trim(), type: kindOf(p, 'other'), pref: isPref(p), label: label(p) })),
    addresses: all('ADR').map((p) => {
      const [pobox = '', extended = '', street = '', city = '', region = '', postcode = '', country = ''] = components(valueOf(p));
      return { street: [pobox, extended, street].filter(Boolean).join('\n'), city, region, postcode, country, type: kindOf(p, 'other'), pref: isPref(p), label: label(p) };
    }),
    birthday: normaliseDate(bday),
    anniversary: normaliseDate(textOf('ANNIVERSARY')),
    note: textOf('NOTE'),
    urls: all('URL').map((p) => ({ value: unescapeText(valueOf(p)).trim(), type: kindOf(p, 'other'), label: label(p) })),
    photo,
    categories: all('CATEGORIES').flatMap((p) => valueOf(p).split(',').map((s) => unescapeText(s).trim()).filter(Boolean)),
    rev: textOf('REV'),
    kind: (textOf('KIND') || 'individual').toLowerCase(),
    extra: props.filter((p) => !KNOWN.has(p.name) && !(p.group && p.name === 'X-ABLABEL')).map((p) => ({ name: p.name, params: p.params, value: p.value, group: p.group || null })),
  };
}

/** 1985-03-15, 19850315, --0315 or --03-15 → 1985-03-15 or --03-15. */
function normaliseDate(v) {
  if (!v) return null;
  const s = String(v).trim();
  let m = /^(\d{4})-?(\d{2})-?(\d{2})/.exec(s);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  m = /^--(\d{2})-?(\d{2})$/.exec(s);
  if (m) return `--${m[1]}-${m[2]}`;
  m = /^(\d{4})-?(\d{2})$/.exec(s);
  if (m) return `${m[1]}-${m[2]}`;
  return s;
}

/** Every card in a file. */
export function readVCards(textOrBytes) {
  const src = typeof textOrBytes === 'string' ? textOrBytes : new TextDecoder('utf-8').decode(textOrBytes);
  const cards = [];
  let current = null;
  for (const line of logicalLines(src)) {
    const prop = parseLine(line);
    // A group prefix: item1.EMAIL — the name is what follows the dot.
    const dot = prop.name.indexOf('.');
    if (dot > 0) {
      prop.group = prop.name.slice(0, dot).toLowerCase();
      prop.name = prop.name.slice(dot + 1);
    }
    if (prop.name === 'BEGIN' && /^VCARD$/i.test(prop.value.trim())) current = [];
    else if (prop.name === 'END' && /^VCARD$/i.test(prop.value.trim())) {
      if (current) cards.push(readCard(current));
      current = null;
    } else if (current) current.push(prop);
  }
  return cards;
}

/* ── writing ───────────────────────────────────────────────────────────── */

function line(name, params, value) {
  let head = name;
  for (const [k, vs] of Object.entries(params || {})) {
    if (!vs || !vs.length) continue;
    head += `;${k}=${vs.map((v) => (/[;:,]/.test(v) ? `"${v}"` : v)).join(',')}`;
  }
  return fold(`${head}:${value}`);
}

const typeParams = (entry, v4) => {
  const params = {};
  if (entry.type && entry.type !== 'other') params.TYPE = [entry.type.toLowerCase()];
  if (entry.pref) {
    if (v4) params.PREF = ['1'];
    else params.TYPE = [...(params.TYPE || []), 'pref'];
  }
  return params;
};

export function writeVCard(c, { version = '3.0' } = {}) {
  const v4 = version === '4.0';
  const out = ['BEGIN:VCARD', `VERSION:${version}`];
  const n = c.name || {};
  const full = n.full || [n.prefix, n.given, n.middle, n.family, n.suffix].filter(Boolean).join(' ') || c.org || '';
  out.push(line('FN', {}, escapeText(full)));
  out.push(line('N', {}, [n.family, n.given, n.middle, n.prefix, n.suffix].map((s) => escapeText(s || '')).join(';')));
  if (c.uid) out.push(line('UID', {}, c.uid));
  if (c.nickname) out.push(line('NICKNAME', {}, escapeText(c.nickname)));
  if (c.org || c.department) out.push(line('ORG', {}, [c.org, c.department].filter((s) => s != null).map(escapeText).join(';')));
  if (c.title) out.push(line('TITLE', {}, escapeText(c.title)));
  if (c.role) out.push(line('ROLE', {}, escapeText(c.role)));
  let group = 0;
  const grouped = (name, params, value, label) => {
    if (!label) return out.push(line(name, params, value));
    group += 1;
    out.push(line(`item${group}.${name}`, params, value));
    out.push(line(`item${group}.X-ABLabel`, {}, escapeText(label)));
    return out.length;
  };
  for (const e of c.emails || []) if (e.value) grouped('EMAIL', { ...typeParams(e, v4), ...(v4 ? {} : { TYPE: [...((typeParams(e, v4).TYPE) || []), 'INTERNET'] }) }, escapeText(e.value), e.label);
  for (const t of c.phones || []) if (t.value) grouped('TEL', typeParams(t, v4), v4 ? `tel:${t.value}` : escapeText(t.value), t.label);
  for (const a of c.addresses || []) {
    const parts = ['', '', a.street || '', a.city || '', a.region || '', a.postcode || '', a.country || ''].map((s) => escapeText(s));
    grouped('ADR', typeParams(a, v4), parts.join(';'), a.label);
  }
  if (c.birthday) out.push(line('BDAY', {}, v4 ? c.birthday.replace(/-/g, '').replace(/^(\d{4})(\d{2})(\d{2})$/, '$1$2$3') : c.birthday));
  if (c.anniversary) out.push(line('ANNIVERSARY', {}, c.anniversary));
  for (const u of c.urls || []) if (u.value) grouped('URL', typeParams(u, v4), u.value, u.label);
  if (c.note) out.push(line('NOTE', {}, escapeText(c.note)));
  if (c.categories?.length) out.push(line('CATEGORIES', {}, c.categories.map(escapeText).join(',')));
  if (c.photo?.data) {
    const type = (c.photo.mediaType || 'image/jpeg').toLowerCase();
    if (v4) out.push(line('PHOTO', {}, `data:${type};base64,${Buffer.from(c.photo.data).toString('base64')}`));
    else out.push(line('PHOTO', { ENCODING: ['b'], TYPE: [type.replace('image/', '').toUpperCase()] }, Buffer.from(c.photo.data).toString('base64')));
  } else if (c.photo?.uri) out.push(line('PHOTO', v4 ? {} : { VALUE: ['uri'] }, c.photo.uri));
  if (c.kind && c.kind !== 'individual' && v4) out.push(line('KIND', {}, c.kind));
  for (const x of c.extra || []) out.push(line(x.group ? `${x.group}.${x.name}` : x.name, x.params, x.value));
  out.push(line('REV', {}, new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d+/, '')));
  out.push('END:VCARD');
  return out.join('\r\n') + '\r\n';
}

export function writeVCards(contacts, options) {
  return contacts.map((c) => writeVCard(c, options)).join('');
}

/* ── names ─────────────────────────────────────────────────────────────── */

/** The name to show: the full name, else the given and family, else the organisation, else an address. */
export function displayName(c) {
  const n = c?.name || {};
  return n.full || [n.given, n.family].filter(Boolean).join(' ') || c?.org || c?.emails?.[0]?.value || '';
}

export function initials(c) {
  const n = c?.name || {};
  const parts = [n.given, n.family].filter(Boolean);
  const source = parts.length ? parts : displayName(c).split(/\s+/);
  return source.slice(0, 2).map((s) => s[0]?.toUpperCase() || '').join('');
}

/** Split a typed name into the parts a card keeps. */
export function splitName(full) {
  const words = String(full || '').trim().split(/\s+/).filter(Boolean);
  if (!words.length) return { full: '', given: '', family: '', middle: '', prefix: '', suffix: '' };
  const prefixes = new Set(['mr', 'mrs', 'ms', 'miss', 'dr', 'prof', 'sir', 'mr.', 'mrs.', 'ms.', 'dr.', 'prof.']);
  const prefix = prefixes.has(words[0].toLowerCase()) ? words.shift() : '';
  const given = words.shift() || '';
  const family = words.pop() || '';
  return { full: String(full).trim(), given, family, middle: words.join(' '), prefix, suffix: '' };
}

export default { readVCards, writeVCards, writeVCard, displayName, initials, splitName };
