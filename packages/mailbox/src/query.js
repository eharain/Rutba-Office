// The search box's language.
//
// The syntax people already know, from Gmail and from Outlook, rather than one
// of our own: words and "quoted phrases"; from:, to:, cc:, bcc:, subject:;
// filename: (attachment: in Outlook); has:attachment (hasattachments:yes);
// is:unread, is:read, is:starred (is:flagged); in: or folder: for where a
// message is filed, and account: (source:, archive:) for which mailbox or
// archive it is in; after:, before:, on:, older_than: and newer_than:, and
// Outlook's received: and sent: with >, >=, <, <= and a range a..b; size
// with size>, size<, size:>…, larger: and smaller:, in bytes, KB, MB or GB;
// a minus before anything to leave it out, and OR between two things to take
// either.
//
// Everything is case-insensitive. Anything that does not parse as one of the
// above is searched for as words, so a pasted address or a stray colon
// never makes a search fail.

const FIELD_ALIASES = {
  from: 'from',
  to: 'to',
  cc: 'cc',
  bcc: 'bcc',
  subject: 'subject',
  filename: 'filename',
  attachment: 'filename',
  attachments: 'filename',
  in: 'folder',
  folder: 'folder',
  label: 'folder',
  account: 'source',
  source: 'source',
  archive: 'source',
  body: 'body',
};

const DATE_FIELDS = new Set(['after', 'before', 'since', 'on', 'date', 'received', 'sent', 'older', 'newer']);

const UNITS = { b: 1, k: 1024, kb: 1024, m: 1024 ** 2, mb: 1024 ** 2, g: 1024 ** 3, gb: 1024 ** 3 };

/** `10mb`, `500K`, `2.5 MB`, `1000` → bytes, or null. */
export function parseSize(text) {
  const m = /^(\d+(?:\.\d+)?)\s*([kmg]?b?)$/i.exec(String(text || '').trim());
  if (!m) return null;
  const unit = UNITS[(m[2] || 'b').toLowerCase()] ?? 1;
  return Math.round(Number(m[1]) * unit);
}

/**
 * A date as the start and end of the span it names, in local time:
 * `2024` is the year, `2024-03` the month, `2024-03-11` (or 2024/03/11) the
 * day, and today and yesterday what they say.
 * @returns {{ from: number, to: number } | null} milliseconds, half-open
 */
export function parseDay(text, now = Date.now()) {
  const s = String(text || '').trim().toLowerCase();
  const day = (y, m, d) => ({ from: new Date(y, m, d).getTime(), to: new Date(y, m, d + 1).getTime() });
  if (s === 'today') {
    const n = new Date(now);
    return day(n.getFullYear(), n.getMonth(), n.getDate());
  }
  if (s === 'yesterday') {
    const n = new Date(now);
    return day(n.getFullYear(), n.getMonth(), n.getDate() - 1);
  }
  let m = /^(\d{4})$/.exec(s);
  if (m) return { from: new Date(+m[1], 0, 1).getTime(), to: new Date(+m[1] + 1, 0, 1).getTime() };
  m = /^(\d{4})[-/.](\d{1,2})$/.exec(s);
  if (m) return { from: new Date(+m[1], +m[2] - 1, 1).getTime(), to: new Date(+m[1], +m[2], 1).getTime() };
  m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/.exec(s);
  if (m) return day(+m[1], +m[2] - 1, +m[3]);
  // Day first, as most of the world writes it: 11/03/2024.
  m = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/.exec(s);
  if (m) return day(+m[3], +m[2] - 1, +m[1]);
  return null;
}

/** `3d`, `2w`, `6m`, `1y` before now, as Gmail's older_than and newer_than take. */
function relative(text, now) {
  const m = /^(\d+)\s*([dwmy])$/i.exec(String(text || '').trim());
  if (!m) return null;
  const n = Number(m[1]);
  const d = new Date(now);
  const unit = m[2].toLowerCase();
  if (unit === 'd') d.setDate(d.getDate() - n);
  else if (unit === 'w') d.setDate(d.getDate() - 7 * n);
  else if (unit === 'm') d.setMonth(d.getMonth() - n);
  else d.setFullYear(d.getFullYear() - n);
  return d.getTime();
}

/** A date condition - `>=2024-01-01`, `2024-01..2024-03`, `2024-03-11` - as a half-open span. */
function dateSpan(field, raw, now) {
  const value = String(raw || '').trim();
  if (field === 'older' || field === 'newer') {
    const at = relative(value, now);
    if (at == null) return null;
    return field === 'older' ? { from: null, to: at } : { from: at, to: null };
  }
  const range = /^(.+?)\.\.(.+)$/.exec(value);
  if (range) {
    const a = parseDay(range[1], now);
    const b = parseDay(range[2], now);
    if (!a || !b) return null;
    return { from: a.from, to: b.to };
  }
  const cmp = /^(>=|<=|>|<|=)?(.+)$/.exec(value);
  const op = cmp[1] || '';
  const span = parseDay(cmp[2], now);
  if (!span) return null;
  if (field === 'after' || field === 'since') return { from: span.from, to: null };
  if (field === 'before') return { from: null, to: span.from };
  if (op === '>') return { from: span.to, to: null };
  if (op === '>=') return { from: span.from, to: null };
  if (op === '<') return { from: null, to: span.from };
  if (op === '<=') return { from: null, to: span.to };
  return span;
}

/** A size condition - `>1mb`, `<=500k`, `1mb` (larger than, as Gmail reads it). */
function sizeSpan(field, raw) {
  const m = /^(>=|<=|>|<|=)?(.+)$/.exec(String(raw || '').trim());
  const bytes = parseSize(m[2]);
  if (bytes == null) return null;
  const op = m[1] || (field === 'smaller' ? '<' : '>');
  if (field === 'larger' || op === '>') return { min: bytes + (field === 'larger' ? 0 : 1), max: null };
  if (op === '>=') return { min: bytes, max: null };
  if (field === 'smaller' || op === '<') return { min: null, max: bytes - 1 };
  if (op === '<=') return { min: null, max: bytes };
  return { min: bytes, max: bytes };
}

/** Split the box into pieces, keeping quoted runs whole. */
function pieces(query) {
  const out = [];
  const re = /(-)?(?:([a-z_]+)(:|>=|<=|>|<))?(?:"([^"]*)"?|(\S+))/gi;
  let m;
  const text = String(query || '');
  while ((m = re.exec(text))) {
    if (!m[0]) {
      re.lastIndex++;
      continue;
    }
    const [, minus, field, sep, quoted, bare] = m;
    out.push({ negate: Boolean(minus), field: field ? field.toLowerCase() : null, sep: sep || null, quoted: quoted !== undefined, value: quoted ?? bare ?? '' });
  }
  return out;
}

/** The words worth matching in a piece of text - the same split the index makes. */
export function words(text) {
  const out = [];
  for (const raw of String(text || '').toLowerCase().split(/[^\p{L}\p{N}@._+-]+/u)) {
    const token = raw.replace(/^[._-]+|[._-]+$/g, '');
    if (token.length >= 2) out.push(token);
  }
  return out;
}

const norm = (s) => String(s || '').toLowerCase().replace(/\s+/g, ' ').trim();

/** One piece of the query as an atom, or null when it says nothing. */
function atomOf(p, now) {
  const value = p.value;
  if (p.field) {
    const field = p.field;
    const sep = p.sep;
    // size>1mb, size:>1mb, larger:5m, smaller:1m
    if (field === 'size' || field === 'larger' || field === 'smaller') {
      const span = sizeSpan(field, sep === ':' ? value : `${sep}${value}`);
      return span ? { kind: 'size', ...span } : null;
    }
    if (DATE_FIELDS.has(field) || field === 'older_than' || field === 'newer_than') {
      const f = field === 'older_than' ? 'older' : field === 'newer_than' ? 'newer' : field;
      const span = dateSpan(f, sep === ':' ? value : `${sep}${value}`, now);
      return span ? { kind: 'date', ...span } : null;
    }
    if (sep === ':') {
      if (field === 'has') {
        const v = value.toLowerCase();
        if (/^attach/.test(v)) return { kind: 'has', value: 'attachment' };
        return null;
      }
      if (field === 'hasattachment' || field === 'hasattachments') {
        return /^(yes|true|1)$/i.test(value) ? { kind: 'has', value: 'attachment' } : { kind: 'has', value: 'attachment', negateSelf: true };
      }
      if (field === 'is') {
        const v = value.toLowerCase();
        if (['unread', 'read', 'starred', 'flagged', 'unstarred'].includes(v)) return { kind: 'is', value: v === 'flagged' ? 'starred' : v };
        return null;
      }
      if (field === 'isread') return { kind: 'is', value: /^(yes|true|1)$/i.test(value) ? 'read' : 'unread' };
      if (FIELD_ALIASES[field]) {
        const f = FIELD_ALIASES[field];
        if (f === 'body') return p.quoted ? { kind: 'phrase', value: norm(value), words: words(value) } : { kind: 'term', value: value.toLowerCase() };
        return { kind: 'field', field: f, value: norm(value) };
      }
    }
    // Not a field this knows - searched for as it was typed.
    return atomOf({ ...p, field: null, value: `${p.field}${sep}${value}` }, now);
  }
  if (p.quoted) {
    const w = words(value);
    return w.length ? { kind: 'phrase', value: norm(value), words: w } : null;
  }
  const w = words(value);
  if (!w.length) return null;
  // A pasted token the index splits ("Q3/2024") becomes a phrase of its words.
  if (w.length > 1 || w[0] !== value.toLowerCase().replace(/^[._-]+|[._-]+$/g, '')) {
    return w.length === 1 ? { kind: 'term', value: w[0] } : { kind: 'phrase', value: norm(w.join(' ')), words: w, loose: true };
  }
  return { kind: 'term', value: w[0] };
}

/**
 * The query as clauses that must all hold, each a list of atoms any one of
 * which will do (OR), each perhaps negated.
 * @param {string} query
 * @param {{ now?: number }} [opts]
 * @returns {{ clauses: Array<{ any: object[], negate: boolean }>, empty: boolean }}
 */
export function parseSearch(query, { now = Date.now() } = {}) {
  const clauses = [];
  let joinNext = false;
  for (const p of pieces(query)) {
    if (!p.field && !p.quoted && !p.negate && p.value === 'OR') {
      joinNext = clauses.length > 0;
      continue;
    }
    if (!p.field && !p.quoted && (p.value === '(' || p.value === ')')) continue;
    const atom = atomOf({ ...p, value: p.quoted ? p.value : p.value.replace(/^\(+|\)+$/g, '') }, now);
    if (!atom) continue;
    let negate = p.negate;
    if (atom.negateSelf) {
      negate = !negate;
      delete atom.negateSelf;
    }
    if (joinNext && !negate && !clauses[clauses.length - 1].negate) clauses[clauses.length - 1].any.push(atom);
    else clauses.push({ any: [atom], negate });
    joinNext = false;
  }
  return { clauses, empty: clauses.length === 0 };
}
