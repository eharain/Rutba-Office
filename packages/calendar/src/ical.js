// iCalendar (RFC 5545), read and written.
//
// The format is lines of NAME;PARAM=VALUE:VALUE, folded at 75 octets, nested
// in BEGIN/END components. Everything a calendar program writes is here —
// Google, Outlook, Apple, Thunderbird, and the invitations mail carries —
// and the differences between them are in the details this file is careful
// about: a folded line broken inside a multi-byte character, a parameter
// value in quotes because it holds a comma, a TZID that is a Windows name, a
// DTEND missing because a DURATION stands in for it, an all-day event whose
// end is the day after.
//
// What is read becomes a plain event object; what is written comes from one.
// Properties this model does not name are kept on the event as `extra` and
// written back, so a file that came in leaves whole.

import { localToUtc, utcToLocal, ianaZone } from './tz.js';

/* ── lines ─────────────────────────────────────────────────────────────── */

/** The logical lines of a file: continuation lines (a leading space or tab) joined to the line before. */
export function unfold(text) {
  const raw = String(text || '').replace(/^\uFEFF/, '').split(/\r\n|\n|\r/);
  const lines = [];
  for (const line of raw) {
    if ((line.startsWith(' ') || line.startsWith('\t')) && lines.length) lines[lines.length - 1] += line.slice(1);
    else if (line.length) lines.push(line);
  }
  return lines;
}

/** Fold a line at 75 octets, never inside a character. */
export function fold(line) {
  const out = [];
  let current = '';
  let bytes = 0;
  for (const ch of line) {
    const size = Buffer.byteLength(ch, 'utf8');
    if (bytes + size > (out.length ? 74 : 75)) {
      out.push(current);
      current = ch;
      bytes = size;
    } else {
      current += ch;
      bytes += size;
    }
  }
  out.push(current);
  return out.map((s, i) => (i ? ' ' + s : s)).join('\r\n');
}

/** NAME;PARAM=VALUE;OTHER="a,b":value — the name and parameters upper-cased, quoted values kept whole. */
export function parseLine(line) {
  let i = 0;
  let inQuotes = false;
  for (; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') inQuotes = !inQuotes;
    else if (ch === ':' && !inQuotes) break;
  }
  const head = line.slice(0, i);
  const value = line.slice(i + 1);
  const parts = [];
  let current = '';
  inQuotes = false;
  for (const ch of head) {
    if (ch === '"') inQuotes = !inQuotes;
    if (ch === ';' && !inQuotes) {
      parts.push(current);
      current = '';
    } else current += ch;
  }
  parts.push(current);
  const name = parts[0].toUpperCase();
  const params = {};
  for (const p of parts.slice(1)) {
    const eq = p.indexOf('=');
    if (eq < 0) {
      // vCard 2.1 writes a bare parameter — TEL;HOME;VOICE — which means TYPE.
      if (p.trim()) params.TYPE = (params.TYPE || []).concat(p.trim().toUpperCase());
      continue;
    }
    const key = p.slice(0, eq).toUpperCase();
    const values = [];
    let v = '';
    let q = false;
    for (const ch of p.slice(eq + 1)) {
      if (ch === '"') q = !q;
      else if (ch === ',' && !q) {
        values.push(v);
        v = '';
      } else v += ch;
    }
    values.push(v);
    params[key] = (params[key] || []).concat(values);
  }
  return { name, params, value };
}

export function unescapeText(v) {
  return String(v ?? '').replace(/\\([\\;,nN])/g, (_, c) => (c === 'n' || c === 'N' ? '\n' : c));
}

export function escapeText(v) {
  return String(v ?? '').replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\r?\n/g, '\\n');
}

/* ── components ────────────────────────────────────────────────────────── */

/** The BEGIN/END tree. */
export function parseComponents(text) {
  const root = { name: 'ROOT', props: [], children: [] };
  const stack = [root];
  for (const line of unfold(text)) {
    const prop = parseLine(line);
    if (prop.name === 'BEGIN') {
      const node = { name: prop.value.trim().toUpperCase(), props: [], children: [] };
      stack[stack.length - 1].children.push(node);
      stack.push(node);
    } else if (prop.name === 'END') {
      if (stack.length > 1) stack.pop();
    } else stack[stack.length - 1].props.push(prop);
  }
  return root.children;
}

const propOf = (node, name) => node.props.find((p) => p.name === name) || null;
const propsOf = (node, name) => node.props.filter((p) => p.name === name);
const text = (node, name) => {
  const p = propOf(node, name);
  return p ? unescapeText(p.value) : null;
};

/* ── dates ─────────────────────────────────────────────────────────────── */

/**
 * A date or date-time value with its parameters, as the model keeps it:
 * `{ date, time, tzid, utc, allDay, at }` — the parts as written, the zone
 * they are in, and `at`, the instant in milliseconds (for an all-day value,
 * midnight UTC of the calendar date, which consumers treat as a date).
 */
export function parseDateTime(value, params = {}) {
  const raw = String(value || '').trim();
  const m = /^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})?(Z)?)?$/.exec(raw);
  if (!m) return null;
  const [, y, mo, d, H, M, S, z] = m;
  const isDate = (params.VALUE || [])[0] === 'DATE' || H === undefined;
  const parts = { y: Number(y), m: Number(mo), d: Number(d), H: Number(H || 0), M: Number(M || 0), S: Number(S || 0) };
  if (isDate) {
    return { date: `${y}-${mo}-${d}`, time: null, tzid: null, utc: false, allDay: true, at: Date.UTC(parts.y, parts.m - 1, parts.d) };
  }
  const tzid = (params.TZID || [])[0] || null;
  const utc = Boolean(z);
  const at = utc ? Date.UTC(parts.y, parts.m - 1, parts.d, parts.H, parts.M, parts.S) : localToUtc(parts, tzid);
  return { date: `${y}-${mo}-${d}`, time: `${H}:${M}:${S || '00'}`, tzid, utc, allDay: false, at };
}

/** A DURATION such as P1DT2H30M or -PT15M, in milliseconds. */
export function parseDuration(value) {
  const m = /^([+-])?P(?:(\d+)W)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$/.exec(String(value || '').trim());
  if (!m) return null;
  const [, sign, w, d, h, mi, s] = m;
  const ms = ((Number(w || 0) * 7 + Number(d || 0)) * 86400 + Number(h || 0) * 3600 + Number(mi || 0) * 60 + Number(s || 0)) * 1000;
  return sign === '-' ? -ms : ms;
}

export function formatDuration(ms) {
  const sign = ms < 0 ? '-' : '';
  let rest = Math.abs(Math.round(ms / 1000));
  const d = Math.floor(rest / 86400);
  rest -= d * 86400;
  const h = Math.floor(rest / 3600);
  rest -= h * 3600;
  const m = Math.floor(rest / 60);
  const s = rest - m * 60;
  let out = `${sign}P`;
  if (d) out += `${d}D`;
  if (h || m || s || !d) out += `T${h ? `${h}H` : ''}${m ? `${m}M` : ''}${s || (!h && !m) ? `${s}S` : ''}`;
  return out;
}

const pad = (n, w = 2) => String(n).padStart(w, '0');

/** A date-time value for the file: DATE, UTC, or wall-clock with a TZID parameter. */
export function formatDateTime(dt) {
  if (!dt) return null;
  if (dt.allDay) return { params: { VALUE: ['DATE'] }, value: dt.date.replace(/-/g, '') };
  const [y, m, d] = dt.date.split('-');
  const [H, M, S] = (dt.time || '00:00:00').split(':');
  const stamp = `${y}${m}${d}T${pad(H)}${pad(M)}${pad(S || '00')}`;
  if (dt.utc) return { params: {}, value: `${stamp}Z` };
  if (dt.tzid) return { params: { TZID: [dt.tzid] }, value: stamp };
  return { params: {}, value: stamp };
}

/** A date-time from an instant, in a zone (or UTC when none). */
export function dateTimeAt(epochMs, tzid = null, allDay = false) {
  if (allDay) {
    const d = new Date(epochMs);
    return { date: `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`, time: null, tzid: null, utc: false, allDay: true, at: Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()) };
  }
  const zone = tzid ? ianaZone(tzid) : null;
  if (!zone) {
    const d = new Date(epochMs);
    return { date: `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`, time: `${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}:${pad(d.getUTCSeconds())}`, tzid: null, utc: true, allDay: false, at: epochMs };
  }
  const p = utcToLocal(epochMs, zone);
  return { date: `${p.y}-${pad(p.m)}-${pad(p.d)}`, time: `${pad(p.H)}:${pad(p.M)}:${pad(p.S)}`, tzid, utc: false, allDay: false, at: epochMs };
}

/* ── recurrence rules ──────────────────────────────────────────────────── */

const DAYS = ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'];

export function parseRRule(value) {
  const rule = { freq: null, interval: 1, count: null, until: null, byDay: null, byMonthDay: null, byMonth: null, bySetPos: null, wkst: 'MO' };
  for (const part of String(value || '').split(';')) {
    const [k, v] = part.split('=');
    if (!k || v === undefined) continue;
    switch (k.toUpperCase()) {
      case 'FREQ': rule.freq = v.toUpperCase(); break;
      case 'INTERVAL': rule.interval = Math.max(1, Number(v) || 1); break;
      case 'COUNT': rule.count = Number(v) || null; break;
      case 'UNTIL': rule.until = parseDateTime(v, /T/.test(v) ? {} : { VALUE: ['DATE'] }); break;
      case 'BYDAY':
        rule.byDay = v.split(',').map((s) => {
          const m = /^([+-]?\d+)?(SU|MO|TU|WE|TH|FR|SA)$/i.exec(s.trim());
          return m ? { n: m[1] ? Number(m[1]) : 0, day: DAYS.indexOf(m[2].toUpperCase()) } : null;
        }).filter(Boolean);
        break;
      case 'BYMONTHDAY': rule.byMonthDay = v.split(',').map(Number).filter((n) => Number.isFinite(n) && n !== 0); break;
      case 'BYMONTH': rule.byMonth = v.split(',').map(Number).filter((n) => n >= 1 && n <= 12); break;
      case 'BYSETPOS': rule.bySetPos = v.split(',').map(Number).filter((n) => Number.isFinite(n) && n !== 0); break;
      case 'WKST': rule.wkst = v.toUpperCase(); break;
      default: break;
    }
  }
  return rule.freq ? rule : null;
}

export function formatRRule(rule) {
  if (!rule?.freq) return null;
  const parts = [`FREQ=${rule.freq}`];
  if (rule.interval && rule.interval !== 1) parts.push(`INTERVAL=${rule.interval}`);
  if (rule.count) parts.push(`COUNT=${rule.count}`);
  if (rule.until) parts.push(`UNTIL=${rule.until.allDay ? rule.until.date.replace(/-/g, '') : formatDateTime({ ...rule.until, utc: true, tzid: null }).value}`);
  if (rule.byDay?.length) parts.push(`BYDAY=${rule.byDay.map((b) => `${b.n ? b.n : ''}${DAYS[b.day]}`).join(',')}`);
  if (rule.byMonthDay?.length) parts.push(`BYMONTHDAY=${rule.byMonthDay.join(',')}`);
  if (rule.byMonth?.length) parts.push(`BYMONTH=${rule.byMonth.join(',')}`);
  if (rule.bySetPos?.length) parts.push(`BYSETPOS=${rule.bySetPos.join(',')}`);
  if (rule.wkst && rule.wkst !== 'MO') parts.push(`WKST=${rule.wkst}`);
  return parts.join(';');
}

/* ── people ────────────────────────────────────────────────────────────── */

function parsePerson(prop) {
  if (!prop) return null;
  const email = String(prop.value || '').replace(/^mailto:/i, '').trim() || null;
  const p = prop.params;
  return {
    email,
    name: (p.CN || [])[0] || null,
    role: (p.ROLE || [])[0] || null,
    partstat: (p.PARTSTAT || [])[0] || null,
    rsvp: /true/i.test((p.RSVP || [])[0] || ''),
    type: (p.CUTYPE || [])[0] || null,
  };
}

function personParams(person, { rsvp = false } = {}) {
  const params = {};
  if (person.name) params.CN = [person.name];
  if (person.role) params.ROLE = [person.role];
  if (person.partstat) params.PARTSTAT = [person.partstat];
  if (rsvp || person.rsvp) params.RSVP = ['TRUE'];
  if (person.type) params.CUTYPE = [person.type];
  return params;
}

/* ── events ────────────────────────────────────────────────────────────── */

const KNOWN = new Set(['UID', 'SUMMARY', 'DESCRIPTION', 'LOCATION', 'URL', 'STATUS', 'TRANSP', 'SEQUENCE', 'DTSTAMP', 'CREATED', 'LAST-MODIFIED', 'DTSTART', 'DTEND', 'DURATION', 'RRULE', 'EXDATE', 'RDATE', 'RECURRENCE-ID', 'ORGANIZER', 'ATTENDEE', 'CATEGORIES', 'CLASS', 'PRIORITY', 'GEO', 'COLOR']);

function readAlarm(node) {
  const trigger = propOf(node, 'TRIGGER');
  const action = text(node, 'ACTION') || 'DISPLAY';
  if (!trigger) return null;
  const isAbsolute = (trigger.params.VALUE || [])[0] === 'DATE-TIME';
  return {
    action,
    description: text(node, 'DESCRIPTION'),
    // Minutes before the start (positive), or an absolute instant.
    before: isAbsolute ? null : -Math.round((parseDuration(trigger.value) ?? 0) / 60000),
    at: isAbsolute ? parseDateTime(trigger.value)?.at ?? null : null,
    related: (trigger.params.RELATED || [])[0] || 'START',
  };
}

function readEvent(node) {
  const start = propOf(node, 'DTSTART');
  const dtstart = start ? parseDateTime(start.value, start.params) : null;
  const endProp = propOf(node, 'DTEND');
  let dtend = endProp ? parseDateTime(endProp.value, endProp.params) : null;
  const durationProp = propOf(node, 'DURATION');
  const duration = durationProp ? parseDuration(durationProp.value) : null;
  if (!dtend && dtstart) {
    if (duration != null) dtend = dateTimeAt(dtstart.at + duration, dtstart.tzid, dtstart.allDay);
    else dtend = dtstart.allDay ? dateTimeAt(dtstart.at + 86400000, null, true) : { ...dtstart };
  }
  const recurrenceProp = propOf(node, 'RECURRENCE-ID');
  const event = {
    uid: text(node, 'UID'),
    summary: text(node, 'SUMMARY') || '',
    description: text(node, 'DESCRIPTION') || '',
    location: text(node, 'LOCATION') || '',
    url: text(node, 'URL') || null,
    status: (text(node, 'STATUS') || '').toUpperCase() || null,
    transparency: (text(node, 'TRANSP') || '').toUpperCase() || null,
    sequence: Number(text(node, 'SEQUENCE') || 0) || 0,
    dtstamp: propOf(node, 'DTSTAMP') ? parseDateTime(propOf(node, 'DTSTAMP').value, propOf(node, 'DTSTAMP').params) : null,
    created: propOf(node, 'CREATED') ? parseDateTime(propOf(node, 'CREATED').value, propOf(node, 'CREATED').params) : null,
    modified: propOf(node, 'LAST-MODIFIED') ? parseDateTime(propOf(node, 'LAST-MODIFIED').value, propOf(node, 'LAST-MODIFIED').params) : null,
    start: dtstart,
    end: dtend,
    allDay: Boolean(dtstart?.allDay),
    rrule: propOf(node, 'RRULE') ? parseRRule(propOf(node, 'RRULE').value) : null,
    exdates: propsOf(node, 'EXDATE').flatMap((p) => p.value.split(',').map((v) => parseDateTime(v, p.params)).filter(Boolean)),
    rdates: propsOf(node, 'RDATE').flatMap((p) => p.value.split(',').map((v) => parseDateTime(v, p.params)).filter(Boolean)),
    recurrenceId: recurrenceProp ? parseDateTime(recurrenceProp.value, recurrenceProp.params) : null,
    organizer: parsePerson(propOf(node, 'ORGANIZER')),
    attendees: propsOf(node, 'ATTENDEE').map(parsePerson).filter(Boolean),
    categories: propsOf(node, 'CATEGORIES').flatMap((p) => p.value.split(',').map((s) => unescapeText(s).trim()).filter(Boolean)),
    privacy: (text(node, 'CLASS') || '').toUpperCase() || null,
    priority: Number(text(node, 'PRIORITY') || 0) || 0,
    colour: text(node, 'COLOR') || null,
    alarms: node.children.filter((c) => c.name === 'VALARM').map(readAlarm).filter(Boolean),
    extra: node.props.filter((p) => !KNOWN.has(p.name)).map((p) => ({ name: p.name, params: p.params, value: p.value })),
  };
  return event;
}

/**
 * Read a calendar. Answers with its events, the method when it is an
 * invitation or a reply, and what wrote it.
 */
export function readCalendar(textOrBytes) {
  const src = typeof textOrBytes === 'string' ? textOrBytes : new TextDecoder('utf-8').decode(textOrBytes);
  const roots = parseComponents(src);
  const cal = roots.find((c) => c.name === 'VCALENDAR') || { name: 'VCALENDAR', props: [], children: roots };
  const events = cal.children.filter((c) => c.name === 'VEVENT').map(readEvent);
  return {
    prodid: text(cal, 'PRODID'),
    version: text(cal, 'VERSION'),
    method: (text(cal, 'METHOD') || '').toUpperCase() || null,
    name: text(cal, 'X-WR-CALNAME') || text(cal, 'NAME') || null,
    timezone: text(cal, 'X-WR-TIMEZONE') || null,
    timezones: cal.children.filter((c) => c.name === 'VTIMEZONE').map((z) => text(z, 'TZID')).filter(Boolean),
    events,
  };
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

function dateLine(name, dt) {
  const f = formatDateTime(dt);
  return f ? line(name, f.params, f.value) : null;
}

const stamp = (ms = Date.now()) => formatDateTime(dateTimeAt(ms, null)).value;

export function writeEvent(event, out) {
  out.push('BEGIN:VEVENT');
  out.push(line('UID', {}, event.uid || `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}@rutba.io`));
  out.push(line('DTSTAMP', {}, event.dtstamp ? formatDateTime({ ...event.dtstamp, utc: true, tzid: null }).value : stamp()));
  if (event.start) out.push(dateLine('DTSTART', event.start));
  if (event.end) {
    if (event.allDay && event.start && event.end.at - event.start.at === 86400000 && !event.explicitEnd) out.push(dateLine('DTEND', event.end));
    else out.push(dateLine('DTEND', event.end));
  }
  if (event.summary) out.push(line('SUMMARY', {}, escapeText(event.summary)));
  if (event.description) out.push(line('DESCRIPTION', {}, escapeText(event.description)));
  if (event.location) out.push(line('LOCATION', {}, escapeText(event.location)));
  if (event.url) out.push(line('URL', {}, event.url));
  if (event.status) out.push(line('STATUS', {}, event.status));
  if (event.transparency) out.push(line('TRANSP', {}, event.transparency));
  if (event.sequence) out.push(line('SEQUENCE', {}, String(event.sequence)));
  if (event.created) out.push(line('CREATED', {}, formatDateTime({ ...event.created, utc: true, tzid: null }).value));
  if (event.modified) out.push(line('LAST-MODIFIED', {}, formatDateTime({ ...event.modified, utc: true, tzid: null }).value));
  if (event.rrule) out.push(line('RRULE', {}, formatRRule(event.rrule)));
  for (const x of event.exdates || []) out.push(dateLine('EXDATE', x));
  for (const x of event.rdates || []) out.push(dateLine('RDATE', x));
  if (event.recurrenceId) out.push(dateLine('RECURRENCE-ID', event.recurrenceId));
  if (event.organizer?.email) out.push(line('ORGANIZER', personParams(event.organizer), `mailto:${event.organizer.email}`));
  for (const a of event.attendees || []) if (a.email) out.push(line('ATTENDEE', personParams(a), `mailto:${a.email}`));
  if (event.categories?.length) out.push(line('CATEGORIES', {}, event.categories.map(escapeText).join(',')));
  if (event.privacy) out.push(line('CLASS', {}, event.privacy));
  if (event.priority) out.push(line('PRIORITY', {}, String(event.priority)));
  if (event.colour) out.push(line('COLOR', {}, event.colour));
  for (const x of event.extra || []) out.push(line(x.name, x.params, x.value));
  for (const alarm of event.alarms || []) {
    out.push('BEGIN:VALARM');
    out.push(line('ACTION', {}, alarm.action || 'DISPLAY'));
    if (alarm.at != null) out.push(line('TRIGGER', { VALUE: ['DATE-TIME'] }, formatDateTime(dateTimeAt(alarm.at, null)).value));
    else out.push(line('TRIGGER', alarm.related && alarm.related !== 'START' ? { RELATED: [alarm.related] } : {}, formatDuration(-(alarm.before || 0) * 60000)));
    out.push(line('DESCRIPTION', {}, escapeText(alarm.description || event.summary || 'Reminder')));
    out.push('END:VALARM');
  }
  out.push('END:VEVENT');
}

/**
 * Write a calendar. `method` makes it an invitation (REQUEST), an answer
 * (REPLY) or a cancellation (CANCEL); without one it is a plain calendar.
 */
export function writeCalendar({ events = [], method = null, name = null, prodid = '-//Rutba//Rutba Office//EN' } = {}) {
  const out = ['BEGIN:VCALENDAR', line('PRODID', {}, prodid), 'VERSION:2.0', 'CALSCALE:GREGORIAN'];
  if (method) out.push(line('METHOD', {}, method));
  if (name) out.push(line('X-WR-CALNAME', {}, escapeText(name)));
  for (const e of events) writeEvent(e, out);
  out.push('END:VCALENDAR');
  return out.filter(Boolean).join('\r\n') + '\r\n';
}

/* ── invitations ───────────────────────────────────────────────────────── */

/**
 * The answer an attendee sends back: the same UID and sequence, this one
 * attendee with the answer on it, and METHOD:REPLY — which is all the
 * organizer's calendar reads.
 */
export function writeReply(event, { email, name = null, partstat = 'ACCEPTED' }) {
  const reply = {
    uid: event.uid,
    sequence: event.sequence || 0,
    dtstamp: null,
    start: event.start,
    end: event.end,
    allDay: event.allDay,
    summary: event.summary,
    organizer: event.organizer,
    attendees: [{ email, name, partstat: partstat.toUpperCase() }],
    recurrenceId: event.recurrenceId || null,
    extra: [],
    alarms: [],
  };
  return writeCalendar({ events: [reply], method: 'REPLY' });
}

/** Whether a calendar is something to answer — an invitation with attendees. */
export function isInvitation(calendar) {
  return calendar?.method === 'REQUEST' && calendar.events.some((e) => e.attendees.length);
}

export default { readCalendar, writeCalendar, writeEvent, writeReply, isInvitation, parseDateTime, parseDuration, formatDuration, parseRRule, formatRRule, unfold, fold, parseLine, escapeText, unescapeText, dateTimeAt };
