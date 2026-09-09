// Recurrence, expanded.
//
// A recurring event is one VEVENT with a rule; what a calendar shows is its
// occurrences in the window on screen. This turns the one into the many:
// daily, weekly, monthly and yearly rules with INTERVAL, COUNT and UNTIL,
// BYDAY with and without ordinals, BYMONTHDAY, BYMONTH and BYSETPOS —
// the rules every calendar program writes — with EXDATE removing instances,
// RDATE adding them, and a VEVENT that carries a RECURRENCE-ID replacing
// the instance it names, which is how "just this Tuesday moved to 3pm" is
// stored.
//
// The arithmetic is done on the event's own wall clock — a 9am meeting in
// London stays a 9am meeting across the clock change — and each occurrence
// is then turned into an instant in the event's zone.

import { localToUtc, ianaZone } from './tz.js';

const DAY_MS = 86400000;
const DAYS = ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'];

function partsOf(dt) {
  const [y, m, d] = dt.date.split('-').map(Number);
  const [H, M, S] = (dt.time || '00:00:00').split(':').map(Number);
  return { y, m, d, H, M: M || 0, S: S || 0 };
}

/** A calendar date as a day number, for arithmetic that ignores the clock. */
const dayNumber = ({ y, m, d }) => Math.round(Date.UTC(y, m - 1, d) / DAY_MS);
const fromDayNumber = (n) => {
  const date = new Date(n * DAY_MS);
  return { y: date.getUTCFullYear(), m: date.getUTCMonth() + 1, d: date.getUTCDate() };
};
const weekday = (p) => new Date(Date.UTC(p.y, p.m - 1, p.d)).getUTCDay();
const daysInMonth = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();

/** The instant of wall-clock parts in the event's zone (or as an all-day date). */
function instantOf(p, start) {
  if (start.allDay) return Date.UTC(p.y, p.m - 1, p.d);
  if (start.utc) return Date.UTC(p.y, p.m - 1, p.d, p.H, p.M, p.S);
  return localToUtc(p, start.tzid);
}

/** The days of one month that a rule's BYDAY picks out, in order. */
function byDayInMonth(y, m, byDay) {
  const days = [];
  const count = daysInMonth(y, m);
  for (const { n, day } of byDay) {
    const matches = [];
    for (let d = 1; d <= count; d++) if (weekday({ y, m, d }) === day) matches.push(d);
    if (!n) days.push(...matches);
    else if (n > 0 && matches[n - 1]) days.push(matches[n - 1]);
    else if (n < 0 && matches[matches.length + n]) days.push(matches[matches.length + n]);
  }
  return [...new Set(days)].sort((a, b) => a - b);
}

/** The days of a month a rule picks: BYMONTHDAY, BYDAY, or the start's own day. */
function daysOfMonth(y, m, rule, startDay) {
  const count = daysInMonth(y, m);
  let days;
  if (rule.byMonthDay?.length) days = rule.byMonthDay.map((n) => (n > 0 ? n : count + 1 + n)).filter((d) => d >= 1 && d <= count);
  else if (rule.byDay?.length) days = byDayInMonth(y, m, rule.byDay);
  else days = startDay <= count ? [startDay] : [];
  days = [...new Set(days)].sort((a, b) => a - b);
  if (rule.bySetPos?.length) days = rule.bySetPos.map((n) => (n > 0 ? days[n - 1] : days[days.length + n])).filter((d) => d != null).sort((a, b) => a - b);
  return days;
}

/**
 * The wall-clock starts of an event's occurrences, in order, from the
 * rule alone — before exceptions and additions.
 */
function* ruleStarts(start, rule, hardLimit) {
  const first = partsOf(start);
  const interval = rule.interval || 1;
  let produced = 0;
  const wkst = DAYS.indexOf(rule.wkst || 'MO');
  const emit = function* (p) {
    produced += 1;
    yield p;
  };
  if (rule.freq === 'DAILY') {
    const startNumber = dayNumber(first);
    for (let i = 0; produced < hardLimit; i++) {
      const p = { ...fromDayNumber(startNumber + i * interval), H: first.H, M: first.M, S: first.S };
      if (rule.byMonth?.length && !rule.byMonth.includes(p.m)) continue;
      if (rule.byDay?.length && !rule.byDay.some((b) => b.day === weekday(p))) continue;
      yield* emit(p);
    }
  } else if (rule.freq === 'WEEKLY') {
    const days = rule.byDay?.length ? rule.byDay.map((b) => b.day) : [weekday(first)];
    // The week the start falls in, beginning on WKST.
    const startNumber = dayNumber(first);
    const back = (weekday(first) - wkst + 7) % 7;
    const weekStart = startNumber - back;
    for (let w = 0; produced < hardLimit; w++) {
      const base = weekStart + w * interval * 7;
      for (let i = 0; i < 7; i++) {
        const n = base + i;
        if (n < startNumber) continue;
        const p = { ...fromDayNumber(n), H: first.H, M: first.M, S: first.S };
        if (!days.includes(weekday(p))) continue;
        if (rule.byMonth?.length && !rule.byMonth.includes(p.m)) continue;
        yield* emit(p);
        if (produced >= hardLimit) return;
      }
    }
  } else if (rule.freq === 'MONTHLY') {
    for (let k = 0; produced < hardLimit; k++) {
      const total = (first.y * 12 + first.m - 1) + k * interval;
      const y = Math.floor(total / 12);
      const m = (total % 12) + 1;
      if (rule.byMonth?.length && !rule.byMonth.includes(m)) continue;
      for (const d of daysOfMonth(y, m, rule, first.d)) {
        const p = { y, m, d, H: first.H, M: first.M, S: first.S };
        if (dayNumber(p) < dayNumber(first)) continue;
        yield* emit(p);
        if (produced >= hardLimit) return;
      }
    }
  } else if (rule.freq === 'YEARLY') {
    const months = rule.byMonth?.length ? rule.byMonth : [first.m];
    for (let k = 0; produced < hardLimit; k++) {
      const y = first.y + k * interval;
      for (const m of months) {
        const days = rule.byMonthDay?.length || rule.byDay?.length ? daysOfMonth(y, m, rule, first.d) : first.d <= daysInMonth(y, m) ? [first.d] : [];
        for (const d of days) {
          const p = { y, m, d, H: first.H, M: first.M, S: first.S };
          if (dayNumber(p) < dayNumber(first)) continue;
          yield* emit(p);
          if (produced >= hardLimit) return;
        }
      }
    }
  } else {
    yield* emit(first);
  }
}

const sameMoment = (a, b) => Math.abs(a - b) < 1000;

/**
 * The occurrences of an event between two instants.
 *
 * `siblings` are the other VEVENTs in the same calendar with this UID — the
 * RECURRENCE-ID exceptions. Answers `[{ start, end, event, original }]`, each
 * with the instants of that occurrence, the event whose properties apply
 * (the exception, when there is one), and the instant the rule had put it at.
 */
export function occurrences(event, { from, to, siblings = [], limit = 1000 } = {}) {
  if (!event?.start) return [];
  const duration = Math.max(0, (event.end?.at ?? event.start.at) - event.start.at);
  const out = [];
  const exceptions = siblings.filter((s) => s.recurrenceId && s.uid === event.uid);
  const exdates = (event.exdates || []).map((x) => x.at);
  const push = (startAt, original, which) => {
    if (to != null && startAt > to) return false;
    const endAt = startAt + ((which.end?.at ?? which.start.at) - which.start.at || duration);
    if (from != null && endAt < from && startAt < from) return true;
    out.push({ start: startAt, end: endAt, event: which, original });
    return out.length < limit;
  };

  if (!event.rrule) {
    push(event.start.at, event.start.at, event);
  } else {
    const rule = event.rrule;
    const hardLimit = Math.min(limit * 4, rule.count || Infinity, 20000);
    let count = 0;
    for (const p of ruleStarts(event.start, rule, hardLimit)) {
      const at = instantOf(p, event.start);
      if (rule.until && at > rule.until.at + (rule.until.allDay ? DAY_MS - 1 : 0)) break;
      if (rule.count && count >= rule.count) break;
      count += 1;
      if (to != null && at > to) break;
      if (exdates.some((x) => sameMoment(x, at))) continue;
      const exception = exceptions.find((s) => sameMoment(s.recurrenceId.at, at));
      if (exception) {
        if (!push(exception.start.at, at, exception)) break;
      } else if (!push(at, at, event)) break;
    }
  }
  for (const r of event.rdates || []) {
    if (exdates.some((x) => sameMoment(x, r.at))) continue;
    if (out.some((o) => sameMoment(o.original, r.at))) continue;
    push(r.at, r.at, event);
  }
  // An exception moved outside the rule's own instants still shows once.
  for (const ex of exceptions) {
    if (!out.some((o) => o.event === ex) && !exdates.some((x) => sameMoment(x, ex.recurrenceId.at))) push(ex.start.at, ex.recurrenceId.at, ex);
  }
  out.sort((a, b) => a.start - b.start);
  return from != null ? out.filter((o) => o.end >= from || o.start >= from) : out;
}

/** A sentence for a rule, the way a calendar shows it under the event. */
export function describeRule(rule, start = null) {
  if (!rule?.freq) return 'Does not repeat';
  const every = rule.interval > 1 ? `every ${rule.interval} ` : 'every ';
  const names = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const ordinal = (n) => ({ 1: 'first', 2: 'second', 3: 'third', 4: 'fourth', 5: 'fifth', '-1': 'last', '-2': 'second-last' })[n] || `${n}th`;
  let what;
  switch (rule.freq) {
    case 'DAILY': what = rule.interval > 1 ? `${every}${rule.interval === 1 ? 'day' : 'days'}` : 'every day'; break;
    case 'WEEKLY': {
      const days = rule.byDay?.length ? rule.byDay.map((b) => names[b.day]).join(', ') : start ? names[weekday(partsOf(start))] : 'week';
      what = `${every}${rule.interval > 1 ? 'weeks' : 'week'} on ${days}`;
      break;
    }
    case 'MONTHLY': {
      const on = rule.byDay?.length
        ? rule.byDay.map((b) => (b.n ? `the ${ordinal(b.n)} ${names[b.day]}` : `every ${names[b.day]}`)).join(', ')
        : rule.byMonthDay?.length ? `day ${rule.byMonthDay.join(', ')}` : start ? `day ${partsOf(start).d}` : '';
      what = `${every}${rule.interval > 1 ? 'months' : 'month'}${on ? ` on ${on}` : ''}`;
      break;
    }
    case 'YEARLY': what = `${every}${rule.interval > 1 ? 'years' : 'year'}`; break;
    default: what = rule.freq.toLowerCase();
  }
  let tail = '';
  if (rule.count) tail = `, ${rule.count} times`;
  else if (rule.until) tail = `, until ${rule.until.date}`;
  return what.charAt(0).toUpperCase() + what.slice(1) + tail;
}

export default { occurrences, describeRule };
