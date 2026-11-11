// When people and rooms are free, on request.
//
// The event editor's "Check availability" asks for the addresses on the
// Invite line and the time chosen. Each is answered from, in this order:
//
// 1. what is already known for the span: a room whose mailbox is run here
//    (calendar-desk.js), and an answer that came before, by a server or by
//    mail, still fresh;
// 2. the server of the account the meeting goes from, where it keeps
//    free/busy: a CalDAV server's scheduling outbox (RFC 6638, dav.js through
//    the sync's hook), Microsoft Graph's getSchedule, Google Calendar's
//    freeBusy (each needs a sign-in that lets it read free/busy: `tokens`);
// 3. else an emailed enquiry, a VFREEBUSY REQUEST (RFC 5546 3.3) sent from
//    that account; the answer comes back by mail, is taken in as the Inboxes
//    are read (take, below), and the editor shows it when it arrives.
//
// What fills anyone's time is never asked for and never kept: only when
// they are busy.

import { writeFreeBusyRequest, readScheduling, mergeBusy, clashes, sameAddress } from '@rutba/calendar';
import { wordsFrom, fill } from './calendar-text.js';

const MINUTE = 60000;
/** How long an answer is trusted. */
const FRESH = 30 * MINUTE;
/** How soon the same address is asked again by mail, when no answer came. */
const AGAIN = 20 * MINUTE;
/** The widest span asked about. */
const WIDEST = 62 * 86400000;

const lower = (s) => String(s || '').trim().toLowerCase();
const iso = (ms) => new Date(ms).toISOString().replace(/\.\d{3}Z$/, 'Z');

/** Microsoft Graph's getSchedule: the busy times of each address, read from its scheduleItems. */
export async function graphSchedule({ token, emails, start, end, base = 'https://graph.microsoft.com/v1.0', fetch: fetchImpl = globalThis.fetch }) {
  const res = await fetchImpl(`${base}/me/calendar/getSchedule`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json', Prefer: 'outlook.timezone="UTC"' },
    body: JSON.stringify({ schedules: emails, startTime: { dateTime: iso(start).slice(0, -1), timeZone: 'UTC' }, endTime: { dateTime: iso(end).slice(0, -1), timeZone: 'UTC' }, availabilityViewInterval: 15 }),
    signal: AbortSignal.timeout(20000),
  });
  if (!res.ok) throw new Error(`Microsoft answered ${res.status} where free and busy times were asked for.`);
  const data = await res.json();
  const at = (t) => Date.parse(/Z$|[+-]\d\d:?\d\d$/.test(t?.dateTime || '') ? t.dateTime : `${t?.dateTime}Z`);
  return (data.value || []).map((s) => ({
    email: lower(s.scheduleId),
    ok: !s.error,
    busy: (s.scheduleItems || []).filter((i) => i.status !== 'free').map((i) => ({ start: at(i.start), end: at(i.end) })).filter((p) => p.end > p.start),
  }));
}

/** Google Calendar's freeBusy: the busy times of each address it can read. */
export async function googleFreeBusy({ token, emails, start, end, base = 'https://www.googleapis.com/calendar/v3', fetch: fetchImpl = globalThis.fetch }) {
  const res = await fetchImpl(`${base}/freeBusy`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ timeMin: iso(start), timeMax: iso(end), items: emails.map((id) => ({ id })) }),
    signal: AbortSignal.timeout(20000),
  });
  if (!res.ok) throw new Error(`Google answered ${res.status} where free and busy times were asked for.`);
  const data = await res.json();
  return Object.entries(data.calendars || {}).map(([email, c]) => ({
    email: lower(email),
    ok: !(c.errors || []).length,
    busy: (c.busy || []).map((p) => ({ start: Date.parse(p.start), end: Date.parse(p.end) })).filter((p) => p.end > p.start),
  }));
}

/**
 * @param {object} o
 * @param {() => object} o.state the calendar's state (`freeBusy`, `freeBusyAsked`)
 * @param {() => void} o.save
 * @param {object} o.link the Mail link (calendar-mail.js)
 * @param {object} o.invites (calendar-invites.js) for the account a meeting goes from and the files Mail attaches
 * @param {object[]} o.sources what is known already (calendar-rooms.js's sources: the calendars here, the room desk)
 * @param {object} o.hooks the sync's hooks: `freeBusy` asks a CalDAV server
 * @param {(account: object) => ({ kind: 'graph' | 'google', token: string } | null)} [o.tokens] a sign-in that reads free/busy, for an account
 */
export function createAvailability({ state, save, link, invites, sources = [], hooks = {}, tokens = () => null, fetch: fetchImpl = globalThis.fetch, now = () => Date.now() }) {
  const known = () => (state().freeBusy ||= {});
  const asked = () => (state().freeBusyAsked ||= {});
  const words = () => wordsFrom(state().words);

  /** An answer kept for an address that covers the span, still fresh. */
  const kept = (email, start, end) => {
    const k = known()[lower(email)];
    return k && k.start <= start && k.end >= end && now() - k.at < FRESH ? k : null;
  };
  const keep = (email, { start, end, busy, source }) => {
    known()[lower(email)] = { start, end, busy: mergeBusy(busy || []), source, at: now() };
  };
  // Answers older than a day are dropped as the next ones are kept.
  const tidy = () => {
    for (const [k, v] of Object.entries(known())) if (now() - v.at > 86400000) delete known()[k];
    for (const [k, v] of Object.entries(asked())) if (now() - v.at > 86400000) delete asked()[k];
  };

  /** What a server says, for those not answered yet: CalDAV first, then Graph or Google for the account's own sign-in. */
  const fromServer = async (account, organizer, emails, start, end) => {
    const out = [];
    if (typeof hooks.freeBusy === 'function') {
      try {
        const got = await hooks.freeBusy({ organizer, attendees: emails, start, end });
        for (const a of got?.answers || []) if (a.ok) out.push({ ...a, source: 'caldav' });
      } catch {
        /* the next way */
      }
    }
    const left = emails.filter((e) => !out.some((o) => sameAddress(o.email, e)));
    const sign = left.length && account ? tokens(account) : null;
    if (sign?.token) {
      try {
        const got = sign.kind === 'google'
          ? await googleFreeBusy({ token: sign.token, emails: left, start, end, base: sign.base, fetch: fetchImpl })
          : await graphSchedule({ token: sign.token, emails: left, start, end, base: sign.base, fetch: fetchImpl });
        for (const a of got) if (a.ok) out.push({ ...a, source: sign.kind });
      } catch {
        /* asked by mail instead */
      }
    }
    return out;
  };

  /** An emailed enquiry to each address not asked lately. */
  const byMail = async (account, organizer, emails, start, end) => {
    const w = words();
    const sent = [];
    const day = (ms) => new Intl.DateTimeFormat(w.language || 'en', { weekday: 'short', day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' }).format(ms);
    for (const email of emails) {
      const before = asked()[lower(email)];
      if (before && now() - before.at < AGAIN && before.start <= start && before.end >= end) {
        sent.push(email);
        continue;
      }
      const uid = `fb-${now().toString(36)}-${Math.random().toString(36).slice(2, 8)}@rutba.io`;
      const ics = writeFreeBusyRequest({ uid, organizer, attendees: [{ email }], start, end });
      const vars = { from: day(start), to: day(end) };
      await link.sendFrom({ account: account.id, to: [email], subject: fill(w.freeBusySubject, vars), text: fill(w.freeBusyBody, vars), attachments: [invites.attach('REQUEST', ics)] });
      asked()[lower(email)] = { uid, start, end, at: now(), from: account.address || account.email };
      sent.push(email);
    }
    return sent;
  };

  return {
    /**
     * Whether each address is free in [start, end): `free` true or false with
     * the busy periods and where they came from, `asked` when an enquiry went
     * by mail and the answer is awaited, else unknown. `ask` false only reads
     * what is known (the editor looking again after a change).
     */
    check: async ({ emails = [], start, end, ask = true }) => {
      tidy();
      const span = { start: Math.min(start, end), end: Math.min(Math.max(start, end), Math.min(start, end) + WIDEST) };
      const wanted = [...new Set(emails.map(lower).filter((e) => /@/.test(e)))];
      const rows = new Map();
      const answer = (email, busy, source) => rows.set(email, { email, free: !clashes(busy, start, end), busy, source });
      for (const email of wanted) {
        const k = kept(email, span.start, span.end);
        if (k) {
          answer(email, k.busy, k.source);
          continue;
        }
        for (const s of sources) {
          const got = s.busy(email, span.start, span.end, null);
          if (got?.whole) {
            answer(email, got.busy, got.source);
            break;
          }
        }
      }
      let account = null;
      let organizer = null;
      const open = wanted.filter((e) => !rows.has(e));
      if (open.length && ask) {
        account = invites.sendingAccount();
        organizer = account ? { email: account.address || account.email, name: account.name || null } : null;
        if (organizer) {
          const found = await fromServer(account, organizer, open, span.start, span.end);
          for (const f of found) {
            keep(f.email, { ...span, busy: f.busy, source: f.source });
            answer(lower(f.email), mergeBusy(f.busy), f.source);
          }
          const still = open.filter((e) => !rows.has(e));
          if (still.length) await byMail(account, organizer, still, span.start, span.end);
          save();
        }
      }
      // What is left: asked by mail and waiting, or not known; a clash in the calendars here says busy all the same.
      for (const email of wanted) {
        if (rows.has(email)) continue;
        const here = sources.map((s) => s.busy(email, span.start, span.end, null)).filter(Boolean).flatMap((g) => g.busy || []);
        const waiting = asked()[email] && now() - asked()[email].at < 86400000;
        rows.set(email, { email, free: clashes(here, start, end) ? false : null, busy: mergeBusy(here), source: clashes(here, start, end) ? 'calendar' : null, asked: Boolean(waiting) });
      }
      return { rows: wanted.map((e) => rows.get(e)), from: organizer?.email || null };
    },

    /** A VFREEBUSY REPLY that arrived by mail, kept as that person's busy times (the scan's `also`). */
    take: ({ part, row }) => {
      if (!part?.text) return null;
      let read;
      try {
        read = readScheduling(part.text);
      } catch {
        return null;
      }
      if (String(read.method || '').toUpperCase() !== 'REPLY' || !read.freebusy?.length) return null;
      const from = row?.from?.address || null;
      let took = 0;
      for (const fb of read.freebusy) {
        for (const a of fb.attendees || []) {
          // An answer for the person who sent it, to an enquiry sent from here.
          if (!a.email || (from && !sameAddress(a.email, from))) continue;
          const ask = asked()[lower(a.email)];
          if (!ask && !(fb.uid && Object.values(asked()).some((x) => x.uid === fb.uid))) continue;
          keep(a.email, { start: fb.start ?? ask?.start, end: fb.end ?? ask?.end, busy: fb.busy || [], source: 'email' });
          delete asked()[lower(a.email)];
          took += 1;
        }
      }
      if (!took) return null;
      save();
      return { kind: 'freebusy', from };
    },

    /** What an answer kept says, as a source for "Add a room" (calendar-rooms.js): the span it answered for. */
    source: {
      busy: (email, from, to) => {
        const k = known()[lower(email)];
        if (!k || now() - k.at > FRESH) return null;
        return { busy: k.busy, whole: k.start <= from && k.end >= to, source: k.source };
      },
    },
  };
}
