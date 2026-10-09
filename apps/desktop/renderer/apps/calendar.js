// Rutba Calendar.
//
// A calendar that lives on this computer: as many calendars as the person
// wants, each with a colour, shown together by month, week, day or as the
// list of what is coming. An event is a dialog away; a recurring one is one
// event with a rule, and moving a single Tuesday makes an exception rather
// than a mess. Files come and go as .ics: a file opened from disk is shown
// beside the person's own calendars with an offer to keep it, and an
// invitation gets Accept, Tentative and Decline, which answer through mail.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  Ribbon, Group, Button, Separator, Icon, Spacer, Chip, Empty, Panel, Content, Field, Input, Select, Dialog,
  useToast, useCommands, useMenu, menuItems, t, tn, language,
} from '@rutba/office-ui';
import { AppFrame, useAppMenu, pickOpen, pickSave, useFileDrop } from '../shell.js';
import { AccountsDialog } from '../dav-accounts.js';

const DAY = 86400000;
// The days, Monday first, and the months, as the window's language names them
// (Mon and January in English): 1 January 2024 was a Monday.
const named = (options, count, at) => Array.from({ length: count }, (_, i) => at(i).toLocaleDateString(language(), options));
const DAYS = named({ weekday: 'short' }, 7, (i) => new Date(2024, 0, 1 + i));
const DAY_LETTERS = named({ weekday: 'narrow' }, 7, (i) => new Date(2024, 0, 1 + i));
const MONTHS = named({ month: 'long' }, 12, (i) => new Date(2024, i, 1));
const MONTHS_SHORT = named({ month: 'short' }, 12, (i) => new Date(2024, i, 1));
const HOUR_PX = 44;
const VIEW_NAMES = { day: t('Day'), week: t('Week'), month: t('Month'), agenda: t('Agenda') };
const ANSWERS = { ACCEPTED: t('accepted'), DECLINED: t('declined'), TENTATIVE: t('tentative'), DELEGATED: t('delegated') };

/* ── dates on this machine's clock ─────────────────────────────────────── */

const startOfDay = (d) => new Date(d.getFullYear(), d.getMonth(), d.getDate());
const addDays = (d, n) => new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
const sameDay = (a, b) => a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
/** Monday of the week that holds the date. */
const startOfWeek = (d) => addDays(startOfDay(d), -((d.getDay() + 6) % 7));
const pad = (n) => String(n).padStart(2, '0');
const isoDate = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const isoTime = (d) => `${pad(d.getHours())}:${pad(d.getMinutes())}`;
const clock = (ms) => {
  const d = new Date(ms);
  return `${d.getHours()}:${pad(d.getMinutes())}`;
};

/** The range a view shows, as instants. */
function rangeFor(view, cursor) {
  if (view === 'month') {
    const first = new Date(cursor.getFullYear(), cursor.getMonth(), 1);
    const from = startOfWeek(first);
    const to = addDays(from, 42);
    return { from, to };
  }
  if (view === 'week') {
    const from = startOfWeek(cursor);
    return { from, to: addDays(from, 7) };
  }
  if (view === 'day') {
    const from = startOfDay(cursor);
    return { from, to: addDays(from, 1) };
  }
  const from = startOfDay(cursor);
  return { from, to: addDays(from, 30) };
}

/** An all-day slot's instants are UTC midnights of calendar dates; place it on those dates locally. */
function slotDays(slot) {
  if (slot.allDay) {
    const s = new Date(slot.start);
    const e = new Date(slot.end);
    const first = new Date(s.getUTCFullYear(), s.getUTCMonth(), s.getUTCDate());
    const last = new Date(e.getUTCFullYear(), e.getUTCMonth(), e.getUTCDate() - 1);
    return { first, last: last < first ? first : last };
  }
  return { first: startOfDay(new Date(slot.start)), last: startOfDay(new Date(Math.max(slot.start, slot.end - 1))) };
}

/* ── the event editor's form ───────────────────────────────────────────── */

function blankForm(at, allDay = false) {
  const start = at ? new Date(at) : new Date();
  if (!at) start.setMinutes(0, 0, 0), start.setHours(start.getHours() + 1);
  const end = new Date(start.getTime() + 3600000);
  return {
    id: null, calendarId: null, summary: '', location: '', description: '',
    allDay, date: isoDate(start), time: isoTime(start), endDate: isoDate(end), endTime: isoTime(end),
    repeat: 'none', repeatUntil: '', reminder: '15', attendees: '', original: null, rule: null, scope: 'all',
  };
}

function formFrom(event, original = null) {
  const s = event.start;
  const e = event.end || event.start;
  const start = event.allDay ? new Date(`${s.date}T00:00:00`) : new Date(s.at);
  const end = event.allDay ? new Date(`${e.date}T00:00:00`) : new Date(e.at);
  if (event.allDay) end.setDate(end.getDate() - 1);
  const rule = event.rrule;
  const repeat = !rule ? 'none' : rule.freq === 'DAILY' && rule.interval === 1 ? 'daily' : rule.freq === 'WEEKLY' && rule.interval === 1 && !(rule.byDay?.length > 1) ? 'weekly' : rule.freq === 'WEEKLY' && rule.byDay?.length === 5 ? 'weekdays' : rule.freq === 'MONTHLY' && rule.interval === 1 ? 'monthly' : rule.freq === 'YEARLY' && rule.interval === 1 ? 'yearly' : 'custom';
  return {
    id: event.id, calendarId: event.calendarId, summary: event.summary || '', location: event.location || '', description: event.description || '',
    allDay: Boolean(event.allDay), date: isoDate(start), time: isoTime(start), endDate: isoDate(end), endTime: isoTime(end),
    repeat, repeatUntil: rule?.until?.date || '', reminder: event.alarms?.[0]?.before != null ? String(event.alarms[0].before) : 'none',
    attendees: (event.attendees || []).map((a) => a.email).filter(Boolean).join(', '), original, rule: rule || null, scope: 'all',
    uid: event.uid, sequence: event.sequence, organizer: event.organizer, extra: event.extra, exdates: event.exdates, rdates: event.rdates, recurrenceId: event.recurrenceId,
  };
}

/** The form back into an event the service keeps. */
function eventFrom(form, tz) {
  const [y, m, d] = form.date.split('-').map(Number);
  const [ey, em, ed] = (form.endDate || form.date).split('-').map(Number);
  let start;
  let end;
  if (form.allDay) {
    start = { date: form.date, time: null, tzid: null, utc: false, allDay: true, at: Date.UTC(y, m - 1, d) };
    const last = new Date(Date.UTC(ey, em - 1, ed + 1));
    end = { date: last.toISOString().slice(0, 10), time: null, tzid: null, utc: false, allDay: true, at: last.getTime() };
    if (end.at <= start.at) end = { ...start, date: new Date(Date.UTC(y, m - 1, d + 1)).toISOString().slice(0, 10), at: start.at + DAY };
  } else {
    const [H, M] = (form.time || '09:00').split(':').map(Number);
    const [eH, eM] = (form.endTime || form.time || '10:00').split(':').map(Number);
    const s = new Date(y, m - 1, d, H, M);
    let e = new Date(ey, em - 1, ed, eH, eM);
    if (e <= s) e = new Date(s.getTime() + 3600000);
    start = { date: form.date, time: `${pad(H)}:${pad(M)}:00`, tzid: tz, utc: false, allDay: false, at: s.getTime() };
    end = { date: isoDate(e), time: `${pad(e.getHours())}:${pad(e.getMinutes())}:00`, tzid: tz, utc: false, allDay: false, at: e.getTime() };
  }
  const until = form.repeatUntil ? { date: form.repeatUntil, time: null, tzid: null, utc: false, allDay: true, at: new Date(`${form.repeatUntil}T00:00:00Z`).getTime() } : null;
  const weekday = (new Date(y, m - 1, d).getDay() + 0) % 7;
  const rules = {
    none: null,
    daily: { freq: 'DAILY', interval: 1, count: null, until, byDay: null, byMonthDay: null, byMonth: null, bySetPos: null, wkst: 'MO' },
    weekly: { freq: 'WEEKLY', interval: 1, count: null, until, byDay: [{ n: 0, day: weekday }], byMonthDay: null, byMonth: null, bySetPos: null, wkst: 'MO' },
    weekdays: { freq: 'WEEKLY', interval: 1, count: null, until, byDay: [1, 2, 3, 4, 5].map((day) => ({ n: 0, day })), byMonthDay: null, byMonth: null, bySetPos: null, wkst: 'MO' },
    monthly: { freq: 'MONTHLY', interval: 1, count: null, until, byDay: null, byMonthDay: [d], byMonth: null, bySetPos: null, wkst: 'MO' },
    yearly: { freq: 'YEARLY', interval: 1, count: null, until, byDay: null, byMonthDay: null, byMonth: null, bySetPos: null, wkst: 'MO' },
    custom: form.rule,
  };
  const attendees = form.attendees.split(/[,;\s]+/).map((s) => s.trim()).filter((s) => s.includes('@')).map((email) => ({ email, name: null, role: 'REQ-PARTICIPANT', partstat: 'NEEDS-ACTION', rsvp: true }));
  return {
    id: form.id, uid: form.uid, sequence: form.sequence, organizer: form.organizer, extra: form.extra || [], exdates: form.exdates || [], rdates: form.rdates || [], recurrenceId: form.recurrenceId || null,
    summary: form.summary.trim() || '(no title)', location: form.location.trim(), description: form.description.trim(),
    allDay: form.allDay, start, end, rrule: rules[form.repeat] ?? null,
    alarms: form.reminder && form.reminder !== 'none' ? [{ action: 'DISPLAY', before: Number(form.reminder), at: null, related: 'START', description: null }] : [],
    attendees, status: 'CONFIRMED',
  };
}

/* ── the window ────────────────────────────────────────────────────────── */

export default function Calendar({ app, shell, boot }) {
  const toast = useToast();
  const menu = useMenu();
  const [view, setView] = useState('month');
  const [cursor, setCursor] = useState(() => new Date());
  const [calendars, setCalendars] = useState([]);
  const [slots, setSlots] = useState([]);
  const [file, setFile] = useState(null);
  const [form, setForm] = useState(null);
  const [opened, setOpened] = useState(null);
  const [tab, setTab] = useState('home');
  const [now, setNow] = useState(() => Date.now());
  const tz = useMemo(() => Intl.DateTimeFormat().resolvedOptions().timeZone, []);
  const range = useMemo(() => rangeFor(view, cursor), [view, cursor]);
  const today = useMemo(() => startOfDay(new Date(now)), [now]);

  useEffect(() => {
    const t = setInterval(() => setNow(Date.now()), 60000);
    return () => clearInterval(t);
  }, []);

  const refresh = useCallback(async () => {
    try {
      const [cals, own] = await Promise.all([shell.calendar.calendars(), shell.calendar.events({ from: range.from.getTime(), to: range.to.getTime() })]);
      setCalendars(cals);
      let all = own;
      if (file) {
        const seen = await shell.calendar.openFile({ path: file.path, from: range.from.getTime(), to: range.to.getTime() });
        all = [...own, ...seen.events];
      }
      setSlots(all);
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    }
  }, [shell, range, file, toast]);

  useEffect(() => {
    refresh();
  }, [refresh]);
  useEffect(() => shell.on('calendar:changed', refresh), [shell, refresh]);

  // A file opened with the window, or dropped on it: shown, and offered.
  const showFile = useCallback(
    async (path) => {
      try {
        const seen = await shell.calendar.openFile({ path, from: range.from.getTime(), to: range.to.getTime() });
        setFile(seen);
        if (seen.first?.start?.at) setCursor(new Date(seen.first.allDay ? new Date(seen.first.start.date + 'T00:00:00') : seen.first.start.at));
        toast(seen.invitation ? t('An invitation from {name}.', { name: seen.invitation.organizer?.name || seen.invitation.organizer?.email || t('someone') }) : tn(seen.count, '{count} event in {name}.', '{count} events in {name}.', { name: seen.name }), { ms: 5000 });
      } catch (err) {
        toast(err.message, { tone: 'bad', ms: 8000 });
      }
    },
    [shell, range, toast]
  );
  useEffect(() => {
    if (boot.file) showFile(boot.file);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  useFileDrop(useCallback((files) => files.forEach((f) => (/\.ics$/i.test(f) ? showFile(f) : null)), [showFile]));

  /* ── actions ──────────────────────────────────────────────────────────── */

  const go = useCallback((n) => setCursor((c) => (view === 'month' ? new Date(c.getFullYear(), c.getMonth() + n, 1) : addDays(c, n * (view === 'week' ? 7 : view === 'day' ? 1 : 30)))), [view]);
  const newEvent = useCallback((at = null, allDay = false) => setForm({ ...blankForm(at, allDay), calendarId: calendars.find((c) => c.visible)?.id || calendars[0]?.id || null }), [calendars]);

  const openSlot = useCallback(
    async (slot) => {
      if (String(slot.id).startsWith('file:')) {
        setOpened({ slot, event: null, readOnly: true });
        return;
      }
      try {
        const event = await shell.calendar.get({ id: slot.id });
        setOpened({ slot, event, readOnly: false });
      } catch (err) {
        toast(err.message, { tone: 'bad' });
      }
    },
    [shell, toast]
  );

  const saveForm = useCallback(async () => {
    if (!form) return;
    try {
      const event = eventFrom(form, tz);
      await shell.calendar.save({ calendarId: form.calendarId, event, scope: form.scope, original: form.original });
      setForm(null);
      setOpened(null);
      await refresh();
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    }
  }, [form, shell, tz, refresh, toast]);

  const removeOpened = useCallback(
    async (scope = 'all') => {
      if (!opened?.event) return;
      await shell.calendar.remove({ id: opened.event.id, scope, original: opened.slot.original });
      setOpened(null);
      await refresh();
    },
    [opened, shell, refresh]
  );

  const importFile = useCallback(
    async (path = null) => {
      const target = path || (await pickOpen(shell, 'calendar'))?.[0];
      if (!target) return;
      try {
        const r = await shell.calendar.importFile({ path: target });
        setFile(null);
        await refresh();
        toast(t('{added} added, {updated} updated in {calendar}.', { added: r.added, updated: r.updated, calendar: r.calendar }), { tone: 'good', ms: 6000 });
      } catch (err) {
        toast(err.message, { tone: 'bad', ms: 8000 });
      }
    },
    [shell, refresh, toast]
  );

  const exportFile = useCallback(async () => {
    const target = await pickSave(shell, 'calendar', 'calendar.ics');
    if (!target) return;
    try {
      const r = await shell.calendar.exportFile({ path: target });
      toast(tn(r.count, 'Wrote {count} event.', 'Wrote {count} events.'), { tone: 'good' });
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    }
  }, [shell, toast]);

  const respond = useCallback(
    async (partstat) => {
      if (!file?.invitation) return;
      try {
        const r = await shell.calendar.respond({ path: file.path, partstat });
        if (r.replyPath && r.organizer?.email) {
          shell.win.create({ app: 'mail', query: { to: r.organizer.email, subject: r.subject, attach: r.replyPath } });
        }
        setFile(null);
        await refresh();
        toast(r.kept ? t('Added to your calendar, and the reply is ready to send.') : t('Declined; the reply is ready to send.'), { tone: 'good', ms: 6000 });
      } catch (err) {
        toast(err.message, { tone: 'bad', ms: 8000 });
      }
    },
    [file, shell, refresh, toast]
  );

  const invite = useCallback(async () => {
    if (!opened?.event?.attendees?.length) return;
    try {
      const r = await shell.calendar.invitationFile({ id: opened.event.id });
      shell.win.create({ app: 'mail', query: { to: r.to.join(', '), subject: r.subject, attach: r.path } });
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    }
  }, [opened, shell, toast]);

  const toggleCalendar = useCallback(
    async (cal) => {
      await shell.calendar.saveCalendar({ calendar: { id: cal.id, visible: !cal.visible } });
      await refresh();
    },
    [shell, refresh]
  );
  // Home → Sync: every account now, and what moved said.
  const [accountsOpen, setAccountsOpen] = useState(false);
  const syncAll = useCallback(async () => {
    const results = (await shell.dav.sync({})) || [];
    if (!results.length) { setAccountsOpen(true); return; }
    const failed = results.find((r) => r.error);
    if (failed) toast(failed.error, { tone: 'warn', ms: 5500 });
    else toast(t('Synced: {received} in, {sent} out', { received: results.reduce((n, r) => n + (r.received || 0), 0), sent: results.reduce((n, r) => n + (r.sent || 0), 0) }), { tone: 'good' });
    await refresh();
  }, [shell, toast, refresh]);
  const addCalendar = useCallback(async () => {
    const name = `Calendar ${calendars.length + 1}`;
    await shell.calendar.saveCalendar({ calendar: { name } });
    await refresh();
  }, [shell, calendars, refresh]);

  const appMenu = useAppMenu({ shell, appKey: 'calendar', onNew: () => newEvent(), onOpen: () => importFile() });
  const commands = useMemo(
    () => ({
      'event.new': { label: t('New event'), icon: 'plus', key: 'Mod+N', run: () => newEvent() },
      'view.today': { label: t('Today'), icon: 'clock', key: 'Mod+T', run: () => setCursor(new Date()) },
      'view.month': { label: t('Month'), icon: 'grid', run: () => setView('month') },
      'view.week': { label: t('Week'), icon: 'table', run: () => setView('week') },
      'view.day': { label: t('Day'), icon: 'file', run: () => setView('day') },
      'view.agenda': { label: t('Agenda'), icon: 'list', run: () => setView('agenda') },
      'file.import': { label: t('Import…'), icon: 'import', key: 'Mod+O', run: () => importFile() },
      'file.export': { label: t('Export…'), icon: 'export', run: exportFile },
    }),
    [newEvent, importFile, exportFile]
  );
  useCommands(commands, [view, cursor]);

  const title = view === 'month' ? `${MONTHS[cursor.getMonth()]} ${cursor.getFullYear()}` : view === 'week' ? t('Week of {date}', { date: `${range.from.getDate()} ${MONTHS_SHORT[range.from.getMonth()]} ${range.from.getFullYear()}` }) : view === 'day' ? `${DAYS[(cursor.getDay() + 6) % 7]} ${cursor.getDate()} ${MONTHS[cursor.getMonth()]} ${cursor.getFullYear()}` : t('Next 30 days');

  /* ── views ────────────────────────────────────────────────────────────── */

  const visible = useMemo(() => slots.filter((s) => s.file || calendars.find((c) => c.id === s.calendarId)?.visible !== false), [slots, calendars]);

  const chip = (slot, key, extra = '') => (
    <button
      type="button"
      key={key}
      className={`cal-chip${slot.allDay ? ' allday' : ''}${extra}`}
      style={{ '--c': slot.colour }}
      title={`${slot.summary}${slot.location ? ` — ${slot.location}` : ''}`}
      onClick={(e) => { e.stopPropagation(); openSlot(slot); }}
    >
      {slot.allDay ? null : <span className="t">{clock(slot.start)}</span>}
      <span className="s">{slot.summary}</span>
    </button>
  );

  const monthView = () => {
    const days = [];
    for (let i = 0; i < 42; i++) days.push(addDays(range.from, i));
    const byDay = new Map();
    for (const s of visible) {
      const { first, last } = slotDays(s);
      for (let d = first; d <= last; d = addDays(d, 1)) {
        const k = isoDate(d);
        if (!byDay.has(k)) byDay.set(k, []);
        byDay.get(k).push(s);
      }
    }
    return (
      <div className="cal-month">
        <div className="cal-month-head">{DAYS.map((d) => <div key={d}>{d}</div>)}</div>
        <div className="cal-month-grid">
          {days.map((d) => {
            const list = (byDay.get(isoDate(d)) || []).sort((a, b) => Number(b.allDay) - Number(a.allDay) || a.start - b.start);
            const shown = list.slice(0, 4);
            return (
              <div
                key={isoDate(d)}
                className={`cal-day${d.getMonth() !== cursor.getMonth() ? ' other' : ''}${sameDay(d, today) ? ' today' : ''}`}
                onClick={() => newEvent(new Date(d.getFullYear(), d.getMonth(), d.getDate(), 9), false)}
                onDoubleClick={(e) => e.stopPropagation()}
              >
                <div className="cal-day-n">{d.getDate() === 1 ? `${d.getDate()} ${MONTHS_SHORT[d.getMonth()]}` : d.getDate()}</div>
                {shown.map((s, i) => chip(s, `${s.id}-${s.original}-${i}`))}
                {list.length > shown.length ? <button type="button" className="cal-more" onClick={(e) => { e.stopPropagation(); setCursor(d); setView('day'); }}>{t('+{count} more', { count: list.length - shown.length })}</button> : null}{/* words-ok: code, not words */}
              </div>
            );
          })}
        </div>
      </div>
    );
  };

  const timeGrid = (days) => {
    const allDay = visible.filter((s) => s.allDay);
    const timed = visible.filter((s) => !s.allDay);
    return (
      <div className="cal-time">
        <div className="cal-time-head">
          <div className="cal-gutter" />
          {days.map((d) => (
            <div key={isoDate(d)} className={`cal-col-head${sameDay(d, today) ? ' today' : ''}`}>
              <span className="dn">{DAYS[(d.getDay() + 6) % 7]}</span> <span className="dd">{d.getDate()}</span>
            </div>
          ))}
        </div>
        <div className="cal-allday">
          <div className="cal-gutter">{t('all day')}</div>
          {days.map((d) => (
            <div key={isoDate(d)} className="cal-allday-cell" onClick={() => newEvent(new Date(d.getFullYear(), d.getMonth(), d.getDate()), true)}>
              {allDay.filter((s) => { const { first, last } = slotDays(s); return d >= first && d <= last; }).map((s, i) => chip(s, `${s.id}-${i}`))}
            </div>
          ))}
        </div>
        <div className="cal-hours">
          <div className="cal-gutter">
            {Array.from({ length: 24 }, (_, h) => <div key={h} className="cal-hour" style={{ height: HOUR_PX }}>{h ? `${h}:00` : ''}</div>)}
          </div>
          {days.map((d) => {
            const dayStart = d.getTime();
            const mine = timed.filter((s) => s.start < dayStart + DAY && s.end > dayStart);
            return (
              <div
                key={isoDate(d)}
                className={`cal-col${sameDay(d, today) ? ' today' : ''}`}
                style={{ height: HOUR_PX * 24 }}
                onClick={(e) => {
                  const rect = e.currentTarget.getBoundingClientRect();
                  const hour = Math.floor(((e.clientY - rect.top) / HOUR_PX) * 2) / 2;
                  newEvent(new Date(d.getFullYear(), d.getMonth(), d.getDate(), Math.floor(hour), (hour % 1) * 60), false);
                }}
              >
                {Array.from({ length: 24 }, (_, h) => <div key={h} className="cal-line" style={{ top: h * HOUR_PX }} />)}
                {sameDay(d, today) ? <div className="cal-now" style={{ top: ((now - dayStart) / 3600000) * HOUR_PX }} /> : null}
                {mine.map((s, i) => {
                  const top = (Math.max(s.start, dayStart) - dayStart) / 3600000 * HOUR_PX;
                  const bottom = (Math.min(s.end, dayStart + DAY) - dayStart) / 3600000 * HOUR_PX;
                  const overlap = mine.filter((o) => o !== s && o.start < s.end && o.end > s.start);
                  const lane = overlap.filter((o) => o.start < s.start || (o.start === s.start && mine.indexOf(o) < i)).length;
                  const lanes = overlap.length + 1;
                  return (
                    <button
                      type="button"
                      key={`${s.id}-${s.original}-${i}`}
                      className="cal-block"
                      style={{ '--c': s.colour, top, height: Math.max(18, bottom - top), left: `${(lane / lanes) * 100}%`, width: `${100 / lanes}%` }}
                      title={s.summary}
                      onClick={(e) => { e.stopPropagation(); openSlot(s); }}
                    >
                      <span className="s">{s.summary}</span>
                      <span className="t">{clock(s.start)} – {clock(s.end)}</span>
                    </button>
                  );
                })}
              </div>
            );
          })}
        </div>
      </div>
    );
  };

  const agendaView = () => {
    const groups = new Map();
    for (const s of visible) {
      const k = isoDate(slotDays(s).first);
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(s);
    }
    const keys = [...groups.keys()].sort();
    return (
      <div className="cal-agenda">
        {keys.length ? keys.map((k) => {
          const d = new Date(`${k}T00:00:00`);
          return (
            <div key={k} className={`cal-agenda-day${sameDay(d, today) ? ' today' : ''}`}>
              <div className="cal-agenda-date"><span className="dn">{DAYS[(d.getDay() + 6) % 7]}</span><span className="dd">{d.getDate()}</span><span className="dm">{MONTHS_SHORT[d.getMonth()]}</span></div>
              <div className="cal-agenda-list">
                {groups.get(k).sort((a, b) => Number(b.allDay) - Number(a.allDay) || a.start - b.start).map((s, i) => (
                  <button type="button" key={`${s.id}-${i}`} className="cal-agenda-item" style={{ '--c': s.colour }} onClick={() => openSlot(s)}>
                    <span className="t">{s.allDay ? t('all day') : `${clock(s.start)} – ${clock(s.end)}`}</span>
                    <span className="s">{s.summary}</span>
                    {s.location ? <span className="l">{s.location}</span> : null}
                  </button>
                ))}
              </div>
            </div>
          );
        }) : <Empty icon="calendar" title={t('Nothing in the next thirty days')}>{t('Press Ctrl+N to add an event, or open a .ics file.')}</Empty>}
      </div>
    );
  };

  const body = view === 'month' ? monthView() : view === 'week' ? timeGrid(Array.from({ length: 7 }, (_, i) => addDays(range.from, i))) : view === 'day' ? timeGrid([startOfDay(cursor)]) : agendaView();

  /* ── dialogs ──────────────────────────────────────────────────────────── */

  const set = (patch) => setForm((f) => ({ ...f, ...patch }));
  const editor = form ? (
    <Dialog
      title={form.id ? t('Edit event') : t('New event')}
      width={560}
      onClose={() => setForm(null)}
      actions={
        <>
          <Button label={t('Cancel')} onClick={() => setForm(null)} />
          <Button primary label={t('Save')} onClick={saveForm} />
        </>
      }
    >
      <div className="cal-form">
        <Input className="rw-input cal-title" value={form.summary} placeholder={t('Title')} autoFocus onChange={(e) => set({ summary: e.target.value })} onKeyDown={(e) => { if (e.key === 'Enter') saveForm(); }} />
        <label className="cal-check"><input type="checkbox" checked={form.allDay} onChange={(e) => set({ allDay: e.target.checked })} /> {t('All day')}</label>
        <div className="cal-when">
          <Field label={t('Starts')}><div className="cal-dt"><Input type="date" value={form.date} onChange={(e) => set({ date: e.target.value, endDate: form.endDate < e.target.value ? e.target.value : form.endDate })} />{form.allDay ? null : <Input type="time" value={form.time} onChange={(e) => set({ time: e.target.value })} />}</div></Field>
          <Field label={t('Ends')}><div className="cal-dt"><Input type="date" value={form.endDate} onChange={(e) => set({ endDate: e.target.value })} />{form.allDay ? null : <Input type="time" value={form.endTime} onChange={(e) => set({ endTime: e.target.value })} />}</div></Field>
        </div>
        <div className="cal-when">
          <Field label={t('Repeat')}>
            <Select value={form.repeat} onChange={(e) => set({ repeat: e.target.value })}>
              <option value="none">{t('Does not repeat')}</option>
              <option value="daily">{t('Every day')}</option>
              <option value="weekdays">{t('Every weekday')}</option>
              <option value="weekly">{t('Every week')}</option>
              <option value="monthly">{t('Every month')}</option>
              <option value="yearly">{t('Every year')}</option>
              {form.rule ? <option value="custom">{t('As it is ({freq})', { freq: form.rule.freq.toLowerCase() })}</option> : null}
            </Select>
          </Field>
          {form.repeat !== 'none' ? <Field label={t('Until')} hint={t('Blank for no end')}><Input type="date" value={form.repeatUntil} onChange={(e) => set({ repeatUntil: e.target.value })} /></Field> : <Field label={t('Reminder')}><Select value={form.reminder} onChange={(e) => set({ reminder: e.target.value })}><option value="none">{t('None')}</option><option value="0">{t('At the time')}</option><option value="5">{t('5 minutes before')}</option><option value="15">{t('15 minutes before')}</option><option value="30">{t('30 minutes before')}</option><option value="60">{t('1 hour before')}</option><option value="1440">{t('1 day before')}</option></Select></Field>}
        </div>
        {form.id && form.rule && form.original != null ? (
          <Field label={t('Change')}>
            <Select value={form.scope} onChange={(e) => set({ scope: e.target.value })}>
              <option value="this">{t('Only this occurrence')}</option>
              <option value="all">{t('Every occurrence')}</option>
            </Select>
          </Field>
        ) : null}
        <Field label={t('Location')}><Input value={form.location} placeholder={t('Where')} onChange={(e) => set({ location: e.target.value })} /></Field>
        <Field label={t('Calendar')}>
          <Select value={form.calendarId || ''} onChange={(e) => set({ calendarId: e.target.value })}>{calendars.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}</Select>
        </Field>
        <Field label={t('Invite')} hint={t('Addresses, separated by commas. Saving with people on the list offers a message with the invitation attached.')}><Input value={form.attendees} placeholder="kim@example.com, sam@example.org" onChange={(e) => set({ attendees: e.target.value })} /></Field>{/* words-ok: example addresses */}
        <Field label={t('Notes')}><textarea className="rw-input cal-notes" rows={3} value={form.description} onChange={(e) => set({ description: e.target.value })} /></Field>
      </div>
    </Dialog>
  ) : null;

  const detail = opened ? (
    <Dialog
      title={opened.slot.summary}
      width={440}
      onClose={() => setOpened(null)}
      actions={
        opened.readOnly ? (
          <>
            <Button label={t('Close')} onClick={() => setOpened(null)} />
            {file ? <Button primary label={t('Add to my calendar')} onClick={() => importFile(file.path)} /> : null}
          </>
        ) : (
          <>
            {opened.event?.rrule ? <Button label={t('Delete this one')} onClick={() => removeOpened('this')} /> : null}
            <Button label={opened.event?.rrule ? t('Delete all') : t('Delete')} onClick={() => removeOpened('all')} />
            {opened.event?.attendees?.length ? <Button label={t('Send invitation')} onClick={invite} /> : null}
            <Button primary label={t('Edit')} onClick={() => { setForm(formFrom(opened.event, opened.slot.original)); setOpened(null); }} />
          </>
        )
      }
    >
      <div className="cal-detail" style={{ '--c': opened.slot.colour }}>
        <div className="cal-detail-when">
          <Icon name="clock" size={14} />
          <span>
            {opened.slot.allDay
              ? t('{when}, all day', { when: `${new Date(opened.slot.start).toUTCString().slice(0, 16)}${opened.slot.end - opened.slot.start > DAY ? ` – ${new Date(opened.slot.end - DAY).toUTCString().slice(0, 16)}` : ''}` })
              : `${new Date(opened.slot.start).toLocaleString(undefined, { weekday: 'short', day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' })} – ${clock(opened.slot.end)}`}
          </span>
        </div>
        {opened.event?.rule && opened.event.rule !== 'Does not repeat' ? <div className="cal-detail-line"><Icon name="refresh" size={14} /><span>{opened.event.rule}</span></div> : null}
        {opened.slot.location ? <div className="cal-detail-line"><Icon name="globe" size={14} /><span>{opened.slot.location}</span></div> : null}
        {opened.event?.attendees?.length ? <div className="cal-detail-line"><Icon name="contacts" size={14} /><span>{opened.event.attendees.map((a) => `${a.name || a.email}${a.partstat && a.partstat !== 'NEEDS-ACTION' ? ` (${ANSWERS[a.partstat] ?? a.partstat.toLowerCase()})` : ''}`).join(', ')}</span></div> : null}
        {opened.event?.description ? <p className="cal-detail-notes">{opened.event.description}</p> : null}
        <div className="cal-detail-cal"><span className="dot" />{opened.slot.calendar}</div>
      </div>
    </Dialog>
  ) : null;

  return (
    <AppFrame
      app={app}
      shell={shell}
      title={title}
      menu={appMenu}
      ribbon={
        <Ribbon
          tabs={[{ id: 'home', label: t('Home') }, { id: 'view', label: t('View') }]}
          active={tab}
          onTab={setTab}
          quick={
            <>
              <Button icon="plus" title={t('New event (Ctrl+N)')} onClick={() => newEvent()} />
              <Button icon="clock" title={t('Today (Ctrl+T)')} onClick={() => setCursor(new Date())} />
            </>
          }
        >
          {tab === 'home' ? (
            <>
              <Group label={t('New')}>
                <Button tall icon="plus" label={t('New event')} onClick={() => newEvent()} />
              </Group>
              <Group label={t('Go to')}>
                <Button tall icon="clock" label={t('Today')} onClick={() => setCursor(new Date())} />
                <Button icon="chevronLeft" title={t('Back')} onClick={() => go(-1)} />
                <Button icon="chevronRight" title={t('Forward')} onClick={() => go(1)} />
              </Group>
              <Group label={t('Arrange')}>
                <Button tall icon="file" label={t('Day')} pressed={view === 'day'} onClick={() => setView('day')} />
                <Button tall icon="table" label={t('Week')} pressed={view === 'week'} onClick={() => setView('week')} />
                <Button tall icon="grid" label={t('Month')} pressed={view === 'month'} onClick={() => setView('month')} />
                <Button tall icon="list" label={t('Agenda')} pressed={view === 'agenda'} onClick={() => setView('agenda')} />
              </Group>
              <Group label={t('Files')}>
                <Button tall icon="import" label={t('Import')} title={t('Events from a .ics')} onClick={() => importFile()} />
                <Button tall icon="export" label={t('Export')} title={t('Your calendars as a .ics')} onClick={exportFile} />
              </Group>
              <Group label={t('Accounts')}>
                <Button tall icon="globe" label={t('Accounts')} title={t('Accounts — a CalDAV or CardDAV server (iCloud, Fastmail, Nextcloud and the like) kept in step with this computer')} onClick={() => setAccountsOpen(true)} />
                <Button tall icon="refresh" label={t('Sync')} title={t('Sync every account now')} onClick={syncAll} />
              </Group>
            </>
          ) : (
            <Group label={t('Calendars')}>
              <Button tall icon="plus" label={t('New calendar')} onClick={addCalendar} />
            </Group>
          )}
        </Ribbon>
      }
    >
      <style>{CSS}</style>
      {accountsOpen ? <AccountsDialog shell={shell} kind="calendar" toast={toast} onClose={() => setAccountsOpen(false)} /> : null}
      <Panel width={220} resizable>
        <div className="cal-side">
          <MiniMonth cursor={cursor} today={today} onPick={(d) => { setCursor(d); if (view === 'month') setView('day'); }} />
          <div className="cal-cals">
            <div className="cal-cals-head">{t('My calendars')}</div>
            {calendars.map((c) => (
              <label key={c.id} className="cal-cal" title={c.account ? t('Kept in step with an account') : undefined}>
                <input type="checkbox" checked={c.visible} onChange={() => toggleCalendar(c)} />
                <span className="dot" style={{ background: c.colour }} />
                <span className="n">{c.name}</span>
                <span className="k">{c.count || ''}</span>
              </label>
            ))}
            {file ? (
              <div className="cal-cal file">
                <span className="dot" style={{ background: '#8c6d1f' }} />
                <span className="n">{file.name}</span>
                <Button ghost icon="close" title={t('Close the file')} onClick={() => setFile(null)} />
              </div>
            ) : null}
          </div>
        </div>
      </Panel>
      <Content>
        {file?.invitation ? (
          <div className="cal-invite">
            <Icon name="mail" size={16} />
            <div className="grow">
              <b>{file.invitation.organizer?.name || file.invitation.organizer?.email || t('Someone')}</b> {t('invites you to')} <b>{file.invitation.summary}</b>
              {file.invitation.me?.partstat && file.invitation.me.partstat !== 'NEEDS-ACTION' ? <span> — {t('you answered {answer}', { answer: ANSWERS[file.invitation.me.partstat] ?? file.invitation.me.partstat.toLowerCase() })}</span> : null}
            </div>
            <Button primary icon="check" label={t('Accept')} onClick={() => respond('ACCEPTED')} />
            <Button label={t('Tentative')} onClick={() => respond('TENTATIVE')} />
            <Button label={t('Decline')} onClick={() => respond('DECLINED')} />
          </div>
        ) : file ? (
          <div className="cal-invite">
            <Icon name="file" size={16} />
            <div className="grow">{tn(file.count, '{count} event from', '{count} events from')} <b>{file.name}</b>, {t('shown beside yours and not yet kept.')}</div>
            <Button primary icon="import" label={t('Add to my calendar')} onClick={() => importFile(file.path)} />
          </div>
        ) : null}
        <div className="cal-toolbar">
          <Button icon="chevronLeft" title={t('Back')} onClick={() => go(-1)} />
          <Button icon="chevronRight" title={t('Forward')} onClick={() => go(1)} />
          <h2 className="cal-title">{title}</h2>
          <Spacer />
          <div className="cal-views">
            {['day', 'week', 'month', 'agenda'].map((v) => <button type="button" key={v} className={`cal-view${view === v ? ' on' : ''}`} onClick={() => setView(v)}>{VIEW_NAMES[v]}</button>)}
          </div>
        </div>
        <div className="cal-body">{body}</div>
      </Content>
      {editor}
      {detail}
    </AppFrame>
  );
}

function MiniMonth({ cursor, today, onPick }) {
  const [shown, setShown] = useState(() => new Date(cursor.getFullYear(), cursor.getMonth(), 1));
  useEffect(() => setShown(new Date(cursor.getFullYear(), cursor.getMonth(), 1)), [cursor]);
  const from = startOfWeek(shown);
  const cells = Array.from({ length: 42 }, (_, i) => addDays(from, i));
  return (
    <div className="cal-mini">
      <div className="cal-mini-head">
        <button type="button" onClick={() => setShown(new Date(shown.getFullYear(), shown.getMonth() - 1, 1))}>‹</button>
        <span>{MONTHS_SHORT[shown.getMonth()]} {shown.getFullYear()}</span>
        <button type="button" onClick={() => setShown(new Date(shown.getFullYear(), shown.getMonth() + 1, 1))}>›</button>
      </div>
      <div className="cal-mini-grid">
        {DAY_LETTERS.map((d, i) => <span key={i} className="h">{d}</span>)}
        {cells.map((d) => (
          <button type="button" key={isoDate(d)} className={`${d.getMonth() !== shown.getMonth() ? 'other' : ''}${sameDay(d, today) ? ' today' : ''}${sameDay(d, cursor) ? ' on' : ''}`} onClick={() => onPick(d)}>{d.getDate()}</button>
        ))}
      </div>
    </div>
  );
}

const CSS = `
.cal-side { display: flex; flex-direction: column; gap: 12px; padding: 10px; }
.cal-mini { font-size: 12px; }
.cal-mini-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 4px; font-weight: 600; }
.cal-mini-head button { border: 0; background: none; font: inherit; color: var(--ink-2); cursor: pointer; padding: 2px 6px; }
.cal-mini-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 1px; }
.cal-mini-grid .h { text-align: center; color: var(--ink-3); font-size: 10px; padding: 2px 0; }
.cal-mini-grid button { border: 0; background: none; font: inherit; font-size: 11px; padding: 3px 0; border-radius: 4px; color: var(--ink); cursor: pointer; }
.cal-mini-grid button.other { color: var(--ink-3); }
.cal-mini-grid button.today { font-weight: 700; color: var(--accent); }
.cal-mini-grid button.on { background: var(--selected); }
.cal-mini-grid button:hover { background: var(--hover); }
.cal-cals-head { font-size: 11px; font-weight: 600; color: var(--ink-3); letter-spacing: .04em; text-transform: uppercase; margin: 6px 0 4px; }
.cal-cal { display: flex; align-items: center; gap: 8px; padding: 4px 2px; font-size: 13px; cursor: pointer; }
.cal-cal .dot { width: 10px; height: 10px; border-radius: 50%; flex: none; }
.cal-cal .n { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.cal-cal .k { color: var(--ink-3); font-size: 11px; }
.cal-invite { display: flex; align-items: center; gap: 10px; margin: 10px 14px 0; padding: 10px 12px; border-radius: var(--r-2); background: var(--sunken); font-size: 13px; }
.cal-invite .grow { flex: 1; }
.cal-toolbar { display: flex; align-items: center; gap: 6px; padding: 10px 14px 6px; }
.cal-title { margin: 0 0 0 6px; font-size: 18px; font-weight: 600; }
.cal-views { display: flex; border: 1px solid var(--line); border-radius: var(--r-2); overflow: hidden; }
.cal-view { border: 0; background: none; padding: 5px 12px; font: inherit; font-size: 12.5px; color: var(--ink-2); cursor: pointer; }
.cal-view.on { background: var(--selected); color: var(--ink); font-weight: 600; }
.cal-body { flex: 1; min-height: 0; overflow: auto; padding: 0 14px 14px; display: flex; flex-direction: column; }
.cal-month { display: flex; flex-direction: column; flex: 1; min-height: 540px; }
.cal-month-head { display: grid; grid-template-columns: repeat(7, 1fr); font-size: 11px; font-weight: 600; color: var(--ink-3); padding: 4px 0; }
.cal-month-head div { padding-left: 6px; }
.cal-month-grid { display: grid; grid-template-columns: repeat(7, 1fr); grid-auto-rows: 1fr; flex: 1; border-top: 1px solid var(--line); border-left: 1px solid var(--line); }
.cal-day { border-right: 1px solid var(--line); border-bottom: 1px solid var(--line); padding: 4px; min-height: 84px; display: flex; flex-direction: column; gap: 2px; cursor: cell; }
.cal-day.other { background: var(--sunken); color: var(--ink-3); }
.cal-day.today .cal-day-n { background: var(--accent); color: white; }
.cal-day-n { align-self: flex-end; font-size: 12px; font-weight: 600; padding: 1px 6px; border-radius: 10px; }
.cal-chip { display: flex; align-items: center; gap: 4px; border: 0; border-left: 3px solid var(--c); border-radius: 3px; background: color-mix(in srgb, var(--c) 14%, var(--surface)); color: var(--ink); font: inherit; font-size: 11.5px; padding: 1px 5px; text-align: left; cursor: pointer; overflow: hidden; white-space: nowrap; }
.cal-chip.allday { background: var(--c); color: white; border-left-color: var(--c); }
.cal-chip .t { color: var(--ink-3); flex: none; }
.cal-chip.allday .t { color: white; }
.cal-chip .s { overflow: hidden; text-overflow: ellipsis; }
.cal-more { border: 0; background: none; font: inherit; font-size: 11px; color: var(--accent); text-align: left; padding: 0 5px; cursor: pointer; }
.cal-time { display: flex; flex-direction: column; flex: 1; min-height: 0; }
.cal-time-head, .cal-allday, .cal-hours { display: grid; grid-template-columns: 56px repeat(auto-fit, minmax(0, 1fr)); }
.cal-gutter { font-size: 10.5px; color: var(--ink-3); text-align: right; padding-right: 6px; }
.cal-col-head { padding: 4px 6px; font-size: 12px; color: var(--ink-2); border-bottom: 1px solid var(--line); }
.cal-col-head.today .dd { background: var(--accent); color: white; border-radius: 10px; padding: 0 6px; }
.cal-col-head .dd { font-weight: 700; }
.cal-allday { min-height: 28px; border-bottom: 1px solid var(--line); }
.cal-allday .cal-gutter { padding-top: 6px; }
.cal-allday-cell { display: flex; flex-direction: column; gap: 2px; padding: 3px; border-left: 1px solid var(--line); cursor: cell; }
.cal-hours { flex: 1; overflow: auto; position: relative; }
.cal-hour { transform: translateY(-7px); }
.cal-col { position: relative; border-left: 1px solid var(--line); cursor: cell; }
.cal-col.today { background: color-mix(in srgb, var(--accent) 4%, transparent); }
.cal-line { position: absolute; left: 0; right: 0; border-top: 1px solid var(--line); }
.cal-now { position: absolute; left: 0; right: 0; border-top: 2px solid var(--bad); z-index: 2; }
.cal-block { position: absolute; box-sizing: border-box; display: flex; flex-direction: column; gap: 1px; border: 0; border-left: 3px solid var(--c); border-radius: 4px; background: color-mix(in srgb, var(--c) 18%, var(--surface)); color: var(--ink); font: inherit; font-size: 11.5px; padding: 3px 6px; text-align: left; overflow: hidden; cursor: pointer; z-index: 1; }
.cal-block .s { font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.cal-block .t { color: var(--ink-3); font-size: 10.5px; }
.cal-agenda { display: flex; flex-direction: column; gap: 6px; padding-top: 6px; }
.cal-agenda-day { display: grid; grid-template-columns: 64px 1fr; gap: 12px; padding: 8px 0; border-top: 1px solid var(--line); }
.cal-agenda-date { display: flex; flex-direction: column; align-items: center; line-height: 1.1; }
.cal-agenda-date .dn { font-size: 11px; color: var(--ink-3); }
.cal-agenda-date .dd { font-size: 22px; font-weight: 700; }
.cal-agenda-date .dm { font-size: 11px; color: var(--ink-3); }
.cal-agenda-day.today .cal-agenda-date .dd { color: var(--accent); }
.cal-agenda-list { display: flex; flex-direction: column; gap: 4px; }
.cal-agenda-item { display: grid; grid-template-columns: 110px 1fr auto; gap: 10px; align-items: baseline; border: 0; border-left: 3px solid var(--c); border-radius: 4px; background: var(--sunken); font: inherit; font-size: 13px; padding: 6px 10px; text-align: left; color: var(--ink); cursor: pointer; }
.cal-agenda-item .t { color: var(--ink-3); font-size: 12px; }
.cal-agenda-item .s { font-weight: 500; }
.cal-agenda-item .l { color: var(--ink-3); font-size: 12px; }
.cal-form { display: flex; flex-direction: column; gap: 10px; padding: 2px 0 8px; }
.cal-title.rw-input, .cal-form .cal-title { font-size: 18px; font-weight: 600; }
.cal-check { display: flex; align-items: center; gap: 6px; font-size: 13px; }
.cal-when { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
.cal-dt { display: grid; grid-template-columns: 1fr 96px; gap: 6px; }
.cal-notes { width: 100%; resize: vertical; font: inherit; }
.cal-detail { display: flex; flex-direction: column; gap: 8px; font-size: 13px; padding: 2px 0 6px; border-left: 4px solid var(--c); padding-left: 12px; }
.cal-detail-when, .cal-detail-line { display: flex; align-items: center; gap: 8px; color: var(--ink-2); }
.cal-detail-notes { margin: 4px 0; white-space: pre-line; }
.cal-detail-cal { display: flex; align-items: center; gap: 6px; color: var(--ink-3); font-size: 12px; }
.cal-detail-cal .dot { width: 8px; height: 8px; border-radius: 50%; background: var(--c); }
`;
