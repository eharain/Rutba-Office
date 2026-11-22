// The calendar's shown view, on paper.
//
// A month prints as its grid, six weeks of seven days with every event of a
// day written in its square: paper has no "3 more" to click, so a square
// grows to hold them all and a busy month runs onto a second sheet between
// two weeks, never through one. A week prints as its seven days side by
// side, a day and the agenda as a list under each date.
//
// The events are the ones the window shows (a hidden calendar's are left
// out by the caller), at the times this computer's clock gives them.

import { t } from '@rutba/office-ui/messages';
import { pageHtml, esc } from '../print-page.js';

const DAY = 86400000;
const startOfDay = (d) => new Date(d.getFullYear(), d.getMonth(), d.getDate());
const addDays = (d, n) => new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
const pad = (n) => String(n).padStart(2, '0');
const isoDate = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const clock = (ms) => {
  const d = new Date(ms);
  return `${d.getHours()}:${pad(d.getMinutes())}`;
};

/** An all-day slot's instants are UTC midnights of calendar dates; place it on those dates locally. */
export function slotDays(slot) {
  if (slot.allDay) {
    const s = new Date(slot.start);
    const e = new Date(slot.end);
    const first = new Date(s.getUTCFullYear(), s.getUTCMonth(), s.getUTCDate());
    const last = new Date(e.getUTCFullYear(), e.getUTCMonth(), e.getUTCDate() - 1);
    return { first, last: last < first ? first : last };
  }
  return { first: startOfDay(new Date(slot.start)), last: startOfDay(new Date(Math.max(slot.start, slot.end - 1))) };
}

const order = (a, b) => Number(b.allDay) - Number(a.allDay) || a.start - b.start;

/** Each day's events between two dates, by the day's ISO date; an event over several days is on each of them, saying which. */
export function eventsByDay(slots, from, to) {
  const days = new Map();
  for (const s of slots || []) {
    const { first, last } = slotDays(s);
    const count = Math.round((last - first) / DAY) + 1;
    for (let n = 0, d = first; d <= last; n++, d = addDays(d, 1)) {
      if (d < from || d >= to) continue;
      const k = isoDate(d);
      if (!days.has(k)) days.set(k, []);
      days.get(k).push(count > 1 ? { ...s, dayOf: { n: n + 1, count } } : s);
    }
  }
  for (const list of days.values()) list.sort(order);
  return days;
}

const colourOf = (s) => (/^#[0-9a-f]{3,8}$/i.test(String(s.colour || '')) ? s.colour : '#5b6b82');
const whenOf = (s) => (s.dayOf ? t('calendar.day-n-count', s.dayOf) : s.allDay ? t('calendar.all-day-2') : `${clock(s.start)} – ${clock(s.end)}`);

/** An event in a square of the grid: its start, then what it is. */
const chip = (s) => `<div class="ev" style="border-inline-start-color:${colourOf(s)}">${s.allDay || s.dayOf ? '' : `<span class="t">${clock(s.start)}</span> `}<span class="s">${esc(s.summary || t('common.untitled'))}</span></div>`;

/** An event on a line of a list: when, what and where. */
const line = (s) => `<tr class="ev-line"><td class="t">${esc(whenOf(s))}</td><td class="s"><span class="dot" style="background:${colourOf(s)}"></span>${esc(s.summary || t('common.untitled'))}${s.location ? `<span class="l"> · ${esc(s.location)}</span>` : ''}</td></tr>`;

const names = (lang) => {
  const of = (d, options) => {
    try {
      return d.toLocaleDateString(lang, options);
    } catch {
      return d.toLocaleDateString('en', options);
    }
  };
  return {
    short: (d) => of(d, { weekday: 'short' }),
    long: (d) => of(d, { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' }),
    month: (d) => of(d, { month: 'short' }),
  };
};

function monthBody({ from, cursor, slots, lang }) {
  const name = names(lang);
  const days = Array.from({ length: 42 }, (_, i) => addDays(from, i));
  const byDay = eventsByDay(slots, from, addDays(from, 42));
  const head = days.slice(0, 7).map((d) => `<th scope="col">${esc(name.short(d))}</th>`).join('');
  const weeks = [];
  for (let w = 0; w < 6; w++) {
    const cells = days.slice(w * 7, w * 7 + 7).map((d) => {
      const other = d.getMonth() !== cursor.getMonth();
      const list = byDay.get(isoDate(d)) || [];
      return `<td class="${other ? 'other' : ''}"><div class="n">${d.getDate()}${d.getDate() === 1 ? ` ${esc(name.month(d))}` : ''}</div>${list.map(chip).join('')}</td>`;
    });
    weeks.push(`<tr>${cells.join('')}</tr>`);
  }
  return `<table class="month"><thead><tr>${head}</tr></thead><tbody>${weeks.join('')}</tbody></table>`;
}

function weekBody({ from, slots, lang }) {
  const name = names(lang);
  const days = Array.from({ length: 7 }, (_, i) => addDays(from, i));
  const byDay = eventsByDay(slots, from, addDays(from, 7));
  const head = days.map((d) => `<th scope="col">${esc(name.short(d))} ${d.getDate()}</th>`).join('');
  const cells = days.map((d) => {
    const list = byDay.get(isoDate(d)) || [];
    return `<td>${list.map((s) => `<div class="ev" style="border-inline-start-color:${colourOf(s)}"><div class="t">${esc(whenOf(s))}</div><div class="s">${esc(s.summary || t('common.untitled'))}</div>${s.location ? `<div class="l">${esc(s.location)}</div>` : ''}</div>`).join('')}</td>`;
  });
  return `<table class="month week"><thead><tr>${head}</tr></thead><tbody><tr>${cells.join('')}</tr></tbody></table>`;
}

function listBody({ from, to, slots, lang, everyDay = false }) {
  const name = names(lang);
  const byDay = eventsByDay(slots, from, to);
  const keys = everyDay ? [isoDate(from)] : [...byDay.keys()].sort();
  if (!keys.some((k) => byDay.get(k)?.length)) return `<p class="none">${esc(t('calendar.print.no-events'))}</p>`;
  return keys.map((k) => {
    const d = new Date(`${k}T00:00:00`);
    const list = byDay.get(k) || [];
    return `<section class="day"><h2>${esc(name.long(d))}</h2><table class="lines"><tbody>${list.map(line).join('')}</tbody></table></section>`;
  }).join('\n');
}

const CSS = `
  h2 { font-size: 11pt; margin: 0 0 1.5mm; padding-bottom: 1mm; border-bottom: 0.25mm solid #9aa0a6; break-after: avoid; }
  .month { width: 100%; table-layout: fixed; }
  .month th { font-size: 9pt; font-weight: 600; color: #444; padding: 1mm; border: 0.25mm solid #9aa0a6; background: #f1f3f5; }
  .month td { vertical-align: top; border: 0.25mm solid #9aa0a6; padding: 1mm; height: 22mm; font-size: 8pt; line-height: 1.3; }
  .month tr { break-inside: avoid; }
  .month td.other { background: #f7f8f9; color: #777; }
  .month .n { font-weight: 600; font-size: 9pt; margin-bottom: 0.8mm; }
  .month .ev { border-inline-start: 0.8mm solid #5b6b82; padding-inline-start: 1mm; margin-bottom: 0.8mm; break-inside: avoid; }
  .month .ev .t { color: #444; }
  .week td { height: 60mm; font-size: 8.5pt; }
  .week .ev .s { font-weight: 600; }
  .week .ev .l { color: #555; }
  .day { margin: 0 0 5mm; }
  .lines { width: 100%; }
  .lines td { vertical-align: top; padding: 0.8mm 0; }
  .lines tr { break-inside: avoid; }
  .lines .t { width: 32mm; color: #444; white-space: nowrap; }
  .lines .dot { display: inline-block; width: 2.2mm; height: 2.2mm; border-radius: 50%; margin-inline-end: 1.5mm; }
  .lines .l { color: #555; }
  .none { color: #555; }
`;

/**
 * The page for the view the window shows. `view` is day, week, month or
 * agenda; `from` and `to` are the range it shows; `cursor` the date it is
 * on; `slots` the events drawn; `title` the heading the window gives it.
 */
export function calendarPrintHtml({ view = 'month', from, to, cursor = from, slots = [], title = '', options = {}, lang = 'en', dir = 'ltr' } = {}) {
  const body = view === 'month' ? monthBody({ from, cursor, slots, lang })
    : view === 'week' ? weekBody({ from, slots, lang })
      : listBody({ from, to, slots, lang, everyDay: view === 'day' });
  return pageHtml({ title, options, css: CSS, body: `<h1>${esc(title)}</h1>\n${body}`, lang, dir });
}
