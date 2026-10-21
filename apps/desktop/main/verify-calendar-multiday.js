// Calendar: events over several days (owner, 2026-10-22, D-2026-10-22-02).
// In the month an all-day event of three days and a timed one from one
// morning to the afternoon two days later are each one bar across their
// days, in the same row on each day, joined from day to day and named where
// they start (and again where a week starts); a single-day event the same
// day sits below them. The list of the next thirty days shows each on every
// one of its days, with which day it is. The fixture is a file the window
// shows beside the person's own calendars. Run alone with
// RUTBA_VERIFY_ONLY=calendar-multiday.

import fs from 'node:fs';
import path from 'node:path';

const pad = (n) => String(n).padStart(2, '0');
const isoDate = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const icsDate = (d) => isoDate(d).replace(/-/g, '');

function fixture(dir) {
  const today = new Date();
  const day = (n) => new Date(today.getFullYear(), today.getMonth(), today.getDate() + n);
  const fair = [day(1), day(2), day(3)];
  const conference = [day(2), day(3), day(4)];
  const lines = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Rutba//multi-day checks//EN', 'X-WR-CALNAME:Multi-day checks',
    'BEGIN:VEVENT', 'UID:fair@calendar.checks', `DTSTAMP:${icsDate(today)}T080000`, `DTSTART;VALUE=DATE:${icsDate(fair[0])}`, `DTEND;VALUE=DATE:${icsDate(day(4))}`, 'SUMMARY:Trade fair', 'END:VEVENT',
    'BEGIN:VEVENT', 'UID:conference@calendar.checks', `DTSTAMP:${icsDate(today)}T080000`, `DTSTART:${icsDate(conference[0])}T090000`, `DTEND:${icsDate(conference[2])}T170000`, 'SUMMARY:Conference', 'END:VEVENT',
    'BEGIN:VEVENT', 'UID:dentist@calendar.checks', `DTSTAMP:${icsDate(today)}T080000`, `DTSTART:${icsDate(day(3))}T150000`, `DTEND:${icsDate(day(3))}T153000`, 'SUMMARY:Dentist', 'END:VEVENT',
    'END:VCALENDAR', ''];
  const own = path.join(dir, 'calendar-multiday-fixtures');
  fs.mkdirSync(own, { recursive: true });
  const file = path.join(own, 'multiday.ics');
  fs.writeFileSync(file, lines.join('\r\n'));
  return { file, fair: fair.map(isoDate), conference: conference.map(isoDate), dentist: isoDate(day(3)) };
}

/** Each piece of a named event on the month: its day, its row's offset in the box, its joins, whether it shows its name. */
const PIECES = (name) => `[...document.querySelectorAll('.cal-month-grid .cal-day')].flatMap((c) => [...c.querySelectorAll('.cal-chip')].filter((n) => n.querySelector('.s')?.textContent === ${JSON.stringify(name)}).map((n) => {
  const index = [...c.parentElement.children].indexOf(c);
  return { date: c.dataset.date, column: index % 7, week: Math.floor(index / 7), top: Math.round(n.getBoundingClientRect().top - c.getBoundingClientRect().top),
    span: n.classList.contains('span'), before: n.classList.contains('from-before'), after: n.classList.contains('to-after'), named: getComputedStyle(n.querySelector('.s')).visibility !== 'hidden',
    reachesLeft: Math.round(n.getBoundingClientRect().left - c.getBoundingClientRect().left - c.clientLeft), reachesRight: Math.round(c.getBoundingClientRect().left + c.clientLeft + c.clientWidth - n.getBoundingClientRect().right) };
}))`;

export async function verifyCalendarMultiDay({ open, check, until, wait }, { dir }) {
  const fx = fixture(dir);
  try {
    const win = await open('calendar', fx.file);
    const js = (code) => win.webContents.executeJavaScript(code);
    win.setContentSize(1456, 910);
    await until(() => js(`${PIECES('Conference')}.length >= 3`), 'the month with the file\'s events', 10000).catch(() => {});
    await wait(300);
    const bars = {};
    const bad = [];
    for (const [name, dates] of [['Trade fair', fx.fair], ['Conference', fx.conference]]) {
      const pieces = (await js(PIECES(name))).filter((p) => dates.includes(p.date));
      bars[name] = pieces;
      if (pieces.length !== 3) { bad.push(`${name}: ${pieces.length} pieces, not 3`); continue; }
      pieces.forEach((p, i) => {
        const first = i === 0 || p.column === 0;
        if (!p.span) bad.push(`${name} ${p.date}: not a piece of a bar`);
        if (p.before !== (i > 0 && p.column > 0)) bad.push(`${name} ${p.date}: joined to the day before ${p.before}`);
        if (p.after !== (i < 2 && p.column < 6)) bad.push(`${name} ${p.date}: joined to the day after ${p.after}`);
        if (p.named !== first) bad.push(`${name} ${p.date}: its name shown ${p.named}`);
        if (p.after && p.reachesRight > 1) bad.push(`${name} ${p.date}: stops ${p.reachesRight} px short of the next day`);
        if (p.before && p.reachesLeft > 1) bad.push(`${name} ${p.date}: starts ${p.reachesLeft} px into its day`);
      });
      // One row through each week it crosses.
      for (const w of new Set(pieces.map((p) => p.week))) {
        const tops = new Set(pieces.filter((p) => p.week === w).map((p) => p.top));
        if (tops.size !== 1) bad.push(`${name}: rows ${[...tops].join(', ')} in one week`);
      }
    }
    const dentist = (await js(PIECES('Dentist'))).find((p) => p.date === fx.dentist);
    const below = dentist && bars.Conference?.length && bars['Trade fair']?.length
      && dentist.top > Math.max(...[...bars.Conference, ...bars['Trade fair']].filter((p) => p.date === fx.dentist).map((p) => p.top));
    if (!below) bad.push(`the dentist is not under the bars on ${fx.dentist}`);
    check('calendar: an event over several days is one bar across them in the month, the same row each day, joined from day to day and named where it starts or a week starts; a day\'s other events sit below it',
      bad.length === 0, bad.join('; ') || JSON.stringify(bars.Conference.map((p) => [p.date, p.top, p.before, p.after, p.named])));

    // The next thirty days: each of its days, which day it is.
    // The views in their order: day, week, month, and the next thirty days.
    await js(`document.querySelectorAll('.cal-view')[3]?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.cal-agenda'))`), 'the list of the next thirty days', 5000).catch(() => {});
    const listed = await js(`[...document.querySelectorAll('.cal-agenda-item')].filter((n) => /Trade fair|Conference/.test(n.querySelector('.s')?.textContent || '')).map((n) => [n.querySelector('.s').textContent, n.querySelector('.t').textContent])`);
    const want = ['Day 1 of 3', 'Day 2 of 3', 'Day 3 of 3'];
    const of = (name) => listed.filter(([s]) => s === name).map(([, t]) => t);
    check('calendar: the next thirty days list an event over several days on each of its days, saying which day it is',
      JSON.stringify(of('Trade fair')) === JSON.stringify(want) && JSON.stringify(of('Conference')) === JSON.stringify(want), JSON.stringify(listed));
    if (process.env.RUTBA_VERIFY_CAPTURE) {
      await js(`document.querySelectorAll('.cal-view')[2]?.click(), 1`);
      await wait(500);
      win.webContents.invalidate();
      await wait(700);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'calendar-multiday.png'), (await win.webContents.capturePage()).toPNG());
    }
  } catch (err) {
    check('calendar: the multi-day checks ran to the end', false, err?.message || String(err));
  }
}
