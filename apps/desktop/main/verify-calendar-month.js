// Calendar's month view shows each day's events as Outlook's does: as many
// as the day's box has room for at the window's height, measured rather than
// a fixed count, and "+N more" at the foot of the box only when something is
// left out. "+N more" opens the day's events in a small list over the month,
// which names the day (it opens the day view) and opens any of the events.
// Made smaller or larger, the window shows fewer or more of a busy day's
// events. Looked at with a month of one busy day (eight events) and quiet
// days at 1300x700 and 1680x1050, light and dark, and with the Hijri day
// beside each date; captures go to
// RUTBA_VERIFY_CAPTURE as calendar-month-*.png. The fixture is a file the
// window shows beside the person's own calendars and never keeps. Run alone
// with RUTBA_VERIFY_ONLY=calendar-month.

import fs from 'node:fs';
import path from 'node:path';

const SIZES = [[1300, 700], [1680, 1050]];
const pad = (n) => String(n).padStart(2, '0');
const isoDate = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const icsDate = (d) => isoDate(d).replace(/-/g, '');
const icsTime = (d, h, m = 0) => `${icsDate(d)}T${pad(h)}${pad(m)}00`;

/** The busy day (today) and the quiet ones, all in this month so all on its grid. */
function fixture(dir) {
  const today = new Date();
  const day = (n) => new Date(today.getFullYear(), today.getMonth(), n);
  const quiet = Array.from({ length: 28 }, (_, i) => i + 1).filter((n) => n !== today.getDate()).map(day);
  const events = [
    [today, 'Quarter close', null, null],
    [today, 'Stand-up', [8, 30], [8, 45]],
    [today, 'Supplier call', [9, 0], [9, 30]],
    [today, 'Budget review', [10, 0], [11, 30]],
    [today, 'Design critique', [11, 30], [12, 0]],
    [today, 'Lunch with Sara', [12, 30], [13, 30]],
    [today, 'Interview: shop manager', [14, 30], [15, 30]],
    [today, 'Train home', [17, 15], [18, 0]],
    [quiet[3], 'Dentist', [9, 0], [9, 30]],
    [quiet[3], 'Team lunch', [12, 30], [13, 30]],
    [quiet[3], 'Book club', [19, 0], [20, 30]],
    [quiet[7], 'Gym', [7, 0], [8, 0]],
    [quiet[7], 'Parents evening', [18, 0], [19, 0]],
    [quiet[10], 'Car service', [8, 30], [9, 30]],
    [quiet[14], 'School run', null, null],
    [quiet[18], 'Payroll', [11, 0], [11, 30]],
    [quiet[22], 'Haircut', [16, 0], [16, 30]],
  ];
  const lines = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Rutba//month checks//EN', 'X-WR-CALNAME:Month checks'];
  events.forEach(([d, summary, from, to], i) => {
    const next = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1);
    lines.push('BEGIN:VEVENT', `UID:month-${i}@calendar.checks`, `DTSTAMP:${icsTime(today, 8)}`,
      from ? `DTSTART:${icsTime(d, ...from)}` : `DTSTART;VALUE=DATE:${icsDate(d)}`,
      to ? `DTEND:${icsTime(d, ...to)}` : `DTEND;VALUE=DATE:${icsDate(next)}`,
      `SUMMARY:${summary}`, 'END:VEVENT');
  });
  lines.push('END:VCALENDAR', '');
  const own = path.join(dir, 'calendar-month-fixtures');
  fs.mkdirSync(own, { recursive: true });
  const file = path.join(own, 'month.ics');
  fs.writeFileSync(file, lines.join('\r\n'));
  const byDay = new Map();
  for (const [d, summary] of events) byDay.set(isoDate(d), [...(byDay.get(isoDate(d)) || []), summary]);
  return { file, today: isoDate(today), byDay };
}

/** Every day box on the month: what it shows, what it says it leaves out, and how it sits in its box. */
const MONTH = `(() => {
  const grid = document.querySelector('.cal-month-grid');
  if (!grid) return null;
  return [...grid.querySelectorAll('.cal-day')].map((c) => {
    const r = c.getBoundingClientRect();
    const cs = getComputedStyle(c);
    const inner = r.top + c.clientTop + c.clientHeight - (parseFloat(cs.paddingBottom) || 0);
    const gap = parseFloat(cs.rowGap) || 0;
    const top = c.querySelector('.cal-day-top');
    const chips = [...c.querySelectorAll('.cal-chip')];
    const more = c.querySelector('.cal-more');
    const m = more ? more.getBoundingClientRect() : null;
    const last = chips.length ? chips[chips.length - 1].getBoundingClientRect() : null;
    const above = last ? last.bottom : top ? top.getBoundingClientRect().bottom : r.top;
    const chipH = chips.length ? chips[0].getBoundingClientRect().height : 0;
    return {
      date: c.dataset.date || '',
      today: c.classList.contains('today'),
      titles: chips.map((n) => n.querySelector('.s')?.textContent || ''),
      hidden: more ? Number((more.textContent.match(/\\d+/) || [0])[0]) : 0,
      moreInTop: Boolean(more && top && top.contains(more)),
      moreBelow: !more || !last || m.top >= last.bottom - 0.5,
      moreAtFoot: !more || Math.abs(inner - m.bottom) <= 2,
      // Room left between the last event and "+N more" (or the foot of the box).
      free: Math.round(((m ? m.top : inner) - above) * 10) / 10,
      clipped: chips.filter((n) => { const b = n.getBoundingClientRect(); return b.bottom > inner + 0.5 || b.right > r.right + 0.5; }).length + (m && m.bottom > inner + 0.5 ? 1 : 0),
      chipH, gap,
    };
  });
})()`;

export async function verifyCalendarMonth({ open, check, until, wait }, { dir }) {
  const shots = process.env.RUTBA_VERIFY_CAPTURE;
  const fx = fixture(dir);
  const snap = async (win, name) => {
    if (!shots) return;
    win.webContents.invalidate();
    await wait(700);
    fs.writeFileSync(path.join(shots, name), (await win.webContents.capturePage()).toPNG());
  };
  try {
    const win = await open('calendar', fx.file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const size = async (w, h) => {
      win.setContentSize(w, h);
      await until(() => js('[innerWidth, innerHeight]').then(([iw, ih]) => Math.abs(iw - w) <= 2 && Math.abs(ih - h) <= 2), `the window at ${w}x${h}`, 5000).catch(() => {});
    };
    const busy = (cells) => cells?.find((c) => c.date === fx.today) || cells?.find((c) => c.today) || null;
    // The busy day settled at this size: its events and "+N more" add up to its eight (or more, beside the person's own).
    const settled = (cells) => { const b = busy(cells); return Boolean(b && b.titles.length + b.hidden >= 8 && b.hidden > 0); };
    await until(() => js(`document.querySelectorAll('.cal-month-grid .cal-chip').length >= 6`), 'the month with the file\'s events', 10000).catch(() => {});

    const shown = {};
    // The month's own measure at each size: events that fit when none is left out / beside "+N more".
    const fits = {};
    const bad = [];
    for (const [w, h] of SIZES) {
      await size(w, h);
      await until(async () => settled(await js(MONTH)), `the month at ${w}x${h}`, 5000).catch(() => {});
      await wait(300);
      for (const theme of ['light', 'dark']) {
        await js(`document.documentElement.dataset.theme = ${JSON.stringify(theme)}, 1`);
        await wait(200);
        const cells = (await js(MONTH)) || [];
        const b = busy(cells);
        // An event's height, from any day that shows one.
        const chipH = Math.max(0, ...cells.map((c) => c.chipH));
        const at = `${w}x${h} ${theme}`;
        if (!b) { bad.push(`${at}: no busy day`); continue; }
        shown[`${w}x${h}`] = b.titles.length;
        fits[`${w}x${h}`] = await js(`document.querySelector(".cal-month-grid")?.dataset.fit || ""`);
        if (!b.hidden) bad.push(`${at}: the busy day says nothing is left out (${b.titles.length} shown)`);
        if (b.titles.length + b.hidden < 8) bad.push(`${at}: the busy day shows ${b.titles.length} and "+${b.hidden} more", short of its 8`);
        if (b.moreInTop || !b.moreBelow || !b.moreAtFoot) bad.push(`${at}: "+N more" is not at the foot of the busy day (in the date's line ${b.moreInTop}, under its events ${b.moreBelow}, at the foot ${b.moreAtFoot})`);
        // As many as fit: no room is left for another event (and the gap under it) above "+N more".
        if (b.hidden && b.free >= chipH + 2 * b.gap - 0.1) bad.push(`${at}: the busy day leaves ${b.free} px empty above "+N more", room for another ${chipH} px event`);
        if (!b.titles.length) bad.push(`${at}: the busy day shows none of its events`);
        // What the month measures with is never taken for a day's event.
        const stray = await js(`[...document.querySelectorAll(".cal-month .cal-chip, .cal-month .cal-day-h")].filter((n) => !n.closest(".cal-day, .cal-peek")).length`);
        if (stray) bad.push(`${at}: ${stray} event or date outside any day`);
        for (const c of cells) {
          const want = fx.byDay.get(c.date) || [];
          if (c.clipped) bad.push(`${at}: ${c.date || '?'} cuts off ${c.clipped} of what it shows`);
          if (c.date !== fx.today && want.length && c.titles.length + c.hidden < want.length) bad.push(`${at}: ${c.date} shows ${c.titles.length} and "+${c.hidden} more" of its ${want.length}`);
          if (c.date !== fx.today && c.hidden && c.free >= chipH + 2 * c.gap - 0.1) bad.push(`${at}: ${c.date} hides events with ${c.free} px to spare`);
          if (c.hidden && (c.moreInTop || !c.moreAtFoot)) bad.push(`${at}: ${c.date}'s "+N more" is not at its foot`);
          if (!c.hidden && want.some((s) => !c.titles.includes(s))) bad.push(`${at}: ${c.date} leaves out ${want.filter((s) => !c.titles.includes(s)).join(', ')} and says nothing`);
        }
        await snap(win, `calendar-month-${w}x${h}-${theme}.png`);
      }
      await js(`document.documentElement.dataset.theme = 'light', 1`);
    }
    check('calendar-month: each day shows as many of its events as its box has room for, "+N more" at the foot only when some are left out, nothing cut off, at 1300x700 and 1680x1050, light and dark',
      !bad.length, bad.slice(0, 4).join('; ') || `busy day shows ${JSON.stringify(shown)}, the month measures ${JSON.stringify(fits)}`);

    // With the Hijri day beside each date the date's line is measured with it,
    // and nothing is cut off; the choice is kept, so it is put back after.
    await size(1300, 700);
    const hijriButton = `[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((b) => b.textContent.trim() === 'Hijri')`;
    const hijriWas = await js(`${hijriButton}?.getAttribute('aria-pressed') === 'true'`);
    const hijriBad = [];
    if (!hijriWas) await js(`${hijriButton}?.click(), 1`);
    const hijriOn = await until(() => js(`document.querySelectorAll('.cal-month-grid .cal-day .cal-day-h').length === 42`), 'the Hijri days', 4000).then(() => true, () => false);
    await until(async () => settled(await js(MONTH)), 'the month with the Hijri days', 4000).catch(() => {});
    for (const c of (await js(MONTH)) || []) {
      if (c.clipped) hijriBad.push(`${c.date} cuts off ${c.clipped}`);
      if (c.hidden && (c.moreInTop || !c.moreAtFoot)) hijriBad.push(`${c.date}'s "+N more" is not at its foot`);
      if (c.date === fx.today && (!c.titles.length || c.titles.length + c.hidden < 8)) hijriBad.push(`the busy day shows ${c.titles.length} and "+${c.hidden} more"`);
    }
    await snap(win, 'calendar-month-hijri-1300x700-light.png');
    if (!hijriWas) {
      await js(`${hijriButton}?.click(), 1`);
      await until(() => js(`!document.querySelector('.cal-month-grid .cal-day-h')`), 'the Hijri days to go', 4000).catch(() => {});
    }
    check('calendar-month: with the Hijri day beside each date, each day still shows what fits and "+N more" at its foot, nothing cut off',
      hijriOn && !hijriBad.length, hijriBad.slice(0, 4).join('; ') || `Hijri ${hijriOn ? 'shown' : 'never shown'}`);

    // Smaller and back: the count follows the window, measured each time.
    await size(1300, 700);
    const small = await until(async () => { const b = busy(await js(MONTH)); return b && b.titles.length === shown['1300x700']; }, 'the busy day at 1300x700 again', 5000).then(() => true, () => false);
    const big = shown['1680x1050'] || 0;
    check('calendar-month: made larger the busy day shows more of its events, and fewer again made smaller',
      big > (shown['1300x700'] || 0) && small, `1300x700 shows ${shown['1300x700']}, 1680x1050 shows ${big}, back at 1300x700 ${small ? 'the same again' : 'not'}`);

    // "+N more" opens the day's list over the month.
    const openPeek = async () => {
      await js(`(() => { const m = document.querySelector('.cal-day[data-date="${fx.today}"] .cal-more') || document.querySelector('.cal-day.today .cal-more'); m?.click(); return Boolean(m); })()`);
      return until(() => js(`document.querySelectorAll('.cal-peek .cal-chip').length >= 8`), 'the day\'s list', 4000).then(() => true, () => false);
    };
    const PEEK = `(() => {
      const p = document.querySelector('.cal-peek');
      if (!p) return null;
      const r = p.getBoundingClientRect();
      const list = p.querySelector('.cal-peek-list');
      return {
        head: p.querySelector('.cal-peek-head')?.textContent.trim() || '',
        titles: [...p.querySelectorAll('.cal-chip .s')].map((n) => n.textContent),
        inside: r.left >= 0 && r.top >= 0 && r.right <= innerWidth && r.bottom <= innerHeight,
        scrolls: list ? list.scrollHeight > list.clientHeight + 1 : false,
        cut: [...p.querySelectorAll('.cal-chip')].filter((n) => { const b = n.getBoundingClientRect(); return b.right > r.right + 0.5 || b.left < r.left - 0.5; }).length,
      };
    })()`;
    const opened = await openPeek();
    const peek = (await js(PEEK)) || {};
    const all = fx.byDay.get(fx.today);
    check('calendar-month: "+N more" opens a list of all the busy day\'s events over the month, named by its day, inside the window',
      opened && all.every((s) => (peek.titles || []).includes(s)) && peek.inside && !peek.cut && /\d/.test(peek.head || ''), JSON.stringify(peek));
    await snap(win, 'calendar-month-more-light.png');
    await js(`document.documentElement.dataset.theme = 'dark', 1`);
    await wait(200);
    await snap(win, 'calendar-month-more-dark.png');
    await js(`document.documentElement.dataset.theme = 'light', 1`);
    await js(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })), 1`);
    const escaped = await until(() => js(`!document.querySelector('.cal-peek')`), 'the list to close', 3000).then(() => true, () => false);
    await openPeek();
    await js(`document.querySelector('.cal-month-head')?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })), 1`);
    const away = await until(() => js(`!document.querySelector('.cal-peek')`), 'the list to close on a click elsewhere', 3000).then(() => true, () => false);
    check('calendar-month: the day\'s list closes with Escape and with a click elsewhere', escaped && away, `Escape ${escaped}, click ${away}`);

    // An event in the list opens it; the day's name opens the day.
    await openPeek();
    await js(`[...document.querySelectorAll('.cal-peek .cal-chip')].find((n) => n.querySelector('.s')?.textContent === 'Train home')?.click(), 1`);
    const detail = await until(() => js(`Boolean(document.querySelector('.rw-dialog .cal-detail')) && !document.querySelector('.cal-peek')`), 'the event', 4000).then(() => true, () => false);
    const title = await js(`document.querySelector('.rw-dialog .rw-dialog-head')?.textContent.trim() || ''`);
    await js(`[...document.querySelectorAll('.rw-dialog .rw-btn')].find((b) => b.textContent.trim() === 'Close')?.click(), 1`);
    await until(() => js(`!document.querySelector('.rw-dialog')`), 'the event to close', 3000).catch(() => {});
    await openPeek();
    await js(`document.querySelector('.cal-peek-head')?.click(), 1`);
    const dayView = await until(() => js(`document.querySelectorAll('.cal-hours .cal-col').length === 1 && !document.querySelector('.cal-peek')`), 'the day view', 4000).then(() => true, () => false);
    const dayBlocks = await js(`[...document.querySelectorAll('.cal-block .s')].map((n) => n.textContent)`);
    check('calendar-month: an event in the day\'s list opens it, and the day\'s name opens that day',
      detail && /Train home/.test(title) && dayView && dayBlocks.includes('Train home'), `event ${detail} "${title}", day view ${dayView} ${JSON.stringify(dayBlocks)}`);
    await js(`[...document.querySelectorAll('.cal-view')].find((b) => b.textContent === 'Month')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.cal-month-grid'))`), 'the month again', 4000).catch(() => {});
  } catch (err) {
    check('calendar-month: the checks ran', false, err.stack || err.message);
  }
}
