// Calendar: the faults the how-to scripts found (2026-10-24, A6.7). Saving
// an event with people invited offers the message with the invitation, and
// the message shows the file's size; editing one occurrence of a repeating
// event edits that occurrence, its own day shown, until Every occurrence is
// chosen; a repeating event shows its Reminder, one not in the list among
// them; an event's time zone can be chosen, and its times there are shown
// beside this computer's; a second zone's hours run beside the day and the
// week; and the small month's arrows move the main view. Run alone with
// RUTBA_VERIFY_ONLY=calendar-fixes.

import { BrowserWindow } from 'electron';
import { localToUtc } from '@rutba/calendar';

const pad = (n) => String(n).padStart(2, '0');
const isoDate = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;

export async function verifyCalendarFixes({ open, check, until, wait, errorsIn }) {
  const local = Intl.DateTimeFormat().resolvedOptions().timeZone;
  const away = local === 'America/New_York' ? 'Asia/Tokyo' : 'America/New_York';
  const second = local === 'Asia/Tokyo' ? 'Europe/London' : 'Asia/Tokyo';
  let mail = null;
  try {
    const win = await open('calendar');
    const js = (code) => win.webContents.executeJavaScript(code);
    win.setContentSize(1456, 910);
    const click = (sel, text) => js(`(() => { const b = [...document.querySelectorAll(${JSON.stringify(sel)})].find((x) => x.textContent.trim() === ${JSON.stringify(text)}); b?.click(); return Boolean(b); })()`);
    const pick = (sel, value) => js(`(() => { const s = ${sel.startsWith('!') ? sel.slice(1) : `document.querySelector(${JSON.stringify(sel)})`}; if (!s) return null; Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(s, ${JSON.stringify(value)}); s.dispatchEvent(new Event('change', { bubbles: true })); return s.value; })()`);
    const type = (sel, value) => js(`(() => { const i = document.querySelector(${JSON.stringify(sel)}); if (!i) return false; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(i, ${JSON.stringify(value)}); i.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
    const REPEAT = `![...document.querySelectorAll('.cal-form select')].find((s) => [...s.options].some((o) => o.value === 'weekly'))`;

    // A weekly event from today at 10:00 here, with a reminder ten minutes before (not one of the editor's own).
    const today = new Date();
    const nextWeek = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 7);
    await js(`(async () => {
      const api = window.rutbaOffice.calendar;
      const cals = await api.calendars({});
      const d = new Date(); d.setHours(10, 0, 0, 0);
      const p = (n) => String(n).padStart(2, '0');
      const at = (ms) => { const x = new Date(ms); return { date: x.getFullYear() + '-' + p(x.getMonth() + 1) + '-' + p(x.getDate()), time: p(x.getHours()) + ':' + p(x.getMinutes()) + ':00', tzid: ${JSON.stringify(local)}, utc: false, allDay: false, at: ms }; };
      await api.save({ calendarId: cals[0].id, event: { summary: 'Weekly sync', start: at(d.getTime()), end: at(d.getTime() + 3600000),
        rrule: { freq: 'WEEKLY', interval: 1, count: null, until: null, byDay: [{ n: 0, day: d.getDay() }], byMonthDay: null, byMonth: null, bySetPos: null, wkst: 'MO' },
        alarms: [{ action: 'DISPLAY', before: 10, at: null, related: 'START', description: null }] } });
      return 1;
    })()`);

    // One occurrence opened from the list of the next thirty days: the editor edits that one.
    await click('.cal-view', 'Agenda');
    await until(() => js(`document.querySelectorAll('.cal-agenda-item').length >= 2`), 'the weekly event in the agenda', 6000).catch(() => {});
    await js(`[...document.querySelectorAll('.cal-agenda-day')].find((d) => d.querySelector('.cal-agenda-date .dd')?.textContent === ${JSON.stringify(String(nextWeek.getDate()))} && /Weekly sync/.test(d.textContent))?.querySelector('.cal-agenda-item')?.click(), 'opened'`);
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-dialog .rw-btn')].find((b) => b.textContent.trim() === 'Edit this one'))`), 'the occurrence', 4000).catch(() => {});
    await click('.rw-dialog .rw-btn', 'Edit this one');
    await until(() => js(`Boolean(document.querySelector('.cal-scope'))`), 'the editor', 4000).catch(() => {});
    const one = await js(`({ scope: document.querySelector('.cal-scope')?.value, date: document.querySelector('.cal-form input[type=date]')?.value, reminder: document.querySelector('.cal-reminder')?.value, reminderText: document.querySelector('.cal-reminder')?.selectedOptions[0]?.textContent, repeat: Boolean(${REPEAT.slice(1)}) })`);
    check('calendar: editing one occurrence of a repeating event edits that occurrence, its own day shown',
      one.scope === 'this' && one.date === isoDate(nextWeek) && !one.repeat, JSON.stringify(one));
    await pick('.cal-scope', 'all');
    await wait(150);
    const all = await js(`({ date: document.querySelector('.cal-form input[type=date]')?.value, repeat: ${REPEAT.slice(1)}?.value, until: [...document.querySelectorAll('.cal-form .rw-field label')].some((l) => l.textContent === 'Until'), reminder: document.querySelector('.cal-reminder')?.value })`);
    check('calendar: choosing every occurrence shows the series\' own first day, and a repeating event keeps its Reminder (one not in the list among them)',
      all.date === isoDate(today) && all.repeat === 'weekly' && all.until && all.reminder === '10' && one.reminder === '10' && /10 minutes before/.test(one.reminderText || ''), JSON.stringify({ all, one }));
    await click('.rw-dialog .rw-btn', 'Cancel');
    await until(() => js(`!document.querySelector('.cal-form')`), 'the editor to close', 3000).catch(() => {});

    // A new event in another zone, with someone invited.
    await click('.rw-ribbon .rw-btn', 'New event');
    await until(() => js(`Boolean(document.querySelector('.cal-zone'))`), 'the editor with its time zone', 4000).catch(() => {});
    const zoneShown = await js(`document.querySelector('.cal-zone')?.value`);
    await pick(REPEAT, 'weekly');
    await wait(100);
    const reminderWhenRepeating = await js(`Boolean(document.querySelector('.cal-reminder'))`);
    await pick(REPEAT, 'none');
    check('calendar: a new event offers its time zone, this computer\'s first, and its Reminder stays while it repeats', zoneShown === local && reminderWhenRepeating, JSON.stringify({ zoneShown, local, reminderWhenRepeating }));
    await type('.cal-form .cal-title', 'Zone call');
    await pick('.cal-zone', away);
    await type('.cal-form input[type=time]', '09:00');
    await type('.cal-invitees', 'kim@example.com');
    await wait(150);
    const date = await js(`document.querySelector('.cal-form input[type=date]')?.value`);
    const hint = await js(`[...document.querySelectorAll('.cal-form .rw-hint')].map((h) => h.textContent).find((h) => /on this computer/.test(h)) || ''`);
    const before = new Set(BrowserWindow.getAllWindows().map((w) => w.id));
    await click('.rw-dialog .rw-btn', 'Save');
    await until(() => js(`Boolean(document.querySelector('.cal-offer'))`), 'the offer to send the invitation', 4000).catch(() => {});
    const offer = await js(`document.querySelector('.cal-offer')?.textContent || ''`);
    const [y, m, d] = date.split('-').map(Number);
    const expected = localToUtc({ y, m, d, H: 9, M: 0 }, away);
    const stored = await js(`(async () => { const all = await window.rutbaOffice.calendar.events({ from: ${expected - 86400000}, to: ${expected + 86400000} }); const s = all.find((e) => e.summary === 'Zone call'); return s ? { start: s.start, ev: await window.rutbaOffice.calendar.get({ id: s.id }) } : null; })()`);
    check('calendar: an event\'s time zone can be chosen: its times kept on that zone\'s clock, and the editor says the time here',
      stored?.start === expected && stored?.ev?.start?.tzid === away && hint.length > 0, JSON.stringify({ expected, start: stored?.start, tzid: stored?.ev?.start?.tzid, hint }));
    check('calendar: saving an event with people invited offers the message with the invitation', /Send invitation/.test(offer) && /1 person invited/.test(offer), offer);

    // Sent: a message in Mail with the invitation, its size shown.
    await click('.cal-offer .rw-btn', 'Send invitation');
    await until(() => BrowserWindow.getAllWindows().some((w) => !before.has(w.id) && /app=mail|mail/.test(w.webContents.getURL())), 'the message to open', 8000).catch(() => {});
    mail = BrowserWindow.getAllWindows().find((w) => !before.has(w.id)) || null;
    let sent = { url: mail?.webContents.getURL() || '' };
    if (mail) {
      await until(() => mail.webContents.executeJavaScript(`Boolean(document.querySelector('.ml-attachment'))`), 'the attachment in the message', 10000).catch(() => {});
      sent = { ...sent, ...(await mail.webContents.executeJavaScript(`({ name: document.querySelector('.ml-attachment .name')?.textContent, size: document.querySelector('.ml-attachment .size')?.textContent })`)) };
    }
    check('calendar: Send invitation opens a message with the invitation attached, its size shown, not 0 B', /\.ics$/.test(sent.name || '') && sent.size && !/^0\s*B$/.test(sent.size.trim()), JSON.stringify(sent));
    if (mail && !mail.isDestroyed()) mail.destroy();
    mail = null;
    win.focus();

    // Its times on its own zone's clock, beside this computer's.
    await until(() => js(`[...document.querySelectorAll('.cal-agenda-item')].some((b) => /Zone call/.test(b.textContent))`), 'the event in the agenda', 4000).catch(() => {});
    await js(`[...document.querySelectorAll('.cal-agenda-item')].find((b) => /Zone call/.test(b.textContent))?.click(), 'opened'`);
    await until(() => js(`Boolean(document.querySelector('.cal-detail'))`), 'the event', 4000).catch(() => {});
    const zoneLine = await js(`document.querySelector('.cal-detail-zone')?.textContent || ''`);
    check('calendar: an event in another zone shows its times on that zone\'s clock too', /^9:00 to 10:00 in /.test(zoneLine), zoneLine);
    await click('.rw-dialog .rw-btn', 'Delete');
    await wait(200);

    // The small month's arrows move the main view.
    await click('.cal-view', 'Month');
    await wait(200);
    const titleBefore = await js(`document.querySelector('.cal-toolbar .cal-title')?.textContent || ''`);
    await js(`document.querySelector('.cal-mini-next')?.click(), 'next'`);
    await wait(250);
    const titleAfter = await js(`document.querySelector('.cal-toolbar .cal-title')?.textContent || ''`);
    const nextMonth = new Date(today.getFullYear(), today.getMonth() + 1, 1).toLocaleDateString('en', { month: 'long' });
    await js(`document.querySelector('.cal-mini-prev')?.click(), 'back'`);
    await wait(250);
    const titleBack = await js(`document.querySelector('.cal-toolbar .cal-title')?.textContent || ''`);
    check('calendar: the small month\'s arrows move the main view a month, and back', titleAfter.startsWith(nextMonth) && titleBack === titleBefore, JSON.stringify({ titleBefore, titleAfter, titleBack }));

    // A second zone's hours beside the week's.
    await click('.rw-tab', 'View');
    await until(() => js(`Boolean(document.querySelector('.cal-second-zone'))`), 'View → Second time zone', 3000).catch(() => {});
    await pick('.cal-second-zone', second);
    await click('.cal-view', 'Week');
    await until(() => js(`Boolean(document.querySelector('.cal-time.two .cal-zone-hours'))`), 'the second zone\'s hours', 3000).catch(() => {});
    const hours = await js(`(() => {
      const shown = [...document.querySelectorAll('.cal-zone-hours .cal-hour')].map((h) => h.textContent);
      const head = document.querySelector('.cal-time-head .cal-zone-name')?.textContent;
      return { shown: shown.slice(0, 4), head };
    })()`);
    // What the second zone's clock says at 1:00 here on the week's first day.
    const weekFirst = await js(`(() => { const days = [...document.querySelectorAll('.cal-col-head .dd')].map((n) => Number(n.textContent)); return days[0]; })()`);
    const base = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    for (let i = 0; i < 7 && base.getDate() !== weekFirst; i++) base.setDate(base.getDate() - 1);
    const oneHere = new Date(base.getFullYear(), base.getMonth(), base.getDate(), 1).getTime();
    const there = new Intl.DateTimeFormat('en-GB', { timeZone: second, hour: 'numeric', minute: '2-digit', hourCycle: 'h23' }).formatToParts(new Date(oneHere));
    const want = `${Number(there.find((p) => p.type === 'hour').value)}:${there.find((p) => p.type === 'minute').value}`;
    check('calendar: a second time zone\'s hours run beside this computer\'s in the week, headed by its place', hours.shown[1] === want && hours.head === second.split('/').pop().replace(/_/g, ' '), JSON.stringify({ hours, want }));
    await pick('.cal-second-zone', '');

    const complaints = await errorsIn(win);
    check('calendar: the fixes report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('calendar: the fixes checks ran', false, err.message);
  } finally {
    if (mail && !mail.isDestroyed()) mail.destroy();
  }
}
