// Calendar and Contacts: an account on a CalDAV and CardDAV server.
//
// A small server of the suite's own (tests/fixtures/dav-server.js) runs on
// this computer for the length of the check. Through the Accounts dialog the
// account is added with its address, user name and password; the server's
// calendar appears in the sidebar with its event on the month; an event
// changed on the server comes back with Sync; an event made here goes to the
// server; the server's card is in Contacts; and removing the account takes
// its calendar away again. Run alone with RUTBA_VERIFY_ONLY=dav.
import { startDavServer } from '../../../tests/fixtures/dav-server.js';

const ics = (uid, summary, start) => [
  'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Other device//EN', 'BEGIN:VEVENT', `UID:${uid}`, 'DTSTAMP:20261001T000000Z',
  `DTSTART:${new Date(start).toISOString().replace(/[-:]/g, '').slice(0, 15)}Z`,
  `DTEND:${new Date(start + 3600000).toISOString().replace(/[-:]/g, '').slice(0, 15)}Z`, `SUMMARY:${summary}`, 'END:VEVENT', 'END:VCALENDAR', '',
].join('\r\n');

export async function verifyDav({ open, check, until, wait, errorsIn }) {
  const server = await startDavServer();
  try {
    const today = new Date();
    const at = Date.UTC(today.getUTCFullYear(), today.getUTCMonth(), today.getUTCDate(), 12);
    const work = server.calendarPath('work');
    server.put(`${work}standup.ics`, ics('standup@other', 'Server stand-up', at));
    server.put(`${server.bookPath('friends')}bo.vcf`, ['BEGIN:VCARD', 'VERSION:3.0', 'UID:bo@other', 'FN:Bo Lund', 'N:Lund;Bo;;;', 'EMAIL:bo@example.com', 'END:VCARD', ''].join('\r\n'));

    const win = await open('calendar');
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-ribbon .rw-btn')].find((b) => b.textContent.trim() === 'Accounts'))`), 'Home → Accounts', 8000);
    await js(`[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((b) => b.textContent.trim() === 'Accounts').click(), 'open'`);
    await until(() => js(`Boolean(document.querySelector('input.dav-url'))`), 'the Accounts dialog', 4000);
    await js(`(() => {
      const set = (sel, v) => { const i = document.querySelector(sel); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(i, v); i.dispatchEvent(new Event('input', { bubbles: true })); };
      set('input.dav-url', ${JSON.stringify(server.url)});
      set('input.dav-user', 'ann');
      set('input.dav-password', 'secret');
      return 1;
    })()`);
    await wait(150);
    await js(`document.querySelector('.dav-add').click(), 'add'`);
    await until(() => js(`Boolean(document.querySelector('.dav-account'))`), 'the account listed', 10000).catch(() => {});
    const listed = await js(`document.querySelector('.dav-account .dav-status')?.textContent || document.querySelector('.dav-error')?.textContent || ''`);
    check('calendar: Accounts adds a CalDAV account with its address, user name and password, and says what it found', /1 calendar, 1 address book · synced/.test(listed), listed);
    await js(`[...document.querySelectorAll('.rw-dialog .rw-btn')].find((b) => b.textContent.trim() === 'Close')?.click(), 'closed'`);
    await until(() => js(`[...document.querySelectorAll('.cal-chip .s')].some((n) => n.textContent === 'Server stand-up')`), 'the server\'s event on the month', 8000).catch(() => {});
    const shown = await js(`({ cals: [...document.querySelectorAll('.cal-cal .n')].map((n) => n.textContent), chips: [...document.querySelectorAll('.cal-chip .s')].map((n) => n.textContent) })`);
    check('calendar: the account\'s calendar is in the sidebar and its event on the month', shown.cals.includes('Work') && shown.chips.includes('Server stand-up'), JSON.stringify(shown));

    // Changed on the server, brought back by Sync.
    server.put(`${work}standup.ics`, ics('standup@other', 'Server stand-up, moved', at));
    await js(`[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((b) => b.textContent.trim() === 'Sync').click(), 'sync'`);
    await until(() => js(`[...document.querySelectorAll('.cal-chip .s')].some((n) => n.textContent === 'Server stand-up, moved')`), 'the change', 8000).catch(() => {});
    check('calendar: Sync brings back an event changed on the server', await js(`[...document.querySelectorAll('.cal-chip .s')].some((n) => n.textContent === 'Server stand-up, moved')`), JSON.stringify(await js(`[...document.querySelectorAll('.cal-chip .s')].map((n) => n.textContent)`)));

    // Made here, sent to the server.
    await js(`(async () => {
      const cals = await window.rutbaOffice.calendar.calendars({});
      const w = cals.find((c) => c.name === 'Work');
      const start = ${at + 3 * 3600000};
      const t = (ms) => ({ date: new Date(ms).toISOString().slice(0, 10), time: new Date(ms).toISOString().slice(11, 19), tzid: null, utc: true, allDay: false, at: ms });
      await window.rutbaOffice.calendar.save({ calendarId: w.id, event: { summary: 'Made here', start: t(start), end: t(start + 3600000) } });
      await window.rutbaOffice.dav.sync({});
      return 1;
    })()`);
    await until(() => [...server.collections.get(work).items.values()].some((i) => /SUMMARY:Made here/.test(i.data)), 'the event on the server', 8000).catch(() => {});
    check('calendar: an event made in an account\'s calendar goes to the server', [...server.collections.get(work).items.values()].some((i) => /SUMMARY:Made here/.test(i.data)), `${server.collections.get(work).items.size} items on the server`);

    // The server's card is in Contacts.
    const people = await open('contacts');
    const pjs = (code) => people.webContents.executeJavaScript(code);
    await until(() => pjs(`document.body.innerText.includes('Bo Lund')`), 'the server\'s card', 8000).catch(() => {});
    check('contacts: the account\'s card is among the contacts', await pjs(`document.body.innerText.includes('Bo Lund')`), 'Bo Lund');

    // Removed: its calendar goes.
    await js(`[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((b) => b.textContent.trim() === 'Accounts').click(), 'open'`);
    await until(() => js(`Boolean(document.querySelector('.dav-account'))`), 'the account', 4000).catch(() => {});
    await js(`[...document.querySelectorAll('.dav-account .rw-btn')].find((b) => b.textContent.trim() === 'Remove')?.click(), 'removed'`);
    await until(() => js(`!document.querySelector('.dav-account')`), 'the account gone', 4000).catch(() => {});
    await js(`[...document.querySelectorAll('.rw-dialog .rw-btn')].find((b) => b.textContent.trim() === 'Close')?.click(), 'closed'`);
    await until(() => js(`![...document.querySelectorAll('.cal-cal .n')].some((n) => n.textContent === 'Work')`), 'the calendar gone', 6000).catch(() => {});
    check('calendar: removing the account takes its calendar away', await js(`![...document.querySelectorAll('.cal-cal .n')].some((n) => n.textContent === 'Work')`), JSON.stringify(await js(`[...document.querySelectorAll('.cal-cal .n')].map((n) => n.textContent)`)));
    const complaints = [...(await errorsIn(win)), ...(await errorsIn(people))];
    check('calendar and contacts: accounts report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('calendar: the accounts checks ran', false, err.message);
  } finally {
    await server.close();
  }
}
