// Calendar: availability on request (2.25). In the event editor, Check
// availability asks for each address on the Invite line: one the CalDAV
// account's server knows is answered from its scheduling outbox (the small
// server in tests/fixtures, on this computer); one it does not is asked by
// email, a VFREEBUSY REQUEST from the account Mail sends from, to the test
// SMTP server (check-smtp.js), and the row says so; when that person's
// calendar answers, the row shows it without being asked again. Run alone
// with RUTBA_VERIFY_ONLY=calendar-availability.

import { parseMessage } from '@rutba/mailbox/mime';
import { readScheduling, writeFreeBusyReply } from '@rutba/calendar';
import { startTestSmtp } from './check-smtp.js';
import { startDavServer } from '../../../tests/fixtures/dav-server.js';

const H = 3600000;
const pad = (n) => String(n).padStart(2, '0');

export async function verifyCalendarAvailability({ open, check, until, wait, errorsIn }) {
  let smtp = null;
  let dav = null;
  let accountId = null;
  let davId = null;
  let mail = null;
  let win = null;
  try {
    // Three days ahead, ten to eleven on this computer's clock.
    const day = new Date();
    day.setHours(0, 0, 0, 0);
    day.setDate(day.getDate() + 3);
    const at = (h) => day.getTime() + h * H;
    smtp = await startTestSmtp();
    globalThis.__rutbaCheckSmtp = { port: smtp.port };
    dav = await startDavServer({ books: [], busy: { 'kim@dav.example': [{ start: at(9.5), end: at(10.5) }] } });

    mail = await open('mail');
    const mjs = (code) => mail.webContents.executeJavaScript(code);
    accountId = (await mjs(`window.rutbaOffice.mail.addAccount({ account: { email: 'ann@dav.example', name: 'Ann Avail', imap: { host: 'mail.dav.example', port: 993, secure: true }, smtp: { host: 'mail.dav.example', port: 587 } }, password: 'check' })`)).id;

    win = await open('calendar');
    const js = (code) => win.webContents.executeJavaScript(code);
    win.setContentSize(1456, 910);
    davId = (await js(`window.rutbaOffice.dav.add({ url: ${JSON.stringify(dav.url)}, user: 'ann', password: 'secret', name: 'Availability check' })`))?.id || null;
    const click = (sel, text) => js(`(() => { const b = [...document.querySelectorAll(${JSON.stringify(sel)})].find((x) => x.textContent.trim() === ${JSON.stringify(text)}); b?.click(); return Boolean(b); })()`);
    const type = (sel, value, nth = 0) => js(`(() => { const i = document.querySelectorAll(${JSON.stringify(sel)})[${nth}]; if (!i) return false; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(i, ${JSON.stringify(value)}); i.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
    const rows = () => js(`[...document.querySelectorAll('.cal-avail .cal-avail-row')].map((r) => ({ email: r.dataset.email, state: r.dataset.state, text: r.textContent }))`);

    await click('.rw-ribbon .rw-btn', 'New event');
    await until(() => js(`Boolean(document.querySelector('.cal-invitees'))`), 'the editor', 5000).catch(() => {});
    await type('.cal-form .cal-title', 'Availability check');
    const date = `${day.getFullYear()}-${pad(day.getMonth() + 1)}-${pad(day.getDate())}`;
    await type('.cal-form input[type=date]', date, 0);
    await type('.cal-form input[type=time]', '10:00', 0);
    await wait(100);
    await type('.cal-form input[type=date]', date, 1);
    await type('.cal-form input[type=time]', '11:00', 1);
    await type('.cal-invitees', 'kim@dav.example, sam@far.example');
    await wait(150);
    await js(`document.querySelector('.rw-dialog .cal-avail-check')?.click(), 1`);
    await until(async () => (await rows()).length === 2 || null, 'the availability', 15000).catch(() => {});
    const first = await rows();
    await until(() => smtp.received.length >= 1 || null, 'the enquiry by mail', 15000).catch(() => {});
    const sent = smtp.received[0] ? parseMessage(smtp.received[0].data) : null;
    const part = (sent?.attachments || []).find((a) => /calendar/i.test(a.type || '') || /\.ics$/i.test(a.filename || ''));
    const ask = part?.bytes ? readScheduling(Buffer.from(part.bytes).toString('utf8')).freebusy[0] : null;
    const kim = first.find((r) => r.email === 'kim@dav.example');
    const sam = first.find((r) => r.email === 'sam@far.example');
    check('calendar availability: Check availability answers from the CalDAV server for the address it knows, and asks the other by email, saying so',
      kim?.state === 'busy' && sam?.state === 'asked' && /Asked by email/.test(sam.text)
        && smtp.received.length === 1 && smtp.received[0].to.includes('sam@far.example') && /^When are you free\?/.test(sent?.subject || '') && ask?.attendees?.[0]?.email === 'sam@far.example',
      JSON.stringify({ first, to: smtp.received[0]?.to, subject: sent?.subject, ask }));

    // Sam's calendar answers by mail: busy half past ten to half past eleven.
    const reply = writeFreeBusyReply({ uid: ask?.uid || 'fb', organizer: ask?.organizer || { email: 'ann@dav.example' }, attendee: { email: 'sam@far.example' }, start: ask?.start ?? at(0), end: ask?.end ?? at(24), busy: [{ start: at(10.5), end: at(11.5) }] });
    const raw = ['From: Sam <sam@far.example>', 'To: Ann Avail <ann@dav.example>', 'Subject: Free/busy', `Date: ${new Date().toUTCString()}`, `Message-ID: <fb-${Date.now()}@far.example>`, 'MIME-Version: 1.0',
      'Content-Type: multipart/mixed; boundary="f1"', '', '--f1', 'Content-Type: text/plain; charset=utf-8', '', 'Busy times.', '--f1',
      'Content-Type: text/calendar; method=REPLY; charset=utf-8', 'Content-Disposition: attachment; filename="freebusy.ics"', '', reply, '--f1--', ''].join('\r\n');
    await mjs(`window.rutbaOffice.mail.deliverTest({ accountId: ${JSON.stringify(accountId)}, folder: 'Inbox', raw: ${JSON.stringify(raw)} })`);
    const taken = await js(`window.rutbaOffice.calendar.checkMail()`);
    await until(async () => (await rows()).find((r) => r.email === 'sam@far.example')?.state === 'busy' || null, "Sam's answer in the editor", 8000).catch(() => {});
    const after = await rows();
    check("calendar availability: when the answer comes by mail, the editor shows it, nobody asked twice",
      taken.some((r) => r.kind === 'freebusy') && after.find((r) => r.email === 'sam@far.example')?.state === 'busy' && smtp.received.length === 1,
      JSON.stringify({ taken, after, sent: smtp.received.length }));

    await click('.rw-dialog .rw-btn', 'Cancel');
    const complaints = await errorsIn(win);
    check('calendar availability: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('calendar availability: the checks ran', false, err.stack || err.message);
  } finally {
    delete globalThis.__rutbaCheckSmtp;
    await smtp?.close().catch(() => {});
    if (win && !win.isDestroyed() && davId) await win.webContents.executeJavaScript(`window.rutbaOffice.dav.remove({ id: ${JSON.stringify(davId)} }).then(() => 1, () => 1)`).catch(() => {});
    await dav?.close().catch(() => {});
    if (mail && !mail.isDestroyed()) {
      if (accountId) await mail.webContents.executeJavaScript(`window.rutbaOffice.mail.removeAccount({ id: ${JSON.stringify(accountId)} }).then(() => 1, () => 1)`).catch(() => {});
      mail.destroy();
    }
    if (win && !win.isDestroyed()) win.destroy();
  }
}
