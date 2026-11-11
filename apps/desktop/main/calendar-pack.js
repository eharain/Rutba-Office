// Calendar's pack: what the core calls to register it (docs/APP-PACKS.md 5.1.2).
//
// The calendar sends invitations and answers through Mail, and reads the
// answers, changes and cancellations that come back out of its Inboxes; Mail
// is reached through `packs` when it is asked (calendar-mail.js), and with
// Mail absent there are no accounts to send from or to read.

import { createCalendarService } from './calendar.js';
import { mailFor } from './calendar-mail.js';

export { mailFor };

/** How often the Inboxes are read for answers, after the first look a little after start. */
const EVERY = 90 * 1000;
const FIRST = 20 * 1000;

export function register({ stores, broadcast, packs }) {
  const calendar = createCalendarService({ stores, broadcast, mail: mailFor(packs), contacts: () => packs.service('contacts') });
  // A window check reads them itself (calendar.checkMail), when it means to.
  if (!process.env.RUTBA_OFFICE_VERIFY_APPS && !process.env.RUTBA_OFFICE_SMOKE) {
    let reading = false;
    const look = async () => {
      if (reading) return;
      reading = true;
      try {
        await calendar.checkMail();
      } catch (err) {
        console.warn('calendar: reading the Inboxes for answers failed:', err?.message || err);
      } finally {
        reading = false;
      }
    };
    setTimeout(look, FIRST).unref?.();
    setInterval(look, EVERY).unref?.();
  }
  return { calendar };
}
