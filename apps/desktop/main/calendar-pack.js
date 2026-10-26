// Calendar's pack: what the core calls to register it (docs/APP-PACKS.md 5.1.2).
//
// The calendar answers an invitation through a message Mail sends; Mail is
// reached through `packs` when it is asked, and with Mail absent there are no
// accounts to answer from.

import { createCalendarService } from './calendar.js';

/** Mail as the calendar reaches it: its accounts, none while Mail is absent. */
export function mailFor(packs) {
  return { accounts: () => packs.service('mail')?.accounts?.() || [] };
}

export function register({ stores, broadcast, packs }) {
  return {
    calendar: createCalendarService({ stores, broadcast, mail: mailFor(packs) }),
  };
}
