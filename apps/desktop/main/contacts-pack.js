// Contacts' pack: what the core calls to register it (docs/APP-PACKS.md 5.1.2).
//
// The address book, with the people Mail has seen behind it for Compose to
// complete from. Mail is reached through `packs` when it is asked, so it may
// register before or after this, or be absent: then there are only the cards.

import { createContactsService } from './contacts.js';

/** The people Mail has seen, for the address book's suggestions: none while Mail is absent. */
export function peopleFrom(packs) {
  return () => {
    try {
      const mail = packs.service('mail');
      const out = [];
      for (const account of mail?.accounts?.() || []) {
        const { rows = [] } = mail.people({ accountId: account.id, limit: 300 }) || {};
        for (const p of rows) out.push({ name: p.name || '', email: p.address });
      }
      return out;
    } catch {
      return [];
    }
  };
}

export function register({ stores, broadcast, packs }) {
  return {
    contacts: createContactsService({ stores, broadcast, people: peopleFrom(packs) }),
  };
}
