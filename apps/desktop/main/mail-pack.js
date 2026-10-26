// Mail's pack: what the core calls to register it (docs/APP-PACKS.md 5.1.2).
//
// The core hands every pack the same context (app-packs.js): the stores, the
// broadcast, the services of the core it may use (oauth, the blob store), and
// `packs`, through which it reaches another pack's service without importing
// it. An address book that is not installed answers no cards: "reply to
// contacts only" and the junk filter's trust in contacts then match no one.

import { createMailService } from './mail.js';
import { createDiscoveryService } from './discover.js';

/** The address book as Mail reaches it: no cards while Contacts is absent. */
export function contactsFor(packs) {
  return { list: (args) => packs.service('contacts')?.list?.(args) || [] };
}

export function register(core) {
  const { stores, holdBlob, broadcast, oauth, packs } = core;
  return {
    mail: createMailService({
      stores,
      holdBlob,
      broadcast,
      userData: stores.dir,
      oauth,
      contacts: contactsFor(packs),
    }),
    // What other mail clients have left on this computer.
    discover: createDiscoveryService(),
  };
}
