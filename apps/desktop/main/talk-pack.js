// Talk's pack: what the core calls to register it (docs/APP-PACKS.md 5.1.2).
//
// Messages and calls with contacts, through the introducer. An invitation
// emailed from the person's own address goes through Mail, a copy in Sent;
// with Mail absent there are no accounts to send from, and a send asked for
// all the same is refused in plain words.

import path from 'node:path';
import { createTalkService } from './talk.js';

/** Mail as Talk reaches it: through `packs`, each time, so Mail may be absent. */
export function mailFor(packs) {
  return {
    accounts: () => packs.service('mail')?.accounts?.() || [],
    send: async (args) => {
      const mail = packs.service('mail');
      if (!mail) throw new Error('Mail is not installed on this computer, so the invitation cannot be sent from your own address. Add Mail with Add or remove apps, or let Rutba send it.');
      return mail.send(args);
    },
  };
}

export function register({ stores, broadcast, packs }) {
  return {
    talk: createTalkService({
      stores,
      broadcast,
      dataDir: path.join(stores.dir, 'talk'),
      // The sign-in opens the system browser; under the window checks the address is noted instead, as a link's is.
      openBrowser: (url) => {
        if (process.env.RUTBA_OFFICE_VERIFY_APPS) { (globalThis.__rutbaCheckOpened ||= []).push(url); return undefined; }
        return import('electron').then(({ shell }) => shell.openExternal(url));
      },
      // A test server and token: the window checks, and the screens build's Talk picture (a stand-in introducer).
      testServers: Boolean(process.env.RUTBA_OFFICE_VERIFY_APPS || process.env.RUTBA_SCREENS_FEATURES),
      mail: mailFor(packs),
    }),
  };
}
