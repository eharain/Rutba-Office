// Presentations' pack: what the core calls to register it (docs/APP-PACKS.md 5.1.2).
// Where a slide show is up to, shared by the audience's window and the presenter's (present.js).
// The decks themselves are the core's document service, which the three document apps share.

import { createPresentService } from './present.js';

export function register({ broadcast }) {
  return { present: createPresentService({ broadcast }) };
}
