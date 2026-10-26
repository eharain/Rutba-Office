// Recorder's pack: what the core calls to register it (docs/APP-PACKS.md 5.1.2).
// Go Live: the destinations, their keys in the keystore, the publishers (live.js).
// Recording the screen itself is the core's (the capture namespace), shared with Video and Presentations.

import { createLiveService } from './live.js';

export function register({ stores }) {
  return { live: createLiveService({ stores }) };
}
