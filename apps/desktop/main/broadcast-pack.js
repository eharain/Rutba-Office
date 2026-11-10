// Live Broadcast's pack: what the core calls to register it (docs/APP-PACKS.md 5.1.2).
// Going live: the destinations, their keys in the keystore, the publishers (live.js).
// The studio's picture and recording the screen are the core's (the capture namespace), shared with the Recorder.

import { sendEvent } from '@rutba/office-shell/electron/main';
import { createLiveService } from './live.js';

export function register({ stores }) {
  return { live: createLiveService({ stores, send: sendEvent }) };
}
