// Live Broadcast: the studio window, going live to the destinations set,
// with several cameras' shots switched by hand, by who is speaking, or on a
// timer (owner, 2026-11-03, D-2026-11-03-01 and -02). The studio is the
// Recorder's (../studio/); going live is this app's own.

import React from 'react';
import { liveBitrate } from '@rutba/stream/endpoints';
import { StudioWindow } from '../studio/window.js';
import { LiveEncoder, AUDIO_KBPS } from './broadcast/live.js';
import { LiveDialog, LIVE_CSS } from './broadcast/live-dialog.js';
import { LiveBar, LIVE_BAR_CSS } from '../studio/live-bar.js';

const LIVE = { LiveEncoder, AUDIO_KBPS, LiveDialog, LIVE_CSS: LIVE_CSS + LIVE_BAR_CSS, LiveBar, liveBitrate };

export default function Broadcast(props) {
  return <StudioWindow {...props} kind="broadcast" live={LIVE} />;
}
