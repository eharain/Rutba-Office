// The Recorder: the studio window, recording to a file. Going live is Live
// Broadcast's (D-2026-11-03-01); the studio itself is in ../studio/.

import React from 'react';
import { StudioWindow } from '../studio/window.js';

export default function Recorder(props) {
  return <StudioWindow {...props} kind="recorder" />;
}
