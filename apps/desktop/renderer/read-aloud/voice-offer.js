// The offer of an added voice beside Read Aloud: one icon, shown only while
// no speech engine is installed (no `speech` on the bridge), that opens
// Voices and languages. Read Aloud reads with the system's voice either way;
// nothing is downloaded until the person chooses a voice there.
//
// Video, the Recorder and Live Broadcast offer it the same way for their
// spoken text layers (D-2026-11-13-31), handing in the added voices they
// found, so an engine with no voice left counts as none, and words of their
// own, since only an added voice can speak there.

import React from 'react';
import { Button, t } from '@rutba/office-ui';

/**
 * Is there no added voice to offer a choice of yet? With `voices` (the
 * added voices, as `speech.voices` answers them), none among them counts too.
 */
export const offersVoice = (shell, voices = null) => !shell?.speech || (Array.isArray(voices) && voices.length === 0);

export function VoiceOffer({ shell, voices = null, title = null, className = '' }) {
  if (!offersVoice(shell, voices)) return null;
  return (
    <Button
      ghost
      icon="globe"
      className={`read-aloud-voice-offer${className ? ` ${className}` : ''}`}
      title={title || t('Add a voice: more natural voices for Read Aloud, from Voices and languages')}
      onClick={() => shell.win.create({ app: 'home', query: { voices: 1 } }).catch?.(() => {})}
    />
  );
}
