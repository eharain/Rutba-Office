// The offer of an added voice beside Read Aloud: one icon, shown only while
// no speech engine is installed (no `speech` on the bridge), that opens
// Voices and languages. Read Aloud reads with the system's voice either way;
// nothing is downloaded until the person chooses a voice there.

import React from 'react';
import { Button, t } from '@rutba/office-ui';

/** Is there no added voice to offer a choice of yet? */
export const offersVoice = (shell) => !shell?.speech;

export function VoiceOffer({ shell }) {
  if (!offersVoice(shell)) return null;
  return (
    <Button
      ghost
      icon="globe"
      className="read-aloud-voice-offer"
      title={t('Add a voice: more natural voices for Read Aloud, from Voices and languages')}
      onClick={() => shell.win.create({ app: 'home', query: { voices: 1 } }).catch?.(() => {})}
    />
  );
}
