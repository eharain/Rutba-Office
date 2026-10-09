// The words the main process shows itself — the question before a window
// with unsaved work closes, the menu — in the language the windows are in.
//
// The same catalogue as the windows' (@rutba/office-ui/messages), with the
// language taken from the "language" setting, else the system's, each time
// the main process is about to speak: a language chosen in a window's app
// menu reaches the next question at once, and the menu at the next start.

import { app } from 'electron';
import { t, tn, setLanguage, registerCatalogue } from '@rutba/office-ui/messages';
import { CATALOGUES } from '@rutba/office-ui/catalogues';

for (const [tag, messages] of Object.entries(CATALOGUES)) registerCatalogue(tag, messages);

/** Settle the main process's language from the settings, before it says something. */
export function speakAs(stores) {
  let chosen = null;
  try {
    chosen = stores?.settings?.get('language', null) || null;
  } catch {
    chosen = null;
  }
  return setLanguage(chosen || (app.isReady() ? app.getLocale() : null) || 'en');
}

export { t, tn };
