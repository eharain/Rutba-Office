// Who Rutba Office is to Windows: the application's id, the id Windows calls
// a notification's press back on, and whether this copy hands Windows any
// notification at all.
//
// Electron makes a notification work by itself, and to do so it writes to
// the person's computer the first time a notification (or the question
// whether one can be shown) is asked of it: a shortcut in the person's Start
// menu, named after the program file's product name, carrying the
// application's id and the press-back id, and a key for that id under
// HKCU\Software\Classes\CLSID that names the program file. A shortcut that is
// there but points elsewhere is rewritten to point at the copy running.
//
// For the person's installed copy that is right: the shortcut is the
// installer's own and stays as it was, with the press-back id added. For any
// other copy it is not. A copy run from the source is "Electron" to Windows:
// it left an Electron shortcut with Rutba Office's id in the Start menu,
// which took the installed copy's place in Windows' list of apps, and a new
// key for every run (24 Nov). A portable copy, or a built copy never
// installed, has the installed copy's product name, so it would turn the
// installed copy's own shortcut towards itself. So only the installed copy
// hands Windows a notification; the others count what waits on their icon
// and show none.
//
// A copy run from the source can be told to show them, for work on the
// notifications themselves: RUTBA_OFFICE_TOASTS=1. It then goes under an id
// of its own, so its Electron shortcut stands beside the installed copy's
// entry and never in its place.

/** The application's id, as the installer writes it on its shortcuts. */
export const APP_ID = 'co.techstyle.rutba.office';

/**
 * The id Windows calls a notification's press back on. One for every
 * installed copy, so a copy installed again over itself keeps it; left
 * unsaid, Electron draws a new one, with a new key, whenever the shortcut has
 * lost the last.
 */
export const TOAST_ACTIVATOR = '{74583220-6F39-4387-A371-FE35BACE5663}';

/** The switch that lets a copy run from the source show Windows notifications. */
export const TOASTS_SWITCH = 'RUTBA_OFFICE_TOASTS';

/**
 * @param {object} run
 * @param {boolean} run.packaged   a built copy, not one run from the source
 * @param {boolean} run.installed  the person's installed copy (installed-copy.js)
 * @param {string} run.platform
 * @param {object} [run.env]
 * @returns {{ appUserModelId: string, toastActivatorClsid: string|null, toasts: boolean }}
 */
export function appIdentity({ packaged, installed, platform, env = {} }) {
  // Only Windows is written to for a notification; elsewhere every copy shows them, as it always did.
  if (platform !== 'win32') return { appUserModelId: APP_ID, toastActivatorClsid: null, toasts: true };
  if (installed) return { appUserModelId: APP_ID, toastActivatorClsid: TOAST_ACTIVATOR, toasts: true };
  const asked = !packaged && String(env[TOASTS_SWITCH] || '') === '1';
  return { appUserModelId: asked ? `${APP_ID}.dev` : APP_ID, toastActivatorClsid: null, toasts: asked };
}
