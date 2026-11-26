// Whether Windows itself lets Rutba Office show a notification.
//
// A notification Office hands to Windows is shown only where Windows'
// own switches allow it: Settings, System, Notifications, for the whole of
// Windows and for each app. With the first off nothing is shown by any app,
// and nothing says so: the person presses "Show one now" and sees nothing
// (the owner's own PC, 26 Nov: the switch for the whole of Windows was off).
// So the switches are read, and the settings window says what it found, with
// the way to Windows' page.
//
// They are read from where Windows keeps them, with reg.exe as the browsers
// and the printers are (browsers.js), and never written:
//   - the administrator's rule, NoToastApplicationNotification, for the
//     person or for the computer;
//   - the switch for the whole of Windows, ToastEnabled;
//   - the app's own switch, Enabled, under the application's id. The key is
//     made by Windows the first time the app shows a notification, so none
//     there is no block.
// Only a plain "off" is taken as one. A value that is not there, a reg that
// did not answer, or a Windows that keeps them elsewhere one day, is no
// block: Windows' own answer when a notification is handed to it (the
// notifier's `refused`) still stands behind this.

import { regRunner } from './browsers.js';

const POLICY = 'Software\\Policies\\Microsoft\\Windows\\CurrentVersion\\PushNotifications';
const ALL = 'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\PushNotifications';
const APPS = 'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Notifications\\Settings';

/** Windows' own page for them, as Windows documents its address. */
export const WINDOWS_PAGE = 'ms-settings:notifications';

/** A number reg.exe printed for a value, or null: not there, not a number, or no answer. */
export function dword(text, name) {
  if (typeof text !== 'string') return null;
  const safe = String(name).replace(/[^A-Za-z0-9_]/g, '');
  const found = new RegExp(`^\\s*${safe}\\s+REG_DWORD\\s+0x([0-9a-f]+)\\s*$`, 'im').exec(text);
  return found ? parseInt(found[1], 16) : null;
}

/**
 * What stands in the way, the widest first: 'policy' (the administrator's
 * rule), 'all' (the switch for the whole of Windows), 'app' (the switch for
 * this app), or null.
 * @param {{ policyUser?: string|null, policyMachine?: string|null, all?: string|null, app?: string|null }} texts what reg.exe printed for each
 */
export function blockOf({ policyUser = null, policyMachine = null, all = null, app = null } = {}) {
  if (dword(policyUser, 'NoToastApplicationNotification') === 1 || dword(policyMachine, 'NoToastApplicationNotification') === 1) return 'policy';
  if (dword(all, 'ToastEnabled') === 0) return 'all';
  if (dword(app, 'Enabled') === 0) return 'app';
  return null;
}

const ask = regRunner({ timeout: 8000 });

/**
 * Reads the switches. Answers 'policy', 'all', 'app' or null; null too
 * anywhere but Windows.
 * @param {{ appId: string, platform?: string, reg?: (args: string[]) => Promise<string|null> }} o `reg` for the tests
 */
export async function readToastBlock({ appId, platform = process.platform, reg = ask }) {
  if (platform !== 'win32') return null;
  const value = (key, name) => Promise.resolve().then(() => reg(['query', key, '/v', name])).catch(() => null);
  const [policyUser, policyMachine, all, app] = await Promise.all([
    value(`HKCU\\${POLICY}`, 'NoToastApplicationNotification'),
    value(`HKLM\\${POLICY}`, 'NoToastApplicationNotification'),
    value(ALL, 'ToastEnabled'),
    // An id is letters, digits, dots and dashes; anything else is not asked about.
    /^[\w.-]+$/.test(String(appId || '')) ? value(`${APPS}\\${appId}`, 'Enabled') : null,
  ]);
  return blockOf({ policyUser, policyMachine, all, app });
}
