// Whether this run is the person's installed copy, the one that may change
// what Windows starts for them: the rutba-office:// links (a sign-in's return
// to Mail) and the pins and shortcuts (pin-repair.js).
//
// A built copy that was never installed (release\win-unpacked) is packaged
// too, and a test run of it took the links for itself on the owner's
// computer (16 Nov): a sign-in would have come back to the test build. So
// the installer's mark is asked for, its uninstaller beside the program, and
// a run that is a test is never the person's copy: a scratch
// `--user-data-dir`, or a check run's switch. Windows never adds that
// switch, nor does our own relaunch; a person who starts their copy with it
// on purpose keeps the links and pins as they were.

import { win32 as path } from 'node:path';

export const UNINSTALLER = 'Uninstall Rutba Office.exe';

const CHECK_SWITCHES = ['RUTBA_OFFICE_VERIFY_EDIT', 'RUTBA_OFFICE_VERIFY_APPS', 'RUTBA_OFFICE_VERIFY_CORPUS', 'RUTBA_OFFICE_SMOKE', 'RUTBA_SCREENS_FEATURES', 'RUTBA_PROFILE_PEER'];

/**
 * @param {{ isPackaged: boolean, platform: string, exe: string, argv: string[], env: object, exists: (file: string) => boolean }} run
 */
export function isInstalledCopy({ isPackaged, platform, exe, argv, env, exists }) {
  if (!isPackaged || platform !== 'win32') return false;
  if (env.PORTABLE_EXECUTABLE_DIR) return false;
  if (CHECK_SWITCHES.some((k) => env[k])) return false;
  if (argv.some((a) => /^--user-data-dir(=|$)/i.test(a))) return false;
  return exists(path.join(path.dirname(exe), UNINSTALLER));
}
