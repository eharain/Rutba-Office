// Making Rutba Office the default.
//
// What is actually possible here differs sharply by platform, and the honest
// thing is to do what each one allows and say plainly what happened rather than
// claim success and leave a button that appears to do nothing.
//
//   Windows   An application cannot make itself the default. Since Windows 8
//             the choice belongs to the person, in Settings, and any program
//             that could change it silently would be malware. So: register the
//             types (the installer does this), then open Settings at the right
//             page and say so.
//   Linux     `xdg-mime default` is exactly this, and works. We can do it.
//   macOS     Requires a native call this build does not make. Open the Finder
//             route instead — Get Info, Change All — and say that.
//
// The registration itself is real on Windows: the installer writes the class
// entries, and this reads them back so the panel can say which formats are
// actually pointing at us rather than guessing.

import { app, shell } from 'electron';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import path from 'node:path';
import { APPS_CHOICE_KEY, chosenApps } from '@rutba/office-formats/registry';
import { fileTypePlan, applyFileTypes } from './file-types.js';

const run = promisify(execFile);
const isWindows = process.platform === 'win32';
const isMac = process.platform === 'darwin';

/** The registry keeps a per-user choice that overrides the machine default. */
async function windowsHandlerFor(ext) {
  const key = `HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\.${ext}\\UserChoice`;
  try {
    const { stdout } = await run('reg', ['query', key, '/v', 'ProgId'], { windowsHide: true });
    return /ProgId\s+REG_SZ\s+(\S+)/.exec(stdout)?.[1] ?? null;
  } catch {
    // No user choice recorded: fall back to what the class root says.
    try {
      const { stdout } = await run('reg', ['query', `HKCR\\.${ext}`, '/ve'], { windowsHide: true });
      return /REG_SZ\s+(\S+)/.exec(stdout)?.[1] ?? null;
    } catch {
      return null;
    }
  }
}

async function linuxHandlerFor(mime) {
  try {
    const { stdout } = await run('xdg-mime', ['query', 'default', mime]);
    return stdout.trim() || null;
  } catch {
    return null;
  }
}

/** What was last brought in line, so it is done again only after a change or an update. */
const APPLIED_KEY = 'apps.fileTypesApplied';

export function createDefaultsService({ associations, associationsFor = null, stores = null }) {
  const ours = (handler) => Boolean(handler && /rutba/i.test(handler));

  /**
   * The types Windows is told about follow the apps chosen (Choose apps).
   * Only for an installed copy: a run from the source tree, or the checks,
   * leave the registry alone. An update's installer registers every type
   * again, so this runs once more after each one.
   */
  const follow = async ({ apps } = {}) => {
    if (!isWindows || !app.isPackaged || process.env.RUTBA_VERIFY || !associationsFor) return { skipped: true };
    const chosen = Array.isArray(apps) ? apps : chosenApps(stores?.settings.get(APPS_CHOICE_KEY, null));
    const result = await applyFileTypes(fileTypePlan(associations, associationsFor(chosen)));
    stores?.settings.set(APPLIED_KEY, `${app.getVersion()}|${chosen.join(',')}`);
    return { skipped: false, withdrawn: result.withdrawn.length, restored: result.restored.length };
  };

  /** At start-up: only when a choice was made and has not been applied to this version yet. */
  const followAtStart = async () => {
    const choice = stores?.settings.get(APPS_CHOICE_KEY, null);
    if (!choice) return { skipped: true };
    const chosen = chosenApps(choice);
    if (stores.settings.get(APPLIED_KEY, '') === `${app.getVersion()}|${chosen.join(',')}`) return { skipped: true };
    return follow({ apps: chosen });
  };

  return {
    follow,
    followAtStart,
    /**
     * Which of the formats we claim currently open with us.
     * @returns {{ platform, canSet, ours: number, total: number, formats: [] }}
     */
    status: async () => {
      const formats = [];
      for (const assoc of associations) {
        let handler = null;
        if (isWindows) handler = await windowsHandlerFor(assoc.ext);
        else if (!isMac) handler = await linuxHandlerFor(assoc.mimeType);
        formats.push({
          ext: assoc.ext,
          name: assoc.name,
          mime: assoc.mimeType,
          handler,
          isDefault: ours(handler),
        });
      }
      return {
        platform: process.platform,
        packaged: app.isPackaged,
        // Only Linux lets an application set this for itself.
        canSet: process.platform === 'linux',
        instructions: isWindows
          ? 'Windows keeps this choice for you: pick Rutba Office in Settings › Default apps. The button below opens it there.'
          : isMac
            ? 'macOS keeps this choice in Finder: select a file, press Command-I, choose Rutba Office under “Open with”, then Change All.'
            : 'These are set with xdg-mime, which this can do directly.',
        ours: formats.filter((f) => f.isDefault).length,
        total: formats.length,
        formats,
      };
    },

    /**
     * Do what this platform permits.
     * @returns {{ changed: number, opened: boolean, message: string }}
     */
    set: async ({ exts } = {}) => {
      const wanted = exts?.length ? associations.filter((a) => exts.includes(a.ext)) : associations;

      if (process.platform === 'linux') {
        const desktop = 'rutba-office.desktop';
        let changed = 0;
        for (const assoc of wanted) {
          try {
            await run('xdg-mime', ['default', desktop, assoc.mimeType]);
            changed++;
          } catch {
            // A type the desktop database does not know about; keep going.
          }
        }
        return {
          changed,
          opened: false,
          message: changed
            ? `${changed} file types now open with Rutba Office.`
            : 'Nothing could be set — the desktop entry may not be installed yet.',
        };
      }

      if (isWindows) {
        // `ms-settings:defaultapps` is the page; deep-linking to one application
        // is not something Windows exposes to the application itself.
        await shell.openExternal('ms-settings:defaultapps');
        return {
          changed: 0,
          opened: true,
          message: 'Settings is open. Choose Rutba Office there — Windows does not let a program set this for you.',
        };
      }

      if (isMac) {
        await shell.openPath(path.dirname(app.getPath('exe')));
        return {
          changed: 0,
          opened: true,
          message: 'macOS sets this in Finder: select a file, press Command-I, pick Rutba Office and choose Change All.',
        };
      }

      return { changed: 0, opened: false, message: 'This platform is not supported yet.' };
    },
  };
}
