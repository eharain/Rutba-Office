// Rutba Office — the main process.
//
// Small on purpose. The shell package owns windows, the protocol, the menu and
// every platform capability; this file only says which app opens which file and
// hands the shell the two namespaces that know about documents and mail.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { app as electron } from 'electron';
import { createShell, holdBlob, broadcast } from '@rutba/office-shell/electron/main';
import { appFor, kindFromExtension } from '@rutba/office-formats/sniff';
import { fileAssociations, APPS } from '@rutba/office-formats/registry';
import { createDocumentService } from './documents.js';
import { createPrintService } from './print.js';
import { createMailService } from './mail.js';
import { createUpdateService } from './updates.js';
import { createOAuthService } from './oauth.js';
import { createPresentService } from './present.js';
import { createAnnouncementService } from './announce.js';
import { createDefaultsService } from './defaults.js';
import { createDiscoveryService } from './discover.js';
import { createContactsService } from './contacts.js';
import { createCalendarService } from './calendar.js';

const here = path.dirname(fileURLToPath(import.meta.url));
const app = path.resolve(here, '..');

// Kept so the smoke run can drive the same services the windows do.
let services = null;
let updates = null;

createShell({
  appName: 'Rutba Office',
  rendererDir: path.join(app, 'build', 'out'),
  preloadPath: path.join(app, 'build', 'out', 'preload.cjs'),
  iconPath: path.join(app, 'resources', process.platform === 'win32' ? 'icon.ico' : 'icon.png'),
  // One tile per app for the taskbar (build/make-file-icons.js). The taskbar
  // reads the file itself, so in a packaged copy it lives unpacked beside
  // the archive, and the path says so.
  appIcons: Object.fromEntries(Object.keys(APPS).map((key) => [key, path.join(app, 'resources', 'apps', `${key}.ico`).replace(`${path.sep}app.asar${path.sep}`, `${path.sep}app.asar.unpacked${path.sep}`)])),
  appNames: Object.fromEntries(Object.entries(APPS).map(([key, a]) => [key, a.name])),

  /**
   * Which window a double-clicked file opens. The extension decides, because
   * this runs before the file is read — the deeper sniff happens once the
   * document service actually opens it, and can still redirect.
   */
  appForFile: (file) => appFor(kindFromExtension(file)) || 'home',

  namespaces: ({ stores, holdBlob: hold }) => {
    // Built before the mail service, which needs it to fetch an access token
    // for an account that was added by signing in rather than by typing a
    // password.
    const oauth = createOAuthService({ stores, broadcast });
    // Unsaved work is written to a copy in the profile every half minute,
    // and the copy is deleted the moment the document is saved or closed. What
    // is left in that folder at start-up is what a crash took.
    const doc = createDocumentService({ holdBlob: hold, recoveryDir: path.join(stores.dir, 'recovery') });

    return (services = {
      doc,
      // Paper and PDFs, for every kind of document. It asks the document
      // service where the pages fall and hands the result to a hidden window.
      print: createPrintService({ docs: doc }),
      update: updates = createUpdateService({ stores, broadcast }),
      announce: createAnnouncementService({ stores, broadcast }),
      defaults: createDefaultsService({ associations: fileAssociations() }),
      discover: createDiscoveryService(),
      present: createPresentService({ broadcast }),
      oauth,
      mail: createMailService({
        stores,
        holdBlob: hold,
        broadcast,
        userData: stores.dir,
        oauth,
      }),
      // The address book, with the people mail has seen behind it for
      // Compose to complete from; and the calendar, which answers an
      // invitation through a message mail sends.
      contacts: createContactsService({
        stores,
        broadcast,
        people: () => {
          try {
            const out = [];
            for (const account of services?.mail?.accounts?.() || []) {
              const { rows = [] } = services.mail.people({ accountId: account.id, limit: 300 }) || {};
              for (const p of rows) out.push({ name: p.name || '', email: p.address });
            }
            return out;
          } catch {
            return [];
          }
        },
      }),
      calendar: createCalendarService({ stores, broadcast, mail: { accounts: () => services?.mail?.accounts?.() || [] } }),
    });
  },

  /**
   * After the first window is up.
   *
   * Normally that means starting the update rhythm — quietly, and not for
   * twenty-five seconds, because the first moments after launch belong to
   * whatever the person opened the application to do.
   *
   * The two verification runs take this over instead: `npm run smoke`
   * photographs every window, `npm run verify:edit` types into two of them.
   * Neither should ever contact GitHub, so neither starts the updater.
   */
  onReady: async ({ windows, stores }) => {
    // A check run contacts nothing: not GitHub for updates (never started
    // here) and not office.rutba.io for the notice board. Its launcher used
    // to make that request, and an aborted one threw a dialog onto the
    // owner's screen in the middle of a run.
    if (process.env.RUTBA_OFFICE_VERIFY_EDIT || process.env.RUTBA_OFFICE_VERIFY_APPS || process.env.RUTBA_OFFICE_VERIFY_CORPUS || process.env.RUTBA_OFFICE_SMOKE) {
      stores.settings.set('announcements.enabled', false);
      stores.settings.set('updates.automatic', false);
    }

    // `--import-accounts=<file>` sets up the accounts in a file before the
    // first window: how an administrator's file, or another client's,
    // becomes a working Mail without a dialog per address. `--accounts=a,b`
    // picks addresses from it, `--sync` fetches each new inbox, and with
    // `--quit` the application leaves rather than opening a window. What
    // happened is printed — addresses and outcomes, never a password.
    const importFlag = process.argv.find((a) => a.startsWith('--import-accounts='));
    if (importFlag) {
      const file = importFlag.slice('--import-accounts='.length).replace(/^"|"$/g, '');
      const onlyFlag = process.argv.find((a) => a.startsWith('--accounts='));
      const only = onlyFlag ? onlyFlag.slice('--accounts='.length).split(',').map((e) => e.trim()).filter(Boolean) : null;
      const report = { file };
      try {
        report.result = await services.mail.importAccounts({ path: file, only, test: true });
        if (process.argv.includes('--sync')) {
          report.synced = [];
          for (const added of report.result.added) {
            try {
              const r = await services.mail.sync({ accountId: added.id, limit: 100 });
              report.synced.push({ email: added.email, added: r.added, total: r.total });
            } catch (err) {
              report.synced.push({ email: added.email, error: err?.message || String(err) });
            }
          }
        }
      } catch (err) {
        report.error = err?.message || String(err);
      }
      // Written synchronously: on Windows a pipe is asynchronous, and an
      // exit right after console.log can lose the line.
      try {
        fs.writeSync(1, `[import-accounts] ${JSON.stringify(report)}\n`);
      } catch {
        console.log(`[import-accounts] ${JSON.stringify(report)}`);
      }
      if (process.argv.includes('--quit')) {
        setTimeout(() => electron.exit(report.error ? 1 : 0), 300);
        return undefined;
      }
    }

    // Every half minute, whatever is unsaved. Nothing is written for a
    // document that has not changed since its last copy, so an open workbook
    // costs nothing to leave open.
    if (!process.env.RUTBA_OFFICE_VERIFY_CORPUS) {
      const timer = setInterval(() => {
        try {
          services.doc?.autosave?.();
        } catch (err) {
          console.error('autosave:', err?.message || err);
        }
      }, Number(process.env.RUTBA_AUTOSAVE_MS || 30000));
      timer.unref?.();
    }

    // A closed window frees the documents it opened. The session is the
    // engine and everything it holds — a deck's pictures, a document's
    // pages — and nothing let go of it before this.
    const { app: electronApp, BrowserWindow } = await import('electron');
    electronApp.on('browser-window-created', (_e, win) => {
      const id = win.id;
      win.once('closed', () => {
        const gone = services.doc?.closeWindow?.(id) ?? [];
        if (!gone.length) return;
        // A presenter window shows a deck its editor holds open. With the
        // editor gone the deck is gone, and a window left showing it would
        // fail every call it made: the show ends with its editor.
        for (const other of BrowserWindow.getAllWindows()) {
          let presenting = null;
          try {
            presenting = new URL(other.webContents.getURL()).searchParams.get('presenter');
          } catch {
            presenting = null;
          }
          if (presenting && gone.includes(presenting)) other.close();
        }
      });
    });

    // A check run ends by exiting, pass or fail — including when the check
    // itself throws. Without this, an error raised before the first window
    // (a corpus run pointed at no folder, say) left an application with no
    // windows, which never quits: `npm run verify:corpus` hung for as long as
    // anyone let it, having printed the reason and then waited forever.
    const finish = async (run) => {
      const { app: electronApp } = await import('electron');
      try {
        return electronApp.exit((await run()) ? 0 : 1);
      } catch (err) {
        console.error(`the check run stopped: ${err?.stack || err?.message || err}`);
        return electronApp.exit(1);
      }
    };

    if (process.env.RUTBA_OFFICE_VERIFY_EDIT) {
      return finish(async () => {
        const { verifyEditing } = await import('./verify-edit.js');
        return verifyEditing({ windows, doc: services.doc });
      });
    }
    if (process.env.RUTBA_OFFICE_VERIFY_APPS) {
      return finish(async () => {
        const { verifyApps } = await import('./verify-apps.js');
        // Always seeded. A check that only passes because the developer happens
        // to have imported an archive last week is not a check.
        const { seedMail } = await import('./seed-mail.js');
        await seedMail({ stores, mail: services?.mail }).catch((e) => console.error('the mail fixture failed:', e.message));
        return verifyApps({ windows, doc: services.doc });
      });
    }
    if (process.env.RUTBA_OFFICE_VERIFY_CORPUS) {
      // Every file in a folder, opened for real, one window at a time.
      return finish(async () => {
        const { verifyCorpus } = await import('./verify-corpus.js');
        return verifyCorpus({ windows, doc: services.doc, appForFile: (file) => appFor(kindFromExtension(file)) || 'home' });
      });
    }
    if (process.env.RUTBA_OFFICE_SMOKE) {
      return finish(async () => {
        const { runSmoke } = await import('./smoke.js');
        return runSmoke({
          windows,
          outDir: process.env.RUTBA_SMOKE_OUT || path.join(app, 'build', 'smoke'),
          stores,
          mail: services?.mail,
        });
      });
    }
    return updates?.start();
  },
});

export { holdBlob };
