// Rutba Office — the main process.
//
// Small on purpose. The shell package owns windows, the protocol, the menu and
// every platform capability; this file only says which app opens which file and
// hands the shell the two namespaces that know about documents and mail.

import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createShell, holdBlob, broadcast } from '@rutba/office-shell/electron/main';
import { appFor, kindFromExtension } from '@rutba/office-formats/sniff';
import { fileAssociations } from '@rutba/office-formats/registry';
import { createDocumentService } from './documents.js';
import { createMailService } from './mail.js';
import { createUpdateService } from './updates.js';
import { createOAuthService } from './oauth.js';
import { createPresentService } from './present.js';
import { createAnnouncementService } from './announce.js';
import { createDefaultsService } from './defaults.js';
import { createDiscoveryService } from './discover.js';

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

    return (services = {
      doc: createDocumentService({ holdBlob: hold }),
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
    if (process.env.RUTBA_OFFICE_VERIFY_EDIT) {
      const { verifyEditing } = await import('./verify-edit.js');
      const ok = await verifyEditing({ windows, doc: services.doc });
      const { app: electronApp } = await import('electron');
      return electronApp.exit(ok ? 0 : 1);
    }
    if (process.env.RUTBA_OFFICE_VERIFY_APPS) {
      const { verifyApps } = await import('./verify-apps.js');
      // Always seeded. A check that only passes because the developer happens
      // to have imported an archive last week is not a check.
      const { seedMail } = await import('./seed-mail.js');
      await seedMail({ stores, mail: services?.mail }).catch((e) => console.error('the mail fixture failed:', e.message));
      const ok = await verifyApps({ windows, doc: services.doc });
      const { app: electronApp } = await import('electron');
      return electronApp.exit(ok ? 0 : 1);
    }
    if (process.env.RUTBA_OFFICE_SMOKE) {
      const { runSmoke } = await import('./smoke.js');
      const ok = await runSmoke({
        windows,
        outDir: process.env.RUTBA_SMOKE_OUT || path.join(app, 'build', 'smoke'),
        stores,
        mail: services?.mail,
      });
      const { app: electronApp } = await import('electron');
      return electronApp.exit(ok ? 0 : 1);
    }
    return updates?.start();
  },
});

export { holdBlob };
