// Rutba Office — the main process.
//
// Small on purpose. The shell package owns windows, the protocol, the menu and
// every platform capability; this file only says which app opens which file and
// hands the shell the two namespaces that know about documents and mail.

import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createShell, holdBlob, broadcast } from '@rutba/office-shell/electron/main';
import { appFor, kindFromExtension } from '@rutba/office-formats/sniff';
import { createDocumentService } from './documents.js';
import { createMailService } from './mail.js';

const here = path.dirname(fileURLToPath(import.meta.url));
const app = path.resolve(here, '..');

// Kept so the smoke run can drive the same services the windows do.
let services = null;

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

  namespaces: ({ stores, holdBlob: hold }) => (services = {
    doc: createDocumentService({ holdBlob: hold }),
    mail: createMailService({
      stores,
      holdBlob: hold,
      broadcast,
      userData: stores.dir,
    }),
  }),

  // `npm run smoke` boots this same process, photographs each app, and exits.
  onReady: process.env.RUTBA_OFFICE_SMOKE
    ? async ({ windows, stores }) => {
        const { runSmoke } = await import('./smoke.js');
        const ok = await runSmoke({ windows, outDir: process.env.RUTBA_SMOKE_OUT || path.join(app, 'build', 'smoke'), stores, mail: services?.mail });
        const { app: electronApp } = await import('electron');
        electronApp.exit(ok ? 0 : 1);
      }
    : undefined,
});

export { holdBlob };
