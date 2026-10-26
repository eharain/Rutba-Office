// Rutba Office — the main process.
//
// Small on purpose. The shell package owns windows, the protocol, the menu and
// every platform capability; this file only says which app opens which file and
// hands the shell the two namespaces that know about documents and mail.

// First: which profile this run is and which template set it uses, before anything below reads either (template-set-choice.js).
import { SETS_DIR, bundledManifest, PROFILE_NAME, PROFILE } from './template-set-choice.js';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { app as electron, BrowserWindow } from 'electron';
import { createShell, holdBlob, releaseBlob, blobUrl, broadcast } from '@rutba/office-shell/electron/main';
import { useClipboard, memoryClipboard } from '@rutba/office-shell/electron/clipboard';
import { useShortcuts } from '@rutba/office-shell/electron/shortcuts';
import { fakeShortcuts } from '@rutba/office-shell/electron/machine-fakes';

// A check run keeps to a clipboard and global shortcuts of its own, so the
// gate's check processes can run side by side and none touches the person's;
// the gate's machine lane alone uses the computer's, to check them.
// A second profile the profile checks start beside them is a check run too (verify-profiles.js).
const PROFILE_PEER = process.env.RUTBA_PROFILE_PEER || null;
const CHECK_RUN = Boolean(process.env.RUTBA_OFFICE_VERIFY_EDIT || process.env.RUTBA_OFFICE_VERIFY_APPS || process.env.RUTBA_OFFICE_VERIFY_CORPUS || process.env.RUTBA_OFFICE_SMOKE || process.env.RUTBA_SCREENS_FEATURES || PROFILE_PEER);
if (CHECK_RUN && process.env.RUTBA_VERIFY_LANE !== 'machine') {
  useClipboard(memoryClipboard());
  useShortcuts(fakeShortcuts());
}
import { appFor, kindFromExtension } from '@rutba/office-formats/sniff';
import { fileAssociations, APPS, APPS_CHOICE_KEY, OFF_UNTIL_CHOSEN, APP_ORDER, APP_EXTENSIONS, OLD_EXTENSIONS } from '@rutba/office-formats/registry';
import { createDocumentHost } from './doc-host.js';
import { createPrintService } from './print.js';
import { createMathMeasurer } from './math-raster.js';
import { createUpdateService } from './updates.js';
import { createTemplateSets } from './template-sets.js';
import { createOAuthService } from './oauth.js';
import { createAnnouncementService } from './announce.js';
import { createDefaultsService } from './defaults.js';
import { createDavService } from './dav-sync.js';
import { createPacksService } from './packs-service.js';
import { planPacks, registerPacks, packView, readPackKeys, readPackSizes, PACK_ORDER } from './app-packs.js';
import { unicodeFont } from './system-fonts.js';
import { LOADED_KEY } from './pack-tools.js';
import { recordPacks } from './file-types.js';
import { installPackResolution } from '@rutba/office-shell/electron/pack-resolve';

const here = path.dirname(fileURLToPath(import.meta.url));
const app = path.resolve(here, '..');

// Kept so the smoke run can drive the same services the windows do.
let services = null;

/**
 * The apps on this computer (app-packs.js, docs/APP-PACKS.md). An installed
 * copy keeps its packs beside the archive, as the installer and Add or remove
 * apps wrote them; a check run may name a folder of its own. With none, a copy
 * run from the source or the portable build, every app is built in.
 */
function packsDir() {
  if (process.env.RUTBA_PACKS_DIR) return path.resolve(process.env.RUTBA_PACKS_DIR);
  if (!electron.isPackaged) return null;
  const dir = path.join(process.resourcesPath, 'packs');
  return fs.existsSync(dir) ? dir : null;
}

/**
 * A copy that carries every pack and has no installer to choose in: the
 * portable build (run from wherever it was unpacked) and the macOS and Linux
 * builds (build/dist.js). It loads its packs like an installed copy, but has
 * nothing to add or remove and writes nothing for an installer to read.
 */
const everyPack = () => electron.isPackaged && !process.env.RUTBA_PACKS_DIR && (process.platform !== 'win32' || Boolean(process.env.PORTABLE_EXECUTABLE_DIR));

/** The keys a pack must be signed with: the build's, in the core; a copy run from the source may be handed a development key's public half. */
const PACK_KEYS_FILE = path.join(app, 'build', 'out', 'pack-keys.json');
function packKeys() {
  const keys = readPackKeys(PACK_KEYS_FILE);
  if (!electron.isPackaged && process.env.RUTBA_PACK_PUBLIC_KEY) keys.push(fs.readFileSync(process.env.RUTBA_PACK_PUBLIC_KEY, 'utf8'));
  return keys;
}

// The taskbar reads an app's tile itself, so in a packaged copy the tiles live unpacked beside the archive.
const builtInTiles = path.join(app, 'resources', 'apps').replace(`${path.sep}app.asar${path.sep}`, () => `${path.sep}app.asar.unpacked${path.sep}`);
let packs = null;
/** The packs this run loads, found once, before the first window. */
const packsNow = () => (packs ||= planPacks({ dir: packsDir(), keys: packKeys(), version: electron.getVersion(), tilesDir: builtInTiles }));
// Each app's taskbar tile, from its pack: filled in once the packs are found.
const appIcons = {};

/**
 * The app a window is made for when the one asked for is not installed:
 * another installed app that opens the file, in the registry's order, or
 * the launcher, which says which app to add.
 */
function resolveApp({ app: key, file }) {
  const p = packsNow();
  if (!key || key === 'home' || !APPS[key] || p.has(key)) return null;
  const ext = file ? path.extname(String(file)).slice(1).toLowerCase() : '';
  const other = ext ? APP_ORDER.find((k) => p.has(k) && [...(APP_EXTENSIONS[k] || []), ...(OLD_EXTENSIONS[k] || [])].includes(ext)) : null;
  return other ? { app: other } : { app: 'home', query: { missing: key } };
}

/**
 * Newer template sets, fetched and checked with Office's own updates
 * (D-2026-10-21-24): signed by Rutba's key, used from the next start when
 * this Office is new enough for them, kept beside the profile.
 */
function templateSets() {
  const fetchFrom = (url) => fetch(url, { headers: { 'user-agent': 'Rutba-Office', accept: 'application/vnd.github+json' } }).then((r) => {
    if (!r.ok) throw new Error(`The template sets could not be fetched (${r.status}).`);
    return r;
  });
  return createTemplateSets({
    root: SETS_DIR(),
    bundled: bundledManifest(),
    office: process.env.RUTBA_OFFICE_VERSION || electron.getVersion(),
    fetchJson: (url) => fetchFrom(url).then((r) => r.json()),
    fetchText: (url) => fetchFrom(url).then((r) => r.text()),
  });
}
let updates = null;

// The document service runs on a thread of its own (doc-host.js, doc-worker.js).
// Started here, its engines load while the shell gets the first window ready.
const docHost = createDocumentHost({ holdBlob, releaseBlob, blobUrl }).start();

/** What a pack wants told when a recent file is renamed (Data Files: what a database may read goes with it). */
const renameHooks = [];

/**
 * Renaming a recent file from the launcher: the file on disk moves, and the
 * recent list is told where it went. Refused — with a plain sentence, never
 * a stack trace — when the file is missing, when a file of that name
 * already exists, when the name is empty or has a path separator in it, or
 * when the document is open in a window; an open session has its own idea
 * of the path and would go on saving to a file that no longer answers to
 * that name, so that one has to close first.
 */
async function renameRecent({ stores, doc, path: from, name }) {
  const trimmed = String(name || '').trim();
  if (!from) throw new Error('No file was given to rename.');
  if (!trimmed) throw new Error('Enter a name.');
  if (/[\\/]/.test(trimmed)) throw new Error('The name cannot contain a path separator.');
  if (!fs.existsSync(from)) throw new Error('That file no longer exists.');

  const dir = path.dirname(from);
  const finalName = path.extname(trimmed) ? trimmed : `${trimmed}${path.extname(from)}`;
  const to = path.join(dir, finalName);

  if (to !== from && fs.existsSync(to)) throw new Error('A file with that name already exists.');
  if ((await doc.sessions()).some((s) => s.path === from)) throw new Error('That file is open in a window. Close it first.');

  if (to !== from) fs.renameSync(from, to);
  // The packs that keep something by a file's path are told: a database renamed here reads what it read before (A2.18).
  if (to !== from) {
    for (const hook of renameHooks) {
      try {
        hook(from, to);
      } catch (err) {
        console.error('rename hook:', err?.message || err);
      }
    }
  }
  return stores.recent.rename({ path: from, to });
}

// This Office's version, for every thread to read: a template says the oldest Office it
// needs (Rutba-Office-Templates' manifest), and the document worker lists the gallery.
// A check may stand in another (RUTBA_OFFICE_VERSION) to see what an older Office shows.
process.env.RUTBA_APP_VERSION = process.env.RUTBA_OFFICE_VERSION || electron.getVersion();

// A screenshot build asks for a device scale of its own (`RUTBA_SCREEN_SCALE=2`),
// so a capture is crisp at twice the window's size whatever display it ran on.
if (process.env.RUTBA_SCREEN_SCALE) electron.commandLine.appendSwitch('force-device-scale-factor', process.env.RUTBA_SCREEN_SCALE);
// A check run's microphone is Chromium's own stand-in: a tone, so Record has something to record.
if (process.env.RUTBA_FAKE_MEDIA) electron.commandLine.appendSwitch('use-fake-device-for-media-stream');

// The profile checks' second profile reports what it sees, and answers its own password window (verify-profiles-peer.js).
const peer = PROFILE_PEER ? import('./verify-profiles-peer.js').then((m) => m.startPeer({ report: PROFILE_PEER, profile: PROFILE })) : null;

createShell({
  // The profile's folder is named by it too (template-set-choice.js works that folder out before this runs).
  appName: PROFILE_NAME,
  // Which profile this is: its folder is already Electron's (profile-choice.js), its windows show it.
  profile: PROFILE,
  rendererDir: path.join(app, 'build', 'out'),
  preloadPath: path.join(app, 'build', 'out', 'preload.cjs'),
  iconPath: path.join(app, 'resources', process.platform === 'win32' ? 'icon.ico' : 'icon.png'),
  // One tile per app for the taskbar (build/make-file-icons.js), from the
  // installed packs once they are found: the taskbar reads the file itself,
  // so it lives unpacked beside an archive.
  appIcons,
  appNames: Object.fromEntries(Object.entries(APPS).map(([key, a]) => [key, a.name])),
  appShortNames: Object.fromEntries(Object.entries(APPS).map(([key, a]) => [key, a.short])),

  /**
   * Which window a double-clicked file opens. The extension decides, because
   * this runs before the file is read — the deeper sniff happens once the
   * document service actually opens it, and can still redirect. An app that
   * is not installed is never opened (resolveApp).
   */
  appForFile: (file) => appFor(kindFromExtension(file)) || 'home',
  resolveApp,
  // What this run offers its windows: the core's channels and the loaded packs' (app-packs.js).
  offers: () => packsNow().offer(),
  // A pack that carries its own window code is served from its archive.
  packRoots: () => Object.fromEntries(packsNow().packs.filter((p) => p.archive).map((p) => [p.id, p.archive])),

  namespaces: async ({ stores, holdBlob: hold }) => {
    // The apps on this computer, before anything is built for them; what was
    // loaded is written down for the AI tools (pack-tools.js).
    const plan = packsNow();
    for (const p of plan.packs) {
      const tile = p.tiles?.find((t) => t.app === p.id)?.icon;
      if (tile) appIcons[p.id] = tile;
    }
    for (const r of plan.refused) console.error(`[packs] ${r.reason}`);
    // Built before the mail service, which needs it to fetch an access token
    // for an account that was added by signing in rather than by typing a
    // password.
    const oauth = createOAuthService({ stores, broadcast });
    // Equations printed as Chromium lays their MathML out — see math-raster.js.
    const math = createMathMeasurer();
    // What the document service's thread asks of the main process: an
    // equation laid out by Chromium, a word taught to a window's spelling
    // underline, a proofing setting written to the store.
    docHost.provide({
      measureMath: (list) => math.measure(list),
      teach: (winId, word, add) => {
        const ses = winId == null ? null : BrowserWindow.fromId(winId)?.webContents?.session;
        if (!ses) return;
        if (add) ses.addWordToSpellCheckerDictionary?.(word);
        else ses.removeWordFromSpellCheckerDictionary?.(word);
      },
      setting: (key, value) => stores.settings.set(key, value),
    });
    // Unsaved work is written to a copy in the profile every half minute,
    // and the copy is deleted the moment the document is saved or closed. What
    // is left in that folder at start-up is what a crash took.
    docHost.init({ recoveryDir: path.join(stores.dir, 'recovery'), locale: electron.getLocale?.() || 'en-GB', settings: stores.settings.all() });
    const { doc } = await docHost.ready();
    stores.settings.set(LOADED_KEY, plan.ids());

    // The apps chosen on the installer's page, taken over before any window
    // reads the choice. The kinds of file Windows is told about are those of
    // the apps installed, and of those, the apps chosen.
    const installed = (apps) => (apps || APP_ORDER).filter((k) => plan.has(k));
    const defaults = createDefaultsService({ associations: fileAssociations({ apps: installed() }), associationsFor: (apps) => fileAssociations({ apps: installed(apps) }), stores });
    await defaults.adoptInstallerChoice().catch(() => {});

    services = {
      // Every call answered from the service's thread. A window's Open waits
      // on the file and its unzipping (documents.js, `openAsync`), and its
      // Save compresses the parts on the thread pool (`saveAsync`), so
      // neither holds the thread's other documents either. The checks,
      // which read a file the moment they save it, are handed the service
      // as a function call (`directDoc` below). A recovered copy's password
      // is tried on the password threads (`recoverAsync`), as Open's is.
      doc: { ...doc, open: doc.openAsync, save: doc.saveAsync, saveNow: doc.save, recover: doc.recoverAsync },
      // Paper and PDFs, for every kind of document. It asks the document
      // service where the pages fall and hands the result to a hidden window.
      print: createPrintService({ docs: doc }),
      update: updates = createUpdateService({
        stores,
        broadcast,
        // A newer signed template set is looked for with Office's own update (template-sets.js).
        templates: templateSets(),
        // A download shows on every window's taskbar button while it runs.
        taskbar: (fraction) => {
          for (const w of BrowserWindow.getAllWindows()) {
            try {
              w.setProgressBar(fraction);
            } catch {
              // A window on its way out.
            }
          }
        },
      }),
      announce: createAnnouncementService({ stores, broadcast }),
      defaults,
      oauth,
      // Renaming touches the file system and the document service both,
      // which the shell's own `app` namespace knows about neither — so this
      // adds just the one method, on top of recent(), addRecent() and
      // removeRecent() the shell already provides.
      app: {
        renameRecent: (p) => renameRecent({ stores, doc, ...p }),
      },
    };

    // Each installed app registers itself (app-packs.js, <id>-pack.js), handed
    // the core's stores and services, and `packs`, through which it reaches
    // another app's service only when that app is there. Mail before Contacts
    // or after makes no difference: each reaches the other when asked.
    const view = packView({ plan, services });
    // A pack that carries its code finds the core's packages as the core does (pack-resolve.js).
    if (plan.source === 'files') installPackResolution({ packsDir: plan.dir, coreFile: fileURLToPath(import.meta.url) });
    await registerPacks(plan, {
      stores,
      broadcast,
      holdBlob: hold,
      doc,
      oauth,
      packs: view,
      onRecentRenamed: (fn) => renameHooks.push(fn),
      // A system font that draws Arabic and Hebrew, for a PDF a pack writes (Mail's archive).
      unicodeFont,
    }, { services, record: (kind, err) => console.error(`${kind}: ${err?.stack || err}`) });

    // Calendar and contacts accounts: CalDAV and CardDAV servers, kept in step
    // with whichever of the two apps are installed (dav-sync.js keeps to those).
    services.dav = createDavService({ stores, calendar: view.service('calendar'), contacts: view.service('contacts'), broadcast });
    // Add or remove apps (packs-service.js).
    services.packs = createPacksService({ plan, stores, broadcast, version: electron.getVersion(), keys: packKeys(), dir: plan.source === 'files' && !everyPack() ? plan.dir : null, sizes: readPackSizes(PACK_KEYS_FILE), app: electron });
    return services;
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
    if (process.env.RUTBA_OFFICE_VERIFY_EDIT || process.env.RUTBA_OFFICE_VERIFY_APPS || process.env.RUTBA_OFFICE_VERIFY_CORPUS || process.env.RUTBA_OFFICE_SMOKE || PROFILE_PEER) {
      stores.settings.set('announcements.enabled', false);
      stores.settings.set('updates.automatic', false);
      // Nor does it take its look from the machine. Left on 'system', a
      // computer in dark mode drew every window dark, and a check that reads
      // an unfilled cell as white or clear failed there and passed everywhere
      // else — a verdict on the operating system's setting, not the code.
      // `RUTBA_SMOKE_THEME` still asks for the dark captures on purpose.
      stores.settings.set('theme', 'light');
      // And its words in English, which is what the checks read, whatever the
      // machine speaks; `RUTBA_LANGUAGE=ur` asks for another on purpose (the
      // Urdu captures, the Urdu checks).
      stores.settings.set('language', process.env.RUTBA_LANGUAGE || 'en');
      // And every app and add-on, as a person who chose them all: the checks
      // and the captures reach every app from the launcher. Choose apps' own
      // check takes this away to see a first start.
      stores.settings.set(APPS_CHOICE_KEY, { off: [], on: OFF_UNTIL_CHOSEN.slice(), addons: ['ai'] });
    }

    // The packs installed, written where the next installer reads them, so it
    // opens on what is here (docs/APP-PACKS.md 5.3), a pack that would not
    // load among them, so the next installer writes it again. An installed copy only.
    const plan = packsNow();
    if (plan.source === 'files' && process.platform === 'win32' && electron.isPackaged && !everyPack() && !CHECK_RUN) {
      recordPacks([...new Set([...plan.ids(), ...plan.refused.map((r) => r.id)])].filter((id) => PACK_ORDER.includes(id))).catch(() => {});
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
        if (!services?.mail) throw new Error('Mail is not installed on this computer: add it with Add or remove apps.');
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
        Promise.resolve()
          .then(() => services.doc?.autosave?.())
          .catch((err) => console.error('autosave:', err?.message || err));
      }, Number(process.env.RUTBA_AUTOSAVE_MS || 30000));
      timer.unref?.();
    }

    // A closed window frees the documents it opened. The session is the
    // engine and everything it holds — a deck's pictures, a document's
    // pages — and nothing let go of it before this.
    const { app: electronApp, BrowserWindow } = await import('electron');
    electronApp.on('browser-window-created', (_e, win) => {
      const id = win.id;
      win.once('closed', async () => {
        const gone = (await Promise.resolve(services.doc?.closeWindow?.(id)).catch(() => [])) ?? [];
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
      // The code a run ends with is its verdict; say what was handed over and
      // what the process ended on, so an exit code the run never chose shows.
      process.once('exit', (code) => console.log(`     [exit] the process ended with ${code}`));
      try {
        const code = (await run()) ? 0 : 1;
        console.log(`     [exit] the run hands app.exit ${code}`);
        return electronApp.exit(code);
      } catch (err) {
        console.error(`the check run stopped: ${err?.stack || err?.message || err}`);
        return electronApp.exit(1);
      }
    };

    // The document service as the checks call it: a function that answers,
    // saving at once, as their reads expect; opening as a window opens.
    const directDoc = () => ({ ...docHost.direct, open: docHost.direct.openAsync, saveNow: docHost.direct.save });
    // tools/upgrade-test.js: what this copy read from the profile an older
    // version left, and the update feed, written to a file before any window.
    if (process.env.RUTBA_OFFICE_UPGRADE_REPORT) {
      return finish(async () => {
        const { upgradeReport } = await import('./upgrade-report.js');
        return upgradeReport({ stores, out: process.env.RUTBA_OFFICE_UPGRADE_REPORT, feed: process.env.RUTBA_UPGRADE_FEED !== '0' });
      });
    }
    if (process.env.RUTBA_OFFICE_VERIFY_EDIT) {
      return finish(async () => {
        const { verifyEditing } = await import('./verify-edit.js');
        return verifyEditing({ windows, doc: directDoc() });
      });
    }
    if (process.env.RUTBA_OFFICE_VERIFY_APPS) {
      return finish(async () => {
        const { verifyApps } = await import('./verify-apps.js');
        // Always seeded. A check that only passes because the developer happens
        // to have imported an archive last week is not a check.
        const { seedMail } = await import('./seed-mail.js');
        await seedMail({ stores, mail: services?.mail }).catch((e) => console.error('the mail fixture failed:', e.message));
        return verifyApps({ windows, doc: directDoc(), broadcast, update: updates, idle: (limit) => docHost.idle(limit) });
      });
    }
    if (process.env.RUTBA_OFFICE_VERIFY_CORPUS) {
      // Every file in a folder, opened for real, one window at a time.
      return finish(async () => {
        const { verifyCorpus } = await import('./verify-corpus.js');
        return verifyCorpus({ windows, doc: directDoc(), appForFile: (file) => appFor(kindFromExtension(file)) || 'home' });
      });
    }
    if (process.env.RUTBA_OFFICE_SMOKE) {
      return finish(async () => {
        // `RUTBA_SCREENS_FEATURES`: the screens build's feature captures,
        // each a window driven into the state that shows one feature.
        if (process.env.RUTBA_SCREENS_FEATURES) {
          const { runFeatureScreens } = await import('./screen-run.js');
          if (process.env.RUTBA_SMOKE_THEME) stores.settings.set('theme', process.env.RUTBA_SMOKE_THEME);
          return runFeatureScreens({
            windows,
            outDir: process.env.RUTBA_SMOKE_OUT || path.join(app, 'build', 'screens', 'features'),
            samples: process.env.RUTBA_SCREENS_SAMPLES,
            doc: directDoc(),
            stores,
            mail: services?.mail,
            broadcast,
          });
        }
        const { runSmoke } = await import('./smoke.js');
        return runSmoke({
          windows,
          outDir: process.env.RUTBA_SMOKE_OUT || path.join(app, 'build', 'smoke'),
          stores,
          mail: services?.mail,
        });
      });
    }
    // The profile checks' second profile: what it was asked to do, beside the first, and no updater.
    if (peer) {
      (await peer).ready({ stores, services, windows });
      return undefined;
    }
    // The kinds of file Windows sends here follow the apps chosen, again after an update's installer
    // registered them all: a little after start, out of the way of the first window.
    setTimeout(() => services?.defaults?.followAtStart().catch(() => {}), 8000).unref?.();
    return updates?.start();
  },
});

export { holdBlob };
