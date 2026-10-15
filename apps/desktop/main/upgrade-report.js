// What this copy found in a profile an older version left behind, for the
// upgrade test (tools/upgrade-test.js).
//
// The test runs an older release over a profile of its own, then this one
// over the same profile with RUTBA_OFFICE_UPGRADE_REPORT=<file>. Before any
// window opens, this writes what the application read: its version, every
// setting, the recent list as the launcher shows it, and the update feed as
// the updater reads it. Then it quits. The judgement is the test's
// (tools/upgrade-check.js), from the file.
//
// The feed is read with the updater's own library and the same feed, with
// downloading off: reading it is the point, fetching an installer is not.
// RUTBA_UPGRADE_FEED=0 leaves it out, for a run with no network.

import fs from 'node:fs';
import { app } from 'electron';
import { FEED } from './updates.js';

async function readFeed() {
  if (!app.isPackaged) return { ok: false, error: 'An unpackaged copy has no updater to read the feed with.' };
  try {
    const mod = await import('electron-updater');
    const updater = mod.autoUpdater ?? mod.default?.autoUpdater;
    updater.autoDownload = false;
    updater.autoInstallOnAppQuit = false;
    updater.allowPrerelease = true;
    updater.logger = null;
    updater.setFeedURL(FEED);
    const result = await updater.checkForUpdates();
    const info = result?.updateInfo;
    if (!info) return { ok: false, error: 'The updater answered nothing.' };
    return {
      ok: true,
      version: info.version || null,
      releaseDate: info.releaseDate || null,
      files: (info.files || []).map((f) => ({ url: f.url, size: f.size ?? null, sha512: f.sha512 || null })),
    };
  } catch (err) {
    return { ok: false, error: String(err?.message || err).slice(0, 400) };
  }
}

export async function upgradeReport({ stores, out, feed = true }) {
  const report = {
    version: app.getVersion(),
    packaged: app.isPackaged,
    userData: stores.dir,
    settings: stores.settings.all(),
    recent: await stores.recent.list(),
    feed: feed ? await readFeed() : null,
  };
  fs.writeFileSync(out, JSON.stringify(report, null, 2));
  return true;
}
