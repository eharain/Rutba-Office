// What this copy found in a profile an older version left behind, for the
// upgrade test (tools/upgrade-test.js).
//
// The test runs an older release over a profile of its own, then this one
// over the same profile with RUTBA_OFFICE_UPGRADE_REPORT=<file>. This writes
// what the application read: its version, every setting, the recent list as
// the launcher shows it, the update feed as the updater reads it, and whether
// the launcher's window came up and drew itself, because a copy can read its
// profile and still show nothing. Then it quits. The judgement is the test's
// (tools/upgrade-check.js), from the file.
//
// From 2.26.0 it also says whether a process the speech engine starts is held
// to its folder in this copy as built (engine-guard-probe.js): the start
// through the core's bootstrap, from inside the archive, with the guard on.
//
// The feed is read with the updater's own library and the same feed, with
// downloading off: reading it is the point, fetching an installer is not.
// RUTBA_UPGRADE_FEED=0 leaves it out, for a run with no network.

import fs from 'node:fs';
import { app, utilityProcess } from 'electron';
import { FEED } from './updates.js';
import { probeEngineChild } from './engine-guard-probe.js';

async function readFeed() {
  if (!app.isPackaged) return { ok: false, error: 'An unpackaged copy has no updater to read the feed with.' };
  try {
    const mod = await import('electron-updater');
    const updater = mod.autoUpdater ?? mod.default?.autoUpdater;
    updater.autoDownload = false;
    updater.autoInstallOnAppQuit = false;
    updater.allowPrerelease = false;
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

const until = async (ask, ms) => {
  const end = Date.now() + ms;
  while (Date.now() < end) {
    try { if (await ask()) return true; } catch { /* not there yet */ }
    await new Promise((r) => setTimeout(r, 250));
  }
  return false;
};

/** The launcher, opened as a person's start opens it: did it show, and did it draw? */
async function firstWindow(windows) {
  try {
    const win = windows.open({ app: 'home' });
    const visible = await until(() => !win.isDestroyed() && win.isVisible(), 30000);
    const drawn = await until(() => win.webContents.executeJavaScript('document.body.innerText.trim().length > 20'), 30000);
    return { visible, drawn };
  } catch (err) {
    return { visible: false, drawn: false, error: String(err?.message || err).slice(0, 400) };
  }
}

export async function upgradeReport({ stores, windows, out, feed = true }) {
  const report = {
    version: app.getVersion(),
    packaged: app.isPackaged,
    userData: stores.dir,
    settings: stores.settings.all(),
    recent: await stores.recent.list(),
    feed: feed ? await readFeed() : null,
    window: await firstWindow(windows),
    engineChild: await probeEngineChild({ utilityProcess }),
  };
  fs.writeFileSync(out, JSON.stringify(report, null, 2));
  return true;
}
