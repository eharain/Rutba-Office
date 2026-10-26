// The windows' side of profiles (profiles.js, D-2026-10-26-02): which profile
// this window belongs to, the list, and making, renaming, colouring,
// choosing the one that opens first, removing, and a password set, changed
// or taken away.
//
// Each profile runs as its own Rutba Office, so opening another profile
// starts the program again with `--profile <id>`: a new one if that profile
// is not running, or, if it is, the running one hears of it through its
// one-instance lock and comes to the front (main.js, 'second-instance').
// A password belongs to the profile it protects, so it is set, changed and
// taken away from that profile's own windows, which hold its key.

import fs from 'node:fs';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { app, shell as electronShell } from 'electron';
import {
  readProfiles, profileFolder, makeProfile, renameProfile, setProfileColour, setOpenFirst, rekeyProfile,
  removeProfile, removalRefused, isRunning, folderSize, makeLock, unlock, markUsed, PROFILES_FILE, DEFAULT_ID, COLOURS,
} from '../profiles.js';

/**
 * @param {object} o
 * @param {object} o.profile  the profile this run is (profile-choice.js), kept current here
 * @param {object} o.stores   this profile's stores, whose vault seals its files
 * @param {(event: string, payload: any) => void} o.broadcast
 */
export function createProfilesService({ profile, stores, broadcast }) {
  const base = profile.base;

  /** This run's own entry, read again after every change (another window, or another profile's, may have made it). */
  const refresh = () => {
    const list = readProfiles(base);
    const me = list.profiles.find((p) => p.id === profile.id);
    if (me) Object.assign(profile, { name: me.name, colour: me.colour, lock: me.lock || null });
    profile.count = list.profiles.length;
    return list;
  };

  const current = () => {
    refresh();
    return { id: profile.id, name: profile.name, colour: profile.colour, count: profile.count, protected: Boolean(profile.lock), isDefault: profile.id === DEFAULT_ID };
  };

  const list = () => {
    const l = refresh();
    return {
      current: profile.id,
      openFirst: l.openFirst,
      colours: COLOURS,
      profiles: l.profiles.map((p) => ({
        id: p.id,
        name: p.name,
        colour: p.colour,
        protected: Boolean(p.lock),
        openFirst: p.id === l.openFirst,
        current: p.id === profile.id,
        running: p.id === profile.id || isRunning(profileFolder(base, p)),
        isDefault: p.id === DEFAULT_ID,
      })),
    };
  };

  // Every window hears of a change, this profile's and the other running
  // profiles' (the list is one file, watched).
  let last = '';
  const tell = () => {
    const now = current();
    const said = JSON.stringify(now);
    if (said === last) return now;
    last = said;
    broadcast?.('profiles:changed', now);
    return now;
  };
  try {
    fs.watchFile(path.join(base, PROFILES_FILE), { interval: 1500, persistent: false }, () => tell());
  } catch {
    // Without a watch, a change made in another profile shows in this one's windows when they next ask.
  }
  last = JSON.stringify(current());
  const changed = () => {
    tell();
    return list();
  };

  // The profile used last is the one a file opened from Explorer goes to.
  markUsed(base, profile.id);
  app.on('browser-window-focus', () => {
    try {
      markUsed(base, profile.id);
    } catch {
      // The list is a nicety here: the file still opens in the profile that opens first.
    }
  });

  /** This program again, for another profile (and a file to open in it). */
  const launch = (id, file = null) => {
    const args = [];
    if (!app.isPackaged) args.push(app.getAppPath());
    if (profile.baseNamed) args.push(`--user-data-dir=${base}`);
    args.push(`--profile=${id}`);
    if (file) args.push(file);
    // A check run starts its profiles as check runs too (verify-profiles.js).
    const env = globalThis.__rutbaCheckProfileEnv?.(id) || { ...process.env };
    delete env.ELECTRON_RUN_AS_NODE;
    const child = spawn(process.execPath, args, { detached: true, stdio: 'ignore', env, windowsHide: false });
    child.on('error', () => {});
    child.unref();
    return child.pid;
  };

  // A profile removed goes to the Recycle Bin, as a file removed in Office does; a check run's temporary one is deleted.
  const discard = async (folder) => {
    if (process.env.RUTBA_OFFICE_VERIFY_APPS || process.env.RUTBA_PROFILE_PEER) return fs.rmSync(folder, { recursive: true, force: true, maxRetries: 20, retryDelay: 250 });
    try {
      await electronShell.trashItem(folder);
    } catch {
      fs.rmSync(folder, { recursive: true, force: true, maxRetries: 20, retryDelay: 250 });
    }
  };

  const mustBeHere = () => {
    if (!stores?.vault) throw new Error('This profile’s files cannot be sealed here.');
  };

  return {
    current,
    list,

    /** ({ id, file }) -> { started | here }: that profile's Rutba Office, started or brought to the front. */
    open: ({ id, file = null } = {}) => {
      const l = readProfiles(base);
      const p = l.profiles.find((x) => x.id === id);
      if (!p) throw new Error('That profile is no longer there.');
      if (p.id === profile.id && !file) return { here: true };
      return { started: Boolean(launch(p.id, file)), running: isRunning(profileFolder(base, p)) };
    },

    /** ({ name, colour }) -> list, with the new profile's id as `made`. */
    make: ({ name, colour } = {}) => {
      const { profile: made } = makeProfile(base, { name, colour });
      return { ...changed(), made: made.id };
    },

    rename: ({ id, name } = {}) => {
      renameProfile(base, id, name);
      return changed();
    },

    setColour: ({ id, colour } = {}) => {
      setProfileColour(base, id, colour);
      return changed();
    },

    setOpenFirst: ({ id } = {}) => {
      setOpenFirst(base, id);
      return changed();
    },

    /** ({ id }) -> { refused, size }: what removing it would take, asked before it is removed. */
    removal: ({ id } = {}) => {
      const refused = removalRefused(base, id, { current: profile.id });
      const p = readProfiles(base).profiles.find((x) => x.id === id);
      return { refused, size: p && !refused ? folderSize(profileFolder(base, p)) : 0 };
    },

    remove: async ({ id } = {}) => {
      await removeProfile(base, id, { current: profile.id, discard });
      return changed();
    },

    /**
     * ({ password }) -> current. This profile asks for it before it opens from
     * now on, and its saved passwords, recent files, calendars, contacts and
     * Talk messages are sealed with a key made from it. Each of the three is
     * written down first and the new lock put in last (profiles.js,
     * rekeyProfile): a stop midway is finished at the next open, with either
     * password, and no file is left sealed with a key the lock does not give.
     */
    setPassword: async ({ password } = {}) => {
      mustBeHere();
      if (profile.lock) throw new Error('This profile already has a password. Change it instead.');
      const { lock, key } = await makeLock(password);
      profile.lock = await rekeyProfile({ base, id: profile.id, vault: stores.vault, from: null, to: key, lock, folder: stores.dir });
      return tell();
    },

    /** ({ old, password }) -> current: the files sealed again with the new password's key. */
    changePassword: async ({ old, password } = {}) => {
      mustBeHere();
      if (!profile.lock) throw new Error('This profile has no password yet.');
      const from = await unlock(profile.lock, old);
      if (!from) throw new Error('The current password is not right.');
      const { lock, key } = await makeLock(password);
      profile.lock = await rekeyProfile({ base, id: profile.id, vault: stores.vault, from, to: key, lock, folder: stores.dir });
      return tell();
    },

    /** ({ old }) -> current: the files plain again (the keychain still protects saved passwords), then the lock gone. */
    removePassword: async ({ old } = {}) => {
      mustBeHere();
      if (!profile.lock) return current();
      if (!(await unlock(profile.lock, old))) throw new Error('The current password is not right.');
      profile.lock = await rekeyProfile({ base, id: profile.id, vault: stores.vault, from: null, to: null, lock: null, folder: stores.dir });
      return tell();
    },
  };
}
