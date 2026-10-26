// The AI Tools window's side in the main process: the person's switch (off,
// for this session, always), the folders the tools may use, the log of what
// they did, and the command an AI app starts them with (D-2026-10-16-06, -07);
// and who may connect: AI apps on this computer only, or also from other
// machines over the network, with a key (owner, 2026-10-21, D-2026-10-21-06:
// "the user should have choice to open remote connections or local access
// only"), served by this running Rutba Office while the tools are on.
//
// "For this session" is this running Rutba Office: the switch names its
// process, and the tools answer only while it is running. When it quits the
// switch goes back to off, and a session left over by a crash is found off
// at the next start.

import fs from 'node:fs';
import path from 'node:path';
import { app } from 'electron';
import { settledProfile } from '@rutba/office-shell/electron/profile-choice';
import { MODE_KEY, SESSION_KEY, FOLDERS_KEY, MODES, foldersOf, readLog, readConnections, openNow } from './mcp/access.js';
import { mcpCommand as commandFor } from './mcp/command.js';
import { createRemote, REACH_KEY, REACHES, REMOTE_KEY, REMOTE_PORT_KEY, DEFAULT_PORT, newKey, addresses } from './mcp/remote.js';
import { LOADED_KEY } from './pack-tools.js';

/**
 * How an AI app starts the tools (mcp/command.js): Rutba Office's own program
 * run as Node on the file beside the archive in an installed copy, or on the
 * source in a copy run from it, in this window's profile.
 */
export function mcpCommand({ execPath = process.execPath, appDir = app.getAppPath() } = {}) {
  return commandFor({ execPath, appDir, packaged: app.isPackaged, resourcesPath: process.resourcesPath, profile: settledProfile(), screens: Boolean(process.env.RUTBA_SCREENS_FEATURES) });
}

export function createAiToolsService({ stores, dataDir = app.getPath('userData'), pid = process.pid }) {
  const settings = stores.settings;
  const modeNow = () => {
    const mode = settings.get(MODE_KEY, 'off');
    if (mode === 'session') return Number(settings.get(SESSION_KEY, 0)) === pid ? 'session' : 'off';
    return MODES.includes(mode) ? mode : 'off';
  };
  const folders = () => foldersOf({ [FOLDERS_KEY]: settings.get(FOLDERS_KEY, []) });

  // A session switched on by an earlier run that did not quit cleanly is over.
  if (settings.get(MODE_KEY, 'off') === 'session' && Number(settings.get(SESSION_KEY, 0)) !== pid) settings.set(MODE_KEY, 'off');
  app.on('will-quit', () => {
    if (settings.get(MODE_KEY, 'off') === 'session' && Number(settings.get(SESSION_KEY, 0)) === pid) settings.set(MODE_KEY, 'off');
  });

  // Who may connect: this computer only, the default, or other machines too.
  const reach = () => (REACHES.includes(settings.get(REACH_KEY, 'local')) ? settings.get(REACH_KEY, 'local') : 'local');
  const port = () => {
    const p = Number(settings.get(REMOTE_PORT_KEY, DEFAULT_PORT));
    return Number.isInteger(p) && p >= 0 && p < 65536 ? p : DEFAULT_PORT;
  };
  const key = () => {
    let k = settings.get(REMOTE_KEY, '');
    if (!k) { k = newKey(); settings.set(REMOTE_KEY, k); }
    return k;
  };
  // The settings the tools obey, as the window keeps them, read for every call.
  const snapshot = () => ({
    'apps.choice': settings.get('apps.choice', null),
    [LOADED_KEY]: settings.get(LOADED_KEY, null),
    [MODE_KEY]: settings.get(MODE_KEY, 'off'),
    [SESSION_KEY]: settings.get(SESSION_KEY, 0),
    [FOLDERS_KEY]: settings.get(FOLDERS_KEY, []),
  });
  const remoteSessions = new Map();
  let remote = null;
  let remoteError = null;
  const allowed = () => (reach() === 'remote'
    ? openNow(snapshot(), { isAlive: (p) => p === pid })
    : { open: false, reason: 'Rutba Office is not taking connections from other machines. To let an AI app on another machine use the tools, choose it in AI Tools.' });

  /** The listener as the switch and the reach say: listening while both allow it, closed otherwise. */
  const sync = async () => {
    const wanted = reach() === 'remote' && modeNow() !== 'off';
    try {
      if (!wanted) {
        await remote?.stop();
        remoteError = null;
        return;
      }
      if (!remote) {
        // The tools' modules are loaded the first time the network is chosen, not at every start.
        const { makeTools, versionOf, INSTRUCTIONS } = await import('./mcp/make-tools.js');
        remote = createRemote({
          tools: makeTools({ settings: snapshot, dataDir, isAlive: (p) => p === pid }),
          version: versionOf(),
          instructions: INSTRUCTIONS,
          allowed,
          key,
          onSession: (id, entry) => {
            if (entry) remoteSessions.set(id, { ...entry, pid: `net-${id}` });
            else remoteSessions.delete(id);
          },
        });
      }
      const at = remote.listening()?.port;
      if (at === undefined || (port() !== 0 && at !== port())) await remote.start(port());
      remoteError = null;
    } catch (err) {
      remoteError = err?.message || String(err);
    }
  };
  app.whenReady().then(() => sync());
  app.on('will-quit', () => { remote?.stop(); });

  const state = () => ({
    mode: modeNow(),
    folders: folders(),
    connect: mcpCommand(),
    log: readLog(dataDir, 60),
    connections: [...readConnections(dataDir), ...remoteSessions.values()],
    reach: reach(),
    remote: reach() === 'remote'
      ? { port: remote?.listening()?.port ?? port(), key: key(), addresses: addresses(), listening: Boolean(remote?.listening()), error: remoteError }
      : null,
  });

  return {
    state,

    /**
     * ({ mode }) -> state. On needs a folder the person chose first (owner,
     * 2026-10-17: "the user should have choice to setup different one"): the
     * window asks for one before it turns them on, and none is made for them.
     */
    setMode: ({ mode }) => {
      if (!MODES.includes(mode)) throw new Error(`No such setting: ${mode}`);
      if (mode !== 'off' && !folders().length) throw new Error('Choose a folder for the AI tools first.');
      if (mode === 'session') settings.set(SESSION_KEY, pid);
      settings.set(MODE_KEY, mode);
      return sync().then(state);
    },

    /**
     * ({ reach, port? }) -> state. 'local': AI apps on this computer only,
     * nothing listening on the network. 'remote': other machines too, at
     * the port, with the key the window shows, while the tools are on.
     */
    setReach: async ({ reach: to, port: at } = {}) => {
      if (!REACHES.includes(to)) throw new Error(`No such setting: ${to}`);
      if (at !== undefined) {
        const p = Number(at);
        if (!Number.isInteger(p) || p < 0 || p > 65535) throw new Error('A port is a whole number from 1 to 65535.');
        settings.set(REMOTE_PORT_KEY, p);
      }
      settings.set(REACH_KEY, to);
      if (to === 'remote') key();
      await sync();
      return state();
    },

    /** () -> state: a new key; the old one, and every connection made with it, stop working at once. */
    renewKey: async () => {
      settings.set(REMOTE_KEY, newKey());
      await remote?.stop();
      await sync();
      return state();
    },

    /** ({ path }) -> state: a folder the tools may use from now on. */
    addFolder: ({ path: folder }) => {
      const full = path.resolve(String(folder || ''));
      if (!folder || !fs.existsSync(full) || !fs.statSync(full).isDirectory()) throw new Error('That folder is not there.');
      settings.set(FOLDERS_KEY, [...new Set([...folders(), full])]);
      return state();
    },

    /** ({ path }) -> state: a folder the tools may no longer use. */
    removeFolder: ({ path: folder }) => {
      const full = path.resolve(String(folder || ''));
      settings.set(FOLDERS_KEY, folders().filter((f) => f.toLowerCase() !== full.toLowerCase()));
      return state();
    },
  };
}
