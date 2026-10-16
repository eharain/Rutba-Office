// The AI Tools window's side in the main process: the person's switch (off,
// for this session, always), the folders the tools may use, the log of what
// they did, and the command an AI app starts them with (D-2026-10-16-06, -07).
//
// "For this session" is this running Rutba Office: the switch names its
// process, and the tools answer only while it is running. When it quits the
// switch goes back to off, and a session left over by a crash is found off
// at the next start.

import fs from 'node:fs';
import path from 'node:path';
import { app } from 'electron';
import { MODE_KEY, SESSION_KEY, FOLDERS_KEY, MODES, foldersOf, readLog, readConnections } from './mcp/access.js';

/**
 * How an AI app starts the tools: Rutba Office's own program run as Node
 * (mcp/node.js), from the app's archive in an installed copy, or from the
 * source in a copy run from it.
 */
export function mcpCommand({ execPath = process.execPath, appDir = app.getAppPath() } = {}) {
  return { command: execPath, args: [path.join(appDir, 'main', 'mcp', 'node.js')], env: { ELECTRON_RUN_AS_NODE: '1' } };
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

  const state = () => ({ mode: modeNow(), folders: folders(), connect: mcpCommand(), log: readLog(dataDir, 60), connections: readConnections(dataDir) });

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
