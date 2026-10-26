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
//
// Beyond files (D-2026-10-26-04, -05, -06): each mail account, archive,
// calendar, address book, Data Files database and folder carries its own
// setting (mcp/reach.js), read and written here for its own app's settings
// and for the AI Tools window alike, so the two always show the one
// setting. The tools for those things run here, in the copy that holds them
// (mcp/reach-tools.js), reached by the tools' own process over this
// profile's channel (mcp/channel.js) while the tools are on; and what they
// ask to send or write elsewhere waits here, in every window, for the
// person's press (mcp/confirm.js).

import fs from 'node:fs';
import path from 'node:path';
import { app, BrowserWindow } from 'electron';
import { settledProfile } from '@rutba/office-shell/electron/profile-choice';
import { heldBlob, releaseBlob } from '@rutba/office-shell/electron/main';
import { MODE_KEY, SESSION_KEY, FOLDERS_KEY, MODES, foldersOf, readLog, readConnections, openNow } from './mcp/access.js';
import { mcpCommand as commandFor } from './mcp/command.js';
import { createRemote, REACH_KEY, REACHES, REMOTE_KEY, REMOTE_PORT_KEY, DEFAULT_PORT, newKey, addresses } from './mcp/remote.js';
import { KINDS, STEPS, DATABASES_KEY, LOCAL_BOOK, stepsFor, stepOf, folderEntries, withFolderStep, databaseSetting } from './mcp/reach.js';
import { createConfirmations } from './mcp/confirm.js';
import { createChannelServer } from './mcp/channel.js';

/**
 * How an AI app starts the tools (mcp/command.js): Rutba Office's own program
 * run as Node on the file beside the archive in an installed copy, or on the
 * source in a copy run from it, in this window's profile.
 */
export function mcpCommand({ execPath = process.execPath, appDir = app.getAppPath() } = {}) {
  return commandFor({ execPath, appDir, packaged: app.isPackaged, resourcesPath: process.resourcesPath, profile: settledProfile(), screens: Boolean(process.env.RUTBA_SCREENS_FEATURES) });
}

const sameFile = (a, b) => (process.platform === 'win32' ? String(a).toLowerCase() === String(b).toLowerCase() : a === b);

/**
 * @param {object} o
 * @param {object} o.stores       this profile's stores
 * @param {() => object} [o.services]  the running copy's services (mail, calendar, contacts, base), asked when used
 * @param {(event: string, payload: any) => void} [o.broadcast]
 */
export function createAiToolsService({ stores, dataDir = app.getPath('userData'), pid = process.pid, services = () => ({}), broadcast = () => {} }) {
  const settings = stores.settings;
  const modeNow = () => {
    const mode = settings.get(MODE_KEY, 'off');
    if (mode === 'session') return Number(settings.get(SESSION_KEY, 0)) === pid ? 'session' : 'off';
    return MODES.includes(mode) ? mode : 'off';
  };
  const folders = () => foldersOf({ [FOLDERS_KEY]: settings.get(FOLDERS_KEY, []) });
  const addon = () => {
    const choice = settings.get('apps.choice', null);
    return !choice || (Array.isArray(choice.addons) && choice.addons.includes('ai'));
  };

  // A session switched on by an earlier run that did not quit cleanly is over.
  if (settings.get(MODE_KEY, 'off') === 'session' && Number(settings.get(SESSION_KEY, 0)) !== pid) settings.set(MODE_KEY, 'off');
  app.on('will-quit', () => {
    if (settings.get(MODE_KEY, 'off') === 'session' && Number(settings.get(SESSION_KEY, 0)) === pid) settings.set(MODE_KEY, 'off');
  });

  /* ── what waits for the person's press ──────────────────────────────── */

  const confirm = createConfirmations({
    tell: (pending) => broadcast('ai:pending', pending),
    // A window to the front for a new request; a check run's windows stay where they are.
    attend: () => {
      if (process.env.RUTBA_WINDOW_DISPLAY === 'offscreen') return;
      const win = BrowserWindow.getFocusedWindow() || BrowserWindow.getAllWindows().find((w) => !w.isDestroyed() && w.isVisible());
      if (!win) return;
      if (win.isMinimized()) win.restore();
      if (!win.isFocused()) win.flashFrame(true);
      win.show();
    },
  });
  app.on('will-quit', () => confirm.clear());

  /* ── the tools for mail, calendars, address books and databases ────── */

  let reachTools = null;
  /** Built the first time the tools are asked for them: the readers they need load then, not at every start. */
  const reach = async () => {
    if (reachTools) return reachTools;
    const [{ createReachTools }, { READERS }, { resolveIn }] = await Promise.all([import('./mcp/reach-tools.js'), import('./mcp/tools.js'), import('./mcp/access.js')]);
    // Built when first asked, by which time every service is.
    const s = () => services() || {};
    reachTools = createReachTools({
      settings,
      mail: s().mail,
      calendar: s().calendar,
      contacts: s().contacts,
      base: s().base,
      confirm,
      files: { readable: (given) => resolveIn(folders(), given) },
      readers: READERS,
      attachmentBytes: async (args) => {
        const held = await s().mail?.attachment(args);
        if (!held?.id) return null;
        const blob = heldBlob(held.id);
        releaseBlob(held.id);
        return blob ? { bytes: blob.bytes, name: held.name, type: held.type } : null;
      },
    });
    return reachTools;
  };
  const reachCall = async (name, args) => (await reach()).call(name, args);

  // This profile's channel, for the tools an AI app on this computer starts: open while the tools are on.
  const channel = createChannelServer({
    folder: dataDir,
    handle: async (method, params) => {
      if (modeNow() === 'off') throw new Error('Rutba Office’s AI tools are off.');
      if (method === 'hello') return { profile: settledProfile()?.id || 'default', unlocked: true, pid };
      if (method === 'call') return reachCall(String(params.name || ''), params.args || {});
      throw new Error(`Rutba Office does not answer ${method}.`);
    },
  });

  // Who may connect: this computer only, the default, or other machines too.
  const reachOf = () => (REACHES.includes(settings.get(REACH_KEY, 'local')) ? settings.get(REACH_KEY, 'local') : 'local');
  const port = () => {
    const p = Number(settings.get(REMOTE_PORT_KEY, DEFAULT_PORT));
    return Number.isInteger(p) && p >= 0 && p < 65536 ? p : DEFAULT_PORT;
  };
  // The network key, kept with the profile's saved passwords (its keystore); where nothing here can keep
  // one, it lasts for this run only, and an AI app on another machine is given the new one at the next.
  let runKey = '';
  const keepKey = (k) => {
    if (stores.secrets?.available?.()) stores.secrets.set(REMOTE_KEY, k);
    else runKey = k;
  };
  const key = () => {
    let k = (stores.secrets?.available?.() ? stores.secrets.get(REMOTE_KEY) : null) || runKey;
    if (!k) { k = newKey(); keepKey(k); }
    return k;
  };
  // The settings the tools obey, as the window keeps them, read for every call.
  const snapshot = () => ({
    'apps.choice': settings.get('apps.choice', null),
    [MODE_KEY]: settings.get(MODE_KEY, 'off'),
    [SESSION_KEY]: settings.get(SESSION_KEY, 0),
    [FOLDERS_KEY]: settings.get(FOLDERS_KEY, []),
  });
  const remoteSessions = new Map();
  let remote = null;
  let remoteError = null;
  const allowed = () => (reachOf() === 'remote'
    ? openNow(snapshot(), { isAlive: (p) => p === pid })
    : { open: false, reason: 'Rutba Office is not taking connections from other machines. To let an AI app on another machine use the tools, choose it in AI Tools.' });

  /** The channel and the network listener as the switch and the reach say: open while they allow it, closed otherwise. */
  const sync = async () => {
    try {
      if (modeNow() !== 'off') await channel.start();
      else await channel.stop();
    } catch (err) {
      console.error('ai tools channel:', err?.message || err);
    }
    const wanted = reachOf() === 'remote' && modeNow() !== 'off';
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
          tools: makeTools({ settings: snapshot, dataDir, isAlive: (p) => p === pid, reach: { call: reachCall } }),
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
  app.on('will-quit', () => { remote?.stop(); channel.stop(); });

  /* ── each thing's own setting ───────────────────────────────────────── */

  /** Every database the person has opened lately or given a setting, for the window to list. */
  const databaseFiles = async () => {
    const set = Object.keys(settings.get(DATABASES_KEY, {}) || {});
    let recent = [];
    try {
      recent = ((await stores.recent.list()) || []).map((r) => r.path).filter((p) => /\.rbase$/i.test(String(p || '')));
    } catch {
      recent = [];
    }
    const out = [];
    for (const p of [...set, ...recent]) if (!out.some((q) => sameFile(q, p))) out.push(p);
    return out;
  };

  /** Everything the tools could reach, each with its step and the steps it may have. */
  const things = async () => {
    const s = services() || {};
    const list = [];
    for (const a of s.mail?.accounts?.() || []) list.push({ kind: 'mail', id: a.id, name: a.email, label: a.name || a.email, local: Boolean(a.local), step: stepOf('mail', a.ai, a), steps: stepsFor('mail', a) });
    for (const a of s.mail?.archives?.() || []) list.push({ kind: 'archive', id: a.id, name: a.name || path.basename(a.path || ''), step: stepOf('archive', a.ai), steps: stepsFor('archive') });
    for (const c of s.calendar?.calendars?.() || []) list.push({ kind: 'calendar', id: c.id, name: c.name, server: Boolean(c.account), step: stepOf('calendar', c.ai), steps: stepsFor('calendar') });
    for (const b of s.contacts?.books?.() || []) list.push({ kind: 'book', id: b.id, name: b.id === LOCAL_BOOK ? null : b.name, server: Boolean(b.server), step: stepOf('book', b.ai), steps: stepsFor('book') });
    const dbs = settings.get(DATABASES_KEY, {}) || {};
    for (const file of await databaseFiles()) {
      const d = databaseSetting(dbs, file);
      list.push({ kind: 'database', id: file, name: path.basename(file).replace(/\.[^.]+$/, ''), file, step: d.step, role: d.role, steps: stepsFor('database'), missing: !fs.existsSync(file) });
    }
    for (const f of folderEntries(settings.get(FOLDERS_KEY, []))) list.push({ kind: 'folder', id: f.path, name: f.path, step: f.step, steps: stepsFor('folder') });
    return list;
  };

  /** One thing's setting, as its own app's settings show it. */
  const stepOfThing = async ({ kind, id }) => {
    const hit = (await things()).find((x) => x.kind === kind && (x.id === id || (kind === 'database' || kind === 'folder' ? sameFile(x.id, id) : false)));
    if (hit) return { addon: addon(), step: hit.step, steps: hit.steps, role: hit.role || null };
    // A database not listed yet: Off until chosen.
    if (kind === 'database') return { addon: addon(), step: 'off', steps: stepsFor('database'), role: null };
    throw new Error('That is no longer there.');
  };

  /** One thing's setting written where it is kept (mcp/reach.js), and every window told. */
  const setStep = async ({ kind, id, step, role = undefined }) => {
    if (!KINDS.includes(kind)) throw new Error(`No such kind: ${kind}`);
    if (!STEPS.includes(step)) throw new Error(`No such setting: ${step}`);
    const s = services() || {};
    if (kind === 'mail') {
      const a = (s.mail?.accounts?.() || []).find((x) => x.id === id);
      if (!a) throw new Error('That account is no longer there.');
      if (!stepsFor('mail', a).includes(step)) throw new Error('A mailbox brought in from a file has no server to send through.');
      // Through Mail's own update, so the account's record keeps it.
      s.mail.updateAccount({ id, patch: { ai: step } });
    } else if (kind === 'archive') {
      if (!stepsFor('archive').includes(step)) throw new Error('An archive opened in place is read only.');
      const list = settings.get('mail.archives', []);
      const at = list.findIndex((x) => x.id === id);
      if (at < 0) throw new Error('That archive is no longer open.');
      list[at] = { ...list[at], ai: step };
      settings.set('mail.archives', list);
    } else if (kind === 'calendar') {
      if (!(s.calendar?.calendars?.() || []).some((c) => c.id === id)) throw new Error('That calendar is no longer there.');
      s.calendar.saveCalendar({ calendar: { id, ai: step } });
    } else if (kind === 'book') {
      s.contacts.setBookAi({ id, ai: step });
    } else if (kind === 'database') {
      const file = path.resolve(String(id || ''));
      const map = { ...(settings.get(DATABASES_KEY, {}) || {}) };
      const was = Object.keys(map).find((p) => sameFile(p, file));
      const before = was ? map[was] : {};
      if (was) delete map[was];
      map[file] = { step, role: role === undefined ? before?.role || null : role || null };
      settings.set(DATABASES_KEY, map);
    } else if (kind === 'folder') {
      const full = path.resolve(String(id || ''));
      if (!stepsFor('folder').includes(step)) throw new Error('The AI tools read a folder and make new files in it; nothing in a folder is sent anywhere.');
      settings.set(FOLDERS_KEY, withFolderStep(settings.get(FOLDERS_KEY, []), full, step, sameFile));
    }
    const now = await stepOfThing({ kind, id });
    broadcast('ai:changed', { kind, id, step: now.step });
    return { kind, id, ...now };
  };

  const state = async () => ({
    mode: modeNow(),
    folders: folders(),
    connect: mcpCommand(),
    log: readLog(dataDir, 60),
    connections: [...readConnections(dataDir), ...remoteSessions.values()],
    reach: reachOf(),
    remote: reachOf() === 'remote'
      ? { port: remote?.listening()?.port ?? port(), key: key(), addresses: addresses(), listening: Boolean(remote?.listening()), error: remoteError }
      : null,
    things: await things(),
  });

  return {
    state,

    /**
     * ({ mode }) -> state. On needs a folder the person chose first (owner,
     * 2026-10-17: "the user should have choice to setup different one"): the
     * window asks for one before it turns them on, and none is made for them.
     */
    setMode: async ({ mode }) => {
      if (!MODES.includes(mode)) throw new Error(`No such setting: ${mode}`);
      if (mode !== 'off' && !folderEntries(settings.get(FOLDERS_KEY, [])).length) throw new Error('Choose a folder for the AI tools first.');
      if (mode === 'session') settings.set(SESSION_KEY, pid);
      settings.set(MODE_KEY, mode);
      await sync();
      return state();
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
      keepKey(newKey());
      await remote?.stop();
      await sync();
      return state();
    },

    /** ({ path }) -> state: a folder the tools may use from now on, to read and make new files in. */
    addFolder: ({ path: folder }) => {
      const full = path.resolve(String(folder || ''));
      if (!folder || !fs.existsSync(full) || !fs.statSync(full).isDirectory()) throw new Error('That folder is not there.');
      const kept = settings.get(FOLDERS_KEY, []);
      if (!folderEntries(kept).some((f) => sameFile(f.path, full))) settings.set(FOLDERS_KEY, [...kept, full]);
      return state();
    },

    /** ({ path }) -> state: a folder the tools may no longer use. */
    removeFolder: ({ path: folder }) => {
      const full = path.resolve(String(folder || ''));
      settings.set(FOLDERS_KEY, (settings.get(FOLDERS_KEY, []) || []).filter((f) => !sameFile(typeof f === 'string' ? f : f?.path || '', full)));
      return state();
    },

    things: async () => ({ addon: addon(), things: await things() }),
    setStep,
    stepOf: stepOfThing,
    pending: () => confirm.pending(),
    /** ({ id, choice }) -> what was done: the person's press, the only way anything the tools asked for is sent. */
    answer: ({ id, choice }) => confirm.answer(String(id || ''), String(choice || '')),

  };
}
