// Network cameras as picture sources in the Recorder and Live Broadcast
// (owner, 22 Nov: "in the sources we need screens, camera and camera on
// (onvif, rtsp etc)").
//
// The cameras a person has added are kept here: each one's name, its address
// and its user name in the settings; its password in the secrets, encrypted
// by the operating system's keystore, and never handed back to a window. An
// address typed with the sign-in inside it is taken apart first, so no
// password is ever written with the address.
//
// A camera is read (@rutba/netcam) only while a studio window holds it open,
// which it does only for a source that is started in the scene it has open.
// When the last window lets go, or closes, the reading stops. Two windows that
// show the same camera share one reading: most cameras allow few. Nothing is
// recorded here; the frames go to the windows that asked and nowhere else.
//
// A camera that drops is tried again after a while, for as long as it is
// wanted. One that says no (a sign-in it does not accept, a stream it does
// not have) is not asked again until its details are saved again: asking a
// camera over and over with a password it refused can lock its account.
//
// Office connects out to the camera, on the one connection. Nothing here
// listens.
//
// Cameras on the network: Office asks the local network which cameras speak
// ONVIF (one question, and the answers to it read for a few seconds), and
// asks the one a person chooses for its stream address, with the sign-in the
// person gives. A window is handed the list and names a camera by a token of
// it; it never names an address for Office to ask.

import crypto from 'node:crypto';
import { RtspReader, parseCameraAddress, cameraNameFrom, discover, readCameraService } from '@rutba/netcam';

const LIST_KEY = 'cameras.list';
const passName = (id) => `cameras.pass.${id}`;
const ID = /^[A-Za-z0-9_-]{1,60}$/;
const MAX_ADDRESS = 400;

/** How long to wait before trying a camera again that dropped, in seconds. */
const RETRY_S = [2, 5, 10, 20, 30];

/** What a camera says that asking again will not change. */
const FINAL = new Set(['address', 'missing', 'sign-in-needed', 'sign-in-refused', 'sign-in-unsupported', 'not-found', 'refused', 'redirect', 'no-video', 'h265', 'codec', 'no-tcp', 'certificate']);

/**
 * How the network is asked which cameras are on it.
 *
 * A run of the window checks asks nothing of the network the computer is on: the cameras a person
 * has there are not the checks' to list, and what answers would differ from one run to the next.
 * It asks what the check stood in (`globalThis.__rutbaCameraFind`, a function that answers as
 * `discover` does), and finds no camera when a check stood nothing in.
 * @param {{ checkRun?: boolean, ask?: () => Promise<object[]> }} [o] `ask` is for the tests
 * @returns {() => Promise<object[]> | object[]}
 */
export function cameraFinder({ checkRun = false, ask = () => discover() } = {}) {
  if (!checkRun) return () => ask();
  return () => (typeof globalThis.__rutbaCameraFind === 'function' ? globalThis.__rutbaCameraFind() : []);
}

/**
 * @param {{ stores, send: (win, event, payload) => void, checkRun?: boolean, makeReader?: (o) => RtspReader, retrySeconds?: number[], find?: () => Promise<object[]>, readService?: (o) => Promise<object> }} o
 * `checkRun` is true in a run of the window checks (`cameraFinder`). `makeReader`, `retrySeconds`,
 * `find` and `readService` are for the tests.
 */
export function createNetCameras({ stores, send, checkRun = false, makeReader = (o) => new RtspReader(o), retrySeconds = RETRY_S, find = cameraFinder({ checkRun }), readService = readCameraService }) {
  const settings = stores.settings;
  const secrets = stores.secrets;
  const read = () => {
    const list = settings.get(LIST_KEY, []);
    return Array.isArray(list) ? list.filter((c) => c && ID.test(String(c.id || '')) && typeof c.url === 'string') : [];
  };
  const write = (list) => settings.set(LIST_KEY, list);
  const shown = (c) => ({ id: c.id, name: c.name || '', url: c.url, user: c.user || '', hasPassword: Boolean(secrets.get(passName(c.id))) });

  /** camera id -> the one reading of it, and the windows it is for */
  const sessions = new Map();
  /** The windows whose cameras are let go when they close or reload. */
  const watched = new WeakSet();
  /** token -> a camera that answered the last time the network was asked */
  let foundNow = new Map();
  /** The asking under way, shared by whoever asks meanwhile. */
  let finding = null;
  const hostOf = (url) => parseCameraAddress(url)?.host || '';

  const stateOf = (s) => ({ id: s.id, status: s.status, code: s.code, width: s.width, height: s.height, sound: s.sound });
  const alive = (win) => win && !win.isDestroyed?.();
  const tell = (s) => {
    for (const win of s.owners) if (alive(win)) send(win, 'camera:state', stateOf(s));
  };

  function halt(s) {
    s.run += 1;
    clearTimeout(s.retry);
    s.retry = null;
    const reader = s.reader;
    s.reader = null;
    try {
      reader?.stop();
    } catch {
      // Already gone.
    }
  }

  function ended(s, code) {
    s.reader = null;
    s.code = code;
    if (FINAL.has(code)) {
      s.status = 'failed';
      tell(s);
      return;
    }
    s.status = 'waiting';
    tell(s);
    const wait = retrySeconds[Math.min(s.tries, retrySeconds.length - 1)];
    s.tries += 1;
    clearTimeout(s.retry);
    s.retry = setTimeout(() => {
      if (sessions.get(s.id) === s && s.owners.size) begin(s);
    }, wait * 1000);
    s.retry.unref?.();
  }

  function begin(s) {
    halt(s);
    const run = s.run;
    const camera = read().find((c) => c.id === s.id);
    if (!camera) {
      s.status = 'failed';
      s.code = 'missing';
      tell(s);
      return;
    }
    s.status = 'connecting';
    s.code = '';
    tell(s);
    const current = () => run === s.run && sessions.get(s.id) === s;
    let reader;
    try {
      reader = makeReader({ url: camera.url, user: camera.user || '', password: secrets.get(passName(camera.id)) || '' });
    } catch {
      ended(s, 'address');
      return;
    }
    s.reader = reader;
    reader.on('config', (c) => {
      if (!current()) return;
      s.codec = c.codec;
      s.width = c.width;
      s.height = c.height;
    });
    reader.on('unit', (u) => {
      if (!current()) return;
      if (s.status !== 'live') {
        if (!u.key) return;
        s.status = 'live';
        s.code = '';
        s.tries = 0;
        tell(s);
      }
      // A copy of its own: the bytes a window is sent share nothing with what else this process holds.
      const frame = { id: s.id, key: u.key, timestamp: u.timestamp, bytes: new Uint8Array(u.data), ...(u.key ? { codec: s.codec, width: s.width, height: s.height } : {}) };
      for (const win of s.owners) if (alive(win)) send(win, 'camera:frame', frame);
    });
    reader.on('end', (err) => {
      if (current()) ended(s, err?.code || 'closed');
    });
    reader.start().then(
      (info) => {
        if (current()) s.sound = Boolean(info?.sound);
      },
      (err) => {
        if (current()) ended(s, err?.code || 'closed');
      },
    );
  }

  function letGo(id, win) {
    const s = sessions.get(id);
    if (!s || !s.owners.has(win)) return false;
    s.owners.delete(win);
    if (!s.owners.size) {
      halt(s);
      sessions.delete(id);
    }
    return true;
  }

  function closeAll(win) {
    for (const id of [...sessions.keys()]) {
      if (win) letGo(id, win);
      else {
        halt(sessions.get(id));
        sessions.delete(id);
      }
    }
  }

  const api = {
    /** () -> [{ id, name, url, user, hasPassword }]: the cameras added on this computer. Never a password. */
    cameraList: () => read().map(shown),

    /**
     * ({ id?, name?, url?, user?, password? }) -> { ok: true, camera, cameras }
     * or { ok: false, reason: 'address' | 'no-keystore' }. A sign-in typed
     * inside the address is lifted out of it. A password given replaces the
     * one kept; none leaves it as it is; a camera left with no user name
     * keeps no password. An `id` that is not known yet makes the camera
     * under that id: a setup brought from another computer names its cameras
     * by id, and the address is given here.
     */
    cameraSave: ({ id, name, url, user, password } = {}) => {
      const list = read();
      const known = id != null && ID.test(String(id));
      const at = known ? list.findIndex((c) => c.id === String(id)) : -1;
      const was = at >= 0 ? list[at] : null;
      const address = parseCameraAddress(url ?? was?.url);
      if (!address || address.url.length > MAX_ADDRESS) return { ok: false, reason: 'address' };
      const who = (user === undefined ? address.user || was?.user || '' : String(user).trim() || address.user).slice(0, 200);
      const pass = String(password || '') || address.password;
      if (pass && !secrets.available()) return { ok: false, reason: 'no-keystore' };
      const entry = {
        id: known ? String(id) : crypto.randomUUID(),
        name: String(name || was?.name || cameraNameFrom(address.url)).trim().slice(0, 80),
        url: address.url,
        user: who,
      };
      if (pass) secrets.set(passName(entry.id), pass);
      else if (!who) secrets.delete(passName(entry.id));
      if (at >= 0) list[at] = entry;
      else list.push(entry);
      write(list);
      // A camera being read is read again as it is now.
      const s = sessions.get(entry.id);
      if (s) {
        s.tries = 0;
        begin(s);
      }
      return { ok: true, camera: shown(entry), cameras: read().map(shown) };
    },

    /** ({ id }) -> cameras: the camera forgotten, its password with it, and its reading stopped. */
    cameraRemove: ({ id } = {}) => {
      write(read().filter((c) => c.id !== id));
      secrets.delete(passName(id));
      const s = sessions.get(id);
      if (s) {
        halt(s);
        s.status = 'failed';
        s.code = 'missing';
        tell(s);
      }
      return read().map(shown);
    },

    /**
     * ({ id }, win) -> state: this window shows the camera, so it is read.
     * Its frames come as camera:frame and what becomes of it as camera:state.
     * state: { id, status: 'connecting' | 'live' | 'waiting' | 'failed', code, width, height, sound }.
     */
    cameraOpen: ({ id } = {}, win) => {
      if (!win || !ID.test(String(id || ''))) throw new Error('A camera is opened by a window, by its id.');
      if (!watched.has(win)) {
        watched.add(win);
        win.once?.('closed', () => closeAll(win));
        win.webContents?.on?.('did-navigate', () => closeAll(win));
      }
      let s = sessions.get(id);
      if (s) {
        s.owners.add(win);
        // Asked for again after it said no: once more, since something may have been put right.
        if (s.status === 'failed') {
          s.tries = 0;
          begin(s);
        }
      } else {
        s = { id, owners: new Set([win]), reader: null, run: 0, retry: null, tries: 0, status: 'connecting', code: '', codec: '', width: 0, height: 0, sound: false };
        sessions.set(id, s);
        begin(s);
      }
      return stateOf(s);
    },

    /** ({ id }, win) -> { closed }: this window no longer shows it; with no window left, the reading stops. */
    cameraClose: ({ id } = {}, win) => ({ closed: letGo(id, win) }),

    /** Lets go of every camera a window held, or of all of them. */
    closeCameras: closeAll,
    /** [{ id, status, code, owners }] for the tests and checks. */
    camerasReading: () => [...sessions.values()].map((s) => ({ ...stateOf(s), owners: s.owners.size })),

    /**
     * () -> { cameras: [{ key, name, hardware, host, known, knownName }] }:
     * the cameras that answer on the local network, by name. `key` names one
     * to cameraAdopt until the network is asked again; `known` is the id of a
     * camera already added at that address, or null.
     */
    cameraFind: () => {
      if (!finding) {
        finding = Promise.resolve().then(find).then((list) => {
          const next = new Map();
          for (const c of Array.isArray(list) ? list : []) {
            if (c && typeof c.service === 'string' && typeof c.host === 'string') next.set(crypto.randomUUID(), c);
          }
          foundNow = next;
        }, () => {
          foundNow = new Map();
        }).then(() => {
          finding = null;
        });
      }
      return finding.then(() => {
        const saved = read();
        const cameras = [...foundNow].map(([key, c]) => {
          const was = saved.find((s) => hostOf(s.url) === c.host) || null;
          return { key, name: String(c.name || '').slice(0, 80), hardware: String(c.hardware || '').slice(0, 80), host: c.host, known: was ? was.id : null, knownName: was ? was.name || '' : '' };
        });
        cameras.sort((a, b) => (a.name || a.hardware || a.host).localeCompare(b.name || b.hardware || b.host) || a.host.localeCompare(b.host));
        return { cameras };
      });
    },

    /**
     * ({ key, user?, password?, name? }) -> { ok: true, camera, cameras } or
     * { ok: false, reason }: the found camera is asked for its stream address
     * with that sign-in, and added as a camera by address is. reason: 'gone'
     * (not in the list any more), 'no-keystore', or what the camera said:
     * 'sign-in-needed', 'sign-in-refused', 'unreachable', 'timeout',
     * 'no-video', 'h265', 'protocol', 'certificate'.
     */
    cameraAdopt: async ({ key, user, password, name } = {}) => {
      const found = foundNow.get(String(key || ''));
      if (!found) return { ok: false, reason: 'gone' };
      const who = String(user || '').trim().slice(0, 200);
      const pass = String(password || '');
      if (pass && !secrets.available()) return { ok: false, reason: 'no-keystore' };
      let stream;
      try {
        stream = await readService({ service: found.service, user: who, password: pass });
      } catch (e) {
        return { ok: false, reason: typeof e?.code === 'string' ? e.code : 'unreachable' };
      }
      // The same camera found again is the same camera, not a second one.
      const was = read().find((c) => c.url === stream.url) || null;
      return api.cameraSave({ ...(was ? { id: was.id } : {}), url: stream.url, user: who, ...(pass ? { password: pass } : {}), name: String(name || '').trim() || found.name || found.hardware || was?.name || '' });
    },
  };
  return api;
}
