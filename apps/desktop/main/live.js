// Going live from the Recorder (owner, 2026-10-19: "stream videos live on
// youtube and facebook etc").
//
// The streams a person has set up are kept here: each one's platform, a name,
// its server address, its shape and whether it is on, in the settings; its
// stream key in the secrets, encrypted by the operating system's keystore, and
// never handed back to a window (the window sees only its last four
// characters).
//
// A broadcast (@rutba/stream/broadcast) is open from Go live until End
// broadcast, and streams join and leave it while it runs. The window encodes
// the program, one picture per shape in use and one sound, H.264 and AAC, and
// sends each piece here; the same bytes go to every stream of that shape.
//
// Nothing connects until the person presses Go live, and only to the servers
// they set up. Nothing here listens.

import crypto from 'node:crypto';
import { RtmpPublisher, parseRtmpUrl, maskKey, PLATFORMS } from '@rutba/stream';
import { Broadcast } from '@rutba/stream/broadcast';
import { sendEvent } from '@rutba/office-shell/electron/main';

const LIST_KEY = 'live.destinations';
const keyName = (id) => `live.key.${id}`;
const shapeName = (s) => (s === 'tall' ? 'tall' : 'wide');

/**
 * @param {{ stores, makePublisher?: (o) => RtmpPublisher }} o `makePublisher`
 * is for the checks, which hand it a local ingest.
 */
export function createLiveService({ stores, makePublisher = (o) => new RtmpPublisher(o) }) {
  const settings = stores.settings;
  const secrets = stores.secrets;
  const read = () => (Array.isArray(settings.get(LIST_KEY, [])) ? settings.get(LIST_KEY, []) : []);
  const write = (list) => settings.set(LIST_KEY, list);

  const shown = (d) => {
    const key = secrets.get(keyName(d.id));
    return { id: d.id, platform: d.platform, name: d.name, url: d.url, on: Boolean(d.on), shape: shapeName(d.shape), hasKey: Boolean(key), keyHint: maskKey(key) };
  };

  // The open broadcast, and the window it answers.
  let broadcast = null;
  let window = null;

  const empty = () => ({ live: false, destinations: [], kbps: 0, target: 0, rates: {} });
  const statusNow = () => (broadcast ? broadcast.status() : empty());

  const join = (d, key) => {
    broadcast.add({ id: d.id, name: d.name, url: d.url, key, shape: shapeName(d.shape), maxKbps: PLATFORMS[d.platform]?.maxKbps || 0 });
  };

  return {
    /** () -> [{ id, platform, name, url, on, shape, hasKey, keyHint }] */
    destinations: () => read().map(shown),

    /** The platforms offered, with their usual server and where their key is found. */
    platforms: () => PLATFORMS,

    /**
     * ({ id?, platform, name?, url?, on?, shape?, key? }) -> destinations. A
     * key given replaces the one kept; no key leaves it as it is. Left out, a
     * field keeps what it was, and a new stream is off until it is turned on.
     */
    saveDestination: ({ id, platform, name, url, on, shape, key } = {}) => {
      const list = read();
      const at = id ? list.findIndex((d) => d.id === id) : -1;
      const was = at >= 0 ? list[at] : null;
      const asked = platform ?? was?.platform;
      const p = PLATFORMS[asked] ? asked : 'custom';
      const address = String(url ?? was?.url ?? PLATFORMS[p].url ?? '').trim();
      if (!parseRtmpUrl(address)) throw new Error('That server address is not an RTMP address: it starts rtmp:// or rtmps://, then the server and the app.');
      const entry = {
        id: at >= 0 ? id : crypto.randomUUID(),
        platform: p,
        name: String(name || was?.name || PLATFORMS[p].name).slice(0, 80),
        url: address,
        on: on === undefined ? Boolean(was?.on) : Boolean(on),
        shape: shapeName(shape ?? was?.shape ?? (PLATFORMS[p].tall ? 'tall' : 'wide')),
      };
      if (key != null && String(key).trim()) {
        if (!secrets.available()) throw new Error('This computer has no keystore to keep a stream key safely, so it is not kept.');
        secrets.set(keyName(entry.id), String(key).trim());
      }
      if (at >= 0) list[at] = entry;
      else list.push(entry);
      write(list);
      return read().map(shown);
    },

    /** ({ id }) -> destinations: the stream and its key both gone, and off the broadcast if it was on it. */
    removeDestination: ({ id } = {}) => {
      broadcast?.remove(id);
      write(read().filter((d) => d.id !== id));
      secrets.delete(keyName(id));
      return read().map(shown);
    },

    /** ({ ids }) -> destinations: which streams are on, until the setup keeps that (plan A12). */
    setOn: ({ ids = [] } = {}) => {
      const wanted = new Set(ids);
      write(read().map((d) => ({ ...d, on: wanted.has(d.id) })));
      return read().map(shown);
    },

    /**
     * ({ ids?, video: { width, height, fps, kbps }, tall?: same, audio: { kbps, sampleRate, channels } }, win)
     * -> { started: [id], skipped: [{ id, name, reason }] }: a broadcast opened,
     * each chosen stream that has a key connecting. Without `ids`, the streams
     * that are on. It stays open with none going: nothing here refuses that.
     */
    start: async ({ ids, video = {}, tall = null, audio = {} } = {}, win) => {
      if (broadcast?.running) throw new Error('Already live.');
      const all = read();
      const wanted = new Set(ids || []);
      const chosen = ids ? all.filter((d) => wanted.has(d.id)) : all.filter((d) => d.on);
      window = win;
      broadcast = new Broadcast({ video, tall, audio, makePublisher, onChange: (status) => { if (window) sendEvent(window, 'live:status', status); } });
      const started = [];
      const skipped = [];
      for (const d of chosen) {
        const key = secrets.get(keyName(d.id));
        if (!key) { skipped.push({ id: d.id, name: d.name, reason: 'no stream key' }); continue; }
        join(d, key);
        started.push(d.id);
      }
      return { started, skipped };
    },

    /** ({ id }) -> status: one stream joins the open broadcast. It may be the first of its shape. */
    startOne: ({ id } = {}) => {
      if (!broadcast?.running) throw new Error('There is no broadcast open.');
      const d = read().find((x) => x.id === id);
      const key = secrets.get(keyName(id));
      if (!d || !key) throw new Error('That stream has no key yet.');
      join(d, key);
      return broadcast.status();
    },

    /** ({ id }) -> status: one stream leaves; the broadcast and the others carry on. */
    stopOne: ({ id } = {}) => {
      broadcast?.remove(id);
      return statusNow();
    },

    /**
     * One piece of the program: ({ kind: 'video' | 'audio', shape?, timestamp
     * (ms), data, key, config }). A configuration is kept for those who join later.
     */
    send: (piece) => (broadcast?.running ? broadcast.send(piece) : { live: 0 }),

    /** () -> { live, destinations: [{ id, name, shape, state, detail, refused, bytes, dropped, since, reconnects }], kbps, target, rates } */
    status: () => statusNow(),

    /** () -> status: every stream closed, nothing tried again. */
    stop: () => (broadcast ? broadcast.close() : empty()),
  };
}
