// Going live from the Recorder (owner, 2026-10-19: "stream videos live on
// youtube and facebook etc").
//
// The destinations a person streams to are kept here: each one's platform, a
// name, its server address and whether it is on, in the settings; its stream
// key in the secrets, encrypted by the operating system's keystore, and never
// handed back to a window (the window sees only its last four characters).
//
// Going live opens one RTMP publisher per destination that is on and has a
// key (@rutba/stream). The window encodes the program once, H.264 and AAC,
// and sends each piece here; the same bytes go to every destination. The
// metadata and the two configurations are kept, so a destination that joins
// late, or comes back after the line dropped, is given them first and then
// waits for the next key frame. A destination that drops is tried again,
// after 2, 4, 8, 15 and then 30 seconds, until Stop.
//
// Nothing connects until the person presses Go Live, and only to the servers
// they set up. Nothing here listens.

import crypto from 'node:crypto';
import { RtmpPublisher, videoBody, audioBody, metadataBody, parseRtmpUrl, maskKey, PLATFORMS } from '@rutba/stream';
import { BitrateGovernor } from '@rutba/stream/adapt';
import { sendEvent } from '@rutba/office-shell/electron/main';

const LIST_KEY = 'live.destinations';
const keyName = (id) => `live.key.${id}`;
const BACKOFF = [2, 4, 8, 15, 30];

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
    return { id: d.id, platform: d.platform, name: d.name, url: d.url, on: Boolean(d.on), hasKey: Boolean(key), keyHint: maskKey(key) };
  };

  // The live session: one entry per destination, and what every late joiner needs first.
  let session = null;

  const statusOf = () => (session ? [...session.dest.values()].map((s) => ({
    id: s.id,
    name: s.name,
    state: s.state,
    detail: s.detail,
    bytes: s.pub ? s.pub.sent : s.bytes,
    dropped: s.pub ? s.pub.dropped : s.dropped,
    since: s.since,
    reconnects: s.reconnects,
  })) : []);

  const tell = () => {
    if (session?.win) sendEvent(session.win, 'live:status', { live: Boolean(session.running), destinations: statusOf(), kbps: session.kbps, target: session.governor?.target || session.kbps });
  };

  /** Bring one destination's publisher up, and give it what it needs before frames. */
  async function connect(s) {
    if (!session?.running || s.stopped) return;
    s.state = s.reconnects ? 'reconnecting' : 'connecting';
    s.detail = '';
    tell();
    let pub;
    try {
      pub = makePublisher({ url: s.url, key: s.key });
    } catch (err) {
      s.state = 'error';
      s.detail = err.message;
      tell();
      return;
    }
    s.pub = pub;
    s.seenDropped = 0;
    pub.on('error', (err) => { s.detail = err.message; });
    pub.on('close', () => {
      if (s.pub !== pub) return;
      s.bytes += pub.sent;
      s.dropped += pub.dropped;
      s.pub = null;
      if (!session?.running || s.stopped) return;
      // Dropped while live: try again, waiting longer each time.
      const wait = BACKOFF[Math.min(s.reconnects, BACKOFF.length - 1)];
      s.reconnects += 1;
      s.state = 'reconnecting';
      s.detail = s.detail || 'The connection dropped.';
      tell();
      s.retry = setTimeout(() => connect(s), wait * 1000);
    });
    try {
      await pub.start();
      if (!session?.running || s.stopped) { pub.close(); return; }
      s.state = 'live';
      s.detail = '';
      s.since = s.since || Date.now();
      if (session.meta) pub.sendMetadata(session.meta);
      if (session.videoConfig) pub.sendVideo(session.videoConfig, session.last, { config: true });
      if (session.audioConfig) pub.sendAudio(session.audioConfig, session.last);
      tell();
    } catch (err) {
      // A refused key is said once and not tried again; anything else is tried again by the close above.
      s.detail = err.message;
      if (/refused/.test(err.message)) {
        s.stopped = true;
        clearTimeout(s.retry);
        s.state = 'error';
        tell();
      }
    }
  }

  return {
    /** () -> [{ id, platform, name, url, on, hasKey, keyHint }] */
    destinations: () => read().map(shown),

    /** The platforms offered, with their usual server and where their key is found. */
    platforms: () => PLATFORMS,

    /**
     * ({ id?, platform, name?, url?, on?, key? }) -> destinations. A key given
     * replaces the one kept; no key leaves it as it is.
     */
    saveDestination: ({ id, platform = 'custom', name, url, on = true, key } = {}) => {
      const p = PLATFORMS[platform] ? platform : 'custom';
      const address = String(url || PLATFORMS[p].url || '').trim();
      if (!parseRtmpUrl(address)) throw new Error('That server address is not an RTMP address: it starts rtmp:// or rtmps://, then the server and the app.');
      const list = read();
      const at = id ? list.findIndex((d) => d.id === id) : -1;
      const entry = { id: at >= 0 ? id : crypto.randomUUID(), platform: p, name: String(name || PLATFORMS[p].name).slice(0, 80), url: address, on: Boolean(on) };
      if (key != null && String(key).trim()) {
        if (!secrets.available()) throw new Error('This computer has no keystore to keep a stream key safely, so it is not kept.');
        secrets.set(keyName(entry.id), String(key).trim());
      }
      if (at >= 0) list[at] = entry;
      else list.push(entry);
      write(list);
      return read().map(shown);
    },

    /** ({ id }) -> destinations: the destination and its key both gone. */
    removeDestination: ({ id } = {}) => {
      write(read().filter((d) => d.id !== id));
      secrets.delete(keyName(id));
      return read().map(shown);
    },

    /**
     * ({ video: { width, height, fps, kbps }, audio: { kbps, sampleRate, channels } }, win)
     * -> { started: [id], skipped: [{ id, name, reason }] }: every destination
     * that is on and has a key, connecting.
     */
    start: async ({ video = {}, audio = {} } = {}, win) => {
      if (session?.running) throw new Error('Already live.');
      const chosen = read().filter((d) => d.on);
      const skipped = [];
      const dest = new Map();
      for (const d of chosen) {
        const key = secrets.get(keyName(d.id));
        if (!key) { skipped.push({ id: d.id, name: d.name, reason: 'no stream key' }); continue; }
        dest.set(d.id, { id: d.id, name: d.name, url: d.url, key, state: 'waiting', detail: '', pub: null, bytes: 0, dropped: 0, since: 0, reconnects: 0, stopped: false, retry: null });
      }
      if (!dest.size) throw new Error(chosen.length ? 'No destination that is on has a stream key yet.' : 'Turn on a destination to go live to.');
      session = {
        running: true,
        win,
        dest,
        meta: metadataBody({ width: video.width, height: video.height, fps: video.fps, videoKbps: video.kbps, audioKbps: audio.kbps, sampleRate: audio.sampleRate || 48000, channels: audio.channels || 2 }),
        videoConfig: null,
        audioConfig: null,
        last: 0,
        // The picture's bitrate follows the slowest destination's line (@rutba/stream/adapt).
        governor: new BitrateGovernor({ target: Number(video.kbps) || 4500 }),
        kbps: Number(video.kbps) || 4500,
        ticker: null,
      };
      session.ticker = setInterval(() => {
        if (!session?.running) return;
        const looks = [...session.dest.values()].filter((s) => s.pub?.state === 'live').map((s) => {
          const dropped = Math.max(0, s.pub.dropped - (s.seenDropped || 0));
          s.seenDropped = s.pub.dropped;
          return { dropped, queued: (s.pub.socket?.writableLength || 0) / (s.pub.maxQueueBytes || 1) };
        });
        if (looks.length) session.kbps = session.governor.tick(looks);
        tell();
      }, 2000);
      for (const s of dest.values()) connect(s);
      tell();
      return { started: [...dest.keys()], skipped };
    },

    /**
     * One piece of the program: ({ kind: 'video' | 'audio', timestamp (ms),
     * data, key, config }). A configuration is kept for those who join later.
     */
    send: ({ kind, timestamp = 0, data, key = false, config = false } = {}) => {
      if (!session?.running || !data) return { live: 0 };
      const bytes = Buffer.from(data.buffer ? new Uint8Array(data.buffer, data.byteOffset, data.byteLength) : data);
      session.last = Math.max(session.last, Math.round(timestamp));
      let live = 0;
      if (kind === 'video') {
        const body = videoBody({ data: bytes, key, config });
        if (config) session.videoConfig = body;
        for (const s of session.dest.values()) if (s.pub?.state === 'live' && s.pub.sendVideo(body, timestamp, { config })) live += 1;
      } else if (kind === 'audio') {
        const body = audioBody({ data: bytes, config });
        if (config) session.audioConfig = body;
        for (const s of session.dest.values()) if (s.pub?.state === 'live' && s.pub.sendAudio(body, timestamp)) live += 1;
      }
      return { live };
    },

    /** () -> { live, destinations: [{ id, name, state, detail, bytes, dropped, since, reconnects }] } */
    status: () => ({ live: Boolean(session?.running), destinations: statusOf(), kbps: session?.kbps || 0, target: session?.governor?.target || 0 }),

    /** () -> status: every destination closed, nothing tried again. */
    stop: () => {
      if (!session) return { live: false, destinations: [] };
      session.running = false;
      clearInterval(session.ticker);
      for (const s of session.dest.values()) {
        s.stopped = true;
        clearTimeout(s.retry);
        s.pub?.close('stopped');
        s.state = 'ended';
      }
      const out = { live: false, destinations: statusOf() };
      tell();
      return out;
    },
  };
}
