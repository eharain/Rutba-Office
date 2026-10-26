// The tools' way into the running Rutba Office of their profile. The tools
// an AI app starts (node.js) are a process of their own: files they read and
// make themselves, but mail, calendars, address books and databases they ask
// the running Rutba Office for, over a local channel (a named pipe on
// Windows, a socket in the profile elsewhere; never a network port). The
// running copy holds those stores, open and unlocked, as its windows use
// them, and is where the person sees and answers what needs confirming: so
// nothing the tools do there races a window's own writes, and a protected
// profile's sealed files are read only while it is open and unlocked.
//
// The running copy writes ai-tools-channel.json in the profile while the
// tools are on: where the channel is and a key made for this run. A request
// without that key is not answered. One JSON message a line each way:
// { id, key, method, params } in, { id, result } or { id, error } out.

import fs from 'node:fs';
import os from 'node:os';
import net from 'node:net';
import path from 'node:path';
import crypto from 'node:crypto';

export const CHANNEL_FILE = 'ai-tools-channel.json';
const MAX_LINE = 8 * 1024 * 1024;

/** The channel's name for a profile folder: the same from both sides, unique to the folder. */
export function channelPath(folder) {
  const id = crypto.createHash('sha256').update(path.resolve(folder).toLowerCase()).digest('hex').slice(0, 20);
  return process.platform === 'win32' ? `\\\\.\\pipe\\rutba-office-ai-${id}` : path.join(os.tmpdir(), `rutba-office-ai-${id}.sock`);
}

/**
 * The running copy's side. `handle(method, params)` answers a request.
 * start() listens and writes the channel file; stop() closes it and takes
 * the file away.
 */
export function createChannelServer({ folder, handle, pid = process.pid }) {
  const where = channelPath(folder);
  const file = path.join(folder, CHANNEL_FILE);
  let server = null;
  let key = null;

  const same = (given) => {
    const a = Buffer.from(String(given || ''));
    const b = Buffer.from(String(key || ''));
    return a.length === b.length && b.length > 0 && crypto.timingSafeEqual(a, b);
  };

  const onConnection = (socket) => {
    socket.setEncoding('utf8');
    let buffered = '';
    const reply = (obj) => { if (!socket.destroyed) socket.write(`${JSON.stringify(obj)}\n`); };
    socket.on('data', (chunk) => {
      buffered += chunk;
      if (buffered.length > MAX_LINE) { socket.destroy(); return; }
      let at = buffered.indexOf('\n');
      while (at >= 0) {
        const line = buffered.slice(0, at).trim();
        buffered = buffered.slice(at + 1);
        at = buffered.indexOf('\n');
        if (!line) continue;
        let msg;
        try {
          msg = JSON.parse(line);
        } catch {
          reply({ id: null, error: 'That line is not JSON.' });
          continue;
        }
        if (!same(msg.key)) {
          reply({ id: msg.id ?? null, error: 'The key is missing or wrong.' });
          continue;
        }
        Promise.resolve()
          .then(() => handle(String(msg.method || ''), msg.params || {}))
          .then((result) => reply({ id: msg.id, result: result ?? null }), (err) => reply({ id: msg.id, error: err?.message || String(err) }));
      }
    });
    socket.on('error', () => {});
  };

  return {
    where,
    async start() {
      if (server) return;
      key = crypto.randomBytes(32).toString('base64url');
      if (process.platform !== 'win32') { try { fs.rmSync(where, { force: true }); } catch { /* none there */ } }
      server = net.createServer(onConnection);
      await new Promise((resolve, reject) => {
        server.once('error', reject);
        server.listen(where, () => { server.off('error', reject); resolve(); });
      });
      fs.mkdirSync(folder, { recursive: true });
      const tmp = `${file}.${pid}.tmp`;
      fs.writeFileSync(tmp, JSON.stringify({ where, key, pid }), { mode: 0o600 });
      fs.renameSync(tmp, file);
    },
    async stop() {
      if (!server) return;
      const s = server;
      server = null;
      try {
        const kept = JSON.parse(fs.readFileSync(file, 'utf8'));
        if (kept.key === key) fs.rmSync(file, { force: true });
      } catch {
        /* already gone */
      }
      key = null;
      await new Promise((resolve) => s.close(() => resolve()));
    },
    listening: () => Boolean(server),
  };
}

/**
 * The tools' side: request(method, params) answers what the running copy
 * said, or throws { notRunning: true } when no Rutba Office of this profile
 * is running with the tools on.
 */
export function createChannelClient({ folder, timeoutMs = 120_000, isAlive = null }) {
  const file = path.join(folder, CHANNEL_FILE);
  let socket = null;
  let connecting = null;
  let buffered = '';
  let next = 1;
  const waiting = new Map();

  const notRunning = () => Object.assign(new Error('Rutba Office is not open in this profile.'), { notRunning: true });
  const drop = (err) => {
    socket = null;
    connecting = null;
    for (const [, w] of waiting) w.reject(err || notRunning());
    waiting.clear();
  };

  const info = () => {
    try {
      const kept = JSON.parse(fs.readFileSync(file, 'utf8'));
      if (isAlive && !isAlive(Number(kept.pid))) return null;
      return kept;
    } catch {
      return null;
    }
  };

  const connect = () => {
    if (socket && !socket.destroyed) return Promise.resolve(socket);
    if (connecting) return connecting;
    const kept = info();
    if (!kept?.where || !kept?.key) return Promise.reject(notRunning());
    connecting = new Promise((resolve, reject) => {
      const s = net.connect(kept.where);
      s.setEncoding('utf8');
      s.once('connect', () => { socket = s; s.key = kept.key; connecting = null; resolve(s); });
      s.once('error', () => { drop(); reject(notRunning()); });
      s.on('close', () => drop());
      s.on('data', (chunk) => {
        buffered += chunk;
        let at = buffered.indexOf('\n');
        while (at >= 0) {
          const line = buffered.slice(0, at);
          buffered = buffered.slice(at + 1);
          at = buffered.indexOf('\n');
          let msg;
          try { msg = JSON.parse(line); } catch { continue; }
          const w = waiting.get(msg.id);
          if (!w) continue;
          waiting.delete(msg.id);
          clearTimeout(w.timer);
          if (msg.error) w.reject(new Error(msg.error));
          else w.resolve(msg.result);
        }
      });
      s.unref?.();
    });
    return connecting;
  };

  return {
    /** Whether a running Rutba Office answers on this profile's channel. */
    async request(method, params = {}) {
      const s = await connect();
      const id = next++;
      return new Promise((resolve, reject) => {
        const timer = setTimeout(() => {
          waiting.delete(id);
          reject(new Error('Rutba Office did not answer in time.'));
        }, timeoutMs);
        timer.unref?.();
        waiting.set(id, { resolve, reject, timer });
        s.write(`${JSON.stringify({ id, key: s.key, method, params })}\n`);
      });
    },
    close() {
      socket?.destroy();
      drop(new Error('closed'));
    },
  };
}
