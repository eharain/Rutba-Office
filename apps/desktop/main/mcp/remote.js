// AI Tools over the network, for an AI app or service on another machine
// (owner, 2026-10-21: "on ai tools i guess the user should have choice to open
// remote connections or local access only"). Off unless the person chooses it
// in the AI Tools window: the tools are then served by this running Rutba
// Office over HTTP, as the Model Context Protocol's Streamable HTTP transport
// asks (one address, a JSON-RPC message posted, its answer returned as JSON),
// to whoever shows the key the window gives. Local only, the default, keeps
// to what it was: nothing listens on the network.
//
// Every request is asked afresh whether the tools are on and the network is
// still chosen, so turning either off in the window refuses the next request
// at once; the listener itself closes a moment later. A request with no key,
// or the wrong one, is refused and counted, and an address that keeps trying
// is refused outright for a while. A request a web page sends (it carries an
// Origin) is refused too: a page in the person's own browser must not reach
// the tools by naming this computer's address.

import http from 'node:http';
import os from 'node:os';
import crypto from 'node:crypto';
import { createServer as createProtocol } from './protocol.js';

export const REACH_KEY = 'ai.tools.reach';
export const REACHES = ['local', 'remote'];
export const REMOTE_KEY = 'ai.tools.remoteKey';
export const REMOTE_PORT_KEY = 'ai.tools.remotePort';
export const DEFAULT_PORT = 47613;
const PATH = '/mcp';
const MAX_BODY = 4 * 1024 * 1024;
const TRIES = 10; // wrong keys from one address in a minute before it is turned away
const SESSION_IDLE = 60 * 60 * 1000;

/** A new key: 32 random bytes, as letters and digits a person can copy. */
export const newKey = () => crypto.randomBytes(32).toString('base64url');

/** The addresses this computer has on its networks, for the window to show. */
export function addresses({ interfaces = os.networkInterfaces() } = {}) {
  const out = [];
  for (const list of Object.values(interfaces)) {
    for (const a of list || []) {
      if (a.internal || a.family !== 'IPv4' || a.address.startsWith('169.254.')) continue;
      out.push(a.address);
    }
  }
  return [...new Set(out)];
}

/** Whether two keys are the same, in the same time whatever they hold. */
function sameKey(given, key) {
  const a = Buffer.from(String(given || ''));
  const b = Buffer.from(String(key || ''));
  return a.length === b.length && b.length > 0 && crypto.timingSafeEqual(a, b);
}

/**
 * The network side. `allowed()` answers { open, reason } for each request:
 * the tools on and the network chosen. `key()` is the key as it is now.
 * `onSession(id, entry | null)` tells who connected, from where, and when
 * they went.
 */
export function createRemote({ tools, version, instructions, allowed, key, onSession = () => {}, onRefused = () => {} }) {
  const sessions = new Map(); // id -> { from, client, since, last }
  const failures = new Map(); // address -> [times]
  let server = null;
  let listening = null; // { port }

  const turnedAway = (from) => {
    const now = Date.now();
    const recent = (failures.get(from) || []).filter((t) => now - t < 60_000);
    failures.set(from, recent);
    return recent.length >= TRIES;
  };
  const failed = (from) => failures.set(from, [...(failures.get(from) || []), Date.now()]);

  const protocolFor = (sessionId, from) => createProtocol({
    name: 'rutba-office',
    version,
    // Each call logged with the address it came from.
    tools: { list: () => tools.list(), call: (name, args) => tools.call(name, args, { where: from }) },
    instructions,
    onInitialize: ({ client }) => {
      const entry = { client: String(client?.name || 'An AI app').slice(0, 80), clientVersion: String(client?.version || '').slice(0, 40), where: from, since: new Date().toISOString(), last: null };
      sessions.set(sessionId, entry);
      onSession(sessionId, entry);
    },
    onCall: (tool) => {
      const s = sessions.get(sessionId);
      if (!s) return;
      s.last = new Date().toISOString();
      s.lastTool = String(tool || '').slice(0, 60);
      onSession(sessionId, s);
    },
  });

  const send = (res, status, body, headers = {}) => {
    const text = body === undefined ? '' : JSON.stringify(body);
    res.writeHead(status, { 'content-type': 'application/json', 'cache-control': 'no-store', ...headers });
    res.end(text);
  };
  const rpcError = (code, message) => ({ jsonrpc: '2.0', id: null, error: { code, message } });

  async function onRequest(req, res) {
    const from = String(req.socket.remoteAddress || '').replace(/^::ffff:/, '');
    const url = new URL(req.url || '/', 'http://x');
    if (url.pathname !== PATH) return send(res, 404, rpcError(-32000, 'Rutba Office answers at /mcp.'));
    if (turnedAway(from)) return send(res, 429, rpcError(-32000, 'Too many wrong keys from this address. Wait a minute.'));
    // A page in a browser names its origin; an AI app or service does not.
    if (req.headers.origin) {
      onRefused({ from, why: 'a web page' });
      return send(res, 403, rpcError(-32000, 'Rutba Office does not answer web pages.'));
    }
    const bearer = /^Bearer\s+(.+)$/i.exec(String(req.headers.authorization || ''))?.[1]?.trim();
    if (!sameKey(bearer, key())) {
      failed(from);
      onRefused({ from, why: 'no key or a wrong key' });
      return send(res, 401, rpcError(-32001, 'The key is missing or wrong. Rutba Office shows the key in its AI Tools window.'), { 'www-authenticate': 'Bearer realm="Rutba Office"' });
    }
    const gate = allowed();
    if (!gate.open) return send(res, 403, rpcError(-32002, gate.reason));

    const sessionId = String(req.headers['mcp-session-id'] || '');
    if (req.method === 'DELETE') {
      if (sessions.delete(sessionId)) onSession(sessionId, null);
      res.writeHead(204);
      return res.end();
    }
    // Nothing is pushed unasked, so there is no stream to open.
    if (req.method !== 'POST') return send(res, 405, rpcError(-32000, 'Post a JSON-RPC message.'), { allow: 'POST, DELETE' });

    let size = 0;
    const chunks = [];
    for await (const chunk of req) {
      size += chunk.length;
      if (size > MAX_BODY) return send(res, 413, rpcError(-32000, 'That message is too large.'));
      chunks.push(chunk);
    }
    let msg;
    try {
      msg = JSON.parse(Buffer.concat(chunks).toString('utf8'));
    } catch {
      return send(res, 400, rpcError(-32700, 'That is not JSON.'));
    }
    const messages = Array.isArray(msg) ? msg : [msg];
    const opening = messages.some((m) => m?.method === 'initialize');
    let id = sessionId;
    if (opening) id = crypto.randomUUID();
    else if (id && !sessions.has(id)) return send(res, 404, rpcError(-32000, 'That session has ended. Connect again.'));
    const protocol = protocolFor(id, from);
    const replies = (await Promise.all(messages.map((m) => protocol.handle(m)))).filter(Boolean);
    const headers = opening ? { 'mcp-session-id': id } : {};
    if (!replies.length) {
      res.writeHead(202, headers);
      return res.end();
    }
    return send(res, 200, Array.isArray(msg) ? replies : replies[0], headers);
  }

  /** Sessions quiet for an hour are taken as gone. */
  const sweep = setInterval(() => {
    const now = Date.now();
    for (const [id, s] of sessions) {
      if (now - Date.parse(s.last || s.since) > SESSION_IDLE) { sessions.delete(id); onSession(id, null); }
    }
  }, 60_000);
  sweep.unref?.();

  return {
    /** Listening on `port`, on every network this computer is on; answers { port } or throws why not. */
    async start(port) {
      if (listening?.port === port) return listening;
      await this.stop();
      server = http.createServer((req, res) => { onRequest(req, res).catch((err) => send(res, 500, rpcError(-32603, String(err?.message || err)))); });
      server.headersTimeout = 15_000;
      server.requestTimeout = 120_000;
      await new Promise((resolve, reject) => {
        server.once('error', (err) => reject(err.code === 'EADDRINUSE' ? new Error(`Port ${port} is in use by another program. Choose another port.`) : err));
        server.listen(port, '0.0.0.0', resolve);
      });
      listening = { port: server.address().port };
      return listening;
    },
    async stop() {
      if (!server) return;
      const s = server;
      server = null;
      listening = null;
      for (const id of sessions.keys()) onSession(id, null);
      sessions.clear();
      await new Promise((resolve) => { s.close(() => resolve()); s.closeAllConnections?.(); });
    },
    listening: () => listening,
    sessions: () => [...sessions.entries()].map(([id, s]) => ({ id, ...s })),
  };
}
