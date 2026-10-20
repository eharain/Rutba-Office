// Talk: calls and messages with the people you are in contact with (owner,
// 2026-10-20, D-2026-10-20-11 and -12: "the calls we only allow between
// contacts that are established on both sides", "the Rutba sign in to begin
// with", "for calls lets assume we have few TURN servers and we can give user
// to choose from or let us choose one of them and to begin with one of user
// hosts the TURN for their own").
//
// The introducer (workers/introducer) is the one server this talks to: it
// knows who you are from your Rutba sign-in, who your contacts are, who is
// online, and passes a call's set-up between the two computers. The call
// itself goes straight between them, or through a relay. This service holds
// the one connection for every window, keeps your messages on this computer,
// and hands each window what the introducer says.
//
// Nothing connects until you sign in, and only to the introducer.

import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';

export const SERVER_KEY = 'talk.server';
export const DEFAULT_SERVER = 'wss://introducer.rutba.io/ws';
const TOKEN_SECRET = 'talk.token';
const OWN_RELAY_KEY = 'talk.ownRelay';
const OWN_RELAY_SECRET = 'talk.ownRelay.credential';
const RELAY_CHOICE_KEY = 'talk.relay';
const PRESENCE_KEY = 'talk.presence';
const BACKOFF = [2, 4, 8, 15, 30];
const HISTORY_CAP = 5000;

/**
 * @param {object} o
 * @param {object} o.stores the settings and the secrets
 * @param {(event: string, payload: object) => void} o.broadcast to every window
 * @param {string} o.dataDir where the messages are kept
 * @param {(url: string) => WebSocket} [o.connect] how a connection is made (the checks hand in their own)
 */
export function createTalkService({ stores, broadcast, dataDir, connect = (url) => new WebSocket(url) }) {
  const settings = stores.settings;
  const secrets = stores.secrets;
  let socket = null;
  let retry = 0;
  let retryTimer = null;
  let wanted = false;
  let state = { phase: 'signed-out', me: null, contacts: [], requests: { incoming: [], outgoing: [] }, blocked: [], relays: [], error: '' };

  const tell = () => broadcast?.('talk:state', snapshot());
  const snapshot = () => ({ ...state, server: settings.get(SERVER_KEY, DEFAULT_SERVER), relay: relayChoice(), presence: settings.get(PRESENCE_KEY, 'online') });

  /* ── the messages kept here, a file per contact ───────────────────── */

  const folder = () => (state.me ? path.join(dataDir, crypto.createHash('sha256').update(String(state.me.id)).digest('hex').slice(0, 24)) : null);
  const fileFor = (contactId) => path.join(folder(), `${crypto.createHash('sha256').update(String(contactId)).digest('hex').slice(0, 24)}.json`);
  const readHistory = (contactId) => {
    if (!folder()) return [];
    try { return JSON.parse(fs.readFileSync(fileFor(contactId), 'utf8')); } catch { return []; }
  };
  const keep = (contactId, entry) => {
    if (!folder()) return;
    const list = readHistory(contactId).filter((m) => m.id !== entry.id);
    list.push(entry);
    fs.mkdirSync(folder(), { recursive: true });
    fs.writeFileSync(fileFor(contactId), JSON.stringify(list.slice(-HISTORY_CAP)));
  };

  /* ── relays: Rutba's, one of them by choice, or the person's own ──── */

  /** 'auto', a relay's id, or 'own'. */
  const relayChoice = () => settings.get(RELAY_CHOICE_KEY, 'auto');
  const ownRelay = () => {
    const own = settings.get(OWN_RELAY_KEY, null);
    if (!own?.urls?.length) return null;
    return { id: 'own', name: own.name || 'My relay', urls: own.urls, username: own.username || '', credential: secrets.get(OWN_RELAY_SECRET) || '' };
  };
  /** The relays a call is made with, as an RTCPeerConnection takes them. */
  const iceServers = () => {
    const ice = (r) => (r.username ? { urls: r.urls, username: r.username, credential: r.credential } : { urls: r.urls });
    const finders = state.relays.filter((r) => r.urls.every((u) => /^stuns?:/.test(u)));
    const turns = state.relays.filter((r) => r.urls.some((u) => /^turns?:/.test(u)));
    const choice = relayChoice();
    if (choice === 'own') return [...finders, ...(ownRelay() ? [ownRelay()] : turns)].map(ice);
    if (choice !== 'auto') return [...finders, ...(turns.filter((r) => r.id === choice).length ? turns.filter((r) => r.id === choice) : turns)].map(ice);
    return [...finders, ...turns].map(ice);
  };

  /* ── the connection ──────────────────────────────────────────────── */

  const handle = (msg) => {
    switch (msg.type) {
      case 'welcome':
        retry = 0;
        state = { ...state, phase: 'online', me: msg.me, contacts: msg.contacts || [], requests: msg.requests || { incoming: [], outgoing: [] }, blocked: msg.blocked || [], relays: msg.relays || [], error: '' };
        break;
      case 'relays':
        state = { ...state, relays: msg.relays || [] };
        break;
      case 'presence':
        state = { ...state, contacts: state.contacts.map((c) => (c.id === msg.id ? { ...c, presence: msg.presence } : c)) };
        break;
      case 'contact.added':
        state = {
          ...state,
          contacts: [...state.contacts.filter((c) => c.id !== msg.contact.id), msg.contact].sort((a, b) => String(a.name).localeCompare(String(b.name))),
          requests: { incoming: state.requests.incoming.filter((r) => r.id !== msg.requestId && r.from?.id !== msg.contact.id), outgoing: state.requests.outgoing.filter((r) => r.id !== msg.requestId && r.email !== msg.contact.email) },
        };
        break;
      case 'contact.removed':
      case 'contact.blocked':
        state = { ...state, contacts: state.contacts.filter((c) => c.id !== msg.id), blocked: msg.type === 'contact.blocked' ? [...new Set([...state.blocked, msg.id])] : state.blocked };
        break;
      case 'contact.unblocked':
        state = { ...state, blocked: state.blocked.filter((b) => b !== msg.id) };
        break;
      case 'contact.requested':
        state = { ...state, requests: { ...state.requests, incoming: [...state.requests.incoming.filter((r) => r.id !== msg.request.id), msg.request] } };
        break;
      case 'contact.request.sent':
        if (msg.request) state = { ...state, requests: { ...state.requests, outgoing: [...state.requests.outgoing.filter((r) => r.id !== msg.request.id), msg.request] } };
        if (msg.contact) handle({ type: 'contact.added', contact: msg.contact });
        break;
      case 'contact.request.gone':
        state = { ...state, requests: { incoming: state.requests.incoming.filter((r) => r.id !== msg.requestId), outgoing: state.requests.outgoing.filter((r) => r.id !== msg.requestId) } };
        break;
      case 'messages': {
        const ids = [];
        for (const m of msg.messages || []) {
          keep(m.from, { id: m.id, from: m.from, to: state.me?.id, text: m.text, at: m.at, mine: false });
          ids.push(m.id);
        }
        // Kept here now, so the introducer lets them go.
        if (ids.length) send({ type: 'msg.ack', ids });
        break;
      }
      case 'msg.copy':
        keep(msg.message.to, { ...msg.message, mine: true });
        break;
      case 'error':
        if (msg.for === 'hello') state = { ...state, phase: 'signed-out', error: msg.message || 'The sign-in was not accepted.' };
        break;
      default:
        break;
    }
    // Every window hears every message: a call's set-up goes to the window that holds the call.
    broadcast?.('talk:event', msg);
    tell();
  };

  const send = (msg) => {
    if (!socket || socket.readyState !== 1) throw new Error('Talk is not connected.');
    socket.send(JSON.stringify(msg));
  };

  const open = () => {
    clearTimeout(retryTimer);
    const token = secrets.get(TOKEN_SECRET);
    if (!token) { state = { ...state, phase: 'signed-out' }; tell(); return; }
    state = { ...state, phase: 'connecting', error: '' };
    tell();
    let ws;
    try {
      ws = connect(settings.get(SERVER_KEY, DEFAULT_SERVER));
    } catch (err) {
      state = { ...state, phase: 'offline', error: err.message };
      tell();
      return;
    }
    socket = ws;
    ws.onopen = () => ws.send(JSON.stringify({ type: 'hello', token, presence: settings.get(PRESENCE_KEY, 'online') }));
    ws.onmessage = (e) => {
      let msg;
      try { msg = JSON.parse(typeof e.data === 'string' ? e.data : Buffer.from(e.data).toString('utf8')); } catch { return; }
      handle(msg);
    };
    ws.onclose = () => {
      if (socket !== ws) return;
      socket = null;
      if (!wanted || state.phase === 'signed-out') { tell(); return; }
      state = { ...state, phase: 'offline' };
      tell();
      const wait = BACKOFF[Math.min(retry, BACKOFF.length - 1)] * 1000;
      retry += 1;
      retryTimer = setTimeout(open, wait);
      retryTimer.unref?.();
    };
    ws.onerror = () => {};
  };

  return {
    state: () => snapshot(),

    /**
     * Sign in with a Rutba sign-in token. (The browser sign-in at
     * auth.rutba.io hands the token here; until it is set up, a token can be
     * given outright for a test server.)
     */
    signIn: ({ token, server } = {}) => {
      if (!token) throw new Error('A sign-in is needed.');
      if (server) {
        if (!/^wss?:\/\//.test(server)) throw new Error('The server address starts with wss://.');
        settings.set(SERVER_KEY, server);
      }
      secrets.set(TOKEN_SECRET, String(token));
      wanted = true;
      open();
      return snapshot();
    },

    signOut: () => {
      wanted = false;
      clearTimeout(retryTimer);
      secrets.delete(TOKEN_SECRET);
      const ws = socket;
      socket = null;
      try { ws?.close(1000, 'signed out'); } catch { /* gone */ }
      state = { phase: 'signed-out', me: null, contacts: [], requests: { incoming: [], outgoing: [] }, blocked: [], relays: [], error: '' };
      tell();
      return snapshot();
    },

    /** Connect again with the sign-in kept, as at start-up. */
    resume: () => {
      if (secrets.get(TOKEN_SECRET)) { wanted = true; if (!socket) open(); }
      return snapshot();
    },

    /** Anything for the introducer: a contact request, presence, a call's set-up. */
    send: (msg) => {
      if (!msg || typeof msg.type !== 'string' || msg.type === 'hello') throw new Error('That is not something Talk sends.');
      if (msg.type === 'presence') settings.set(PRESENCE_KEY, msg.presence);
      send(msg);
      return { sent: true };
    },

    /** A message to a contact: kept here, and sent. */
    message: ({ to, text }) => {
      const id = `m${crypto.randomBytes(9).toString('base64url')}`;
      const entry = { id, from: state.me?.id, to, text: String(text || ''), at: Date.now(), mine: true };
      send({ type: 'msg.send', to, id, text: entry.text });
      keep(to, entry);
      return entry;
    },

    history: ({ contactId, limit = 200 } = {}) => readHistory(contactId).slice(-limit),

    relays: () => ({ choice: relayChoice(), relays: state.relays.map(({ id, name, region, urls }) => ({ id, name, region, urls })), own: ownRelay() ? { ...ownRelay(), credential: ownRelay().credential ? '••••' : '' } : null, ice: iceServers() }),

    /** Rutba's relays automatically, one of them, or the person's own. */
    setRelay: ({ choice, own } = {}) => {
      if (own) {
        const urls = [].concat(own.urls || []).map(String).filter((u) => /^(turn|turns|stun|stuns):\S+$/.test(u));
        if (!urls.length) throw new Error('A relay address starts with turn: or turns:.');
        settings.set(OWN_RELAY_KEY, { name: String(own.name || 'My relay').slice(0, 80), urls, username: String(own.username || '').slice(0, 200) });
        if (own.credential) secrets.set(OWN_RELAY_SECRET, String(own.credential));
      }
      if (choice) settings.set(RELAY_CHOICE_KEY, String(choice));
      return { choice: relayChoice(), ice: iceServers() };
    },

    /** What a call is made with: the relays, and the person's own one to offer the other side. */
    callSetup: () => ({ iceServers: iceServers(), offerRelay: relayChoice() === 'own' && ownRelay() ? { urls: ownRelay().urls, username: ownRelay().username, credential: ownRelay().credential } : null }),
  };
}
