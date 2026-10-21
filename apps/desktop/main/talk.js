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
// Inviting someone (owner, 2026-10-21, D-2026-10-21-27 and -28): asking an
// address to be in contact also emails it, from the person's own Mail account
// when they have one, else from Rutba (no-reply@rutba.io, with a daily and a
// weekly limit the introducer keeps), and the person sees which it was. Or
// they make an invitation link and send it over any channel of their own.
// Either way the email or the link opens Talk on the invitation, or, for a
// friend without Rutba Office, a page where a Rutba account is enough to
// accept.
//
// Nothing connects until you sign in, and only to the introducer. Signing in
// is Rutba's own, in the browser (talk-signin.js, workers/introducer
// docs/SIGN-IN.md): the access token goes to the introducer, refreshed before
// it runs out with the refresh token kept beside it; a token can still be
// given outright for a test or self-run introducer.

import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { ISSUER, signInWithBrowser, refreshTokens, revoke, expiresAt } from './talk-signin.js';

export const SERVER_KEY = 'talk.server';
export const DEFAULT_SERVER = 'wss://introducer.rutba.io/ws';
const TOKEN_SECRET = 'talk.token';
const REFRESH_SECRET = 'talk.refresh';
/** Where Rutba sign-in is: auth.rutba.io, or a stand-in under the checks. */
const ISSUER_KEY = 'talk.issuer';
/** How long before its end an access token is renewed. */
const RENEW_BEFORE = 5 * 60 * 1000;
const OWN_RELAY_KEY = 'talk.ownRelay';
const OWN_RELAY_SECRET = 'talk.ownRelay.credential';
const RELAY_CHOICE_KEY = 'talk.relay';
const PRESENCE_KEY = 'talk.presence';
/** How each request was told to its person: { [requestId]: { by, from?, sent, reason? } }, the introducer knows only Rutba's. */
const TOLD_KEY = 'talk.told';
const TOLD_CAP = 500;
/** How long an answer to a question put to the introducer is waited for. */
const ASK_MS = 15000;
/** Asking with Rutba's email waits for the email to be decided: up to 20 s on a slow relay. */
const ASK_MAIL_MS = 30000;
/** Where a friend without Rutba Office gets it (D-2026-10-21-28). */
export const DOWNLOAD_URL = 'https://office.rutba.io/download';
const BACKOFF = [2, 4, 8, 15, 30];
const HISTORY_CAP = 5000;

/**
 * @param {object} o
 * @param {object} o.stores the settings and the secrets
 * @param {(event: string, payload: object) => void} o.broadcast to every window
 * @param {string} o.dataDir where the messages are kept
 * @param {(url: string) => WebSocket} [o.connect] how a connection is made (the checks hand in their own)
 * @param {(url: string) => Promise<void>|void} [o.openBrowser] opens the system browser, for the sign-in
 * @param {boolean} [o.testServers] whether the browser sign-in may be pointed at another introducer (the checks only)
 */
export function createTalkService({ stores, broadcast, dataDir, connect = (url) => new WebSocket(url), openBrowser = null, testServers = false, fetchImpl = (...a) => fetch(...a), mail = null }) {
  const settings = stores.settings;
  const secrets = stores.secrets;
  let socket = null;
  let retry = 0;
  let retryTimer = null;
  let wanted = false;
  let state = { phase: 'signed-out', me: null, contacts: [], requests: { incoming: [], outgoing: [] }, blocked: [], relays: [], invites: [], mailAllowance: null, error: '' };

  const tell = () => broadcast?.('talk:state', snapshot());
  const snapshot = () => {
    const told = settings.get(TOLD_KEY, {}) || {};
    // Each request with how its person was told, as the asker sees it.
    const outgoing = state.requests.outgoing.map((r) => ({ ...r, told: told[r.id] || null }));
    return { ...state, requests: { ...state.requests, outgoing }, server: settings.get(SERVER_KEY, DEFAULT_SERVER), relay: relayChoice(), presence: settings.get(PRESENCE_KEY, 'online') };
  };
  const keepTold = (requestId, told) => {
    const all = { ...(settings.get(TOLD_KEY, {}) || {}), [requestId]: { ...told, at: Date.now() } };
    const ids = Object.keys(all).sort((a, b) => all[a].at - all[b].at);
    for (const id of ids.slice(0, Math.max(0, ids.length - TOLD_CAP))) delete all[id];
    settings.set(TOLD_KEY, all);
  };

  /* ── questions put to the introducer, and their answers ──────────── */

  // Each waits for the answer that names its ref (or, for an answer that
  // carries none, the first of its type), or for a refusal of what it sent.
  const asking = [];
  let refs = 0;
  const answerOf = (msg) => {
    const i = asking.findIndex((a) => (msg.type === 'error' ? a.sent === msg.for : a.match(msg)));
    if (i < 0) return;
    const [a] = asking.splice(i, 1);
    clearTimeout(a.timer);
    a.resolve(msg.type === 'error' ? { error: msg.code || 'server' } : msg);
  };
  /** Sends `msg` and answers with what the introducer said: the answer, or { error: code }. */
  const ask = (msg, answer, match = null, waitMs = ASK_MS) => new Promise((resolve) => {
    const ref = `r${++refs}`;
    const a = { sent: msg.type, resolve, match: (m) => m.type === answer && (match ? match(m) : m.ref == null || m.ref === ref) };
    a.timer = setTimeout(() => { const i = asking.indexOf(a); if (i >= 0) asking.splice(i, 1); resolve({ error: 'no-answer' }); }, waitMs);
    a.timer.unref?.();
    asking.push(a);
    try {
      send({ ...msg, ref });
    } catch {
      asking.splice(asking.indexOf(a), 1);
      clearTimeout(a.timer);
      resolve({ error: 'offline' });
    }
  });
  const leaveAsking = () => { for (const a of asking.splice(0)) { clearTimeout(a.timer); a.resolve({ error: 'offline' }); } };

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

  /* ── an invitation from the person's own address ────────────────── */

  /** The invitation emailed from one of the person's Mail accounts: how it went. */
  const mailFromOwn = async ({ accountId, to, link, letter }) => {
    const account = (await Promise.resolve(mail?.accounts?.() || [])).find((a) => a.id === accountId);
    // An imported archive has no server to send through.
    if (!account || account.local || !mail?.send) return { by: 'own', sent: false, reason: 'no-account' };
    if (!link) return { by: 'own', from: account.email, sent: false, reason: 'no-link' };
    const join = `${DOWNLOAD_URL}?invite=${encodeURIComponent(inviteCode(link) || '')}`;
    const fill = (text) => String(text || '').replaceAll('{accept}', () => link).replaceAll('{join}', () => join);
    const subject = fill(letter?.subject) || 'Talk with me on Rutba Talk';
    const text = fill(letter?.text) || `I would like to talk with you on Rutba Talk.\n\nAccept: ${link}\n\nNo Rutba Office yet? Accept with a free Rutba account on that page, then get Rutba Office: ${join}\n`;
    try {
      await mail.send({ accountId, draft: { to, subject, text, html: letterHtml(text) } });
      return { by: 'own', from: account.email, sent: true };
    } catch (err) {
      return { by: 'own', from: account.email, sent: false, reason: 'failed', detail: err?.message || String(err) };
    }
  };

  /* ── the connection ──────────────────────────────────────────────── */

  const handle = (msg) => {
    switch (msg.type) {
      case 'welcome':
        retry = 0;
        state = { ...state, phase: 'online', me: msg.me, contacts: msg.contacts || [], requests: msg.requests || { incoming: [], outgoing: [] }, blocked: msg.blocked || [], relays: msg.relays || [], invites: msg.invites || [], mailAllowance: msg.mailAllowance || null, error: '' };
        break;
      case 'mail.allowance':
        state = { ...state, mailAllowance: msg.allowance || null };
        break;
      case 'invites':
        state = { ...state, invites: msg.invites || [] };
        break;
      case 'invite.created':
        if (msg.invite) state = { ...state, invites: [...state.invites.filter((i) => i.code !== msg.invite.code), msg.invite] };
        break;
      case 'invite.revoked':
        state = { ...state, invites: state.invites.filter((i) => i.code !== msg.code) };
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
        if (msg.mail?.allowance) state = { ...state, mailAllowance: msg.mail.allowance };
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
        // A token that has run out is renewed and the connection made again; any other refusal signs out.
        if (msg.for === 'hello' && msg.code === 'expired' && secrets.get(REFRESH_SECRET)) expired = true;
        else if (msg.for === 'hello') state = { ...state, phase: 'signed-out', error: msg.message || 'The sign-in was not accepted.' };
        break;
      default:
        break;
    }
    answerOf(msg);
    // Every window hears every message: a call's set-up goes to the window that holds the call.
    broadcast?.('talk:event', msg);
    tell();
  };

  const send = (msg) => {
    if (!socket || socket.readyState !== 1) throw new Error('Talk is not connected.');
    socket.send(JSON.stringify(msg));
  };

  /* ── the sign-in's tokens ────────────────────────────────────────── */

  const issuer = () => settings.get(ISSUER_KEY, ISSUER);
  const keepTokens = (t) => {
    secrets.set(TOKEN_SECRET, String(t.access_token));
    if (t.refresh_token) secrets.set(REFRESH_SECRET, String(t.refresh_token));
  };
  const forget = () => {
    secrets.delete(TOKEN_SECRET);
    secrets.delete(REFRESH_SECRET);
  };
  let renewing = null;
  /**
   * A fresh access token when it is near its end, or when the introducer
   * said it had run out. 'fresh' or 'kept' when there is one to connect
   * with; 'ended' when the sign-in has ended (signed out at auth.rutba.io,
   * by an administrator, or 30 days gone), and the tokens are forgotten.
   */
  const renew = (force = false) => {
    const refresh = secrets.get(REFRESH_SECRET);
    if (!refresh) return Promise.resolve('kept');
    const until = expiresAt(secrets.get(TOKEN_SECRET));
    if (!force && until && until - Date.now() > RENEW_BEFORE) return Promise.resolve('kept');
    renewing ||= refreshTokens({ issuer: issuer(), refreshToken: refresh, fetchImpl })
      .then((t) => { keepTokens(t); return 'fresh'; })
      .catch((err) => {
        if (err.code === 'invalid_grant') { forget(); return 'ended'; }
        // Out of reach for now: the person is still signed in; the next try renews it.
        return 'kept';
      })
      .finally(() => { renewing = null; });
    return renewing;
  };
  let expired = false;

  const open = async () => {
    clearTimeout(retryTimer);
    if ((await renew(expired)) === 'ended') {
      expired = false;
      state = { ...state, phase: 'signed-out', error: 'Your Rutba sign-in has ended. Sign in again to use Talk.' };
      tell();
      return;
    }
    expired = false;
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
      leaveAsking();
      if (!wanted || state.phase === 'signed-out') { tell(); return; }
      state = { ...state, phase: 'offline' };
      tell();
      // Turned away for a token that ran out: again at once, with a fresh one.
      if (expired) { open(); return; }
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
      // A token given outright has no refresh of its own.
      secrets.delete(REFRESH_SECRET);
      secrets.set(TOKEN_SECRET, String(token));
      wanted = true;
      open();
      return snapshot();
    },

    /**
     * Sign in with Rutba in the browser: auth.rutba.io in the system
     * browser, the code back to Office on 127.0.0.1, then Rutba's
     * introducer. Answers once the sign-in has come back (or why not).
     */
    signInBrowser: async ({ server } = {}) => {
      if (!openBrowser) throw new Error('This copy of Rutba Office cannot open a browser.');
      state = { ...state, phase: 'signing-in', error: '' };
      tell();
      try {
        const tokens = await signInWithBrowser({ issuer: issuer(), open: openBrowser, fetchImpl });
        settings.set(SERVER_KEY, testServers && server ? server : DEFAULT_SERVER);
        keepTokens(tokens);
      } catch (err) {
        state = { ...state, phase: 'signed-out', error: err?.message || String(err) };
        tell();
        throw err;
      }
      wanted = true;
      open();
      return snapshot();
    },

    signOut: () => {
      wanted = false;
      clearTimeout(retryTimer);
      // The refresh token given back, so the sign-in cannot be renewed from this copy again.
      const refresh = secrets.get(REFRESH_SECRET);
      if (refresh) revoke({ issuer: issuer(), refreshToken: refresh, fetchImpl }).catch(() => {});
      forget();
      const ws = socket;
      socket = null;
      try { ws?.close(1000, 'signed out'); } catch { /* gone */ }
      state = { phase: 'signed-out', me: null, contacts: [], requests: { incoming: [], outgoing: [] }, blocked: [], relays: [], error: '' };
      tell();
      return snapshot();
    },

    /** Connect again with the sign-in kept, as at start-up. */
    resume: () => {
      if (secrets.get(TOKEN_SECRET) || secrets.get(REFRESH_SECRET)) { wanted = true; if (!socket) open(); }
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

    /**
     * Ask an address to be in contact, and tell them: `how` is 'own' (an
     * email from the person's Mail account `accountId`, its words `letter`
     * with {accept} and {join} where the links go), 'rutba' (Rutba emails
     * them, within its limits) or 'none' (they see it in Talk if they use
     * it). Answers { request, told } or { contact } when they were one
     * already, or { error }.
     */
    invite: async ({ email, note = '', how = 'none', accountId = null, letter = null } = {}) => {
      const address = String(email || '').trim();
      const r = await ask({ type: 'contact.request', email: address, note: String(note || ''), notify: how === 'rutba' ? 'rutba' : 'none' }, 'contact.request.sent', null, how === 'rutba' ? ASK_MAIL_MS : ASK_MS);
      if (r.error) return { error: r.error };
      if (r.contact) return { contact: r.contact };
      const request = r.request;
      let told = { by: 'none', sent: false };
      if (how === 'rutba') {
        // An introducer from before invitation emails says nothing of mail.
        told = r.mail ? { by: 'rutba', sent: Boolean(r.mail.sent), reason: r.mail.reason || null } : { by: 'rutba', sent: false, reason: 'unsupported' };
      } else if (how === 'own') {
        told = await mailFromOwn({ accountId, to: address, link: request.link, letter });
      }
      keepTold(request.id, told);
      tell();
      return { request: { ...request, told }, told, allowance: state.mailAllowance };
    },

    /** How many emails Rutba will still send for this person, today and this week. */
    allowance: async () => {
      const r = await ask({ type: 'mail.allowance' }, 'mail.allowance');
      return r.error ? { error: r.error } : r.allowance;
    },

    /** An invitation link to send over any channel: { invite } or { error }. */
    inviteLink: async ({ note = '' } = {}) => {
      const r = await ask({ type: 'invite.create', note: String(note || '').slice(0, 300) }, 'invite.created');
      return r.error ? { error: r.error } : { invite: r.invite };
    },

    revokeInvite: async ({ code } = {}) => {
      const r = await ask({ type: 'invite.revoke', code: String(code || '') }, 'invite.revoked', (m) => m.code === String(code || ''));
      return r.error ? { error: r.error } : { revoked: true };
    },

    /** Who an invitation (a link, or its code) is from, before it is accepted. */
    peekInvite: async ({ link } = {}) => {
      const code = inviteCode(link);
      if (!code) return { error: 'no-invite' };
      const r = await ask({ type: 'invite.peek', code }, 'invite.peeked');
      return r.error ? { error: r.error, code } : { code, from: r.from, note: r.note || '', kind: r.kind };
    },

    /** Accept an invitation: the two are contacts. { contact } or { error }. */
    acceptInvite: async ({ link } = {}) => {
      const code = inviteCode(link);
      if (!code) return { error: 'no-invite' };
      // The introducer names the code it accepted; one from before it did is taken as this one.
      const r = await ask({ type: 'invite.accept', code }, 'contact.added', (m) => m.code == null || m.code === code);
      return r.error ? { error: r.error } : { contact: r.contact };
    },

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

/**
 * An invitation's code from what someone has: the page's address
 * (https://introducer.rutba.io/i/<code>, or a test introducer's), the
 * rutba-office:// link, or the code itself.
 */
export function inviteCode(link) {
  const text = String(link || '').trim();
  const m = /^(?:https?:\/\/[^/\s]+\/i\/|rutba-office:\/\/talk\/invite\/)?([A-Za-z0-9_-]{8,128})\/?(?:[?#].*)?$/.exec(text);
  return m ? m[1] : null;
}

const escapeHtml = (t) => String(t).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);

/** The letter as HTML too, its addresses as links: a mail of a bare link and nothing else reads as spam. */
export function letterHtml(text) {
  const paragraphs = String(text || '').split(/\n{2,}/).map((p) => escapeHtml(p).replace(/(https?:\/\/[^\s<]+)/g, '<a href="$1">$1</a>').replace(/\n/g, '<br>'));
  return `<div style="font-family:Segoe UI,Arial,sans-serif;font-size:15px;line-height:1.5">${paragraphs.map((p) => `<p>${p}</p>`).join('')}</div>`;
}
