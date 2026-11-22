// The Talk service on a stand-in introducer, for tests that need no real
// one: what the service sends is kept in `out`, what it tells its windows in
// `told`, and the test says what the introducer says with `say`.
//
// Nothing here waits for a length of time. `online()` waits for the service
// to tell its windows it is connected; `tick()` says a mark of its own after
// everything already on its way and waits to hear it back, again and again
// until nothing more was sent meanwhile.

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createTalkService } from '../apps/desktop/main/talk.js';

export const AMIRA = { id: 'u-a', name: 'Amira Khan', email: 'amira@example.com' };
export const BEN = { id: 'u-ben', name: 'Ben Ortiz', email: 'ben@example.com', presence: 'online' };
export const CARA = { id: 'u-cara', name: 'Cara Lund', email: 'cara@example.com', presence: 'offline' };

const map = (m) => ({ get: (k, d) => (m.has(k) ? m.get(k) : d), set: (k, v) => m.set(k, v), delete: (k) => m.delete(k) });

/**
 * @param {object} [o]
 * @param {(msg: object) => object|object[]|null} [o.answer] what the introducer answers to each thing sent
 * @param {object} [o.welcome] more for the welcome (blocked, blockedPeople, ...)
 * @param {object[]} [o.after] said once the welcome has gone
 * @param {Map} [o.settings] the person's settings, to keep between two services
 * @param {boolean} [o.signedIn] whether a sign-in is kept
 */
export function standIn({ answer = () => null, welcome = {}, after = [], contacts = [BEN, CARA], settings = new Map(), signedIn = true, dataDir = null } = {}) {
  const out = [];
  const told = [];
  const dir = dataDir || fs.mkdtempSync(path.join(os.tmpdir(), 'talk-stand-in-'));
  let ws = null;
  let connections = 0;
  const say = (m) => setImmediate(() => ws.onmessage({ data: JSON.stringify(m) }));
  const watchers = new Set();
  const until = (seen) => new Promise((resolve) => {
    const w = (t) => { if (seen(t)) { watchers.delete(w); resolve(t); } };
    watchers.add(w);
  });
  let marks = 0;
  const svc = createTalkService({
    stores: { settings: map(settings), secrets: map(new Map(signedIn ? [['talk.token', 'amira']] : [])) },
    dataDir: dir,
    broadcast: (event, payload) => { const t = { event, payload }; told.push(t); for (const w of [...watchers]) w(t); },
    connect: () => {
      connections += 1;
      ws = { readyState: 1, send: (text) => {
        const msg = JSON.parse(text);
        out.push(msg);
        if (msg.type === 'hello') { say({ type: 'welcome', me: AMIRA, contacts, requests: { incoming: [], outgoing: [] }, ...welcome }); for (const m of after) say(m); return; }
        for (const reply of [].concat(answer(msg) || [])) say(reply);
      }, close: () => {} };
      setImmediate(() => ws.onopen());
      return ws;
    },
  });
  const it = {
    svc,
    out,
    told,
    dir,
    settings,
    say,
    until,
    connections: () => connections,
    /** What Talk said is worth a notification, in order. */
    notified: () => told.filter((t) => t.event === 'talk:notify').map((t) => t.payload),
    /** Connected and welcomed, however the connection was asked for. */
    async online(how = () => svc.resume()) {
      const there = until(() => svc.state().phase === 'online');
      how();
      await there;
    },
    /** Everything said so far has been heard and answered. */
    async tick() {
      let sent;
      do {
        sent = out.length;
        const mark = ++marks;
        const heard = until((t) => t.event === 'talk:event' && t.payload?.type === 'test.mark' && t.payload.mark === mark);
        say({ type: 'test.mark', mark });
        await heard;
      } while (out.length !== sent);
    },
    /** The connection is lost, as when the network goes. */
    drop() { const was = ws; was.readyState = 3; was.onclose(); },
    clean() { if (!dataDir) fs.rmSync(dir, { recursive: true, force: true }); },
  };
  return it;
}
