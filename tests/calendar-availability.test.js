// Availability on request: free/busy from the server where it keeps it (a
// CalDAV scheduling outbox, Microsoft Graph's getSchedule, Google's
// freeBusy), else a VFREEBUSY enquiry by mail whose answer, when it comes,
// is what the editor shows; a room run here answers from its own bookings.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import http from 'node:http';
import { readScheduling, writeFreeBusyReply } from '../packages/calendar/src/index.js';
import { createCalendarService } from '../apps/desktop/main/calendar.js';
import { createMailLink } from '../apps/desktop/main/calendar-mail.js';
import { createDavClient } from '../apps/desktop/main/dav.js';
import { createDavService } from '../apps/desktop/main/dav-sync.js';
import { graphSchedule, googleFreeBusy } from '../apps/desktop/main/calendar-availability.js';
import { startDavServer } from './fixtures/dav-server.js';

const H = 3600000;
const start = Date.UTC(2026, 10, 20, 9);
const end = start + 8 * H;

/** A server on this computer answering one path with what `answer` makes of the request. */
async function serve(answer) {
  const seen = [];
  const server = http.createServer((req, res) => {
    let body = '';
    req.on('data', (d) => { body += d; });
    req.on('end', () => {
      seen.push({ method: req.method, url: req.url, headers: req.headers, body: body ? JSON.parse(body) : null });
      const out = answer(seen.at(-1));
      res.writeHead(out.status || 200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify(out.json));
    });
  });
  await new Promise((r) => server.listen(0, '127.0.0.1', r));
  return { base: `http://127.0.0.1:${server.address().port}`, seen, close: () => new Promise((r) => server.close(r)) };
}

test("a CalDAV server's scheduling outbox answers for the people it knows", async () => {
  const dav = await startDavServer({ busy: { 'kim@dav.example': [{ start: start + H, end: start + 2 * H }] } });
  try {
    const client = createDavClient({ url: dav.url, user: 'ann', password: 'secret' });
    const got = await client.freeBusy({ organizer: { email: 'ann@dav.example' }, attendees: ['kim@dav.example', 'nobody@dav.example'], start, end });
    assert.deepEqual(got.map((g) => [g.email, g.ok, g.busy.map((b) => [b.start, b.end])]), [['kim@dav.example', true, [[start + H, start + 2 * H]]], ['nobody@dav.example', false, []]]);
    assert.ok(dav.log.some((r) => r.method === 'POST' && r.path === '/dav/outbox/ann/'));
  } finally {
    await dav.close();
  }
});

test("Graph's getSchedule and Google's freeBusy are asked with the sign-in, and read back as busy times", async () => {
  const graph = await serve(() => ({ json: { value: [
    { scheduleId: 'Kim@Contoso.example', scheduleItems: [{ status: 'busy', start: { dateTime: '2026-11-20T10:00:00.0000000', timeZone: 'UTC' }, end: { dateTime: '2026-11-20T11:00:00.0000000', timeZone: 'UTC' } }, { status: 'free', start: { dateTime: '2026-11-20T12:00:00', timeZone: 'UTC' }, end: { dateTime: '2026-11-20T13:00:00', timeZone: 'UTC' } }] },
    { scheduleId: 'gone@contoso.example', error: { message: 'not found' } },
  ] } }));
  const google = await serve(() => ({ json: { calendars: { 'sam@gmail.example': { busy: [{ start: '2026-11-20T14:00:00Z', end: '2026-11-20T15:30:00Z' }] }, 'x@gmail.example': { errors: [{ reason: 'notFound' }] } } } }));
  try {
    const g = await graphSchedule({ token: 'tok-graph', emails: ['kim@contoso.example', 'gone@contoso.example'], start, end, base: graph.base });
    assert.equal(graph.seen[0].url, '/me/calendar/getSchedule');
    assert.equal(graph.seen[0].headers.authorization, 'Bearer tok-graph');
    assert.deepEqual(graph.seen[0].body.schedules, ['kim@contoso.example', 'gone@contoso.example']);
    assert.deepEqual(g.map((x) => [x.email, x.ok, x.busy.map((b) => [b.start, b.end])]), [['kim@contoso.example', true, [[start + H, start + 2 * H]]], ['gone@contoso.example', false, []]]);

    const o = await googleFreeBusy({ token: 'tok-google', emails: ['sam@gmail.example', 'x@gmail.example'], start, end, base: google.base });
    assert.equal(google.seen[0].url, '/freeBusy');
    assert.deepEqual(google.seen[0].body.items, [{ id: 'sam@gmail.example' }, { id: 'x@gmail.example' }]);
    assert.equal(google.seen[0].body.timeMin, '2026-11-20T09:00:00Z');
    assert.deepEqual(o.map((x) => [x.email, x.ok, x.busy.map((b) => [b.start, b.end])]), [['sam@gmail.example', true, [[start + 5 * H, start + 6.5 * H]]], ['x@gmail.example', false, []]]);
  } finally {
    await graph.close();
    await google.close();
  }
});

function fakeMail() {
  const sent = [];
  const inbox = [];
  const service = {
    accounts: () => [{ id: 'me', email: 'ann@dav.example', name: 'Ann' }],
    folders: () => [{ path: 'INBOX', role: 'inbox' }],
    messages: ({ limit }) => ({ rows: inbox.slice(-limit).reverse() }),
    message: ({ id }) => inbox.find((m) => m.id === id)?.message || null,
    attachmentText: ({ id }) => inbox.find((m) => m.id === id)?.part || null,
    send: async () => ({ refused: [] }),
    sendFrom: async (msg) => { sent.push(msg); return { refused: [] }; },
  };
  const deliver = (from, text) => inbox.push({ id: `in-${inbox.length + 1}`, from: { address: from }, date: new Date().toISOString(), hasAttachments: true, attachmentCount: 1, message: { attachments: [{ type: 'text/calendar', filename: 'reply.ics' }] }, part: { name: 'reply.ics', type: 'text/calendar', text } });
  return { service, sent, deliver };
}

function memory() {
  const map = new Map();
  return { get: (k) => map.get(k), set: (k, v) => { map.set(k, v); }, delete: (k) => { map.delete(k); } };
}

test('availability: the CalDAV server first, then a sign-in that reads free/busy, else asked by mail and answered later', async (t) => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-avail-'));
  t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
  const dav = await startDavServer({ busy: { 'kim@dav.example': [{ start: start + H, end: start + 2 * H }] } });
  const graph = await serve(({ body }) => ({ json: { value: body.schedules.filter((s) => s === 'lee@contoso.example').map((s) => ({ scheduleId: s, scheduleItems: [] })) } }));
  t.after(() => Promise.all([dav.close(), graph.close()]));
  const mail = fakeMail();
  const stores = { dir, settings: memory(), secrets: memory() };
  const calendar = createCalendarService({ stores, broadcast: null, mail: createMailLink(() => mail.service), freeBusyTokens: (account) => (account.address === 'ann@dav.example' ? { kind: 'graph', token: 'tok', base: graph.base } : null) });
  const sync = createDavService({ stores, calendar, contacts: null, timers: false });
  await sync.add({ url: dav.url, user: 'ann', password: 'secret' });

  const first = await calendar.availability({ emails: ['kim@dav.example', 'lee@contoso.example', 'sam@far.example'], start: start + H, end: start + 2 * H });
  const by = Object.fromEntries(first.rows.map((r) => [r.email, r]));
  assert.equal(by['kim@dav.example'].free, false);
  assert.equal(by['kim@dav.example'].source, 'caldav');
  assert.equal(by['lee@contoso.example'].free, true);
  assert.equal(by['lee@contoso.example'].source, 'graph');
  assert.equal(by['sam@far.example'].free, null);
  assert.equal(by['sam@far.example'].asked, true);
  assert.equal(first.from, 'ann@dav.example');

  // The one asked by mail: a VFREEBUSY REQUEST from Ann's account, to Sam alone.
  assert.equal(mail.sent.length, 1);
  const ask = mail.sent[0];
  assert.deepEqual(ask.to, ['sam@far.example']);
  assert.match(ask.subject, /^When are you free\?/);
  const request = readScheduling(ask.attachments[0].content);
  assert.equal(request.method, 'REQUEST');
  assert.equal(request.freebusy[0].attendees[0].email, 'sam@far.example');
  // Asked again soon after: answered from what is kept, Sam not asked twice.
  await calendar.availability({ emails: ['kim@dav.example', 'sam@far.example'], start: start + H, end: start + 2 * H });
  assert.equal(mail.sent.length, 1);
  assert.equal(dav.log.filter((r) => r.method === 'POST').length, 2, 'the server asked for Sam the second time, not for Kim');

  // Sam's calendar answers; the editor, asking again without asking anyone, has it.
  const fb = request.freebusy[0];
  mail.deliver('sam@far.example', writeFreeBusyReply({ uid: fb.uid, organizer: fb.organizer, attendee: { email: 'sam@far.example' }, start: fb.start, end: fb.end, busy: [{ start: start + 1.5 * H, end: start + 3 * H }] }));
  const taken = await calendar.checkMail();
  assert.equal(taken[0]?.kind, 'freebusy');
  const later = await calendar.availability({ emails: ['sam@far.example'], start: start + H, end: start + 2 * H, ask: false });
  assert.equal(later.rows[0].free, false);
  assert.equal(later.rows[0].source, 'email');
  assert.equal(mail.sent.length, 1);
  // An answer nobody asked for is not taken.
  mail.deliver('eve@far.example', writeFreeBusyReply({ uid: 'other', organizer: { email: 'ann@dav.example' }, attendee: { email: 'eve@far.example' }, start, end, busy: [] }));
  assert.deepEqual(await calendar.checkMail(), []);
});
