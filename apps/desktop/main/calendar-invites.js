// Meeting invitations by email: the calendar's side of iTIP over iMIP.
//
// The organizer's Send mails a REQUEST from the account the meeting is
// organized from, through Mail (calendar-mail.js); an edit sent again is an
// update with the SEQUENCE the edit raised, and the people taken off the list
// get a CANCEL of their own; a meeting deleted, or one occurrence of it,
// sends a CANCEL. Answers come back as REPLYs, read from the Inbox as they
// arrive (scan, below) or when the message is opened, and each one goes onto
// that person's line. A person invited answers with Accept, Tentative or
// Decline, and the REPLY goes from the address they were invited at.
//
// What was last sent is kept on the event (`sentAs`: the time, the place and
// the rule; `sentTo`: the addresses), to tell an update that asks everyone
// again from one that does not, and to know who was taken off the list.

import fs from 'node:fs';
import path from 'node:path';
import { writeRequest, writeCancel, readScheduling, applyReply, significantChange, sameAddress, dateTimeAt } from '@rutba/calendar';
import { wordsFrom, inviteText, subjectOf, answerText, sendFailure } from './calendar-text.js';

const lower = (s) => String(s || '').trim().toLowerCase();
const DAY = 86400000;

/**
 * @param {object} o
 * @param {object} o.stores the profile's stores (`dir`)
 * @param {() => object} o.state the calendar's state
 * @param {(id: string) => ({ cal, event } | null)} o.find
 * @param {() => void} o.save
 * @param {(cal: object, uid: string) => void} o.touch
 * @param {object} o.mail the Mail link (calendar-mail.js)
 * @param {() => string[]} o.myAddresses
 */
export function createInvites({ stores, state, find, save, touch, mail, myAddresses }) {
  /** The words last handed over by a window, kept for messages sent with none open. */
  const wordsNow = (given) => {
    if (given && typeof given === 'object') state().words = wordsFrom(given);
    return wordsFrom(state().words);
  };
  const mine = (address) => Boolean(address) && myAddresses().includes(lower(address));
  /** Sent through Mail, a failure said plainly in the person's words. */
  const send = async (message, words, address = null) => {
    try {
      return await mail.sendFrom(message);
    } catch (err) {
      throw sendFailure(err, words, address || message.account);
    }
  };

  /** A calendar file for a message, kept a day in the profile's outgoing folder for Mail to attach. */
  const outgoing = (kind, text) => {
    const dir = path.join(stores.dir, 'outgoing');
    fs.mkdirSync(dir, { recursive: true });
    try {
      for (const name of fs.readdirSync(dir)) {
        const full = path.join(dir, name);
        if (Date.now() - fs.statSync(full).mtimeMs > DAY) fs.rmSync(full, { force: true });
      }
    } catch {
      /* an old file still in use is removed next time */
    }
    const file = path.join(dir, `${kind}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}.ics`);
    fs.writeFileSync(file, text, 'utf8');
    return file;
  };
  const attachment = (method, text, file) => ({
    filename: { REPLY: 'reply.ics', CANCEL: 'cancel.ics' }[method] || 'invite.ics',
    content: text,
    contentType: `text/calendar; method=${method}; charset=utf-8`,
    path: file,
  });

  /** The event and the exceptions of its series: what an invitation carries. */
  const seriesOf = (cal, uid) => cal.events.filter((e) => e.uid && e.uid === uid);
  const masterOf = (cal, event) => (event.recurrenceId ? cal.events.find((e) => e.uid === event.uid && !e.recurrenceId) || event : event);

  /** Every event with this UID, across the calendars, the series' own first. */
  const byUid = (uid) => {
    const out = [];
    for (const cal of state().calendars) for (const e of cal.events) if (e.uid && e.uid === uid) out.push({ cal, event: e });
    return out.sort((a, b) => Number(Boolean(a.event.recurrenceId)) - Number(Boolean(b.event.recurrenceId)));
  };

  /** The organizer a meeting goes from: its own, else the account new mail goes from. */
  const organizerOf = (event) => {
    if (event.organizer?.email) return event.organizer;
    const account = mail.defaultAccount();
    return account ? { email: account.address, name: account.name || null } : null;
  };

  const invited = (event, organizer) => (event.attendees || []).filter((a) => a.email && !sameAddress(a.email, organizer?.email));

  /** See takeMail below; named so scan can call it. */
  let takeMail = null;

  const api = {
    /** The account an invitation would go from, or null when there is none to send from. */
    sendingAccount: () => {
      let a = mail.defaultAccount();
      // Never a room's mailbox run here (calendar-desk.js): the person's own first address instead.
      if (a && !mine(a.address)) a = myAddresses().map((e) => mail.accountFor(e)).find(Boolean) || null;
      return a ? { id: a.id, address: a.address, name: a.name } : null;
    },

    /**
     * Send the invitation for an event, or the change to one already sent.
     * Answers { to, from, update, significant, removed, refused }.
     */
    sendInvitation: async ({ id, words: given = null }) => {
      const words = wordsNow(given);
      const hit = find(id);
      if (!hit) throw new Error('That event is no longer there.');
      const master = masterOf(hit.cal, hit.event);
      if (master.organizer?.email && !mine(master.organizer.email)) throw new Error('Only the organizer sends the changes to this meeting.');
      const organizer = organizerOf(master);
      if (!organizer?.email) throw new Error('There is no Mail account to send the invitation from.');
      const people = invited(master, organizer);
      if (!people.length) throw new Error('No one is invited.');
      const account = mail.accountFor(organizer.email) || mail.defaultAccount();
      const update = Boolean(master.sentAs);
      const significant = update && significantChange(master.sentAs, master);
      // A new time or place asks everyone again.
      if (significant) master.attendees = master.attendees.map((a) => (sameAddress(a.email, organizer.email) ? a : { ...a, partstat: 'NEEDS-ACTION', rsvp: true, answeredAt: undefined }));
      master.organizer = { email: organizer.email, name: organizer.name || null };
      const text = writeRequest(seriesOf(hit.cal, master.uid), { organizer: master.organizer });
      const sent = await send({
        account: account?.id || null,
        to: people.map((a) => a.email),
        subject: subjectOf(update ? 'UPDATE' : 'REQUEST', master, words),
        text: inviteText(master, words, { update }),
        attachments: [attachment('REQUEST', text, outgoing('invite', text))],
      }, words, account?.address);
      // Taken off the list: told it is off for them.
      const removed = (master.sentTo || []).filter((e) => !people.some((a) => sameAddress(a.email, e)));
      if (removed.length) {
        const cancel = writeCancel([master], { organizer: master.organizer, attendees: removed.map((email) => ({ email })) });
        await send({ account: account?.id || null, to: removed, subject: subjectOf('CANCEL', master, words), text: inviteText(master, words, { method: 'CANCEL' }), attachments: [attachment('CANCEL', cancel, outgoing('cancel', cancel))] }, words, account?.address);
      }
      master.sentAs = { start: master.start, end: master.end, allDay: Boolean(master.allDay), rrule: master.rrule || null, location: master.location || '' };
      master.sentTo = people.map((a) => a.email);
      master.sentSequence = master.sequence || 0;
      touch(hit.cal, master.uid);
      save();
      return { to: master.sentTo, from: account?.address || organizer.email, update, significant, removed, refused: sent?.refused || [] };
    },

    /**
     * Before a meeting this person organizes is deleted: a CANCEL to the
     * people invited, for the whole of it or (`scope` 'this') the one
     * occurrence at `original`. Answers { sent, to } ({ sent: false } for an
     * event with nobody to tell, or one someone else organizes).
     */
    sendCancel: async ({ id, scope = 'all', original = null, words: given = null }) => {
      const words = wordsNow(given);
      const hit = find(id);
      if (!hit) return { sent: false, to: [] };
      const master = masterOf(hit.cal, hit.event);
      const organizer = master.organizer?.email ? master.organizer : null;
      if (!organizer || !mine(organizer.email) || !master.sentTo?.length) return { sent: false, to: [] };
      const people = invited(master, organizer);
      if (!people.length) return { sent: false, to: [] };
      master.sequence = (master.sequence || 0) + 1;
      let events;
      if (scope === 'this' && master.rrule && original != null) {
        const length = (master.end?.at ?? master.start.at) - master.start.at;
        const tz = master.start.tzid || null;
        const at = (ms) => dateTimeAt(ms, tz, master.allDay);
        events = [{ ...master, rrule: null, exdates: [], rdates: [], recurrenceId: at(original), start: at(original), end: at(original + length) }];
      } else events = seriesOf(hit.cal, master.uid).map((e) => ({ ...e, sequence: master.sequence }));
      const text = writeCancel(events, { organizer });
      const account = mail.accountFor(organizer.email) || mail.defaultAccount();
      const shown = { ...master, ...(scope === 'this' ? { start: events[0].start, end: events[0].end, rrule: null } : {}) };
      const sent = await send({
        account: account?.id || null,
        to: people.map((a) => a.email),
        subject: subjectOf('CANCEL', shown, words),
        text: inviteText(shown, words, { method: 'CANCEL' }),
        attachments: [attachment('CANCEL', text, outgoing('cancel', text))],
      }, words, account?.address);
      touch(hit.cal, master.uid);
      save();
      return { sent: true, to: people.map((a) => a.email), refused: sent?.refused || [] };
    },

    /**
     * An answer, sent: the REPLY `respond` wrote goes to the organizer from
     * the account this person was invited at. Answers what `respond` did,
     * with `sent` and `from`; `sent: false` with `reason` when there is no
     * account of theirs to send it from (the window then opens it in Mail).
     */
    sendAnswer: async (r, { partstat, comment = null, words: given = null, summary = '' }) => {
      const words = wordsNow(given);
      if (!r.replyPath || !r.organizer?.email) return { ...r, sent: false, reason: 'no-organizer' };
      const account = (r.me && mail.accountFor(r.me)) || null;
      if (!account) return { ...r, sent: false, reason: 'no-account' };
      const text = fs.readFileSync(r.replyPath, 'utf8');
      const subject = subjectOf(partstat, { summary }, words);
      const body = [answerText(partstat, account.name || account.address, words), comment ? `\n\n${comment}` : ''].join('');
      // From the address invited: Mail takes an account by its address.
      const sent = await send({ account: account.address, to: [r.organizer.email], subject, text: body, attachments: [attachment('REPLY', text, r.replyPath)] }, words, account.address);
      return { ...r, subject, sent: true, from: account.address, refused: sent?.refused || [] };
    },

    /**
     * A scheduling message that arrived, taken into the calendar: a REPLY onto
     * the organizer's event, a CANCEL onto the copy kept of a meeting, a
     * REQUEST that changes a meeting already in the calendar onto it. `from`
     * is who sent it: a REPLY is taken only for the people on the list or
     * from that address, a CANCEL or a change only from the organizer.
     * Answers what changed ({ kind, eventId, summary, answers? }) or null.
     */
    takeMail: ({ text, from = null }) => {
      let read;
      try {
        read = readScheduling(text);
      } catch {
        return null;
      }
      const incoming = read.events[0];
      if (!incoming?.uid) return null;
      const hits = byUid(incoming.uid);
      const exact = (e) => hits.find((h) => (e.recurrenceId ? h.event.recurrenceId && Math.abs(h.event.recurrenceId.at - e.recurrenceId.at) < 1000 : !h.event.recurrenceId));
      if (read.method === 'REPLY') {
        const hit = exact(incoming) || hits[0];
        if (!hit) return null;
        const organizer = hit.event.organizer?.email || masterOf(hit.cal, hit.event).organizer?.email;
        if (organizer && !mine(organizer)) return null;
        const lines = incoming.attendees.filter((a) => !from || sameAddress(a.email, from) || (hit.event.attendees || []).some((x) => sameAddress(x.email, a.email)));
        const r = applyReply(hit.event, { ...incoming, attendees: lines });
        if (!r.changed) return { kind: 'reply', eventId: hit.event.id, summary: hit.event.summary, answers: [], already: true };
        hit.event.attendees = r.attendees;
        hit.event.updatedAt = Date.now();
        touch(hit.cal, hit.event.uid);
        save();
        return { kind: 'reply', eventId: hit.event.id, summary: hit.event.summary, answers: r.answers };
      }
      if (!hits.length) return null;
      const master = hits[0];
      const organizer = master.event.organizer?.email;
      if (organizer && mine(organizer)) return null; // this person's own meeting: its changes are made here
      if (from && organizer && !sameAddress(from, organizer)) return null;
      if (read.method === 'CANCEL') {
        if (incoming.recurrenceId && master.event.rrule) {
          const at = incoming.recurrenceId.at;
          if ((master.event.exdates || []).some((x) => Math.abs(x.at - at) < 1000)) return { kind: 'cancel', eventId: master.event.id, summary: master.event.summary, already: true };
          master.event.exdates = [...(master.event.exdates || []), incoming.recurrenceId];
          master.cal.events = master.cal.events.filter((e) => !(e.uid === incoming.uid && e.recurrenceId && Math.abs(e.recurrenceId.at - at) < 1000));
        } else {
          if (master.event.status === 'CANCELLED') return { kind: 'cancel', eventId: master.event.id, summary: master.event.summary, already: true };
          master.event.status = 'CANCELLED';
        }
        master.event.updatedAt = Date.now();
        touch(master.cal, master.event.uid);
        save();
        return { kind: 'cancel', eventId: master.event.id, summary: master.event.summary };
      }
      if (read.method === 'REQUEST') {
        const hit = exact(incoming);
        if (!hit) return null;
        const newer = (incoming.sequence || 0) > (hit.event.sequence || 0)
          || ((incoming.sequence || 0) === (hit.event.sequence || 0) && (incoming.dtstamp?.at || 0) > (hit.event.dtstamp?.at || 0) && significantChange(hit.event, incoming));
        if (!newer) return null;
        const significant = significantChange(hit.event, incoming);
        const kept = myAddresses();
        // This person's own answer stays, unless the meeting moved: then it asks again.
        const attendees = incoming.attendees.map((a) => {
          if (!kept.includes(lower(a.email))) return a;
          const was = (hit.event.attendees || []).find((x) => sameAddress(x.email, a.email));
          return significant || !was ? { ...a, partstat: 'NEEDS-ACTION' } : { ...a, partstat: was.partstat };
        });
        Object.assign(hit.event, { ...incoming, attendees, id: hit.event.id, alarms: hit.event.alarms, remote: hit.event.remote, status: incoming.status || 'CONFIRMED', updatedAt: Date.now() });
        touch(hit.cal, hit.event.uid);
        save();
        return { kind: 'update', eventId: hit.event.id, summary: hit.event.summary, significant };
      }
      return null;
    },

    /**
     * The Inboxes read for what arrived: each message not looked at before
     * that carries a calendar part is taken in (takeMail), and `also` sees
     * every new one (a room's mailbox, calendar-desk.js); the accounts in
     * `rooms` are a room's, their calendar parts left to `also`. Answers what
     * changed, each with the account and message it came from.
     */
    scan: async ({ also = null, rooms = [] } = {}) => {
      const st = state();
      st.mailSeen ||= {};
      const results = [];
      let looked = false;
      for (const a of mail.accounts()) {
        const folder = mail.inboxOf(a.id);
        if (!folder) continue;
        const seen = new Set(st.mailSeen[a.id] || []);
        const fresh = mail.latest({ accountId: a.id, folder, limit: 60 }).filter((row) => !seen.has(row.id));
        for (const row of fresh.reverse()) {
          seen.add(row.id);
          looked = true;
          const { part } = row.hasAttachments || row.attachmentCount ? mail.calendarPart({ accountId: a.id, folder, id: row.id }) : { part: null };
          const from = row.from?.address || null;
          let r = null;
          // A room's Inbox is answered for the room (`also`), not taken into this person's calendar.
          if (part?.text && !rooms.includes(a.id)) {
            try {
              r = takeMail({ text: part.text, from });
            } catch {
              r = null;
            }
          }
          if (r) results.push({ ...r, accountId: a.id, messageId: row.id });
          let more = null;
          try {
            more = (await also?.({ account: a, folder, row, part, taken: r })) || null;
          } catch (err) {
            more = { kind: 'failed', reason: err?.message || String(err) };
          }
          if (more) results.push({ ...more, accountId: a.id, messageId: row.id });
        }
        if (fresh.length) st.mailSeen[a.id] = [...seen].slice(-1000);
      }
      if (looked) save();
      return results;
    },
  };
  takeMail = api.takeMail;
  /** A calendar part for Mail to attach: the file kept a day, and the text. */
  api.attach = (method, text) => attachment(method, text, outgoing(method.toLowerCase(), text));
  return api;
}
