// Guests in Live Broadcast and the Recorder (Office 2.27.0, the Talk, Meet
// and Live plan's phase 3): the person asks up to three of their Talk
// contacts in, or sends them a link; each who comes takes a seat, which is a
// source like any other (packages/studio/src/guests.js, guest-source.js).
//
// The window holds a small Talk meeting itself. Live Broadcast's is set up
// to be streamed live and the Recorder's to be recorded, which is chosen at
// set-up and said to everyone before they join (D-2026-11-22-19): nothing of
// a guest goes out before they were told. This computer's seat sends the
// programme, so the guests see and hear it. A guest waits off air until a
// scene or the list shows them, and each is told when it is on air and
// whether they are in the picture.
//
// A meeting set up in Talk to be streamed is taken over here with
// ?meet=<code>, and goes back to Talk as a meeting with "Continue in Talk"
// (D-2026-11-22-18): the same people stay connected.
//
// Kept plain (owner, 2026-10-21: "for chats and conference call etc build
// clean simple user interface do not complicate them like Ms teams").

import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Button, Dialog, Icon, Spinner, t } from '@rutba/office-ui';
import { normaliseStack } from '@rutba/studio/stack';
import { GUEST_SCENES, SEATS } from '@rutba/studio/guests';
import { GuestSession } from './guest-session.js';
import { meetRefusal, goneText, meetingCode, liveNotice, recordedNotice } from '../apps/talk/meet-lines.js';
import { ConfirmDialog } from './parts.js';

const MOST = SEATS.length;

export const GUESTS_CSS = `
.st-guests-btn { display: inline-flex; align-items: center; gap: 6px; height: 28px; padding: 0 10px; border: 1px solid var(--line); border-radius: var(--r-2); background: transparent; color: inherit; font: inherit; font-size: 12px; cursor: pointer; }
.st-guests-btn:hover { background: var(--hover); }
.st-guests-btn.on { border-color: var(--accent); }
.st-guests-btn b { font-weight: 600; }
.st-guests p { margin: 0 0 10px; line-height: 1.5; }
.st-guests h3 { margin: 14px 0 6px; font-size: 12px; font-weight: 600; color: var(--ink-2); }
.st-guests-note { display: flex; align-items: flex-start; gap: 8px; padding: 8px 12px; border-radius: var(--r-2); border: 1px solid #d9822b; background: rgba(217,130,43,0.12); font-size: 13px; }
.st-guests-note svg { flex: none; margin-top: 2px; color: #d9822b; }
.st-guests-list { display: flex; flex-direction: column; gap: 4px; max-height: 190px; overflow: auto; }
.st-guests-row { display: flex; align-items: center; gap: 8px; padding: 6px 8px; border-radius: var(--r-2); }
.st-guests-row:hover { background: var(--hover); }
.st-guests-row > span { flex: 1; min-width: 0; display: flex; flex-direction: column; }
.st-guests-row strong { font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.st-guests-row small { color: var(--ink-2); }
.st-guests-row small.air { color: #b3261e; font-weight: 600; }
.st-guests-foot { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 14px; }
`;

/** What the names of the ready-made scenes are, in the person's language. */
const sceneNames = () => ({ [GUEST_SCENES[0]]: t('studio.guests.one-guest'), [GUEST_SCENES[1]]: t('studio.guests.two-guests'), [GUEST_SCENES[2]]: t('studio.guests.three-guests') });

/** The meeting's name, as those asked in see it. */
export function sessionName(how, me) {
  const first = String(me?.name || '').split(/\s+/)[0];
  if (how === 'live') return first ? t('studio.guests.live-broadcast-name', { name: first }) : t('studio.live-broadcast');
  return first ? t('studio.guests.recording-with-name', { name: first }) : t('common.recording');
}

/** Where a guest is (GuestSession.where), in words. */
export function whereText(where) {
  switch (where) {
    case 'on-air': return t('studio.live-panel.on-air');
    case 'off-air': return t('studio.guests.off-air');
    case 'unseated': return t('studio.guests.every-seat-taken');
    case 'failed': return t('talk.meeting.could-not-connect');
    default: return t('talk.meeting.connecting');
  }
}

/** Why a meeting could not be held here, or people not asked in, in words. */
export function guestRefusal(r) {
  if (r.error === 'not-live') return t('studio.guests.meeting-not-set-streamed');
  if (r.error === 'not-recorded') return t('studio.guests.meeting-not-set');
  return meetRefusal(r.error, r);
}

/**
 * The studio's guests. `how` is 'live' (Live Broadcast) or 'record' (the
 * Recorder); `onAir` whether the broadcast or the recording is running.
 * What is done is GuestSession's (guest-session.js); this draws it. Answers
 * what the window draws and asks: { enabled, button, dialog, addItem,
 * nameOf, note, rowMenu, seats }.
 */
export function useGuests({ shell, how, enabled, loaded, media, canvas, stack, airStack, changeStack, setAir, onAir, toast }) {
  const [, setTick] = useState(0);
  const [open, setOpen] = useState(false);
  const [busy, setBusy] = useState(false);
  const [ask, setAsk] = useState(null);
  const now = useRef({});
  now.current = { changeStack, setAir, toast, canvas };

  const session = useMemo(() => new GuestSession({
    shell,
    how,
    guests: media.guests,
    feed: async () => media.guests.feed(now.current.canvas),
    change: (fn) => {
      now.current.changeStack(fn, true, { record: false });
      now.current.setAir((a) => (a ? normaliseStack(fn(a)) : a));
    },
    say: (what, d) => {
      const words = what === 'came' ? t('studio.guests.name-waiting-off', { name: d.name || t('common.someone') }) : goneText(d.why, d.by);
      if (words) now.current.toast(words, { ms: 5000 });
    },
    onChange: () => setTick((n) => n + 1),
    sceneNames,
  }), [shell, how, media]);
  session.look({ stack, airStack, onAir });

  // Talk as it is: who is signed in, their contacts, whether it is connected.
  useEffect(() => {
    if (!enabled) return undefined;
    let gone = false;
    const take = (st) => { if (!gone) session.setTalk(st); };
    shell.talk.state().then(take).catch(() => take({ phase: 'off', contacts: [] }));
    const off = shell.on('talk:state', take);
    return () => { gone = true; off?.(); };
  }, [shell, session, enabled]);

  // What the introducer says of the meeting, as every window hears it.
  useEffect(() => (enabled ? shell.on('talk:event', (msg) => session.handle(msg)) : undefined), [shell, session, enabled]);

  // The frames as they are: who sits where, whose sound goes out, what everyone in the meeting is told.
  useEffect(() => { if (enabled) session.sync({ stack, airStack, onAir }); }, [session, enabled, stack, airStack, onAir]);

  // The window closing takes this seat out of the meeting; it goes on for the others.
  useEffect(() => {
    const out = () => session.leave();
    session.closed = false;
    window.addEventListener('pagehide', out);
    return () => { window.removeEventListener('pagehide', out); session.close(); };
  }, [session]);

  const bad = (r) => toast(guestRefusal(r), { tone: 'bad', ms: 7000 });

  // A meeting set up in Talk for this, taken over: ?meet=<code>.
  const asked = useRef(false);
  const online = session.online;
  const known = Boolean(session.talk);
  useEffect(() => {
    if (!enabled || !loaded || !known || asked.current) return;
    const code = meetingCode(new URLSearchParams(location.search).get('meet'));
    if (!code) { asked.current = true; return; }
    if (!online) return;
    asked.current = true;
    session.hold(code).then((res) => { if (res.error) bad(res); else setOpen(true); });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [enabled, loaded, known, online, session]);

  if (!enabled) return { enabled: false, button: null, dialog: null, addItem: null, nameOf: () => '', note: () => '', rowMenu: () => [], seats: {} };

  const { seats, room } = session;
  const people = session.people().map((p) => ({ id: p.id, name: p.name, muted: p.muted, connected: p.connected, where: session.where(p.id) }));

  const invite = async (ids) => {
    setBusy(true);
    const r = await session.invite(ids, { name: sessionName(how, session.talk?.me) });
    setBusy(false);
    if (r.error) bad(r);
    return !r.error;
  };
  const toTalk = () => {
    const id = session.leave();
    if (!id) return;
    setOpen(false);
    shell.win.create({ app: 'talk', query: { meet: id }, reuse: true }).catch((err) => toast(err?.message || String(err), { tone: 'bad', ms: 6000 }));
  };
  const endAll = async () => { const r = await session.end(); if (r.error) bad(r); else setOpen(false); };
  const takeOut = async (who) => { const r = await session.remove(who.id); if (r?.error) bad(r); };
  const copy = async () => {
    await shell.clipboard.writeText({ text: room?.meeting?.link || '' }).catch(() => {});
    toast(t('talk.meeting.link-copied-paste'), { tone: 'good', ms: 2500 });
  };

  /** What a seat's row says under its name. */
  const note = (item) => (seats[item.ref] ? whereText(session.where(seats[item.ref].id)) : t('studio.guests.nobody-here-yet'));

  const rowMenu = (item) => {
    const who = seats[item.ref];
    const p = who ? session.person(who.id) : null;
    return [
      ...(who ? [
        { label: t('common.mute'), icon: 'micOff', disabled: !p?.connected || p.muted, run: () => session.mute(who.id) },
        { label: t('talk.groups.take-out'), icon: 'close', run: () => setAsk({ what: 'out', who }) },
      ] : []),
      { label: t('studio.guests.guests-2'), icon: 'contacts', run: () => setOpen(true) },
    ];
  };

  const count = Object.keys(seats).length;
  const button = (
    <button type="button" className={`st-guests-btn${room ? ' on' : ''}`} data-guests={count} data-tip={t('studio.guests.guests-from-talk')} aria-label={t('studio.guests.guests-from-talk')} onClick={() => setOpen(true)}>
      <Icon name="contacts" size={15} /><span>{t('studio.guests.guests')}</span>{count ? <b>{count}</b> : null}
    </button>
  );
  const dialog = (
    <>
      {open ? <GuestsDialog how={how} talk={session.talk} holding={Boolean(room)} people={people} onAir={Boolean(onAir)} busy={busy}
        onInvite={invite} onCopy={copy} onMute={(who) => session.mute(who.id)} onTakeOut={(who) => setAsk({ what: 'out', who })} onToTalk={toTalk} onEnd={() => setAsk({ what: 'end' })}
        onOpenTalk={() => shell.win.create({ app: 'talk', reuse: true }).catch(() => {})} onClose={() => setOpen(false)} /> : null}
      {ask?.what === 'out' ? <ConfirmDialog title={t('talk.meeting.take-name-out', { name: ask.who.name || t('common.someone') })} body={t('talk.meeting.cannot-join-meeting')} okLabel={t('talk.groups.take-out')} onOk={() => takeOut(ask.who)} onClose={() => setAsk(null)} /> : null}
      {ask?.what === 'end' ? <ConfirmDialog title={t('talk.meeting.end-the-meeting-for-everyone')} body={t('talk.meeting.everyone-taken-out')} okLabel={t('talk.meeting.end-meeting')} onOk={endAll} onClose={() => setAsk(null)} /> : null}
    </>
  );
  return {
    enabled: true,
    button,
    dialog,
    addItem: { label: t('studio.guests.guest-from-talk'), icon: 'contacts', run: () => setOpen(true) },
    nameOf: (item) => seats[item.ref]?.name || '',
    note,
    rowMenu,
    seats,
  };
}

/**
 * Guests: who is here and whether they are on air, the contacts who can be
 * asked in, the link, and going back to Talk or ending it for everyone.
 * `people` is [{ id, name, muted, connected, where }] (GuestSession.where),
 * `holding` whether this window holds a meeting yet.
 */
export function GuestsDialog({ how, talk, holding, people, onAir, busy, onInvite, onCopy, onMute, onTakeOut, onToTalk, onEnd, onOpenTalk, onClose }) {
  const [chosen, setChosen] = useState([]);
  const online = talk?.phase === 'online';
  const here = new Set(people.map((p) => p.id));
  const free = Math.max(0, MOST - people.length);
  const contacts = (talk?.contacts || []).filter((c) => !here.has(c.id));
  const toggle = (id) => setChosen((c) => (c.includes(id) ? c.filter((x) => x !== id) : c.length < free ? [...c, id] : c));
  const send = async () => { if (await onInvite(chosen)) setChosen([]); };
  return (
    <Dialog
      title={t('studio.guests.guests-from-talk')}
      width={480}
      className="st-guests"
      onClose={onClose}
      actions={online ? (
        <><Button label={t('common.close')} onClick={onClose} /><Button primary className="st-guests-invite" label={t('calendar.invite')} disabled={busy || !chosen.length} onClick={send} /></>
      ) : <Button label={t('common.close')} onClick={onClose} />}
    >
      {!talk ? <Spinner /> : !online ? (
        <>
          <p className="st-guests-off">{t('studio.guests.guests-come-through')}</p>
          <Button icon="talk" className="st-guests-open-talk" label={t('studio.guests.open-talk')} onClick={onOpenTalk} />
        </>
      ) : (
        <>
          <p>{how === 'live'
            ? t('studio.guests.ask-three-talk-contacts')
            : t('studio.guests.ask-three-talk')}</p>
          <p className="st-guests-note" role="note" data-how={how}><Icon name={how === 'live' ? 'globe' : 'record'} size={15} /><span>{t('studio.guests.everyone-asked-told')} {how === 'live' ? liveNotice() : recordedNotice()}</span></p>
          {people.length ? (
            <>
              <h3>{t('studio.guests.here-now')}</h3>
              <div className="st-guests-list st-guests-here">
                {people.map((p) => (
                  <div key={p.id} className="st-guests-row" data-person={p.id} data-where={p.where}>
                    <span><strong>{p.name || t('common.someone')}</strong><small className={p.where === 'on-air' ? 'air' : ''}>{whereText(p.where)}{p.muted ? ` · ${t('talk.meet-lines.muted')}` : ''}</small></span>
                    <Button ghost label={t('common.mute')} className="st-guests-mute" disabled={p.muted || !p.connected} onClick={() => onMute(p)} />
                    <Button ghost label={t('talk.groups.take-out')} className="st-guests-out" onClick={() => onTakeOut(p)} />
                  </div>
                ))}
              </div>
            </>
          ) : null}
          {free > 0 ? (
            <>
              <h3>{t('common.contacts')}</h3>
              {contacts.length ? (
                <div className="st-guests-list st-guests-pick" role="group" aria-label={t('common.contacts')}>
                  {contacts.map((c) => (
                    <label key={c.id} className="st-guests-row" data-contact={c.id}>
                      <input type="checkbox" checked={chosen.includes(c.id)} disabled={!chosen.includes(c.id) && chosen.length >= free} onChange={() => toggle(c.id)} />
                      <span><strong>{c.name || c.email || t('common.someone')}</strong>{c.presence === 'online' ? null : <small>{t('studio.guests.not-talk-just')}</small>}</span>
                    </label>
                  ))}
                </div>
              ) : <p className="st-guests-none">{t('studio.guests.no-other-talk')}</p>}
            </>
          ) : <p className="st-guests-full">{t('studio.guests.all-three-seats')}</p>}
          {holding ? (
            <div className="st-guests-foot">
              <Button ghost icon="link" className="st-guests-copy" label={t('talk.meeting.copy-meeting-link')} title={t('studio.guests.send-link-contact')} onClick={onCopy} />
              <Button ghost icon="talk" className="st-guests-talk" label={t('studio.guests.continue-in-talk')} disabled={onAir} title={onAir ? (how === 'live' ? t('studio.guests.end-broadcast-first') : t('studio.guests.stop-recording-first')) : t('studio.guests.window-leaves-talk')} onClick={onToTalk} />
              <Button ghost className="st-guests-end" label={t('talk.meeting.end-meeting-everyone')} onClick={onEnd} />
            </div>
          ) : null}
        </>
      )}
    </Dialog>
  );
}
