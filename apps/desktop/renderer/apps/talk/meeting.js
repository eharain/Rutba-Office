// A small meeting in Talk (Office 2.28): up to four people, each computer
// connected to each other.
//
// Kept plain on purpose (owner, 2026-10-21: "for chats and conference call
// etc build clean simple user interface do not complicate them like Ms
// teams"): one dialog sets a meeting up, one shows it before joining, and the
// meeting itself is the people, one row of buttons, and a side that shows
// either who is here or what was written.
//
// A meeting set up to be streamed live says so where it is set up, before
// anyone joins and all the while they are in it (D-2026-11-22-19): that is
// chosen at set-up and there is no switch for it in the meeting.

import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Button, Dialog, Icon, Input, Spinner, formatWhen } from '@rutba/office-ui';
import { t, tn } from '../talk-words.js';
import { liveNotice, meetRefusal, meetingName, personMarks, tilesAcross, whenText } from './meet-lines.js';

const nameOf = (p) => p?.name || p?.email || t('common.someone');
const initials = (name) => String(name || '?').split(/\s+/).filter(Boolean).slice(0, 2).map((w) => w[0].toUpperCase()).join('') || '?';
/** A hand, raised: a sign every language reads. */
const HAND = '✋';

/** The notice of a meeting set up to be streamed live. */
export function LiveNotice({ className = '' }) {
  return <p className={`tk-live ${className}`} role="note" data-live="yes"><Icon name="globe" size={15} /><span>{liveNotice()}</span></p>;
}

/* ── setting one up, to start now ─────────────────────────────────── */

/**
 * Meet now: a name, who to ask (in a group conversation, everyone in it is
 * told), and the one switch. `group` is { id, title } when it starts there.
 */
export function MeetSetup({ shell, state, group = null, toast, onClose, onStarted }) {
  const [name, setName] = useState(() => meetingName({ group: group?.title || null, me: state.me }));
  const [chosen, setChosen] = useState([]);
  const [live, setLive] = useState(false);
  const [busy, setBusy] = useState(false);
  const most = 3;
  const toggle = (id) => setChosen((c) => (c.includes(id) ? c.filter((x) => x !== id) : c.length < most ? [...c, id] : c));
  const start = async (video) => {
    setBusy(true);
    try {
      const r = await shell.talk.meetCreate({ name: name.trim(), live, groupId: group?.id || null, invite: group ? [] : chosen, now: true });
      if (r.error) { toast(meetRefusal(r.error, r), { tone: 'bad', ms: 5000 }); return; }
      onClose();
      onStarted(r.meeting, { video });
    } catch (err) {
      toast(err?.message || String(err), { tone: 'bad', ms: 5000 });
    } finally {
      setBusy(false);
    }
  };
  const ready = name.trim() && (group || chosen.length > 0);
  return (
    <Dialog
      title={t('talk.meet-now')}
      width={460}
      className="tk-meet-setup"
      onClose={onClose}
      actions={<><Button label={t('common.cancel')} onClick={onClose} /><Button className="tk-meet-start" label={t('common.start')} disabled={busy || !ready} onClick={() => start(false)} /><Button primary icon="camera" className="tk-meet-start-video" label={t('talk.meeting.start-with-video')} disabled={busy || !ready} onClick={() => start(true)} /></>}
    >
      <p className="tk-lead">{group ? t('talk.meeting.everyone-group-told') : t('talk.meeting.meeting-holds-four')}</p>
      <label className="tk-field"><span>{t('talk.meeting.name-meeting')}</span><Input className="tk-meet-name" value={name} maxLength={80} onChange={(e) => setName(e.target.value)} /></label>
      {group ? null : state.contacts.length ? (
        <div className="tk-pick-list" role="group" aria-label={t('common.contacts')}>
          {state.contacts.map((c) => (
            <label key={c.id} className="tk-pick" data-pick={c.id}>
              <input type="checkbox" checked={chosen.includes(c.id)} disabled={!chosen.includes(c.id) && chosen.length >= most} onChange={() => toggle(c.id)} />
              <span className="tk-who"><strong>{nameOf(c)}</strong><small>{c.presence && c.presence !== 'offline' ? c.email : t('talk.offline')}</small></span>
            </label>
          ))}
        </div>
      ) : <p className="tk-none">{t('contacts.no-contacts-yet')}</p>}
      <label className="tk-meet-live-switch">
        <input type="checkbox" className="tk-meet-live" checked={live} onChange={(e) => setLive(e.target.checked)} />
        <span><strong>{t('calendar.meeting-streamed-live')}</strong><small>{t('talk.meeting.everyone-asked-told')}</small></span>
      </label>
    </Dialog>
  );
}

/* ── before joining ───────────────────────────────────────────────── */

/** A meeting looked at before joining: what it is, who set it up, who is in it, and the notice when it will be streamed live. */
export function MeetJoin({ shell, link, online, onClose, onJoin }) {
  const [seen, setSeen] = useState(null);
  useEffect(() => {
    if (!online) return undefined;
    let gone = false;
    setSeen(null);
    shell.talk.meetPeek({ link }).then((r) => { if (!gone) setSeen(r); }, () => { if (!gone) setSeen({ error: 'offline' }); });
    return () => { gone = true; };
  }, [shell, link, online]);
  const m = seen?.meeting || null;
  const inIt = (m?.people || []).map(nameOf);
  return (
    <Dialog
      title={m ? m.name : t('talk.meeting.join-a-meeting')}
      width={440}
      className="tk-meet-join"
      onClose={onClose}
      actions={m ? (
        <><Button label={t('common.cancel')} onClick={onClose} /><Button className="tk-meet-join-voice" label={t('common.join')} onClick={() => { onClose(); onJoin(m, { video: false }); }} /><Button primary icon="camera" className="tk-meet-join-video" label={t('talk.meeting.join-with-video')} onClick={() => { onClose(); onJoin(m, { video: true }); }} /></>
      ) : <Button label={t('common.close')} onClick={onClose} />}
    >
      {!online ? <p className="tk-lead">{t('talk.invite.talk-not-connected')}</p>
        : !seen ? <Spinner />
          : seen.error ? <p className="tk-lead tk-meet-refused" data-refused={seen.error}>{meetRefusal(seen.error, seen)}</p> : (
            <div className="tk-meet-seen" data-meeting={m.id}>
              <p className="tk-lead">{[t('talk.meeting.set-name', { name: nameOf(m.starter) }), whenText(m.at)].filter(Boolean).join(' · ')}</p>
              <p className="tk-lead">{inIt.length ? t('talk.meeting.now-names', { names: inIt.join(', ') }) : t('talk.meeting.nobody-yet')}</p>
              {m.locked ? <p className="tk-lead">{t('talk.meet-lines.meeting-locked-nobody')}</p> : null}
              {m.live ? <LiveNotice /> : null}
            </div>
          )}
    </Dialog>
  );
}

/** Asked into a meeting that is on now: who asks, and Join. */
export function MeetCard({ asked, onJoin, onDismiss }) {
  const m = asked.meeting;
  return (
    <div className="tk-ring tk-meet-card" role="alertdialog" aria-label={t('common.meeting')} data-meeting={m.id}>
      <span className="tk-glyph" aria-hidden="true"><Icon name="camera" size={16} /></span>
      <span className="tk-who"><strong>{m.name}</strong><small>{[t('talk.meeting.name-started-meeting', { name: nameOf(asked.by) }), m.live ? t('talk.meeting.streamed-live') : ''].filter(Boolean).join(' · ')}</small></span>
      <Button primary label={t('common.join')} className="tk-meet-card-join" onClick={onJoin} />
      <Button label={t('common.not-now')} className="tk-meet-card-later" onClick={onDismiss} />
    </div>
  );
}

/* ── the meeting ──────────────────────────────────────────────────── */

/** A stream shown in a video element, kept in step as tracks come and go. */
function Screen({ stream, muted = false, className }) {
  const ref = useRef(null);
  useEffect(() => {
    if (ref.current && ref.current.srcObject !== stream) ref.current.srcObject = stream || null;
  });
  return <video ref={ref} className={className} autoPlay playsInline muted={muted} />;
}

function Tile({ person, local }) {
  const picture = person.me ? Boolean(local) : (person.camera || person.sharing) && person.stream;
  return (
    <div className={`tk-tile${person.hand ? ' hand' : ''}`} data-person={person.id} data-connected={person.connected ? 'yes' : 'no'}>
      {person.me
        ? (local ? <Screen stream={local} muted className={`tk-tile-video${person.sharing ? '' : ' mirror'}`} /> : null)
        : person.stream ? <Screen stream={person.stream} className={picture ? 'tk-tile-video' : 'tk-hidden'} /> : null}
      {picture ? null : <span className="tk-tile-face" aria-hidden="true">{initials(person.name)}</span>}
      <span className="tk-tile-name">
        {person.muted ? <Icon name="micOff" size={13} /> : null}
        {person.hand ? <span aria-hidden="true">{HAND}</span> : null}
        <span>{person.me ? t('common.you') : nameOf(person)}</span>
        {!person.me && !person.connected ? <small>{person.failed ? t('talk.meeting.could-not-connect') : t('talk.meeting.connecting')}</small> : null}
      </span>
    </div>
  );
}

/**
 * The meeting: the people, one row of buttons, and a side for who is here
 * or what was written. `onShare` shares a screen the way a call does.
 */
export function MeetingView({ shell, room, toast, asked, onShare }) {
  const [side, setSide] = useState(null);
  const [text, setText] = useState('');
  const end = useRef(null);
  const mesh = room.mesh;
  const meeting = room.meeting;
  const people = room.people();
  const me = people[0];
  const local = useMemo(() => new MediaStream(), []);
  const shown = mesh?.screen || mesh?.camera || null;
  for (const tr of local.getTracks()) if (tr !== shown) local.removeTrack(tr);
  if (shown && !local.getTracks().includes(shown)) local.addTrack(shown);

  useEffect(() => { if (side === 'messages') { room.seen(); end.current?.scrollIntoView({ block: 'end' }); } }, [side, room, room.says.length]);

  if (!mesh || !meeting || !me) return <div className="tk-callview tk-meeting" data-state={room.state}><div className="tk-stage"><Spinner /></div></div>;

  const refuse = (r) => { if (r?.error) toast(meetRefusal(r.error, r), { tone: 'bad', ms: 5000 }); };
  const say = () => {
    const words = text.trim();
    if (!words) return;
    setText('');
    room.say(words);
  };
  const copy = async () => {
    await shell.clipboard.writeText({ text: meeting.link }).catch(() => {});
    toast(t('talk.meeting.link-copied-paste'), { tone: 'good', ms: 2500 });
  };
  const takeOut = async (p) => {
    if (!(await asked(t('talk.meeting.take-name-out', { name: nameOf(p) }), t('talk.meeting.cannot-join-meeting'), t('talk.groups.take-out')))) return;
    refuse(await room.remove(p.id));
  };
  const endAll = async () => {
    if (!(await asked(t('talk.meeting.end-the-meeting-for-everyone'), t('talk.meeting.everyone-taken-out'), t('talk.meeting.end-meeting')))) return;
    refuse(await room.end());
  };
  const across = tilesAcross(people.length);

  return (
    <div className="tk-callview tk-meeting" data-state={room.state} data-meeting={meeting.id} data-live={meeting.live ? 'yes' : 'no'}>
      <header className="tk-meet-head">
        <strong>{meeting.name}</strong>
        <span>{tn(people.length, 'talk.count-person')}</span>
        {meeting.locked ? <span className="tk-meet-locked"><Icon name="lock" size={13} />{t('sheets.ribbon.locked')}</span> : null}
      </header>
      {meeting.live ? <LiveNotice className="in" /> : null}
      <div className="tk-meet-body">
        <div className="tk-tiles" style={{ gridTemplateColumns: `repeat(${across}, minmax(0, 1fr))` }} role="group" aria-label={t('common.people')}>
          {people.map((p) => <Tile key={p.id} person={p} local={p.me ? (shown ? local : null) : null} />)}
        </div>
        {side ? (
          <aside className="tk-meet-side" data-side={side}>
            <header><strong>{side === 'people' ? t('common.people') : t('mail.messages')}</strong><Button ghost icon="close" title={t('common.close')} aria-label={t('common.close')} onClick={() => setSide(null)} /></header>
            {side === 'people' ? (
              <>
                <div className="tk-meet-list">
                  {people.map((p) => (
                    <div key={p.id} className="tk-meet-row" data-person={p.id}>
                      <span className="tk-who"><strong>{nameOf(p)}</strong><small>{personMarks(p, { me: p.me, starter: p.starter })}</small></span>
                      {room.mine && !p.me ? (
                        <>
                          <Button ghost label={t('common.mute')} className="tk-meet-mute" disabled={p.muted || !p.connected} onClick={() => room.mute(p.id)} />
                          <Button ghost label={t('talk.groups.take-out')} className="tk-meet-remove" onClick={() => takeOut(p)} />
                        </>
                      ) : null}
                    </div>
                  ))}
                </div>
                <div className="tk-meet-foot">
                  <Button ghost icon="link" label={t('talk.meeting.copy-meeting-link')} className="tk-meet-copy" onClick={copy} />
                  {room.mine ? (
                    <>
                      <label className="tk-meet-lock"><input type="checkbox" checked={Boolean(meeting.locked)} onChange={async (e) => refuse(await room.lock(e.target.checked))} /><span>{t('talk.meeting.lock-meeting-nobody')}</span></label>
                      <Button ghost label={t('talk.meeting.end-meeting-everyone')} className="tk-meet-end" onClick={endAll} />
                    </>
                  ) : null}
                </div>
              </>
            ) : (
              <>
                <div className="tk-meet-says" role="log" aria-live="polite">
                  {room.says.length ? room.says.map((s) => (
                    <div key={s.id} className={`tk-bubble${s.mine ? ' mine' : ''}`} data-say={s.id}>
                      {s.mine ? null : <strong className="tk-from">{s.name || t('common.someone')}</strong>}
                      <span className="tk-text">{s.text}</span>
                      <span className="tk-meta"><time>{formatWhen(s.at)}</time></span>
                    </div>
                  )) : <p className="tk-none">{t('talk.meeting.written-seen-people')}</p>}
                  <div ref={end} />
                </div>
                <form className="tk-compose" onSubmit={(e) => { e.preventDefault(); say(); }}>
                  <textarea className="tk-input tk-meet-input" rows={1} value={text} placeholder={t('common.write-a-message')} aria-label={t('common.write-a-message')} onChange={(e) => setText(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); say(); } }} />
                  <Button primary icon="send" title={t('common.send')} className="tk-meet-send" disabled={!text.trim()} onClick={say} />
                </form>
              </>
            )}
          </aside>
        ) : null}
      </div>
      <div className="tk-controls">
        <button type="button" className={`tk-round${me.muted ? ' off' : ''}`} aria-pressed={me.muted} disabled={!mesh.mic} title={!mesh.mic ? t('talk.meeting.no-microphone-found') : me.muted ? t('common.unmute') : t('common.mute')} onClick={() => mesh.setMuted(!mesh.muted)}><Icon name={me.muted ? 'micOff' : 'mic'} size={20} /></button>
        <button type="button" className={`tk-round${mesh.camera ? '' : ' off'}`} aria-pressed={!mesh.camera} title={mesh.camera ? t('talk.turn-camera-off') : t('talk.turn-camera')} onClick={() => mesh.setCamera(!mesh.camera)}><Icon name="camera" size={20} /></button>
        <button type="button" className={`tk-round${mesh.screen ? ' on' : ''}`} aria-pressed={Boolean(mesh.screen)} title={mesh.screen ? t('talk.stop-sharing') : t('talk.share-screen')} onClick={() => onShare(mesh)}><Icon name="screen" size={20} /></button>
        <button type="button" className={`tk-round tk-meet-hand${mesh.hand ? ' on' : ''}`} aria-pressed={mesh.hand} title={mesh.hand ? t('talk.meeting.lower-your-hand') : t('talk.meeting.raise-your-hand')} aria-label={mesh.hand ? t('talk.meeting.lower-your-hand') : t('talk.meeting.raise-your-hand')} onClick={() => mesh.setHand(!mesh.hand)}><span aria-hidden="true">{HAND}</span></button>
        <button type="button" className={`tk-round tk-meet-messages${side === 'messages' ? ' on' : ''}`} aria-pressed={side === 'messages'} title={t('mail.messages')} aria-label={t('mail.messages')} onClick={() => setSide(side === 'messages' ? null : 'messages')}><Icon name="talk" size={20} />{room.unseen && side !== 'messages' ? <i className="tk-round-count">{room.unseen > 9 ? '9+' : room.unseen}</i> : null}</button>
        <button type="button" className={`tk-round tk-meet-people${side === 'people' ? ' on' : ''}`} aria-pressed={side === 'people'} title={t('common.people')} aria-label={t('common.people')} onClick={() => setSide(side === 'people' ? null : 'people')}><Icon name="contacts" size={20} /></button>
        <button type="button" className="tk-round end tk-meet-leave" title={t('talk.meeting.leave-the-meeting')} aria-label={t('talk.meeting.leave-the-meeting')} onClick={() => room.leave()}><Icon name="phone" size={20} /></button>
      </div>
    </div>
  );
}
