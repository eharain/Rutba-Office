// Rutba Talk: messages and calls with the people you are in contact with.
//
// Kept plain on purpose (owner, 2026-10-21: "for chats and conference call
// etc build clean simple user interface do not complicate them like Ms
// teams"): the people on the left, the conversation on the right, Call and
// Video above it, and a call that fills the conversation with four buttons.
//
// Someone is a contact once both of you have said so: you ask by their
// address, they accept. Only contacts can call or write to each other.
// Asking also emails them, or a link goes to friends any way the person
// likes (talk/invite.js).

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Button, Icon, Input, Select, Panel, Content, Empty, Dialog, Spinner, useToast, useMenu, formatWhen, t, tn } from '@rutba/office-ui';
import { AppFrame, useAppMenu } from '../shell.js';
import { Call } from './talk/call.js';
import { TALK_CSS } from './talk/css.js';
import { AddContact, InviteLinks, OpenInvitation, Invitation, toldText, errorText, againLetter } from './talk/invite.js';
import { InviteByEmail, InviteByEmailButton } from './talk/invite-email.js';
import { MyAddress, MyAddressButton } from './talk/address.js';
import { NewGroup, GroupPeople } from './talk/groups.js';
import { Report } from './talk/report.js';
import { MissedCalls } from './talk/missed.js';
import { TalkSettings } from './talk/settings.js';
import { createRinger, soundFor, incomingCall, onAirNow } from './talk/ring.js';
import { MeetingRoom } from '../meeting/meet-room.js';
import { MeetSetup, MeetJoin, MeetCard, MeetingView, AirChip } from './talk/meeting.js';
import { airMarks, meetingCode, meetingLine, meetRefusal, goneText, serverFaultText, withholdOf } from './talk/meet-lines.js';
import { groupKey, groupIdOf, countText, titleWithCount, waitingIn, groupTitle, groupPeopleLine, senderName, eventText, missedText, markOf, leaveGroupText, refusalText } from './talk/lines.js';

const PRESENCE = () => [['online', t('talk.available')], ['busy', t('common.busy')], ['away', t('talk.away')], ['hidden', t('talk.appear-offline')]];
// The shapes of an address, shown in empty fields: examples, not words to translate.
const SERVER_SHAPE = 'wss://introducer.example/ws';
const RELAY_SHAPE = 'turn:relay.example.com:3478';
/** What the introducer says that changes a conversation's own lines, so the one shown is read again. */
const HOLDS = new Set(['messages', 'msg.copy', 'msg.sent', 'group.messages', 'group.copy', 'group.sent', 'group.created', 'group.added', 'group.changed', 'notices', 'call.ended', 'call.declined', 'meet.invited', 'meet.created', 'meet.ended']);
const callId = () => `c${Array.from(crypto.getRandomValues(new Uint8Array(12)), (b) => b.toString(16).padStart(2, '0')).join('')}`;
const initials = (name) => String(name || '?').split(/\s+/).filter(Boolean).slice(0, 2).map((w) => w[0].toUpperCase()).join('') || '?';

function Avatar({ person, size = 32 }) {
  return (
    <span className="tk-avatar" style={{ width: size, height: size, fontSize: size * 0.4 }} aria-hidden="true">
      {initials(person?.name || person?.email)}
      {person?.presence ? <i className={`tk-dot ${person.presence}`} /> : null}
    </span>
  );
}

/** How many wait in a conversation, as a small mark at the end of its row. */
function Count({ n, label }) {
  return n > 0 ? <span className="tk-count" role="img" aria-label={label}>{countText(n)}</span> : null;
}

/** A stream shown in a video element, kept in step as tracks come and go. */
function Screen({ stream, muted = false, className }) {
  const ref = useRef(null);
  useEffect(() => {
    if (ref.current && ref.current.srcObject !== stream) ref.current.srcObject = stream || null;
  });
  return <video ref={ref} className={className} autoPlay playsInline muted={muted} />;
}

/* ── signed out ───────────────────────────────────────────────────── */

/**
 * Signed out: Sign in with Rutba. The box for a server address and a token is
 * drawn only where the service takes them (state.testServers: the window
 * checks and the screens build); a copy in use has no place to give a server.
 */
function SignIn({ shell, state, toast }) {
  const [server, setServer] = useState(state.server || '');
  const [token, setToken] = useState('');
  const [busy, setBusy] = useState(false);
  const connect = async () => {
    setBusy(true);
    try {
      await shell.talk.signIn({ token: token.trim(), server: server.trim() || undefined });
    } catch (err) {
      toast(err?.message || String(err), { tone: 'bad', ms: 6000 });
    } finally {
      setBusy(false);
    }
  };
  // Rutba sign-in in the browser; the window waits while it is open there.
  const signInWithRutba = () => shell.talk.signInBrowser().catch(() => { /* the state says why */ });
  const waiting = state.phase === 'signing-in';
  return (
    <div className="tk-signin" data-waiting={waiting ? 'yes' : 'no'}>
      <Empty icon="talk" title={t('talk.talk-contacts')}>
        <p>{t('talk.messages-calls-people')}</p>
        {state.error && !waiting ? <p className="tk-error" role="alert">{state.error}</p> : null}
        {waiting ? (
          // The browser is open at auth.rutba.io: the wait can be given up, and the browser opened again if its tab was closed.
          <div className="tk-signin-waiting" role="status">
            <p className="tk-signin-wait"><Spinner /><span>{t('talk.waiting')}</span></p>
            <p className="tk-soon">{t('talk.finish-signing-browser')}</p>
            <div className="tk-signin-actions">
              <Button label={t('talk.open-browser-again')} className="tk-rutba-again" onClick={() => shell.talk.signInReopen().catch(() => {})} />
              <Button label={t('common.cancel')} className="tk-rutba-cancel" onClick={() => shell.talk.signInCancel().catch(() => {})} />
            </div>
          </div>
        ) : (
          <>
            <Button primary label={t('talk.sign-rutba')} className="tk-rutba" onClick={signInWithRutba} />
            <p className="tk-soon">{t('talk.browser-opens-auth')}</p>
          </>
        )}
      </Empty>
      {state.testServers ? (
      <details className="tk-test">
        <summary>{t('talk.connect-test-server')}</summary>
        <label className="tk-field"><span>{t('common.server')}</span><Input className="tk-server" value={server} placeholder={SERVER_SHAPE} onChange={(e) => setServer(e.target.value)} /></label>
        <label className="tk-field"><span>{t('talk.sign-in-token')}</span><Input className="tk-token" value={token} onChange={(e) => setToken(e.target.value)} /></label>
        <Button primary label={busy ? t('common.connecting') : t('common.connect')} className="tk-connect" disabled={busy || !token.trim()} onClick={connect} />
      </details>
      ) : null}
    </div>
  );
}

/* ── the people on the left ───────────────────────────────────────── */

function People({ state, current, onPick, onAdd, onAnswerRequest, onRequestMenu, onCopy, onContactMenu, onGroupMenu, onWithdraw, onRemind, onMissed, onNewGroup, onMeet }) {
  const { contacts, requests } = state;
  const groups = state.groups || [];
  const unread = state.unread || {};
  const me = state.me?.id;
  return (
    <div className="tk-people">
      <button type="button" className="tk-contact tk-meet-now" disabled={!contacts.length} onClick={onMeet}>
        <span className="tk-glyph" aria-hidden="true"><Icon name="camera" size={15} /></span>
        <span className="tk-who"><strong>{t('talk.meet-now')}</strong></span>
      </button>
      <button type="button" className="tk-contact tk-missed-row" onClick={onMissed}>
        <span className="tk-glyph" aria-hidden="true"><Icon name="phone" size={15} /></span>
        <span className="tk-who"><strong>{t('talk.missed-calls')}</strong></span>
        <Count n={state.missed || 0} label={t('talk.count-not-looked', { count: state.missed || 0 })} />
      </button>
      {requests.incoming.length ? (
        <section className="tk-section">
          <h3>{t('talk.asking-contact')}</h3>
          {requests.incoming.map((r) => (
            <div key={r.id} className="tk-request" data-request={r.id}>
              <Avatar person={r.from} size={28} />
              <span className="tk-who"><strong>{r.from?.name || r.from?.email}</strong>{r.note ? <small>{r.note}</small> : <small>{r.from?.email}</small>}</span>
              <Button primary label={t('common.accept')} className="tk-accept" onClick={() => onAnswerRequest(r, 'accept')} />
              <Button icon="close" title={t('common.decline')} aria-label={t('common.decline')} className="tk-decline" onClick={() => onAnswerRequest(r, 'decline')} />
              {/* Someone unknown asking: they can be blocked or reported from here, without being accepted first. */}
              {r.from?.id ? <Button ghost icon="more" title={t('common.more')} aria-label={t('common.more')} className="tk-request-more" onClick={(e) => onRequestMenu(e, r)} /> : null}
            </div>
          ))}
        </section>
      ) : null}
      <section className="tk-section">
        <div className="tk-section-head">
          <h3>{t('common.contacts')}</h3>
          <Button ghost icon="plus" label={t('common.add')} className="tk-add" onClick={onAdd} />
        </div>
        {contacts.length ? contacts.map((c) => (
          <button key={c.id} type="button" className={`tk-contact${current === c.id ? ' on' : ''}`} data-contact={c.id} onClick={() => onPick(c.id)} onContextMenu={(e) => { e.preventDefault(); onContactMenu(e, c); }}>
            <Avatar person={c} />
            <span className="tk-who"><strong>{c.name || c.email}</strong><small>{c.presence === 'offline' ? t('talk.offline') : PRESENCE().find(([k]) => k === c.presence)?.[1] || ''}</small></span>
            <Count n={unread[c.id] || 0} label={t('talk.count-not-read', { count: unread[c.id] || 0 })} />
          </button>
        )) : <p className="tk-none">{t('talk.no-contacts-yet')}</p>}
        {requests.outgoing.length ? (
          <div className="tk-waiting">
            {requests.outgoing.map((r) => (
              <div key={r.id} className="tk-asked" data-asked={r.email} data-told={r.told ? `${r.told.by}:${r.told.sent ? 'sent' : r.told.reason || 'not'}` : ''}>
                <small>
                  <strong>{t('talk.asked-email', { email: r.email })}</strong>
                  <span className="tk-told">{[r.at ? t('talk.sent-when', { when: formatWhen(r.reminded || r.at) }) : '', r.told ? toldText(r.told) : ''].filter(Boolean).join(' · ')}</span>
                </small>
                {r.link ? <Button ghost icon="link" title={t('talk.copy-invitation-link')} className="tk-asked-copy" onClick={() => onCopy(r.link)} /> : null}
                {/* Again once the introducer allows it (two days on); an introducer from before says nothing, so not yet. */}
                <Button ghost icon="refresh" title={r.remindAt && Date.now() >= r.remindAt ? t('talk.invite-email-again', { email: r.email }) : r.remindAt ? t('talk.invite-again', { when: formatWhen(r.remindAt) }) : t('talk.invite-again-not')} aria-label={t('talk.invite-email-again', { email: r.email })} className="tk-asked-again" disabled={!r.remindAt || Date.now() < r.remindAt} onClick={() => onRemind(r)} />
                <Button ghost icon="close" title={t('talk.withdraw-invitation', { email: r.email })} aria-label={t('talk.withdraw-invitation', { email: r.email })} className="tk-asked-withdraw" onClick={() => onWithdraw(r)} />
              </div>
            ))}
          </div>
        ) : null}
      </section>
      <section className="tk-section tk-groups">
        <div className="tk-section-head">
          <h3>{t('talk.groups')}</h3>
          <Button ghost icon="plus" label={t('common.new')} className="tk-group-new" disabled={contacts.length < 2} title={contacts.length < 2 ? t('talk.group-needs-two') : ''} onClick={onNewGroup} />
        </div>
        {groups.length ? groups.map((g) => (
          <button key={g.id} type="button" className={`tk-contact${current === groupKey(g.id) ? ' on' : ''}`} data-group={g.id} onClick={() => onPick(groupKey(g.id))} onContextMenu={(e) => { e.preventDefault(); onGroupMenu(e, g); }}>
            <Avatar person={{ name: groupTitle(g, me) }} />
            <span className="tk-who"><strong>{groupTitle(g, me)}</strong><small>{tn(g.members.length, 'talk.count-person')}</small></span>
            <Count n={unread[groupKey(g.id)] || 0} label={t('talk.count-not-read', { count: unread[groupKey(g.id)] || 0 })} />
          </button>
        )) : <p className="tk-none">{t('talk.group-three-more')}</p>}
      </section>
    </div>
  );
}

/* ── the conversation ─────────────────────────────────────────────── */

/** One thing in a conversation: a message, or a line saying a call was missed or what changed in a group. */
function Entry({ m, me, group, online = true, onJoinMeeting, onAgain, onForget }) {
  // A meeting known to be over offers no Join.
  if (m.kind === 'meeting') return <div className="tk-line" data-line="meeting" data-ended={m.ended ? 'yes' : 'no'}><Icon name="camera" size={13} /><span>{meetingLine(m, me?.id)}</span>{m.ended ? null : airMarks(m.meeting).map((a) => <AirChip key={a.kind} mark={a} />)}<time>{formatWhen(m.at)}</time>{m.ended ? <small className="tk-line-ended">{t('talk.ended')}</small> : <Button ghost label={t('common.join')} className="tk-line-join" onClick={() => onJoinMeeting?.(m.meeting.id)} />}</div>;
  if (m.kind === 'missed-call') return <div className="tk-line" data-line="missed-call"><Icon name="phone" size={13} /><span>{missedText(m)}</span><time>{formatWhen(m.at)}</time></div>;
  if (m.kind === 'group-event') {
    const words = eventText(m.event, me?.id);
    return words ? <div className="tk-line" data-line={m.event.kind}><span>{words}</span></div> : null;
  }
  const mark = markOf(m, { online });
  return (
    <div className={`tk-bubble${m.mine ? ' mine' : ''}`} data-message={m.id} data-state={m.mine ? m.state || '' : undefined}>
      {group && !m.mine ? <strong className="tk-from">{senderName(group, m.from)}</strong> : null}
      <span className="tk-text">{m.text}</span>
      <span className="tk-meta">
        <time>{formatWhen(m.at)}</time>
        {mark ? <span className={`tk-mark${mark.bad ? ' bad' : ''}`} role="img" title={mark.words} aria-label={mark.words}>{mark.icon ? <><Icon name={mark.icon} size={12} />{mark.double ? <Icon name={mark.icon} size={12} /> : null}</> : mark.words}</span> : null}
      </span>
      {mark?.waiting ? (
        <span className="tk-unsent tk-waiting">
          <Button ghost label={t('talk.send-again')} className="tk-waiting-again" onClick={() => onAgain?.(m)} />
        </span>
      ) : null}
      {mark?.failed ? (
        <span className="tk-unsent">
          <Button ghost label={t('voices-dialog.try-again')} className="tk-unsent-again" onClick={() => onAgain?.(m)} />
          <Button ghost label={t('common.remove')} className="tk-unsent-forget" onClick={() => onForget?.(m)} />
        </span>
      ) : null}
    </div>
  );
}

function Conversation({ shell, contact, group, me, history, online = true, onCall, onMore, onPeople, onMeet, onJoinMeeting, onChanged, toast }) {
  const [text, setText] = useState('');
  const conversation = group ? groupKey(group.id) : contact.id;
  // A message that was not sent: sent again under its own id, or taken out of the conversation here.
  const again = async (m) => {
    const r = await shell.talk.messageAgain({ conversation, id: m.id }).catch(() => ({ error: 'offline' }));
    if (r.error === 'offline') toast(t('talk.invite.talk-not-connected'), { tone: 'bad', ms: 5000 });
    onChanged?.();
  };
  const forget = async (m) => {
    await shell.talk.messageForget({ conversation, id: m.id }).catch(() => {});
    onChanged?.();
  };
  const end = useRef(null);
  useEffect(() => { end.current?.scrollIntoView({ block: 'end' }); }, [history.length, contact?.id, group?.id]);
  const send = async () => {
    const words = text.trim();
    if (!words) return;
    setText('');
    try {
      await shell.talk.message(group ? { group: group.id, text: words } : { to: contact.id, text: words });
    } catch (err) {
      setText(words);
      toast(err?.message || String(err), { tone: 'bad', ms: 5000 });
    }
  };
  const reachable = contact?.presence && contact.presence !== 'offline';
  const title = group ? groupTitle(group, me?.id) : contact.name || contact.email;
  return (
    <div className="tk-conversation" data-conversation={group ? groupKey(group.id) : contact.id}>
      {group ? (
        <header className="tk-head">
          <Avatar person={{ name: title }} size={36} />
          <span className="tk-who"><strong>{title}</strong><small>{groupPeopleLine(group, me?.id)}</small></span>
          <span className="tk-grow" />
          <Button primary icon="camera" label={t('talk.meet-now')} className="tk-group-meet" onClick={onMeet} />
          <Button label={t('common.people')} className="tk-group-people-btn" onClick={onPeople} />
        </header>
      ) : (
        <header className="tk-head">
          <Avatar person={contact} size={36} />
          <span className="tk-who"><strong>{title}</strong><small>{contact.email}</small></span>
          <span className="tk-grow" />
          <Button icon="phone" label={t('talk.call')} className="tk-call-voice" disabled={!reachable} title={reachable ? '' : t('talk.calls-reach')} onClick={() => onCall(contact, false)} />
          <Button primary icon="camera" label={t('common.video')} className="tk-call-video" disabled={!reachable} onClick={() => onCall(contact, true)} />
          <Button ghost icon="more" title={t('common.more')} aria-label={t('common.more')} className="tk-more" onClick={(e) => onMore(e, contact)} />
        </header>
      )}
      <div className="tk-messages" role="log" aria-live="polite">
        {history.length ? history.map((m) => <Entry key={m.id} m={m} me={me} group={group} online={online} onJoinMeeting={onJoinMeeting} onAgain={again} onForget={forget} />)
          : <p className="tk-none">{group ? t('talk.write-first-message') : t('talk.say-hello-name', { name: title })}</p>}
        <div ref={end} />
      </div>
      <form className="tk-compose" onSubmit={(e) => { e.preventDefault(); send(); }}>
        <textarea
          className="tk-input"
          rows={1}
          value={text}
          placeholder={t('common.write-a-message')}
          aria-label={t('common.write-a-message')}
          onChange={(e) => setText(e.target.value)}
          onKeyDown={(e) => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); send(); } }}
        />
        <Button primary icon="send" title={t('common.send')} className="tk-send" disabled={!text.trim()} onClick={send} />
      </form>
    </div>
  );
}

/* ── a call ───────────────────────────────────────────────────────── */

function CallView({ shell, call, onShare }) {
  const [, setTick] = useState(0);
  useEffect(() => {
    const timer = setInterval(() => setTick((n) => n + 1), 1000);
    return () => clearInterval(timer);
  }, []);
  const remoteVideo = call.remote.getVideoTracks().length > 0;
  const local = useMemo(() => new MediaStream(), []);
  for (const tr of local.getTracks()) if (tr !== (call.screen || call.camera)) local.removeTrack(tr);
  const shown = call.screen || call.camera;
  if (shown && !local.getTracks().includes(shown)) local.addTrack(shown);
  const secs = call.startedAt ? Math.floor((Date.now() - call.startedAt) / 1000) : 0;
  const words = { calling: t('talk.calling'), ringing: t('talk.ringing'), connecting: t('common.connecting'), connected: `${Math.floor(secs / 60)}:${String(secs % 60).padStart(2, '0')}`, ended: t('talk.call-ended') };
  return (
    <div className="tk-callview" data-state={call.state}>
      <div className="tk-stage">
        {remoteVideo ? <Screen stream={call.remote} className="tk-remote" /> : (
          <div className="tk-remote-still"><Avatar person={call.peer} size={96} /><strong>{call.peer.name}</strong></div>
        )}
        {/* The other side's voice, when there is no picture to carry it. */}
        {!remoteVideo ? <Screen stream={call.remote} className="tk-hidden" /> : null}
        {shown ? <Screen stream={local} muted className="tk-self" /> : null}
        <span className="tk-status">{words[call.state] || call.state}</span>
      </div>
      <div className="tk-controls">
        <button type="button" className={`tk-round${call.muted ? ' off' : ''}`} aria-pressed={call.muted} title={call.muted ? t('common.unmute') : t('common.mute')} onClick={() => call.setMuted(!call.muted)}><Icon name={call.muted ? 'micOff' : 'mic'} size={20} /></button>
        <button type="button" className={`tk-round${call.camera ? '' : ' off'}`} aria-pressed={!call.camera} title={call.camera ? t('talk.turn-camera-off') : t('talk.turn-camera')} onClick={() => call.setCamera(!call.camera)}><Icon name={call.camera ? 'camera' : 'cameraOff'} size={20} /></button>
        <button type="button" className={`tk-round${call.screen ? ' on' : ''}`} aria-pressed={Boolean(call.screen)} title={call.screen ? t('talk.stop-sharing') : t('talk.share-screen')} onClick={() => onShare(call)}><Icon name="screen" size={20} /></button>
        <button type="button" className="tk-round end" title={t('talk.hang-up')} onClick={() => call.end('hung-up')}><Icon name="phone" size={20} /></button>
      </div>
      {call.code ? <p className="tk-code" data-code={call.code}>{t('talk.safety-code-code', { code: call.code, name: call.peer.name })}</p> : null}
    </div>
  );
}

/** The conversation a link or a notification names: a person's id, or a group's key; null when it names none. */
const conversationOf = (q) => (q?.group ? groupKey(String(q.group)) : q?.contact ? String(q.contact) : null);

/* ── the window ───────────────────────────────────────────────────── */

export default function Talk({ app, shell }) {
  const toast = useToast();
  const [state, setState] = useState(null);
  // Opened at a conversation, as from a notification (?contact= or ?group=), or at none.
  const [current, setCurrent] = useState(() => conversationOf(Object.fromEntries(new URLSearchParams(location.search))));
  const [history, setHistory] = useState([]);
  const [adding, setAdding] = useState(false);
  const [linksOpen, setLinksOpen] = useState(false);
  // Invite by email, open on its own or on someone another app named (Contacts' "Invite to Talk": ?inviteEmail=).
  const [emailing, setEmailing] = useState(() => {
    const q = new URLSearchParams(location.search);
    return q.get('inviteEmail') ? [{ email: q.get('inviteEmail'), name: q.get('inviteName') || '' }] : false;
  });
  const [pasting, setPasting] = useState(false);
  // Taking a meeting to Live Broadcast or the Recorder, asked first: { room, app }.
  const [studioAsk, setStudioAsk] = useState(null);
  // An invitation this window was opened on (a rutba-office:// link), or one pasted.
  const [invitation, setInvitation] = useState(() => new URLSearchParams(location.search).get('invite') || null);
  const [relaysOpen, setRelaysOpen] = useState(false);
  const [addressOpen, setAddressOpen] = useState(false);
  // A new group being made, a group's people shown, someone being reported, the missed calls.
  const [groupNew, setGroupNew] = useState(false);
  const [peopleOpen, setPeopleOpen] = useState(false);
  const [reporting, setReporting] = useState(null);
  const [missedOpen, setMissedOpen] = useState(() => new URLSearchParams(location.search).get('show') === 'calls');
  const [settingsOpen, setSettingsOpen] = useState(false);
  // A meeting being set up to start now ({ group } or {}), one looked at before joining (its code: from a link, a card, a pasted address), and the one this window is in.
  const [meetSetup, setMeetSetup] = useState(null);
  const [meetAsk, setMeetAsk] = useState(() => meetingCode(new URLSearchParams(location.search).get('meet')));
  const [room, setRoom] = useState(null);
  const roomRef = useRef(null);
  roomRef.current = room;
  const [call, setCall] = useState(null);
  const [ringing, setRinging] = useState(null);
  const [, setCallTick] = useState(0);
  const callRef = useRef(null);
  callRef.current = call;
  // The call whose card is showing, known the moment it is taken: two calls can come in before the window draws again.
  const ringRef = useRef(null);
  // The calls taken elsewhere or ended lately, so one that ends while its ring is being decided rings no more.
  const overRef = useRef(new Set());
  ringRef.current = ringing;
  const currentRef = useRef(null);
  currentRef.current = current;
  const menu = useMenu();
  const appMenu = useAppMenu({ shell, appKey: 'talk' });

  useEffect(() => {
    shell.talk.resume().then(setState).catch(() => {});
    return shell.on('talk:state', setState);
  }, [shell]);
  // A link opened while this window was already open comes here, and so does what a notification opens Talk at:
  // a conversation ({ contact } or { group }), the missed calls ({ show: 'calls' }), a meeting ({ meet }).
  useEffect(() => {
    const at = (q) => {
      if (q?.invite) setInvitation(q.invite);
      if (meetingCode(q?.meet)) setMeetAsk(meetingCode(q.meet));
      if (q?.inviteEmail) setEmailing([{ email: q.inviteEmail, name: q.inviteName || '' }]);
      if (conversationOf(q)) setCurrent(conversationOf(q));
      if (q?.show === 'calls') setMissedOpen(true);
    };
    const offLink = shell.on('app:link', at);
    const offCommand = shell.on('app:command', ({ command, args } = {}) => { if (command === 'notify.open') at(args); });
    return () => { offLink?.(); offCommand?.(); };
  }, [shell]);
  const endInvitation = useCallback(() => setInvitation(null), []);
  // A contact taken off or blocked, an invitation withdrawn or sent again (D-2026-10-22-09).
  const asked = async (message, detail, yes) => (await shell.dialog.message({ type: 'question', message, detail, buttons: [yes, t('common.cancel')], cancelId: 1 })).response === 0;
  const contactMenu = (e, c) => menu.open(e, [
    { label: t('talk.remove-from'), run: async () => {
      if (!(await asked(t('talk.remove-name-contacts', { name: c.name || c.email }), t('talk.neither-write-call'), t('common.remove')))) return;
      const r = await shell.talk.removeContact({ contactId: c.id });
      if (r.error) toast(errorText(r.error), { tone: 'bad', ms: 5000 });
      else { if (currentRef.current === c.id) setCurrent(null); toast(t('talk.name-no-longer', { name: c.name || c.email }), { ms: 3000 }); }
    } },
    { label: t('talk.block-2'), run: async () => {
      if (!(await asked(t('talk.block-name', { name: c.name || c.email }), t('talk.taken-off-contacts'), t('talk.block')))) return;
      const r = await shell.talk.block({ contactId: c.id });
      if (r.error) toast(errorText(r.error), { tone: 'bad', ms: 5000 });
      else { if (currentRef.current === c.id) setCurrent(null); toast(t('talk.name-is-blocked', { name: c.name || c.email }), { ms: 3000 }); }
    } },
    { label: t('talk.report'), run: () => setReporting({ person: c }) },
  ]);
  // Someone asking to be a contact: blocked or reported without being accepted first. Their request is put away with the block.
  const requestMenu = (e, r) => {
    const name = r.from?.name || r.from?.email || '';
    menu.open(e, [
      { label: t('talk.block-2'), run: async () => {
        if (!(await asked(t('talk.block-name', { name }), t('talk.cannot-ask-contact'), t('talk.block')))) return;
        const res = await shell.talk.block({ contactId: r.from.id, requestId: r.id }).catch(() => ({ error: 'server' }));
        if (res.error) toast(errorText(res.error), { tone: 'bad', ms: 5000 });
        else toast(t('talk.name-is-blocked', { name }), { ms: 3000 });
      } },
      { label: t('talk.report'), run: () => setReporting({ person: r.from, requestId: r.id }) },
    ]);
  };
  // A group, from its row: who is in it, and leaving it without opening it first.
  const leaveGroup = async (g) => {
    if (!(await asked(t('talk.groups.leave-this-group'), leaveGroupText(g, state?.me?.id), t('talk.groups.leave-group')))) return;
    const res = await shell.talk.groupLeave({ groupId: g.id }).catch(() => ({ error: 'offline' }));
    if (res.error) toast(refusalText(res.error, res), { tone: 'bad', ms: 5000 });
    else if (currentRef.current === groupKey(g.id)) setCurrent(null);
  };
  const groupMenu = (e, g) => menu.open(e, [
    { label: t('common.people'), run: () => {
      if (currentRef.current === groupKey(g.id)) { setPeopleOpen(true); return; }
      peopleWanted.current = true;
      setCurrent(groupKey(g.id));
    } },
    { label: t('talk.groups.leave-group'), run: () => leaveGroup(g) },
  ]);
  const withdraw = async (r) => {
    const res = await shell.talk.withdraw({ requestId: r.id }).catch(() => ({ error: 'server' }));
    if (res.error) toast(errorText(res.error), { tone: 'bad', ms: 5000 });
    else toast(t('talk.invitation-email', { email: r.email }), { ms: 3000 });
  };
  const remind = async (r) => {
    const res = await shell.talk.remind({ requestId: r.id, letter: againLetter(state.me) }).catch(() => ({ error: 'server' }));
    if (res.error === 'too-soon') toast(t('talk.you-can-invite-them-again-when', { when: formatWhen(res.next) }), { ms: 5000 });
    else if (res.error === 'too-many') toast(t('talk.invitation-sent-again'), { ms: 5000 });
    else if (res.error === 'no-request') toast(t('talk.invitation-no-longer'), { ms: 5000 });
    else if (res.error) toast(errorText(res.error), { tone: 'bad', ms: 5000 });
    else toast([t('talk.invited-email-again', { email: r.email }), toldText(res.told)].join(' '), { tone: res.told?.sent || res.told?.by === 'none' ? 'good' : 'warn', ms: 5000 });
  };
  const copyLink = useCallback(async (link) => {
    await shell.clipboard.writeText({ text: link }).catch(() => {});
    toast(t('talk.link-copied-paste'), { tone: 'good', ms: 2500 });
  }, [shell, toast]);

  // `current` is the conversation shown: a contact's id, or a group's key.
  const loadHistory = useCallback((key) => {
    if (!key) { setHistory([]); return; }
    const groupId = groupIdOf(key);
    shell.talk.history(groupId ? { groupId } : { contactId: key }).then(setHistory).catch(() => setHistory([]));
  }, [shell]);
  useEffect(() => loadHistory(current), [current, loadHistory]);
  // Another conversation closes the list of a group's people, unless the group was opened to show them.
  const peopleWanted = useRef(false);
  useEffect(() => { setPeopleOpen(peopleWanted.current); peopleWanted.current = false; }, [current]);

  // A conversation in front of the person is read: its count goes. Not while a call covers it, nor in a window behind others.
  const unreadHere = current && !call ? state?.unread?.[current] || 0 : 0;
  useEffect(() => {
    if (!unreadHere) return undefined;
    const read = () => { if (document.hasFocus()) shell.talk.read({ conversation: current }).catch(() => {}); };
    read();
    window.addEventListener('focus', read);
    return () => window.removeEventListener('focus', read);
  }, [current, unreadHere, shell]);

  // While this window holds a call or a meeting it is no source for a screenshot, a recording or a broadcast,
  // and a call or a meeting not set up to be streamed or recorded is kept out of a capture of the screen too.
  const hold = withholdOf({ call: call?.state || null, room: Boolean(room), meeting: room?.meeting || null });
  useEffect(() => { Promise.resolve(shell.capture?.withhold?.({ on: hold.on, protect: hold.protect })).catch(() => {}); }, [shell, hold.on, hold.protect]);
  useEffect(() => () => { Promise.resolve(shell.capture?.withhold?.({ on: false })).catch(() => {}); }, [shell]);

  // A call already ringing when this window opened (Talk connects as Office starts): shown, and answered, as one heard here.
  useEffect(() => {
    let gone = false;
    shell.talk.ringing?.().then((calls) => { if (!gone && calls?.[0] && !callRef.current && !roomRef.current) setRinging((r) => r || calls[0]); }).catch(() => {});
    return () => { gone = true; };
  }, [shell]);

  // Talk's sounds: a ring while a call comes in, a soft tone while one's own call rings at the other side, one short
  // tone for being asked into a meeting. None in a build made for the checks, and none when the person turned them off.
  const ringer = useMemo(() => createRinger(), []);
  useEffect(() => () => ringer.close(), [ringer]);
  const sounds = Boolean(state) && state.ring !== false && !state.testServers;
  const sound = soundFor({ on: sounds, incoming: Boolean(ringing), call: call?.state || null });
  useEffect(() => { if (sound) ringer.play(sound); else ringer.stop(); }, [ringer, sound]);
  const askedInto = state?.meetings?.[0]?.meeting?.id || null;
  useEffect(() => { if (askedInto && sounds && !ringing && !callRef.current && !roomRef.current) ringer.once('chime'); }, [ringer, askedInto, sounds]);

  const track = useCallback((c) => { setCall(c); setCallTick((n) => n + 1); if (c.state === 'ended') setTimeout(() => setCall((x) => (x === c ? null : x)), 1500); }, []);

  // What the introducer says: messages, and a call's set-up.
  useEffect(() => shell.on('talk:event', async (msg) => {
    const c = callRef.current;
    if (msg.type === 'welcome' || msg.type.startsWith('meet.')) roomRef.current?.handle(msg);
    // Anything that changes what a conversation holds: a message, its marks, a missed call, a change in a group.
    if (HOLDS.has(msg.type) || (msg.type === 'error' && /\.send$/.test(msg.for || ''))) {
      if (currentRef.current) loadHistory(currentRef.current);
      if (!msg.type.startsWith('call.')) return;
    }
    if (msg.type === 'call.incoming') {
      // On air in Live Broadcast: turned away as busy, and the person told who called (a missed call kept).
      const onAir = await onAirNow(shell);
      // Read again after that question: another call may ring by now, or this one be over.
      if (overRef.current.has(msg.callId)) return;
      const what = incomingCall({ callId: msg.callId, call: callRef.current?.state || null, room: Boolean(roomRef.current), ringing: ringRef.current?.callId || null, onAir });
      if (what === 'busy' && onAir) toast(t('talk.name-called-while', { name: msg.from?.name || msg.from?.email || t('common.someone') }), { ms: 8000 });
      // Turned away as busy it is a missed call: its line is in the conversation, shown at once when that one is open.
      if (what === 'busy') shell.talk.send({ type: 'call.decline', callId: msg.callId, reason: 'busy' }).then(() => { if (currentRef.current) loadHistory(currentRef.current); }).catch(() => {});
      if (what === 'ring') { ringRef.current = msg; setRinging(msg); }
      return;
    }
    if (msg.type === 'call.taken' || msg.type === 'call.ended') {
      const over = overRef.current;
      over.add(msg.callId);
      if (over.size > 50) over.delete(over.values().next().value);
      // Its card goes, read from the ref: the ringing this handler was made with can be a render old.
      if (ringRef.current?.callId === msg.callId) ringRef.current = null;
      setRinging((r) => (r?.callId === msg.callId ? null : r));
    }
    if (!c || msg.callId !== c.callId) return;
    if (msg.type === 'call.answered') await c.answered(msg.sdp);
    else if (msg.type === 'call.candidate') await c.candidate(msg.candidate);
    else if (msg.type === 'call.declined') { toast(msg.reason === 'busy' ? t('talk.name-another-call', { name: c.peer.name }) : t('talk.name-did-not', { name: c.peer.name }), { ms: 4000 }); c.end('declined', { tell: false }); }
    else if (msg.type === 'call.unavailable') { toast(t('talk.name-not-available', { name: c.peer.name }), { ms: 4000 }); c.end('unavailable', { tell: false }); }
    else if (msg.type === 'call.ended') c.end(msg.reason || 'hung-up', { tell: false });
    else if (msg.type === 'error' && msg.callId === c.callId) { toast(t('talk.call-could-not'), { tone: 'bad', ms: 5000 }); c.end('failed', { tell: false }); }
  }), [shell, toast, loadHistory, ringing]);

  const startCall = useCallback(async (contact, video) => {
    const c = new Call({ shell, callId: callId(), peer: contact, video, outgoing: true, onChange: track });
    track(c);
    try {
      await c.place();
    } catch (err) {
      toast(err?.message || String(err), { tone: 'bad', ms: 6000 });
      c.end('failed', { tell: false });
    }
  }, [shell, track, toast]);

  // Into a meeting: this window holds it until the person leaves, is taken out, or it ends.
  const enterMeeting = useCallback(async (meeting, { video = false, muted = false, camera = null } = {}) => {
    // Not coming in after all: the camera they looked at is let go.
    const drop = () => { try { camera?.stop(); } catch { /* stopped */ } };
    if (roomRef.current) { drop(); return; }
    if (callRef.current && callRef.current.state !== 'ended') { drop(); toast(t('talk.finish-call-first'), { ms: 4000 }); return; }
    const r = new MeetingRoom({
      shell,
      me: { id: state?.me?.id, name: state?.me?.name || '' },
      onChange: (x) => {
        setCallTick((n) => n + 1);
        if (x.state !== 'gone' || roomRef.current !== x) return;
        // A meeting held on a Talk Server that could not be started, come into or stayed in: why, in a sentence.
        if (x.why === 'server') toast(serverFaultText(x.fault, x.meeting?.starter), { tone: 'bad', ms: 9000 });
        const words = goneText(x.why, x.by);
        if (words) toast(words, { ms: 5000 });
        setRoom(null);
      },
      onMuted: (x) => toast(t('talk.name-muted-unmute', { name: x.meeting?.starter?.name || t('common.someone') }), { ms: 5000 }),
    });
    setRoom(r);
    roomRef.current = r;
    const res = await r.join(meeting.id, { video, muted, camera });
    // Given up while still coming in: nothing to say of it.
    if (res.error) { if (res.error !== 'left') toast(meetRefusal(res.error, res), { tone: 'bad', ms: 6000 }); setRoom((x) => (x === r ? null : x)); }
    else if (meeting.groupId) setCurrent(groupKey(meeting.groupId));
  }, [shell, state, toast]);
  // A meeting set up to be streamed or recorded, taken by the person who set it up to Live Broadcast or the
  // Recorder: this window leaves it and that one comes in, with the others in it as its guests. Asked
  // first; this window leaves only once the other has opened, so a window that cannot open never
  // leaves the meeting with nobody organising it.
  const toStudio = useCallback((r, app) => { if (r?.id) setStudioAsk({ room: r, app }); }, []);
  const goStudio = useCallback(async () => {
    const ask = studioAsk;
    setStudioAsk(null);
    if (!ask?.room?.id) return;
    try {
      await shell.win.create({ app: ask.app, query: { meet: ask.room.id } });
      ask.room.leave();
    } catch (err) {
      toast(t('talk.meeting-stays-reason', { reason: err?.message || String(err) }), { tone: 'bad', ms: 7000 });
    }
  }, [shell, toast, studioAsk]);
  // The window closing takes this person out of the meeting, so nobody is left looking at someone who has gone.
  useEffect(() => {
    const out = () => roomRef.current?.leave();
    window.addEventListener('pagehide', out);
    return () => { window.removeEventListener('pagehide', out); out(); };
  }, []);

  const accept = useCallback(async (video) => {
    const r = ringing;
    setRinging(null);
    if (!r) return;
    const peer = state?.contacts.find((x) => x.id === r.from.id) || r.from;
    setCurrent(peer.id);
    const c = new Call({ shell, callId: r.callId, peer, video, outgoing: false, onChange: track });
    track(c);
    try {
      await c.answer({ sdp: r.sdp, relay: r.relay, video });
    } catch (err) {
      toast(err?.message || String(err), { tone: 'bad', ms: 6000 });
      c.end('failed');
    }
  }, [ringing, state, shell, track, toast]);

  const decline = useCallback(() => {
    if (ringing) shell.talk.send({ type: 'call.decline', callId: ringing.callId }).catch(() => {});
    setRinging(null);
  }, [ringing, shell]);

  /** Share: the first screen, chosen the way the recorder chooses one. */
  const share = useCallback(async (c) => {
    if (c.screen) { c.share(null); return; }
    try {
      const sources = (await shell.capture.sources()) || [];
      const screen = sources.find((s) => s.kind === 'screen') || sources[0];
      if (!screen) throw new Error(t('talk.no-screen-share'));
      await shell.capture.choose({ id: screen.id, audio: false });
      const stream = await navigator.mediaDevices.getDisplayMedia({ video: { frameRate: 15 }, audio: false });
      c.share(stream);
    } catch (err) {
      toast(err?.message || String(err), { tone: 'bad', ms: 5000 });
    }
  }, [shell, toast]);

  const answerRequest = useCallback((r, how) => {
    shell.talk.send(how === 'accept' ? { type: 'contact.accept', requestId: r.id } : { type: 'contact.decline', requestId: r.id }).catch((err) => toast(err.message, { tone: 'bad' }));
  }, [shell, toast]);

  if (!state) return <AppFrame app={app} shell={shell} title={t('common.talk')} menu={appMenu}><Content><div className="tk-wait"><Spinner /></div></Content></AppFrame>;

  const contact = state.contacts.find((c) => c.id === current) || null;
  const group = (state.groups || []).find((g) => groupKey(g.id) === current) || null;
  // Signing in in the browser is still the sign-in screen, waiting.
  const signedIn = state.phase !== 'signed-out' && state.phase !== 'signing-in';
  const presence = state.presence || 'online';

  return (
    <AppFrame
      app={app}
      shell={shell}
      title={room?.meeting ? room.meeting.name : call && call.state !== 'ended' ? t('talk.call-with-name', { name: call.peer.name }) : titleWithCount(t('common.talk'), waitingIn(state))}
      menu={appMenu}
      status={<span className="tk-statusbar">{signedIn ? (state.phase === 'online' ? t('talk.signed-name', { name: state.me?.name || state.me?.email || '' }) : state.phase === 'offline' ? t('talk.reconnecting') : t('common.connecting')) : t('talk.signed-out')}</span>}
    >
      <style>{TALK_CSS}</style>
      {menu.node}
      {invitation ? <Invitation shell={shell} link={invitation} online={state.phase === 'online'} toast={toast} onDone={endInvitation} onAccepted={(c) => c && setCurrent(c.id)} /> : null}
      {!signedIn ? <Content><SignIn shell={shell} state={state} toast={toast} /></Content> : (
        <>
          <Panel name="people" width={280} title={state.me?.name || t('common.talk')}
            right={(
              <Select value={presence} className="tk-presence" aria-label={t('talk.your-status')} onChange={(e) => shell.talk.send({ type: 'presence', presence: e.target.value }).catch(() => {})}>
                {PRESENCE().map(([k, label]) => <option key={k} value={k}>{label}</option>)}
              </Select>
            )}
          >
            <People state={state} current={current} onPick={setCurrent} onAdd={() => setAdding(true)} onAnswerRequest={answerRequest} onRequestMenu={requestMenu} onCopy={copyLink} onContactMenu={contactMenu} onGroupMenu={groupMenu} onWithdraw={withdraw} onRemind={remind} onMissed={() => setMissedOpen(true)} onNewGroup={() => setGroupNew(true)} onMeet={() => setMeetSetup({})} />
            <div className="tk-invite-actions">
              <InviteByEmailButton onClick={() => setEmailing(true)} />
              <Button ghost icon="link" label={t('talk.invite-link')} className="tk-links-btn" onClick={() => setLinksOpen(true)} />
              <Button ghost label={t('talk.open-an')} className="tk-paste-btn" onClick={() => setPasting(true)} />
            </div>
            <div className="tk-foot">
              <MyAddressButton onClick={() => setAddressOpen(true)} />
              <Button ghost icon="settings" label={t('talk.relays-2')} className="tk-relays-btn" onClick={() => setRelaysOpen(true)} />
              <Button ghost label={t('common.settings')} className="tk-settings-btn" onClick={() => setSettingsOpen(true)} />
              <Button ghost label={t('common.sign-out')} className="tk-signout" onClick={() => shell.talk.signOut().catch(() => {})} />
            </div>
          </Panel>
          <Content>
            {room ? <MeetingView shell={shell} room={room} toast={toast} asked={asked} onShare={share} onStudio={toStudio} />
              : call ? <CallView shell={shell} call={call} onShare={share} />
              : contact ? <Conversation shell={shell} contact={contact} me={state.me} history={history} online={state.phase === 'online'} onCall={startCall} onMore={contactMenu} onJoinMeeting={setMeetAsk} onChanged={() => loadHistory(currentRef.current)} toast={toast} />
                : group ? <Conversation shell={shell} group={group} me={state.me} history={history} online={state.phase === 'online'} onChanged={() => loadHistory(currentRef.current)} onPeople={() => setPeopleOpen(true)} onMeet={() => setMeetSetup({ group: { id: group.id, title: groupTitle(group, state.me?.id) } })} onJoinMeeting={setMeetAsk} toast={toast} />
                : <div className="tk-empty"><Empty icon="talk" title={t('talk.choose-someone-talk')}>{t('talk.messages-calls-go')}</Empty></div>}
          </Content>
        </>
      )}
      {ringing ? (
        <div className="tk-ring" role="alertdialog" aria-label={t('talk.incoming-call')} data-sound={sound === 'ring' ? 'on' : 'off'}>
          <Avatar person={ringing.from} size={44} />
          <span className="tk-who"><strong>{ringing.from.name || ringing.from.email}</strong><small>{ringing.video ? t('talk.video-call') : t('talk.call')}</small></span>
          <Button icon="phone" label={t('common.answer')} className="tk-answer" onClick={() => accept(false)} />
          {ringing.video ? <Button primary icon="camera" label={t('talk.with-video')} className="tk-answer-video" onClick={() => accept(true)} /> : null}
          <Button label={t('common.decline')} className="tk-ring-decline" onClick={decline} />
        </div>
      ) : null}
      {adding ? <AddContact shell={shell} state={state} toast={toast} onClose={() => setAdding(false)} /> : null}
      {emailing && signedIn ? <InviteByEmail key={Array.isArray(emailing) ? emailing[0].email : 'own'} shell={shell} state={state} toast={toast} initial={Array.isArray(emailing) ? emailing : []} onClose={() => setEmailing(false)} /> : null}
      {linksOpen ? <InviteLinks shell={shell} state={state} toast={toast} onClose={() => setLinksOpen(false)} /> : null}
      {pasting ? <OpenInvitation onOpen={(link) => (meetingCode(link) ? setMeetAsk(meetingCode(link)) : setInvitation(link))} onClose={() => setPasting(false)} /> : null}
      {/* Asked into a meeting that is on now; one at a time, and not while this window is in a meeting or a call. */}
      {signedIn && !room && !call && !ringing && !meetAsk && state.meetings?.length ? <MeetCard asked={state.meetings[0]} onJoin={() => setMeetAsk(state.meetings[0].meeting.id)} onDismiss={() => shell.talk.meetDismiss({ id: state.meetings[0].meeting.id }).catch(() => {})} /> : null}
      {meetSetup && signedIn && !room ? <MeetSetup shell={shell} state={state} group={meetSetup.group || null} toast={toast} onClose={() => setMeetSetup(null)} onStarted={enterMeeting} /> : null}
      {meetAsk && signedIn && !room ? <MeetJoin shell={shell} link={meetAsk} online={state.phase === 'online'} me={state.me} onClose={() => setMeetAsk(null)} onJoin={enterMeeting} onSeen={(r) => { if (r?.error === 'no-meeting') loadHistory(currentRef.current); }} /> : null}
      {addressOpen && signedIn ? <MyAddress shell={shell} state={state} toast={toast} onClose={() => setAddressOpen(false)} /> : null}
      {relaysOpen ? <RelaysDialog shell={shell} toast={toast} onClose={() => setRelaysOpen(false)} /> : null}
      {settingsOpen && signedIn ? <TalkSettings shell={shell} state={state} toast={toast} onClose={() => setSettingsOpen(false)} /> : null}
      {groupNew && signedIn ? <NewGroup shell={shell} state={state} toast={toast} onClose={() => setGroupNew(false)} onMade={(g) => setCurrent(groupKey(g.id))} /> : null}
      {peopleOpen && group ? <GroupPeople shell={shell} state={state} group={group} toast={toast} asked={asked} onClose={() => setPeopleOpen(false)} onReport={(m) => setReporting({ person: m, groupId: group.id })} /> : null}
      {missedOpen && signedIn ? <MissedCalls shell={shell} state={state} onClose={() => setMissedOpen(false)} onOpen={(c) => setCurrent(c.id)} onCallBack={(c, video) => { setCurrent(c.id); startCall(c, video); }} /> : null}
      {studioAsk ? (
        <Dialog
          title={studioAsk.app === 'broadcast' ? t('talk.meeting.open-live-broadcast') : t('talk.meeting.open-recorder')}
          className="tk-studio-ask"
          onClose={() => setStudioAsk(null)}
          actions={<><Button label={t('common.cancel')} className="tk-studio-no" onClick={() => setStudioAsk(null)} /><Button primary className="tk-studio-go" label={studioAsk.app === 'broadcast' ? t('talk.meeting.open-live-broadcast') : t('talk.meeting.open-recorder')} onClick={goStudio} /></>}
        >
          <p>{studioAsk.app === 'broadcast' ? t('talk.meeting.people-become-guests') : t('talk.meeting.people-become-guests-recording')}</p>
          <p>{t('talk.window-leaves-meeting')}</p>
        </Dialog>
      ) : null}
      {reporting ? <Report shell={shell} person={reporting.person} groupId={reporting.groupId || null} requestId={reporting.requestId || null} toast={toast} onClose={() => setReporting(null)} onBlocked={(p) => { if (currentRef.current === p.id) setCurrent(null); }} /> : null}
    </AppFrame>
  );
}

function RelaysDialog({ shell, toast, onClose }) {
  const [info, setInfo] = useState(null);
  const [own, setOwn] = useState({ urls: '', username: '', credential: '' });
  useEffect(() => { shell.talk.relays().then((r) => { setInfo(r); if (r.own) setOwn({ urls: r.own.urls.join(' '), username: r.own.username, credential: '' }); }).catch(() => {}); }, [shell]);
  const save = async (choice) => {
    try {
      const args = { choice };
      if (choice === 'own') args.own = { urls: own.urls.split(/\s+/).filter(Boolean), username: own.username, credential: own.credential || undefined };
      await shell.talk.setRelay(args);
      setInfo(await shell.talk.relays());
      toast(t('talk.calls-use-relay'), { tone: 'good', ms: 3000 });
    } catch (err) {
      toast(err?.message || String(err), { tone: 'bad', ms: 5000 });
    }
  };
  return (
    <Dialog title={t('talk.relays')} width={520} onClose={onClose} actions={<Button primary label={t('common.done')} onClick={onClose} />}>
      <p className="tk-lead">{t('talk.call-goes-straight')}</p>
      {!info ? <Spinner /> : (
        <div className="tk-relay-list" role="radiogroup" aria-label={t('talk.relays')}>
          <label className="tk-relay"><input type="radio" name="relay" checked={info.choice === 'auto'} onChange={() => save('auto')} /><span><strong>{t('talk.choose-for-me')}</strong><small>{t('talk.nearest-rutbas-relays')}</small></span></label>
          {info.relays.filter((r) => r.urls.some((u) => /^turns?:/.test(u))).map((r) => (
            <label key={r.id} className="tk-relay"><input type="radio" name="relay" checked={info.choice === r.id} onChange={() => save(r.id)} /><span><strong>{r.name}</strong><small>{r.region || r.urls[0]}</small></span></label>
          ))}
          <label className="tk-relay"><input type="radio" name="relay" checked={info.choice === 'own'} onChange={() => save('own')} disabled={!own.urls.trim()} /><span><strong>{t('talk.my-own-relay')}</strong><small>{t('talk.turn-server-run')}</small></span></label>
          <div className="tk-own">
            <Input className="tk-own-urls" value={own.urls} placeholder={RELAY_SHAPE} aria-label={t('talk.relay-address')} onChange={(e) => setOwn({ ...own, urls: e.target.value })} />
            <Input className="tk-own-user" value={own.username} placeholder={t('common.user-name')} aria-label={t('common.user-name')} onChange={(e) => setOwn({ ...own, username: e.target.value })} />
            <Input className="tk-own-pass" type="password" value={own.credential} placeholder={info.own?.credential ? t('talk.kept') : t('common.password')} aria-label={t('common.password')} onChange={(e) => setOwn({ ...own, credential: e.target.value })} />
            <Button label={t('talk.use-my-relay')} className="tk-own-save" disabled={!own.urls.trim()} onClick={() => save('own')} />
          </div>
        </div>
      )}
    </Dialog>
  );
}
