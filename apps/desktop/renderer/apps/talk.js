// Rutba Talk: messages and calls with the people you are in contact with.
//
// Kept plain on purpose (owner, 2026-10-21: "for chats and conference call
// etc build clean simple user interface do not complicate them like Ms
// teams"): the people on the left, the conversation on the right, Call and
// Video above it, and a call that fills the conversation with four buttons.
//
// Someone is a contact once both of you have said so: you ask by their
// address, they accept. Only contacts can call or write to each other.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Button, Icon, Input, Select, Panel, Content, Empty, Dialog, Spinner, useToast, useMenu, formatWhen, t } from '@rutba/office-ui';
import { AppFrame, useAppMenu } from '../shell.js';
import { Call } from './talk/call.js';
import { TALK_CSS } from './talk/css.js';

const PRESENCE = () => [['online', t('Available')], ['busy', t('Busy')], ['away', t('Away')], ['hidden', t('Appear offline')]];
// The shapes of an address, shown in empty fields: examples, not words to translate.
const SERVER_SHAPE = 'wss://introducer.example/ws';
const RELAY_SHAPE = 'turn:relay.example.com:3478';
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

/** A stream shown in a video element, kept in step as tracks come and go. */
function Screen({ stream, muted = false, className }) {
  const ref = useRef(null);
  useEffect(() => {
    if (ref.current && ref.current.srcObject !== stream) ref.current.srcObject = stream || null;
  });
  return <video ref={ref} className={className} autoPlay playsInline muted={muted} />;
}

/* ── signed out ───────────────────────────────────────────────────── */

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
    <div className="tk-signin">
      <Empty icon="talk" title={t('Talk with your contacts')}>
        <p>{t('Messages and calls with the people you know. Sign in with your Rutba account; someone becomes a contact once both of you have said so.')}</p>
        {state.error && !waiting ? <p className="tk-error">{state.error}</p> : null}
        <Button primary label={waiting ? t('Waiting for the browser…') : t('Sign in with Rutba')} className="tk-rutba" disabled={waiting} onClick={signInWithRutba} />
        <p className="tk-soon">{waiting ? t('Finish signing in in your browser, then come back here.') : t('Your browser opens at auth.rutba.io to sign in, as everywhere on Rutba.')}</p>
      </Empty>
      <details className="tk-test">
        <summary>{t('Connect to a test server')}</summary>
        <label className="tk-field"><span>{t('Server')}</span><Input className="tk-server" value={server} placeholder={SERVER_SHAPE} onChange={(e) => setServer(e.target.value)} /></label>
        <label className="tk-field"><span>{t('Sign-in token')}</span><Input className="tk-token" value={token} onChange={(e) => setToken(e.target.value)} /></label>
        <Button primary label={busy ? t('Connecting…') : t('Connect')} className="tk-connect" disabled={busy || !token.trim()} onClick={connect} />
      </details>
    </div>
  );
}

/* ── the people on the left ───────────────────────────────────────── */

function People({ state, current, onPick, onAdd, onAnswerRequest }) {
  const { contacts, requests } = state;
  return (
    <div className="tk-people">
      {requests.incoming.length ? (
        <section className="tk-section">
          <h3>{t('Asking to be in contact')}</h3>
          {requests.incoming.map((r) => (
            <div key={r.id} className="tk-request" data-request={r.id}>
              <Avatar person={r.from} size={28} />
              <span className="tk-who"><strong>{r.from?.name || r.from?.email}</strong>{r.note ? <small>{r.note}</small> : <small>{r.from?.email}</small>}</span>
              <Button primary label={t('Accept')} className="tk-accept" onClick={() => onAnswerRequest(r, 'accept')} />
              <Button icon="close" title={t('Decline')} className="tk-decline" onClick={() => onAnswerRequest(r, 'decline')} />
            </div>
          ))}
        </section>
      ) : null}
      <section className="tk-section">
        <div className="tk-section-head">
          <h3>{t('Contacts')}</h3>
          <Button ghost icon="plus" label={t('Add')} className="tk-add" onClick={onAdd} />
        </div>
        {contacts.length ? contacts.map((c) => (
          <button key={c.id} type="button" className={`tk-contact${current === c.id ? ' on' : ''}`} data-contact={c.id} onClick={() => onPick(c.id)}>
            <Avatar person={c} />
            <span className="tk-who"><strong>{c.name || c.email}</strong><small>{c.presence === 'offline' ? t('Offline') : PRESENCE().find(([k]) => k === c.presence)?.[1] || ''}</small></span>
          </button>
        )) : <p className="tk-none">{t('No contacts yet. Add someone by their email address; once they accept, you can write and call.')}</p>}
        {requests.outgoing.length ? (
          <div className="tk-waiting">
            {requests.outgoing.map((r) => <small key={r.id} data-asked={r.email}>{t('Asked {email}', { email: r.email })}</small>)}
          </div>
        ) : null}
      </section>
    </div>
  );
}

/* ── the conversation ─────────────────────────────────────────────── */

function Conversation({ shell, contact, me, history, onCall, toast }) {
  const [text, setText] = useState('');
  const end = useRef(null);
  useEffect(() => { end.current?.scrollIntoView({ block: 'end' }); }, [history.length, contact?.id]);
  const send = async () => {
    const words = text.trim();
    if (!words) return;
    setText('');
    try {
      await shell.talk.message({ to: contact.id, text: words });
    } catch (err) {
      setText(words);
      toast(err?.message || String(err), { tone: 'bad', ms: 5000 });
    }
  };
  const reachable = contact.presence && contact.presence !== 'offline';
  return (
    <div className="tk-conversation">
      <header className="tk-head">
        <Avatar person={contact} size={36} />
        <span className="tk-who"><strong>{contact.name || contact.email}</strong><small>{contact.email}</small></span>
        <span className="tk-grow" />
        <Button icon="phone" label={t('Call')} className="tk-call-voice" disabled={!reachable} title={reachable ? '' : t('Calls reach a contact who is online')} onClick={() => onCall(contact, false)} />
        <Button primary icon="camera" label={t('Video')} className="tk-call-video" disabled={!reachable} onClick={() => onCall(contact, true)} />
      </header>
      <div className="tk-messages" role="log" aria-live="polite">
        {history.length ? history.map((m) => (
          <div key={m.id} className={`tk-bubble${m.mine ? ' mine' : ''}`} data-message={m.id}>
            <span className="tk-text">{m.text}</span>
            <time>{formatWhen(m.at)}</time>
          </div>
        )) : <p className="tk-none">{t('Say hello to {name}.', { name: contact.name || contact.email })}</p>}
        <div ref={end} />
      </div>
      <form className="tk-compose" onSubmit={(e) => { e.preventDefault(); send(); }}>
        <textarea
          className="tk-input"
          rows={1}
          value={text}
          placeholder={t('Write a message')}
          aria-label={t('Write a message')}
          onChange={(e) => setText(e.target.value)}
          onKeyDown={(e) => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); send(); } }}
        />
        <Button primary icon="send" title={t('Send')} className="tk-send" disabled={!text.trim()} onClick={send} />
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
  const words = { calling: t('Calling…'), ringing: t('Ringing…'), connecting: t('Connecting…'), connected: `${Math.floor(secs / 60)}:${String(secs % 60).padStart(2, '0')}`, ended: t('Call ended') };
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
        <button type="button" className={`tk-round${call.muted ? ' off' : ''}`} aria-pressed={call.muted} title={call.muted ? t('Unmute') : t('Mute')} onClick={() => call.setMuted(!call.muted)}><Icon name={call.muted ? 'micOff' : 'mic'} size={20} /></button>
        <button type="button" className={`tk-round${call.camera ? '' : ' off'}`} aria-pressed={!call.camera} title={call.camera ? t('Turn camera off') : t('Turn camera on')} onClick={() => call.setCamera(!call.camera)}><Icon name="camera" size={20} /></button>
        <button type="button" className={`tk-round${call.screen ? ' on' : ''}`} aria-pressed={Boolean(call.screen)} title={call.screen ? t('Stop sharing') : t('Share screen')} onClick={() => onShare(call)}><Icon name="screen" size={20} /></button>
        <button type="button" className="tk-round end" title={t('Hang up')} onClick={() => call.end('hung-up')}><Icon name="phone" size={20} /></button>
      </div>
      {call.code ? <p className="tk-code" data-code={call.code}>{t('Safety code {code}: if {name} sees the same, nobody is listening in.', { code: call.code, name: call.peer.name })}</p> : null}
    </div>
  );
}

/* ── the window ───────────────────────────────────────────────────── */

export default function Talk({ app, shell }) {
  const toast = useToast();
  const [state, setState] = useState(null);
  const [current, setCurrent] = useState(null);
  const [history, setHistory] = useState([]);
  const [adding, setAdding] = useState(false);
  const [relaysOpen, setRelaysOpen] = useState(false);
  const [call, setCall] = useState(null);
  const [ringing, setRinging] = useState(null);
  const [, setCallTick] = useState(0);
  const callRef = useRef(null);
  callRef.current = call;
  const currentRef = useRef(null);
  currentRef.current = current;
  const menu = useMenu();
  const appMenu = useAppMenu({ shell, appKey: 'talk' });

  useEffect(() => {
    shell.talk.resume().then(setState).catch(() => {});
    return shell.on('talk:state', setState);
  }, [shell]);

  const loadHistory = useCallback((id) => {
    if (!id) { setHistory([]); return; }
    shell.talk.history({ contactId: id }).then(setHistory).catch(() => setHistory([]));
  }, [shell]);
  useEffect(() => loadHistory(current), [current, loadHistory]);

  const track = useCallback((c) => { setCall(c); setCallTick((n) => n + 1); if (c.state === 'ended') setTimeout(() => setCall((x) => (x === c ? null : x)), 1500); }, []);

  // What the introducer says: messages, and a call's set-up.
  useEffect(() => shell.on('talk:event', async (msg) => {
    const c = callRef.current;
    if (msg.type === 'messages' || msg.type === 'msg.copy' || msg.type === 'msg.sent') {
      if (currentRef.current) loadHistory(currentRef.current);
      return;
    }
    if (msg.type === 'call.incoming') {
      if (c && c.state !== 'ended') { shell.talk.send({ type: 'call.decline', callId: msg.callId, reason: 'busy' }).catch(() => {}); return; }
      setRinging(msg);
      return;
    }
    if (msg.type === 'call.taken' || (msg.type === 'call.ended' && ringing?.callId === msg.callId)) { setRinging((r) => (r?.callId === msg.callId ? null : r)); }
    if (!c || msg.callId !== c.callId) return;
    if (msg.type === 'call.answered') await c.answered(msg.sdp);
    else if (msg.type === 'call.candidate') await c.candidate(msg.candidate);
    else if (msg.type === 'call.declined') { toast(msg.reason === 'busy' ? t('{name} is on another call.', { name: c.peer.name }) : t('{name} did not take the call.', { name: c.peer.name }), { ms: 4000 }); c.end('declined', { tell: false }); }
    else if (msg.type === 'call.unavailable') { toast(t('{name} is not available.', { name: c.peer.name }), { ms: 4000 }); c.end('unavailable', { tell: false }); }
    else if (msg.type === 'call.ended') c.end(msg.reason || 'hung-up', { tell: false });
    else if (msg.type === 'error' && msg.callId === c.callId) { toast(t('The call could not be made.'), { tone: 'bad', ms: 5000 }); c.end('failed', { tell: false }); }
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
      if (!screen) throw new Error(t('There is no screen to share.'));
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

  if (!state) return <AppFrame app={app} shell={shell} title={t('Talk')} menu={appMenu}><Content><div className="tk-wait"><Spinner /></div></Content></AppFrame>;

  const contact = state.contacts.find((c) => c.id === current) || null;
  // Signing in in the browser is still the sign-in screen, waiting.
  const signedIn = state.phase !== 'signed-out' && state.phase !== 'signing-in';
  const presence = state.presence || 'online';

  return (
    <AppFrame
      app={app}
      shell={shell}
      title={call && call.state !== 'ended' ? t('Call with {name}', { name: call.peer.name }) : t('Talk')}
      menu={appMenu}
      status={<span className="tk-statusbar">{signedIn ? (state.phase === 'online' ? t('Signed in as {name}', { name: state.me?.name || state.me?.email || '' }) : state.phase === 'offline' ? t('Reconnecting…') : t('Connecting…')) : t('Signed out')}</span>}
    >
      <style>{TALK_CSS}</style>
      {menu.node}
      {!signedIn ? <Content><SignIn shell={shell} state={state} toast={toast} /></Content> : (
        <>
          <Panel name="people" width={280} title={state.me?.name || t('Talk')}
            right={(
              <Select value={presence} className="tk-presence" aria-label={t('Your status')} onChange={(e) => shell.talk.send({ type: 'presence', presence: e.target.value }).catch(() => {})}>
                {PRESENCE().map(([k, label]) => <option key={k} value={k}>{label}</option>)}
              </Select>
            )}
          >
            <People state={state} current={current} onPick={setCurrent} onAdd={() => setAdding(true)} onAnswerRequest={answerRequest} />
            <div className="tk-foot">
              <Button ghost icon="settings" label={t('Relays…')} className="tk-relays-btn" onClick={() => setRelaysOpen(true)} />
              <Button ghost label={t('Sign out')} className="tk-signout" onClick={() => shell.talk.signOut().catch(() => {})} />
            </div>
          </Panel>
          <Content>
            {call ? <CallView shell={shell} call={call} onShare={share} />
              : contact ? <Conversation shell={shell} contact={contact} me={state.me} history={history} onCall={startCall} toast={toast} />
                : <div className="tk-empty"><Empty icon="talk" title={t('Choose someone to talk to')}>{t('Messages and calls go only between contacts who have both said yes.')}</Empty></div>}
          </Content>
        </>
      )}
      {ringing ? (
        <div className="tk-ring" role="alertdialog" aria-label={t('Incoming call')}>
          <Avatar person={ringing.from} size={44} />
          <span className="tk-who"><strong>{ringing.from.name || ringing.from.email}</strong><small>{ringing.video ? t('Video call') : t('Call')}</small></span>
          <Button icon="phone" label={t('Answer')} className="tk-answer" onClick={() => accept(false)} />
          {ringing.video ? <Button primary icon="camera" label={t('With video')} className="tk-answer-video" onClick={() => accept(true)} /> : null}
          <Button label={t('Decline')} className="tk-ring-decline" onClick={decline} />
        </div>
      ) : null}
      {adding ? <AddContact shell={shell} toast={toast} onClose={() => setAdding(false)} /> : null}
      {relaysOpen ? <RelaysDialog shell={shell} toast={toast} onClose={() => setRelaysOpen(false)} /> : null}
    </AppFrame>
  );
}

function AddContact({ shell, toast, onClose }) {
  const [email, setEmail] = useState('');
  const [note, setNote] = useState('');
  const ask = async () => {
    try {
      await shell.talk.send({ type: 'contact.request', email: email.trim(), note: note.trim() });
      toast(t('Asked {email}. You can talk once they accept.', { email: email.trim() }), { tone: 'good', ms: 4000 });
      onClose();
    } catch (err) {
      toast(err?.message || String(err), { tone: 'bad', ms: 5000 });
    }
  };
  return (
    <Dialog title={t('Add a contact')} width={440} onClose={onClose} actions={<><Button label={t('Cancel')} onClick={onClose} /><Button primary label={t('Ask')} className="tk-ask" disabled={!/^\S+@\S+\.\S+$/.test(email.trim())} onClick={ask} /></>}>
      <p className="tk-lead">{t('They see your name and your note, and become a contact when they accept. Nobody is told whether an address has an account.')}</p>
      <label className="tk-field"><span>{t('Email address')}</span><Input className="tk-email" value={email} autoFocus onChange={(e) => setEmail(e.target.value)} /></label>
      <label className="tk-field"><span>{t('A note (optional)')}</span><Input className="tk-note" value={note} onChange={(e) => setNote(e.target.value)} /></label>
    </Dialog>
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
      toast(t('Calls use that relay when they cannot connect directly.'), { tone: 'good', ms: 3000 });
    } catch (err) {
      toast(err?.message || String(err), { tone: 'bad', ms: 5000 });
    }
  };
  return (
    <Dialog title={t('Relays')} width={520} onClose={onClose} actions={<Button primary label={t('Done')} onClick={onClose} />}>
      <p className="tk-lead">{t('A call goes straight between the two computers. When a network stops that, it goes through a relay, still encrypted end to end.')}</p>
      {!info ? <Spinner /> : (
        <div className="tk-relay-list" role="radiogroup" aria-label={t('Relays')}>
          <label className="tk-relay"><input type="radio" name="relay" checked={info.choice === 'auto'} onChange={() => save('auto')} /><span><strong>{t('Choose for me')}</strong><small>{t('The nearest of Rutba’s relays')}</small></span></label>
          {info.relays.filter((r) => r.urls.some((u) => /^turns?:/.test(u))).map((r) => (
            <label key={r.id} className="tk-relay"><input type="radio" name="relay" checked={info.choice === r.id} onChange={() => save(r.id)} /><span><strong>{r.name}</strong><small>{r.region || r.urls[0]}</small></span></label>
          ))}
          <label className="tk-relay"><input type="radio" name="relay" checked={info.choice === 'own'} onChange={() => save('own')} disabled={!own.urls.trim()} /><span><strong>{t('My own relay')}</strong><small>{t('A TURN server you run; offered to the other side for your calls')}</small></span></label>
          <div className="tk-own">
            <Input className="tk-own-urls" value={own.urls} placeholder={RELAY_SHAPE} aria-label={t('Relay address')} onChange={(e) => setOwn({ ...own, urls: e.target.value })} />
            <Input className="tk-own-user" value={own.username} placeholder={t('User name')} aria-label={t('User name')} onChange={(e) => setOwn({ ...own, username: e.target.value })} />
            <Input className="tk-own-pass" type="password" value={own.credential} placeholder={info.own?.credential ? t('Kept') : t('Password')} aria-label={t('Password')} onChange={(e) => setOwn({ ...own, credential: e.target.value })} />
            <Button label={t('Use my relay')} className="tk-own-save" disabled={!own.urls.trim()} onClick={() => save('own')} />
          </div>
        </div>
      )}
    </Dialog>
  );
}
