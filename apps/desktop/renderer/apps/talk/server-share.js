// Who else may hold meetings on a Talk Server of one's own (Office 2.27.0):
// a part of Talk's settings that only somebody with a server sees, and only
// where the Talk service says it knows a meeting held on one (the same rule
// as the switch on the set-up screen: meet-lines.js, serverChoice).
//
// There are four ways (the Talk server contract, "Naming people"): only the
// owner, the owner's contacts as they stand when a meeting is asked for, the
// people the owner chooses, or everybody in the owner's organisation. The
// people chosen are ticked among the owner's contacts here, because Talk's
// service keeps their account ids and no names; somebody chosen who is no
// longer a contact is shown as that, with the way to take them out.
//
// Nobody is told that they were chosen. Where they set a meeting up, their
// Talk offers to hold it on a Talk Server and never says whose.
//
// Nothing here names a server's address, and there is no place to type one.

import React, { useEffect, useRef, useState } from 'react';
import { Button, Select } from '@rutba/office-ui';
import { t, tn } from '../talk-words.js';
import { serverWhyText } from './meet-lines.js';

/** The four ways, in the order they are offered. */
export const WHO = ['me', 'contacts', 'named', 'org'];
/** The most people one server's list holds (the Talk service takes no more). */
export const NAMED_MOST = 500;

/** One way, as the choice names it. */
export function whoLabel(who) {
  switch (who) {
    case 'contacts': return t('talk.server-share.my-contacts');
    case 'named': return t('talk.server-share.people-choose');
    case 'org': return t('talk.server-share.everybody-organisation');
    default: return t('talk.server-share.only-me');
  }
}

/** What one way means, in a sentence; `count` is how many people are chosen. */
export function whoNote(who, count = 0) {
  switch (who) {
    case 'contacts': return t('talk.server-share.whoever-contact-moment');
    case 'named': return count
      ? tn(count, 'talk.server-share.count-person-chosen', { count })
      : t('talk.server-share.nobody-chosen-yet');
    case 'org': return t('talk.server-share.everybody-organisation-rutba');
    default: return t('talk.server-share.nobody-else-hold');
  }
}

/** A contact as a list names them. */
const nameOf = (c) => c?.name || c?.email || '';

/**
 * The people chosen on a server, for the chooser: the owner's contacts, each
 * with whether they are chosen (`on`), and the ids chosen that are no longer
 * among the contacts (`gone`), which have no name to show.
 */
export function namedList(named, contacts) {
  const chosen = new Set([].concat(named || []).map(String));
  const known = new Set();
  const people = [].concat(contacts || []).filter((c) => c?.id).map((c) => { known.add(String(c.id)); return { id: String(c.id), name: nameOf(c), email: c.email || '', on: chosen.has(String(c.id)) }; });
  return { people, gone: [...chosen].filter((id) => !known.has(id)) };
}

/** The list as it should stand with one person put in or taken out: it is sent whole. */
export function namedWith(named, id, on) {
  const rest = [...new Set([].concat(named || []).map(String))].filter((x) => x !== String(id));
  return on && rest.length < NAMED_MOST ? [...rest, String(id)] : rest;
}

/**
 * The owner's servers, each with who may hold meetings on it. `mine` is what
 * Talk's service said of them where the caller has it already (the checks
 * draw it so). Somebody with no server is shown nothing, and so is everybody
 * where the Talk service does not say it knows a meeting held on a server.
 */
const serversOf = (m) => (m && !m.error && m.knows === true && Array.isArray(m.servers) ? m.servers : []);

export function ServerShare({ shell, state, toast, mine = null }) {
  const [servers, setServers] = useState(() => (mine ? serversOf(mine) : null));
  const online = state.phase === 'online';
  const shown = useRef(true);
  /** Changes go one after another, so that the list sent last is the list kept. */
  const turn = useRef(Promise.resolve());
  const read = () => Promise.resolve(shell.talk.serverMine()).then((r) => { if (shown.current && !r?.error) setServers(serversOf(r)); }, () => {});
  useEffect(() => {
    shown.current = true;
    if (!mine && online && typeof shell.talk.serverMine === 'function') read();
    return () => { shown.current = false; };
  }, [shell, online]);
  if (!servers?.length) return null;

  const change = (s, patch) => {
    // Shown at once; if the Talk service does not take it, what it holds is read again and the person is told.
    setServers((all) => (all || []).map((x) => (x.id === s.id ? { ...x, ...patch } : x)));
    turn.current = turn.current.then(async () => {
      const r = await Promise.resolve(shell.talk.serverShare({ id: s.id, ...patch })).catch(() => ({ error: 'offline' }));
      if (!r?.error || !shown.current) return;
      toast(t('talk.server-share.talk-could-not', { name: s.name || t('talk.server-share.your-talk-server') }), { tone: 'bad', ms: 5000 });
      await read();
    });
  };

  return (
    <>
      <h3 className="tk-settings-head">{t('talk.server-share.your-talk-servers')}</h3>
      {servers.map((s) => {
        const who = WHO.includes(s.who) ? s.who : 'me';
        const { people, gone } = namedList(s.named, state.contacts);
        const count = people.filter((p) => p.on).length + gone.length;
        const full = count >= NAMED_MOST;
        return (
          <div key={s.id} className="tk-own-server" data-server={s.id} data-who={who}>
            <label className="tk-field">
              <span>{t('talk.server-share.may-hold-meetings', { name: s.name || t('talk.server-share.your-talk-server') })}</span>
              <Select className="rw-select tk-server-who" value={who} disabled={!online} title={online ? '' : t('talk.settings.talk-not-connected')} onChange={(e) => change(s, { who: e.target.value })}>
                {WHO.map((w) => <option key={w} value={w}>{whoLabel(w)}</option>)}
              </Select>
            </label>
            <p className="tk-lead tk-server-note">{whoNote(who, count)}</p>
            {s.canCarry ? null : <p className="tk-lead tk-server-why" data-why={s.why || 'no'}>{serverWhyText(s)}</p>}
            {who !== 'named' ? null : (
              <>
                {people.length ? (
                  <div className="tk-pick-list tk-server-people" role="group" aria-label={t('talk.server-share.people-choose')}>
                    {people.map((p) => (
                      <label key={p.id} className="tk-pick" data-pick={p.id}>
                        <input type="checkbox" checked={p.on} disabled={!online || (!p.on && full)} onChange={(e) => change(s, { named: namedWith(s.named, p.id, e.target.checked) })} />
                        <span className="tk-who"><strong>{p.name}</strong>{p.name !== p.email && p.email ? <small>{p.email}</small> : null}</span>
                      </label>
                    ))}
                  </div>
                ) : <p className="tk-none">{t('contacts.no-contacts-yet')}</p>}
                {gone.map((id) => (
                  <div key={id} className="tk-member tk-server-gone" data-gone={id}>
                    <span className="tk-who"><strong>{t('talk.server-share.somebody-no-longer')}</strong></span>
                    <Button className="tk-server-take-out" label={t('talk.groups.take-out')} disabled={!online} onClick={() => change(s, { named: namedWith(s.named, id, false) })} />
                  </div>
                ))}
              </>
            )}
          </div>
        );
      })}
    </>
  );
}
