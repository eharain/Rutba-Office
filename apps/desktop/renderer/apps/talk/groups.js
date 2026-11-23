// Group conversations (Office 2.27): three or more people writing to each
// other. A group is made of one's own contacts; anyone in it can add one of
// theirs, so a group may hold someone the person does not know, and says so.
// One dialog makes a group and one shows who is in it: the name, the people,
// adding, and leaving. Nothing else to learn.

import React, { useState } from 'react';
import { Button, Dialog, Input } from '@rutba/office-ui';
import { t } from '../talk-words.js';
import { canJoin, groupTitle, leaveGroupText, memberNote, refusalText } from './lines.js';

const nameOf = (p) => p.name || p.email || '';

/** Contacts to tick, by name. */
function Pick({ contacts, chosen, onChange }) {
  const toggle = (id) => onChange(chosen.includes(id) ? chosen.filter((x) => x !== id) : [...chosen, id]);
  if (!contacts.length) return <p className="tk-none">{t('talk.groups.none-contacts-left')}</p>;
  return (
    <div className="tk-pick-list" role="group" aria-label={t('common.contacts')}>
      {contacts.map((c) => (
        <label key={c.id} className="tk-pick" data-pick={c.id}>
          <input type="checkbox" checked={chosen.includes(c.id)} onChange={() => toggle(c.id)} />
          <span className="tk-who"><strong>{nameOf(c)}</strong><small>{c.email}</small></span>
        </label>
      ))}
    </div>
  );
}

/** A new group: a name if the person wants one, and two contacts or more. */
export function NewGroup({ shell, state, toast, onClose, onMade }) {
  const [name, setName] = useState('');
  const [chosen, setChosen] = useState([]);
  const [busy, setBusy] = useState(false);
  const make = async () => {
    setBusy(true);
    try {
      const r = await shell.talk.groupCreate({ name: name.trim(), members: chosen });
      if (r.error) { toast(refusalText(r.error, r), { tone: 'bad', ms: 5000 }); return; }
      onMade?.(r.group);
      onClose();
    } catch (err) {
      toast(err?.message || String(err), { tone: 'bad', ms: 5000 });
    } finally {
      setBusy(false);
    }
  };
  return (
    <Dialog
      title={t('talk.groups.new-group')}
      width={460}
      className="tk-new-group"
      onClose={onClose}
      actions={<><Button label={t('common.cancel')} onClick={onClose} /><Button primary className="tk-group-make" label={t('talk.groups.start-group')} disabled={busy || chosen.length < 2} onClick={make} /></>}
    >
      <p className="tk-lead">{t('talk.groups.choose-two-more')}</p>
      <label className="tk-field"><span>{t('talk.groups.name-want-one')}</span><Input className="tk-group-name" value={name} maxLength={80} onChange={(e) => setName(e.target.value)} /></label>
      <Pick contacts={state.contacts} chosen={chosen} onChange={setChosen} />
    </Dialog>
  );
}

/** Who is in a group: its name, its people, adding one's own contacts, and leaving. */
export function GroupPeople({ shell, state, group, toast, onClose, onReport, asked }) {
  const me = state.me?.id;
  const mine = group.starter === me;
  const [name, setName] = useState(group.name || '');
  const [adding, setAdding] = useState(false);
  const [chosen, setChosen] = useState([]);
  const [busy, setBusy] = useState(false);
  const run = async (what, done) => {
    setBusy(true);
    try {
      const r = await what();
      if (r.error) toast(refusalText(r.error, r), { tone: 'bad', ms: 5000 });
      else done?.(r);
    } catch (err) {
      toast(err?.message || String(err), { tone: 'bad', ms: 5000 });
    } finally {
      setBusy(false);
    }
  };
  const rename = () => run(() => shell.talk.groupRename({ groupId: group.id, name: name.trim() }));
  const add = () => run(() => shell.talk.groupAdd({ groupId: group.id, members: chosen }), () => { setChosen([]); setAdding(false); });
  const remove = async (m) => {
    if (!(await asked(t('talk.groups.take-name-out', { name: nameOf(m) }), t('talk.groups.no-longer-see'), t('talk.groups.take-out')))) return;
    run(() => shell.talk.groupRemove({ groupId: group.id, member: m.id }));
  };
  const leave = async () => {
    if (!(await asked(t('talk.groups.leave-this-group'), leaveGroupText(group, me), t('talk.groups.leave-group')))) return;
    run(() => shell.talk.groupLeave({ groupId: group.id }), onClose);
  };
  const free = canJoin(state.contacts, group);
  return (
    <Dialog
      title={groupTitle(group, me)}
      width={480}
      className="tk-group-people"
      onClose={onClose}
      actions={<><Button className="tk-group-leave" label={t('talk.groups.leave-group')} disabled={busy} onClick={leave} /><Button primary label={t('common.done')} onClick={onClose} /></>}
    >
      <div className="tk-group-rename">
        <label className="tk-field"><span>{t('talk.groups.name-group')}</span><Input className="tk-group-name" value={name} maxLength={80} onChange={(e) => setName(e.target.value)} /></label>
        <Button className="tk-group-rename-save" label={t('common.save')} disabled={busy || name.trim() === (group.name || '')} onClick={rename} />
      </div>
      <div className="tk-member-list">
        {group.members.map((m) => (
          <div key={m.id} className="tk-member" data-member={m.id}>
            <span className="tk-who"><strong>{nameOf(m)}</strong><small>{memberNote(m, group, me)}</small></span>
            {m.id !== me ? <Button ghost className="tk-member-report" label={t('common.report')} onClick={() => onReport(m)} /> : null}
            {mine && m.id !== me ? <Button ghost className="tk-member-remove" label={t('talk.groups.take-out')} disabled={busy} onClick={() => remove(m)} /> : null}
          </div>
        ))}
      </div>
      {adding ? (
        <div className="tk-group-add">
          <Pick contacts={free} chosen={chosen} onChange={setChosen} />
          <Button primary className="tk-group-add-save" label={t('talk.groups.add-to-group')} disabled={busy || !chosen.length} onClick={add} />
        </div>
      ) : <Button icon="plus" className="tk-group-add-open" label={t('talk.groups.add-people')} onClick={() => setAdding(true)} />}
    </Dialog>
  );
}
