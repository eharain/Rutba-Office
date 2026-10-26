// Profiles in the windows (D-2026-10-26-02; the owner, 2026-10-26: "we should
// offer multiple profiles in office so user can switch between them and
// manage them as well"). Every window's title bar shows the profile it
// belongs to, its name and colour, once there is more than one; the
// launcher's always does. A click lists the profiles, to open or switch to
// each (each runs as its own Rutba Office, side by side), with New Profile
// and Manage Profiles: make, rename, colour, the one that opens first,
// remove, and this profile's password.

import React, { useCallback, useEffect, useState } from 'react';
import { Dialog, Button, Icon, Input, Field, useMenu, useToast, formatBytes, t } from '@rutba/office-ui';

/** A profile's name as shown: the first one's own name in the window's language while it keeps it. */
export const profileName = (p) => (p?.isDefault && p.name === 'Default' ? t('Default') : p?.name || '');

/** The letter in a profile's circle. */
const initialOf = (p) => (Array.from(profileName(p))[0] || '?').toUpperCase();

/** This window's profile and the list, kept current as profiles change here or in another running profile. */
export function useProfiles(shell) {
  const [me, setMe] = useState(null);
  const [list, setList] = useState(null);
  const load = useCallback(() => {
    if (!shell.profiles?.current) return;
    shell.profiles.current().then(setMe).catch(() => {});
    shell.profiles.list().then(setList).catch(() => {});
  }, [shell]);
  useEffect(() => { load(); }, [load]);
  useEffect(() => (shell.on ? shell.on('profiles:changed', load) : undefined), [shell, load]);
  return { me, list, reload: load };
}

/** The circle with a profile's first letter, in its colour. */
export function ProfileMark({ profile, size = 18 }) {
  return (
    <span className="pf-mark" aria-hidden="true" style={{ '--pf-colour': profile?.colour || 'var(--accent)', width: size, height: size, fontSize: Math.round(size * 0.55) }}>
      {initialOf(profile)}
    </span>
  );
}

/** Manage Profiles, from the menu: opened in whichever window asked (AppFrame listens). */
export const openManageProfiles = (mode = null) => window.dispatchEvent(new CustomEvent('rutba-profiles', { detail: { mode } }));

/** The profile menu's items: each profile to open or switch to, then New Profile and Manage Profiles. */
export function profileMenuItems(shell, list, toast) {
  const items = [{ heading: true, label: t('Profiles') }];
  for (const p of list?.profiles || []) {
    items.push({
      label: p.current ? t('{name} (this window)', { name: profileName(p) }) : p.running ? t('{name} (open)', { name: profileName(p) }) : profileName(p),
      icon: p.current ? 'check' : p.protected ? 'lock' : 'contacts',
      title: p.current ? undefined : p.running ? t('Switch to {name}', { name: profileName(p) }) : t('Open {name} beside this profile', { name: profileName(p) }),
      run: () => {
        if (p.current) return;
        shell.profiles.open({ id: p.id }).catch((err) => toast?.(err.message, { tone: 'bad' }));
      },
    });
  }
  items.push('-');
  items.push({ label: t('New Profile…'), icon: 'plus', run: () => openManageProfiles('new') });
  items.push({ label: t('Manage Profiles…'), icon: 'settings', run: () => openManageProfiles() });
  return items;
}

/**
 * The title bar's profile badge: its circle and name. In every window once
 * there are two profiles or more; in the launcher always, as its way into
 * the profiles.
 */
export function ProfileBadge({ shell, always = false }) {
  const { me, list } = useProfiles(shell);
  const menu = useMenu();
  const toast = useToast();
  if (!me || (!always && me.count < 2)) return null;
  const tip = me.count > 1 ? t('Profile: {name}. Click to open another profile or manage them.', { name: profileName(me) }) : t('Profiles: open Rutba Office as another person, or for other work, side by side.');
  const open = (e) => {
    const r = e.currentTarget.getBoundingClientRect();
    menu.open({ preventDefault() {}, stopPropagation() {}, currentTarget: e.currentTarget, type: 'click', detail: e.detail, clientX: r.left, clientY: r.bottom + 4 }, profileMenuItems(shell, list, toast));
  };
  return (
    <>
      <button type="button" className="rw-btn ghost pf-badge" data-tip={tip} aria-label={tip} aria-haspopup="menu" data-profile={me.id} style={{ '--pf-colour': me.colour }} onClick={open}>
        <ProfileMark profile={me} size={18} />
        <span className="pf-badge-name">{profileName(me)}</span>
        {me.protected ? <Icon name="lock" size={11} /> : null}
      </button>
      {menu.node}
    </>
  );
}

/** Set, change or take away this profile's password. */
export function PasswordDialog({ shell, mode, profile, onClose, onDone }) {
  const [old, setOld] = useState('');
  const [next, setNext] = useState('');
  const [again, setAgain] = useState('');
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);
  // How far the mail store has got, sealed again with the new key: a large one takes a while.
  const [sealing, setSealing] = useState(null);
  useEffect(() => (busy ? shell.on?.('mail:sealing', (p) => setSealing(p)) : undefined), [busy, shell]);
  const title = mode === 'set' ? t('Set a Password') : mode === 'change' ? t('Change the Password') : t('Remove the Password');
  const go = async () => {
    setError('');
    if (mode !== 'remove' && next !== again) {
      setError(t('The two passwords are not the same.'));
      return;
    }
    setBusy(true);
    try {
      const done = mode === 'set'
        ? await shell.profiles.setPassword({ password: next })
        : mode === 'change'
          ? await shell.profiles.changePassword({ old, password: next })
          : await shell.profiles.removePassword({ old });
      onDone?.(done);
      onClose();
    } catch (err) {
      setError(err.message);
      setBusy(false);
    }
  };
  const onKey = (e) => { if (e.key === 'Enter') { e.preventDefault(); go(); } };
  return (
    <Dialog title={title} width={440} className="pf-password-dialog" onClose={onClose}
      actions={<><Button label={t('Cancel')} onClick={onClose} /><Button primary busy={busy} className="pf-password-go" label={mode === 'remove' ? t('Remove Password') : t('Save Password')} onClick={go} /></>}>
      <p className="pf-lead">
        {mode === 'remove'
          ? t('{name} will open without asking, and what the password sealed is kept as before, protected by Windows for your account only.', { name: profileName(profile) })
          : t('{name} will ask for this password before it opens. Its saved passwords, recent files, mail, calendars, contacts and Talk messages are sealed with it, so another person using this Windows account cannot read them.', { name: profileName(profile) })}
      </p>
      {mode !== 'set' ? (
        <Field label={t('Current password')}>
          <Input type="password" className="rw-input pf-old" autoFocus value={old} onChange={(e) => setOld(e.target.value)} onKeyDown={onKey} />
        </Field>
      ) : null}
      {mode !== 'remove' ? (
        <>
          <Field label={t('New password')}>
            <Input type="password" className="rw-input pf-new" autoFocus={mode === 'set'} value={next} onChange={(e) => setNext(e.target.value)} onKeyDown={onKey} />
          </Field>
          <Field label={t('Type it again')}>
            <Input type="password" className="rw-input pf-again" value={again} onChange={(e) => setAgain(e.target.value)} onKeyDown={onKey} />
          </Field>
          <p className="pf-warn" role="note">
            <Icon name="info" size={13} />
            <span>{t('A forgotten password cannot be recovered. Nobody, not even Rutba Office, can open what it sealed without it.')}</span>
          </p>
        </>
      ) : null}
      {busy && sealing?.total ? (
        <p className="pf-sealing" role="status">{t('Your mail: {done} of {total} files done', { done: Math.min(sealing.done, sealing.total), total: sealing.total })}</p>
      ) : null}
      {error ? <p className="pf-error" role="alert">{error}</p> : null}
    </Dialog>
  );
}

/** Manage Profiles. */
export function ManageProfiles({ shell, mode = null, onClose }) {
  const toast = useToast();
  const { me, list, reload } = useProfiles(shell);
  const [renaming, setRenaming] = useState(null);
  const [newName, setNewName] = useState('');
  const [newColour, setNewColour] = useState(null);
  const [colourFor, setColourFor] = useState(null);
  const [removing, setRemoving] = useState(null);
  const [password, setPassword] = useState(null);

  const act = async (fn, said = null) => {
    try {
      await fn();
      reload();
      if (said) toast(said, { tone: 'good', ms: 3000 });
      return true;
    } catch (err) {
      toast(err.message, { tone: 'bad', ms: 6000 });
      return false;
    }
  };
  const make = async () => {
    const name = newName.trim();
    if (await act(() => shell.profiles.make({ name, colour: newColour || undefined }), t('Profile {name} made. Open it from this list or the profile menu.', { name }))) {
      setNewName('');
      setNewColour(null);
    }
  };
  const askRemove = async (p) => {
    try {
      const r = await shell.profiles.removal({ id: p.id });
      if (r.refused) toast(r.refused, { tone: 'bad', ms: 6000 });
      else setRemoving({ ...p, size: r.size });
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    }
  };
  const colours = list?.colours || [];

  return (
    <Dialog title={t('Profiles')} width={640} className="pf-dialog" onClose={onClose} actions={<Button label={t('Close')} primary onClick={onClose} />}>
      <p className="pf-lead">
        {t('Each profile keeps its own settings, saved passwords, recent files, mail, calendars, contacts, Talk sign-in and AI Tools settings. Profiles open side by side, each in windows of its own, with its name and colour in the title bar.')}
      </p>
      <ul className="pf-list">
        {(list?.profiles || []).map((p) => (
          <li key={p.id} className={`pf-row${p.current ? ' current' : ''}`} data-profile={p.id}>
            <ProfileMark profile={p} size={28} />
            <span className="pf-row-text">
              {renaming === p.id ? (
                <Input
                  className="rw-input pf-rename"
                  aria-label={t('Profile name')}
                  autoFocus
                  defaultValue={p.name}
                  onKeyDown={(e) => {
                    e.stopPropagation();
                    if (e.key === 'Enter') { const name = e.target.value; setRenaming(null); act(() => shell.profiles.rename({ id: p.id, name })); }
                    else if (e.key === 'Escape') setRenaming(null);
                  }}
                  onBlur={() => setRenaming(null)}
                />
              ) : (
                <strong className="pf-row-name">{profileName(p)}</strong>
              )}
              <span className="pf-tags">
                {p.current ? <span className="pf-tag">{t('This window')}</span> : p.running ? <span className="pf-tag">{t('Open')}</span> : null}
                {p.openFirst ? <span className="pf-tag">{t('Opens first')}</span> : null}
                {p.protected ? <span className="pf-tag"><Icon name="lock" size={10} /> {t('Password')}</span> : null}
              </span>
            </span>
            <span className="pf-row-actions">
              {!p.current ? <Button label={p.running ? t('Switch to') : t('Open')} className="pf-open" onClick={() => act(() => shell.profiles.open({ id: p.id }))} /> : null}
              <Button ghost icon="textbox" className="pf-rename-btn" title={t('Rename')} onClick={() => setRenaming(p.id)} />
              <Button ghost icon="bucket" className="pf-colour-btn" title={t('Colour')} pressed={colourFor === p.id} onClick={() => setColourFor(colourFor === p.id ? null : p.id)} />
              {!p.openFirst ? <Button ghost icon="star" className="pf-first-btn" title={t('Open this profile first when Rutba Office starts')} onClick={() => act(() => shell.profiles.setOpenFirst({ id: p.id }))} /> : null}
              {!p.isDefault && !p.current ? <Button ghost icon="trash" className="pf-remove-btn" title={t('Remove this profile')} onClick={() => askRemove(p)} /> : null}
            </span>
            {colourFor === p.id ? (
              <span className="pf-swatches" role="radiogroup" aria-label={t('Colour')}>
                {colours.map((c) => (
                  <button key={c} type="button" role="radio" aria-checked={c === p.colour} className={`pf-swatch${c === p.colour ? ' on' : ''}`} style={{ background: c }} data-colour={c} aria-label={c}
                    onClick={() => { setColourFor(null); act(() => shell.profiles.setColour({ id: p.id, colour: c })); }} />
                ))}
              </span>
            ) : null}
          </li>
        ))}
      </ul>

      <h3 className="pf-h">{t('New profile')}</h3>
      <div className="pf-new-row">
        <Input className="rw-input pf-new-name" aria-label={t('Name of the new profile')} placeholder={t('Name, for example Work')} autoFocus={mode === 'new'} value={newName} onChange={(e) => setNewName(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') make(); }} />
        <span className="pf-swatches inline" role="radiogroup" aria-label={t('Colour')}>
          {colours.map((c) => (
            <button key={c} type="button" role="radio" aria-checked={c === newColour} className={`pf-swatch${c === newColour ? ' on' : ''}`} style={{ background: c }} data-colour={c} aria-label={c} onClick={() => setNewColour(c)} />
          ))}
        </span>
        <Button primary icon="plus" label={t('Make Profile')} className="pf-make" disabled={!newName.trim()} onClick={make} />
      </div>

      {me ? (
        <>
          <h3 className="pf-h">{t('Password for {name}', { name: profileName(me) })}</h3>
          <p className="pf-note">
            {me.protected
              ? t('This profile asks for its password before it opens.')
              : t('A profile can ask for a password before it opens. Its saved passwords, recent files, mail, calendars, contacts and Talk messages are then sealed with it.')}
          </p>
          <div className="pf-password-row">
            {me.protected ? (
              <>
                <Button icon="lock" label={t('Change Password…')} className="pf-change-password" onClick={() => setPassword('change')} />
                <Button icon="unlock" label={t('Remove Password…')} className="pf-remove-password" onClick={() => setPassword('remove')} />
              </>
            ) : (
              <Button icon="lock" label={t('Set a Password…')} className="pf-set-password" onClick={() => setPassword('set')} />
            )}
          </div>
          <p className="pf-note">{t('Another profile’s password is set from that profile’s own windows.')}</p>
        </>
      ) : null}

      {removing ? (
        <Dialog title={t('Remove {name}?', { name: profileName(removing) })} width={440} className="pf-remove-dialog" onClose={() => setRemoving(null)}
          actions={<><Button label={t('Cancel')} onClick={() => setRemoving(null)} /><Button primary className="pf-remove-go" label={t('Remove Profile')} onClick={() => { const p = removing; setRemoving(null); act(() => shell.profiles.remove({ id: p.id }), t('Profile {name} removed.', { name: profileName(p) })); }} /></>}>
          <p className="pf-lead">
            {t('Its settings, saved passwords, recent files, mail, calendars, contacts, Talk messages and AI Tools settings ({size}) go to the Recycle Bin. Documents you saved elsewhere, such as in Documents, stay where they are.', { size: formatBytes(removing.size || 0) })}
          </p>
        </Dialog>
      ) : null}
      {password ? <PasswordDialog shell={shell} mode={password} profile={me} onClose={() => setPassword(null)} onDone={() => { reload(); toast(password === 'remove' ? t('The password is removed.') : t('The password is saved.'), { tone: 'good', ms: 3000 }); }} /> : null}
    </Dialog>
  );
}

export const PROFILES_CSS = `
.pf-mark { display: inline-grid; place-items: center; border-radius: 50%; background: var(--pf-colour); color: #fff; font-weight: 600; line-height: 1; flex: none; }
.pf-badge { display: inline-flex; align-items: center; gap: 6px; max-width: 180px; padding: 0 8px 0 4px; }
.pf-badge-name { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-size: 12px; }
.pf-lead { margin-top: 0; color: var(--ink-2); line-height: 1.45; }
.pf-list { list-style: none; margin: 0 0 8px; padding: 0; display: grid; gap: 6px; }
.pf-row { display: grid; grid-template-columns: auto 1fr auto; align-items: center; gap: 10px; padding: 8px 10px; border: 1px solid var(--line); border-radius: var(--r-2); }
.pf-row.current { border-color: var(--accent-line); background: var(--accent-soft); }
.pf-row-text { display: grid; gap: 3px; min-width: 0; }
.pf-row-name { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.pf-tags { display: flex; gap: 4px; flex-wrap: wrap; }
.pf-tag { display: inline-flex; align-items: center; gap: 3px; padding: 1px 6px; border-radius: 9px; background: var(--sunken); color: var(--ink-2); font-size: 11px; }
.pf-row-actions { display: flex; align-items: center; gap: 2px; }
.pf-swatches { grid-column: 1 / -1; display: flex; gap: 6px; flex-wrap: wrap; padding: 2px 0 0 38px; }
.pf-swatches.inline { grid-column: auto; padding: 0; }
.pf-swatch { width: 20px; height: 20px; border-radius: 50%; border: 2px solid transparent; padding: 0; cursor: pointer; box-shadow: 0 0 0 1px var(--line); }
.pf-swatch.on { border-color: var(--surface); box-shadow: 0 0 0 2px var(--ink); }
.pf-h { font-size: 12px; text-transform: uppercase; letter-spacing: .04em; color: var(--ink-3); margin: 16px 0 6px; }
.pf-new-row { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
.pf-new-row .pf-new-name { flex: 1; min-width: 180px; }
.pf-note { color: var(--ink-3); font-size: 12.5px; margin: 0 0 6px; }
.pf-password-row { display: flex; gap: 8px; margin: 0 0 6px; }
.pf-warn { display: flex; align-items: flex-start; gap: 6px; padding: 8px 10px; border-radius: var(--r-2); background: color-mix(in srgb, var(--warn) 12%, var(--sunken)); color: var(--ink-2); font-size: 12.5px; line-height: 1.45; }
.pf-error { color: var(--bad); margin: 6px 0 0; }
.pf-sealing { color: var(--ink-2); margin: 6px 0 0; font-size: 12.5px; font-variant-numeric: tabular-nums; }
.pf-password-dialog .rw-field { margin-bottom: 8px; }
`;
