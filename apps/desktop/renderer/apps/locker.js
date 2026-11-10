// The Locker's window: folders, their locked files, and the choices that
// belong to the person. Its engine lives in the main process and every file
// operation goes through it; this module only asks and shows.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Button, Icon, Spacer, Empty, Field, Input, Select, Dialog, useToast, formatBytes, formatWhen } from '@rutba/office-ui';
import { AppFrame, useAppMenu } from '../shell.js';
import { t } from './locker-words.js';

const CSS = `
.lk { position: relative; flex: 1; min-width: 0; min-height: 0; display: flex; flex-direction: column; background: var(--window); color: var(--ink); font-size: 13px; }
.lk-bar { display: flex; align-items: center; gap: 10px; padding: 10px 16px; background: var(--chrome); border-bottom: 1px solid var(--line-soft); }
.lk-pill { display: inline-flex; align-items: center; gap: 6px; height: 28px; padding: 0 12px; border-radius: 14px; background: color-mix(in srgb, var(--good) 14%, transparent); color: var(--good); font-weight: 600; font-size: 12.5px; }
.lk-pill.closed { background: var(--hover); color: var(--ink-2, var(--ink)); }
.lk-body { flex: 1; display: grid; grid-template-columns: 300px minmax(0, 1fr); min-height: 0; }
.lk-side { display: flex; flex-direction: column; gap: 6px; padding: 14px; background: var(--chrome); border-right: 1px solid var(--line-soft); overflow: auto; }
.lk-side-head { font-size: 11px; font-weight: 700; letter-spacing: .06em; text-transform: uppercase; color: var(--ink-3, var(--ink)); }
.lk-folder { display: flex; align-items: center; gap: 10px; padding: 8px 10px; border-radius: 8px; border: 1px solid transparent; background: none; color: inherit; font: inherit; text-align: left; cursor: pointer; width: 100%; }
.lk-folder:hover { background: var(--hover); }
.lk-folder.on { background: var(--selected); border-color: var(--accent-line); }
.lk-folder small { display: block; font-size: 11.5px; color: var(--ink-3, var(--ink)); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.lk-folder b { font-weight: 600; }
.lk-folder-text { flex: 1; min-width: 0; }
.lk-note { font-size: 12px; color: var(--ink-3, var(--ink)); line-height: 1.45; }
.lk-note.foot { margin-top: auto; }
.lk-main { display: flex; flex-direction: column; gap: 12px; padding: 16px 20px; min-width: 0; min-height: 0; }
.lk-head { display: flex; align-items: center; gap: 10px; }
.lk-head h1 { margin: 0; font-size: 19px; }
.lk-head-note { font-size: 12.5px; color: var(--ink-3, var(--ink)); min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.lk-alert { display: flex; align-items: center; gap: 10px; padding: 10px 12px; border-radius: 8px; background: color-mix(in srgb, var(--bad) 12%, transparent); color: var(--bad); }
.lk-alert b { font-weight: 600; }
.lk-table-wrap { flex: 1; min-height: 0; overflow: auto; background: var(--surface); border: 1px solid var(--line-soft); border-radius: 10px; }
.lk-table { border-collapse: collapse; width: 100%; font-size: 13px; }
.lk-table th { position: sticky; top: 0; background: var(--surface); text-align: left; font-size: 11.5px; font-weight: 600; color: var(--ink-3, var(--ink)); padding: 6px 10px; border-bottom: 1px solid var(--line-soft); }
.lk-table td { padding: 8px 10px; border-bottom: 1px solid var(--line-soft); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 320px; }
.lk-table tr.sel td { background: var(--selected); }
.lk-table tbody tr { cursor: default; }
.lk-table tbody tr:hover td { background: var(--hover); }
.lk-table tr.sel:hover td { background: var(--selected); }
.lk-table tr.lk-dir td, .lk-table tr.lk-dir:hover td { padding: 6px 10px; max-width: none; background: color-mix(in srgb, var(--accent) 12%, var(--surface)); border-top: 1px solid var(--accent-line); border-bottom: 1px solid var(--accent-line); }
.lk-dir-in { display: flex; align-items: center; gap: 8px; color: var(--accent); }
.lk-dir-in b { font-weight: 700; overflow: hidden; text-overflow: ellipsis; }
.lk-table td.lk-indent { padding-inline-start: 34px; }
.lk-table td.lk-row-acts { max-width: none; width: 1%; overflow: visible; padding: 4px 10px; }
.lk-row-acts > div { display: flex; justify-content: flex-end; gap: 6px; }
.lk-table tr.lk-more td, .lk-table tr.lk-more:hover td { max-width: 0; white-space: normal; background: var(--selected); padding: 6px 10px 8px; }
.lk-more-in { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
.lk-more-in .lk-detail { min-width: 0; overflow-wrap: anywhere; }
.lk-search { display: flex; align-items: center; gap: 10px; }
.lk-search > :first-child { flex: 1; min-width: 0; }
.lk-state { display: inline-flex; align-items: center; height: 21px; padding: 0 8px; border-radius: 11px; font-size: 11.5px; font-weight: 600; }
.lk-state.locked { background: color-mix(in srgb, var(--good) 14%, transparent); color: var(--good); }
.lk-state.waiting { background: color-mix(in srgb, var(--warn) 16%, transparent); color: var(--warn); }
.lk-state.excluded { background: var(--hover); color: var(--ink-3, var(--ink)); }
.lk-state.conflict, .lk-state.damaged { background: color-mix(in srgb, var(--bad) 14%, transparent); color: var(--bad); }
.lk-why { color: var(--ink-3, var(--ink)); }
.lk-actions { display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }
.lk-choice-group { display: flex; flex-direction: column; gap: 6px; }
.lk-choice-head { font-size: 13px; font-weight: 600; }
.lk-choices { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 6px; }
.lk-choice { display: flex; gap: 8px; align-items: flex-start; padding: 8px 10px; border: 1px solid var(--line); border-radius: 8px; background: var(--surface); line-height: 1.4; cursor: pointer; }
.lk-choice.on { border-color: var(--accent); background: var(--selected); }
.lk-choice b { display: block; font-weight: 600; }
.lk-choice small { color: var(--ink-3, var(--ink)); font-size: 12px; }
.lk-sug { display: inline-block; margin-left: 5px; padding: 0 5px; border-radius: 3px; background: color-mix(in srgb, var(--good) 15%, transparent); color: var(--good); font-size: 10px; font-weight: 700; vertical-align: 1px; }
.lk-dialog-body { display: flex; flex-direction: column; gap: 14px; }
.lk-warn { display: flex; gap: 8px; padding: 8px 10px; border-radius: 8px; background: color-mix(in srgb, var(--warn) 14%, transparent); color: var(--warn); font-size: 12.5px; line-height: 1.45; }
.lk-rule-area { width: 100%; height: 220px; resize: vertical; font-family: var(--mono); font-size: 12.5px; }
.lk-center { position: absolute; inset: 0; display: grid; place-items: center; background: var(--window); }
.lk-card { width: 440px; display: flex; flex-direction: column; gap: 14px; padding: 22px 24px; background: var(--surface); border: 1px solid var(--line-soft); border-radius: 14px; }
.lk-card h1 { margin: 0; font-size: 17px; }
.lk-card-head { display: flex; align-items: center; gap: 12px; }
.lk-card-mark { width: 44px; height: 44px; border-radius: 12px; background: #2f4a6b; color: #fff; display: grid; place-items: center; flex: none; }
.lk-card-sub { font-size: 12.5px; color: var(--ink-3, var(--ink)); }
.lk-or { display: flex; align-items: center; gap: 10px; font-size: 12px; color: var(--ink-3, var(--ink)); }
.lk-or span { flex: 1; height: 1px; background: var(--line-soft); }
.lk-card-foot { display: flex; justify-content: space-between; font-size: 12.5px; }
.lk-link { border: none; background: none; color: var(--accent); cursor: pointer; font: inherit; padding: 0; text-decoration: underline; }
.lk-error { color: var(--bad); font-size: 12.5px; }
.lk-key { font-family: var(--mono); font-size: 13px; padding: 10px; background: var(--sunken); border-radius: 8px; overflow-wrap: anywhere; }
.lk-fullbtn { display: inline-flex; align-items: center; justify-content: center; gap: 8px; height: 40px; border-radius: 8px; border: 1px solid var(--line); background: var(--surface); color: var(--ink); font: inherit; cursor: pointer; font-size: 13.5px; }
.lk-fullbtn.primary { background: var(--accent); border-color: var(--accent); color: var(--ink-on-accent); font-weight: 600; }
.lk-fullbtn:disabled { opacity: .55; cursor: default; }
`;

const NAMES = {
  hidden: () => t('Names hidden in the cloud'),
  visible: () => t('Names visible'),
  delete: () => t('plain originals deleted once locked'),
  keep: () => t('plain originals kept'),
};

// Three words for a person. What lies behind them is in the line under the list.
const stateName = (e) => ({
  locked: t('Locked'),
  waiting: t('Not locked'),
  excluded: t('Not locked'),
  conflict: t('Needs you'),
  damaged: t('Needs you'),
}[e.state]);

const stateDetail = (e) => ({
  locked: e.pieces ? t('Locked, in {n} pieces', { n: e.pieces }) : t('Locked'),
  waiting: t('Not locked'),
  excluded: t('Stays plain'),
  conflict: t('{n} versions to choose between', { n: e.versions || 2 }),
  damaged: e.newer ? t('Needs a newer Rutba Office') : t('Damaged'),
}[e.state]);

// What the engine said, without the wrapping a message picks up on its way across.
const said = (err) => String(err?.message || err).replace(/^(?:Error invoking remote method '[^']*':\s*)?(?:\w*Error:\s*)*/, '');

/** The files by the directory they are in: the folder's own first, then each subfolder in order. */
function byDirectory(entries) {
  const groups = new Map();
  for (const e of entries) {
    const cut = e.rel.lastIndexOf('/');
    const dir = cut < 0 ? '' : e.rel.slice(0, cut);
    if (!groups.has(dir)) groups.set(dir, []);
    groups.get(dir).push(e);
  }
  return [...groups.entries()].sort(([a], [b]) => a.localeCompare(b));
}

/** Every word typed must appear somewhere in the file's path, in any case. */
function matching(entries, query) {
  const words = query.trim().toLowerCase().split(/\s+/).filter(Boolean);
  if (!words.length) return entries;
  return entries.filter((e) => {
    const rel = e.rel.toLowerCase();
    return words.every((w) => rel.includes(w));
  });
}

const whyText = (e) => e.message || e.why || (e.state === 'locked' ? t('Locked by the folder\'s rules') : '');

function useLocker(shell) {
  const [state, setState] = useState(null);
  const refresh = useCallback(() => shell.locker.state().then(setState).catch(() => {}), [shell]);
  useEffect(() => {
    refresh();
    const off = shell.on('locker:state', setState);
    return off;
  }, [shell, refresh]);
  return [state, setState, refresh];
}

/** First start: a password and the recovery key, said once. */
function Setup({ shell, onDone, onOpenSent }) {
  const [password, setPassword] = useState('');
  const [again, setAgain] = useState('');
  const [recovery, setRecovery] = useState(null);
  const [error, setError] = useState(null);
  const [busy, setBusy] = useState(false);
  const can = password.length >= 8 && password === again && !busy;
  const make = async () => {
    setBusy(true);
    setError(null);
    try {
      const out = await shell.locker.setup({ password });
      setRecovery(out.recovery);
    } catch (err) {
      setError(said(err));
    } finally {
      setBusy(false);
    }
  };
  return (
    <div className="lk-center">
      <div className="lk-card" role="dialog" aria-label={t('Make a locker')}>
        <div className="lk-card-head">
          <div className="lk-card-mark"><Icon name="lock" size={22} /></div>
          <div>
            <h1>{t('Make your locker')}</h1>
            <div className="lk-card-sub">{t('A password that opens your locked folders on any computer.')}</div>
          </div>
        </div>
        {recovery ? (
          <>
            <div>{t('Your recovery key. Write it down and keep it somewhere safe.')}</div>
            <div className="lk-key" data-recovery>{recovery}</div>
            <div className="lk-card-sub">{t('A forgotten password with no recovery key means the files cannot be opened. It is shown only now.')}</div>
            <button type="button" className="lk-fullbtn primary" onClick={onDone}>{t('I have kept it')}</button>
          </>
        ) : (
          <>
            <Field label={t('Locker password')}><Input type="password" value={password} onChange={(e) => setPassword(e.target.value)} aria-label={t('Locker password')} /></Field>
            <Field label={t('The same password again')}><Input type="password" value={again} onChange={(e) => setAgain(e.target.value)} aria-label={t('The same password again')} /></Field>
            <div className="lk-card-sub">{t('At least eight characters. Rutba Office cannot get it back for you.')}</div>
            {error ? <div className="lk-error" role="alert">{error}</div> : null}
            <button type="button" className="lk-fullbtn primary" disabled={!can} onClick={make}>{t('Make the locker')}</button>
            <button type="button" className="lk-link lk-open-sent" onClick={onOpenSent}>{t('Open a file someone sent you…')}</button>
          </>
        )}
      </div>
    </div>
  );
}

/** The locker is shut: Windows Hello if it was turned on here, the password, or the recovery key. */
function Unlock({ shell, state, onOpened, onOpenSent }) {
  const [password, setPassword] = useState('');
  const [recovery, setRecovery] = useState('');
  const [useRecovery, setUseRecovery] = useState(false);
  const [error, setError] = useState(null);
  const [busy, setBusy] = useState(false);
  const go = async (how) => {
    setBusy(true);
    setError(null);
    try {
      const s = await shell.locker.unlock(how);
      if (!s.cancelled) onOpened(s);
    } catch (err) {
      setError(said(err));
    } finally {
      setBusy(false);
      setPassword('');
    }
  };
  return (
    <div className="lk-center">
      <div className="lk-card" role="dialog" aria-label={t('Unlock the locker')}>
        <div className="lk-card-head">
          <div className="lk-card-mark"><Icon name="lock" size={22} /></div>
          <div>
            <h1>{t('Unlock the locker')}</h1>
            <div className="lk-card-sub">{t('Your locked folders stay locked until you do.')}</div>
          </div>
        </div>
        {state.helloOn ? (
          <>
            <button type="button" className="lk-fullbtn primary" disabled={busy} onClick={() => go({ hello: true })}>
              <Icon name="key" size={16} />{t('Use Windows Hello (PIN, fingerprint or face)')}
            </button>
            <div className="lk-or"><span />{t('or')}<span /></div>
          </>
        ) : null}
        {useRecovery ? (
          <>
            <Field label={t('Recovery key')}><Input value={recovery} onChange={(e) => setRecovery(e.target.value)} aria-label={t('Recovery key')} /></Field>
            <button type="button" className="lk-fullbtn" disabled={busy || !recovery} onClick={() => go({ recovery })}>{t('Unlock with the recovery key')}</button>
          </>
        ) : (
          <>
            <Field label={t('Locker password')}>
              <Input type="password" value={password} autoFocus onChange={(e) => setPassword(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && password) go({ password }); }} aria-label={t('Locker password')} />
            </Field>
            <button type="button" className="lk-fullbtn" disabled={busy || !password} onClick={() => go({ password })}>{t('Unlock with password')}</button>
          </>
        )}
        {error ? <div className="lk-error" role="alert">{error}</div> : null}
        <div className="lk-card-foot">
          <button type="button" className="lk-link" onClick={() => setUseRecovery(!useRecovery)}>{useRecovery ? t('Use the password') : t('Use the recovery key')}</button>
          <span className="lk-card-sub">{t('Stays unlocked for {n} minutes', { n: state.idleMinutes })}</span>
        </div>
        <button type="button" className="lk-link lk-open-sent" onClick={onOpenSent}>{t('Open a file someone sent you…')}</button>
      </div>
    </div>
  );
}

/** One of the two-way choices, with the suggested one marked. */
function Choice({ legend, name, value, onChange, options }) {
  return (
    <div className="lk-choice-group">
      <span className="lk-choice-head">{legend}</span>
      <div className="lk-choices" role="radiogroup" aria-label={legend}>
        {options.map((o) => (
          <label key={o.value} className={`lk-choice${value === o.value ? ' on' : ''}`}>
            <input type="radio" name={name} checked={value === o.value} onChange={() => onChange(o.value)} />
            <span><b>{o.title}{o.suggested ? <i className="lk-sug" style={{ fontStyle: 'normal' }}>{t('SUGGESTED')}</i> : null}</b><small>{o.hint}</small></span>
          </label>
        ))}
      </div>
    </div>
  );
}

const CLOUD = /(onedrive|dropbox|google ?drive|icloud)/i;
const whichCloud = (p) => {
  const m = CLOUD.exec(p || '');
  if (!m) return null;
  return /onedrive/i.test(m[1]) ? 'OneDrive' : /dropbox/i.test(m[1]) ? 'Dropbox' : /icloud/i.test(m[1]) ? 'iCloud Drive' : 'Google Drive';
};

function Choices({ c, set, keepDirPick, showWhat }) {
  return (
    <>
      {showWhat ? (
        <Choice
          legend={t('What to lock')}
          name="what"
          value={c.what}
          onChange={(what) => set({ ...c, what })}
          options={[
            { value: 'every', title: t('Every file'), hint: t('All but the temporary files Office and Windows make'), suggested: true },
            { value: 'types', title: t('Only these types'), hint: c.types },
          ]}
        />
      ) : null}
      {showWhat && c.what === 'types' ? (
        <Field label={t('File types, separated by commas')}><Input value={c.types} onChange={(e) => set({ ...c, types: e.target.value })} aria-label={t('File types')} /></Field>
      ) : null}
      <Choice
        legend={t('File names in the cloud')}
        name="names"
        value={c.hidden ? 'hidden' : 'visible'}
        onChange={(v) => set({ ...c, hidden: v === 'hidden' })}
        options={[
          { value: 'hidden', title: t('Hidden'), hint: t('Your cloud sees random names. Folder names stay visible.'), suggested: true },
          { value: 'visible', title: t('Visible'), hint: t('Lease.pdf stays Lease.pdf.rlock, easier to find in Explorer') },
        ]}
      />
      <Choice
        legend={t('Plain files once they are locked')}
        name="after"
        value={c.afterLock}
        onChange={(afterLock) => set({ ...c, afterLock })}
        options={[
          { value: 'delete', title: t('Delete the plain copy'), hint: t('Only after the locked copy has been read back and matched'), suggested: true },
          { value: 'keep', title: t('Keep it'), hint: t('Moved to a folder you choose, outside the locker') },
        ]}
      />
      {c.afterLock === 'keep' ? (
        <div className="lk-actions">
          <Button label={t('Choose a folder…')} onClick={keepDirPick} />
          <span className="lk-why">{c.keepDir || t('No folder chosen yet')}</span>
        </div>
      ) : null}
      <Choice
        legend={t('While a file is open')}
        name="openin"
        value={c.openIn || 'beside'}
        onChange={(openIn) => set({ ...c, openIn })}
        options={[
          { value: 'beside', title: t('Beside the file'), hint: t('In a hidden folder next to it, written back as you save and removed when you close it'), suggested: true },
          { value: 'profile', title: t('In Office\'s own folder'), hint: t('Away from this folder, for a folder a cloud drive uploads') },
        ]}
      />
      <Choice
        legend={t('Versions kept of each file')}
        name="versions"
        value={String(c.keepVersions)}
        onChange={(v) => set({ ...c, keepVersions: Number(v) })}
        options={[
          { value: '3', title: t('Three'), hint: t('The current one and the two before it'), suggested: true },
          { value: '1', title: t('One'), hint: t('Only the current one') },
        ]}
      />
    </>
  );
}

/** A chosen folder, looked at first: one that is already locked is opened, never locked again. */
function AddFolder({ shell, root, onClose, onAdded }) {
  const [seen, setSeen] = useState(null);
  useEffect(() => {
    let live = true;
    shell.locker.lookAtFolder({ root }).then((v) => { if (live) setSeen(v); }, (err) => { if (live) setSeen({ trouble: said(err) }); });
    return () => { live = false; };
  }, [shell, root]);
  if (!seen) return null;
  if (seen.trouble) {
    return (
      <Dialog title={t('This folder cannot be added')} width={520} onClose={onClose} actions={<Button primary label={t('Close')} onClick={onClose} />}>
        <div className="lk-dialog-body">
          <div className="lk-card-sub">{root}</div>
          <div className="lk-error" role="alert">{seen.trouble}</div>
        </div>
      </Dialog>
    );
  }
  if (seen.locked) return <OpenLockedFolder shell={shell} root={root} opens={seen.opens} onClose={onClose} onAdded={onAdded} />;
  return <LockNewFolder shell={shell} root={root} onClose={onClose} onAdded={onAdded} />;
}

/** A folder locked before, here or on another computer: its password or recovery key opens it, and its files stay as they are. */
function OpenLockedFolder({ shell, root, opens, onClose, onAdded }) {
  const [useRecovery, setUseRecovery] = useState(false);
  const [password, setPassword] = useState('');
  const [recovery, setRecovery] = useState('');
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState(null);
  const can = !busy && (opens || Boolean(useRecovery ? recovery : password));
  const go = async () => {
    if (!can) return;
    setBusy(true);
    setError(null);
    try {
      const out = await shell.locker.addFolder({ root, choices: {}, password: !opens && !useRecovery ? password : null, recovery: !opens && useRecovery ? recovery : null });
      onAdded(out.id);
    } catch (err) {
      setError(said(err));
      setBusy(false);
    }
  };
  const onKeyDown = (e) => { if (e.key === 'Enter') go(); };
  return (
    <Dialog
      title={t('Open the locked folder {name}', { name: root.split(/[\\/]/).pop() })}
      width={560}
      className="lk-open-locked"
      onClose={onClose}
      actions={<>
        <Button label={t('Cancel')} onClick={onClose} />
        <Button primary className="lk-open-go" label={busy ? t('Opening…') : t('Open this folder')} disabled={!can} onClick={go} />
      </>}
    >
      <div className="lk-dialog-body">
        <div className="lk-card-sub">{root}</div>
        {opens ? <div>{t('This folder is already locked, and this locker can open it. Its files stay locked.')}</div> : (
          <>
            <div>{t('This folder is already locked. Enter the password that locked it, or its recovery key. Nothing in the folder changes.')}</div>
            {useRecovery ? (
              <Field label={t('Recovery key')}>
                <Input value={recovery} onChange={(e) => setRecovery(e.target.value)} onKeyDown={onKeyDown} autoFocus aria-label={t('Recovery key that opens this folder')} />
              </Field>
            ) : (
              <Field label={t('Password')}>
                <Input type="password" value={password} onChange={(e) => setPassword(e.target.value)} onKeyDown={onKeyDown} autoFocus aria-label={t('Password that locked this folder')} />
              </Field>
            )}
            <button type="button" className="lk-link lk-open-switch" style={{ alignSelf: 'flex-start' }} onClick={() => { setUseRecovery(!useRecovery); setError(null); }}>{useRecovery ? t('Use the password') : t('Use the recovery key')}</button>
          </>
        )}
        {error ? <div className="lk-error" role="alert">{error}</div> : null}
      </div>
    </Dialog>
  );
}

/** Locking a folder for the first time: the choices asked once, with a suggestion on each. */
function LockNewFolder({ shell, root, onClose, onAdded }) {
  const toast = useToast();
  const [c, set] = useState({ what: 'every', types: '.pdf, .docx, .xlsx, .jpg', hidden: true, afterLock: 'delete', keepDir: null, keepVersions: 3, openIn: 'beside' });
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState(null);
  const [more, setMore] = useState(false);
  const cloud = whichCloud(root);
  const pickKeep = async () => {
    const paths = await shell.dialog.open({ title: t('Choose a folder for the plain files'), directory: true });
    if (paths?.[0]) set({ ...c, keepDir: paths[0] });
  };
  const can = !busy && (c.afterLock !== 'keep' || c.keepDir);
  const add = async () => {
    setBusy(true);
    setError(null);
    try {
      const out = await shell.locker.addFolder({ root, choices: { hidden: c.hidden, afterLock: c.afterLock, keepDir: c.keepDir, keepVersions: c.keepVersions, openIn: c.openIn } });
      if (c.what === 'types') {
        const lines = c.types.split(',').map((s) => s.trim()).filter(Boolean).map((s) => (s.startsWith('.') ? `*${s}` : s));
        await shell.locker.writeRules({ id: out.id, text: `# What to lock in this folder, in the wording of .gitignore.\n${lines.join('\n')}\n` });
      }
      const res = await shell.locker.lockWaiting({ id: out.id });
      if (res.failed.length) toast?.(t('{n} files could not be locked yet. They are shown as not locked.', { n: res.failed.length }), { tone: 'warn' });
      onAdded(out.id);
    } catch (err) {
      setError(said(err));
      setBusy(false);
    }
  };
  return (
    <Dialog
      title={t('Lock the folder {name}', { name: root.split(/[\\/]/).pop() })}
      width={760}
      className="lk-add-dialog"
      onClose={onClose}
      actions={<>
        <Button label={t('Cancel')} onClick={onClose} />
        <Button primary className="lk-add-go" label={busy ? t('Locking…') : t('Lock this folder')} disabled={!can} onClick={add} />
      </>}
    >
      <div className="lk-dialog-body">
        <div className="lk-card-sub">{root}</div>
        <div>{t('Every file in this folder is locked where it is. Files you add later stay unlocked until you lock them.')}</div>
        {more ? (
          <>
            <Choices c={c} set={set} keepDirPick={pickKeep} showWhat />
            <div className="lk-card-sub">{t('Finer rules, with exceptions, go in the folder\'s .rutbalocker file, in the same wording as .gitignore.')}</div>
          </>
        ) : <button type="button" className="lk-link lk-change" onClick={() => setMore(true)} style={{ alignSelf: 'flex-start' }}>{t('Change how')}</button>}
        {cloud ? (
          <div className="lk-warn" role="note">
            <Icon name="warn" size={16} />
            <span>{t('This folder is in {cloud}, which keeps earlier versions of each file. Those earlier plain versions stay there after locking; clear them in {cloud}\'s version history. While a file is open beside itself, {cloud} can upload its plain copy too.', { cloud })}</span>
          </div>
        ) : null}
        {error ? <div className="lk-error" role="alert">{error}</div> : null}
      </div>
    </Dialog>
  );
}

/** A folder's choices, changed later. */
function FolderSettings({ shell, folder, onClose }) {
  const [c, set] = useState({ ...folder.choices, what: 'every', types: '' });
  const [error, setError] = useState(null);
  const pickKeep = async () => {
    const paths = await shell.dialog.open({ title: t('Choose a folder for the plain files'), directory: true });
    if (paths?.[0]) set({ ...c, keepDir: paths[0] });
  };
  const save = async () => {
    try {
      await shell.locker.setChoices({ id: folder.id, patch: { hidden: c.hidden, afterLock: c.afterLock, keepDir: c.keepDir, keepVersions: c.keepVersions, openIn: c.openIn } });
      onClose();
    } catch (err) {
      setError(said(err));
    }
  };
  return (
    <Dialog title={t('Folder settings')} width={700} onClose={onClose} actions={<><Button label={t('Cancel')} onClick={onClose} /><Button primary label={t('Save')} onClick={save} /></>}>
      <div className="lk-dialog-body">
        <div className="lk-card-sub">{folder.root}</div>
        <Choices c={c} set={set} keepDirPick={pickKeep} />
        <div className="lk-card-sub">{t('A change to names applies to files locked from now on.')}</div>
        {error ? <div className="lk-error" role="alert">{error}</div> : null}
      </div>
    </Dialog>
  );
}

function Rules({ shell, folder, onClose }) {
  const [text, setText] = useState('');
  const [ready, setReady] = useState(false);
  useEffect(() => {
    shell.locker.readRules({ id: folder.id }).then((v) => { setText(v); setReady(true); });
  }, [shell, folder.id]);
  const save = async () => {
    await shell.locker.writeRules({ id: folder.id, text });
    onClose();
  };
  return (
    <Dialog title={t('Edit rules for {name}', { name: folder.name })} width={560} onClose={onClose} actions={<><Button label={t('Cancel')} onClick={onClose} /><Button primary label={t('Save')} disabled={!ready} onClick={save} /></>}>
      <div className="lk-dialog-body">
        <div className="lk-card-sub">{t('Lines say what to lock; a line starting with ! makes an exception. With no lines, every file is locked.')}</div>
        <textarea className="rw-input lk-rule-area" value={text} spellCheck={false} onChange={(e) => setText(e.target.value)} placeholder={t('For example:\n*.pdf\n!Public/')} aria-label={t('Rules')} />
      </div>
    </Dialog>
  );
}

function Share({ shell, folder, entry, onClose, toast }) {
  const [password, setPassword] = useState('');
  const [error, setError] = useState(null);
  const [saved, setSaved] = useState(null);
  const go = async () => {
    const dest = await shell.dialog.save({ title: t('Send to someone'), defaultPath: `${entry.name}.rlock` });
    if (!dest) return;
    try {
      const out = await shell.locker.exportShare({ id: folder.id, rel: entry.rel, dest, password });
      setPassword('');
      setError(null);
      setSaved({ dest, size: out.size });
    } catch (err) {
      setError(said(err));
    }
  };
  const mailIt = (file) => shell.win.create({ app: 'mail', query: { subject: t('A locked file for you'), attach: file.dest, size: file.size } });
  const email = () => mailIt(saved).catch((err) => setError(said(err)));
  const emailNow = async () => {
    try {
      await mailIt(await shell.locker.shareForMail({ id: folder.id, rel: entry.rel, password }));
      toast?.(t('Give the password a different way, such as a call or a text message.'));
      onClose();
    } catch (err) {
      setError(said(err));
    }
  };
  if (saved) {
    return (
      <Dialog title={t('{name} is ready to send', { name: entry.name })} width={460} onClose={onClose} actions={<><Button className="lk-share-close" label={t('Close')} onClick={onClose} /><Button className="lk-share-show" label={t('Show in Explorer')} onClick={() => shell.shell.showInFolder({ path: saved.dest })} /><Button primary className="lk-share-mail" label={t('Attach to an email')} onClick={email} /></>}>
        <div className="lk-dialog-body lk-share-done">
          <div>{t('The locked copy is saved. Send it any way you like.')}</div>
          <div className="lk-card-sub">{t('Give the password a different way, such as a call or a text message, and never in the same email. The other person opens the file with Open a sent file in their Locker.')}</div>
          {error ? <div className="lk-error" role="alert">{error}</div> : null}
        </div>
      </Dialog>
    );
  }
  return (
    <Dialog title={t('Send {name} to someone', { name: entry.name })} width={440} onClose={onClose} actions={<><Button label={t('Cancel')} onClick={onClose} /><Button label={t('Save the locked copy…')} disabled={password.length < 4} onClick={go} /><Button primary className="lk-share-email" label={t('Attach to an email')} disabled={password.length < 4} onClick={emailNow} /></>}>
      <div className="lk-dialog-body">
        <Field label={t('A password for this one file')}><Input type="password" value={password} onChange={(e) => setPassword(e.target.value)} aria-label={t('A password for this one file')} /></Field>
        <div className="lk-card-sub">{t('Attach it to an email here, or save it to send another way. Send the password separately from the file.')}</div>
        {error ? <div className="lk-error" role="alert">{error}</div> : null}
      </div>
    </Dialog>
  );
}

/** A new name for one file, in the same place. */
function RenameFile({ shell, folder, entry, onClose, onDone }) {
  const [name, setName] = useState(entry.name);
  const [error, setError] = useState(null);
  const ready = Boolean(name.trim()) && name.trim() !== entry.name;
  const go = async () => {
    try {
      await shell.locker.renameFile({ id: folder.id, rel: entry.rel, name, state: entry.state });
      onDone();
    } catch (err) {
      setError(said(err));
    }
  };
  return (
    <Dialog title={t('Rename {name}', { name: entry.name })} width={440} onClose={onClose} actions={<><Button label={t('Cancel')} onClick={onClose} /><Button primary className="lk-rename-go" label={t('Rename')} disabled={!ready} onClick={go} /></>}>
      <div className="lk-dialog-body">
        <Field label={t('New name')}><Input value={name} autoFocus onChange={(e) => setName(e.target.value)} aria-label={t('New name')} onKeyDown={(e) => { if (e.key === 'Enter' && ready) go(); }} /></Field>
        {error ? <div className="lk-error" role="alert">{error}</div> : null}
      </div>
    </Dialog>
  );
}

/** A file someone sent: its own password, then a plain copy, or into one of this person's locked folders. Needs no locker for the plain copy. */
function OpenSent({ shell, state, file, into, onClose, toast }) {
  const [password, setPassword] = useState('');
  const folders = state?.open ? state.folders : [];
  const [folderId, setFolderId] = useState(into || folders[0]?.id || '');
  const [error, setError] = useState(null);
  const [busy, setBusy] = useState(false);
  const [saved, setSaved] = useState(null);
  const plainName = file.split(/[\\/]/).pop().replace(/\.rlock$/i, '');
  const run = async (work) => {
    setBusy(true);
    setError(null);
    try {
      await work();
    } catch (err) {
      setError(said(err));
    } finally {
      setBusy(false);
    }
  };
  const savePlain = () => run(async () => {
    const dest = await shell.dialog.save({ title: t('Save a plain copy'), defaultPath: plainName });
    if (!dest) return;
    await shell.locker.openShared({ file, dest, password });
    setPassword('');
    setSaved(dest);
  });
  const addToFolder = () => run(async () => {
    const out = await shell.locker.importShared({ file, id: folderId, password });
    toast?.(t('{name} is locked in {folder}.', { name: out.rel, folder: folders.find((f) => f.id === folderId)?.name || '' }));
    onClose();
  });
  if (saved) {
    return (
      <Dialog title={t('{name} is open', { name: plainName })} width={440} onClose={onClose} actions={<><Button className="lk-sent-close" label={t('Close')} onClick={onClose} /><Button primary className="lk-sent-show" label={t('Show in Explorer')} onClick={() => shell.shell.showInFolder({ path: saved })} /></>}>
        <div className="lk-dialog-body"><div>{t('A plain copy is saved. That copy is not locked.')}</div></div>
      </Dialog>
    );
  }
  const can = Boolean(password) && !busy;
  return (
    <Dialog title={t('Open {name}', { name: plainName })} width={440} onClose={onClose} actions={<><Button label={t('Cancel')} onClick={onClose} /><Button primary={!folders.length} label={t('Save a plain copy…')} disabled={!can} onClick={savePlain} />{folders.length ? <Button primary label={t('Add to this folder')} disabled={!can || !folderId} onClick={addToFolder} /> : null}</>}>
      <div className="lk-dialog-body">
        <Field label={t('The password for this file')}><Input type="password" value={password} autoFocus onChange={(e) => setPassword(e.target.value)} aria-label={t('The password for this file')} /></Field>
        {folders.length ? (
          <Field label={t('Keep it locked in')}>
            <Select value={folderId} onChange={(e) => setFolderId(e.target.value)} aria-label={t('Keep it locked in')}>
              {folders.map((f) => <option key={f.id} value={f.id}>{f.name}</option>)}
            </Select>
          </Field>
        ) : <div className="lk-card-sub">{state?.exists ? t('Unlock your locker first if you want to keep it in a locked folder.') : t('You do not need a locker of your own to open this file.')}</div>}
        {error ? <div className="lk-error" role="alert">{error}</div> : null}
      </div>
    </Dialog>
  );
}

function Compare({ shell, folder, list, onClose }) {
  const conflicts = list.filter((e) => e.state === 'conflict');
  const exportOne = async (e) => {
    const dest = await shell.dialog.save({ title: t('Save an unlocked copy'), defaultPath: e.name });
    if (dest) await shell.locker.exportCopy({ id: folder.id, rel: e.rel, dest, overwrite: true, allowConflict: true });
  };
  return (
    <Dialog title={t('Two versions')} width={640} onClose={onClose} actions={<Button label={t('Close')} onClick={onClose} />}>
      <div className="lk-dialog-body">
        <div className="lk-card-sub">{t('Your cloud kept both copies when two computers saved the file. Neither has been changed. Export each to compare them, then delete the one you do not want in File Explorer.')}</div>
        <div className="lk-choices">
          {conflicts.map((e) => (
            <div key={e.rel} className="lk-choice" style={{ flexDirection: 'column', cursor: 'default' }}>
              <b>{e.name}</b>
              <small>{formatBytes(e.size)} · {formatWhen(e.modified)}</small>
              <Button label={t('Save an unlocked copy…')} onClick={() => exportOne(e)} />
            </div>
          ))}
        </div>
      </div>
    </Dialog>
  );
}

function Settings({ shell, state, onClose, toast }) {
  const [minutes, setMinutes] = useState(state.idleMinutes);
  const [current, setCurrent] = useState('');
  const [next, setNext] = useState('');
  const [error, setError] = useState(null);
  const [recovery, setRecovery] = useState(null);
  const save = async () => {
    await shell.locker.setIdle({ minutes });
    onClose();
  };
  const changePw = async () => {
    setError(null);
    try {
      await shell.locker.changePassword({ current, next });
      setCurrent('');
      setNext('');
      toast?.(t('The password was changed.'));
    } catch (err) {
      setError(said(err));
    }
  };
  const newRecovery = async () => {
    setError(null);
    try {
      const out = await shell.locker.replaceRecovery({ current });
      setRecovery(out.recovery);
      setCurrent('');
    } catch (err) {
      setError(said(err));
    }
  };
  const hello = async (on) => {
    setError(null);
    try {
      if (on) await shell.locker.enableHello();
      else await shell.locker.disableHello();
    } catch (err) {
      setError(said(err));
    }
  };
  return (
    <Dialog title={t('Locker settings')} width={520} onClose={onClose} actions={<><Button label={t('Cancel')} onClick={onClose} /><Button primary label={t('Save')} onClick={save} /></>}>
      <div className="lk-dialog-body">
        <Field label={t('Lock again after this many idle minutes')}>
          <Select value={String(minutes)} onChange={(e) => setMinutes(Number(e.target.value))} aria-label={t('Idle minutes')}>
            {[5, 10, 15, 30, 60, 120].map((m) => <option key={m} value={String(m)}>{m}</option>)}
          </Select>
        </Field>
        <label className="lk-advanced" style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
          <input type="checkbox" checked={state.advanced === true} onChange={(e) => shell.locker.setAdvanced({ on: e.target.checked })} />
          <span>{t('Show advanced options')}</span>
          <span className="lk-card-sub">{t('Folder rules, folder settings, sizes and reasons')}</span>
        </label>
        <Field label={t('When the Locker locks, temporary copies of files you opened')}>
          <Select value={state.keepCopies ? 'keep' : 'remove'} onChange={(e) => shell.locker.setWhenLocked({ keepCopies: e.target.value === 'keep' })} aria-label={t('Temporary copies')}>
            <option value="remove">{t('Remove them (suggested)')}</option>
            <option value="keep">{t('Keep them until the next unlock')}</option>
          </Select>
        </Field>
        {state.keepCopies ? <div className="lk-warn" role="note"><Icon name="warn" size={16} /><span>{t('Anyone at this computer can read those copies until you unlock the Locker again.')}</span></div> : null}
        <Field label={t('When the Locker locks, files you have unlocked')}>
          <Select value={state.relock ? 'lock' : 'leave'} onChange={(e) => shell.locker.setWhenLocked({ relock: e.target.value === 'lock' })} aria-label={t('Unlocked files')}>
            <option value="leave">{t('Leave them unlocked (suggested)')}</option>
            <option value="lock">{t('Lock them again')}</option>
          </Select>
        </Field>
        {state.hello ? (
          <div className="lk-actions">
            <Button label={state.helloOn ? t('Turn off Windows Hello') : t('Turn on Windows Hello')} onClick={() => hello(!state.helloOn)} />
            <span className="lk-card-sub">{t('On this computer only. The password stays the master.')}</span>
          </div>
        ) : null}
        <Field label={t('Current password')}><Input type="password" value={current} onChange={(e) => setCurrent(e.target.value)} aria-label={t('Current password')} /></Field>
        <Field label={t('New password')}><Input type="password" value={next} onChange={(e) => setNext(e.target.value)} aria-label={t('New password')} /></Field>
        <div className="lk-actions">
          <Button label={t('Change password')} disabled={!current || next.length < 8} onClick={changePw} />
          <Button label={t('Replace the recovery key')} disabled={!current} onClick={newRecovery} />
        </div>
        {recovery ? <div className="lk-key">{recovery}</div> : null}
        {error ? <div className="lk-error" role="alert">{error}</div> : null}
      </div>
    </Dialog>
  );
}

export default function Locker({ app, shell }) {
  const toast = useToast();
  const [state, setState, refresh] = useLocker(shell);
  const [current, setCurrent] = useState(null);
  const [entries, setEntries] = useState([]);
  const [selected, setSelected] = useState(null);
  const [query, setQuery] = useState('');
  const [dialog, setDialog] = useState(null);
  const [tick, setTick] = useState(0);
  // Setting up stays on screen until the person has seen the recovery key: making
  // the locker opens it, and the state that says so must not take the key away.
  const [settingUp, setSettingUp] = useState(false);
  useEffect(() => { if (state && !state.exists) setSettingUp(true); }, [state]);
  const menu = useAppMenu({ shell, appKey: 'locker' });

  // The countdown, once a minute.
  useEffect(() => {
    const id = setInterval(() => setTick((n) => n + 1), 30000);
    return () => clearInterval(id);
  }, []);

  // Clicking or typing here is using the Locker, which puts off the idle lock. Sent at most every few seconds.
  useEffect(() => {
    let last = 0;
    const used = () => {
      const now = Date.now();
      if (now - last < 5000) return;
      last = now;
      shell.locker.touch().then(({ lockAt }) => setState((s) => (s?.open ? { ...s, lockAt } : s))).catch(() => {});
    };
    const kinds = ['pointerdown', 'keydown', 'wheel'];
    kinds.forEach((k) => window.addEventListener(k, used, true));
    return () => kinds.forEach((k) => window.removeEventListener(k, used, true));
  }, [shell, setState]);

  useEffect(() => shell.on('locker:changed', (e) => {
    if (e?.left?.length) toast?.(t('Still open in another program, so not cleared yet: {files}. Close it to clear the temporary copy.', { files: e.left.join(', ') }), { tone: 'warn' });
    if (e?.open?.length) toast?.(t('Still open in Office: {files}. Keep working and save as usual. Your saves are locked in when you next unlock.', { files: e.open.join(', ') }), { tone: 'warn' });
    if (e?.notLocked?.length) toast?.(t('{n} files could not be locked again. They are listed as not locked.', { n: e.notLocked.length }), { tone: 'warn' });
  }), [shell, toast]);

  useEffect(() => {
    if (!state?.open) return;
    const off = shell.on('locker:changed', (e) => {
      if (e?.problem) toast?.(e.problem, { tone: 'error' });
      if (current) shell.locker.list({ id: current }).then(setEntries).catch(() => {});
      refresh();
    });
    return off;
  }, [shell, state?.open, current, refresh]);

  useEffect(() => {
    if (!state?.open) return;
    if (!state.folders.length) {
      setCurrent(null);
      return;
    }
    setCurrent((c) => (c && state.folders.some((f) => f.id === c) ? c : state.folders[0].id));
  }, [state]);

  const reload = useCallback(() => {
    if (!current) {
      setEntries([]);
      return;
    }
    shell.locker.list({ id: current }).then((list) => {
      setEntries(list);
      setSelected((sel) => (list.some((e) => e.rel === sel) ? sel : null));
    }).catch(() => setEntries([]));
  }, [shell, current]);
  useEffect(() => { reload(); }, [reload]);

  const folder = useMemo(() => state?.folders?.find((f) => f.id === current) || null, [state, current]);
  const selEntry = entries.find((e) => e.rel === selected) || null;
  const shown = matching(entries, query);

  const addFolder = async () => {
    const paths = await shell.dialog.open({ title: t('Choose a folder to lock'), directory: true });
    if (paths?.[0]) setDialog({ kind: 'add', root: paths[0] });
  };

  const makeHomeFolder = () => act(async () => {
    const out = await shell.locker.makeHomeFolder();
    setCurrent(out.id);
    toast?.(out.made ? t('Your Locker folder is ready. Put files in it, then lock them here.') : t('This is your Locker folder.'));
  });

  const importFiles = async () => {
    const paths = await shell.dialog.open({ title: t('Add files'), multiple: true });
    if (!paths?.length || !current) return;
    for (const source of paths) {
      try {
        await shell.locker.importFile({ id: current, source });
      } catch (err) {
        toast?.(said(err), { tone: 'error' });
      }
    }
    reload();
  };

  const act = async (fn) => {
    try {
      await fn();
    } catch (err) {
      toast?.(said(err), { tone: 'error' });
    }
    reload();
  };

  const openEntry = (e) => act(async () => {
    const out = await shell.locker.open({ id: current, rel: e.rel });
    if (out.app) await shell.win.create({ app: out.app, file: out.path });
    else await shell.shell.openPath({ path: out.path });
  });

  const exportCopy = (e) => act(async () => {
    const dest = await shell.dialog.save({ title: t('Save an unlocked copy'), defaultPath: e.name });
    if (dest) await shell.locker.exportCopy({ id: current, rel: e.rel, dest, overwrite: true });
  });

  const unlockForGood = (e) => act(async () => {
    const { response } = await shell.dialog.message({ type: 'warning', message: t('Unlock {name}?', { name: e.name }), detail: t('It becomes an ordinary file in the same place, and the earlier versions the Locker kept of it are removed. You can lock it again whenever you like.'), buttons: [t('Unlock'), t('Cancel')], defaultId: 1, cancelId: 1 });
    if (response === 0) {
      await shell.locker.unlockForGood({ id: current, rel: e.rel });
      setSelected(null);
    }
  });

  const deleteFile = (e) => act(async () => {
    const { response } = await shell.dialog.message({ type: 'warning', message: t('Delete {name}?', { name: e.name }), detail: t('It is removed for good, with the earlier versions the Locker kept of it. This cannot be undone.'), buttons: [t('Delete'), t('Cancel')], defaultId: 1, cancelId: 1 });
    if (response === 0) {
      await shell.locker.deleteFile({ id: current, rel: e.rel, state: e.state });
      setSelected(null);
    }
  });

  /** One file or a subfolder's: each tried, so one a program holds does not stop the rest. */
  const lockFiles = (list) => act(async () => {
    let problem = null;
    for (const e of list) {
      try {
        await shell.locker.lockOne({ id: current, rel: e.rel });
      } catch (err) {
        problem = err;
      }
    }
    setSelected(null);
    if (problem) throw problem;
  });

  const removeFolder = () => act(async () => {
    const { response } = await shell.dialog.message({ type: 'warning', message: t('Stop locking {name}?', { name: folder.name }), detail: t('Every file in it is unlocked first, and the earlier versions it kept are removed. The folder itself stays where it is.'), buttons: [t('Stop locking'), t('Cancel')], defaultId: 1, cancelId: 1 });
    if (response === 0) {
      const out = await shell.locker.removeFolder({ id: current });
      if (!out.removed) toast?.(t('{n} files could not be unlocked, so the folder was not removed.', { n: out.failed.length }), { tone: 'error' });
    }
  });

  const lockNow = () => shell.locker.lock().then(setState).catch(() => {});

  const openSent = async () => {
    const paths = await shell.dialog.open({ title: t('Open a sent file'), filters: [{ name: t('Locked files'), extensions: ['rlock'] }] });
    if (paths?.[0]) setDialog({ kind: 'sent', file: paths[0] });
  };

  const minutesLeft = state?.lockAt ? Math.max(0, Math.ceil((state.lockAt - Date.now()) / 60000)) : null;
  void tick;

  let inner;
  if (!state) inner = null;
  else if (!state.exists || settingUp) inner = <Setup shell={shell} onDone={() => { setSettingUp(false); refresh(); }} onOpenSent={openSent} />;
  else if (!state.open) inner = <Unlock shell={shell} state={state} onOpened={setState} onOpenSent={openSent} />;
  else {
    const conflicts = entries.filter((e) => e.state === 'conflict');
    const waiting = entries.filter((e) => e.state === 'waiting').length;
    const opts = folder?.choices;
    const describe = folder && [
      opts.hidden ? NAMES.hidden() : NAMES.visible(),
      opts.afterLock === 'delete' ? NAMES.delete() : NAMES.keep(),
      t('rules from .rutbalocker'),
    ].join(' · ');
    inner = (
      <>
        <div className="lk-bar">
          <span className="lk-pill"><Icon name="unlock" size={14} />{minutesLeft !== null ? t('Unlocked · locks again in {n} min', { n: minutesLeft }) : t('Unlocked')}</span>
          <Button className="lk-lock" label={t('Lock now')} onClick={lockNow} />
          <Spacer />
          <Button className="lk-add" label={t('Add folder…')} onClick={addFolder} />
          <Button className="lk-import" label={t('Add files…')} disabled={!current} onClick={importFiles} />
          <Button className="lk-open-sent" label={t('Open a sent file…')} onClick={openSent} />
          <Button className="lk-settings" label={t('Settings')} onClick={() => setDialog({ kind: 'settings' })} />
        </div>
        <div className="lk-body">
          <aside className="lk-side" aria-label={t('Locked folders')}>
            <span className="lk-side-head">{t('Locked folders')}</span>
            {state.folders.map((f) => (
              <button key={f.id} type="button" className={`lk-folder${f.id === current ? ' on' : ''}`} data-folder={f.name} onClick={() => setCurrent(f.id)}>
                <Icon name="folder" size={18} />
                <span className="lk-folder-text">
                  <b>{f.name}</b>
                  <small>{state.advanced ? `${f.root} · ` : ''}{t('{n} files', { n: (f.summary?.locked ?? 0) + (f.summary?.conflict ?? 0) })}</small>
                </span>
                {f.summary?.waiting ? <span className="lk-state waiting">{t('{n} not locked', { n: f.summary.waiting })}</span> : null}
              </button>
            ))}
            {!state.folders.length ? <div className="lk-note">{t('No folders yet. Add one to lock its files where they are.')}</div> : null}
            <div className="lk-note foot">{t('Files are locked one by one where they are.')}</div>
          </aside>
          <main className="lk-main">
            {folder ? (
              <>
                <div className="lk-head">
                  <h1>{folder.name}</h1>
                  {state.advanced ? <span className="lk-head-note">{describe}</span> : null}
                  <Spacer />
                  {waiting ? <Button className="lk-lockall" label={t('Lock {n} now', { n: waiting })} onClick={() => act(() => shell.locker.lockWaiting({ id: current }))} /> : null}
                  <Button className="lk-show" label={t('Show in Explorer')} onClick={() => shell.shell.openPath({ path: folder.root })} />
                  {state.advanced ? <Button className="lk-rules" label={t('Edit rules')} onClick={() => setDialog({ kind: 'rules' })} /> : null}
                  {state.advanced ? <Button className="lk-folder-settings" label={t('Folder settings')} onClick={() => setDialog({ kind: 'folder' })} /> : null}
                  <Button className="lk-remove" label={t('Stop locking this folder…')} onClick={removeFolder} />
                </div>
                <div className="lk-search">
                  <Input value={query} placeholder={t('Search files in this folder')} aria-label={t('Search files')} onChange={(e) => setQuery(e.target.value)} onKeyDown={(e) => { if (e.key === 'Escape') setQuery(''); }} />
                  {query.trim() ? <span className="lk-card-sub lk-search-count">{t('{n} of {total} files', { n: shown.length, total: entries.length })}</span> : null}
                  {query ? <Button className="lk-search-clear" label={t('Clear')} onClick={() => setQuery('')} /> : null}
                </div>
                {conflicts.length ? (
                  <div className="lk-alert" role="alert">
                    <span style={{ flex: 1 }}><b>{t('Two versions of {name}.', { name: conflicts[0].name })}</b> {t('Your cloud made a conflicted copy when two computers saved it.')}</span>
                    <Button className="lk-compare" label={t('Compare…')} onClick={() => setDialog({ kind: 'compare' })} />
                  </div>
                ) : null}
                <div className="lk-table-wrap">
                  <table className="lk-table">
                    <thead><tr><th>{t('Name')}</th><th>{t('State')}</th><th>{t('Changed')}</th>{state.advanced ? <><th>{t('Size')}</th><th>{t('Why')}</th></> : null}<th aria-label={t('Actions')} /></tr></thead>
                    <tbody>
                      {byDirectory(shown).map(([dir, files]) => {
                        const cols = state.advanced ? 6 : 4;
                        const notLocked = files.filter((e) => e.state === 'waiting');
                        const lockedHere = files.filter((e) => e.state === 'locked').length;
                        return (
                          <React.Fragment key={dir || '.'}>
                            {dir ? (
                              <tr className="lk-dir" data-dir={dir}>
                                <td colSpan={cols}>
                                  <div className="lk-dir-in">
                                    <Icon name="folder" size={16} />
                                    <b title={dir}>{dir.split('/').join(' / ')}</b>
                                    <span className="lk-card-sub">{t('{n} locked', { n: lockedHere })}{notLocked.length ? ` · ${t('{n} not locked', { n: notLocked.length })}` : ''}</span>
                                    <Spacer />
                                    {notLocked.length ? <Button className="lk-lockdir" label={t('Lock {n} here', { n: notLocked.length })} onClick={() => lockFiles(notLocked)} /> : null}
                                  </div>
                                </td>
                              </tr>
                            ) : null}
                            {files.map((e) => (
                              <React.Fragment key={e.rel}>
                                <tr data-rel={e.rel} data-state={e.state} className={e.rel === selected ? 'sel' : ''} onClick={() => setSelected(e.rel)} onDoubleClick={() => { if (e.state === 'locked') openEntry(e); }}>
                                  <td title={e.rel} className={dir ? 'lk-indent' : ''}>{e.name}</td>
                                  <td><span className={`lk-state ${e.state}`}>{stateName(e)}</span></td>
                                  <td>{e.modified ? formatWhen(e.modified) : ''}</td>
                                  {state.advanced ? (
                                    <>
                                      <td>{e.size !== undefined ? formatBytes(e.size) : ''}</td>
                                      <td className="lk-why" title={whyText(e)}>{whyText(e)}</td>
                                    </>
                                  ) : null}
                                  <td className="lk-row-acts" onClick={(ev) => ev.stopPropagation()} onDoubleClick={(ev) => ev.stopPropagation()}>
                                    <div>
                                      {e.state === 'locked' ? <Button primary className="lk-open" label={t('Open')} onClick={() => openEntry(e)} /> : null}
                                      {e.state === 'locked' ? <Button className="lk-unlock" label={t('Unlock…')} onClick={() => unlockForGood(e)} /> : null}
                                      {e.state === 'waiting' ? <Button primary className="lk-lockfile" label={t('Lock')} onClick={() => lockFiles([e])} /> : null}
                                    </div>
                                  </td>
                                </tr>
                                {e.rel === selected ? (
                                  <tr className="lk-more">
                                    <td colSpan={cols}>
                                      <div className="lk-more-in">
                                        <span className="lk-detail lk-card-sub">{e.rel} · {stateDetail(e)}{e.size !== undefined ? ` · ${formatBytes(e.size)}` : ''}{whyText(e) ? ` · ${whyText(e)}` : ''}</span>
                                        <Spacer />
                                        {e.state === 'locked' ? <Button className="lk-export" label={t('Save an unlocked copy…')} onClick={() => exportCopy(e)} /> : null}
                                        {e.state === 'locked' ? <Button className="lk-share" label={t('Send to someone…')} onClick={() => setDialog({ kind: 'share', entry: e })} /> : null}
                                        {e.state === 'locked' || e.state === 'waiting' || e.state === 'excluded' ? <Button className="lk-rename" label={t('Rename…')} onClick={() => setDialog({ kind: 'rename', entry: e })} /> : null}
                                        {e.state !== 'conflict' ? <Button className="lk-delete" label={t('Delete…')} onClick={() => deleteFile(e)} /> : null}
                                      </div>
                                    </td>
                                  </tr>
                                ) : null}
                              </React.Fragment>
                            ))}
                          </React.Fragment>
                        );
                      })}
                    </tbody>
                  </table>
                  {!entries.length ? <Empty title={t('Nothing in this folder yet')} /> : null}
                  {entries.length && !shown.length ? <Empty title={t('No files match that search')} /> : null}
                </div>
              </>
            ) : (
              <Empty
                icon="lock"
                title={t('Make a safe folder')}
                action={<><Button primary className="lk-make" label={t('Make my Locker folder')} onClick={makeHomeFolder} /> <Button label={t('Lock a folder I already have…')} onClick={addFolder} /></>}
              />
            )}
          </main>
        </div>
      </>
    );
  }

  return (
    <AppFrame app={app} shell={shell} title={t('Locker')} menu={menu}>
      <style>{CSS}</style>
      <div className="lk">{inner}</div>
      {dialog?.kind === 'add' ? <AddFolder shell={shell} root={dialog.root} onClose={() => setDialog(null)} onAdded={(id) => { setDialog(null); setCurrent(id); refresh(); }} /> : null}
      {dialog?.kind === 'folder' && folder ? <FolderSettings shell={shell} folder={folder} onClose={() => { setDialog(null); refresh(); }} /> : null}
      {dialog?.kind === 'rules' && folder ? <Rules shell={shell} folder={folder} onClose={() => { setDialog(null); reload(); refresh(); }} /> : null}
      {dialog?.kind === 'rename' && folder ? <RenameFile shell={shell} folder={folder} entry={dialog.entry} onClose={() => setDialog(null)} onDone={() => { setDialog(null); setSelected(null); reload(); }} /> : null}
      {dialog?.kind === 'share' && folder ? <Share shell={shell} folder={folder} entry={dialog.entry} toast={toast} onClose={() => setDialog(null)} /> : null}
      {dialog?.kind === 'sent' ? <OpenSent shell={shell} state={state} file={dialog.file} into={current} toast={toast} onClose={() => { setDialog(null); refresh(); }} /> : null}
      {dialog?.kind === 'compare' && folder ? <Compare shell={shell} folder={folder} list={entries} onClose={() => setDialog(null)} /> : null}
      {dialog?.kind === 'settings' ? <Settings shell={shell} state={state} toast={toast} onClose={() => { setDialog(null); refresh(); }} /> : null}
    </AppFrame>
  );
}
