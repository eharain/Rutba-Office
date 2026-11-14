// Review → Restrict Editing: the pane Word opens on the right, and the two
// dialogs it leads to — Start Enforcing Protection (an optional password,
// typed twice) and Unprotect Document (the password asked for).
//
// Before protection is enforced the pane holds Word's three numbered steps:
// formatting limited to styles, the one kind of editing allowed (No changes,
// Tracked changes, Comments, Filling in forms) with Everyone as an exception
// for the selected paragraphs, and "Yes, Start Enforcing Protection". Once
// enforced it says what this person may do where the caret is, finds the
// next region they can edit, highlights them all, and offers Stop Protection.

import React, { useEffect, useState } from 'react';
import { Dialog, Button, Icon, Select, t, tn } from '@rutba/office-ui';
import { SecretInput, PROTECT_CSS } from '../../protect.js';

export const EDIT_KINDS = [
  { id: 'readOnly', label: t('word.restrict.no-changes-read') },
  { id: 'trackedChanges', label: t('word.tracked-changes') },
  { id: 'comments', label: t('common.comments') },
  { id: 'forms', label: t('word.restrict.filling-in-forms') },
];

/** What Word's pane says about the caret's place, under each kind of protection. */
export function permissionSentence(protection, editable) {
  switch (protection?.edit) {
    case 'trackedChanges':
      return t('word.restrict.document-protected-unintentional-editing-may');
    case 'none':
      return t('word.restrict.document-protected-unintentional-editing-may-edit');
    case 'comments':
      return editable ? t('word.restrict.document-protected') : t('word.restrict.document-protected-unintentional-editing-may-only');
    case 'forms':
      return t('word.restrict.document-protected-unintentional-editing');
    default:
      return editable ? t('word.restrict.document-protected') : t('word.restrict.document-protected-unintentional');
  }
}

/** The sentence the page shows when a keystroke is refused. */
export function refusalSentence(protection) {
  switch (protection?.edit) {
    case 'comments':
      return t('word.restrict.document-protected-only');
    case 'forms':
      return t('common.document-protected');
    default:
      return t('common.modification-not-allowed');
  }
}

export function RestrictPane({ protection, selection, highlight, onHighlight, onPermission, onStart, onStop, onFindNext, onClose }) {
  const enforced = Boolean(protection?.enforced);
  const [formatting, setFormatting] = useState(Boolean(protection?.formatting));
  const [limit, setLimit] = useState(Boolean(protection && protection.edit !== 'none'));
  const [edit, setEdit] = useState(protection?.edit && protection.edit !== 'none' ? protection.edit : 'readOnly');
  // The file's own settings, when they change under the pane (a reopened
  // file, Stop Protection) — the pane shows what the document holds.
  // Only what the file says is taken up; a document with only exception
  // regions (edit "none") leaves the person's own ticks alone.
  useEffect(() => {
    if (protection?.formatting) setFormatting(true);
    if (protection?.edit && protection.edit !== 'none') {
      setLimit(true);
      setEdit(protection.edit);
    }
  }, [protection?.edit, protection?.formatting, enforced]);

  const from = selection?.from ?? 0;
  const to = selection?.to ?? from;
  const regions = protection?.regions || [];
  const inRegion = regions.some((r) => from >= r.from && to <= r.to);
  const touched = regions.some((r) => r.to >= from && r.from <= to);
  const exceptions = limit && (edit === 'readOnly' || edit === 'comments');

  return (
    <aside className="wd-restrict" aria-label={t('word.restrict')}>
      <div className="wd-restrict-head">
        <span>{t('word.restrict')}</span>
        <button type="button" className="wd-nav-close" onClick={onClose} data-tip={t('word.restrict.close-restrict-editing')} aria-label={t('common.close')}>
          <Icon name="close" size={12} />
        </button>
      </div>
      {enforced ? (
        <div className="wd-restrict-body">
          <section className="wd-restrict-step">
            <h4>{t('word.restrict.your-permissions')}</h4>
            <p className="wd-restrict-perm">{permissionSentence(protection, inRegion || protection.edit === 'trackedChanges' || protection.edit === 'none')}</p>
            {regions.length && protection.edit !== 'trackedChanges' && protection.edit !== 'forms' && protection.edit !== 'none' ? (
              <>
                <Button label={t('word.restrict.find-next')} className="wd-restrict-next" onClick={onFindNext} />
                <label className="wd-restrict-check">
                  <input type="checkbox" className="wd-restrict-highlight" checked={highlight} onChange={(e) => onHighlight(e.target.checked)} />
                  <span>{t('word.restrict.highlight-regions')}</span>
                </label>
              </>
            ) : null}
            {protection.formatting ? <p className="wd-restrict-note">{t('word.restrict.formatting-limited')}</p> : null}
          </section>
          <div className="wd-restrict-foot">
            <Button label={t('word.restrict.stop')} className="wd-restrict-stop" onClick={onStop} />
          </div>
        </div>
      ) : (
        <div className="wd-restrict-body">
          <section className="wd-restrict-step">
            <h4><span className="wd-restrict-n">1</span>{t('word.restrict.formatting')}</h4>
            <label className="wd-restrict-check">
              <input type="checkbox" className="wd-restrict-formatting" checked={formatting} onChange={(e) => setFormatting(e.target.checked)} />
              <span>{t('word.restrict.limit-formatting')}</span>
            </label>
          </section>
          <section className="wd-restrict-step">
            <h4><span className="wd-restrict-n">2</span>{t('word.restrict.editing-restrictions')}</h4>
            <label className="wd-restrict-check">
              <input type="checkbox" className="wd-restrict-limit" checked={limit} onChange={(e) => setLimit(e.target.checked)} />
              <span>{t('word.restrict.allow-only-type')}</span>
            </label>
            <Select className="wd-restrict-kind" value={edit} disabled={!limit} onChange={(e) => setEdit(e.target.value)}>
              {EDIT_KINDS.map((k) => (
                <option key={k.id} value={k.id}>{k.label}</option>
              ))}
            </Select>
            {exceptions ? (
              <div className="wd-restrict-exceptions">
                <div className="wd-restrict-sub">{t('word.restrict.exceptions-optional')}</div>
                <p className="wd-restrict-note">{t('word.restrict.select-parts-document')}</p>
                <div className="wd-restrict-groups">
                  <div className="wd-restrict-sub2">{t('word.restrict.groups')}</div>
                  <label className="wd-restrict-check wd-restrict-everyone-row">
                    <input type="checkbox" className="wd-restrict-everyone" checked={touched} onChange={(e) => onPermission(e.target.checked)} />
                    <span>{t('common.everyone')}</span>
                  </label>
                </div>
                {regions.length ? <p className="wd-restrict-count">{tn(regions.length, 'word.restrict.one-region-marked')}</p> : null}
              </div>
            ) : null}
          </section>
          <section className="wd-restrict-step">
            <h4><span className="wd-restrict-n">3</span>{t('word.restrict.start-enforcement')}</h4>
            <p className="wd-restrict-note">{t('word.restrict.ready-apply-settings')}</p>
            <Button
              primary
              label={t('word.restrict.yes-start')}
              className="wd-restrict-start"
              disabled={!limit && !formatting}
              title={!limit && !formatting ? t('word.restrict.yes-start-enforcing-protection') : t('word.restrict.yes-start-enforcing')}
              onClick={() => onStart({ edit: limit ? edit : 'none', formatting })}
            />
          </section>
        </div>
      )}
    </aside>
  );
}

/** Start Enforcing Protection: Word's password method, the password optional and typed twice. */
export function StartEnforcingDialog({ onClose, onStart }) {
  const [password, setPassword] = useState('');
  const [again, setAgain] = useState('');
  // While the password is hashed (100,000 rounds) OK is held with a spinner.
  const [busy, setBusy] = useState(false);
  const mismatch = again !== '' && again !== password;
  const ok = password === again;
  const submit = async () => {
    if (!ok || busy) return;
    setBusy(true);
    try {
      await onStart(password);
    } finally {
      setBusy(false);
    }
  };
  const enter = (e) => {
    if (e.key === 'Enter') submit();
  };
  return (
    <Dialog
      title={t('word.restrict.start-enforcing')}
      width={440}
      onClose={busy ? undefined : onClose}
      actions={
        <>
          <Button label={t('common.cancel')} className="wd-enforce-cancel" disabled={busy} onClick={onClose} />
          <Button primary label={t('common.ok')} className="wd-enforce-ok" disabled={!ok} busy={busy} onClick={submit} />
        </>
      }
    >
      <style>{PROTECT_CSS}</style>
      <div className="pw-set">
        <p className="pw-set-lead">{t('word.restrict.protection-method')} <b>{t('common.password')}</b>. {t('word.restrict.document-not-encrypted')}</p>
        <label className="pw-label">{t('word.restrict.enter-new-password')}</label>
        <SecretInput className="wd-enforce-password" label={t('word.restrict.enter-new-password')} value={password} onChange={setPassword} onKeyDown={enter} autoFocus />
        <label className="pw-label">{t('word.restrict.reenter-password')}</label>
        <SecretInput className="wd-enforce-again" label={t('word.restrict.reenter-password')} value={again} onChange={setAgain} onKeyDown={enter} invalid={mismatch} />
        {mismatch ? (
          <div className="pw-error" role="alert">
            <Icon name="info" size={14} />
            <span>{t('common.two-passwords-not')}</span>
          </div>
        ) : null}
      </div>
    </Dialog>
  );
}

/** Unprotect Document: the password Stop Protection needs, and Word's refusal when it is wrong. */
export function UnprotectDialog({ error = '', onClose, onSubmit }) {
  const [password, setPassword] = useState('');
  // While the password is checked OK says so with a spinner and cannot be pressed twice.
  const [busy, setBusy] = useState(false);
  useEffect(() => {
    if (error) setPassword('');
  }, [error]);
  const submit = async () => {
    if (busy) return;
    setBusy(true);
    try {
      await onSubmit(password);
    } finally {
      setBusy(false);
    }
  };
  return (
    <Dialog
      title={t('word.restrict.unprotect-document')}
      width={400}
      onClose={busy ? undefined : onClose}
      actions={
        <>
          <Button label={t('common.cancel')} className="wd-unprotect-cancel" disabled={busy} onClick={onClose} />
          <Button primary label={busy ? t('common.checking') : t('common.ok')} className="wd-unprotect-ok" busy={busy} onClick={submit} />
        </>
      }
    >
      <style>{PROTECT_CSS}</style>
      <div className="pw-set">
        <label className="pw-label">{t('common.password')}</label>
        <SecretInput className="wd-unprotect-password" label={t('common.password')} value={password} onChange={setPassword} invalid={Boolean(error) && !password} onKeyDown={(e) => { if (e.key === 'Enter') submit(); }} autoFocus />
        {error ? (
          <div className="pw-error wd-unprotect-error" role="alert">
            <Icon name="info" size={14} />
            <span>{error}</span>
          </div>
        ) : null}
      </div>
    </Dialog>
  );
}

export const RESTRICT_CSS = `
/* A pane beside the page: the flow and the pane side by side. */
.wd.wd-side { flex-direction: row; }
.wd-restrict {
  width: 288px; flex: none; border-left: 1px solid var(--line); background: var(--chrome);
  display: flex; flex-direction: column; min-height: 0;
}
.wd-restrict-head {
  display: flex; align-items: center; justify-content: space-between; padding: 9px 12px;
  font-size: 11px; text-transform: uppercase; letter-spacing: .06em; color: var(--ink-3);
  border-bottom: 1px solid var(--line-soft);
}
.wd-restrict-body { overflow: auto; padding: 4px 0 12px; display: flex; flex-direction: column; min-height: 0; flex: 1; }
.wd-restrict-step { padding: 12px 14px; border-bottom: 1px solid var(--line-soft); display: flex; flex-direction: column; gap: 8px; }
.wd-restrict-step:last-child { border-bottom: 0; }
.wd-restrict-step h4 { margin: 0; display: flex; align-items: center; gap: 8px; font-size: 12.5px; font-weight: 650; color: var(--ink); }
.wd-restrict-n { width: 18px; height: 18px; border-radius: 50%; display: grid; place-items: center; font-size: 10.5px; font-weight: 700; background: var(--accent-soft); color: var(--accent-text); flex: none; }
.wd-restrict-check { display: flex; align-items: flex-start; gap: 8px; font-size: 12.5px; line-height: 1.4; color: var(--ink); cursor: pointer; }
.wd-restrict-check input { margin: 2px 0 0; accent-color: var(--accent); flex: none; }
.wd-restrict .rw-select { width: 100%; }
.wd-restrict-note { margin: 0; font-size: 12px; line-height: 1.45; color: var(--ink-2); }
.wd-restrict-exceptions { display: flex; flex-direction: column; gap: 6px; margin-top: 4px; padding: 10px; border-radius: 8px; background: var(--surface); border: 1px solid var(--line-soft); }
.wd-restrict-sub { font-size: 12px; font-weight: 650; color: var(--ink); }
.wd-restrict-sub2 { font-size: 11.5px; color: var(--ink-3); }
.wd-restrict-groups { display: flex; flex-direction: column; gap: 4px; }
.wd-restrict-everyone-row { padding: 5px 8px; border-radius: 6px; background: var(--surface-2); }
.wd-restrict-count { margin: 0; font-size: 11.5px; color: var(--ink-3); }
.wd-restrict-perm { margin: 0; font-size: 12.5px; line-height: 1.5; color: var(--ink); padding: 9px 10px; border-radius: 8px; background: color-mix(in srgb, var(--warn) 11%, var(--surface)); border: 1px solid color-mix(in srgb, var(--warn) 30%, transparent); }
.wd-restrict .rw-btn { justify-content: center; min-height: 30px; }
.wd-restrict .rw-btn:not(.primary) { border: 1px solid var(--line); background: var(--surface); }
.wd-restrict .rw-btn:not(.primary):hover:not(:disabled) { background: var(--hover); border-color: var(--line-strong); }
.wd-restrict-foot { margin-top: auto; padding: 12px 14px 0; border-top: 1px solid var(--line-soft); display: flex; justify-content: flex-end; }
/* The regions Everyone may edit, as Word marks them: a light yellow wash
   with a bracket at each end of the region. */
.wd-page.wd-perm-show .wd-block.wd-perm { background: color-mix(in srgb, #ffd83d 26%, transparent); }
:root[data-theme='dark'] .wd-page.wd-perm-show .wd-block.wd-perm { background: color-mix(in srgb, #ffd83d 16%, transparent); }
.wd-page.wd-perm-show .wd-block.wd-perm-first::before { content: '['; position: absolute; margin-left: -12px; color: #b8860b; font-weight: 600; }
.wd-page.wd-perm-show .wd-block.wd-perm-last::after { content: ']'; margin-left: 2px; color: #b8860b; font-weight: 600; }
.wd-page.wd-locked { caret-color: var(--ink-3); }
`;
