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
import { Dialog, Button, Icon, Select } from '@rutba/office-ui';
import { SecretInput, PROTECT_CSS } from '../../protect.js';

export const EDIT_KINDS = [
  { id: 'readOnly', label: 'No changes (Read only)' },
  { id: 'trackedChanges', label: 'Tracked changes' },
  { id: 'comments', label: 'Comments' },
  { id: 'forms', label: 'Filling in forms' },
];

/** What Word's pane says about the caret's place, under each kind of protection. */
export function permissionSentence(protection, editable) {
  const lead = 'This document is protected from unintentional editing.';
  switch (protection?.edit) {
    case 'trackedChanges':
      return `${lead} You may edit it, and every change is tracked.`;
    case 'none':
      return `${lead} You may edit it; formatting is limited to styles.`;
    case 'comments':
      return editable ? `${lead} You may edit in this region.` : `${lead} You may only insert comments into this region.`;
    case 'forms':
      return `${lead} You may only fill in forms in this region.`;
    default:
      return editable ? `${lead} You may edit in this region.` : `${lead} You may only view this region.`;
  }
}

/** The sentence the page shows when a keystroke is refused. */
export function refusalSentence(protection) {
  switch (protection?.edit) {
    case 'comments':
      return 'This document is protected: only comments can be added here. Review → Restrict Editing shows the regions you can edit.';
    case 'forms':
      return 'This document is protected for filling in forms: only its form fields can be changed.';
    default:
      return 'This modification is not allowed because the selection is locked.';
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
    <aside className="wd-restrict" aria-label="Restrict Editing">
      <div className="wd-restrict-head">
        <span>Restrict Editing</span>
        <button type="button" className="wd-nav-close" onClick={onClose} data-tip="Close the Restrict Editing pane" aria-label="Close">
          <Icon name="close" size={12} />
        </button>
      </div>
      {enforced ? (
        <div className="wd-restrict-body">
          <section className="wd-restrict-step">
            <h4>Your permissions</h4>
            <p className="wd-restrict-perm">{permissionSentence(protection, inRegion || protection.edit === 'trackedChanges' || protection.edit === 'none')}</p>
            {regions.length && protection.edit !== 'trackedChanges' && protection.edit !== 'forms' && protection.edit !== 'none' ? (
              <>
                <Button label="Find Next Region I Can Edit" className="wd-restrict-next" onClick={onFindNext} />
                <label className="wd-restrict-check">
                  <input type="checkbox" className="wd-restrict-highlight" checked={highlight} onChange={(e) => onHighlight(e.target.checked)} />
                  <span>Highlight the regions I can edit</span>
                </label>
              </>
            ) : null}
            {protection.formatting ? <p className="wd-restrict-note">Formatting is limited to styles: direct formatting is refused.</p> : null}
          </section>
          <div className="wd-restrict-foot">
            <Button label="Stop Protection" className="wd-restrict-stop" onClick={onStop} />
          </div>
        </div>
      ) : (
        <div className="wd-restrict-body">
          <section className="wd-restrict-step">
            <h4><span className="wd-restrict-n">1</span>Formatting restrictions</h4>
            <label className="wd-restrict-check">
              <input type="checkbox" className="wd-restrict-formatting" checked={formatting} onChange={(e) => setFormatting(e.target.checked)} />
              <span>Limit formatting to a selection of styles</span>
            </label>
          </section>
          <section className="wd-restrict-step">
            <h4><span className="wd-restrict-n">2</span>Editing restrictions</h4>
            <label className="wd-restrict-check">
              <input type="checkbox" className="wd-restrict-limit" checked={limit} onChange={(e) => setLimit(e.target.checked)} />
              <span>Allow only this type of editing in the document:</span>
            </label>
            <Select className="wd-restrict-kind" value={edit} disabled={!limit} onChange={(e) => setEdit(e.target.value)}>
              {EDIT_KINDS.map((k) => (
                <option key={k.id} value={k.id}>{k.label}</option>
              ))}
            </Select>
            {exceptions ? (
              <div className="wd-restrict-exceptions">
                <div className="wd-restrict-sub">Exceptions (optional)</div>
                <p className="wd-restrict-note">Select parts of the document and choose users who are allowed to freely edit them.</p>
                <div className="wd-restrict-groups">
                  <div className="wd-restrict-sub2">Groups:</div>
                  <label className="wd-restrict-check wd-restrict-everyone-row">
                    <input type="checkbox" className="wd-restrict-everyone" checked={touched} onChange={(e) => onPermission(e.target.checked)} />
                    <span>Everyone</span>
                  </label>
                </div>
                {regions.length ? <p className="wd-restrict-count">{regions.length === 1 ? 'One region' : `${regions.length} regions`} marked for Everyone.</p> : null}
              </div>
            ) : null}
          </section>
          <section className="wd-restrict-step">
            <h4><span className="wd-restrict-n">3</span>Start enforcement</h4>
            <p className="wd-restrict-note">Are you ready to apply these settings? (You can turn them off later)</p>
            <Button
              primary
              label="Yes, Start Enforcing Protection"
              className="wd-restrict-start"
              disabled={!limit && !formatting}
              title={!limit && !formatting ? 'Yes, Start Enforcing Protection — choose a restriction above first' : 'Yes, Start Enforcing Protection — with an optional password'}
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
  const mismatch = again !== '' && again !== password;
  const ok = password === again;
  const submit = () => ok && onStart(password);
  const enter = (e) => {
    if (e.key === 'Enter') submit();
  };
  return (
    <Dialog
      title="Start Enforcing Protection"
      width={440}
      onClose={onClose}
      actions={
        <>
          <Button label="Cancel" className="wd-enforce-cancel" onClick={onClose} />
          <Button primary label="OK" className="wd-enforce-ok" disabled={!ok} onClick={submit} />
        </>
      }
    >
      <style>{PROTECT_CSS}</style>
      <div className="pw-set">
        <p className="pw-set-lead">Protection method: <b>Password</b>. The document is not encrypted — people can still open it; the password only stops them taking the protection off.</p>
        <label className="pw-label">Enter new password (optional)</label>
        <SecretInput className="wd-enforce-password" label="Enter new password (optional)" value={password} onChange={setPassword} onKeyDown={enter} autoFocus />
        <label className="pw-label">Reenter password to confirm</label>
        <SecretInput className="wd-enforce-again" label="Reenter password to confirm" value={again} onChange={setAgain} onKeyDown={enter} invalid={mismatch} />
        {mismatch ? (
          <div className="pw-error" role="alert">
            <Icon name="info" size={14} />
            <span>The two passwords are not the same.</span>
          </div>
        ) : null}
      </div>
    </Dialog>
  );
}

/** Unprotect Document: the password Stop Protection needs, and Word's refusal when it is wrong. */
export function UnprotectDialog({ error = '', onClose, onSubmit }) {
  const [password, setPassword] = useState('');
  useEffect(() => {
    if (error) setPassword('');
  }, [error]);
  const submit = () => onSubmit(password);
  return (
    <Dialog
      title="Unprotect Document"
      width={400}
      onClose={onClose}
      actions={
        <>
          <Button label="Cancel" className="wd-unprotect-cancel" onClick={onClose} />
          <Button primary label="OK" className="wd-unprotect-ok" onClick={submit} />
        </>
      }
    >
      <style>{PROTECT_CSS}</style>
      <div className="pw-set">
        <label className="pw-label">Password</label>
        <SecretInput className="wd-unprotect-password" label="Password" value={password} onChange={setPassword} invalid={Boolean(error) && !password} onKeyDown={(e) => { if (e.key === 'Enter') submit(); }} autoFocus />
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
.wd-restrict-n { width: 18px; height: 18px; border-radius: 50%; display: grid; place-items: center; font-size: 10.5px; font-weight: 700; background: var(--accent-soft); color: var(--accent); flex: none; }
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
