// What the AI tools may do with this database (D-2026-10-26-04, -06): Off,
// Read, or Read and prepare (new and changed records, never the design), and,
// for a protected database, the role an Admin chooses for them, never one
// that may change the design. Kept on this computer, not in the file, and
// the same setting the AI Tools window lists. A change the tools ask to
// write to a table linked to a database engine needs Read, prepare and act
// and is shown to the person first.

import React, { useEffect, useRef, useState } from 'react';
import { Button, Dialog, Field, Icon, Select, useToast, t } from '@rutba/office-ui';
import { STEP_NAMES } from '../../ai-tools.js';

/** The side's button: shown when the AI tools are among the person's apps and they may set this. */
export function AiToolsButton({ shell, db, onOpen }) {
  const [s, setS] = useState(null);
  useEffect(() => {
    if (!shell?.ai?.stepOf || !db?.file) return undefined;
    let live = true;
    const read = () => shell.ai.stepOf({ kind: 'database', id: db.file }).then((x) => { if (live) setS(x); }).catch(() => {});
    read();
    const off = shell.on?.('ai:changed', (c) => { if (c?.kind === 'database') read(); });
    return () => { live = false; off?.(); };
  }, [shell, db?.file]);
  if (!s?.addon) return null;
  return (
    <button type="button" className="bs-table bs-ai-tools" data-ai-step={s.step} onClick={onOpen}>
      <Icon name="wand" size={14} />
      <span>{t('base.ai-tools.ai-tools-step', { step: STEP_NAMES()[s.step] || s.step })}</span>
    </button>
  );
}

export function AiToolsDialog({ shell, db, onClose }) {
  const toast = useToast();
  const [s, setS] = useState(null);
  const [step, setStep] = useState('off');
  const [role, setRole] = useState('');
  const [roles, setRoles] = useState([]);
  const guarded = Boolean(db.access?.protected);
  // What the user chose before the current setting was read stays chosen (11 Nov gate: read late, it put Off back over
  // Read and prepare, and Save, enabled by that read, saved Off).
  const chosen = useRef({ step: false, role: false });

  useEffect(() => {
    shell.ai.stepOf({ kind: 'database', id: db.file }).then((x) => {
      setS(x);
      if (!chosen.current.step) setStep(x.step);
      if (!chosen.current.role) setRole(x.role || '');
    }).catch((err) => toast(err.message, { tone: 'bad' }));
    // Only the roles that may not change the design: the tools never work as one that may.
    if (guarded) shell.base.roles({ id: db.id }).then((r) => setRoles((r || []).filter((x) => !x.builtIn && !x.rights?.design))).catch(() => setRoles([]));
  }, [shell, db.id, db.file, guarded, toast]);

  const save = async () => {
    try {
      await shell.ai.setStep({ kind: 'database', id: db.file, step, role: guarded ? role || null : null });
      onClose();
    } catch (err) {
      toast(err.message, { tone: 'bad', ms: 6000 });
    }
  };
  const names = STEP_NAMES();
  const notes = {
    off: t('base.ai-tools.ai-tools-cannot'),
    read: t('base.ai-tools.may-read-records'),
    prepare: t('base.ai-tools.may-also-add'),
    act: t('base.ai-tools.may-also-write'),
  };

  return (
    <Dialog
      title={t('base.ai-tools.ai-tools-name', { name: db.name })}
      width={480}
      className="bs-ai-dialog"
      onClose={onClose}
      actions={(
        <>
          <Button label={t('common.cancel')} onClick={onClose} />
          <Button primary label={t('common.save')} className="bs-ai-save" disabled={!s || (guarded && step !== 'off' && !role)} onClick={save} />
        </>
      )}
    >
      <fieldset className="ai-modes" aria-label={t('common.ai-tools-2')}>
        {(s?.steps || ['off', 'read', 'prepare', 'act']).map((k) => (
          <label key={k} className={`ai-mode${step === k ? ' on' : ''}`}>
            <input type="radio" name="bs-ai-step" value={k} checked={step === k} onChange={() => { chosen.current.step = true; setStep(k); }} />
            <span className="ai-mode-label">{names[k]}</span>
            <span className="ai-mode-note">{notes[k]}</span>
          </label>
        ))}
      </fieldset>
      {guarded ? (
        <Field label={t('base.ai-tools.work-as')} hint={t('base.ai-tools.protected-database')}>
          {roles.length ? (
            <Select className="bs-ai-role" value={role} onChange={(e) => { chosen.current.role = true; setRole(e.target.value); }}>
              <option value="">{t('base.ai-tools.choose-a-role')}</option>
              {roles.map((r) => <option key={r.name} value={r.name}>{r.name}</option>)}
            </Select>
          ) : <p className="ai-note">{t('base.ai-tools.make-role-first')}</p>}
        </Field>
      ) : null}
      <p className="ai-note">{t('base.ai-tools.kept-computer-not')}</p>
    </Dialog>
  );
}
