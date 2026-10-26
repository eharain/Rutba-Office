// An account whose password is not saved on this computer any more (the
// profile's keystore would not open here, D-2026-10-26-09, or it was never
// kept): Mail asks for it as the window opens, once a session for each,
// rather than failing at every fetch. Saved, the account fetches at once.

import React, { useState } from 'react';
import { Button, Dialog, Field, Input, t } from '@rutba/office-ui';

/** Whether an account signs in with a password that is not saved here. */
export const needsPassword = (a) => Boolean(a) && !a.local && !a.archive && a.auth !== 'oauth' && a.imap?.auth !== 'oauth' && !a.hasPassword;

export function PasswordAgain({ shell, accounts, onSaved, toast }) {
  const [skipped, setSkipped] = useState(() => new Set());
  const [value, setValue] = useState('');
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const due = (accounts || []).find((a) => needsPassword(a) && !skipped.has(a.id));
  if (!due) return null;

  const skip = () => {
    setSkipped(new Set([...skipped, due.id]));
    setValue('');
    setError('');
  };
  const save = async () => {
    if (!value) return;
    setBusy(true);
    setError('');
    try {
      await shell.mail.updateAccount({ id: due.id, patch: {}, password: value });
      setValue('');
      await onSaved?.(due);
      toast?.(t('Saved. {address} is fetching its mail.', { address: due.email }), { tone: 'good' });
    } catch (err) {
      setError(String(err?.message || err).replace(/^Error invoking remote method '[^']+': (Error: )?/, ''));
    } finally {
      setBusy(false);
    }
  };

  return (
    <Dialog
      title={t('Password for {address}', { address: due.email })}
      width={460}
      className="ml-password-again"
      onClose={skip}
      actions={(
        <>
          <Button label={t('Not now')} className="ml-password-later" onClick={skip} />
          <Button primary label={busy ? t('Saving…') : t('Save')} className="ml-password-save" disabled={!value || busy} onClick={save} />
        </>
      )}
    >
      <p className="rw-hint" style={{ marginTop: 0 }}>
        {t('The password for this account is not saved on this computer, so Mail cannot fetch or send its mail. Type it again to keep it in this profile.')}
      </p>
      <Field label={t('Password')}>
        <Input className="ml-password-input" type="password" value={value} autoFocus onChange={(e) => setValue(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') save(); }} />
      </Field>
      {error ? <p className="ml-password-error" style={{ color: 'var(--bad, #c62828)', fontSize: 13 }}>{error}</p> : null}
    </Dialog>
  );
}
