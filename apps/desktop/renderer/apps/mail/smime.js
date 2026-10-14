// Signed and encrypted mail in the Mail window: the strip over a message that
// says who signed it and whether it can be trusted, or that it is encrypted
// and how it was opened; the Certificates dialog where a person imports their
// own certificate (a .p12 or .pfx) and other people's; and the Sign and
// Encrypt choices of a new message.
//
// The engine is @rutba/mailbox/smime behind the main process's mail-smime.js;
// this is only what the window shows of it. Nothing is asked of any server:
// revocation is not checked, and the strip says so.

import React, { useCallback, useEffect, useState } from 'react';
import { Button, Dialog, Field, Input, Icon, formatWhen, t } from '@rutba/office-ui';

const day = (iso) => (iso ? new Date(iso).toLocaleDateString(undefined, { year: 'numeric', month: 'long', day: 'numeric' }) : '');
const who = (s) => (s ? (s.email ? `${s.name} <${s.email}>` : s.name) : t('an unknown signer'));

/** The sentence for a signature, by what was found. */
function signedSentence(signed, from) {
  const name = who(signed.signer);
  switch (signed.status) {
    case 'valid':
      return t('Signed by {name}. The signature is valid and the message has not changed since it was signed.', { name });
    case 'other-sender':
      return t('Signed by {name}, who is not the sender: the certificate is for {email}, and the message says it is from {from}.', { name, email: signed.signer?.email || signed.signer?.name || '', from });
    case 'no-from':
      return t('Signed by {name}, but the message names no sender to compare the certificate with.', { name });
    case 'changed':
      return t('Signed by {name}, but the message has changed since it was signed. Do not rely on what it says.', { name });
    case 'bad-signature':
      return t('The signature does not belong to the certificate it names, so who signed this cannot be known.');
    case 'untrusted':
      return t('Signed by {name}, and the message has not changed, but no authority this computer trusts vouches for the certificate.', { name });
    case 'expired':
      return t('Signed by {name}, and the message has not changed, but the certificate is out of date.', { name });
    case 'not-for-mail':
      return t('Signed by {name} with a certificate that is not meant for mail.', { name });
    case 'no-certificate':
      return t('This message is signed, but the signer\'s certificate is neither in it nor on this computer, so the signature cannot be checked.');
    case 'damaged':
      return t('The signature is damaged and cannot be checked.');
    default:
      return t('This message is signed in a way this cannot check yet.');
  }
}

/** The sentence for an encrypted message. */
function encryptedSentence(encrypted) {
  switch (encrypted.status) {
    case 'decrypted':
      return t('Encrypted ({cipher}), and opened with your certificate for {email}.', { cipher: encrypted.cipher, email: encrypted.recipient?.email || encrypted.recipient?.name || '' });
    case 'no-identity':
      return t('This message is encrypted, and no certificate on this computer opens it. Import yours (a .p12 or .pfx file) to read it.');
    case 'damaged':
      return t('This message is encrypted, but it is damaged and cannot be opened.');
    default:
      return t('This message is encrypted in a way this cannot open yet.');
  }
}

const GOOD = new Set(['valid']);

/**
 * Is a decrypted message one whose words must not be drawn as a web page:
 * encrypted, but not signed by someone trusted, so it could have been changed
 * on the way to send its own contents elsewhere when opened (EFAIL). Such a
 * message is shown as plain text, with no links and nothing fetched.
 */
export const unsafeDecrypted = (smime) => smime?.encrypted?.status === 'decrypted' && smime?.signed?.status !== 'valid';

/**
 * The strip over a signed or encrypted message.
 * @param {{ smime: { signed, encrypted, error?, certificateChanged? }, from?: string, onImport?: () => void, onCertificates?: () => void }} props
 */
export function SmimeStrip({ smime, from = '', onImport, onCertificates }) {
  const [open, setOpen] = useState(false);
  if (!smime) return null;
  if (smime.error) {
    return (
      <div className="ml-smime caution" data-signed="" data-encrypted="" data-error={smime.error}>
        <div className="ml-smime-row">
          <Icon name="info" size={15} />
          <span className="ml-smime-text">
            {smime.error === 'too-large' ? t('This message is too large to check its signature or encryption.') : t('This message\'s signature or encryption could not be read.')}
          </span>
        </div>
      </div>
    );
  }
  if (!smime.signed && !smime.encrypted) return null;
  const { signed, encrypted } = smime;
  const cautions = [...(signed?.cautions || []), ...(encrypted?.cautions || [])];
  const bad = (signed && !GOOD.has(signed.status)) || (encrypted && encrypted.status !== 'decrypted') || cautions.length > 0 || Boolean(smime.certificateChanged);
  const warn = signed && ['changed', 'bad-signature', 'damaged'].includes(signed.status);
  return (
    <div className={`ml-smime${warn ? ' bad' : bad ? ' caution' : ' good'}`} data-signed={signed?.status || ''} data-encrypted={encrypted?.status || ''} data-cautions={cautions.join(' ')}>
      {encrypted ? (
        <div className="ml-smime-row ml-smime-encrypted">
          <Icon name="lock" size={15} />
          <span className="ml-smime-text">
            {encryptedSentence(encrypted)}
            {encrypted.cautions?.includes('weak-cipher') ? ` ${t('That cipher is too weak to keep a message private now.')}` : ''}
            {unsafeDecrypted(smime) ? ` ${t('It is shown as plain text, with its links and pictures left out: an encrypted message that is not signed by someone trusted could have been changed on the way to send its words elsewhere.')}` : ''}
          </span>
          {encrypted.status === 'no-identity' && onImport ? (
            <button type="button" className="ml-smime-import" onClick={onImport}>
              {t('Import certificate…')}
            </button>
          ) : null}
        </div>
      ) : null}
      {signed ? (
        <div className="ml-smime-row ml-smime-signed">
          <Icon name={warn ? 'info' : 'shield'} size={15} />
          <span className="ml-smime-text">
            {signedSentence(signed, from)}
            {signed.cautions?.includes('sha1') ? ` ${t('It was signed with SHA-1, which no longer protects a signature from forgery.')}` : ''}
            {smime.certificateChanged ? ` ${t('The certificate is not the one kept for {email} before, so it was not kept: check with them, then choose in Certificates.', { email: smime.certificateChanged.email })}` : ''}
          </span>
          {signed.signer ? (
            <button type="button" className="ml-smime-more" aria-expanded={open} onClick={() => setOpen((o) => !o)}>
              {open ? t('Fewer details') : t('Details')}
            </button>
          ) : null}
        </div>
      ) : null}
      {signed?.signer && open ? (
        <dl className="ml-smime-details">
          <dt>{t('Signed by')}</dt>
          <dd>{who(signed.signer)}</dd>
          {signed.signedAt ? (
            <>
              <dt>{t('Signed on')}</dt>
              <dd>{formatWhen(signed.signedAt)}</dd>
            </>
          ) : null}
          <dt>{t('Issued by')}</dt>
          <dd>{signed.signer.issuer}</dd>
          <dt>{t('Valid')}</dt>
          <dd>{t('{from} to {to}', { from: day(signed.signer.validFrom), to: day(signed.signer.validTo) })}</dd>
          {signed.trustedBy ? (
            <>
              <dt>{t('Trusted through')}</dt>
              <dd>{signed.trustedBy}</dd>
            </>
          ) : null}
          <dt>{t('Thumbprint')}</dt>
          <dd dir="ltr" className="ml-smime-thumb">{signed.signer.thumbprint}</dd>
        </dl>
      ) : null}
      {signed ? (
        <div className="ml-smime-note">
          {t('Revocation is not checked: that would mean asking the certificate\'s authority over the network, which this never does.')}
          {onCertificates ? (
            <button type="button" className="ml-smime-certs" onClick={onCertificates}>
              {t('Certificates…')}
            </button>
          ) : null}
        </div>
      ) : null}
    </div>
  );
}

/**
 * Certificates for signed and encrypted mail: one's own (with their keys, from
 * a .p12 or .pfx) and other people's (kept from their signed mail, or
 * imported from a .cer).
 */
export function CertificatesDialog({ shell, onClose, toast, start = null }) {
  const [identities, setIdentities] = useState([]);
  const [others, setOthers] = useState({});
  const [held, setHeld] = useState({});
  const [path, setPath] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);

  const load = useCallback(async () => {
    setIdentities(await shell.mail.smimeIdentities());
    setOthers(await shell.mail.smimeCertificates());
    setHeld(await shell.mail.pendingCertificates());
  }, [shell]);
  useEffect(() => {
    load().catch(() => {});
  }, [load]);
  useEffect(() => {
    if (start === 'import') document.querySelector('.ml-cert-path input')?.focus();
  }, [start]);

  const browse = async () => {
    const paths = await shell.dialog.open({ title: t('Import your certificate'), filters: [{ name: t('Certificates with their keys'), extensions: ['p12', 'pfx'] }, { name: t('All files'), extensions: ['*'] }] });
    if (paths?.[0]) setPath(paths[0]);
  };
  const importOwn = async () => {
    setBusy(true);
    setError('');
    try {
      const got = await shell.mail.importIdentity({ path, password });
      setPassword('');
      setPath('');
      toast?.(t('Imported the certificate for {email}.', { email: got.email || got.name }));
      await load();
    } catch (err) {
      setError(err.message);
    } finally {
      setBusy(false);
    }
  };
  const importOther = async () => {
    const paths = await shell.dialog.open({ title: t('Import a certificate'), filters: [{ name: t('Certificates'), extensions: ['cer', 'crt', 'pem', 'der'] }, { name: t('All files'), extensions: ['*'] }] });
    if (!paths?.[0]) return;
    try {
      await shell.mail.importCertificate({ path: paths[0] });
      await load();
    } catch (err) {
      setError(err.message);
    }
  };
  const remove = async (thumbprint) => {
    await shell.mail.removeIdentity({ thumbprint });
    await load();
  };

  const people = Object.entries(others).filter(([email]) => !identities.some((i) => i.emails?.includes(email)));
  return (
    <Dialog title={t('Certificates for signed and encrypted mail')} width={620} onClose={onClose} actions={<Button primary label={t('Close')} onClick={onClose} />}>
      <div className="ml-certs">
        <div className="ml-hold-section">{t('Your certificates')}</div>
        {identities.length ? (
          <ul className="ml-cert-list">
            {identities.map((c) => (
              <li key={c.thumbprint} className="ml-cert" data-thumbprint={c.thumbprint}>
                <Icon name="shield" size={16} />
                <span className="ml-cert-who">
                  <strong>{c.name}</strong>
                  <span dir="ltr">{c.email}</span>
                </span>
                <span className="ml-cert-dates">{t('Issued by {issuer}, valid until {date}', { issuer: c.issuer, date: day(c.validTo) })}</span>
                <button type="button" className="ml-cert-remove" onClick={() => remove(c.thumbprint)}>
                  {t('Remove')}
                </button>
              </li>
            ))}
          </ul>
        ) : (
          <p className="ml-cert-empty">{t('None yet. Import yours to sign your mail and to read mail encrypted for you.')}</p>
        )}
        <div className="ml-cert-import">
          <Field label={t('Certificate file (.p12 or .pfx)')}>
            <div className="ml-cert-path">
              <Input value={path} onChange={(e) => setPath(e.target.value)} placeholder={t('Choose the file Windows or your mail program exported')} />
              <Button icon="folder" label={t('Browse…')} onClick={browse} />
            </div>
          </Field>
          <Field label={t('Its password')}>
            <Input className="rw-input ml-cert-password" type="password" value={password} onChange={(e) => setPassword(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && path && importOwn()} />
          </Field>
          <div className="ml-cert-actions">
            <Button primary className="ml-cert-ok" label={busy ? t('Importing…') : t('Import')} disabled={!path || busy} onClick={importOwn} />
            {error ? <span className="ml-cert-error" role="alert">{error}</span> : null}
          </div>
        </div>
        {Object.keys(held).length ? (
          <>
            <div className="ml-hold-section">{t('Certificates that changed')}</div>
            <ul className="ml-cert-list">
              {Object.entries(held).map(([email, c]) => (
                <li key={email} className="ml-cert changed" data-email={email}>
                  <Icon name="info" size={16} />
                  <span className="ml-cert-who">
                    <strong>{c.name}</strong>
                    <span dir="ltr">{email}</span>
                  </span>
                  <span className="ml-cert-dates">{t('A different certificate arrived for this address. Use it only if they told you it changed. Issued by {issuer}, thumbprint {thumbprint}.', { issuer: c.issuer, thumbprint: c.thumbprint })}</span>
                  <span className="ml-cert-choice">
                    <button type="button" className="ml-cert-accept" onClick={async () => { await shell.mail.acceptCertificate({ email }); await load(); }}>
                      {t('Use the new one')}
                    </button>
                    <button type="button" className="ml-cert-dismiss" onClick={async () => { await shell.mail.dismissCertificate({ email }); await load(); }}>
                      {t('Keep the old one')}
                    </button>
                  </span>
                </li>
              ))}
            </ul>
          </>
        ) : null}
        <div className="ml-hold-section">{t('Other people\'s certificates')}</div>
        {people.length ? (
          <ul className="ml-cert-list">
            {people.map(([email, list]) => (
              <li key={email} className="ml-cert other" data-email={email}>
                <Icon name="lock" size={16} />
                <span className="ml-cert-who">
                  <strong>{list[0]?.name}</strong>
                  <span dir="ltr">{email}</span>
                </span>
                <span className="ml-cert-dates">{list[0]?.how === 'signed' ? t('Kept from their signed mail, valid until {date}', { date: day(list[0]?.validTo) }) : t('Imported, valid until {date}', { date: day(list[0]?.validTo) })}</span>
              </li>
            ))}
          </ul>
        ) : (
          <p className="ml-cert-empty">{t('None yet. A certificate is kept when someone sends you signed mail, so you can encrypt to them.')}</p>
        )}
        <Button className="ml-cert-other" icon="import" label={t('Import a certificate…')} onClick={importOther} />
        <p className="ml-cert-note">{t('Your keys are kept in the protected store of this computer\'s account, never in a file. Certificates are checked against the authorities this computer trusts, offline; revocation is not checked.')}</p>
      </div>
    </Dialog>
  );
}

/** Sign and Encrypt for a new message. */
export function SecurityRow({ value = {}, identity, onChange, onCertificates }) {
  const none = !identity;
  return (
    <div className="ml-security">
      <button type="button" className={`ml-security-sign${value.sign ? ' on' : ''}`} aria-pressed={Boolean(value.sign)} disabled={none} data-tip={none ? t('Sign: import your certificate first') : t('Sign: the message carries your signature, so its recipients know it is from you and unchanged')} onClick={() => onChange({ ...value, sign: !value.sign })}>
        <Icon name="shield" size={14} />
        {t('Sign')}
      </button>
      <button type="button" className={`ml-security-encrypt${value.encrypt ? ' on' : ''}`} aria-pressed={Boolean(value.encrypt)} disabled={none} data-tip={none ? t('Encrypt: import your certificate first') : t('Encrypt: only the recipients, and you, can read it')} onClick={() => onChange({ ...value, encrypt: !value.encrypt })}>
        <Icon name="lock" size={14} />
        {t('Encrypt')}
      </button>
      {none ? (
        <button type="button" className="ml-security-certs" onClick={onCertificates}>
          {t('Import your certificate to sign or encrypt…')}
        </button>
      ) : (
        <span className="ml-security-who" dir="ltr">
          {identity.email}
        </span>
      )}
    </div>
  );
}

let installed = false;

/** The styles of the strip, the dialog and the compose row, once. */
export function installSmimeStyles() {
  if (installed || typeof document === 'undefined') return;
  installed = true;
  const style = document.createElement('style');
  style.id = 'rutba-mail-smime-css';
  style.textContent = SMIME_CSS;
  document.head.appendChild(style);
}

const SMIME_CSS = `
.ml-smime { margin: 10px 22px 2px; border: 1px solid var(--line); border-radius: var(--r-2); background: var(--surface); font-size: 12px; overflow: hidden; }
.ml-smime.good { border-color: color-mix(in srgb, var(--good, #2e7d32) 45%, var(--line)); }
.ml-smime.caution { border-color: color-mix(in srgb, var(--warn, #b26a00) 50%, var(--line)); }
.ml-smime.bad { border-color: color-mix(in srgb, var(--bad, #c62828) 60%, var(--line)); background: color-mix(in srgb, var(--bad, #c62828) 6%, var(--surface)); }
.ml-smime-row { display: flex; align-items: flex-start; gap: 8px; padding: 7px 10px; line-height: 1.45; }
.ml-smime-row + .ml-smime-row { border-top: 1px solid var(--line-soft); }
.ml-smime-row > svg { flex: none; margin-top: 1px; }
.ml-smime.good .ml-smime-signed > svg { color: var(--good, #2e7d32); }
.ml-smime.caution .ml-smime-row > svg { color: var(--warn, #b26a00); }
.ml-smime.bad .ml-smime-row > svg { color: var(--bad, #c62828); }
.ml-smime-encrypted > svg { color: var(--accent); }
.ml-smime-text { flex: 1; min-width: 0; color: var(--ink); }
.ml-smime-row button, .ml-smime-note button { flex: none; border: 0; background: none; color: var(--accent); font: inherit; font-weight: 600; cursor: pointer; padding: 0; white-space: nowrap; }
.ml-smime-row button:hover, .ml-smime-note button:hover { text-decoration: underline; }
.ml-smime-details { display: grid; grid-template-columns: max-content 1fr; gap: 3px 14px; margin: 0; padding-block: 6px 8px; padding-inline: 33px 10px; border-top: 1px solid var(--line-soft); }
.ml-smime-details dt { color: var(--ink-3); }
.ml-smime-details dd { margin: 0; color: var(--ink); overflow-wrap: anywhere; }
.ml-smime-thumb { font-family: var(--mono); font-size: 11px; }
.ml-smime-note { display: flex; gap: 12px; align-items: baseline; padding-block: 5px 7px; padding-inline: 33px 10px; border-top: 1px solid var(--line-soft); color: var(--ink-3); font-size: 11px; }
.ml-smime-note button { margin-inline-start: auto; font-size: 11px; }

.ml-certs { display: flex; flex-direction: column; gap: 8px; }
.ml-cert-list { list-style: none; margin: 0; padding: 0; border: 1px solid var(--line); border-radius: var(--r-2); }
.ml-cert { display: grid; grid-template-columns: 18px 1fr auto; gap: 2px 10px; align-items: center; padding: 7px 10px; }
.ml-cert + .ml-cert { border-top: 1px solid var(--line-soft); }
.ml-cert > svg { grid-row: 1 / 3; color: var(--accent); }
.ml-cert-who { display: flex; gap: 8px; align-items: baseline; min-width: 0; }
.ml-cert-who span { color: var(--ink-3); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.ml-cert-dates { grid-column: 2; color: var(--ink-3); font-size: 11.5px; }
.ml-cert-choice { grid-column: 3; grid-row: 1 / 3; display: flex; flex-direction: column; gap: 4px; align-items: flex-end; }
.ml-cert-choice button { border: 0; background: none; color: var(--accent); font: inherit; font-weight: 600; cursor: pointer; padding: 0; white-space: nowrap; }
.ml-cert.changed > svg { color: var(--warn, #b26a00); }
.ml-cert-remove { grid-column: 3; grid-row: 1 / 3; border: 0; background: none; color: var(--accent); font: inherit; font-weight: 600; cursor: pointer; }
.ml-cert-empty, .ml-cert-note { margin: 0; color: var(--ink-3); font-size: 12px; }
.ml-cert-import { display: flex; flex-direction: column; gap: 6px; padding: 10px; border: 1px dashed var(--line); border-radius: var(--r-2); }
.ml-cert-path { display: flex; gap: 6px; }
.ml-cert-path > :first-child { flex: 1; }
.ml-cert-actions { display: flex; gap: 10px; align-items: center; }
.ml-cert-error { color: var(--bad, #c62828); font-size: 12px; }
.ml-cert-other { align-self: flex-start; }

.ml-security { display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }
.ml-security button { display: inline-flex; align-items: center; gap: 5px; border: 1px solid var(--line); border-radius: 20px; background: var(--surface); color: var(--ink-2); font: inherit; font-size: 12px; padding: 2px 10px; cursor: pointer; }
.ml-security button.on { background: var(--selected); color: var(--ink); border-color: var(--accent); font-weight: 600; }
.ml-security button:disabled { opacity: .55; cursor: default; }
.ml-security .ml-security-certs { border: 0; background: none; color: var(--accent); font-weight: 600; padding: 2px 4px; }
.ml-security-who { color: var(--ink-3); font-size: 11.5px; }
`;
