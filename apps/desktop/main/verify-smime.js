// Signed and encrypted mail (S/MIME) in the Mail window.
//
// Certificates and keys are made for the check when it runs: a test
// authority (trusted for the run only, as a person would add one to
// Windows), Amina's certificate, the account owner's as a .p12 with a
// password, and a stranger's from an authority nobody trusts. Into an account
// of the check's own arrive: a message Amina signed, the same message with a
// word changed after signing, one the stranger signed, and one Amina signed
// and encrypted for the owner; and the hostile ones: Amina's signature on a
// message from Bo, a second certificate for Amina's address from someone
// else, and an encrypted message nobody signed whose HTML has a link and a
// remote picture. In the window each says what it is: signed by Amina, valid
// and unchanged; changed since it was signed; signed by someone no authority
// here vouches for; signed by someone who is not the sender (not shown as
// good); a certificate different from the one kept (held aside, not
// swapped in, and the Certificates dialog asks); encrypted, with no
// certificate here to open it. The owner's .p12 is imported from that strip
// (a wrong password refused in a sentence first), the encrypted message then
// opens, its signature checked inside, and the unsigned one opens as plain
// text with nothing fetched. A new message to Amina, signed and encrypted,
// is refused while it has a Bcc, then goes out: its Sent copy opens with the
// owner's own certificate, and Amina's key opens what was sent. Then the strip
// in Urdu, right to left. Everything the check made is taken away at the end.
// Run alone with RUTBA_VERIFY_ONLY=smime.

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import { protect, readSmime, certificate } from '@rutba/mailbox/smime';
import { makeIdentities, makePkcs12, makeCertificate } from '../../../tests/fixtures/smime-keys.js';

const PASSWORD = 'Harbour 2026';

const message = ({ from, to, subject, body }) =>
  Buffer.from(
    `From: ${from.name} <${from.email}>\r\nTo: ${to.name} <${to.email}>\r\nSubject: ${subject}\r\nDate: ${new Date().toUTCString()}\r\nMessage-ID: <${crypto.randomUUID()}@check.example>\r\nMIME-Version: 1.0\r\nContent-Type: text/plain; charset=utf-8\r\nContent-Transfer-Encoding: 7bit\r\n\r\n${body}\r\n`,
    'latin1'
  );

/** Would this document fetch anything: an http(s) address in a src, href or url(). */
const fetchesRemote = (html) => /(?:src|href)\s*=\s*["']?https?:|url\(\s*["']?https?:/i.test(html);

/** Files under `dir` named `name`, newest first. */
function findFiles(dir, name, out = []) {
  let entries = [];
  try {
    entries = fs.readdirSync(dir, { withFileTypes: true });
  } catch {
    return out;
  }
  for (const e of entries) {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) findFiles(p, name, out);
    else if (e.name === name) out.push(p);
  }
  return out.sort((a, b) => fs.statSync(b).mtimeMs - fs.statSync(a).mtimeMs);
}

/**
 * @param {object} h the harness: open, check, until, wait, capture, errorsIn
 */
export async function verifySmime(h) {
  const { open, check, until, wait, capture, errorsIn } = h;
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-verify-smime-'));
  let win = null;
  let js = null;
  let accountId = null;
  let language = null;
  let undo = null;
  let thumbprint = null;
  try {
    const people = makeIdentities({ org: 'Rutba Check Authority' });
    const p12 = path.join(root, 'Archive Owner.p12');
    fs.writeFileSync(p12, makePkcs12(people.owner, PASSWORD, { chain: [people.ca.cert] }));

    win = await open('mail');
    js = (code) => win.webContents.executeJavaScript(code);
    const mail = (method, args) => js(`window.rutbaOffice.mail.${method}(${JSON.stringify(args ?? {})})`);
    const shot = async (w, name) => {
      w.webContents.invalidate();
      await wait(700);
      await capture(w, name);
    };
    const waitFor = async (fn, what, ms) => {
      let got = null;
      await until(async () => {
        got = await fn();
        return Boolean(got);
      }, what, ms).catch(() => {});
      return got || null;
    };
    language = await js(`window.rutbaOffice.store.get({ key: 'language', fallback: null })`);
    undo = await js(`window.rutbaOffice.store.get({ key: 'mail.undoSeconds', fallback: null })`);
    await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the sidebar', 10000);

    // The account, the authority trusted for the run, and the mail that arrives.
    const made = await mail('addAccount', { account: { email: people.owner.email, name: 'Archive Owner', local: true } });
    accountId = made.id;
    await mail('trustForTest', { pem: people.ca.cert });
    const signed = protect(message({ from: people.amina, to: people.owner, subject: 'Harbour lease: signed copy', body: 'The quartzberry clause stands as agreed.' }), { sign: { key: people.amina.key, cert: people.amina.cert, chain: [people.ca.cert] } });
    const altered = Buffer.from(protect(message({ from: people.amina, to: people.owner, subject: 'Harbour lease: altered copy', body: 'The rent is 900 a month.' }), { sign: { key: people.amina.key, cert: people.amina.cert } }).toString('latin1').replace('900 a month', '990 a month'), 'latin1');
    const stranger = protect(message({ from: people.stranger, to: people.owner, subject: 'An offer from an unknown signer', body: 'Sign here.' }), { sign: { key: people.stranger.key, cert: people.stranger.cert } });
    const sealed = protect(message({ from: people.amina, to: people.owner, subject: 'Harbour lease: the sealed terms', body: 'The sealed quartzberry terms: twelve months.' }), { sign: { key: people.amina.key, cert: people.amina.cert, chain: [people.ca.cert] }, encryptFor: [people.owner.cert, people.amina.cert] });
    const fromBo = protect(message({ from: people.bo, to: people.owner, subject: 'Invoice from Bo, signed by Amina', body: 'Pay the new account.' }), { sign: { key: people.amina.key, cert: people.amina.cert, chain: [people.ca.cert] } });
    // Someone else with a certificate for Amina's address (the authority fooled, or a key not hers).
    const otherKeys = crypto.generateKeyPairSync('rsa', { modulusLength: 2048 });
    const otherCert = makeCertificate({ subject: { cn: 'Amina Rahman', email: people.amina.email }, issuer: { cn: 'Rutba Check Authority Root', o: 'Rutba Check Authority' }, publicKey: otherKeys.publicKey, signingKey: people.ca.key, email: people.amina.email });
    const secondKey = protect(message({ from: people.amina, to: people.owner, subject: 'Harbour lease: from a second key', body: 'Use my new certificate from now on.' }), { sign: { key: otherKeys.privateKey, cert: otherCert, chain: [people.ca.cert] } });
    // Encrypted for the owner, signed by nobody, its HTML with a link and a remote picture (EFAIL's shape).
    const unsignedHtml = Buffer.from(`From: ${people.amina.name} <${people.amina.email}>\r\nTo: ${people.owner.name} <${people.owner.email}>\r\nSubject: Harbour lease: unsigned and sealed\r\nMIME-Version: 1.0\r\nContent-Type: text/html; charset=utf-8\r\n\r\n<p>The unsigned quartzberry note.</p><p><a href="https://collector.example/?q=secret">Open</a><img src="https://collector.example/pixel.png"></p>\r\n`, 'latin1');
    const unsignedSealed = protect(unsignedHtml, { encryptFor: [people.owner.cert] });
    for (const raw of [signed, altered, stranger, sealed, fromBo, secondKey, unsignedSealed]) await mail('deliverTest', { accountId, folder: 'Inbox', raw: raw.toString('latin1') });

    await js(`(() => { const b = [...document.querySelectorAll('.ml-accounts button')].find((n) => n.textContent.trim().startsWith('Archive Owner')); b?.click(); return Boolean(b); })()`);
    const openSubject = (subject) => js(`(() => { const r = [...document.querySelectorAll('.ml-row .ml-subject')].find((n) => n.textContent.trim() === ${JSON.stringify(subject)}); r?.click(); return Boolean(r); })()`);
    const strip = (want) =>
      waitFor(async () => {
        const s = await js(`(() => { const el = document.querySelector('.ml-smime'); return el ? { signed: el.dataset.signed, encrypted: el.dataset.encrypted, tone: el.classList.contains('good') ? 'good' : el.classList.contains('bad') ? 'bad' : 'caution', text: [...el.querySelectorAll('.ml-smime-text')].map((n) => n.textContent).join(' | '), note: el.querySelector('.ml-smime-note')?.textContent || '', head: document.querySelector('.ml-head h2')?.textContent || '', body: document.querySelector('.ml-body iframe')?.srcdoc || '' } : null; })()`);
        return s && want(s) ? s : false;
      }, 'the S/MIME strip', 10000);
    await until(() => js(`document.querySelectorAll('.ml-row').length >= 7`), 'the seven messages', 10000).catch(() => {});

    // Signed, valid, unchanged.
    await openSubject('Harbour lease: signed copy');
    const valid = await strip((s) => s.head.includes('signed copy') && s.signed === 'valid');
    check('smime: a signed message says who signed it, that the signature is valid and the message unchanged, and that revocation is not checked',
      Boolean(valid) && /Signed by Amina Rahman <amina@northwind\.example>\. The signature is valid and the message has not changed since it was signed\./.test(valid.text) && /Revocation is not checked/.test(valid.note) && /quartzberry clause/.test(valid.body),
      JSON.stringify(valid && { text: valid.text, note: valid.note }));
    await js(`document.querySelector('.ml-smime-more')?.click(), 1`);
    const details = await waitFor(() => js(`(() => { const d = document.querySelector('.ml-smime-details'); return d ? d.textContent : null; })()`), 'the certificate details', 4000);
    check('smime: its details name the issuer, the dates, the day it was signed and the authority it is trusted through', Boolean(details) && /Rutba Check Authority Root/.test(details) && /Trusted through/.test(details) && details.includes(String(new Date().getFullYear())) && !/\b2[1-9]\d\d\b/.test(details.replace(/Thumbprint.*/, '')), details);
    await shot(win, 'smime-1-signed-valid.png');

    // Amina's good signature on a message from Bo: not shown as good, and her certificate not kept for Bo.
    await openSubject('Invoice from Bo, signed by Amina');
    const otherSender = await strip((s) => s.head.includes('Invoice from Bo'));
    const boKept = (await mail('smimeCertificates'))['bo@contoso.example'];
    check('smime: a good signature by someone other than the sender says so, is not shown as good, and keeps no certificate for the sender',
      otherSender?.signed === 'other-sender' && otherSender.tone !== 'good' && /who is not the sender/.test(otherSender.text) && /amina@northwind\.example/.test(otherSender.text) && !boKept,
      JSON.stringify(otherSender && { signed: otherSender.signed, tone: otherSender.tone, text: otherSender.text, boKept: Boolean(boKept) }));
    await shot(win, 'smime-8-other-sender.png');

    // A second certificate for Amina's address: held aside, the one kept stays, the strip says so.
    const keptBefore = (await mail('smimeCertificates'))['amina@northwind.example']?.[0]?.thumbprint || null;
    await openSubject('Harbour lease: from a second key');
    const second = await strip((s) => s.head.includes('second key'));
    const keptAfter = (await mail('smimeCertificates'))['amina@northwind.example']?.[0]?.thumbprint || null;
    const held = (await mail('pendingCertificates'))['amina@northwind.example'];
    check('smime: a different certificate for an address already kept is held aside, never swapped in, and the strip says so',
      Boolean(second) && Boolean(keptBefore) && keptAfter === keptBefore && Boolean(held) && held.thumbprint !== keptBefore && second.tone !== 'good' && /is not the one kept for amina@northwind\.example before/.test(second.text),
      JSON.stringify(second && { signed: second.signed, tone: second.tone, text: second.text, keptBefore, keptAfter, held: held?.thumbprint }));
    await shot(win, 'smime-9-certificate-changed.png');

    // Changed after signing.
    await openSubject('Harbour lease: altered copy');
    const changed = await strip((s) => s.head.includes('altered copy'));
    check('smime: a message changed after it was signed says so, and not that it is valid', changed?.signed === 'changed' && /has changed since it was signed/.test(changed.text), JSON.stringify(changed && { signed: changed.signed, text: changed.text }));
    await shot(win, 'smime-2-changed.png');

    // A signer no authority here vouches for.
    await openSubject('An offer from an unknown signer');
    const untrusted = await strip((s) => s.head.includes('unknown signer'));
    check('smime: a signature no authority on this computer vouches for says so', untrusted?.signed === 'untrusted' && /no authority this computer trusts vouches for the certificate/.test(untrusted.text), JSON.stringify(untrusted && { signed: untrusted.signed, text: untrusted.text }));

    // Encrypted, and nothing here to open it.
    await openSubject('Harbour lease: the sealed terms');
    const locked = await strip((s) => s.head.includes('sealed terms') && s.encrypted === 'no-identity');
    const offered = await js(`Boolean(document.querySelector('.ml-smime-import'))`);
    check('smime: an encrypted message with no certificate here to open it says so and offers to import one; nothing of its body shows',
      Boolean(locked) && offered && /no certificate on this computer opens it/.test(locked.text) && !/quartzberry/.test(locked.body),
      JSON.stringify(locked && { encrypted: locked.encrypted, text: locked.text }));
    await shot(win, 'smime-3-encrypted-locked.png');

    // Import the owner's .p12 from the strip: a wrong password first.
    await js(`document.querySelector('.ml-smime-import')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.ml-cert-path input'))`), 'the Certificates dialog', 5000).catch(() => {});
    const setInput = (selector, value) => js(`(() => { const el = document.querySelector(${JSON.stringify(selector)}); if (!el) return false; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, ${JSON.stringify(value)}); el.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
    await setInput('.ml-cert-path input', p12);
    await setInput('input.ml-cert-password', 'harbour 2026');
    await until(() => js(`document.querySelector('.ml-cert-ok')?.disabled === false`), 'Import to wake', 3000).catch(() => {});
    await js(`document.querySelector('.ml-cert-ok')?.click(), 1`);
    const refused = await waitFor(() => js(`document.querySelector('.ml-cert-error')?.textContent || ''`), 'the wrong password refused', 8000);
    check('smime: a wrong password for the .p12 is refused in a sentence', /^That password is not right for this certificate\.$/.test(refused || ''), refused);
    await setInput('input.ml-cert-password', PASSWORD);
    await js(`document.querySelector('.ml-cert-ok')?.click(), 1`);
    const listed = await waitFor(() => js(`(() => { const c = document.querySelector('.ml-cert[data-thumbprint]'); return c ? { thumbprint: c.dataset.thumbprint, text: c.textContent } : null; })()`), 'the certificate listed', 8000);
    thumbprint = listed?.thumbprint || null;
    await shot(win, 'smime-4-certificates.png');
    check('smime: the right password imports the certificate, listed with its address, issuer and date', Boolean(listed) && /owner@example\.com/.test(listed.text) && /Rutba Check Authority Root/.test(listed.text), JSON.stringify(listed));
    const asked = await waitFor(() => js(`(() => { const c = document.querySelector('.ml-cert.changed[data-email="amina@northwind.example"]'); return c ? { text: c.textContent, accept: Boolean(c.querySelector('.ml-cert-accept')), dismiss: Boolean(c.querySelector('.ml-cert-dismiss')) } : null; })()`), 'the changed certificate in the dialog', 4000);
    await js(`document.querySelector('.ml-cert.changed[data-email="amina@northwind.example"] .ml-cert-dismiss')?.click(), 1`);
    const cleared = await waitFor(() => js(`!document.querySelector('.ml-cert.changed')`), 'the held certificate dismissed', 4000);
    const stillKept = (await mail('smimeCertificates'))['amina@northwind.example']?.[0]?.thumbprint || null;
    check('smime: Certificates lists the certificate that changed, to use the new one or keep the old; keeping the old forgets the new',
      Boolean(asked) && asked.accept && asked.dismiss && /Use it only if they told you it changed/.test(asked.text) && Boolean(cleared) && stillKept === keptBefore,
      JSON.stringify({ asked, cleared, stillKept, keptBefore }));
    await js(`[...document.querySelectorAll('.rw-dialog-foot button')].find((b) => b.textContent.trim() === 'Close')?.click(), 1`);
    await until(() => js(`!document.querySelector('.ml-cert-path')`), 'the dialog closed', 4000).catch(() => {});

    // The encrypted message now opens, its signature checked inside.
    const opened = await strip((s) => s.head.includes('sealed terms') && s.encrypted === 'decrypted');
    check('smime: once the certificate is imported the encrypted message opens, says how, and its signature inside is checked',
      Boolean(opened) && opened.signed === 'valid' && /Encrypted \(AES-256\), and opened with your certificate for owner@example\.com\./.test(opened.text) && /sealed quartzberry terms/.test(opened.body),
      JSON.stringify(opened && { signed: opened.signed, encrypted: opened.encrypted, text: opened.text }));
    await shot(win, 'smime-5-encrypted-opened.png');

    // Encrypted but signed by nobody: plain text, no link, nothing fetched (EFAIL).
    await openSubject('Harbour lease: unsigned and sealed');
    const unsigned = await strip((s) => s.head.includes('unsigned and sealed') && s.encrypted === 'decrypted');
    check('smime: an encrypted message nobody signed opens as plain text, with no link and nothing fetched, and says why',
      Boolean(unsigned) && /unsigned quartzberry note/.test(unsigned.body) && /<pre class="plain">/.test(unsigned.body) && !/<a\b|<img\b/i.test(unsigned.body) && /shown as plain text/.test(unsigned.text) && !fetchesRemote(unsigned.body),
      JSON.stringify(unsigned && { encrypted: unsigned.encrypted, signed: unsigned.signed, text: unsigned.text, body: unsigned.body.slice(-400) }));
    await shot(win, 'smime-10-unsigned-plain.png');

    // A new message to Amina, signed and encrypted (her certificate was kept from her signed mail).
    await js(`window.rutbaOffice.store.set({ key: 'mail.undoSeconds', value: 0 }).then(() => 1)`);
    const setField = (label, value) => js(`(() => {
      const el = [...document.querySelectorAll('.rw-field')].find((f) => f.querySelector('label')?.textContent.trim() === ${JSON.stringify(label)})?.querySelector('input, textarea');
      if (!el) return false;
      const proto = el.tagName === 'TEXTAREA' ? window.HTMLTextAreaElement.prototype : window.HTMLInputElement.prototype;
      Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, ${JSON.stringify(value)});
      el.dispatchEvent(new Event('input', { bubbles: true }));
      return true;
    })()`);
    await js(`document.querySelector('.ml-compose-cta button')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.ml-to-row'))`), 'the composer', 5000).catch(() => {});
    await setField('To', people.amina.email);
    await setField('Subject', 'Re: the sealed terms');
    await until(() => js(`document.querySelector('.ml-security-sign')?.disabled === false`), 'Sign to be offered', 5000).catch(() => {});
    await js(`document.querySelector('.ml-security-sign')?.click(), 1`);
    await js(`document.querySelector('.ml-security-encrypt')?.click(), 1`);
    const toggles = await waitFor(() => js(`(() => { const s = document.querySelector('.ml-security-sign'); const e = document.querySelector('.ml-security-encrypt'); return s?.getAttribute('aria-pressed') === 'true' && e?.getAttribute('aria-pressed') === 'true' ? { who: document.querySelector('.ml-security-who')?.textContent || '' } : null; })()`), 'Sign and Encrypt pressed', 4000);
    await shot(win, 'smime-6-compose.png');
    await js(`(() => { if ([...document.querySelectorAll('.rw-field label')].some((l) => l.textContent.trim() === 'Bcc')) return 1; [...document.querySelectorAll('.ml-linkchip')].find((b) => b.textContent.trim() === 'Cc / Bcc')?.click(); return 1; })()`);
    await until(() => js(`[...document.querySelectorAll('.rw-field label')].some((l) => l.textContent.trim() === 'Bcc')`), 'the Bcc line', 4000).catch(() => {});
    const bccSet = await setField('Bcc', people.bo.email);
    await js(`[...document.querySelectorAll('.rw-dialog-foot button')].find((b) => b.textContent.trim() === 'Send')?.click(), 1`);
    const bccRefused = await waitFor(() => js(`[...document.querySelectorAll('.rw-toast')].map((n) => n.textContent).find((x) => /Bcc/.test(x)) || ''`), 'the Bcc refusal', 5000);
    const stillOpen = await js(`Boolean(document.querySelector('.ml-to-row'))`);
    check('smime: an encrypted message with a Bcc is refused, saying why, and the message stays open', /Encrypted mail is not sent to Bcc recipients/.test(bccRefused || '') && stillOpen, JSON.stringify({ bccSet, bccRefused, stillOpen }));
    await shot(win, 'smime-11-bcc-refused.png');
    await setField('Bcc', '');
    await js(`[...document.querySelectorAll('.rw-dialog-foot button')].find((b) => b.textContent.trim() === 'Send')?.click(), 1`);
    const sentRow = await waitFor(async () => {
      const rows = (await mail('messages', { accountId, folder: 'Sent', limit: 50 })).rows;
      return rows.find((r) => r.subject === 'Re: the sealed terms') || false;
    }, 'the message in Sent', 20000);
    let sentView = null;
    if (sentRow) sentView = await mail('message', { accountId, folder: 'Sent', id: sentRow.id });
    // What went out, opened with Amina's key.
    const { app } = await import('electron');
    const original = findFiles(path.join(app.getPath('userData'), 'mail'), 'smime-original.eml').find((f) => fs.readFileSync(f, 'latin1').includes('Re: the sealed terms'));
    const asAmina = original ? readSmime(fs.readFileSync(original), { identities: [{ cert: certificate(people.amina.cert), key: people.amina.key }], anchors: [certificate(people.ca.cert)] }) : null;
    check('smime: a new message signed and encrypted to Amina goes out; its Sent copy opens with your own certificate, and Amina\'s key opens what was sent, signed by you',
      Boolean(toggles) && sentView?.smime?.encrypted?.status === 'decrypted' && sentView?.smime?.signed?.status === 'valid' && asAmina?.encrypted?.status === 'decrypted' && asAmina?.signed?.status === 'valid' && asAmina?.signed?.signer?.email === 'owner@example.com',
      JSON.stringify({ toggles, sent: sentView?.smime && { signed: sentView.smime.signed?.status, encrypted: sentView.smime.encrypted?.status }, asAmina: asAmina && { encrypted: asAmina.encrypted?.status, signed: asAmina.signed?.status, by: asAmina.signed?.signer?.email }, original: Boolean(original) }));

    // Less the Bcc refusal the check itself provoked.
    const complaints = (await errorsIn(win)).filter((c) => !/not sent to Bcc recipients/.test(c));
    check('smime: the window reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');

    // The strip in Urdu, right to left.
    await js(`window.rutbaOffice.store.set({ key: 'language', value: 'ur' }).then(() => 1)`);
    const w = await open('mail');
    const wjs = (code) => w.webContents.executeJavaScript(code);
    await until(() => wjs(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the Urdu window', 10000).catch(() => {});
    await wjs(`(() => { const b = [...document.querySelectorAll('.ml-accounts button')].find((n) => n.textContent.trim().startsWith('Archive Owner')); b?.click(); return Boolean(b); })()`);
    await until(() => wjs(`document.querySelectorAll('.ml-row').length >= 7`), 'the messages in Urdu', 10000).catch(() => {});
    await wjs(`(() => { const r = [...document.querySelectorAll('.ml-row .ml-subject')].find((n) => n.textContent.trim() === 'Harbour lease: the sealed terms'); r?.click(); return Boolean(r); })()`);
    const urdu = await waitFor(async () => {
      const s = await wjs(`(() => { const el = document.querySelector('.ml-smime'); if (!el) return null; const r = (n) => n?.getBoundingClientRect(); const row = el.querySelector('.ml-smime-signed'); return { dir: document.documentElement.dir, encrypted: el.dataset.encrypted, signed: el.dataset.signed, text: el.textContent, iconRight: row ? r(row.querySelector('svg')).left > r(row.querySelector('.ml-smime-text')).left : false }; })()`);
      return s && s.encrypted === 'decrypted' && /[؀-ۿ]/.test(s.text) ? s : false;
    }, 'the strip in Urdu', 10000);
    await shot(w, 'smime-7-urdu.png');
    check('smime: in Urdu the strip reads right to left in Urdu', Boolean(urdu) && urdu.dir === 'rtl' && urdu.iconRight && !/Signed by|Encrypted \(/.test(urdu.text), JSON.stringify(urdu));
    w.destroy();
  } catch (err) {
    check('smime: the checks ran', false, err.stack || err.message);
  } finally {
    try {
      if (js) {
        await js(`window.rutbaOffice.store.set({ key: 'language', value: ${JSON.stringify(language ?? null)} }).then(() => 1)`).catch(() => {});
        await js(`window.rutbaOffice.store.set({ key: 'mail.undoSeconds', value: ${JSON.stringify(undo ?? 8)} }).then(() => 1)`).catch(() => {});
        if (thumbprint) await js(`window.rutbaOffice.mail.removeIdentity({ thumbprint: ${JSON.stringify(thumbprint)} })`).catch(() => {});
        await js(`window.rutbaOffice.mail.forgetCertificates({ emails: ['amina@northwind.example', 'stranger@elsewhere.example', 'owner@example.com', 'bo@contoso.example'] })`).catch(() => {});
        if (accountId) await js(`window.rutbaOffice.mail.removeAccount({ id: ${JSON.stringify(accountId)}, keepMail: false })`).catch(() => {});
      }
    } finally {
      fs.rmSync(root, { recursive: true, force: true });
    }
  }
}
