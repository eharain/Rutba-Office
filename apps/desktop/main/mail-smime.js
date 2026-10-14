// Signed and encrypted mail (S/MIME) for the Mail window.
//
// A person's own certificate and key come from a .p12 or .pfx they import,
// opened with its password: the key is kept in the system's secret store
// (Windows' data protection, the same place account passwords go), never in
// a settings file, and the certificate's facts are listed in the settings.
//
// Other people's certificates - to encrypt to them - are kept from the signed
// mail they send, but only from a signature that is good in every way: an
// authority this computer trusts vouches for it, it is for mail, in date, and
// for the very address the message is from. Anything less (a certificate
// nobody vouches for, one for another address) is never kept on its own: a
// person can import it deliberately, from a .cer they were given. And a
// certificate kept for an address is never replaced behind the person's back:
// a different one arriving for the same address is held aside and the window
// says so, until the person chooses it in Certificates. Encrypting to someone
// uses only a certificate that is still trusted, or one they imported
// themselves.
//
// Reading a message, and opening a .p12, happen on a thread of their own with
// a time and a memory budget (@rutba/mailbox/smime-host): a message built to
// hang or swell the reader costs that thread, not the application.
// Nothing is sent anywhere to check a certificate, and revocation is not
// checked, which the window says.

import fs from 'node:fs';
import crypto from 'node:crypto';
import { protect, certificate, describe, pem, pemBlocks, forEncrypting, entityFromAttachment, systemAnchors, chainOf, headEnd } from '@rutba/mailbox/smime';
import { createSmimeHost } from '@rutba/mailbox/smime-host';

const IDENTITIES = 'mail.smime.identities';
const CERTIFICATES = 'mail.smime.certificates';
/** Certificates that arrived for an address that already has one: held aside until the person chooses. */
const PENDING = 'mail.smime.pending';
const secretKey = (thumbprint) => `mail.smime.identity.${thumbprint}`;

/** A certificate file's certificates: PEM text or DER bytes. */
function certificatesIn(bytes) {
  const text = Buffer.from(bytes).toString('latin1');
  const blocks = pemBlocks(text);
  return (blocks.length ? blocks : [Buffer.from(bytes)]).map((der) => certificate(der));
}

/**
 * @param {{ stores: { settings, secrets }, checkRun?: boolean, host?: object }} deps
 */
export function createSmimeDesk({ stores, checkRun = false, host = null }) {
  const settings = stores.settings;
  const secrets = stores.secrets;
  const thread = host || createSmimeHost();
  let loaded = null; // thumbprint -> { cert, key, keyPem, chain }
  const extraAnchors = [];

  const listed = () => settings.get(IDENTITIES, []);
  const kept = () => settings.get(CERTIFICATES, {});
  const pending = () => settings.get(PENDING, {});

  /** The identities with their keys, read from the secret store once. */
  function identities() {
    if (loaded) return loaded;
    loaded = new Map();
    for (const entry of listed()) {
      try {
        const stored = JSON.parse(secrets.get(secretKey(entry.thumbprint)) || 'null');
        if (!stored) continue;
        loaded.set(entry.thumbprint, { cert: certificate(stored.cert), key: crypto.createPrivateKey(stored.key), keyPem: stored.key, chain: (stored.chain || []).map((c) => certificate(c)) });
      } catch {
        /* a key the secret store no longer opens: the identity is listed but cannot be used */
      }
    }
    return loaded;
  }

  /** Every certificate kept for other people (and one's own), as certificate objects. */
  function others() {
    const out = [];
    for (const list of Object.values(kept())) {
      for (const c of list) {
        try {
          out.push(certificate(c.pem));
        } catch {
          /* a certificate that no longer reads */
        }
      }
    }
    for (const id of identities().values()) out.push(id.cert, ...id.chain);
    return out;
  }

  const anchors = () => (extraAnchors.length ? [...systemAnchors(), ...extraAnchors] : systemAnchors());
  const entryOf = (c, how) => ({ thumbprint: c.thumbprint, pem: pem(c.der), how, kept: new Date().toISOString(), validTo: c.validTo.toISOString() });

  /**
   * Keep a certificate for one address. An address with none takes it; an
   * address that has another is not changed - the new one is held aside and
   * the change is reported, for the person to decide.
   * @returns {'kept'|'same'|'changed'}
   */
  function keepFor(email, c, how) {
    const map = { ...kept() };
    const list = map[email] || [];
    if (list.some((x) => x.thumbprint === c.thumbprint)) return 'same';
    if (list.length) {
      settings.set(PENDING, { ...pending(), [email]: { ...entryOf(c, how), name: c.name } });
      return 'changed';
    }
    map[email] = [entryOf(c, how)];
    settings.set(CERTIFICATES, map);
    return 'kept';
  }

  /**
   * The certificate to encrypt to an address: the person's own, one they
   * imported themselves, or one kept from signed mail that is still trusted -
   * in date, for mail, an RSA key.
   */
  function certificateFor(email, now = new Date()) {
    const address = String(email || '').toLowerCase();
    const own = [...identities().values()].find((id) => id.cert.emails.includes(address));
    if (own) return own.cert;
    const pool = others();
    const usable = (kept()[address] || [])
      .map((entry) => {
        try {
          return { entry, cert: certificate(entry.pem) };
        } catch {
          return null;
        }
      })
      .filter(Boolean)
      .filter(({ cert }) => cert.validFrom <= now && now <= cert.validTo && cert.x509.publicKey.asymmetricKeyType === 'rsa' && forEncrypting(cert))
      .filter(({ entry, cert }) => entry.how === 'imported' || chainOf(cert, pool, anchors()).trusted);
    return usable.sort((a, b) => b.cert.validTo - a.cert.validTo)[0]?.cert || null;
  }

  // The address in the message's own From header (its headers only, never a line in the body).
  const fromOf = (raw) => {
    const head = Buffer.from(raw.subarray(0, headEnd(raw).head)).toString('latin1');
    const m = /^from:[ \t]*(.*(?:\r?\n[ \t].*)*)/im.exec(head);
    const v = m ? m[1].replace(/\r?\n[ \t]+/g, ' ') : '';
    // Found by scanning, not by a pattern that backtracks: a From line far past
    // the standard's length, one long word with no spaces, held the main
    // process for minutes under the pattern this replaces.
    const lt = v.indexOf('<');
    const gt = lt < 0 ? -1 : v.indexOf('>', lt + 1);
    if (gt > lt + 1) {
      const inner = v.slice(lt + 1, gt);
      const at = inner.indexOf('@');
      if (at > 0 && at < inner.length - 1 && !/[\s<>]/.test(inner) && inner.indexOf('@', at + 1) < 0) return inner.toLowerCase();
    }
    for (const word of v.split(/[\s<>",;]+/)) {
      const at = word.indexOf('@');
      if (at > 0 && at < word.length - 1) return word.toLowerCase();
    }
    return '';
  };

  const desk = {
    /** The person's own certificates, without their keys. */
    identities: () => listed(),

    /** Import a .p12/.pfx, opened on the S/MIME thread: its key goes to the secret store, its facts to the list. */
    importIdentity: async ({ path: file, password = '' }) => {
      if (!secrets.available()) throw new Error('This computer offers no protected store for keys, so a certificate with its key cannot be kept here.');
      const bytes = fs.readFileSync(file);
      let got;
      try {
        got = await thread.pkcs12({ bytes, password });
      } catch (err) {
        // The sentence alone reaches the window, not the name of the error.
        throw Object.assign(new Error(err.message), { code: err.code });
      }
      const cert = certificate(Buffer.from(got.cert));
      const chain = got.chain.map((c) => certificate(Buffer.from(c)));
      const facts = { ...describe(cert), thumbprint: cert.thumbprint, signs: true, encrypts: cert.x509.publicKey.asymmetricKeyType === 'rsa', imported: new Date().toISOString() };
      secrets.set(secretKey(cert.thumbprint), JSON.stringify({ cert: pem(cert.der), key: got.key, chain: chain.map((c) => pem(c.der)) }));
      settings.set(IDENTITIES, [...listed().filter((e) => e.thumbprint !== cert.thumbprint), facts]);
      loaded = null;
      return facts;
    },

    removeIdentity: ({ thumbprint }) => {
      secrets.delete(secretKey(thumbprint));
      settings.set(IDENTITIES, listed().filter((e) => e.thumbprint !== thumbprint));
      loaded = null;
      return { removed: true };
    },

    /**
     * Someone else's certificate, from a .cer/.crt/.pem they sent: the
     * person's own decision, so it is kept even when no authority here
     * vouches for it, for each address it is for (an address that has
     * another keeps it, the new one held aside to choose).
     */
    importCertificate: ({ path: file }) => {
      const certs = certificatesIn(fs.readFileSync(file)).filter((c) => !c.ca);
      const out = [];
      for (const c of certs) for (const email of c.emails) out.push({ ...describe(c), email, result: keepFor(email, c, 'imported') });
      return out;
    },

    /** Certificates held aside: a different one arrived for an address that has one. */
    pendingCertificates: () => Object.fromEntries(Object.entries(pending()).map(([email, e]) => [email, { ...describe(certificate(e.pem)), how: e.how, seen: e.kept }])),

    /** Use the certificate held aside for an address instead of the one kept. */
    acceptCertificate: ({ email }) => {
      const address = String(email).toLowerCase();
      const held = pending()[address];
      if (!held) return { accepted: false };
      settings.set(CERTIFICATES, { ...kept(), [address]: [{ ...held, kept: new Date().toISOString() }] });
      const rest = { ...pending() };
      delete rest[address];
      settings.set(PENDING, rest);
      return { accepted: true };
    },

    /** Keep the certificate already kept, and forget the one held aside. */
    dismissCertificate: ({ email }) => {
      const rest = { ...pending() };
      delete rest[String(email).toLowerCase()];
      settings.set(PENDING, rest);
      return { dismissed: true };
    },

    /** Forget the certificates kept for these addresses. */
    forgetCertificates: ({ emails = [] }) => {
      const map = { ...kept() };
      const held = { ...pending() };
      for (const e of emails) {
        delete map[String(e).toLowerCase()];
        delete held[String(e).toLowerCase()];
      }
      settings.set(CERTIFICATES, map);
      settings.set(PENDING, held);
      return { forgotten: emails.length };
    },

    /** Certificates kept for other people, by address. */
    certificates: () => Object.fromEntries(Object.entries(kept()).map(([email, list]) => [email, list.map((c) => ({ ...describe(certificate(c.pem)), how: c.how, kept: c.kept }))])),

    /** Who among these addresses a message could be encrypted for, and who not. */
    canEncryptFor: ({ addresses = [] }) => {
      const missing = addresses.filter((a) => !certificateFor(a));
      return { ok: missing.length === 0, missing };
    },

    /** For a check run only: trust a test authority, as a person would add one to Windows. */
    trustForTest: ({ pem: text }) => {
      if (!checkRun) throw new Error('Test authorities are trusted only during verification.');
      extraAnchors.push(...certificatesIn(Buffer.from(text, 'latin1')));
      return { anchors: extraAnchors.length };
    },
  };

  /**
   * Read an S/MIME message (as it arrived), on the S/MIME thread. A signer's
   * certificate is kept only from a signature that is good in every way.
   * @returns {Promise<{ message, smime: { signed, encrypted, error?, certificateChanged? } } | null>} null when it is not S/MIME;
   *   `message` null when it could not be read (smime.error says why)
   */
  async function read(raw) {
    if (!raw) return null;
    let r;
    try {
      r = await thread.read({
        raw: Buffer.from(raw),
        identities: [...identities().values()].map((id) => ({ cert: id.cert.der, key: id.keyPem })),
        certificates: others().map((c) => c.der),
        extraAnchors: extraAnchors.map((c) => c.der),
      });
    } catch (err) {
      return { message: null, smime: { signed: null, encrypted: null, error: 'unreadable' } };
    }
    if (r.error) return { message: r.message, smime: { signed: null, encrypted: null, error: r.error } };
    if (!r.signed && !r.encrypted) return null;
    const smime = { signed: r.signed, encrypted: r.encrypted };
    if (r.signed?.status === 'valid') {
      const from = fromOf(Buffer.from(raw));
      for (const derBytes of r.certificates) {
        let c;
        try {
          c = certificate(Buffer.from(derBytes));
        } catch {
          continue;
        }
        const signer = r.signed.signers.find((s) => s.status === 'valid' && s.signer?.thumbprint === c.thumbprint);
        if (!signer || !from || !c.emails.includes(from) || c.ca) continue;
        if (keepFor(from, c, 'signed') === 'changed') smime.certificateChanged = { email: from };
      }
    }
    return { message: r.message, smime };
  }

  /** The bytes to read for a message kept by a store or an archive: its original, or an Outlook data file's smime.p7m. */
  function rawOf(message, original = null) {
    if (original) return original;
    if (message?.smimeRaw) return message.smimeRaw;
    const p7m = (message?.attachments || []).find((a) => /^smime\.p7m$/i.test(a.filename || '') && a.bytes);
    const entity = p7m ? entityFromAttachment(p7m) : null;
    if (!entity) return null;
    const who = (list) => (list || []).map((p) => (p.name ? `"${p.name}" <${p.address}>` : p.address)).join(', ');
    const from = Array.isArray(message.from) ? message.from[0] : message.from;
    const head = [from ? `From: ${who([from])}` : null, message.to?.length ? `To: ${who(message.to)}` : null, `Subject: ${message.subject || ''}`, message.date ? `Date: ${new Date(message.date).toUTCString()}` : null].filter(Boolean).join('\r\n');
    return Buffer.concat([Buffer.from(`${head}\r\n`, 'utf8'), entity]);
  }

  /**
   * Sign and/or encrypt a message about to be sent. `from` names the
   * identity; encryption is for every recipient and the sender, and refused
   * (naming who) when a recipient's certificate is not kept. Blind copies are
   * never encrypted together with the rest: each recipient's certificate is
   * named in the message, which would show the others who had one.
   */
  function protectOutgoing(raw, { sign = false, encrypt = false, from, recipients = [], bcc = [] }) {
    const own = [...identities().values()].find((id) => id.cert.emails.includes(String(from || '').toLowerCase()));
    if ((sign || encrypt) && !own) throw Object.assign(new Error(`There is no certificate for ${from} on this computer. Import yours (a .p12 or .pfx) to sign or encrypt.`), { code: 'no-identity' });
    let encryptFor = null;
    if (encrypt) {
      if (bcc.length) throw Object.assign(new Error('Encrypted mail is not sent to Bcc recipients: every recipient\'s certificate is named inside the message, so the others would see who had a blind copy. Send it to them separately.'), { code: 'bcc' });
      const missing = [];
      const certs = [];
      for (const r of recipients) {
        const c = certificateFor(r);
        if (c) certs.push(c.der);
        else missing.push(r);
      }
      if (missing.length) throw Object.assign(new Error(`This cannot be encrypted for ${missing.join(', ')}: there is no certificate for them here. Ask them for a signed message, or import their certificate.`), { code: 'no-certificate', missing });
      encryptFor = [...certs, own.cert.der];
    }
    return protect(raw, { sign: sign ? { key: own.key, cert: own.cert.der, chain: own.chain.map((c) => c.der) } : null, encryptFor });
  }

  return { desk, read, rawOf, protectOutgoing, keepFor, stop: () => thread.stop() };
}
