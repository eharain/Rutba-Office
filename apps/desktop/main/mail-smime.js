// Signed and encrypted mail (S/MIME) for the Mail window.
//
// A person's own certificate and key come from a .p12 or .pfx they import,
// opened with its password: the key is kept in the system's secret store
// (Windows' data protection, the same place account passwords go), never in
// a settings file, and the certificate's facts are listed in the settings.
// Other people's certificates - to encrypt to them - are kept from the signed
// mail they send, or imported from a .cer they give.
//
// A message is read through readSmime (@rutba/mailbox/smime): what the
// window is shown is the message inside, with what was found about its
// signature and its encryption beside it. Nothing is sent anywhere to check a
// certificate, and revocation is not checked, which the window says.

import fs from 'node:fs';
import crypto from 'node:crypto';
import { readSmime, protect, readPkcs12, certificate, describe, pem, pemBlocks, forMail, entityFromAttachment, setSystemAnchors, systemAnchors } from '@rutba/mailbox/smime';

const IDENTITIES = 'mail.smime.identities';
const CERTIFICATES = 'mail.smime.certificates';
const secretKey = (thumbprint) => `mail.smime.identity.${thumbprint}`;

/** A certificate file's certificates: PEM text or DER bytes. */
function certificatesIn(bytes) {
  const text = Buffer.from(bytes).toString('latin1');
  const blocks = pemBlocks(text);
  return (blocks.length ? blocks : [Buffer.from(bytes)]).map((der) => certificate(der));
}

/**
 * @param {{ stores: { settings, secrets }, checkRun?: boolean }} deps
 */
export function createSmimeDesk({ stores, checkRun = false }) {
  const settings = stores.settings;
  const secrets = stores.secrets;
  let loaded = null; // thumbprint -> { cert, key, chain }
  const extraAnchors = [];

  const listed = () => settings.get(IDENTITIES, []);
  const kept = () => settings.get(CERTIFICATES, {});

  /** The identities with their keys, read from the secret store once. */
  function identities() {
    if (loaded) return loaded;
    loaded = new Map();
    for (const entry of listed()) {
      try {
        const stored = JSON.parse(secrets.get(secretKey(entry.thumbprint)) || 'null');
        if (!stored) continue;
        loaded.set(entry.thumbprint, { cert: certificate(stored.cert), key: crypto.createPrivateKey(stored.key), chain: (stored.chain || []).map((c) => certificate(c)) });
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

  /** Keep certificates for the addresses they are for (the newest per thumbprint). */
  function keep(certs, how) {
    const map = { ...kept() };
    let changed = false;
    for (const c of certs) {
      if (!forMail(c) || c.ca) continue;
      for (const email of c.emails) {
        const list = map[email] || [];
        if (list.some((x) => x.thumbprint === c.thumbprint)) continue;
        map[email] = [...list, { thumbprint: c.thumbprint, pem: pem(c.der), how, kept: new Date().toISOString(), validTo: c.validTo.toISOString() }];
        changed = true;
      }
    }
    if (changed) settings.set(CERTIFICATES, map);
  }

  const anchors = () => (extraAnchors.length ? [...systemAnchors(), ...extraAnchors] : systemAnchors());

  /** The best certificate kept for an address: in its dates, the latest to expire. */
  function certificateFor(email, now = new Date()) {
    const address = String(email || '').toLowerCase();
    const own = [...identities().values()].find((id) => id.cert.emails.includes(address));
    if (own) return own.cert;
    const list = (kept()[address] || []).map((c) => certificate(c.pem)).filter((c) => c.validFrom <= now && now <= c.validTo && c.x509.publicKey.asymmetricKeyType === 'rsa');
    return list.sort((a, b) => b.validTo - a.validTo)[0] || null;
  }

  const desk = {
    /** The person's own certificates, without their keys. */
    identities: () => listed(),

    /** Import a .p12/.pfx: its key goes to the secret store, its facts to the list. */
    importIdentity: ({ path: file, password = '' }) => {
      if (!secrets.available()) throw new Error('This computer offers no protected store for keys, so a certificate with its key cannot be kept here.');
      const bytes = fs.readFileSync(file);
      let got;
      try {
        got = readPkcs12(bytes, password);
      } catch (err) {
        // The sentence alone reaches the window, not the name of the error.
        throw Object.assign(new Error(err.message), { code: err.code });
      }
      const facts = { ...describe(got.cert), thumbprint: got.cert.thumbprint, signs: true, encrypts: got.cert.x509.publicKey.asymmetricKeyType === 'rsa', imported: new Date().toISOString() };
      secrets.set(secretKey(got.cert.thumbprint), JSON.stringify({ cert: pem(got.cert.der), key: got.key.export({ type: 'pkcs8', format: 'pem' }), chain: got.chain.map((c) => pem(c.der)) }));
      settings.set(IDENTITIES, [...listed().filter((e) => e.thumbprint !== got.cert.thumbprint), facts]);
      loaded = null;
      keep(got.chain, 'imported');
      return facts;
    },

    removeIdentity: ({ thumbprint }) => {
      secrets.delete(secretKey(thumbprint));
      settings.set(IDENTITIES, listed().filter((e) => e.thumbprint !== thumbprint));
      loaded = null;
      return { removed: true };
    },

    /** Someone else's certificate, from a .cer/.crt/.pem they sent. */
    importCertificate: ({ path: file }) => {
      const certs = certificatesIn(fs.readFileSync(file));
      keep(certs, 'imported');
      return certs.map(describe);
    },

    /** Forget the certificates kept for these addresses. */
    forgetCertificates: ({ emails = [] }) => {
      const map = { ...kept() };
      for (const e of emails) delete map[String(e).toLowerCase()];
      settings.set(CERTIFICATES, map);
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
      setSystemAnchors(null);
      return { anchors: extraAnchors.length };
    },
  };

  /**
   * Read an S/MIME message (as it arrived). Signers' certificates found are
   * kept, to encrypt to them later.
   * @returns {{ message, smime: { signed, encrypted } } | null} null when it is not S/MIME
   */
  function read(raw) {
    if (!raw) return null;
    const r = readSmime(raw, { identities: [...identities().values()], certificates: others(), anchors: anchors() });
    if (!r.signed && !r.encrypted) return null;
    if (r.signed && ['valid', 'untrusted', 'expired'].includes(r.signed.status)) keep(r.certificates.filter((c) => r.signed.signers.some((s) => s.signer?.thumbprint === c.thumbprint)), 'signed');
    return { message: r.message, smime: { signed: r.signed, encrypted: r.encrypted } };
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
   * (naming who) when a recipient's certificate is not kept.
   */
  function protectOutgoing(raw, { sign = false, encrypt = false, from, recipients = [] }) {
    const own = [...identities().values()].find((id) => id.cert.emails.includes(String(from || '').toLowerCase()));
    if ((sign || encrypt) && !own) throw Object.assign(new Error(`There is no certificate for ${from} on this computer. Import yours (a .p12 or .pfx) to sign or encrypt.`), { code: 'no-identity' });
    let encryptFor = null;
    if (encrypt) {
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

  return { desk, read, rawOf, protectOutgoing, keep };
}
