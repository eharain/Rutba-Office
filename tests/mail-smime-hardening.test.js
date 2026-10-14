// S/MIME, against messages and certificates written to do harm: each test
// proves one fix from the security review. Reading stays linear in the size
// of what arrived, its structures are capped, it runs on a thread of its own
// with a time limit, a certificate is kept only from a signature good in every
// way and never replaced unasked, a chain is checked link by link for mail,
// a good signature by someone else is not shown as good, weak algorithms are
// said, and a mismatched signature algorithm is refused.

import test from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { readSmime, protect, certificate, headEnd, canonical, chainOf, forEncrypting, forSigning, MAX_SMIME_BYTES } from '@rutba/mailbox/smime';
import { createSmimeHost } from '@rutba/mailbox/smime-host';
import { signDetached } from '../packages/mailbox/src/smime/cms.js';
import { parse, intOf, smallInt, seq, set, int, oid, octets, explicit, algorithm, tlv, MAX_NODES, MAX_DEPTH } from '../packages/mailbox/src/smime/der.js';
import { OID } from '../packages/mailbox/src/smime/oids.js';
import { readPkcs12, writePkcs12 } from '../packages/mailbox/src/smime/pkcs12.js';
import { createSmimeDesk } from '../apps/desktop/main/mail-smime.js';
import { makeIdentities, makeCertificate } from './fixtures/smime-keys.js';

const DAY = 86_400_000;
const people = makeIdentities();
const ca = { cn: 'Rutba Test Authority Root', o: 'Rutba Test Authority' };
const anchors = [certificate(people.ca.cert)];
const rsa = () => crypto.generateKeyPairSync('rsa', { modulusLength: 2048 });
const signer = (p, chain = [people.ca.cert]) => ({ key: p.key, cert: p.cert, chain });

/** A person whose certificate the test authority (or `issuer`) issues, with any of makeCertificate's awkward options. */
function person(cn, email, opts = {}) {
  const keys = rsa();
  const cert = makeCertificate({ subject: { cn, email }, issuer: opts.issuer || ca, publicKey: keys.publicKey, signingKey: opts.issuerKey || people.ca.key, email, ...opts });
  return { key: keys.privateKey, publicKey: keys.publicKey, cert, email, name: cn };
}

/** An authority under the test root (or another issuer). */
function authority(cn, opts = {}) {
  const keys = rsa();
  const subject = { cn, o: 'Rutba Test Middle' };
  const cert = makeCertificate({ subject, issuer: opts.issuer || ca, publicKey: keys.publicKey, signingKey: opts.issuerKey || people.ca.key, ca: true, ...opts });
  return { key: keys.privateKey, cert, subject };
}

const message = (from, body = 'The quartzberry clause stands.\r\n') =>
  Buffer.from(`From: ${from.name} <${from.email}>\r\nTo: Archive Owner <owner@example.com>\r\nSubject: Lease\r\nDate: ${new Date().toUTCString()}\r\nMIME-Version: 1.0\r\nContent-Type: text/plain; charset=utf-8\r\n\r\n${body}`, 'latin1');

/** An opaque signed message (the content inside the signature), from `from`, signed as given. */
function opaque(from, cms) {
  return Buffer.from(`From: ${from.name} <${from.email}>\r\nTo: owner@example.com\r\nSubject: Opaque\r\nMIME-Version: 1.0\r\nContent-Type: application/pkcs7-mime; smime-type=signed-data; name="smime.p7m"\r\nContent-Transfer-Encoding: base64\r\n\r\n${cms.toString('base64').replace(/.{76}/g, '$&\r\n')}\r\n`, 'latin1');
}
const inner = Buffer.from('Content-Type: text/plain; charset=utf-8\r\n\r\nOpaque quartzberry.\r\n', 'latin1');

const timed = (fn) => {
  const t = process.hrtime.bigint();
  const value = fn();
  return { value, ms: Number(process.hrtime.bigint() - t) / 1e6 };
};

// H1 (a): finding where the headers end, and making line ends canonical, were regular expressions that took
// time in the square of a run of carriage returns. Now one pass over the bytes.
test('H1: headEnd and canonical take time in proportion to a run of carriage returns, not its square', () => {
  for (const n of [1_000_000, 4_000_000]) {
    const afterNewline = Buffer.from(`A\n${'\r'.repeat(n)}x`, 'latin1');
    const alone = Buffer.from(`${'\r'.repeat(n)}x`, 'latin1');
    const a = timed(() => headEnd(afterNewline));
    const b = timed(() => canonical(afterNewline));
    const c = timed(() => canonical(alone));
    // The expressions they replace took minutes at a million; a byte loop takes milliseconds.
    assert.ok(a.ms < 1500 && b.ms < 1500 && c.ms < 1500, `${n} carriage returns: headEnd ${a.ms} ms, canonical ${b.ms} / ${c.ms} ms`);
    assert.equal(a.value.head, afterNewline.length, 'no blank line in it');
  }
  // What it finds is unchanged.
  assert.deepEqual(headEnd(Buffer.from('A: 1\r\nB: 2\r\n\r\nbody', 'latin1')), { head: 10, body: 14 });
  assert.deepEqual(headEnd(Buffer.from('A: 1\nB: 2\n\nbody', 'latin1')), { head: 9, body: 11 });
  assert.deepEqual(headEnd(Buffer.from('A: 1\r\r\nB: 2\r\r\n\r\r\nbody', 'latin1')), { head: 11, body: 17 }, 'CR CR LF, as a text-mode program on Windows writes it');
  assert.equal(canonical(Buffer.from('a\nb\r\nc\r\r\nd\re', 'latin1')).toString('latin1'), 'a\r\nb\r\nc\r\nd\re');
});

// H1 (b): a number in a structure was read through a string a digit at a time (quadratic), and a count could be any length.
test('H1: intOf reads a long number at once, and a count longer than it could be is refused', () => {
  const huge = { value: Buffer.alloc(100_000, 0x41) };
  const r = timed(() => intOf(huge));
  assert.ok(r.ms < 200, `a 100,000-byte number read in ${r.ms} ms`);
  assert.equal(r.value.toString(16).length, 200_000);
  assert.equal(smallInt(parse(int(20))), 20);
  assert.equal(smallInt(parse(int(10_000_000)), 4), 10_000_000);
  assert.throws(() => smallInt(parse(int(2n ** 72n)), 8), /count that cannot be real/, 'nine bytes');
  assert.throws(() => smallInt(parse(int(2n ** 40n)), 4), /count that cannot be real/);
  assert.throws(() => smallInt(parse(int(-1))), /count that cannot be real/, 'a negative count');
  assert.throws(() => smallInt({ value: Buffer.alloc(0) }), /count that cannot be real/, 'an empty one');
});

// H1, the reviewer's own message: a signed part that is nothing but carriage returns.
test('H1: a signed message made of a million carriage returns is read in proportion to its size', () => {
  const sha = '2.16.840.1.101.3.4.2.1';
  const attrs = set(seq(oid('1.2.840.113549.1.9.4'), set(octets(Buffer.alloc(32)))));
  const si = seq(int(1), seq(seq(), int(1)), algorithm(sha), tlv(0xa0, attrs.subarray(2)), algorithm('1.2.840.113549.1.1.1'), octets(Buffer.alloc(8)));
  const ci = seq(oid('1.2.840.113549.1.7.2'), explicit(0, seq(int(1), set(algorithm(sha)), seq(oid('1.2.840.113549.1.7.1')), set(si))));
  const build = (n) =>
    Buffer.from(`From: a@b.example\r\nSubject: x\r\nContent-Type: multipart/signed; protocol="application/pkcs7-signature"; boundary=B\r\n\r\n--B\r\nContent-Type: text/plain\n\nhi\n${'\r'.repeat(n)}x\r\n--B\r\nContent-Type: application/pkcs7-signature\r\nContent-Transfer-Encoding: base64\r\n\r\n${ci.toString('base64')}\r\n--B--\r\n`, 'latin1');
  const r = timed(() => readSmime(build(1_280_000), { anchors: [] }));
  assert.ok(r.ms < 3000, `read in ${r.ms} ms`);
  assert.equal(r.value.signed.status, 'changed');
});

test('H1: a message over 50 MB is not opened, and says so, keeping its headers', () => {
  const head = 'From: Amina Rahman <amina@northwind.example>\r\nSubject: Too big\r\nMIME-Version: 1.0\r\nContent-Type: application/pkcs7-mime; smime-type=enveloped-data\r\nContent-Transfer-Encoding: base64\r\n\r\n';
  const raw = Buffer.concat([Buffer.from(head, 'latin1'), Buffer.alloc(MAX_SMIME_BYTES + 1 - head.length, 0x41)]);
  const r = readSmime(raw, { anchors });
  assert.equal(r.error, 'too-large');
  assert.equal(r.message.subject, 'Too big');
  assert.equal(r.encrypted, null);
});

// H2: a structure of millions of tiny values, or nested a hundred thousand deep, swelled memory or the stack.
test('H2: DER holds at most 100,000 values and nests at most 64 deep', () => {
  const flat = (n) => {
    const b = Buffer.alloc(4 + 2 * n);
    b[0] = 0x30;
    b[1] = 0x80;
    for (let i = 0; i < n; i++) b[2 + 2 * i] = 5;
    return b;
  };
  assert.equal(parse(flat(MAX_NODES - 1)).children.length, MAX_NODES - 1);
  assert.throws(() => parse(flat(MAX_NODES)), /more values than any/);
  const deep = (d) => {
    let node = Buffer.from([5, 0]);
    for (let i = 0; i < d; i++) node = seq(node);
    return node;
  };
  assert.doesNotThrow(() => parse(deep(MAX_DEPTH)));
  assert.throws(() => parse(deep(MAX_DEPTH + 2)), /nested deeper/);
  const indefinite = Buffer.alloc(400_000);
  for (let i = 0; i < 200_000; i++) [indefinite[2 * i], indefinite[2 * i + 1]] = [0x30, 0x80];
  assert.throws(() => parse(indefinite), /nested deeper/, 'the reviewer\'s 200,000-deep message');
});

test('H1: the S/MIME thread ends a read that runs past its time, says it could not be read, and reads the next', async () => {
  const sent = protect(message(people.amina), { sign: signer(people.amina) });
  const anchorsDer = [certificate(people.ca.cert).der];
  const hurried = createSmimeHost({ readMs: 1 });
  try {
    await assert.rejects(hurried.read({ raw: sent, anchors: anchorsDer }), (err) => err.code === 'unreadable');
    assert.equal(hurried.stats.timeouts, 1);
  } finally {
    await hurried.stop();
  }
  const host = createSmimeHost();
  try {
    const r = await host.read({ raw: sent, anchors: anchorsDer });
    assert.equal(r.signed.status, 'valid');
    // A structure that throws comes back as a failure, not a crash.
    const broken = Buffer.from(`From: a@b.example\r\nContent-Type: application/pkcs7-mime; smime-type=signed-data\r\nContent-Transfer-Encoding: base64\r\n\r\n${Buffer.alloc(400, 0x30).toString('base64')}\r\n`, 'latin1');
    const got = await host.read({ raw: broken, anchors: anchorsDer }).catch((err) => ({ failed: err }));
    assert.ok(got.failed || got.signed?.status === 'damaged', 'a broken structure fails or reads as damaged');
    assert.equal((await host.read({ raw: sent, anchors: anchorsDer })).signed.status, 'valid', 'the thread carries on');
  } finally {
    await host.stop();
  }
});

// M2: what is shown as good.
test('M2: a good signature by someone else, or with no From, is not shown as good', () => {
  const other = readSmime(protect(message(people.bo), { sign: signer(people.amina) }), { anchors });
  assert.equal(other.signed.status, 'other-sender');
  const bare = Buffer.from(opaque(people.amina, signDetached(inner, { key: people.amina.key, cert: people.amina.cert, chain: [people.ca.cert], encapsulate: true })).toString('latin1').replace(/^From:.*\r\n/m, ''), 'latin1');
  assert.equal(readSmime(bare, { anchors }).signed.status, 'no-from');
});

test('M2: a certificate out of date now is not good, whatever date the signer wrote in the signature', () => {
  const old = people.expired;
  const back = new Date(Date.now() - 500 * DAY); // while it was in date
  const r = readSmime(opaque(old, signDetached(inner, { key: old.key, cert: old.cert, chain: [people.ca.cert], now: back, encapsulate: true })), { anchors });
  assert.equal(r.signed.status, 'expired', 'signed when it was in date, but out of date now');
  // A signing time in the future proves nothing either.
  const future = readSmime(opaque(people.amina, signDetached(inner, { key: people.amina.key, cert: people.amina.cert, chain: [people.ca.cert], now: new Date(Date.now() + 3650 * DAY), encapsulate: true })), { anchors });
  assert.equal(future.signed.status, 'valid', 'judged by today, the certificate being in date today');
  const later = readSmime(opaque(people.amina, signDetached(inner, { key: people.amina.key, cert: people.amina.cert, chain: [people.ca.cert], encapsulate: true })), { anchors, now: new Date(Date.now() + 400 * DAY) });
  assert.equal(later.signed.status, 'expired', 'read after the certificate ran out');
});

// L3: weak algorithms are said, a signature algorithm naming another hash is refused, and 3DES is not offered.
test('L3: SHA-1 is a caution, and a signature whose algorithm names another hash than its digest is refused', () => {
  const sha1 = signDetached(inner, { key: people.amina.key, cert: people.amina.cert, chain: [people.ca.cert], digest: 'sha1', encapsulate: true });
  const r = readSmime(opaque(people.amina, sha1), { anchors });
  assert.equal(r.signed.status, 'valid');
  assert.deepEqual(r.signed.cautions, ['sha1']);
  // The signer's algorithm rewritten from rsaEncryption to sha256WithRSAEncryption (same length): its digest is SHA-1.
  const rsaOid = oid(OID.rsaEncryption);
  const at = sha1.lastIndexOf(rsaOid);
  const swapped = Buffer.from(sha1);
  oid(OID.sha256WithRSA).copy(swapped, at);
  const bad = readSmime(opaque(people.amina, swapped), { anchors });
  assert.equal(bad.signed.status, 'bad-signature');
  const good = readSmime(opaque(people.amina, signDetached(inner, { key: people.amina.key, cert: people.amina.cert, chain: [people.ca.cert], encapsulate: true })), { anchors });
  assert.deepEqual(good.signed.cautions, []);
});

test('L3: a message encrypted with Triple DES opens with a caution, and Triple DES is no longer offered to senders', () => {
  const cek = crypto.randomBytes(24);
  const iv = crypto.randomBytes(8);
  const c = crypto.createCipheriv('des-ede3-cbc', cek, iv);
  const body = Buffer.concat([c.update(Buffer.from('Content-Type: text/plain\r\n\r\nOld quartzberry.\r\n', 'latin1')), c.final()]);
  const owner = certificate(people.owner.cert);
  const wrapped = crypto.publicEncrypt({ key: owner.x509.publicKey, padding: crypto.constants.RSA_PKCS1_PADDING }, cek);
  const ktri = seq(int(0), seq(owner.issuerRaw, tlv(0x02, owner.serialRaw)), algorithm(OID.rsaEncryption), octets(wrapped));
  const cms = seq(oid(OID.envelopedData), explicit(0, seq(int(0), set(ktri), seq(oid(OID.data), algorithm(OID.desEde3CBC, octets(iv)), tlv(0x80, body)))));
  const raw = Buffer.from(`From: a@b.example\r\nMIME-Version: 1.0\r\nContent-Type: application/pkcs7-mime; smime-type=enveloped-data\r\nContent-Transfer-Encoding: base64\r\n\r\n${cms.toString('base64')}\r\n`, 'latin1');
  const r = readSmime(raw, { identities: [{ cert: owner, key: people.owner.key }], anchors });
  assert.equal(r.encrypted.status, 'decrypted');
  assert.deepEqual(r.encrypted.cautions, ['weak-cipher']);
  assert.match(r.message.text, /Old quartzberry/);
  const signed = signDetached(inner, { key: people.amina.key, cert: people.amina.cert, encapsulate: true });
  assert.equal(signed.indexOf(oid(OID.desEde3CBC)), -1, 'capabilities name AES only');
  assert.ok(signed.indexOf(oid(OID.aes256CBC)) >= 0);
});

// M1: a chain that links up is checked link by link for mail.
test('M1: every authority in a chain must be one, may sign certificates, keeps its path length, and is for mail', () => {
  const pool = (...xs) => xs.map((x) => certificate(x.cert));
  const leafUnder = (mid, email = 'amina@northwind.example', opts = {}) => person('Amina Rahman', email, { issuer: mid.subject, issuerKey: mid.key, ...opts });
  const judge = (leaf, ...mids) => chainOf(certificate(leaf.cert), pool(...mids), anchors);

  const good = authority('Good Middle');
  assert.equal(judge(leafUnder(good), good).trusted, true);

  // Its key may sign certificates, but it is not an authority (basicConstraints CA:FALSE).
  const notCa = authority('Not An Authority', { ca: false, usage: null, keyUsage: 0x06 });
  assert.equal(judge(leafUnder(notCa), notCa).reason, 'not-a-ca', 'an ordinary certificate cannot issue one');

  const noCertSign = authority('No Certificate Signing', { keyUsage: 0x80 });
  const unsigned = judge(leafUnder(noCertSign), noCertSign);
  assert.equal(unsigned.trusted, false, 'an authority whose key may not sign certificates');
  assert.ok(['not-a-ca', 'no-authority'].includes(unsigned.reason), unsigned.reason);

  const webOnly = authority('Web Only Middle', { caUsage: ['1.3.6.1.5.5.7.3.1'] });
  assert.equal(judge(leafUnder(webOnly), webOnly).reason, 'not-for-mail', 'an authority for web servers only');
  const anyPurpose = authority('Any Purpose Middle', { caUsage: [OID.anyExtendedKeyUsage] });
  assert.equal(judge(leafUnder(anyPurpose), anyPurpose).trusted, true);

  // A root allowing no authority below it.
  const rootKeys = rsa();
  const rootName = { cn: 'Short Path Root', o: 'Short' };
  const root = makeCertificate({ subject: rootName, publicKey: rootKeys.publicKey, signingKey: rootKeys.privateKey, ca: true, pathLen: 0 });
  const roots = [certificate(root)];
  const under = authority('Too Far Middle', { issuer: rootName, issuerKey: rootKeys.privateKey });
  assert.equal(chainOf(certificate(leafUnder(under).cert), pool(under), roots).reason, 'path-length');
  const direct = person('Amina Rahman', 'amina@northwind.example', { issuer: rootName, issuerKey: rootKeys.privateKey });
  assert.equal(chainOf(certificate(direct.cert), [], roots).trusted, true, 'a signer issued by the root itself is within it');

  const unknown = leafUnder(good, 'amina@northwind.example', { unknownCritical: true });
  assert.equal(judge(unknown, good).reason, 'unknown-critical');
});

test('M1: an authority\'s name constraints on mail addresses are kept', () => {
  const northwind = authority('Northwind Only', { permitted: ['northwind.example'] });
  const inside = person('Amina Rahman', 'amina@northwind.example', { issuer: northwind.subject, issuerKey: northwind.key });
  const outside = person('Bo Lindqvist', 'bo@contoso.example', { issuer: northwind.subject, issuerKey: northwind.key });
  assert.equal(chainOf(certificate(inside.cert), [certificate(northwind.cert)], anchors).trusted, true);
  assert.equal(chainOf(certificate(outside.cert), [certificate(northwind.cert)], anchors).reason, 'name-constraints');
  const m = certificate(northwind.cert);
  assert.deepEqual(m.nameConstraints.permitted, ['northwind.example']);
});

test('M1: the signer\'s own certificate must be for mail and allowed to sign', () => {
  const web = person('Web Server', 'web@northwind.example', { usage: ['1.3.6.1.5.5.7.3.1'] });
  assert.equal(readSmime(protect(message(web), { sign: signer(web) }), { anchors }).signed.status, 'not-for-mail');
  const encryptOnly = person('Encrypt Only', 'enc@northwind.example', { keyUsage: 0x20 });
  assert.equal(forSigning(certificate(encryptOnly.cert)), false);
  assert.equal(forEncrypting(certificate(encryptOnly.cert)), true);
  assert.equal(readSmime(protect(message(encryptOnly), { sign: signer(encryptOnly) }), { anchors }).signed.status, 'not-for-mail');
  const signOnly = person('Sign Only', 'sig@northwind.example', { keyUsage: 0x80 });
  assert.equal(forEncrypting(certificate(signOnly.cert)), false, 'never encrypted to');
});

// L5: a .p12 that asks for billions of key derivations, or a key of a megabyte.
test('L5: a .p12 asking for more than 10,000,000 iterations is refused before any work', () => {
  const bytes = writePkcs12({ key: people.amina.key, cert: people.amina.cert, chain: [], password: 'pw', name: 'Amina' });
  // The same file with its integrity check asking for 2^31 - 1 derivations (hours of work).
  const [version, authSafe, mac] = parse(bytes).children;
  const swollen = seq(Buffer.from(version.raw), Buffer.from(authSafe.raw), seq(Buffer.from(mac.children[0].raw), Buffer.from(mac.children[1].raw), int(2 ** 31 - 1)));
  const t = timed(() => assert.throws(() => readPkcs12(swollen, 'pw'), (err) => err.code === 'algorithm'));
  assert.ok(t.ms < 2000, `refused in ${t.ms} ms`);
  // Ten million exactly is still allowed (the check is the cap, not a guess).
  const ten = seq(Buffer.from(version.raw), Buffer.from(authSafe.raw), seq(Buffer.from(mac.children[0].raw), Buffer.from(mac.children[1].raw), int(10_000_001)));
  assert.throws(() => readPkcs12(ten, 'pw'), (err) => err.code === 'algorithm');
  assert.equal(readPkcs12(bytes, 'pw').cert.email, 'amina@northwind.example');
});

/** A settings and secrets store in memory, as the desktop gives the desk. */
function memoryStores() {
  const values = new Map();
  const secrets = new Map();
  return {
    settings: { get: (k, d) => (values.has(k) ? structuredClone(values.get(k)) : d), set: (k, v) => values.set(k, structuredClone(v)) },
    secrets: { available: () => true, get: (k) => secrets.get(k) ?? null, set: (k, v) => secrets.set(k, v), delete: (k) => secrets.delete(k) },
  };
}

// H3: key substitution. A certificate is kept from mail only when the signature is good in every way, never replaces one kept.
test('H3: a certificate is kept from signed mail only when trusted, for mail and for the From address', async () => {
  const stores = memoryStores();
  const desk = createSmimeDesk({ stores, checkRun: true });
  try {
    desk.desk.trustForTest({ pem: people.ca.cert });
    // Untrusted: kept? No.
    const stranger = await desk.read(protect(message(people.stranger), { sign: signer(people.stranger, []) }));
    assert.equal(stranger.smime.signed.status, 'untrusted');
    assert.deepEqual(desk.desk.certificates(), {}, 'a certificate nobody vouches for is never kept on its own');
    // Signed by Amina, sent from Bo: neither address keeps Amina's certificate.
    const other = await desk.read(protect(message(people.bo), { sign: signer(people.amina) }));
    assert.equal(other.smime.signed.status, 'other-sender');
    assert.deepEqual(desk.desk.certificates(), {});
    // Good in every way: kept, and Amina can be encrypted to.
    const good = await desk.read(protect(message(people.amina), { sign: signer(people.amina) }));
    assert.equal(good.smime.signed.status, 'valid');
    assert.deepEqual(Object.keys(desk.desk.certificates()), ['amina@northwind.example']);
    assert.equal(desk.desk.canEncryptFor({ addresses: ['amina@northwind.example'] }).ok, true);
    assert.equal(desk.desk.canEncryptFor({ addresses: ['stranger@elsewhere.example'] }).ok, false);
  } finally {
    await desk.stop();
  }
});

test('H3: a different certificate for an address already kept is held aside and reported, never swapped in', async () => {
  const stores = memoryStores();
  const desk = createSmimeDesk({ stores, checkRun: true });
  try {
    desk.desk.trustForTest({ pem: people.ca.cert });
    await desk.read(protect(message(people.amina), { sign: signer(people.amina) }));
    const first = certificate(people.amina.cert).thumbprint;
    // Someone with a certificate the same authority issued for Amina's address (the authority fooled, or a second key).
    const impostor = person('Amina Rahman', 'amina@northwind.example');
    const r = await desk.read(protect(message(impostor), { sign: signer(impostor) }));
    assert.equal(r.smime.signed.status, 'valid', 'the signature itself is good');
    assert.deepEqual(r.smime.certificateChanged, { email: 'amina@northwind.example' });
    assert.equal(desk.desk.certificates()['amina@northwind.example'][0].thumbprint, first, 'the one kept stays');
    assert.equal(desk.desk.pendingCertificates()['amina@northwind.example'].thumbprint, certificate(impostor.cert).thumbprint);
    // Reading it again changes nothing.
    await desk.read(protect(message(impostor), { sign: signer(impostor) }));
    assert.equal(desk.desk.certificates()['amina@northwind.example'].length, 1);
    // The person keeps the old one...
    desk.desk.dismissCertificate({ email: 'amina@northwind.example' });
    assert.deepEqual(desk.desk.pendingCertificates(), {});
    assert.equal(desk.desk.certificates()['amina@northwind.example'][0].thumbprint, first);
    // ...or, told by Amina it changed, chooses the new one.
    await desk.read(protect(message(impostor), { sign: signer(impostor) }));
    assert.equal(desk.desk.acceptCertificate({ email: 'amina@northwind.example' }).accepted, true);
    assert.equal(desk.desk.certificates()['amina@northwind.example'][0].thumbprint, certificate(impostor.cert).thumbprint);
  } finally {
    await desk.stop();
  }
});

test('H3: an untrusted certificate is used to encrypt only when the person imported it, and a kept one only while still trusted', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-smime-h-'));
  const stores = memoryStores();
  const trusting = createSmimeDesk({ stores, checkRun: true });
  try {
    trusting.desk.trustForTest({ pem: people.ca.cert });
    await trusting.read(protect(message(people.amina), { sign: signer(people.amina) }));
    assert.equal(trusting.desk.canEncryptFor({ addresses: ['amina@northwind.example'] }).ok, true);
    const file = path.join(dir, 'stranger.cer');
    fs.writeFileSync(file, people.stranger.cert);
    assert.equal(trusting.desk.importCertificate({ path: file })[0].result, 'kept');
    assert.equal(trusting.desk.canEncryptFor({ addresses: ['stranger@elsewhere.example'] }).ok, true, 'imported on purpose');
  } finally {
    await trusting.stop();
  }
  // The same kept certificates, on a computer that does not trust the test authority.
  const other = createSmimeDesk({ stores });
  try {
    assert.deepEqual(other.desk.canEncryptFor({ addresses: ['amina@northwind.example'] }).missing, ['amina@northwind.example'], 'kept from mail, but no longer trusted');
    assert.equal(other.desk.canEncryptFor({ addresses: ['stranger@elsewhere.example'] }).ok, true);
  } finally {
    await other.stop();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('H1: a message the S/MIME thread cannot read in its time is shown as it arrived, saying so', async () => {
  const desk = createSmimeDesk({ stores: memoryStores(), host: createSmimeHost({ readMs: 1 }) });
  try {
    const r = await desk.read(protect(message(people.amina), { sign: signer(people.amina) }));
    assert.equal(r.message, null);
    assert.equal(r.smime.error, 'unreadable');
  } finally {
    await desk.stop();
  }
});

// L2: encrypting to Bcc would name every blind recipient's certificate inside the message.
test('L2: encrypted mail with Bcc recipients is refused, saying why', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-smime-h-'));
  const stores = memoryStores();
  const desk = createSmimeDesk({ stores, checkRun: true });
  try {
    desk.desk.trustForTest({ pem: people.ca.cert });
    const p12 = path.join(dir, 'owner.p12');
    fs.writeFileSync(p12, writePkcs12({ key: people.owner.key, cert: people.owner.cert, chain: [people.ca.cert], password: 'pw', name: 'Owner' }));
    await desk.desk.importIdentity({ path: p12, password: 'pw' });
    await desk.read(protect(message(people.amina), { sign: signer(people.amina) }));
    const raw = Buffer.from('From: owner@example.com\r\nTo: amina@northwind.example\r\nSubject: x\r\n\r\nhi\r\n', 'latin1');
    assert.throws(() => desk.protectOutgoing(raw, { encrypt: true, from: 'owner@example.com', recipients: ['amina@northwind.example'], bcc: ['bo@contoso.example'] }), (err) => err.code === 'bcc' && /Bcc/.test(err.message));
    const sent = desk.protectOutgoing(raw, { encrypt: true, from: 'owner@example.com', recipients: ['amina@northwind.example'] });
    assert.match(sent.toString('latin1'), /enveloped-data/);
    // Signed only, Bcc is fine: nothing names the recipients.
    assert.doesNotThrow(() => desk.protectOutgoing(raw, { sign: true, from: 'owner@example.com', recipients: ['amina@northwind.example'], bcc: ['bo@contoso.example'] }));
  } finally {
    await desk.stop();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('the From kept against is the header, not a From line in the body', async () => {
  const stores = memoryStores();
  const desk = createSmimeDesk({ stores, checkRun: true });
  try {
    desk.desk.trustForTest({ pem: people.ca.cert });
    // No From header; the body has one naming Amina.
    const raw = Buffer.from(`To: owner@example.com\r\nSubject: x\r\nMIME-Version: 1.0\r\nContent-Type: text/plain\r\n\r\nFrom: Amina Rahman <amina@northwind.example>\r\n`, 'latin1');
    const r = await desk.read(protect(raw, { sign: signer(people.amina) }));
    assert.equal(r.smime.signed.status, 'no-from');
    assert.deepEqual(desk.desk.certificates(), {});
  } finally {
    await desk.stop();
  }
});

