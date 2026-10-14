// S/MIME: signed and encrypted mail read, verified, opened and written, with
// certificates, keys and .p12 files made when the test runs.

import test from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import { readSmime, protect, smimeKind, entityFromAttachment, readPkcs12, Pkcs12Error, certificate, splitEntity } from '@rutba/mailbox/smime';
import { signDetached } from '../packages/mailbox/src/smime/cms.js';
import { RC2_PITABLE, rc2Block, rc2Encrypt, rc2Decrypt } from '../packages/mailbox/src/smime/rc2.js';
import { parse, oid, oidOf, int, intOf } from '../packages/mailbox/src/smime/der.js';
import { makeIdentities, makePkcs12 } from './fixtures/smime-keys.js';

const people = makeIdentities();
const anchors = [certificate(people.ca.cert)];
const identity = (p) => ({ cert: certificate(p.cert), key: p.key });
const signer = (p) => ({ key: p.key, cert: p.cert, chain: [people.ca.cert] });

/** A message as a mail program would hand it over to be sent. */
const message = ({ from = people.amina, to = people.owner, subject = 'Harbour lease: signed copy', body = 'The quartzberry clause stands.\r\nSee you on Monday.\r\n' } = {}) =>
  Buffer.from(
    `From: ${from.name} <${from.email}>\r\nTo: ${to.name} <${to.email}>\r\nSubject: ${subject}\r\nDate: ${new Date().toUTCString()}\r\nMessage-ID: <${crypto.randomUUID()}@test.example>\r\nMIME-Version: 1.0\r\nContent-Type: multipart/mixed; boundary="mix"\r\n\r\n--mix\r\nContent-Type: text/plain; charset=utf-8\r\nContent-Transfer-Encoding: 7bit\r\n\r\n${body}\r\n--mix\r\nContent-Type: application/pdf; name="lease.pdf"\r\nContent-Transfer-Encoding: base64\r\nContent-Disposition: attachment; filename="lease.pdf"\r\n\r\n${Buffer.from('%PDF-1.4 the lease').toString('base64')}\r\n--mix--\r\n`,
    'latin1'
  );

test('RC2 is RFC 2268 RC2: its table is a permutation, and it gives the RFC test vectors', () => {
  assert.equal(new Set(RC2_PITABLE).size, 256);
  const vectors = [
    ['0000000000000000', 63, '0000000000000000', 'ebb773f993278eff'],
    ['ffffffffffffffff', 64, 'ffffffffffffffff', '278b27e42e2f0d49'],
    ['3000000000000000', 64, '1000000000000001', '30649edf9be7d2c2'],
    ['88', 64, '0000000000000000', '61a8a244adacccf0'],
    ['88bca90e90875a', 64, '0000000000000000', '6ccf4308974c267f'],
    ['88bca90e90875a7f0f79c384627bafb2', 64, '0000000000000000', '1a807d272bbe5db1'],
    ['88bca90e90875a7f0f79c384627bafb2', 128, '0000000000000000', '2269552ab0f85ca6'],
  ];
  for (const [k, bits, p, c] of vectors) {
    assert.equal(rc2Block(Buffer.from(k, 'hex'), bits, Buffer.from(p, 'hex')).toString('hex'), c, `${k} at ${bits} bits`);
    assert.equal(rc2Block(Buffer.from(k, 'hex'), bits, Buffer.from(c, 'hex'), true).toString('hex'), p);
  }
  // RC2-40-CBC as OpenSSL writes it (openssl enc -rc2-40-cbc -K 0123456789 -iv 0011223344556677).
  const key = Buffer.from('0123456789', 'hex');
  const iv = Buffer.from('0011223344556677', 'hex');
  const out = rc2Encrypt(Buffer.from('hello rc2 world, padded'), key, iv, 40);
  assert.equal(out.toString('hex'), 'd703a852a12c0742449de1a0c63f6cd284faedacf62da608');
  assert.equal(rc2Decrypt(out, key, iv, 40).toString(), 'hello rc2 world, padded');
});

test('DER reads back what it writes', () => {
  assert.equal(oidOf(parse(oid('1.2.840.113549.1.9.16.1.23'))), '1.2.840.113549.1.9.16.1.23');
  for (const n of [0n, 127n, 128n, 255n, 256n, 2n ** 64n + 5n]) assert.equal(intOf(parse(int(n))), n);
});

test('a .p12 opens with its password, as Windows writes it now (AES) and as it once did (Triple DES and RC2)', () => {
  for (const legacy of [false, true]) {
    const bytes = makePkcs12(people.amina, 'Correct horse 7', { legacy, chain: [people.ca.cert] });
    const got = readPkcs12(bytes, 'Correct horse 7');
    assert.equal(got.cert.email, 'amina@northwind.example', `legacy ${legacy}`);
    assert.equal(got.chain.length, 1);
    assert.equal(got.chain[0].name, 'Rutba Test Authority Root');
    const signed = crypto.sign('sha256', Buffer.from('x'), got.key);
    assert.ok(crypto.verify('sha256', Buffer.from('x'), people.amina.publicKey, signed), 'the key is the certificate\'s');
    assert.throws(() => readPkcs12(bytes, 'correct horse 7'), (err) => err instanceof Pkcs12Error && err.code === 'password');
  }
  assert.throws(() => readPkcs12(Buffer.from('not a certificate'), 'x'), (err) => err.code === 'not-pkcs12');
});

test('a signed message says who signed it, that the signature is valid and that it is unchanged', () => {
  const sent = protect(message(), { sign: signer(people.amina) });
  assert.equal(smimeKind(sent), 'signed');
  const r = readSmime(sent, { anchors });
  assert.equal(r.signed.status, 'valid');
  assert.equal(r.signed.signer.name, 'Amina Rahman');
  assert.equal(r.signed.signer.email, 'amina@northwind.example');
  assert.equal(r.signed.trustedBy, 'Rutba Test Authority Root');
  assert.equal(r.signed.fromMatches, true);
  assert.equal(r.signed.revocation, 'not-checked');
  assert.ok(Math.abs(new Date(r.signed.signedAt) - Date.now()) < 3_600_000, `signed just now, read back as ${r.signed.signedAt}`);
  assert.equal(r.encrypted, null);
  assert.match(r.message.text, /quartzberry clause stands/);
  assert.equal(r.message.subject, 'Harbour lease: signed copy');
  assert.deepEqual(r.message.attachments.map((a) => a.filename), ['lease.pdf'], 'the signature itself is not shown as an attachment');
  assert.ok(r.certificates.some((c) => c.email === 'amina@northwind.example'), 'the signer\'s certificate is kept, to encrypt to them later');

  // One character of the body changed after signing.
  const tampered = Buffer.from(sent.toString('latin1').replace('quartzberry', 'quartzbarry'), 'latin1');
  assert.equal(readSmime(tampered, { anchors }).signed.status, 'changed');
  // Stored with bare newlines on the way: still the same message.
  assert.equal(readSmime(Buffer.from(sent.toString('latin1').replace(/\r\n/g, '\n'), 'latin1'), { anchors }).signed.status, 'valid');
  // Written with a carriage return too many at each line's end (a program
  // writing CRLF in text mode on Windows), as OpenSSL there leaves it: valid,
  // and shown with its words.
  const doubled = readSmime(Buffer.from(sent.toString('latin1').replace(/\r\n/g, '\r\r\n'), 'latin1'), { anchors });
  assert.equal(doubled.signed.status, 'valid');
  assert.match(doubled.message.text, /quartzberry clause stands/);
});

test('a signature says so when no authority here vouches for it, the certificate is out of date, or it is for another address', () => {
  assert.equal(readSmime(protect(message(), { sign: signer(people.amina) }), { anchors: [] }).signed.status, 'untrusted');
  assert.equal(readSmime(protect(message({ from: people.stranger }), { sign: signer(people.stranger) }), { anchors }).signed.status, 'untrusted');
  const old = readSmime(protect(message({ from: people.expired }), { sign: signer(people.expired) }), { anchors });
  assert.equal(old.signed.status, 'expired');
  const other = readSmime(protect(message({ from: people.bo }), { sign: signer(people.amina) }), { anchors });
  assert.equal(other.signed.status, 'other-sender', 'a good signature by someone other than the sender is not shown as good');
  assert.equal(other.signed.fromMatches, false, 'signed by Amina, sent from Bo\'s address');
  const ec = readSmime(protect(message({ from: people.ecSigner }), { sign: signer(people.ecSigner) }), { anchors });
  assert.equal(ec.signed.status, 'valid', 'an elliptic-curve signature');
});

test('an opaque signed message (application/pkcs7-mime, signed-data) is read and checked too', () => {
  const entity = Buffer.from('Content-Type: text/plain; charset=utf-8\r\n\r\nOpaque quartzberry.\r\n', 'latin1');
  const cms = signDetached(entity, { key: people.amina.key, cert: people.amina.cert, encapsulate: true });
  const raw = Buffer.from(`From: Amina Rahman <amina@northwind.example>\r\nTo: owner@example.com\r\nSubject: Opaque\r\nMIME-Version: 1.0\r\nContent-Type: application/pkcs7-mime; smime-type=signed-data; name="smime.p7m"\r\nContent-Transfer-Encoding: base64\r\n\r\n${cms.toString('base64').replace(/.{76}/g, '$&\r\n')}\r\n`, 'latin1');
  assert.equal(smimeKind(raw), 'opaque-signed');
  const r = readSmime(raw, { anchors });
  assert.equal(r.signed.status, 'valid');
  assert.match(r.message.text, /Opaque quartzberry/);
});

test('an encrypted message opens with the certificate it was encrypted for, and only that one', () => {
  const sent = protect(message(), { encryptFor: [people.owner.cert, people.amina.cert] });
  assert.equal(smimeKind(sent), 'enveloped');
  assert.doesNotMatch(sent.toString('latin1'), /quartzberry/, 'nothing of the body is readable in what is sent');
  const r = readSmime(sent, { identities: [identity(people.owner)], anchors });
  assert.equal(r.encrypted.status, 'decrypted');
  assert.equal(r.encrypted.cipher, 'AES-256');
  assert.equal(r.encrypted.recipient.email, 'owner@example.com');
  assert.match(r.message.text, /quartzberry clause stands/);
  assert.deepEqual(r.message.attachments.map((a) => a.filename), ['lease.pdf']);
  // The sender's own copy opens with the sender's own certificate.
  assert.equal(readSmime(sent, { identities: [identity(people.amina)], anchors }).encrypted.status, 'decrypted');
  const nobody = readSmime(sent, { identities: [identity(people.bo)], anchors });
  assert.equal(nobody.encrypted.status, 'no-identity');
  assert.equal(nobody.message.text, null);
  assert.equal(nobody.message.subject, 'Harbour lease: signed copy', 'the headers are still shown');
  // Cut short on the way.
  const { headText, body } = splitEntity(sent);
  const cut = Buffer.concat([Buffer.from(`${headText}\r\n\r\n`, 'latin1'), Buffer.from(Buffer.from(body.toString('latin1'), 'base64').subarray(0, -7).toString('base64'), 'latin1')]);
  assert.ok(['damaged', 'failed'].includes(readSmime(cut, { identities: [identity(people.owner)], anchors }).encrypted.status));
});

test('a message signed and then encrypted opens and is checked, layer by layer', () => {
  const sent = protect(message(), { sign: signer(people.amina), encryptFor: [people.owner.cert, people.amina.cert] });
  const r = readSmime(sent, { identities: [identity(people.owner)], anchors });
  assert.equal(r.encrypted.status, 'decrypted');
  assert.equal(r.signed.status, 'valid');
  assert.equal(r.signed.signer.name, 'Amina Rahman');
  assert.match(r.message.text, /quartzberry/);
  // An encrypted message to someone else's certificate only: refused in a sentence, not a crash.
  assert.throws(() => protect(message(), { encryptFor: [people.ecSigner.cert] }), /not an RSA one/);
});

test('a message that is not S/MIME reads as it is', () => {
  const plain = Buffer.from('From: a@example.com\r\nSubject: Hi\r\n\r\nJust words.\r\n', 'latin1');
  const r = readSmime(plain, { anchors });
  assert.equal(r.signed, null);
  assert.equal(r.encrypted, null);
  assert.match(r.message.text, /Just words/);
});

test('an Outlook data file\'s smime.p7m attachment gives back the message it holds', () => {
  const sent = protect(message(), { sign: signer(people.amina) });
  const { headText, body } = splitEntity(sent);
  const contentType = headText.split(/\r\n(?=\S)/).find((l) => /^content-type:/i.test(l));
  // As Outlook keeps it: the whole entity, its Content-Type first ...
  const whole = entityFromAttachment({ filename: 'smime.p7m', bytes: Buffer.concat([Buffer.from(`${contentType}\r\n\r\n`, 'latin1'), body]) });
  assert.equal(readSmime(Buffer.concat([Buffer.from('From: Amina Rahman <amina@northwind.example>\r\n', 'latin1'), whole]), { anchors }).signed.status, 'valid');
  // ... or the body alone, its boundary on its first line.
  const bare = entityFromAttachment({ filename: 'smime.p7m', bytes: body.subarray(body.indexOf('--')) });
  assert.equal(readSmime(bare, { anchors }).signed.status, 'no-from', 'with no From to compare, it is said so');
  assert.equal(entityFromAttachment({ filename: 'report.pdf', bytes: Buffer.from('x') }), null);
});
