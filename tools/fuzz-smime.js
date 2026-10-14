// The S/MIME reader against damaged and made-up messages, as tools/fuzz-pdf.js
// holds the PDF reader: node tools/fuzz-smime.js [messages] [seed]
//
// A signed or encrypted message comes from a stranger, and reading it means
// parsing structures they wrote. Each message here is read on a worker with a
// two-second limit and a 512 MB heap, as the Mail window's S/MIME thread reads
// one: signed (multipart/signed and opaque), encrypted, signed then encrypted,
// encrypted with Triple DES, and random bytes behind S/MIME headers, damaged
// in the message's own bytes or inside its CMS structure (a byte flipped, a
// length changed, a run cut out or repeated, noise spliced in). A result, or a
// refusal in a sentence, is fine; a fault, a hang past the limit or a thread
// that runs out of memory is not.

import fs from 'node:fs';
import os from 'node:os';
import crypto from 'node:crypto';
import { Worker, isMainThread, parentPort, workerData } from 'node:worker_threads';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const LIMIT_MS = 2000;

// fuzz-open.js's judgement, copied rather than imported, as fuzz-pdf.js does.
const FAULT_NAMES = new Set(['RangeError', 'TypeError', 'ReferenceError', 'EvalError', 'InternalError', 'AggregateError']);
const FAULT_WORDS = /Maximum call stack|Invalid array length|Array buffer allocation|out of memory|outside the bounds|Offset is outside|is not a function|is not iterable|Cannot read propert|Cannot set propert|undefined is not|null is not/i;
function judge(err) {
  const links = [];
  for (let e = err, n = 0; e && n < 8; e = e.cause, n++) links.push({ name: e?.name || typeof e, message: String(e?.message ?? e) });
  const fault = links.find((l) => FAULT_NAMES.has(l.name) || FAULT_WORDS.test(l.message));
  const top = links[0]?.message || '';
  return fault || !/[a-z]{3}.*[ .]/.test(top)
    ? { outcome: 'FAULT', message: `${(fault || links[0]).name}: ${(fault || links[0]).message}`.slice(0, 200) }
    : { outcome: 'refused', message: top.slice(0, 200) };
}

if (!isMainThread) {
  const { readSmime, certificate } = await import('@rutba/mailbox/smime');
  const identities = [{ cert: certificate(workerData.ownerCert), key: crypto.createPrivateKey(workerData.ownerKey) }];
  const anchors = [certificate(workerData.caCert)];
  parentPort.on('message', ({ bytes }) => {
    const t = performance.now();
    try {
      const r = readSmime(Buffer.from(bytes), { identities, anchors });
      parentPort.postMessage({ outcome: 'read', status: r.error || r.encrypted?.status || r.signed?.status || 'plain', ms: performance.now() - t });
    } catch (err) {
      parentPort.postMessage({ ...judge(err), ms: performance.now() - t });
    }
  });
} else if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const { protect, certificate } = await import('@rutba/mailbox/smime');
  const { signDetached } = await import('../packages/mailbox/src/smime/cms.js');
  const { seq, set, int, oid, octets, explicit, algorithm, tlv } = await import('../packages/mailbox/src/smime/der.js');
  const { OID } = await import('../packages/mailbox/src/smime/oids.js');
  const { makeIdentities } = await import('../tests/fixtures/smime-keys.js');

  const total = Number(process.argv[2] || 2000);
  let state = (Number(process.argv[3] || Date.now() % 100000) >>> 0) || 1;
  const seed = state;
  const rand = () => { state ^= state << 13; state ^= state >>> 17; state ^= state << 5; return ((state >>> 0) % 1000000) / 1000000; };
  const pick = (n) => Math.floor(rand() * n);

  const people = makeIdentities();
  const signer = (p) => ({ key: p.key, cert: p.cert, chain: [people.ca.cert] });
  const message = (from) =>
    Buffer.from(`From: ${from.name} <${from.email}>\r\nTo: Archive Owner <owner@example.com>\r\nSubject: Lease\r\nMIME-Version: 1.0\r\nContent-Type: multipart/mixed; boundary="mix"\r\n\r\n--mix\r\nContent-Type: text/plain; charset=utf-8\r\n\r\nThe quartzberry clause stands.\r\n--mix\r\nContent-Type: application/pdf; name="lease.pdf"\r\nContent-Transfer-Encoding: base64\r\n\r\n${Buffer.from('%PDF-1.4 the lease').toString('base64')}\r\n--mix--\r\n`, 'latin1');
  const b64 = (b) => b.toString('base64').replace(/.{76}/g, '$&\r\n');

  // Each source: a message, and where its CMS structure is (to damage inside it and put it back).
  const asOpaque = (type) => (cms) => Buffer.from(`From: Amina Rahman <amina@northwind.example>\r\nTo: owner@example.com\r\nSubject: x\r\nMIME-Version: 1.0\r\nContent-Type: application/pkcs7-mime; smime-type=${type}; name="smime.p7m"\r\nContent-Transfer-Encoding: base64\r\n\r\n${b64(cms)}\r\n`, 'latin1');
  const asDetached = (content) => (cms) => Buffer.from(`From: Amina Rahman <amina@northwind.example>\r\nTo: owner@example.com\r\nSubject: x\r\nMIME-Version: 1.0\r\nContent-Type: multipart/signed; protocol="application/pkcs7-signature"; micalg=sha-256; boundary="sig"\r\n\r\n--sig\r\n${content}\r\n--sig\r\nContent-Type: application/pkcs7-signature; name="smime.p7s"\r\nContent-Transfer-Encoding: base64\r\n\r\n${b64(cms)}\r\n--sig--\r\n`, 'latin1');
  const cmsOf = (raw) => {
    const text = raw.toString('latin1');
    const at = text.lastIndexOf('base64\r\n');
    const start = text.indexOf('\r\n\r\n', at) + 4;
    const end = text.indexOf('\r\n--', start) > 0 ? text.indexOf('\r\n--', start) : text.length;
    return Buffer.from(text.slice(start, end).replace(/\s+/g, ''), 'base64');
  };
  const inner = 'Content-Type: text/plain; charset=utf-8\r\n\r\nThe quartzberry clause stands.\r\n';
  const owner = certificate(people.owner.cert);
  const tripleDes = (() => {
    const cek = crypto.randomBytes(24);
    const iv = crypto.randomBytes(8);
    const c = crypto.createCipheriv('des-ede3-cbc', cek, iv);
    const body = Buffer.concat([c.update(Buffer.from(inner, 'latin1')), c.final()]);
    const wrapped = crypto.publicEncrypt({ key: owner.x509.publicKey, padding: crypto.constants.RSA_PKCS1_PADDING }, cek);
    const ktri = seq(int(0), seq(owner.issuerRaw, tlv(0x02, owner.serialRaw)), algorithm(OID.rsaEncryption), octets(wrapped));
    return seq(oid(OID.envelopedData), explicit(0, seq(int(0), set(ktri), seq(oid(OID.data), algorithm(OID.desEde3CBC, octets(iv)), tlv(0x80, body)))));
  })();
  const SOURCES = {
    'signed (multipart/signed)': { build: asDetached(inner), cms: signDetached(Buffer.from(inner, 'latin1'), { key: people.amina.key, cert: people.amina.cert, chain: [people.ca.cert] }) },
    'signed (opaque)': { build: asOpaque('signed-data'), cms: signDetached(Buffer.from(inner, 'latin1'), { key: people.amina.key, cert: people.amina.cert, chain: [people.ca.cert], encapsulate: true }) },
    'signed with SHA-1': { build: asOpaque('signed-data'), cms: signDetached(Buffer.from(inner, 'latin1'), { key: people.amina.key, cert: people.amina.cert, digest: 'sha1', encapsulate: true }) },
    'signed by a curve key': { build: asOpaque('signed-data'), cms: signDetached(Buffer.from(inner, 'latin1'), { key: people.ecSigner.key, cert: people.ecSigner.cert, encapsulate: true }) },
    'encrypted (AES)': { build: asOpaque('enveloped-data'), cms: cmsOf(protect(message(people.amina), { encryptFor: [people.owner.cert] })) },
    'signed then encrypted': { build: asOpaque('enveloped-data'), cms: cmsOf(protect(message(people.amina), { sign: signer(people.amina), encryptFor: [people.owner.cert, people.amina.cert] })) },
    'encrypted (Triple DES)': { build: asOpaque('enveloped-data'), cms: tripleDes },
  };
  // Whole-message sources, damaged as they are.
  const WHOLE = {
    'signed, as sent': protect(message(people.amina), { sign: signer(people.amina) }),
    'signed then encrypted, as sent': protect(message(people.amina), { sign: signer(people.amina), encryptFor: [people.owner.cert] }),
  };

  const noise = (n) => { const o = Buffer.alloc(n); for (let i = 0; i < n; i++) o[i] = pick(256); return o; };
  const DAMAGE = {
    'flip a byte': (b) => { const o = Buffer.from(b); o[pick(o.length)] ^= 1 << pick(8); return o; },
    'flip several bytes': (b) => { const o = Buffer.from(b); for (let i = 1 + pick(8); i > 0; i--) o[pick(o.length)] ^= 1 << pick(8); return o; },
    'a length changed': (b) => {
      // DER lengths sit after a tag: change a byte just after one of the common tags.
      const o = Buffer.from(b);
      const tags = [];
      for (let i = 0; i < o.length - 1; i++) if ([0x30, 0x31, 0x04, 0x02, 0x06, 0xa0, 0x80].includes(o[i])) tags.push(i + 1);
      if (tags.length) o[tags[pick(tags.length)]] = [0x00, 0x7f, 0x80, 0x81, 0x84, 0x88, 0xff, pick(256)][pick(8)];
      return o;
    },
    'indefinite lengths': (b) => { const o = Buffer.from(b); for (let i = 0; i < o.length - 1; i++) if (o[i] === 0x30 && rand() < 0.3) o[i + 1] = 0x80; return o; },
    'zero a run': (b) => { const o = Buffer.from(b); const at = pick(o.length); o.fill(0, at, Math.min(o.length, at + 1 + pick(64))); return o; },
    'cut the tail': (b) => Buffer.from(b).subarray(0, 1 + pick(b.length)),
    'cut a chunk out': (b) => { const at = pick(b.length); const to = Math.min(b.length, at + 1 + pick(256)); return Buffer.concat([b.subarray(0, at), b.subarray(to)]); },
    'repeat a chunk': (b) => { const at = pick(b.length); const chunk = b.subarray(at, Math.min(b.length, at + 1 + pick(64))); return Buffer.concat([b.subarray(0, at), ...Array(1 + pick(200)).fill(chunk), b.subarray(at)]); },
    'splice in noise': (b) => { const at = pick(b.length); return Buffer.concat([b.subarray(0, at), noise(1 + pick(128)), b.subarray(at)]); },
    'carriage returns': (b) => { const at = pick(b.length); return Buffer.concat([b.subarray(0, at), Buffer.alloc(1 + pick(200_000), 13), b.subarray(at)]); },
  };
  const names = Object.keys(DAMAGE);
  const randomMessage = () => {
    const types = ['signed-data', 'enveloped-data', 'authenveloped-data'];
    if (rand() < 0.5) return asOpaque(types[pick(types.length)])(noise(1 + pick(4096)));
    // A random DER-looking structure: tags and lengths with random content.
    const parts = [];
    for (let i = 0, n = 1 + pick(40); i < n; i++) parts.push(Buffer.from([[0x30, 0x31, 0x02, 0x04, 0x06, 0xa0][pick(6)], pick(256)]), noise(pick(32)));
    return rand() < 0.5 ? asOpaque(types[pick(types.length)])(Buffer.concat(parts)) : asDetached(inner)(Buffer.concat(parts));
  };

  const ownerKey = people.owner.key.export({ type: 'pkcs8', format: 'pem' });
  const data = { ownerCert: people.owner.cert, ownerKey, caCert: people.ca.cert };
  let worker = null;
  const start = () => {
    worker = new Worker(fileURLToPath(import.meta.url), { workerData: data, resourceLimits: { maxOldGenerationSizeMb: 512, maxYoungGenerationSizeMb: 64 } });
    worker.unref();
  };
  const run = (bytes) => new Promise((resolve) => {
    if (!worker) start();
    const w = worker;
    const done = (m) => { clearTimeout(timer); w.off('message', done); w.off('error', failed); resolve(m); };
    const failed = (err) => { worker = null; done({ outcome: 'FAULT', message: String(err?.message || err).slice(0, 200), ms: LIMIT_MS }); };
    // A fresh thread first loads the reader: its first message gets that time too (the read itself is timed inside).
    const grace = w.warm ? 0 : 3000;
    w.warm = true;
    const timer = setTimeout(() => { worker = null; w.terminate(); done({ outcome: 'FAULT', message: `did not finish in ${LIMIT_MS / 1000} seconds`, ms: LIMIT_MS }); }, LIMIT_MS + grace);
    w.on('message', done);
    w.on('error', failed);
    w.postMessage({ bytes });
  });

  const sources = Object.keys(SOURCES);
  const wholes = Object.keys(WHOLE);
  console.log(`fuzz-smime: ${total} messages from ${sources.length + wholes.length} sources and random bytes, seed ${seed}, ${LIMIT_MS / 1000} s each`);
  const counts = { read: 0, refused: 0, FAULT: 0 };
  const statuses = {};
  const faults = [];
  let slowest = { ms: 0 };
  for (let i = 0; i < total; i++) {
    let bytes;
    let what;
    const roll = rand();
    if (roll < 0.1) {
      bytes = randomMessage();
      what = 'random';
    } else if (roll < 0.3) {
      const source = wholes[pick(wholes.length)];
      const how = names[pick(names.length)];
      bytes = DAMAGE[how](WHOLE[source]);
      what = `${source}, ${how} in the message`;
    } else {
      const source = sources[pick(sources.length)];
      const how = names[pick(names.length)];
      bytes = SOURCES[source].build(DAMAGE[how](SOURCES[source].cms));
      what = `${source}, ${how} in the CMS`;
    }
    const r = await run(bytes);
    if (r.outcome !== 'FAULT' && r.ms > LIMIT_MS) Object.assign(r, { outcome: 'FAULT', message: `took ${r.ms.toFixed(0)} ms` });
    counts[r.outcome] = (counts[r.outcome] || 0) + 1;
    if (r.status) statuses[r.status] = (statuses[r.status] || 0) + 1;
    if (r.ms > slowest.ms) slowest = { ms: r.ms, what, size: bytes.length };
    if (r.outcome === 'FAULT') faults.push({ what, message: r.message, bytes });
  }
  await worker?.terminate();
  console.log(`${total} messages: read ${counts.read}, refused in a sentence ${counts.refused}, faults ${counts.FAULT}`);
  console.log(`what they read as: ${Object.entries(statuses).sort((a, b) => b[1] - a[1]).map(([k, v]) => `${k} ${v}`).join(', ')}`);
  console.log(`slowest: ${slowest.ms.toFixed(0)} ms (${slowest.what}, ${slowest.size} bytes)`);
  if (faults.length) {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-fuzz-smime-'));
    faults.forEach((f, i) => fs.writeFileSync(path.join(dir, `${i}.eml`), f.bytes));
    console.log(`faults (messages kept in ${dir}):`);
    for (const f of faults.slice(0, 20)) console.log(`  ${f.what}: ${f.message}`);
    process.exit(1);
  }
  console.log('every message was read or refused in a sentence, each within its time.');
}
