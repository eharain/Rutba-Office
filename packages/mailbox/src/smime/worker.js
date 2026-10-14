// S/MIME read on a thread of its own (host.js starts it).
//
// A signed or encrypted message comes from a stranger, and reading it means
// parsing structures they wrote: so it is never done on the main process,
// which answers every window. This thread reads one message (or opens one
// .p12) at a time; the host gives it a time limit and a memory limit and
// starts another if it runs over. What goes back is plain data: the message
// inside, what was found about it, and certificates as their DER bytes.

import crypto from 'node:crypto';
import { parentPort } from 'node:worker_threads';
import { readSmime } from './index.js';
import { readPkcs12 } from './pkcs12.js';
import { certificate } from './cert.js';
import { systemAnchors } from './trust.js';

const der = (c) => Buffer.from(c.der);

const OPS = {
  read({ raw, identities = [], certificates = [], extraAnchors = [], anchors = null, now = null }) {
    const ids = identities.map((i) => ({ cert: certificate(Buffer.from(i.cert)), key: crypto.createPrivateKey(i.key) }));
    const pool = certificates.map((c) => certificate(Buffer.from(c)));
    const trusted = anchors ? anchors.map((c) => certificate(Buffer.from(c))) : [...systemAnchors(), ...extraAnchors.map((c) => certificate(Buffer.from(c)))];
    const r = readSmime(Buffer.from(raw), { identities: ids, certificates: pool, anchors: trusted, now: now ? new Date(now) : new Date() });
    return { message: r.message, signed: r.signed, encrypted: r.encrypted, error: r.error || null, certificates: r.certificates.map(der) };
  },
  pkcs12({ bytes, password }) {
    const got = readPkcs12(Buffer.from(bytes), password);
    return { key: got.key.export({ type: 'pkcs8', format: 'pem' }), cert: der(got.cert), chain: got.chain.map(der) };
  },
};

parentPort.on('message', ({ seq, op, args }) => {
  try {
    const value = OPS[op](args || {});
    parentPort.postMessage({ seq, ok: true, value });
  } catch (err) {
    parentPort.postMessage({ seq, ok: false, error: { message: err?.message || String(err), name: err?.name, code: err?.code } });
  }
});
