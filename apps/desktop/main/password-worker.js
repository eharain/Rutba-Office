// A thread that works out password hashes, so the document service's own
// thread goes on answering every window while 100,000 rounds of SHA-512 are
// spun (password-host.js). The same functions the engines call, from the
// same modules: what comes back is byte for byte what the engine would
// have worked out itself.

import { parentPort } from 'node:worker_threads';
import { hashPassword } from '@rutba/ooxml/protection';
import { spinPassword, decryptPackage } from '@rutba/office-formats/crypt';

/** A Buffer over a copy's bytes, as the engines were written to take. */
const buffer = (bytes) => Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);

/** Only the bytes of a view, copied out, so the copy can be handed over whole. */
function own(bytes) {
  const copy = new Uint8Array(bytes.byteLength);
  copy.set(bytes);
  return copy;
}

const JOBS = {
  hash: ({ password, salt, spinCount, algorithm }) => ({ value: hashPassword(password, { salt, spinCount, algorithm }) }),
  spin: ({ password, salt, spinCount, alg }) => {
    const value = own(spinPassword(password, buffer(salt), spinCount, alg));
    return { value, transfer: [value.buffer] };
  },
  decrypt: ({ bytes, password }) => {
    const out = decryptPackage(buffer(bytes), password);
    const plain = own(out.bytes);
    return { value: { bytes: plain, kind: out.kind, integrity: out.integrity }, transfer: [plain.buffer] };
  },
};

parentPort.on('message', ({ id, job, args }) => {
  try {
    const fn = JOBS[job];
    if (!fn) throw new Error(`the password thread has no ${job}`);
    const { value, transfer = [] } = fn(args);
    parentPort.postMessage({ id, ok: true, value }, transfer);
  } catch (err) {
    parentPort.postMessage({ id, ok: false, error: { message: err?.message || String(err), name: err?.name || 'Error', code: err?.code } });
  }
});
