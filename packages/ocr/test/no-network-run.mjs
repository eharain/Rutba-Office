// Run by ocr.test.js under --require no-network.cjs: a recognition in a
// process where every way to the network is refused, saying whether the
// guard reached the threads and how many requests were tried.

import fs from 'node:fs';
import { Worker } from 'node:worker_threads';
import { createOcrHost } from '@rutba/ocr';

const guarded = await new Promise((resolve) => {
  const w = new Worker('process.getBuiltinModule("node:worker_threads").parentPort.postMessage(globalThis.__networkGuard === true)', { eval: true });
  w.once('message', (m) => { resolve(m); w.terminate(); });
});
const host = createOcrHost({ threads: 1 });
const r = await host.recognize(fs.readFileSync(process.argv[2]), { languages: ['urd', 'eng'] });
const s = host.stats;
await host.stop();
console.log(JSON.stringify({ guarded, inMain: globalThis.__networkGuard === true, attempts: globalThis.__networkAttempts.length, words: r.words.length, failed: s.failed }));
