// A recognition thread (host.js starts them). It starts the engine on its
// first job and keeps it, with the last languages it loaded, until the host
// stops the thread.

import { parentPort, workerData, threadId } from 'node:worker_threads';
import { createEngine } from './engine.js';
import { recognizePage } from './page.js';

let engine = null;

parentPort.on('message', async ({ id, args }) => {
  try {
    engine ||= await createEngine({ vendor: workerData?.vendor || undefined });
    const value = { ...recognizePage(engine, args), thread: threadId };
    parentPort.postMessage({ id, ok: true, value });
  } catch (err) {
    parentPort.postMessage({ id, ok: false, error: { message: err?.message || String(err), name: err?.name, code: err?.code } });
  }
});
