// The document service on its own thread (apps/desktop/main/doc-host.js and
// doc-worker.js): the windows' calls answered with promises, the checks'
// calls answered as function calls, a picture minted on the thread and held
// on the main process for the window that asked before the answer naming
// it arrives, an equation laid out by the main process for a method that
// was asynchronous already, faults said in sentences, and a thread that
// stops started again.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createDocumentHost } from '../apps/desktop/main/doc-host.js';

test('a thread that does not get ready in time leaves the service to the main process, which still opens documents', async (t) => {
  const host = createDocumentHost({ holdBlob: (b, ty, n, o) => ({ id: o?.id || 'b1', url: 'rutba://blob/b1' }), releaseBlob: () => {}, blobUrl: (id) => `rutba://blob/${id}`, readyLimitMs: 1 }).start();
  t.after(() => host.stop());
  host.provide({ measureMath: async () => new Map(), setting: () => {}, teach: () => {} });
  host.init({ recoveryDir: fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-host-')), locale: 'en-GB', settings: {} });
  const { doc, direct } = await host.ready();
  assert.equal(host.threaded, false, 'on the main process');
  const made = await doc.new({ kind: 'doc' });
  assert.ok(made.id);
  assert.ok(direct.sessions().some((s) => s.id === made.id), 'both faces reach the same service');
});

test('the document service answers from its own thread, as a promise and as a function call', async (t) => {
  const recoveryDir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-host-'));
  const held = [];
  const released = [];
  const asked = [];
  const settings = [];
  const host = createDocumentHost({
    holdBlob: (bytes, type, name, options) => { held.push({ bytes, type, name, options }); return { id: options.id }; },
    releaseBlob: (id) => released.push(id),
    blobUrl: (id) => `rutba://blob/${id}`,
  }).start();
  t.after(() => host.stop());
  host.provide({
    measureMath: async (list) => { asked.push(list); return new Map([['deck-eq', { widthEm: 4, heightEm: 1.5, png: null }]]); },
    setting: (key, value) => settings.push([key, value]),
    teach: () => {},
  });
  host.init({ recoveryDir, locale: 'en-GB', settings: {} });
  const { doc, direct } = await host.ready();
  assert.equal(host.threaded, true, 'on its own thread');
  assert.equal(typeof doc.apply, 'function');
  assert.equal(typeof direct.model, 'function');

  // A window's call: a promise.
  const made = await doc.new({ kind: 'doc' }, { id: 7 });
  assert.ok(made.id && made.model, 'a new document, its model with it');
  // A check's call: the answer itself, on the next line.
  const typed = direct.apply({ id: made.id, ops: [{ op: 'setSelection', anchor: { block: 0, offset: 0 }, focus: { block: 0, offset: 0 } }, { op: 'insertText', text: 'Hello from the thread' }] });
  assert.equal(typeof typed?.then, 'undefined', 'not a promise');
  assert.match(JSON.stringify(direct.model({ id: made.id })), /Hello from the thread/);
  assert.ok(direct.sessions().some((s) => s.id === made.id));

  // A picture handed to window 7: held on the main side, under the thread's own id, before the answer.
  const asset = direct.asset({ id: made.id, ref: 'word/document.xml' }, { id: 7 });
  assert.match(asset.url, /^rutba:\/\/blob\/w1-\d+$/);
  const hold = held.find((h) => h.options.id === asset.id);
  assert.ok(hold, 'held before the answer naming it came back');
  assert.equal(hold.options.owner, 7, 'owned by the window that asked');
  assert.equal(hold.bytes.byteLength, asset.size);
  assert.match(Buffer.from(hold.bytes).toString('utf8'), /Hello from the thread/);

  // Asynchronous in the service: the check is handed a promise, and the
  // thread's question to the main process is answered while it waits.
  const deck = await doc.new({ kind: 'deck' }, { id: 7 });
  const pending = direct.deckEquation({ id: deck.id, linear: 'a^2+b^2=c^2' });
  assert.equal(typeof pending.then, 'function', 'a promise, as from the service itself');
  const shape = await pending;
  assert.ok(shape, 'the equation placed');
  assert.equal(asked.length, 1, 'laid out by the main process');
  assert.match(asked[0][0].mathml, /<math/);

  // Faults in sentences, both ways.
  assert.throws(() => direct.model({ id: 'nope' }), /./);
  await assert.rejects(doc.model({ id: 'nope' }), /./);

  // A closed window frees what it opened.
  const gone = await doc.closeWindow(7);
  assert.deepEqual(gone.sort(), [made.id, deck.id].sort());

  // The thread stops: what was waiting on it is told, and it is started again.
  const again = await doc.new({ kind: 'sheet' }, { id: 8 });
  const before = host.threadId;
  const inFlight = doc.sessions();
  await host.crashForCheck();
  await inFlight.catch(() => null);
  const after = await host.ready();
  assert.notEqual(host.threadId, before, 'a new thread');
  await assert.rejects(after.doc.model({ id: again.id }), /./, 'the documents it held are gone');
  const fresh = await after.doc.new({ kind: 'doc' }, { id: 9 });
  assert.ok(fresh.id, 'and new ones open');
  const freshAsset = after.direct.asset({ id: fresh.id, ref: 'word/document.xml' }, { id: 9 });
  assert.match(freshAsset.url, /^rutba:\/\/blob\/w2-\d+$/, 'its pictures named apart from the first thread\'s');
});
