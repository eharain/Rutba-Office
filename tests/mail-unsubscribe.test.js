// One-click unsubscribe (RFC 8058, A6.5): the one request the dialog says it
// sends, and nothing else. A POST of "List-Unsubscribe=One-Click" to the
// sender's https address, with no cookies, no sign-in and no redirect
// followed, made from Mail itself rather than by opening the browser.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import https from 'node:https';
import { createMailService } from '../apps/desktop/main/mail.js';
import { unsubscribeOffer } from '../apps/desktop/main/mail-insight.js';
import { imapServer, certificate } from '../packages/mailbox/test/imap-server.js';

function fakeStores() {
  const settings = new Map();
  const secrets = new Map();
  return {
    settings: { get: (k, fallback) => (settings.has(k) ? structuredClone(settings.get(k)) : fallback), set: (k, v) => settings.set(k, structuredClone(v)), delete: (k) => settings.delete(k), all: () => Object.fromEntries(settings) },
    secrets: { available: () => true, get: (k) => secrets.get(k) ?? null, set: (k, v) => secrets.set(k, v), delete: (k) => secrets.delete(k), keys: () => [...secrets.keys()] },
  };
}

/** The sender's unsubscribe address: every request it is sent, and how it answers. */
async function listServer() {
  const heard = [];
  let answer = { status: 200 };
  const server = https.createServer(certificate(), (req, res) => {
    let body = '';
    req.on('data', (c) => { body += c; });
    req.on('end', () => {
      heard.push({ method: req.method, url: req.url, headers: req.headers, body });
      res.writeHead(answer.status, answer.location ? { Location: answer.location } : {});
      res.end();
    });
  });
  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
  return { heard, port: server.address().port, answer: (a) => { answer = a; }, close: () => new Promise((resolve) => server.close(resolve)) };
}

const letter = (n, unsubscribe, post = 'List-Unsubscribe=One-Click') => [
  'From: Offers <news@shop.example>', 'To: Sam <sam@checks.example>', `Subject: Offer ${n}`,
  `Date: ${new Date().toUTCString()}`, `Message-ID: <offer${n}@shop.example>`, 'List-Id: Shop offers <offers.shop.example>',
  `List-Unsubscribe: ${unsubscribe}`, ...(post ? [`List-Unsubscribe-Post: ${post}`] : []), '', `Offer ${n}.`, '',
].join('\r\n');

test('one-click unsubscribe sends the one request it says it sends, and says when the sender did not take it', async () => {
  const list = await listServer();
  const imap = await imapServer();
  const base = `https://127.0.0.1:${list.port}`;
  imap.deliver('INBOX', letter(1, `<${base}/u/one?token=abc>, <mailto:leave@shop.example?subject=leave>`));
  imap.deliver('INBOX', letter(2, `<${base}/u/two>`));
  imap.deliver('INBOX', letter(3, `<${base}/u/three>`));
  imap.deliver('INBOX', letter(4, `<${base}/u/four>`, null));
  imap.deliver('INBOX', letter(5, `<http://127.0.0.1:${list.port}/u/five>`));
  const service = createMailService({
    stores: fakeStores(),
    holdBlob: () => ({ url: '' }),
    broadcast: () => {},
    userData: fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-unsub-')),
    smtp: { async send() { return { messageId: '<x@here>' }; } },
    // The test's own certificate is trusted here alone.
    unsubscribeAgent: new https.Agent({ rejectUnauthorized: false }),
  });
  try {
    const { id: accountId } = service.addAccount({ account: imap.account(), password: 'check' });
    await service.sync({ accountId });
    const idOf = (n) => service.messages({ accountId, folder: 'INBOX', limit: 50 }).rows.find((r) => r.subject === `Offer ${n}`).id;

    const done = await service.unsubscribe({ accountId, folder: 'INBOX', id: idOf(1) });
    assert.deepEqual(done, { done: true, status: 200 });
    assert.equal(list.heard.length, 1, 'one request');
    const [req] = list.heard;
    assert.equal(req.method, 'POST');
    assert.equal(req.url, '/u/one?token=abc');
    assert.equal(req.headers['content-type'], 'application/x-www-form-urlencoded');
    assert.equal(req.body, 'List-Unsubscribe=One-Click');
    assert.equal(req.headers.cookie, undefined, 'no cookies');
    assert.equal(req.headers.authorization, undefined, 'no sign-in');

    // Refused by the sender: said, not taken for done.
    list.answer({ status: 500 });
    assert.deepEqual(await service.unsubscribe({ accountId, folder: 'INBOX', id: idOf(2) }), { done: false, reason: 'refused', status: 500 });
    // A redirect is not followed: a POST sent on is not the one request it said.
    list.answer({ status: 302, location: `${base}/elsewhere` });
    assert.deepEqual(await service.unsubscribe({ accountId, folder: 'INBOX', id: idOf(3) }), { done: false, reason: 'refused', status: 302 });
    assert.equal(list.heard.length, 3);
    assert.ok(!list.heard.some((h) => h.url === '/elsewhere'));

    // Without List-Unsubscribe-Post there is no one-click, and nothing is sent.
    assert.deepEqual(await service.unsubscribe({ accountId, folder: 'INBOX', id: idOf(4) }), { done: false, reason: 'not-offered' });
    assert.equal(list.heard.length, 3);

    // Only over https: a plain http address is opened in the browser, as any web page is.
    const plain = service.insight({ accountId, folder: 'INBOX', id: idOf(5) }).unsubscribe;
    assert.equal(plain.method, 'web');
    assert.deepEqual(await service.unsubscribe({ accountId, folder: 'INBOX', id: idOf(5) }), { done: false, reason: 'not-offered' });
    assert.equal(list.heard.length, 3);

    // The sender gone: said, in a few seconds.
    await list.close();
    const away = await service.unsubscribe({ accountId, folder: 'INBOX', id: idOf(1) });
    assert.equal(away.done, false);
    assert.equal(away.reason, 'unreachable');
  } finally {
    await imap.close();
    await list.close().catch(() => {});
  }
});

test('the offer names one-click only for an https address with List-Unsubscribe-Post', () => {
  const headers = (unsub, post) => ({ headers: [{ key: 'list-unsubscribe', value: unsub }, ...(post ? [{ key: 'list-unsubscribe-post', value: post }] : [])] });
  assert.equal(unsubscribeOffer(headers('<https://a.example/u>', 'List-Unsubscribe=One-Click')).method, 'one-click');
  assert.equal(unsubscribeOffer(headers('<http://a.example/u>', 'List-Unsubscribe=One-Click')).method, 'web');
  assert.equal(unsubscribeOffer(headers('<https://a.example/u>', null)).method, 'web');
  assert.equal(unsubscribeOffer(headers('<mailto:x@a.example>', 'List-Unsubscribe=One-Click')).method, 'email');
});
