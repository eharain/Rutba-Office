// The mail server's own automatic reply over ManageSieve (D-2026-10-22-01):
// the Sieve script Office writes, and turning it on and off on a stand-in
// server that keeps one active script, as a real one does.

import test from 'node:test';
import assert from 'node:assert/strict';
import { openSieve, turnOn, turnOff, readState, vacationScript, readScript, sieveString, SCRIPT_NAME } from '@rutba/mailbox/managesieve';
import { standInSieve } from '@rutba/mailbox/managesieve-standin';

test('the script: the reply with its words, its window in UTC, and the filters that were active run after it', () => {
  const s = vacationScript({ subject: 'Away "now"', message: 'Back on Monday.\nThanks.', start: '2026-10-22T09:00:00.000Z', end: '2026-10-29T17:30:00.000Z' }, { addresses: ['amira@example.com'], previous: 'my-filters' });
  assert.match(s, /^require \["vacation", "date", "relational", "include"\];$/m);
  assert.match(s, /if allof\(currentdate :zone "\+0000" :value "ge" "iso8601" "2026-10-22T09:00:00", currentdate :zone "\+0000" :value "le" "iso8601" "2026-10-29T17:30:00"\) \{/);
  assert.match(s, /vacation :days 4 :subject "Away \\"now\\"" :addresses \["amira@example.com"\] "Back on Monday\.\r\nThanks\.";/);
  assert.match(s, /include :personal "my-filters";\r\n$/);
  assert.deepEqual(readScript(s), { subject: 'Away "now"', message: 'Back on Monday.\nThanks.', start: '2026-10-22T09:00:00.000Z', end: '2026-10-29T17:30:00.000Z', previous: 'my-filters' });
  // No window, no filters before it: the reply alone, needing nothing more.
  // The subject left blank: "Automatic reply" alone where the server cannot echo the original (A6.11).
  const plain = vacationScript({ message: 'Away.' });
  assert.match(plain, /^require \["vacation"\];$/m);
  assert.match(plain, /^vacation :days 4 :subject "Automatic reply" "Away\.";$/m);
  // And "Automatic reply: " with the subject answered, where it can (Sieve's variables).
  const echoed = vacationScript({ message: 'Away.' }, { variables: true });
  assert.match(echoed, /^require \["vacation", "variables"\];$/m);
  assert.match(echoed, /^if header :matches "subject" "\*" \{ set "subject" "\$\{1\}"; \}$/m);
  assert.match(echoed, /^vacation :days 4 :subject "Automatic reply: \$\{subject\}" "Away\.";$/m);
  // A subject of one's own keeps its place for the original.
  assert.match(vacationScript({ subject: 'Away: <original subject>', message: 'x' }, { variables: true }), /:subject "Away: \$\{subject\}"/);
  assert.doesNotMatch(vacationScript({ subject: 'Away now', message: 'x' }, { variables: true }), /variables/);
  assert.equal(readScript('require "fileinto";'), null);
  assert.equal(sieveString('a\\b"c'), '"a\\\\b\\"c"');
});

test('on and off at the server, the person\'s own filters kept and given back', async () => {
  const server = standInSieve({ scripts: { 'my-filters': 'require "fileinto";\r\n' } });
  server.setActive('my-filters');
  const port = await server.listen();
  const session = await openSieve({ host: '127.0.0.1', port, user: 'amira@checks.example', pass: 'check', plain: true });
  try {
    assert.deepEqual(await readState(session), { on: false, settings: null });
    assert.deepEqual(await turnOn(session, { message: 'Away.' }, { addresses: ['amira@checks.example'] }), { on: true, previous: 'my-filters' });
    assert.equal(server.active(), SCRIPT_NAME);
    // The server has Sieve's variables: the subject promised is the one written (A6.11).
    assert.match(server.scripts.get(SCRIPT_NAME), /:subject "Automatic reply: \$\{subject\}"/);
    // On again with new words: what was before is still what comes back.
    assert.deepEqual(await turnOn(session, { message: 'Still away.' }), { on: true, previous: 'my-filters' });
    const state = await readState(session);
    assert.equal(state.on, true);
    assert.equal(state.settings.message, 'Still away.');
    assert.deepEqual(await turnOff(session), { on: false, previous: 'my-filters' });
    assert.equal(server.active(), 'my-filters');
    // With no filters of their own, off is no active script at all.
    server.scripts.delete('my-filters');
    await turnOn(session, { message: 'Away.' });
    await turnOff(session);
    assert.equal(server.active(), null);
  } finally {
    await session.close();
    await server.close();
  }
});

test('a wrong password, a server without automatic replies, and a password never sent in the clear', async () => {
  const server = standInSieve();
  const port = await server.listen();
  await assert.rejects(openSieve({ host: '127.0.0.1', port, user: 'amira@checks.example', pass: 'wrong', plain: true }), (e) => e.code === 'auth');
  await server.close();
  // The stand-in offers no STARTTLS: Office does not sign in over it at all.
  const open = standInSieve();
  const p1 = await open.listen();
  await assert.rejects(openSieve({ host: '127.0.0.1', port: p1, user: 'amira@checks.example', pass: 'check' }), (e) => e.code === 'no-tls');
  assert.equal(open.log.filter((l) => l[0] === 'AUTHENTICATE').length, 0);
  await open.close();
  const without = standInSieve({ vacation: false });
  const p2 = await without.listen();
  await assert.rejects(openSieve({ host: '127.0.0.1', port: p2, user: 'amira@checks.example', pass: 'check', plain: true }), (e) => e.code === 'no-vacation');
  await without.close();
});
