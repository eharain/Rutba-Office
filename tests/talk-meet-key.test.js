// The key of a meeting held on a Talk Server, sealed for each person in the
// meeting's set-up (apps/desktop/renderer/meeting/meet-key.js).

import test from 'node:test';
import assert from 'node:assert/strict';
import { askForKey, newMeetingKey, readSetUp, sealKeyFor, wireSeal } from '../apps/desktop/renderer/meeting/meet-key.js';

const WHO = { meeting: 'meet-AAAAAAAAAAAAAAAAAAAA', from: 'u-amira', to: 'u-ben' };

test('a new key is 32 bytes, and never the same twice', () => {
  const a = newMeetingKey();
  const b = newMeetingKey();
  assert.equal(a.length, 32);
  assert.notDeepEqual([...a], [...b]);
});

test('the key sealed for the one who asked opens for them, as the same key', async () => {
  const key = newMeetingKey();
  const ask = await askForKey(WHO);
  const sealed = await sealKeyFor(ask.pub, key, WHO);
  assert.deepEqual(Object.keys(sealed).sort(), ['box', 'iv', 'pub']);
  for (const v of Object.values(sealed)) assert.match(v, /^[\w-]+$/);
  assert.deepEqual([...(await ask.open(sealed))], [...key]);
  // What travels is not the key, nor anything that holds it in the clear.
  const travelled = JSON.stringify(sealed);
  assert.ok(!travelled.includes(Buffer.from(key).toString('base64url')));
});

test('sealed twice for the same asking, it never travels the same way twice', async () => {
  const key = newMeetingKey();
  const ask = await askForKey(WHO);
  const one = await sealKeyFor(ask.pub, key, WHO);
  const two = await sealKeyFor(ask.pub, key, WHO);
  assert.notEqual(one.box, two.box);
  assert.notEqual(one.pub, two.pub);
});

test('it opens for nobody else: not another asking, another person, another meeting, or another sender', async () => {
  const key = newMeetingKey();
  const ask = await askForKey(WHO);
  const sealed = await sealKeyFor(ask.pub, key, WHO);
  const other = await askForKey(WHO);
  assert.equal(await other.open(sealed), null);
  for (const who of [{ ...WHO, to: 'u-cara' }, { ...WHO, meeting: 'meet-BBBBBBBBBBBBBBBBBBBB' }, { ...WHO, from: 'u-mallory' }]) {
    const wrong = await sealKeyFor(ask.pub, key, who);
    assert.equal(await ask.open(wrong), null, JSON.stringify(who));
  }
});

test('a box that was changed on the way opens as nothing', async () => {
  const key = newMeetingKey();
  const ask = await askForKey(WHO);
  const sealed = await sealKeyFor(ask.pub, key, WHO);
  const flip = (s) => `${s.slice(0, -2)}${s.at(-2) === 'A' ? 'B' : 'A'}${s.at(-1)}`;
  assert.equal(await ask.open({ ...sealed, box: flip(sealed.box) }), null);
  assert.equal(await ask.open({ ...sealed, iv: flip(sealed.iv) }), null);
  assert.equal(await ask.open({ ...sealed, pub: (await askForKey(WHO)).pub }), null);
  for (const junk of [null, {}, { pub: 'x', iv: 'y', box: 'z' }, { ...sealed, box: '' }, { ...sealed, iv: 'AAAA' }, { ...sealed, box: 'not base64!' }]) assert.equal(await ask.open(junk), null);
});

test('nothing is sealed for what is not a public half, or with what is not a key', async () => {
  const ask = await askForKey(WHO);
  assert.equal(await sealKeyFor('nonsense', newMeetingKey(), WHO), null);
  assert.equal(await sealKeyFor('', newMeetingKey(), WHO), null);
  assert.equal(await sealKeyFor(ask.pub, new Uint8Array(16), WHO), null);
  assert.equal(await sealKeyFor(ask.pub, 'a string of thirty-two characters', WHO), null);
  // A point that is not on the curve.
  const off = Buffer.concat([Buffer.from([4]), Buffer.alloc(64, 1)]).toString('base64url');
  assert.equal(await sealKeyFor(off, newMeetingKey(), WHO), null);
});

test('what the apps say to each other in the set-up is checked as it arrives', async () => {
  const ask = await askForKey(WHO);
  const sealed = await sealKeyFor(ask.pub, newMeetingKey(), WHO);
  assert.deepEqual(readSetUp({ t: 'srv-ask', pub: ask.pub, more: 1 }), { t: 'srv-ask', pub: ask.pub });
  assert.deepEqual(readSetUp({ t: 'srv-let', mid: WHO.meeting, ...sealed, extra: true }), { t: 'srv-let', mid: WHO.meeting, ...sealed });
  assert.deepEqual(readSetUp({ t: 'srv-no', why: 'no-server' }), { t: 'srv-no', why: 'no-server' });
  assert.deepEqual(readSetUp({ t: 'srv-no', why: '<b>anything</b>' }), { t: 'srv-no', why: 'refused' });
  assert.deepEqual(readSetUp({ t: 'srv-moved' }), { t: 'srv-moved' });
  for (const junk of [null, 'srv-ask', {}, { t: 'srv-ask' }, { t: 'srv-ask', pub: 'short' }, { t: 'srv-let', mid: 'x', ...sealed }, { t: 'srv-let', mid: `${WHO.meeting}/../x`, ...sealed }, { t: 'srv-let', mid: WHO.meeting, ...sealed, box: '' }, { t: 'offer', sdp: 'v=0' }]) {
    assert.equal(readSetUp(junk), null, JSON.stringify(junk));
  }
});

test('the line beside sound and picture: sealed with the meeting\'s key, opened only as from its sender, for whom it was meant, and once', async () => {
  const key = newMeetingKey();
  const who = (me) => ({ meeting: 'meet-AAAAAAAAAAAAAAAAAAAA', me });
  const amira = await wireSeal(key, who('u-amira'));
  const ben = await wireSeal(key, who('u-ben'));
  const cara = await wireSeal(key, who('u-cara'));
  const words = { t: 'say', id: 's1', text: 'Shall we start?' };

  const toAll = await amira.seal(words);
  assert.ok(!Buffer.from(toAll).toString('latin1').includes('start'), 'nothing of it can be read as it travels');
  const benHeard = { run: null, seq: 0 };
  assert.deepEqual(JSON.parse(await ben.open(toAll, 'u-amira', benHeard)), words);
  assert.equal(benHeard.seq, 1);
  // Passed on a second time, or as somebody else's: not opened.
  assert.equal(await ben.open(toAll, 'u-amira', benHeard), null);
  assert.equal(await cara.open(toAll, 'u-ben', { run: null, seq: 0 }), null);
  // The next one is taken, an earlier one after it is not.
  const second = await amira.seal({ t: 'mute' });
  const third = await amira.seal({ t: 'state' });
  assert.ok(await ben.open(third, 'u-amira', benHeard));
  assert.equal(await ben.open(second, 'u-amira', benHeard), null);
  assert.equal(benHeard.seq, 3);

  // Meant for one person: opens for them, not for another, and not as if it were for everyone.
  const forBen = await amira.seal({ t: 'air', on: true, you: true }, 'u-ben');
  assert.equal(await cara.open(forBen, 'u-amira', { run: null, seq: 0 }), null);
  const asForAll = forBen.slice();
  asForAll[1] = 0;
  assert.equal(await ben.open(asForAll, 'u-amira', { run: null, seq: 0 }), null);
  assert.ok(await ben.open(forBen, 'u-amira', { run: null, seq: 0 }));

  // Changed on the way, cut short, another meeting, another key, or no seal at all: nothing.
  const bent = toAll.slice();
  bent[bent.length - 1] ^= 1;
  assert.equal(await cara.open(bent, 'u-amira', { run: null, seq: 0 }), null);
  const counted = toAll.slice();
  counted[15] += 1;
  assert.equal(await cara.open(counted, 'u-amira', { run: null, seq: 0 }), null, 'the count is part of what is sealed');
  assert.equal(await cara.open(toAll.slice(0, 30), 'u-amira', { run: null, seq: 0 }), null);
  assert.equal(await (await wireSeal(key, { meeting: 'meet-BBBBBBBBBBBBBBBBBBBB', me: 'u-cara' })).open(toAll, 'u-amira', { run: null, seq: 0 }), null);
  assert.equal(await (await wireSeal(newMeetingKey(), who('u-cara'))).open(toAll, 'u-amira', { run: null, seq: 0 }), null);
  assert.equal(await cara.open(new TextEncoder().encode(JSON.stringify(words).padEnd(60)), 'u-amira', { run: null, seq: 0 }), null);
  assert.equal(await cara.open('not bytes', 'u-amira', { run: null, seq: 0 }), null);
  // Somebody who came in again counts from the start; whoever knew their earlier stay does not take the new one for it.
  const amiraAgain = await wireSeal(key, who('u-amira'));
  assert.equal(await ben.open(await amiraAgain.seal(words), 'u-amira', benHeard), null);
  assert.ok(await ben.open(await amiraAgain.seal(words), 'u-amira', { run: null, seq: 0 }));
});
