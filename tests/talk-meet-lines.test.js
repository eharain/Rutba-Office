// What Talk, Calendar and Mail say about a small meeting (Office 2.28,
// renderer/apps/talk/meet-lines.js): its link read out of an event's words,
// the notice that it will be streamed live, what an invitation carries, why
// joining was refused, and the marks beside a person. Read without a window.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { isKey } from '../tools/label-keys.mjs';
import { isLabelKey } from './label-english.js';
import {
  meetingCode, isLive, meetingIn, liveNotice, meetingNote, withMeeting, meetingName, meetRefusal, goneText, personMarks, meetingLine, tilesAcross, whenText,
} from '../apps/desktop/renderer/apps/talk/meet-lines.js';
import { meetingCode as serviceCode } from '../apps/desktop/main/talk-meetings.js';

const PLAIN = 'meet-AAAAAAAAAAAAAAAAAAAA';
const LIVE = 'live-BBBBBBBBBBBBBBBBBBBB';
const link = (code) => `https://introducer.rutba.io/m/${code}`;

/** A line is words a person reads: never a bare key, never empty. */
const words = (s) => { assert.equal(typeof s, 'string'); assert.ok(s.trim().length > 0, 'not empty'); assert.ok(!isKey(s), `"${s}" is a key, not words`); return s; };

test("a meeting's code is read as the service reads it", () => {
  for (const x of [PLAIN, LIVE, link(PLAIN), `${link(LIVE)}/`, `rutba-office://talk/meet/${PLAIN}`, '', null, 'meet-short', `${PLAIN}A`, `https://introducer.rutba.io/i/${PLAIN}`, `${link(PLAIN)} and more`]) {
    assert.equal(meetingCode(x), serviceCode(x), String(x));
  }
  assert.equal(meetingCode(link(LIVE)), LIVE);
});

test('whether a meeting will be streamed live is in its code, so nothing kept beside the link can change it', () => {
  assert.equal(isLive(LIVE), true);
  assert.equal(isLive(PLAIN), false);
  assert.equal(isLive(''), false);
  assert.equal(isLive(null), false);
});

test("the meeting an event carries is found in its place or its notes, and a link that is nearly one is not", () => {
  assert.deepEqual(meetingIn(link(PLAIN), ''), { code: PLAIN, link: link(PLAIN), live: false });
  assert.deepEqual(meetingIn('Room 4', `Agenda attached.\n\nJoin: ${link(LIVE)}.\nSee you there`), { code: LIVE, link: link(LIVE), live: true });
  assert.equal(meetingIn(`<a href="${link(PLAIN)}">join</a>`).code, PLAIN);
  assert.equal(meetingIn('Room 4', 'No link here'), null);
  assert.equal(meetingIn(undefined, null), null);
  assert.equal(meetingIn(`${link(PLAIN)}AB`), null, 'a longer code is another thing');
  assert.equal(meetingIn(`https://introducer.rutba.io/i/${PLAIN}`), null);
  assert.equal(meetingIn(`rutba-office://talk/meet/${PLAIN}`), null, 'an event carries the page address, which works in mail');
  assert.equal(meetingIn(link(PLAIN), link(LIVE)).code, PLAIN, 'the place is read first');
});

test('the invitation says how to join, and that the meeting will be streamed live when it will', () => {
  const plain = meetingNote({ link: link(PLAIN), live: false });
  assert.ok(words(plain).includes(link(PLAIN)));
  assert.ok(!plain.includes(words(liveNotice())));
  const live = meetingNote({ link: link(LIVE), live: true });
  assert.ok(live.includes(link(LIVE)));
  assert.ok(live.includes(liveNotice()), 'the notice travels with the invitation');
  assert.equal(live.split('\n').length, 2);
});

test("an event takes its meeting: the link where the place was empty, the note under what was written, once", () => {
  const meeting = { link: link(LIVE), live: true };
  const empty = withMeeting({ location: '', description: '' }, meeting);
  assert.equal(empty.location, link(LIVE));
  assert.equal(empty.description, meetingNote(meeting));

  const kept = withMeeting({ location: 'Room 4', description: 'Bring the figures.\n' }, meeting);
  assert.equal(kept.location, 'Room 4', "a place the person wrote stays");
  assert.equal(kept.description, `Bring the figures.\n\n${meetingNote(meeting)}`);
  assert.deepEqual(meetingIn(kept.location, kept.description), { code: LIVE, link: link(LIVE), live: true }, 'and the meeting is found again');

  const again = withMeeting(kept, meeting);
  assert.equal(again.description, kept.description, 'not written twice');
  assert.deepEqual(withMeeting(undefined, { link: link(PLAIN), live: false }).location, link(PLAIN));
});

test('a meeting started now is named after its group, or the person', () => {
  assert.equal(meetingName({ group: 'Launch' }), 'Launch');
  assert.ok(words(meetingName({ me: { name: 'Amira Khan' } })).includes('Amira'));
  assert.ok(!meetingName({ me: { name: 'Amira Khan' } }).includes('Khan'));
  words(meetingName({}));
  words(meetingName());
});

test('every refusal has its own words, and says whose meeting it is and how many it holds', () => {
  const codes = ['no-meeting', 'not-allowed', 'removed', 'locked', 'full', 'not-yours', 'nobody', 'name', 'time', 'not-a-contact', 'no-group', 'too-many', 'unknown', 'offline', 'something-else'];
  const said = codes.map((c) => words(meetRefusal(c)));
  assert.equal(new Set(said).size, codes.length);
  assert.equal(meetRefusal('no-answer'), meetRefusal('offline'));
  assert.ok(meetRefusal('not-allowed', { starter: { name: 'Dana Ruiz' } }).includes('Dana Ruiz'));
  assert.notEqual(meetRefusal('not-allowed', { starter: { name: 'Dana Ruiz' } }), meetRefusal('not-allowed'));
  assert.ok(meetRefusal('full', { most: 4 }).includes('4'));
  assert.ok(meetRefusal('full').includes('4'), 'four, when the introducer did not say');
});

test('why this person is out of a meeting is said, except when they left themselves', () => {
  assert.ok(words(goneText('removed', { name: 'Ben Ortiz' })).includes('Ben Ortiz'));
  words(goneText('removed'));
  words(goneText('elsewhere'));
  assert.ok(words(goneText('ended', { name: 'Ben Ortiz' })).includes('Ben Ortiz'));
  words(goneText('ended'));
  words(goneText('lost'));
  assert.equal(goneText('left'), '');
  assert.equal(goneText('refused'), '', 'a refusal has its own words');
});

test('the marks beside a person, as a screen reader says them', () => {
  assert.equal(personMarks({}), '');
  assert.equal(personMarks({ muted: true, hand: true, sharing: true }, { starter: true, me: true }).split(' · ').length, 5);
  const each = [personMarks({}, { me: true }), personMarks({}, { starter: true }), personMarks({ hand: true }), personMarks({ sharing: true }), personMarks({ muted: true })].map(words);
  assert.equal(new Set(each).size, 5);
});

test("the line a group's conversation keeps for a meeting says who started it", () => {
  const entry = { kind: 'meeting', meeting: { id: PLAIN, name: 'Launch' }, by: { id: 'u-ben', name: 'Ben Ortiz' } };
  const theirs = words(meetingLine(entry, 'u-a'));
  assert.ok(theirs.includes('Ben Ortiz') && theirs.includes('Launch'));
  const mine = words(meetingLine(entry, 'u-ben'));
  assert.ok(mine.includes('Launch') && !mine.includes('Ben Ortiz'));
  words(meetingLine({ kind: 'meeting' }, 'u-a'));
});

test('up to four tiles: one alone, then two across', () => {
  assert.deepEqual([0, 1, 2, 3, 4].map(tilesAcross), [1, 1, 2, 2, 2]);
});

test('a planned meeting says when it is; one started now says nothing', () => {
  assert.equal(whenText(null), '');
  assert.equal(whenText(0), '');
  assert.match(whenText(Date.UTC(2026, 10, 23, 9, 30), 'en-GB'), /23 Nov/);
});

test('every label of a meeting is named by its key, and every key has its wording', () => {
  for (const f of ['talk/meet-lines.js', 'talk/meeting.js']) {
    const src = fs.readFileSync(new URL(`../apps/desktop/renderer/apps/${f}`, import.meta.url), 'utf8');
    const named = [...src.matchAll(/\bt\(\s*'((?:[^'\\]|\\.)*)'/g)].map((m) => m[1]);
    assert.ok(named.length > 0, `${f} names labels`);
    for (const k of named) {
      assert.ok(isKey(k), `${f}: t('${k}') is not a key`);
      assert.ok(isLabelKey(k), `${f}: t('${k}') has no wording in the catalogue`);
    }
  }
});
