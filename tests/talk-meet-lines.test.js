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
  meetingCode, isLive, meetingIn, liveNotice, recordedNotice, airText, studioFor, meetingNote, withMeeting, meetingName, meetRefusal, goneText, personMarks, meetingLine, tilesAcross, whenText,
  cannotJoin, whoCanJoin, ownMeetingRule, deviceWords,
  airMarks, plainNotice, liveSwitchText, controlWord, notConnectedText,
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

test('the invitation says how to join and who can, and first of all that the meeting will be streamed live when it will', () => {
  const plain = meetingNote({ link: link(PLAIN), live: false });
  assert.ok(words(plain).includes(link(PLAIN)));
  assert.ok(!plain.includes(words(liveNotice())));
  assert.ok(plain.includes('Only Talk contacts of the organiser can join'), 'who can join travels with the invitation, so nobody is turned away unwarned');
  assert.equal(plain.split('\n').length, 2);
  const live = meetingNote({ link: link(LIVE), live: true });
  assert.ok(live.includes(link(LIVE)));
  assert.ok(live.includes(liveNotice()), 'the notice travels with the invitation');
  assert.equal(live.split('\n').length, 3);
  assert.equal(live.split('\n')[0], liveNotice(), 'and comes first, so nobody reads past it');
  const taped = meetingNote({ link: link(PLAIN), live: false, recorded: true });
  assert.ok(taped.includes(words(recordedNotice())), 'and so does the notice that it will be recorded');
  assert.equal(taped.split('\n')[0], recordedNotice());
  assert.ok(!taped.includes(liveNotice()));
});

test('on air or being recorded has its words, with or without this person in the picture, and nothing is said while it is neither', () => {
  const said = [['live', true], ['live', false], ['record', true], ['record', false]].map(([how, you]) => words(airText({ on: true, you, how })));
  assert.equal(new Set(said).size, 4);
  assert.equal(airText({ on: false, you: false, how: null }), '');
  assert.equal(airText(undefined), '');
});

test('a meeting goes to Live Broadcast when set up to be streamed, to the Recorder when set up to be recorded, and nowhere otherwise', () => {
  assert.equal(studioFor({ live: true, recorded: false }), 'broadcast');
  assert.equal(studioFor({ live: false, recorded: true }), 'recorder');
  assert.equal(studioFor({ live: false, recorded: false }), null);
  assert.equal(studioFor(null), null);
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

/* ── who can come in, and a microphone or camera that cannot be used (the screen review of 2026-11-22) ── */

test('who cannot come in is known from what the meeting says of itself: full or locked, never someone already in it', () => {
  const people = (n) => Array.from({ length: n }, (_, i) => ({ id: `u-${i}` }));
  const base = { starter: { id: 'u-0' }, most: 4, locked: false };
  assert.equal(cannotJoin({ ...base, people: people(3) }, 'u-me'), null);
  assert.equal(cannotJoin({ ...base, people: people(4) }, 'u-me'), 'full');
  assert.equal(cannotJoin({ ...base, people: people(4) }, 'u-2'), null, 'already in it on another computer');
  assert.equal(cannotJoin({ ...base, locked: true, people: people(1) }, 'u-me'), 'locked');
  assert.equal(cannotJoin({ ...base, locked: true, people: [] }, 'u-0'), null, 'the person who set it up');
  assert.equal(cannotJoin({ people: people(4) }, 'u-me'), 'full', 'four when the meeting does not say how many it holds');
  assert.equal(cannotJoin({ ...base }, 'u-me'), null);
  assert.equal(cannotJoin(null, 'u-me'), null);
  for (const code of ['full', 'locked']) words(meetRefusal(code, { most: 4 }));
  assert.ok(meetRefusal('full', { most: 4 }).includes('4'));
});

test('who can join with the link is said as Talk has it: the contacts of the person who set it up, and the group for a meeting started in one', () => {
  const mine = { starter: { id: 'u-a', name: 'Amira Khan' }, most: 4 };
  const own = words(whoCanJoin(mine, 'u-a'));
  assert.equal(own, ownMeetingRule(4));
  assert.ok(own.includes('your Talk contacts') && own.includes('4'));
  const theirs = words(whoCanJoin(mine, 'u-ben'));
  assert.ok(theirs.includes('Talk contacts of Amira Khan'));
  const ownGroup = words(whoCanJoin({ ...mine, groupId: 'g_1' }, 'u-a'));
  const theirGroup = words(whoCanJoin({ ...mine, groupId: 'g_1' }, 'u-ben'));
  assert.ok(ownGroup.includes('this group') && theirGroup.includes('this group') && theirGroup.includes('Amira Khan'));
  assert.equal(new Set([own, theirs, ownGroup, theirGroup]).size, 4, 'four things to say, each its own words');
  words(whoCanJoin({ starter: { id: 'u-x' } }, 'u-ben'));
  words(whoCanJoin(null, null));
});

test('a microphone or a camera that cannot be used has words for what it means and what to check', () => {
  const said = [['microphone', 'none'], ['microphone', 'blocked'], ['camera', 'none'], ['camera', 'blocked']].map(([kind, fault]) => words(deviceWords(kind, fault)));
  assert.equal(new Set(said).size, 4);
  assert.ok(said[0].includes('Others cannot hear you') && said[1].includes('Others cannot hear you'));
  assert.ok(said[1].includes('privacy settings') && said[3].includes('privacy settings'));
  assert.ok(said.every((s) => !/Windows/.test(s)), 'Office also runs on a Mac and on Linux');
  assert.equal(deviceWords('microphone', null), '');
  assert.equal(deviceWords('camera', undefined), '');
});

/* ── the screen review's "should fix" (2.27.0) ── */

test('the mark of a meeting to be streamed or recorded: a few words each, and none for any other', () => {
  assert.deepEqual(airMarks({ live: false, recorded: false }), []);
  assert.deepEqual(airMarks(null), []);
  const live = airMarks({ live: true });
  assert.deepEqual(live.map((m) => m.kind), ['live']);
  words(live[0].words);
  const both = airMarks({ live: true, recorded: true });
  assert.deepEqual(both.map((m) => m.kind), ['live', 'recorded']);
  assert.notEqual(words(both[0].words), words(both[1].words));
});

test('a meeting set up neither to be streamed nor recorded: said as how it was set up, and that this cannot change', () => {
  const said = words(plainNotice());
  assert.ok(said.includes('not set up to be streamed live or recorded'));
  assert.ok(said.includes('cannot be changed'));
  assert.notEqual(said, liveNotice());
});

test('the switch at set-up says one thing while it is off and another once it is on', () => {
  const off = words(liveSwitchText(false));
  const on = words(liveSwitchText(true));
  assert.notEqual(off, on);
  assert.ok(off.includes('cannot be made live once it is set up'));
  // Turning it on streams nothing by itself: that is done in Live Broadcast.
  assert.ok(on.includes('Nothing is streamed until') && on.includes('Live Broadcast'));
  assert.ok(on.includes('told before they join'));
});

test('the word under each round button says how it stands now', () => {
  for (const which of ['microphone', 'camera', 'screen', 'hand']) {
    assert.notEqual(words(controlWord(which, true)), words(controlWord(which, false)), which);
  }
  assert.equal(controlWord('camera', false), 'Camera off');
  assert.equal(controlWord('microphone', false), 'Muted');
  assert.equal(controlWord('screen', true), 'Sharing');
  assert.equal(controlWord('something-else', true), '');
});

test('someone who could not be connected to: who, and the likely reason, never a bare "could not connect"', () => {
  const said = words(notConnectedText('Cara Lund'));
  assert.ok(said.includes('Cara Lund') && said.includes('network'));
  words(notConnectedText(''));
  assert.ok(!notConnectedText('').includes('{name}'));
});

test('the person who set a meeting up is marked "Organiser", the word the invitation and the join screen use', () => {
  assert.equal(personMarks({}, { starter: true }), 'Organiser');
});
