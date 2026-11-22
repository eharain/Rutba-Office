// What Talk says about a conversation beside its messages (Office 2.28):
// counts, a missed call, whether a message arrived, what changed in a group,
// and why something was refused. The words are the window's own
// (talk/lines.js), read here without a window.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { isKey } from '../tools/label-keys.mjs';
import { isLabelKey, englishOf, pluralOf } from './label-english.js';
import {
  groupKey, groupIdOf, countText, titleWithCount, waitingIn, groupTitle, groupPeopleLine, senderName, canJoin,
  eventText, missedText, missedWhy, markOf, refusalText, reportReasons, memberNote,
} from '../apps/desktop/renderer/apps/talk/lines.js';
import { REPORT_REASONS, groupKey as serviceGroupKey } from '../apps/desktop/main/talk-conversations.js';

const amira = { id: 'u-a', name: 'Amira Khan', email: 'amira@example.com', contact: false };
const ben = { id: 'u-ben', name: 'Ben Ortiz', email: 'ben@example.com', contact: true };
const cara = { id: 'u-cara', name: 'Cara Lund', contact: false, needsUpdate: true };
const group = { id: 'g1', name: '', starter: 'u-ben', created: 1, members: [amira, ben, cara] };

/** A line is words a person reads: never a bare key, never empty. */
const words = (s) => { assert.equal(typeof s, 'string'); assert.ok(s.trim().length > 0, 'not empty'); assert.ok(!isKey(s), `"${s}" is a key, not words`); return s; };

test('a group conversation is named as the service names it, and never as a person', () => {
  assert.equal(groupKey('g1'), 'group:g1');
  assert.equal(groupKey('g1'), serviceGroupKey('g1'));
  assert.equal(groupIdOf('group:g1'), 'g1');
  assert.equal(groupIdOf('u-ben'), null);
  assert.equal(groupIdOf(null), null);
});

test('a count shows nothing for none and stays narrow', () => {
  assert.equal(countText(0), '');
  assert.equal(countText(undefined), '');
  assert.equal(countText(7), '7');
  assert.equal(countText(99), '99');
  assert.equal(countText(100), '99+');
});

test("Talk's title carries what waits in it: messages not read and calls not looked at", () => {
  assert.equal(waitingIn(null), 0);
  assert.equal(waitingIn({ unreadTotal: 3, missed: 2 }), 5);
  assert.equal(titleWithCount('Talk', 0), 'Talk');
  assert.equal(titleWithCount('Talk', 5), 'Talk (5)');
  assert.equal(titleWithCount('Talk', 250), 'Talk (99+)');
});

test('a group without a name is named by the others in it', () => {
  assert.equal(groupTitle({ ...group, name: 'Budget' }, 'u-a'), 'Budget');
  assert.equal(groupTitle(group, 'u-a'), 'Ben, Cara');
  assert.equal(groupTitle(group, 'u-ben'), 'Amira, Cara');
  words(groupTitle({ ...group, members: [amira] }, 'u-a'));
  const line = groupPeopleLine(group, 'u-a');
  assert.ok(line.startsWith('Ben Ortiz, Cara Lund, '));
  assert.equal(senderName(group, 'u-cara'), 'Cara Lund');
  words(senderName(group, 'u-gone'));
});

test('only contacts not in the group already can be put in it', () => {
  const contacts = [{ id: 'u-ben' }, { id: 'u-dan' }];
  assert.deepEqual(canJoin(contacts, group), [{ id: 'u-dan' }]);
  assert.deepEqual(canJoin(null, group), []);
});

test('what changed in a group is said as the reader sees it', () => {
  const by = { id: 'u-ben', name: 'Ben Ortiz' };
  const who = [{ id: 'u-dan', name: 'Dan Moyo' }];
  assert.ok(words(eventText({ kind: 'added-you', by }, 'u-a')).includes('Ben Ortiz'));
  const theirs = words(eventText({ kind: 'added', by, who }, 'u-a'));
  assert.ok(theirs.includes('Ben Ortiz') && theirs.includes('Dan Moyo'));
  const mine = words(eventText({ kind: 'added', by, who }, 'u-ben'));
  assert.ok(!mine.includes('Ben Ortiz') && mine.includes('Dan Moyo'), 'my own change does not name me');
  assert.ok(words(eventText({ kind: 'renamed', by, name: 'Budget' }, 'u-a')).includes('Budget'));
  words(eventText({ kind: 'renamed', by, name: '' }, 'u-a'));
  words(eventText({ kind: 'renamed', by, name: '' }, 'u-ben'));
  assert.ok(words(eventText({ kind: 'removed', by, who }, 'u-a')).includes('Dan Moyo'));
  assert.ok(words(eventText({ kind: 'removed', by, who }, 'u-ben')).includes('Dan Moyo'));
  assert.ok(words(eventText({ kind: 'left', by: who[0], who }, 'u-a')).includes('Dan Moyo'));
  assert.equal(eventText({ kind: 'something-new' }, 'u-a'), '', 'a change this Office does not know draws no line');
  assert.equal(eventText(null, 'u-a'), '');
});

test('a missed call says what kind it was and, when it helps, why', () => {
  assert.notEqual(words(missedText({ video: true })), words(missedText({ video: false })));
  // The same person calling more than once is one line that says how many.
  assert.equal(missedText({ video: false, count: 1 }), missedText({ video: false }));
  assert.ok(words(missedText({ video: true, count: 3 })).includes('3'));
  assert.notEqual(missedText({ count: 3 }), missedText({ count: 4 }));
  assert.equal(missedWhy('unanswered'), '');
  words(missedWhy('busy'));
  words(missedWhy('offline'));
});

test("a message of mine is marked as it goes: sending, sent, arrived, or not sent", () => {
  assert.equal(markOf({ mine: false, state: 'sent' }), null, 'only my own messages are marked');
  assert.equal(markOf({ mine: true }), null, 'a message from before the marks has none');
  const sending = markOf({ mine: true, state: 'sending' });
  assert.equal(sending.icon, null);
  const sent = markOf({ mine: true, state: 'sent' });
  const arrived = markOf({ mine: true, state: 'arrived' });
  assert.equal(sent.icon, 'check');
  assert.equal(sent.double, false);
  assert.equal(arrived.double, true);
  const failed = markOf({ mine: true, state: 'failed' });
  assert.equal(failed.bad, true);
  const said = [sending, sent, arrived, failed].map((m) => words(m.words));
  assert.equal(new Set(said).size, 4, 'each state has its own words, for a screen reader');
});

test('every refusal has words, and the most a group holds is said', () => {
  for (const code of ['not-a-contact', 'too-few', 'too-many', 'too-many-groups', 'no-group', 'not-yours', 'nobody', 'offline', 'no-answer', 'unknown', 'server', 'something-else']) words(refusalText(code));
  assert.ok(refusalText('too-many', { most: 50 }).includes('50'));
  assert.notEqual(refusalText('too-many'), refusalText('too-many', { most: 50 }));
  // Someone else being in as many groups as Talk takes is said of them, by name, not of the person reading.
  const full = words(refusalText('too-many-groups', { who: [{ id: 'u-cara', name: 'Cara Lund' }, { id: 'u-ben', name: 'Ben Ortiz' }] }));
  assert.ok(full.includes('Cara Lund') && full.includes('Ben Ortiz'));
  assert.notEqual(full, refusalText('too-many-groups'));
  assert.equal(refusalText('too-many-groups', { who: [] }), refusalText('too-many-groups'));
});

test('the reasons for a report are the ones the service takes', () => {
  const reasons = reportReasons();
  assert.deepEqual(reasons.map(([k]) => k), REPORT_REASONS);
  for (const [, label] of reasons) words(label);
});

test("a group's list of people says who started it, who is not a contact, and who cannot see it yet", () => {
  words(memberNote(amira, group, 'u-a'));
  assert.equal(memberNote({ ...ben }, { ...group, starter: 'u-a' }, 'u-a'), '', 'a contact with a current Office needs no note');
  const starter = memberNote(ben, group, 'u-a');
  words(starter);
  const stranger = memberNote(cara, group, 'u-a');
  assert.equal(stranger.split(' · ').length, 2, 'not a contact, and needs a newer Office');
});

test('the window names every label by its key', () => {
  for (const f of ['talk/lines.js', 'talk/groups.js', 'talk/report.js', 'talk/missed.js']) {
    const src = fs.readFileSync(new URL(`../apps/desktop/renderer/apps/${f}`, import.meta.url), 'utf8');
    const named = [...src.matchAll(/\bt\('([^']+)'/g), ...src.matchAll(/\btn\([^,()]+(?:\([^()]*\))?[^,()]*,\s*'([^']+)'/g)];
    assert.ok(named.length > 0, `${f} names no label at all`);
    for (const m of named) {
      assert.ok(isKey(m[1]), `${f}: '${m[1]}' is not a key`);
      // A key that only looks like one has no wording: the window would show the key itself.
      assert.ok(isLabelKey(m[1]), `${f}: '${m[1]}' has no wording in the catalogue`);
      assert.ok(String(englishOf(m[1])).trim().length > 0, `${f}: '${m[1]}' has empty wording`);
    }
  }
  // A count's label has both its forms.
  assert.ok(pluralOf('talk.lines.count-missed-call'), 'one form and the other');
});
