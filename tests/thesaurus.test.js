/**
 * Review → Thesaurus: the suite's own word list, read into meanings — a
 * word in several sets has several, each with its part of speech — and a
 * word with an ending found by the word it was made from.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { thesaurusLookUp, thesaurusSize } from '@rutba/proofing';
import { THESAURUS_EN } from '../packages/proofing/src/thesaurus-en.js';
import { wordAround } from '../apps/desktop/renderer/word-around.js';

test('A word\'s meanings come back grouped by sense, each with its part of speech, the word itself left out', () => {
  const happy = thesaurusLookUp('happy');
  assert.equal(happy.base, 'happy');
  assert.ok(happy.meanings.length >= 2, 'glad, and willing: two senses at least');
  assert.ok(happy.meanings.every((m) => m.part === 'adjective'));
  assert.ok(happy.meanings[0].words.includes('cheerful'));
  assert.ok(!happy.meanings.some((m) => m.words.includes('happy')));
  const run = thesaurusLookUp('run');
  assert.ok(run.meanings.some((m) => m.part === 'verb' && m.words.includes('sprint')));
  assert.ok(run.meanings.some((m) => m.part === 'verb' && m.words.includes('manage')), 'run a business, as well as run a race');
});

test('A word with an ending is looked up by the word it was made from, and only the senses that ending fits', () => {
  assert.equal(thesaurusLookUp('Running').base, 'run');
  assert.equal(thesaurusLookUp('cities').base, 'city');
  assert.equal(thesaurusLookUp('cities').meanings[0].part, 'noun');
  const happier = thesaurusLookUp('happier');
  assert.equal(happier.base, 'happy');
  assert.equal(happier.how, 'comparative');
  assert.ok(happier.meanings.every((m) => m.part === 'adjective'));
  assert.equal(thesaurusLookUp('decided').base, 'decide');
  assert.deepEqual(thesaurusLookUp('qwertyuiop').meanings, []);
  assert.deepEqual(thesaurusLookUp('').meanings, []);
});

test('The list is sound: thousands of words, every set two words or more and a known part of speech, no set twice', () => {
  assert.ok(thesaurusSize() > 5000, `${thesaurusSize()} words`);
  const lines = THESAURUS_EN.split(/\r?\n/).filter((l) => l.trim());
  const seen = new Set();
  for (const line of lines) {
    assert.match(line, /^[nvar]: \S/, line);
    const words = line.slice(3).split(',').map((w) => w.trim());
    assert.ok(words.length >= 2 && words.every(Boolean), line);
    assert.ok(!seen.has(line.toLowerCase()), `twice: ${line}`);
    seen.add(line.toLowerCase());
  }
});

test('The word a caret is in, or a selection covers, is found in its paragraph', () => {
  assert.deepEqual(wordAround('The plan was good.', 14), { word: 'good', start: 13, end: 17 });
  assert.deepEqual(wordAround('The plan was good.', 13, 17), { word: 'good', start: 13, end: 17 });
  assert.deepEqual(wordAround('a well-known name', 5), { word: 'well-known', start: 2, end: 12 });
  assert.equal(wordAround('  ', 1), null);
});
