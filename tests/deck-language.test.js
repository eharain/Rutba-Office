// Review → Language in Presentations: a run's language and "do not check"
// read from a:rPr and written back on it, so an edit no longer turns every
// run into US English, and the spelling pass reads past what is marked.

import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';
import { deckSegments } from '@rutba/proofing';

const deck = () => Deck.open(buildPptx({ slides: [{ layout: 'title', title: 'Bonjour tout le monde', body: 'Hello there' }] }));
const titleOf = (d) => d.slide(0).shapes.find((s) => s.placeholder?.type === 'title' || s.placeholder?.type === 'ctrTitle');

test('a run marked French and not to be checked is written so, and read back so', () => {
  const d = deck();
  const title = titleOf(d);
  const paragraphs = title.text.paragraphs.map(({ plain, ...p }) => ({ ...p, runs: p.runs.map((r) => ({ ...r, lang: 'fr-FR', noProof: true })) }));
  d.setText(0, title.id, paragraphs);
  const back = Deck.open(d.save());
  const run = titleOf(back).text.paragraphs[0].runs[0];
  assert.equal(run.lang, 'fr-FR');
  assert.equal(run.noProof, true);
  assert.match(back.pkg.text(back.slideParts[0].part), /<a:rPr lang="fr-FR"[^>]*noProof="1"/);
});

test('an edit keeps the language each run already had, rather than making it US English', () => {
  const d = deck();
  const title = titleOf(d);
  const part = d.slideParts[0].part;
  d.pkg.write_(part, Buffer.from(d.pkg.text(part).replace(/<a:rPr lang="en-US"/g, () => '<a:rPr lang="en-GB"'), 'utf8'));
  const fresh = titleOf(d);
  assert.equal(fresh.text.paragraphs[0].runs[0].lang, 'en-GB');
  // A format press: bold over the whole box, everything else as read.
  d.setText(0, fresh.id, fresh.text.paragraphs.map(({ plain, ...p }) => ({ ...p, runs: p.runs.map((r) => ({ ...r, bold: true })) })));
  assert.match(d.pkg.text(part), /<a:rPr lang="en-GB"[^>]*b="1"/);
  assert.doesNotMatch(d.pkg.text(part).slice(d.pkg.text(part).indexOf('Bonjour') - 200, d.pkg.text(part).indexOf('Bonjour')), /lang="en-US"/);
});

test('Spelling reads past runs marked French or not to be checked', () => {
  const d = deck();
  const title = titleOf(d);
  d.setText(0, title.id, title.text.paragraphs.map(({ plain, ...p }) => ({ ...p, runs: p.runs.map((r) => ({ ...r, lang: 'fr-FR' })) })));
  const texts = deckSegments(d).map((s) => s.text);
  assert.ok(!texts.some((t) => t.includes('Bonjour')), texts.join(' | '));
  assert.ok(texts.some((t) => t.includes('Hello')), 'the English body is still read');
});
