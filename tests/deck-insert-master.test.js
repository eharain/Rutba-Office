// Slide Master → Insert Slide Master: a second master — the first's copy,
// its layouts and its theme in parts of their own — listed in the
// presentation with ids no other master or layout has, kept while unused,
// and a slide moved onto one of its layouts draws from it.

import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';

const deck = () => Deck.open(buildPptx({ slides: [{ layout: 'title', title: 'One' }, { layout: 'obj', title: 'Two', body: ['A point'] }] }));
const relsOf = (d, part) => d.pkg.text(part.replace(/([^/]+)$/, '_rels/$1.rels'));

test('a second master is the first\'s copy in parts of its own: its layouts, its theme, its relationships', () => {
  const d = deck();
  const [first] = d.masterParts();
  const added = d.insertMaster();
  assert.deepEqual(d.masterParts(), [first, added]);
  const [one, two] = d.masterList();
  assert.equal(two.name, 'Custom Design');
  assert.equal(two.preserve, true, 'kept even while no slide uses it');
  assert.equal(two.layouts.length, one.layouts.length);
  assert.ok(two.layouts.every((l) => !one.layouts.some((o) => o.part === l.part)), 'layouts of its own');
  for (const l of two.layouts) assert.match(relsOf(d, l.part), new RegExp(`Target="../slideMasters/${added.split('/').pop()}"`), `${l.part} points at the new master`);
  const themeOf = (m) => /Type="[^"]*\/theme" Target="([^"]+)"/.exec(relsOf(d, m))?.[1];
  assert.notEqual(themeOf(added), themeOf(first), 'a theme of its own');
  assert.ok(d.pkg.has(`ppt/theme/${themeOf(added).split('/').pop()}`));
});

test('every master and layout id is its own, above 2^31, and the deck opens again with both masters', () => {
  const d = deck();
  d.insertMaster();
  const pres = d.pkg.text('ppt/presentation.xml');
  const ids = [...pres.matchAll(/<p:sldMasterId\b[^>]*\bid="(\d+)"/g)].map((m) => Number(m[1]));
  for (const m of d.masterParts()) for (const x of d.pkg.text(m).matchAll(/<p:sldLayoutId\b[^>]*\bid="(\d+)"/g)) ids.push(Number(x[1]));
  assert.equal(new Set(ids).size, ids.length, 'no id twice');
  assert.ok(ids.every((n) => n >= 2147483648));
  const back = Deck.open(d.save());
  assert.equal(back.masterParts().length, 2);
  assert.equal(back.slideCount, 2);
  assert.ok(back.slide(1).shapes.length > 0, 'the slides still draw');
});

test('a slide put on one of the new master\'s layouts draws from the new master', () => {
  const d = deck();
  const added = d.insertMaster();
  const layout = d.masterList()[1].layouts.find((l) => l.type === 'obj') || d.masterList()[1].layouts[0];
  d.applyLayout(1, layout.part);
  assert.equal(d.masterOf(1), added);
  assert.equal(d.masterOf(0), d.masterParts()[0], 'the other slide stays on the first');
});
