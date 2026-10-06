// Animations → Trigger: an effect started by a click on a shape in the
// show, kept as PowerPoint keeps it — an interactive sequence of its own,
// out of the slide's click order — read back, moved back, and gone with the
// shape that triggered it. A trigger PowerPoint wrote is read and kept.

import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';

function deck() {
  const d = Deck.open(buildPptx({ slides: [{ layout: 'title', title: 'Press the button', body: 'The answer' }, { layout: 'obj', title: 'Next' }] }));
  d.addTextBox(0, { x: 40, y: 40, w: 160, h: 50, paragraphs: [{ runs: [{ text: 'Button' }] }] });
  const shapes = d.slide(0).shapes;
  return { d, answer: shapes[1].id, button: shapes[shapes.length - 1].id };
}
const timingOf = (d) => { const x = d.pkg.text(d.slideParts[0].part); return x.slice(x.indexOf('<p:timing'), x.indexOf('</p:timing>') + 11); };

test('an effect put on a trigger leaves the slide\'s click order for an interactive sequence of that shape', () => {
  const { d, answer, button } = deck();
  d.addAnimation(0, answer, { kind: 'entr', effect: 'fade' });
  d.setAnimation(0, 0, { triggerShape: button });
  const [e] = d.animations(0);
  assert.deepEqual([e.kind, e.effect, e.trigger, e.triggerShape], ['entr', 'fade', 'onClick', String(button)]);
  const t = timingOf(d);
  assert.doesNotMatch(t, /nodeType="mainSeq"/, 'nothing is left in the slide\'s own order');
  assert.match(t, new RegExp(`nodeType="interactiveSeq"><p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${button}"/>`));
  assert.match(t, new RegExp(`<p:nextCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${button}"/></p:tgtEl></p:cond></p:nextCondLst></p:seq>`));
  assert.match(t, new RegExp(`<p:bldP spid="${answer}" grpId="0"`), 'the build entry is kept');

  // Saved and opened again, the trigger is read back.
  const back = Deck.open(d.save());
  assert.equal(back.animations(0)[0].triggerShape, String(button));
});

test('the slide\'s own effects and a trigger\'s live side by side, and an effect goes back to the slide\'s order', () => {
  const { d, answer, button } = deck();
  d.addAnimation(0, answer, { kind: 'entr', effect: 'fade' });
  d.addAnimation(0, button, { kind: 'emph', effect: 'pulse' });
  d.setAnimation(0, 0, { triggerShape: button });
  assert.deepEqual(d.animations(0).map((e) => [e.shapeId, e.triggerShape]), [[String(button), null], [String(answer), String(button)]], 'the slide\'s own first, the trigger\'s after');
  const t = timingOf(d);
  assert.ok(t.indexOf('nodeType="mainSeq"') < t.indexOf('nodeType="interactiveSeq"'), 'the main sequence first, as PowerPoint orders them');
  d.setAnimation(0, 1, { triggerShape: null });
  assert.deepEqual(d.animations(0).map((e) => e.triggerShape), [null, null]);
  assert.doesNotMatch(timingOf(d), /interactiveSeq/);
  assert.throws(() => d.setAnimation(0, 0, { triggerShape: 999 }), /not on this slide/);
});

test('deleting the trigger shape takes the effects it started with it', () => {
  const { d, answer, button } = deck();
  d.addAnimation(0, answer, { kind: 'entr', effect: 'fade' });
  d.setAnimation(0, 0, { triggerShape: button });
  d.removeShape(0, button);
  assert.deepEqual(d.animations(0), []);
});
