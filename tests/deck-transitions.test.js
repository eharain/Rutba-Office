// Transitions — the engine's half: a slide's `<p:transition>` written where
// the schema puts it (after p:clrMapOvr, before p:timing), in PowerPoint's
// own forms (the bare element when the duration is one of the three speeds,
// mc:AlternateContent with p14:dur when it is not), read back — including
// what PowerPoint itself wrote — changed, applied to every slide and taken
// off, with the rest of the slide untouched.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx, TRANSITIONS, readTransition, transitionXml } from '@rutba/presentation';

const DECK = () => Deck.open(buildPptx({
  title: 'Transitions',
  slides: [
    { layout: 'title', title: 'Opening', body: 'A deck that moves' },
    { layout: 'obj', title: 'Middle', body: ['One', 'Two'] },
    { layout: 'obj', title: 'Closing', body: ['Thanks'] },
  ],
}));
const xmlOf = (deck, i) => deck.pkg.text(deck.slideParts[i].part);

test('each classic transition is written and read back as itself, with its gallery direction and duration', () => {
  for (const [type, defaults] of Object.entries(TRANSITIONS)) {
    const deck = DECK();
    assert.equal(deck.setTransition(1, { type }), true, `${type} changes the slide`);
    const t = deck.transition(1);
    assert.equal(t.type, type);
    assert.equal(t.known, true);
    assert.equal(t.direction, defaults.direction, `${type} direction`);
    assert.equal(t.duration, defaults.duration, `${type} duration`);
    assert.equal(t.advanceOnClick, true);
    assert.equal(t.advanceAfter, null);
    assert.match(xmlOf(deck, 1), new RegExp(`<p:${type}\\b`), `${type} element written`);
    // The scene and the outline carry it too.
    assert.equal(deck.slide(1).transition.type, type);
    assert.deepEqual(deck.outline().map((o) => o.transition), [null, type, null]);
  }
});

test('directions are written the way PowerPoint writes them, and read back', () => {
  const deck = DECK();
  const cases = [
    ['push', 'u', '<p:push dir="u"/>'],
    ['push', 'l', '<p:push/>'],
    ['wipe', 'r', '<p:wipe dir="r"/>'],
    ['cover', 'rd', '<p:cover dir="rd"/>'],
    ['pull', 'd', '<p:pull dir="d"/>'],
    ['split', 'vert-in', '<p:split orient="vert" dir="in"/>'],
    ['split', 'horz-out', '<p:split/>'],
    ['randomBar', 'vert', '<p:randomBar dir="vert"/>'],
    ['zoom', 'out', '<p:zoom/>'],
    ['fade', 'black', '<p:fade thruBlk="1"/>'],
  ];
  for (const [type, direction, element] of cases) {
    deck.setTransition(0, { type, direction });
    assert.ok(xmlOf(deck, 0).includes(element), `${type} ${direction} → ${element}`);
    assert.equal(deck.transition(0).direction, direction);
  }
  // Effect Options alone keeps the effect and changes only the direction.
  deck.setTransition(0, { type: 'push', direction: 'u', duration: 1.25 });
  deck.setTransition(0, { direction: 'r' });
  assert.deepEqual([deck.transition(0).type, deck.transition(0).direction, deck.transition(0).duration], ['push', 'r', 1.25]);
});

test('a duration that is one of the three speeds is a bare element; any other takes the p14 form with a fallback', () => {
  const deck = DECK();
  deck.setTransition(0, { type: 'fade', duration: 0.75 });
  assert.match(xmlOf(deck, 0), /<p:transition spd="med"><p:fade\/><\/p:transition>/);
  assert.doesNotMatch(xmlOf(deck, 0), /AlternateContent/);
  assert.equal(deck.transition(0).duration, 0.75);

  deck.setTransition(0, { duration: 1.25 });
  const xml = xmlOf(deck, 0);
  assert.match(xml, /<mc:AlternateContent xmlns:mc="http:\/\/schemas\.openxmlformats\.org\/markup-compatibility\/2006"><mc:Choice xmlns:p14="http:\/\/schemas\.microsoft\.com\/office\/powerpoint\/2010\/main" Requires="p14"><p:transition spd="slow" p14:dur="1250"><p:fade\/><\/p:transition><\/mc:Choice><mc:Fallback><p:transition spd="slow"><p:fade\/><\/p:transition><\/mc:Fallback><\/mc:AlternateContent>/);
  assert.equal(deck.transition(0).duration, 1.25);
  // Only one transition on the slide, whatever the form.
  assert.equal((xml.match(/<mc:AlternateContent/g) || []).length, 1);
});

test('advance settings: On Mouse Click off and After a time, with or without an effect', () => {
  const deck = DECK();
  deck.setTransition(2, { type: 'wipe', advanceOnClick: false, advanceAfter: 3.5 });
  assert.match(xmlOf(deck, 2), /<p:transition spd="slow" advClick="0" advTm="3500"><p:wipe\/><\/p:transition>/);
  let t = deck.transition(2);
  assert.equal(t.advanceOnClick, false);
  assert.equal(t.advanceAfter, 3.5);

  // "None" keeps the advance settings with no effect.
  deck.setTransition(2, { type: 'none' });
  assert.match(xmlOf(deck, 2), /<p:transition advClick="0" advTm="3500"\/>/);
  t = deck.transition(2);
  assert.deepEqual([t.type, t.advanceOnClick, t.advanceAfter], ['none', false, 3.5]);

  // With nothing left to say, the element goes.
  deck.setTransition(2, { advanceOnClick: true, advanceAfter: null });
  assert.equal(deck.transition(2), null);
  assert.doesNotMatch(xmlOf(deck, 2), /<p:transition/);
});

test('Apply To All copies the slide’s transition, exactly, onto every slide; a slide with none clears them', () => {
  const deck = DECK();
  deck.setTransition(1, { type: 'push', direction: 'd', duration: 1.6, advanceAfter: 2 });
  deck.setTransition(2, { type: 'fade' });
  assert.equal(deck.applyTransitionToAll(1), 2);
  const block = /<mc:AlternateContent[\s\S]*?<\/mc:AlternateContent>/.exec(xmlOf(deck, 1))[0];
  for (const i of [0, 2]) {
    assert.ok(xmlOf(deck, i).includes(block), `slide ${i + 1} carries the same block`);
    assert.equal((xmlOf(deck, i).match(/<mc:AlternateContent/g) || []).length, 1, 'and only that one');
    assert.deepEqual(deck.transition(i), deck.transition(1));
  }
  // Applying again changes nothing.
  assert.equal(deck.applyTransitionToAll(1), 0);
  // From a slide with none: every transition goes.
  deck.setTransition(0, null);
  assert.equal(deck.applyTransitionToAll(0), 2);
  assert.deepEqual(deck.outline().map((o) => o.transition), [null, null, null]);
});

test('Remove takes the transition off and leaves the slide byte for byte as it was', () => {
  const deck = DECK();
  const before = xmlOf(deck, 1);
  deck.setTransition(1, { type: 'split', direction: 'vert-out', duration: 1.5 });
  assert.notEqual(xmlOf(deck, 1), before);
  assert.equal(deck.setTransition(1, null), true);
  assert.equal(xmlOf(deck, 1), before);
  assert.equal(deck.setTransition(1, null), false, 'removing nothing is not a change');
});

test('what PowerPoint wrote is read: speeds, p14:dur choices, a PowerPoint 2010 effect with its fallback, and a timing-only transition', () => {
  const deck = DECK();
  const part = deck.slideParts[0].part;
  const base = xmlOf(deck, 0);
  const put = (block) => deck.pkg.write_(part, Buffer.from(base.replace('</p:clrMapOvr>', `</p:clrMapOvr>${block}`), 'utf8'));

  put('<p:transition spd="med"><p:push dir="r"/></p:transition>');
  assert.deepEqual(pick(deck.transition(0)), { type: 'push', direction: 'r', duration: 0.75, advanceOnClick: true, advanceAfter: null, known: true });

  put('<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14"><p:transition spd="slow" p14:dur="2000"><p:cut/></p:transition></mc:Choice><mc:Fallback><p:transition spd="slow"><p:cut/></p:transition></mc:Fallback></mc:AlternateContent>');
  assert.deepEqual(pick(deck.transition(0)), { type: 'cut', direction: 'smooth', duration: 2, advanceOnClick: true, advanceAfter: null, known: true });

  // A 2010 effect: read as itself, played as its fallback, kept when only its timing changes.
  const prism = '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14"><p:transition spd="med"><p14:prism dir="l"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="med"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>';
  put(prism);
  const t = deck.transition(0);
  assert.equal(t.type, 'prism');
  assert.equal(t.known, false);
  assert.equal(t.fallback, 'fade');
  deck.setTransition(0, { advanceAfter: 4 });
  const kept = xmlOf(deck, 0);
  assert.ok(kept.includes('<p14:prism dir="l"/>'), 'the 2010 effect is kept');
  assert.equal((kept.match(/advTm="4000"/g) || []).length, 2, 'both the choice and the fallback carry the new timing');
  // Picking a classic transition replaces it.
  deck.setTransition(0, { type: 'wipe' });
  assert.doesNotMatch(xmlOf(deck, 0), /prism/);
  assert.equal(deck.transition(0).type, 'wipe');
  assert.equal(deck.transition(0).advanceAfter, 4, 'the advance time carries over');

  // A transition that says only how the slide advances.
  put('<p:transition xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" spd="med" advClick="1"/>');
  assert.deepEqual(pick(deck.transition(0)), { type: 'none', direction: null, duration: 0.75, advanceOnClick: true, advanceAfter: null, known: true });
});

test('the transition sits after p:clrMapOvr and before p:timing, and survives a save and reopen', () => {
  const deck = DECK();
  const part = deck.slideParts[1].part;
  const timing = '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"/></p:par></p:tnLst></p:timing>';
  deck.pkg.write_(part, Buffer.from(xmlOf(deck, 1).replace('</p:clrMapOvr>', `</p:clrMapOvr>${timing}`), 'utf8'));
  deck.setTransition(1, { type: 'push', direction: 'u', duration: 1.1 });
  const xml = xmlOf(deck, 1);
  const at = { ovr: xml.indexOf('</p:clrMapOvr>'), tr: xml.indexOf('<mc:AlternateContent'), timing: xml.indexOf('<p:timing>') };
  assert.ok(at.ovr >= 0 && at.ovr < at.tr && at.tr < at.timing, JSON.stringify(at));

  const reopened = Deck.open(deck.save());
  assert.deepEqual(pick(reopened.transition(1)), { type: 'push', direction: 'u', duration: 1.1, advanceOnClick: true, advanceAfter: null, known: true });
  assert.equal(reopened.transition(0), null);
  assert.ok(reopened.pkg.text(reopened.slideParts[1].part).includes(timing), 'the timing is untouched');

  // The writer on its own, with no slide.
  assert.equal(transitionXml({ type: 'dissolve', duration: 0.5 }), '<p:transition spd="fast"><p:dissolve/></p:transition>');
  assert.equal(readTransition('<p:sld><p:cSld/></p:sld>'), null);
});

function pick(t) {
  const { type, direction, duration, advanceOnClick, advanceAfter, known } = t;
  return { type, direction, duration, advanceOnClick, advanceAfter, known };
}
