// Animations — the engine's half: a slide's main sequence read from the
// `<p:timing>` tree PowerPoint writes (tmRoot → mainSeq → click groups →
// effects, and the build list beside it), and written back the same way:
// effects added, changed, moved and taken out, with PowerPoint's preset
// numbers, triggers grouped into click groups, and everything this does not
// model — an effect it does not know, a trigger's interactive sequence, a
// media node — kept as it was.
import test from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { Deck, buildPptx, readAnimations } from '@rutba/presentation';

const RICH = join(import.meta.dirname, 'fixtures', 'rich');

// A slide with a title, a body and three shapes of our own, ids 4, 5 and 6.
function deckWithShapes() {
  const deck = Deck.open(buildPptx({
    title: 'Animations',
    slides: [
      { layout: 'obj', title: 'Moving parts', body: ['One', 'Two'] },
      { layout: 'obj', title: 'Second', body: ['Three'] },
    ],
  }));
  const a = deck.addShape(0, { preset: 'rect', x: 100, y: 300, w: 200, h: 120 });
  const b = deck.addShape(0, { preset: 'ellipse', x: 400, y: 300, w: 200, h: 120 });
  const c = deck.addTextBox(0, { x: 700, y: 300, w: 200, h: 80, paragraphs: [{ runs: [{ text: 'Words' }] }] });
  return { deck, ids: [String(a), String(b), String(c)] };
}
const xmlOf = (deck, i = 0) => deck.pkg.text(deck.slideParts[i].part);
const timingOf = (deck, i = 0) => /<p:timing>[\s\S]*<\/p:timing>/.exec(xmlOf(deck, i))?.[0] || '';
const pick = (list) => list.map(({ shapeId, kind, effect, direction, trigger, duration, delay, group }) => ({ shapeId, kind, effect, direction, trigger, duration, delay, group }));

// PowerPoint's own tree for three effects: Fly In from the bottom on the
// first click; Fade on the second, with a legacy Basic Zoom (preset 23)
// starting with it — as the showcase deck's slide 7 carries them.
const POWERPOINT_TIMING = '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst><p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="5" presetID="2" presetClass="entr" presetSubtype="4" fill="hold" grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="6" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="SPA"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set><p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="7" dur="500" fill="hold"/><p:tgtEl><p:spTgt spid="SPA"/></p:tgtEl><p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="#ppt_x"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="#ppt_x"/></p:val></p:tav></p:tavLst></p:anim><p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="8" dur="500" fill="hold"/><p:tgtEl><p:spTgt spid="SPA"/></p:tgtEl><p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="1+#ppt_h/2"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="#ppt_y"/></p:val></p:tav></p:tavLst></p:anim></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par><p:par><p:cTn id="9" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="10" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="11" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="12" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="SPB"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set><p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="13" dur="500"/><p:tgtEl><p:spTgt spid="SPB"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par><p:par><p:cTn id="14" presetID="23" presetClass="entr" presetSubtype="16" fill="hold" grpId="0" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="15" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="SPC"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set><p:anim calcmode="lin" valueType="num"><p:cBhvr><p:cTn id="16" dur="500" fill="hold"/><p:tgtEl><p:spTgt spid="SPC"/></p:tgtEl><p:attrNameLst><p:attrName>ppt_w</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:fltVal val="0"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="#ppt_w"/></p:val></p:tav></p:tavLst></p:anim><p:anim calcmode="lin" valueType="num"><p:cBhvr><p:cTn id="17" dur="500" fill="hold"/><p:tgtEl><p:spTgt spid="SPC"/></p:tgtEl><p:attrNameLst><p:attrName>ppt_h</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:fltVal val="0"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="#ppt_h"/></p:val></p:tav></p:tavLst></p:anim></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst><p:bldP spid="SPA" grpId="0" animBg="1"/><p:bldP spid="SPB" grpId="0" animBg="1"/><p:bldP spid="SPC" grpId="0" animBg="1"/></p:bldLst></p:timing>';

function withPowerPointTiming() {
  const { deck, ids } = deckWithShapes();
  const timing = POWERPOINT_TIMING.replaceAll('SPA', ids[0]).replaceAll('SPB', ids[1]).replaceAll('SPC', ids[2]);
  const part = deck.slideParts[0].part;
  deck.pkg.write_(part, Buffer.from(xmlOf(deck).replace('</p:clrMapOvr>', `</p:clrMapOvr>${timing}`), 'utf8'));
  return { deck, ids, timing };
}

test('PowerPoint\'s own timing tree is read as the Animation Pane lists it', () => {
  const { deck, ids } = withPowerPointTiming();
  const list = deck.animations(0);
  assert.deepEqual(pick(list), [
    { shapeId: ids[0], kind: 'entr', effect: 'fly', direction: 'bottom', trigger: 'onClick', duration: 0.5, delay: 0, group: 1 },
    { shapeId: ids[1], kind: 'entr', effect: 'fade', direction: null, trigger: 'onClick', duration: 0.5, delay: 0, group: 2 },
    { shapeId: ids[2], kind: 'entr', effect: 'zoom', direction: null, trigger: 'withPrevious', duration: 0.5, delay: 0, group: 2 },
  ]);
  assert.deepEqual(list.map((e) => e.known), [true, true, false], 'Basic Zoom (23) is read and played, not rewritten');
  assert.deepEqual(deck.slide(0).animations.map((e) => e.effect), ['fly', 'fade', 'zoom']);
  assert.deepEqual(deck.outline().map((o) => o.animated), [true, false]);

  // The showcase deck, as PowerPoint saved it.
  const showcase = join(RICH, 'showcase.pptx');
  if (existsSync(showcase)) {
    const d = Deck.open(readFileSync(showcase));
    const i = d.slideParts.findIndex((s) => d.pkg.text(s.part).includes('<p:timing'));
    assert.ok(i >= 0);
    assert.deepEqual(d.animations(i).map((e) => [e.effect, e.trigger, e.group]), [['fly', 'onClick', 1], ['fade', 'onClick', 2], ['zoom', 'withPrevious', 2]]);
  }
});

test('each effect is written with PowerPoint\'s preset class, id and subtype, and read back as itself', () => {
  const cases = [
    ['entr', 'appear', 1, 0], ['entr', 'fade', 10, 0], ['entr', 'fly', 2, 4], ['entr', 'float', 42, 0], ['entr', 'split', 16, 37], ['entr', 'wipe', 22, 4], ['entr', 'zoom', 53, 16],
    ['emph', 'pulse', 26, 0], ['emph', 'spin', 8, 0], ['emph', 'grow', 6, 0],
    ['exit', 'appear', 1, 0], ['exit', 'fade', 10, 0], ['exit', 'fly', 2, 4], ['exit', 'float', 42, 0], ['exit', 'split', 16, 21], ['exit', 'wipe', 22, 4], ['exit', 'zoom', 53, 16],
  ];
  for (const [kind, effect, id, subtype] of cases) {
    const { deck, ids } = deckWithShapes();
    assert.equal(deck.addAnimation(0, ids[0], { kind, effect }), 0);
    const t = timingOf(deck);
    assert.match(t, new RegExp(`presetID="${id}" presetClass="${kind}" presetSubtype="${subtype}" fill="hold" grpId="0" nodeType="clickEffect"`), `${kind} ${effect}`);
    const [e] = deck.animations(0);
    assert.deepEqual([e.kind, e.effect, e.known], [kind, effect, true], `${kind} ${effect} read back`);
    // Entrances make the shape visible; exits hide it at the end.
    if (kind === 'entr') assert.match(t, /<p:strVal val="visible"\/>/);
    if (kind === 'exit') assert.match(t, /<p:strVal val="hidden"\/>/);
  }
  // The behaviours PowerPoint uses: a fade filter, ppt_x/ppt_y from off the slide, a full turn.
  const { deck, ids } = deckWithShapes();
  deck.addAnimation(0, ids[0], { kind: 'entr', effect: 'fly', direction: 'left' });
  deck.addAnimation(0, ids[1], { kind: 'entr', effect: 'fade' });
  deck.addAnimation(0, ids[1], { kind: 'emph', effect: 'spin' });
  deck.addAnimation(0, ids[2], { kind: 'entr', effect: 'wipe', direction: 'left' });
  const t = timingOf(deck);
  assert.match(t, /presetID="2" presetClass="entr" presetSubtype="8"/, 'Fly In from the left is subtype 8');
  assert.match(t, /<p:attrName>ppt_x<\/p:attrName><\/p:attrNameLst><\/p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="0-#ppt_w\/2"\/>/);
  assert.match(t, /<p:animEffect transition="in" filter="fade">/);
  assert.match(t, /<p:animRot by="21600000">/);
  assert.match(t, /<p:animEffect transition="in" filter="wipe\(right\)">/, 'a wipe from the left moves its edge right');
  assert.deepEqual(deck.animations(0).map((e) => e.direction), ['left', null, 'clockwise', 'left']);
  // A second effect on the same shape takes the next build group.
  assert.match(t, /presetClass="emph" presetSubtype="0" fill="hold" grpId="1"/);
});

test('a fresh sequence is the tree PowerPoint writes: tmRoot, mainSeq, a click group, the build list, ids 1, 2, 3…', () => {
  const { deck, ids } = deckWithShapes();
  deck.addAnimation(0, ids[0], { kind: 'entr', effect: 'fade' });
  const t = timingOf(deck);
  assert.ok(t.startsWith('<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst><p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="5" presetID="10"'), t.slice(0, 400));
  assert.match(t, /<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt\/><\/p:tgtEl><\/p:cond><\/p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt\/><\/p:tgtEl><\/p:cond><\/p:nextCondLst><\/p:seq>/);
  assert.match(t, new RegExp(`<p:bldLst><p:bldP spid="${ids[0]}" grpId="0" animBg="1"/></p:bldLst></p:timing>$`));
  const cTnIds = [...t.matchAll(/<p:cTn id="(\d+)"/g)].map((m) => Number(m[1]));
  assert.deepEqual(cTnIds, cTnIds.map((_, i) => i + 1), 'numbered in document order');
  // A picture or a group has no build entry; a text box does (a shape, not a placeholder).
  deck.addAnimation(0, ids[2], { kind: 'entr', effect: 'appear' });
  assert.match(timingOf(deck), new RegExp(`<p:bldP spid="${ids[2]}" grpId="0" animBg="1"/>`));
});

test('triggers group into click groups: With Previous shares a start, After Previous starts when the one before ends', () => {
  const { deck, ids } = deckWithShapes();
  deck.addAnimation(0, ids[0], { kind: 'entr', effect: 'fade', trigger: 'onClick' }); // 0.5 s
  deck.addAnimation(0, ids[1], { kind: 'entr', effect: 'float', trigger: 'withPrevious' }); // 1 s
  deck.addAnimation(0, ids[2], { kind: 'entr', effect: 'fly', trigger: 'afterPrevious', delay: 0.25 }); // after the float
  deck.addAnimation(0, ids[0], { kind: 'exit', effect: 'fade', trigger: 'onClick' });
  const list = deck.animations(0);
  assert.deepEqual(list.map((e) => [e.trigger, e.group]), [['onClick', 1], ['withPrevious', 1], ['afterPrevious', 1], ['onClick', 2]]);
  const t = timingOf(deck);
  // Two click groups under the main sequence.
  assert.equal((t.match(/<p:cond delay="indefinite"\/>/g) || []).length, 2);
  // The first group's second inner group starts when the float ends: 1000 ms.
  assert.match(t, /<p:par><p:cTn id="\d+" fill="hold"><p:stCondLst><p:cond delay="1000"\/><\/p:stCondLst>/);
  assert.match(t, /nodeType="withEffect"/);
  assert.match(t, /nodeType="afterEffect"><p:stCondLst><p:cond delay="250"\/>/, 'the effect\'s own delay is on its own cTn');
  assert.equal(list[2].delay, 0.25);

  // An effect that starts with the slide: an onBegin condition on the main sequence.
  const second = deckWithShapes();
  second.deck.addAnimation(0, second.ids[0], { kind: 'entr', effect: 'fade', trigger: 'afterPrevious' });
  const t2 = timingOf(second.deck);
  assert.match(t2, /<p:stCondLst><p:cond delay="indefinite"\/><p:cond evt="onBegin" delay="0"><p:tn val="2"\/><\/p:cond><\/p:stCondLst>/);
  assert.deepEqual(second.deck.animations(0).map((e) => [e.trigger, e.group]), [['afterPrevious', 0]]);
});

test('Move Earlier and Move Later reorder the sequence, each effect keeping its own trigger', () => {
  const { deck, ids } = deckWithShapes();
  deck.addAnimation(0, ids[0], { kind: 'entr', effect: 'fade' });
  deck.addAnimation(0, ids[1], { kind: 'entr', effect: 'fly' });
  deck.addAnimation(0, ids[2], { kind: 'entr', effect: 'zoom', trigger: 'withPrevious' });
  assert.equal(deck.moveAnimation(0, 1, 'earlier'), 0);
  assert.deepEqual(deck.animations(0).map((e) => e.effect), ['fly', 'fade', 'zoom']);
  assert.equal(deck.moveAnimation(0, 0, 'later'), 1);
  assert.deepEqual(deck.animations(0).map((e) => e.effect), ['fade', 'fly', 'zoom']);
  assert.equal(deck.moveAnimation(0, 2, 0), 0);
  assert.deepEqual(deck.animations(0).map((e) => [e.effect, e.trigger]), [['zoom', 'withPrevious'], ['fade', 'onClick'], ['fly', 'onClick']]);
  assert.deepEqual(deck.animations(0).map((e) => e.group), [0, 1, 2], 'a With Previous moved first starts with the slide');
  // Moving past either end is not a change.
  assert.equal(deck.moveAnimation(0, 0, 'earlier'), 0);
});

test('taking effects out leaves a valid tree: an emptied click group goes, and the last effect takes the timing with it', () => {
  const { deck, ids } = deckWithShapes();
  const before = xmlOf(deck);
  deck.addAnimation(0, ids[0], { kind: 'entr', effect: 'fade' });
  deck.addAnimation(0, ids[1], { kind: 'entr', effect: 'wipe' });
  deck.addAnimation(0, ids[2], { kind: 'entr', effect: 'split', trigger: 'withPrevious' });
  deck.removeAnimation(0, 0);
  let t = timingOf(deck);
  assert.equal((t.match(/<p:cond delay="indefinite"\/>/g) || []).length, 1, 'one click group left');
  assert.doesNotMatch(t, new RegExp(`spid="${ids[0]}"`), 'the shape\'s build entry went with its effect');
  assert.deepEqual(deck.animations(0).map((e) => [e.effect, e.group]), [['wipe', 1], ['split', 1]]);
  deck.removeAnimation(0, 1);
  deck.removeAnimation(0, 0);
  assert.equal(timingOf(deck), '');
  assert.equal(xmlOf(deck), before, 'the slide is byte for byte as it was');

  // Deleting a shape takes its effects out; the gallery's None does too.
  deck.addAnimation(0, ids[0], { kind: 'entr', effect: 'fade' });
  deck.addAnimation(0, ids[1], { kind: 'entr', effect: 'fade' });
  deck.addAnimation(0, ids[1], { kind: 'exit', effect: 'fade' });
  deck.removeShape(0, ids[0]);
  assert.deepEqual(deck.animations(0).map((e) => e.shapeId), [ids[1], ids[1]]);
  assert.equal(deck.removeShapeAnimations(0, ids[1]), true);
  assert.equal(timingOf(deck), '');
});

test('Start, Duration, Delay, Effect Options and another effect from the gallery change one effect', () => {
  const { deck, ids } = deckWithShapes();
  deck.addAnimation(0, ids[0], { kind: 'entr', effect: 'fade' });
  deck.addAnimation(0, ids[1], { kind: 'entr', effect: 'fly' });
  deck.setAnimation(0, 1, { trigger: 'withPrevious' });
  assert.deepEqual(deck.animations(0).map((e) => [e.trigger, e.group]), [['onClick', 1], ['withPrevious', 1]]);
  deck.setAnimation(0, 1, { duration: 2, delay: 0.5 });
  let e = deck.animations(0)[1];
  assert.deepEqual([e.duration, e.delay, e.effect, e.direction], [2, 0.5, 'fly', 'bottom']);
  assert.match(timingOf(deck), /<p:cTn id="\d+" dur="2000" fill="hold"\/>/);
  deck.setAnimation(0, 1, { direction: 'right' });
  e = deck.animations(0)[1];
  assert.deepEqual([e.direction, e.duration, e.delay, e.trigger], ['right', 2, 0.5, 'withPrevious'], 'Effect Options keeps the timing');
  assert.match(timingOf(deck), /presetID="2" presetClass="entr" presetSubtype="2"/);
  // Another effect from the gallery: same shape, same build group, its own gallery length.
  deck.setAnimation(0, 0, { effect: 'zoom' });
  e = deck.animations(0)[0];
  assert.deepEqual([e.effect, e.shapeId, e.duration], ['zoom', ids[0], 0.5]);
  assert.equal((timingOf(deck).match(new RegExp(`<p:bldP spid="${ids[0]}"`, 'g')) || []).length, 1);
});

test('what this does not model is kept: an unknown effect, a trigger\'s interactive sequence and a media node', () => {
  const { deck, ids } = deckWithShapes();
  const path = `<p:par><p:cTn id="5" presetID="42" presetClass="path" presetSubtype="0" accel="50000" decel="50000" fill="hold" grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animMotion origin="layout" path="M 0 0 L 0.25 0.1 E" pathEditMode="relative" rAng="0" ptsTypes="AA"><p:cBhvr><p:cTn id="6" dur="2000" fill="hold"/><p:tgtEl><p:spTgt spid="${ids[2]}"/></p:tgtEl><p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr><p:rCtr x="12500" y="5000"/></p:animMotion></p:childTnLst></p:cTn></p:par>`;
  const interactive = `<p:seq concurrent="1" nextAc="seek"><p:cTn id="7" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq"><p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${ids[1]}"/></p:tgtEl></p:cond></p:stCondLst><p:endSync evt="end" delay="0"><p:rtn val="all"/></p:endSync><p:childTnLst><p:par><p:cTn id="8" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="9" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="10" presetID="1" presetClass="exit" presetSubtype="0" fill="hold" grpId="1" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="11" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${ids[1]}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="hidden"/></p:to></p:set></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn><p:nextCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${ids[1]}"/></p:tgtEl></p:cond></p:nextCondLst></p:seq>`;
  const media = `<p:audio><p:cMediaNode vol="80000"><p:cTn id="12" fill="hold" display="0"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${ids[1]}"/></p:tgtEl></p:cMediaNode></p:audio>`;
  const timing = `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst><p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${path}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>${interactive}${media}</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst><p:bldP spid="${ids[2]}" grpId="0" animBg="1"/><p:bldP spid="${ids[1]}" grpId="1" animBg="1"/></p:bldLst></p:timing>`;
  const part = deck.slideParts[0].part;
  deck.pkg.write_(part, Buffer.from(xmlOf(deck).replace('</p:clrMapOvr>', `</p:clrMapOvr>${timing}`), 'utf8'));

  const [only] = deck.animations(0);
  assert.deepEqual([only.kind, only.name, only.known, only.duration], ['path', 'Motion Path', false, 2]);
  const noIds = (s) => s.replace(/ id="\d+"/g, ' id=""').replace(/nodeType="\w+Effect"/g, 'nodeType=""');

  // An effect added in front of it, the path moved about, its start changed.
  deck.addAnimation(0, ids[0], { kind: 'entr', effect: 'fade' }, 0);
  deck.moveAnimation(0, 1, 'earlier');
  deck.setAnimation(0, 0, { trigger: 'withPrevious' });
  const t = timingOf(deck);
  assert.ok(noIds(t).includes(noIds(path)), 'the motion path is kept, byte for byte but its ids and node type');
  assert.ok(noIds(t).includes(noIds(interactive)), 'the trigger\'s interactive sequence is kept');
  assert.ok(noIds(t).includes(noIds(media)), 'the media node is kept');
  assert.match(t, new RegExp(`<p:bldP spid="${ids[1]}" grpId="1" animBg="1"/>`), 'the build entry the trigger\'s effect uses is kept');
  assert.deepEqual(deck.animations(0).map((e) => [e.kind, e.trigger]), [['path', 'withPrevious'], ['entr', 'onClick']]);
  // Every time node numbered once, in order.
  const cTnIds = [...t.matchAll(/<p:cTn id="(\d+)"/g)].map((m) => Number(m[1]));
  assert.deepEqual(cTnIds, cTnIds.map((_, i) => i + 1));
  // A duration on an effect this does not write stretches its behaviours instead.
  deck.setAnimation(0, 0, { duration: 3 });
  assert.equal(deck.animations(0)[0].duration, 3);
  assert.match(timingOf(deck), /<p:animMotion origin="layout" path="M 0 0 L 0\.25 0\.1 E"/);
});

test('the sequence survives a save and reopen, and sits after the transition, before the slide\'s extLst', () => {
  const { deck, ids } = deckWithShapes();
  deck.setTransition(0, { type: 'fade' });
  deck.addAnimation(0, ids[0], { kind: 'entr', effect: 'fly', direction: 'top-left' });
  deck.addAnimation(0, ids[1], { kind: 'emph', effect: 'pulse', trigger: 'afterPrevious', delay: 1 });
  deck.addAnimation(0, ids[2], { kind: 'exit', effect: 'wipe', direction: 'right' });
  const before = pick(deck.animations(0));
  const reopened = Deck.open(deck.save());
  assert.deepEqual(pick(reopened.animations(0)), before);
  assert.deepEqual(before.map((e) => [e.effect, e.direction, e.trigger, e.group]), [['fly', 'top-left', 'onClick', 1], ['pulse', null, 'afterPrevious', 1], ['wipe', 'right', 'onClick', 2]]);
  const xml = xmlOf(reopened);
  const at = { ovr: xml.indexOf('</p:clrMapOvr>'), tr: xml.indexOf('<p:transition'), timing: xml.indexOf('<p:timing>') };
  assert.ok(at.ovr < at.tr && at.tr < at.timing, JSON.stringify(at));

  // With a root extension list, the timing goes before it.
  const part = deck.slideParts[1].part;
  const ext = '<p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"><p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1"/></p:ext></p:extLst>';
  deck.pkg.write_(part, Buffer.from(xmlOf(deck, 1).replace('</p:sld>', `${ext}</p:sld>`), 'utf8'));
  const titleId = /<p:cNvPr id="(\d+)"/.exec(xmlOf(deck, 1).split('<p:sp>')[1])[1];
  deck.addAnimation(1, titleId, { kind: 'entr', effect: 'appear' });
  const x1 = xmlOf(deck, 1);
  assert.ok(x1.indexOf('</p:timing>') < x1.indexOf(ext) && x1.endsWith(`${ext}</p:sld>`));
  // A title placeholder with words builds without animBg.
  assert.match(x1, new RegExp(`<p:bldP spid="${titleId}" grpId="0"/>`));
  assert.deepEqual(readAnimations(x1).map((e) => e.effect), ['appear']);
});

test('an effect on one paragraph keeps its paragraph when changed, and a sequence laid out another way is read but not rewritten', () => {
  const { deck, ids } = deckWithShapes();
  deck.addAnimation(0, ids[2], { kind: 'entr', effect: 'fade', paragraph: 0 });
  assert.match(timingOf(deck), new RegExp(`<p:spTgt spid="${ids[2]}"><p:txEl><p:pRg st="0" end="0"/></p:txEl></p:spTgt>`));
  deck.setAnimation(0, 0, { effect: 'fly', direction: 'left' });
  const [e] = deck.animations(0);
  assert.deepEqual([e.effect, e.direction, e.paragraph], ['fly', 'left', 0]);
  assert.equal((timingOf(deck).match(/<p:pRg st="0" end="0"\/>/g) || []).length, 3, 'the set and both motions target the paragraph');

  // A main sequence whose click group holds an effect directly, with no
  // inner group, is not PowerPoint's shape: it is read, and left alone.
  const odd = deckWithShapes();
  const timing = `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst><p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst></p:cTn></p:par></p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
  const part = odd.deck.slideParts[0].part;
  odd.deck.pkg.write_(part, Buffer.from(xmlOf(odd.deck).replace('</p:clrMapOvr>', `</p:clrMapOvr>${timing}`), 'utf8'));
  assert.deepEqual(odd.deck.animations(0), []);
  assert.throws(() => odd.deck.addAnimation(0, odd.ids[0], { effect: 'fade' }), /cannot rewrite/);
  // Deleting a shape there still works; the odd timing stays as it was.
  odd.deck.removeShape(0, odd.ids[0]);
  assert.ok(timingOf(odd.deck).includes('<p:cTn id="3" fill="hold">'));
  // A duplicated slide carries its animations with it.
  const dup = deck.duplicateSlide(0);
  assert.deepEqual(deck.animations(dup).map((x) => x.effect), ['fly']);
});
