// View → Notes Master and Handout Master — the engine's half.
//
// A deck with neither gets the master PowerPoint makes: its placeholders
// laid out for the notes size, its own theme, related from the presentation
// and listed after the slide masters, in the schema's order. A notes page
// written for a slide belongs to the notes master and to its slide. The
// Placeholders boxes take one off and put it back; the orientation turns
// the notes size and lays the placeholders out again; and the master is a
// page of the notes' size the stage can draw.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx, renderSlide } from '@rutba/presentation';

const deck2 = () => Deck.open(buildPptx({ title: 'Notes', slides: [{ layout: 'title', title: 'One' }, { layout: 'obj', title: 'Two' }] }));

test('a deck without a notes master gets the one PowerPoint makes, listed and related as PowerPoint lists it', () => {
  const deck = deck2();
  assert.equal(deck.masterFor('notes'), null);
  const part = deck.ensureMaster('notes');
  assert.equal(part, 'ppt/notesMasters/notesMaster1.xml');
  assert.equal(deck.ensureMaster('notes'), part, 'made once');
  const xml = deck.pkg.text(part);
  assert.match(xml, /^<\?xml[^>]*\?><p:notesMaster\b/);
  for (const type of ['hdr', 'dt', 'sldImg', 'body', 'ftr', 'sldNum']) assert.match(xml, new RegExp(`<p:ph type="${type}"`));
  assert.match(xml, /<p:notesStyle><a:lvl1pPr marL="0"[^>]*><a:defRPr sz="1200"/);
  assert.match(deck.pkg.contentTypeOf(part), /presentationml\.notesMaster\+xml$/);
  const theme = [...deck.pkg.rels(part)].find((r) => /\/theme$/.test(r.Type));
  assert.ok(theme && deck.pkg.has('ppt/' + theme.Target.replace('../', '')), 'its own theme');
  const px = deck.pkg.text('ppt/presentation.xml');
  assert.match(px, /<\/p:sldMasterIdLst><p:notesMasterIdLst><p:notesMasterId r:id="rId\d+"\/><\/p:notesMasterIdLst>/);
  // The handout master after it.
  deck.ensureMaster('handout');
  assert.match(deck.pkg.text('ppt/presentation.xml'), /<\/p:notesMasterIdLst><p:handoutMasterIdLst><p:handoutMasterId r:id="rId\d+"\/><\/p:handoutMasterIdLst>/);
  assert.deepEqual(deck.masterPlaceholders(deck.masterFor('handout')), { hdr: true, dt: true, ftr: true, sldNum: true });
});

test('a notes page belongs to the notes master and to its slide', () => {
  const deck = deck2();
  deck.setNotes(1, 'Speak slowly.');
  const notes = [...deck.pkg.rels('ppt/slides/slide2.xml')].find((r) => /notesSlide$/.test(r.Type));
  const part = 'ppt/notesSlides/' + notes.Target.split('/').pop();
  const rels = [...deck.pkg.rels(part)];
  assert.ok(rels.some((r) => /\/notesMaster$/.test(r.Type) && r.Target === '../notesMasters/notesMaster1.xml'));
  assert.ok(rels.some((r) => /\/slide$/.test(r.Type) && r.Target === '../slides/slide2.xml'));
  assert.equal(deck.slide(1).notes, 'Speak slowly.');
  // A second edit rewrites the page; the master stays one.
  deck.setNotes(1, 'Speak slowly, then pause.');
  assert.equal(deck.pkg.partNames().filter((n) => /^ppt\/notesMasters\/notesMaster\d+\.xml$/.test(n)).length, 1);
});

test('a placeholder comes off and goes back, and the page turns with its placeholders laid out again', () => {
  const deck = deck2();
  const part = deck.ensureMaster('notes');
  assert.equal(deck.setMasterPlaceholder(part, 'hdr', false), true);
  assert.equal(deck.masterPlaceholders(part).hdr, false);
  assert.equal(deck.setMasterPlaceholder(part, 'hdr', false), false, 'already off');
  deck.setMasterPlaceholder(part, 'hdr', true);
  assert.equal(deck.masterPlaceholders(part).hdr, true);
  assert.throws(() => deck.setMasterPlaceholder(deck.ensureMaster('handout'), 'body', true), /no body placeholder/);

  const before = deck.notesSize;
  assert.ok(before.cy > before.cx, 'portrait to start');
  assert.equal(deck.setNotesOrientation(true), false, 'already portrait');
  assert.equal(deck.setNotesOrientation(false), true);
  const after = deck.notesSize;
  assert.deepEqual([after.cx, after.cy], [before.cy, before.cx]);
  const ftr = /<p:sp\b(?:(?!<\/p:sp>)[\s\S])*?type="ftr"[\s\S]*?<a:off x="(\d+)" y="(\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"/.exec(deck.pkg.text(part));
  assert.equal(Number(ftr[2]) + Number(ftr[4]), after.cy, 'the footer stands on the turned page\'s foot');
});

test('the notes master is a page of the notes size the stage draws, its notes in their own text style', () => {
  const deck = deck2();
  const part = deck.ensureMaster('notes');
  const scene = deck.partScene(part);
  assert.deepEqual([scene.size.cx, scene.size.cy], [deck.notesSize.cx, deck.notesSize.cy]);
  const body = scene.shapes.find((s) => s.placeholder?.type === 'body');
  assert.ok(body, 'the notes placeholder');
  const svg = renderSlide(scene, { width: 600, placeholderFrames: true });
  assert.match(svg, /Click to edit Master text styles/);
  assert.match(svg, /stroke="#000000"/, 'the slide image, outlined');
});
