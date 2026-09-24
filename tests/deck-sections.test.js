// Sections in a deck — the engine's half: PowerPoint's p14:sectionLst on the
// presentation part, read as it comes, written as PowerPoint writes it, kept
// in step when a slide is added, moved or taken away, and taken off again.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';
import { OoxmlPackage } from '@rutba/ooxml/package';

const DECK = buildPptx({
  title: 'Sections',
  slides: [
    { layout: 'title', title: 'Opening', body: 'A deck in parts' },
    { layout: 'obj', title: 'First point', body: ['One'] },
    { layout: 'obj', title: 'Second point', body: ['Two'] },
    { layout: 'obj', title: 'Closing', body: ['Thanks'] },
  ],
});
const names = (deck) => deck.sections().map((s) => [s.name, s.slides]);
const presXml = (deck) => OoxmlPackage.read(deck.save()).text('ppt/presentation.xml');
const GUID = '\\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\\}';

test('a section is added before a slide, the slides before it in a Default Section, written as PowerPoint writes the list, read back, renamed and removed', () => {
  const deck = Deck.open(DECK);
  assert.deepEqual(deck.sections(), [], 'a built deck has no sections');
  assert.equal(deck.sectionOf(2), -1);
  assert.equal(deck.outline()[2].section, null);

  assert.equal(deck.addSection(2, 'The points'), 1, 'the new section is the second');
  assert.deepEqual(names(deck), [['Default Section', [0, 1]], ['The points', [2, 3]]]);
  assert.equal(deck.sectionOf(3), 1);
  assert.deepEqual(deck.outline().map((o) => o.section), [0, 0, 1, 1]);
  assert.ok(deck.dirty);

  // The file: one p:ext with PowerPoint's URI, the p14 namespace on the list,
  // a GUID and the sldId ids on each section, at the end of the presentation.
  const xml = presXml(deck);
  assert.match(xml, new RegExp(
    '<p:extLst><p:ext uri="\\{521415D9-36F7-43E2-AB2F-B90AF26B5E84\\}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">' +
    `<p14:section name="Default Section" id="${GUID}"><p14:sldIdLst><p14:sldId id="256"/><p14:sldId id="257"/></p14:sldIdLst></p14:section>` +
    `<p14:section name="The points" id="${GUID}"><p14:sldIdLst><p14:sldId id="258"/><p14:sldId id="259"/></p14:sldIdLst></p14:section>` +
    '</p14:sectionLst></p:ext></p:extLst></p:presentation>$'));
  assert.equal((xml.match(/<p:extLst>/g) || []).length, 1);
  assert.deepEqual(names(Deck.open(deck.save())), [['Default Section', [0, 1]], ['The points', [2, 3]]], 'read back through a save');

  // A slide that starts a section starts it still.
  assert.equal(deck.addSection(2, 'Again'), 1);
  assert.deepEqual(names(deck), [['Default Section', [0, 1]], ['The points', [2, 3]]]);

  // A section split in two; the one it came from keeps the slides before.
  assert.equal(deck.addSection(3, 'The end'), 2);
  assert.deepEqual(names(deck), [['Default Section', [0, 1]], ['The points', [2]], ['The end', [3]]]);
  const ids = () => [...presXml(deck).matchAll(/<p14:section name="([^"]*)" id="([^"]*)"/g)].map((m) => m[2]);
  assert.equal(new Set(ids()).size, 3, 'three GUIDs, all different');

  assert.equal(deck.renameSection(1, ' Points & <arguments> '), true);
  assert.equal(deck.sections()[1].name, 'Points & <arguments>', 'trimmed, and the XML-hostile name survives');
  assert.match(presXml(deck), /<p14:section name="Points &amp; &lt;arguments&gt;"/);
  assert.equal(deck.renameSection(1, 'Points & <arguments>'), false, 'the same name is no change');
  assert.equal(deck.sections()[1].name, 'Points & <arguments>');
  assert.equal(deck.renameSection(1, '   ').valueOf(), true);
  assert.equal(deck.sections()[1].name, 'Untitled Section', 'a blank name is the default one');
  assert.throws(() => deck.renameSection(7, 'x'), /no section/);

  // Removed: its slides join the section before; the first section's join the one after.
  assert.equal(deck.removeSection(2), true);
  assert.deepEqual(names(deck), [['Default Section', [0, 1]], ['Untitled Section', [2, 3]]]);
  assert.equal(deck.removeSection(0), true);
  assert.deepEqual(names(deck), [['Untitled Section', [0, 1, 2, 3]]]);
  // The last one takes the list with it.
  assert.equal(deck.removeSection(0), true);
  assert.deepEqual(deck.sections(), []);
  assert.doesNotMatch(presXml(deck), /sectionLst|<p:extLst>/);
  assert.match(presXml(deck), /<\/p:defaultTextStyle><\/p:presentation>$/, 'the extLst went with it');
  assert.equal(deck.removeAllSections(), false, 'nothing to remove');

  assert.equal(deck.addSection(0, 'All of it'), 0, 'a section at the first slide needs no Default Section before it');
  assert.deepEqual(names(deck), [['All of it', [0, 1, 2, 3]]]);
  assert.equal(deck.removeAllSections(), true);
  assert.deepEqual(names(Deck.open(deck.save())), []);
  assert.throws(() => deck.addSection(9), /no slide/);
});

test('a slide added, duplicated, moved or taken away keeps the sections in step, and one added to a file with an extLst already sits inside it', () => {
  const deck = Deck.open(DECK);
  deck.addSection(2, 'The points');

  // A new slide after slide 1 joins slide 1's section; after slide 3, the other.
  assert.equal(deck.insertSlide(1, { layout: 'obj', title: 'In between', body: ['x'] }), 2);
  assert.deepEqual(names(deck), [['Default Section', [0, 1, 2]], ['The points', [3, 4]]]);
  assert.equal(deck.duplicateSlide(4), 5);
  assert.deepEqual(names(deck), [['Default Section', [0, 1, 2]], ['The points', [3, 4, 5]]]);

  // Moved down into the second section, it belongs there; moved back up, here.
  assert.equal(deck.moveSlide(1, 4), true);
  assert.deepEqual(names(deck), [['Default Section', [0, 1]], ['The points', [2, 3, 4, 5]]]);
  assert.equal(deck.moveSlide(4, 1), true);
  assert.deepEqual(names(deck), [['Default Section', [0, 1, 2]], ['The points', [3, 4, 5]]]);
  // Moved to the very front, it joins the first section.
  assert.equal(deck.moveSlide(5, 0), true);
  assert.deepEqual(names(deck), [['Default Section', [0, 1, 2, 3]], ['The points', [4, 5]]]);

  assert.equal(deck.removeSlide(4), true);
  assert.deepEqual(names(deck), [['Default Section', [0, 1, 2, 3]], ['The points', [4]]]);
  assert.equal(deck.removeSlide(4), true);
  assert.deepEqual(names(deck), [['Default Section', [0, 1, 2, 3]], ['The points', []]], 'a section with no slides stays, as PowerPoint keeps one');
  const xml = presXml(deck);
  assert.match(xml, /<p14:section name="The points" id="[^"]+"><p14:sldIdLst\/><\/p14:section>/);
  assert.deepEqual(names(Deck.open(deck.save())), [['Default Section', [0, 1, 2, 3]], ['The points', []]]);

  // A presentation part with an extLst of its own keeps what is in it.
  const pkg = OoxmlPackage.read(DECK);
  const other = '<p:extLst><p:ext uri="{EFAFB233-063F-42B5-8137-9DF3F51BA10A}"><p15:sldGuideLst xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main"/></p:ext></p:extLst>';
  pkg.write_('ppt/presentation.xml', Buffer.from(pkg.text('ppt/presentation.xml').replace('</p:presentation>', `${other}</p:presentation>`), 'utf8'));
  const kept = new Deck(pkg);
  kept.addSection(1, 'Rest');
  const keptXml = presXml(kept);
  assert.match(keptXml, /<p:extLst><p:ext uri="\{EFAFB233[^"]*"><p15:sldGuideLst[^>]*\/><\/p:ext><p:ext uri="\{521415D9[^"]*"><p14:sectionLst/);
  assert.equal((keptXml.match(/<p:extLst>/g) || []).length, 1);
  kept.removeAllSections();
  assert.match(presXml(kept), /<p:extLst><p:ext uri="\{EFAFB233[^"]*"><p15:sldGuideLst[^>]*\/><\/p:ext><\/p:extLst><\/p:presentation>$/, 'the other extension stays');
});

test('a list PowerPoint wrote is read, a slide it names that is not in the show is set aside, and a slide it leaves out joins the section before it', () => {
  const pkg = OoxmlPackage.read(DECK);
  const list =
    '<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">' +
    '<p14:section name="Start &amp; end" id="{0D8B4C6E-1F2A-4B3C-9D4E-5F6A7B8C9D0E}"><p14:sldIdLst><p14:sldId id="256"/><p14:sldId id="999"/></p14:sldIdLst></p14:section>' +
    '<p14:section name="Middle" id="{1D8B4C6E-1F2A-4B3C-9D4E-5F6A7B8C9D0E}"><p14:sldIdLst><p14:sldId id="258"/></p14:sldIdLst></p14:section>' +
    '</p14:sectionLst></p:ext></p:extLst>';
  pkg.write_('ppt/presentation.xml', Buffer.from(pkg.text('ppt/presentation.xml').replace('</p:presentation>', `${list}</p:presentation>`), 'utf8'));
  const deck = new Deck(pkg);
  assert.deepEqual(names(deck), [['Start & end', [0]], ['Middle', [2]]], 'slide 999 is not in the show; slides 257 and 259 are in no section');
  assert.equal(deck.sections()[0].id, '{0D8B4C6E-1F2A-4B3C-9D4E-5F6A7B8C9D0E}');
  assert.deepEqual(deck.outline().map((o) => o.section), [0, null, 1, null]);
  assert.equal(deck.dirty, false, 'reading changes nothing');

  // Any slide edit brings the list into step: the unlisted slides join the
  // section of the slide before them, and the one not in the show is dropped.
  deck.insertSlide(3, { layout: 'blank' });
  assert.deepEqual(names(deck), [['Start & end', [0, 1]], ['Middle', [2, 3, 4]]]);
  assert.doesNotMatch(presXml(deck), /id="999"/);
  assert.match(presXml(deck), /id="\{0D8B4C6E-1F2A-4B3C-9D4E-5F6A7B8C9D0E\}"/, 'PowerPoint’s own GUID is kept');

  // A section added at an unlisted slide of an untouched file first brings the list into step.
  const fresh = OoxmlPackage.read(DECK);
  fresh.write_('ppt/presentation.xml', Buffer.from(fresh.text('ppt/presentation.xml').replace('</p:presentation>', `${list}</p:presentation>`), 'utf8'));
  const again = new Deck(fresh);
  assert.deepEqual(names(again), [['Start & end', [0]], ['Middle', [2]]]);
  assert.equal(again.addSection(3, 'Last'), 2);
  assert.deepEqual(names(again), [['Start & end', [0, 1]], ['Middle', [2]], ['Last', [3]]]);
  assert.equal(again.addSection(1, 'Second'), 1, 'a slide a file left out of every section, once placed, starts one');
  assert.deepEqual(names(again), [['Start & end', [0]], ['Second', [1]], ['Middle', [2]], ['Last', [3]]]);
});
