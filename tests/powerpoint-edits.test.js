// Slides held to PowerPoint (tools/powerpoint-check.mjs built the deck there,
// made the same table edits, and opened what we write).
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Deck } from '@rutba/presentation';
import { OoxmlPackage } from '@rutba/ooxml/package';
import { applyTableEdit, spansOf, COLUMN_EDITS, ROW_EDITS } from '../tools/powerpoint-check.mjs';

const bytes = fs.readFileSync(new URL('./fixtures/powerpoint-edits.pptx', import.meta.url));
const answers = JSON.parse(fs.readFileSync(new URL('./fixtures/powerpoint-edits.json', import.meta.url), 'utf8'));
const tableOf = (deck) => deck.slide(1).shapes.find((s) => s.kind === 'table' || s.table);

test('the table PowerPoint wrote reads the same here', () => {
  assert.deepEqual(spansOf(bytes), answers.base);
});

[...COLUMN_EDITS.map((e) => ['col', e]), ...ROW_EDITS.map((e) => ['row', e])].forEach(([axis, edit], i) => {
  test(`as PowerPoint: ${edit.name}`, () => {
    assert.deepEqual(spansOf(applyTableEdit(bytes, axis, edit)), answers.edits[i].powerpoint);
  });
});

test('a table\'s frame moves, in its own p:xfrm', () => {
  const deck = Deck.open(bytes);
  const id = tableOf(deck).id;
  deck.setGeometry(1, id, { x: 100, y: 120, w: 500, h: 200 });
  const xml = deck.pkg.text(deck.slideParts[1].part);
  assert.match(xml, /<p:xfrm><a:off x="952500" y="1143000"\/><a:ext cx="4762500" cy="1905000"\/><\/p:xfrm>/);
});

test('a copied slide has its own notes: changing one leaves the other', () => {
  const deck = Deck.open(bytes);
  const copy = deck.duplicateSlide(0);
  deck.setNotes(copy, 'Only the copy');
  const again = Deck.open(deck.save());
  const notes = again.outline().map((s) => s.notes);
  assert.match(notes[0], /^Notes of the first slide/);
  assert.match(notes[copy], /^Only the copy/);
});

const allText = (buf) => {
  const pkg = OoxmlPackage.read(buf);
  return pkg.partNames().filter((n) => /\.(xml|rels)$/.test(n)).map((n) => pkg.text(n)).join('\n');
};

test('a deleted slide does not travel with the file: words, notes or title', () => {
  assert.match(allText(bytes), /Secret plan/);
  const deck = Deck.open(bytes);
  deck.removeSlide(2);
  const saved = deck.save();
  const text = allText(saved);
  assert.doesNotMatch(text, /Words nobody should keep/);
  assert.doesNotMatch(text, /Secret notes/);
  assert.doesNotMatch(text, /Secret plan/);
  const pkg = OoxmlPackage.read(saved);
  for (const part of pkg.partNames()) {
    if (!part.endsWith('.rels')) continue;
    const owner = part === '_rels/.rels' ? '' : part.replace(/_rels\/([^/]+)\.rels$/, '$1');
    for (const r of pkg.rels(owner)) {
      if (r.TargetMode === 'External') continue;
      assert.ok(pkg.has(OoxmlPackage.resolveTarget(owner, r.Target)), `${part} points at ${r.Target}, which is gone`);
    }
  }
  assert.doesNotMatch(pkg.text('ppt/slides/slide1.xml'), /<a:hlinkClick\b[^>]*action="ppaction:\/\/hlinksldjump"/, 'the link to it goes too');
});

test('undo brings a deleted slide back, whole', () => {
  const deck = Deck.open(bytes);
  const before = deck.snapshot();
  deck.removeSlide(2);
  deck.pushUndo(before);
  deck.undo();
  assert.equal(deck.slideCount, 3);
  assert.match(deck.outline()[2].notes, /^Secret notes/);
  assert.match(allText(deck.save()), /Words nobody should keep/);
});

const bodyOf = (deck) => deck.slide(0).shapes.find((s) => s.text?.paragraphs?.length > 1);
const bodyXml = (deck, id) => {
  const xml = deck.pkg.text(deck.slideParts[0].part);
  return [...xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)].map((m) => m[0]).find((s) => s.includes(`<p:cNvPr id="${id}"`));
};

test('editing words keeps what the reader has no field for: theme colour, Urdu face, slide number, paragraph end', () => {
  const deck = Deck.open(bytes);
  const shape = bodyOf(deck);
  const before = bodyXml(deck, shape.id);
  assert.match(before, /<a:schemeClr val="accent1"/, 'the fixture has them');
  assert.match(before, /<a:cs typeface="Noto Nastaliq Urdu"/);
  assert.match(before, /<a:fld\b[^>]*type="slidenum"/);
  const fieldId = /<a:fld\b[^>]*\bid="([^"]*)"/.exec(before)[1];
  deck.setText(0, shape.id, shape.text.paragraphs.map((p) => ({ ...p, runs: p.runs.map((r) => (r.field || r.break ? r : { ...r, text: r.text.replace('First', 'Changed') })) })));
  const after = bodyXml(deck, shape.id);
  assert.match(after, /Changed line of words/);
  assert.match(after, /<a:schemeClr val="accent1"/);
  assert.match(after, /<a:cs typeface="Noto Nastaliq Urdu"/);
  assert.match(after, new RegExp(`<a:fld\\b[^>]*id="${fieldId.replace(/[{}]/g, '\\$&')}"[^>]*type="slidenum"`));
  assert.equal((after.match(/<a:endParaRPr\b/g) || []).length, (before.match(/<a:endParaRPr\b/g) || []).length);
});

test('Replace keeps them too', () => {
  const deck = Deck.open(bytes);
  assert.equal(deck.replaceText('First line', 'Opening line'), 1);
  const after = bodyXml(deck, bodyOf(deck).id);
  assert.match(after, /Opening line/);
  assert.match(after, /<a:schemeClr val="accent1"/);
  assert.match(after, /<a:fld\b[^>]*type="slidenum"/);
});

test('a change of look keeps the rest: bold on a theme-coloured run stays theme-coloured', () => {
  const deck = Deck.open(bytes);
  const shape = bodyOf(deck);
  deck.setText(0, shape.id, shape.text.paragraphs.map((p, i) => (i === 0 ? { ...p, runs: p.runs.map((r) => ({ ...r, bold: true })) } : p)));
  const after = bodyXml(deck, shape.id);
  assert.match(after, /b="1"[^>]*>\s*<a:solidFill><a:schemeClr val="accent1"/);
});

test('a new slide takes the layout of the slide before it', () => {
  const deck = Deck.open(bytes);
  const layoutOf = (i) => deck.pkg.rels(deck.slideParts[i].part).find((r) => r.Type.endsWith('/slideLayout')).Target;
  deck.insertSlide(0, { title: 'After the first' });
  assert.equal(layoutOf(1), layoutOf(0));
});

test('a new slide after a blank one is blank, and words given for it get a layout with room', () => {
  const deck = Deck.open(bytes);
  const after = deck.insertSlide(1);
  assert.equal(deck.slide(after).shapes.length, 0);
  const titled = deck.insertSlide(1, { title: 'Words' });
  assert.notEqual(deck.pkg.rels(deck.slideParts[titled].part).find((r) => r.Type.endsWith('/slideLayout')).Target, deck.pkg.rels(deck.slideParts[1].part).find((r) => r.Type.endsWith('/slideLayout')).Target);
});
