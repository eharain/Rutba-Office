// Clear all formatting — the deck engine's half: every run of the shape
// keeps its text and its link, a field or a break, and nothing else; the
// paragraphs' own properties stay.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';
import { OoxmlPackage } from '@rutba/ooxml/package';

const DECK = buildPptx({ title: 'Clear', slides: [{ layout: 'obj', title: 'A shape', body: ['One'] }] });
const runs = (s) => (s?.text?.paragraphs || []).flatMap((p) => p.runs || []);

test('clearTextFormat leaves the words, their link and the paragraph properties, and nothing else', () => {
  const deck = Deck.open(DECK);
  const id = deck.addTextBox(0, { x: 10, y: 10, w: 300, h: 80, paragraphs: [
    { align: 'center', runs: [{ text: 'Big ', bold: true, size: 40, color: '#ff0000', font: 'Georgia' }, { text: 'red', italic: true, underline: true }] },
    { level: 1, bullet: { type: 'char', char: '•' }, runs: [{ text: 'and small', size: 9, strike: true }] },
  ] });
  const box = () => deck.slide(0).shapes.find((s) => String(s.id) === String(id));
  deck.setLink(0, id, 'https://office.rutba.io');
  assert.ok(runs(box()).some((r) => r.bold) && runs(box()).every((r) => r.link), 'formatted and linked to begin with');

  assert.equal(deck.clearTextFormat(0, id), true);
  const after = runs(box());
  assert.deepEqual(after.map((r) => r.text), ['Big ', 'red', 'and small'], 'the words');
  assert.ok(after.every((r) => !r.bold && !r.italic && !r.underline && !r.strike && !r.size && !r.color && !r.font), 'nothing else');
  assert.ok(after.every((r) => r.link?.url === 'https://office.rutba.io'), 'the link stays');
  assert.equal(box().text.paragraphs[0].align, 'center', 'the paragraph keeps its alignment');
  assert.equal(box().text.paragraphs[1].level, 1, 'and its level');

  const xml = OoxmlPackage.read(deck.save()).text(deck.slideParts[0].part);
  const sp = xml.slice(xml.indexOf(`<p:cNvPr id="${id}"`));
  assert.match(sp, /<a:rPr lang="en-US"><a:hlinkClick r:id="rId\d+"[^>]*\/><\/a:rPr><a:t xml:space="preserve">Big <\/a:t>/, 'a run with only its link');
  assert.doesNotMatch(sp, / b="1"| sz="| i="1"|<a:solidFill>|<a:latin/, 'no formatting left in the file');

  const title = deck.slide(0).shapes.find((s) => s.placeholder?.type === 'title');
  assert.equal(deck.clearTextFormat(0, title.id), true, 'a placeholder clears too');
  assert.throws(() => deck.clearTextFormat(0, 9999), /not found/);
});
