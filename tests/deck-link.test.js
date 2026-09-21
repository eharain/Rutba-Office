// A link on a shape's words — the deck engine's half: hlinkClick on every
// run pointing at an External relationship the slide carries, read back
// resolved to its address, kept when the words are edited, taken off again.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';
import { OoxmlPackage } from '@rutba/ooxml/package';

const DECK = buildPptx({ title: 'Link', slides: [{ layout: 'obj', title: 'A shape', body: ['One'] }] });
const runs = (s) => (s?.text?.paragraphs || []).flatMap((p) => p.runs || []);

test('setLink puts an address on every run, resolved on the way back, written as PowerPoint writes it, kept through an edit and taken off', () => {
  const deck = Deck.open(DECK);
  const id = deck.addTextBox(0, { x: 10, y: 10, w: 300, h: 60, paragraphs: [{ runs: [{ text: 'See the ', bold: true }, { text: 'help', italic: true }] }] });
  const box = () => deck.slide(0).shapes.find((s) => String(s.id) === String(id));
  assert.ok(runs(box()).every((r) => !r.link), 'no link to begin with');

  const rId = deck.setLink(0, id, 'https://office.rutba.io/help');
  assert.match(rId, /^rId\d+$/);
  assert.deepEqual(runs(box()).map((r) => [r.text, r.link?.id, r.link?.url, r.bold ?? null, r.italic ?? null]),
    [['See the ', rId, 'https://office.rutba.io/help', true, null], ['help', rId, 'https://office.rutba.io/help', null, true]], 'every run, its look kept');

  const pkg = OoxmlPackage.read(deck.save());
  const part = deck.slideParts[0].part;
  const sp = pkg.text(part).slice(pkg.text(part).indexOf(`<p:cNvPr id="${id}"`));
  assert.match(sp, new RegExp('<a:rPr lang="en-US" b="1"><a:hlinkClick r:id="' + rId + '" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></a:rPr><a:t xml:space="preserve">See the </a:t>'));
  const rels = pkg.text(part.replace(/\/([^/]+)$/, '/_rels/$1.rels'));
  assert.match(rels, new RegExp('<Relationship Id="' + rId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://office.rutba.io/help" TargetMode="External"/>'));
  assert.equal(runs(Deck.open(deck.save()).slide(0).shapes.find((s) => String(s.id) === String(id)))[0].link.url, 'https://office.rutba.io/help', 'through a save');

  // The words edited the way the window edits them — each run's props carried — keep the link.
  const edited = box().text.paragraphs.map((p) => { const { runs: rs, plain, ...props } = p; return { ...props, runs: rs.map((r) => ({ ...r, text: r.text.toUpperCase() })) }; });
  deck.setText(0, id, edited);
  assert.deepEqual(runs(box()).map((r) => [r.text, r.link?.url]), [['SEE THE ', 'https://office.rutba.io/help'], ['HELP', 'https://office.rutba.io/help']]);

  assert.equal(deck.setLink(0, id, null), null);
  assert.ok(runs(box()).every((r) => !r.link), 'taken off every run');
  assert.doesNotMatch(OoxmlPackage.read(deck.save()).text(part), /hlinkClick/);
  assert.throws(() => deck.setLink(0, 9999, 'https://x.example'), /not found/);
});
