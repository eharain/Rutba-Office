// Layout → Hyphenation: None, Automatic, Manual and the Hyphenation Options,
// written to settings.xml as Word writes them, a paragraph left whole, the
// optional hyphen as Word's own element, and the printout breaking words
// where the page does.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { renderPdf } from '@rutba/doc-view/export/pdf';
import { hyphenPoints, hyphenated, breakableWord, hyphenationRules } from '@rutba/doc-view/hyphenate';
import { buildDocx, OoxmlPackage } from '@rutba/ooxml';

const long = 'Internationalization characteristically incomprehensibilities notwithstanding, responsibility demonstrations overwhelmingly institutionalized interdisciplinary consideration. ';

/** A document with a settings part holding what Word writes round the hyphenation settings. */
function withSettings(paragraphs) {
  const pkg = OoxmlPackage.read(buildDocx({ styles: true, paragraphs }));
  const settings = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>';
  if (pkg.has('word/settings.xml')) pkg.write_('word/settings.xml', Buffer.from(settings, 'utf8'));
  else {
    pkg.addPart('word/settings.xml', Buffer.from(settings, 'utf8'), 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');
    pkg.addRelationshipTo('word/document.xml', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings', 'settings.xml');
  }
  return openDocx(pkg.write());
}
const settingsXml = (view) => view.doc.doc.pkg.text('word/settings.xml');

test('the hyphenator breaks English words at their syllables and leaves short words whole', () => {
  assert.equal(hyphenated('running'), 'run-ning');
  assert.equal(hyphenated('walking'), 'walk-ing');
  assert.equal(hyphenated('management'), 'manage-ment');
  assert.equal(hyphenated('information'), 'in-for-ma-tion');
  assert.equal(hyphenated('window'), 'win-dow');
  assert.equal(hyphenated('spring'), 'spring', 'a first cluster stays whole');
  assert.equal(hyphenated('shipped'), 'shipped');
  assert.deepEqual(hyphenPoints('hello'), [], 'five letters: never broken');
  for (const word of ['characteristically', 'institutionalized', 'demonstrations', 'interdisciplinary']) {
    const pts = hyphenPoints(word);
    assert.ok(pts.length >= 1, `${word} breaks: ${hyphenated(word)}`);
    assert.ok(pts.every((p) => p >= 2 && word.length - p >= 3), 'two letters before a break and three after, as TeX\'s English keeps');
  }
  const rules = hyphenationRules({ auto: true, zoneTwips: 360, limit: 0, caps: false });
  assert.equal(breakableWord('NATO', rules), false, 'capitals left whole when the option says so');
  assert.equal(breakableWord('Nation', rules), true);
  assert.equal(rules.zonePx, 24, 'the zone in px: a quarter inch');
});

test('None, Automatic and the options are written to settings.xml where Word writes them', () => {
  const view = withSettings([{ text: long }]);
  assert.deepEqual(view.hyphenation(), { auto: false, zoneTwips: 360, limit: 0, caps: true });
  view.setHyphenation({ auto: true, zoneTwips: 567, limit: 2, caps: false });
  const xml = settingsXml(view);
  assert.match(xml, /<w:defaultTabStop w:val="720"\/><w:autoHyphenation\/><w:consecutiveHyphenLimit w:val="2"\/><w:hyphenationZone w:val="567"\/><w:doNotHyphenateCaps\/><w:characterSpacingControl/, 'in the schema\'s order, after the default tab stop');
  assert.deepEqual(view.hyphenation(), { auto: true, zoneTwips: 567, limit: 2, caps: false });
  // Back to Word's defaults: nothing written for them.
  view.setHyphenation({ auto: false, zoneTwips: 360, limit: 0, caps: true });
  assert.doesNotMatch(settingsXml(view), /autoHyphenation|hyphenationZone|consecutiveHyphenLimit|doNotHyphenateCaps/);
  // One undo puts the settings back.
  view.undo();
  assert.equal(view.hyphenation().auto, true);
  const reopened = openDocx(view.save());
  assert.equal(reopened.hyphenation().zoneTwips, 567, 'in the file');
  assert.equal(reopened.render({ pages: false }).hyphenation.limit, 2, 'and in the frame the page draws from');
});

test('a paragraph left whole is w:suppressAutoHyphens, in its place in the paragraph\'s properties', () => {
  const view = withSettings([{ text: long }, { text: long }]);
  view.setSelection({ block: 0, offset: 0 });
  view.setParagraphFormat({ align: 'center', noHyphens: true });
  const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(view.doc.doc.editParagraph(0).xml)[0];
  assert.match(pPr, /<w:suppressAutoHyphens\/>[\s\S]*<w:jc w:val="center"\/>/, 'before the alignment, as the schema orders them');
  assert.equal(view.formatAtCaret().noHyphens, true);
  assert.equal(view.render({ pages: false }).blocks[0].noHyphens, true);
  view.setParagraphFormat({ noHyphens: false });
  assert.doesNotMatch(view.doc.doc.editParagraph(0).xml, /suppressAutoHyphens/);
});

test('an optional hyphen is Word\'s w:softHyphen in the file and one character of the words', () => {
  const view = withSettings([{ text: 'Hyphenation here.' }]);
  view.setSelection({ block: 0, offset: 6 });
  view.insertText('­');
  assert.equal(view.blocks[0].text, 'Hyphen­ation here.');
  const xml = view.doc.doc.editParagraph(0).xml;
  assert.match(xml, /<w:t xml:space="preserve">Hyphen<\/w:t><w:softHyphen\/><w:t xml:space="preserve">ation here\.<\/w:t>/);
  assert.equal((xml.match(/<w:softHyphen\/>/g) || []).length, 1, 'written once');
  // Typing beside it keeps it, once.
  view.setSelection({ block: 0, offset: view.blocks[0].text.length });
  view.insertText(' More.');
  assert.equal((view.doc.doc.editParagraph(0).xml.match(/<w:softHyphen\/>/g) || []).length, 1);
  const again = openDocx(view.save());
  assert.equal(again.blocks[0].text, 'Hyphen­ation here. More.');
});

test('the printout breaks words at line ends with a hyphen, within the zone and the limit, and never in a paragraph left whole', () => {
  const paragraphs = Array.from({ length: 3 }, () => ({ text: long.repeat(3) }));
  const view = withSettings(paragraphs);
  const hyphenLines = () => view.pages.pages.flatMap((p) => p.fragments).filter((f) => f.kind === 'paragraph').flatMap((f) => f.lines.map((l) => ({ ...l, block: f.paragraphIndex })));
  assert.equal(hyphenLines().filter((l) => l.hyphen).length, 0, 'none while hyphenation is off');
  view.setHyphenation({ auto: true });
  const on = hyphenLines();
  const broken = on.filter((l) => l.hyphen);
  assert.ok(broken.length >= 3, `lines end mid-word: ${broken.length}`);
  for (const l of broken) {
    const text = view.blocks[l.block].text;
    assert.match(text[l.end - 1], /[A-Za-z]/, 'the break is inside a word');
    assert.match(text[l.end], /[A-Za-z]/);
  }
  // A limit of one: never two hyphenated lines in a row.
  view.setHyphenation({ limit: 1 });
  const limited = hyphenLines();
  for (let i = 1; i < limited.length; i++) assert.ok(!(limited[i].hyphen && limited[i - 1].hyphen && limited[i].block === limited[i - 1].block), 'no two in a row');
  // A wide zone: a line is only broken when it would end very short.
  view.setHyphenation({ limit: 0, zoneTwips: 3000 });
  assert.ok(hyphenLines().filter((l) => l.hyphen).length < broken.length, 'fewer hyphens with a wider zone');
  // A paragraph left whole.
  view.setHyphenation({ zoneTwips: 360 });
  view.setSelection({ block: 1, offset: 0 });
  view.setParagraphFormat({ noHyphens: true });
  assert.equal(hyphenLines().filter((l) => l.hyphen && l.block === 1).length, 0);
  const pdf = renderPdf(view, { created: '2026-09-25T00:00:00Z' }).buffer.toString('latin1');
  assert.match(pdf, /[a-z]-\) Tj/, 'the hyphen is printed at the break');
});

test('an optional hyphen breaks the word on paper even with hyphenation off, and is invisible elsewhere', () => {
  const word = 'Supercalifragilistic­expialidocious';
  const view = withSettings([{ text: Array.from({ length: 12 }, () => word).join(' ') }]);
  const lines = view.pages.pages[0].fragments.filter((f) => f.kind === 'paragraph').flatMap((f) => f.lines);
  assert.ok(lines.some((l) => l.hyphen), 'a line ends at the optional hyphen');
  const pdf = renderPdf(view, { created: '2026-09-25T00:00:00Z' }).buffer.toString('latin1');
  assert.ok(!pdf.includes('­'), 'the optional hyphen itself is never printed');
  assert.match(pdf, /Supercalifragilistic-\) Tj/, 'only the hyphen at the break');
});
