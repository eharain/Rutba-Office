// Citations & Bibliography — the engine's half.
//
// The sources live in the document's custom XML part (<b:Sources>, as Word
// writes it, with its itemProps and relationships); a citation is a
// CITATION field in a content control marked w:citation; the bibliography a
// BIBLIOGRAPHY field in a w:bibliography control inside the Bibliographies
// building block. Each style is formatted as its manual sets it out, in the
// text and in the list.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';
import {
  parseSources, sourcesXml, formatCitation, formatBibliography, formatBibliographyEntry, segmentsText,
  makeTag, parseNames, citationInstr, parseCitationInstr, styleFromAttributes,
} from '@rutba/ooxml/bibliography';

const BOOK = { tag: 'Smi20', type: 'Book', guid: '{11111111-1111-4111-8111-111111111111}', people: { Author: { list: [{ last: 'Smith', first: 'John', middle: 'A' }] } }, fields: { Title: 'The Book of Things', Year: '2020', City: 'London', Publisher: 'Penguin' } };
const ARTICLE = { tag: 'Jon19', type: 'JournalArticle', guid: '{22222222-2222-4222-8222-222222222222}', people: { Author: { list: [{ last: 'Jones', first: 'Bob' }, { last: 'Brown', first: 'Carl' }] } }, fields: { Title: 'On testing', JournalName: 'Journal of Tests', Year: '2019', Volume: '12', Issue: '3', Pages: '45-67' } };
const WEB = { tag: 'Wor21', type: 'InternetSite', guid: '{33333333-3333-4333-8333-333333333333}', people: { Author: { corporate: 'World Health Organization' } }, fields: { Title: 'Fact sheet', InternetSiteTitle: 'WHO', Year: '2021', Month: '3', Day: '5', YearAccessed: '2022', MonthAccessed: 'June', DayAccessed: '1', URL: 'https://who.int/x' } };
const SECTION = { tag: 'Gre18', type: 'BookSection', guid: '{44444444-4444-4444-8444-444444444444}', people: { Author: { list: [{ last: 'Green', first: 'Dee' }] }, Editor: { list: [{ last: 'White', first: 'Eve' }] } }, fields: { Title: 'A chapter', BookTitle: 'Collected Papers', Year: '2018', Pages: '10-20', City: 'Oxford', Publisher: 'OUP' } };
const REPORT = { tag: 'Nat17', type: 'Report', guid: '{55555555-5555-4555-8555-555555555555}', people: { Author: { corporate: 'National Audit Office' } }, fields: { Title: 'Annual review', Year: '2017', City: 'London', Publisher: 'NAO' } };
const PAPER = { tag: 'Lee16', type: 'ConferenceProceedings', guid: '{66666666-6666-4666-8666-666666666666}', people: { Author: { list: [{ last: 'Lee', first: 'Ann' }] } }, fields: { Title: 'Fast things', ConferenceName: 'Proceedings of Speed', Year: '2016', Pages: '1-9', City: 'Paris', Publisher: 'ACM' } };

/** Segments with italics marked *so*, for comparing to a manual. */
const shown = (segs) => segs.map((s) => (s.italic ? '*' + s.text + '*' : s.text)).join('');
const cite = (list, style, opts) => segmentsText(formatCitation(list, style, opts));

test('APA 7: in-text by surname and year, & for two, et al. from three; the list by the seventh edition', () => {
  assert.equal(cite([BOOK], 'apa7'), '(Smith, 2020)');
  assert.equal(cite([ARTICLE], 'apa7', { pages: '50' }), '(Jones & Brown, 2019, p. 50)');
  assert.equal(cite([{ ...ARTICLE, people: { Author: { list: [...ARTICLE.people.Author.list, { last: 'Grey', first: 'Al' }] } } }], 'apa7'), '(Jones et al., 2019)');
  assert.equal(cite([BOOK, ARTICLE], 'apa7'), '(Smith, 2020; Jones & Brown, 2019)');
  assert.equal(shown(formatBibliographyEntry(BOOK, 'apa7')), 'Smith, J. A. (2020). *The Book of Things*. Penguin.');
  assert.equal(shown(formatBibliographyEntry(ARTICLE, 'apa7')), 'Jones, B., & Brown, C. (2019). On testing. *Journal of Tests*, *12*(3), 45–67.');
  assert.equal(shown(formatBibliographyEntry(SECTION, 'apa7')), 'Green, D. (2018). A chapter. In E. White (Ed.), *Collected Papers* (pp. 10–20). OUP.');
  assert.equal(shown(formatBibliographyEntry(WEB, 'apa7')), 'World Health Organization. (2021, March 5). *Fact sheet*. WHO. Retrieved June 1, 2022, from https://who.int/x');
  assert.equal(shown(formatBibliographyEntry(REPORT, 'apa7')), 'National Audit Office. (2017). *Annual review*. NAO.');
  assert.equal(shown(formatBibliographyEntry(PAPER, 'apa7')), 'Lee, A. (2016). Fast things. In *Proceedings of Speed* (pp. 1–9). ACM.');
});

test('APA 6: the place of publication stays, and three to five authors are all named in the text', () => {
  const three = { ...ARTICLE, people: { Author: { list: [...ARTICLE.people.Author.list, { last: 'Grey', first: 'Al' }] } } };
  assert.equal(cite([three], 'apa6'), '(Jones, Brown, & Grey, 2019)');
  assert.equal(shown(formatBibliographyEntry(BOOK, 'apa6')), 'Smith, J. A. (2020). *The Book of Things*. London: Penguin.');
});

test('MLA 7: no year in the text, the page after the name, the medium at the end of the entry', () => {
  assert.equal(cite([BOOK], 'mla7'), '(Smith)');
  assert.equal(cite([BOOK], 'mla7', { pages: '23' }), '(Smith 23)');
  assert.equal(cite([ARTICLE], 'mla7'), '(Jones and Brown)');
  assert.equal(shown(formatBibliographyEntry(BOOK, 'mla7')), 'Smith, John A. *The Book of Things*. London: Penguin, 2020. Print.');
  assert.equal(shown(formatBibliographyEntry(ARTICLE, 'mla7')), 'Jones, Bob, and Carl Brown. “On testing.” *Journal of Tests* 12.3 (2019): 45–67. Print.');
  assert.equal(shown(formatBibliographyEntry(WEB, 'mla7')), 'World Health Organization. “Fact sheet.” *WHO*. 5 Mar. 2021. Web. 1 June 2022.');
});

test('Chicago 16: author and year with no comma in the text; the bibliography by the notes style', () => {
  assert.equal(cite([BOOK], 'chicago16'), '(Smith 2020)');
  assert.equal(cite([BOOK], 'chicago16', { pages: '23' }), '(Smith 2020, 23)');
  assert.equal(shown(formatBibliographyEntry(BOOK, 'chicago16')), 'Smith, John A. *The Book of Things*. London: Penguin, 2020.');
  assert.equal(shown(formatBibliographyEntry(ARTICLE, 'chicago16')), 'Jones, Bob, and Carl Brown. “On testing.” *Journal of Tests* 12, no. 3 (2019): 45–67.');
  assert.equal(shown(formatBibliographyEntry(SECTION, 'chicago16')), 'Green, Dee. “A chapter.” In *Collected Papers*, edited by Eve White, 10–20. Oxford: OUP, 2018.');
});

test('Harvard (Anglia 2008): surnames in capitals and the year straight after them', () => {
  assert.equal(cite([BOOK], 'harvard'), '(Smith, 2020)');
  assert.equal(cite([BOOK], 'harvard', { pages: '23' }), '(Smith, 2020, p.23)');
  assert.equal(cite([ARTICLE], 'harvard'), '(Jones and Brown, 2019)');
  assert.equal(shown(formatBibliographyEntry(BOOK, 'harvard')), 'SMITH, J.A., 2020. *The Book of Things*. London: Penguin.');
  assert.equal(shown(formatBibliographyEntry(ARTICLE, 'harvard')), 'JONES, B. and BROWN, C., 2019. On testing. *Journal of Tests*, 12(3), pp.45–67.');
  assert.equal(shown(formatBibliographyEntry(WEB, 'harvard')), 'World Health Organization, 2021. *Fact sheet*. [Online] Available at: https://who.int/x [Accessed 1 June 2022].');
});

test('IEEE: numbers in brackets by the order of first citation, the list in that order', () => {
  const numbers = new Map([['Jon19', 1], ['Smi20', 2]]);
  assert.equal(cite([BOOK], 'ieee', { numbers }), '[2]');
  assert.equal(cite([BOOK, ARTICLE], 'ieee', { numbers }), '[2, 1]');
  const list = formatBibliography([BOOK, ARTICLE], 'ieee', { citedTags: ['Jon19', 'Smi20'] });
  assert.deepEqual(list.map((e) => e.tag), ['Jon19', 'Smi20']);
  assert.equal(shown(list[1].segments), '[2]\tJ. A. Smith, *The Book of Things*, London: Penguin, 2020.');
});

test('the alphabetical styles sort by the lead name, a corporate author as a name, a title without its article', () => {
  const noAuthor = { tag: 'The', type: 'Book', people: {}, fields: { Title: 'The Anonymous Book', Year: '2001' } };
  const list = formatBibliography([BOOK, WEB, ARTICLE, noAuthor], 'apa7');
  assert.deepEqual(list.map((e) => e.tag), ['The', 'Jon19', 'Smi20', 'Wor21']);
  assert.equal(cite([noAuthor], 'apa7'), '(The Anonymous Book, 2001)');
});

test('<b:Sources> round trips: style, people, fields, and an element we do not edit, kept', () => {
  const xml = sourcesXml({ style: 'mla7', sources: [BOOK, ARTICLE, WEB] });
  assert.ok(xml.includes('SelectedStyle="\\MLASeventhEditionOfficeOnline.xsl" StyleName="MLA" Version="7"'));
  assert.ok(xml.includes('<b:Author><b:Author><b:NameList><b:Person><b:Last>Smith</b:Last><b:First>John</b:First><b:Middle>A</b:Middle></b:Person></b:NameList></b:Author></b:Author>'));
  assert.ok(xml.includes('<b:Author><b:Author><b:Corporate>World Health Organization</b:Corporate></b:Author></b:Author>'));
  const back = parseSources(xml);
  assert.equal(back.style, 'mla7');
  assert.equal(sourcesXml(back), xml);
  // Word's own, with a field we have no editor for (b:LCID) and a Sources attribute order of its own.
  const word = '<?xml version="1.0"?><b:Sources SelectedStyle="\\APASixthEditionOfficeOnline.xsl" StyleName="APA" Version="6" xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"><b:Source><b:Tag>Doe10</b:Tag><b:SourceType>Book</b:SourceType><b:Guid>{AAAAAAAA-0000-4000-8000-000000000000}</b:Guid><b:LCID>2057</b:LCID><b:Author><b:Author><b:NameList><b:Person><b:Last>Doe</b:Last><b:First>Jane</b:First></b:Person></b:NameList></b:Author></b:Author><b:Title>Things &amp; Stuff</b:Title><b:Year>2010</b:Year><b:RefOrder>1</b:RefOrder></b:Source></b:Sources>';
  const read = parseSources(word);
  assert.equal(read.style, 'apa6');
  assert.equal(read.sources[0].fields.Title, 'Things & Stuff');
  assert.equal(read.sources[0].refOrder, 1);
  const again = sourcesXml(read);
  assert.ok(again.includes('<b:LCID>2057</b:LCID>'), 'an element we do not edit rides through');
  assert.ok(again.includes('<b:Title>Things &amp; Stuff</b:Title>'));
  assert.deepEqual(parseSources(again), read);
  assert.equal(styleFromAttributes({ selected: '\\HarvardAnglia2008OfficeOnline.xsl' }), 'harvard');
});

test('tags, names and the CITATION field code, as Word writes them', () => {
  assert.equal(makeTag(BOOK, []), 'Smi20');
  assert.equal(makeTag(BOOK, ['Smi20', 'smi201']), 'Smi202');
  assert.deepEqual(parseNames('Smith, John A.; Bob Jones'), [{ last: 'Smith', first: 'John', middle: 'A.' }, { last: 'Jones', first: 'Bob', middle: '' }]);
  const instr = citationInstr({ tags: ['Smi20', 'Jon19'], pages: '23', lcid: 2057 });
  assert.equal(instr, ' CITATION Smi20 \\l 2057 \\p 23 \\m Jon19 ');
  assert.deepEqual(parseCitationInstr(instr), { tags: ['Smi20', 'Jon19'], lcid: 2057, pages: '23', suppressAuthor: false, suppressYear: false, suppressTitle: false });
});

const letter = () => openDocx(buildDocx({ styles: true, paragraphs: [
  { text: 'Testing is well studied' },
  { text: 'The end.' },
] }));

test('a citation is a CITATION field in a w:citation control, the sources in customXml with their itemProps', () => {
  const view = letter();
  view.setSources({ sources: [BOOK, ARTICLE] });
  view.setSelection({ block: 0, offset: view.blocks[0].text.length });
  view.insertCitation({ tags: ['Smi20'] });
  assert.equal(view.blocks[0].text, 'Testing is well studied (Smith, 2020)');
  const engine = view.doc.doc;
  const p = engine.editParagraph(0);
  assert.ok(/<w:sdt><w:sdtPr><w:id w:val="-?\d+"\/><w:citation\/><\/w:sdtPr><w:sdtContent>/.test(p.xml));
  assert.ok(p.xml.includes('<w:instrText xml:space="preserve"> CITATION Smi20 \\l 1033 </w:instrText>'));
  assert.equal(p.structural, false, 'a paragraph with a citation stays editable');
  // The window's frame carries them, for the ribbon and Manage Sources.
  const frame = view.render({ pages: false });
  assert.deepEqual(frame.references.sources.map((s) => s.tag), ['Smi20', 'Jon19']);
  assert.deepEqual(frame.references.citations.map((c) => c.tags), [['Smi20']]);

  // Typing after it keeps the control byte for byte.
  view.insertText('.');
  assert.equal(view.blocks[0].text, 'Testing is well studied (Smith, 2020).');
  assert.ok(engine.editParagraph(0).xml.includes('<w:citation/>'));

  const saved = openDocx(view.save());
  const doc = saved.doc.doc;
  const part = doc._bibliographyPart();
  assert.equal(part, 'customXml/item1.xml');
  assert.ok(doc.pkg.has('customXml/itemProps1.xml'));
  assert.ok(doc.pkg.text('customXml/itemProps1.xml').includes('ds:schemaRef ds:uri="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"'));
  assert.ok(doc.pkg.text('customXml/_rels/item1.xml.rels').includes('relationships/customXmlProps'));
  assert.ok(doc.pkg.text('word/_rels/document.xml.rels').includes('Target="../customXml/item1.xml"'));
  assert.ok(doc.pkg.text('[Content_Types].xml').includes('/customXml/itemProps1.xml'));
  assert.deepEqual(doc.bibliographySources().sources.map((s) => s.tag), ['Smi20', 'Jon19']);
  assert.deepEqual(doc.citations().map((c) => c.tags), [['Smi20']]);
  assert.equal(saved.blocks[0].structural, false);
});

test('Style re-draws every citation and the bibliography; a new source joins the list; a placeholder shows its name', () => {
  const view = letter();
  view.setSources({ sources: [BOOK] });
  view.setSelection({ block: 0, offset: view.blocks[0].text.length });
  view.insertCitation({ tags: ['Smi20'], pages: '12' });
  view.setSelection({ block: 1, offset: 0 });
  view.insertCitation({ tags: ['Placeholder1'] });
  assert.equal(view.blocks[1].text, '(Placeholder1)The end.');
  view.setSelection({ block: 1, offset: view.blocks[1].text.length });
  view.insertBibliography({ heading: 'References' });
  const engine = view.doc.doc;
  assert.deepEqual(engine.bibliographyEntries(), { heading: 'References', entries: ['Smith, J. A. (2020). The Book of Things. Penguin.'] });
  const body = engine._body().body;
  assert.ok(body.includes('<w:docPartGallery w:val="Bibliographies"/>'));
  assert.ok(body.includes('<w:bibliography/>'));
  assert.ok(body.includes(' BIBLIOGRAPHY '));
  assert.ok(engine.pkg.text('word/styles.xml').includes('w:styleId="Bibliography"'));

  view.setBibliographyStyle('mla7');
  assert.equal(view.blocks[0].text, 'Testing is well studied (Smith 12)');
  assert.deepEqual(engine.bibliographyEntries().entries, ['Smith, John A. The Book of Things. London: Penguin, 2020. Print.']);
  assert.equal(engine.bibliographySources().style, 'mla7');

  view.setSources({ sources: [BOOK, ARTICLE] });
  assert.deepEqual(engine.bibliographyEntries().entries, [
    'Jones, Bob, and Carl Brown. “On testing.” Journal of Tests 12.3 (2019): 45–67. Print.',
    'Smith, John A. The Book of Things. London: Penguin, 2020. Print.',
  ]);

  // The placeholder becomes a real source: its citation reads from it at once.
  view.setSources({ sources: [BOOK, ARTICLE, { ...REPORT, tag: 'Placeholder1' }] });
  assert.equal(view.blocks[1].text.startsWith('(National Audit Office)'), true);
});

test('undo takes a source list back with the words', () => {
  const view = letter();
  view.setSources({ sources: [BOOK] });
  view.setSources({ sources: [BOOK, ARTICLE] });
  assert.equal(view.doc.doc.bibliographySources().sources.length, 2);
  view.undo();
  assert.equal(view.doc.doc.bibliographySources().sources.length, 1);
  view.redo();
  assert.equal(view.doc.doc.bibliographySources().sources.length, 2);
});

test('a citation Word wrote is read as one run and F9 brings it up to date', () => {
  const sdt = '<w:sdt><w:sdtPr><w:id w:val="-1745637220"/><w:citation/></w:sdtPr><w:sdtContent>'
    + '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> CITATION Smi20 \\l 2057 </w:instrText></w:r>'
    + '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:noProof/><w:lang w:val="en-GB"/></w:rPr><w:t xml:space="preserve"> (Smith, 2019)</w:t></w:r>'
    + '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:sdtContent></w:sdt>';
  const view = letter();
  view.setSources({ sources: [BOOK] });
  const engine = view.doc.doc;
  const p = engine.editParagraph(0);
  engine._spliceBody(p.end - '</w:p>'.length, p.end - '</w:p>'.length, sdt);
  view._invalidate();
  const b = view.blocks[0];
  assert.equal(b.structural, false);
  assert.equal(b.text, 'Testing is well studied (Smith, 2019)');
  assert.equal(b.runs.filter((r) => r.field?.kind === 'citation').length, 1);
  view.updateFields();
  assert.equal(view.blocks[0].text, 'Testing is well studied (Smith, 2020)');
  assert.ok(engine.editParagraph(0).xml.includes('<w:id w:val="-1745637220"/>'), 'the control is Word\'s own, kept');
});
