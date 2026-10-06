// Mark Citation and Insert Table of Authorities — the engine's half.
//
// Mark Citation writes a hidden TA field after the cited words — the long
// citation, the short one and the category the first time, the short one
// alone after — and the paragraph stays editable; Insert Table of
// Authorities writes a TOA field per category: its heading, then each
// authority, sorted by its long citation, with the pages the window laid
// its citations on, or "passim" for five pages or more. Update Table keeps
// the choices and brings the pages up to date.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';
import { taInstr, parseTaInstr, toaInstr, parseToaInstr, buildAuthorities, authorityTail, categoryName, CITATION_HINTS } from '@rutba/ooxml/wordtoa';

test('TA and TOA field codes are written and read as Word writes them', () => {
  assert.equal(taInstr({ long: 'Smith v. Jones, 1 U.S. 2 (1990)', short: 'Smith v. Jones', category: 1 }), ' TA \\l "Smith v. Jones, 1 U.S. 2 (1990)" \\s "Smith v. Jones" \\c 1 ');
  assert.equal(taInstr({ short: 'Smith v. Jones' }), ' TA \\s "Smith v. Jones" ');
  assert.equal(taInstr({ long: 'The "Act"\nof 1990', category: 2 }), ' TA \\l "The \'Act\' of 1990" \\s "The \'Act\' of 1990" \\c 2 ');
  assert.deepEqual(parseTaInstr(' TA \\l "Smith v. Jones" \\s "Smith" \\c 3 \\b '), { long: 'Smith v. Jones', short: 'Smith', category: 3, bookmark: '', bold: true, italic: false });
  assert.deepEqual(parseTaInstr(' TA \\s "Smith" '), { long: '', short: 'Smith', category: 1, bookmark: '', bold: false, italic: false });
  assert.equal(parseTaInstr(' XE "Smith" '), null);
  assert.equal(toaInstr({ category: 1 }), ' TOA \\h \\c "1" \\p ');
  assert.equal(toaInstr({ category: 2, passim: false, keepFormatting: false }), ' TOA \\h \\c "2" \\f ');
  assert.deepEqual(parseToaInstr(' TOA \\h \\c "4" \\e "\t" \\p '), { category: 4, heading: true, passim: true, keepFormatting: true });
  assert.equal(categoryName(2), 'Statutes');
  assert.equal(categoryName(9), '9');
  assert.ok(CITATION_HINTS.test('as held in Smith v. Jones, the'));
  assert.ok(!CITATION_HINTS.test('an ordinary sentence'));
});

test('authorities sort by their long citations, short citations count for theirs, and five pages or more read passim', () => {
  const list = buildAuthorities([
    { long: 'Zeta v. Alpha', short: 'Zeta', category: 1, page: 4 },
    { long: 'Brown v. Board', short: 'Brown', category: 1, page: 2 },
    { short: 'Brown', page: 3 },
    { short: 'brown', page: 3, bold: true },
    { long: 'Clean Air Act', category: 2, page: 1 },
    ...[5, 6, 7, 8].map((page) => ({ short: 'Zeta', page })),
  ]);
  assert.deepEqual(list.map((a) => [a.category, a.long]), [[1, 'Brown v. Board'], [2, 'Clean Air Act'], [1, 'Zeta v. Alpha']]);
  assert.deepEqual(list[0].pages, [{ text: '2', bold: false, italic: false }, { text: '3', bold: true, italic: false }]);
  assert.equal(list[2].passim, true);
  const text = (a) => a.long + authorityTail(a).map((s) => s.text).join('');
  assert.equal(text(list[0]), 'Brown v. Board\t2, 3');
  assert.equal(text(list[2]), 'Zeta v. Alpha\tpassim');
  assert.deepEqual(buildAuthorities(list.map((a) => ({ ...a, page: 1 })), { category: 2 }).map((a) => a.long), ['Clean Air Act']);
  assert.equal(buildAuthorities([{ long: 'Zeta v. Alpha', page: 1 }, ...[2, 3, 4, 5].map((page) => ({ short: 'Zeta v. Alpha', page }))], { usePassim: false })[0].passim, false);
});

const brief = () => openDocx(buildDocx({ styles: true, paragraphs: [
  { text: 'Argument' },
  { text: 'The court in Brown v. Board, 347 U.S. 483 (1954) held otherwise.' },
  { text: 'Later, Brown was read narrowly; Brown v. Board, 347 U.S. 483 (1954) again.' },
  { text: 'The Clean Air Act applies.' },
  { text: 'Nothing here.' },
  { text: '' },
] }));

test('Mark Citation writes a hidden TA field after the words, and Mark All marks the rest with the short form', () => {
  const view = brief();
  const long = 'Brown v. Board, 347 U.S. 483 (1954)';
  const p1 = view.blocks[1].text;
  view.setSelection({ block: 1, offset: p1.indexOf('Brown') }, { block: 1, offset: p1.indexOf('Brown') + long.length });
  const n = view.markCitation({ long, short: 'Brown', category: 1, all: true, text: long });
  const engine = view.doc.doc;
  assert.equal(n, 3, 'the selection, then Brown and the long form in the next paragraph');
  assert.ok(engine.editParagraph(1).xml.includes('<w:instrText xml:space="preserve"> TA \\l "' + long + '" \\s "Brown" \\c 1 </w:instrText>'));
  assert.ok(engine.editParagraph(1).xml.includes('<w:vanish/>'), 'a TA field is hidden text');
  assert.equal(view.blocks[1].text, p1, 'a TA adds no words');
  assert.equal((engine.editParagraph(2).xml.match(/ TA \\s "Brown" /g) || []).length, 2);
  assert.equal(engine.editParagraph(1).structural, false, 'the paragraph stays editable');
  assert.ok(view.blocks[1].runs.some((r) => r.field?.kind === 'ta' && r.text === ''));
  // An authority marked already is marked again with its short form.
  view.setSelection({ block: 3, offset: 4 }, { block: 3, offset: 17 });
  view.markCitation({ long: 'Clean Air Act', category: 2 });
  view.setSelection({ block: 4, offset: 0 }, { block: 4, offset: 7 });
  view.markCitation({ long: 'clean air act', category: 2 });
  assert.ok(engine.editParagraph(4).xml.includes(' TA \\s "Clean Air Act" '));
  assert.deepEqual(engine.authorities(), [{ long, short: 'Brown', category: 1 }, { long: 'Clean Air Act', short: 'Clean Air Act', category: 2 }]);
  // Typing elsewhere in the paragraph keeps the field.
  view.setSelection({ block: 1, offset: 0 });
  view.insertText('First, ');
  assert.ok(engine.editParagraph(1).xml.includes(' TA \\l "' + long));
});

test('Insert Table of Authorities writes a table per category, pages from the window, and Update Table keeps its choices', () => {
  const view = brief();
  const long = 'Brown v. Board, 347 U.S. 483 (1954)';
  const p1 = view.blocks[1].text;
  view.setSelection({ block: 1, offset: p1.indexOf('Brown') }, { block: 1, offset: p1.indexOf('Brown') + long.length });
  view.markCitation({ long, short: 'Brown', category: 1, all: true, text: long });
  view.setSelection({ block: 3, offset: 4 }, { block: 3, offset: 17 });
  view.markCitation({ long: 'Clean Air Act', category: 2 });
  view.setSelection({ block: 0, offset: 0 });
  view.insertTableOfAuthorities({ category: 'all', pages: { 1: 2, 2: 3, 3: 5 } });
  const engine = view.doc.doc;
  let tables = engine.tablesOfAuthorities();
  assert.deepEqual(tables.map((t) => t.category), [1, 2]);
  assert.deepEqual(tables[0].lines, [{ style: 'TOAHeading', text: 'Cases' }, { style: 'TableofAuthorities', text: long + '\t2, 3' }]);
  assert.deepEqual(tables[1].lines, [{ style: 'TOAHeading', text: 'Statutes' }, { style: 'TableofAuthorities', text: 'Clean Air Act\t5' }]);
  assert.equal(tables[0].leader, 'dot');
  assert.ok(engine.pkg.text('word/styles.xml').includes('w:styleId="TableofAuthorities"'));
  assert.ok(engine.xml.includes(' TOA \\h \\c "1" \\p '));
  assert.deepEqual(engine.referencesInfo().toa.map((t) => t.category), [1, 2]);
  // The table's paragraphs are the field's: the words round it stay where they were.
  assert.equal(view.blocks[0].text, 'Argument');
  // Update with no pages keeps the pages the lines had; with pages, the new ones.
  assert.equal(view.updateTablesOfAuthorities({}), 2);
  assert.equal(engine.tablesOfAuthorities()[0].lines[1].text, long + '\t2, 3');
  const at = view.blocks.findIndex((b) => b.text === 'The Clean Air Act applies.');
  view.updateTablesOfAuthorities({ pages: { [at - 1]: 7, [at]: 9 } });
  tables = engine.tablesOfAuthorities();
  assert.equal(tables[1].lines[1].text, 'Clean Air Act\t9');
  // One category, again: only that table is replaced, where it stands.
  view.insertTableOfAuthorities({ category: 2, passim: false, leader: 'hyphen', pages: { [at]: 4 } });
  tables = engine.tablesOfAuthorities();
  assert.deepEqual(tables.map((t) => [t.category, t.passim, t.leader]), [[1, true, 'dot'], [2, false, 'hyphen']]);
  assert.equal(tables[1].lines[1].text, 'Clean Air Act\t4');
  // The saved file opens again with its fields.
  const again = openDocx(view.save());
  assert.equal(again.doc.doc.tablesOfAuthorities().length, 2);
  assert.equal(again.doc.doc.authorities().length, 2);
});

test('a table with nothing marked says so, as Word does', () => {
  const view = brief();
  view.setSelection({ block: 0, offset: 0 });
  view.insertTableOfAuthorities({ category: 'all' });
  assert.deepEqual(view.doc.doc.tablesOfAuthorities()[0].lines.map((l) => l.text), ['Cases', 'No table of authorities entries found.']);
});
