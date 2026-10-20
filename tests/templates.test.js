// The template gallery's files: every template the gallery lists is kept,
// pictured, and exactly what its specification makes through the suite's own
// engines today (so none is a file nobody can make again); each opens in the
// engine of the app that makes it with what it promises: a document's styles,
// headings and table of contents, a workbook's formulas worked out, a deck's
// slides on its layouts; and the gallery's words are in the catalogue.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { TEMPLATES, CATEGORIES, templatePath, thumbPath, templateById, templateForLegacy } from '@rutba/templates';
import { makeTemplate } from '@rutba/templates/make';
import { OoxmlPackage } from '@rutba/ooxml/package';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { SheetView } from '@rutba/sheet-view';
import { Deck } from '@rutba/presentation';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const MESSAGES = new Set(JSON.parse(fs.readFileSync(path.join(ROOT, 'packages/office-ui/src/catalogues/messages.json'), 'utf8')).messages.map((m) => m.message));

/** A part's text with what changes from one making to the next taken out: the time it was made, and a workbook's worked-out values (TODAY moves). */
function steady(name, text) {
  if (name === 'docProps/core.xml') return text.replace(/<dcterms:(created|modified)[^>]*>[^<]*</g, '<dcterms:$1>');
  if (/^xl\/worksheets\//.test(name)) return text.replace(/<v>[^<]*<\/v>/g, '');
  // A slide number's field carries an id of its own, new each time, and a
  // notes master's date field the day it was made (PowerPoint shows today's
  // on opening), so a run past midnight is not a changed template.
  if (/^ppt\//.test(name)) {
    return text.replace(/<a:fld id="\{[0-9A-F-]+\}"/g, '<a:fld id=""')
      .replace(/(type="datetimeFigureOut"><a:rPr[^>]*\/><a:t>)[^<]*</g, '$1<');
  }
  return text;
}

function partsOf(bytes) {
  const pkg = OoxmlPackage.read(bytes);
  const out = new Map();
  for (const name of pkg.partNames()) {
    const data = pkg.read(name);
    out.set(name, /\.(xml|rels)$/.test(name) ? steady(name, data.toString('utf8')) : data.toString('base64'));
  }
  return out;
}

const cell = (view, ref) => {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  const col = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  return view.displayValue(Number(m[2]) - 1, col).text;
};

test('the gallery lists twenty-seven templates: ten documents, ten workbooks and seven decks, each with its file and its picture', () => {
  const by = (app) => TEMPLATES.filter((t) => t.app === app).length;
  assert.deepEqual([by('word'), by('sheets'), by('slides')], [10, 10, 7]);
  assert.equal(new Set(TEMPLATES.map((t) => t.id)).size, TEMPLATES.length);
  for (const t of TEMPLATES) {
    assert.ok(fs.statSync(templatePath(t)).size > 1000, `${t.file} is kept`);
    const png = fs.readFileSync(thumbPath(t));
    assert.equal(png.subarray(1, 4).toString('latin1'), 'PNG', `${t.thumb} is a picture`);
    assert.ok(png.readUInt32BE(16) >= 300, `${t.thumb} is drawn at the gallery's size`);
    assert.ok(t.categories.length && t.categories.every((c) => CATEGORIES.some((k) => k.id === c)), `${t.id} is in known categories`);
  }
});

test('every template kept is exactly what its specification makes through the engines today', () => {
  for (const t of TEMPLATES) {
    const kept = partsOf(fs.readFileSync(templatePath(t)));
    const made = partsOf(makeTemplate(t.id));
    assert.deepEqual([...kept.keys()].sort(), [...made.keys()].sort(), `${t.id}: the same parts`);
    for (const [name, text] of made) assert.equal(kept.get(name), text, `${t.id}: ${name} as the specification makes it (run node tools/make-templates.mjs)`);
  }
});

test('a launcher quick start stands behind its gallery template, in its own app only', () => {
  assert.equal(templateForLegacy('word', 'letter').id, 'word-letter');
  assert.equal(templateForLegacy('sheets', 'budget').id, 'sheet-budget');
  assert.equal(templateForLegacy('sheets', 'invoice').id, 'sheet-invoice');
  assert.equal(templateForLegacy('slides', 'pitch').id, 'deck-pitch');
  assert.equal(templateForLegacy('word', 'invoice'), null);
  assert.equal(templateById('nothing'), null);
});

test('the documents carry a theme, styles their headings follow, and what each promises', () => {
  for (const t of TEMPLATES.filter((x) => x.app === 'word')) {
    const view = openDocx(fs.readFileSync(templatePath(t)));
    const doc = view.doc.doc;
    assert.ok(doc.themePart(), `${t.id} has a theme`);
    const styles = doc.pkg.text('word/styles.xml');
    for (const id of ['Normal', 'Title', 'Heading1', 'Heading2', 'Heading3', 'Quote', 'Caption']) assert.match(styles, new RegExp(`w:styleId="${id}"`), `${t.id} has ${id}`);
    assert.match(styles, /w:styleId="Heading1"[\s\S]*?<w:outlineLvl w:val="0"\/>[\s\S]*?w:themeColor="accent1"/, `${t.id}: Heading 1 is an outline level and wears the theme's accent`);
    const page = doc.section();
    assert.equal(Math.round(page.widthPx), 794, `${t.id} is on A4`);
  }
  const headings = (id) => {
    const doc = openDocx(fs.readFileSync(templatePath(templateById(id)))).doc.doc;
    return doc.paragraphs().map((e) => ({ style: doc.paragraph(e.index).style, text: e.text })).filter((p) => /^(Title|Heading\d)$/.test(p.style || ''));
  };
  // The report: a cover, a contents page listing the eight headings with their pages, numbered next steps.
  const report = openDocx(fs.readFileSync(templatePath(templateById('word-report')))).doc.doc;
  const toc = report.tableOfContents();
  assert.deepEqual(toc.entries.map((e) => `${e.text}@${e.page}`), ['Introduction@3', 'Background@3', 'Scope@3', 'Findings@4', 'Results@4', 'What the results mean@4', 'Recommendations@5', 'Next steps@5']);
  assert.ok(report.pkg.has('word/numbering.xml'), 'the next steps are a numbered list');
  assert.match(report.pkg.text(report.headerFooters().footers.default.part), /PAGE/, 'the footer numbers the pages');
  assert.deepEqual(headings('word-cv').filter((h) => h.style === 'Heading1').map((h) => h.text), ['Profile', 'Experience', 'Education', 'Skills']);
  assert.deepEqual(headings('word-minutes').filter((h) => h.style === 'Heading1').map((h) => h.text), ['Agenda', 'Actions']);
  assert.equal(headings('word-letter')[0].style, 'Title');
});

test('the workbooks work their figures out: totals, payments, hours and balances', () => {
  const open = (id) => SheetView.open(fs.readFileSync(templatePath(templateById(id))), { viewportWidth: 1100, viewportHeight: 620 });
  const budget = open('sheet-budget');
  assert.deepEqual(['B7', 'C7', 'B20', 'C20', 'C22'].map((r) => cell(budget, r)), ['£3,500', '£3,450', '£2,680', '£2,676', '£774']);
  assert.deepEqual((budget.drawings.get('Budget') || []).map((d) => d.kind), ['chart', 'chart'], 'two charts of the spending');
  // A figure typed in moves the totals.
  budget.setCell(4, 2, '3400');
  assert.equal(cell(budget, 'C7'), '£3,650');
  assert.equal(cell(budget, 'C22'), '£974');
  const invoice = open('sheet-invoice');
  assert.deepEqual(['D12', 'D21', 'D23', 'D24', 'D16'].map((r) => cell(invoice, r)), ['£390.00', '£1,760.00', '£352.00', '£2,112.00', '']);
  const loan = open('sheet-loan');
  assert.equal(cell(loan, 'B8'), '£1,111.66', 'the monthly payment on £200,000 at 4.5% over 25 years');
  assert.equal(cell(loan, 'I28'), '£0', 'paid off in the last year');
  assert.equal(cell(loan, 'E29'), '', 'no schedule past the term');
  const timesheet = open('sheet-timesheet');
  assert.deepEqual(['F8', 'F12', 'F15', 'B18', 'B19'].map((r) => cell(timesheet, r)), ['8.00', '7.00', '39.00', '1.50', '£616.13']);
  const todo = open('sheet-todo');
  assert.equal(todo.sheetTables().length, 1, 'the tasks are a table');
  assert.equal(cell(todo, 'B2'), '17%');
  const ledger = open('sheet-ledger');
  assert.deepEqual(['F6', 'F15', 'F16', 'I7'].map((r) => cell(ledger, r)), ['£4,340.00', '£6,075.95', '', '£6,075.95']);
});

test('the decks have five to eight slides on the theme\'s layouts, each with its title', () => {
  for (const t of TEMPLATES.filter((x) => x.app === 'slides')) {
    const deck = Deck.open(fs.readFileSync(templatePath(t)));
    assert.ok(deck.slideCount >= 5 && deck.slideCount <= 8, `${t.id}: ${deck.slideCount} slides`);
    assert.deepEqual(deck.layoutList().map((l) => l.name), ['Title Slide', 'Title and Content', 'Section Header', 'Title Only', 'Picture with Caption']);
    for (let i = 0; i < deck.slideCount; i++) {
      const title = deck.slide(i).shapes.find((s) => s.placeholder && /title/i.test(s.placeholder.type || ''));
      assert.ok(title, `${t.id} slide ${i + 1} has a title placeholder`);
    }
  }
  const counts = Object.fromEntries(TEMPLATES.filter((x) => x.app === 'slides').map((t) => [t.id, Deck.open(fs.readFileSync(templatePath(t))).slideCount]));
  assert.deepEqual(counts, { 'deck-pitch': 8, 'deck-status': 7, 'deck-album': 6, 'deck-training': 8, 'deck-quarterly': 7, 'deck-lesson': 7, 'deck-portfolio': 7 });
});

test('the gallery speaks every template\'s name and description and every category in the catalogue', () => {
  for (const t of TEMPLATES) {
    assert.ok(MESSAGES.has(t.name), `"${t.name}" is in the message list`);
    assert.ok(MESSAGES.has(t.description), `"${t.description}" is in the message list`);
  }
  for (const c of CATEGORIES) assert.ok(MESSAGES.has(c.name), `"${c.name}" is in the message list`);
});
