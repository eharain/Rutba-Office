/**
 * Typing in a line that holds a tab (2.18.0, A3.0).
 *
 * In 2.17.1 a paragraph's tab-only run was kept verbatim by every rebuild
 * while the rebuild also wrote the tab back from the text, so each keystroke
 * in the Letter template's address line added a tab, and the file was saved
 * with them (a letter made in the Templates episode carried 72). A page break
 * did the same with a line break. And a triple-click's selection ran to the
 * start of the next paragraph, so typing over the name joined it to the
 * address line, where a space then froze the window.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { OoxmlPackage } from '@rutba/ooxml';
import { buildDocx } from '@rutba/ooxml/build';
import { TEMPLATES, templatePath } from '@rutba/templates';

const letter = () => openDocx(fs.readFileSync(templatePath(TEMPLATES.find((t) => t.file === 'word-letter.docx'))));
const savedXml = (view) => OoxmlPackage.read(view.save()).text('word/document.xml');
const tabsIn = (xml) => (xml.match(/<w:tab\/>/g) || []).length;
const typeEach = (view, words) => { for (const ch of words) view.insertText(ch); };

/** A paragraph whose first run is followed by `extra`. */
function withExtra(extra) {
  const pkg = OoxmlPackage.read(buildDocx({ paragraphs: ['Our letterhead', 'Body text'] }));
  const xml = pkg.text('word/document.xml');
  const marker = 'Our letterhead</w:t></w:r>';
  pkg.write_('word/document.xml', xml.replace(marker, marker + extra));
  return openDocx(pkg.write());
}

test('typing in the Letter template\'s address line keeps its one tab', () => {
  const view = letter();
  const address = view.block(1).text;
  assert.equal(address, 'Street address, Town, Postcode\tphone  |  you@example.com');
  const before = tabsIn(savedXml(view));
  view.collapseTo({ block: 1, offset: 0 });
  typeEach(view, '12 Mill Lane, ');
  view.deleteBackward();
  view.insertText(' ');
  assert.equal(view.block(1).text, '12 Mill Lane, Street address, Town, Postcode\tphone  |  you@example.com');
  const xml = savedXml(view);
  assert.equal(tabsIn(xml), before, 'one tab stays one tab, keystroke after keystroke');
  const reopened = openDocx(view.save());
  assert.equal(reopened.block(1).text, view.block(1).text, 'and the saved file reads the same');
});

test('a run of only a tab, a line break or a carriage return is written once', () => {
  for (const [extra, ch] of [['<w:r><w:tab/></w:r>', '\t'], ['<w:r><w:br/></w:r>', '\n'], ['<w:r><w:cr/></w:r>', '\n']]) {
    const view = withExtra(extra + '<w:r><w:t>More</w:t></w:r>');
    for (let i = 0; i < 4; i++) {
      view.collapseTo({ block: 0, offset: 3 + i });
      view.insertText('X');
    }
    assert.equal(view.block(0).text, 'OurXXXX letterhead' + ch + 'More', extra);
    assert.equal(openDocx(view.save()).block(0).text, 'OurXXXX letterhead' + ch + 'More', extra + ' saved');
  }
});

test('a page break is no line break of the words, and typing beside it adds none', () => {
  const view = withExtra('<w:r><w:br w:type="page"/></w:r>');
  assert.equal(view.block(0).text, 'Our letterhead');
  typeEach(view, 'Re: ');
  const xml = savedXml(view);
  assert.equal((xml.match(/<w:br w:type="page"\/>/g) || []).length, 1, 'the break is kept, once');
  assert.ok(!/<w:br\/>/.test(xml), 'and no line break grows beside it');
  const mixed = withExtra('<w:r><w:t>A</w:t><w:br w:type="column"/><w:t>B</w:t></w:r>');
  mixed.collapseTo({ block: 0, offset: 0 });
  mixed.insertText('X');
  assert.equal((savedXml(mixed).match(/<w:br w:type="column"\/>/g) || []).length, 1, 'a column break inside a run of words is kept too');
});

test('typing over a triple-clicked line keeps its paragraph mark, as in Word', () => {
  const view = letter();
  const count = view.blocks.length;
  // A triple-click selects the line and runs to the start of the next one.
  view.setSelection({ block: 0, offset: 0 }, { block: 1, offset: 0 });
  view.insertText('L');
  typeEach(view, 'arkspur');
  assert.equal(view.block(0).text, 'Larkspur');
  assert.equal(view.blocks.length, count, 'the name and the address are still two lines');
  assert.equal(view.block(1).text, 'Street address, Town, Postcode\tphone  |  you@example.com');
  view.undo();
  view.undo();
  assert.equal(view.block(0).text, 'Your Name', 'undone, the name comes back');
});

test('a space typed into the joined name and address line lays out at once', () => {
  // The line the freeze was found on: the name joined to the address, its
  // tab and its right stop, then a space.
  const view = letter();
  view.setSelection({ block: 0, offset: 0 }, { block: 0, offset: view.block(0).text.length });
  view.insertText('Larkspur');
  view.collapseTo({ block: 0, offset: 'Larkspur'.length });
  view.deleteForward();
  assert.match(view.block(0).text, /^LarkspurStreet address/);
  const started = Date.now();
  view.collapseTo({ block: 0, offset: 'Larkspur'.length });
  typeEach(view, ' Corner ');
  for (let i = 0; i < 40; i++) view.insertText(' ');
  const pages = view.pages;
  assert.ok(pages && pages.pages.length >= 1);
  assert.ok(Date.now() - started < 5000, 'laid out in ' + (Date.now() - started) + ' ms');
  assert.equal(tabsIn(savedXml(view)), tabsIn(savedXml(letter())), 'and still one tab');
});
