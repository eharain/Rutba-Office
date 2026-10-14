'use strict';

// A five-page book for the tests and the window checks: an outline (nested,
// one item through a named destination, one through a GoTo action, one a
// web link), page labels (i, ii, then A-1 onward), a link on page 1 through
// a named destination to page 4, and a text field with widgets on pages 2
// and 4. The suite's writer makes none of these, so an update adds them as
// another tool would.

const { PdfDocument } = require('../src/document');
const { Dict, Name, Ref } = require('../src/reader');
const { textString } = require('../src/assemble');
const { PdfUpdate } = require('../src/update');

const created = '2026-10-14T09:00:00Z';

function book(prefix = 'Page', { fieldName = 'reviewer' } = {}) {
  const doc = new PdfDocument({ title: `${prefix} book`, created });
  for (let i = 0; i < 5; i++) doc.addPage().text(`${prefix} ${i + 1}`, 72, 100, { font: 'Helvetica', size: 24 });
  const up = new PdfUpdate(doc.toBuffer());
  const pages = up.file.pages();
  const ref = (i) => pages[i].ref;
  // A named destination in a name tree, to page 4.
  const names = up.add(new Dict([['Dests', new Dict([['Names', [textString('chapter-two'), new Dict([['D', [ref(3), new Name('Fit')]]])]]])]]));
  // The outline: Contents > [Intro (page 1), Body (page 3) > [Detail (named, page 4), End (GoTo, page 5)]], Web (a link).
  const root = up.alloc();
  const [contents, intro, body, detail, end, web] = [up.alloc(), up.alloc(), up.alloc(), up.alloc(), up.alloc(), up.alloc()];
  const r = (n) => new Ref(n, 0);
  up.set(contents, new Dict([['Title', textString('Contents')], ['Parent', r(root)], ['Next', r(web)], ['First', r(intro)], ['Last', r(body)], ['Count', 4], ['Dest', [ref(0), new Name('XYZ'), 0, 792, 0]]]));
  up.set(intro, new Dict([['Title', textString('Intro')], ['Parent', r(contents)], ['Next', r(body)], ['Dest', [ref(0), new Name('Fit')]]]));
  up.set(body, new Dict([['Title', textString('Body')], ['Parent', r(contents)], ['Prev', r(intro)], ['First', r(detail)], ['Last', r(end)], ['Count', 2], ['Dest', [ref(2), new Name('Fit')]], ['F', 2], ['C', [0.8, 0, 0]]]));
  up.set(detail, new Dict([['Title', textString('Detail')], ['Parent', r(body)], ['Next', r(end)], ['Dest', textString('chapter-two')]]));
  up.set(end, new Dict([['Title', textString('End')], ['Parent', r(body)], ['Prev', r(detail)], ['A', new Dict([['S', new Name('GoTo')], ['D', [ref(4), new Name('Fit')]]])]]));
  up.set(web, new Dict([['Title', textString('Web')], ['Parent', r(root)], ['Prev', r(contents)], ['A', new Dict([['S', new Name('URI')], ['URI', textString('https://rutba.io/')]])]]));
  up.set(root, new Dict([['Type', new Name('Outlines')], ['First', r(contents)], ['Last', r(web)], ['Count', 6]]));
  // A link on page 1 to the named destination, and a text field with widgets on pages 2 and 4.
  const link = up.add(new Dict([['Type', new Name('Annot')], ['Subtype', new Name('Link')], ['Rect', [72, 600, 200, 620]], ['P', ref(0)], ['Dest', textString('chapter-two')]]));
  const field = up.alloc();
  const w1 = up.add(new Dict([['Type', new Name('Annot')], ['Subtype', new Name('Widget')], ['Parent', r(field)], ['P', ref(1)], ['Rect', [72, 500, 300, 520]], ['F', 4]]));
  const w2 = up.add(new Dict([['Type', new Name('Annot')], ['Subtype', new Name('Widget')], ['Parent', r(field)], ['P', ref(3)], ['Rect', [72, 500, 300, 520]], ['F', 4]]));
  up.set(field, new Dict([['FT', new Name('Tx')], ['T', textString(fieldName)], ['V', textString(`${prefix} reviewer`)], ['Kids', [w1, w2]], ['DA', textString('/Helv 0 Tf 0 g')]]));
  const helv = up.add(new Dict([['Type', new Name('Font')], ['Subtype', new Name('Type1')], ['BaseFont', new Name('Helvetica')]]));
  const form = up.add(new Dict([['Fields', [r(field)]], ['DR', new Dict([['Font', new Dict([['Helv', helv]])]])], ['DA', textString('/Helv 0 Tf 0 g')]]));
  const annotsOn = { 0: [link], 1: [w1], 3: [w2] };
  for (const [i, list] of Object.entries(annotsOn)) {
    const d = new Dict(pages[i].dict.entries());
    d.set('Annots', list);
    up.set(ref(Number(i)), d);
  }
  const catalog = new Dict(up.file.catalog().entries());
  catalog.set('Outlines', r(root));
  catalog.set('Names', names);
  catalog.set('AcroForm', form);
  catalog.set('PageLabels', new Dict([['Nums', [0, new Dict([['S', new Name('r')]]), 2, new Dict([['S', new Name('D')], ['P', textString('A-')]])]]]));
  up.set(up.file.trailer.get('Root'), catalog);
  return up.toBuffer();
}

module.exports = { book };
