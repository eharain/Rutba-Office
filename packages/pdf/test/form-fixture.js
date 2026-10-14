'use strict';

// A one-page form for the tests and the window checks, its fields added by an
// update as a form tool would add them (the suite's writer makes no forms):
// a text field sized to fit, a multi-line one, one aligned right, a check
// box, a radio group of three, a drop-down, a list and a read-only field.

const { PdfDocument } = require('../src/document');
const { Dict, Name, Ref } = require('../src/reader');
const { textString } = require('../src/assemble');
const { PdfUpdate, flateStream } = require('../src/update');

/** A one-page form, its fields added by an update as a form tool would: text, multi-line, check box, radio group, combo and list. */
function formPdf() {
  const doc = new PdfDocument({ title: 'Application', created: '2026-10-14T09:00:00Z' });
  const page = doc.addPage();
  page.text('Application form', 72, 80, { font: 'Helvetica-Bold', size: 20 });
  // Each field's label, its baseline level with the field (the writer measures from the top of an A4 page).
  const label = (words, y) => page.text(words, 72, 842 - y, { font: 'Helvetica', size: 11 });
  label('Name', 686); label('Address', 646); label('Amount', 566); label('I agree', 523);
  label('Size', 483); label('Country', 446); label('Colours', 406); label('Reference', 326);
  const up = new PdfUpdate(doc.toBuffer());
  const pg = up.file.pages()[0];
  const helv = up.add(new Dict([['Type', new Name('Font')], ['Subtype', new Name('Type1')], ['BaseFont', new Name('Helvetica')], ['Encoding', new Name('WinAnsiEncoding')]]));
  const look = (w, h, body) => up.add(flateStream([['Type', new Name('XObject')], ['Subtype', new Name('Form')], ['BBox', [0, 0, w, h]], ['Resources', new Dict()]], body));
  const widget = (extra) => new Dict([['Type', new Name('Annot')], ['Subtype', new Name('Widget')], ['P', pg.ref], ['F', 4], ...extra]);
  const annots = [];
  const fields = [];
  const field = (entries) => { const r = up.add(widget(entries)); annots.push(r); fields.push(r); return r; };
  field([['FT', new Name('Tx')], ['T', textString('name')], ['Rect', [150, 680, 400, 700]], ['DA', textString('/Helv 0 Tf 0 g')]]);
  field([['FT', new Name('Tx')], ['T', textString('address')], ['Ff', 1 << 12], ['Rect', [150, 600, 400, 660]], ['DA', textString('/Helv 10 Tf 0 0 0.5 rg')]]);
  field([['FT', new Name('Tx')], ['T', textString('amount')], ['Q', 2], ['Rect', [150, 560, 300, 580]], ['DA', textString('/Helv 11 Tf 0 g')]]);
  const yes = look(14, 14, '0 g 2 7 m 6 3 l 12 12 l S');
  const off = look(14, 14, '');
  field([['FT', new Name('Btn')], ['T', textString('agree')], ['Rect', [150, 520, 164, 534]], ['V', new Name('Off')], ['AS', new Name('Off')], ['AP', new Dict([['N', new Dict([['Yes', yes], ['Off', off]])]])]]);
  // A radio group: the field, and three kids that are only widgets.
  const group = up.alloc();
  const kids = ['S', 'M', 'L'].map((state, i) => {
    const on = look(14, 14, '0 g 4 4 6 6 re f');
    const r = up.add(widget([['Parent', new Ref(group, 0)], ['Rect', [150 + i * 40, 480, 164 + i * 40, 494]], ['AS', new Name('Off')], ['AP', new Dict([['N', new Dict([[state, on], ['Off', off]])]])]]));
    annots.push(r);
    return r;
  });
  up.set(group, new Dict([['FT', new Name('Btn')], ['Ff', (1 << 15) | (1 << 14)], ['T', textString('size')], ['V', new Name('Off')], ['Kids', kids]]));
  fields.push(new Ref(group, 0));
  field([['FT', new Name('Ch')], ['Ff', 1 << 17], ['T', textString('country')], ['Opt', [textString('Pakistan'), textString('United Kingdom'), [textString('IN'), textString('India')]]], ['Rect', [150, 440, 300, 460]], ['DA', textString('/Helv 0 Tf 0 g')]]);
  field([['FT', new Name('Ch')], ['T', textString('colours')], ['Ff', 1 << 21], ['Opt', [textString('Red'), textString('Green'), textString('Blue')]], ['Rect', [150, 360, 300, 420]], ['DA', textString('/Helv 12 Tf 0 g')]]);
  field([['FT', new Name('Tx')], ['T', textString('locked')], ['Ff', 1], ['V', textString('fixed')], ['Rect', [150, 320, 300, 340]], ['DA', textString('/Helv 0 Tf 0 g')]]);
  const form = up.add(new Dict([['Fields', fields], ['DR', new Dict([['Font', new Dict([['Helv', helv]])]])], ['DA', textString('/Helv 0 Tf 0 g')]]));
  const catalog = new Dict(up.file.catalog().entries());
  catalog.set('AcroForm', form);
  up.set(up.file.trailer.get('Root'), catalog);
  const pageDict = new Dict(pg.dict.entries());
  pageDict.set('Annots', annots);
  up.set(pg.ref, pageDict);
  return up.toBuffer();
}

module.exports = { formPdf };
