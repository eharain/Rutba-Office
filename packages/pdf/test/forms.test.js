'use strict';

// Forms (forms.js): the fields of an AcroForm read with their names, kinds,
// values and widgets; filled as an incremental update whose appearance
// streams carry the words (sized to fit when the form says 0, wrapped in a
// multi-line field, centred when the field says so), check boxes and radio
// buttons turned to their states, choices written with their indices; and
// words Windows-1252 cannot spell left for the reader to draw. pdf.js reads
// the values back.

const test = require('node:test');
const assert = require('node:assert');
const { openPdf, Stream } = require('../src/reader');
const { readFields, fillFields, textAppearance } = require('../src/forms');
const { formPdf } = require('./form-fixture');

const appearanceText = (file, widgetId) => {
  const w = file.getObject(Number(widgetId));
  const n = file.resolve(file.resolve(w.get('AP')).get('N'));
  return n instanceof Stream ? file.decode(n).toString('latin1') : '';
};

test('the fields are read with their names, kinds, values, options and widgets', () => {
  const fields = readFields(openPdf(formPdf()));
  assert.deepEqual(fields.map((f) => `${f.name}:${f.type}`), ['name:text', 'address:text', 'amount:text', 'agree:checkbox', 'size:radio', 'country:combo', 'colours:list', 'locked:text']);
  const byName = Object.fromEntries(fields.map((f) => [f.name, f]));
  assert.equal(byName.address.multiline, true);
  assert.equal(byName.amount.align, 2);
  assert.equal(byName.agree.value, 'Off');
  assert.equal(byName.agree.widgets[0].onState, 'Yes');
  assert.deepEqual(byName.size.widgets.map((w) => w.onState), ['S', 'M', 'L']);
  assert.deepEqual(byName.country.options, [{ value: 'Pakistan', label: 'Pakistan' }, { value: 'United Kingdom', label: 'United Kingdom' }, { value: 'IN', label: 'India' }]);
  assert.equal(byName.locked.readOnly, true);
  assert.equal(byName.locked.value, 'fixed');
  assert.ok(fields.every((f) => f.widgets.every((w) => w.page === 0 && w.rect)));
});

test('filling writes the values and their appearances as an incremental update', () => {
  const original = formPdf();
  const ids = Object.fromEntries(readFields(openPdf(original)).map((f) => [f.name, f.id]));
  const { bytes, changed } = fillFields(original, {
    name: 'Ejaz Arain', address: 'Flat 2, 10 Long Street, Gulberg III, Lahore 54000, Punjab, Pakistan, which is long enough to wrap', amount: '1,250.00',
    agree: true, size: 'M', country: 'IN', colours: ['Green', 'Blue'], locked: 'changed?',
  });
  assert.ok(bytes.subarray(0, original.length).equals(original));
  assert.equal(changed.length, 7, 'the read-only field is left alone');
  const file = openPdf(bytes);
  assert.equal(file.repaired, false);
  const f = Object.fromEntries(readFields(file).map((x) => [x.name, x]));
  assert.equal(f.name.value, 'Ejaz Arain');
  assert.equal(f.agree.value, 'Yes');
  assert.equal(f.agree.widgets[0].state, 'Yes');
  assert.equal(f.size.value, 'M');
  assert.deepEqual(f.size.widgets.map((w) => w.state), ['Off', 'M', 'Off']);
  assert.equal(f.country.value, 'IN');
  assert.deepEqual(f.colours.value, ['Green', 'Blue']);
  assert.equal(f.locked.value, 'fixed');
  // Appearances: the words in the field's font, auto-sized, wrapped, aligned right, the label of a choice.
  const name = appearanceText(file, ids.name);
  assert.match(name, /\/Tx BMC/);
  assert.match(name, /\/Helv [\d.]+ Tf/);
  assert.match(name, /\(Ejaz Arain\) Tj/);
  const address = appearanceText(file, ids.address);
  assert.ok((address.match(/Tj/g) || []).length >= 2, 'wrapped');
  assert.match(address, /0 0 0.5 rg/);
  assert.match(appearanceText(file, ids.country), /\(India\) Tj/);
  const list = appearanceText(file, ids.colours);
  assert.equal((list.match(/0.6 0.75 0.86 rg/g) || []).length, 2, 'two options chosen on a band');
  // Unchecking turns the box off again.
  const again = fillFields(bytes, { agree: false, size: 'L' });
  const g = Object.fromEntries(readFields(openPdf(again.bytes)).map((x) => [x.name, x]));
  assert.equal(g.agree.value, 'Off');
  assert.deepEqual(g.size.widgets.map((w) => w.state), ['Off', 'Off', 'L']);
});

test('words Windows-1252 cannot spell are left for the reader to draw, the form marked to say so', () => {
  const { bytes } = fillFields(formPdf(), { name: 'اعجاز' });
  const file = openPdf(bytes);
  const f = readFields(file).find((x) => x.name === 'name');
  assert.equal(f.value, 'اعجاز');
  assert.equal(file.getObject(Number(f.id)).get('AP'), undefined);
  assert.equal(file.resolve(file.catalog().get('AcroForm')).get('NeedAppearances'), true);
});

test('a field that is not there is refused by name', () => {
  assert.throws(() => fillFields(formPdf(), { nothing: 'x' }), /no field named nothing/);
});

test('an auto-sized appearance shrinks a long value to its box', () => {
  const short = textAppearance({ width: 100, height: 20, text: 'Hi', da: '/Helv 0 Tf 0 g' }).content;
  const long = textAppearance({ width: 100, height: 20, text: 'A very long value that cannot fit at full size', da: '/Helv 0 Tf 0 g' }).content;
  const size = (s) => Number(/\/Helv ([\d.]+) Tf/.exec(s)[1]);
  assert.ok(size(long) < size(short));
});

test('pdf.js reads the values filled', async (t) => {
  const { readWithPdfjs } = require('./pdfjs');
  const { bytes } = fillFields(formPdf(), { name: 'Ejaz Arain', agree: true, size: 'M', country: 'IN' });
  const read = await readWithPdfjs(bytes);
  if (!read) { t.skip('pdf.js is not vendored'); return; }
  const value = (name) => read.fields[name]?.[0]?.value;
  assert.equal(value('name'), 'Ejaz Arain');
  assert.equal(value('agree'), 'Yes');
  assert.equal(read.fields.size.find((w) => w.exportValues === 'M' || w.exportValues?.includes?.('M'))?.value ?? value('size'), 'M');
  assert.equal(value('country'), 'IN');
});

