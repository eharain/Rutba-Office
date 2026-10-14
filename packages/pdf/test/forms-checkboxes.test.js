'use strict';

// Check box fields made where a scanned form's tick boxes are printed
// (addCheckboxes): an incremental update that leaves every original byte,
// each box a field and its widget on its page, ticked ones ticked, a form
// made for a file that had none and joined for one that had one, names kept
// unique; the fields read back, fill like any other, and pdf.js sees them.

const test = require('node:test');
const assert = require('node:assert');
const { openPdf } = require('../src/reader');
const { PdfDocument } = require('../src/document');
const { readFields, fillFields, addCheckboxes } = require('../src/forms');
const { formPdf } = require('./form-fixture');

const scanned = () => {
  const doc = new PdfDocument({ title: 'Scanned form', created: '2026-10-14T09:00:00Z' });
  doc.addPage().text('Type of request:  [ ] New   [ ] Renewal   [ ] Change', 72, 700, { font: 'Helvetica', size: 12 });
  doc.addPage().text('Contact by:  [ ] Email   [ ] Phone', 72, 700, { font: 'Helvetica', size: 12 });
  return doc.toBuffer();
};

test('addCheckboxes: a form made for a file with none, each box a field on its page, ticked ones ticked, the original bytes kept', () => {
  const original = scanned();
  const { bytes, added } = addCheckboxes(original, [
    { page: 0, rect: [180, 695, 192, 707], name: 'New', checked: false },
    { page: 0, rect: [250, 695, 262, 707], name: 'Renewal', checked: true },
    { page: 1, rect: [160, 695, 172, 707], name: 'Email', checked: true },
  ]);
  assert.equal(added.length, 3);
  assert.ok(Buffer.from(bytes).subarray(0, original.length).equals(original), 'an incremental update');
  const fields = readFields(openPdf(bytes));
  assert.deepEqual(fields.map((f) => [f.name, f.type, f.value]), [['New', 'checkbox', 'Off'], ['Renewal', 'checkbox', 'Yes'], ['Email', 'checkbox', 'Yes']]);
  assert.deepEqual(fields.map((f) => f.widgets[0].page), [0, 0, 1]);
  assert.deepEqual(fields[1].widgets[0].rect, [250, 695, 262, 707]);
  assert.equal(fields[1].widgets[0].onState, 'Yes');

  // They fill like any field: New ticked, Renewal cleared.
  const filled = fillFields(bytes, { New: true, Renewal: false });
  const after = readFields(openPdf(filled.bytes));
  assert.deepEqual(after.map((f) => f.value), ['Yes', 'Off', 'Yes']);
});

test('addCheckboxes: joined to a form that has fields, names made unique, nothing asked nothing changed', () => {
  const base = formPdf();
  const before = readFields(openPdf(base));
  const { bytes } = addCheckboxes(base, [
    { page: 0, rect: [400, 400, 412, 412], name: before[0].name, checked: true },
    { page: 0, rect: [420, 400, 432, 412] },
  ]);
  const fields = readFields(openPdf(bytes));
  assert.equal(fields.length, before.length + 2);
  assert.deepEqual(fields.slice(0, before.length).map((f) => f.name), before.map((f) => f.name), 'the form\'s own fields stay');
  assert.deepEqual(fields.slice(before.length).map((f) => f.name), [`${before[0].name} 2`, 'Check Box 2']);
  assert.ok(Buffer.from(addCheckboxes(base, []).bytes).equals(base));
  assert.throws(() => addCheckboxes(base, [{ page: 9, rect: [0, 0, 10, 10] }]), /no page 10/);
});

test('pdf.js reads the check boxes made', async (t) => {
  const { readWithPdfjs } = require('./pdfjs');
  const { bytes } = addCheckboxes(scanned(), [{ page: 0, rect: [180, 695, 192, 707], name: 'New', checked: true }, { page: 0, rect: [250, 695, 262, 707], name: 'Renewal' }]);
  const read = await readWithPdfjs(bytes);
  if (!read) { t.skip('pdf.js is not vendored'); return; }
  assert.equal(read.fields.New?.[0]?.value, 'Yes');
  assert.equal(read.fields.Renewal?.[0]?.value, 'Off');
});
