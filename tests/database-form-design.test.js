// Data Files' form Design (2.18, Part A2 of the plan): a form's fields
// moved, sized, taken off and put back and their labels renamed, kept with
// the form; and a subform's lines changed and added as the form points them.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Base, wizardForm, placeControl, takeOff, putBack, relabel, fieldsOff, fitForm } from '../packages/database/src/index.js';
import { labelLeft } from '../packages/database/src/form-design.js';

const shop = () => {
  const b = Base.open(':memory:');
  b.createTable({ name: 'Products', fields: [{ name: 'Name', type: 'text' }, { name: 'Price', type: 'currency' }, { name: 'In Stock', type: 'integer' }, { name: 'Notes', type: 'longtext' }, { name: 'Sold Out', type: 'yesno' }] });
  b.createTable({ name: 'Sales', fields: [{ name: 'Sold On', type: 'date' }, { name: 'Payment', type: 'text' }] });
  b.createTable({ name: 'Sale Lines', fields: [{ name: 'Sale', type: 'integer' }, { name: 'Product', type: 'integer' }, { name: 'Quantity', type: 'integer' }] });
  b.addRelationship({ from: { table: 'Sale Lines', field: 'Product' }, to: { table: 'Products', field: 'ID' } });
  b.addRelationship({ from: { table: 'Sale Lines', field: 'Sale' }, to: { table: 'Sales', field: 'ID' } });
  for (const [n, p, s] of [['Milk', 1.1, 3], ['Bread', 1.4, 10], ['Tea', 3, 4]]) b.insert('Products', { Name: n, Price: p, 'In Stock': s });
  for (const d of ['2026-10-21', '2026-10-22']) b.insert('Sales', { 'Sold On': d, Payment: 'Cash' });
  for (const [s, p, q] of [[1, 1, 2], [1, 2, 1], [2, 3, 5]]) b.insert('Sale Lines', { Sale: s, Product: p, Quantity: q });
  return b;
};

test('a form field moved and sized snaps to the grid, its label going with it, and the form grows to hold it', () => {
  const b = shop();
  const form = b.makeForm('Products');
  const price = form.controls.find((c) => c.field === 'Price');
  assert.equal(labelLeft(price), 16);
  const moved = placeControl(form, 'Price', { x: price.x + 203, y: price.y + 37, w: 301 });
  const now = moved.controls.find((c) => c.field === 'Price');
  assert.deepEqual([now.x, now.y, now.w, now.h], [price.x + 200, price.y + 40, 304, price.h]);
  assert.equal(labelLeft(now), 216);
  assert.equal(moved.width, now.x + now.w + 16);
  // Room for the label is kept on the left; nothing goes above the top; nothing is smaller than the grid allows.
  const squeezed = placeControl(form, 'Price', { x: 4, y: -50, w: 2, h: 1 }).controls.find((c) => c.field === 'Price');
  assert.deepEqual([squeezed.x, squeezed.y, squeezed.w, squeezed.h], [136, 0, 32, 16]);
  // A tick box keeps its own size.
  const tick = placeControl(form, 'Sold Out', { w: 300, h: 90 }).controls.find((c) => c.field === 'Sold Out');
  assert.equal(tick.w, form.controls.find((c) => c.field === 'Sold Out').w);
  assert.throws(() => placeControl(form, 'Nope', { x: 0 }), /Nope is not on the form/);
  // The wizard's own form is the size fitForm works out.
  assert.deepEqual([fitForm(form).width, fitForm(form).height], [form.width, form.height]);
});

test('a field taken off the form stays in the table and can be put back, below the others, a subform moved down for it', () => {
  const b = shop();
  const fields = b.fields('Products');
  const form = b.makeForm('Products');
  const off = takeOff(form, 'Notes');
  assert.deepEqual(off.controls.map((c) => c.field), ['ID', 'Name', 'Price', 'In Stock', 'Sold Out']);
  // The others stay where they were; the subform below holds the form's height.
  assert.deepEqual(off.controls.map((c) => c.y), form.controls.filter((c) => c.field !== 'Notes').map((c) => c.y));
  assert.equal(off.height, form.height);
  assert.deepEqual(fieldsOff(off, fields).map((f) => f.name), ['Notes']);
  const back = putBack(off, fields.find((f) => f.name === 'Notes'), (t) => b.fields(t), b.relationships());
  const notes = back.controls.at(-1);
  assert.equal(notes.field, 'Notes');
  assert.equal(notes.kind, 'longtext');
  const above = off.controls.at(-1);
  assert.ok(notes.y >= above.y + above.h);
  // The subform stands below the field put back.
  const sub = back.subforms[0];
  assert.ok(sub.y >= notes.y + notes.h, JSON.stringify({ sub, notes }));
  assert.ok(back.height >= sub.y + sub.h);
  assert.throws(() => putBack(back, fields[1], () => [], []), /Name is on the form already/);
  // The last field on the form stays.
  let one = form;
  for (const c of form.controls.slice(1)) one = takeOff(one, c.field);
  assert.throws(() => takeOff(one, 'ID'), /one field at least/);
  // A field that points at another table goes back as a list of its records.
  const lines = b.makeForm('Sale Lines');
  const product = putBack(takeOff(lines, 'Product'), b.fields('Sale Lines').find((f) => f.name === 'Product'), (t) => b.fields(t), b.relationships()).controls.at(-1);
  assert.deepEqual([product.kind, product.lookup], ['list', { table: 'Products', field: 'ID', show: 'Name' }]);
});

test('a label renamed is kept with the form; an empty one reads as the field, and the layout is saved and opened again', () => {
  const b = shop();
  const form = b.makeForm('Products');
  let changed = relabel(form, 'In Stock', '  Units on the shelf ');
  changed = relabel(changed, 'Name', '');
  changed = placeControl(changed, 'Price', { y: 400 });
  b.saveForm(changed);
  const kept = b.forms().find((f) => f.name === 'Products').layout;
  assert.equal(kept.controls.find((c) => c.field === 'In Stock').label, 'Units on the shelf');
  assert.equal(kept.controls.find((c) => c.field === 'Name').label, 'Name');
  assert.equal(kept.controls.find((c) => c.field === 'Price').y, 400);
  assert.equal(kept.height, changed.height);
  // A form made before 2.18 has no more than the wizard gave it, and reads as it did.
  const old = wizardForm('Products', (t) => b.fields(t), b.relationships());
  assert.deepEqual(Object.keys(old).sort(), ['controls', 'grid', 'height', 'name', 'subforms', 'table', 'title', 'width']);
});

test("a subform's line is added pointing at the record shown, changed where it stands, and deleted, as its table allows", () => {
  const b = shop();
  const form = b.makeForm('Sales');
  const sub = form.subforms[0];
  assert.deepEqual([sub.table, sub.field, sub.to], ['Sale Lines', 'Sale', 'ID']);
  // What the window does: a line typed at the foot, with the record shown's key in the field that points at it.
  const sale = b.rows('Sales')[0];
  b.insert(sub.table, { Product: 3, Quantity: 2, [sub.field]: sale[sub.to] });
  const lines = b.rows(sub.table, { where: { [sub.field]: sale[sub.to] } });
  assert.deepEqual(lines.map((r) => [r.Product, r.Quantity]), [[1, 2], [2, 1], [3, 2]]);
  b.update(sub.table, lines[1]._rowid, { Quantity: 4 });
  assert.equal(b.rows(sub.table, { where: { Sale: 1 } })[1].Quantity, 4);
  // A line pointing at a product that is not there is refused in plain words.
  assert.throws(() => b.insert(sub.table, { Product: 99, Quantity: 1, Sale: 1 }), /Products has no record 99, so Product cannot point at it/);
  b.remove(sub.table, [lines[2]._rowid]);
  assert.equal(b.rows(sub.table, { where: { Sale: 1 } }).length, 2);
});
