// The print template wizard's fields: a receipt read "Number 1 Number S1001"
// when the table kept a number of its own beside the record's.
import test from 'node:test';
import assert from 'node:assert/strict';
import { wizardDocument } from '../packages/database/src/document-design.js';

const make = (fields) => wizardDocument({ table: 'Sales', fieldsOf: () => fields, relationships: [] });
const labels = (doc) => {
  const block = (doc.blocks || doc.layout?.blocks || []).find((b) => b.type === 'fields');
  return block.items.map((i) => i.label || i.field);
};

test('a table with a number of its own shows that one, not the record\'s as well', () => {
  const shown = labels(make([{ name: 'ID', type: 'autonumber' }, { name: 'Number', type: 'text' }, { name: 'Customer', type: 'text' }]));
  assert.deepEqual(shown.filter((l) => /number/i.test(l)), ['Number']);
  assert.ok(!shown.includes('ID'));
});

test('a table without one shows the record\'s number, called Number', () => {
  const shown = labels(make([{ name: 'ID', type: 'autonumber' }, { name: 'Customer', type: 'text' }]));
  assert.equal(shown[0], 'Number');
});
