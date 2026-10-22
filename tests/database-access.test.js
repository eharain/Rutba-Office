// People, roles and protection (D-2026-10-22-11): a protected database asks
// who you are, and each role sees and does what the Admin allowed, enforced
// beneath queries, reports, print templates and lookups alike.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Base, buildFromSpec } from '../packages/database/src/index.js';
import { hashPassword, checkPassword } from '../packages/database/src/access.js';

const ref = (t) => ({ type: 'integer', points: t });

function shop(file = ':memory:') {
  const b = Base.open(file);
  buildFromSpec(b, {
    tables: [
      { name: 'Products', fields: [{ name: 'Name', type: 'text' }, { name: 'Price', type: 'currency' }, { name: 'Cost', type: 'currency' }] },
      { name: 'Sales', fields: [{ name: 'Product', ...ref('Products') }, { name: 'Quantity', type: 'integer' }] },
      { name: 'Wages', fields: [{ name: 'Who', type: 'text' }, { name: 'Amount', type: 'currency' }] },
    ],
  });
  b.insert('Products', { Name: 'Bread', Price: 2.5, Cost: 1.1 });
  b.insert('Products', { Name: 'Milk', Price: 1.2, Cost: 0.7 });
  b.insert('Wages', { Who: 'Amira', Amount: 900 });
  b.saveQuery('Margins', { tables: ['Products'], columns: [{ table: 'Products', field: 'Name' }, { expr: '[Price] - [Cost]', as: 'Margin' }] });
  b.saveQuery('Prices', { tables: ['Products'], columns: [{ table: 'Products', field: 'Name' }, { table: 'Products', field: 'Price' }] });
  return b;
}

const CASHIER = {
  tables: {
    default: 'read',
    each: {
      Products: { access: 'read', fields: { Cost: 'hidden' } },
      Sales: { access: 'change', own: true },
      Wages: 'hidden',
    },
  },
  hidden: { reports: ['Takings'] },
};

/** A shop protected by Amira, with Ben as a cashier. */
function protectedShop(file) {
  const b = shop(file);
  b.protect({ name: 'Amira', password: 'owner-pass' });
  b.saveRole('Cashier', CASHIER);
  const ben = b.addPerson({ name: 'Ben', role: 'Cashier', password: 'till-pass' });
  return { b, ben };
}

test('a password is kept as scrypt\'s answer, and checked against it', () => {
  const kept = hashPassword('open sesame');
  assert.match(kept, /^scrypt\$16384\$8\$1\$/);
  assert.ok(!kept.includes('open sesame'));
  assert.equal(checkPassword('open sesame', kept), true);
  assert.equal(checkPassword('open Sesame', kept), false);
  assert.throws(() => hashPassword('abc'), /at least 4/);
});

test('a database that is not protected works as it always did', () => {
  const b = shop();
  try {
    assert.equal(b.isProtected(), false);
    assert.equal(b.person(), null);
    assert.equal(b.rights(), null);
    assert.equal(b.tables().length, 3);
    assert.equal(b.rows('Products')[0].Cost, 1.1);
    assert.equal(b.query('SELECT count(*) AS n FROM Wages').rows[0].n, 1);
  } finally { b.close(); }
});

test('protected, it opens locked until someone signs in, and a wrong password is refused', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'access-'));
  const file = path.join(dir, 'shop.rbase');
  const { b } = protectedShop(file);
  b.close();
  const again = Base.open(file);
  try {
    assert.equal(again.isProtected(), true);
    assert.deepEqual(again.tables(), []);
    assert.throws(() => again.rows('Products'), /Sign in/);
    assert.throws(() => again.query('SELECT 1'), /Sign in/);
    assert.deepEqual(again.signInNames(), ['Amira', 'Ben']);
    assert.throws(() => again.signIn('Ben', 'wrong'), /do not match/);
    assert.throws(() => again.signIn('Nobody', 'till-pass'), /do not match/);
    assert.deepEqual(again.signIn('ben', 'till-pass'), { id: 2, name: 'Ben', role: 'Cashier', admin: false, seeing: null });
    again.signOut();
    assert.throws(() => again.rows('Products'), /Sign in/);
  } finally { again.close(); fs.rmSync(dir, { recursive: true, force: true }); }
});

test('a role sees the tables and fields it may, and nothing reads past it: queries, lookups, SQL', () => {
  const { b } = protectedShop();
  try {
    b.signOut();
    b.signIn('Ben', 'till-pass');
    assert.deepEqual(b.tables().map((t) => t.name), ['Products', 'Sales']);
    assert.deepEqual(b.fields('Products').map((f) => [f.name, Boolean(f.readOnly)]), [['ID', true], ['Name', true], ['Price', true]]);
    assert.deepEqual(Object.keys(b.rows('Products')[0]).sort(), ['ID', 'Name', 'Price', '_rowid']);
    assert.throws(() => b.rows('Wages'), /no table called Wages/);
    // A saved query over a hidden field fails rather than show it; one over what the role sees runs.
    assert.deepEqual(b.queries().map((x) => x.name), ['Margins', 'Prices']);
    assert.throws(() => b.runQuery({ name: 'Margins' }));
    assert.deepEqual(b.runQuery({ name: 'Prices' }).rows.map((r) => r.Price), [2.5, 1.2]);
    // A design that names the hidden field is refused too, in the designer's own words or SQLite's.
    assert.throws(() => b.runQuery({ design: { tables: ['Products'], columns: [{ table: 'Products', field: 'Cost' }] } }));
    assert.throws(() => b.query('SELECT * FROM main.Wages'), /design/);
    assert.deepEqual(b.lookup({ table: 'Products', field: 'ID', show: ['Name', 'Cost'] }).map((o) => o.text), ['Bread', 'Milk']);
    assert.deepEqual(b.lookup({ table: 'Wages', field: 'ID', show: 'Who' }), []);
  } finally { b.close(); }
});

test('a role adds and changes what it may, its own records only where it says so, and the trail says who', () => {
  const { b } = protectedShop();
  try {
    const mine = b.insert('Sales', { Product: 1, Quantity: 2 });
    b.signOut();
    b.signIn('Ben', 'till-pass');
    // Amira's sale is not Ben's to see or change.
    assert.deepEqual(b.rows('Sales'), []);
    assert.throws(() => b.update('Sales', mine, { Quantity: 9 }), /that record/);
    const his = b.insert('Sales', { Product: 2, Quantity: 1 });
    assert.equal(b.update('Sales', his, { Quantity: 3 }), 1);
    assert.deepEqual(b.rows('Sales').map((r) => r.Quantity), [3]);
    assert.throws(() => b.remove('Sales', [his]), /may not delete/);
    assert.throws(() => b.insert('Products', { Name: 'Eggs' }), /may not add/);
    assert.throws(() => b.update('Products', 1, { Price: 9 }), /may not change/);
    assert.deepEqual(b.trailOf('Sales', his).map((t) => [t.person, t.action, t.fields]), [['Ben', 'add', null], ['Ben', 'change', ['Quantity']]]);
    // Design is not the cashier's.
    assert.throws(() => b.createTable({ name: 'Notes', fields: [] }), /design/);
    assert.throws(() => b.people(), /Admin/);
  } finally { b.close(); }
});

test('a read-only field may be given as it is, and is refused when changed', () => {
  const { b } = protectedShop();
  try {
    b.saveRole('Pricing', { tables: { default: 'read', each: { Products: { access: 'change', fields: { Name: 'read', Cost: 'hidden' } } } } });
    b.addPerson({ name: 'Chen', role: 'Pricing', password: 'price-pass' });
    b.signOut();
    b.signIn('Chen', 'price-pass');
    assert.equal(b.update('Products', 1, { Name: 'Bread', Price: 2.75 }), 1);
    assert.throws(() => b.update('Products', 1, { Name: 'Loaf' }), /read only/);
    assert.throws(() => b.update('Products', 1, { Cost: 0 }), /no field/);
    b.signOut();
    b.signIn('Amira', 'owner-pass');
    assert.deepEqual(b.rows('Products')[0], { _rowid: 1, ID: 1, Name: 'Bread', Price: 2.75, Cost: 1.1 });
  } finally { b.close(); }
});

test('saved items a role hides, or that stand on a table it hides, are not listed', () => {
  const { b } = protectedShop();
  try {
    b.makeReport({ table: 'Sales', name: 'Takings' });
    b.makeReport({ table: 'Wages', name: 'Payroll' });
    b.makeReport({ table: 'Products', name: 'Price list' });
    b.makeForm('Wages');
    b.makeForm('Sales');
    b.makeDocument({ table: 'Sales', name: 'Receipt' });
    b.signOut();
    b.signIn('Ben', 'till-pass');
    assert.deepEqual(b.reports().map((r) => r.name), ['Price list']);
    assert.ok(b.forms().some((f) => f.layout.table === 'Sales'));
    assert.ok(!b.forms().some((f) => f.layout.table === 'Wages'));
    assert.deepEqual(b.documents().map((d) => d.name), ['Receipt']);
  } finally { b.close(); }
});

test('the Admin sees the database as a role would, reading only, then goes back', () => {
  const { b } = protectedShop();
  try {
    assert.equal(b.person().admin, true);
    assert.equal(b.seeAs('Cashier').seeing, 'Cashier');
    assert.deepEqual(b.tables().map((t) => t.name), ['Products', 'Sales']);
    assert.throws(() => b.insert('Sales', { Product: 1 }), /seeing the database as Cashier/);
    assert.throws(() => b.saveRole('Cashier', {}), /Admin/);
    assert.equal(b.seeAs(null).seeing, null);
    assert.equal(b.tables().length, 3);
    assert.equal(b.rows('Products')[0].Cost, 1.1);
  } finally { b.close(); }
});

test('people and roles: the last Admin stays, a role in use stays, and a password is changed with the one before it', () => {
  const { b, ben } = protectedShop();
  try {
    assert.deepEqual(b.roles().map((r) => [r.name, r.builtIn]), [['Admin', true], ['Staff', false], ['Read only', false], ['Cashier', false]]);
    assert.throws(() => b.saveRole('Admin', {}), /everything/);
    assert.throws(() => b.removeRole('Cashier'), /Give them another/);
    assert.throws(() => b.changePerson(1, { role: 'Staff' }), /at least one Admin/);
    assert.throws(() => b.removePerson(1), /yourself/);
    assert.throws(() => b.addPerson({ name: 'ben', role: 'Staff', password: 'xxxx' }), /already/);
    b.setPassword(ben, 'new-till');
    b.signOut();
    b.signIn('Ben', 'new-till');
    assert.throws(() => b.setPassword(ben, 'another', { current: 'wrong' }), /not right/);
    b.setPassword(ben, 'another', { current: 'new-till' });
    assert.throws(() => b.setPassword(1, 'take-over'), /Admin/);
    b.signOut();
    assert.ok(b.signIn('Ben', 'another'));
  } finally { b.close(); }
});

test('the protection taken away by an Admin with their password: the people go, the records stay', () => {
  const { b } = protectedShop();
  try {
    assert.throws(() => b.unprotect({ password: 'wrong' }), /not right/);
    b.unprotect({ password: 'owner-pass' });
    assert.equal(b.isProtected(), false);
    assert.equal(b.person(), null);
    assert.equal(b.rows('Wages').length, 1);
    assert.deepEqual(b.signInNames(), []);
  } finally { b.close(); }
});
