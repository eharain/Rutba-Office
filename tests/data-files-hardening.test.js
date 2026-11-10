// Data Files where a role, a long key or a redirect could have gone wrong, and marks measured on a PDF's own glyphs.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { Base } from '@rutba/database';
import { fromEngine as fromPostgres } from '../apps/desktop/main/engines/postgres.js';
import { fromEngine as fromMysql } from '../apps/desktop/main/engines/mysql.js';
import { readJsonSource } from '../apps/desktop/main/base-sources.js';
import { createBaseService } from '../apps/desktop/main/base.js';

const require = createRequire(import.meta.url);

function roles() {
  const b = Base.open(':memory:');
  b.createTable({ name: 'Pay', fields: [{ name: 'Who', type: 'text' }, { name: 'Salary', type: 'number' }] });
  b.protect({ name: 'Boss', password: 'secret-for-the-test' });
  return b;
}
const rightsOf = (b, role) => JSON.parse(b.db.prepare('SELECT rights FROM rutba_roles WHERE name = ?').get(role).rights);

test('a table a role has hidden stays hidden after it is renamed', () => {
  const b = roles();
  b.saveRole('Clerk', { tables: { default: 'read', each: { Pay: 'hidden' } } });
  b.renameTable('Pay', 'Wages');
  const each = rightsOf(b, 'Clerk').tables.each;
  assert.deepEqual(Object.keys(each), ['Wages']);
  assert.equal(each.Wages.access ?? each.Wages, 'hidden');
  b.close();
});

test('a field a role has hidden stays hidden after it is renamed', () => {
  const b = roles();
  b.saveRole('Clerk', { tables: { default: 'read', each: { Pay: { access: 'read', fields: { Salary: 'hidden' } } } } });
  const field = b.fields('Pay').find((f) => f.name === 'Salary');
  b.changeField('Pay', 'Salary', { ...field, name: 'Monthly' });
  assert.deepEqual(rightsOf(b, 'Clerk').tables.each.Pay.fields, { Monthly: 'hidden' });
  b.close();
});

test('the records a person added are still theirs after the table is renamed', () => {
  const b = roles();
  b.db.prepare("INSERT INTO rutba_trail (at, person, tbl, row, action, detail) VALUES ('2026-11-10', 1, 'Pay', 1, 'add', NULL)").run();
  b.renameTable('Pay', 'Wages');
  assert.deepEqual(b.db.prepare('SELECT tbl FROM rutba_trail').all().map((r) => r.tbl), ['Wages']);
  b.close();
});

test('a locked database answers no query, saved or designed', () => {
  const b = roles();
  b.access = { locked: true };
  const design = { tables: ['Pay'], columns: [{ table: 'Pay', field: 'Who' }] };
  assert.throws(() => b.runQuery({ design }), /Sign in/);
  assert.throws(() => b.queryRows({ design }), /Sign in/);
  b.access = null;
  b.close();
});

test('a role narrowed to some tables cannot run SQL of its own through a query page', () => {
  const b = roles();
  b.access = { locked: false, full: false, rights: { tables: { default: 'read', each: {} }, hidden: {} } };
  assert.throws(() => b.queryRows({ design: { sql: 'SELECT * FROM main."Pay"' } }), /role/);
  b.access = null;
  b.close();
});

const LONG = '9007199254740993';

test('a key longer than a JavaScript number keeps every digit from each engine', () => {
  assert.equal(fromPostgres('integer', LONG), LONG);
  assert.equal(fromMysql('integer', LONG), LONG);
  assert.equal(fromPostgres('integer', '42'), 42);
  assert.equal(fromMysql('integer', '42'), 42);
  assert.equal(fromPostgres('number', '1.50'), 1.5);
});

test('a linked table with such keys reads, and names the record exactly', () => {
  const b = Base.open(':memory:');
  const link = b.saveLink({ name: 'Away', engine: 'postgres', settings: { host: 'db.example', database: 'shop', user: 'office' } });
  b.linkTable({ table: 'Stock', link, remote: { schema: 'public', name: 'stock' }, key: 'id', fields: [{ name: 'id', type: 'integer' }, { name: 'name', type: 'text' }] });
  b.loadLinked('Stock', [{ id: LONG, name: 'Bread' }, { id: 7, name: 'Milk' }]);
  const rows = b.rows('Stock', { limit: 10 });
  assert.deepEqual(rows.map((r) => r.id).sort(), [7, LONG].sort());
  const long = rows.find((r) => r.name === 'Bread');
  assert.equal(b.linkedKey('Stock', long.rowid ?? LONG), LONG);
  b.close();
});

const answer = (status, body, headers = {}) => ({ status, ok: status >= 200 && status < 300, statusText: '', headers: { get: (k) => headers[k.toLowerCase()] ?? null }, json: async () => body });

test('a secret header stops at the origin it was given for, through a redirect too', async () => {
  const seen = [];
  const fetchImpl = async (url, opts) => {
    seen.push({ url, headers: opts.headers, redirect: opts.redirect });
    if (url.startsWith('https://api.example/')) return answer(302, null, { location: 'https://elsewhere.example/data' });
    return answer(200, [{ a: 1 }]);
  };
  const out = await readJsonSource({ where: 'https://api.example/items', headers: [{ name: 'X-Api-Key', secret: true }] }, { secretOf: () => 'the-key', fetchImpl });
  assert.equal(seen.length, 2);
  assert.equal(seen[0].redirect, 'manual');
  assert.equal(seen[0].headers['X-Api-Key'], 'the-key');
  assert.equal(seen[1].headers['X-Api-Key'], undefined);
  assert.ok(out.rows.length === 1 || out.length === 1);
});

test('a redirect within the origin keeps the header, and a loop ends', async () => {
  const seen = [];
  const fetchImpl = async (url, opts) => {
    seen.push(opts.headers['X-Api-Key']);
    return answer(302, null, { location: '/again' });
  };
  await assert.rejects(readJsonSource({ where: 'https://api.example/items', headers: [{ name: 'X-Api-Key', secret: true }] }, { secretOf: () => 'the-key', fetchImpl }), /could not be reached/);
  assert.equal(seen.length, 11);
  assert.ok(seen.every((v) => v === 'the-key'));
});

test('a live table past the limit says so and keeps what it read', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'live-cut-'));
  const file = path.join(dir, 'shop.rbase');
  const asked = [];
  const fields = [{ name: 'id', type: 'integer' }];
  const engine = {
    describe: async () => ({ fields, key: 'id' }),
    read: async (remote, o) => { asked.push(o.limit); return Array.from({ length: o.limit }, (_, i) => ({ id: i + 1 })); },
    close: async () => {},
  };
  const svc = createBaseService({ secrets: null, connect: async () => engine });
  try {
    const b = Base.open(file);
    const link = b.saveLink({ name: 'Away', engine: 'postgres', settings: { host: 'db.example', database: 'shop', user: 'office' } });
    b.linkTable({ table: 'Stock', link, remote: { schema: 'public', name: 'stock' }, key: 'id', fields });
    b.close();
    const opened = await svc.open({ file });
    const s = await svc.allowSources({ id: opened.id, addresses: opened.ask.map((a) => a.address) });
    const stock = s.tables.find((t) => t.name === 'Stock');
    assert.equal(asked[0], 100001);
    assert.equal(stock.count, 100000);
    assert.equal(stock.cut, 100000);
  } finally {
    svc.closeAll();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

/* Mark Matches */

function pdfWith(content, font) {
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 600 800] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>',
    `<< /Length ${content.length} >>\nstream\n${content}\nendstream`,
    font,
  ];
  let out = '%PDF-1.4\n';
  const at = [];
  objects.forEach((o, i) => { at.push(out.length); out += `${i + 1} 0 obj\n${o}\nendobj\n`; });
  const start = out.length;
  out += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n${at.map((a) => `${String(a).padStart(10, '0')} 00000 n \n`).join('')}trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${start}\n%%EOF\n`;
  return Buffer.from(out, 'latin1');
}
// Every "i" is thin and every "m" wide: a stand-in font would put a word in the middle of this line far from where it is.
const FONT = (() => {
  const widths = Array.from({ length: 256 }, (_, c) => (c === 105 ? 100 : c === 109 ? 1500 : 500));
  return `<< /Type /Font /Subtype /TrueType /BaseFont /Thin /FirstChar 0 /LastChar 255 /Widths [${widths.join(' ')}] /Encoding /WinAnsiEncoding >>`;
})();

test('a word is measured where the page draws it, not where a stand-in font would', () => {
  const { pageGlyphs, quadsOf, applyRedactions, bytesContain } = require('@rutba/pdf/redact');
  const pdf = pdfWith('BT /F1 10 Tf 50 700 Td (iiiiiiiiii SECRET mmmm) Tj ET', FONT);
  const { text, boxes, unsure } = pageGlyphs(pdf, 0);
  assert.equal(unsure, false);
  assert.equal(text.trim(), 'iiiiiiiiii SECRET mmmm');
  const at = text.indexOf('SECRET');
  const quads = quadsOf(boxes, at, at + 6);
  assert.equal(quads.length, 1);
  // Ten letters of 1, a space of 5, then six of 5, at size 10.
  assert.deepEqual([quads[0][0], quads[0][2]], [50 + 10 + 5, 50 + 10 + 5 + 30]);
  const rect = [quads[0][4], quads[0][5], quads[0][2], quads[0][3]];
  const done = applyRedactions(pdf, { areas: [{ page: 0, rects: [rect] }] });
  assert.equal(done.removed.glyphs, 6);
  assert.equal(bytesContain(done.bytes, 'SECRET'), false);
  assert.equal(bytesContain(done.bytes, 'iiiiiiiiii'), true);
});

test('two lines give a mark each, and words in separate strings are parted', () => {
  const { pageGlyphs, quadsOf } = require('@rutba/pdf/redact');
  const pdf = pdfWith('BT /F1 10 Tf 50 700 Td (one) Tj 0 -20 Td (two) Tj ET', FONT);
  const { text, boxes } = pageGlyphs(pdf, 0);
  assert.equal(text.trim(), 'one two');
  const start = text.indexOf('one');
  assert.equal(quadsOf(boxes, start, text.indexOf('two') + 3).length, 2);
});
