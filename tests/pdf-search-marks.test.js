// Where a search marks a page for redaction: the page's own glyphs, words
// inside a form at the form's place on the page, and the window using those
// boxes only when they are sure and agree with what it found (second review, 10 Nov 2026).
import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { findAll } from '@rutba/editing/find';
import { exactQuads } from '../apps/desktop/renderer/apps/pdf/quads.js';

const require = createRequire(import.meta.url);
const { pageGlyphs, quadsOf } = require('@rutba/pdf/redact');

const FONT = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>';

/** A PDF of one page from the objects given, numbered from 4. */
function pdf({ content, resources = '', objects = [] }) {
  const streamOf = (data, dict = '') => `<< ${dict} /Length ${Buffer.byteLength(data, 'latin1')} >>\nstream\n${data}\nendstream`;
  const all = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> ${resources} >> >>`,
    streamOf(content),
    FONT,
    ...objects.map((o) => (typeof o === 'string' ? o : streamOf(o.data, o.dict))),
  ];
  let out = '%PDF-1.7\n';
  const offsets = [];
  all.forEach((body, i) => {
    offsets.push(Buffer.byteLength(out, 'latin1'));
    out += `${i + 1} 0 obj\n${body}\nendobj\n`;
  });
  const at = Buffer.byteLength(out, 'latin1');
  out += `xref\n0 ${all.length + 1}\n0000000000 65535 f \n${offsets.map((o) => `${String(o).padStart(10, '0')} 00000 n \n`).join('')}`;
  out += `trailer\n<< /Size ${all.length + 1} /Root 1 0 R >>\nstartxref\n${at}\n%%EOF\n`;
  return Buffer.from(out, 'latin1');
}

/** What the main process answers the window: each match as QuadPoints on the page's glyphs. */
function matchQuads(bytes, page, query) {
  const { text, boxes, unsure } = pageGlyphs(bytes, page);
  const matches = unsure ? [] : findAll(text, query);
  return { count: matches.length, unsure, quads: matches.flatMap((m) => quadsOf(boxes, m.start, m.end)) };
}

test('words inside a form are read where the form puts them on the page', () => {
  const form = { data: 'BT /F1 12 Tf 5 5 Td (inside) Tj ET', dict: '/Type /XObject /Subtype /Form /BBox [0 0 200 50] /Matrix [1 0 0 1 10 20] /Resources << /Font << /F1 5 0 R >> >>' };
  const file = pdf({ content: 'BT /F1 12 Tf 50 700 Td (outside) Tj ET q 1 0 0 1 300 400 cm /Fm Do Q', resources: '/XObject << /Fm 6 0 R >>', objects: [form] });
  const read = pageGlyphs(file, 0);
  assert.match(read.text, /outside/);
  assert.match(read.text, /inside/);
  const box = read.boxes[read.text.indexOf('inside')];
  // 300 + 10 + 5 across, 400 + 20 + 5 up: the page's place, the form's own and the text's.
  assert.ok(Math.abs(box[0] - 315) < 0.5, `starts at ${box[0]}`);
  assert.ok(box[1] <= 425 && box[3] > 425, `on the line at 425: ${box}`);
});

test('a search marks each match by its own glyphs, one quad a line, inside a form too', () => {
  const form = { data: 'BT /F1 12 Tf 5 5 Td (Ledger total) Tj ET', dict: '/Type /XObject /Subtype /Form /BBox [0 0 200 50] /Resources << /Font << /F1 5 0 R >> >>' };
  const file = pdf({ content: 'BT /F1 12 Tf 50 700 Td (Ledger) Tj ET q 1 0 0 1 300 400 cm /Fm Do Q', resources: '/XObject << /Fm 6 0 R >>', objects: [form] });
  const found = matchQuads(file, 0, 'Ledger');
  assert.equal(found.unsure, false);
  assert.equal(found.count, 2);
  assert.equal(found.quads.length, 2);
  const [onPage, inForm] = found.quads.map((q) => ({ x0: q[0], x1: q[2], top: q[1], bottom: q[5] }));
  assert.ok(Math.abs(onPage.x0 - 50) < 0.5 && onPage.bottom < 700 && onPage.top > 700, JSON.stringify(onPage));
  assert.ok(Math.abs(inForm.x0 - 305) < 0.5 && inForm.bottom < 405 && inForm.top > 405, JSON.stringify(inForm));
  // Six letters of Helvetica at 12 points, not the whole line.
  assert.ok(inForm.x1 - inForm.x0 > 30 && inForm.x1 - inForm.x0 < 45, `as wide as the word: ${inForm.x1 - inForm.x0}`);
});

test('the window uses the page\'s own boxes only when they are sure and found what it found', () => {
  const quads = [[1, 2, 3, 2, 1, 0, 3, 0]];
  assert.deepEqual(exactQuads(1, { count: 1, unsure: false, quads }), quads);
  assert.equal(exactQuads(2, { count: 1, unsure: false, quads }), null, 'a match the glyphs did not find: the window widens its own');
  assert.equal(exactQuads(1, { count: 2, unsure: false, quads }), null, 'more found on the glyphs than in the text');
  assert.equal(exactQuads(1, { count: 1, unsure: true, quads }), null, 'a font whose widths are a guess');
  assert.equal(exactQuads(1, null), null, 'no answer from the main process');
  assert.equal(exactQuads(0, { count: 0, unsure: false, quads: [] }), null, 'nothing found');
});
