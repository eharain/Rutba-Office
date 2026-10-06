// Insert → Equation on a sheet: written as Excel writes one — a text box
// whose paragraph holds the OMML (a14:m) in DrawingML's own run properties,
// beside a fallback box of the linear form, in mc:AlternateContent — drawn
// in the grid as MathML, read back, changed, undone, and kept in the file.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { OoxmlPackage } from '@rutba/ooxml';
import { createDocumentService } from '../apps/desktop/main/documents.js';

const drawingXml = (bytes) => {
  const pkg = OoxmlPackage.read(bytes);
  const part = pkg.partNames().find((n) => /^xl\/drawings\/drawing\d+\.xml$/.test(n));
  return part ? pkg.text(part) : '';
};

test('an equation goes on the sheet as Excel writes one, and is drawn as math', async () => {
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const { id } = docs.new({ kind: 'sheet' });
  docs.apply({ id, ops: [{ op: 'select', row: 1, col: 1 }, { op: 'insertEquation', linear: 'x^2+y^2=r^2' }] });
  const [eq] = docs.model({ id }).drawings;
  assert.equal(eq.kind, 'equation');
  assert.match(eq.svg, /<foreignObject\b[\s\S]*<math\b[\s\S]*<msup>/, 'drawn as MathML');
  assert.equal(eq.linear.replace(/\s+/g, ''), 'x^2+y^2=r^2', 'and read back in its linear form');
  assert.equal(eq.anchor.row, 1);

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-sheq-'));
  const file = path.join(dir, 'equation.xlsx');
  docs.save({ id, path: file });
  const xml = drawingXml(fs.readFileSync(file));
  assert.match(xml, /<mc:AlternateContent xmlns:mc="[^"]+"><mc:Choice xmlns:a14="http:\/\/schemas\.microsoft\.com\/office\/drawing\/2010\/main" Requires="a14"><xdr:sp\b/);
  assert.match(xml, /<a14:m><m:oMathPara xmlns:m="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/math">/);
  assert.doesNotMatch(xml, /<w:rPr\b/, 'the runs speak DrawingML, not WordprocessingML');
  assert.match(xml, /<a:latin typeface="Cambria Math"/);
  assert.match(xml, /<mc:Fallback><xdr:sp\b[\s\S]*<a:t>x\^2\+y\^2=r\^2<\/a:t>/, 'a reader without math gets the words');

  // Opened again, it is still an equation.
  const again = await docs.open({ path: file });
  assert.equal(docs.model({ id: again.id }).drawings[0].kind, 'equation');
});

test('an equation typed again keeps its box, and one undo puts the old one back', () => {
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const { id } = docs.new({ kind: 'sheet' });
  docs.apply({ id, ops: [{ op: 'insertEquation', linear: 'a+b' }] });
  const before = docs.model({ id }).drawings[0];
  docs.apply({ id, ops: [{ op: 'setEquation', id: before.id, linear: '\\sqrt(a^2+b^2)' }] });
  const after = docs.model({ id }).drawings[0];
  assert.equal(after.id, before.id);
  assert.deepEqual([after.x, after.y, after.width, after.height], [before.x, before.y, before.width, before.height], 'the same box');
  assert.match(after.svg, /<msqrt>/);
  docs.undo({ id });
  assert.equal(docs.model({ id }).drawings[0].linear.replace(/\s+/g, ''), 'a+b');
  assert.throws(() => docs.apply({ id, ops: [{ op: 'insertEquation', linear: '   ' }] }), /Type the equation/);
});
