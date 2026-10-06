// Insert → WordArt on a sheet: written as Excel writes it — a text box with
// no fill and no line, sized to its words, the words centred and carrying
// their look in the run (an outline, no fill for an outline alone, a glow
// and a shadow) — drawn in the grid with a stroke round the letters and the
// glow and shadow round the words; its words changed by a double-click
// keep the look; and the file keeps all of it.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { OoxmlPackage } from '@rutba/ooxml';
import { textLookXml } from '@rutba/ooxml/build';
import { createDocumentService } from '../apps/desktop/main/documents.js';

const drawingXml = (bytes) => {
  const pkg = OoxmlPackage.read(bytes);
  const part = pkg.partNames().find((n) => /^xl\/drawings\/drawing\d+\.xml$/.test(n));
  return part ? pkg.text(part) : '';
};
const glow = { radiusPt: 5, color: '#4472C4', alpha: 0.4 };
const shadow = { blurPx: 4, distPx: 4, dir: 45, color: '#000000', alpha: 0.43 };

test('a run\'s WordArt look is written in the order the schema wants', () => {
  assert.equal(textLookXml({ color: '#FFFFFF', outline: { width: 1, color: '#4472C4' }, textEffects: { glow, shadow } }),
    '<a:ln w="12700"><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></a:ln><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
    + '<a:effectLst><a:glow rad="63500"><a:srgbClr val="4472C4"><a:alpha val="40000"/></a:srgbClr></a:glow>'
    + '<a:outerShdw blurRad="38100" dist="38100" dir="2700000" algn="tl" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="43000"/></a:srgbClr></a:outerShdw></a:effectLst>');
  assert.equal(textLookXml({ noFill: true, outline: { width: 1.5, color: '#ED7D31' } }), '<a:ln w="19050"><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill></a:ln><a:noFill/>');
  assert.equal(textLookXml({}), '');
});

test('WordArt goes on the sheet as Excel writes it, is drawn with its look, and its words change without losing it', () => {
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const { id } = docs.new({ kind: 'sheet' });
  docs.apply({ id, ops: [{ op: 'select', row: 2, col: 1 }, { op: 'insertWordArt', text: 'Your text here', style: { color: '#FFFFFF', outline: { width: 1, color: '#4472C4' }, textEffects: { glow, shadow } } }] });
  const [art] = docs.model({ id }).drawings;
  assert.equal(art.kind, 'shape');
  assert.equal(art.textBox, true);
  assert.equal(art.text, 'Your text here');
  assert.equal(art.anchor.row, 2);
  assert.match(art.svg, /<text\b[^>]*fill="#ffffff"[^>]*stroke="#4472c4"[^>]*stroke-width="1\.33"[^>]*style="filter:drop-shadow\(0 0 [\d.]+px rgba\(68,114,196,0\.4\)\) drop-shadow\([\d.]+px [\d.]+px 2px rgba\(0,0,0,0\.43\)\)"[^>]*>Your text here<\/text>/);

  // Its words changed, as a double-click changes them: the look stays.
  docs.apply({ id, ops: [{ op: 'setShapeText', id: art.id, text: 'Sales\nby region' }] });
  const [again] = docs.model({ id }).drawings;
  assert.equal(again.text.replace(/\s+/g, ' '), 'Sales by region');
  assert.match(again.svg, /stroke="#4472c4"/);

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-shwa-'));
  const file = path.join(dir, 'wordart.xlsx');
  docs.save({ id, path: file });
  const xml = drawingXml(fs.readFileSync(file));
  assert.match(xml, /<xdr:cNvSpPr txBox="1"\/>/);
  assert.match(xml, /<xdr:spPr><a:prstGeom prst="rect"><a:avLst\/><\/a:prstGeom><a:noFill\/><\/xdr:spPr>/);
  assert.match(xml, /<a:bodyPr wrap="none"[^>]*><a:spAutoFit\/><\/a:bodyPr>/);
  assert.equal((xml.match(/<a:p><a:pPr algn="ctr"\/><a:r><a:rPr lang="en-US" sz="3600" b="0" cap="none" spc="0"><a:ln w="12700">/g) || []).length, 2, 'each line a centred paragraph with the look');
  assert.match(xml, /<a:t>by region<\/a:t>/);

  // Opened again, it is still WordArt.
  const back = docs.open({ path: file });
  const [read] = docs.model({ id: back.id }).drawings;
  assert.match(read.svg, /stroke="#4472c4"/);
  assert.match(read.svg, /drop-shadow/);
});

test('an outline alone is drawn unfilled, and a plain text box\'s words change too', () => {
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const { id } = docs.new({ kind: 'sheet' });
  docs.apply({ id, ops: [{ op: 'insertWordArt', text: 'Hollow', style: { noFill: true, outline: { width: 1.5, color: '#ED7D31' } } }] });
  assert.match(docs.model({ id }).drawings[0].svg, /<text\b[^>]*fill="none"[^>]*stroke="#ed7d31"/);
  docs.apply({ id, ops: [{ op: 'select', row: 10, col: 1 }, { op: 'insertShape', geometry: 'rect', text: 'Text' }] });
  const box = docs.model({ id }).drawings[1];
  docs.apply({ id, ops: [{ op: 'setShapeText', id: box.id, text: 'Notes on Q3' }] });
  assert.equal(docs.model({ id }).drawings[1].text, 'Notes on Q3');
  assert.match(docs.model({ id }).drawings[1].svg, />Notes on Q3<\/text>/);
});
