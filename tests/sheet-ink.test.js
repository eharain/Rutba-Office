// Draw on a sheet — the engine's half.
//
// A stroke goes on the sheet as a freeform drawing named "Ink N", anchored
// to the cells under it to the offset, a smooth path in the pen's colour
// and width, see-through for a highlighter; the grid draws it so, and the
// file keeps it. Ink to Shape's rectangle, oval or triangle goes at the box
// drawn, outlined in the pen's colour and not filled.

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

test('strokes go on the sheet as ink, drawn as they were laid, and kept in the file', () => {
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const { id } = docs.new({ kind: 'sheet' });
  docs.apply({ id, ops: [{ op: 'insertInk', strokes: [
    { points: [[100, 50], [130, 60], [170, 90], [200, 95]], color: '#E81123', width: 1, alpha: 1, cap: 'rnd' },
    { points: [[80, 200], [400, 200]], color: '#FFFF00', width: 8, alpha: 0.5, cap: 'sq' },
  ] }] });
  const drawings = docs.model({ id }).drawings;
  assert.deepEqual(drawings.map((d) => d.name), ['Ink 1', 'Ink 2']);
  const [pen, hl] = drawings;
  // The box holds the stroke's width too: half of it, and a pixel, on every side.
  const pad = (1 * 96) / 72 / 2 + 1;
  assert.ok(Math.abs(pen.x - (100 - pad)) < 1.5 && Math.abs(pen.y - (50 - pad)) < 1.5 && Math.abs(pen.width - (100 + 2 * pad)) < 1.5 && Math.abs(pen.height - (45 + 2 * pad)) < 1.5, JSON.stringify(pen));
  assert.ok(hl.height > 10, 'a straight highlighter stroke is as tall as it is thick');
  assert.match(pen.svg, /<path d="M[^"]*Q[^"]*" fill="none" stroke="#e81123" stroke-width="1" stroke-linecap="round" stroke-linejoin="round"/i);
  assert.match(hl.svg, /stroke-linecap="square"[^>]*opacity="0\.5"/);

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-shink-'));
  const file = path.join(dir, 'ink.xlsx');
  docs.save({ id, path: file });
  const xml = drawingXml(fs.readFileSync(file));
  assert.match(xml, /<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>1<\/xdr:col><xdr:colOff>\d+<\/xdr:colOff>/);
  assert.match(xml, /<xdr:cNvPr id="\d+" name="Ink 1"\/>/);
  assert.match(xml, /<a:custGeom>[\s\S]*<a:quadBezTo>/);
  assert.match(xml, /<a:ln w="101600" cap="sq"><a:solidFill><a:srgbClr val="FFFF00"><a:alpha val="50000"\/><\/a:srgbClr>/);
  const back = docs.open({ path: file });
  assert.match(docs.model({ id: back.id }).drawings[0].svg, /stroke-linecap="round"/);

  // The Eraser's way: the drawings by id.
  docs.apply({ id, ops: [{ op: 'deleteDrawings', ids: [pen.id] }] });
  assert.deepEqual(docs.model({ id }).drawings.map((d) => d.name), ['Ink 2']);
});

test('Ink to Shape puts the shape at the box drawn, outlined in the pen\'s colour and not filled', () => {
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const { id } = docs.new({ kind: 'sheet' });
  docs.apply({ id, ops: [{ op: 'insertShapeAt', geometry: 'ellipse', x: 150, y: 80, width: 200, height: 100, line: { color: '#0078D4', width: 2 } }] });
  const [shape] = docs.model({ id }).drawings;
  assert.ok(Math.abs(shape.x - 150) < 1.5 && Math.abs(shape.width - 200) < 1.5);
  assert.match(shape.svg, /<ellipse\b[^>]*fill="none"[^>]*stroke="#0078d4"/i);
});
