// OpenDocument custom shapes, worked out into their outlines.
//
// Office and LibreOffice write a shape to ODF as an enhanced geometry: a
// path over equations and modifiers. Worked out, the oval from PowerPoint's
// showcase.odp lies on its ellipse, the smile's mouth is a line not
// filled, and a shape the transform turned and moved stands where
// PowerPoint had it, turned as it was.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { readZip } from '@rutba/ooxml/zip';
import { parse, all, first } from '../packages/office-formats/src/xml.js';
import { enhancedFigures, figuresGeometryXml } from '../packages/office-formats/src/odf-geometry.js';
import { readOdf } from '../packages/office-formats/src/odf.js';
import { createDocumentService } from '../apps/desktop/main/documents.js';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const fixture = (name) => fs.readFileSync(path.join(HERE, 'fixtures', 'rich', name));
/** An enhanced geometry element, from its attributes and equations. */
const geometry = (attrs, equations = {}) => ({
  name: 'draw:enhanced-geometry',
  attrs,
  children: Object.entries(equations).map(([name, formula]) => ({ name: 'draw:equation', attrs: { 'draw:name': name, 'draw:formula': formula }, children: [] })),
});
const ends = (figure) => figure.commands.filter((c) => c.op !== 'Z').map((c) => c.pts[c.pts.length - 1].map((v) => Math.round(v * 100) / 100));

test('equations over the view box, the modifiers and each other, and the functions', () => {
  const g = geometry({ 'svg:viewBox': '0 0 100 200', 'draw:modifiers': '25', 'draw:enhanced-path': 'M 0 0 L ?f0 ?f1 ?f2 0 ?f3 ?f4 Z N' }, {
    f0: 'width/2', f1: 'bottom-$0', f2: '?f0+min(10,20)*2', f3: 'if(-1,5,7)', f4: 'sqrt(?f5)+abs(-3)', f5: '16',
  });
  const [f] = enhancedFigures(g, { width: 100, height: 200 });
  assert.deepEqual(ends(f), [[0, 0], [50, 175], [70, 0], [7, 7]]);
  assert.equal(f.fill, true);
  assert.equal(enhancedFigures(geometry({ 'draw:enhanced-path': 'M 0 0 L ?nothing 5 N' }, {}), { width: 10, height: 10 }).length, 1, 'an equation it lacks counts as nought');
  assert.equal(enhancedFigures(geometry({ 'draw:enhanced-path': 'M 0 0 L 5' }), { width: 10, height: 10 }), null, 'a path short of numbers gives way to the preset');
});

test('arcs: a full ellipse by centre and angles, counter-clockwise and clockwise arcs by box and points, quarter ellipses', () => {
  const box = { width: 200, height: 100 };
  const [circle] = enhancedFigures(geometry({ 'draw:enhanced-path': 'U 10800 10800 10800 10800 0 23592960 Z N' }), box);
  const onEllipse = ends(circle).every(([x, y]) => Math.abs(((x - 100) / 100) ** 2 + ((y - 50) / 50) ** 2 - 1) < 1e-3);
  assert.ok(onEllipse, "Office's 16.16 angles: a whole ellipse");
  const sweep = (cmd) => {
    const [f] = enhancedFigures(geometry({ 'svg:viewBox': '0 0 100 100', 'draw:enhanced-path': `M 100 50 ${cmd} 0 0 100 100 100 50 50 0 N` }), { width: 100, height: 100 });
    return ends(f);
  };
  // From three o'clock to twelve: counter-clockwise is a quarter, clockwise three quarters.
  assert.equal(sweep('A').length, 3, 'A: the line to the start, then a quarter');
  assert.deepEqual(sweep('A').at(-1), [50, 0]);
  assert.equal(sweep('W').length, 5, 'W: the line, then three quarters round');
  const [q] = enhancedFigures(geometry({ 'svg:viewBox': '0 0 10 10', 'draw:enhanced-path': 'M 0 10 X 10 0 Y 0 10 N' }), { width: 10, height: 10 });
  assert.deepEqual(q.commands.map((c) => c.op), ['M', 'C', 'C']);
  assert.equal(q.commands[1].pts[0][1], 10, 'leaving across: its first handle level with the start');
});

test("Office's own shapes come out as their outlines: the oval on its ellipse, the smile's mouth not filled", () => {
  const root = parse(readZip(fixture('showcase.odp')).entries.find((e) => e.name === 'content.xml').data.toString('utf8'));
  const shape = (name) => first(all(root, 'draw:custom-shape').find((s) => s.attrs['draw:name'] === name), 'draw:enhanced-geometry');
  const [oval] = enhancedFigures(shape('Shape Oval'), { width: 160, height: 107 });
  assert.ok(ends(oval).every(([x, y]) => Math.abs(((x - 80) / 80) ** 2 + ((y - 53.5) / 53.5) ** 2 - 1) < 1e-3));
  const smile = enhancedFigures(shape('Shape Smile'), { width: 160, height: 107 });
  assert.deepEqual(smile.map((f) => f.fill), [true, true, true, false], 'the face and two eyes filled, the mouth a line');
  for (const name of ['Shape Star', 'Shape Heart', 'Shape Cloud', 'Shape Callout', 'Shape Chevron', 'Shape Hexagon']) assert.ok(enhancedFigures(shape(name), { width: 160, height: 107 }), name);
  const xml = figuresGeometryXml(smile, 160, 107);
  assert.match(xml, /^<a:custGeom>[\s\S]*<a:path w="1524000" h="1019175" fill="none">/);
});

test('a shape turned and moved by its transform stands where PowerPoint had it, turned as it was', () => {
  const odp = readOdf(fixture('showcase.odp'));
  const turned = odp.slides[2].shapes.find((s) => s.name === 'Shape Parallelogram');
  assert.deepEqual([Math.round(turned.x), Math.round(turned.y), Math.round(turned.w), Math.round(turned.h), turned.rotation], [1113, 373, 133, 147, 25], 'as showcase.pptx has it');
  const ods = readOdf(fixture('showcase.ods'));
  assert.equal(ods.sheets.find((s) => s.name === 'Charts').drawings.find((d) => d.name === 'Rectangle 23').rotation, 30);
});

test('opened, the shapes are custom geometry where no preset names them, and the turned one turned', () => {
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-odf-')), 'showcase.odp');
  fs.writeFileSync(file, fixture('showcase.odp'));
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const s = docs.open({ path: file });
  const out = file.replace(/\.odp$/, '.pptx');
  docs.save({ id: s.id, path: out });
  const slide = readZip(fs.readFileSync(out)).entries.find((e) => e.name === 'ppt/slides/slide3.xml').data.toString('utf8');
  assert.match(slide, /name="Shape Oval"[\s\S]*?<a:custGeom>[\s\S]*?<a:cubicBezTo>/);
  assert.match(slide, /name="Shape Smile"[\s\S]*?<a:path [^>]*fill="none"/);
  assert.match(slide, /name="Shape Parallelogram"[\s\S]*?<a:xfrm rot="1500000"><a:off x="10601325" y="3552825"\/>/);
});
