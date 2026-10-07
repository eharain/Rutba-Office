// Presentations: a table drawn in its table style, and the preset shapes
// drawn as their outlines — read from tests/fixtures/rich/showcase.pptx,
// which PowerPoint made, and from tables the deck writes itself.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { Deck, buildPptx, renderSlide } from '@rutba/presentation';
import { presetPath, PRESET_PATHS } from '@rutba/drawing/presets';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const showcase = () => Deck.open(fs.readFileSync(path.join(HERE, 'fixtures', 'rich', 'showcase.pptx')));
const tableOf = (deck, i) => deck.slide(i).shapes.find((s) => s.kind === 'table').table;

test('a table takes its style: the header in the accent with bold light words and a rule under it, the rows banded, its own fill over the style\'s', () => {
  const table = tableOf(showcase(), 3);
  const [header, north, south] = table.rows;
  assert.equal(table.styleId, '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}');
  assert.equal(header.cells[0].fill.color, '#156082');
  assert.deepEqual([header.cells[0].text.paragraphs[0].runs[0].bold, header.cells[0].text.paragraphs[0].runs[0].color], [true, '#ffffff']);
  assert.deepEqual([header.cells[0].edges.bottom.width, header.cells[0].edges.bottom.color], [3, '#ffffff'], 'the thicker rule under the header');
  assert.deepEqual([north.cells[0].edges.left.width, north.cells[0].edges.left.color], [1, '#ffffff'], 'light lines between the cells');
  assert.notEqual(north.cells[1].fill.color, south.cells[1].fill.color, 'banded');
  assert.equal(north.cells[1].text.paragraphs[0].runs[0].bold, undefined, 'the body\'s words plain');
  assert.equal(table.rows[4].cells[3].fill.color, '#ffe699', 'the cell picked out in its own fill');
});

test('a table naming the deck\'s default style is drawn in it though the file does not write it out', () => {
  const deck = Deck.open(buildPptx({ slides: [{ layout: 'blank' }] }));
  deck.addTable(0, { rows: 3, cols: 2, x: 40, y: 40, w: 400, h: 120 });
  const table = tableOf(deck, 0);
  assert.ok(table.rows[0].cells[0].fill?.color, 'the header filled');
  assert.notEqual(table.rows[0].cells[0].fill.color, table.rows[1].cells[0].fill.color);
  const svg = renderSlide(deck.slide(0), { width: 960 });
  assert.match(svg, /<line [^>]*stroke="#ffffff"/, 'its light lines drawn');
});

test('a table with no style draws only the borders and fills its cells give', () => {
  const pkg = buildPptx({ slides: [{ layout: 'blank' }] });
  const deck = Deck.open(pkg);
  deck.addTable(0, { rows: 2, cols: 2, x: 40, y: 40, w: 400, h: 120 });
  const part = 'ppt/slides/slide1.xml';
  const red = '<a:tcPr><a:lnB w="25400"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:lnB><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:tcPr>';
  const xml = deck.pkg.text(part).replace(/<a:tblPr\b[\s\S]*?<\/a:tblPr>/, '<a:tblPr/>').replace(/<a:tcPr\/>|<a:tcPr>[\s\S]*?<\/a:tcPr>/, () => red);
  deck.pkg.write_(part, xml);
  const table = tableOf(Deck.open(deck.pkg.write()), 0);
  assert.equal(table.styleId, null);
  assert.equal(table.rows[0].cells[0].fill.color.toLowerCase(), '#ff0000');
  assert.deepEqual([table.rows[0].cells[0].edges.bottom.width, table.rows[0].cells[0].edges.bottom.color.toLowerCase()], [2, '#00ff00']);
  assert.equal(table.rows[1].cells[1].fill, null, 'no fill where none is given');
  assert.equal(table.rows[1].cells[1].edges.top, null, 'and no border');
});

test('presets drawn as their outlines: a heart from its cleft, a can with its top, a smiley\'s face, eyes and mouth', () => {
  const heart = presetPath('heart', 0, 0, 100, 100);
  assert.match(heart, /^M50 25 C/, 'from the cleft a quarter down');
  assert.equal((heart.match(/C/g) || []).length, 2);
  assert.equal((presetPath('can', 0, 0, 100, 100).match(/M/g) || []).length, 2, 'the body, and the top over it');
  assert.equal((presetPath('smileyFace', 0, 0, 100, 100).match(/M/g) || []).length, 4);
  assert.match(presetPath('wedgeRectCallout', 0, 0, 120, 80), /L\d/, 'its tail');
  assert.equal(presetPath('flowChartDecision', 0, 0, 10, 10), presetPath('diamond', 0, 0, 10, 10), 'a decision is a diamond');
  assert.equal(presetPath('noSuchShape', 0, 0, 10, 10), null);
  assert.ok(PRESET_PATHS.includes('cloud') && PRESET_PATHS.includes('upArrow'));
  // On a slide: the heart a path, not a box.
  const deck = showcase();
  const svg = renderSlide(deck.slide(2), { width: 960, tagShapes: true });
  assert.ok((svg.match(/<path d="M/g) || []).length >= 6, 'the presets once drawn as boxes');
});
