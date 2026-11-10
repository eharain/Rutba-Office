// Files written to do harm to the readers of plain formats. Each of these was
// found by tools/fuzz-grammar.js: an overflowed stack, a crash, or a pattern
// that read a long line once for every character on it, which a person meets
// as a window that stops. A reader answers quickly or refuses in a sentence.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readRtfDocument } from '../packages/office-formats/src/rtf-read.js';
import { parseHtml, readHtmlDocument } from '../packages/office-formats/src/html-read.js';
import { parseMarkdown, inlineRuns, serializeMarkdown, runsToMarkdown } from '../packages/office-formats/src/markdown.js';
import { parse as parseXml } from '../packages/office-formats/src/xml.js';
import { exceptionsOf, occurrences } from '../packages/calendar/src/recur.js';
import { readOdf } from '../packages/office-formats/src/odf.js';
import { writeZip, ZipEntry } from '../packages/ooxml/src/zip.js';

/** Runs it and says how long it took, against a limit generous enough for a loaded machine. */
function within(ms, fn) {
  const t = performance.now();
  const out = fn();
  assert.ok(performance.now() - t < ms, `took ${(performance.now() - t).toFixed(0)} ms`);
  return out;
}
const sentence = /^This (file|page|document) [a-z ]+\.$/;

test('RTF: groups nested past reason are refused in a sentence', () => {
  assert.throws(() => within(3000, () => readRtfDocument(`{\\rtf1 ${'{'.repeat(200000)}x`)), (e) => e.name === 'Error' && sentence.test(e.message));
});

test('RTF: a Unicode value out of range does not stop the file', () => {
  const doc = readRtfDocument('{\\rtf1\\uc0 a\\u-99999999 b\\u5000000 c\\par}');
  assert.equal(doc.blocks[0].runs.map((r) => r.text).join(''), 'a\ufffdb\ufffdc');
});

test('RTF: a Unicode value in range still reads', () => {
  const doc = readRtfDocument('{\\rtf1 \\uc0\\u8364 x\\par}');
  assert.equal(doc.blocks[0].runs.map((r) => r.text).join(''), '€x');
});

test('RTF: a row of a hundred thousand cells does not take minutes', () => {
  within(4000, () => readRtfDocument(`{\\rtf1 \\trowd${Array.from({ length: 100000 }, (_, i) => `\\cellx${i + 1}`).join('')}\\row}`));
});

test('RTF: a table claiming millions of cells is refused in a sentence', () => {
  const edges = Array.from({ length: 3000 }, (_, i) => `\\cellx${i + 1}`).join('');
  const rows = '\\trowd\\cellx1\\intbl a\\cell\\row '.repeat(1000);
  assert.throws(() => within(6000, () => readRtfDocument(`{\\rtf1 \\trowd${edges}\\intbl a\\cell\\row ${rows}\\par x\\par}`)), (e) => sentence.test(e.message));
});

test('RTF: a long run of spaces inside a paragraph is read once', () => {
  within(3000, () => readRtfDocument(`{\\rtf1 ${' '.repeat(400000)}x\\par}`));
});

test('HTML: comments that never end are read once', () => {
  within(3000, () => parseHtml('<!--'.repeat(300000)));
});

test('HTML: tags that never close are read once', () => {
  within(3000, () => parseHtml(`${'<a b '.repeat(200000)}/x`));
});

test('HTML: elements nested past reason do not overflow the stack', () => {
  const doc = within(4000, () => readHtmlDocument(`${'<div>'.repeat(300000)}words`));
  assert.ok(doc.blocks.length >= 1);
});

test('HTML: a very long page of paragraphs inside one element does not overflow the stack', () => {
  const doc = readHtmlDocument(`<div>${'<p>x</p>'.repeat(200000)}</div>`);
  assert.equal(doc.blocks.length, 200000);
});

test('HTML: spans claiming the whole table are refused or finish, quickly', () => {
  const html = `<table>${'<tr><td colspan=99999999 rowspan=99999999>x'.repeat(5000)}</table>`;
  try {
    within(6000, () => readHtmlDocument(html));
  } catch (err) {
    assert.match(err.message, sentence);
  }
});

test('HTML: a page people write still parses as it did', () => {
  const doc = readHtmlDocument(`<!doctype html><title>T</title><!-- gone --><p align="center" class=x>One <b>two</b><br/>three</p>
    <a href='https://example.com/a?b=1&amp;c=2' title="x > y">link</a> 1 < 2 and <![CDATA[kept out]]>
    <ul><li>a<li>b</ul><table border=1><tr><td colspan="2">c<td>d</table>`);
  assert.equal(doc.title, 'T');
  assert.equal(doc.blocks[0].align, 'center');
  assert.deepEqual(doc.blocks[0].runs.map((r) => r.text ?? 'BR'), ['One ', 'two', 'BR', 'three']);
  const link = doc.blocks[1].runs.find((r) => r.link);
  assert.equal(link.link, 'https://example.com/a?b=1&c=2');
  assert.match(doc.blocks[1].runs.map((r) => r.text).join(''), /1 < 2 and/);
  assert.doesNotMatch(doc.blocks[1].runs.map((r) => r.text).join(''), /kept out/);
  const table = doc.blocks.find((b) => b.type === 'table');
  assert.equal(table.rows[0][0].colspan, 2);
  assert.equal(table.rows[0][1].covered, true);
});

for (const [what, text] of [
  ['stars', '*'.repeat(300000)],
  ['underscores', '_'.repeat(300000)],
  ['open brackets', '['.repeat(300000)],
  ['backticks', '`'.repeat(300000)],
  ['footnote openings', '[^'.repeat(150000)],
  ['angle brackets', '<a'.repeat(150000)],
  ['nested links', `${'['.repeat(100000)}x${'](u)'.repeat(100000)}`],
  ['stars around backticks', `*${'`'.repeat(300000)}`],
]) {
  test(`Markdown: a line of ${what} is read in time`, () => {
    within(5000, () => inlineRuns(text));
  });
}

for (const [what, text] of [
  ['one heading padded with spaces', `# a${' '.repeat(100000)}b`],
  ['one heading of hashes and a word', `# ${'#'.repeat(400000)}x`],
  ['a table rule of dashes that fails at the end', `a|b\n${'-'.repeat(400000)}x`],
  ['a paragraph padded with spaces', `x${' '.repeat(400000)}y\nnext`],
]) {
  test(`Markdown: ${what} is read in time`, () => {
    within(5000, () => parseMarkdown(text));
  });
}

test('Markdown: what people write still reads as it did', () => {
  const runs = inlineRuns('a *b* **c** `d` [e](u "t") ![f](g) ~~h~~ [^n] <http://x.y> [r][k] \\*i\\*', new Map([['k', { url: 'v', title: null, label: 'k' }]]));
  const by = (key) => runs.find((r) => r[key]);
  assert.equal(by('italic').text, 'b');
  assert.equal(by('bold').text, 'c');
  assert.equal(by('code').text, 'd');
  assert.equal(by('strike').text, 'h');
  assert.equal(by('image').image, 'g');
  assert.equal(by('footnote').footnote, 'n');
  assert.deepEqual([runs.find((r) => r.text === 'e').link, runs.find((r) => r.text === 'e').title], ['u', 't']);
  assert.equal(runs.find((r) => r.text === 'r').link, 'v');
  assert.match(runs.at(-1).text, /\*i\*/);
  assert.equal(runs.find((r) => r.angle).link, 'http://x.y');
  const blocks = parseMarkdown('### Three ###\n\n#nope\n\n####### seven\n\n| a | b |\n|:--|--:|\n| 1 | 2 |\n').blocks;
  assert.deepEqual(blocks.slice(0, 3).map((b) => [b.type, b.level ?? null, b.text]), [['heading', 3, 'Three'], ['paragraph', null, '#nope'], ['paragraph', null, '####### seven']]);
  assert.deepEqual(blocks[3].align, ['left', 'right']);
  const hard = parseMarkdown('one  \ntwo').blocks[0];
  assert.equal(hard.text, 'one\ntwo');
});

test('Markdown: writing a very long code block back does not overflow the stack', () => {
  const out = serializeMarkdown([{ type: 'code', language: '', text: 'x\n'.repeat(300000) }]);
  assert.ok(out.startsWith('```'));
});

test('Markdown: writing back a long line of marks is linear and still escapes', () => {
  within(4000, () => runsToMarkdown([{ text: '[*_'.repeat(150000) }]));
  assert.equal(runsToMarkdown([{ text: '[a](b) *c* d' }]), '\\[a\\](b) \\*c* d');
});

test('XML: elements nested past reason are refused in a sentence', () => {
  assert.throws(() => parseXml('<a>'.repeat(100000)), (e) => e.name === 'Error' && sentence.test(e.message));
  assert.equal(parseXml(`${'<a>'.repeat(400)}x${'</a>'.repeat(400)}`).children.length, 1);
});

const ODS_NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"';
function ods(rows) {
  const entry = (name, text) => {
    const e = new ZipEntry({ name, method: 0, crc: 0, compressedSize: 0, uncompressedSize: 0, compressed: Buffer.alloc(0), dosTime: 0, dosDate: 0x2821 });
    e.data = Buffer.from(text, 'utf8');
    return e;
  };
  return Buffer.from(writeZip([entry('mimetype', 'application/vnd.oasis.opendocument.spreadsheet'), entry('content.xml', `<office:document-content ${ODS_NS}><office:body><office:spreadsheet><table:table table:name="S">${rows}</table:table></office:spreadsheet></office:body></office:document-content>`)]));
}

test('ODS: a few bytes that ask for millions of filled cells are refused, quickly', () => {
  const row = '<table:table-row table:number-rows-repeated="3000"><table:table-cell table:number-columns-repeated="1000" office:value-type="float" office:value="1"/></table:table-row>';
  assert.throws(() => within(2500, () => readOdf(ods(row))), (e) => /more cells/.test(e.message));
});

test('ODS: repeats as a real sheet writes them still read', () => {
  const row = '<table:table-row table:number-rows-repeated="3"><table:table-cell table:number-columns-repeated="2" office:value-type="float" office:value="7"/></table:table-row><table:table-row table:number-rows-repeated="1000000"><table:table-cell table:number-columns-repeated="1024"/></table:table-row>';
  const out = readOdf(ods(row));
  assert.deepEqual(out.sheets[0].rows.map((r) => r.map((c) => c.value)), [['7', '7'], ['7', '7'], ['7', '7']]);
});

test('Calendar: a series owns its exceptions, once, and a stray event gets none', () => {
  const at = (n) => ({ at: n, date: true });
  const parent = { uid: 'a', start: at(1) };
  const stray = { uid: 'a', start: at(2) };
  const moved = { uid: 'a', start: at(3), recurrenceId: at(1) };
  const other = { uid: 'b', start: at(4) };
  const of = exceptionsOf([parent, stray, moved, other]);
  assert.deepEqual(of(parent), [moved]);
  assert.deepEqual(of(stray), []);
  assert.deepEqual(of(other), []);
});

test('Calendar: a limit holds, and an exception moved outside the rule still shows', () => {
  const start = { at: 1700000000000, date: '2023-11-14', time: '00:00', utc: true };
  const moved = { uid: 'a', start: { ...start, at: start.at + 5000 }, recurrenceId: { at: start.at + 9000 } };
  const shown = occurrences({ uid: 'a', start }, { from: 0, to: 1800000000000, siblings: [moved], limit: 10 });
  assert.deepEqual(shown.map((o) => o.event === moved), [false, true]);
  const many = Array.from({ length: 50 }, (_, i) => ({ uid: 'a', start: { ...start, at: start.at + i * 7000 }, recurrenceId: { at: start.at + i * 7000 + 500 } }));
  assert.equal(occurrences({ uid: 'a', start }, { from: 0, to: 1800000000000, siblings: many, limit: 10 }).length, 10);
});

test('Calendar: sixty thousand events of one UID are expanded in time', () => {
  const events = [];
  for (let i = 0; i < 30000; i++) {
    events.push({ uid: 'a', start: { at: 1700000000000 + i * 1000, date: '2023-11-14', time: '00:00', utc: true } });
    events.push({ uid: 'a', start: { at: 1700000000000 + i * 1000, date: '2023-11-14', time: '00:00', utc: true }, recurrenceId: { at: 1700000000000 + i * 1000 } });
  }
  within(6000, () => {
    const of = exceptionsOf(events);
    for (const e of events) if (!e.recurrenceId) occurrences(e, { from: 0, to: 1800000000000, siblings: of(e), limit: 500 });
  });
});
