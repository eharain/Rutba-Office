// Review → Compare — the engine's half.
//
// Two documents compared make the revised one with what changed marked as
// Word marks revisions: words taken out and put in within a paragraph that
// changed, a paragraph removed or added whole with its mark, a paragraph
// that did not change left exactly as it was.
import test from 'node:test';
import assert from 'node:assert/strict';
import { buildDocx } from '@rutba/ooxml';
import { OoxmlPackage } from '@rutba/ooxml';
import { compareDocx, compareBodies } from '@rutba/ooxml/compare';
import { openDocx } from '@rutba/doc-view/backends/ooxml';

const docx = (paragraphs) => buildDocx({ styles: true, paragraphs: paragraphs.map((text, i) => ({ text, ...(i === 0 ? { style: 'Heading1' } : {}) })) });
const original = docx(['Report', 'The quick brown fox jumps over the dog.', 'Remove this one entirely.', 'Last line stays.']);
const revised = docx(['Report', 'The quick red fox leaps over the dog.', 'Last line stays.', 'Something else altogether new.']);

test('a changed paragraph has its words taken out and put in; whole paragraphs come and go with their marks', () => {
  const a = OoxmlPackage.read(original).text('word/document.xml');
  const b = OoxmlPackage.read(revised).text('word/document.xml');
  const { xml, changes } = compareBodies(a, b, { author: 'Reviewer', date: '2026-10-07T05:00:00Z' });
  assert.match(xml, /<w:del w:id="\d+" w:author="Reviewer" w:date="2026-10-07T05:00:00Z"><w:r>(<w:rPr>[\s\S]*?<\/w:rPr>)?<w:delText xml:space="preserve">brown<\/w:delText><\/w:r><\/w:del><w:ins w:id="\d+" w:author="Reviewer" w:date="2026-10-07T05:00:00Z"><w:r>(<w:rPr>[\s\S]*?<\/w:rPr>)?<w:t xml:space="preserve">red<\/w:t><\/w:r><\/w:ins>/);
  assert.match(xml, /<w:delText xml:space="preserve">jumps<\/w:delText>[\s\S]*?<w:t xml:space="preserve">leaps<\/w:t>/);
  // The paragraph gone: its runs deleted and its mark too.
  assert.match(xml, /<w:p><w:pPr>[\s\S]*?<w:rPr><w:del w:id="\d+" w:author="Reviewer"[^>]*\/>[\s\S]*?<w:del w:id[^>]*><w:r>[\s\S]*?<w:delText[^>]*>Remove this one entirely\.<\/w:delText>/);
  // The paragraph new: its runs inserted and its mark too, after the line that stayed.
  assert.match(xml, /Last line stays\.[\s\S]*<w:rPr><w:ins w:id="\d+" w:author="Reviewer"[^>]*\/>[\s\S]*<w:ins w:id[^>]*><w:r>[\s\S]*Something else altogether new\./);
  // The heading did not change: it is the revised one as it was.
  const heading = /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*Report[\s\S]*?<\/w:p>/.exec(b)[0];
  assert.ok(xml.includes(heading));
  assert.ok(changes >= 6, String(changes));
  const ids = [...xml.matchAll(/<w:(?:ins|del) w:id="(\d+)"/g)].map((m) => m[1]);
  assert.equal(new Set(ids).size, ids.length, 'every revision its own id');
});

test('the compared document opens with the revisions to accept or reject, and the same text twice has none', () => {
  const { bytes } = compareDocx(original, revised, { author: 'Reviewer' });
  const view = openDocx(bytes);
  const runs = view.blocks.flatMap((b) => b.runs || []);
  assert.ok(runs.some((r) => r.ins), 'insertions');
  assert.ok(runs.some((r) => r.del), 'deletions');
  const same = compareDocx(original, original);
  assert.equal(same.changes, 0);
});
