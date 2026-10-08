// Update Field over a table reaches Word's own = fields, which are complex
// fields (begin, instrText, separate, result, end), not only fldSimple ones;
// and a number picture, quoted or not, is honoured.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Document, OoxmlPackage, buildDocx } from '@rutba/ooxml';

const cell = (inner) => `<w:tc><w:p>${inner}</w:p></w:tc>`;
const text = (t) => `<w:r><w:t>${t}</w:t></w:r>`;
const complex = (instr, result) => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr.slice(0, 4)}</w:instrText></w:r><w:r><w:instrText xml:space="preserve">${instr.slice(4)} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>${result}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;

const withTotal = (field) => {
  const pkg = OoxmlPackage.read(buildDocx({ paragraphs: [{ text: 'PLACEHOLDER' }, { text: 'After' }] }));
  const table = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr>${cell(text('4'))}</w:tr><w:tr>${cell(text('1,206'))}</w:tr><w:tr>${cell(field)}</w:tr></w:tbl>`;
  pkg.write_('word/document.xml', pkg.text('word/document.xml').replace(/<w:p\b(?:(?!<\/w:p>)[\s\S])*?PLACEHOLDER[\s\S]*?<\/w:p>/, () => table));
  return Document.open(pkg.write());
};

test('a complex = field in a table is worked out again, in its own run, in its picture', () => {
  const doc = withTotal(complex(String.raw`=SUM(ABOVE) \# "#,##0.00"`, '0'));
  assert.equal(doc.updateTableFormulas(), 1);
  const xml = OoxmlPackage.read(doc.save()).text('word/document.xml');
  assert.match(xml, /fldCharType="separate"\/><\/w:r><w:r><w:rPr><w:b\/><\/w:rPr><w:t xml:space="preserve">1,210\.00<\/w:t><\/w:r><w:r><w:fldChar w:fldCharType="end"/);
  assert.match(xml, /<w:instrText xml:space="preserve"> =SUM<\/w:instrText>/, 'the instruction is left as it was');
});

test('a picture written without quotes is read, in a complex field and a simple one', () => {
  const unquoted = withTotal(complex(String.raw`=SUM(ABOVE) \# 0.0`, '0'));
  assert.equal(unquoted.updateTableFormulas(), 1);
  assert.match(OoxmlPackage.read(unquoted.save()).text('word/document.xml'), />1210\.0<\/w:t>/);
  const simple = withTotal(String.raw`<w:fldSimple w:instr=" =SUM(ABOVE) \# 0.00 "><w:r><w:t>0</w:t></w:r></w:fldSimple>`);
  assert.equal(simple.updateTableFormulas(), 1);
  assert.match(OoxmlPackage.read(simple.save()).text('word/document.xml'), />1210\.00<\/w:t>/);
});
