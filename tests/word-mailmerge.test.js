// Mailings → mail merge — the engine's half.
//
// A merge is a main document, a recipient list, and fields that say where
// each recipient's words go. These pin what the file holds (settings.xml's
// `w:mailMerge`, the fields as Word's own complex fields), what each field
// says for a record (the Address Block that leaves blank lines out, the
// Greeting Line that falls back to "Dear Sir or Madam,", an IF), and what
// Finish & Merge makes of it all: a document with a section per record, or a
// message per record handed to a sender.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import {
  parseDelimited, sourceFromRows, contactsToSource, matchFields, formatAddressBlock, formatGreeting,
  addressBlockInstr, greetingLineInstr, evaluateField, ifInstr, mergeOrder, findDuplicates, sendMergedMessages,
} from '@rutba/ooxml/mailmerge';
import { readMergeSource, writeMergeList } from '../apps/desktop/main/mailmerge-source.js';
import { createDocumentService } from '../apps/desktop/main/documents.js';

const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-mailmerge-'));
const CSV = '\uFEFFTitle,First Name,Last Name,Company Name,Address Line 1,Address Line 2,City,State,ZIP Code,Country or Region,E-mail Address\r\n'
  + 'Mr.,Joshua,Randall,,"12 High Street, Flat 2",,London,,SW1A 1AA,United Kingdom,josh@example.com\r\n'
  + 'Ms.,Cynthia,Gartner,Contoso Ltd,1 Main Road,Suite 5,Leeds,West Yorkshire,LS1 1AA,United Kingdom,cynthia@example.com\r\n'
  + 'Dr.,Amira,Haddad,,4 Rue Neuve,,Paris,,75001,France,amira@example.com\r\n';
const source = () => sourceFromRows(parseDelimited(CSV), { kind: 'csv', name: 'list.csv', path: path.join(dir, 'list.csv') });

/** A letter: an address block, a greeting, then a line with a field in it. */
function letter() {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: '' }, { text: '' }, { text: 'Thank you for your order.' }] }));
  view.startMailMerge('formLetters');
  view.attachMergeSource(source());
  view.setSelection({ block: 0, offset: 0 });
  view.insertAddressBlock({ except: 'United Kingdom' });
  view.setSelection({ block: 1, offset: 0 });
  view.insertGreetingLine();
  view.setSelection({ block: 2, offset: 0 });
  view.insertMergeField('First Name');
  view.insertText(', ');
  return view;
}

test('settings.xml carries w:mailMerge as Word writes it, reads back, and comes off again with its source link', () => {
  const view = openDocx(buildDocx({ paragraphs: ['Hello'] }));
  const engine = view.doc.doc;
  engine.setTrackRevisions(true);
  engine.setMailMerge({ type: 'formLetters', source: { kind: 'csv', path: 'C:\\Lists\\list.csv' }, activeRecord: 2, viewMergedData: true });

  const settings = engine.pkg.text('word/settings.xml');
  assert.match(settings, /<w:mailMerge><w:mainDocumentType w:val="formLetters"\/><w:linkToQuery\/><w:dataType w:val="textFile"\/><w:connectString w:val=""\/><w:query w:val="SELECT \* FROM C:\\Lists\\list.csv"\/><w:dataSource r:id="rId\d+"\/><w:viewMergedData\/><w:activeRecord w:val="2"\/><\/w:mailMerge><w:trackRevisions\/>/, 'the schema\'s order, before trackRevisions');
  const rels = engine.pkg.text('word/_rels/settings.xml.rels');
  assert.match(rels, /Type="http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships\/mailMergeSource" Target="file:\/\/\/C:\\Lists\\list.csv" TargetMode="External"/);

  const back = openDocx(view.save()).doc.doc.mailMerge();
  assert.equal(back.type, 'formLetters');
  assert.equal(back.path, 'C:\\Lists\\list.csv');
  assert.equal(back.activeRecord, 2);
  assert.equal(back.viewMergedData, true);

  // A workbook: Word's ACE provider, the sheet in the query.
  engine.setMailMerge({ type: 'mailingLabels', source: { kind: 'xlsx', path: 'C:\\Lists\\people.xlsx', sheet: 'Customers' } });
  const wb = engine.mailMerge();
  assert.equal(wb.type, 'mailingLabels');
  assert.equal(wb.dataType, 'native');
  assert.equal(wb.sheet, 'Customers');
  assert.equal(wb.path, 'C:\\Lists\\people.xlsx');
  assert.match(wb.connectString, /^Provider=Microsoft\.ACE\.OLEDB\.12\.0;.*Data Source=C:\\Lists\\people\.xlsx;/);
  assert.equal((engine.pkg.text('word/_rels/settings.xml.rels').match(/mailMergeSource/g) || []).length, 1, 'the old link replaced, not added to');

  // Normal Word Document: the element and its link gone, the rest untouched.
  engine.setMailMerge(null);
  assert.equal(engine.mailMerge(), null);
  assert.doesNotMatch(engine.pkg.text('word/_rels/settings.xml.rels'), /mailMergeSource/);
  assert.equal(engine.trackRevisions(), true);
});

test('Insert Merge Field writes Word\'s complex field — MERGEFIELD, \\* MERGEFORMAT, «name» — and the paragraph stays editable', () => {
  const view = openDocx(buildDocx({ paragraphs: [{ runs: [{ text: 'Hello ', bold: true }] }] }));
  view.attachMergeSource(source());
  view.setSelection({ block: 0, offset: 6 });
  view.insertMergeField('First Name');
  const xml = view.doc.doc.paragraph(0).xml;
  assert.equal(xml,
    '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Hello </w:t></w:r>'
    + '<w:r><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>'
    + '<w:r><w:rPr><w:b/></w:rPr><w:instrText xml:space="preserve"> MERGEFIELD First_Name \\* MERGEFORMAT </w:instrText></w:r>'
    + '<w:r><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>'
    + '<w:r><w:rPr><w:b/><w:noProof/></w:rPr><w:t xml:space="preserve">«First_Name»</w:t></w:r>'
    + '<w:r><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p>');
  const b = view.blocks[0];
  assert.equal(b.structural, false, 'a merge field does not lock its paragraph');
  const field = b.runs.find((r) => r.field);
  assert.equal(field.field.kind, 'mergefield');
  assert.equal(field.field.name, 'First_Name');
  assert.equal(field.text, '«First_Name»');

  // Typing after it rebuilds the paragraph and writes the field back as it was.
  const fieldXml = xml.slice(xml.indexOf('<w:r><w:rPr><w:b/></w:rPr><w:fldChar'), -'</w:p>'.length);
  view.setSelection({ block: 0, offset: b.text.length });
  view.insertText('!');
  assert.ok(view.doc.doc.paragraph(0).xml.includes(fieldXml), 'the field\'s runs survive the rebuild byte for byte');
  assert.equal(view.blocks[0].text, 'Hello «First_Name»!');
  assert.equal((view.doc.doc.paragraph(0).xml.match(/fldCharType="begin"/g) || []).length, 1, 'written once, not twice');
});

test('Address Block leaves out a blank company, a blank second line and a blank region, and the home country', () => {
  const s = source();
  const map = matchFields(s.fields);
  const instr = addressBlockInstr({ except: 'United Kingdom' });
  assert.match(instr, /^ ADDRESSBLOCK \\f "<<_TITLE0_ >><<_FIRST0_>><< _LAST0_>><< _SUFFIX0_>>\r<<_COMPANY_\r>><<_STREET1_\r>><<_STREET2_\r>><<_CITY_>><<, _STATE_>><< _POSTAL_>><<\r_COUNTRY_>>" \\l 1033 \\c 2 \\e "United Kingdom" \\d $/);
  assert.deepEqual(formatAddressBlock(instr, s, s.records[0], map), ['Mr. Joshua Randall', '12 High Street, Flat 2', 'London SW1A 1AA']);
  assert.deepEqual(formatAddressBlock(instr, s, s.records[1], map), ['Ms. Cynthia Gartner', 'Contoso Ltd', '1 Main Road', 'Suite 5', 'Leeds, West Yorkshire LS1 1AA']);
  assert.deepEqual(formatAddressBlock(instr, s, s.records[2], map), ['Dr. Amira Haddad', '4 Rue Neuve', 'Paris 75001', 'France'], 'a country other than the home one is kept');
  // No company, name as "Joshua Randall Jr.", never the country.
  const plain = addressBlockInstr({ name: 'firstLastSuffix', company: false, country: 'never' });
  assert.deepEqual(formatAddressBlock(plain, s, s.records[1], map), ['Cynthia Gartner', '1 Main Road', 'Suite 5', 'Leeds, West Yorkshire LS1 1AA']);
});

test('Greeting Line says Dear Mr. Randall, and falls back to Dear Sir or Madam, when the name is missing', () => {
  const s = sourceFromRows([['Title', 'First Name', 'Last Name'], ['Mr.', 'Joshua', 'Randall'], ['Dr.', '', ''], ['', 'Ali', '']]);
  const map = matchFields(s.fields);
  const instr = greetingLineInstr();
  assert.equal(instr, ' GREETINGLINE \\f "<<_BEFORE_ Dear >><<_TITLE0_ >><<_LAST0_>><<_AFTER_ ,>>" \\l 1033 \\e "Dear Sir or Madam," ');
  assert.equal(formatGreeting(instr, s, s.records[0], map), 'Dear Mr. Randall,');
  assert.equal(formatGreeting(instr, s, s.records[1], map), 'Dear Sir or Madam,', 'a title with no name is not "Dear Dr.,"');
  assert.equal(formatGreeting(instr, s, s.records[2], map), 'Dear Sir or Madam,', 'no title or surname is not "Dear ,"');
  const to = greetingLineInstr({ salutation: 'To', name: 'first', punctuation: ':', fallback: 'To Whom It May Concern:' });
  assert.equal(formatGreeting(to, s, s.records[2], map), 'To Ali:');
  assert.equal(formatGreeting(to, s, s.records[1], map), 'To Whom It May Concern:');
});

test('If…Then…Else nests a MERGEFIELD as Word does and picks its words per record', () => {
  const view = letter();
  view.setSelection({ block: 2, offset: view.blocks[2].text.length });
  view.insertMergeRule('if', { field: 'City', comparison: '=', value: 'London', then: 'See you at the shop.', otherwise: 'We will post it.' });
  const xml = view.doc.doc.paragraph(2).xml;
  assert.match(xml, /<w:instrText xml:space="preserve"> IF <\/w:instrText><\/w:r><w:r><w:fldChar w:fldCharType="begin"\/><\/w:r><w:r><w:instrText xml:space="preserve"> MERGEFIELD City <\/w:instrText>/, 'the MERGEFIELD is a field inside the IF');
  assert.equal(view.blocks[2].structural, false);
  const field = view.blocks[2].runs.find((r) => r.field?.kind === 'if');
  assert.equal(field.field.instr, ' IF { MERGEFIELD City } = "London" "See you at the shop." "We will post it." ');

  view.setMergePreview({ on: true, record: 1 });
  assert.match(view.render({ pages: false }).blocks[2].text, /See you at the shop\.$/);
  view.setMergePreview({ record: 3 });
  assert.match(view.render({ pages: false }).blocks[2].text, /We will post it\.$/);
  // Numbers compare as numbers; "is blank" is a comparison with "".
  const s = sourceFromRows([['Qty', 'Note'], ['12', ''], ['9', 'x']]);
  assert.equal(evaluateField(' IF {MERGEFIELD Qty} > "10" "bulk" "single" ', { source: s, record: s.records[0] }).text, 'bulk');
  assert.equal(evaluateField(' IF {MERGEFIELD Qty} > "10" "bulk" "single" ', { source: s, record: s.records[1] }).text, 'single');
  assert.equal(evaluateField(ifInstr({ field: 'Note', comparison: 'blank', then: 'none', otherwise: 'some' }), { source: s, record: s.records[0] }).text, 'none');
});

test('Preview Results shows record 2\'s words in place of the fields, in the field\'s own formatting', () => {
  const view = letter();
  // The first name bold, the way a person would make it stand out.
  view.setSelection({ block: 2, offset: 0 }, { block: 2, offset: '«First_Name»'.length });
  view.toggleFormat('b');
  view.setMergePreview({ on: true, record: 2 });
  const frame = view.render({ pages: false });
  assert.equal(frame.mailMerge.preview, true);
  assert.equal(frame.mailMerge.record, 2);
  assert.equal(frame.blocks[0].text, 'Ms. Cynthia Gartner\nContoso Ltd\n1 Main Road\nSuite 5\nLeeds, West Yorkshire LS1 1AA');
  assert.equal(frame.blocks[1].text, 'Dear Ms. Gartner,');
  const first = frame.blocks[2].runs[0];
  assert.equal(first.text, 'Cynthia');
  assert.equal(first.bold, true, 'the value keeps the field\'s bold');
  assert.equal(first.merged, true);
  // The file is untouched by previewing: the fields are still fields.
  assert.equal(view.blocks[2].runs[0].text, '«First_Name»');
  view.setMergePreview({ on: false });
  assert.equal(view.render({ pages: false }).blocks[2].runs[0].text, '«First_Name»');
});

test('Edit Individual Documents makes a section per record, each starting a page, with no fields left and no merge settings', () => {
  const view = letter();
  const merged = view.mergeToDocument();
  assert.equal(merged.copies, 3);
  const out = openDocx(merged.bytes);
  const texts = out.blocks.map((b) => b.text);
  assert.deepEqual(texts.filter((t) => t.startsWith('Dear')), ['Dear Mr. Randall,', 'Dear Ms. Gartner,', 'Dear Dr. Haddad,']);
  assert.equal(out.render({ pages: false }).sectionCount, 3);
  const sections = out.doc.doc.sections();
  assert.deepEqual(sections.map((s) => s.type), ['nextPage', 'nextPage', null]);
  assert.equal(out.blocks.filter((b) => b.pageBreakBefore).length, 2, 'records two and three each start a page');
  assert.equal(out.blocks[3].text.split('\n')[0], 'Ms. Cynthia Gartner');
  const body = out.doc.doc.xml;
  assert.doesNotMatch(body, /fldChar|MERGEFIELD|ADDRESSBLOCK/, 'merged words are words');
  assert.match(body, /<w:br\/>/, 'the address lines are line breaks');
  assert.equal(out.doc.doc.mailMerge(), null);
  // Current record, and From–To.
  view.setMergePreview({ on: true, record: 3 });
  assert.equal(openDocx(view.mergeToDocument({ range: 'current' }).bytes).blocks[1].text, 'Dear Dr. Haddad,');
  assert.equal(view.mergeToDocument({ range: { from: 2, to: 3 } }).copies, 2);
});

test('a .csv with a byte-order mark, quoted commas and CRLF, a .tsv and a workbook are recipient lists', () => {
  const csvPath = path.join(dir, 'people.csv');
  fs.writeFileSync(csvPath, CSV);
  const csv = readMergeSource(csvPath);
  assert.equal(csv.fields[0], 'Title', 'no byte-order mark in the first field\'s name');
  assert.equal(csv.records.length, 3);
  assert.equal(csv.records[0][4], '12 High Street, Flat 2', 'a quoted comma stays in its field');
  assert.equal(csv.kind, 'csv');

  const tsvPath = path.join(dir, 'people.tsv');
  fs.writeFileSync(tsvPath, 'Name\tCity\nAda, Countess\tLondon\n');
  assert.deepEqual(readMergeSource(tsvPath).records, [['Ada, Countess', 'London']]);

  const xlsxPath = path.join(dir, 'people.xlsx');
  fs.writeFileSync(xlsxPath, buildXlsx({ sheets: [
    { name: 'Notes', rows: [['nothing here']] },
    { name: 'Customers', rows: [['First Name', 'Last Name', 'Orders', 'Double'], ['Joshua', 'Randall', 3, '=C2*2'], ['Cynthia', 'Gartner', 12, '=C3*2']] },
  ] }));
  const first = readMergeSource(xlsxPath);
  assert.equal(first.sheet, 'Notes', 'the first sheet unless another is chosen');
  assert.deepEqual(first.sheets, ['Notes', 'Customers']);
  const wb = readMergeSource(xlsxPath, { sheet: 'Customers' });
  assert.deepEqual(wb.fields, ['First Name', 'Last Name', 'Orders', 'Double']);
  assert.deepEqual(wb.records[1], ['Cynthia', 'Gartner', '12', '24'], 'numbers and formulas as the grid shows them');

  // Type a New List saves a .csv Excel reads — and reads back the same.
  const newList = path.join(dir, 'typed.csv');
  writeMergeList(newList, ['First Name', 'City'], [['Zoë', 'Lahore, Punjab']]);
  assert.equal(fs.readFileSync(newList)[0], 0xef, 'a UTF-8 byte-order mark first');
  assert.deepEqual(readMergeSource(newList).records, [['Zoë', 'Lahore, Punjab']]);
});

test('Choose from Contacts makes a list in Word\'s own column names, which Address Block matches unasked', () => {
  const s = contactsToSource([
    { name: { full: 'Joshua Randall', given: 'Joshua', family: 'Randall', prefix: 'Mr.' }, org: 'Contoso', emails: [{ value: 'josh@example.com' }], addresses: [{ street: '12 High St\nFlat 2', city: 'London', region: '', postcode: 'SW1A 1AA', country: 'United Kingdom' }], phones: [{ type: 'work', value: '020 7946 0000' }] },
    { name: { full: 'Ali' }, emails: [] },
  ]);
  assert.equal(s.kind, 'contacts');
  assert.deepEqual(s.records[0].slice(0, 10), ['Mr.', 'Joshua', 'Randall', 'Contoso', '12 High St', 'Flat 2', 'London', '', 'SW1A 1AA', 'United Kingdom']);
  assert.equal(s.records[0][12], 'josh@example.com');
  const map = matchFields(s.fields);
  assert.deepEqual(formatAddressBlock(addressBlockInstr({ except: 'United Kingdom' }), s, s.records[0], map), ['Mr. Joshua Randall', 'Contoso', '12 High St', 'Flat 2', 'London SW1A 1AA']);
  assert.equal(s.records[1][1], 'Ali');
});

test('Edit Recipient List: unticked records are left out, the sort orders the merge, duplicates are found', () => {
  const view = letter();
  view.setMergeRecipients({ excluded: [0], sort: { field: 'Last Name', descending: false } });
  const frame = view.render({ pages: false });
  assert.equal(frame.mailMerge.included, 2);
  assert.deepEqual(view.mergeOrder(), [1, 2], 'Gartner before Haddad; Randall left out');
  view.setMergeRecipients({ excluded: [], sort: { field: 'Last Name', descending: true } });
  assert.deepEqual(view.mergeOrder(), [0, 2, 1]);
  view.setMergePreview({ on: true, record: 1 });
  assert.equal(view.render({ pages: false }).blocks[1].text, 'Dear Mr. Randall,', 'record 1 is the first in the sorted list');
  const s = sourceFromRows([['A', 'B'], ['x', 'y'], ['X ', 'y'], ['z', 'y']]);
  assert.deepEqual(findDuplicates(s), [[0, 1]]);
  assert.deepEqual(mergeOrder(s, { excluded: [1] }), [0, 2]);
  assert.equal(view.findMergeRecipient('amira'), 2, 'Find Recipient goes to Haddad, second in this order');
});

test('Send E-mail Messages hands a message per record to the sender: its address, the subject, the merged letter', async () => {
  const view = letter();
  view.setMergeEmail({ toField: 'E-mail Address', subject: 'Your order' });
  const messages = view.mergeMessages({ format: 'html' });
  assert.equal(messages.length, 3);
  assert.equal(messages[1].to, 'cynthia@example.com');
  assert.equal(messages[1].subject, 'Your order');
  assert.equal(messages[1].text, 'Ms. Cynthia Gartner\nContoso Ltd\n1 Main Road\nSuite 5\nLeeds, West Yorkshire LS1 1AA\nDear Ms. Gartner,\nCynthia, Thank you for your order.\n');
  assert.match(messages[1].html, /<p style="margin:0 0 10px">Dear Ms\. Gartner,<\/p>/);

  const sent = [];
  const progress = [];
  const fake = async (m) => { if (m.to === 'amira@example.com') throw new Error('mailbox full'); sent.push(m); };
  const out = await sendMergedMessages(messages, fake, (done, total) => progress.push(`${done}/${total}`));
  assert.equal(out.sent, 2);
  assert.deepEqual(out.failed, [{ record: 3, to: 'amira@example.com', error: 'mailbox full' }], 'one failure is named and the others still go');
  assert.deepEqual(progress, ['1/3', '2/3', '3/3']);
  assert.deepEqual(sent.map((m) => m.to), ['josh@example.com', 'cynthia@example.com']);
  // The To column and subject are kept in the file, as Word keeps them.
  const saved = openDocx(view.save()).doc.doc.mailMerge();
  assert.equal(saved.addressField, 'E-mail Address');
  assert.equal(saved.subject, 'Your order');
});

test('Next Record fills a sheet of labels: each copy takes the records its «Next Record» fields move on to; Skip Record If leaves one out', () => {
  const view = openDocx(buildDocx({ paragraphs: [{ text: '' }, { text: '' }] }));
  view.startMailMerge('mailingLabels');
  view.attachMergeSource(source());
  view.setSelection({ block: 0, offset: 0 });
  view.insertMergeField('Last Name');
  view.setSelection({ block: 1, offset: 0 });
  view.insertMergeRule('next');
  view.insertMergeField('Last Name');
  view.setMergePreview({ on: true, record: 1 });
  assert.deepEqual(view.render({ pages: false }).blocks.map((b) => b.text), ['Randall', 'Gartner']);
  const merged = openDocx(view.mergeToDocument().bytes);
  assert.deepEqual(merged.blocks.map((b) => b.text), ['Randall', 'Gartner', 'Haddad', ''], 'two labels a copy; the last copy runs out');

  const skipping = letter();
  skipping.setSelection({ block: 0, offset: 0 });
  skipping.insertMergeRule('skipif', { field: 'Country or Region', comparison: '=', value: 'France' });
  const kept = openDocx(skipping.mergeToDocument().bytes).blocks.map((b) => b.text).filter((t) => t.startsWith('Dear'));
  assert.deepEqual(kept, ['Dear Mr. Randall,', 'Dear Ms. Gartner,']);
});

test('a merge letter saved and opened again finds its list where it left it, on the record it was showing', () => {
  const service = createDocumentService({ holdBlob: () => ({ url: 'blob://held' }) });
  const listPath = path.join(dir, 'reopen.csv');
  fs.writeFileSync(listPath, CSV);
  const docPath = path.join(dir, 'letter.docx');
  fs.writeFileSync(docPath, buildDocx({ paragraphs: [{ text: 'Dear ' }] }));

  const first = service.open({ path: docPath, kind: 'doc' });
  service.apply({ id: first.id, ops: [{ op: 'startMailMerge', type: 'formLetters' }] });
  const attached = service.mailMerge({ id: first.id, action: 'attach', path: listPath });
  assert.equal(attached.source.count, 3);
  service.apply({ id: first.id, ops: [{ op: 'setSelection', anchor: { block: 0, offset: 5 } }, { op: 'insertMergeField', name: 'First Name' }, { op: 'mergePreview', on: true, record: 2 }] });
  service.save({ id: first.id });
  service.close({ id: first.id });

  const again = service.open({ path: docPath, kind: 'doc' });
  const mm = again.model.mailMerge;
  assert.equal(mm.type, 'formLetters');
  assert.equal(mm.source?.name, 'reopen.csv');
  assert.equal(mm.source.count, 3);
  assert.equal(mm.record, 2);
  assert.equal(mm.preview, true);
  assert.equal(again.model.blocks[0].text, 'Dear Cynthia');
  assert.equal(again.dirty, false, 'reopening is not an edit');

  // The list moved: the letter still opens, and says where it looked.
  service.close({ id: again.id });
  fs.renameSync(listPath, listPath + '.moved');
  const lost = service.open({ path: docPath, kind: 'doc' });
  assert.equal(lost.model.mailMerge.source, null);
  assert.deepEqual(lost.model.mailMerge.pending, { path: listPath, sheet: null });
  service.close({ id: lost.id });
});

test('Finish & Merge through the service makes an unsaved Letters1 a window can take over', () => {
  const service = createDocumentService({ holdBlob: () => ({ url: 'blob://held' }) });
  const listPath = path.join(dir, 'finish.csv');
  fs.writeFileSync(listPath, CSV);
  const opened = service.new({ kind: 'word' });
  service.mailMerge({ id: opened.id, action: 'attach', path: listPath });
  service.apply({ id: opened.id, ops: [{ op: 'setSelection', anchor: { block: 0, offset: 0 } }, { op: 'insertGreetingLine', spec: {} }] });
  assert.deepEqual(service.mailMerge({ id: opened.id, action: 'errors' }), []);
  const made = service.mailMerge({ id: opened.id, action: 'finish', range: 'all' });
  assert.equal(made.name, 'Letters1');
  assert.equal(made.copies, 3);
  const adopted = service.adopt({ id: made.id }, { id: 42 });
  assert.equal(adopted.dirty, true, 'a merged document is unsaved work');
  assert.equal(adopted.model.sectionCount, 3);
  assert.deepEqual(adopted.model.blocks.map((b) => b.text).filter(Boolean), ['Dear Mr. Randall,', 'Dear Ms. Gartner,', 'Dear Dr. Haddad,']);
  service.close({ id: made.id });
  service.close({ id: opened.id });
});
