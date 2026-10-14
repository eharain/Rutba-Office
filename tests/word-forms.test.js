// Forms in Documents: content controls (rich text, plain text, check box,
// combo box, drop-down list, date picker, picture) and the legacy form
// fields, inserted, filled and given their properties through the view, and
// read back from the saved file as Word writes them, the w14 check box,
// w:showingPlcHdr while empty, a list's w:lastValue, a date's w:fullDate, a
// placeholder of one's own in the glossary, a legacy field's w:ffData.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx, OoxmlPackage } from '@rutba/ooxml';
import { readControl, foldControls, formatControlDate, withControlValue, CONTROL_MARK } from '@rutba/ooxml/forms';

const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const documentXml = (bytes) => OoxmlPackage.read(bytes).text('word/document.xml');
const fresh = (paragraphs = [{ text: 'Name: ' }, { text: 'Agreed: ' }]) => openDocx(buildDocx({ styles: true, paragraphs }));
const caretAt = (view, block, offset) => view.setSelection({ block, offset }, { block, offset });
const fieldsOf = (view) => view.formFields();

test('every kind of content control goes in at the caret as one character, and the file has it as Word writes it', () => {
  const view = fresh([{ text: 'Form' }]);
  const kinds = ['richText', 'plainText', 'checkBox', 'comboBox', 'dropDownList', 'date', 'picture'];
  for (const kind of kinds) {
    caretAt(view, 0, view.blocks[0].text.length);
    view.insertControl({ kind, props: { title: kind + ' field', tag: kind, ...(kind === 'comboBox' || kind === 'dropDownList' ? { items: ['Red', 'Green'] } : {}) } });
  }
  assert.equal(view.blocks[0].text, 'Form' + CONTROL_MARK.repeat(7));
  assert.equal(view.blocks[0].structural, false, 'the paragraph stays one a person can type in');
  const bytes = view.save();
  const xml = documentXml(bytes);
  assert.match(xml, new RegExp('xmlns:w14="' + W14.replace(/[./]/g, '\\$&') + '"'));
  assert.match(xml, /mc:Ignorable="[^"]*w14/, 'w14 is declared ignorable, as Word declares it');
  assert.match(xml, /<w14:checkbox><w14:checked w14:val="0"\/><w14:checkedState w14:val="2612" w14:font="MS Gothic"\/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"\/><\/w14:checkbox>/);
  assert.match(xml, /<w:text\/>/);
  assert.match(xml, /<w:comboBox><w:listItem w:displayText="Red" w:value="Red"\/><w:listItem w:displayText="Green" w:value="Green"\/><\/w:comboBox>/);
  assert.match(xml, /<w:date><w:dateFormat w:val="dd\/MM\/yyyy"\/><w:lid w:val="en-GB"\/><w:storeMappedDataAs w:val="dateTime"\/><w:calendar w:val="gregorian"\/><\/w:date>/);
  assert.match(xml, /<w:picture\/>/);
  assert.match(xml, /<w:placeholder><w:docPart w:val="DefaultPlaceholder_-1854013440"\/><\/w:placeholder><w:showingPlcHdr\/>/);
  assert.match(xml, /<w:rStyle w:val="PlaceholderText"\/><\/w:rPr><w:t xml:space="preserve">Click or tap here to enter text\.<\/w:t>/);
  assert.match(OoxmlPackage.read(bytes).text('word/styles.xml'), /w:styleId="PlaceholderText"/, 'the Placeholder Text style is added');
  const ids = [...xml.matchAll(/<w:id w:val="(-?\d+)"\/>/g)].map((m) => m[1]);
  assert.equal(new Set(ids).size, 7, 'each control its own id');

  const again = openDocx(bytes);
  const read = fieldsOf(again);
  assert.deepEqual(read.map((f) => f.kind), kinds);
  assert.deepEqual(read.map((f) => f.title), kinds.map((k) => k + ' field'));
  assert.deepEqual(read.map((f) => f.tag), kinds);
  assert.ok(read.filter((f) => f.kind !== 'checkBox' && f.kind !== 'picture').every((f) => f.showingPlaceholder));
  const picture = again.render({ pages: false }).blocks[0].runs.find((r) => r.control?.kind === 'picture').control;
  assert.match(picture.href || '', /^data:image\/png;base64,/, 'the picture control shows Word\'s placeholder picture');
  assert.equal((again.render({ pages: false }).blocks[0].images || []).length, 0, 'and the picture is the control\'s, not drawn again as the paragraph\'s');
});

test('filling each control writes what Word writes, and an empty one shows its placeholder again', () => {
  const view = fresh([{ text: 'x' }]);
  for (const kind of ['plainText', 'checkBox', 'dropDownList', 'comboBox', 'date']) {
    caretAt(view, 0, view.blocks[0].text.length);
    view.insertControl({ kind, props: { items: [{ text: 'Yes please', value: 'Y' }, { text: 'No thanks', value: 'N' }] } });
  }
  const at = (i) => ({ block: 0, offset: 1 + i });
  view.setControlValue({ ...at(0), value: 'Ayesha Khan' });
  view.setControlValue({ ...at(1), value: true });
  view.setControlValue({ ...at(2), value: 'N' });
  view.setControlValue({ ...at(3), value: 'Something else' });
  view.setControlValue({ ...at(4), value: '2026-10-14' });
  const xml = documentXml(view.save());
  const sdts = foldControls(/<w:p\b[\s\S]*?<\/w:p>/.exec(xml)[0]).controls;
  assert.equal(sdts.length, 5);
  assert.doesNotMatch(sdts[0], /showingPlcHdr|PlaceholderText/, 'filled: no placeholder flag or style');
  assert.match(sdts[0], /<w:t xml:space="preserve">Ayesha Khan<\/w:t>/);
  assert.match(sdts[1], /<w14:checked w14:val="1"\/>/);
  assert.match(sdts[1], /<w:t xml:space="preserve">☒<\/w:t>/, 'the ticked box\'s glyph');
  assert.match(sdts[2], /<w:dropDownList w:lastValue="N">/);
  assert.match(sdts[2], /<w:t xml:space="preserve">No thanks<\/w:t>/, 'a list shows the item\'s words');
  assert.match(sdts[3], /<w:t xml:space="preserve">Something else<\/w:t>/, 'a combo box takes words of its own');
  assert.match(sdts[4], /<w:date w:fullDate="2026-10-14T00:00:00Z">/);
  assert.match(sdts[4], /<w:t xml:space="preserve">14\/10\/2026<\/w:t>/);

  const again = openDocx(view.save());
  const values = fieldsOf(again).map((f) => f.value);
  assert.deepEqual(values, ['Ayesha Khan', true, 'No thanks', 'Something else', '14/10/2026']);
  assert.equal(fieldsOf(again)[4].date, '2026-10-14');

  again.setControlValue({ block: 0, offset: 1, value: '' });
  const cleared = fieldsOf(again)[0];
  assert.equal(cleared.showingPlaceholder, true);
  assert.equal(cleared.text, 'Click or tap here to enter text.');
  assert.throws(() => again.setControlValue({ block: 0, offset: 3, value: 'Maybe' }), /not one of the list's items/);
  again.undo();
  assert.equal(fieldsOf(again)[0].value, 'Ayesha Khan', 'Undo takes the filling back');
});

test('Properties: title, tag, locks, a placeholder of one\'s own kept in the glossary, list items, a date\'s format', () => {
  const view = fresh([{ text: '' }]);
  caretAt(view, 0, 0);
  view.insertControl({ kind: 'plainText' });
  view.insertControl({ kind: 'date' });
  view.insertControl({ kind: 'dropDownList', props: { items: ['A'] } });
  view.setControlProps({ block: 0, offset: 0, props: { title: 'Full name', tag: 'name', placeholder: 'Your name as on your passport', noDelete: true, multiLine: true } });
  view.setControlValue({ block: 0, offset: 1, value: '2026-03-05' });
  view.setControlProps({ block: 0, offset: 1, props: { dateFormat: 'dddd, d MMMM yyyy' } });
  view.setControlProps({ block: 0, offset: 2, props: { items: [{ text: 'Lahore', value: 'LHE' }, { text: 'Karachi', value: 'KHI' }] } });
  const bytes = view.save();
  const pkg = OoxmlPackage.read(bytes);
  const xml = pkg.text('word/document.xml');
  assert.match(xml, /<w:alias w:val="Full name"\/><w:tag w:val="name"\/><w:id w:val="-?\d+"\/><w:lock w:val="sdtLocked"\/>/);
  assert.match(xml, /<w:text w:multiLine="1"\/>/);
  assert.match(xml, /Your name as on your passport/, 'the new placeholder is shown');
  assert.match(xml, /<w:t xml:space="preserve">Thursday, 5 March 2026<\/w:t>/, 'the date written again in its new format');
  assert.match(xml, /<w:listItem w:displayText="Lahore" w:value="LHE"\/><w:listItem w:displayText="Karachi" w:value="KHI"\/>/);
  assert.ok(pkg.has('word/glossary/document.xml'), 'a glossary is made for the placeholder');
  const glossary = pkg.text('word/glossary/document.xml');
  const name = /<w:placeholder><w:docPart w:val="([0-9A-F]{32})"\/>/.exec(xml)?.[1];
  assert.ok(name, 'the control names its own glossary entry');
  assert.match(glossary, new RegExp('<w:name w:val="' + name + '"/><w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/>'));
  assert.match(pkg.text('word/_rels/document.xml.rels'), /relationships\/glossaryDocument" Target="glossary\/document\.xml"/);
  assert.match(pkg.text('[Content_Types].xml'), /glossary\+xml/);

  const again = openDocx(bytes);
  again.setControlValue({ block: 0, offset: 0, value: 'Bilal' });
  again.setControlValue({ block: 0, offset: 0, value: '' });
  assert.equal(fieldsOf(again)[0].text, 'Your name as on your passport', 'emptied, it shows the placeholder its glossary keeps');
  assert.equal(fieldsOf(again)[0].noDelete, true);
});

test('a control whose contents are locked refuses to be filled', () => {
  const view = fresh([{ text: '' }]);
  caretAt(view, 0, 0);
  view.insertControl({ kind: 'richText', props: { noEdit: true } });
  assert.throws(() => view.setControlValue({ block: 0, offset: 0, value: 'x' }), /cannot be edited/);
});

test('legacy form fields go in, are filled and come back with their w:ffData', () => {
  const view = fresh([{ text: 'Legacy: ' }]);
  caretAt(view, 0, 8);
  view.insertControl({ kind: 'formText', legacy: true });
  view.insertControl({ kind: 'formCheckBox', legacy: true });
  view.insertControl({ kind: 'formDropDown', legacy: true, props: { items: ['One', 'Two', 'Three'] } });
  let xml = documentXml(view.save());
  assert.match(xml, /<w:bookmarkStart w:id="\d+" w:name="Text1"\/><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Text1"\/><w:enabled\/><w:calcOnExit w:val="0"\/><w:textInput\/><\/w:ffData><\/w:fldChar><\/w:r><w:r><w:instrText xml:space="preserve"> FORMTEXT <\/w:instrText><\/w:r>/);
  assert.match(xml, /<w:checkBox><w:sizeAuto\/><w:default w:val="0"\/><\/w:checkBox>/);
  assert.match(xml, /<w:ddList><w:listEntry w:val="One"\/><w:listEntry w:val="Two"\/><w:listEntry w:val="Three"\/><\/w:ddList>/);
  view.setControlValue({ block: 0, offset: 8, value: 'Typed in' });
  view.setControlValue({ block: 0, offset: 9, value: true });
  view.setControlValue({ block: 0, offset: 10, value: 'Three' });
  xml = documentXml(view.save());
  assert.match(xml, /FORMTEXT <\/w:instrText><\/w:r><w:r><w:fldChar w:fldCharType="separate"\/><\/w:r><w:r><w:rPr><w:noProof\/><\/w:rPr><w:t xml:space="preserve">Typed in<\/w:t><\/w:r><w:r><w:fldChar w:fldCharType="end"\/><\/w:r>/);
  assert.match(xml, /<w:checkBox><w:sizeAuto\/><w:default w:val="0"\/><w:checked w:val="1"\/><\/w:checkBox>/);
  assert.match(xml, /<w:ddList><w:result w:val="2"\/>/);
  const again = openDocx(view.save());
  assert.equal(again.blocks[0].structural, false);
  assert.deepEqual(fieldsOf(again).map((f) => [f.kind, f.value]), [['formText', 'Typed in'], ['formCheckBox', true], ['formDropDown', 'Three']]);
});

test('typing beside a field leaves its XML exactly as it was', () => {
  const view = fresh([{ text: 'ab' }]);
  caretAt(view, 0, 1);
  view.insertControl({ kind: 'checkBox' });
  view.setControlValue({ block: 0, offset: 1, value: true });
  const before = foldControls(/<w:p\b[\s\S]*?<\/w:p>/.exec(documentXml(view.save()))[0]).controls[0];
  caretAt(view, 0, 0);
  view.insertText('X');
  caretAt(view, 0, 4);
  view.insertText('Y');
  assert.equal(view.blocks[0].text, 'Xa' + CONTROL_MARK + 'bY');
  const after = foldControls(/<w:p\b[\s\S]*?<\/w:p>/.exec(documentXml(view.save()))[0]).controls[0];
  assert.equal(after, before);
});

test('controls as Word writes them in a table cell are read and filled; a control round whole paragraphs stays as it is', () => {
  const pkg = OoxmlPackage.read(buildDocx({ styles: true, paragraphs: [{ text: 'Top' }] }));
  const cell = '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
    + '<w:p><w:r><w:t xml:space="preserve">Colour: </w:t></w:r><w:sdt><w:sdtPr><w:rPr><w:b/></w:rPr><w:alias w:val="Colour"/><w:tag w:val="colour"/><w:id w:val="-1416542330"/><w:placeholder><w:docPart w:val="DefaultPlaceholder_-1854013438"/></w:placeholder><w:showingPlcHdr/>'
    + '<w:dropDownList><w:listItem w:value="Choose an item."/><w:listItem w:displayText="Blue" w:value="Blue"/><w:listItem w:displayText="Gold" w:value="Gold"/></w:dropDownList></w:sdtPr><w:sdtContent><w:r w:rsidRPr="00A1"><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>Choose an item.</w:t></w:r></w:sdtContent></w:sdt></w:p></w:tc></w:tr></w:tbl>'
    + '<w:sdt><w:sdtPr><w:id w:val="77"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Block level</w:t></w:r></w:p></w:sdtContent></w:sdt>';
  const xml = pkg.text('word/document.xml').replace('<w:sectPr/>', cell + '<w:sectPr/>');
  pkg.write_('word/document.xml', xml);
  const view = openDocx(pkg.write());
  const cellBlock = view.blocks.findIndex((b) => b.container && b.text.startsWith('Colour'));
  assert.ok(cellBlock > 0);
  assert.equal(view.blocks[cellBlock].structural, false);
  const f = view.formFields().find((x) => x.kind === 'dropDownList');
  assert.equal(f.title, 'Colour');
  assert.deepEqual(f.items.map((i) => i.text), ['Choose an item.', 'Blue', 'Gold']);
  view.setControlValue({ block: f.block, offset: f.offset, value: 'Gold' });
  const saved = documentXml(view.save());
  assert.match(saved, /<w:dropDownList w:lastValue="Gold">/);
  assert.match(saved, /<w:r><w:rPr><w:b\/><\/w:rPr><w:t xml:space="preserve">Gold<\/w:t><\/w:r>/, 'the words in the look the control keeps for them');
  assert.match(saved, /<w:sdt><w:sdtPr><w:id w:val="77"\/><\/w:sdtPr><w:sdtContent><w:p><w:r><w:t>Block level<\/w:t><\/w:r><\/w:p><\/w:sdtContent><\/w:sdt>/);
  assert.equal(view.blocks.find((b) => b.text === 'Block level').structural, true);
});

test('Word\'s date formats, in the control\'s language', () => {
  assert.equal(formatControlDate('2026-10-04', 'dd/MM/yyyy'), '04/10/2026');
  assert.equal(formatControlDate('2026-10-04', 'M/d/yyyy', 'en-US'), '10/4/2026');
  assert.equal(formatControlDate('2026-10-04', 'dddd, MMMM d, yyyy', 'en-US'), 'Sunday, October 4, 2026');
  assert.equal(formatControlDate('2026-10-04', 'd-MMM-yy', 'en-GB'), '4-Oct-26');
  assert.equal(formatControlDate('2026-10-04', "d 'of' MMMM", 'en-GB'), '4 of October');
  assert.match(formatControlDate('2026-10-04', 'd MMMM yyyy', 'ur-PK'), /^4 .+ 2026$/);
});

test('a check box Word ticked reads ticked and a control holding a run of several pieces reads its words whole', () => {
  const xml = '<w:sdt><w:sdtPr><w:id w:val="5"/><w14:checkbox><w14:checked w14:val="1"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic" w:hint="eastAsia"/></w:rPr><w:t>☒</w:t></w:r></w:sdtContent></w:sdt>';
  const c = readControl(xml);
  assert.equal(c.kind, 'checkBox');
  assert.equal(c.checked, true);
  const off = withControlValue(xml, false);
  assert.match(off, /<w14:checked w14:val="0"\/>/);
  assert.match(off, /<w:t xml:space="preserve">☐<\/w:t>/);
  const rich = readControl('<w:sdt><w:sdtPr><w:id w:val="9"/></w:sdtPr><w:sdtContent><w:r><w:t xml:space="preserve">Two </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>runs</w:t></w:r></w:sdtContent></w:sdt>');
  assert.equal(rich.kind, 'richText');
  assert.equal(rich.value, 'Two runs');
});
