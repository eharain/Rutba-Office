/**
 * Review → comments: Excel 365's threaded comments, in the engine.
 *
 * A thread lives in `xl/threadedComments/threadedCommentN.xml` (related from
 * the sheet), its authors in `xl/persons/person.xml` (related from the
 * workbook), and keeps a classic shadow — a comment authored `tc={id}`
 * reading "[Threaded comment] … Comment: … Reply: …" with its VML box — so
 * an Excel without threaded comments shows something. These read what the
 * engine writes, what it reads back, and a file written the way Excel
 * writes one.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml';
import { SheetView } from '@rutba/sheet-view';

const open = () => new SheetView(buildXlsx({
  sheets: [
    { name: 'Plan', rows: [['Item', 'Cost'], ['Paint', 40], ['Brushes', 12]] },
    { name: 'Notes', rows: [['x']] },
  ],
}));
const text = (view, name) => view.pkg.text(name);
const sheetRels = (view) => text(view, 'xl/worksheets/_rels/sheet1.xml.rels');
const legacy = (view) => {
  const rel = /Target="([^"]*comments\d*\.xml)"/.exec(sheetRels(view))?.[1];
  return rel ? text(view, 'xl/' + rel.replace('../', '')) : '';
};
const WHEN = new Date('2026-09-25T09:30:12.345Z');

test('New Comment writes the threaded part, the person, and the classic shadow with its box', () => {
  const view = open();
  view.select(1, 1);
  const id = view.addComment({ text: 'Is this the trade price?', author: 'Ejaz', date: WHEN });
  assert.match(id, /^\{[0-9A-F-]{36}\}$/);

  const part = text(view, 'xl/threadedComments/threadedComment1.xml');
  assert.match(part, /<ThreadedComments xmlns="http:\/\/schemas\.microsoft\.com\/office\/spreadsheetml\/2018\/threadedcomments"/);
  assert.match(part, new RegExp('<threadedComment ref="B2" dT="2026-09-25T09:30:12\\.34" personId="\\{[0-9A-F-]{36}\\}" id="' + id.replace(/[{}]/g, '\\$&') + '"><text>Is this the trade price\\?</text>'));
  assert.match(sheetRels(view), /Type="http:\/\/schemas\.microsoft\.com\/office\/2017\/10\/relationships\/threadedComment" Target="\.\.\/threadedComments\/threadedComment1\.xml"/);

  const persons = text(view, 'xl/persons/person.xml');
  assert.match(persons, /<personList xmlns="http:\/\/schemas\.microsoft\.com\/office\/spreadsheetml\/2018\/threadedcomments"/);
  assert.match(persons, /<person displayName="Ejaz" id="\{[0-9A-F-]{36}\}" userId="Ejaz" providerId="None"\/>/);
  assert.match(text(view, 'xl/_rels/workbook.xml.rels'), /relationships\/person" Target="persons\/person\.xml"/);

  const types = text(view, '[Content_Types].xml');
  assert.match(types, /PartName="\/xl\/threadedComments\/threadedComment1\.xml" ContentType="application\/vnd\.ms-excel\.threadedcomments\+xml"/);
  assert.match(types, /PartName="\/xl\/persons\/person\.xml" ContentType="application\/vnd\.ms-excel\.person\+xml"/);

  const shadow = legacy(view);
  assert.match(shadow, new RegExp('<author>tc=' + id.replace(/[{}]/g, '\\$&') + '</author>'));
  assert.match(shadow, /\[Threaded comment\]/);
  assert.match(shadow, /Comment:\n {4}Is this the trade price\?/);
  const vml = Object.keys(view.workbook.snapshotParts(view._noteParts())).find((n) => n.endsWith('.vml'));
  assert.ok(vml, 'a VML box for the shadow');
  assert.match(text(view, vml), /ObjectType="Note"[\s\S]*<x:Row>1<\/x:Row>\s*<x:Column>1<\/x:Column>/);

  const frame = view.render();
  const cell = frame.cells.find((c) => c.ref === 'B2');
  assert.deepEqual({ ...cell.thread, id: null }, { id: null, done: false, replies: 0, author: 'Ejaz', text: 'Is this the trade price?' });
  assert.equal(cell.note, null, 'the shadow is not drawn as a note');
  assert.equal(frame.threads, 1);
  assert.equal(frame.notes, 0);
  assert.equal(frame.thread.comments[0].date, '2026-09-25T09:30:12.34');
});

test('a reply joins the thread; edit, resolve and reopen are written where Excel keeps them', () => {
  const view = open();
  view.select(1, 1);
  const top = view.addComment({ text: 'Is this the trade price?', author: 'Ejaz', date: WHEN });
  const reply = view.addComment({ text: 'Yes, before VAT.', author: 'Sana', date: WHEN });
  let part = text(view, 'xl/threadedComments/threadedComment1.xml');
  assert.match(part, new RegExp('id="' + reply.replace(/[{}]/g, '\\$&') + '" parentId="' + top.replace(/[{}]/g, '\\$&') + '"><text>Yes, before VAT\\.</text>'));
  assert.equal(view.workbook.persons().length, 2);
  assert.match(legacy(view), /Comment:\n {4}Is this the trade price\?\nReply:\n {4}Yes, before VAT\./);
  const thread = view.render().thread;
  assert.deepEqual(thread.comments.map((c) => [c.author, c.text]), [['Ejaz', 'Is this the trade price?'], ['Sana', 'Yes, before VAT.']]);

  view.editComment({ id: reply, text: 'Yes — before VAT.' });
  assert.match(text(view, 'xl/threadedComments/threadedComment1.xml'), /<text>Yes — before VAT\.<\/text>/);
  assert.match(legacy(view), /Reply:\n {4}Yes — before VAT\./);

  view.resolveComment({ done: true });
  part = text(view, 'xl/threadedComments/threadedComment1.xml');
  assert.match(part, new RegExp('id="' + top.replace(/[{}]/g, '\\$&') + '" done="1"'));
  assert.equal(view.render().cells.find((c) => c.ref === 'B2').thread.done, true);
  view.resolveComment({ done: false });
  assert.doesNotMatch(text(view, 'xl/threadedComments/threadedComment1.xml'), /done=/);
});

test('deleting a reply leaves the thread; deleting the first comment, or the thread, takes all of it', () => {
  const view = open();
  view.select(1, 1);
  const top = view.addComment({ text: 'One', author: 'A' });
  const reply = view.addComment({ text: 'Two', author: 'B' });
  view.deleteComment({ id: reply });
  assert.equal(view.render().thread.comments.length, 1);
  assert.doesNotMatch(legacy(view), /Reply:/);
  view.addComment({ text: 'Three', author: 'B' });
  view.deleteComment({ id: top });
  assert.equal(view.render().thread, null);
  assert.doesNotMatch(text(view, 'xl/threadedComments/threadedComment1.xml'), /<threadedComment\b/);
  assert.doesNotMatch(legacy(view), /<comment\b/, 'the shadow goes with it');

  view.select(2, 1);
  view.addComment({ text: 'Again', author: 'A' });
  view.deleteThread();
  assert.equal(view.render().threads, 0);
  assert.throws(() => view.deleteThread(), /no comment/);
});

test('each change is one undo step, and undoing the first comment takes its parts away', () => {
  const view = open();
  view.select(1, 1);
  view.addComment({ text: 'One', author: 'A' });
  view.addComment({ text: 'Two', author: 'A' });
  view.undo();
  assert.equal(view.render().thread.comments.length, 1);
  view.undo();
  assert.equal(view.render().thread, null);
  assert.equal(view.pkg.has('xl/threadedComments/threadedComment1.xml'), false);
  assert.equal(view.pkg.has('xl/persons/person.xml'), false);
  view.redo();
  assert.equal(view.render().thread.comments[0].text, 'One');
});

test('a cell with a note refuses a comment; a comment needs words', () => {
  const view = open();
  view.setNote({ row: 1, col: 1, author: 'A', text: 'A plain note' });
  assert.throws(() => view.addComment({ row: 1, col: 1, text: 'Hello', author: 'A' }), /has a note/);
  assert.throws(() => view.addComment({ row: 2, col: 1, text: '   ', author: 'A' }), /needs some words/);
  view.addComment({ row: 2, col: 1, text: 'Hello', author: 'A' });
  const frame = view.render();
  assert.equal(frame.notes, 1, 'the note stays a note');
  assert.equal(frame.threads, 1);
});

test('Previous and Next walk the threads across the sheets and round again; the pane lists them all', () => {
  const view = open();
  view.addComment({ row: 2, col: 0, text: 'third row', author: 'A' });
  view.addComment({ row: 0, col: 1, text: 'first row', author: 'A' });
  view.selectSheet('Notes');
  view.addComment({ row: 0, col: 0, text: 'other sheet', author: 'A' });
  view.selectSheet('Plan');
  view.select(0, 0);
  assert.equal(view.stepComment('next'), 'B1');
  assert.equal(view.stepComment('next'), 'A3');
  assert.equal(view.stepComment('next'), 'A1');
  assert.equal(view.activeSheet, 'Notes');
  assert.equal(view.stepComment('next'), 'B1', 'round again');
  assert.equal(view.activeSheet, 'Plan');
  assert.equal(view.stepComment('prev'), 'A1');
  assert.equal(view.activeSheet, 'Notes');
  view.commentsOpen = true;
  assert.deepEqual(view.render().comments.map((t) => t.sheet + '!' + t.ref), ['Plan!B1', 'Plan!A3', 'Notes!A1']);
});

test('saved and opened again, the conversation is as it was', () => {
  const view = open();
  view.select(1, 1);
  view.addComment({ text: 'Is this right?', author: 'Ejaz' });
  view.addComment({ text: 'It is.', author: 'Sana' });
  view.resolveComment({ done: true });
  const again = new SheetView(view.save());
  const t = again.workbook.threadAt('Plan', 'B2');
  assert.equal(t.done, true);
  assert.deepEqual(t.comments.map((c) => c.author + ': ' + c.text), ['Ejaz: Is this right?', 'Sana: It is.']);
});

test('threaded comments in a file written by Excel are read, their shadow is not taken for a note', () => {
  const pkg = OoxmlPackage.read(buildXlsx({ sheets: [{ name: 'Sheet1', rows: [['a', 'b']] }] }));
  const P1 = '{1F9D2A4C-0000-4000-8000-000000000001}';
  const P2 = '{1F9D2A4C-0000-4000-8000-000000000002}';
  const T = '{7A1B2C3D-0000-4000-8000-00000000000A}';
  const R = '{7A1B2C3D-0000-4000-8000-00000000000B}';
  pkg.addPart('xl/threadedComments/threadedComment1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    + `<threadedComment ref="A1" dT="2024-03-01T10:15:00.00" personId="${P1}" id="${T}" done="1"><text>Check the total &amp; the date</text></threadedComment>`
    + `<threadedComment ref="A1" dT="2024-03-01T11:00:00.00" personId="${P2}" id="${R}" parentId="${T}"><text>Done</text><mentions><mention mentionpersonId="${P1}" mentionId="{X}" startIndex="0" length="4"/></mentions></threadedComment>`
    + '</ThreadedComments>', 'application/vnd.ms-excel.threadedcomments+xml');
  pkg.addRelationshipTo('xl/worksheets/sheet1.xml', 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment', '../threadedComments/threadedComment1.xml');
  pkg.addPart('xl/persons/person.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    + `<person displayName="Ada Lovelace" id="${P1}" userId="ada@example.com" providerId="AD"/><person displayName="Grace Hopper" id="${P2}" userId="grace@example.com" providerId="AD"/></personList>`,
    'application/vnd.ms-excel.person+xml');
  pkg.addRelationshipTo('xl/workbook.xml', 'http://schemas.microsoft.com/office/2017/10/relationships/person', 'persons/person.xml');
  pkg.addPart('xl/comments1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors>'
    + `<author>tc=${T}</author></authors><commentList><comment ref="A1" authorId="0" shapeId="0"><text><t>[Threaded comment]\n\nComment:\n    Check the total</t></text></comment></commentList></comments>`,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml');
  pkg.addRelationshipTo('xl/worksheets/sheet1.xml', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments', '../comments1.xml');

  const view = new SheetView(pkg.write());
  const frame = view.render();
  const cell = frame.cells.find((c) => c.ref === 'A1');
  assert.equal(cell.note, null, 'the tc= shadow is the thread, not a note');
  assert.equal(frame.notes, 0);
  assert.deepEqual({ done: cell.thread.done, replies: cell.thread.replies, author: cell.thread.author, text: cell.thread.text },
    { done: true, replies: 1, author: 'Ada Lovelace', text: 'Check the total & the date' });
  assert.deepEqual(frame.thread.comments.map((c) => c.author + ': ' + c.text), ['Ada Lovelace: Check the total & the date', 'Grace Hopper: Done']);

  // A reply from here keeps what Excel wrote — the mention, the people — intact.
  view.addComment({ row: 0, col: 0, text: 'Thanks', author: 'Grace Hopper' });
  const written = view.pkg.text('xl/threadedComments/threadedComment1.xml');
  assert.match(written, /<mentions><mention mentionpersonId=/);
  assert.equal(view.workbook.persons().length, 2, 'Grace is known already');
  assert.equal((written.match(/<threadedComment\b/g) || []).length, 3);
});
