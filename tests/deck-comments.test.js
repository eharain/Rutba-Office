// Review → Comments — the deck engine's half.
//
// Modern comments as PowerPoint 365 writes them: a p188:cmLst part per
// slide named for the slide's id and creation id, related from the slide
// and named again in its extension list, anchored by PowerPoint's own
// monikers, the people in ppt/authors.xml. Older p:cmLst comments are read.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';
import { OoxmlPackage } from '@rutba/ooxml/package';

const DECK = buildPptx({ title: 'Comments', slides: [{ layout: 'title', title: 'One', body: 'a' }, { layout: 'obj', title: 'Two', body: ['b'] }, { layout: 'obj', title: 'Three', body: ['c'] }] });
const partOf = (deck, re) => deck.pkg.partNames().find((p) => re.test(p)) || null;

test('a new comment is written as PowerPoint 365 writes one: its part, its relationship and extension, the slide\'s creation id and the authors part', () => {
  const deck = Deck.open(DECK);
  const id = deck.addComment(1, { text: 'Is this the final figure?', author: 'Jane Doe' });
  assert.match(id, /^\{[0-9A-F-]{36}\}$/);
  const slide = deck.pkg.text('ppt/slides/slide2.xml');
  const cId = /<p14:creationId xmlns:p14="http:\/\/schemas\.microsoft\.com\/office\/powerpoint\/2010\/main" val="(\d+)"\/>/.exec(slide)?.[1];
  assert.ok(cId, 'the slide carries a creation id');
  const part = `ppt/comments/modernComment_101_${Number(cId).toString(16).toUpperCase()}.xml`;
  assert.ok(deck.pkg.has(part), `the part is named for sldId 257 and the creation id: ${partOf(deck, /modernComment/)}`);
  assert.match(deck.pkg.text('[Content_Types].xml'), new RegExp(`PartName="/${part.replace(/\./g, '\\.')}" ContentType="application/vnd\\.ms-powerpoint\\.comments\\+xml"`));
  assert.match(deck.pkg.text('ppt/slides/_rels/slide2.xml.rels'), /Type="http:\/\/schemas\.microsoft\.com\/office\/2018\/10\/relationships\/comments" Target="\.\.\/comments\/modernComment_101_/);
  assert.match(slide, /<\/p:clrMapOvr><p:extLst><p:ext uri="\{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E\}">[\s\S]*<p:ext uri="\{6950BFC3-D8DA-4A85-94F7-54DA5524770B\}"><p188:commentRel xmlns:p188="http:\/\/schemas\.microsoft\.com\/office\/powerpoint\/2018\/8\/main" r:id="rId\d+"\/><\/p:ext><\/p:extLst><\/p:sld>$/);
  const xml = deck.pkg.text(part);
  assert.match(xml, new RegExp(`<p188:cm id="\\${id.slice(0, -1)}\\}" authorId="\\{[0-9A-F-]{36}\\}" created="\\d{4}-\\d\\d-\\d\\dT\\d\\d:\\d\\d:\\d\\d\\.\\d{3}"><pc:sldMkLst xmlns:pc="http://schemas\\.microsoft\\.com/office/powerpoint/2013/main/command"><pc:docMk/><pc:sldMk cId="${cId}" sldId="257"/></pc:sldMkLst><p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Is this the final figure\\?</a:t></a:r></a:p></p188:txBody></p188:cm>`));
  assert.match(deck.pkg.text('ppt/authors.xml'), /<p188:author id="\{[0-9A-F-]{36}\}" name="Jane Doe" initials="JD" userId="Jane Doe" providerId="None"\/>/);
  assert.match(deck.pkg.text('ppt/_rels/presentation.xml.rels'), /relationships\/authors" Target="authors\.xml"/);
  assert.match(deck.pkg.text('[Content_Types].xml'), /PartName="\/ppt\/authors\.xml" ContentType="application\/vnd\.ms-powerpoint\.authors\+xml"/);
  const [c] = deck.comments();
  assert.deepEqual({ slide: c.slide, author: c.author, initials: c.initials, text: c.text, status: c.status, shapeId: c.shapeId, legacy: c.legacy }, { slide: 1, author: 'Jane Doe', initials: 'JD', text: 'Is this the final figure?', status: 'active', shapeId: null, legacy: false });
});

test('a comment on a shape names it by id and creation id, and the shape is given one', () => {
  const deck = Deck.open(DECK);
  const title = deck.slide(2).shapes.find((s) => s.placeholder?.type === 'title');
  deck.addComment(2, { text: 'Shorter title?', author: 'Sam Lee', shape: title.id });
  const slide = deck.pkg.text('ppt/slides/slide3.xml');
  const shapeGuid = new RegExp(`<p:cNvPr id="${title.id}" name="Title 1"><a:extLst><a:ext uri="\\{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236\\}"><a16:creationId xmlns:a16="http://schemas\\.microsoft\\.com/office/drawing/2014/main" id="(\\{[0-9A-F-]{36}\\})"/></a:ext></a:extLst></p:cNvPr>`).exec(slide)?.[1];
  assert.ok(shapeGuid, 'the shape carries a creation id');
  const xml = deck.pkg.text(partOf(deck, /modernComment/));
  assert.match(xml, new RegExp(`<ac:deMkLst xmlns:ac="http://schemas\\.microsoft\\.com/office/drawing/2013/main/command"><pc:docMk xmlns:pc="[^"]+"/><pc:sldMk xmlns:pc="[^"]+" cId="\\d+" sldId="258"/><ac:spMk id="${title.id}" creationId="\\${shapeGuid.slice(0, -1)}\\}"/></ac:deMkLst>`));
  assert.equal(deck.comments()[0].shapeId, String(title.id));
  assert.equal(deck.slide(2).shapes.find((s) => s.id === title.id).text.paragraphs[0].runs[0].text, 'Three', 'the shape is otherwise untouched');
});

test('replies go in the thread before its words, each author listed once', () => {
  const deck = Deck.open(DECK);
  const id = deck.addComment(0, { text: 'Start here', author: 'Jane Doe' });
  deck.replyComment(0, id, { text: 'Agreed', author: 'Sam Lee' });
  deck.replyComment(0, id, { text: 'Done', author: 'Jane Doe' });
  const xml = deck.pkg.text(partOf(deck, /modernComment/));
  assert.match(xml, /<\/pc:sldMkLst><p188:replyLst><p188:reply id="\{[^"]+\}" authorId="\{[^"]+\}" created="[^"]+"><p188:txBody>[\s\S]*?Agreed[\s\S]*?<\/p188:reply><p188:reply [\s\S]*?Done[\s\S]*?<\/p188:replyLst><p188:txBody>[\s\S]*?Start here/);
  const t = deck.comments()[0];
  assert.deepEqual(t.replies.map((r) => [r.author, r.text]), [['Sam Lee', 'Agreed'], ['Jane Doe', 'Done']]);
  assert.equal((deck.pkg.text('ppt/authors.xml').match(/<p188:author /g) || []).length, 2);
  assert.throws(() => deck.replyComment(0, '{00000000-0000-0000-0000-000000000000}', { text: 'x' }), /no such comment/);
  assert.throws(() => deck.addComment(0, { text: '   ' }), /some words/);
});

test('Resolve marks a thread resolved and Reopen takes it back', () => {
  const deck = Deck.open(DECK);
  const id = deck.addComment(1, { text: 'Fix the chart', author: 'Jane Doe' });
  assert.equal(deck.resolveComment(1, id, true), true);
  assert.match(deck.pkg.text(partOf(deck, /modernComment/)), /<p188:cm id="[^"]+" authorId="[^"]+" status="resolved" created=/);
  assert.equal(deck.comments()[0].status, 'resolved');
  assert.equal(deck.resolveComment(1, id, false), true);
  assert.equal(deck.comments()[0].status, 'active');
  assert.doesNotMatch(deck.pkg.text(partOf(deck, /modernComment/)), /status=/);
});

test('Delete takes a reply, then a thread; the last one takes the part, its relationship and the extension with it', () => {
  const deck = Deck.open(DECK);
  const before = deck.pkg.text('ppt/slides/_rels/slide2.xml.rels');
  const a = deck.addComment(1, { text: 'First', author: 'Jane Doe' });
  const b = deck.addComment(1, { text: 'Second', author: 'Jane Doe' });
  const r = deck.replyComment(1, a, { text: 'Reply', author: 'Sam Lee' });
  deck.removeComment(1, a, { reply: r });
  assert.equal(deck.comments().find((c) => c.id === a).replies.length, 0);
  assert.doesNotMatch(deck.pkg.text(partOf(deck, /modernComment/)), /replyLst/);
  deck.removeComment(1, a);
  assert.deepEqual(deck.comments().map((c) => c.text), ['Second']);
  const part = partOf(deck, /modernComment/);
  deck.removeComment(1, b);
  assert.equal(deck.pkg.has(part), false);
  assert.equal(deck.pkg.text('ppt/slides/_rels/slide2.xml.rels'), before);
  assert.doesNotMatch(deck.pkg.text('ppt/slides/slide2.xml'), /commentRel/);
  assert.match(deck.pkg.text('ppt/slides/slide2.xml'), /p14:creationId/, 'the slide keeps its creation id');
  assert.doesNotMatch(deck.pkg.text('[Content_Types].xml'), /modernComment/);
  assert.ok(Deck.open(deck.save()));
});

test('an older PowerPoint\'s comments are read, replies threaded under them, deleted with them, and not replied to', () => {
  const pkg = OoxmlPackage.read(DECK);
  pkg.addPart('ppt/commentAuthors.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cmAuthor id="0" name="Old Hand" initials="OH" lastIdx="2" clrIdx="0"/><p:cmAuthor id="1" name="New Voice" initials="NV" lastIdx="1" clrIdx="1"/></p:cmAuthorLst>', 'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml');
  pkg.addRelationshipTo('ppt/presentation.xml', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors', 'commentAuthors.xml');
  pkg.addPart('ppt/comments/comment1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main">' +
    '<p:cm authorId="0" dt="2019-03-01T10:00:00.000" idx="1"><p:pos x="480" y="96"/><p:text>Legacy note</p:text></p:cm>' +
    '<p:cm authorId="1" dt="2019-03-02T10:00:00.000" idx="1"><p:pos x="480" y="96"/><p:text>Legacy reply</p:text><p:extLst><p:ext uri="{C676402C-5697-4E1C-873F-D02D1690AC5C}"><p15:threadingInfo timeZoneBias="0"><p15:parentCm authorId="0" idx="1"/></p15:threadingInfo></p:ext></p:extLst></p:cm>' +
    '<p:cm authorId="0" dt="2019-03-03T10:00:00.000" idx="2"><p:pos x="10" y="10"/><p:text>Another</p:text></p:cm></p:cmLst>', 'application/vnd.openxmlformats-officedocument.presentationml.comments+xml');
  pkg.addRelationshipTo('ppt/slides/slide1.xml', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments', '../comments/comment1.xml');
  const deck = Deck.open(pkg.write());
  const threads = deck.comments();
  assert.deepEqual(threads.map((t) => [t.author, t.text, t.replies.map((r) => r.text).join('|'), t.legacy]), [['Old Hand', 'Legacy note', 'Legacy reply', true], ['Old Hand', 'Another', '', true]]);
  assert.deepEqual(threads[0].pos, { x: 80, y: 16 }, 'eighths of a point, as pixels');
  assert.throws(() => deck.replyComment(0, threads[0].id, { text: 'x' }), /older version/);
  deck.removeComment(0, threads[0].id);
  assert.deepEqual(deck.comments().map((t) => t.text), ['Another'], 'the thread went with its reply');
  deck.removeComment(0, deck.comments()[0].id);
  assert.equal(deck.pkg.has('ppt/comments/comment1.xml'), false);
  assert.doesNotMatch(deck.pkg.text('ppt/slides/_rels/slide1.xml.rels'), /comments/);
});

test('comments survive a save, and Delete All clears a slide or the whole deck', () => {
  const deck = Deck.open(DECK);
  deck.addComment(0, { text: 'a', author: 'X' });
  deck.addComment(0, { text: 'b', author: 'X' });
  deck.addComment(2, { text: 'c', author: 'Y', x: 200, y: 100 });
  const again = Deck.open(deck.save());
  assert.deepEqual(again.comments().map((c) => [c.slide, c.text]), [[0, 'a'], [0, 'b'], [2, 'c']]);
  assert.deepEqual(again.comments()[2].pos, { x: 200, y: 100 });
  assert.equal(again.removeAllComments(0), 2);
  assert.deepEqual(again.comments().map((c) => c.text), ['c']);
  assert.equal(again.removeAllComments(), 1);
  assert.equal(again.comments().length, 0);
});

test('adding a comment is one undo step: undone, the new parts and relationships are gone', () => {
  const deck = Deck.open(DECK);
  const types = deck.pkg.text('[Content_Types].xml');
  const snap = deck.snapshot();
  deck.addComment(1, { text: 'Undo me', author: 'Jane Doe' });
  deck.pushUndo(snap);
  assert.equal(deck.comments().length, 1);
  deck.undo();
  assert.equal(deck.comments().length, 0);
  assert.equal(partOf(deck, /modernComment|authors\.xml/), null);
  assert.equal(deck.pkg.text('[Content_Types].xml'), types);
  deck.redo();
  assert.equal(deck.comments()[0].text, 'Undo me');
});
