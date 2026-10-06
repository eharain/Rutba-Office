// Text holding $&, $', $` or $$, written by every engine as it was typed.
//
// In a replacement string those four are instructions, and the engines spliced
// what a person typed — or a stranger's mail carried — into their XML as one.
// 1.28.1 fixed the workbook; the document, the deck and Mail's automatic
// replies had the same splice. Each case below went through one of them.

import test from 'node:test';
import assert from 'node:assert/strict';
import { Document } from '@rutba/ooxml';
import { buildDocx } from '@rutba/ooxml/build';
import { Deck, buildPptx } from '@rutba/presentation';
import { buildReply } from '../apps/desktop/main/mail-ooo.js';

const NASTY = "costs $' and $& — and $$, and $`";

test('a comment in a document keeps every dollar, and the comments part stays whole', () => {
  const doc = Document.open(buildDocx({ paragraphs: ['First.', 'Second.'] }));
  doc.addComment(0, { author: 'A. Reader', text: 'one' });
  doc.addComment(1, { author: 'A. Reader', text: NASTY });
  const back = Document.open(doc.save());
  const texts = back.comments().map((c) => c.text);
  assert.deepEqual(texts, ['one', NASTY]);
});

test('a deck keeps a dollar in a section name and a slide comment', () => {
  const deck = Deck.open(buildPptx({ title: 'D', slides: [{ layout: 'title', title: 'One' }, { layout: 'title', title: 'Two' }] }));
  deck.addSection(0, 'Prices in US$');
  deck.renameSection(0, NASTY);
  deck.addComment(1, { text: NASTY, author: 'A. Reader' });
  const back = Deck.open(deck.save());
  assert.equal(back.sections()[0]?.name, NASTY, 'the section name');
  assert.equal(back.comments().find((c) => c.slide === 1)?.text, NASTY, 'the comment');
});

test('an automatic reply quotes a subject with dollars in it as the sender wrote it', () => {
  const reply = buildReply({ subject: "Invoice $'000 & $&", from: [{ address: 'a@example.com' }] }, {}, {});
  assert.equal(reply.subject, "Automatic reply: Invoice $'000 & $&");
});
