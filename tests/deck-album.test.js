// Insert → Photo Album: a new deck of pictures — a title slide, then the
// pictures one, two or four to a slide, each fitted to its share without
// stretching, captioned with its name when asked — and the document service
// making it a session for a window of its own.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Deck, photoAlbum } from '@rutba/presentation';
import { createDocumentService } from '../apps/desktop/main/documents.js';

const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC', 'base64');
const pictures = (sizes) => sizes.map(([width, height], i) => ({ data: PNG, contentType: 'image/png', name: `Holiday ${i + 1}.png`, width, height }));
const picturesOn = (deck, i) => deck.slide(i).shapes.filter((s) => s.kind === 'picture');
const overlap = (a, b) => a.x < b.x + b.w && b.x < a.x + a.w && a.y < b.y + b.h && b.y < a.y + a.h;

test('one to a slide: a title slide, then each picture fitted inside the slide in its own shape', () => {
  const deck = photoAlbum({ pictures: pictures([[4000, 3000], [1080, 1920]]), title: 'Summer', subtitle: 'by Jo' });
  assert.equal(deck.slideCount, 3);
  assert.match(deck.slide(0).shapes.map((s) => (s.text?.paragraphs || []).map((p) => p.runs.map((r) => r.text).join('')).join(' ')).join(' | '), /Summer.*by Jo/);
  const { width, height } = deck.size;
  for (const [i, [iw, ih]] of [[1, [4000, 3000]], [2, [1080, 1920]]]) {
    const [p] = picturesOn(deck, i);
    const g = p.geometry;
    assert.ok(g.x >= 0 && g.y >= 0 && g.x + g.w <= width && g.y + g.h <= height, `slide ${i} picture inside the slide`);
    assert.ok(Math.abs(g.w / g.h - iw / ih) < 0.02, `slide ${i} keeps its shape: ${g.w}x${g.h}`);
  }
});

test('four to a slide, captioned: a two-by-two grid with nothing over anything, and each name under its picture', () => {
  const deck = Deck.open(photoAlbum({ pictures: pictures([[800, 600], [600, 800], [1000, 1000], [1600, 900], [800, 600]]), perSlide: 4, captions: true }).save());
  assert.equal(deck.slideCount, 3, 'a title, four on the first slide, one on the second');
  const four = picturesOn(deck, 1).map((s) => s.geometry);
  assert.equal(four.length, 4);
  for (let a = 0; a < 4; a++) for (let b = a + 1; b < 4; b++) assert.ok(!overlap(four[a], four[b]), `pictures ${a} and ${b} apart`);
  const captions = deck.slide(1).shapes.filter((s) => s.kind !== 'picture' && s.text).map((s) => ({ text: s.text.paragraphs[0].runs.map((r) => r.text).join(''), g: s.geometry }));
  assert.deepEqual(captions.map((c) => c.text), ['Holiday 1', 'Holiday 2', 'Holiday 3', 'Holiday 4']);
  captions.forEach((c, i) => assert.ok(c.g.y >= four[i].y + four[i].h - 1, `caption ${i + 1} under its picture`));
  assert.equal(picturesOn(deck, 2).length, 1);
});

test('two to a slide sit side by side; any other count, or no picture, is refused', () => {
  const deck = photoAlbum({ pictures: pictures([[800, 600], [800, 600]]), perSlide: 2 });
  const [a, b] = picturesOn(deck, 1).map((s) => s.geometry);
  assert.ok(Math.abs(a.y - b.y) <= 1 && a.x + a.w <= b.x, 'level, the first on the left');
  assert.throws(() => photoAlbum({ pictures: pictures([[1, 1]]), perSlide: 3 }), /1, 2 or 4/);
  assert.throws(() => photoAlbum({ pictures: [] }), /needs a picture/);
});

test('the document service makes the album a session for a window to adopt, and refuses a file that is not a picture', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-album-'));
  const files = ['a.png', 'b.png', 'c.png'].map((n) => { const f = path.join(dir, n); fs.writeFileSync(f, PNG); return f; });
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const made = docs.photoAlbum({ files, perSlide: 2, captions: true, title: 'Trip' });
  assert.equal(made.name, 'Photo Album1');
  assert.equal(made.slides, 3);
  const adopted = docs.adopt({ id: made.id });
  assert.equal(adopted.model.count, 3);
  assert.equal(adopted.dirty, true, 'unsaved, as a new album is');
  assert.equal(docs.photoAlbum({ files: [files[0]] }).name, 'Photo Album2');
  const text = path.join(dir, 'notes.txt');
  fs.writeFileSync(text, 'hello');
  assert.throws(() => docs.photoAlbum({ files: [text] }), /notes\.txt is not a picture/);
});
