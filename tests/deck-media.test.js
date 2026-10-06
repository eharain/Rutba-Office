// Insert → Video and Audio — the engine's half.
//
// The media goes in the package as PowerPoint puts it: the bytes in
// ppt/media/mediaN.ext with a default content type, related from the slide
// twice — the 2007 video/audio link and PowerPoint 2010's media embed, to
// the same part — and a picture whose nvPr names it, drawn as its poster,
// a click on it playing it in the show. The reader gives the media back,
// and the window's model carries a URL to play it from.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';
import { createDocumentService } from '../apps/desktop/main/documents.js';

// A one-pixel PNG for the poster, and stand-in media bytes.
const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64');
const MP4 = Buffer.from('000000186674797069736f6d0000020069736f6d69736f32', 'hex');
const MP3 = Buffer.from('4944330300000000000a', 'hex');
const blank = () => Deck.open(buildPptx({ title: 'Media', slides: [{ layout: 'blank' }] }));

test('a video goes in as PowerPoint puts one: its part, both relationships, and a picture that names it', () => {
  const deck = blank();
  const { id, part } = deck.addMedia(0, { kind: 'video', data: MP4, contentType: 'video/mp4', poster: { data: PNG, contentType: 'image/png' }, name: 'Launch film', x: 100, y: 80, w: 640, h: 360 });
  assert.equal(part, 'ppt/media/media1.mp4');
  assert.deepEqual([...deck.media(part)], [...MP4]);
  assert.match(deck.pkg.text('[Content_Types].xml'), /<Default Extension="mp4" ContentType="video\/mp4"\/>/);
  const rels = [...deck.pkg.rels('ppt/slides/slide1.xml')];
  const media = rels.find((r) => r.Type === 'http://schemas.microsoft.com/office/2007/relationships/media');
  const video = rels.find((r) => r.Type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video');
  assert.equal(media.Target, '../media/media1.mp4');
  assert.equal(video.Target, '../media/media1.mp4');
  const xml = deck.pkg.text('ppt/slides/slide1.xml');
  assert.match(xml, new RegExp(`<p:cNvPr id="${id}" name="Launch film"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr>`));
  assert.match(xml, new RegExp(`<a:videoFile r:link="${video.Id}"/><p:extLst><p:ext uri="\\{DAA4B4D4-6D71-4841-9C94-3DA8AD8B0F8F\\}"><p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="${media.Id}"/>`));
  const shape = deck.slide(0).shapes.find((s) => String(s.id) === String(id));
  assert.equal(shape.kind, 'picture');
  assert.equal(shape.media.kind, 'video');
  assert.equal(shape.media.source.part, 'ppt/media/media1.mp4');
  assert.ok(shape.source?.part?.startsWith('ppt/media/image'), 'drawn as its poster');
});

test('audio goes in the same way, and a second clip takes the next part', () => {
  const deck = blank();
  deck.addMedia(0, { kind: 'audio', data: MP3, contentType: 'audio/mpeg', poster: { data: PNG, contentType: 'image/png' }, w: 48, h: 48 });
  const { part } = deck.addMedia(0, { kind: 'audio', data: MP3, contentType: 'audio/mpeg', poster: { data: PNG, contentType: 'image/png' }, w: 48, h: 48 });
  assert.equal(part, 'ppt/media/media2.mp3');
  assert.match(deck.pkg.text('ppt/slides/slide1.xml'), /<a:audioFile r:link="rId\d+"\/>/);
  assert.deepEqual(deck.slide(0).shapes.filter((s) => s.media).map((s) => s.media.kind), ['audio', 'audio']);
  assert.throws(() => deck.addMedia(0, { kind: 'video', data: MP3, contentType: 'audio/mpeg', poster: { data: PNG, contentType: 'image/png' }, w: 4, h: 4 }), /unsupported video type/);
  assert.throws(() => deck.addMedia(0, { kind: 'video', data: MP4, contentType: 'video/x-flv', poster: { data: PNG, contentType: 'image/png' }, w: 4, h: 4 }), /unsupported video type/);
});

test('the window\'s model carries the media and a URL to play it, and the saved deck keeps it', () => {
  const held = [];
  const docs = createDocumentService({ holdBlob: (bytes, type) => { held.push({ size: bytes.length, type }); return { url: `blob:${held.length}` }; } });
  const { id } = docs.new({ kind: 'deck' });
  docs.apply({ id, ops: [{ op: 'addMedia', slide: 0, kind: 'video', data: new Uint8Array(MP4), contentType: 'video/mp4', poster: new Uint8Array(PNG), name: 'Clip', x: 10, y: 10, w: 320, h: 180 }] });
  const shape = docs.model({ id }).slide.shapes.find((s) => s.media);
  assert.equal(shape.media.kind, 'video');
  assert.match(shape.media.url, /^blob:/);
  assert.ok(held.some((h) => h.type === 'video/mp4' && h.size === MP4.length));
});
