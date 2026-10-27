// Windows' own stars and tags, read (2.21 D4): the EXIF Rating and
// RatingPercent and Windows' XPKeywords, and the XMP xmp:Rating and
// dc:subject keywords Windows Photos and Lightroom write, read from the file
// and put together the way Explorer shows them; and Pictures' Use, which
// takes them as its own beside the tags it has.
import test from 'node:test';
import assert from 'node:assert/strict';
import { readXmp, readMetadata, fileMarks, starsFromPercent } from '@rutba/imaging/exif';
import { adoptMarks } from '../apps/desktop/renderer/apps/pictures/marks.js';
import { tiffBlock, xp, jpegWith, pngWith, webpWith, heifWith, XMP } from './exif-fixtures.js';

test('XMP: the rating and keywords as attributes or elements, under whatever prefix the packet gives them', () => {
  assert.deepEqual(readXmp(XMP), { rating: 3, ratingPercent: 50, keywords: ['Holiday', 'Beach', 'Tom & Ann'] });
  const elements = '<x:xmpmeta><rdf:RDF><rdf:Description xmlns:ap="http://ns.adobe.com/xap/1.0/" xmlns:d="http://purl.org/dc/elements/1.1/"><ap:Rating>5</ap:Rating><d:subject><rdf:Bag><rdf:li>One</rdf:li><rdf:li xml:lang="x-default">Two&#x2019;s</rdf:li></rdf:Bag></d:subject></rdf:Description></rdf:RDF></x:xmpmeta>';
  assert.deepEqual(readXmp(elements), { rating: 5, keywords: ['One', 'Two’s'] });
  assert.deepEqual(readXmp('<rdf:Description xmp:Rating="-1"/>'), { rejected: true }, 'Lightroom\'s rejected');
  assert.deepEqual(readXmp('<rdf:Description xmp:Rating="9"/>'), {}, 'a rating out of range is not one');
  // Windows' own copy of the keywords beside dc:subject, duplicates once.
  const both = '<rdf:Description xmlns:MicrosoftPhoto="http://ns.microsoft.com/photo/1.0/"><dc:subject><rdf:Bag><rdf:li>Cat</rdf:li></rdf:Bag></dc:subject><MicrosoftPhoto:LastKeywordXMP><rdf:Bag><rdf:li>cat</rdf:li><rdf:li>Garden</rdf:li></rdf:Bag></MicrosoftPhoto:LastKeywordXMP></rdf:Description>';
  assert.deepEqual(readXmp(both).keywords, ['Cat', 'Garden']);
});

test('XMP is bounded: a packet of a thousand keywords keeps two hundred, and a huge one reads quickly', () => {
  const many = `<dc:subject><rdf:Bag>${Array.from({ length: 1000 }, (_, i) => `<rdf:li>k${i}</rdf:li>`).join('')}</rdf:Bag></dc:subject>`;
  assert.equal(readXmp(many).keywords.length, 200);
  const started = Date.now();
  readXmp(`<x>${'a'.repeat(5 * 1024 * 1024)}`);
  assert.ok(Date.now() - started < 1500);
});

test('Windows\' percent rating as stars, the way Explorer turns one into the other', () => {
  assert.deepEqual([0, 1, 12, 13, 25, 37, 38, 50, 62, 63, 75, 87, 88, 99, 100].map(starsFromPercent), [0, 1, 1, 2, 2, 2, 3, 3, 3, 4, 4, 4, 5, 5, 5]);
});

test('which rating wins: XMP\'s, then EXIF Rating, then RatingPercent; keywords from XMP and XPKeywords together', () => {
  assert.deepEqual(fileMarks({ exif: { rating: 2 }, xmp: { rating: 4 } }), { rating: 4, rejected: false, tags: [] });
  assert.equal(fileMarks({ exif: { rating: 2, ratingPercent: 99 } }).rating, 2);
  assert.equal(fileMarks({ exif: { ratingPercent: 75 } }).rating, 4);
  assert.equal(fileMarks({ xmp: { ratingPercent: 1 } }).rating, 1);
  assert.equal(fileMarks({}).rating, null, 'no rating is not a rating of nought');
  assert.deepEqual(fileMarks({ exif: { xpKeywords: 'beach; family;;Beach' }, xmp: { keywords: ['Holiday'] } }).tags, ['Holiday', 'beach', 'family']);
});

test('the file\'s marks read from a JPEG, PNG, WebP and HEIC, from EXIF alone (Explorer) and from XMP (Lightroom)', () => {
  const windows = tiffBlock({ ifd0: [[0x4746, 3, 2], [0x4749, 3, 25], [0x9c9e, 1, xp('Family;Garden')]] });
  for (const make of [(o) => jpegWith(o), (o) => pngWith(o), (o) => webpWith(o), (o) => heifWith(o)]) {
    assert.deepEqual(readMetadata(make({ block: windows })).marks, { rating: 2, rejected: false, tags: ['Family', 'Garden'] });
    assert.deepEqual(readMetadata(make({ xmp: XMP })).marks, { rating: 3, rejected: false, tags: ['Holiday', 'Beach', 'Tom & Ann'] });
  }
});

test('Use takes the file\'s rating, and adds its tags to the ones Pictures has, each once', () => {
  assert.deepEqual(adoptMarks({ rating: 1, tags: ['beach'] }, { rating: 4, tags: ['Beach', 'Holiday'] }), { rating: 4, tags: ['beach', 'Holiday'], changed: true });
  assert.deepEqual(adoptMarks({ rating: 3 }, { rating: null, tags: [] }), { rating: 3, tags: [], changed: false }, 'no rating in the file leaves Pictures\' own');
  assert.equal(adoptMarks({ rating: 4, tags: ['Holiday'] }, { rating: 4, tags: ['holiday'] }).changed, false, 'nothing to take: no Use');
});
