/**
 * What a file is, decided from its bytes.
 *
 * This is the first thing that happens to every file the suite is handed, and
 * it had no test of its own: which app a double-click opens, whether a
 * document is refused, and what the refusal says all follow from here.
 *
 * The case that prompted the file: a text file saved as Unicode on Windows
 * begins FF FE, and FF FE is also a valid MPEG frame sync, so every UTF-16
 * text file in a folder was identified as an MP3 and refused with "Rutba
 * Office cannot open MP3 Audio files yet." Notepad writes that encoding when
 * you choose Unicode, and PowerShell writes it whenever it redirects.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { sniff, appFor, kindFromExtension } from '@rutba/office-formats/sniff';
import { decodeText } from '@rutba/office-formats/text';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { buildPptx } from '@rutba/presentation';

const wide = (text, encoding = 'utf16le') =>
  Buffer.concat([
    Buffer.from(encoding === 'utf16le' ? [0xff, 0xfe] : [0xfe, 0xff]),
    encoding === 'utf16le' ? Buffer.from(text, 'utf16le') : Buffer.from(text, 'utf16le').swap16(),
  ]);

test('a UTF-16 text file is text, not an MP3', () => {
  const le = sniff(wide('Hello there.\nSecond line.\n'), 'notes.txt');
  assert.equal(le.kind, 'txt', 'little-endian');
  assert.equal(le.app, 'word');
  assert.equal(le.encoding, 'utf-16le');

  const be = sniff(wide('Hello there.\n', 'utf16be'), 'notes.txt');
  assert.equal(be.kind, 'txt', 'big-endian');
  assert.equal(be.encoding, 'utf-16be');

  // The content decides as it does for any other text: a comma-separated file
  // is a workbook, a page is a page.
  assert.equal(sniff(wide('a,b,c\n1,2,3\n'), 'table.csv').kind, 'csv');
  assert.equal(sniff(wide('<!doctype html><p>hi</p>'), 'page.html').kind, 'html');

  // And the reader agrees with the identification.
  assert.equal(decodeText(wide('Hello there.\n')), 'Hello there.\n');
});

test('a real MP3 is still an MP3, and UTF-32 is neither', () => {
  assert.equal(sniff(Buffer.from([0xff, 0xfb, 0x90, 0x00, 1, 2, 3, 4]), 'song.mp3').kind, 'mp3', 'a frame sync');
  assert.equal(sniff(Buffer.concat([Buffer.from('ID3'), Buffer.alloc(20)]), 'song.mp3').kind, 'mp3', 'a tag');
  // FF FE 00 00 is UTF-32LE, which shares UTF-16LE's mark. Nothing here reads
  // it; calling it audio would be a refusal about the wrong thing.
  assert.notEqual(sniff(Buffer.from([0xff, 0xfe, 0x00, 0x00, 0x41, 0, 0, 0]), 'x.bin').kind, 'mp3');
});

test('the formats the suite edits are identified from their bytes, and the extension only breaks ties', () => {
  const docx = buildDocx({ paragraphs: ['One.'] });
  const xlsx = buildXlsx({ sheets: [{ name: 'S', rows: [[1]] }] });
  const pptx = buildPptx({ title: 'D', slides: [{ layout: 'title', title: 'T' }] });
  assert.equal(sniff(docx, 'a.docx').kind, 'docx');
  assert.equal(sniff(xlsx, 'a.xlsx').kind, 'xlsx');
  assert.equal(sniff(pptx, 'a.pptx').kind, 'pptx');
  assert.equal(appFor(sniff(docx, 'a.docx').kind), 'word');
  assert.equal(appFor(sniff(xlsx, 'a.xlsx').kind), 'sheets');
  assert.equal(appFor(sniff(pptx, 'a.pptx').kind), 'slides');
  // An OOXML package whose extension says nothing is deep-scanned by the caller.
  const nameless = sniff(docx, 'a.zip');
  assert.equal(nameless.container, 'ooxml');
  assert.ok(nameless.needsDeepScan, 'the caller is told to read the part names');
});

test('a file that is not what its name says is identified by its bytes', () => {
  const docx = buildDocx({ paragraphs: ['One.'] });
  // A document renamed to .xlsx: the container says OOXML and the extension is
  // believed, which is why the app that opens it has to check the kind it got.
  assert.equal(sniff(docx, 'renamed.xlsx').container, 'ooxml');
  // Text under a document's name keeps the name — the extension is all there
  // is to go on, and the refusal when it fails to open says what it found.
  const named = sniff(Buffer.from('this is not a zip archive, whatever the name says'), 'a.docx');
  assert.equal(named.kind, 'docx');
  assert.equal(named.confidence, 'extension');
  assert.equal(sniff(Buffer.from('plain words with no extension to go on'), 'a.unknownext').kind, 'txt');
  assert.equal(sniff(Buffer.alloc(0), 'a.docx').kind, 'docx', 'an empty file is only its name');
  assert.equal(kindFromExtension('REPORT.DOCX'), 'docx', 'case does not matter');
});
