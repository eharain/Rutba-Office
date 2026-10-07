/**
 * Windows-1252 read as Windows-1252, in every runtime.
 *
 * Node 24 — the Node in Electron's main process, where files are opened —
 * decodes `new TextDecoder('windows-1252')` as Latin-1: the curly quotes,
 * dashes and euro sign of an RTF or Word 97 file came out as invisible C1
 * control characters. The readers decode it by table instead.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { decode1252, encode1252, decodeIn } from '@rutba/office-formats/codepage';
import { readRtf } from '@rutba/office-formats/rtf';
import { decodeText } from '@rutba/office-formats/text';

test('Windows-1252\'s quotes, dashes and euro sign are themselves', () => {
  assert.equal(decode1252(Uint8Array.of(0x93, 0x48, 0x69, 0x94, 0x20, 0x96, 0x20, 0x97, 0x20, 0x80, 0x20, 0x85)), '“Hi” – — € …');
  assert.deepEqual([...encode1252('“€”')], [0x93, 0x80, 0x94]);
  assert.equal(decodeIn('cp850', Uint8Array.of(0x43, 0x61, 0x66, 0x82)), 'Café');
  assert.equal(decodeIn('windows-1251', Uint8Array.of(0xcf, 0xf0)), 'Пр');
});

test('an RTF\'s escaped curly quotes are curly quotes', () => {
  const rtf = readRtf(new TextEncoder().encode("{\\rtf1\\ansi\\ansicpg1252 \\'93Quoted\\'94 \\'96 dash\\par}"));
  const text = rtf.blocks.map((b) => (b.runs || []).map((r) => r.text).join('') || b.text || '').join('');
  assert.equal(text, '“Quoted” – dash');
});

test('a text file that is not UTF-8 is read as Windows-1252', () => {
  assert.equal(decodeText(Uint8Array.of(0x93, 0x6f, 0x6b, 0x94, 0xe9, 0xe9, 0xe9, 0xe9, 0xe9, 0xe9)), '“ok”éééééé');
});
