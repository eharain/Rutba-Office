// The Locker asks whether an Office window shows one of its copies by reading
// the window's address. windows.js writes that address with URLSearchParams,
// which turns a space into "+", and 2.25.0 searched the decoded text for the
// path: a copy with a space anywhere in its path was never seen as open, and
// locking wiped it from under the document.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { addressParam, addressShows, addressApp, sameFile } from '../apps/desktop/main/extras/locker/window-address.js';

// As windows.js urlFor writes it.
const urlFor = (app, file, extra = {}) => {
  const params = new URLSearchParams({ app });
  if (file) params.set('file', file);
  for (const [k, v] of Object.entries(extra)) params.set(k, v);
  return `rutba-office://app/index.html?${params.toString()}`;
};

const WIN = 'C:\\Users\\Sam Khan\\AppData\\Roaming\\Rutba Office\\locker\\open\\a1b2\\Larkspur staff handbook.docx';
const PLUS = 'C:\\Users\\sam\\AppData\\Roaming\\RutbaOffice\\locker\\open\\a1b2\\C++ notes+more.docx';
const ODD = 'C:\\Users\\sam\\locker\\open\\a1b2\\50% off & more #2 \u00e9t\u00e9.docx';

test('a path with a space is read back whole, where the decoded address has a "+" in its place', () => {
  const url = urlFor('word', WIN);
  assert.equal(addressParam(url, 'file'), WIN);
  assert.equal(addressShows(url, WIN, 'win32'), true);
  // What 2.25.0 tested, kept here so the cause stays written down.
  assert.equal(decodeURIComponent(url).includes(WIN), false);
});

test('a "+" in a name stays a "+", and other characters an address escapes come back as they were', () => {
  assert.equal(addressShows(urlFor('word', PLUS), PLUS, 'win32'), true);
  assert.equal(addressShows(urlFor('sheets', ODD, { lang: 'ur' }), ODD, 'win32'), true);
});

test('on Windows case does not count; elsewhere it does', () => {
  assert.equal(addressShows(urlFor('word', WIN), WIN.toUpperCase(), 'win32'), true);
  assert.equal(sameFile('/home/sam/A.docx', '/home/sam/a.docx', 'linux'), false);
  assert.equal(sameFile('/home/sam/a.docx', '/home/sam/a.docx', 'linux'), true);
});

test('a file whose path only begins the same way is not the file shown', () => {
  // The old test searched the text, so a window showing "a.docx.bak" counted as showing "a.docx".
  assert.equal(addressShows(urlFor('word', `${WIN}.bak`), WIN, 'win32'), false);
  assert.equal(addressShows(urlFor('word', null), WIN, 'win32'), false);
  assert.equal(addressShows('not an address', WIN, 'win32'), false);
  assert.equal(addressShows(urlFor('word', WIN), '', 'win32'), false);
});

test('the Locker\'s own window is known by its app parameter alone', () => {
  assert.equal(addressApp(urlFor('locker', null)), 'locker');
  assert.equal(addressApp(urlFor('word', 'C:\\locker\\app=locker.docx')), 'word');
  assert.equal(addressApp('about:blank'), null);
});
