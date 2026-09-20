// A picture taller than the page's inside is printed to fit it.
//
// A scanned card in a letter is taller than it is wide; scaled to the
// column's width it can still be taller than a sheet, and print placed it
// whole at the head of a page and let the sheet's edge cut it. The screen
// draws such a picture to fit the page; print does the same.

import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';
import { gradientPng, joinPictureParagraphs } from '../apps/desktop/main/sample-picture.js';

test('a picture taller than the page prints to fit it, proportions kept; one that fits is left alone', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'A letter' }, { text: 'With a scan.' }, { text: 'And a photograph.' }, { text: 'After.' }] }));
  view.setSelection({ block: 1, offset: 0 });
  view.insertImage({ name: 'scan', contentType: 'image/png', data: gradientPng(30, 70, [50, 90, 200], [250, 200, 60]), widthPx: 600, heightPx: 1400 });
  const blocks = view.render({ pages: false }).blocks;
  const scanBlock = blocks.findIndex((b) => (b.images || []).length);
  view.setSelection({ block: scanBlock + 2, offset: 0 });
  view.insertImage({ name: 'photo', contentType: 'image/png', data: gradientPng(30, 30, [60, 160, 90], [200, 50, 50]), widthPx: 300, heightPx: 300 });

  const laid = view.pages;
  const inside = laid.contentHeightPx;
  const drawn = laid.pages.flatMap((p) => p.fragments.filter((f) => f.kind === 'images')).flatMap((f) => f.images);
  assert.equal(drawn.length, 2, 'both pictures are placed as blocks');
  const [scan, photo] = drawn;
  assert.ok(scan.heightPx <= inside, `the scan fits the page's inside: ${scan.heightPx} of ${inside}`);
  assert.ok(scan.heightPx > inside * 0.9, 'and is as tall as the page allows');
  assert.ok(Math.abs(scan.widthPx / scan.heightPx - 600 / 1400) < 0.01, 'its proportions are kept');
  assert.equal(photo.widthPx, 300, 'a picture that fits is drawn at its own size');
  assert.equal(photo.heightPx, 300);
  for (const page of laid.pages) {
    for (const f of page.fragments.filter((x) => x.kind === 'images')) {
      const cost = f.images.reduce((t, i) => t + i.heightPx, 0);
      assert.ok(cost <= inside, 'no picture block is taller than the page');
    }
  }
});

test('a paragraph of four wide pictures, one to a row, prints as many rows to a page as fit, whole, and the rest on the next', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Cards' }] }));
  view.setSelection({ block: 0, offset: 0 });
  for (let i = 0; i < 4; i++) {
    view.insertImage({ name: `card ${i + 1}`, contentType: 'image/png', data: gradientPng(60, 40, [40, 120, 200], [230, 200, 60]), widthPx: 600, heightPx: 400 });
  }
  const joined = openDocx(joinPictureParagraphs(view.save()));
  const blocks = joined.render({ pages: false }).blocks;
  assert.equal(blocks.filter((b) => (b.images || []).length).length, 1, 'one paragraph holds the pictures');
  assert.equal(blocks.find((b) => (b.images || []).length).images.length, 4);

  const laid = joined.pages;
  const inside = laid.contentHeightPx;
  const perPage = laid.pages.map((p) => p.fragments.filter((f) => f.kind === 'images').flatMap((f) => f.images).length);
  assert.equal(perPage.reduce((a, b) => a + b, 0), 4, 'every picture is placed once');
  // The heading's line and two rows with their gaps come to 838 px of the
  // page's 930; a third would not fit, so two rows to a page.
  assert.deepEqual(perPage, [2, 2], `two to a page: ${JSON.stringify(perPage)}`);
  for (const page of laid.pages) {
    const used = page.fragments.filter((f) => f.kind === 'images').flatMap((f) => f.images).reduce((t, i) => t + i.heightPx, 0);
    assert.ok(used <= inside, 'no page holds more picture than fits');
    for (const img of page.fragments.filter((f) => f.kind === 'images').flatMap((f) => f.images)) assert.equal(img.heightPx, 400, 'a picture that fits is not scaled');
  }
});

test('a paragraph of only pictures prints them in rows, as many to a row as fit the column, and a paragraph with words one to a row', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Cards' }] }));
  view.setSelection({ block: 0, offset: 0 });
  for (let i = 0; i < 4; i++) {
    view.insertImage({ name: `card ${i + 1}`, contentType: 'image/png', data: gradientPng(28, 45, [40, 120, 200], [230, 200, 60]), widthPx: 280, heightPx: 450 });
  }
  const joined = openDocx(joinPictureParagraphs(view.save()));
  const laid = joined.pages;
  const fragments = laid.pages.flatMap((p) => p.fragments.filter((f) => f.kind === 'images'));
  const rows = fragments.flatMap((f) => f.rows);
  assert.deepEqual(rows.map((r) => r.count), [2, 2], 'two to a row: the column is 602 px wide and two cards are 560');
  assert.equal(rows[0].heightPx, 450, 'a row is as tall as its tallest picture');
  assert.equal(fragments.flatMap((f) => f.images).length, 4, 'every card is placed once');
  // The heading's line and two rows with their gaps come to more than the
  // page's 930 px, so the second row heads the next page, whole.
  assert.equal(laid.pages.length, 2);
  assert.deepEqual(laid.pages.map((p) => p.fragments.filter((f) => f.kind === 'images').flatMap((f) => f.rows).length), [1, 1], 'one row a page');
  assert.equal(fragments[0].align, null, 'a plain paragraph: the row starts at the left');

  // The same four cards under words stay one to a row, as blocks under the text.
  const worded = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Cards' }] }));
  worded.setSelection({ block: 0, offset: 0 });
  worded.insertImage({ name: 'card', contentType: 'image/png', data: gradientPng(28, 45, [40, 120, 200], [230, 200, 60]), widthPx: 280, heightPx: 450 });
  worded.insertImage({ name: 'card', contentType: 'image/png', data: gradientPng(28, 45, [40, 120, 200], [230, 200, 60]), widthPx: 280, heightPx: 450 });
  const stacked = worded.pages.pages.flatMap((p) => p.fragments.filter((f) => f.kind === 'images'));
  for (const f of stacked) for (const r of f.rows) assert.equal(r.count, 1, 'one picture a row');
});
