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
import { gradientPng } from '../apps/desktop/main/sample-picture.js';

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
