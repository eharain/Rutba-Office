// A ribbon button whose label has a short version (L3) carries both, and the
// theme draws the short one only in a group or status bar too narrow for the
// medium. The short one's words come from data-label, drawn by the theme, so
// the button's text in the page is its medium label alone, whichever version
// is drawn: with lang-en present, "Track Changes" reads "Track Changes", not
// "Track ChangesTracking" (part run 2 of 14 Nov 2026, 07.0250, 07.0251 and
// the Custom Slide Show flow; the owner, D-2026-11-14-07: English shows its
// full wording, the short one only where space is tight).
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import esbuild from 'esbuild';
import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';

const here = path.dirname(fileURLToPath(import.meta.url));

/** The UI kit bundled for node, React shared with this test, and its messages module so versions can be set. */
async function kit() {
  const out = await esbuild.build({
    stdin: {
      contents: "export { Button } from './packages/office-ui/src/index.js'; export { setLanguage, setFallback, setVersions, t } from './packages/office-ui/src/messages.js';",
      resolveDir: path.join(here, '..'),
      loader: 'js',
    },
    bundle: true,
    format: 'esm',
    platform: 'node',
    write: false,
    logLevel: 'silent',
    loader: { '.js': 'jsx', '.css': 'empty' },
    jsx: 'automatic',
    external: ['react', 'react-dom', 'react/jsx-runtime'],
  });
  const file = path.join(here, '..', 'node_modules', `.button-short-label-${process.pid}.mjs`);
  fs.writeFileSync(file, out.outputFiles[0].text);
  try { return await import(`file://${file.replace(/\\/g, '/')}`); } finally { fs.rmSync(file, { force: true }); }
}

const textOf = (html) => html.replace(/<[^>]*>/g, '');

test('a button with a short version has the medium label as its only text, the short one in data-label', async () => {
  const { Button, setLanguage, setFallback, setVersions, t } = await kit();
  setLanguage('en');
  setFallback(null);
  const english = {
    'Track Changes': { short: 'Tracking', complete: 'Turn Track Changes on or off' },
    'Read Aloud': { short: 'Read' },
    'Custom Slide Show': { short: 'Custom Show' },
    'Two Words Here': { short: 'Two Short Words' },
  };
  // An English window: lang-en's versions are its own, and the fallback's for a label it lacks.
  setVersions(english, english);
  try {
    const html = (props) => renderToStaticMarkup(React.createElement(Button, props));

    const small = html({ label: t('Track Changes') });
    assert.equal(textOf(small), 'Track Changes');
    assert.match(small, /<span class="rw-medium">Track Changes<\/span><span class="rw-short" data-label="Tracking"><\/span>/);
    assert.match(small, /data-tip="Turn Track Changes on or off"/);

    // A large button with a menu: the last short word kept with the arrow, as withCaret keeps the medium's.
    const menu = html({ label: t('Custom Slide Show'), tall: true, dropdown: true });
    assert.equal(textOf(menu), 'Custom Slide Show');
    assert.match(menu, /<span class="rw-short" data-label="Custom "><span class="rw-btn-last" data-label="Show"><svg/);
    const oneWord = html({ label: t('Read Aloud'), tall: true, dropdown: true });
    assert.equal(textOf(oneWord), 'Read Aloud');
    assert.match(oneWord, /<span class="rw-short" data-label="Read"><svg/);

    // A large button on two lines: the short words split the same way, as two lines of data-label.
    const two = html({ label: t('Two Words Here'), tall: true, twoLine: true });
    assert.equal(textOf(two), 'Two Words Here');
    assert.match(two, /data-label="Two Short\nWords"|data-label="Two\nShort Words"/);

    // A label with no short version has no short span at all.
    const plain = html({ label: t('Next') });
    assert.equal(textOf(plain), 'Next');
    assert.doesNotMatch(plain, /rw-short/);
  } finally {
    setVersions(null, null);
  }
});

test('the theme draws a short label from data-label, and only where its group is short', () => {
  const css = fs.readFileSync(path.join(here, '..', 'packages', 'office-ui', 'src', 'theme.css'), 'utf8');
  assert.match(css, /\.rw-btn > \.rw-short::before, \.rw-btn > \.rw-short > \.rw-btn-last::before \{ content: attr\(data-label\); \}/);
  // No white-space of its own in general: the short words keep the button's nowrap, as their text did.
  assert.doesNotMatch(css, /\.rw-btn > \.rw-short::before[^{]*\{[^}]*pre-line/);
  assert.match(css, /\.rw-btn\.tall\.two-line > \.rw-short::before \{ white-space: pre; \}/);
  assert.match(css, /:is\(\.rw-group\.compact, \.rw-group\.small\) \.rw-btn\.two-line > \.rw-short::before \{ white-space: nowrap; \}/);
  assert.match(css, /\.rw-btn > \.rw-short \{ display: none; \}/);
  assert.match(css, /:is\(\.rw-group\.short, \.rw-status\.short\) \.rw-btn > \.rw-short \{ display: inline; \}/);
  assert.match(css, /:is\(\.rw-group\.short, \.rw-status\.short\) \.rw-btn > \.rw-medium \{ display: none; \}/);
});
