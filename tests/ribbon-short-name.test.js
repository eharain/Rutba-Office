// A ribbon button drawn in its short version is named by the words drawn.
//
// The short words come from data-label through the theme (7ca18007), so the
// button's text stays its medium label; in a short group that text is
// hidden, and the accessibility audit counted no name at all: 05.0012 to
// 05.0014 with the English copy, "button-name" on every button of a
// ".short.compact.rw-group". WCAG 2.5.3 (Label in Name) wants the name to
// hold the words seen, so while the short version is drawn the button
// carries them as its aria-label; drawn medium again, the label goes and its
// own text names it. Session-proposed.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import esbuild from 'esbuild';
import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';

const here = path.dirname(fileURLToPath(import.meta.url));

async function kit() {
  const out = await esbuild.build({
    stdin: {
      contents: "export { Button, nameShortButtons } from './packages/office-ui/src/index.js'; export { setLanguage, setFallback, setVersions, t } from './packages/office-ui/src/messages.js';",
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
  const file = path.join(here, '..', 'node_modules', `.ribbon-short-name-${process.pid}.mjs`);
  fs.writeFileSync(file, out.outputFiles[0].text);
  try { return await import(`file://${file.replace(/\\/g, '/')}`); } finally { fs.rmSync(file, { force: true }); }
}

/** Just enough of an element for nameShortButtons: attributes, dataset, children by data-label. */
function element({ label = null, children = [] } = {}) {
  const attrs = new Map();
  const el = {
    dataset: {},
    parentElement: null,
    children,
    getAttribute: (k) => (attrs.has(k) ? attrs.get(k) : null),
    setAttribute: (k, v) => attrs.set(k, String(v)),
    removeAttribute: (k) => attrs.delete(k),
    hasAttribute: (k) => attrs.has(k),
    querySelectorAll: () => children.filter((c) => c.dataset.label != null),
  };
  if (label != null) el.dataset.label = label;
  for (const c of children) c.parentElement = el;
  return el;
}

/** A ribbon of buttons, each with its short span; `drawn` says which short spans the theme shows. */
function ribbon(buttons) {
  const shorts = buttons.map(({ parts }) => {
    const [first, ...rest] = parts;
    return element({ label: first, children: rest.map((p) => element({ label: p })) });
  });
  const btns = shorts.map((s) => {
    const b = element({ children: [s] });
    return b;
  });
  return { el: { querySelectorAll: () => shorts }, btns, shorts };
}

function withStyle(drawn, fn) {
  const was = globalThis.getComputedStyle;
  globalThis.getComputedStyle = (n) => ({ display: drawn.has(n) ? 'inline' : 'none' });
  try { return fn(); } finally { globalThis.getComputedStyle = was; }
}

test('a button drawn short is named by the short words, and loses the name when drawn medium again', async () => {
  const { nameShortButtons } = await kit();
  const r = ribbon([{ parts: ['Cond. Format'] }, { parts: ['Custom ', 'Show'] }, { parts: ['Two Short\nWords'] }]);
  const drawn = new Set(r.shorts);
  withStyle(drawn, () => nameShortButtons(r.el));
  assert.deepEqual(r.btns.map((b) => b.getAttribute('aria-label')), ['Cond. Format', 'Custom Show', 'Two Short Words']);

  // The group has room again: the medium is drawn, and its own text names the button.
  withStyle(new Set(), () => nameShortButtons(r.el));
  assert.deepEqual(r.btns.map((b) => b.getAttribute('aria-label')), [null, null, null]);
});

test('a label the button had from elsewhere is never taken away', async () => {
  const { nameShortButtons } = await kit();
  const r = ribbon([{ parts: ['Tracking'] }]);
  r.btns[0].setAttribute('aria-label', 'Given by its app');
  withStyle(new Set(), () => nameShortButtons(r.el));
  assert.equal(r.btns[0].getAttribute('aria-label'), 'Given by its app');
});

test('the Button gives no aria-label of its own where it has a short version, and its text stays the medium', async () => {
  const { Button, setLanguage, setFallback, setVersions, t } = await kit();
  setLanguage('en');
  setFallback(null);
  const english = { 'Conditional Formatting': { short: 'Cond. Format', complete: 'Conditional Formatting: rules that colour cells' } };
  setVersions(english, english);
  try {
    const html = renderToStaticMarkup(React.createElement(Button, { label: t('Conditional Formatting'), tall: true, dropdown: true }));
    assert.doesNotMatch(html, /aria-label=/);
    assert.equal(html.replace(/<[^>]*>/g, ''), 'Conditional Formatting');
  } finally {
    setVersions(null, null);
  }
});

test('the ribbon and the status bar name their buttons after every fitting', () => {
  const src = fs.readFileSync(path.join(here, '..', 'packages', 'office-ui', 'src', 'index.js'), 'utf8');
  assert.match(src, /function fitGroups\(el, iconsOnly = false\) \{\s*fitStages\(el, iconsOnly\);\s*nameShortButtons\(el\);\s*\}/);
  assert.match(src, /if \(el\.scrollWidth > el\.clientWidth \+ 1\) el\.classList\.add\('short'\);\s*nameShortButtons\(el\);/);
});
