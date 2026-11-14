// A ribbon button drawn in its short version is named by the words drawn.
//
// The short words come from data-label through the theme (7ca18007), so the
// button's text stays its medium label; in a short group that text is
// hidden, and the accessibility audit counted no name at all: 05.0012 to
// 05.0014 with the English copy, "button-name" on every button of a
// ".short.compact.rw-group". WCAG 2.5.3 (Label in Name) wants the name to
// hold the words seen, so while the short version is drawn the button
// carries them as its aria-label, and is described by its full words; drawn
// medium again, or with no short version left, the label goes and its own
// text names it. A label from anywhere else is never touched.
// Session-proposed; the review of d5305ff7 asked for the stale-name cases.
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

/* ── just enough of a page: elements, their attributes and dataset, text, and the selectors the code uses ── */

const kebab = (k) => k.replace(/[A-Z]/g, (c) => `-${c.toLowerCase()}`);

const doc = {
  createTreeWalker(root) {
    const texts = [];
    const walk = (n) => { for (const c of n.childNodes) (c.nodeType === 3 ? texts.push(c) : walk(c)); };
    walk(root);
    let i = 0;
    return { nextNode: () => texts[i++] || null };
  },
};

function el(tag, { cls = '', attrs = {}, kids = [] } = {}) {
  const attributes = new Map(Object.entries(attrs));
  const node = {
    nodeType: 1,
    tagName: tag.toUpperCase(),
    classList: new Set(cls.split(' ').filter(Boolean)),
    childNodes: [],
    parentElement: null,
    ownerDocument: doc,
    sets: 0,
    getAttribute: (k) => (attributes.has(k) ? attributes.get(k) : null),
    setAttribute: (k, v) => { node.sets += 1; attributes.set(k, String(v)); },
    removeAttribute: (k) => attributes.delete(k),
    hasAttribute: (k) => attributes.has(k),
    closest: () => null,
    append(...more) { for (const c of more) { c.parentElement = node; node.childNodes.push(c); } return node; },
    descendants() { const all = []; const walk = (n) => { for (const c of n.childNodes) if (c.nodeType === 1) { all.push(c); walk(c); } }; walk(node); return all; },
    querySelectorAll(sel) {
      const match = {
        '.rw-btn > .rw-short': (n) => n.classList.has('rw-short') && n.parentElement?.classList.has('rw-btn'),
        '.rw-btn[data-named-short]': (n) => n.classList.has('rw-btn') && n.hasAttribute('data-named-short'),
        '[data-label]': (n) => n.hasAttribute('data-label'),
      }[sel];
      if (!match) throw new Error(`selector not in the fake page: ${sel}`);
      return node.descendants().filter(match);
    },
    querySelector(sel) {
      if (sel === ':scope > .rw-short') return node.childNodes.find((c) => c.nodeType === 1 && c.classList.has('rw-short')) || null;
      if (sel === 'img[alt]:not([alt=""]), svg > title') return null;
      throw new Error(`selector not in the fake page: ${sel}`);
    },
  };
  node.dataset = new Proxy({}, {
    get: (_, k) => node.getAttribute(`data-${kebab(String(k))}`) ?? undefined,
    set: (_, k, v) => { node.setAttribute(`data-${kebab(String(k))}`, v); return true; },
    has: (_, k) => node.hasAttribute(`data-${kebab(String(k))}`),
    deleteProperty: (_, k) => { node.removeAttribute(`data-${kebab(String(k))}`); return true; },
  });
  node.append(...kids);
  return node;
}
const text = (data) => ({ nodeType: 3, data, parentElement: null });

/** A ribbon button as the Button draws it: its medium words, and its short version in data-label. */
function button({ medium, short, tip, label } = {}) {
  const b = el('button', { cls: 'rw-btn', attrs: { 'data-tip': tip ?? medium, ...(label ? { 'aria-label': label } : {}) } });
  const m = el('span', { cls: 'rw-medium' });
  m.append(text(medium));
  m.childNodes[0].parentElement = m;
  b.append(m);
  if (short) b.append(el('span', { cls: 'rw-short', attrs: { 'data-label': short } }));
  return b;
}
const page = (...buttons) => el('div', { cls: 'rw-groups' }).append(...buttons);
const shortSpan = (b) => b.querySelector(':scope > .rw-short');

/** The theme: the short spans in `drawn` shown, every other one display: none. */
function drawing(drawn, fn) {
  const was = globalThis.getComputedStyle;
  globalThis.getComputedStyle = (n) => ({ display: drawn.has(n) ? 'inline' : 'none' });
  try { return fn(); } finally { globalThis.getComputedStyle = was; }
}

test('drawn short, a button is named by its short words and described by its full ones; drawn medium, both go back', async () => {
  const { nameShortButtons } = await kit();
  const b = button({ medium: 'Conditional Formatting', short: 'Cond. Format', tip: 'Conditional Formatting: rules that colour cells' });
  const root = page(b);
  drawing(new Set([shortSpan(b)]), () => nameShortButtons(root));
  assert.equal(b.getAttribute('aria-label'), 'Cond. Format');
  assert.equal(b.getAttribute('aria-description'), 'Conditional Formatting: rules that colour cells');

  drawing(new Set(), () => nameShortButtons(root));
  assert.equal(b.getAttribute('aria-label'), null);
  assert.equal(b.hasAttribute('data-named-short'), false);
  // Its own words name it again; the description is the rest of the tip, as the window's naming gives it.
  assert.equal(b.getAttribute('aria-description'), 'rules that colour cells');
});

test('a large button split round its arrow, and one on two lines, are named by the words as one line', async () => {
  const { nameShortButtons } = await kit();
  const split = button({ medium: 'Custom Slide Show', short: 'Custom ' });
  shortSpan(split).append(el('span', { cls: 'rw-btn-last', attrs: { 'data-label': 'Show' } }));
  const two = button({ medium: 'Two Words Here', short: 'Two Short\nWords' });
  const root = page(split, two);
  drawing(new Set([shortSpan(split), shortSpan(two)]), () => nameShortButtons(root));
  assert.deepEqual([split.getAttribute('aria-label'), two.getAttribute('aria-label')], ['Custom Show', 'Two Short Words']);
});

test('a short version removed while the button is named short takes its name with it (a language with none)', async () => {
  const { nameShortButtons } = await kit();
  const b = button({ medium: 'Conditional Formatting', short: 'Cond. Format' });
  const root = page(b);
  drawing(new Set([shortSpan(b)]), () => nameShortButtons(root));
  assert.equal(b.getAttribute('aria-label'), 'Cond. Format');
  // The new language has no short version: React drops the span, and leaves the attribute it never set.
  b.childNodes.splice(b.childNodes.indexOf(shortSpan(b)), 1);
  drawing(new Set(), () => nameShortButtons(root));
  assert.equal(b.getAttribute('aria-label'), null);
  assert.equal(b.hasAttribute('data-named-short'), false);
});

test('new short words in data-label rename the button, and an unchanged pass writes nothing', async () => {
  const { nameShortButtons } = await kit();
  const b = button({ medium: 'Conditional Formatting', short: 'Cond. Format' });
  const root = page(b);
  const drawn = new Set([shortSpan(b)]);
  drawing(drawn, () => nameShortButtons(root));
  const before = b.sets;
  drawing(drawn, () => nameShortButtons(root));
  assert.equal(b.sets, before, 'a second pass with nothing changed sets no attribute');
  shortSpan(b).setAttribute('data-label', 'Rules');
  drawing(drawn, () => nameShortButtons(root));
  assert.equal(b.getAttribute('aria-label'), 'Rules');
});

test("a caller's own label is kept, drawn short and drawn medium", async () => {
  const { nameShortButtons } = await kit();
  const b = button({ medium: 'Track Changes', short: 'Tracking', label: 'Track Changes, on' });
  const root = page(b);
  drawing(new Set([shortSpan(b)]), () => nameShortButtons(root));
  assert.equal(b.getAttribute('aria-label'), 'Track Changes, on');
  assert.equal(b.hasAttribute('data-named-short'), false);
  drawing(new Set(), () => nameShortButtons(root));
  assert.equal(b.getAttribute('aria-label'), 'Track Changes, on');
});

test('a label the Button sets after the short version is gone is not the one taken away', async () => {
  const { nameShortButtons } = await kit();
  const b = button({ medium: 'Track Changes', short: 'Tracking' });
  const root = page(b);
  drawing(new Set([shortSpan(b)]), () => nameShortButtons(root));
  // A language with no short version but a complete one the label is part of: the Button names it so.
  b.childNodes.splice(b.childNodes.indexOf(shortSpan(b)), 1);
  b.setAttribute('aria-label', 'Turn Track Changes on or off');
  drawing(new Set(), () => nameShortButtons(root));
  assert.equal(b.getAttribute('aria-label'), 'Turn Track Changes on or off');
  assert.equal(b.hasAttribute('data-named-short'), false);
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

test('the ribbon and the status bar name their buttons after every fitting, the status bar with no short version left too', () => {
  const src = fs.readFileSync(path.join(here, '..', 'packages', 'office-ui', 'src', 'index.js'), 'utf8');
  assert.match(src, /function fitGroups\(el, iconsOnly = false\) \{\s*fitStages\(el, iconsOnly\);\s*nameShortButtons\(el\);\s*\}/);
  assert.match(src, /if \(!el\.querySelector\('\.rw-short'\)\) \{[\s\S]{0,160}nameShortButtons\(el\);\s*return;/);
  assert.match(src, /if \(el\.scrollWidth > el\.clientWidth \+ 1\) el\.classList\.add\('short'\);\s*nameShortButtons\(el\);/);
  // A change to the short words alone fits the ribbon again.
  assert.match(src, /mo\?\.observe\(el, \{[^}]*attributeFilter: \['data-label'\] \}\)/);
});
