// Every labelled element the UI kit draws carries its label's key in two
// attributes (D-2026-11-14-11, stage 3 of the meaning keys): data-key, the
// key's name, and data-scope, where it applies (never shown). They come from
// the key t() made the words from, so no call changes, and a check or a screen
// reader can find the element by what it means rather than by its words.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import esbuild from 'esbuild';
import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';

const here = path.dirname(fileURLToPath(import.meta.url));

/**
 * React for the kit, with its layout effects recorded rather than dropped (a
 * page drawn here never commits): `commit()` runs the kit's endDraw that the
 * draws since the last commit asked for, as React runs layout effects once a
 * draw is committed, after every element has been drawn.
 */
const COMMITS = `import * as R from 'react';
export * from 'react';
export default R.default || R;
const asked = [];
export function useLayoutEffect(fn) { asked.push(fn); }
export function commit(only) { const run = asked.splice(0).filter((fn) => fn === only); run.slice(0, 1).forEach((fn) => fn()); return run.length; }`;
const withCommits = {
  name: 'commits',
  setup(build) {
    build.onResolve({ filter: /^react$/ }, (a) => (a.namespace === 'commits' ? { path: 'react', external: true } : { path: 'react', namespace: 'commits' }));
    build.onLoad({ filter: /.*/, namespace: 'commits' }, () => ({ contents: COMMITS, loader: 'js', resolveDir: path.join(here, '..') }));
  },
};

let bundles = 0;
/** The UI kit bundled for node, React shared with this test, with its messages module; `{ commits: true }`: with commit() (COMMITS). */
async function kit({ commits = false } = {}) {
  const out = await esbuild.build({
    stdin: {
      contents: "export { Button, Item, Group, Field, Dialog, Empty, Panel, Menu, Ribbon, keyAttrs } from './packages/office-ui/src/index.js'; export { setLanguage, registerCatalogue, t, tn, setBare, endDraw } from './packages/office-ui/src/messages.js';"
        + (commits ? " export { commit } from 'react';" : ''),
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
    external: commits ? ['react-dom', 'react/jsx-runtime'] : ['react', 'react-dom', 'react/jsx-runtime'],
    plugins: commits ? [withCommits] : [],
  });
  // A name of its own for each bundle: import() keeps a module by its address, and one with commit() is another bundle.
  const file = path.join(here, '..', 'node_modules', `.label-key-attrs-${process.pid}-${(bundles += 1)}.mjs`);
  fs.writeFileSync(file, out.outputFiles[0].text);
  try { return await import(`file://${file.replace(/\\/g, '/')}`); } finally { fs.rmSync(file, { force: true }); }
}

const html = (C, props) => renderToStaticMarkup(React.createElement(C, props));

test("a button, a list item, a group, a field, a dialog, a panel and an empty page carry their label's key and scope", async () => {
  const { Button, Item, Group, Field, Dialog, Empty, Panel, setLanguage, t, tn } = await kit();
  setLanguage('en');
  assert.match(html(Button, { label: t('common.save') }), /<button[^>]* data-key="save" data-scope="common"[^>]*>/);
  assert.match(html(Button, { label: t('Close') }), /data-key="close" data-scope="common"/, 'a call by English too');
  assert.match(html(Item, { label: tn(3, 'common.count-word') }), /<button[^>]* data-key="count-word" data-scope="common"/);
  assert.match(html(Group, { label: t('ui.proofing.errors') }), /<div class="rw-group-label" data-key="errors" data-scope="ui\.proofing">Errors<\/div>/);
  assert.match(html(Field, { label: t('common.save') }), /<label id="[^"]*" data-key="save" data-scope="common">Save<\/label>/);
  assert.match(html(Dialog, { title: t('common.close') }), /<div class="rw-dialog-head" data-key="close" data-scope="common">Close<\/div>/);
  assert.match(html(Empty, { title: t('common.save') }), /<h3 data-key="save" data-scope="common">Save<\/h3>/);
  assert.match(html(Panel, { title: t('common.close') }), /<span style="flex:1" data-key="close" data-scope="common">Close<\/span>/);
});

test('in another language the words drawn still carry the key; words no key gave carry none', async () => {
  const { Button, Group, registerCatalogue, setLanguage, t, keyAttrs } = await kit();
  registerCatalogue('fr', { Save: 'Enregistrer' });
  try {
    setLanguage('fr');
    assert.match(html(Button, { label: t('common.save') }), /data-key="save" data-scope="common"[^>]*>.*Enregistrer/);
  } finally {
    setLanguage('en');
  }
  const plain = html(Button, { label: 'Words no key gave' });
  assert.doesNotMatch(plain, /data-key|data-scope/);
  assert.doesNotMatch(html(Group, { label: 'Nor here' }), /data-key/);
  assert.equal(keyAttrs(''), null);
  assert.equal(keyAttrs(null), null);
  assert.doesNotMatch(html(Button, { icon: 'save', title: 'An icon alone' }), /data-key/);
});

test('a menu item, a submenu and a menu heading carry their keys', async () => {
  const { Menu, t } = await kit();
  const menu = html(Menu, {
    x: 0,
    y: 0,
    onClose() {},
    items: [
      { heading: true, label: t('ui.proofing.errors') },
      { label: t('common.save'), run() {} },
      { label: t('common.close'), items: [{ label: t('Close'), run() {} }] },
    ],
  });
  assert.match(menu, /class="rw-menu-head" role="presentation" data-key="errors" data-scope="ui\.proofing"/);
  assert.match(menu, /role="menuitem"[^>]* data-key="save" data-scope="common"/);
  assert.match(menu, /aria-haspopup="menu"[^>]* data-key="close" data-scope="common"/);
});

test("a ribbon's tabs carry their keys", async () => {
  const { Ribbon, t } = await kit();
  const ribbon = html(Ribbon, { tabs: [{ id: 'a', label: t('common.save') }, { id: 'b', label: t('common.close') }], active: 'a', onTab() {} });
  assert.match(ribbon, /role="tab"[^>]* data-key="save" data-scope="common"/);
  assert.match(ribbon, /role="tab"[^>]* data-key="close" data-scope="common"/);
});

test('with no pack wording a label, the element shows its name and has the full key as its tip (D-2026-11-14-16 (2))', async () => {
  const { Button, Item, Group, Menu, setLanguage, setBare, t, tn } = await kit();
  setLanguage('en');
  setBare(true);
  try {
    // A button's tip is its own (data-tip): the full key, never a native title as well.
    const save = html(Button, { label: t('common.save') });
    assert.match(save, /data-tip="common\.save"/);
    assert.match(save, /data-key="save" data-scope="common"/);
    assert.match(save, />save<\/span>/);
    assert.doesNotMatch(save, / title=/);
    // A tip that is a label of its own, shown bare: its full key.
    const change = html(Button, { label: t('common.change'), title: t('ui.proofing.change-this-one') });
    assert.match(change, /data-tip="ui\.proofing\.change-this-one"/);
    assert.match(change, /data-tip-key="change-this-one"/);
    // An icon button named by its tip alone.
    assert.match(html(Button, { icon: 'save', title: t('common.save') }), /data-tip="common\.save"/);
    // A tip no key gave stays as it is.
    assert.match(html(Button, { label: t('common.change'), title: 'Report.docx' }), /data-tip="Report\.docx"/);
    // The other elements the code draws: the full key as their title.
    assert.match(html(Group, { label: t('ui.proofing.errors') }), /<div class="rw-group-label" data-key="errors" data-scope="ui\.proofing" title="ui\.proofing\.errors">errors<\/div>/);
    // A list item and a menu item may be the person's own words: their name, with no key and no full key.
    const item = html(Item, { label: tn(3, 'common.count-word') });
    assert.match(item, />count word</);
    assert.doesNotMatch(item, /data-key|title="common\.count-word"/);
    const menu = html(Menu, { x: 0, y: 0, onClose() {}, items: [{ label: t('common.close'), run() {} }] });
    assert.match(menu, />close</);
    assert.doesNotMatch(menu, /data-key|title="common\.close"/);
  } finally {
    setBare(null);
  }
  // With the English pack wording it, no title is added.
  assert.doesNotMatch(html(Group, { label: t('ui.proofing.errors') }), /title=/);
  assert.match(html(Button, { label: t('common.save') }), /data-tip="Save"/);
});

test('with no pack, two keys with the same name keep their own keys, each element its own call; words no call gave are never marked (D-2026-11-14-16 (2))', async () => {
  const { Button, Item, Menu, setLanguage, setBare, t } = await kit();
  setLanguage('en');
  setBare(true);
  try {
    // common.save and profiles.save both show "save": each button has its own key, as its call gave it.
    const [first, second] = renderToStaticMarkup(React.createElement('div', null,
      React.createElement(Button, { label: t('common.save') }),
      React.createElement(Button, { label: t('profiles.save') }))).split('</button>');
    assert.match(first, /data-tip="common\.save"/);
    assert.match(first, /data-key="save" data-scope="common"/);
    assert.match(second, /data-tip="profiles\.save"/);
    assert.match(second, /data-key="save" data-scope="profiles"/);
    // A file named "save" in a list is no label, though a call gave "save" in this task and no
    // element has taken it: a list's or a menu's item never takes a call's key (round 3 (3)).
    t('common.save');
    const file = html(Item, { label: 'save' });
    assert.doesNotMatch(file, /data-key|data-scope/);
    assert.match(file, /title="save"/);
    assert.doesNotMatch(html(Menu, { x: 0, y: 0, onClose() {}, items: [{ label: 'save', run() {} }, { label: t('common.close'), run() {} }] }), /data-key|title="[a-z.]*(save|close)"/);
    // The call the file did not take is still the button's that the code draws with it.
    assert.match(html(Button, { label: 'save' }), /data-tip="common\.save"/);
    // Nor once the task the name was shown in is over.
    t('common.save');
    await new Promise((resolve) => setTimeout(resolve, 5));
    assert.doesNotMatch(html(Item, { label: 'save' }), /data-key|title="common\.save"/);
    assert.doesNotMatch(html(Button, { label: 'save' }), /data-key|common\.save/);
  } finally {
    setBare(null);
  }
});

test('with no pack, the limits of a key taken from its call: two same-named elements drawn in the other order swap keys, never words; a call from an earlier task gives no key (D-2026-11-14-16 (2))', async () => {
  const { Button, setLanguage, setBare, t } = await kit();
  setLanguage('en');
  setBare(true);
  try {
    // The calls in one order, the buttons drawn in the other: each takes the oldest call it finds.
    const profiles = t('profiles.save');
    const common = t('common.save');
    const [first, second] = renderToStaticMarkup(React.createElement('div', null,
      React.createElement(Button, { label: common }),
      React.createElement(Button, { label: profiles }))).split('</button>');
    // What is shown is right whatever the order: both are the words "save", as each call gave them.
    assert.equal(common, 'save');
    assert.equal(profiles, 'save');
    assert.match(first, />save<\/span>/);
    assert.match(second, />save<\/span>/);
    // Only the key, and the tip that shows it, can be the other's: the limit, kept here.
    assert.match(first, /data-tip="profiles\.save"/);
    assert.match(second, /data-tip="common\.save"/);
    // Each key is given once: no element takes a key no call gave, nor one twice.
    assert.deepEqual([first, second].map((b) => b.match(/data-scope="([a-z.]+)"/)[1]).sort(), ['common', 'profiles']);

    // A button whose words are the person's own, the same as a bare name, drawn in the same draw
    // as a call of that name no element took (four step names made, one shown): it takes the key.
    // The limit, kept here: its words are right, its key and tip are the call's.
    t('common.save');
    const own = renderToStaticMarkup(React.createElement('div', null, React.createElement(Button, { label: 'save' })));
    assert.match(own, />save<\/span>/);
    assert.match(own, /data-tip="common\.save"/);

    // A label set from an IPC answer: its call ran in an earlier task. It shows its name, and
    // has no key and no full key as its tip, its words as they are. Accepted: with no pack the
    // words are still the key's name, and the next draw that calls t() again marks it.
    const later = t('common.save');
    await new Promise((resolve) => setTimeout(resolve, 5));
    const drawn = html(Button, { label: later });
    assert.match(drawn, />save<\/span>/);
    assert.match(drawn, /data-tip="save"/);
    assert.doesNotMatch(drawn, /data-key|common\.save/);
  } finally {
    setBare(null);
  }
});

test('with no pack, a list item "save" beside a button the code labels save never takes the key, in the same draw, before it or after it', async () => {
  const { Button, Item, setLanguage, setBare, t } = await kit();
  setLanguage('en');
  setBare(true);
  try {
    for (const itemFirst of [true, false]) {
      const label = t('Save');
      const item = React.createElement(Item, { key: 'i', label: 'save' });
      const button = React.createElement(Button, { key: 'b', label });
      const page = renderToStaticMarkup(React.createElement('div', null, ...(itemFirst ? [item, button] : [button, item])));
      const [drawnItem] = page.match(/<button[^>]*class="rw-item[^"]*"[^>]*>/);
      assert.doesNotMatch(drawnItem, /data-key|data-scope|common\.save/, drawnItem);
      assert.match(page, /data-tip="common\.save"/, 'the button keeps its call');
      assert.match(page, /data-key="save" data-scope="common"/);
    }
  } finally {
    setBare(null);
  }
});

test('with no pack, a call no element took is dropped once the draw is committed: a later draw in the same task never takes it', async () => {
  const { Button, setLanguage, setBare, t, endDraw, commit } = await kit({ commits: true });
  setLanguage('en');
  setBare(true);
  try {
    // Four step names made, one shown: "save" is left over.
    t('common.save');
    assert.match(html(Button, { label: t('common.close') }), /data-tip="common\.close"/);
    // React commits that draw, and the kit's layout effect drops what was left.
    assert.ok(commit(endDraw) >= 1, 'the drawn button asks for endDraw after its commit');
    // The same task, a later draw: a button whose words are the person's own "save" takes nothing.
    const later = html(Button, { label: 'save' });
    assert.match(later, />save<\/span>/);
    assert.match(later, /data-tip="save"/);
    assert.doesNotMatch(later, /data-key|common\.save/);
    // A call made for the new draw is still taken by it.
    assert.match(html(Button, { label: t('common.save') }), /data-tip="common\.save"/);
  } finally {
    setBare(null);
  }
});

test('a caller that names an element itself keeps its own data-key', async () => {
  const { Button, t } = await kit();
  assert.match(html(Button, { label: t('common.save'), 'data-key': 'mine' }), /data-key="mine"/);
  assert.doesNotMatch(html(Button, { label: t('common.save'), 'data-key': 'mine' }), /data-key="save"/);
});

test("a button whose tip is a label of its own carries the tip's key too, and the spelling pane's Custom Dictionary link its key", async () => {
  const { Button, setLanguage, t } = await kit();
  setLanguage('en');
  // The spelling pane's Change: its label is Change, its tip "Change, this one" in its own words.
  const change = html(Button, { label: t('common.change'), title: t('ui.proofing.change-this-one') });
  assert.match(change, /data-key="change" data-scope="common"/);
  assert.match(change, /data-tip-key="change-this-one" data-tip-scope="ui.proofing"/);
  // A tip no key gave carries none.
  assert.doesNotMatch(html(Button, { label: t('common.change'), title: 'Report.docx' }), /data-tip-key/);
  const proofing = fs.readFileSync(path.join(here, '..', 'packages', 'office-ui', 'src', 'proofing.js'), 'utf8');
  assert.match(proofing, /className="pf-link" onClick=\{onDictionary\} \{\.\.\.keyAttrs\(t\('ui\.proofing\.custom-dictionary-2'\)\)\}/);
});
