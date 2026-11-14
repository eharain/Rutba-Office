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

/** The UI kit bundled for node, React shared with this test, with its messages module. */
async function kit() {
  const out = await esbuild.build({
    stdin: {
      contents: "export { Button, Item, Group, Field, Dialog, Empty, Panel, Menu, Ribbon, keyAttrs } from './packages/office-ui/src/index.js'; export { setLanguage, registerCatalogue, t, tn, setBare } from './packages/office-ui/src/messages.js';",
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
  const file = path.join(here, '..', 'node_modules', `.label-key-attrs-${process.pid}.mjs`);
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
    // The other elements: the full key as their title.
    assert.match(html(Item, { label: tn(3, 'common.count-word') }), /title="common\.count-word"/);
    assert.match(html(Group, { label: t('ui.proofing.errors') }), /<div class="rw-group-label" data-key="errors" data-scope="ui\.proofing" title="ui\.proofing\.errors">errors<\/div>/);
    assert.match(html(Menu, { x: 0, y: 0, onClose() {}, items: [{ label: t('common.close'), run() {} }] }), /title="common\.close"/);
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
    // A menu of three "close" items, each with its own key and its own full key as its tip.
    const items = html(Menu, { x: 0, y: 0, onClose() {}, items: [{ label: t('mail.close'), run() {} }, { label: t('word.ribbon.close'), run() {} }, { label: t('common.close'), run() {} }] }).split('</button>');
    ['mail', 'word.ribbon', 'common'].forEach((scope, i) => {
      assert.ok(items[i].includes(`title="${scope}.close"`) && items[i].includes(`data-key="close" data-scope="${scope}"`), items[i]);
    });
    // Every call taken: a file named "save" in a list is no label, and has no key and no full key as its tip.
    const file = html(Item, { label: 'save' });
    assert.doesNotMatch(file, /data-key|data-scope/);
    assert.match(file, /title="save"/);
    // Nor once the task the name was shown in is over.
    t('common.save');
    await new Promise((resolve) => setTimeout(resolve, 5));
    assert.doesNotMatch(html(Item, { label: 'save' }), /data-key|title="common\.save"/);
    assert.doesNotMatch(html(Menu, { x: 0, y: 0, onClose() {}, items: [{ label: 'close', run() {} }] }), /data-key|title="[a-z.]*close"/);
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
