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
      contents: "export { Button, Item, Group, Field, Dialog, Empty, Panel, Menu, Ribbon, keyAttrs } from './packages/office-ui/src/index.js'; export { setLanguage, registerCatalogue, t, tn } from './packages/office-ui/src/messages.js';",
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

test('a caller that names an element itself keeps its own data-key', async () => {
  const { Button, t } = await kit();
  assert.match(html(Button, { label: t('common.save'), 'data-key': 'mine' }), /data-key="mine"/);
  assert.doesNotMatch(html(Button, { label: t('common.save'), 'data-key': 'mine' }), /data-key="save"/);
});
