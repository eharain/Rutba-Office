// Documents: Review → Language, pressed on the ribbon — words selected by
// the keyboard are marked French in the dialog, the page stops underlining
// them, they are then marked not to be checked, the document's default is
// set, and all of it is in the saved file as Word keeps it.

import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml';

/**
 * @param {object} h the harness: open, check, until, wait, doc, sessionFor
 */
export async function verifyWordLanguage(h, { dir }) {
  const { open, check, until, wait, doc, sessionFor } = h;
  const file = path.join(dir, 'language.docx');
  try {
    fs.writeFileSync(file, buildDocx({ styles: true, paragraphs: [{ text: 'Bonjour mes amis and then English words.' }] }));
    const win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const session = sessionFor('doc');
    const model = () => doc.model({ id: session.id });
    const key = async (keyCode, modifiers = []) => {
      wc.sendInputEvent({ type: 'keyDown', keyCode, modifiers });
      wc.sendInputEvent({ type: 'keyUp', keyCode, modifiers });
      await wait(60);
    };
    const openDialog = () => js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Review')?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Language');
      if (!b) return 'no button'; if (b.disabled) return 'disabled'; b.click(); return 'clicked';
    })()`);
    const dialogUp = () => until(() => js(`Boolean(document.querySelector('.pf-lang-list'))`), 'the Language dialog', 4000).then(() => true).catch(() => false);
    const runOf = (word) => (model().blocks[0].runs || []).find((r) => r.text.includes(word)) || {};

    // Select "Bonjour mes amis" with the keyboard.
    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="0"]'))`), 'the paragraph', 8000);
    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    await key('Home', ['control']);
    for (let i = 0; i < 3; i++) await key('Right', ['control', 'shift']);
    await until(() => { const s = model().selection; return s && s.focus.offset >= 16 && s.anchor.offset === 0; }, 'three words selected', 3000).catch(() => {});

    // 1. Marked French in the dialog.
    const pressed = await openDialog();
    const up = await dialogUp();
    await js(`(() => {
      const find = document.querySelector('.pf-lang-find');
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(find, 'French');
      find.dispatchEvent(new Event('input', { bubbles: true }));
      return 1;
    })()`);
    await wait(150);
    const picked = await js(`(() => { const r = document.querySelector('.pf-lang-row[data-tag="fr-FR"]'); if (!r) return 'no row'; r.click(); return 'picked'; })()`);
    await wait(100);
    await js(`document.querySelector('.pf-lang-ok')?.click(), 1`);
    const marked = await until(() => runOf('Bonjour').lang === 'fr-FR' && runOf('English').lang !== 'fr-FR', 'the French mark', 4000).then(() => true).catch(() => false);
    const span = await until(() => js(`Boolean([...document.querySelectorAll('.wd-page span[lang="fr-FR"]')].find((s) => s.textContent.includes('Bonjour') && s.spellcheck === false))`), 'the page to drop the underline', 3000).then(() => true).catch(() => false);
    check('documents: Review → Language marks the selected words French, and the page stops underlining them',
      pressed === 'clicked' && up && picked === 'picked' && marked && span, JSON.stringify({ pressed, up, picked, marked, span, run: runOf('Bonjour') }));

    // 2. Opened again on the same words: it says French; marked not to be checked.
    await openDialog();
    await dialogUp();
    const onFrench = await js(`document.querySelector('.pf-lang-row.on')?.dataset.tag || null`);
    await js(`(() => { document.querySelector('.pf-lang-noproof').click(); document.querySelector('.pf-lang-ok').click(); return 1; })()`);
    const unproofed = await until(() => runOf('Bonjour').noProof === true, 'the noProof mark', 4000).then(() => true).catch(() => false);
    check('documents: the Language dialog opens on the selection\'s language, and "Do not check spelling or grammar" marks the words',
      onFrench === 'fr-FR' && unproofed, JSON.stringify({ onFrench, unproofed }));

    // 3. Set As Default: English (United States) for the document.
    await openDialog();
    await dialogUp();
    await js(`(() => { document.querySelector('.pf-lang-row[data-tag="en-US"]')?.click(); return 1; })()`);
    await wait(100);
    await js(`document.querySelector('.pf-lang-default')?.click(), 1`);
    const byDefault = await until(async () => (await doc.proof({ id: session.id, action: 'language' }))?.default === 'en-US', 'the default language', 4000).then(() => true).catch(() => false);
    await until(() => js(`document.querySelector('.pf-lang-current')?.textContent === 'English (United States)'`), 'the dialog to say so', 3000).catch(() => {});
    const shown = await js(`document.querySelector('.pf-lang-current')?.textContent || null`);
    await js(`[...document.querySelectorAll('.rw-dialog .rw-btn, .rw-btn')].find((b) => b.textContent.trim() === 'Cancel')?.click(), 1`);
    check('documents: Set As Default makes English (United States) the document\'s language, and the dialog says so',
      byDefault && shown === 'English (United States)', JSON.stringify({ byDefault, shown }));

    // 4. Saved: the marks and the default are in the file.
    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => {
      try {
        const pkg = OoxmlPackage.read(fs.readFileSync(file));
        return /<w:noProof\/>[\s\S]*?<w:lang w:val="fr-FR"\/><\/w:rPr><w:t[^>]*>Bonjour/.test(pkg.text('word/document.xml'))
          && /<w:rPrDefault><w:rPr>[\s\S]*?<w:lang w:val="en-US"\/>/.test(pkg.text('word/styles.xml'));
      } catch {
        return false;
      }
    }, 'the saved marks', 8000).then(() => true).catch(() => false);
    check('documents: the saved file carries w:lang, w:noProof and the default language, as Word keeps them', saved, saved ? 'in the file' : 'not in the file');
  } catch (err) {
    check('documents: the Language checks ran', false, err.message);
  }
}
