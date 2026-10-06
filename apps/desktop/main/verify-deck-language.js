// Presentations: Review → Language, pressed on the ribbon — the title box is
// clicked, marked French and not to be checked in the dialog, the body is
// left as it was, and the saved deck carries both on the title's runs.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait, doc
 */
export async function verifyDeckLanguage(h, { dir }) {
  const { open, check, until, wait, doc } = h;
  const file = path.join(dir, 'language.pptx');
  try {
    fs.writeFileSync(file, buildPptx({ title: 'Langue', slides: [{ layout: 'title', title: 'Bonjour tout le monde', body: 'Hello there' }] }));
    const win = await open('slides', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.sl-hit[data-shape]').length > 0`), 'the slide\'s shapes', 8000);
    const session = doc.sessions().filter((s) => s.kind === 'deck' && s.path === file).pop();
    const runOf = (word) => (doc.model({ id: session.id }).slide?.shapes || [])
      .flatMap((s) => (s.text?.paragraphs || []).flatMap((p) => p.runs || []))
      .find((r) => String(r.text || '').includes(word)) || {};
    const press = () => js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Review')?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Language');
      if (!b) return 'no button'; if (b.disabled) return 'disabled'; b.click(); return 'clicked';
    })()`);

    // The title's box, selected.
    await js(`(() => { document.querySelector('.sl-hit[data-shape]')?.click(); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.sl-hit.selected'))`), 'the title selected', 3000).catch(() => {});
    const pressed = await press();
    const up = await until(() => js(`Boolean(document.querySelector('.pf-lang-list'))`), 'the Language dialog', 4000).then(() => true).catch(() => false);
    const noDefault = await js(`!document.querySelector('.pf-lang-default')`);
    await js(`(() => {
      document.querySelector('.pf-lang-row[data-tag="fr-FR"]')?.click();
      return 1;
    })()`);
    await wait(100);
    await js(`(() => { document.querySelector('.pf-lang-noproof').click(); document.querySelector('.pf-lang-ok').click(); return 1; })()`);
    const marked = await until(() => runOf('Bonjour').lang === 'fr-FR' && runOf('Bonjour').noProof === true, 'the title marked', 4000).then(() => true).catch(() => false);
    const bodyKept = runOf('Hello').lang !== 'fr-FR' && !runOf('Hello').noProof;
    check('presentations: Review → Language marks the selected box\'s words French and not to be checked, and leaves the other box alone',
      pressed === 'clicked' && up && noDefault && marked && bodyKept, JSON.stringify({ pressed, up, noDefault, title: runOf('Bonjour'), body: runOf('Hello') }));

    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => {
      try {
        const d = Deck.open(fs.readFileSync(file));
        return /<a:rPr lang="fr-FR"[^>]*noProof="1"[^>]*>(?:(?!<\/a:r>)[\s\S])*Bonjour/.test(d.pkg.text(d.slideParts[0].part));
      } catch {
        return false;
      }
    }, 'the saved marks', 8000).then(() => true).catch(() => false);
    check('presentations: the saved deck carries lang and noProof on the title\'s runs, as PowerPoint keeps them', saved, saved ? 'in the file' : 'not in the file');
  } catch (err) {
    check('presentations: the Language checks ran', false, err.message);
  }
}
