// Templates of one's own: Documents' app menu → Save as Template writes a
// .dotx into Office's own folder for them (Documents\Custom Office
// Templates), labelled as a template; Home lists it under Start something;
// choosing it there opens a new, untitled document made from it, the
// template left as it was. Documents points at a folder of the run's own
// while this runs. Run alone with RUTBA_VERIFY_ONLY=templates.
import fs from 'node:fs';
import path from 'node:path';
import { app, BrowserWindow } from 'electron';
import { buildDocx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml/package';

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn, doc
 * @param {{ dir: string }} args where the files are written
 */
export async function verifyPersonalTemplates({ open, check, until, wait, errorsIn, doc }, { dir }) {
  const documentsWas = app.getPath('documents');
  const documents = path.join(dir, 'Documents');
  fs.mkdirSync(documents, { recursive: true });
  app.setPath('documents', documents);
  const file = path.join(dir, 'letterhead.docx');
  const template = path.join(documents, 'Custom Office Templates', 'Letterhead.dotx');
  try {
    fs.writeFileSync(file, buildDocx({ styles: true, paragraphs: [{ text: 'Rutba Letterhead', style: 'Heading1' }, { text: 'Template words, kept for every letter.' }] }));
    const win = await open('word', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="1"]'))`), 'the paragraphs', 8000);
    await js(`document.querySelector('.rw-appmark').click(), 1`);
    await until(() => js(`[...document.querySelectorAll('.rw-menu button')].some((b) => b.textContent.trim() === 'Save as Template…')`), 'the Save as Template item', 4000);
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: template });
    await js(`[...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === 'Save as Template…').click(), 1`);
    const saved = await until(() => {
      try {
        const pkg = OoxmlPackage.read(fs.readFileSync(template));
        return /wordprocessingml\.template\.main\+xml$/.test(pkg.contentTypeOf(pkg.mainDocument()));
      } catch { return false; }
    }, 'the template written', 6000).then(() => true).catch(() => false);
    check('templates: Save as Template writes a .dotx labelled as a template, in the folder Office keeps them in', saved, saved ? template.replace(dir, '…') : 'no template');

    // Home lists it; choosing it makes a new document from it.
    const stamp = saved ? fs.statSync(template).mtimeMs : 0;
    const home = await open('home');
    const hjs = (code) => home.webContents.executeJavaScript(code);
    const listed = await until(() => hjs(`[...document.querySelectorAll('.home-template.personal')].some((b) => b.textContent.trim() === 'Letterhead')`), 'the template on Home', 6000).then(() => true).catch(() => false);
    const before = new Set(doc.sessions().map((s) => s.id));
    const windowsBefore = new Set(BrowserWindow.getAllWindows().map((w) => w.id));
    await hjs(`[...document.querySelectorAll('.home-template.personal')].find((b) => b.textContent.trim() === 'Letterhead')?.click(), 1`);
    const fresh = () => doc.sessions().find((s) => !before.has(s.id) && s.kind === 'doc' && !s.path);
    await until(() => Boolean(fresh()), 'the new document', 8000).catch(() => null);
    const made = fresh() || null;
    const words = made ? (doc.model({ id: made.id }).blocks || []).map((b) => b.text).join(' ') : '';
    const untouched = saved && fs.statSync(template).mtimeMs === stamp;
    check('templates: Home lists a template of one\'s own, and choosing it opens a new, untitled document made from it, the template left as it was',
      listed && Boolean(made) && /Template words/.test(words) && untouched,
      JSON.stringify({ listed, made: made ? { name: made.name, path: made.path } : null, words: words.slice(0, 60), untouched }));
    const complaints = [...await errorsIn(win), ...await errorsIn(home)];
    // The window Home opened is not one of the run's: it is closed here.
    for (const w of BrowserWindow.getAllWindows()) if (!windowsBefore.has(w.id) && !w.isDestroyed()) w.destroy();
    check('templates: templates of one\'s own report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    void wait;
  } catch (err) {
    check('templates: the template checks ran', false, err?.message || String(err));
  } finally {
    app.setPath('documents', documentsWas);
  }
}
