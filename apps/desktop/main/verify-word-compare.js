// Documents: Review → Compare, pressed on the ribbon in the revised
// document — the Compare Documents box takes this document as the revised
// one and the original picked with Browse; OK opens a new window on the
// compared document, the revised one with what changed marked as
// revisions: the words taken out and put in, the paragraph removed.

import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';

/**
 * @param {object} h the harness: open, check, until, wait, doc
 */
export async function verifyWordCompare(h, { dir }) {
  const { open, check, until, wait, doc } = h;
  const original = path.join(dir, 'compare-original.docx');
  const revised = path.join(dir, 'compare-revised.docx');
  try {
    fs.writeFileSync(original, buildDocx({ styles: true, paragraphs: [{ text: 'Budget' }, { text: 'We will spend ten thousand on travel.' }, { text: 'This paragraph is dropped later.' }, { text: 'Thanks.' }] }));
    fs.writeFileSync(revised, buildDocx({ styles: true, paragraphs: [{ text: 'Budget' }, { text: 'We will spend eight thousand on travel.' }, { text: 'Thanks.' }] }));
    const win = await open('word', revised);
    const js = (code) => win.webContents.executeJavaScript(code);
    const before = new Set(doc.sessions().map((s) => s.id));

    await until(() => js(`document.querySelectorAll('.wd-page [data-block]').length >= 3`), 'the page', 8000);
    const pressed = await js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Review')?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Compare');
      if (!b || b.disabled) return 'no button'; b.click(); return 'clicked';
    })()`);
    await until(() => js(`Boolean(document.querySelector('.wd-compare-ok'))`), 'the Compare Documents box', 4000).catch(() => {});
    const revisedShown = await js(`document.querySelector('.wd-compare-revised')?.textContent || ''`);
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'open', answer: [original] });
    await js(`document.querySelector('.wd-compare-browse-original')?.click(), 1`);
    await until(() => js(`(document.querySelector('.wd-compare-original')?.textContent || '').includes('compare-original')`), 'the original picked', 4000).catch(() => {});
    await js(`document.querySelector('.wd-compare-ok')?.click(), 1`);
    const made = await until(() => doc.sessions().some((s) => !before.has(s.id) && s.kind === 'doc'), 'the compared document', 8000).then(() => true).catch(() => false);
    const session = doc.sessions().find((s) => !before.has(s.id) && s.kind === 'doc');
    const runs = session ? (doc.model({ id: session.id }).blocks || []).flatMap((b) => (b.runs || []).map((r) => ({ text: r.text || r.del?.text || '', ins: Boolean(r.ins), del: Boolean(r.del) }))) : [];
    const took = runs.some((r) => r.del && /ten/.test(r.text));
    const put = runs.some((r) => r.ins && /eight/.test(r.text));
    const dropped = runs.some((r) => r.del && /dropped later/.test(r.text));
    await wait(300);
    check('documents: Review → Compare makes a new document of the revised one with the original\'s differences marked as revisions — words taken out and put in, a paragraph removed',
      pressed === 'clicked' && /compare-revised/.test(revisedShown) && made && took && put && dropped,
      JSON.stringify({ pressed, revisedShown, made, took, put, dropped, name: session?.name }));
  } catch (err) {
    check('documents: the Compare checks ran', false, err.message);
  }
}
