// Documents: Review → Delete removes comments from the window and the file.
//
// The three ways in, as a person meets them: the caret in a commented
// paragraph and Delete → Delete; the trash button beside one comment in the
// Comments list; Delete → Delete All Comments in Document. After each the
// engine's count is read, and at the end the saved file is.

import fs from 'node:fs';
import path from 'node:path';
import { Document } from '@rutba/ooxml';
import { buildDocx } from '@rutba/ooxml/build';

/** A document with a comment on each of its three paragraphs. */
function commented() {
  const doc = Document.open(buildDocx({ paragraphs: ['First point.', 'Second point.', 'Third point.'] }));
  doc.addComment(0, { author: 'A. Reader', text: 'On the first.' });
  doc.addComment(1, { author: 'A. Reader', text: 'On the second.' });
  doc.addComment(2, { author: 'B. Colleague', text: 'On the third.' });
  return doc.save();
}

/**
 * @param {object} h the harness: open, check, until, wait, capture, doc
 */
export async function verifyWordComments(h, { dir }) {
  const { open, check, until, wait, capture, doc } = h;
  const file = path.join(dir, 'commented.docx');
  try {
    fs.writeFileSync(file, commented());
    const win = await open('word', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="2"]'))`), 'the commented document', 10000);
    const session = doc.sessions().filter((s) => s.kind === 'doc' && s.path === file).pop();
    const count = () => (doc.model({ id: session.id }).comments || []).length;
    const clickTab = (name) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(), 'tab'`);
    const openDelete = () => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Delete' && (n.title || n.dataset.tip || '').startsWith('Delete — the comment'));
      if (!b) return 'no button';
      if (b.disabled) return 'disabled';
      b.click();
      return 'clicked';
    })()`);
    const pick = (label) => js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); if (!b) return 'no item'; if (b.disabled) return 'disabled'; b.click(); return 'picked'; })()`);

    await clickTab('Review');
    await wait(250);

    // 1. The caret in the second paragraph — a real click, so the window
    // knows where it is — and Delete → Delete takes its comment.
    const { id } = session;
    const at = await js(`(() => { const b = document.querySelector('.wd-page [data-block="1"]'); b.scrollIntoView({ block: 'center' }); const r = b.getBoundingClientRect(); return { x: Math.round(r.left + 6), y: Math.round(r.top + r.height / 2) }; })()`);
    win.webContents.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
    win.webContents.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1 });
    await until(() => doc.model({ id }).selection?.focus?.block === 1, 'the caret in the second paragraph', 4000).catch(() => {});
    await wait(300);
    const opened = await openDelete();
    await until(() => js(`Boolean(document.querySelector('.rw-menu'))`), 'the Delete menu', 3000).catch(() => {});
    const picked = await pick('Delete');
    const afterOne = await until(() => count() === 2, 'one comment gone', 5000).then(() => true).catch(() => false);
    const left = (doc.model({ id }).comments || []).map((c) => c.text);
    check('documents: the caret in a commented paragraph and Delete → Delete takes that comment and no other',
      opened === 'clicked' && picked === 'picked' && afterOne === true && left.join('|') === 'On the first.|On the third.', JSON.stringify({ opened, picked, left }));

    // 2. The Comments list: the trash beside the third comment.
    await js(`(() => { [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => /^Show/.test(n.textContent.trim()))?.click(); return 1; })()`);
    await until(() => js(`document.querySelectorAll('.wd-comment-row').length === 2`), 'the Comments list', 4000).catch(() => {});
    const trashed = await js(`(() => { const row = [...document.querySelectorAll('.wd-comment-row')].find((r) => r.textContent.includes('On the third.')); const b = row?.querySelector('.wd-comment-delete'); if (!b) return 'no trash'; b.click(); return 'clicked'; })()`);
    const afterTwo = await until(() => count() === 1, 'the third comment gone', 5000).then(() => true).catch(() => false);
    const rows = await until(() => js(`document.querySelectorAll('.wd-comment-row').length === 1`), 'the list to follow', 3000).then(() => true).catch(() => false);
    check('documents: the trash beside a comment in the Comments list deletes it, and the list follows',
      trashed === 'clicked' && afterTwo === true && rows === true, JSON.stringify({ trashed, afterTwo, rows, count: count() }));
    await js(`(() => { [...document.querySelectorAll('.rw-dialog button')].find((b) => b.textContent.trim() === 'Close')?.click(); return 1; })()`);
    await wait(200);

    // 3. Delete → Delete All Comments in Document, then Save.
    await openDelete();
    await until(() => js(`Boolean(document.querySelector('.rw-menu'))`), 'the Delete menu again', 3000).catch(() => {});
    const all = await pick('Delete All Comments in Document');
    const afterAll = await until(() => count() === 0, 'every comment gone', 5000).then(() => true).catch(() => false);
    const greyed = await until(() => js(`[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Delete' && (n.title || n.dataset.tip || '').startsWith('Delete —'))?.disabled === true`), 'Delete to grey', 3000).then(() => true).catch(() => false);
    await capture(win, 'word-comments-deleted.png');
    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => {
      try { return Document.open(fs.readFileSync(file)).comments().length === 0; } catch { return false; }
    }, 'the saved file', 8000).then(() => true).catch(() => false);
    const body = fs.readFileSync(file) && Document.open(fs.readFileSync(file)).pkg.text('word/document.xml');
    check('documents: Delete All Comments in Document empties the list, greys Delete, and the saved file carries no comment and no mark of one',
      all === 'picked' && afterAll === true && greyed === true && saved === true && !/<w:commentReference\b/.test(body),
      JSON.stringify({ all, afterAll, greyed, saved }));
  } catch (err) {
    check('documents: the comment delete checks ran', false, err.message);
  }
}
