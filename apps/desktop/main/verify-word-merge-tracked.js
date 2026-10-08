// Documents: merging table cells while Track Changes is on asks first, as
// Word does — "This action will not be marked as a change" — and Cancel
// leaves the table as it was while OK merges. Run alone with
// RUTBA_VERIFY_ONLY=mergetracked.

import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';

export async function verifyWordMergeTracked({ open, check, until, doc, sessionFor }, { dir }) {
  const file = path.join(dir, 'merge-tracked.docx');
  try {
    fs.writeFileSync(file, buildDocx({ styles: true, paragraphs: [
      { text: 'Before the table' },
      { table: { rows: [['Region', 'Q1'], ['North', '120'], ['South', '140']] } },
      { text: 'After the table' },
    ] }));
    const win = await open('word', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const model = () => doc.model({ id: sessionFor('doc').id });
    const index = (text) => model().blocks.findIndex((b) => b.text === text);
    await until(() => js(`document.querySelectorAll('.wd-page table.wd-table').length === 1`), 'the table', 8000);
    // Track Changes on, as the Review tab's button turns it on.
    doc.apply({ id: sessionFor('doc').id, ops: [{ op: 'toggleTrackChanges', on: true }] });
    await until(() => js(`Boolean(document.querySelector('.rw-tab'))`), 'the ribbon', 4000);
    const select = (from, to) => js(`(() => {
      const el = (i) => document.querySelector('.wd-page [data-block="' + i + '"]');
      const a = el(${index(from)}); const b = el(${index(to)});
      a.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }));
      const r = document.createRange();
      const first = (n) => { const w = document.createTreeWalker(n, NodeFilter.SHOW_TEXT); return w.nextNode() || n; };
      r.setStart(first(a), 0); r.setEnd(first(b), 0);
      const s = getSelection(); s.removeAllRanges(); s.addRange(r);
      document.querySelector('.wd-page').dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
      return true;
    })()`);
    const rows = () => js(`[...document.querySelector('.wd-page table.wd-table').tBodies[0].rows].map((r) => r.cells.length)`);
    const button = `[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Merge Cells')`;
    await select('North', 'South');
    await until(() => js(`(() => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Table Layout')?.click(); return Boolean(${button}) && ${button}.disabled === false; })()`), 'Merge Cells', 4000).catch(() => {});
    const before = await rows();
    const queue = globalThis.__rutbaCheckDialogAnswers;
    queue?.push({ kind: 'message', answer: { response: 1, checked: false } });
    await js(`${button}?.click(), 1`);
    await new Promise((r) => setTimeout(r, 800));
    const afterCancel = await rows();
    await select('North', 'South');
    queue?.push({ kind: 'message', answer: { response: 0, checked: false } });
    await js(`${button}?.click(), 1`);
    await until(async () => JSON.stringify(await rows()) !== JSON.stringify(before), 'the merge', 4000).catch(() => {});
    const afterOk = await rows();
    const tracking = model().trackRevisions;
    check('documents: merging cells while Track Changes is on asks first, as Word does; Cancel leaves the table, OK merges',
      tracking === true && JSON.stringify(afterCancel) === JSON.stringify(before) && JSON.stringify(afterOk) !== JSON.stringify(before),
      JSON.stringify({ tracking, before, afterCancel, afterOk }));
  } catch (err) {
    check('documents: the tracked merge checks ran', false, err?.message || String(err));
  }
}
