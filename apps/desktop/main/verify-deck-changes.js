// Presentations: Review → Show Changes — a deck opened for the first time
// has nothing to compare with; the file is then changed elsewhere (words on
// one slide and a shape added, the last slide moved to the front) and opened
// again: a toast says how many slides differ, the strip marks them, Show
// Changes lists each with what changed, a click goes to the slide, and
// Dismiss clears the list. Run alone with RUTBA_VERIFY_ONLY=deckchanges.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';

export async function verifyDeckChanges({ open, check, until, wait, errorsIn, doc }, { dir }) {
  const file = path.join(dir, 'changes.pptx');
  try {
    fs.writeFileSync(file, Deck.open(buildPptx({ title: 'Plan', slides: [
      { layout: 'title', title: 'The plan' },
      { layout: 'obj', title: 'Steps', body: ['Plan', 'Build'] },
      { layout: 'obj', title: 'Costs', body: ['Small'] },
    ] })).save());
    const sessionOf = () => doc.sessions().filter((s) => s.kind === 'deck' && s.path === file).pop();
    const showChanges = async (js) => {
      await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Review')?.click(), 1`);
      await until(() => js(`Boolean(document.querySelector('.sl-show-changes'))`), 'the Show Changes button', 4000).catch(() => {});
      await js(`(() => { const b = document.querySelector('.sl-show-changes'); if (b && b.getAttribute('aria-pressed') !== 'true') b.click(); return 1; })()`);
      await until(() => js(`Boolean(document.querySelector('.sl-changes, .rw-panel .sl-pane-empty'))`), 'the Changes pane', 4000).catch(() => {});
    };

    // A window of its own stays open throughout: closing the last window would end the run.
    const keep = path.join(dir, 'changes-keep.pptx');
    fs.writeFileSync(keep, Deck.open(buildPptx({ title: 'Keep', slides: [{ layout: 'title', title: 'Keep' }] })).save());
    await open('slides', keep);

    // First open: nothing to compare with yet.
    let win = await open('slides', file);
    let js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.sl-thumb').length === 3`), 'the strip', 8000).catch(() => {});
    await wait(500);
    await showChanges(js);
    const first = await js(`document.querySelector('.rw-panel .sl-pane-empty')?.textContent || ''`);
    check('presentations: Show Changes on a deck open here for the first time says there is nothing yet to compare with', /first time this deck has been open/.test(first), first.slice(0, 80));
    const firstSession = sessionOf();
    win.close();
    await until(() => win.isDestroyed(), 'the window closed', 4000).catch(() => {});
    if (firstSession && doc.sessions().some((s) => s.id === firstSession.id)) await doc.close({ id: firstSession.id }).catch(() => {});

    // The file changed elsewhere.
    const d = Deck.open(fs.readFileSync(file));
    const body = d.slide(1).shapes.find((s) => (s.text?.paragraphs || []).some((p) => p.runs?.some((r) => r.text === 'Plan')));
    d.setText(1, body.id, [{ runs: [{ text: 'Plan' }] }, { runs: [{ text: 'Build' }] }, { runs: [{ text: 'Ship' }] }]);
    d.addShape(1, { preset: 'rect', x: 40, y: 40, w: 90, h: 40, name: 'Badge' });
    d.moveSlide(2, 0);
    fs.writeFileSync(file, d.save());

    win = await open('slides', file);
    js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.sl-thumb').length === 3`), 'the strip again', 8000).catch(() => {});
    const toasted = await until(() => js(`[...document.querySelectorAll('.rw-toast, [role="status"]')].some((t) => /different since this deck was last open here/.test(t.textContent))`), 'the toast', 6000).then(() => true).catch(() => false);
    const marked = await until(() => js(`[...document.querySelectorAll('.sl-thumb')].map((t, i) => (t.querySelector('.sl-thumb-chg') ? i : -1)).filter((i) => i >= 0)`), 'the marks', 4000).catch(() => []);
    const marks = await js(`[...document.querySelectorAll('.sl-thumb')].map((t, i) => (t.querySelector('.sl-thumb-chg') ? i : -1)).filter((i) => i >= 0)`);
    await showChanges(js);
    const items = await js(`[...document.querySelectorAll('.sl-change')].map((b) => ({ kind: b.dataset.kind, text: b.textContent }))`);
    check('presentations: a deck changed elsewhere opens with a toast and its changed slides marked in the strip, and Show Changes lists what changed on each',
      toasted && JSON.stringify(marks) === '[0,2]' && items.length === 2
        && items.some((x) => x.kind === 'moved' && /Costs/.test(x.text) && /moved from slide 3/.test(x.text))
        && items.some((x) => x.kind === 'changed' && /Steps/.test(x.text) && /words of .* changed/.test(x.text) && /Badge added/.test(x.text)),
      JSON.stringify({ toasted, marks, items }));

    // A click goes to the slide; Dismiss clears the list.
    await js(`[...document.querySelectorAll('.sl-change')].find((b) => b.dataset.kind === 'changed')?.click(), 1`);
    const went = await until(() => js(`[...document.querySelectorAll('.sl-thumb')].findIndex((t) => t.classList.contains('active')) === 2`), 'the slide', 3000).then(() => true).catch(() => false);
    await js(`document.querySelector('.sl-changes-dismiss')?.click(), 1`);
    const cleared = await until(() => js(`!document.querySelector('.sl-change') && !document.querySelector('.sl-thumb-chg')`), 'the list cleared', 3000).then(() => true).catch(() => false);
    check('presentations: a change clicked goes to its slide, and Dismiss clears the list and the marks', went && cleared, JSON.stringify({ went, cleared }));

    const complaints = await errorsIn(win);
    check('presentations: Show Changes reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('presentations: the Show Changes checks ran', false, err?.message || JSON.stringify(err));
  }
}
