// Presentations: Review → Compare. A reviewed copy of the deck is merged
// in; the Revisions pane lists the slide's changes and the presentation's;
// markers sit by the changed shapes, and a marker's list ticks a change in
// (the words on the slide become the reviewer's) and out again; Accept
// takes one, this slide's or all; Reject puts them back; the deck saved
// holds what was accepted and opens as it was left. Run alone with
// RUTBA_VERIFY_ONLY=deck-compare.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';

function decks() {
  const bytes = buildPptx({ title: 'Plan', slides: [
    { layout: 'title', title: 'Plan for 2027', body: 'Draft for the board' },
    { layout: 'obj', title: 'Costs', body: ['Rent 100', 'Staff 200'] },
    { layout: 'obj', title: 'Risks', body: ['Late supplies'] },
  ] });
  const d = Deck.open(bytes);
  const body = d.slide(1).shapes.find((s) => s.placeholder?.type !== 'title' && s.text);
  d.setText(1, body.id, [{ runs: [{ text: 'Rent 100' }] }, { runs: [{ text: 'Staff 250 from March' }] }]);
  const title = d.slide(1).shapes.find((s) => s.placeholder?.type === 'title');
  d.setGeometry(1, title.id, { ...title.geometry, x: title.geometry.x + 60 });
  d.addTextBox(1, { x: 120, y: 420, w: 360, h: 60, paragraphs: [{ runs: [{ text: 'Ask the board' }] }], name: 'Reviewer note' });
  d.moveSlide(2, 1);
  return { original: bytes, reviewed: Buffer.from(d.save()) };
}

/**
 * @param {object} h the harness: open, check, until, wait, capture, doc
 */
export async function verifyDeckCompare(h, { dir }) {
  const { open, check, wait, capture, doc } = h;
  // Waits as the harness's until does, and answers what the condition found once it held.
  const until = async (fn, what, ms) => { let found; await h.until(async () => { found = await fn(); return found; }, what, ms); return found; };
  const own = path.join(dir, 'deck-compare');
  const file = path.join(own, 'plan.pptx');
  const reviewed = path.join(own, 'plan (reviewed).pptx');
  try {
    fs.mkdirSync(own, { recursive: true });
    const { original, reviewed: r } = decks();
    fs.writeFileSync(file, original);
    fs.writeFileSync(reviewed, r);
  } catch (err) {
    return check('deck-compare: the fixtures were made', false, err.message);
  }
  try {
    const win = await open('slides', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const session = doc.sessions().filter((s) => s.kind === 'deck' && s.path === file).pop();
    const fresh = async (name) => { wc.invalidate(); await wait(700); await capture(win, name); };
    const ribbon = (tip) => js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || n.title || '').startsWith(${JSON.stringify(tip)})); if (!b) return 'no button'; if (b.disabled) return 'disabled'; b.click(); return 'clicked'; })()`);
    const menuItem = (label) => js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); b?.click(); return Boolean(b); })()`);
    const changes = () => js(`[...document.querySelectorAll('.sl-rev [data-change]')].map((n) => ({ id: n.dataset.change, kind: n.dataset.kind, on: n.querySelector('input').checked, text: n.textContent }))`);
    const words = (i) => JSON.stringify((doc.model({ id: session.id, slide: i }).slide?.shapes || []).map((s) => s.text?.paragraphs?.map((p) => p.plain)));

    await until(() => js(`Boolean(document.querySelector('.sl-slide'))`), 'the deck', 10000);
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Review')?.click(), 1`);
    await wait(250);
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'open', answer: [reviewed] });
    const pressed = await ribbon('Compare');
    const pane = await until(() => js(`Boolean(document.querySelector('.sl-rev'))`), 'the Revisions pane', 8000).catch(() => false);
    const deckList = await until(async () => { const l = await changes(); return l.some((c) => c.kind === 'slideMoved') ? l : null; }, 'the presentation changes', 6000).catch(() => null);
    check(
      'deck-compare: Review → Compare merges the reviewed copy and opens the Revisions pane, the presentation\'s changes listed (a slide moved)',
      pressed === 'clicked' && pane === true && Array.isArray(deckList),
      JSON.stringify({ pressed, pane, deckList }),
    );

    // The slide with the reviewer's changes: its markers and its list.
    await js(`(() => { const b = document.querySelector('.sl-rev-jump .sl-rev-goto'); b?.click(); return Boolean(b); })()`);
    const costs = await until(async () => { const l = await changes(); return l.filter((c) => ['text', 'moved', 'shapeAdded'].includes(c.kind)).length >= 3 ? l : null; }, 'the slide\'s changes', 6000).catch(() => null);
    const markers = await js(`document.querySelectorAll('.sl-slide .sl-rev-marker').length`);
    const ghost = await js(`document.querySelectorAll('.sl-slide .sl-rev-ghost').length`);
    await fresh('deck-compare-pane.png');
    const textChange = (costs || []).find((c) => c.kind === 'text');
    check(
      'deck-compare: the slide\'s changes are listed by shape, the words put in and taken out, a title moved, a box inserted, with a marker by each shape and the inserted box outlined where it would go',
      Boolean(textChange) && /Inserted "250 from March"/.test(textChange.text) && /Deleted "200"/.test(textChange.text) && markers >= 3 && ghost === 1,
      JSON.stringify({ costs, markers, ghost }),
    );

    // A marker's list: ticking the text change makes the words the reviewer's; unticking puts them back.
    // The Costs slide, found by its title: the slide the window shows is the renderer's to know.
    const at = doc.model({ id: session.id }).outline.findIndex((o) => /Costs/.test(o.title || ''));
    await js(`(() => { const m = [...document.querySelectorAll('.sl-slide .sl-rev-marker')].find((n) => /Inserted/.test(n.dataset.tip || '')); m?.click(); return Boolean(m); })()`);
    await until(() => js(`Boolean(document.querySelector('.sl-rev-pop .sl-rev-tick'))`), 'the marker\'s list', 4000).catch(() => {});
    await fresh('deck-compare-marker.png');
    await js(`document.querySelector('.sl-rev-pop [data-kind="text"] input')?.click(), 1`);
    const ticked = await until(() => /Staff 250 from March/.test(words(at)), 'the reviewer\'s words', 6000).catch(() => false);
    await until(() => js(`Boolean(document.querySelector('.sl-rev [data-kind="text"] input')?.checked)`), 'the pane showing the tick', 4000).catch(() => {});
    await js(`document.querySelector('.sl-rev [data-kind="text"] input')?.click(), 1`);
    const unticked = await until(() => /Staff 200/.test(words(at)) && !/Staff 250/.test(words(at)), 'the words put back', 6000).catch(() => false);
    check('deck-compare: a tick in a marker\'s list accepts a change on the slide and an untick rejects it', ticked === true && unticked === true, JSON.stringify({ ticked, unticked, now: words(at) }));

    // Accept → All Changes to This Slide, then Accept → All Changes to the Presentation.
    await ribbon('Accept');
    await menuItem('Accept All Changes to This Slide');
    const slideDone = await until(async () => (await changes()).filter((c) => c.kind !== 'slideMoved').every((c) => c.on), 'this slide\'s changes', 6000).catch(() => false);
    const boxThere = (doc.model({ id: session.id, slide: at }).slide?.shapes || []).some((s) => /^Reviewer note/.test(s.name || ''));
    await ribbon('Accept');
    await menuItem('Accept All Changes to the Presentation');
    const allDone = await until(async () => { const l = await changes(); return l.length && l.every((c) => c.on); }, 'every change', 6000).catch(() => false);
    const order = doc.model({ id: session.id }).outline.map((o) => o.title || '');
    check(
      'deck-compare: Accept takes all of this slide\'s changes (the reviewer\'s box is put on it) and then all of the presentation\'s (the slide moved)',
      slideDone === true && boxThere && allDone === true && /Risks/.test(order[1] || ''),
      JSON.stringify({ slideDone, boxThere, allDone, order }),
    );

    // Reject → All Changes to the Presentation puts it back; Accept again; save and open again.
    await ribbon('Reject');
    await menuItem('Reject All Changes to the Presentation');
    const rejected = await until(async () => (await changes()).every((c) => !c.on), 'everything rejected', 6000).catch(() => false);
    const back = doc.model({ id: session.id }).outline.map((o) => o.title || '');
    await ribbon('Accept');
    await menuItem('Accept All Changes to the Presentation');
    await until(async () => (await changes()).every((c) => c.on), 'everything accepted again', 6000).catch(() => {});
    await fresh('deck-compare-accepted.png');
    await ribbon('End Review');
    const ended = await until(() => js(`!document.querySelector('.sl-rev')`), 'the review ended', 4000).catch(() => false);
    await js(`(() => { const b = [...document.querySelectorAll('.rw-btn')].find((n) => (n.dataset.tip || n.title || '').startsWith('Save')); b?.click(); return 1; })()`);
    await until(() => doc.meta({ id: session.id }).dirty === false, 'the save', 15000).catch(() => {});
    const again = Deck.open(fs.readFileSync(file));
    const costsAt = [0, 1, 2].find((i) => again.slide(i).shapes.some((s) => /Costs/.test(JSON.stringify(s.text || ''))));
    const savedOk = again.slideCount === 3 && costsAt === 2 && again.slide(2).shapes.some((s) => /^Reviewer note/.test(s.name || '')) && /Staff 250 from March/.test(JSON.stringify(again.slide(2).shapes.map((s) => s.text)));
    check(
      'deck-compare: Reject puts every change back, End Review closes the review, and the saved deck holds what was accepted',
      rejected === true && /Costs/.test(back[1] || '') && ended === true && savedOk,
      JSON.stringify({ rejected, back, ended, savedOk }),
    );
  } catch (err) {
    check('deck-compare: the Compare checks ran', false, err.stack || err.message);
  }
}
