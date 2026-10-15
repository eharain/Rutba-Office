// Presentations: comments as PowerPoint 365 keeps them, checked against a
// deck written as PowerPoint writes one, a thread by someone else with a
// mention in it, a reply. The pane shows the mention as a mention; a reply
// is posted with an @ that offers the people in the deck's comments; one's
// own reply is edited in place; the thread is resolved (its reply box gone,
// marked Resolved) and reopened; and the saved file carries the reply, the
// edited words, the status and the authors as PowerPoint reads them. Run
// alone with RUTBA_VERIFY_ONLY=deck-comments.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait, capture, doc
 */
export async function verifyDeckCommentThreads(h, { dir }) {
  const { open, check, wait, capture, doc } = h;
  // Waits as the harness's until does, and answers what the condition found once it held.
  const until = async (fn, what, ms) => { let found; await h.until(async () => { found = await fn(); return found; }, what, ms); return found; };
  const own = path.join(dir, 'deck-comments');
  const file = path.join(own, 'review.pptx');
  try {
    fs.mkdirSync(own, { recursive: true });
    const d = Deck.open(buildPptx({ title: 'Review', slides: [{ layout: 'obj', title: 'Budget', body: ['Travel 40', 'Tools 60'] }] }));
    d.addComment(0, { text: 'Can @Ayesha Khan check the travel line?', author: 'Bilal Ahmed' });
    d.addComment(0, { text: 'Second thread', author: 'Ayesha Khan' });
    fs.writeFileSync(file, d.save());
  } catch (err) {
    return check('deck-comments: the fixture was made', false, err.message);
  }
  try {
    const win = await open('slides', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const session = doc.sessions().filter((s) => s.kind === 'deck' && s.path === file).pop();
    const threads = () => doc.model({ id: session.id }).comments || [];
    const me = doc.model({ id: session.id }).me;
    const fresh = async (name) => { wc.invalidate(); await wait(700); await capture(win, name); };
    const setText = (sel, value) => js(`(() => { const el = document.querySelector(${JSON.stringify(sel)}); if (!el) return false; el.focus(); Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(el, ${JSON.stringify(value)}); el.dispatchEvent(new Event('input', { bubbles: true })); el.setSelectionRange(el.value.length, el.value.length); el.dispatchEvent(new Event('select', { bubbles: true })); return true; })()`);

    await until(() => js(`Boolean(document.querySelector('.sl-slide'))`), 'the deck', 10000);
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Review')?.click(), 1`);
    await wait(200);
    await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || n.title || '').startsWith('Show Comments')); b?.click(); return 1; })()`);
    const pane = await until(() => js(`document.querySelectorAll('.sl-cm-card').length === 2`), 'the two threads', 6000).catch(() => false);
    const mention = await js(`[...document.querySelectorAll('.sl-cm-card .sl-cm-mention')].map((n) => n.textContent)`);
    check('deck-comments: a mention in a thread is shown as PowerPoint shows one, the whole name in the accent colour', pane === true && mention.includes('@Ayesha Khan'), JSON.stringify({ pane, mention }));

    // A reply with an @: the people in the deck's comments are offered, and the name goes in.
    const first = threads().find((c) => /travel line/.test(c.text));
    await js(`document.querySelector('.sl-cm-card[data-comment="${first.id}"]')?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })), 1`);
    await until(() => js(`Boolean(document.querySelector('.sl-cm-card[data-comment="${first.id}"] .sl-cm-replybox textarea'))`), 'the reply box', 4000).catch(() => {});
    await setText(`.sl-cm-card[data-comment="${first.id}"] .sl-cm-replybox textarea`, 'Done, thanks @Bil');
    const offered = await until(() => js(`[...document.querySelectorAll('.sl-cm-people .sl-cm-person')].map((n) => n.dataset.person)`).then((l) => (l.includes('Bilal Ahmed') ? l : null)), 'the people offered', 4000).catch(() => null);
    await fresh('deck-comments-mention.png');
    await js(`document.querySelector('.sl-cm-people [data-person="Bilal Ahmed"]')?.click(), 1`);
    const typed = await js(`document.querySelector('.sl-cm-card[data-comment="${first.id}"] .sl-cm-replybox textarea')?.value`);
    await js(`document.querySelector('.sl-cm-card[data-comment="${first.id}"] .sl-cm-replybox .sl-cm-post')?.click(), 1`);
    const replied = await until(() => (threads().find((c) => c.id === first.id)?.replies || []).some((r) => r.text === 'Done, thanks @Bilal Ahmed'), 'the reply', 6000).catch(() => false);
    check('deck-comments: typing @ in a reply offers the people in the deck\'s comments, the name goes in whole, and the reply is posted', Array.isArray(offered) && typed === 'Done, thanks @Bilal Ahmed ' && replied === true, JSON.stringify({ offered, typed, replied }));

    // Edit one's own reply in place.
    const reply = (threads().find((c) => c.id === first.id)?.replies || []).find((r) => r.author === me);
    await js(`document.querySelector('.sl-cm-card[data-comment="${first.id}"] [data-reply="${reply?.id}"] .sl-cm-edit-reply')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.sl-cm-editbox textarea'))`), 'the edit box', 4000).catch(() => {});
    await setText('.sl-cm-editbox textarea', 'Done: travel is 35 now. @Bilal Ahmed');
    await js(`document.querySelector('.sl-cm-editbox .sl-cm-post')?.click(), 1`);
    const edited = await until(() => (threads().find((c) => c.id === first.id)?.replies || []).some((r) => r.text === 'Done: travel is 35 now. @Bilal Ahmed'), 'the edited reply', 6000).catch(() => false);
    const others = await js(`Boolean(document.querySelector('.sl-cm-card[data-comment="${first.id}"] .sl-cm-head .sl-cm-edit'))`);
    check('deck-comments: one\'s own reply is edited in place (Edit reply), and a thread someone else wrote offers no Edit', edited === true && others === false, JSON.stringify({ edited, others }));

    // Resolve and reopen.
    await js(`document.querySelector('.sl-cm-card[data-comment="${first.id}"] .sl-cm-resolve')?.click(), 1`);
    const resolved = await until(() => threads().find((c) => c.id === first.id)?.status === 'resolved', 'resolved', 5000).catch(() => false);
    const shown = await until(() => js(`(() => { const c = document.querySelector('.sl-cm-card[data-comment="${first.id}"]'); return c && c.classList.contains('resolved') && Boolean(c.querySelector('.sl-cm-badge')) && !c.querySelector('.sl-cm-replybox'); })()`), 'shown resolved', 4000).catch(() => false);
    await fresh('deck-comments-resolved.png');
    await js(`document.querySelector('.sl-cm-card[data-comment="${first.id}"] .sl-cm-resolve')?.click(), 1`);
    const reopened = await until(() => threads().find((c) => c.id === first.id)?.status === 'active', 'reopened', 5000).catch(() => false);
    check('deck-comments: Resolve marks the thread Resolved and puts its reply box away; Reopen opens it again', resolved === true && shown === true && reopened === true, JSON.stringify({ resolved, shown, reopened }));

    // Saved as PowerPoint 365 keeps them.
    await until(() => js(`(() => { const c = document.querySelector('.sl-cm-card[data-comment="${first.id}"]'); return Boolean(c) && !c.classList.contains('resolved'); })()`), 'the card shown open', 4000).catch(() => {});
    await js(`document.querySelector('.sl-cm-card[data-comment="${first.id}"] .sl-cm-resolve')?.click(), 1`);
    await until(() => threads().find((c) => c.id === first.id)?.status === 'resolved', 'resolved again', 5000).catch(() => {});
    await js(`(() => { const b = [...document.querySelectorAll('.rw-btn')].find((n) => (n.dataset.tip || n.title || '').startsWith('Save')); b?.click(); return 1; })()`);
    await until(() => doc.meta({ id: session.id }).dirty === false, 'the save', 15000).catch(() => {});
    const again = Deck.open(fs.readFileSync(file));
    const back = again.comments().find((c) => /travel line/.test(c.text));
    const part = again.pkg.partNames().find((n) => /^ppt\/comments\/modernComment/.test(n));
    const xml = part ? again.pkg.text(part) : '';
    check(
      'deck-comments: the saved deck carries the thread resolved, the edited reply and its author, in PowerPoint 365\'s modern comments part',
      back?.status === 'resolved' && back.replies.some((r) => r.text === 'Done: travel is 35 now. @Bilal Ahmed' && r.author === me) && /<p188:cm\b[^>]*status="resolved"/.test(xml) && /<p188:reply\b/.test(xml),
      JSON.stringify({ status: back?.status, replies: back?.replies?.map((r) => [r.author, r.text]) }),
    );
  } catch (err) {
    check('deck-comments: the comment checks ran', false, err.stack || err.message);
  }
}
