// Word: Mailings → Rules → Fill-in and Ask, driven through the ribbon — a
// letter with a list puts a Fill-in after "Seat: " and an Ask (kept in a
// bookmark, shown straight after it) after "See you at "; Finish & Merge →
// Edit Individual Documents asks the merge's questions first — the Ask once,
// the Fill-in for each record — and each letter carries its answers. Run
// alone with RUTBA_VERIFY_ONLY=mergerules.
import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { CHANNEL_PREFIX } from '@rutba/office-shell/contract';

const CSV = 'First Name,City\r\nJoshua,London\r\nCynthia,Leeds\r\nAmira,Paris\r\n';

export async function verifyWordMergeRules({ open, check, until, wait, errorsIn, doc, sessionFor }, { dir }) {
  try {
    const letter = path.join(dir, 'merge-rules.docx');
    const list = path.join(dir, 'merge-rules.csv');
    fs.writeFileSync(letter, buildDocx({ styles: true, paragraphs: [{ text: 'Seat: ' }, { text: 'See you at ' }, { text: 'Yours, the Rutba team' }] }));
    fs.writeFileSync(list, CSV, 'utf8');
    const win = await open('word', letter);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = sessionFor('doc');
    const model = () => doc.model({ id: session.id });
    const clickRibbon = (title) => js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}) && !n.disabled); if (!b) return 'no button'; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked'; })()`);
    const clickMenu = async (label) => {
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}))`), `the menu item ${label}`, 3000).catch(() => {});
      return js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(label)} && !n.disabled); if (!b) return 'no item'; b.click(); return 'clicked'; })()`);
    };
    const type = (selector, value) => js(`(() => { const el = document.querySelector(${JSON.stringify(selector)}); if (!el) return false; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, ${JSON.stringify(value)}); el.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
    const blockText = (i) => (model().blocks?.[i]?.text ?? '').replace(/​/g, '');
    // A real click just after a paragraph's last word, as a person puts the caret there.
    const caretAtEnd = async (block) => {
      const at = await js(`(() => {
        const b = document.querySelector('.wd-page [data-block="${block}"]');
        if (!b) return null;
        b.scrollIntoView({ block: 'center' });
        const walker = document.createTreeWalker(b, NodeFilter.SHOW_TEXT);
        let last = null;
        for (let n = walker.nextNode(); n; n = walker.nextNode()) if (n.nodeValue.length) last = n;
        if (!last) return null;
        const r = document.createRange();
        r.setStart(last, last.nodeValue.length - 1); r.setEnd(last, last.nodeValue.length);
        const rect = [...r.getClientRects()].pop();
        return { x: Math.round(rect.right + 2), y: Math.round(rect.top + rect.height / 2) };
      })()`);
      if (!at) return false;
      win.webContents.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      win.webContents.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      return until(() => model().selection?.focus?.block === block, `the caret in paragraph ${block}`, 5000).then(() => true).catch(() => false);
    };

    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="2"]'))`), 'the letter', 8000);
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Mailings')?.click(), 1`);
    await wait(250);
    await clickRibbon('Start Mail Merge');
    await clickMenu('Letters');
    await until(() => model().mailMerge?.type === 'formLetters', 'a form letter', 5000).catch(() => {});
    win.webContents.send(`${CHANNEL_PREFIX}/event/app:command`, { command: 'mailings.useList', args: { path: list } });
    await until(() => (model().mailMerge?.source?.count ?? 0) === 3, 'the list', 6000).catch(() => {});

    // Rules → Fill-in… after "Seat:".
    await caretAtEnd(0);
    await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Rules'); b?.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: 500, clientY: 110 })); return 1; })()`);
    const fillPicked = await clickMenu('Fill-in…');
    await until(() => js(`Boolean(document.querySelector('.wd-mm-prompt-text'))`), 'the Fill-in dialog', 3000).catch(() => {});
    await type('.wd-mm-prompt-text', 'Which seat?');
    await type('.wd-mm-prompt-default', 'any seat');
    await wait(150);
    await js(`document.querySelector('.wd-mm-prompt-ok')?.click(), 1`);
    await until(() => blockText(0).includes('any seat'), 'the Fill-in', 5000).catch(() => {});
    check('word: Mailings → Rules → Fill-in asks for a prompt and a default and puts the field at the caret, showing the default',
      fillPicked === 'clicked' && /^Seat:\s?any seat$/.test(blockText(0)), JSON.stringify({ fillPicked, text: blockText(0) }));

    // Rules → Ask… after "See you at", asked once, shown after it.
    await caretAtEnd(1);
    await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Rules'); b?.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: 500, clientY: 110 })); return 1; })()`);
    const askPicked = await clickMenu('Ask…');
    await until(() => js(`Boolean(document.querySelector('.wd-mm-prompt-name'))`), 'the Ask dialog', 3000).catch(() => {});
    await type('.wd-mm-prompt-name', 'Event');
    await type('.wd-mm-prompt-text', 'Which event?');
    await type('.wd-mm-prompt-default', 'the dinner');
    await js(`(() => { const box = [...document.querySelectorAll('.rw-dialog input[type=checkbox]')][0]; if (box && !box.checked) box.click(); return 1; })()`);
    await wait(150);
    await js(`document.querySelector('.wd-mm-prompt-ok')?.click(), 1`);
    await until(() => blockText(1).includes('«Ask Event»'), 'the Ask', 5000).catch(() => {});
    check('word: Rules → Ask puts an ASK at the caret and, after it, a REF that will show its answer', askPicked === 'clicked' && blockText(1).includes('«Ask Event»'), JSON.stringify({ askPicked, text: blockText(1) }));

    // Finish & Merge: the questions first.
    await clickRibbon('Finish & Merge');
    await clickMenu('Edit Individual Documents…');
    await until(() => js(`Boolean(document.querySelector('.rw-dialog[aria-label="Merge to New Document"]'))`), 'the merge dialog', 4000).catch(() => {});
    await js(`(() => { const b = [...document.querySelectorAll('.rw-dialog[aria-label="Merge to New Document"] .rw-dialog-foot button')].find((n) => n.textContent.trim() === 'OK'); b?.click(); return 1; })()`);
    const asked = await until(() => js(`Boolean(document.querySelector('.wd-mm-answers-ok'))`), 'the questions', 5000).then(() => true).catch(() => false);
    const shape = await js(`({ once: document.querySelectorAll('.wd-mm-answer-once').length, perRecord: document.querySelectorAll('.wd-mm-answer-record').length })`);
    await type('.wd-mm-answer-record[data-record="2"]', 'B2');
    await type('.wd-mm-answer-once', 'the summer party');
    await wait(150);
    await js(`document.querySelector('.wd-mm-answers-ok')?.click(), 1`);
    let merged = null;
    await until(() => { merged = doc.sessions().find((s) => s.kind === 'doc' && /^Letters\d+$/.test(s.name)); return Boolean(merged); }, 'the merged document', 8000).catch(() => {});
    const texts = merged ? (doc.model({ id: merged.id }).blocks || []).map((b) => (b.text || '').replace(/​/g, '')).filter((t) => /^(Seat|See you)/.test(t)) : [];
    check('word: Finish & Merge asks the merge\'s questions first — the Ask once, the Fill-in for each record — and each letter carries its answers',
      asked && shape.once === 1 && shape.perRecord === 3
        && JSON.stringify(texts) === JSON.stringify(['Seat: any seat', 'See you at the summer party', 'Seat: B2', 'See you at the summer party', 'Seat: any seat', 'See you at the summer party']),
      JSON.stringify({ asked, shape, texts }));

    const complaints = await errorsIn(win);
    check('word: the merge rules report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the merge rule checks ran', false, err?.message || JSON.stringify(err));
  }
}
