// Word: Mailings → mail merge, driven through the ribbon.
//
// A letter is started as a mail merge, given a recipient list from a .csv
// (handed to the window the way the OS hands it a file — the Select Data
// Source dialog is the system's own and a check never touches the desktop),
// and gets an Address Block, a Greeting Line and a merge field from the
// ribbon at a caret placed by a real click. Preview Results then shows record
// 2 on the page; Finish & Merge → Edit Individual Documents opens Letters1 in
// a window of its own with a section per record; Send E-mail Messages sends
// one message per record through the Mail service — under the checks, its
// fake transport — and the Sent folder holds them. Run alone with
// RUTBA_VERIFY_ONLY=mailmerge.
import fs from 'node:fs';
import path from 'node:path';
import { BrowserWindow } from 'electron';
import { buildDocx } from '@rutba/ooxml/build';
import { CHANNEL_PREFIX } from '@rutba/office-shell/contract';

const CSV = '\uFEFFTitle,First Name,Last Name,Company Name,Address Line 1,Address Line 2,City,State,ZIP Code,Country or Region,E-mail Address\r\n'
  + 'Mr.,Joshua,Randall,,"12 High Street, Flat 2",,London,,SW1A 1AA,United Kingdom,josh@example.com\r\n'
  + 'Ms.,Cynthia,Gartner,Contoso Ltd,1 Main Road,Suite 5,Leeds,West Yorkshire,LS1 1AA,United Kingdom,cynthia@example.com\r\n'
  + 'Dr.,Amira,Haddad,,4 Rue Neuve,,Paris,,75001,France,amira@example.com\r\n';

/** The letter and its list, this check's own, in the run's folder. */
export function makeMailMergeFixture(dir) {
  const letter = path.join(dir, 'merge-letter.docx');
  fs.writeFileSync(letter, buildDocx({
    styles: true,
    paragraphs: [
      { text: '' },
      { text: '' },
      { text: 'Your order is on its way, ' },
      { text: 'Thank you for shopping with us.' },
      { text: 'Yours sincerely,' },
      { text: 'The Rutba team' },
    ],
  }));
  const list = path.join(dir, 'merge-recipients.csv');
  fs.writeFileSync(list, CSV, 'utf8');
  return { letter, list };
}

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn, capture, doc, sessionFor
 */
export async function verifyWordMailMerge(h, { dir }) {
  const { open, check, until, wait, errorsIn, capture, doc, sessionFor } = h;
  const SUBJECT = `Verify mail merge ${Date.now()}`;
  let win = null;
  let accountId = null;
  try {
    const { letter, list } = makeMailMergeFixture(dir);
    win = await open('word', letter);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = sessionFor('doc');
    const model = () => doc.model({ id: session.id });

    const clickRibbon = (title) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}) && !n.disabled);
      if (!b) return 'no button ' + ${JSON.stringify(title)};
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      b.click();
      return 'clicked';
    })()`);
    const clickMenu = (label) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(label)} && !n.disabled);
      if (!b) return 'no menu item ' + ${JSON.stringify(label)};
      b.click();
      return 'clicked';
    })()`);
    const clickDialog = (dialog, label) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-dialog[aria-label=${JSON.stringify(dialog)}] .rw-dialog-foot button')].find((n) => n.textContent.trim() === ${JSON.stringify(label)} && !n.disabled);
      if (!b) return 'no ' + ${JSON.stringify(label)} + ' in ' + ${JSON.stringify(dialog)};
      b.click();
      return 'clicked';
    })()`);
    const clickTab = (name) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(), 'tab'`);
    const blockText = (i) => js(`(() => { const b = document.querySelector('.wd-page [data-block="${i}"]'); return b ? b.innerText.replace(/\\u200b/g, '') : null; })()`);
    // A real click at the start or the end of a paragraph's words: a caret
    // placed by script stops being believed once the page has been edited.
    const caretAt = async (block, where) => {
      const at = await js(`(() => {
        const b = document.querySelector('.wd-page [data-block="${block}"]');
        if (!b) return null;
        b.scrollIntoView({ block: 'center' });
        // The last character's own box: a range collapsed at the end of an
        // element has no rectangle to click beside.
        const r = document.createRange();
        const walker = document.createTreeWalker(b, NodeFilter.SHOW_TEXT);
        let last = null;
        for (let n = walker.nextNode(); n; n = walker.nextNode()) if (n.nodeValue.length) last = n;
        if (${where === 'end' ? 'true' : 'false'} && last) { r.setStart(last, last.nodeValue.length - 1); r.setEnd(last, last.nodeValue.length); }
        else { r.selectNodeContents(b); r.collapse(true); }
        const rects = [...r.getClientRects()];
        const box = b.getBoundingClientRect();
        const rect = rects.length ? rects[rects.length - 1] : null;
        const x = rect && rect.width + rect.height > 0 ? rect.right + ${where === 'end' ? 2 : 1} : box.left + 3;
        const y = rect && rect.height > 0 ? rect.top + rect.height / 2 : box.top + Math.min(box.height, 20) / 2;
        return { x: Math.round(x), y: Math.round(y) };
      })()`);
      if (!at) return false;
      win.webContents.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      win.webContents.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      return until(() => model().selection?.focus?.block === block, `the caret in paragraph ${block}`, 5000).catch(() => false);
    };

    // An off-screen window hands back the last frame it painted: ask for a
    // fresh one, twice, the first capture only prompting the paint.
    const snap = async (name) => {
      win.webContents.invalidate();
      await wait(700);
      await win.webContents.capturePage().catch(() => null);
      win.webContents.invalidate();
      await wait(600);
      await capture(win, name);
    };

    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="2"]'))`), 'the letter', 8000);

    // Start Mail Merge → Letters.
    await clickTab('Mailings');
    await wait(250);
    const startPressed = await clickRibbon('Start Mail Merge');
    await until(() => js(`Boolean(document.querySelector('.rw-menu'))`), 'the Start Mail Merge menu', 3000).catch(() => {});
    const letters = await clickMenu('Letters');
    const started = await until(() => model().mailMerge?.type === 'formLetters', 'the letter to be a merge letter', 5000).catch(() => false);
    check('word: Mailings → Start Mail Merge → Letters makes the document a form letter', startPressed === 'clicked' && letters === 'clicked' && started === true, `${startPressed}; ${letters}; type ${model().mailMerge?.type}`);

    // Select Recipients: the list, handed over by path, through the same door.
    win.webContents.send(`${CHANNEL_PREFIX}/event/app:command`, { command: 'mailings.useList', args: { path: list } });
    const attached = await until(() => model().mailMerge?.source?.count === 3, 'the list to be attached', 8000).catch(() => false);
    const ribbonLive = await until(() => js(`Boolean([...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || '').startsWith('Address Block') && !n.disabled))`), 'Address Block to come alive', 5000).catch(() => false);
    const fields = model().mailMerge?.source?.fields || [];
    check(
      'word: Select Recipients reads the .csv — byte-order mark gone, quoted comma kept — and the Write & Insert Fields buttons come alive',
      attached === true && ribbonLive === true && fields[0] === 'Title' && fields.includes('E-mail Address'),
      JSON.stringify({ count: model().mailMerge?.source?.count, first: fields[0], live: ribbonLive })
    );

    // Edit Recipient List: a table of the three, each ticked.
    await clickRibbon('Edit Recipient List');
    const listed = await until(() => js(`document.querySelectorAll('.rw-dialog[aria-label="Mail Merge Recipients"] tbody tr').length === 3`), 'the recipient list', 5000).catch(() => false);
    const table = await js(`(() => {
      const d = document.querySelector('.rw-dialog[aria-label="Mail Merge Recipients"]');
      if (!d) return null;
      const rows = [...d.querySelectorAll('tbody tr')];
      return { rows: rows.length, ticked: rows.filter((r) => r.querySelector('input[type=checkbox]')?.checked).length, street: rows[0]?.children[5]?.textContent || '', heads: [...d.querySelectorAll('thead th')].length };
    })()`);
    check(
      'word: Edit Recipient List shows every record with a tick and every column',
      listed === true && table?.ticked === 3 && table?.street === '12 High Street, Flat 2' && table?.heads === 12,
      JSON.stringify(table)
    );
    await snap('word-mailmerge-recipients.png');
    await clickDialog('Mail Merge Recipients', 'OK');
    await until(() => js(`!document.querySelector('.rw-dialog[aria-label="Mail Merge Recipients"]')`), 'the list to close', 4000).catch(() => {});

    // Address Block in the first paragraph.
    const caret0 = await caretAt(0, 'start');
    await clickRibbon('Address Block');
    await until(() => js(`(document.querySelector('.wd-mm-address-preview')?.innerText || '').includes('Joshua')`), 'the address block preview', 5000).catch(() => {});
    const dialogPreview = await js(`document.querySelector('.wd-mm-address-preview')?.innerText || ''`);
    await snap('word-mailmerge-addressblock.png');
    await clickDialog('Insert Address Block', 'OK');
    const address = await until(() => (model().blocks?.[0]?.runs || []).some((r) => r.field?.kind === 'addressblock'), 'the address block field', 5000).catch(() => false);

    // Greeting Line in the second.
    const caret1 = await caretAt(1, 'start');
    await clickRibbon('Greeting Line');
    await until(() => js(`Boolean(document.querySelector('.rw-dialog[aria-label="Insert Greeting Line"]'))`), 'the greeting dialog', 4000).catch(() => {});
    await clickDialog('Insert Greeting Line', 'OK');
    const greeting = await until(() => (model().blocks?.[1]?.runs || []).some((r) => r.field?.kind === 'greetingline'), 'the greeting line field', 5000).catch(() => false);

    // «First_Name» at the end of the third.
    const caret2 = await caretAt(2, 'end');
    await clickRibbon('Insert Merge Field');
    await until(() => js(`Boolean(document.querySelector('.rw-menu'))`), 'the field menu', 3000).catch(() => {});
    const picked = await clickMenu('First Name');
    const field = await until(() => (model().blocks?.[2]?.runs || []).some((r) => r.field?.kind === 'mergefield'), 'the merge field', 5000).catch(() => false);
    const onPage = await until(async () => (await blockText(2)) === 'Your order is on its way, «First_Name»', 'the field drawn on the page', 5000).catch(() => false);
    check(
      'word: Address Block, Greeting Line and Insert Merge Field put Word\'s fields at the caret, and the page shows «First_Name»',
      caret0 !== false && caret1 !== false && caret2 !== false && address === true && greeting === true && field === true && picked === 'clicked' && onPage === true
        && /Mr\. Joshua Randall/.test(dialogPreview),
      JSON.stringify({ caret0, caret1, caret2, address, greeting, field, picked, onPage, block2: await blockText(2), dialogPreview })
    );

    // Preview Results, then Next record: record 2's words on the page.
    await clickRibbon('Preview Results');
    await until(() => model().mailMerge?.preview === true, 'the preview', 5000).catch(() => {});
    await until(async () => (await blockText(1)) === 'Dear Mr. Randall,', 'record 1 on the page', 5000).catch(() => {});
    await clickRibbon('Next record');
    const shown = await until(async () => (await blockText(1)) === 'Dear Ms. Gartner,', 'record 2 on the page', 6000).catch(() => false);
    const page = { address: await blockText(0), greeting: await blockText(1), line: await blockText(2), box: await js(`document.querySelector('.wd-mm-record')?.value`) };
    check(
      'word: Preview Results → Next record shows record 2 in place of the fields — the address block without a blank line, the greeting, the name',
      shown === true && page.box === '2' && page.address === 'Ms. Cynthia Gartner\nContoso Ltd\n1 Main Road\nSuite 5\nLeeds, West Yorkshire LS1 1AA' && page.line === 'Your order is on its way, Cynthia',
      JSON.stringify(page)
    );
    await snap('word-mailmerge-preview.png');

    // Finish & Merge → Edit Individual Documents: Letters1, a section each.
    const before = new Set(BrowserWindow.getAllWindows().map((w) => w.id));
    await clickRibbon('Finish & Merge');
    await until(() => js(`Boolean(document.querySelector('.rw-menu'))`), 'the finish menu', 3000).catch(() => {});
    await clickMenu('Edit Individual Documents…');
    await until(() => js(`Boolean(document.querySelector('.rw-dialog[aria-label="Merge to New Document"]'))`), 'the merge dialog', 4000).catch(() => {});
    await clickDialog('Merge to New Document', 'OK');
    let merged = null;
    await until(() => {
      merged = doc.sessions().find((s) => s.kind === 'doc' && /^Letters\d+$/.test(s.name));
      return Boolean(merged);
    }, 'the merged document', 8000).catch(() => {});
    let newWin = null;
    await until(() => {
      newWin = BrowserWindow.getAllWindows().find((w) => !before.has(w.id) && !w.isDestroyed() && w.webContents.getURL().includes('session='));
      return Boolean(newWin);
    }, 'the Letters1 window', 8000).catch(() => {});
    let drawn = null;
    if (newWin) {
      const njs = (code) => newWin.webContents.executeJavaScript(code);
      await until(() => njs(`document.querySelectorAll('.wd-sheet').length >= 3 && document.querySelectorAll('.wd-page .wd-block').length >= 18`), 'Letters1 to lay out', 12000).catch(() => {});
      drawn = await njs(`({ sheets: document.querySelectorAll('.wd-sheet').length, greetings: [...document.querySelectorAll('.wd-page .wd-block')].map((b) => b.innerText).filter((t) => /^Dear /.test(t)), title: document.title })`).catch(() => null);
    }
    const mm = merged ? doc.model({ id: merged.id }) : null;
    check(
      'word: Finish & Merge → Edit Individual Documents opens Letters1 in its own window, a section per record, each letter on a page of its own',
      Boolean(merged) && mm?.sectionCount === 3 && Boolean(drawn) && drawn.sheets >= 3
        && JSON.stringify(drawn.greetings) === JSON.stringify(['Dear Mr. Randall,', 'Dear Ms. Gartner,', 'Dear Dr. Haddad,'])
        && (mm.blocks || []).filter((b) => b.pageBreakBefore).length === 2,
      JSON.stringify({ name: merged?.name, sections: mm?.sectionCount, drawn })
    );
    if (newWin) {
      const complaints = await errorsIn(newWin).catch(() => []);
      check('word: the merged letters\' window reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    }

    // Finish & Merge → Send E-mail Messages, through the fake transport.
    win.focus?.();
    await clickRibbon('Finish & Merge');
    await until(() => js(`Boolean(document.querySelector('.rw-menu'))`), 'the finish menu', 3000).catch(() => {});
    await clickMenu('Send E-mail Messages…');
    await until(() => js(`Boolean(document.querySelector('.wd-mm-account option[value]:not([value=""])'))`), 'the send dialog with an account', 6000).catch(() => {});
    accountId = await js(`document.querySelector('.wd-mm-account')?.value || null`);
    const toColumn = await js(`document.querySelector('.wd-mm-to')?.value || null`);
    await js(`(() => {
      const el = document.querySelector('.wd-mm-subject');
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, ${JSON.stringify(SUBJECT)});
      el.dispatchEvent(new Event('input', { bubbles: true }));
      return true;
    })()`);
    await until(() => js(`document.querySelector('.wd-mm-subject')?.value === ${JSON.stringify(SUBJECT)}`), 'the subject typed', 3000).catch(() => {});
    await clickDialog('Merge to E-mail', 'Send');
    const sentAll = await until(() => js(`(document.querySelector('.wd-mm-progress-text')?.textContent || '') === 'Sent 3 of 3.'`), 'the three messages to go', 12000).catch(() => false);
    const inSent = accountId ? await js(`(async () => {
      const rows = (await window.rutbaOffice.mail.messages({ accountId: ${JSON.stringify(accountId)}, folder: 'Sent', limit: 400 })).rows.filter((r) => r.subject === ${JSON.stringify(SUBJECT)});
      const one = rows.find((r) => /cynthia/.test(JSON.stringify(r.to || '')));
      const body = one ? await window.rutbaOffice.mail.message({ accountId: ${JSON.stringify(accountId)}, folder: 'Sent', id: one.id }) : null;
      return { count: rows.length, ids: rows.map((r) => r.id), to: rows.map((r) => JSON.stringify(r.to)).sort(), text: body?.text || '' };
    })()`) : null;
    check(
      'word: Send E-mail Messages sends one message per record to the E-mail Address column, the merged letter its body',
      sentAll === true && toColumn === 'E-mail Address' && inSent?.count === 3 && /Dear Ms\. Gartner,/.test(inSent.text) && /Your order is on its way, Cynthia/.test(inSent.text) && /Contoso Ltd/.test(inSent.text),
      JSON.stringify({ sentAll, toColumn, count: inSent?.count, to: inSent?.to, text: (inSent?.text || '').slice(0, 160) })
    );
    await snap('word-mailmerge-email.png');
    await clickDialog('Merge to E-mail', 'Close');
    // Sent goes back to what it was: other mail checks share the account.
    if (accountId && inSent?.ids?.length) await js(`window.rutbaOffice.mail.delete({ accountId: ${JSON.stringify(accountId)}, folder: 'Sent', ids: ${JSON.stringify(inSent.ids)} })`).catch(() => {});

    const complaints = await errorsIn(win);
    check('word: the mail merge checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    // The letter and its list are this check's own, made fresh in the run's
    // folder; nothing shared was touched except Sent, put back above.
  } catch (err) {
    check('word: the mail merge checks ran', false, err.message);
  }
}
