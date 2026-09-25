// Word: Review → Restrict Editing — the pane, the protection written as Word
// writes it, and the window really enforcing it.
//
// A four-paragraph agreement: the pane limits editing to No changes, the
// two form lines are marked for Everyone, and Yes, Start Enforcing
// Protection takes a password. Then the page refuses a keystroke in the
// locked text and takes one in a region, highlights the regions, and Find
// Next Region I Can Edit selects one; the saved settings.xml carries Word's
// w:documentProtection with the SHA-512 hash; Stop Protection refuses a
// wrong password and takes the right one. Tracked changes mode turns
// recording on and greys the button; comments mode refuses typing; and a
// file already protected (written as Word writes it) opens enforced with
// the pane up. Run alone with RUTBA_VERIFY_ONLY=restrict.
import fs from 'node:fs';
import path from 'node:path';
import { OoxmlPackage } from '@rutba/ooxml';
import { buildDocx } from '@rutba/ooxml/build';
import { wordProtectionAttrs } from '@rutba/ooxml/protection';

const PARAGRAPHS = ['Service agreement — the terms below are fixed.', 'Client name: ', 'Signed on: ', 'Standard clauses follow and may not be changed.'];
const PASSWORD = 'Stamp 9';

/** A document with a settings part, as Word always writes one. */
function agreement() {
  const pkg = OoxmlPackage.read(buildDocx({ paragraphs: PARAGRAPHS, styles: true }));
  if (!pkg.has('word/settings.xml')) {
    pkg.addPart('word/settings.xml', Buffer.from('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/></w:settings>', 'utf8'), 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');
    pkg.addRelationshipTo('word/document.xml', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings', 'settings.xml');
  }
  return pkg;
}

/**
 * @param {object} h the harness: open, check, until, wait, capture, doc
 */
export async function verifyWordRestrict(h, { dir }) {
  const { open, check, until, wait, capture, doc } = h;
  const file = path.join(dir, 'agreement.docx');
  const protectedFile = path.join(dir, 'agreement-protected.docx');
  try {
    fs.writeFileSync(file, agreement().write());
    // The same agreement as Word leaves it protected: read only, a password,
    // the signature line open to Everyone.
    const pkg = agreement();
    const a = wordProtectionAttrs('Contract');
    pkg.write_('word/settings.xml', pkg.text('word/settings.xml').replace('<w:defaultTabStop', `<w:documentProtection w:edit="readOnly" w:enforcement="1" ${Object.entries(a).map(([k, v]) => `${k}="${v}"`).join(' ')}/><w:defaultTabStop`));
    pkg.write_('word/document.xml', pkg.text('word/document.xml').replace(/(<w:p\b[^>]*>(?:<w:pPr>[\s\S]*?<\/w:pPr>)?)(<w:r>(?:(?!<\/w:p>)[\s\S])*?Signed on: (?:(?!<\/w:p>)[\s\S])*?)(<\/w:p>)/, '$1<w:permStart w:id="773" w:edGrp="everyone"/>$2<w:permEnd w:id="773"/>$3'));
    fs.writeFileSync(protectedFile, pkg.write());
  } catch (err) {
    return check('restrict: the fixtures were made', false, err.message);
  }

  try {
    const win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const session = doc.sessions().filter((s) => s.kind === 'doc' && s.path === file).pop();
    const model = () => doc.model({ id: session.id });
    const fresh = async (name) => {
      wc.invalidate();
      await wait(700);
      await capture(win, name);
    };
    const clickRibbon = (title) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}));
      if (!b) return 'no button ' + ${JSON.stringify(title)};
      if (b.disabled) return 'disabled';
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      b.click();
      return 'clicked';
    })()`);
    const clickTab = (name) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(), 'tab'`);
    const click = (selector) => js(`(() => { const b = document.querySelector(${JSON.stringify(selector)}); if (!b) return 'none'; if (b.disabled) return 'disabled'; b.click(); return 'clicked'; })()`);
    /** A real click in a paragraph — at its first words, or past its last. */
    const caretTo = async (block, atEnd = false) => {
      const at = await js(`(() => { const b = document.querySelector('.wd-page [data-block="${block}"]'); if (!b) return null; b.scrollIntoView({ block: 'center' }); const q = b.getBoundingClientRect(); const r = document.createRange(); r.selectNodeContents(b); const t = r.getBoundingClientRect(); return { x: Math.round(${atEnd ? 't.right + 6' : 'q.left + 4'}), y: Math.round(q.top + q.height / 2) }; })()`);
      if (!at) return false;
      wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      wc.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      await wait(350);
      return true;
    };
    const type = async (text) => {
      for (const ch of text) {
        wc.sendInputEvent({ type: 'keyDown', keyCode: ch });
        wc.sendInputEvent({ type: 'char', keyCode: ch });
        wc.sendInputEvent({ type: 'keyUp', keyCode: ch });
        await wait(70);
      }
    };
    const textOf = (i) => js(`(() => { const b = document.querySelector('.wd-page [data-block="${i}"]'); return b ? b.textContent : null; })()`);
    const toasts = () => js(`[...document.querySelectorAll('.rw-toast')].map((n) => n.textContent).join(' | ')`);
    const typeInto = async (selector, text) => {
      await js(`(() => { const i = document.querySelector(${JSON.stringify(selector)}); if (!i) return false; i.focus(); i.select?.(); return true; })()`);
      await wait(80);
      wc.insertText(text);
      await wait(150);
    };

    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="3"]'))`), 'the agreement', 10000);

    // 1. Review → Restrict Editing opens the pane with Word's three steps.
    await clickTab('Review');
    await wait(250);
    const pressed = await clickRibbon('Restrict Editing');
    const pane = await until(() => js(`Boolean(document.querySelector('.wd-restrict .wd-restrict-start'))`), 'the Restrict Editing pane', 5000).catch(() => false);
    const steps = await js(`[...document.querySelectorAll('.wd-restrict-step h4')].map((h) => h.textContent.replace(/^\\d/, ''))`);
    const kinds = await js(`[...document.querySelectorAll('.wd-restrict-kind option')].map((o) => o.textContent)`);
    check(
      'restrict: Review → Restrict Editing opens the pane — formatting restrictions, editing restrictions (No changes, Tracked changes, Comments, Filling in forms), start enforcement',
      pressed === 'clicked' && pane === true && steps.join('|') === 'Formatting restrictions|Editing restrictions|Start enforcement' && kinds.join('|') === 'No changes (Read only)|Tracked changes|Comments|Filling in forms',
      JSON.stringify({ pressed, steps, kinds })
    );

    // 2. Only No changes, and the two form lines marked for Everyone.
    await click('.wd-restrict-limit');
    await until(() => js(`Boolean(document.querySelector('.wd-restrict-everyone'))`), 'the Exceptions', 3000).catch(() => {});
    await caretTo(1);
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'Down', modifiers: ['shift'] });
    wc.sendInputEvent({ type: 'keyUp', keyCode: 'Down', modifiers: ['shift'] });
    await wait(500);
    await until(() => (model().selection?.focus?.block ?? 0) === 2, 'the selection to reach the second form line', 3000).catch(() => {});
    await click('.wd-restrict-everyone');
    const marked = await until(() => (model().protection?.regions || []).length === 1, 'the region', 5000).catch(() => false);
    const region = (model().protection?.regions || [])[0] || null;
    const washed = await until(() => js(`document.querySelectorAll('.wd-page.wd-perm-show .wd-block.wd-perm').length >= 1`), 'the highlight', 4000).catch(() => false);
    check(
      'restrict: Exceptions → Everyone marks the selected paragraphs with w:permStart/w:permEnd, highlighted on the page',
      marked === true && region?.from === 1 && region?.to >= 1 && washed === true && (await js(`document.querySelector('.wd-restrict-everyone')?.checked`)) === true,
      JSON.stringify({ region, washed })
    );
    await fresh('word-restrict-pane.png');

    // 3. Yes, Start Enforcing Protection, with a password typed twice.
    await click('.wd-restrict-start');
    const asked = await until(() => js(`Boolean(document.querySelector('.rw-dialog[aria-label="Start Enforcing Protection"] .wd-enforce-password'))`), 'the password dialog', 4000).catch(() => false);
    await typeInto('.wd-enforce-password', PASSWORD);
    await typeInto('.wd-enforce-again', 'Stamp 8');
    const mismatch = await until(() => js(`document.querySelector('.wd-enforce-ok')?.disabled === true`), 'OK greyed', 2000).catch(() => false);
    await typeInto('.wd-enforce-again', PASSWORD);
    await fresh('word-restrict-enforce.png');
    await click('.wd-enforce-ok');
    const enforced = await until(() => model().protection?.enforced === true, 'protection to be enforced', 20000).catch(() => false);
    const p = model().protection || {};
    check(
      'restrict: Yes, Start Enforcing Protection asks for an optional password twice and enforces read-only with it',
      asked === true && mismatch === true && enforced === true && p.edit === 'readOnly' && p.hasPassword === true,
      JSON.stringify({ asked, mismatch, edit: p.edit, hasPassword: p.hasPassword })
    );
    const perms = await until(() => js(`/You may only view this region/.test(document.querySelector('.wd-restrict-perm')?.textContent || '') || /You may edit in this region/.test(document.querySelector('.wd-restrict-perm')?.textContent || '')`), 'Your permissions', 4000).catch(() => false);
    check('restrict: once enforced the pane says what may be done here, with Find Next Region I Can Edit and Stop Protection', perms === true && (await js(`Boolean(document.querySelector('.wd-restrict-next')) && Boolean(document.querySelector('.wd-restrict-stop'))`)));

    // 4. The page enforces it: locked text refuses a keystroke, a region takes one.
    const before0 = await textOf(0);
    await caretTo(0);
    await type('X');
    const refusal = await until(async () => /not allowed because the selection is locked/.test(await toasts()), 'the refusal', 4000).catch(() => false);
    await wait(300);
    const after0 = await textOf(0);
    await caretTo(1, true);
    await type('Ada');
    const typed = await until(async () => /Client name: Ada/.test((await textOf(1)) || ''), 'the typing in the region', 6000).catch(() => false);
    check(
      'restrict: read-only refuses a keystroke in locked text with Word\'s sentence, and takes one inside a region anyone may edit',
      refusal === true && after0 === before0 && typed === true,
      JSON.stringify({ refusal, before0, after0, one: await textOf(1) })
    );

    // 5. Find Next Region I Can Edit: from the locked first paragraph, the region.
    await caretTo(0);
    await click('.wd-restrict-next');
    const found = await until(() => (model().selection?.anchor?.block ?? -1) === 1, 'the next region', 4000).catch(() => false);
    check('restrict: Find Next Region I Can Edit selects the region after the caret', found === true, JSON.stringify(model().selection));
    await fresh('word-restrict-enforced.png');

    // 6. Saved: Word's element in settings.xml.
    await clickRibbon('Save');
    await until(() => doc.meta({ id: session.id }).dirty === false, 'the save', 15000).catch(() => false);
    const saved = OoxmlPackage.read(fs.readFileSync(file));
    const settings = saved.text('word/settings.xml');
    const body = saved.text('word/document.xml');
    check(
      'restrict: the saved file carries w:documentProtection as Word writes it — w:edit="readOnly", w:enforcement="1", SHA-512 (sid 14), 100,000 rounds, hash and salt — and the region as w:permStart/w:permEnd',
      /<w:documentProtection w:edit="readOnly" w:enforcement="1" w:cryptProviderType="rsaAES" w:cryptAlgorithmClass="hash" w:cryptAlgorithmType="typeAny" w:cryptAlgorithmSid="14" w:cryptSpinCount="100000" w:hash="[^"]{80,}" w:salt="[^"]{20,}"\/>/.test(settings) && /<w:permStart w:id="\d+" w:edGrp="everyone"\/>/.test(body) && /<w:permEnd w:id="\d+"\/>/.test(body),
      (/<w:documentProtection[^>]*>/.exec(settings) || ['none'])[0].slice(0, 200)
    );

    // 7. Stop Protection: the wrong password refused, the right one takes it off.
    await click('.wd-restrict-stop');
    await until(() => js(`Boolean(document.querySelector('.wd-unprotect-password'))`), 'Unprotect Document', 4000).catch(() => {});
    await typeInto('.wd-unprotect-password', 'stamp 9');
    await click('.wd-unprotect-ok');
    const wrong = await until(() => js(`/not right/.test(document.querySelector('.wd-unprotect-error')?.textContent || '')`), 'the wrong password', 20000).catch(() => false);
    const stillOn = model().protection?.enforced === true;
    await fresh('word-restrict-unprotect.png');
    await typeInto('.wd-unprotect-password', PASSWORD);
    await click('.wd-unprotect-ok');
    const off = await until(() => model().protection?.enforced === false, 'protection off', 20000).catch(() => false);
    check('restrict: Stop Protection asks for the password, refuses a wrong one, and takes the protection off with the right one', wrong === true && stillOn && off === true, JSON.stringify({ wrong, stillOn, off }));

    // 8. Tracked changes: recording forced on, the button greyed.
    await until(() => js(`Boolean(document.querySelector('.wd-restrict-kind'))`), 'the settings again', 4000).catch(() => {});
    await js(`(() => { const s = document.querySelector('.wd-restrict-kind'); const set = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set; set.call(s, 'trackedChanges'); s.dispatchEvent(new Event('change', { bubbles: true })); return s.value; })()`);
    await wait(200);
    await click('.wd-restrict-start');
    await until(() => js(`Boolean(document.querySelector('.wd-enforce-ok'))`), 'the password dialog', 4000).catch(() => {});
    await click('.wd-enforce-ok');
    const tracking = await until(() => model().protection?.lockedTracking === true && model().trackRevisions === true, 'tracked changes mode', 8000).catch(() => false);
    const greyed = await until(() => js(`[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || '').startsWith('Track Changes'))?.disabled === true`), 'the greyed button', 4000).catch(() => false);
    await caretTo(3, true);
    await type(' Z');
    const recorded = await until(() => Boolean(model().blocks?.[3]?.tracked), 'the recorded change', 6000).catch(() => false);
    check('restrict: tracked-changes protection turns recording on, greys Track Changes, and a keystroke is recorded as a tracked change', tracking === true && greyed === true && recorded === true, JSON.stringify({ tracking, greyed, recorded }));
    await fresh('word-restrict-tracked.png');
    await click('.wd-restrict-stop');
    await until(() => model().protection?.enforced === false, 'protection off', 8000).catch(() => {});

    // 9. Comments: typing refused, the sentence says comments only.
    await js(`(() => { const s = document.querySelector('.wd-restrict-kind'); const set = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set; set.call(s, 'comments'); s.dispatchEvent(new Event('change', { bubbles: true })); return s.value; })()`);
    await wait(200);
    await click('.wd-restrict-start');
    await until(() => js(`Boolean(document.querySelector('.wd-enforce-ok'))`), 'the password dialog', 4000).catch(() => {});
    await click('.wd-enforce-ok');
    await until(() => model().protection?.edit === 'comments' && model().protection?.enforced === true, 'comments mode', 8000).catch(() => {});
    const three = await textOf(3);
    await caretTo(3);
    await type('Q');
    const onlyComments = await until(async () => /only comments/.test(await toasts()), 'the comments refusal', 4000).catch(() => false);
    await wait(300);
    check('restrict: comments protection refuses typing in the text and says only comments can be added', onlyComments === true && (await textOf(3)) === three, JSON.stringify({ onlyComments }));
    await click('.wd-restrict-stop');
    await until(() => model().protection?.enforced === false, 'protection off', 8000).catch(() => {});

    // The dark theme, for the capture only.
    await js(`document.documentElement.dataset.theme = 'dark'`);
    await fresh('word-restrict-pane-dark.png');
    await js(`document.documentElement.dataset.theme = 'light'`);
  } catch (err) {
    check('restrict: the Restrict Editing checks ran', false, err.message);
  }

  // 10. A file Word protected opens enforced, the pane up and the region marked.
  try {
    const win = await open('word', protectedFile);
    const js = (code) => win.webContents.executeJavaScript(code);
    const up = await until(() => js(`/protected from unintentional editing/.test(document.querySelector('.wd-restrict-perm')?.textContent || '')`), 'the pane on open', 10000).catch(() => false);
    const marked = await until(() => js(`document.querySelectorAll('.wd-page.wd-perm-show .wd-block.wd-perm').length === 1`), 'the region highlighted', 5000).catch(() => false);
    const session = doc.sessions().filter((s) => s.kind === 'doc' && s.path === protectedFile).pop();
    const p = session ? doc.model({ id: session.id }).protection : null;
    check(
      'restrict: a document Word protected opens enforced — the Restrict Editing pane up with "Your permissions", its one region highlighted',
      up === true && marked === true && p?.enforced === true && p?.hasPassword === true && p?.regions?.[0]?.from === 2,
      JSON.stringify({ up, marked, protection: p && { edit: p.edit, regions: p.regions } })
    );
    win.webContents.invalidate();
    await wait(700);
    await capture(win, 'word-restrict-opened.png');
  } catch (err) {
    check('restrict: the protected-file check ran', false, err.message);
  }
}
