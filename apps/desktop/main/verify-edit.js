// Does typing actually work?
//
// The defect this exists to catch was invisible to every other kind of test:
// the editor was wired to React's `onBeforeInput`, which is a polyfill carrying
// no `inputType`, so every keystroke was cancelled and then ignored. The build
// was clean, the window rendered, the unit tests passed, and typing did
// nothing.
//
// So this drives the real window — real key events through Chromium's own input
// pipeline — and then asks the *engine*, not the DOM, what the document says.
// Checking the DOM would only prove the browser did what browsers do; checking
// the model proves the edit reached the thing that saves the file.

import { clipboard } from 'electron';

const wait = (ms) => new Promise((r) => setTimeout(r, ms));

/** Wait for a condition rather than for a guess at how long it takes. */
async function until(condition, what, timeout = 6000) {
  const deadline = Date.now() + timeout;
  for (;;) {
    let held = false;
    try {
      held = await condition();
    } catch {
      held = false;
    }
    if (held) return true;
    if (Date.now() > deadline) throw new Error(`waited ${timeout} ms for ${what} and it never happened`);
    await wait(80);
  }
}

/**
 * Chromium hands a key event only to a focused view, and on a busy machine
 * the page can lose it between two keys; every key after that is dropped
 * without a word. Each key gives the page its focus back first — inside its
 * own window only: taking the desktop's focus from whoever is at the machine
 * is something a check never does.
 */
function focusFor(wc) {
  try {
    wc.focus();
  } catch {
    // A window on its way out.
  }
}

/**
 * Press a key the way a keyboard does.
 *
 * Keys that produce a character need a "char" event between the down and the
 * up, or Chromium's editing pipeline never sees them — Return is one of those
 * and Backspace is not, which is a distinction worth writing down because it
 * costs an hour to rediscover.
 */
async function press(wc, keyCode, { modifiers = [], char = false } = {}) {
  focusFor(wc);
  wc.sendInputEvent({ type: 'keyDown', keyCode, modifiers });
  if (char) wc.sendInputEvent({ type: 'char', keyCode, modifiers });
  wc.sendInputEvent({ type: 'keyUp', keyCode, modifiers });
  await wait(120);
}

/** Type a printable character, which is a keyDown, a char, and a keyUp. */
async function typeChar(wc, ch) {
  focusFor(wc);
  wc.sendInputEvent({ type: 'keyDown', keyCode: ch });
  wc.sendInputEvent({ type: 'char', keyCode: ch });
  wc.sendInputEvent({ type: 'keyUp', keyCode: ch });
  await wait(70);
}

/** until(), but for a check: false when it never happens, so the check says what it holds. */
const settle = (condition, what, timeout = 8000) => until(condition, what, timeout).catch(() => false);

export async function verifyEditing({ windows, doc }) {
  const results = [];
  const opened = [];
  const check = (name, ok, detail = '') => {
    results.push({ name, ok, detail });
    console.log(`${ok ? 'ok  ' : 'FAIL'} ${name}${detail ? ` — ${detail}` : ''}`);
  };

  const openWindow = async (app) => {
    const win = windows.create({ app });
    await new Promise((resolve) => {
      win.webContents.once('did-finish-load', () => setTimeout(resolve, 1100));
    });
    win.focus();
    win.webContents.focus();
    // Keys reach only a page that believes it has focus, and the desktop's
    // focus is not this run's to keep: any program started beside it — a
    // console window, a build — can take it for a moment, and a key sent then
    // is dropped. Chromium's own focus emulation keeps the page focused
    // whatever the desktop does; the keys still go through the real input
    // pipeline, which is what this pass exists to prove.
    try {
      win.webContents.debugger.attach('1.3');
      await win.webContents.debugger.sendCommand('Emulation.setFocusEmulationEnabled', { enabled: true });
    } catch {
      // Without it the pass runs as before, needing the desktop's focus.
    }
    opened.push(win);
    return win;
  };

  /** The session the window opened, read straight from the document service. */
  const sessionFor = (kind) => doc.sessions().filter((s) => s.kind === kind).pop();

  /* ── Word ───────────────────────────────────────────────────────────── */

  const word = await openWindow('word');
  const blockText = (i) => doc.model({ id: sessionFor('doc')?.id })?.blocks?.[i]?.runs?.map((r) => r.text).join('') ?? '';
  // On a loaded machine the page re-renders a paragraph after an edit a
  // beat later than a fixed pause allows, and a key sent before focus and
  // caret are back goes nowhere: every check after it then failed together.
  // Each key goes to a focused page with the caret where the check means it.
  const caretAtEnd = () => word.webContents.executeJavaScript(`(() => {
    const page = document.querySelector('.wd-page');
    const block = page && page.querySelector('[data-block="0"]');
    if (!block) return false;
    page.focus();
    const range = document.createRange();
    range.selectNodeContents(block);
    range.collapse(false);
    const sel = getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    return document.activeElement === page;
  })()`);
  try {
    await settle(() => word.webContents.executeJavaScript(`Boolean(document.querySelector('.wd-page [data-block="0"]'))`), 'the page to draw', 12000);
    // Put the caret in the first paragraph the way a click would.
    await word.webContents.executeJavaScript(`(() => {
      const page = document.querySelector('.wd-page');
      const block = page && page.querySelector('[data-block="0"]');
      if (!block) return 'no block';
      page.focus();
      const range = document.createRange();
      range.selectNodeContents(block);
      range.collapse(true);
      const sel = getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      return 'ok';
    })()`);
    await wait(250);

    // insertText goes through Chromium's editing pipeline and fires a real
    // `beforeinput` with inputType 'insertText' — exactly what a keyboard does.
    word.webContents.insertText('Hello');
    await settle(() => blockText(0).startsWith('Hello'), 'the typed words in the engine');
    await wait(150);

    let session = sessionFor('doc');
    let model = session && doc.model({ id: session.id });
    const afterType = model?.blocks?.[0]?.runs?.map((r) => r.text).join('') ?? '';
    check('word: typing reaches the engine', afterType.startsWith('Hello'), `first block is ${JSON.stringify(afterType)}`);

    // Backspace never arrives through React's polyfill: keypress does not fire
    // for it. This is the half of the defect a text-only check would miss.
    word.webContents.focus();
    await caretAtEnd();
    await wait(120);
    await press(word.webContents, 'Backspace');
    await settle(() => blockText(0) === 'Hell', 'the backspace in the engine');
    session = sessionFor('doc');
    model = session && doc.model({ id: session.id });
    const afterDelete = model?.blocks?.[0]?.runs?.map((r) => r.text).join('') ?? '';
    check('word: backspace reaches the engine', afterDelete === 'Hell', `first block is ${JSON.stringify(afterDelete)}`);

    const blocksBefore = model?.blocks?.length ?? 0;

    // Record what the editing pipeline actually reports, so a failure here says
    // which half is wrong: the event not arriving, or the operation not landing.
    await word.webContents.executeJavaScript(`(() => {
      window.__seen = [];
      document.querySelector('.wd-page')
        ?.addEventListener('beforeinput', (e) => window.__seen.push(e.inputType), true);
      return 'listening';
    })()`);

    await caretAtEnd();
    await wait(120);
    await press(word.webContents, 'Return', { char: true });
    await settle(() => (doc.model({ id: sessionFor('doc').id })?.blocks?.length ?? 0) > blocksBefore, 'the paragraph split in the engine');
    const seen = await word.webContents.executeJavaScript('window.__seen || []');
    if (!seen.length) console.log('       (no beforeinput arrived for Return)');
    else console.log(`       (beforeinput saw: ${seen.join(', ')})`);
    session = sessionFor('doc');
    model = session && doc.model({ id: session.id });
    const blocksAfter = model?.blocks?.length ?? 0;
    check('word: Enter splits the paragraph', blocksAfter === blocksBefore + 1, `${blocksBefore} → ${blocksAfter} blocks`);

    // And the window shows it, which is the other half of the loop.
    const shownNow = () => word.webContents.executeJavaScript(
      `document.querySelector('.wd-page [data-block="0"]')?.textContent ?? ''`
    );
    await settle(() => shownNow().then((t) => t.trim() === 'Hell'), 'the page to show the engine');
    const shown = await shownNow();
    check('word: the page shows what the engine holds', shown.trim() === 'Hell', `page shows ${JSON.stringify(shown.trim())}`);

    // Bold. The ribbon says "bold" and the engine says "b"; when nothing
    // translated between them every press raised "unknown format: bold" and
    // the only thing that happened was a red notice.
    await word.webContents.executeJavaScript(`(() => {
      const block = document.querySelector('.wd-page [data-block="0"]');
      const range = document.createRange();
      range.selectNodeContents(block);
      const sel = getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      document.querySelector('.wd-page').dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
      return 'selected';
    })()`);
    await settle(() => doc.model({ id: sessionFor('doc').id })?.selection?.collapsed === false, 'the range in the engine');

    // The engine must have the *range*, not a caret; a collapsed selection here
    // means the sync lost it, and bold would silently apply to nothing.
    const before = doc.model({ id: sessionFor('doc').id });
    check(
      'word: selecting text reaches the engine as a range',
      before.selection && !before.selection.collapsed,
      `engine has ${JSON.stringify(before.selection)}`
    );

    await press(word.webContents, 'b', { modifiers: ['control'] });
    await settle(() => (doc.model({ id: sessionFor('doc').id })?.blocks?.[0]?.runs ?? []).some((r) => r.bold), 'bold in the engine');

    session = sessionFor('doc');
    model = session && doc.model({ id: session.id });
    const bolded = model?.blocks?.[0]?.runs ?? [];
    check(
      'word: bold reaches the engine',
      bolded.some((r) => r.bold),
      bolded.length ? `runs are ${JSON.stringify(bolded.map((r) => ({ text: r.text, bold: !!r.bold })))}` : 'no runs'
    );

    const complaints = await word.webContents.executeJavaScript(
      `[...document.querySelectorAll('.rw-toast.bad')].map((n) => n.textContent)`
    );
    check('word: bold raises no error', complaints.length === 0, complaints.join(' | ') || 'nothing was reported');

    // A selection that ends on the page itself, not in a paragraph — what a
    // drag into the margin, a triple-click or Ctrl+A leaves — must still
    // reach the engine as a range, or Delete takes one character and a paste
    // lands beside the selected words instead of over them (owner, 2026-09-24).
    // A second paragraph first, so the selection can run past the first.
    await word.webContents.executeJavaScript(`(() => {
      const page = document.querySelector('.wd-page');
      const block = page.querySelector('[data-block="0"]');
      const range = document.createRange();
      range.selectNodeContents(block);
      range.collapse(false);
      const sel = getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      page.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
      return 'ok';
    })()`);
    await wait(250);
    const blocksNow = doc.model({ id: sessionFor('doc').id })?.blocks?.length ?? 0;
    await press(word.webContents, 'Return', { char: true });
    await settle(() => (doc.model({ id: sessionFor('doc').id })?.blocks?.length ?? 0) > blocksNow, 'the second paragraph in the engine');
    word.webContents.insertText('World');
    await settle(() => blockText(1).includes('World'), 'the words of the second paragraph in the engine');
    const two = doc.model({ id: sessionFor('doc').id });
    const twoBlocks = two?.blocks?.length ?? 0;
    const selectedPast = await word.webContents.executeJavaScript(`(() => {
      const page = document.querySelector('.wd-page');
      const first = page.querySelector('[data-block="0"]');
      const second = page.querySelector('[data-block="1"]');
      if (!first || !second) return 'no second paragraph';
      const walker = document.createTreeWalker(first, NodeFilter.SHOW_TEXT);
      const text = walker.nextNode();
      if (!text) return 'no text';
      const range = document.createRange();
      range.setStart(text, 2);
      // The end on the page itself, after the second paragraph's element.
      const parent = second.parentNode;
      range.setEnd(parent, [...parent.childNodes].indexOf(second) + 1);
      const sel = getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      page.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
      return sel.focusNode === parent ? 'on the page' : 'in ' + (sel.focusNode.nodeName || '?');
    })()`);
    await settle(() => { const r = doc.model({ id: sessionFor('doc').id })?.selection; return r && !r.collapsed && r.to?.block >= 1; }, 'the range past the paragraph in the engine');
    const ranged = doc.model({ id: sessionFor('doc').id })?.selection;
    check('word: a selection that ends on the page itself reaches the engine as a range',
      twoBlocks >= 2 && ranged && !ranged.collapsed && ranged.from?.block === 0 && ranged.from?.offset === 2 && ranged.to?.block >= 1,
      `${twoBlocks} paragraphs; selection end ${selectedPast}; engine has ${JSON.stringify(ranged)}`);

    await press(word.webContents, 'Delete');
    await settle(() => (doc.model({ id: sessionFor('doc').id })?.blocks?.length ?? 0) === 1, 'the delete in the engine');
    const afterRange = doc.model({ id: sessionFor('doc').id });
    const left = afterRange?.blocks?.map((b) => b.runs?.map((r) => r.text).join('') ?? '') ?? [];
    check('word: Delete over that selection takes the selected words and the paragraph mark between them',
      left.length === 1 && left[0] === 'He',
      `paragraphs now ${JSON.stringify(left)}`);

    // Paste over a selection replaces it — the clipboard's words where the
    // selected ones were, nothing of the old ones left.
    clipboard.writeText('Pasted');
    await word.webContents.executeJavaScript(`(() => {
      const page = document.querySelector('.wd-page');
      const block = page.querySelector('[data-block="0"]');
      const range = document.createRange();
      range.selectNodeContents(block);
      const sel = getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      page.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
      return 'selected';
    })()`);
    await settle(() => doc.model({ id: sessionFor('doc').id })?.selection?.collapsed === false, 'the range to paste over in the engine');
    word.webContents.paste();
    // The menu's paste reaches a page only while it truly holds the desktop's
    // focus — and with the focus emulation above it does not reach it at
    // all — so when it has not landed, the page is handed the event a paste
    // raises (beforeinput, insertFromPaste, the words on its dataTransfer).
    // What this proves is the page's handling of a paste over a selection;
    // the operating system's clipboard is not part of it.
    if (!(await settle(() => blockText(0) === 'Pasted', 'the paste in the engine', 3000))) {
      console.log('       (the menu paste did not reach the page; the event a paste raises is handed over instead)');
      await word.webContents.executeJavaScript(`(() => {
        const dt = new DataTransfer();
        dt.setData('text/plain', 'Pasted');
        const target = document.activeElement?.closest?.('.wd-page') || document.querySelector('.wd-page');
        target.dispatchEvent(new InputEvent('beforeinput', { inputType: 'insertFromPaste', dataTransfer: dt, bubbles: true, cancelable: true }));
        return 'pasted';
      })()`);
      await settle(() => blockText(0) === 'Pasted', 'the paste in the engine');
    }
    const pasted = doc.model({ id: sessionFor('doc').id })?.blocks?.map((b) => b.runs?.map((r) => r.text).join('') ?? '') ?? [];
    check('word: a paste over a selection replaces the selected words', pasted.length === 1 && pasted[0] === 'Pasted', `paragraphs now ${JSON.stringify(pasted)}`);
  } catch (err) {
    check('word: the checks ran', false, err.message);
  }

  /* ── Worksheets ─────────────────────────────────────────────────────── */

  const sheets = await openWindow('sheets');
  try {
    await sheets.webContents.executeJavaScript(`document.querySelector('.sh')?.focus(), 'ok'`);
    await wait(200);

    // Typing a printable character starts an edit; the rest goes into the cell.
    for (const ch of '42') await typeChar(sheets.webContents, ch);
    await wait(200);
    await press(sheets.webContents, 'Return', { char: true });
    await wait(450);

    const session = sessionFor('sheet');
    const cell = (ref) => doc.model({ id: session.id })?.cells?.find((c) => c.ref === ref);
    await until(() => cell('A1')?.text === '42', 'the typed value to reach the cell');
    const a1 = cell('A1');
    // A number right-aligns; text left-aligns. They display identically, and
    // only one of them can be multiplied — which is what the next check finds.
    check('sheets: typing into a cell reaches the engine', a1?.text === '42' && a1?.align === 'right', `A1 is ${JSON.stringify(a1?.text ?? null)}, aligned ${a1?.align} (so it is ${a1?.align === 'right' ? 'a number' : 'text'})`);

    // A formula has to survive the same path, and then compute.
    for (const ch of '=A1*2') await typeChar(sheets.webContents, ch);
    await wait(200);
    await press(sheets.webContents, 'Return', { char: true });
    await wait(600);

    try {
      await until(() => cell('A2')?.text === '84', 'the formula to compute');
    } catch {
      // Fall through to the check, which reports what it actually holds.
    }
    const a2 = cell('A2');
    check('sheets: a formula computes', a2?.text === '84', `A2 is ${JSON.stringify(a2?.text ?? null)}${a2?.isFormula ? ' (stored as a formula)' : ' (not a formula)'}`);
  } catch (err) {
    check('sheets: the checks ran', false, err.message);
  }

  for (const win of opened) win.destroy();

  const failed = results.filter((r) => !r.ok);
  console.log(`\n${results.length - failed.length}/${results.length} editing checks passed`);
  return failed.length === 0;
}
