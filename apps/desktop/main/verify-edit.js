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
 * Press a key the way a keyboard does.
 *
 * Keys that produce a character need a "char" event between the down and the
 * up, or Chromium's editing pipeline never sees them — Return is one of those
 * and Backspace is not, which is a distinction worth writing down because it
 * costs an hour to rediscover.
 */
async function press(wc, keyCode, { modifiers = [], char = false } = {}) {
  wc.sendInputEvent({ type: 'keyDown', keyCode, modifiers });
  if (char) wc.sendInputEvent({ type: 'char', keyCode, modifiers });
  wc.sendInputEvent({ type: 'keyUp', keyCode, modifiers });
  await wait(120);
}

/** Type a printable character, which is a keyDown, a char, and a keyUp. */
async function typeChar(wc, ch) {
  wc.sendInputEvent({ type: 'keyDown', keyCode: ch });
  wc.sendInputEvent({ type: 'char', keyCode: ch });
  wc.sendInputEvent({ type: 'keyUp', keyCode: ch });
  await wait(70);
}

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
    opened.push(win);
    return win;
  };

  /** The session the window opened, read straight from the document service. */
  const sessionFor = (kind) => doc.sessions().filter((s) => s.kind === kind).pop();

  /* ── Word ───────────────────────────────────────────────────────────── */

  const word = await openWindow('word');
  try {
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
    await wait(450);

    let session = sessionFor('doc');
    let model = session && doc.model({ id: session.id });
    const afterType = model?.blocks?.[0]?.runs?.map((r) => r.text).join('') ?? '';
    check('word: typing reaches the engine', afterType.startsWith('Hello'), `first block is ${JSON.stringify(afterType)}`);

    // Backspace never arrives through React's polyfill: keypress does not fire
    // for it. This is the half of the defect a text-only check would miss.
    await press(word.webContents, 'Backspace');
    await wait(300);
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

    await press(word.webContents, 'Return', { char: true });
    await wait(320);
    const seen = await word.webContents.executeJavaScript('window.__seen || []');
    if (!seen.length) console.log('       (no beforeinput arrived for Return)');
    else console.log(`       (beforeinput saw: ${seen.join(', ')})`);
    session = sessionFor('doc');
    model = session && doc.model({ id: session.id });
    const blocksAfter = model?.blocks?.length ?? 0;
    check('word: Enter splits the paragraph', blocksAfter === blocksBefore + 1, `${blocksBefore} → ${blocksAfter} blocks`);

    // And the window shows it, which is the other half of the loop.
    const shown = await word.webContents.executeJavaScript(
      `document.querySelector('.wd-page [data-block="0"]')?.textContent ?? ''`
    );
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
    await wait(280);

    // The engine must have the *range*, not a caret; a collapsed selection here
    // means the sync lost it, and bold would silently apply to nothing.
    const before = doc.model({ id: sessionFor('doc').id });
    check(
      'word: selecting text reaches the engine as a range',
      before.selection && !before.selection.collapsed,
      `engine has ${JSON.stringify(before.selection)}`
    );

    await press(word.webContents, 'b', { modifiers: ['control'] });
    await wait(380);

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
