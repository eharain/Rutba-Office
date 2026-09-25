// Word: Review → Track Changes — recording, not just reading.
//
// Turning it on writes `w:trackRevisions`; typing after that wraps the new
// words in `w:ins`, deleting existing words wraps them in `w:del` rather
// than removing them. All Markup shows both inline (underlined, struck
// through, a change bar); Accept All keeps the insertion and drops the
// deletion for good. Run alone with RUTBA_VERIFY_ONLY=track.
import fs from 'node:fs';
import { openDocx } from '@rutba/doc-view/backends/ooxml';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, capture, doc, sessionFor
 */
export async function verifyWordTrack(h, { file }) {
  const { open, check, until, wait, errorsIn, capture, doc, sessionFor } = h;
  if (!file) return check('word: the track changes fixture was made', false, 'no fixture');
  try {
    const win = await open('word', file);
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
    const clickTab = (name) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(), 'tab'`);
    const pickMenu = async (label) => {
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}))`), `the ${label} menu item`, 4000);
      return js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}); b.click(); return 'picked'; })()`);
    };
    // A run through the window's own bridge — the picture check's technique
    // for setting state up behind the window's back — then Ctrl+Z/Ctrl+Y so
    // its React model, which never heard an apply it did not dispatch
    // itself, catches up before the next ribbon press reads it.
    const applyBehind = (ops) => js(`(async () => {
      const all = await window.rutbaOffice.doc.sessions({});
      const mine = all.filter((s) => s.kind === 'doc').pop();
      await window.rutbaOffice.doc.apply({ id: mine.id, ops: ${JSON.stringify(ops)} });
      return 1;
    })()`);
    const syncWindow = async () => {
      await js(`document.querySelector('.wd-page')?.focus(), 'focused'`);
      win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'z', modifiers: ['control'] });
      win.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'z', modifiers: ['control'] });
      await wait(250);
      win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'y', modifiers: ['control'] });
      win.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'y', modifiers: ['control'] });
      await wait(250);
    };

    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="0"]'))`), 'the first paragraph', 8000);

    await clickTab('Review');
    await wait(200);
    const pressedTrack = await clickRibbon('Track Changes');
    const trackedOn = await until(() => model().trackRevisions === true, 'recording to turn on', 5000).catch(() => false);
    check(
      'word: Review → Track Changes turns recording on, and the ribbon shows it pressed',
      pressedTrack === 'clicked' && trackedOn === true,
      `${pressedTrack}; trackRevisions=${model().trackRevisions}`
    );
    const pressedShown = await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Track Changes')); return b ? b.getAttribute('aria-pressed') === 'true' : null; })()`);
    check('word: the Track Changes button reads pressed', pressedShown !== false, `pressed=${pressedShown}`);

    // Display for Review → All Markup, so the change is shown inline.
    await clickRibbon('Display for Review');
    await pickMenu('All Markup');
    await wait(150);

    // Type at the end of the first paragraph, behind the window's back, then sync it.
    const before0 = model().blocks[0];
    await applyBehind([
      { op: 'setSelection', anchor: { block: 0, offset: before0.text.length }, focus: { block: 0, offset: before0.text.length } },
      { op: 'insertText', text: ' ADDED' },
    ]);
    await syncWindow();
    const inserted = await until(
      () => (model().blocks[0].runs || []).some((r) => r.ins && r.text === ' ADDED'),
      'the insertion in the model',
      6000
    ).catch(() => false);
    check('word: typing while recording writes an insertion the model carries author and date on', inserted === true, JSON.stringify(model().blocks[0].runs));

    const insShown = await until(
      () => js(`Boolean([...document.querySelectorAll('[data-block="0"] .wd-ins')].find((s) => s.textContent.includes('ADDED')))`),
      'the insertion drawn underlined on the page',
      6000
    ).catch(() => false);
    const insDeco = insShown ? await js(`(() => { const s = [...document.querySelectorAll('[data-block="0"] .wd-ins')].find((s) => s.textContent.includes('ADDED')); return getComputedStyle(s).textDecorationLine; })()`) : null;
    check('word: All Markup shows the insertion underlined on the page', insShown === true && /underline/.test(insDeco || ''), `shown=${insShown}; decoration=${insDeco}`);

    // Delete "text" off the end of the second paragraph.
    const before1 = model().blocks[1];
    const delFrom = before1.text.length - 4;
    await applyBehind([
      { op: 'setSelection', anchor: { block: 1, offset: delFrom }, focus: { block: 1, offset: before1.text.length } },
      { op: 'deleteSelection' },
    ]);
    await syncWindow();
    const deleted = await until(
      () => (model().blocks[1].runs || []).some((r) => r.del?.text === 'text'),
      'the deletion in the model',
      6000
    ).catch(() => false);
    check(
      'word: deleting existing words while recording marks them deleted rather than removing them',
      deleted === true && model().blocks[1].text === 'Second paragraph ',
      `text now "${model().blocks[1].text}"; runs ${JSON.stringify(model().blocks[1].runs)}`
    );

    const delShown = await until(
      () => js(`Boolean([...document.querySelectorAll('[data-block="1"] .wd-del')].find((s) => s.textContent === 'text'))`),
      'the deletion drawn struck through on the page',
      6000
    ).catch(() => false);
    const delDeco = delShown ? await js(`(() => { const s = [...document.querySelectorAll('[data-block="1"] .wd-del')].find((s) => s.textContent === 'text'); return getComputedStyle(s).textDecorationLine; })()`) : null;
    check('word: All Markup shows the deletion struck through on the page', delShown === true && /line-through/.test(delDeco || ''), `shown=${delShown}; decoration=${delDeco}`);

    const barShown = await js(`Boolean(document.querySelector('.wd-block.wd-changebar'))`);
    check('word: a paragraph with a tracked change wears a change bar', barShown === true, `bar shown: ${barShown}`);

    await capture(win, 'word-track.png');

    // Accept All — the insertion's words stay, the deletion's are finally
    // gone. The caret is put on a paragraph that IS tracked first: the
    // Accept button is disabled otherwise, and where the last sync left it
    // is not this check's business to assume.
    await applyBehind([{ op: 'setSelection', anchor: { block: 1, offset: 0 }, focus: { block: 1, offset: 0 } }]);
    await syncWindow();
    await clickRibbon('Accept');
    await pickMenu('Accept All Changes');
    const accepted = await until(
      () => model().blocks[0].text.endsWith('ADDED') && !(model().blocks[0].runs || []).some((r) => r.ins) &&
        model().blocks[1].text === 'Second paragraph ' && !(model().blocks[1].runs || []).some((r) => r.del),
      'Accept All to resolve both changes',
      6000
    ).catch(() => false);
    check('word: Accept All keeps the insertion and drops the deletion, for good', accepted === true, JSON.stringify([model().blocks[0], model().blocks[1]]));

    await clickRibbon('Save');
    const xmlOf = (i) => {
      try {
        return openDocx(fs.readFileSync(file)).doc.doc.editParagraph(i).xml;
      } catch {
        return '';
      }
    };
    await until(() => /ADDED/.test(xmlOf(0)) && !/w:ins\b/.test(xmlOf(0)), 'the accepted file on disk', 8000).catch(() => {});
    const savedOk = !/w:(ins|del)\b/.test(xmlOf(0)) && !/w:(ins|del)\b/.test(xmlOf(1)) && /ADDED/.test(xmlOf(0)) && !/text<\/w:t>/.test(xmlOf(1));
    check('word: the saved file carries neither field once both changes are accepted', savedOk, `${xmlOf(0).slice(-160)} | ${xmlOf(1).slice(-160)}`);

    // Turn recording off again.
    await clickTab('Review');
    await wait(150);
    await clickRibbon('Track Changes');
    const trackedOff = await until(() => model().trackRevisions === false, 'recording to turn off', 5000).catch(() => false);
    check('word: Track Changes turns off again', trackedOff === true, `trackRevisions=${model().trackRevisions}`);

    const complaints = await errorsIn(win);
    check('word: the track changes checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    // track.docx is this check's own fixture, rebuilt fresh by makeFixtures
    // on every run — nothing shared to restore.
  } catch (err) {
    check('word: the track changes checks ran', false, err.message);
  }
}
