// Word: typing in the Letter template's address line (2.18.0, A3.0). In
// 2.17.1 every keystroke in a line holding a tab added another tab, saved
// with the file; a triple-click's selection ran on to the next line, so the
// name typed over it joined the address line; and a space typed there then
// froze the window. Now a triple-click selects the line, typing keeps its
// paragraph mark and its one tab, the window keeps answering, and a letter
// already carrying 72 tabs takes a space without a stall. Run alone with
// RUTBA_VERIFY_ONLY=wordtyping.
import fs from 'node:fs';
import path from 'node:path';
import { OoxmlPackage } from '@rutba/ooxml';
import { TEMPLATES, templatePath } from '@rutba/templates';

const tabsIn = (xml) => (xml.match(/<w:tab\/>/g) || []).length;

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where the fixtures are written
 */
export async function verifyWordTypingTabs(h, { dir }) {
  const { open, check, until, wait, press, errorsIn, doc, sessionFor } = h;
  const source = fs.readFileSync(templatePath(TEMPLATES.find((t) => t.file === 'word-letter.docx')));
  const tabsBefore = tabsIn(OoxmlPackage.read(source).text('word/document.xml'));
  const file = path.join(dir, 'letter-typing.docx');
  fs.writeFileSync(file, source);
  // The letter the Templates episode saved: 72 tabs after its address.
  const heavy = path.join(dir, 'letter-72-tabs.docx');
  {
    const pkg = OoxmlPackage.read(source);
    const xml = pkg.text('word/document.xml');
    const at = xml.indexOf('you@example.com</w:t></w:r>') + 'you@example.com</w:t></w:r>'.length;
    pkg.write_('word/document.xml', xml.slice(0, at) + '<w:r><w:tab/></w:r>'.repeat(72) + xml.slice(at));
    fs.writeFileSync(heavy, pkg.write());
  }

  const typeChars = async (wc, text) => {
    for (const ch of text) {
      wc.sendInputEvent({ type: 'keyDown', keyCode: ch });
      wc.sendInputEvent({ type: 'char', keyCode: ch });
      wc.sendInputEvent({ type: 'keyUp', keyCode: ch });
      await wait(60);
    }
  };
  // The window answers a small question in good time, or it is frozen.
  const answers = (wc, ms) => Promise.race([
    wc.executeJavaScript('1').then(() => true),
    new Promise((resolve) => setTimeout(() => resolve(false), ms)),
  ]);

  let win = null;
  try {
    win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const model = () => doc.model({ id: sessionFor('doc').id });
    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="1"]'))`), 'the letter', 8000);

    // Triple-click the first "Your Name".
    const at = await js(`(() => { const r = document.querySelector('.wd-page [data-block="0"]').getBoundingClientRect(); return { x: Math.round(r.left + 30), y: Math.round(r.top + r.height / 2) }; })()`);
    for (const clickCount of [1, 2, 3]) {
      wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount });
      wc.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount });
      await wait(90);
    }
    const selected = await until(() => {
      const s = model().selection;
      return s && s.from.block === 0 && s.from.offset === 0 && s.to.block === 0 && s.to.offset === 'Your Name'.length;
    }, 'the line selected', 4000).then(() => true).catch(() => false);
    const domSelected = await js(`getSelection().toString()`);
    check('word: a triple-click selects the line it is on, and not the start of the next one',
      selected && !/Street/.test(domSelected), JSON.stringify({ engine: model().selection, page: domSelected }));

    await typeChars(wc, 'L');
    await typeChars(wc, 'arkspur');
    await typeChars(wc, ' ');
    const alive = await answers(wc, 5000);
    const typed = await until(() => model().blocks[0].text === 'Larkspur ', 'the name typed', 5000).then(() => true).catch(() => false);
    const blocks = model().blocks;
    check('word: the name typed over the selected line stays on its own line, the address below it with its one tab, and the window answers',
      alive && typed && blocks[1].text === 'Street address, Town, Postcode\tphone  |  you@example.com',
      JSON.stringify({ alive, first: blocks[0].text, second: blocks[1].text }));

    // Typed in the address line itself: each keystroke leaves the one tab.
    const line = await js(`(() => { const r = document.querySelector('.wd-page [data-block="1"]').getBoundingClientRect(); return { x: Math.round(r.left + 4), y: Math.round(r.top + r.height / 2) }; })()`);
    wc.sendInputEvent({ type: 'mouseDown', x: line.x, y: line.y, button: 'left', clickCount: 1 });
    wc.sendInputEvent({ type: 'mouseUp', x: line.x, y: line.y, button: 'left', clickCount: 1 });
    await wait(200);
    await press(wc, 'Home');
    await typeChars(wc, '12 Mill Lane, ');
    await until(() => model().blocks[1].text.startsWith('12 Mill Lane, '), 'the address typed', 5000).catch(() => false);
    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    await press(wc, 's', { modifiers: ['control'] });
    let tabsAfter = -1;
    const saved = await until(() => {
      try {
        const xml = OoxmlPackage.read(fs.readFileSync(file)).text('word/document.xml');
        tabsAfter = tabsIn(xml);
        return xml.includes('12 Mill Lane, ');
      } catch { return false; }
    }, 'the save', 6000).then(() => true).catch(() => false);
    check('word: typing in a line that holds a tab adds no tab, and the saved letter has as many tabs as the template',
      saved && tabsAfter === tabsBefore && model().blocks[1].text.split('\t').length === 2, `${tabsAfter} tabs saved, ${tabsBefore} in the template; ${JSON.stringify(model().blocks[1].text)}`);
    const complaints = await errorsIn(win);
    check('word: typing in the letter reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the Letter typing checks ran', false, err?.message || String(err));
  }

  // The line the freeze was found on, in a letter that already carries the tabs.
  try {
    win = await open('word', heavy);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const model = () => doc.model({ id: sessionFor('doc').id });
    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="1"]'))`), 'the letter', 8000);
    const end = await js(`(() => { const r = document.querySelector('.wd-page [data-block="0"]').getBoundingClientRect(); return { x: Math.round(r.right - 2), y: Math.round(r.top + r.height / 2) }; })()`);
    wc.sendInputEvent({ type: 'mouseDown', x: end.x, y: end.y, button: 'left', clickCount: 1 });
    wc.sendInputEvent({ type: 'mouseUp', x: end.x, y: end.y, button: 'left', clickCount: 1 });
    await wait(200);
    await press(wc, 'End');
    await press(wc, 'Delete');
    await until(() => /^Your NameStreet/.test(model().blocks[0].text), 'the lines joined', 4000).catch(() => false);
    await typeChars(wc, ' ');
    const started = Date.now();
    const alive = await answers(wc, 5000);
    const spaced = await until(() => /^Your Name Street/.test(model().blocks[0].text), 'the space', 5000).then(() => true).catch(() => false);
    check('word: a space typed into the joined name and address line, 72 tabs after it, is laid out and the window answers',
      alive && spaced, JSON.stringify({ alive, ms: Date.now() - started, text: model().blocks[0].text.slice(0, 60) }));
  } catch (err) {
    check('word: the joined-line check ran', false, err?.message || String(err));
  }
}
