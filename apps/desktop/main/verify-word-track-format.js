// Documents: Track Changes records formatting — Home → Bold pressed on a
// word with recording on writes a change of formatting from how the word
// looked; All Markup tells who formatted it, Original draws it as it was,
// the Reviewing Pane lists it, and Reject All puts it back. A paragraph
// Word added while tracking — its mark recorded — is accepted and saved
// whole, where it used to be written back broken.

import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml';

/**
 * @param {object} h the harness: open, check, until, wait, doc
 */
export async function verifyWordTrackFormat(h, { dir }) {
  const { open, check, until, wait, doc } = h;
  const file = path.join(dir, 'track-format.docx');
  try {
    // A paragraph Word added while tracking, then one to format.
    const pkg = OoxmlPackage.read(buildDocx({ styles: true, paragraphs: [{ text: 'PLACEHOLDER' }, { text: 'The figures were checked twice.' }] }));
    const added = '<w:p><w:pPr><w:rPr><w:ins w:id="1" w:author="Ann" w:date="2026-01-01T00:00:00Z"/></w:rPr></w:pPr>'
      + '<w:r><w:t xml:space="preserve">Ann wrote this </w:t></w:r><w:ins w:id="2" w:author="Ann" w:date="2026-01-01T00:00:00Z"><w:r><w:t>paragraph</w:t></w:r></w:ins></w:p>';
    pkg.write_('word/document.xml', pkg.text('word/document.xml').replace(/<w:p\b(?:(?!<\/w:p>)[\s\S])*?PLACEHOLDER[\s\S]*?<\/w:p>/, () => added));
    fs.writeFileSync(file, pkg.write());

    const win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const session = doc.sessions().filter((s) => s.kind === 'doc' && s.path === file).pop();
    const model = () => doc.model({ id: session.id });
    const tab = (name) => js(`(() => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(); return 1; })()`);
    const ribbon = (title) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => ((n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}) || n.textContent.trim() === ${JSON.stringify(title)}) && !n.disabled);
      if (!b) return 'no button';
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      b.click();
      return 'clicked';
    })()`);
    const menu = async (label) => {
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}))`), `the ${label} item`, 4000);
      return js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}).click(); return 'picked'; })()`);
    };
    const span = (word) => js(`(() => { const s = [...document.querySelectorAll('.wd-page [data-block="1"] span')].find((n) => n.children.length === 0 && n.textContent.trim() === ${JSON.stringify(word)}); return s ? { weight: getComputedStyle(s).fontWeight, title: s.title || '' } : null; })()`);

    await until(() => js(`document.querySelectorAll('.wd-page [data-block]').length >= 2`), 'the page', 8000);
    await tab('Review');
    await wait(200);
    await ribbon('Track Changes');
    await until(() => model().trackRevisions === true, 'recording on', 5000).catch(() => {});
    await ribbon('Display for Review');
    await menu('All Markup');

    // Double-click "figures" — the page selects the word — then Home → Bold.
    const at = await js(`(() => {
      const b = document.querySelector('.wd-page [data-block="1"]');
      const range = document.createRange();
      const walker = document.createTreeWalker(b, NodeFilter.SHOW_TEXT);
      for (let n = walker.nextNode(); n; n = walker.nextNode()) {
        const i = n.textContent.indexOf('figures');
        if (i >= 0) { range.setStart(n, i); range.setEnd(n, i + 7); const r = range.getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) }; }
      }
      return null;
    })()`);
    if (at) {
      wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      wc.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 2 });
      wc.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 2 });
    }
    await until(() => { const s = model().selection; return s && s.anchor.block === 1 && Math.abs(s.focus.offset - s.anchor.offset) >= 7; }, 'the word selected', 4000).catch(() => {});
    await tab('Home');
    await wait(150);
    const bolded = await ribbon('Bold');
    const recorded = await until(() => (model().blocks[1].runs || []).some((r) => r.text.trim() === 'figures' && r.bold && r.formatChange?.author && r.formatChange.was.bold === false), 'the change of formatting', 6000).then(() => true).catch(() => false);
    const shown = await until(async () => /^Formatted by /.test((await span('figures'))?.title || ''), 'the tip on the page', 4000).then(() => span('figures')).catch(() => span('figures'));
    check('documents: Bold pressed with Track Changes on records a change of formatting, the word bold and its tip saying who formatted it',
      Boolean(at) && bolded === 'clicked' && recorded && Number(shown?.weight) >= 600, JSON.stringify({ at, bolded, recorded, shown, selection: model().selection }));

    // Original draws it as it was; the Reviewing Pane lists it.
    await tab('Review');
    await wait(150);
    await ribbon('Display for Review');
    await menu('Original');
    const original = await until(async () => Number((await span('figures'))?.weight) < 600, 'the word as it was', 4000).then(() => span('figures')).catch(() => span('figures'));
    await ribbon('Display for Review');
    await menu('All Markup');
    await ribbon('Reviewing Pane');
    const kinds = () => js(`[...document.querySelectorAll('.wd-tracked-kinds')].map((n) => n.textContent)`);
    await until(async () => (await kinds()).length > 0, 'the Reviewing Pane', 4000).catch(() => {});
    const pane = await kinds();
    await js(`(() => { [...document.querySelectorAll('.rw-dialog button, .rw-btn')].find((b) => b.textContent.trim() === 'Close')?.click(); return 1; })()`);
    check('documents: Original draws a formatted word as it was, and the Reviewing Pane lists the formatting and the new paragraph',
      Number(original?.weight) < 600 && Array.isArray(pane) && pane.some((t) => /formatting/.test(t)) && pane.some((t) => /a new paragraph/.test(t)),
      JSON.stringify({ original, pane }));

    // Reject on the formatted paragraph: the word plain again, the record gone.
    await js(`document.querySelector('.wd-page [data-block="1"]')?.focus(), 1`);
    const rejected = await (async () => {
      const r = await ribbon('Reject');
      if (r !== 'clicked') return r;
      return menu('Reject This Change');
    })();
    const plain = await until(() => (model().blocks[1].runs || []).every((r) => !r.bold && !r.formatChange), 'the word plain again', 6000).then(() => true).catch(() => false);
    check('documents: Reject puts a recorded change of formatting back', rejected === 'picked' && plain, JSON.stringify({ rejected, runs: model().blocks[1].runs }));

    // Accept All: Ann's paragraph kept, its records gone, the file whole —
    // pressed with the caret in hers, where Accept is live.
    const inAnn = await js(`(() => { const b = document.querySelector('.wd-page [data-block="0"]'); const r = b.getBoundingClientRect(); return { x: Math.round(r.left + 4), y: Math.round(r.top + r.height / 2) }; })()`);
    wc.sendInputEvent({ type: 'mouseDown', x: inAnn.x, y: inAnn.y, button: 'left', clickCount: 1 });
    wc.sendInputEvent({ type: 'mouseUp', x: inAnn.x, y: inAnn.y, button: 'left', clickCount: 1 });
    await until(() => model().selection?.focus?.block === 0, 'the caret in her paragraph', 4000).catch(() => {});
    await tab('Review');
    await wait(150);
    await ribbon('Accept');
    await menu('Accept All Changes');
    await until(() => !model().blocks.some((b) => b.tracked), 'everything accepted', 6000).catch(() => {});
    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => {
      try {
        const xml = OoxmlPackage.read(fs.readFileSync(file)).text('word/document.xml');
        const body = /<w:body>([\s\S]*)<\/w:body>/.exec(xml)[1];
        const stack = [];
        for (const m of body.matchAll(/<(\/?)([\w:]+)[^>]*?(\/?)>/g)) {
          if (m[3] === '/') continue;
          if (m[1] === '/') { if (stack.pop() !== m[2]) return false; } else stack.push(m[2]);
        }
        return stack.length === 0 && /Ann wrote this <\/w:t><\/w:r><w:r><w:t>paragraph/.test(body) && !/<w:(?:ins|del|rPrChange)\b/.test(body);
      } catch { return false; }
    }, 'the accepted file', 8000).then(() => true).catch(() => false);
    check('documents: Accept All on a paragraph Word added while tracking keeps it and saves the file whole', saved, saved ? 'accepted and well formed' : 'not in the file, or broken');
    await wait(50);
  } catch (err) {
    check('documents: the tracked formatting checks ran', false, err.message);
  }
}
