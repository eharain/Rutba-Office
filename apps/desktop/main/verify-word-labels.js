// Word: Mailings → Labels and Envelopes, driven through their dialogs.
//
// Labels: an address typed into the Labels dialog, Avery L7160 chosen in
// Label Options, New Document — a Labels window of its own with twenty-one
// labels on an A4 sheet at the sheet's margins, and a PDF of it whose every
// address prints inside its own label. Envelopes: Add to Document puts a DL
// envelope in front of a letter as a section of its own; the page draws its
// sheet at the envelope's size with the delivery address in its frame, and
// the PDF's first page is DL. Run alone with RUTBA_VERIFY_ONLY=labels.
import fs from 'node:fs';
import path from 'node:path';
import { BrowserWindow } from 'electron';
import { buildDocx } from '@rutba/ooxml/build';
import { labelProduct } from '@rutba/ooxml/labels';

const ADDRESS = ['Ms. Cynthia Gartner', 'Contoso Ltd', '1 Main Road', 'Leeds LS1 1AA'];
const PT = 72 / 25.4;

/** Every text a PDF draws, per page, with where it starts. */
function pdfTexts(buffer) {
  const s = buffer.toString('latin1');
  const pages = [];
  for (const stream of s.matchAll(/stream\r?\n([\s\S]*?)endstream/g)) {
    const out = [];
    for (const m of stream[1].matchAll(/1 0 0 1 ([\d.-]+) ([\d.-]+) Tm\s*\(((?:\\.|[^\\)])*)\) Tj/g)) out.push({ text: m[3], x: Number(m[1]), y: Number(m[2]) });
    if (out.length) pages.push(out);
  }
  return pages;
}
const mediaBoxes = (buffer) => [...buffer.toString('latin1').matchAll(/\/MediaBox \[([^\]]*)\]/g)].map((m) => m[1].split(' ').map(Number));

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn, capture, doc, sessionFor
 */
export async function verifyWordLabels(h, { dir }) {
  const { open, check, until, wait, press, errorsIn, capture, doc, sessionFor } = h;
  let win = null;
  let saved = null;
  try {
    const letter = path.join(dir, 'labels-letter.docx');
    fs.writeFileSync(letter, buildDocx({ styles: true, paragraphs: [{ text: 'Dear Cynthia,' }, { text: 'Your parcel is on its way.' }, { text: 'Yours sincerely,' }] }));
    win = await open('word', letter);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = sessionFor('doc');
    const model = () => doc.model({ id: session.id });
    // The return address and label Word remembers are the person's settings:
    // what this check changes, it puts back.
    saved = await js(`(async () => ({ ret: await window.rutbaOffice.store.get({ key: 'word.mailings.returnAddress', fallback: null }), label: await window.rutbaOffice.store.get({ key: 'word.mailings.label', fallback: null }), env: await window.rutbaOffice.store.get({ key: 'word.mailings.envelope', fallback: null }) }))()`);

    const clickRibbon = (title) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}) && !n.disabled);
      if (!b) return 'no button ' + ${JSON.stringify(title)};
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      b.click();
      return 'clicked';
    })()`);
    const clickDialog = (dialog, label) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-dialog[aria-label=${JSON.stringify(dialog)}] button')].find((n) => n.textContent.trim() === ${JSON.stringify(label)} && !n.disabled);
      if (!b) return 'no ' + ${JSON.stringify(label)} + ' in ' + ${JSON.stringify(dialog)};
      b.click();
      return 'clicked';
    })()`);
    const typeInto = (selector, value) => js(`(() => {
      const el = document.querySelector(${JSON.stringify(selector)});
      if (!el) return false;
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(el, ${JSON.stringify(value)});
      el.dispatchEvent(new Event('input', { bubbles: true }));
      return true;
    })()`);
    const snap = async (w, name) => {
      w.webContents.invalidate();
      await wait(700);
      await w.webContents.capturePage().catch(() => null);
      w.webContents.invalidate();
      await wait(600);
      await capture(w, name);
    };

    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="0"]'))`), 'the letter', 8000);
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Mailings')?.click(), 'tab'`);
    await wait(250);

    /* ── Labels → New Document ─────────────────────────────────────────── */
    const labelsPressed = await clickRibbon('Labels');
    await until(() => js(`Boolean(document.querySelector('.wd-el-labeltext'))`), 'the Labels dialog', 5000).catch(() => {});
    const typed = await typeInto('.wd-el-labeltext', ADDRESS.join('\n'));
    await clickDialog('Labels', 'Options…');
    await until(() => js(`Boolean(document.querySelector('.wd-el-products [data-product]'))`), 'Label Options', 4000).catch(() => {});
    await js(`(() => { const s = document.querySelector('.wd-el-vendor'); Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(s, 'Avery A4/A5'); s.dispatchEvent(new Event('change', { bubbles: true })); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.wd-el-products [data-product="L7160"]'))`), 'the A4 products', 3000).catch(() => {});
    await js(`document.querySelector('.wd-el-products [data-product="L7160"]')?.click(), 1`);
    await clickDialog('Label Options', 'OK');
    await until(() => js(`/L7160/.test(document.querySelector('.wd-el-product-name')?.textContent || '')`), 'L7160 chosen', 3000).catch(() => {});
    await snap(win, 'word-labels-dialog.png');
    const before = new Set(BrowserWindow.getAllWindows().map((w) => w.id));
    const made = await clickDialog('Labels', 'New Document');
    let labelsWin = null;
    await until(() => {
      labelsWin = BrowserWindow.getAllWindows().find((w) => !before.has(w.id) && !w.isDestroyed() && w.webContents.getURL().includes('session='));
      return Boolean(labelsWin);
    }, 'the Labels window', 8000).catch(() => {});
    let sheet = null;
    let geometry = null;
    let labelsSession = null;
    if (labelsWin) {
      const ljs = (code) => labelsWin.webContents.executeJavaScript(code);
      await until(() => ljs(`document.querySelectorAll('.wd-page .wd-table td .wd-block').length >= 35`), 'the sheet to draw', 12000).catch(() => {});
      await wait(500);
      sheet = await ljs(`(() => {
        const page = document.querySelector('.wd-page');
        const pageRect = page.getBoundingClientRect();
        const z = pageRect.width / page.offsetWidth;
        const table = page.querySelector('.wd-table');
        const cells = [...table.querySelectorAll('td')];
        const filled = cells.filter((td) => td.innerText.trim());
        const first = filled[0].getBoundingClientRect();
        const second = filled[1].getBoundingClientRect();
        const row2 = filled[3].getBoundingClientRect();
        return {
          sheets: page.querySelectorAll('.wd-sheet').length,
          cells: cells.length,
          labels: filled.length,
          same: filled.every((td) => td.innerText.trim() === filled[0].innerText.trim()),
          top: Math.round((table.getBoundingClientRect().top - pageRect.top) / z),
          left: Math.round((first.left - pageRect.left) / z),
          width: Math.round(first.width / z),
          height: Math.round(first.height / z),
          pitchX: Math.round((second.left - first.left) / z),
          pitchY: Math.round((row2.top - first.top) / z),
          bordered: getComputedStyle(cells[0]).borderTopStyle,
        };
      })()`).catch((e) => ({ error: e.message }));
      labelsSession = doc.sessions().find((s) => /^Labels\d+$/.test(s.name));
      geometry = labelsSession ? doc.model({ id: labelsSession.id }).section : null;
    }
    const p = labelProduct('L7160');
    const mm = (px) => (px * 25.4) / 96;
    check(
      'word: Labels → New Document makes a sheet of twenty-one L7160 labels on A4, at the sheet\'s margins and pitch, with no borders',
      labelsPressed === 'clicked' && typed === true && made === 'clicked' && Boolean(sheet) && sheet.labels === 21 && sheet.same && sheet.sheets === 1
        && Math.abs(mm(sheet.top) - p.top) < 0.6 && Math.abs(mm(sheet.left) - p.side) < 0.6 && Math.abs(mm(sheet.width) - p.w) < 0.6 && Math.abs(mm(sheet.height) - p.h) < 0.6
        && Math.abs(mm(sheet.pitchX) - p.pitchX) < 0.6 && Math.abs(mm(sheet.pitchY) - p.pitchY) < 0.6 && sheet.bordered === 'none'
        && Math.round(geometry?.widthPx) === 794 && Math.round(geometry?.heightPx) === 1123,
      JSON.stringify({ sheet, page: geometry ? { w: Math.round(geometry.widthPx), h: Math.round(geometry.heightPx), top: Math.round(geometry.margins.top), left: Math.round(geometry.margins.left) } : null })
    );
    if (labelsWin) {
      await snap(labelsWin, 'word-labels.png');
      const complaints = await errorsIn(labelsWin).catch(() => []);
      check('word: the labels window reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    }

    // The same sheet as a PDF: A4, and every address inside its own label.
    if (labelsSession) {
      const pdfPath = path.join(dir, 'labels.pdf');
      doc.export({ id: labelsSession.id, format: 'pdf', path: pdfPath });
      const buffer = fs.readFileSync(pdfPath);
      const [box] = mediaBoxes(buffer);
      const firsts = (pdfTexts(buffer)[0] || []).filter((t) => t.text === ADDRESS[0]);
      const inside = firsts.filter((t) => {
        const col = Math.floor((t.x / PT - p.side) / p.pitchX);
        const row = Math.floor(((box[3] - t.y) / PT - p.top) / p.pitchY);
        const left = (p.side + col * p.pitchX) * PT;
        const top = (p.top + row * p.pitchY) * PT;
        return col >= 0 && col < p.cols && row >= 0 && row < p.rows && t.x > left && t.x < left + p.w * PT - 60 && box[3] - t.y > top && box[3] - t.y < top + p.h * PT;
      });
      check(
        'word: the sheet\'s PDF is one A4 page with every address printed inside its own label',
        Math.round(box[2]) === 595 && Math.round(box[3]) === 842 && firsts.length === 21 && inside.length === 21,
        JSON.stringify({ box: box.map(Math.round), addresses: firsts.length, inside: inside.length })
      );
    }

    /* ── Envelopes → Add to Document ───────────────────────────────────── */
    win.focus?.();
    const envPressed = await clickRibbon('Envelopes');
    await until(() => js(`Boolean(document.querySelector('.wd-el-delivery'))`), 'the Envelopes dialog', 5000).catch(() => {});
    await typeInto('.wd-el-delivery', ADDRESS.join('\n'));
    await typeInto('.wd-el-return', 'Rutba Office\n12 High Street\nLondon SW1A 1AA');
    await js(`(() => { const s = document.querySelector('.wd-el-size'); Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(s, 'DL'); s.dispatchEvent(new Event('change', { bubbles: true })); return 1; })()`);
    await wait(200);
    await snap(win, 'word-envelopes-dialog.png');
    const added = await clickDialog('Envelopes', 'Add to Document');
    const inModel = await until(() => model().envelope?.size === 'DL', 'the envelope in the document', 6000).catch(() => false);
    const env = model().envelope;
    const onPage = await until(() => js(`(() => {
      const s = document.querySelector('.wd-sheet.wd-envelope-sheet');
      const f = document.querySelector('.wd-page .wd-frame');
      return Boolean(s && f && /Cynthia Gartner/.test(f.innerText));
    })()`), 'the envelope sheet and its frame', 6000).catch(() => false);
    const drawn = await js(`(() => {
      const page = document.querySelector('.wd-page');
      const z = page.getBoundingClientRect().width / page.offsetWidth;
      const s = document.querySelector('.wd-sheet.wd-envelope-sheet').getBoundingClientRect();
      const f = document.querySelector('.wd-page .wd-frame').getBoundingClientRect();
      const letter = [...page.querySelectorAll('.wd-block')].find((b) => b.innerText.trim() === 'Dear Cynthia,')?.getBoundingClientRect();
      return {
        width: Math.round(s.width / z), height: Math.round(s.height / z),
        frameInside: f.left >= s.left && f.right <= s.right + 1 && f.top >= s.top && f.bottom <= s.bottom + 1,
        frameX: Math.round((f.left - s.left) / z), frameY: Math.round((f.top - s.top) / z),
        letterBelow: Boolean(letter) && letter.top > s.bottom,
        sheets: document.querySelectorAll('.wd-sheet').length,
      };
    })()`).catch((e) => ({ error: e.message }));
    check(
      'word: Envelopes → Add to Document puts a DL envelope in front as a section of its own, drawn at 220 × 110 mm with the delivery address in its frame, the letter on the next page',
      envPressed === 'clicked' && added === 'clicked' && inModel === true && onPage === true && Math.round(env?.widthPx) === 831 && Math.round(env?.heightPx) === 416
        && drawn?.width === 831 && drawn?.height === 416 && drawn?.frameInside && drawn?.letterBelow && model().sectionCount === 2,
      JSON.stringify({ size: env?.size, w: Math.round(env?.widthPx), h: Math.round(env?.heightPx), drawn, sections: model().sectionCount })
    );
    await snap(win, 'word-envelope.png');
    const pdfPath = path.join(dir, 'envelope.pdf');
    doc.export({ id: session.id, format: 'pdf', path: pdfPath });
    const boxes = mediaBoxes(fs.readFileSync(pdfPath)).map((b) => b.map(Math.round));
    check(
      'word: the letter\'s PDF begins with the envelope on DL paper, then the letter on A4',
      JSON.stringify(boxes) === JSON.stringify([[0, 0, 624, 312], [0, 0, 595, 842]]),
      JSON.stringify(boxes)
    );

    // Undo takes the envelope away again: the letter as it was.
    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    await press(win.webContents, 'z', { modifiers: ['control'] });
    const undone = await until(() => !model().envelope && js(`!document.querySelector('.wd-sheet.wd-envelope-sheet')`), 'the envelope undone', 5000).catch(() => false);
    check('word: Undo takes the envelope back off the letter', undone === true, JSON.stringify({ envelope: model().envelope, sections: model().sectionCount }));

    const complaints = await errorsIn(win);
    check('word: the envelopes and labels checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the envelopes and labels checks ran', false, err.message);
  } finally {
    if (win && saved) {
      await win.webContents.executeJavaScript(`(async () => {
        const put = (key, value) => value === null ? window.rutbaOffice.store.delete({ key }) : window.rutbaOffice.store.set({ key, value });
        await put('word.mailings.returnAddress', ${JSON.stringify(saved.ret)});
        await put('word.mailings.label', ${JSON.stringify(saved.label)});
        await put('word.mailings.envelope', ${JSON.stringify(saved.env)});
        return 1;
      })()`).catch(() => {});
    }
  }
}
