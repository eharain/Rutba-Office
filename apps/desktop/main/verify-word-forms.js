// Documents: forms. Developer puts a plain text control, a drop-down list,
// a date picker, a check box, a picture control and legacy form fields at
// the caret; Properties gives one a title, a tag and a placeholder; each is
// filled the way Word's is (words typed in place, an item picked from the
// list that drops, a day from the calendar, a tick, a picture from a file),
// Design Mode shows the tags; the saved file carries what Word writes; and a
// copy protected for Filling in forms takes nothing but the fields, Tab
// walking them. Run alone with RUTBA_VERIFY_ONLY=word-forms.

import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';
import { OoxmlPackage, crc32 } from '@rutba/ooxml';
import { buildDocx } from '@rutba/ooxml/build';

/** A small landscape as a PNG, made here: sky, a sun, a green hill. */
function landscapePng(w = 160, h = 120) {
  const rows = [];
  for (let y = 0; y < h; y++) {
    const row = Buffer.alloc(1 + w * 3);
    for (let x = 0; x < w; x++) {
      let c = [120 + Math.round((y / h) * 80), 170 + Math.round((y / h) * 50), 235];
      if ((x - 118) ** 2 + (y - 32) ** 2 < 220) c = [250, 210, 80];
      if (y > 78 + 18 * Math.sin(x / 26)) c = [70, 140 - Math.round((y - 78) / 2), 70];
      row.set(c, 1 + x * 3);
    }
    rows.push(row);
  }
  const chunk = (type, data) => {
    const len = Buffer.alloc(4);
    len.writeUInt32BE(data.length);
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crc32(Buffer.concat([Buffer.from(type), data])) >>> 0);
    return Buffer.concat([len, Buffer.from(type), data, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0);
  ihdr.writeUInt32BE(h, 4);
  ihdr[8] = 8;
  ihdr[9] = 2;
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(Buffer.concat(rows))), chunk('IEND', Buffer.alloc(0))]);
}
const PNG = landscapePng();

/** A form to fill: six labelled lines and a settings part, as Word always writes one. */
function form() {
  const pkg = OoxmlPackage.read(buildDocx({ styles: true, paragraphs: [{ text: 'Name: ' }, { text: 'City: ' }, { text: 'Start date: ' }, { text: 'I agree: ' }, { text: 'Photo: ' }, { text: 'Legacy: ' }] }));
  pkg.addPart('word/settings.xml', Buffer.from('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/></w:settings>', 'utf8'), 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');
  pkg.addRelationshipTo('word/document.xml', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings', 'settings.xml');
  return pkg;
}

/**
 * @param {object} h the harness: open, check, until, wait, capture, doc
 */
export async function verifyWordForms(h, { dir }) {
  const { open, check, until, wait, capture, doc } = h;
  const own = path.join(dir, 'word-forms');
  const file = path.join(own, 'application.docx');
  const picture = path.join(own, 'photo.png');
  const lockedFile = path.join(own, 'application-locked.docx');
  try {
    fs.mkdirSync(own, { recursive: true });
    fs.writeFileSync(file, form().write());
    fs.writeFileSync(picture, PNG);
  } catch (err) {
    return check('word-forms: the fixtures were made', false, err.message);
  }

  let win = null;
  try {
    win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const session = doc.sessions().filter((s) => s.kind === 'doc' && s.path === file).pop();
    const model = () => doc.model({ id: session.id });
    const fresh = async (name) => {
      wc.invalidate();
      await wait(700);
      await capture(win, name);
    };
    const clickTab = (name) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(), 'tab'`);
    const clickRibbon = (tip) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(tip)}));
      if (!b) return 'no button ' + ${JSON.stringify(tip)};
      if (b.disabled) return 'disabled';
      b.click();
      return 'clicked';
    })()`);
    const caretAtEnd = async (block) => {
      const text = model().blocks[block].text;
      await doc.apply({ id: session.id, ops: [{ op: 'setSelection', anchor: { block, offset: text.length }, focus: { block, offset: text.length } }] });
    };
    const host = (kind) => `document.querySelector('.wd-page .wd-cc[data-kind="${kind}"]')`;
    const hostState = (kind) => js(`(() => { const h = ${host(kind)}; return h ? { value: h.dataset.value, title: h.dataset.ccTitle || '', placeholder: h.dataset.placeholder === '1', design: h.classList.contains('design'), shown: (h.shadowRoot?.querySelector('.box')?.textContent || '').trim() } : null; })()`);
    const press = (kind) => js(`(() => { const h = ${host(kind)}; if (!h) return 'none'; h.scrollIntoView({ block: 'center' }); h.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 })); h.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, button: 0 })); return 'pressed'; })()`);
    const fields = () => (model().blocks || []).flatMap((b, block) => (b.runs || []).filter((r) => r.control).map((r) => ({ block, ...r.control })));
    const insert = async (block, tip, kind) => {
      await caretAtEnd(block);
      await wait(150);
      const pressed = await clickRibbon(tip);
      const there = await until(() => js(`Boolean(${host(kind)})`), `the ${kind} control on the page`, 6000).catch(() => false);
      return pressed === 'clicked' && there;
    };

    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="5"]'))`), 'the form', 10000);
    await clickTab('Developer');
    await until(() => js(`[...document.querySelectorAll('.rw-ribbon .rw-btn')].some((n) => (n.dataset.tip || '').startsWith('Rich Text Content Control'))`), 'the Developer tab', 4000).catch(() => {});
    const tabOk = await js(`['Rich Text Content Control', 'Plain Text Content Control', 'Picture Content Control', 'Check Box Content Control', 'Combo Box Content Control', 'Drop-Down List Content Control', 'Date Picker Content Control', 'Legacy Tools', 'Design Mode', 'Properties', 'Restrict Editing'].every((tip) => [...document.querySelectorAll('.rw-ribbon .rw-btn')].some((n) => (n.dataset.tip || n.title || '').startsWith(tip)))`);

    // 1. Each kind at the caret.
    const put = [
      await insert(0, 'Plain Text Content Control', 'plainText'),
      await insert(1, 'Drop-Down List Content Control', 'dropDownList'),
      await insert(2, 'Date Picker Content Control', 'date'),
      await insert(3, 'Check Box Content Control', 'checkBox'),
      await insert(4, 'Picture Content Control', 'picture'),
    ];
    const plain = await until(async () => (await hostState('plainText'))?.placeholder === true, 'the placeholder', 4000).then(() => hostState('plainText')).catch(() => null);
    check(
      'word-forms: Developer puts a plain text control, a drop-down list, a date picker, a check box and a picture control at the caret, each showing Word\'s placeholder',
      tabOk === true && put.every(Boolean) && plain?.shown === 'Click or tap here to enter text.' && fields().length === 5,
      JSON.stringify({ tabOk, put, plain, kinds: fields().map((f) => f.kind) }),
    );

    // 2. Properties: a title, a tag, a placeholder of one's own; a list's items.
    await press('plainText');
    await wait(250);
    await js(`document.activeElement?.blur?.(), 1`);
    await clickRibbon('Properties');
    const dialog = await until(() => js(`Boolean(document.querySelector('.wd-cc-title'))`), 'Content Control Properties', 4000).catch(() => false);
    const setInput = (sel, value) => js(`(() => { const el = document.querySelector(${JSON.stringify(sel)}); if (!el) return false; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, ${JSON.stringify(value)}); el.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
    await setInput('.wd-cc-title', 'Full name');
    await setInput('.wd-cc-tag', 'name');
    await setInput('.wd-cc-placeholder', 'Your name as on your passport');
    await fresh('word-forms-properties.png');
    await js(`document.querySelector('.wd-cc-props-ok')?.click(), 1`);
    const titled = await until(async () => { const s = await hostState('plainText'); return s?.title === 'Full name' && s.shown.includes('Your name as on your passport'); }, 'the title and placeholder', 6000).catch(() => false);

    await press('dropDownList');
    await wait(200);
    await js(`document.activeElement?.blur?.(), 1`);
    await js(`document.querySelectorAll('.wd-cc-pop').length && document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })), 1`);
    await clickRibbon('Properties');
    await until(() => js(`Boolean(document.querySelector('.wd-cc-add'))`), 'the list properties', 4000).catch(() => {});
    for (const [text, value] of [['Lahore', 'LHE'], ['Karachi', 'KHI'], ['Gilgit', 'GIL']]) {
      await js(`document.querySelector('.wd-cc-add')?.click(), 1`);
      await wait(120);
      await js(`(() => { const rows = document.querySelectorAll('.wd-cc-item-text'); const vals = document.querySelectorAll('.wd-cc-item-value'); const set = (el, v) => { Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, v); el.dispatchEvent(new Event('input', { bubbles: true })); }; set(rows[rows.length - 1], ${JSON.stringify(text)}); set(vals[vals.length - 1], ${JSON.stringify(value)}); return rows.length; })()`);
      await wait(80);
    }
    await js(`document.querySelector('.wd-cc-props-ok')?.click(), 1`);
    const listed = await until(() => (fields().find((f) => f.kind === 'dropDownList')?.items || []).length === 4, 'the list items', 6000).catch(() => false);
    check(
      'word-forms: Properties sets a control\'s title, tag and a placeholder of its own, and a drop-down list\'s items with their values',
      dialog === true && titled === true && listed === true && fields()[0].tag === 'name',
      JSON.stringify({ dialog, titled, listed, items: fields().find((f) => f.kind === 'dropDownList')?.items }),
    );

    // 3. Filled as Word fills them.
    await press('plainText');
    const editing = await until(() => js(`Boolean(${host('plainText')}?.shadowRoot?.querySelector('.edit'))`), 'typing in place', 4000).catch(() => false);
    wc.insertText('Ayesha Khan');
    await wait(150);
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'Tab' });
    wc.sendInputEvent({ type: 'keyUp', keyCode: 'Tab' });
    const typed = await until(async () => (await hostState('plainText'))?.value === 'Ayesha Khan', 'the words in the control', 6000).catch(() => false);
    // Tab went on to the drop-down list, whose list drops.
    const dropped = await until(() => js(`Boolean(document.querySelector('.wd-cc-list [data-item="Karachi"]'))`), 'the list to drop', 5000).catch(() => false);
    await fresh('word-forms-list.png');
    await js(`document.querySelector('.wd-cc-list [data-item="Karachi"]')?.click(), 1`);
    const picked = await until(async () => (await hostState('dropDownList'))?.value === 'Karachi', 'the item chosen', 6000).catch(() => false);
    await press('date');
    const calendar = await until(() => js(`Boolean(document.querySelector('.wd-cc-cal .wd-cc-day'))`), 'the calendar', 5000).catch(() => false);
    await fresh('word-forms-calendar.png');
    const today = await js(`document.querySelector('.wd-cc-cal .wd-cc-day.today')?.dataset.date || null`);
    await js(`document.querySelector('.wd-cc-cal .wd-cc-day.today')?.click(), 1`);
    const dated = await until(() => fields().find((f) => f.kind === 'date')?.date === today, 'the date', 6000).catch(() => false);
    await press('checkBox');
    const ticked = await until(async () => (await hostState('checkBox'))?.value === 'true', 'the tick', 5000).catch(() => false);
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'open', answer: [picture] });
    await press('picture');
    const pictured = await until(() => fields().find((f) => f.kind === 'picture')?.showingPlaceholder === false && /^data:image\/png/.test(fields().find((f) => f.kind === 'picture')?.href || ''), 'the picture', 8000).catch(() => false);
    check(
      'word-forms: words are typed in place and Tab goes on, the list drops and an item is picked, the calendar gives a day, a press ticks the box, a picture comes from a file',
      editing === true && typed === true && dropped === true && picked === true && calendar === true && dated === true && ticked === true && pictured === true,
      JSON.stringify({ editing, typed, dropped, picked, calendar, today, dated, ticked, pictured }),
    );
    await fresh('word-forms-filled.png');

    // 4. Legacy Tools: a text form field and a check box form field, filled.
    await caretAtEnd(5);
    await wait(150);
    const legacyMenu = async (label) => {
      await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || '').startsWith('Legacy Tools')); b?.click(); return Boolean(b); })()`);
      await until(() => js(`[...document.querySelectorAll('.rw-menu button, .rw-menu [role="menuitem"]')].some((n) => n.textContent.trim() === ${JSON.stringify(label)})`), 'the Legacy Tools menu', 3000).catch(() => {});
      return js(`(() => { const b = [...document.querySelectorAll('.rw-menu button, .rw-menu [role="menuitem"]')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); b?.click(); return Boolean(b); })()`);
    };
    const textField = await legacyMenu('Text Form Field');
    await until(() => js(`Boolean(${host('formText')})`), 'the text form field', 5000).catch(() => {});
    const boxField = await legacyMenu('Check Box Form Field');
    await until(() => js(`Boolean(${host('formCheckBox')})`), 'the check box form field', 5000).catch(() => {});
    await press('formText');
    await until(() => js(`Boolean(${host('formText')}?.shadowRoot?.querySelector('.edit'))`), 'typing in the field', 4000).catch(() => {});
    wc.insertText('Old form');
    await wait(120);
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'Return' });
    wc.sendInputEvent({ type: 'keyUp', keyCode: 'Return' });
    const legacyTyped = await until(async () => (await hostState('formText'))?.value === 'Old form', 'the legacy text', 5000).catch(() => false);
    await press('formCheckBox');
    const legacyTicked = await until(async () => (await hostState('formCheckBox'))?.value === 'true', 'the legacy tick', 5000).catch(() => false);
    const shaded = await js(`getComputedStyle(${host('formText')}.shadowRoot.querySelector('.legacy')).backgroundColor`);
    check(
      'word-forms: Legacy Tools puts a text form field and a check box form field in, shaded grey as Word shades them, and they are filled',
      textField === true && boxField === true && legacyTyped === true && legacyTicked === true && shaded === 'rgb(217, 217, 217)',
      JSON.stringify({ textField, boxField, legacyTyped, legacyTicked, shaded }),
    );

    // 5. Design Mode shows each control's tags.
    await clickRibbon('Design Mode');
    const design = await until(async () => (await hostState('plainText'))?.design === true, 'Design Mode', 4000).catch(() => false);
    const tag = await js(`(() => { const t = ${host('plainText')}?.shadowRoot?.querySelector('.tag.start'); return t ? { text: t.textContent, shown: getComputedStyle(t).display !== 'none' } : null; })()`);
    await fresh('word-forms-design.png');
    await clickRibbon('Design Mode');
    await until(async () => (await hostState('plainText'))?.design === false, 'Design Mode off', 4000).catch(() => {});
    check('word-forms: Design Mode shows each control\'s start and end tags, named by its title', design === true && tag?.shown === true && tag.text === 'Full name', JSON.stringify({ design, tag }));

    // 6. Saved as Word saves them.
    await clickRibbon('Save');
    await until(() => doc.meta({ id: session.id }).dirty === false, 'the save', 15000).catch(() => false);
    const saved = OoxmlPackage.read(fs.readFileSync(file));
    const body = saved.text('word/document.xml');
    const fileOk = {
      plain: /<w:alias w:val="Full name"\/><w:tag w:val="name"\/>/.test(body) && /<w:t xml:space="preserve">Ayesha Khan<\/w:t>/.test(body),
      list: /<w:dropDownList w:lastValue="KHI">/.test(body),
      date: new RegExp(`<w:date w:fullDate="${today}T00:00:00Z">`).test(body),
      box: /<w14:checked w14:val="1"\/>/.test(body),
      picture: /<w:picture\/><\/w:sdtPr><w:sdtContent><w:r><w:rPr><w:noProof\/><\/w:rPr><w:drawing>/.test(body) && !/<w:picture\/>[\s\S]{0,40}showingPlcHdr/.test(body),
      glossary: saved.has('word/glossary/document.xml') && /Your name as on your passport/.test(saved.text('word/glossary/document.xml')),
      legacy: /FORMTEXT <\/w:instrText><\/w:r><w:r><w:fldChar w:fldCharType="separate"\/><\/w:r><w:r><w:rPr><w:noProof\/><\/w:rPr><w:t xml:space="preserve">Old form<\/w:t>/.test(body) && /<w:checked w:val="1"\/><\/w:checkBox>/.test(body),
      placeholdersGone: (body.match(/<w:showingPlcHdr\/>/g) || []).length === 0,
    };
    check(
      'word-forms: the saved file holds what Word writes, the words without the placeholder flag, the list\'s w:lastValue, the date\'s w:fullDate, w14:checked, the picture, the placeholder in the glossary, the legacy field\'s result and w:checked',
      Object.values(fileOk).every(Boolean),
      JSON.stringify(fileOk),
    );

    // 7. Restrict Editing → Filling in forms: only the fields take input, Tab walks them.
    const locked = OoxmlPackage.read(fs.readFileSync(file));
    locked.write_('word/settings.xml', locked.text('word/settings.xml').replace('<w:defaultTabStop', '<w:documentProtection w:edit="forms" w:enforcement="1"/><w:defaultTabStop'));
    fs.writeFileSync(lockedFile, locked.write());
    const win2 = await open('word', lockedFile);
    const wc2 = win2.webContents;
    const js2 = (code) => wc2.executeJavaScript(code);
    const s2 = doc.sessions().filter((s) => s.kind === 'doc' && s.path === lockedFile).pop();
    await until(() => js2(`Boolean(document.querySelector('.wd-page .wd-cc[data-kind="plainText"]'))`), 'the protected form', 10000);
    // A keystroke in the words is refused.
    const at = await js2(`(() => { const b = document.querySelector('.wd-page [data-block="0"]'); const q = b.getBoundingClientRect(); return { x: Math.round(q.left + 6), y: Math.round(q.top + q.height / 2) }; })()`);
    wc2.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
    wc2.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1 });
    await wait(300);
    const before = doc.model({ id: s2.id }).blocks[0].text;
    wc2.sendInputEvent({ type: 'keyDown', keyCode: 'Z' });
    wc2.sendInputEvent({ type: 'char', keyCode: 'Z' });
    wc2.sendInputEvent({ type: 'keyUp', keyCode: 'Z' });
    const refused = await until(() => js2(`/filling in forms|form fields/i.test([...document.querySelectorAll('.rw-toast')].map((n) => n.textContent).join(' '))`), 'the refusal', 5000).catch(() => false);
    const unchanged = doc.model({ id: s2.id }).blocks[0].text === before;
    // Tab goes to the first field, typed in place; a Tab more to the list.
    wc2.sendInputEvent({ type: 'keyDown', keyCode: 'Tab' });
    wc2.sendInputEvent({ type: 'keyUp', keyCode: 'Tab' });
    const tabbed = await until(() => js2(`Boolean(document.querySelector('.wd-page .wd-cc[data-kind="plainText"]')?.shadowRoot?.querySelector('.edit'))`), 'Tab to the first field', 5000).catch(() => false);
    await js2(`(() => { const e = document.querySelector('.wd-page .wd-cc[data-kind="plainText"]').shadowRoot.querySelector('.edit'); const r = document.createRange(); r.selectNodeContents(e); const s = e.getRootNode().getSelection ? e.getRootNode().getSelection() : getSelection(); s.removeAllRanges(); s.addRange(r); return 1; })()`);
    wc2.insertText('Bilal Ahmed');
    await wait(120);
    wc2.sendInputEvent({ type: 'keyDown', keyCode: 'Tab' });
    wc2.sendInputEvent({ type: 'keyUp', keyCode: 'Tab' });
    const filled = await until(() => (doc.model({ id: s2.id }).blocks[0].runs || []).some((r) => r.control?.value === 'Bilal Ahmed'), 'the field filled under protection', 6000).catch(() => false);
    const walked = await until(() => js2(`Boolean(document.querySelector('.wd-cc-list'))`), 'Tab on to the list', 5000).catch(() => false);
    await win2.webContents.invalidate();
    await wait(700);
    await capture(win2, 'word-forms-protected.png');
    check(
      'word-forms: protected for Filling in forms, a keystroke in the words is refused with Word\'s sentence, Tab goes to the first field and on to the next, and the fields are filled',
      refused === true && unchanged && tabbed === true && filled === true && walked === true,
      JSON.stringify({ refused, unchanged, tabbed, filled, walked }),
    );
  } catch (err) {
    check('word-forms: the forms checks ran', false, err.stack || err.message);
  }
}
