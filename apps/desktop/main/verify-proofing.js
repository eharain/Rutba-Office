// Review → Check Accessibility and Review → Spelling, in all three document
// windows, on fixtures of their own written beside the shared ones (no other
// block opens them).
//
//   a11y      Word: the pane lists a picture with no description; Edit Alt
//             Text writes one and a colour fix clears the rest, and the
//             status bar reads "Good to go". Worksheets: a picture marked
//             decorative in one press. Presentation: a slide with no title
//             given one through the fix's prompt.
//   spelling  Each app: Review → Spelling finds the planted misspelling from
//             the caret, Change writes the suggestion, Add to Dictionary
//             stops another being flagged, "Spelling check complete" ends it.
//             Word's right-click on a misspelt word offers the suggestions;
//             Worksheets offers the other sheet when the first is done.
//
// Run alone with RUTBA_VERIFY_ONLY=a11y or RUTBA_VERIFY_ONLY=spelling.

import fs from 'node:fs';
import path from 'node:path';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml/package';
import { SheetView } from '@rutba/sheet-view';
import { buildXlsx } from '@rutba/ooxml/build';
import { Deck, buildPptx } from '@rutba/presentation';
import { writeTitle } from '@rutba/proofing';
import { gradientPng } from './sample-picture.js';

/** The window helpers every step below uses. */
function helpers(win, wait) {
  const wc = win.webContents;
  const js = (code) => wc.executeJavaScript(code);
  return {
    wc,
    js,
    capture: async (name) => {
      if (!process.env.RUTBA_VERIFY_CAPTURE) return;
      wc.invalidate();
      await wait(700);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG());
    },
    tab: (name) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(), 'tab'`),
    ribbon: (title) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}) && !n.disabled);
      if (!b) return 'no button ' + ${JSON.stringify(title)};
      b.click();
      return 'clicked';
    })()`),
    click: (selector) => js(`(() => { const b = document.querySelector(${JSON.stringify(selector)}); if (!b) return 'missing ' + ${JSON.stringify(selector)}; if (b.disabled) return 'disabled ' + ${JSON.stringify(selector)}; b.click(); return 'clicked'; })()`),
    button: (tip) => js(`(() => {
      const b = [...document.querySelectorAll('.pf-pane .rw-btn, .rw-dialog .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(tip)}));
      if (!b) return 'no button ' + ${JSON.stringify(tip)};
      if (b.disabled) return 'disabled ' + ${JSON.stringify(tip)};
      b.click();
      return 'clicked';
    })()`),
    type: (selector, value) => js(`(() => {
      const el = document.querySelector(${JSON.stringify(selector)});
      if (!el) return 'no field ' + ${JSON.stringify(selector)};
      el.focus();
      const proto = el.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
      Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, ${JSON.stringify(String(value))});
      el.dispatchEvent(new Event('input', { bubbles: true }));
      return 'typed';
    })()`),
    rules: () => js(`[...document.querySelectorAll('.pf-a11y .pf-rule')].map((r) => r.dataset.rule)`),
    item: (rule) => js(`(() => { const b = document.querySelector('.pf-a11y .pf-rule[data-rule="${rule}"] .pf-item-main'); if (!b) return 'no ${rule} item'; b.click(); return 'clicked'; })()`),
    missed: () => js(`document.querySelector('.pf-editor .pf-miss')?.textContent || null`),
    suggestions: () => js(`[...document.querySelectorAll('.pf-editor .pf-suggestion')].map((b) => b.textContent.trim())`),
    done: () => js(`Boolean(document.querySelector('.pf-editor .pf-done'))`),
  };
}

/* ── the fixtures ─────────────────────────────────────────────────────── */

function a11yFixtures(dir) {
  // Word: a picture with no description and pale words; the title set, so
  // the two fixes leave nothing to find.
  const view = openDocx(buildDocx({ styles: true, paragraphs: [
    { text: 'Quarterly results', style: 'Heading1' },
    { text: 'Orders rose through the spring, as the chart shows.' },
    { text: 'These notes are printed in a very pale grey.', colour: 'CFCFCF' },
  ] }));
  view.setSelection({ block: 1, offset: 0 });
  view.insertImage({ name: 'orders.png', contentType: 'image/png', data: gradientPng(160, 90, [40, 110, 200], [230, 120, 40]), widthPx: 240, heightPx: 135 });
  const pkg = OoxmlPackage.read(view.save());
  writeTitle(pkg, 'Quarterly results');
  fs.writeFileSync(path.join(dir, 'a11y.docx'), pkg.write());

  // Worksheets: a named sheet with a logo that needs no description.
  const sheet = new SheetView(buildXlsx({ sheets: [{ name: 'Orders', rows: [['Month', 'Orders'], ['April', 120], ['May', 180], ['June', 310]] }] }));
  sheet.select(1, 3);
  sheet.insertPicture({ name: 'divider.png', contentType: 'image/png', data: gradientPng(80, 20, [200, 200, 200], [240, 240, 240]), widthPx: 120, heightPx: 30 });
  fs.writeFileSync(path.join(dir, 'a11y.xlsx'), sheet.save());

  // Presentation: a title slide, and a slide whose title placeholder is empty.
  const deck = Deck.open(buildPptx({ title: 'Checks', slides: [
    { layout: 'title', title: 'Quarterly review', body: 'Spring' },
    { layout: 'obj', title: '', body: ['Hire for support', 'Ship the mobile app'] },
  ] }));
  fs.writeFileSync(path.join(dir, 'a11y.pptx'), deck.save());
  return { docx: path.join(dir, 'a11y.docx'), xlsx: path.join(dir, 'a11y.xlsx'), pptx: path.join(dir, 'a11y.pptx') };
}

function spellFixtures(dir) {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [
    { text: 'Spelling checks', style: 'Heading1' },
    { text: 'The quick brwn fox jumps over the lazy dog.' },
    { text: 'The Zorblat team will send the report on Friday.' },
  ] }));
  fs.writeFileSync(path.join(dir, 'spell.docx'), view.save());
  fs.writeFileSync(path.join(dir, 'spell.xlsx'), buildXlsx({ sheets: [
    { name: 'Stock', rows: [['Item', 'Shade', 'Qty'], ['Pens', 'Blakc', 12], ['Ink', 'Blue', 3], ['Zorblat pads', 'Red', 5]] },
    { name: 'Notes', rows: [['Reorder next mnoth']] },
  ] }));
  const deck = Deck.open(buildPptx({ title: 'Checks', slides: [
    { layout: 'title', title: 'Quarterly reveiw', body: 'Spring' },
    { layout: 'obj', title: 'Next steps', body: ['Ask the Zorblat team'] },
  ] }));
  fs.writeFileSync(path.join(dir, 'spell.pptx'), deck.save());
  return { docx: path.join(dir, 'spell.docx'), xlsx: path.join(dir, 'spell.xlsx'), pptx: path.join(dir, 'spell.pptx') };
}

/* ── Check Accessibility ─────────────────────────────────────────────── */

export async function verifyAccessibility(h, { dir }) {
  const { open, check, until, wait, press, errorsIn, doc, sessionFor } = h;
  let files;
  try {
    files = a11yFixtures(dir);
  } catch (err) {
    return check('a11y: the fixtures were written', false, err.message);
  }

  // Rutba Word.
  try {
    const win = await open('word', files.docx);
    const w = helpers(win, wait);
    const id = sessionFor('doc').id;
    await until(() => w.js(`Boolean(document.querySelector('.wd-page [data-block="0"]'))`), 'the page', 8000);
    await w.tab('Review');
    await wait(200);
    const pressed = await w.ribbon('Check Accessibility');
    const listed = await until(async () => (await w.rules()).includes('altText') && (await w.rules()).includes('contrast'), 'the findings', 8000).catch(() => false);
    const rules = await w.rules();
    const tiers = await w.js(`[...document.querySelectorAll('.pf-a11y .pf-tier-label')].map((n) => n.textContent)`);
    check('a11y: Word — Check Accessibility opens the pane with the picture under Errors → Missing alternative text and the pale words under Warnings',
      pressed === 'clicked' && listed === true && tiers[0] === 'Errors' && tiers.includes('Warnings'), `${pressed}; rules ${JSON.stringify(rules)}; tiers ${JSON.stringify(tiers)}`);

    // Select the finding: the picture is picked on the page, the pane says why and offers the fixes.
    const selected = await w.item('altText');
    const picked = await until(() => w.js(`Boolean(document.querySelector('.wd-picture-handles, .wd-handles, [data-image].picked, .wd-image.picked, .wd-pick'))`), 'the picture picked', 3000).catch(() => false);
    const card = await until(() => w.js(`Boolean(document.querySelector('.pf-item.on .pf-action[data-fix="altText"]')) && Boolean(document.querySelector('.pf-info'))`), 'the recommended actions', 4000).catch(() => false);
    await w.capture('word-a11y.png');
    check('a11y: Word — selecting the finding goes to the picture and shows Recommended actions and why to fix it',
      selected === 'clicked' && card === true, `${selected}; card ${card}; picked ${picked}`);

    // Edit Alt Text, from the recommended action.
    const opened = await w.click('.pf-item.on .pf-action[data-fix="altText"]');
    await until(() => w.js(`Boolean(document.querySelector('.pf-alt-text'))`), 'the Alt Text dialog', 4000).catch(() => {});
    const typed = await w.type('.pf-alt-text', 'A bar chart of orders by month, rising from 120 in April to 310 in June.');
    await wait(150);
    await w.capture('word-alt-text.png');
    const ok = await w.click('.pf-alt-ok');
    const described = await until(async () => {
      const r = await doc.proof({ id, action: 'accessibility' });
      return !r.issues.some((i) => i.rule === 'altText');
    }, 'the description written', 6000).catch(() => false);
    const inFile = (await doc.proof({ id, action: 'altTextOf', drawing: 0 }))?.descr || '';
    const gone = await until(async () => !(await w.rules()).includes('altText'), 'the finding gone from the pane', 6000).catch(() => false);
    check('a11y: Word — Edit Alt Text writes the description on the picture (wp:docPr descr) and the finding leaves the pane',
      opened === 'clicked' && typed === 'typed' && ok === 'clicked' && described === true && gone === true && /bar chart of orders/.test(inFile), `${opened}/${typed}/${ok}; written ${described}; pane ${gone}; "${inFile}"`);

    // The contrast fix, then the verdict in the status bar.
    await w.item('contrast');
    await until(() => w.js(`Boolean(document.querySelector('.pf-item.on .pf-action[data-fix="textColour"]'))`), 'the colour fix', 4000).catch(() => {});
    const coloured = await w.click('.pf-item.on .pf-action[data-fix="textColour"]');
    const clean = await until(() => w.js(`Boolean(document.querySelector('.pf-a11y .pf-clean'))`), 'no issues', 6000).catch(() => false);
    await until(async () => (await w.js(`document.querySelector('.pf-status')?.textContent || ''`)).includes('Good to go'), 'the status bar verdict', 4000).catch(() => {});
    const status = await w.js(`document.querySelector('.pf-status')?.textContent || ''`);
    await w.capture('word-a11y-clean.png');
    check('a11y: Word — the colour fix clears the last finding: "No accessibility issues found", and the status bar says Accessibility: Good to go',
      coloured === 'clicked' && clean === true && status === 'Accessibility: Good to go', `${coloured}; clean ${clean}; status "${status}"`);
    // Keep accessibility checker running: the pane closed, the status bar
    // follows the document — undo the colour and it says Investigate, redo
    // and it is good again.
    const kept = await w.click('.pf-keep input');
    await w.click('.rw-panel.right .rw-panel-head .rw-btn');
    await until(async () => !(await w.js(`Boolean(document.querySelector('.pf-a11y'))`)), 'the pane closed', 3000).catch(() => {});
    await w.js(`document.querySelector('.wd-page')?.focus(), 'focused'`);
    await press(w.wc, 'z', { modifiers: ['control'] });
    const investigate = await until(async () => (await w.js(`document.querySelector('.pf-status')?.textContent || ''`)) === 'Accessibility: Investigate', 'the status bar to say Investigate', 6000).catch(() => false);
    await w.capture('word-a11y-status.png');
    await w.js(`document.querySelector('.wd-page')?.focus(), 'focused'`);
    await press(w.wc, 'y', { modifiers: ['control'] });
    const goodAgain = await until(async () => (await w.js(`document.querySelector('.pf-status')?.textContent || ''`)) === 'Accessibility: Good to go', 'the status bar good again', 6000).catch(() => false);
    check('a11y: Word — with the checker kept running the status bar follows the edits: Investigate after undoing the fix, Good to go after redoing it',
      kept === 'clicked' && investigate === true && goodAgain === true, `${kept}; investigate ${investigate}; good again ${goodAgain}`);
    // The setting is the profile's: put it back.
    await win.webContents.executeJavaScript(`window.rutbaOffice.store.set({ key: 'proofing.keepChecking', value: false })`).catch(() => {});
    const complaints = await errorsIn(win);
    check('a11y: Word — the accessibility checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('a11y: the Word accessibility checks ran', false, err.message);
  }

  // Worksheets.
  try {
    const win = await open('sheets', files.xlsx);
    const w = helpers(win, wait);
    const id = sessionFor('sheet').id;
    await until(() => w.js(`Boolean(document.querySelector('.sh-drawing'))`), 'the picture on the grid', 8000).catch(() => {});
    await w.tab('Review');
    await wait(200);
    const pressed = await w.ribbon('Check Accessibility');
    const listed = await until(async () => (await w.rules()).includes('altText'), 'the picture finding', 8000).catch(() => false);
    await w.item('altText');
    await until(async () => (await w.js(`document.querySelector('.sh-namebox')?.textContent || ''`)) === 'D2', 'the anchor cell selected', 4000).catch(() => {});
    const goneTo = await w.js(`document.querySelector('.sh-namebox')?.textContent || ''`);
    await until(() => w.js(`Boolean(document.querySelector('.pf-item.on .pf-action[data-fix="decorative"]'))`), 'Mark as decorative', 4000).catch(() => {});
    await w.capture('sheets-a11y.png');
    const marked = await w.click('.pf-item.on .pf-action[data-fix="decorative"]');
    const decorative = await until(async () => (await doc.proof({ id, action: 'altTextOf', sheet: 'Orders', anchor: 0 }))?.decorative === true, 'decorative written', 6000).catch(() => false);
    const clean = await until(() => w.js(`Boolean(document.querySelector('.pf-a11y .pf-clean'))`), 'no issues', 6000).catch(() => false);
    check('a11y: Worksheets — the picture is listed at its cell, and Mark as decorative writes the adec extension on xdr:cNvPr in one press',
      pressed === 'clicked' && listed === true && goneTo === 'D2' && marked === 'clicked' && decorative === true && clean === true, `${pressed}; listed ${listed}; went to ${goneTo}; ${marked}; decorative ${decorative}; clean ${clean}`);
    const complaints = await errorsIn(win);
    check('a11y: Worksheets — the accessibility checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('a11y: the Worksheets accessibility checks ran', false, err.message);
  }

  // Presentation.
  try {
    const win = await open('slides', files.pptx);
    const w = helpers(win, wait);
    const id = sessionFor('deck').id;
    await until(() => w.js(`document.querySelectorAll('.sl-thumb').length >= 2`), 'the slide sorter', 8000).catch(() => {});
    await w.tab('Review');
    await wait(200);
    const pressed = await w.ribbon('Check Accessibility');
    const listed = await until(async () => (await w.rules()).includes('slideTitle'), 'the missing title', 8000).catch(() => false);
    await w.item('slideTitle');
    const turned = await until(() => w.js(`document.querySelectorAll('.sl-thumb')[1]?.classList.contains('active')`), 'slide 2 shown', 4000).catch(() => false);
    await until(() => w.js(`Boolean(document.querySelector('.pf-item.on .pf-action[data-fix="slideTitle"]'))`), 'the title fix', 4000).catch(() => {});
    await w.capture('slides-a11y.png');
    const asked = await w.click('.pf-item.on .pf-action[data-fix="slideTitle"]');
    await until(() => w.js(`Boolean(document.querySelector('.pf-slidetitle'))`), 'the title prompt', 4000).catch(() => {});
    const typed = await w.type('input.pf-slidetitle', 'Next steps');
    const ok = await w.click('.pf-prompt-ok');
    const titled = await until(() => (doc.model({ id, slide: 1 }).slide?.shapes || []).some((s) => (s.text?.paragraphs || []).some((p) => p.plain === 'Next steps')), 'the title written', 6000).catch(() => false);
    const gone = await until(async () => !(await w.rules()).includes('slideTitle'), 'the finding gone', 6000).catch(() => false);
    check('a11y: Presentation — a slide with no title is listed, the fix turns to it and its prompt types the title into the placeholder',
      pressed === 'clicked' && listed === true && turned === true && asked === 'clicked' && typed === 'typed' && ok === 'clicked' && titled === true && gone === true,
      `${pressed}; listed ${listed}; turned ${turned}; ${asked}/${typed}/${ok}; titled ${titled}; gone ${gone}`);
    const complaints = await errorsIn(win);
    check('a11y: Presentation — the accessibility checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('a11y: the Presentation accessibility checks ran', false, err.message);
  }
}

/* ── Spelling ────────────────────────────────────────────────────────── */

export async function verifySpelling(h, { dir }) {
  const { open, check, until, wait, press, errorsIn, doc, sessionFor } = h;
  let files;
  try {
    files = spellFixtures(dir);
  } catch (err) {
    return check('spelling: the fixtures were written', false, err.message);
  }
  // The person's own dictionary is the profile's: put it back as it was.
  const dictionaryBefore = (await doc.proof({ action: 'dictionary' }))?.words || [];

  /** One pass: find `word`, Change it to `to`; find `add`, Add to Dictionary; then complete. */
  const pass = async (w, { word, to, add }) => {
    // `until` answers true or throws: wait, then read what is there.
    const found = (await until(async () => (await w.missed()) === word, `"${word}" in the Editor pane`, 15000).catch(() => false)) ? word : await w.missed();
    // The suggestions follow the word from a thread of their own.
    await until(async () => (await w.suggestions()).includes(to), `the suggestions for "${word}"`, 10000).catch(() => {});
    const suggestions = await w.suggestions();
    const chosen = await w.js(`(() => { const b = [...document.querySelectorAll('.pf-editor .pf-suggestion')].find((n) => n.textContent.trim() === ${JSON.stringify(to)}); if (!b) return 'no suggestion'; b.click(); return 'chosen'; })()`);
    await wait(150);
    const changed = await w.button('Change — this one');
    const next = (await until(async () => (await w.missed()) === add, `"${add}" next`, 10000).catch(() => false)) ? add : await w.missed();
    const added = await w.button('Add to Dictionary');
    const done = await until(() => w.done(), 'Spelling check complete', 10000).catch(() => false);
    const message = await w.js(`document.querySelector('.pf-editor .pf-done')?.textContent || ''`);
    return { found, suggestions, chosen, changed, next, added, done, message };
  };

  // Rutba Word, from the caret at the start of the document.
  try {
    const win = await open('word', files.docx);
    const w = helpers(win, wait);
    const id = sessionFor('doc').id;
    await until(() => w.js(`Boolean(document.querySelector('.wd-page [data-block="1"]'))`), 'the page', 8000);

    // Right-click on the misspelt word: the suggestions head the menu.
    const menu = await w.js(`(() => {
      const block = document.querySelector('.wd-page [data-block="1"]');
      const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
      for (let n = walker.nextNode(); n; n = walker.nextNode()) {
        const at = n.nodeValue.indexOf('brwn');
        if (at < 0) continue;
        const r = document.createRange(); r.setStart(n, at + 1); r.setEnd(n, at + 2);
        const box = r.getBoundingClientRect();
        block.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: box.left + box.width / 2, clientY: box.top + box.height / 2 }));
        return 'opened';
      }
      return 'no word';
    })()`);
    const menuItems = () => w.js(`[...document.querySelectorAll('.rw-menu button')].map((b) => b.textContent.trim())`);
    await until(async () => { const l = await menuItems(); return l.includes('brown') && l.includes('Add to Dictionary'); }, 'the spelling menu', 15000).catch(() => {});
    const offered = await menuItems();
    await w.capture('word-spelling-menu.png');
    check('spelling: Word — a right-click on a misspelt word offers the suggestions, Ignore All and Add to Dictionary above the usual verbs',
      menu === 'opened' && offered.includes('brown') && offered.includes('Add to Dictionary') && offered.indexOf('brown') < offered.findIndex((t) => t.startsWith('Undo')), `${menu}; ${JSON.stringify(offered)}`);
    await press(w.wc, 'Escape');
    await wait(200);

    await w.tab('Review');
    await wait(200);
    const pressed = await w.ribbon('Spelling');
    await until(async () => (await w.missed()) === 'brwn', 'the first misspelling', 15000).catch(() => {});
    const first = await w.missed();
    const context = await w.js(`document.querySelector('.pf-editor .pf-context')?.textContent || ''`);
    await until(async () => (await w.js(`window.getSelection()?.toString() || ''`)) === 'brwn', 'the word selected on the page', 4000).catch(() => {});
    const selectedOnPage = await w.js(`window.getSelection()?.toString() || ''`);
    await w.capture('word-editor.png');
    const r = await pass(w, { word: 'brwn', to: 'brown', add: 'Zorblat' });
    const text = (doc.model({ id }).blocks || []).map((b) => b.text).join(' | ');
    check('spelling: Word — Review → Spelling finds the planted misspelling from the caret, in its sentence, selected on the page, with suggestions',
      pressed === 'clicked' && first === 'brwn' && /quick brwn fox/.test(context) && selectedOnPage === 'brwn' && r.suggestions.includes('brown'), `${pressed}; ${first}; "${context}"; selected "${selectedOnPage}"; ${JSON.stringify(r.suggestions)}`);
    check('spelling: Word — Change writes the suggestion into the document', r.chosen === 'chosen' && r.changed === 'clicked' && text.includes('The quick brown fox'), `${r.chosen}/${r.changed}; ${text}`);
    const dictionary = (await doc.proof({ action: 'dictionary' }))?.words || [];
    check('spelling: Word — Add to Dictionary puts the word in the person\'s own dictionary and the pass ends with "Spelling check complete"',
      r.next === 'Zorblat' && r.added === 'clicked' && dictionary.includes('Zorblat') && r.done === true && /Spelling check complete/.test(r.message), `${r.next}; ${r.added}; ${JSON.stringify(dictionary)}; ${r.done} "${r.message}"`);
    await w.capture('word-editor-done.png');

    // A second pass: nothing left, the added word included.
    await press(w.wc, 'F7');
    await wait(600);
    const again = await until(async () => (await w.done()) && !(await w.missed()), 'a second pass to complete at once', 10000).catch(() => false);
    const flagged = await w.missed();
    check('spelling: Word — F7 again finds nothing: the added word is no longer flagged', again === true && !flagged, `done ${again}; flagged ${flagged}`);

    // Options → Custom Dictionary: the word is listed; removed there, it is flagged again.
    await w.click('.pf-options > .pf-link');
    await until(() => w.js(`Boolean(document.querySelector('.pf-options-body'))`), 'the options', 3000).catch(() => {});
    await w.capture('word-editor-options.png');
    await w.js(`[...document.querySelectorAll('.pf-options-body .pf-link')].find((b) => b.textContent.includes('Custom Dictionary'))?.click(), 'opened'`);
    const listed = await until(() => w.js(`Boolean(document.querySelector('.pf-dict-row[data-word="Zorblat"]'))`), 'the dictionary dialog', 4000).catch(() => false);
    await w.capture('word-dictionary.png');
    await w.click('.pf-dict-row[data-word="Zorblat"] .pf-dict-remove');
    await w.click('.pf-dict-ok');
    const removed = await until(async () => !((await doc.proof({ action: 'dictionary' }))?.words || []).includes('Zorblat'), 'the word out of the dictionary', 4000).catch(() => false);
    await press(w.wc, 'F7');
    const flaggedAgain = await until(async () => (await w.missed()) === 'Zorblat', 'the word flagged again', 10000).catch(() => false);
    check('spelling: Word — the Custom Dictionary dialog lists the added word; removed there, the next pass flags it again',
      listed === true && removed === true && flaggedAgain === true, `listed ${listed}; removed ${removed}; flagged again ${flaggedAgain}`);
    const complaints = await errorsIn(win);
    check('spelling: Word — the spelling checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('spelling: the Word spelling checks ran', false, err.message);
  }

  // Worksheets: the active sheet, then the other one when asked.
  try {
    await doc.proof({ action: 'dictionary', words: dictionaryBefore });
    const win = await open('sheets', files.xlsx);
    const w = helpers(win, wait);
    const id = sessionFor('sheet').id;
    await until(() => w.js(`document.querySelectorAll('.sh-cell').length > 3`), 'the grid', 8000).catch(() => {});
    await w.tab('Review');
    await wait(200);
    const pressed = await w.ribbon('Spelling');
    await until(async () => Boolean(await w.js(`document.querySelector('.pf-editor .pf-context-where')?.textContent || ''`)), 'the cell named', 15000).catch(() => {});
    const where = await w.js(`document.querySelector('.pf-editor .pf-context-where')?.textContent || ''`);
    await w.capture('sheets-editor.png');
    const r = await pass(w, { word: 'Blakc', to: 'Black', add: 'Zorblat' });
    const blue = await until(() => w.js(`[...document.querySelectorAll('.sh-cell')].some((c) => c.textContent.trim() === 'Black')`), 'the cell drawn with "Black"', 6000).catch(() => false);
    const more = await w.js(`[...document.querySelectorAll('.pf-editor .pf-done .rw-btn')].map((b) => b.textContent.trim())`);
    check('spelling: Worksheets — the pass reads the active sheet\'s cells from A1, Change writes "Black" into the cell, Add to Dictionary, then complete for the sheet',
      pressed === 'clicked' && r.found === 'Blakc' && r.changed === 'clicked' && blue === true && r.next === 'Zorblat' && r.added === 'clicked' && r.done === true && where === 'Stock!B2', `${pressed}; ${r.found} at ${where}; ${r.changed}; cell ${blue}; ${r.next}/${r.added}; ${r.done} "${r.message}"; id ${id}`);
    const other = await w.js(`(() => { const b = document.querySelector('.pf-editor .pf-more'); if (!b) return 'no button'; b.click(); return 'clicked'; })()`);
    await until(async () => (await w.missed()) === 'mnoth', 'the other sheet word', 15000).catch(() => {});
    const onOther = await w.missed();
    check('spelling: Worksheets — when the sheet is done the pane offers the other sheet, and that pass finds its word', other === 'clicked' && onOther === 'mnoth', `${JSON.stringify(more)}; ${other}; ${onOther}`);
    await w.capture('sheets-editor-other.png');
    const complaints = await errorsIn(win);
    check('spelling: Worksheets — the spelling checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('spelling: the Worksheets spelling checks ran', false, err.message);
  }

  // Presentation: every slide from the first.
  try {
    await doc.proof({ action: 'dictionary', words: dictionaryBefore });
    const win = await open('slides', files.pptx);
    const w = helpers(win, wait);
    const id = sessionFor('deck').id;
    await until(() => w.js(`document.querySelectorAll('.sl-thumb').length >= 2`), 'the slide sorter', 8000).catch(() => {});
    await w.tab('Review');
    await wait(200);
    const pressed = await w.ribbon('Spelling');
    await until(async () => (await w.missed()) === 'reveiw', 'the first word', 15000).catch(() => {});
    await w.capture('slides-editor.png');
    const r = await pass(w, { word: 'reveiw', to: 'review', add: 'Zorblat' });
    const titled = (doc.model({ id, slide: 0 }).slide?.shapes || []).some((s) => (s.text?.paragraphs || []).some((p) => p.plain === 'Quarterly review'));
    const turned = await w.js(`document.querySelectorAll('.sl-thumb')[1]?.classList.contains('active')`);
    check('spelling: Presentation — the pass finds the slide title\'s misspelling, Change rewrites it in its run, turns to slide 2 for the next, and completes',
      pressed === 'clicked' && r.found === 'reveiw' && r.changed === 'clicked' && titled && r.next === 'Zorblat' && r.added === 'clicked' && r.done === true,
      `${pressed}; ${r.found}; ${r.changed}; titled ${titled}; ${r.next} (slide 2 shown ${turned}); ${r.added}; ${r.done}`);
    await w.capture('slides-editor-done.png');
    const complaints = await errorsIn(win);
    check('spelling: Presentation — the spelling checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('spelling: the Presentation spelling checks ran', false, err.message);
  } finally {
    await doc.proof({ action: 'dictionary', words: dictionaryBefore }).catch(() => {});
  }
}
