// PDF: comments and markup, pressed in the window and read back from the
// file it saves. A two-page PDF made here already carries a rectangle
// another writer drew; the Comment tab opens the Comments pane with it
// listed under its author. Words selected are highlighted at a press of
// Highlight; with Underline in hand, letting go of a selection underlines
// it; a sticky note and a text box are placed and typed into; a freehand
// line, a rectangle, an ellipse, a line and an arrow are dragged out on the
// second page; a comment is answered, marked resolved, and deleted with
// Undo bringing it back and Redo taking it away again; Save As writes a
// file that begins with every byte of the original, its comments after it,
// each read back with its words, its author and its appearance, the other
// writer's rectangle among them.
// Run alone with RUTBA_VERIFY_ONLY=pdf-annot.

import fs from 'node:fs';
import path from 'node:path';
import pdfWriter from '@rutba/pdf';
import pdfReader from '@rutba/pdf/reader';
import pdfUpdate from '@rutba/pdf/update';
import pdfAssemble from '@rutba/pdf/assemble';
import pdfAnnotations from '@rutba/pdf/annotations';

const { PdfDocument } = pdfWriter;
const { openPdf, Dict, Name } = pdfReader;
const { PdfUpdate } = pdfUpdate;
const { textString } = pdfAssemble;
const { readAnnotations } = pdfAnnotations;

/** Two pages of words, the first with a rectangle another writer drew over it as a comment. */
function makePdf(file) {
  const doc = new PdfDocument({ title: 'Commented report', created: '2026-10-14T09:00:00Z' });
  const one = doc.addPage();
  one.text('Quarterly report', 72, 110, { font: 'Helvetica-Bold', size: 28 });
  one.text('The quick brown fox jumps over the lazy dog.', 72, 160, { font: 'Helvetica', size: 14 });
  one.text('Sales rose in every region this quarter.', 72, 190, { font: 'Helvetica', size: 14 });
  one.text('Costs were held at last year’s level.', 72, 220, { font: 'Helvetica', size: 14 });
  const two = doc.addPage();
  two.text('Second page', 72, 110, { font: 'Helvetica-Bold', size: 28 });
  two.text('Room to draw on.', 72, 160, { font: 'Helvetica', size: 14 });
  const up = new PdfUpdate(doc.toBuffer());
  const page = up.file.pages()[0];
  const box = up.add(new Dict([
    ['Type', new Name('Annot')], ['Subtype', new Name('Square')], ['Rect', [380, 520, 540, 600]], ['C', [0, 0.3, 0.9]],
    ['T', textString('Ayesha')], ['Contents', textString('Please check these figures.')], ['M', textString("D:20261013100000+05'00'")], ['F', 4],
  ]));
  const dict = new Dict(page.dict.entries());
  dict.set('Annots', [box]);
  up.set(page.ref, dict);
  fs.writeFileSync(file, up.toBuffer());
}

function driver(win, { until, wait }) {
  const wc = win.webContents;
  const js = (code) => wc.executeJavaScript(code);
  const capture = async (name) => {
    if (!process.env.RUTBA_VERIFY_CAPTURE) return;
    wc.invalidate();
    await wait(700);
    fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG());
  };
  const tab = (name) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(), 1`);
  const role = async (r) => {
    const sel = `[data-role="${r}"]`;
    await until(() => js(`Boolean(document.querySelector(${JSON.stringify(sel)}) && !document.querySelector(${JSON.stringify(sel)}).disabled)`), `the ${r} button`, 5000).catch(() => {});
    return js(`(() => { const b = document.querySelector(${JSON.stringify(sel)}); if (!b) return 'no button ${r}'; if (b.disabled) return 'disabled ${r}'; b.click(); return 'clicked'; })()`);
  };
  const tip = (label) => js(`(() => {
    const b = [...document.querySelectorAll('.rw-btn')].find((n) => (n.dataset.tip || n.title || '').startsWith(${JSON.stringify(label)}));
    if (!b) return 'no button ${label}';
    if (b.disabled) return 'disabled ${label}';
    b.click();
    return 'clicked';
  })()`);
  /**
   * Words on page `index` selected, from the first word to the last given, as
   * a person drags across them, once the page's words are laid over it again
   * after the last change (a selection made in a layer about to be redrawn is
   * lost with it).
   */
  const select = async (index, from, to) => {
    await wait(250);
    await until(() => js(`Boolean(document.querySelector('.pdf-page[data-index="${index}"] .textLayer[data-ready]'))`), 'the words laid over the page', 8000).catch(() => {});
    return selectNow(index, from, to);
  };
  const selectNow = (index, from, to) => js(`(() => {
    const layer = document.querySelector('.pdf-page[data-index="${index}"] .textLayer[data-ready]');
    if (!layer) return 'no text';
    const nodes = [];
    const walker = document.createTreeWalker(layer, NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) nodes.push(n);
    const a = nodes.find((n) => n.data.includes(${JSON.stringify(from)}));
    const b = nodes.find((n) => n.data.includes(${JSON.stringify(to)}));
    if (!a || !b) return 'no words';
    const r = document.createRange();
    r.setStart(a, a.data.indexOf(${JSON.stringify(from)}));
    r.setEnd(b, b.data.indexOf(${JSON.stringify(to)}) + ${JSON.stringify(to)}.length);
    getSelection().removeAllRanges();
    getSelection().addRange(r);
    return getSelection().toString();
  })()`);
  /** A drag on page `index`'s comment layer, from and to fractions of the page, through `steps` points between. */
  const dragOn = (index, points) => js(`(async () => {
    const svg = document.querySelector('.pdf-page[data-index="${index}"] .pdf-annots');
    if (!svg) return 'no layer';
    const r = svg.getBoundingClientRect();
    const at = ([fx, fy]) => ({ clientX: r.left + r.width * fx, clientY: r.top + r.height * fy, bubbles: true, cancelable: true, button: 0, buttons: 1 });
    const pts = ${JSON.stringify(points)};
    svg.dispatchEvent(new MouseEvent('mousedown', at(pts[0])));
    for (const p of pts.slice(1)) {
      window.dispatchEvent(new MouseEvent('mousemove', at(p)));
      await new Promise((res) => setTimeout(res, 20));
    }
    window.dispatchEvent(new MouseEvent('mouseup', at(pts[pts.length - 1])));
    return 'dragged';
  })()`);
  const setText = (sel, value) => js(`(() => {
    const el = document.querySelector(${JSON.stringify(sel)});
    if (!el) return false;
    const proto = el.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, ${JSON.stringify(String(value))});
    el.dispatchEvent(new Event('input', { bubbles: true }));
    return true;
  })()`);
  /** The comments the pane lists: type, author, words, status, and how many replies. */
  const cards = () => js(`[...document.querySelectorAll('.pdf-comments .pdf-comment')].map((c) => ({
    id: c.dataset.id, type: c.dataset.type, status: c.dataset.status,
    who: c.querySelector('.pdf-comment-who')?.textContent || '',
    text: c.querySelector(':scope > .pdf-comment-text')?.textContent || '',
    replies: c.querySelectorAll('.pdf-reply').length,
  }))`);
  /** The comments drawn over the pages: their type per page. */
  const onPages = () => js(`[...document.querySelectorAll('.pdf-page')].map((p) => [...p.querySelectorAll('.pdf-annot')].map((a) => a.dataset.type))`);
  return { wc, js, capture, tab, role, tip, select, dragOn, setText, cards, onPages };
}

const answer = (kind, value) => {
  (globalThis.__rutbaCheckDialogAnswers ||= []).push({ kind, answer: value });
};

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args where the files are made (a folder of this block's own is made in it)
 */
export async function verifyPdfAnnot(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const root = path.join(dir, 'pdf-annot');
  fs.mkdirSync(root, { recursive: true });
  const file = path.join(root, 'commented-report.pdf');
  makePdf(file);
  const original = fs.readFileSync(file);

  const win = await open('pdf', file);
  const d = driver(win, { until, wait });
  const count = async (n, label) => until(async () => (await d.cards()).length === n, label, 12000).then(() => true).catch(() => false);

  await until(() => d.js(`document.querySelectorAll('.pdf-page canvas[data-drawn]').length >= 1 && Boolean(document.querySelector('.pdf-page[data-index="0"] .textLayer[data-ready]'))`), 'the first page drawn with its words', 20000).catch(() => {});

  // The Comment tab: the pane opens with the other writer's rectangle listed under her name.
  await d.tab('Comment');
  const paneOpen = await until(() => d.js(`Boolean(document.querySelector('[data-role="comments-pane"]'))`), 'the Comments pane', 5000).then(() => true).catch(() => false);
  await until(async () => (await d.cards()).length === 1, 'the other writer’s comment', 8000).catch(() => {});
  const foreign = await d.cards();
  await until(() => d.js(`Boolean(document.querySelector('.pdf-page[data-index="0"] .pdf-annots[data-ready]'))`), 'the comment layer', 8000).catch(() => {});
  // She saved it with no appearance of its own, so the page draws it from what it is.
  const shown = await until(() => d.js(`Boolean(document.querySelector('.pdf-page[data-index="0"] .pdf-unpainted'))`), 'her rectangle drawn', 6000).then(() => true).catch(() => false);
  check('pdf-annot: the Comment tab opens the Comments pane, listing a rectangle another writer drew with her name and words, and the page shows it though she saved no appearance', paneOpen && shown && foreign.length === 1 && foreign[0].type === 'Square' && foreign[0].who === 'Ayesha' && /check these figures/.test(foreign[0].text),
    JSON.stringify({ paneOpen, shown, foreign }));

  // Highlight: words selected, then the button.
  const selected = await d.select(0, 'quick', 'fox');
  const hl = await d.role('tool-highlight');
  const highlighted = await count(2, 'the highlight listed');
  const drawnHl = await until(async () => (await d.onPages())[0]?.includes('Highlight'), 'the highlight on the page', 8000).then(() => true).catch(() => false);
  await d.capture('pdf-annot-01-highlight.png');
  check('pdf-annot: words selected and Highlight pressed are highlighted, listed in the pane and drawn on the page', selected === 'quick brown fox' && hl === 'clicked' && highlighted && drawnHl,
    JSON.stringify({ selected, hl, highlighted, drawnHl, cards: await d.cards() }));

  // Underline in hand: letting go of a selection underlines it.
  const inHand = await d.role('tool-underline');
  const pressed = await until(() => d.js(`document.querySelector('[data-role="tool-underline"]').getAttribute('aria-pressed') === 'true'`), 'Underline in hand', 3000).then(() => true).catch(() => false);
  await d.select(0, 'Sales', 'region');
  await d.js(`document.querySelector('.pdf-view').dispatchEvent(new MouseEvent('mouseup', { bubbles: true })), 1`);
  const underlined = await count(3, 'the underline listed');
  check('pdf-annot: with Underline in hand, letting go of a selection underlines the words', inHand === 'clicked' && pressed && underlined && (await d.cards()).some((c) => c.type === 'Underline'),
    JSON.stringify({ inHand, pressed, underlined }));

  // Strikethrough, the same way.
  await d.role('tool-strike');
  await d.select(0, 'Costs', 'held');
  await d.js(`document.querySelector('.pdf-view').dispatchEvent(new MouseEvent('mouseup', { bubbles: true })), 1`);
  const struck = await count(4, 'the strikethrough listed');
  check('pdf-annot: Strikethrough draws a line through the words selected', struck && (await d.cards()).some((c) => c.type === 'StrikeOut'), JSON.stringify(await d.cards()));

  // A sticky note, placed and typed into.
  await d.role('tool-note');
  await d.dragOn(0, [[0.82, 0.12], [0.82, 0.12]]);
  const editor = await until(() => d.js(`Boolean(document.querySelector('[data-role="note-editor"] textarea'))`), 'the note editor', 4000).then(() => true).catch(() => false);
  await d.setText('[data-role="note-text"]', 'Is this the final figure?');
  await d.capture('pdf-annot-02-note-editor.png');
  await d.role('note-post');
  const noted = await count(5, 'the note listed');
  const note = (await d.cards()).find((c) => c.type === 'Text');
  check('pdf-annot: a sticky note is placed where the page is clicked, typed into and listed with its words', editor && noted && note?.text === 'Is this the final figure?', JSON.stringify({ editor, noted, note }));

  // A text box, dragged out and typed into.
  await d.role('tool-text');
  await d.dragOn(0, [[0.12, 0.62], [0.3, 0.66], [0.55, 0.7]]);
  await until(() => d.js(`Boolean(document.querySelector('[data-role="note-editor"] textarea'))`), 'the text box editor', 4000).catch(() => {});
  await d.setText('[data-role="note-text"]', 'Approved by the board');
  await d.role('note-post');
  const boxed = await count(6, 'the text box listed');
  check('pdf-annot: a text box dragged out on the page carries the words typed', boxed && (await d.cards()).some((c) => c.type === 'FreeText' && c.text === 'Approved by the board'), JSON.stringify(await d.cards()));

  // Drawing on the second page: a freehand line, a rectangle, an ellipse, a line and an arrow.
  await d.js(`(() => { const v = document.querySelector('.pdf-view'); const p = document.querySelector('.pdf-page[data-index="1"]'); v.scrollTop += p.getBoundingClientRect().top - v.getBoundingClientRect().top - 10; return 1; })()`);
  await until(() => d.js(`Boolean(document.querySelector('.pdf-page[data-index="1"] .pdf-annots[data-ready]'))`), 'the second page’s layer', 8000).catch(() => {});
  const draws = {};
  await d.role('tool-ink');
  draws.ink = await d.dragOn(1, [[0.15, 0.3], [0.2, 0.26], [0.26, 0.32], [0.32, 0.27], [0.38, 0.33], [0.44, 0.28]]);
  await count(7, 'the drawing');
  await d.role('tool-rect');
  draws.rect = await d.dragOn(1, [[0.15, 0.42], [0.3, 0.48], [0.4, 0.52]]);
  await count(8, 'the rectangle');
  await d.role('tool-ellipse');
  draws.ellipse = await d.dragOn(1, [[0.5, 0.42], [0.6, 0.48], [0.75, 0.52]]);
  await count(9, 'the ellipse');
  await d.role('tool-line');
  draws.line = await d.dragOn(1, [[0.15, 0.62], [0.3, 0.63], [0.45, 0.64]]);
  await count(10, 'the line');
  await d.role('tool-arrow');
  draws.arrow = await d.dragOn(1, [[0.5, 0.7], [0.6, 0.66], [0.75, 0.6]]);
  const drawn = await count(11, 'the arrow');
  const kinds = (await d.cards()).map((c) => c.type);
  const second = await until(async () => ((await d.onPages())[1] || []).length === 5, 'five on the second page', 8000).then(async () => (await d.onPages())[1]).catch(async () => (await d.onPages())[1]);
  await d.js(`document.querySelector('[data-role="tool-arrow"]').click(), 1`);
  await d.capture('pdf-annot-03-drawings.png');
  check('pdf-annot: a freehand line, a rectangle, an ellipse, a line and an arrow are dragged out on the second page', drawn && JSON.stringify(second) === JSON.stringify(['Ink', 'Square', 'Circle', 'Line', 'Line']),
    JSON.stringify({ draws, kinds, second }));

  // Reply, resolve, and delete with Undo and Redo.
  const noteId = (await d.cards()).find((c) => c.type === 'Text')?.id;
  await d.js(`document.querySelector('.pdf-comment[data-id="${noteId}"] .pdf-comment-meta')?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })), 1`);
  await until(() => d.js(`Boolean(document.querySelector('.pdf-comment.picked [data-role="reply-text"]'))`), 'the reply box', 4000).catch(() => {});
  await d.setText('.pdf-comment.picked [data-role="reply-text"]', 'Yes, it is final.');
  await d.role('reply-post');
  const replied = await until(async () => (await d.cards()).find((c) => c.id === noteId)?.replies === 1, 'the reply', 8000).then(() => true).catch(() => false);
  await d.js(`document.querySelector('.pdf-comment[data-id="${noteId}"] .pdf-comment-resolve').click(), 1`);
  const resolved = await until(async () => (await d.cards()).find((c) => c.id === noteId)?.status === 'Completed', 'resolved', 8000).then(() => true).catch(() => false);
  await d.capture('pdf-annot-04-pane.png');
  check('pdf-annot: a comment is answered and marked resolved in the pane', replied && resolved, JSON.stringify({ replied, resolved, card: (await d.cards()).find((c) => c.id === noteId) }));

  const rectId = (await d.cards()).find((c) => c.type === 'Square' && c.who !== 'Ayesha')?.id;
  await d.js(`document.querySelector('.pdf-comment[data-id="${rectId}"] .pdf-comment-delete').click(), 1`);
  const deleted = await count(10, 'the rectangle deleted');
  await d.tip('Undo');
  const back = await count(11, 'the rectangle back');
  await d.tip('Redo');
  const gone = await count(10, 'the rectangle gone again');
  check('pdf-annot: a comment is deleted; Undo brings it back and Redo takes it away again', deleted && back && gone, JSON.stringify({ deleted, back, gone }));

  // Save As: the original, then the comments, read back.
  await d.tab('Home');
  const saved = path.join(root, 'commented-report-saved.pdf');
  answer('save', saved);
  await d.tip('Save As');
  const wrote = await until(() => fs.existsSync(saved) && fs.statSync(saved).size > original.length, 'the saved file', 10000).then(() => true).catch(() => false);
  let readBack = {};
  if (wrote) {
    const bytes = fs.readFileSync(saved);
    const pdf = openPdf(bytes);
    const list = readAnnotations(pdf);
    readBack = {
      prefix: bytes.subarray(0, original.length).equals(original),
      repaired: pdf.repaired,
      top: list.filter((a) => !a.replyTo).map((a) => `${a.page}:${a.type}${a.arrow ? '>' : ''}`).join(' '),
      ayesha: list.some((a) => a.author === 'Ayesha' && a.type === 'Square'),
      note: list.find((a) => a.type === 'Text' && !a.replyTo)?.contents,
      status: list.find((a) => a.type === 'Text' && !a.replyTo)?.status,
      reply: list.find((a) => a.replyTo && !a.state)?.contents,
      appearances: list.filter((a) => !a.replyTo && a.author !== 'Ayesha').every((a) => a.hasAppearance),
      signed: list.filter((a) => a.author !== 'Ayesha').every((a) => a.author),
    };
  }
  check('pdf-annot: Save As writes the original bytes unchanged with the comments appended, each read back with its words, author and appearance', wrote && readBack.prefix && !readBack.repaired && readBack.ayesha
    && readBack.top === '0:Square 0:Highlight 0:Underline 0:StrikeOut 0:Text 0:FreeText 1:Ink 1:Circle 1:Line 1:Line>' && readBack.note === 'Is this the final figure?' && readBack.status === 'Completed'
    && readBack.reply === 'Yes, it is final.' && readBack.appearances && readBack.signed, JSON.stringify({ wrote, ...readBack }));
  await d.tab('Comment');
  await d.capture('pdf-annot-05-saved.png');
  const errors = await errorsIn(win);
  check('pdf-annot: no error was shown', !errors.length, errors.join(' | '));
}
