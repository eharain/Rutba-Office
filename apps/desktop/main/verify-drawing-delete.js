// Deleting a selected shape or picture, in each app, the way a person does.
//
// The owner's report: inserted shapes and pictures, selected, would not go —
// an error, and nothing removed. In Documents a shape or chart was sent to
// the engine as a picture, which refused it; in Presentations a press on a
// shape kept the focus where it was, so the Delete key never reached the
// slide. The checks before this one focused the page or stage themselves
// before pressing Delete, which hid the second. Here the drawing is pressed
// with the mouse, or has just been inserted, and Delete is pressed with the
// focus wherever that leaves it. Run alone with RUTBA_VERIFY_ONLY=delete.
import fs from 'node:fs';
import path from 'node:path';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { buildPptx } from '@rutba/presentation';
import { openDocx } from '@rutba/doc-view/backends/ooxml';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where the files are written
 */
export async function verifyDrawingDelete(h, { dir }) {
  const { open, check, until, wait, press, errorsIn, doc, sessionFor } = h;

  // A real press and release at the middle of an element: what a mouse does.
  const clickAt = async (wc, selector, button = 'left') => {
    const at = await wc.executeJavaScript(`(() => {
      const el = document.querySelector(${JSON.stringify(selector)});
      if (!el) return null;
      el.scrollIntoView({ block: 'center', inline: 'center' });
      const r = el.getBoundingClientRect();
      return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) };
    })()`);
    if (!at) return false;
    wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button, clickCount: 1 });
    wc.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button, clickCount: 1 });
    await wait(150);
    return true;
  };
  const blur = (wc) => wc.executeJavaScript(`(() => { document.activeElement?.blur?.(); return document.activeElement?.tagName || 'none'; })()`);
  const said = (wc) => wc.executeJavaScript(`document.body.innerText`);
  const clickIn = async (wc, title) => {
    const find = `[...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || n.textContent || '').trim().startsWith(${JSON.stringify(title)}) && !n.disabled)`;
    await until(() => wc.executeJavaScript(`Boolean(${find})`), `the ${title} button`, 3000).catch(() => {});
    return wc.executeJavaScript(`(() => { const b = ${find}; if (!b) return 'no button'; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked'; })()`);
  };
  const clickMenu = async (wc, label) => {
    await until(() => wc.executeJavaScript(`[...document.querySelectorAll('.rw-menu button')].some((n) => n.textContent.trim() === ${JSON.stringify(label)})`), `the ${label} item`, 3000).catch(() => {});
    return wc.executeJavaScript(`(() => { const item = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); if (!item) return 'no item'; item.click(); return 'clicked'; })()`);
  };
  const clickTab = (wc, label) => wc.executeJavaScript(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(label)})?.click(), 'tab'`);

  /* ── Documents: a shape from Insert → Shapes, pressed and deleted ─────── */
  try {
    const file = path.join(dir, 'delete-shape.docx');
    const made = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Before the shape.' }, { text: 'After the shape.' }] }));
    made.collapseTo({ block: 0, offset: 0 });
    made.insertShape({ preset: 'rect', widthPx: 160, heightPx: 90 });
    fs.writeFileSync(file, made.save());
    const win = await open('word', file);
    const wc = win.webContents;
    const session = sessionFor('doc');
    const shapes = () => (doc.model({ id: session.id }).drawings || []).filter((d) => d.kind === 'shape').length;
    await until(() => wc.executeJavaScript(`Boolean(document.querySelector('.wd-page img[data-kind="shape"]'))`), 'the shape drawn', 8000);
    await blur(wc);
    const pressed = await clickAt(wc, '.wd-page img[data-kind="shape"]');
    const picked = await until(() => wc.executeJavaScript(`Boolean(document.querySelector('.wd-page img[data-kind="shape"].picked'))`), 'the shape picked', 3000).catch(() => false);
    await press(wc, 'Delete');
    const gone = await until(() => shapes() === 0, 'the shape deleted', 5000).catch(() => false);
    const blocks = doc.model({ id: session.id }).blocks.map((b) => b.text);
    const text = await said(wc);
    check('word: a shape from Insert → Shapes, pressed and deleted, goes — and the paragraph it stood in with it',
      pressed && picked === true && gone === true && blocks.join('|') === 'Before the shape.|After the shape.' && !/no picture/i.test(text),
      `pressed ${pressed}, picked ${picked}, gone ${gone}; blocks ${JSON.stringify(blocks)}${/no picture/i.test(text) ? '; said "no picture"' : ''}`);
    const complaints = await errorsIn(win);
    check('word: deleting a shape reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the shape-deleting checks ran', false, err.message);
  }

  /* ── Presentations: just inserted, and pressed ─────────────────────────── */
  try {
    const file = path.join(dir, 'delete-shape.pptx');
    fs.writeFileSync(file, buildPptx({ title: 'Shapes', slides: [{ layout: 'title', title: 'Shapes', body: 'Delete them.' }] }));
    const win = await open('slides', file);
    const wc = win.webContents;
    const session = sessionFor('deck');
    const shapes = () => doc.model({ id: session.id }).slide?.shapes || [];
    await until(() => wc.executeJavaScript(`Boolean(document.querySelector('.sl-stage .sl-hit'))`), 'the slide drawn', 8000);
    const insertOval = async () => {
      const before = shapes().length;
      await clickTab(wc, 'Home');
      await wait(150);
      await clickIn(wc, 'Shapes');
      await wait(200);
      await clickMenu(wc, 'Oval');
      await until(() => shapes().length > before, 'an oval to land', 4000).catch(() => {});
      return shapes().slice(-1)[0];
    };

    // Just inserted, and selected by the insert: Delete takes it out.
    const first = await insertOval();
    // As a person would: once it shows selected.
    await until(() => wc.executeJavaScript(`document.querySelector('.sl-hit.selected')?.dataset.shape === ${JSON.stringify(String(first?.id))}`), 'the new oval selected', 4000).catch(() => {});
    await wait(100);
    const holder = await wc.executeJavaScript(`(() => { const a = document.activeElement; return a ? a.tagName + '.' + String(a.className || '').split(' ')[0] : 'none'; })()`);
    await press(wc, 'Delete');
    const firstGone = await until(() => !shapes().some((s) => s.id === first?.id), 'the new oval deleted', 4000).catch(() => false);
    check('slides: a shape just inserted from Home → Shapes goes with the Delete key, without a click first',
      Boolean(first) && holder.startsWith('DIV.sl-stage') && firstGone === true, `oval ${first?.id}; focus on ${holder}; gone ${firstGone}; shapes ${shapes().map((s) => s.kind).join(',')}`);

    // Pressed with the mouse, the focus elsewhere: Delete takes it out.
    const second = await insertOval();
    await blur(wc);
    const pressed = await clickAt(wc, `.sl-hit[data-shape="${second?.id}"]`);
    const focused = await wc.executeJavaScript(`document.activeElement?.classList.contains('sl-stage') || false`);
    await press(wc, 'Delete');
    const secondGone = await until(() => !shapes().some((s) => s.id === second?.id), 'the pressed oval deleted', 4000).catch(() => false);
    check('slides: a shape pressed with the mouse takes the keyboard, and goes with the Delete key',
      pressed && focused && secondGone === true, `pressed ${pressed}, stage focused ${focused}, gone ${secondGone}`);
    const complaints = await errorsIn(win);
    check('slides: deleting shapes reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('slides: the shape-deleting checks ran', false, err.message);
  }

  /* ── Worksheets: pressed and deleted, and deleted from its own menu ────── */
  try {
    const file = path.join(dir, 'delete-shape.xlsx');
    fs.writeFileSync(file, buildXlsx({ sheets: [{ name: 'Sheet1', rows: [['Shapes', 1], ['here', 2]] }] }));
    const made = await doc.open({ path: file });
    doc.apply({ id: made.id, ops: [{ op: 'insertShape', geometry: 'rect' }, { op: 'select', row: 1, col: 5 }, { op: 'insertShape', geometry: 'ellipse' }] });
    doc.save({ id: made.id, path: file });
    doc.close({ id: made.id });
    const win = await open('sheets', file);
    const wc = win.webContents;
    const session = sessionFor('sheet');
    const drawings = () => doc.model({ id: session.id }).drawings || [];
    await until(() => wc.executeJavaScript(`document.querySelectorAll('.sh-drawing').length === 2`), 'the two shapes drawn', 8000);

    const firstId = drawings()[0]?.id;
    await blur(wc);
    await clickAt(wc, `.sh-drawing[data-id="${firstId}"]`);
    await press(wc, 'Delete');
    const firstGone = await until(() => !drawings().some((d) => d.id === firstId), 'the pressed shape deleted', 4000).catch(() => false);
    check('sheets: a shape pressed with the mouse goes with the Delete key', firstGone === true, `gone ${firstGone}; ${drawings().length} left`);

    const secondId = drawings()[0]?.id;
    await clickAt(wc, `.sh-drawing[data-id="${secondId}"]`, 'right');
    const chose = await clickMenu(wc, 'Delete');
    const secondGone = await until(() => drawings().length === 0, 'the shape deleted from its menu', 4000).catch(() => false);
    check('sheets: a shape\'s right-click menu offers Delete, and it goes', chose === 'clicked' && secondGone === true, `${chose}; ${drawings().length} left`);
    const complaints = await errorsIn(win);
    check('sheets: deleting shapes reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('sheets: the shape-deleting checks ran', false, err.message);
  }
}
