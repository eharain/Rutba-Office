// Worksheets: Page Layout → Arrange for a sheet's pictures, shapes and charts.
//
// A click picks a shape — eight handles and a rotation handle; Ctrl+click
// adds two more and Align Left lines them up; Bring to Front puts one on
// top; Rotate Right 90° and Flip Horizontal turn one; Group gathers two into
// one that moves as one, and Ungroup puts them back; the Selection Pane
// lists everything front first, hides one with its eye and renames one; a
// picture is dragged and resized by hand, and a shape turned by its handle.
// Saved, the drawing part says it all as Excel writes it. Run alone with
// RUTBA_VERIFY_ONLY=sheetarrange.
import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { gradientPng } from './sample-picture.js';

/** The fixture: a small list and three shapes and a picture beside it. */
function fixture() {
  const view = new SheetView(buildXlsx({ sheets: [{ name: 'Plan', rows: [['Task', 'Days'], ['Draft', 3], ['Review', 2], ['Print', 1]] }] }));
  for (const [row, col, geometry, text] of [[1, 3, 'roundRect', 'Draft'], [4, 5, 'ellipse', 'Review'], [8, 4, 'rightArrow', 'Print']]) {
    view.select(row, col);
    view.select(row + 2, col + 1, { extend: true });
    view.insertShape({ geometry, text });
  }
  view.select(2, 9);
  view.insertPicture({ name: 'Logo', contentType: 'image/png', data: gradientPng(60, 40, [43, 95, 217], [15, 157, 88]), widthPx: 120, heightPx: 80 });
  return view.save();
}

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, capture, doc, sessionFor
 * @param {{ dir: string }} where the fixtures are written
 */
export async function verifySheetArrange(h, { dir }) {
  const { open, check, until, wait, errorsIn, capture, doc, sessionFor } = h;
  const file = path.join(dir, 'arrange.xlsx');
  fs.writeFileSync(file, fixture());
  try {
    const win = await open('sheets', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = sessionFor('sheet');
    const model = () => doc.model({ id: session.id });
    await until(() => js(`document.querySelectorAll('.sh-drawing').length === 4`), 'the drawings', 8000);

    const clickIn = async (title) => {
      const find = `[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || n.textContent || '').trim().startsWith(${JSON.stringify(title)}) && !n.disabled)`;
      await until(() => js(`Boolean(${find})`), `the ${title} button`, 3000).catch(() => {});
      return js(`(() => { const b = ${find}; if (!b) return 'no button ' + ${JSON.stringify(title)}; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked'; })()`);
    };
    const menuItem = async (label) => {
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}))`), `the ${label} item`, 4000).catch(() => {});
      return js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}); if (!b) return 'no item'; b.click(); return 'clicked'; })()`);
    };
    const tab = (name) => js(`(() => { [...document.querySelectorAll('.rw-ribbon button, .rw-tabs button, button')].find((b) => b.textContent.trim() === ${JSON.stringify(name)})?.click(); return 1; })()`);
    const refresh = async () => { win.webContents.invalidate(); await wait(700); };
    const pick = (name, add = false) => js(`(() => {
      const n = document.querySelector('.sh-drawing[data-name=${JSON.stringify(name)}]');
      if (!n) return 'no ' + ${JSON.stringify(name)};
      n.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0, ctrlKey: ${add} }));
      window.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, button: 0 }));
      return 'picked';
    })()`);
    const drawing = (name) => (model().drawings || []).find((d) => d.name === name);
    const rect = (selector) => js(`(() => { const n = document.querySelector(${JSON.stringify(selector)}); if (!n) return null; const b = n.getBoundingClientRect(); return { x: b.left, y: b.top, w: b.width, h: b.height }; })()`);
    const drag = async (selector, dx, dy) => {
      const r = await rect(selector);
      if (!r) return 'no ' + selector;
      const x0 = Math.round(r.x + r.w / 2);
      const y0 = Math.round(r.y + r.h / 2);
      win.webContents.sendInputEvent({ type: 'mouseMove', x: x0, y: y0 });
      win.webContents.sendInputEvent({ type: 'mouseDown', x: x0, y: y0, button: 'left', clickCount: 1 });
      for (let i = 1; i <= 8; i++) {
        win.webContents.sendInputEvent({ type: 'mouseMove', x: Math.round(x0 + (dx * i) / 8), y: Math.round(y0 + (dy * i) / 8), button: 'left' });
        await wait(40);
      }
      win.webContents.sendInputEvent({ type: 'mouseUp', x: x0 + dx, y: y0 + dy, button: 'left', clickCount: 1 });
      return 'dragged';
    };

    // ── A click picks a shape: its ring, eight handles and a rotation handle.
    await pick('Shape 2');
    const handles = await until(() => js(`document.querySelectorAll('.sh-obj-handle').length === 8 && Boolean(document.querySelector('.sh-obj-rotate'))`), 'the handles', 4000).catch(() => false);
    await tab('Page Layout');
    const live = await js(`['Bring Forward', 'Send Backward', 'Selection Pane', 'Align', 'Group', 'Rotate'].map((l) => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || n.title || n.textContent || '').trim().startsWith(l)); return b ? (b.disabled ? l + ' off' : l) : l + ' missing'; })`);
    check('sheets: a click picks a shape — a ring, eight handles and a rotation handle — and Page Layout\'s Arrange buttons are live',
      handles === true && JSON.stringify(live) === JSON.stringify(['Bring Forward', 'Send Backward', 'Selection Pane', 'Align off', 'Group', 'Rotate']),
      `handles ${handles}; ${JSON.stringify(live)}`);
    await refresh();
    await capture(win, 'sheets-arrange-picked.png');

    // ── Ctrl+click two more, Align Left.
    await pick('Shape 3', true);
    await pick('Shape 4', true);
    await until(() => js(`document.querySelectorAll('.sh-obj-ring.several').length === 3`), 'three picked', 4000).catch(() => {});
    await clickIn('Align');
    await menuItem('Align Left');
    const aligned = await until(() => { const xs = ['Shape 2', 'Shape 3', 'Shape 4'].map((n) => drawing(n)?.x); return xs.every((x) => x === xs[0]); }, 'one left edge', 6000).catch(() => false);
    check('sheets: Ctrl+click picks several; Align Left lines them up on the leftmost edge', aligned === true,
      JSON.stringify(['Shape 2', 'Shape 3', 'Shape 4'].map((n) => drawing(n)?.x)));
    await refresh();
    await capture(win, 'sheets-arrange-align.png');

    // ── Bring to Front.
    await pick('Shape 2');
    await clickIn('Bring Forward');
    await menuItem('Bring to Front');
    const front = await until(() => { const o = model().objects || []; return o[o.length - 1]?.name === 'Shape 2'; }, 'Shape 2 in front', 6000).catch(() => false);
    check('sheets: Bring to Front moves the picked shape to the end of the drawing part — drawn over the others', front === true, JSON.stringify((model().objects || []).map((o) => o.name)));

    // ── Rotate Right 90° and Flip Horizontal.
    await pick('Shape 4');
    await clickIn('Rotate');
    await menuItem('Rotate Right 90°');
    await until(() => drawing('Shape 4')?.rot === 90, 'turned', 6000).catch(() => {});
    await clickIn('Rotate');
    await menuItem('Flip Horizontal');
    const turned = await until(() => drawing('Shape 4')?.rot === 90 && drawing('Shape 4')?.flipH === true, 'turned and flipped', 6000).catch(() => false);
    const drawnTurn = await until(() => js(`/rotate\\(90 /.test(document.querySelector('.sh-drawing[data-name="Shape 4"]')?.innerHTML || '')`), 'the turn drawn', 4000).catch(() => false);
    check('sheets: Rotate Right 90° and Flip Horizontal turn the arrow and mirror it, drawn so', turned === true && drawnTurn === true, JSON.stringify({ rot: drawing('Shape 4')?.rot, flipH: drawing('Shape 4')?.flipH, drawnTurn }));

    // ── Group two, drag the group, Ungroup.
    await pick('Shape 2');
    await pick('Shape 3', true);
    await clickIn('Group');
    await menuItem('Group');
    const grouped = await until(() => (model().drawings || []).some((d) => d.kind === 'group' && d.members?.length === 2), 'the group', 6000).catch(() => false);
    const g = (model().drawings || []).find((d) => d.kind === 'group');
    const members = await js(`document.querySelectorAll('.sh-drawing[data-kind="group"] .sh-member svg').length`);
    await refresh();
    await capture(win, 'sheets-arrange-group.png');
    await drag('.sh-drawing[data-kind="group"]', 60, 40);
    const groupMoved = await until(() => { const n = (model().drawings || []).find((d) => d.kind === 'group'); return n && Math.abs(n.x - g.x - 60) <= 2 && Math.abs(n.y - g.y - 40) <= 2; }, 'the group moved', 6000).catch(() => false);
    const s3 = drawing('Shape 3');
    await clickIn('Group');
    await menuItem('Ungroup');
    const ungrouped = await until(() => !(model().drawings || []).some((d) => d.kind === 'group') && drawing('Shape 3'), 'ungrouped', 6000).catch(() => false);
    check('sheets: Group gathers two shapes into one drawn member by member, dragged as one; Ungroup puts each back where the group took it',
      grouped === true && members === 2 && groupMoved === true && ungrouped === true && s3 === undefined,
      `grouped ${grouped}, members ${members}, moved ${groupMoved}, ungrouped ${ungrouped}; Shape 3 now ${JSON.stringify(drawing('Shape 3') && { x: drawing('Shape 3').x, y: drawing('Shape 3').y })}`);

    // ── The Selection Pane.
    await clickIn('Selection Pane');
    const pane = await until(() => js(`document.querySelectorAll('.sh-selpane-row').length === 4`), 'the pane', 5000).catch(() => false);
    const rows = await js(`[...document.querySelectorAll('.sh-selpane-row')].map((r) => r.dataset.name)`);
    await js(`(() => { document.querySelector('.sh-selpane-row[data-name="Logo"] .sh-selpane-eye').click(); return 1; })()`);
    const hidden = await until(() => js(`!document.querySelector('.sh-drawing[data-name="Logo"]') && document.querySelector('.sh-selpane-row[data-name="Logo"]').classList.contains('hidden')`), 'the logo hidden', 5000).catch(() => false);
    await js(`(() => { document.querySelector('.sh-selpane-row[data-name="Shape 4"]').dispatchEvent(new MouseEvent('dblclick', { bubbles: true })); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.sh-selpane-name'))`), 'the name box', 3000).catch(() => {});
    await js(`(() => {
      const el = document.querySelector('.sh-selpane-name');
      Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set.call(el, 'Next step arrow');
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
      return 1;
    })()`);
    const renamed = await until(() => (model().objects || []).some((o) => o.name === 'Next step arrow'), 'renamed', 5000).catch(() => false);
    await refresh();
    await capture(win, 'sheets-selection-pane.png');
    check('sheets: the Selection Pane lists every drawing front first; its eye hides the picture and a double-click renames the arrow',
      pane === true && rows[0] === 'Shape 2' && rows.length === 4 && hidden === true && renamed === true,
      `pane ${pane}; ${JSON.stringify(rows)}; hidden ${hidden}; renamed ${renamed}`);
    await js(`(() => { document.querySelector('.sh-selpane-row[data-name="Logo"] .sh-selpane-eye').click(); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.sh-drawing[data-name="Logo"]'))`), 'the logo back', 5000).catch(() => {});
    await js(`(() => { document.querySelector('.sh-selpane-head .rw-btn')?.click(); return 1; })()`);

    // ── A picture moved and resized by hand; a shape turned by its handle.
    const logo = drawing('Logo');
    await pick('Logo');
    await drag('.sh-drawing[data-name="Logo"]', -80, 60);
    const moved = await until(() => { const d = drawing('Logo'); return d && Math.abs(d.x - logo.x + 80) <= 2 && Math.abs(d.y - logo.y - 60) <= 2; }, 'the picture moved', 6000).catch(() => false);
    await drag('.sh-obj-handle[data-handle="se"]', 40, 30);
    const resized = await until(() => { const d = drawing('Logo'); return d && Math.abs(d.width - logo.width - 40) <= 2 && Math.abs(d.height - logo.height - 30) <= 2; }, 'the picture resized', 6000).catch(() => false);
    await pick('Shape 3');
    await until(() => js(`Boolean(document.querySelector('.sh-obj-rotate'))`), 'the rotation handle', 3000).catch(() => {});
    const r = await rect('.sh-drawing[data-name="Shape 3"]');
    const knob = await rect('.sh-obj-rotate');
    // From the knob above the centre to a point level with the centre, to its right: a quarter turn.
    const cx = r.x + r.w / 2;
    const cy = r.y + r.h / 2;
    win.webContents.sendInputEvent({ type: 'mouseMove', x: Math.round(knob.x + 6), y: Math.round(knob.y + 6) });
    win.webContents.sendInputEvent({ type: 'mouseDown', x: Math.round(knob.x + 6), y: Math.round(knob.y + 6), button: 'left', clickCount: 1 });
    for (let i = 1; i <= 6; i++) {
      const a = (Math.PI / 2) * (i / 6);
      win.webContents.sendInputEvent({ type: 'mouseMove', x: Math.round(cx + 80 * Math.sin(a)), y: Math.round(cy - 80 * Math.cos(a)), button: 'left' });
      await wait(40);
    }
    win.webContents.sendInputEvent({ type: 'mouseUp', x: Math.round(cx + 80), y: Math.round(cy), button: 'left', clickCount: 1 });
    const spun = await until(() => Math.abs((drawing('Shape 3')?.rot ?? 0) - 90) <= 2, 'the shape turned by hand', 6000).catch(() => false);
    check('sheets: a picture is dragged and resized by its handles, and a shape turned a quarter by its rotation handle — the anchor and xfrm follow',
      moved === true && resized === true && spun === true, `moved ${moved}; resized ${resized}; turned ${drawing('Shape 3')?.rot}`);
    await refresh();
    await capture(win, 'sheets-arrange-turned.png');

    // Saved: the drawing part as Excel writes it.
    doc.save({ id: session.id });
    const saved = new SheetView(fs.readFileSync(file));
    const part = saved.pkg.partNames().find((p) => /^xl\/drawings\/drawing\d+\.xml$/.test(p));
    const xml = saved.pkg.text(part);
    const order = [...xml.matchAll(/<xdr:cNvPr id="\d+" name="([^"]+)"/g)].map((m) => m[1]);
    check('sheets: saved, the drawing part keeps the order, the turns (rot, flipH), the new name and the moved picture\'s anchor',
      /name="Next step arrow"/.test(xml) && /<a:xfrm rot="5400000" flipH="1">/.test(xml) && order[order.length - 1] === 'Shape 2' && /<xdr:ext cx="1524000" cy="1047750"\/>/.test(xml),
      `order ${JSON.stringify(order)}; ${(xml.match(/<a:xfrm[^>]*>/g) || []).join(' ')}`);

    const complaints = await errorsIn(win);
    check('sheets: the Arrange checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('sheets: the Arrange checks ran', false, err.message);
  }
}
