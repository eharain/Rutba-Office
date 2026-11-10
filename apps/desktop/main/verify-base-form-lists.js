// A form's lists and subforms, as the sample films found them (2026-10-22):
// a list's words fit in its box, not cut to their tops, and a subform's
// field that points at another table shows that record's words, not its
// number. On Larkspur Corner Store's Sales form. Run alone with
// RUTBA_VERIFY_ONLY=base-form-lists.

import fs from 'node:fs';
import path from 'node:path';
import { Base, buildSample } from '@rutba/database';

export async function verifyBaseFormLists({ open, check, until, wait }, { dir }) {
  const file = path.join(dir, 'Form lists.rbase');
  try {
    for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
    const b = Base.open(file);
    buildSample(b, 'shop', { withData: true });
    // The Sales form: lists of its own (the customer, the payment) and the sale's lines, each pointing at a product.
    const pointing = (f) => (f.layout.subforms || []).some((s) => s.columns.some((c) => b.relationships().some((r) => r.from.table === s.table && r.from.field === c && c !== s.field)));
    const form = b.forms().find((f) => f.name === 'Sales' && pointing(f)) || b.forms().find((f) => pointing(f) && f.layout.controls.some((c) => c.kind === 'list'));
    const sub = form.layout.subforms.find((s) => s.columns.some((c) => b.relationships().some((r) => r.from.table === s.table && r.from.field === c && c !== s.field)));
    const col = sub.columns.find((c) => c !== sub.field && b.relationships().some((r) => r.from.table === sub.table && r.from.field === c));
    b.close();

    const win = await open('base', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('[data-form=${JSON.stringify(form.name)}]'))`), 'the database', 10000);
    await js(`document.querySelector('[data-form=${JSON.stringify(form.name)}]').click(), 1`);
    await until(() => js(`document.querySelectorAll('[data-subform=${JSON.stringify(sub.table)}] tbody tr').length > 0`), 'the subform\'s records', 8000);
    await wait(500);
    const seen = await js(`(() => {
      const list = document.querySelector('.bs-fc-list');
      const cs = list ? getComputedStyle(list) : null;
      const line = cs ? parseFloat(cs.fontSize) * 1.2 : 0;
      const room = list ? list.clientHeight - parseFloat(cs.paddingTop) - parseFloat(cs.paddingBottom) : 0;
      const cells = [...document.querySelectorAll('[data-subform=${JSON.stringify(sub.table)}] tbody td[data-col=${JSON.stringify(col)}]')].map((n) => n.textContent);
      return { form: ${JSON.stringify(form.name)}, field: list?.dataset.field, height: list?.clientHeight, room, line, cells: cells.slice(0, 4) };
    })()`);
    // A date and time typed in the form (its box answers to the minute) is kept.
    const when = await js(`(() => { const el = document.querySelector('.bs-fc-datetime'); if (!el) return null; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, '2025-11-01T08:15'); el.dispatchEvent(new Event('input', { bubbles: true })); el.dispatchEvent(new Event('change', { bubbles: true })); el.focus(); el.blur(); return el.dataset.field; })()`);
    const readWhen = () => { const x = Base.open(file); try { return x.rows(form.layout.table, { limit: 1 })[0]?.[when]; } finally { x.close(); } };
    // Kept once the save has reached the file, which a busy machine takes longer over.
    if (when) await until(() => readWhen() === '2025-11-01T08:15:00', 'the date and time kept', 8000).catch(() => {});
    const keptWhen = when ? readWhen() : null;
    seen.when = { field: when, kept: keptWhen };
    // A list too long to hold (a line's Sale, one of 20,000): a box showing the sale by its receipt number, searched as typed.
    const lines = await js(`Boolean(document.querySelector('[data-form="Sale Lines"]'))`);
    if (lines) {
      await js(`document.querySelector('[data-form="Sale Lines"]').click(), 1`);
      await until(() => js(`Boolean(document.querySelector('.bs-fc-search[data-field="Sale"] input'))`), 'the sale box', 8000).catch(() => {});
      await until(() => js(`(document.querySelector('.bs-fc-search[data-field="Sale"] input')?.value || '') !== ''`), 'the sale named', 5000).catch(() => {});
      const shownSale = await js(`document.querySelector('.bs-fc-search[data-field="Sale"] input')?.value || ''`);
      await js(`(() => { const el = document.querySelector('.bs-fc-search[data-field="Sale"] input'); el.focus(); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, '00012'); el.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
      await until(() => js(`document.querySelectorAll('.bs-fc-search-list li').length > 0`), 'the matches', 15000).catch(() => {});
      const matches = await js(`[...document.querySelectorAll('.bs-fc-search-list li')].map((n) => n.textContent).slice(0, 3)`);
      // The first match is drawn above the fields under the box, so the mouse finds it there.
      const onTop = await js(`(() => { const li = document.querySelector('.bs-fc-search-list li'); if (!li) return false; const r = li.getBoundingClientRect(); return li.contains(document.elementFromPoint(r.left + r.width / 2, r.top + r.height / 2)); })()`);
      // Typed in, then Enter: the focus stays in the box, its record's words shown. Clicked again there, the words are chosen, so typing replaces them.
      const wc = win.webContents;
      const at = await js(`(() => { const el = document.querySelector('.bs-fc-search[data-field="Sale"] input'); el.blur(); const r = el.getBoundingClientRect(); return { x: Math.round(r.right - 12), y: Math.round(r.top + r.height / 2) }; })()`);
      await wait(300);
      const click = () => { wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 }); wc.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1 }); };
      click();
      await wait(200);
      for (const ch of '00012') { wc.sendInputEvent({ type: 'keyDown', keyCode: ch }); wc.sendInputEvent({ type: 'char', keyCode: ch }); wc.sendInputEvent({ type: 'keyUp', keyCode: ch }); }
      await until(() => js(`[...document.querySelectorAll('.bs-fc-search-list li')].some((n) => /00012/.test(n.textContent))`), 'the matches typed', 15000).catch(() => {});
      wc.sendInputEvent({ type: 'keyDown', keyCode: 'Enter' });
      wc.sendInputEvent({ type: 'keyUp', keyCode: 'Enter' });
      await wait(400);
      const read = () => js(`(() => { const el = document.querySelector('.bs-fc-search[data-field="Sale"] input'); return { value: el.value, focused: document.activeElement === el, all: el.value.length > 0 && el.selectionStart === 0 && el.selectionEnd === el.value.length }; })()`);
      const picked = await read();
      click();
      await wait(200);
      const chosen = { picked: picked.value, focused: picked.focused, all: (await read()).all };
      seen.search = { shownSale, matches, onTop, chosen };
    }
    if (process.env.RUTBA_VERIFY_CAPTURE) {
      win.webContents.invalidate();
      await wait(700);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'base-form-lists.png'), (await win.webContents.capturePage()).toPNG());
    }
    check('base: a form\'s list has room for its words, and a subform\'s field that points at another table shows that record\'s words, not its number; a date and time typed in a form is kept; a list too long to hold shows its record by its reference and finds others as they are typed, drawn above the fields under it, its words chosen when clicked so typing replaces them',
      Boolean(seen.field) && seen.room >= seen.line && seen.cells.length > 0 && seen.cells.every((c) => c && !/^\d+$/.test(c)) && Boolean(seen.when.field) && seen.when.kept === '2025-11-01T08:15:00'
        && Boolean(seen.search) && /^RCP-/.test(seen.search.shownSale) && seen.search.matches.length > 0 && seen.search.matches.every((m) => /00012/.test(m)) && seen.search.onTop === true
        && seen.search.chosen?.focused === true && seen.search.chosen?.all === true,
      JSON.stringify(seen));
  } catch (err) {
    check('base: the form list checks ran to the end', false, err?.message || String(err));
  }
}
